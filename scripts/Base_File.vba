Option Explicit



'=============================================================

'  UpdateHyperlinksFromAPI

'=============================================================





Public Sub UpdateHyperlinksFromAPI()

    '----- Suspend UI -----

    Dim doc As Document: Set doc = ActiveDocument

    Dim links As Hyperlinks: Set links = doc.Hyperlinks

    Dim oldSU As Boolean: oldSU = Application.ScreenUpdating

    Application.ScreenUpdating = False

    On Error GoTo CleanExit



    '----- STEP 1 - Remove blank text external links -----

    Dim linkError As Collection

    Set linkError = RemoveInvisibleExternalHyperlinks(links)

    Set links = doc.Hyperlinks    'refresh collection after deletions



    '----- STEP 2 - Collect unique Lookup_IDs -----

    Dim rx As Object: Set rx = CreateObject("VBScript.RegExp")

    With rx

        .Pattern = "(TSRC-[^-]+-[0-9]{6}|CMS-[^-]+-[0-9]{6})"

        .IgnoreCase = True

    End With



    Dim idDict As Object: Set idDict = CreateObject("Scripting.Dictionary")

    idDict.CompareMode = vbTextCompare



    Dim hl As Hyperlink, lookupID As String

    For Each hl In links

        lookupID = ExtractLookupID(hl.Address, hl.SubAddress, rx)

        If Len(lookupID) > 0 Then

            If Not idDict.Exists(lookupID) Then idDict.Add lookupID, True

        End If

    Next hl



    If idDict.Count = 0 Then

        MsgBox "No valid Lookup_IDs found.", vbExclamation

        GoTo CleanExit

    End If



    '----- STEP 3 - Build JSON & POST -----



    Dim arrIDs() As String, idx As Long, vKey As Variant

    ReDim arrIDs(0 To idDict.Count - 1)

    idx = 0

    For Each vKey In idDict.Keys

        arrIDs(idx) = """" & vKey & """"   'wrap in quotes

        idx = idx + 1

    Next vKey



    Dim jsonBody As String

    jsonBody = "{""Lookup_ID"": [" & Join(arrIDs, ",") & "]}"



    Dim http As Object:
<Insert HTTP Request>

    http.setRequestHeader "Content-Type", "application/json"

    http.send jsonBody



    If http.Status <> 200 Then

        MsgBox "API call failed: " & http.Status & " " & http.StatusText, vbCritical

        GoTo CleanExit

    End If



    '----- STEP 4 - Parse JSON -----



    Dim parser As Json2VBA: Set parser = New Json2VBA

    Dim json As Object: Set json = parser.parse(http.responseText)

    Dim currentVersion As String

    Dim updateNotes As String

    Dim needsUpdate As Boolean

    Dim flowVersion As String





    ' CURRENT VERSION USED FOR UPDATING

    currentVersion = "2.1"

    needsUpdate = False

    flowVersion = json("Version")



    If Not currentVersion = json("Version") Then needsUpdate = True

    updateNotes = json("Changes")



    Dim recDict As Object: Set recDict = CreateObject("Scripting.Dictionary")

    recDict.CompareMode = vbTextCompare



    Dim itm

    For Each itm In json("Results")

        If Not recDict.Exists(itm("Document_ID")) Then recDict.Add itm("Document_ID"), itm

        If Not recDict.Exists(itm("Content_ID")) Then recDict.Add itm("Content_ID"), itm

    Next itm



    '----- STEP 5 - Update hyperlinks -----

    Dim results As New Collection, notFound As New Collection, docExpired As New Collection, updatedURL As New Collection

    Dim dispText As String, last6 As String, last5 As String

    Dim changedURL As Boolean, appended As Boolean

    Dim alreadyExpired As Boolean, alreadyNotFound As Boolean

    Dim rec As Object, targetAddress As String, targetSub As String



    For Each hl In links

            On Error Resume Next

            If Err.Number <> 0 Then

                Err.Clear

                On Error GoTo 0

                GoTo NextHL

            End If

        lookupID = ExtractLookupID(hl.Address, hl.SubAddress, rx)

        If Len(lookupID) = 0 Then GoTo NextHL



        dispText = hl.TextToDisplay

        alreadyExpired = InStr(1, dispText, " - Expired", vbTextCompare) > 0

        alreadyNotFound = InStr(1, dispText, " - Not Found", vbTextCompare) > 0



        If recDict.Exists(lookupID) Then

            Set rec = recDict(lookupID)

            targetAddress = "https://thesource.cvshealth.com/nuxeo/thesource/"

            targetSub = "!/view?docid=" & rec("Document_ID")



            changedURL = (hl.Address <> targetAddress) Or (hl.SubAddress <> targetSub)

            If changedURL Then

                With hl

                    .Address = targetAddress

                    .SubAddress = targetSub

                End With

            End If



            If Not alreadyExpired And Not alreadyNotFound Then

                last6 = Right$(rec("Content_ID"), 6)

                last5 = Right$(last6, 5)

                appended = False



                If Right$(dispText, Len(" (" & last5 & ")")) = " (" & last5 & ")" _

                   And Right$(dispText, Len(" (" & last6 & ")")) <> " (" & last6 & ")" Then

                    dispText = Left$(dispText, Len(dispText) - Len(" (" & last5 & ")")) & " (" & last6 & ")"

                    hl.TextToDisplay = dispText

                    appended = True

                ElseIf InStr(1, dispText, " (" & last6 & ")", vbTextCompare) = 0 Then

                    hl.TextToDisplay = Trim$(dispText) & " (" & last6 & ")"

                    appended = True

                End If

                If Not Left(hl.TextToDisplay, Len(hl.TextToDisplay) - 9) = rec("Title") Then

                    updatedURL.Add LogChangedURL(hl, Left(hl.TextToDisplay, Len(hl.TextToDisplay) - 9), rec("Title"), rec("Content_ID"))

                End If

            End If



            If rec("Status") = "Expired" And Not alreadyExpired Then

                hl.TextToDisplay = hl.TextToDisplay & " - Expired"

                docExpired.Add LogLine(hl, "Expired", rec)

            ElseIf changedURL Or appended Then

                results.Add LogLine(hl, _

                    IIf(changedURL, "URL Updated", "") & _

                    IIf(appended, IIf(changedURL, ", ", "") & "Appended Content ID", ""), rec)

            End If

        ElseIf Not alreadyNotFound And Not alreadyExpired Then

            hl.TextToDisplay = hl.TextToDisplay & " - Not Found"

            notFound.Add LogLine(hl, "Not Found", Nothing)

        End If

NextHL:

    Next hl



CleanExit:

    Application.ScreenUpdating = oldSU

    '----- STEP 6 - Write changelog -----



    If results Is Nothing Then Set results = New Collection

    If notFound Is Nothing Then Set notFound = New Collection

    If docExpired Is Nothing Then Set docExpired = New Collection

    If updatedURL Is Nothing Then Set updatedURL = New Collection

    If linkError Is Nothing Then Set linkError = New Collection

    If updateNotes = "" Then updateNotes = "No update notes available."



    WriteChangelog results, notFound, docExpired, linkError, updatedURL, needsUpdate, flowVersion, updateNotes





End Sub



'========================  Helpers  ========================



Private Function ExtractLookupID(addr As String, subAddr As String, rx As Object) As String

    Dim full As String: full = addr & IIf(Len(subAddr) > 0, "#" & subAddr, "")

    If rx.Test(full) Then

        ExtractLookupID = UCase$(rx.Execute(full)(0).value)

    ElseIf InStr(1, full, "docid=", vbTextCompare) > 0 Then

        ExtractLookupID = Trim$(Split(Split(full, "docid=")(1), "&")(0))

    End If

End Function



Private Function LogLine(hLink As Hyperlink, note As String, rec As Object) As String

    Dim pageNum As Variant, lineNum As Variant, result As String



    On Error Resume Next

    pageNum = hLink.Range.Information(wdActiveEndPageNumber)

    lineNum = hLink.Range.Information(wdFirstCharacterLineNumber)

    On Error GoTo 0



    result = "    Page:" & pageNum & " | Line:" & lineNum & " | " & note



    If rec Is Nothing Then

        result = result & vbCrLf & "        " & hLink.TextToDisplay

    Else

        result = result & vbCrLf & "        " & rec("Title") & vbCrLf & "        " & rec("Content_ID")

    End If



    LogLine = result

End Function

Private Function LogChangedURL(hLink As Hyperlink, currentTitle As String, newTitle As String, contentID As String) As String

    LogChangedURL = "    Page:" & hLink.Range.Information(wdActiveEndPageNumber) & _

              " | Line:" & hLink.Range.Information(wdFirstCharacterLineNumber) & _

              " | " & "Possible Title Change" & _

              vbCrLf & "        Current Title: " & currentTitle & vbCrLf & "        New Title:     " & newTitle & _

              vbCrLf & "        Content ID:    " & contentID

End Function



Private Function RemoveInvisibleExternalHyperlinks(links As Hyperlinks) As Collection

    Dim invisibleLinks As New Collection

    Dim pageNum As Long, lineNum As Long

    Dim printLine As String

    printLine = ""

    Dim i As Long

    For i = links.Count To 1 Step -1

        If Trim$(links(i).TextToDisplay) = "" And Len(links(i).Address) > 0 Then

            pageNum = links(i).Range.Information(wdActiveEndPageNumber)

            lineNum = links(i).Range.Information(wdFirstCharacterLineNumber)

            printLine = "    Page:" & pageNum & " | Line:" & lineNum & " | Invisible Hyperlink Deleted"

            invisibleLinks.Add printLine

            links(i).Delete

            End If

    Next i

    Set RemoveInvisibleExternalHyperlinks = invisibleLinks

End Function



Private Function GetDownloadsFolder() As String

    On Error Resume Next

    Dim sh As Object: Set sh = CreateObject("Shell.Application")

    Dim downloadsPath As String

    downloadsPath = sh.Namespace("shell:Downloads").Self.Path

    If downloadsPath = "" Then downloadsPath = Environ("USERPROFILE") & "\Downloads"

    GetDownloadsFolder = downloadsPath

End Function



Private Sub WriteChangelog(updated As Collection, nf As Collection, de As Collection, fe As Collection, uu As Collection, update As Boolean, newVersion As String, changes As String)

    'Creates or appends a uniquely named changelog in the user's Downloads folder.

    Dim folderPath As String, fp As String, cnt As Long, ff As Integer



    folderPath = GetDownloadsFolder() & "\"

    fp = folderPath & "Changelog.txt"



    'If file exists, increment suffix: Changelog_1.txt, Changelog_2.txt, etc.

    Do While Len(Dir$(fp)) > 0

        cnt = cnt + 1

        fp = folderPath & "Changelog_" & cnt & ".txt"

    Loop



    ff = FreeFile

    Open fp For Output As #ff



    Print #ff, "Updated Links (" & updated.Count & "):"

    DumpColl ff, updated



    Print #ff, vbCrLf & "Found Expired (" & de.Count & "):"

    DumpColl ff, de



    Print #ff, vbCrLf & "Not Found (" & nf.Count & "):"

    DumpColl ff, nf



    Print #ff, vbCrLf & "Found Error (" & fe.Count & "):"

    DumpColl ff, fe



    Print #ff, vbCrLf & "Potential Outdated Titles (" & uu.Count & "):"

    DumpColl ff, uu



    If update Then

        Dim formattedChanges As String

        formattedChanges = Replace(changes, "\""", """")

        formattedChanges = Replace(formattedChanges, "\n", vbCrLf)

        formattedChanges = Replace(formattedChanges, "\", "")

        formattedChanges = Replace(formattedChanges, "&nbsp;", "    ")

        Print #ff, vbCrLf & vbCrLf & "****NEW UPDATE (" & newVersion & ")****"

        Print #ff, vbCrLf & formattedChanges

    End If



    Close #ff



    MsgBox "Changelog saved to " & fp, vbInformation

    Shell "cmd /c start """" """ & fp & """", vbNormalFocus

End Sub



Private Sub DumpColl(ff As Integer, c As Collection)

    Dim v

    If Not c Is Nothing Then

        For Each v In c

            Print #ff, v

        Next v

    End If

End Sub