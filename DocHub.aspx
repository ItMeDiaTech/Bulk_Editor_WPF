<%@ Page Language="C#" MasterPageFile="~masterurl/default.master" Inherits="Microsoft.SharePoint.WebPartPages.WebPartPage,Microsoft.SharePoint,Version=15.0.0.0,Culture=neutral,PublicKeyToken=XXXXXXXXXXXXXXX" %>
<%@ Register TagPrefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=XXXXXXXXXXXXXXX" %>
<%@ Register TagPrefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=XXXXXXXXXXXXXXX" %>
<%@ Register TagPrefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=XXXXXXXXXXXXXXX" %>

<asp:Content ID="PageHead" ContentPlaceHolderID="PlaceHolderAdditionalPageHead" runat="server">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>The Good, The Bad, and The Documented</title>
    
    <style type="text/css">
        /* Modern SharePoint Theme - Fluent UI */
        :root {
            --theme-primary: #0078d4;
            --theme-lighter-alt: #eff6fc;
            --theme-lighter: #deecf9;
            --theme-light: #c7e0f4;
            --theme-tertiary: #71afe5;
            --theme-secondary: #2b88d8;
            --theme-dark-alt: #106ebe;
            --theme-dark: #005a9e;
            --theme-darker: #004578;
            --neutral-lighter-alt: #faf9f8;
            --neutral-lighter: #f3f2f1;
            --neutral-light: #edebe9;
            --neutral-quaternary-alt: #e1dfdd;
            --neutral-quaternary: #d0d0d0;
            --neutral-tertiary-alt: #c8c6c4;
            --neutral-tertiary: #a19f9d;
            --neutral-secondary: #605e5c;
            --neutral-primary-alt: #3b3a39;
            --neutral-primary: #323130;
            --neutral-dark: #201f1e;
            --black: #000000;
            --white: #ffffff;
        }

        /* Reset SharePoint default styles */
        #s4-workspace {
            background: var(--white) !important;
        }
        
        .ms-webpartPage-root {
            border: none !important;
            background: transparent !important;
        }
        
        #s4-bodyContainer {
            background: var(--white) !important;
        }

        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, 'Roboto', 'Helvetica Neue', Arial, sans-serif !important;
            line-height: 1.5;
            color: var(--neutral-primary);
            background-color: var(--white);
            overflow-x: hidden;
        }

        /* Header Section */
        .site-header {
            background: linear-gradient(135deg, var(--theme-primary) 0%, var(--theme-dark) 100%);
            color: var(--white);
            padding: 2rem 0;
            text-align: center;
            position: relative;
            overflow: hidden;
            margin: -10px -10px 0 -10px;
        }

        .site-header::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: url('data:image/svg+xml,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><defs><pattern id="grid" width="10" height="10" patternUnits="userSpaceOnUse"><path d="M 10 0 L 0 0 0 10" fill="none" stroke="rgba(255,255,255,0.1)" stroke-width="0.5"/></pattern></defs><rect width="100" height="100" fill="url(%23grid)"/></svg>');
            opacity: 0.3;
        }

        .site-header .container {
            position: relative;
            z-index: 1;
        }

        .site-title {
            font-size: 3rem;
            font-weight: 600;
            margin-bottom: 0.5rem;
            text-shadow: 0 2px 4px rgba(0,0,0,0.2);
        }

        .site-subtitle {
            font-size: 1.25rem;
            opacity: 0.9;
            font-weight: 300;
        }

        /* Container */
        .container {
            max-width: 1200px;
            margin: 0 auto;
            padding: 0 24px;
        }

        /* Main Content */
        .main-content {
            padding: 3rem 0;
        }

        /* Quick Actions */
        .quick-actions {
            margin-bottom: 3rem;
        }

        .section-title {
            font-size: 2rem;
            font-weight: 600;
            color: var(--neutral-primary);
            margin-bottom: 1.5rem;
            text-align: center;
        }

        .actions-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
            gap: 1.5rem;
            margin-top: 2rem;
        }

        .action-card {
            background: var(--white);
            border: 1px solid var(--neutral-light);
            border-radius: 8px;
            padding: 2rem;
            text-align: center;
            transition: all 0.3s ease;
            cursor: pointer;
            box-shadow: 0 1px 3px rgba(0,0,0,0.08);
            position: relative;
            overflow: hidden;
        }

        .action-card::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            height: 4px;
            background: var(--theme-primary);
            transform: scaleX(0);
            transition: transform 0.3s ease;
        }

        .action-card:hover {
            transform: translateY(-4px);
            box-shadow: 0 8px 24px rgba(0,0,0,0.12);
            border-color: var(--theme-primary);
        }

        .action-card:hover::before {
            transform: scaleX(1);
        }

        .action-icon {
            width: 48px;
            height: 48px;
            margin: 0 auto 1rem;
            background: var(--theme-lighter);
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 24px;
            color: var(--theme-primary);
        }

        .action-title {
            font-size: 1.25rem;
            font-weight: 600;
            color: var(--neutral-primary);
            margin-bottom: 0.5rem;
        }

        .action-description {
            color: var(--neutral-secondary);
            font-size: 0.875rem;
        }

        /* Stats Section */
        .stats-section {
            background: var(--neutral-lighter-alt);
            border-radius: 12px;
            padding: 2rem;
            margin-bottom: 3rem;
        }

        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 2rem;
            margin-top: 1.5rem;
        }

        .stat-item {
            text-align: center;
        }

        .stat-number {
            font-size: 2.5rem;
            font-weight: 700;
            color: var(--theme-primary);
            display: block;
        }

        .stat-label {
            color: var(--neutral-secondary);
            font-size: 0.875rem;
            font-weight: 500;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }

        /* Recent Issues */
        .recent-issues {
            margin-bottom: 3rem;
        }

        .issues-container {
            background: var(--white);
            border-radius: 8px;
            border: 1px solid var(--neutral-light);
            overflow: hidden;
        }

        .issue-item {
            padding: 1rem 1.5rem;
            border-bottom: 1px solid var(--neutral-light);
            display: flex;
            align-items: center;
            justify-content: space-between;
            transition: background-color 0.2s ease;
        }

        .issue-item:last-child {
            border-bottom: none;
        }

        .issue-item:hover {
            background-color: var(--neutral-lighter-alt);
        }

        .issue-info {
            flex: 1;
        }

        .issue-title {
            font-weight: 600;
            color: var(--neutral-primary);
            margin-bottom: 0.25rem;
        }

        .issue-meta {
            font-size: 0.875rem;
            color: var(--neutral-secondary);
        }

        .issue-priority {
            padding: 0.25rem 0.75rem;
            border-radius: 16px;
            font-size: 0.75rem;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }

        .priority-critical {
            background-color: #ffebee;
            color: #c62828;
        }

        .priority-high {
            background-color: #fff3e0;
            color: #ef6c00;
        }

        .priority-medium {
            background-color: #f3e5f5;
            color: #7b1fa2;
        }

        .priority-low {
            background-color: #e8f5e8;
            color: #2e7d32;
        }

        /* News Section */
        .news-section {
            margin-bottom: 3rem;
        }

        .news-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
            gap: 1.5rem;
            margin-top: 1.5rem;
        }

        .news-card {
            background: var(--white);
            border-radius: 8px;
            border: 1px solid var(--neutral-light);
            overflow: hidden;
            transition: transform 0.2s ease, box-shadow 0.2s ease;
        }

        .news-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 16px rgba(0,0,0,0.1);
        }

        .news-image {
            width: 100%;
            height: 160px;
            background: linear-gradient(45deg, var(--theme-lighter), var(--theme-light));
            display: flex;
            align-items: center;
            justify-content: center;
            color: var(--theme-primary);
            font-size: 2rem;
        }

        .news-content {
            padding: 1.5rem;
        }

        .news-title {
            font-size: 1.125rem;
            font-weight: 600;
            color: var(--neutral-primary);
            margin-bottom: 0.5rem;
            line-height: 1.3;
        }

        .news-excerpt {
            color: var(--neutral-secondary);
            font-size: 0.875rem;
            margin-bottom: 1rem;
            line-height: 1.4;
        }

        .news-date {
            font-size: 0.75rem;
            color: var(--neutral-tertiary);
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }

        /* Primary Button */
        .btn-primary {
            background: var(--theme-primary);
            color: var(--white);
            border: none;
            padding: 0.75rem 2rem;
            border-radius: 4px;
            font-size: 1rem;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s ease;
            text-decoration: none;
            display: inline-block;
            text-align: center;
        }

        .btn-primary:hover {
            background: var(--theme-dark-alt);
            transform: translateY(-1px);
            box-shadow: 0 4px 12px rgba(0,120,212,0.3);
        }

        /* Footer */
        .site-footer {
            background: var(--neutral-lighter);
            padding: 2rem 0;
            text-align: center;
            color: var(--neutral-secondary);
            font-size: 0.875rem;
            margin: 0 -10px -10px -10px;
        }

        /* Mobile Responsive */
        @media (max-width: 768px) {
            .container {
                padding: 0 16px;
            }

            .site-title {
                font-size: 2rem;
            }

            .action-card {
                padding: 1.5rem;
            }

            .actions-grid {
                grid-template-columns: 1fr;
                gap: 1rem;
            }

            .stats-grid {
                grid-template-columns: repeat(2, 1fr);
                gap: 1rem;
            }

            .news-grid {
                grid-template-columns: 1fr;
            }

            .issue-item {
                flex-direction: column;
                align-items: flex-start;
                gap: 0.5rem;
            }
        }

        /* Animations */
        @keyframes fadeInUp {
            from {
                opacity: 0;
                transform: translateY(30px);
            }
            to {
                opacity: 1;
                transform: translateY(0);
            }
        }

        .fade-in-up {
            animation: fadeInUp 0.6s ease-out;
        }

        /* Focus styles */
        .action-card:focus {
            outline: 2px solid var(--theme-primary);
            outline-offset: 2px;
        }

        /* SharePoint Override Styles */
        #s4-ribbonrow {
            display: none !important;
        }
        
        .ms-breadcrumb-top {
            display: none !important;
        }
        
        #s4-titlerow {
            display: none !important;
        }
        
        .ms-webpart-chrome-title {
            display: none !important;
        }
        
        .ms-webpart-chrome {
            border: none !important;
            background: transparent !important;
        }
    </style>
</asp:Content>

<asp:Content ID="PageTitle" ContentPlaceHolderID="PlaceHolderPageTitle" runat="server">
    The Good, The Bad, and The Documented
</asp:Content>

<asp:Content ID="PageTitleInTitleArea" ContentPlaceHolderID="PlaceHolderPageTitleInTitleArea" runat="server">
    The Good, The Bad, and The Documented
</asp:Content>

<asp:Content ID="Main" ContentPlaceHolderID="PlaceHolderMain" runat="server">
    
    <!-- Site Header -->
    <header class="site-header">
        <div class="container">
            <h1 class="site-title">The Good, The Bad, and The Documented</h1>
            <p class="site-subtitle">Streamlined issue tracking and documentation for modern teams</p>
        </div>
    </header>

    <!-- Main Content -->
    <main class="main-content">
        <div class="container">
            <!-- Key Items (formerly Quick Actions) -->
            <section class="quick-actions fade-in-up">
                <h2 class="section-title">Key Items</h2>
                <div class="actions-grid">
                    <div class="action-card" onclick="submitIssue()" tabindex="0" role="button" aria-label="Submit a new issue">
                        <div class="action-icon">📋</div>
                        <h3 class="action-title">Submit Issue</h3>
                        <p class="action-description">Report a new issue for a tool</p>
                    </div>
                    
                    <div class="action-card" onclick="viewHyperlinkMacro()" tabindex="0" role="button" aria-label="View hyperlink fix macro work instructions">
                        <div class="action-icon">📊</div>
                        <h3 class="action-title">Hyperlink Fix Macro: Work Instructions</h3>
                        <p class="action-description">Work instructions on how to install, update, and troubleshoot</p>
                    </div>
                    
                    <div class="action-card" onclick="viewFeedbackScript()" tabindex="0" role="button" aria-label="Access feedback importer script work instructions">
                        <div class="action-icon">📖</div>
                        <h3 class="action-title">Feedback Importer Script: Work Instructions</h3>
                        <p class="action-description">Work instructions on how to install and troubleshoot</p>
                    </div>
                    
                    <div class="action-card" onclick="viewReports()" tabindex="0" role="button" aria-label="View reports">
                        <div class="action-icon">📈</div>
                        <h3 class="action-title">Reports</h3>
                        <p class="action-description">Annual Review Reports, Tool Usage Reports, and more</p>
                    </div>
                </div>
            </section>

            <!-- Current Status -->
            <section class="stats-section fade-in-up">
                <h2 class="section-title">Current Status</h2>
                <div class="stats-grid">
                    <div class="stat-item">
                        <span class="stat-number" id="issuesThisWeek">3</span>
                        <span class="stat-label">Issues this Week</span>
                    </div>
                    <div class="stat-item">
                        <span class="stat-number" id="issuesThisMonth">12</span>
                        <span class="stat-label">Issues this Month</span>
                    </div>
                    <div class="stat-item">
                        <span class="stat-number" id="amountResolved">8</span>
                        <span class="stat-label">Amount Resolved</span>
                    </div>
                    <div class="stat-item">
                        <span class="stat-number" id="openIssues">5</span>
                        <span class="stat-label">Open Issues</span>
                    </div>
                    <div class="stat-item">
                        <span class="stat-number" id="resolvedThisMonth">15</span>
                        <span class="stat-label">Resolved this Month</span>
                    </div>
                    <div class="stat-item">
                        <span class="stat-number" id="resolvedAllTime">247</span>
                        <span class="stat-label">Resolved All Time</span>
                    </div>
                </div>
            </section>

            <!-- Latest Updates (Tool Releases) -->
            <section class="recent-issues fade-in-up">
                <h2 class="section-title">Latest Updates</h2>
                <div class="issues-container" id="latest-updates-container">
                    <div class="issue-item">
                        <div class="issue-info">
                            <div class="issue-title">Hyperlink Fix Macro v2.1</div>
                            <div class="issue-meta">Fixed bug with Excel hyperlinks, improved performance • Released 3 days ago</div>
                        </div>
                        <span class="issue-priority priority-low">Released</span>
                    </div>
                    
                    <div class="issue-item">
                        <div class="issue-info">
                            <div class="issue-title">Feedback Importer Script v1.8</div>
                            <div class="issue-meta">Added support for CSV imports, enhanced error handling • Released 1 week ago</div>
                        </div>
                        <span class="issue-priority priority-low">Released</span>
                    </div>
                    
                    <div class="issue-item">
                        <div class="issue-info">
                            <div class="issue-title">Data Analysis Tool v3.2</div>
                            <div class="issue-meta">New dashboard features, improved data visualization • Released 2 weeks ago</div>
                        </div>
                        <span class="issue-priority priority-low">Released</span>
                    </div>
                    
                    <div class="issue-item">
                        <div class="issue-info">
                            <div class="issue-title">Report Generator v1.5</div>
                            <div class="issue-meta">Updated templates, fixed formatting issues • Released 3 weeks ago</div>
                        </div>
                        <span class="issue-priority priority-low">Released</span>
                    </div>
                </div>
            </section>

            <!-- Recent Issues (from SharePoint list) -->
            <section class="recent-issues fade-in-up">
                <h2 class="section-title">Recent Issues</h2>
                <div class="issues-container" id="recent-issues-container">
                    <div class="issue-item">
                        <div class="issue-info">
                            <div class="issue-title">Database Connection Timeout</div>
                            <div class="issue-meta">SQL Management Studio • Reported 2 hours ago</div>
                        </div>
                        <span class="issue-priority priority-critical">Critical</span>
                    </div>
                    
                    <div class="issue-item">
                        <div class="issue-info">
                            <div class="issue-title">Printer Not Responding</div>
                            <div class="issue-meta">HP LaserJet Pro • Reported 4 hours ago</div>
                        </div>
                        <span class="issue-priority priority-medium">Medium</span>
                    </div>
                    
                    <div class="issue-item">
                        <div class="issue-info">
                            <div class="issue-title">License Expiration Alert</div>
                            <div class="issue-meta">Adobe Creative Suite • Reported 1 day ago</div>
                        </div>
                        <span class="issue-priority priority-low">Low</span>
                    </div>
                    
                    <div class="issue-item">
                        <div class="issue-info">
                            <div class="issue-title">Network Connectivity Issues</div>
                            <div class="issue-meta">Building A WiFi • Reported 3 hours ago</div>
                        </div>
                        <span class="issue-priority priority-high">High</span>
                    </div>
                </div>
            </section>

            <!-- Important News (formerly Latest Updates) -->
            <section class="news-section fade-in-up">
                <h2 class="section-title">Important News</h2>
                <div class="news-grid" id="important-news-container">
                    <article class="news-card">
                        <div class="news-image">📢</div>
                        <div class="news-content">
                            <h3 class="news-title">Tool Management Hub Now Live</h3>
                            <p class="news-excerpt">
                                Our new streamlined platform is now available! Report issues faster and access 
                                documentation more efficiently than ever before.
                            </p>
                            <time class="news-date">Today</time>
                        </div>
                    </article>
                    
                    <article class="news-card">
                        <div class="news-image">🔧</div>
                        <div class="news-content">
                            <h3 class="news-title">New Mobile App Available</h3>
                            <p class="news-excerpt">
                                Submit issues on-the-go with our new mobile-responsive interface. 
                                Access everything from your smartphone or tablet.
                            </p>
                            <time class="news-date">2 days ago</time>
                        </div>
                    </article>
                    
                    <article class="news-card">
                        <div class="news-image">⚡</div>
                        <div class="news-content">
                            <h3 class="news-title">Performance Improvements</h3>
                            <p class="news-excerpt">
                                We've optimized load times and improved the user experience based on 
                                your feedback. Pages now load 50% faster!
                            </p>
                            <time class="news-date">1 week ago</time>
                        </div>
                    </article>
                </div>
            </section>
        </div>
    </main>

    <!-- Footer -->
    <footer class="site-footer">
        <div class="container">
            <p>&copy; 2025 Tool Management Hub. Built with modern SharePoint for efficient team collaboration.</p>
        </div>
    </footer>

    <script type="text/javascript">
        // Initialize page
        document.addEventListener('DOMContentLoaded', function() {
            // Add fade-in animation to elements
            var elements = document.querySelectorAll('.fade-in-up');
            for (var i = 0; i < elements.length; i++) {
                elements[i].style.animationDelay = (i * 0.1) + 's';
            }

            // Update stats
            updateStats();
            
            // Add keyboard navigation
            setupKeyboardNavigation();
            
            // Load real data if SharePoint context is available
            if (typeof _spPageContextInfo !== 'undefined') {
                loadSharePointData();
            }
        });

        // Navigation functions (replace with your actual SharePoint URLs)
        function submitIssue() {
            window.location.href = _spPageContextInfo.webAbsoluteUrl + '/Lists/Issue%20Tracker/NewForm.aspx';
        }

        function viewHyperlinkMacro() {
            window.location.href = _spPageContextInfo.webAbsoluteUrl + '/Work%20Instructions/Hyperlink%20Fix%20Macro.aspx';
        }

        function viewFeedbackScript() {
            window.location.href = _spPageContextInfo.webAbsoluteUrl + '/Work%20Instructions/Feedback%20Importer%20Script.aspx';
        }

        function viewReports() {
            window.location.href = _spPageContextInfo.webAbsoluteUrl + '/Reporting%20Documents/Forms/AllItems.aspx';
        }

        // Update statistics
        function updateStats() {
            var stats = {
                issuesThisWeek: '3',
                issuesThisMonth: '12', 
                amountResolved: '8',
                openIssues: '5',
                resolvedThisMonth: '15',
                resolvedAllTime: '247'
            };

            for (var key in stats) {
                var element = document.getElementById(key);
                if (element) {
                    element.textContent = stats[key];
                }
            }
        }

        // Keyboard navigation
        function setupKeyboardNavigation() {
            var actionCards = document.querySelectorAll('.action-card');
            for (var i = 0; i < actionCards.length; i++) {
                actionCards[i].addEventListener('keydown', function(e) {
                    if (e.key === 'Enter' || e.key === ' ') {
                        e.preventDefault();
                        this.click();
                    }
                });
            }
        }

        // Load real SharePoint data (if available)
        function loadSharePointData() {
            // Statistics
            loadStatistics();
            
            // Recent Issues
            loadRecentIssues();
            
            // Latest Updates
            loadLatestUpdates();
            
            // Important News
            loadImportantNews();
        }

        function loadStatistics() {
            // Get issues this week
            var oneWeekAgo = new Date();
            oneWeekAgo.setDate(oneWeekAgo.getDate() - 7);
            var weekFilter = oneWeekAgo.toISOString();
            
            var oneMonthAgo = new Date();
            oneMonthAgo.setMonth(oneMonthAgo.getMonth() - 1);
            var monthFilter = oneMonthAgo.toISOString();
            
            // Issues this week
            fetch(_spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/getbytitle('Issue Tracker')/items?$filter=Created ge datetime'" + weekFilter + "'&$select=Id", {
                headers: { 'Accept': 'application/json;odata=verbose' }
            })
            .then(function(response) { return response.json(); })
            .then(function(data) {
                var element = document.getElementById('issuesThisWeek');
                if (element) element.textContent = data.d.results.length;
            })
            .catch(function(error) { console.log('Error loading issues this week:', error); });
            
            // Issues this month
            fetch(_spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/getbytitle('Issue Tracker')/items?$filter=Created ge datetime'" + monthFilter + "'&$select=Id", {
                headers: { 'Accept': 'application/json;odata=verbose' }
            })
            .then(function(response) { return response.json(); })
            .then(function(data) {
                var element = document.getElementById('issuesThisMonth');
                if (element) element.textContent = data.d.results.length;
            })
            .catch(function(error) { console.log('Error loading issues this month:', error); });
            
            // Open issues
            fetch(_spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/getbytitle('Issue Tracker')/items?$filter=IssueStatus ne 'Closed' and IssueStatus ne 'Resolved'&$select=Id", {
                headers: { 'Accept': 'application/json;odata=verbose' }
            })
            .then(function(response) { return response.json(); })
            .then(function(data) {
                var element = document.getElementById('openIssues');
                if (element) element.textContent = data.d.results.length;
            })
            .catch(function(error) { console.log('Error loading open issues:', error); });
        }

        function loadRecentIssues() {
            fetch(_spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/getbytitle('Issue Tracker')/items?$select=Id,Title,IssuePriority,IssueStatus,ToolName,Created&$orderby=Created desc&$top=4", {
                headers: { 'Accept': 'application/json;odata=verbose' }
            })
            .then(function(response) { return response.json(); })
            .then(function(data) {
                updateRecentIssuesUI(data.d.results);
            })
            .catch(function(error) { console.log('Error loading recent issues:', error); });
        }

        function loadLatestUpdates() {
            fetch(_spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/getbytitle('Tool Updates')/items?$select=Id,Title,ToolName,Version,ReleaseDate,ChangeDescription&$orderby=ReleaseDate desc&$top=4", {
                headers: { 'Accept': 'application/json;odata=verbose' }
            })
            .then(function(response) { return response.json(); })
            .then(function(data) {
                updateLatestUpdatesUI(data.d.results);
            })
            .catch(function(error) { console.log('Error loading latest updates:', error); });
        }

        function loadImportantNews() {
            var today = new Date().toISOString();
            fetch(_spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/getbytitle('Important News')/items?$select=Id,Title,NewsContent,PublishDate,Priority&$filter=IsActive eq true and PublishDate le datetime'" + today + "'&$orderby=Priority desc,PublishDate desc&$top=3", {
                headers: { 'Accept': 'application/json;odata=verbose' }
            })
            .then(function(response) { return response.json(); })
            .then(function(data) {
                updateImportantNewsUI(data.d.results);
            })
            .catch(function(error) { console.log('Error loading important news:', error); });
        }

        function updateRecentIssuesUI(issues) {
            var container = document.querySelector('#recent-issues-container');
            if (!container) return;
            
            container.innerHTML = '';
            
            for (var i = 0; i < issues.length; i++) {
                var issue = issues[i];
                var issueElement = document.createElement('div');
                issueElement.className = 'issue-item';
                issueElement.innerHTML = 
                    '<div class="issue-info">' +
                        '<div class="issue-title">' + escapeHtml(issue.Title) + '</div>' +
                        '<div class="issue-meta">' + escapeHtml(issue.ToolName || 'Unknown Tool') + ' • ' + formatDate(new Date(issue.Created)) + '</div>' +
                    '</div>' +
                    '<span class="issue-priority priority-' + (issue.IssuePriority || 'medium').toLowerCase() + '">' + (issue.IssuePriority || 'Medium') + '</span>';
                container.appendChild(issueElement);
            }
        }

        function updateLatestUpdatesUI(updates) {
            var container = document.querySelector('#latest-updates-container');
            if (!container) return;
            
            container.innerHTML = '';
            
            for (var i = 0; i < updates.length; i++) {
                var update = updates[i];
                var updateElement = document.createElement('div');
                updateElement.className = 'issue-item';
                updateElement.innerHTML = 
                    '<div class="issue-info">' +
                        '<div class="issue-title">' + escapeHtml(update.Title) + '</div>' +
                        '<div class="issue-meta">' + escapeHtml(update.ChangeDescription || 'No description available') + ' • Released ' + formatDate(new Date(update.ReleaseDate)) + '</div>' +
                    '</div>' +
                    '<span class="issue-priority priority-low">Released</span>';
                container.appendChild(updateElement);
            }
        }

        function updateImportantNewsUI(newsItems) {
            var container = document.querySelector('#important-news-container');
            if (!container) return;
            
            container.innerHTML = '';
            
            for (var i = 0; i < newsItems.length; i++) {
                var item = newsItems[i];
                var newsElement = document.createElement('article');
                newsElement.className = 'news-card';
                newsElement.innerHTML = 
                    '<div class="news-image">📢</div>' +
                    '<div class="news-content">' +
                        '<h3 class="news-title">' + escapeHtml(item.Title) + '</h3>' +
                        '<p class="news-excerpt">' + escapeHtml(truncateText(item.NewsContent || '', 120)) + '</p>' +
                        '<time class="news-date">' + formatDate(new Date(item.PublishDate)) + '</time>' +
                    '</div>';
                container.appendChild(newsElement);
            }
        }

        // Utility functions
        function formatDate(date) {
            var now = new Date();
            var diffInHours = Math.floor((now - date) / (1000 * 60 * 60));
            
            if (diffInHours < 1) return 'Just now';
            if (diffInHours < 24) return diffInHours + ' hour' + (diffInHours > 1 ? 's' : '') + ' ago';
            
            var diffInDays = Math.floor(diffInHours / 24);
            if (diffInDays < 7) return diffInDays + ' day' + (diffInDays > 1 ? 's' : '') + ' ago';
            
            var diffInWeeks = Math.floor(diffInDays / 7);
            return diffInWeeks + ' week' + (diffInWeeks > 1 ? 's' : '') + ' ago';
        }

        function escapeHtml(text) {
            var div = document.createElement('div');
            div.textContent = text;
            return div.innerHTML;
        }

        function truncateText(text, maxLength) {
            if (text.length <= maxLength) return text;
            return text.substring(0, maxLength).trim() + '...';
        }
    </script>

</asp:Content>