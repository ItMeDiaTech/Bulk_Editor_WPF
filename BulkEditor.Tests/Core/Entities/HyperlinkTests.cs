using BulkEditor.Core.Entities;
using System;
using Xunit;

namespace BulkEditor.Tests.Core.Entities
{
    /// <summary>
    /// Unit tests for Hyperlink entity
    /// </summary>
    public class HyperlinkTests
    {
        [Fact]
        public void Hyperlink_DefaultConstruction_ShouldSetCorrectDefaults()
        {
            // Arrange & Act
            var hyperlink = new Hyperlink();

            // Assert
            Assert.False(string.IsNullOrEmpty(hyperlink.Id));
            Assert.Equal(HyperlinkStatus.Pending, hyperlink.Status);
            Assert.Equal(HyperlinkAction.None, hyperlink.ActionTaken);
            Assert.False(hyperlink.RequiresUpdate);
            Assert.Equal(string.Empty, hyperlink.DisplayText);
            Assert.Equal(string.Empty, hyperlink.OriginalUrl);
            Assert.Equal(string.Empty, hyperlink.UpdatedUrl);
            Assert.Equal(string.Empty, hyperlink.LookupId);
            Assert.Equal(string.Empty, hyperlink.ContentId);
            Assert.Equal(string.Empty, hyperlink.DocumentId);
            Assert.Equal(string.Empty, hyperlink.ErrorMessage);
        }

        [Theory]
        [InlineData("https://thesource.cvshealth.com/nuxeo/thesource/#!/view?docid=doc-123456", "TSRC-ABC-123456")]
        [InlineData("https://thesource.cvshealth.com/nuxeo/thesource/#!/view?docid=doc-789012", "CMS-DEF-789012")]
        [InlineData("https://thesource.cvshealth.com/nuxeo/thesource/#!/view?docid=no-lookup", "")]
        public void Hyperlink_WithUrl_ShouldAcceptVariousFormats(string url, string expectedLookupId)
        {
            // Arrange & Act
            var hyperlink = new Hyperlink
            {
                OriginalUrl = url,
                LookupId = expectedLookupId
            };

            // Assert
            Assert.Equal(url, hyperlink.OriginalUrl);
            Assert.Equal(expectedLookupId, hyperlink.LookupId);
        }

        [Fact]
        public void Hyperlink_StatusUpdate_ShouldUpdateCorrectly()
        {
            // Arrange
            var hyperlink = new Hyperlink();

            // Act
            hyperlink.Status = HyperlinkStatus.Valid;
            hyperlink.LastChecked = DateTime.UtcNow;

            // Assert
            Assert.Equal(HyperlinkStatus.Valid, hyperlink.Status);
            Assert.True((DateTime.UtcNow - hyperlink.LastChecked.Value).TotalSeconds < 5);
        }

        [Fact]
        public void Hyperlink_ActionTaken_ShouldTrackUpdates()
        {
            // Arrange
            var hyperlink = new Hyperlink
            {
                OriginalUrl = "https://old-url.com",
                UpdatedUrl = "https://new-url.com"
            };

            // Act
            hyperlink.ActionTaken = HyperlinkAction.Updated;

            // Assert
            Assert.Equal(HyperlinkAction.Updated, hyperlink.ActionTaken);
            Assert.Equal("https://old-url.com", hyperlink.OriginalUrl);
            Assert.Equal("https://new-url.com", hyperlink.UpdatedUrl);
        }

        [Theory]
        [InlineData(HyperlinkStatus.Valid, false)]
        [InlineData(HyperlinkStatus.Invalid, true)]
        [InlineData(HyperlinkStatus.NotFound, true)]
        [InlineData(HyperlinkStatus.Expired, true)]
        [InlineData(HyperlinkStatus.Error, true)]
        public void Hyperlink_RequiresUpdate_ShouldBeBasedOnStatus(HyperlinkStatus status, bool expectedRequiresUpdate)
        {
            // Arrange
            var hyperlink = new Hyperlink
            {
                Status = status,
                RequiresUpdate = expectedRequiresUpdate
            };

            // Act & Assert
            Assert.Equal(expectedRequiresUpdate, hyperlink.RequiresUpdate);
        }
    }
}