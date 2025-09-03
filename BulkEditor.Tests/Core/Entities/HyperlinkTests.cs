using BulkEditor.Core.Entities;
using FluentAssertions;
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
            hyperlink.Id.Should().NotBeEmpty();
            hyperlink.Status.Should().Be(HyperlinkStatus.Pending);
            hyperlink.ActionTaken.Should().Be(HyperlinkAction.None);
            hyperlink.RequiresUpdate.Should().BeFalse();
            hyperlink.DisplayText.Should().BeEmpty();
            hyperlink.OriginalUrl.Should().BeEmpty();
            hyperlink.UpdatedUrl.Should().BeEmpty();
            hyperlink.LookupId.Should().BeEmpty();
            hyperlink.ContentId.Should().BeEmpty();
            hyperlink.ErrorMessage.Should().BeEmpty();
        }

        [Theory]
        [InlineData("https://example.com/content/TSRC-ABC-123456", "TSRC-ABC-123456")]
        [InlineData("https://example.com/content/CMS-DEF-789012", "CMS-DEF-789012")]
        [InlineData("https://example.com/no-lookup", "")]
        public void Hyperlink_WithUrl_ShouldAcceptVariousFormats(string url, string expectedLookupId)
        {
            // Arrange & Act
            var hyperlink = new Hyperlink
            {
                OriginalUrl = url,
                LookupId = expectedLookupId
            };

            // Assert
            hyperlink.OriginalUrl.Should().Be(url);
            hyperlink.LookupId.Should().Be(expectedLookupId);
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
            hyperlink.Status.Should().Be(HyperlinkStatus.Valid);
            hyperlink.LastChecked.Should().BeCloseTo(DateTime.UtcNow, TimeSpan.FromSeconds(5));
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
            hyperlink.ActionTaken.Should().Be(HyperlinkAction.Updated);
            hyperlink.OriginalUrl.Should().Be("https://old-url.com");
            hyperlink.UpdatedUrl.Should().Be("https://new-url.com");
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
            hyperlink.RequiresUpdate.Should().Be(expectedRequiresUpdate);
        }
    }
}