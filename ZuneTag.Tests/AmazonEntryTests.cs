namespace DrunkenBakery.ZuneTag.Tests
{
    using DrunkenBakery.ZuneTag;

    using TMDbLib.Objects.Search;

    using Xunit;

    public class AmazonEntryTests
    {
        [Fact]
        public void Constructor_SetsDefaults()
        {
            var entry = new AmazonEntry();

            Assert.False(entry.ExtraData);
            Assert.Equal("unknown", entry.Genre);
            Assert.Equal("unknown", entry.Director);
            Assert.Null(entry.Url);
        }

        [Fact]
        public void ToString_ReturnsMovieTitle()
        {
            var entry = new AmazonEntry { Movie = new SearchMovie { Title = "The Matrix" } };

            Assert.Equal("The Matrix", entry.ToString());
        }
    }
}
