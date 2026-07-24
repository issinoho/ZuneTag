namespace DrunkenBakery.ZuneTag.Tests
{
    using DrunkenBakery.ZuneTag;

    using Xunit;

    public class HexEncodingTests
    {
        [Theory]
        [InlineData("1A2B", 2)]
        [InlineData("1A2", 1)]
        [InlineData("1A-2B", 2)]
        [InlineData("", 0)]
        public void GetByteCount_ReturnsExpectedCount(string hexString, int expected)
        {
            Assert.Equal(expected, HexEncoding.GetByteCount(hexString));
        }

        [Fact]
        public void GetBytes_ParsesContiguousHexPairs()
        {
            var bytes = HexEncoding.GetBytes("1A2B", out var discarded);

            Assert.Equal(new byte[] { 0x1A, 0x2B }, bytes);
            Assert.Equal(0, discarded);
        }

        [Fact]
        public void GetBytes_SkipsNonHexCharactersAndCountsThemAsDiscarded()
        {
            var bytes = HexEncoding.GetBytes("1A-2B", out var discarded);

            Assert.Equal(new byte[] { 0x1A, 0x2B }, bytes);
            Assert.Equal(1, discarded);
        }

        [Fact]
        public void GetBytes_DropsTrailingCharacterOnOddLength()
        {
            var bytes = HexEncoding.GetBytes("1A2", out var discarded);

            Assert.Equal(new byte[] { 0x1A }, bytes);
            Assert.Equal(1, discarded);
        }

        [Fact]
        public void ToString_RoundTripsGetBytesOutput()
        {
            var bytes = HexEncoding.GetBytes("1A2B", out _);

            Assert.Equal("1A2B", HexEncoding.ToString(bytes));
        }

        [Theory]
        [InlineData("1A2B", true)]
        [InlineData("1A2G", false)]
        [InlineData("", true)]
        public void InHexFormat_ValidatesEveryCharacter(string hexString, bool expected)
        {
            Assert.Equal(expected, HexEncoding.InHexFormat(hexString));
        }

        [Theory]
        [InlineData('0', true)]
        [InlineData('9', true)]
        [InlineData('A', true)]
        [InlineData('F', true)]
        [InlineData('a', true)]
        [InlineData('f', true)]
        [InlineData('G', false)]
        [InlineData(':', false)]
        public void IsHexDigit_ClassifiesCharacter(char c, bool expected)
        {
            Assert.Equal(expected, HexEncoding.IsHexDigit(c));
        }
    }
}
