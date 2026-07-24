namespace DrunkenBakery.ZuneTag.Tests
{
    using DrunkenBakery.ZuneTag;

    using WMFSDKWrapper;

    using Xunit;

    public class AttributeTests
    {
        [Fact]
        public void Constructor_SetsAllProperties()
        {
            var attribute = new Attribute(3, "WM/Genre", "Action", WMT_ATTR_DATATYPE.WMT_TYPE_STRING);

            Assert.Equal((ushort)3, attribute.Index);
            Assert.Equal("WM/Genre", attribute.Name);
            Assert.Equal("Action", attribute.Value);
            Assert.Equal(WMT_ATTR_DATATYPE.WMT_TYPE_STRING, attribute.Type);
        }

        [Fact]
        public void ToString_FormatsNameValueAndType()
        {
            var attribute = new Attribute(0, "Title", "Foo", WMT_ATTR_DATATYPE.WMT_TYPE_STRING);

            Assert.Equal("Title = Foo (WMT_TYPE_STRING)", attribute.ToString());
        }
    }
}
