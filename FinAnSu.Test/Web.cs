using FinAnSu;
using NUnit.Framework;

namespace FinAnSu.Test.Web
{
    [TestFixture]
    public class WebTest
    {
        [Test]
        public void CanRetrieveQuote()
        {
            // Hopefully I can reduce the hideousness of this at some point
            Assert.AreNotEqual(double.Parse(FinAnSu.Web.Quote("WFC", "b", "p", false, 15, false)[0, 0].ToString()), 0);
            Assert.AreNotEqual(double.Parse(FinAnSu.Web.Quote("US0003M", "b", "p", false, 15, false)[0, 0].ToString()), 0);
        }

        [Test]
        public void CanParseQuoteExtensions()
        {
            Assert.AreEqual(FinAnSu.Web.Quote("WFC", "b", "p", false, 15, false), FinAnSu.Web.Quote("WFC:US", "b", "p", false, 15, false));
        }

        [Test]
        public void CanParseIndexExtensions()
        {
            Assert.AreEqual(FinAnSu.Web.Quote("US0003M", "b", "p", false, 15, false), FinAnSu.Web.Quote("US0003M:IND", "b", "p", false, 15, false));
        }

        [Test]
        public void CanRetrieveForeignQuote()
        {
            Assert.AreNotEqual(double.Parse(FinAnSu.Web.Quote("7267:JP", "b", "p", false, 15, false)[0, 0].ToString()), 0);
        }
    }
}
