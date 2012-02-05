using NUnit.Framework;

namespace FinAnSu.Test.Main
{
    [TestFixture]
    public class MainTest
    {
        [Test]
        public void RetrievesVersion()
        {
            Assert.AreNotEqual(FinAnSu.Main.LatestVersion(), "");
        }
    }
}