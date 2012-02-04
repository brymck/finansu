using NUnit.Framework;

namespace FinAnSu.Test.FinMath
{
    [TestFixture]
    public class PayoffsTest
    {
        [Test]
        public void Can_Calculate_Long_Asset_Payoff()
        {
            Assert.That(Options.Payoff(true, true, false, 100, 0, 100), Is.EqualTo(0));
            Assert.That(Options.LongAsset(100, 100), Is.EqualTo(0));
            Assert.That(Options.Payoff(true, true, false, 100, 0, 90), Is.EqualTo(-10));
            Assert.That(Options.LongAsset(100, 90), Is.EqualTo(-10));
            Assert.That(Options.Payoff(true, true, false, 100, 0, 110), Is.EqualTo(10));
            Assert.That(Options.LongAsset(100, 110), Is.EqualTo(10));
        }

        [Test]
        public void Can_Calculate_Short_Asset_Payoff()
        {
            Assert.That(Options.Payoff(true, false, false, 100, 0, 100), Is.EqualTo(0));
            Assert.That(Options.ShortAsset(100, 100), Is.EqualTo(0));
            Assert.That(Options.Payoff(true, false, false, 100, 0, 90), Is.EqualTo(10));
            Assert.That(Options.ShortAsset(100, 90), Is.EqualTo(10));
            Assert.That(Options.Payoff(true, false, false, 100, 0, 110), Is.EqualTo(-10));
            Assert.That(Options.ShortAsset(100, 110), Is.EqualTo(-10));
        }

        [Test]
        public void Can_Calculate_Long_Call_Payoff()
        {
            Assert.That(Options.Payoff(false, true, true, 100, 10, 80), Is.EqualTo(-10));
            Assert.That(Options.LongCall(100, 10, 80), Is.EqualTo(-10));
            Assert.That(Options.Payoff(false, true, true, 100, 10, 100), Is.EqualTo(-10));
            Assert.That(Options.LongCall(100, 10, 100), Is.EqualTo(-10));
            Assert.That(Options.Payoff(false, true, true, 100, 10, 110), Is.EqualTo(0));
            Assert.That(Options.LongCall(100, 10, 110), Is.EqualTo(0));
            Assert.That(Options.Payoff(false, true, true, 100, 10, 120), Is.EqualTo(10));
            Assert.That(Options.LongCall(100, 10, 120), Is.EqualTo(10));
        }

        [Test]
        public void Can_Calculate_Long_Put_Payoff()
        {
            Assert.That(Options.Payoff(false, true, false, 100, 10, 110), Is.EqualTo(-10));
            Assert.That(Options.LongPut(100, 10, 110), Is.EqualTo(-10));
            Assert.That(Options.Payoff(false, true, false, 100, 10, 100), Is.EqualTo(-10));
            Assert.That(Options.LongPut(100, 10, 100), Is.EqualTo(-10));
            Assert.That(Options.Payoff(false, true, false, 100, 10, 90), Is.EqualTo(0));
            Assert.That(Options.LongPut(100, 10, 90), Is.EqualTo(0));
            Assert.That(Options.Payoff(false, true, false, 100, 10, 80), Is.EqualTo(10));
            Assert.That(Options.LongPut(100, 10, 80), Is.EqualTo(10));
        }

        [Test]
        public void Can_Calculate_Short_Call_Payoff()
        {
            Assert.That(Options.Payoff(false, false, true, 100, 10, 80), Is.EqualTo(10));
            Assert.That(Options.ShortCall(100, 10, 80), Is.EqualTo(10));
            Assert.That(Options.Payoff(false, false, true, 100, 10, 90), Is.EqualTo(10));
            Assert.That(Options.ShortCall(100, 10, 90), Is.EqualTo(10));
            Assert.That(Options.Payoff(false, false, true, 100, 10, 100), Is.EqualTo(10));
            Assert.That(Options.ShortCall(100, 10, 110), Is.EqualTo(0));
            Assert.That(Options.Payoff(false, false, true, 100, 10, 120), Is.EqualTo(-10));
            Assert.That(Options.ShortCall(100, 10, 120), Is.EqualTo(-10));
            Assert.That(Options.Payoff(false, false, true, 100, 10, 130), Is.EqualTo(-20));
            Assert.That(Options.ShortCall(100, 10, 130), Is.EqualTo(-20));
        }

        [Test]
        public void Can_Calculate_Short_Put_Payoff()
        {
            Assert.That(Options.Payoff(false, false, false, 100, 10, 120), Is.EqualTo(10));
            Assert.That(Options.ShortPut(100, 10, 120), Is.EqualTo(10));
            Assert.That(Options.Payoff(false, false, false, 100, 10, 110), Is.EqualTo(10));
            Assert.That(Options.ShortPut(100, 10, 110), Is.EqualTo(10));
            Assert.That(Options.Payoff(false, false, false, 100, 10, 100), Is.EqualTo(10));
            Assert.That(Options.ShortPut(100, 10, 90), Is.EqualTo(0));
            Assert.That(Options.Payoff(false, false, false, 100, 10, 80), Is.EqualTo(-10));
            Assert.That(Options.ShortPut(100, 10, 70), Is.EqualTo(-20));
        }
    }
}
