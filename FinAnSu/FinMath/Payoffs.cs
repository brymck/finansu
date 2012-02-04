
using ExcelDna.Integration;

namespace FinAnSu
{
    public static partial class Options
    {
        [ExcelFunction("Calculates the payoff from holding an asset or option, long or short, put or call.")]
        public static double Payoff(
            [ExcelArgument("true if calculating the payoff for an asset position, false if calculating the payoff for an option", Name="is_asset")]
            bool isAsset, 
            [ExcelArgument("true if a long position is held in the asset/option, false if the position is short", Name="is_long")]
            bool isLong, 
            [ExcelArgument("true if the option is a call, false if a put (ignored for asset positions)", Name="is_call")]
            bool isCall, 
            [ExcelArgument("the strike price of the option, or purchase/short sale price for asset positions)", Name="strike_price")]
            double strikePrice,
            [ExcelArgument("the premium of the option (ignored for asset positions)", Name="premium")]
            double premium, 
            [ExcelArgument("the current asset price", Name="asset_price")]
            double assetPrice)
        {
            double payoff = isAsset
                         ? assetPrice - strikePrice
                         : (isCall
                                ? (assetPrice > strikePrice ? assetPrice - strikePrice - premium : -premium)
                                : (assetPrice < strikePrice ? strikePrice - assetPrice - premium : -premium));
            return isLong ? payoff : -payoff;
        }

        [ExcelFunction("Calculates the payoff from a long position in an asset")]
        public static double LongAsset(int strikePrice, int assetPrice)
        {
            return Payoff(true, true, false, strikePrice, 0, assetPrice);
        }

        [ExcelFunction("Calculates the payoff from a short position in an asset")]
        public static double ShortAsset(int strikePrice, int assetPrice)
        {
            return Payoff(true, false, false, strikePrice, 0, assetPrice);
        }

        [ExcelFunction("Calculates the payoff from a long position in a call option")]
        public static double LongCall(int strikePrice, int premium, int assetPrice)
        {
            return Payoff(false, true, true, strikePrice, premium, assetPrice);
        }

        [ExcelFunction("Calculates the payoff from a long position in a put option")]
        public static double LongPut(int strikePrice, int premium, int assetPrice)
        {
            return Payoff(false, true, false, strikePrice, premium, assetPrice);
        }

        [ExcelFunction("Calculates the payoff from a short position in a call option")]
        public static double ShortCall(int strikePrice, int premium, int assetPrice)
        {
            return Payoff(false, false, true, strikePrice, premium, assetPrice);
        }

        [ExcelFunction("Calculates the payoff from a short position in a put option")]
        public static double ShortPut(int strikePrice, int premium, int assetPrice)
        {
            return Payoff(false, false, false, strikePrice, premium, assetPrice);
        }
    }
}