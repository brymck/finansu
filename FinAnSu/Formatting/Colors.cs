/* 
 * FinAnSu
 * http://code.google.com/p/finansu/
 * 
 * Copyright 2011, Bryan McKelvey
 * Licensed under the MIT license
 * http://www.opensource.org/licenses/mit-license.php
 */

using System;
using System.Drawing;
using System.Globalization;
using ExcelDna.Integration;

namespace FinAnSu
{
    public static partial class Formats
    {
        private const string TWO_DIGITS_HEX = "x2";

        #region Conversion to Color
        private static Color RGBToColor(int red, int green, int blue)
        {
            return Color.FromArgb(red, green, blue);
        }

        private static int ParseHex(string hex)
        {
            return int.Parse(hex, NumberStyles.HexNumber) * (hex.Length == 1 ? 17 : 1);
        }

        private static Color HexToColor(string hex)
        {
            hex = hex.TrimStart('#');
            int colorLength = (hex.Length == 3 ? 1 : 2);

            int red = ParseHex(hex.Substring(0, colorLength));
            int green = ParseHex(hex.Substring(colorLength, colorLength));
            int blue = ParseHex(hex.Substring(2 * colorLength, colorLength));
            return Color.FromArgb(red, green, blue);
        }

        private static Color HSVToColor(double hue, double saturation, double value)
        {
            int hi = (int)(Math.Floor(hue / 60)) % 6;
            double f = hue / 60 - Math.Floor(hue / 60);

            value = value * 255;
            int v = (int)(value);
            int p = (int)(value * (1 - saturation));
            int q = (int)(value * (1 - f * saturation));
            int t = (int)(value * (1 - (1 - f) * saturation));

            switch (hi)
            {
                case 0:
                    return Color.FromArgb(v, t, p);
                case 1:
                    return Color.FromArgb(q, v, p);
                case 2:
                    return Color.FromArgb(p, v, t);
                case 3:
                    return Color.FromArgb(p, q, v);
                case 4:
                    return Color.FromArgb(t, p, v);
                default:
                    return Color.FromArgb(v, p, q);
            }
        }
        #endregion

        #region Conversion from Color
        private static string ColorToHex(Color color)
        {
            return "#" + color.R.ToString(TWO_DIGITS_HEX) + color.G.ToString(TWO_DIGITS_HEX) + color.B.ToString(TWO_DIGITS_HEX);
        }

        private static object[] ColorToRGB(Color color, string flag)
        {
            int red = color.R;
            int green = color.G;
            int blue = color.B;

            switch (flag.ToLower())
            {
                case "red":
                case "r":
                    return new object[] { red };
                case "green":
                case "g":
                    return new object[] { green };
                case "blue":
                case "b":
                    return new object[] { blue };
                default:
                    return new object[] { red, green, blue };
            }
        }

        private static object[] ColorToHSV(Color color, string flag)
        {
            int max = Math.Max(color.R, Math.Max(color.G, color.B));
            int min = Math.Min(color.R, Math.Min(color.G, color.B));

            double hue = color.GetHue();
            double saturation = (max == 0) ? 0 : 1d - (1d * min / max);
            double value = max / 255d;

            switch (flag.ToLower())
            {
                case "hue":
                case "h":
                    return new object[] { hue };
                case "saturation":
                case "sat":
                case "s":
                    return new object[] { saturation };
                case "value":
                case "val":
                case "v":
                case "brightness":
                case "b":
                    return new object[] { value };
                default:
                    return new object[] { hue, saturation, value };
            }
        }
        #endregion

        #region Excel color functions
        [ExcelFunction("Converts an RGB color to hexadecimal format.")]
        public static string RGBToHex(
            [ExcelArgument("is the level of red in the color (0‒255).")] int red,
            [ExcelArgument("is the level of green in the color (0‒255).")] int green,
            [ExcelArgument("is the level of blue in the color (0‒255).")] int blue)
        {
            return ColorToHex(RGBToColor(red, green, blue));
        }

        [ExcelFunction("Converts an RGB color to HSV format.")]
        public static object[] RGBToHSV(
            [ExcelArgument("is the level of red in the color (0‒255).")] int red,
            [ExcelArgument("is the level of green in the color (0‒255).")] int green,
            [ExcelArgument("is the level of blue in the color (0‒255).")] int blue,
            [ExcelArgument("is text describing which HSV value you are requesting. Defaults to a horizontal array containing all three.")] string flag)
        {
            return ColorToHSV(RGBToColor(red, green, blue), flag);
        }

        [ExcelFunction("Converts a hexadecimal color to RGB format.")]
        public static object[] HexToRGB(
            [ExcelArgument("is the color in three- or six-digit hexadecimal format.")] string hex,
            [ExcelArgument("is text describing which RGB value you are requesting. Defaults to a horizontal array containing all three.")] string flag)
        {
            return ColorToRGB(HexToColor(hex), flag);
        }

        [ExcelFunction("Converts a hexadecimal color to HSV format.")]
        public static object[] HexToHSV(
            [ExcelArgument("is the color in three- or six-digit hexadecimal format.")] string hex,
            [ExcelArgument("is text describing which HSV value you are requesting. Defaults to a horizontal array containing all three.")] string flag)
        {
            return ColorToHSV(HexToColor(hex), flag);
        }

        [ExcelFunction("Converts an HSV color to RGB format.")]
        public static object[] HSVToRGB(
            [ExcelArgument("is the level of hue in the color (0‒100%).")] double hue,
            [ExcelArgument("is the level of saturation in the color (0‒100%).")] double saturation,
            [ExcelArgument("is the level of value (or brightness) in the color (0‒360).")] double value,
            [ExcelArgument("is text describing which RGB value you are requesting. Defaults to a horizontal array containing all three.")] string flag)
        {
            return ColorToRGB(HSVToColor(hue, saturation, value), flag);
        }

        [ExcelFunction("Converts an HSV color to hexadecimal format.")]
        public static string HSVToHex(
            [ExcelArgument("is the level of hue in the color (0‒100%).")] double hue,
            [ExcelArgument("is the level of saturation in the color (0‒100%).")] double saturation,
            [ExcelArgument("is the level of value (or brightness) in the color (0‒360).")] double value)
        {
            return ColorToHex(HSVToColor(hue, saturation, value));
        }
        #endregion
    }
}
