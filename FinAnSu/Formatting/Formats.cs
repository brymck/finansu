/* 
 * FinAnSu
 * http://code.google.com/p/finansu/
 * 
 * Copyright 2011, Bryan McKelvey
 * Licensed under the MIT license
 * http://www.opensource.org/licenses/mit-license.php
 */

using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;

namespace FinAnSu
{
    public static partial class Formats
    {
        private const string DOLLARS_FORMAT = "_({0}* #,##0_);_({0}* (#,##0);_({0}* \" - \"_);_(@_)";
        private const string CENTS_FORMAT = "_({0}* #,##0.00_);_({0}* (#,##0.00);_({0}* \" - \"??_);_(@_)";
        private const string RATES_FORMAT = "_({0}* #,##0.0000_);_({0}* (#,##0.0000);_({0}* \" - \"????_);_(@_)";

        #region Helpers
        public static void CallFormatByName(string name)
        {
            MethodInfo info = MethodBase.GetCurrentMethod().DeclaringType.GetMethod(name);
            info.Invoke(null, null);
        }

        private static Excel.Range GetSelection()
        {
            Excel.Application app = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Range selection = (Excel.Range)app.Selection;
            return selection;
        }

        private static void SetNumberFormat(string format)
        {
            GetSelection().NumberFormat = format;
        }
        #endregion

        #region Currencies
        [ExcelCommand(Description = "Format cell in whole numbers.")]
        public static void WholeFormat()
        {
            GetSelection().NumberFormat = string.Format(DOLLARS_FORMAT, "");
        }

        [ExcelCommand(Description = "Format cell to two decimal places.")]
        public static void TwoDecimalsFormat()
        {
            GetSelection().NumberFormat = string.Format(CENTS_FORMAT, "");
        }

        [ExcelCommand(Description = "Format cell to four decimal places.")]
        public static void FourDecimalsFormat()
        {
            GetSelection().NumberFormat = string.Format(RATES_FORMAT, "");
        }

        [ExcelCommand(Description = "Format cell in dollars.")]
        public static void DollarsFormat()
        {
            GetSelection().NumberFormat = string.Format(DOLLARS_FORMAT, "$");
        }

        [ExcelCommand(Description = "Format cell in dollars and cents.")]
        public static void DollarsCentsFormat()
        {
            GetSelection().NumberFormat = string.Format(CENTS_FORMAT, "$");
        }

        [ExcelCommand(Description = "Format cell in euros.")]
        public static void EurosFormat()
        {
            GetSelection().NumberFormat = string.Format(DOLLARS_FORMAT, "€");
        }

        [ExcelCommand(Description = "Format cell in euros and cents.")]
        public static void EurosCentsFormat()
        {
            GetSelection().NumberFormat = string.Format(CENTS_FORMAT, "€");
        }

        [ExcelCommand(Description = "Format cell in yen.")]
        public static void YenFormat()
        {
            GetSelection().NumberFormat = string.Format(DOLLARS_FORMAT, "¥");
        }

        [ExcelCommand(Description = "Format cell in yen and sen.")]
        public static void YenSenFormat()
        {
            GetSelection().NumberFormat = string.Format(CENTS_FORMAT, "¥");
        }

        [ExcelCommand(Description = "Format cell in pounds.")]
        public static void PoundsFormat()
        {
            GetSelection().NumberFormat = string.Format(DOLLARS_FORMAT, "£");
        }

        [ExcelCommand(Description = "Format cell in pounds and pence.")]
        public static void PoundsPenceFormat()
        {
            GetSelection().NumberFormat = string.Format(CENTS_FORMAT, "£");
        }
        #endregion

        #region Other formats
        public static void AccountingUnderline()
        {
            GetSelection().Font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleSingleAccounting;
        }

        public static void CenterAcross()
        {
            GetSelection().HorizontalAlignment = Excel.XlHAlign.xlHAlignCenterAcrossSelection;
        }
        #endregion
    }
}
