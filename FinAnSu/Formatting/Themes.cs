/* 
 * FinAnSu
 * http://code.google.com/p/finansu/
 * 
 * Copyright 2011, Bryan McKelvey
 * Licensed under the MIT license
 * http://www.opensource.org/licenses/mit-license.php
 */

using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;

namespace FinAnSu
{
    public static partial class Themes
    {
        #region ThemeSettings class
        internal sealed class ThemeSettings : ApplicationSettingsBase
        {
            private static ThemeSettings defaultInstance = ((ThemeSettings)(ApplicationSettingsBase.Synchronized(new ThemeSettings())));

            public static ThemeSettings Default
            {
                get { return defaultInstance; }
            }

            [UserScopedSetting()]
            public Theme.CellFormat Header
            {
                get { return ((Theme.CellFormat)(this["Header"])); }
                set { this["Header"] = value; }
            }

            [UserScopedSetting()]
            public Theme.CellFormat Normal
            {
                get { return ((Theme.CellFormat)(this["Normal"])); }
                set { this["Normal"] = value; }
            }

            [UserScopedSetting()]
            public Theme.CellFormat ColumnHeader
            {
                get { return ((Theme.CellFormat)(this["ColumnHeader"])); }
                set { this["ColumnHeader"] = value; }
            }
        }
        #endregion

        #region Internal Theme class
        internal class Theme
        {
            #region Internal CellFormat class
            internal class CellFormat
            {
                private bool bold;
                private bool italic;
                private int underline;
                private string font;

                public CellFormat(string font,
                                  bool bold, bool italic, int underline)
                {
                    this.font = font;
                    this.bold = bold;
                    this.italic = italic;
                    this.underline = underline;
                }

                #region CellFormat properties
                public bool Bold
                {
                    get { return bold; }
                    set { bold = value; }
                }

                public bool Italic
                {
                    get { return italic; }
                    set { italic = value; }
                }

                public int Underline
                {
                    get { return underline; }
                    set { underline = value; }
                }

                public string Font
                {
                    get { return font; }
                    set { font = value; }
                }
                #endregion

                public void Retrieve()
                {
                    Excel.Application app = (Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;
                    Excel.Range cell = (Excel.Range)app.ActiveCell;
                    this.font = cell.Font.Name;
                    this.underline = cell.Font.Underline;
                    this.bold = cell.Font.Bold;
                    this.italic = cell.Font.Italic;
                }

                public void Apply()
                {
                    MessageBox.Show("Applying!");
                    Excel.Application app = (Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;
                    Excel.Range range = (Excel.Range)app.Selection;
                    range.Font.Name = font;
                    range.Font.Bold = bold;
                    range.Font.Italic = italic;
                    range.Font.Underline = underline;
                }
            }
            #endregion

            private CellFormat header;
            private CellFormat normal;
            private CellFormat columnHeader;

            public Theme()
            {
                MessageBox.Show("Loading theme");
                header = settings.Header;
                MessageBox.Show("Loading theme");
                normal = settings.Normal;
                columnHeader = settings.Header;
                MessageBox.Show("Loaded theme");
            }

            public void Save()
            {
                settings.Save();
            }

            private CellFormat GetFormatByName(string name)
            {
                switch (name.ToLower())
                {
                    case "header":
                        return header;
                    case "normal":
                        return normal;
                    case "columnheader":
                        return columnHeader;
                    default:
                        return null;
                }
            }

            public void Retrieve(string styleName)
            {
                GetFormatByName(styleName).Retrieve();
                Save();
            }

            public void Apply(string styleName)
            {
                CellFormat toApply = GetFormatByName(styleName);
                MessageBox.Show(toApply.Font);
                toApply.Apply();
                MessageBox.Show("Done applying!");
            }

            #region Theme properties
            public CellFormat Header
            {
                get { return header; }
                set { header = value; }
            }

            public CellFormat Normal
            {
                get { return normal; }
                set { normal = value; }
            }

            public CellFormat ColumnHeader
            {
                get { return columnHeader; }
                set { columnHeader = value; }
            }
            #endregion
        }
        #endregion

        private static Excel.Application app = (Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;
        private static ThemeSettings settings = new ThemeSettings();
        private static Theme theme = new Theme();

        public static void Retrieve(string styleName)
        {
            theme.Retrieve(styleName);
        }

        public static void Apply(string styleName)
        {
            theme.Apply(styleName);
            MessageBox.Show("hi");
        }
    }
}
