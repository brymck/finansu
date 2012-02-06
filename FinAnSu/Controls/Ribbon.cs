/* 
 * FinAnSu
 * http://code.google.com/p/finansu/
 * 
 * Copyright 2011, Bryan McKelvey
 * Licensed under the MIT license
 * http://www.opensource.org/licenses/mit-license.php
 */

using System.Diagnostics;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Xml;
using ExcelDna.Integration.CustomUI;
using FinAnSu;
using Excel = Microsoft.Office.Interop.Excel;

// Needs to be COM-visible to work
[ComVisible(true)]
public class MyRibbon : ExcelDna.Integration.CustomUI.ExcelRibbon
{
    private const string HELP_URI = "http://brymck.github.com/finansu/";

    public override string GetCustomUI(string uiName)
    {
        XmlDocument doc = new XmlDocument();
        doc.Load(Assembly.GetExecutingAssembly().GetManifestResourceStream("FinAnSu.Controls.Ribbon.xml"));

        // For some reason doc.getElementById() isn't working
        XmlNodeList buttons = doc.GetElementsByTagName("button");
        foreach (XmlNode button in buttons)
        {
            if (button.Attributes["id"].Value == "UpdateButton")
            {
                if (Main.UpdateAvailable())
                {
                    button.Attributes["visible"].Value = "true";
                }
                break;
            }
        }
        return doc.InnerXml;
    }

    public void HelpButton_Click(IRibbonControl control)
    {
        Process.Start(HELP_URI);
    }

    public void AboutButton_Click(IRibbonControl control)
    {
        new FinAnSu.AboutForm().Show();
    }

    public void UpdateButton_Click(IRibbonControl control)
    {
        new FinAnSu.Controls.DownloadForm().Show();
    }

    public void FunctionsClick(IRibbonControl control)
    {
        Excel.Application app = (Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;
        Excel.Range activeCell = app.ActiveCell;
        string origFormula = app.ActiveCell.Formula;

        // Add formula to Excel
        if (origFormula == "")
        {
            // Add equals sign if formula is blank
            activeCell.Formula = "=" + control.Tag + "()";
        }
        else
        {
            // Add plus sign if no valid operator exists at end of formula
            switch (origFormula[origFormula.Length - 1])
            {
                case ',':
                case '+':
                case '-':
                case '*':
                case '/':
                    activeCell.Formula += control.Tag + "()";
                    break;
                default:
                    activeCell.Formula += "+" + control.Tag + "()";
                    break;
            }
        }

        // Replace original formula if user escapes function dialog wizard
        if (!app.Dialogs[Excel.XlBuiltInDialog.xlDialogFunctionWizard].Show())
        {
            activeCell.Formula = origFormula;
        }
    }

    public void ThemesGetClick(IRibbonControl control)
    {
        Themes.Apply(control.Tag);
    }

    public void ThemesSetClick(IRibbonControl control)
    {
        Themes.Retrieve(control.Tag);
    }

    public void FormatsClick(IRibbonControl control)
    {
        Formats.CallFormatByName(control.Tag);
    }

    /*
    public void ThemesNewButton_Click()
    {
    }
    */

    /// <summary>
    /// Returns a custom icon to a Ribbon control. Requires that Ribbon control to use this
    /// as the callback in their getImage attribute and to specify the name of the resource
    /// in their tag attribute.
    /// </summary>
    /// <param name="control">The caller Ribbon control.</param>
    /// <returns>The specified bitmap from the project resources.</returns>
    public Bitmap GetImage(IRibbonControl control)
    {
        return (Bitmap)FinAnSu.Properties.Resources.ResourceManager.GetObject(control.Tag);
    }
}