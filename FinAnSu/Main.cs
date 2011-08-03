/* 
 * FinAnSu
 * http://code.google.com/p/finansu/
 * 
 * Copyright 2011, Bryan McKelvey
 * Licensed under the MIT license
 * http://www.opensource.org/licenses/mit-license.php
 */

using System;
using System.Reflection;
using System.Text.RegularExpressions;
using ExcelDna.Integration;

namespace FinAnSu
{
    public static class Main
    {
        public static Version installedVersion = Assembly.GetExecutingAssembly().GetName().Version;
        private const int LATEST_IS_INSTALLED = -1;
        private static Regex trimRevisionZeroes = new Regex("(?:\\.0){0,2}$");

        [ExcelFunction("Returns latest version number for FinAnSu.")]
        public static string LatestVersion()
        {
            return Web.GetWebData("http://code.google.com/p/finansu/",
                new string[] { "FinAnSu\\-([0-9]{1,2}\\.[0-9]{1,2}(?:\\.[0-9]{1,2})?)\\.zip" }, 1, null)[0, 0].ToString();
        }

        [ExcelFunction("Returns installed version number for FinAnSu.")]
        public static string InstalledVersion()
        {
            return trimRevisionZeroes.Replace(installedVersion.ToString(), "");
        }

        [ExcelFunction("Returns whether an update exists for FinAnSu.")]
        public static bool UpdateAvailable()
        {
            try
            {
                Version latest = new Version(LatestVersion());
                return (installedVersion.CompareTo(latest) == LATEST_IS_INSTALLED);
            }
            catch (Exception)
            {
                return true;
            }
        }
    }
}