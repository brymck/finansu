/* 
 * FinAnSu
 * http://code.google.com/p/finansu/
 * 
 * Copyright 2011, Bryan McKelvey
 * Licensed under the MIT license
 * http://www.opensource.org/licenses/mit-license.php
 */

using System;
using ExcelDna.Integration;

namespace FinAnSu
{
    public static class Calendar
    {
        [ExcelFunction("Returns whether a given day is a US Federal Reserve holiday.")]
        public static bool IsHoliday(
            [ExcelArgument("is the date you wish to check.", Name = "date")] double lngDate,
            [ExcelArgument("is whether you wish to allow weekends to be holidays.",
                Name = "include_weekends")] bool includeWeekends)
        {
            DateTime date = new DateTime((long)lngDate);
            int day = date.Day;
            int dayOfWeek = (int)date.DayOfWeek;

            if (!includeWeekends)
            {
                switch (dayOfWeek)
                {
                    case 0:
                    case 6:
                        return false;
                    default:
                        break;
                }
            }

            switch (date.Month)
            {
                case 1:
                    // New Year's Day
                    if (day == 1 || day == 2) { return IsObservedHoliday(day, dayOfWeek, 1); }
                    // Birthday of Martin Luther King, Jr.
                    if (dayOfWeek == 1) { return (day >= 22 && day <= 28); }
                    else { return false; }
                case 2:
                    // Washington's Birthday
                    if (dayOfWeek == 1) { return (day >= 15 && day <= 21); }
                    return false;
                case 5:
                    // Memorial Day
                    if (dayOfWeek == 1) { return (day >= 25 && day <= 31); }
                    else { return false; }
                case 7:
                    // Independence Day
                    if (day >= 3 && day <= 5) { return IsObservedHoliday(day, dayOfWeek, 4); }
                    else { return false; }
                case 9:
                    // Labor Day
                    if (dayOfWeek == 1) { return (day >= 1 && day <= 7); }
                    else { return false; }
                case 10:
                    // Columbus Day
                    if (dayOfWeek == 1) { return (day >= 8 && day <= 14); }
                    else { return false; }
                case 11:
                    // Veterans Day
                    if (day >= 10 && day <= 12) { return IsObservedHoliday(day, dayOfWeek, 11); }
                    // Thanksgiving
                    else if (dayOfWeek == 4) { return (day >= 22 && day <= 28); }
                    else { return false; }
                case 12:
                    // Christmas Day
                    if (day >= 24 && day <= 26) { return IsObservedHoliday(day, dayOfWeek, 25); }
                    // New Year's Day
                    if (day == 31 && dayOfWeek == 5) { return true; }
                    else { return false; }
                default:
                    return true;
            }
        }

        private static bool IsObservedHoliday(int day, int dayOfWeek, int holiday)
        {
            // Determine days after (before) holiday
            switch (day - holiday)
            {
                case -1: // If 1 day before holiday, test if Friday
                    return (dayOfWeek == 5);
                case 0: // If on holiday, test if weekday
                    return (dayOfWeek != 0 && dayOfWeek != 6);
                case 1: // If 1 day after holiday, test if Monday
                    return (dayOfWeek == 1);
                default:
                    return false;
            }
        }
    }
}
