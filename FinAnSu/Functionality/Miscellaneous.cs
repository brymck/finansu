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
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;

namespace FinAnSu
{
    public static partial class Functionality
    {
        [ExcelFunction("Returns the portion of text between two matched search strings.")]
        public static string BetweenString(
            [ExcelArgument("is the text from which you want to extract a portion.")] string text,
            [ExcelArgument("is the beginning text you want to find.", Name = "begin_text")] string beginText,
            [ExcelArgument("is the ending text you want to find.", Name = "end_text")] string endText,
            [ExcelArgument("is the character number in Text, starting from the left, at which you want to start searching. " +
                "Defaults to 1.", Name = "start_num")] int startNum)
        {
            if (startNum > 0) startNum -= 1;
            int beginIndex = text.IndexOf(beginText, startNum) + beginText.Length;
            int endIndex = (endText == "") ? text.Length : text.IndexOf(endText, beginIndex);
            return text.Substring(beginIndex, endIndex - beginIndex);
        }

        [ExcelFunction("Overrides error values with a specified replacement value.")]
        public static object NoError(
            [ExcelArgument("is the value you wish to display if no errors are present.")] object value,
            [ExcelArgument("is the value with which you wish to replace errors. Defaults to 0.")] object replacement)
        {
            return ((bool)XlCall.Excel(XlCall.xlfIserror, value)) ? replacement : value;
        }

        [ExcelFunction("Returns a weighted average.")]
        public static object WeightedAverage(
            [ExcelArgument("is a range representing the weights you will apply to each value.")] object[] weights,
            [ExcelArgument("is a range representing the values to weight.")] object[] values)
        {
            if (weights.Length != values.Length)
            {
                return ExcelError.ExcelErrorValue;
            }
            else
            {
                double numerator = 0;
                double denominator = 0;
                for (int index = 0; index < weights.Length; index++)
                {
                    double weight = (double)weights[index];
                    double value = (double)values[index];
                    numerator += weight * value;
                    denominator += weight;
                }
                return numerator / denominator;
            }
        }

        /// <summary>
        /// Looks for a value in the leftmost column of a table, and then returns a value in the same row from a column you specify.
        /// </summary>
        /// <param name="lookupValue">The value to be found in the first column of the table, and can be a value, reference, or a text string.</param>
        /// <param name="tableArray">A table of text, numbers, or logical values, in which data is retrieved. Can be a reference to a range or a range name.</param>
        /// <param name="colIndexName">The header of the column in table_array from which the matching value should be returned. Defaults to the header of the first column.</param>
        /// <returns></returns>
        [ExcelFunction("Looks for a value in the leftmost column of a table, and then returns a value in the same row from a column you specify.")]
        public static object DLookup(
            [ExcelArgument("is the value to be found in the first column of the table, and can be a value, reference, or a text string.",
                Name = "lookup_value")] string lookupValue,
            [ExcelArgument("is a table of text, numbers, or logical values, in which data is retrieved. " +
                "Table_array can be a reference to a range or a range name.", Name = "table_array")] object[,] tableArray,
            [ExcelArgument("is the header of the column in table_array from which the matching value should be returned. " +
                "Defaults to the header of the first column.", Name = "col_index_name")] string colIndexName)
        {
            int height = tableArray.GetLength(0);
            int width = tableArray.GetLength(1);
            int rowIndex = -1;
            int colIndex = -1;

            for (int row = 0; row < height; row++)
            {
                if ((string)tableArray[row, 0] == lookupValue)
                {
                    rowIndex = row;
                    break;
                }
            }

            if (rowIndex == -1)
            {
                return "";
            }
            else
            {
                if (colIndexName == "")
                {
                    return tableArray[rowIndex, 0];
                }
                else
                {
                    for (int col = 0; col < width; col++)
                    {
                        if ((string)tableArray[0, col] == colIndexName)
                        {
                            colIndex = col;
                            return tableArray[rowIndex, colIndex];
                        }
                    }
                    return "";
                }
            }
        }

        /// <summary>
        /// Rolls dice with a designated number of sides the specified number of times.
        /// </summary>
        /// <param name="rollText">A text string containing a dice roll formula (such as "3d6", where a 6-sided die is rolled 3 times).</param>
        /// <param name="resultCount">The number of results to return in a vertical array.</param>
        /// <param name="dropCount">How many of the lowest rolls to drop.</param>
        /// <returns></returns>
        [ExcelFunction("Rolls dice with a designated number of sides the specified number of times.")]
        public static object[,] RollDice(
            [ExcelArgument("is a text string containing a dice roll formula (such as \"3d6\", where a 6-sided die is rolled 3 times).",
                Name = "roll_text")] string rollText,
            [ExcelArgument("is the number of results to return in a vertical array.", Name = "result_count")] int resultCount,
            [ExcelArgument("is how many of the lowest rolls to drop.", Name = "drop_count")] int dropCount)
        {
            int dIndex = rollText.IndexOf('d');
            if (dIndex == -1)
            {
                return new object[,] { { "" } };
            }
            else
            {
                if (resultCount == 0) resultCount = 1;
                object[,] results = new object[resultCount, 1];
                int signIndex = Math.Max(rollText.IndexOf('+'), rollText.IndexOf('-'));
                int diceRolls = (dIndex == 0 ? 1 : int.Parse(rollText.Substring(0, dIndex)));
                int diceSides = (signIndex == -1 ? int.Parse(rollText.Substring(dIndex + 1)) :
                                                   int.Parse(rollText.Substring(dIndex + 1, signIndex - dIndex - 1)));
                int modifier = (signIndex == -1 ? 0 : int.Parse(rollText.Substring(signIndex)));
                Random rand = new Random();

                for (int i = 0; i < resultCount; i++)
                {
                    List<int> temp = new List<int>();
                    for (int j = 0; j < diceRolls; j++)
                    {
                        temp.Add(rand.Next(diceSides) + 1);
                    }
                    temp.Sort();
                    temp.RemoveRange(0, dropCount);
                    results[i, 0] = temp.Sum() + modifier;
                }
                return results;
            }
        }
    }
}
