/* 
 * FinAnSu
 * http://code.google.com/p/finansu/
 * 
 * Copyright 2011, Bryan McKelvey
 * Licensed under the MIT license
 * http://www.opensource.org/licenses/mit-license.php
 */

using System;
using System.Text.RegularExpressions;
using System.Linq;
using System.Collections;
using System.Collections.Generic;
using ExcelDna.Integration;

namespace FinAnSu
{
    public static partial class Functionality
    {
        #region Array sorting helper functions
        private static T[][] ToJagged<T>(this T[,] array, bool vertical)
        {
            int height = vertical ? array.GetLength(0) : array.GetLength(1);
            int width = vertical ? array.GetLength(1) : array.GetLength(0);
            T[][] jagged = new T[height][];

            for (int i = 0; i < height; i++)
            {
                T[] row = new T[width];
                for (int j = 0; j < width; j++)
                {
                    row[j] = array[vertical ? i : j, vertical ? j : i];
                }
                jagged[i] = row;
            }
            return jagged;
        }

        private static T[,] ToRectangular<T>(this T[][] array, bool vertical)
        {
            int height = vertical ? array.Length : array[0].Length;
            int width = vertical ? array[0].Length : array.Length;
            T[,] rect = new T[height, width];
            for (int i = 0; i < (vertical ? height : width); i++)
            {
                T[] row = array[i];
                for (int j = 0; j < (vertical ? width : height); j++)
                {
                    rect[vertical ? i : j, vertical ? j : i] = row[j];
                }
            }
            return rect;
        }

        private static void Sort<T>(T[][] data, int index, bool ascending)
        {
            Comparer<T> comparer = Comparer<T>.Default;
            Array.Sort<T[]>(data, (x, y) => comparer.Compare((ascending ? x[index] : y[index]), (ascending ? y[index] : x[index])));
        }
        #endregion

        [ExcelFunction("Automatically sorts an array in Excel when the range it refers to is updated.")]
        public static object[,] AutoSort(
            [ExcelArgument("is a list of cells to monitor for changes in value")] object[,] range,
            [ExcelArgument("is the index of the row of column you wish to sort on. Defaults to 1.")] int index,
            [ExcelArgument("is whether you wish to sort based on vertical data. Defaults to TRUE.",
                Name = "sort_vertical")] object sortVertical,
            [ExcelArgument("is whether you wish to sort in ascending or alphabetical order. Defaults to TRUE.",
                Name = "sort_ascending")] object sortAscending)
        {
            // Determines whether to sort vertically and ascending, defaulting to true if no argument is supplied
            bool vertical = (sortVertical is bool ? (bool)sortVertical : true);
            bool ascending = (sortAscending is bool ? (bool)sortAscending : true);

            // Convert object to jagged array, as rectangular arrays cannot be enumerated or queried
            object[][] temp = range.ToJagged<object>(vertical);

            // Sort jagged array, with index defaulting to first row or column
            Sort<object>(temp, index < 1 ? 0 : index - 1, ascending);

            // Convert back to rectangular array and return
            return temp.ToRectangular<object>(vertical);
        }
    }
}
