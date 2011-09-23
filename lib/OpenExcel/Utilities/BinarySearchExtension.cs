﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OpenExcel.Utilities
{
    /// <summary>
    /// Implements IList.BinarySearch() extension method.
    /// </summary>
	public static class BinarySearchExtension
	{
        public static int BinarySearch<T>(this IList<T> list, int index, int length, T value, IComparer<T> comparer)
        {
            if (list == null)
                throw new ArgumentNullException("list");
            else if (index < 0 || length < 0)
                throw new ArgumentOutOfRangeException((index < 0) ? "index" : "length");
            else if (list.Count - index < length)
                throw new ArgumentException();

            int lower = index;
            int upper = (index + length) - 1;

            while (lower <= upper)
            {
                int adjustedIndex = lower + ((upper - lower) >> 1);
                int comparison = comparer.Compare(list[adjustedIndex], value);
                if (comparison == 0)
                    return adjustedIndex;
                else if (comparison < 0)
                    lower = adjustedIndex + 1;
                else
                    upper = adjustedIndex - 1;
            }

            return ~lower;
        }

        public static int BinarySearch<T>(this IList<T> list, T value, IComparer<T> comparer)
        {
            return list.BinarySearch(0, list.Count, value, comparer);
        }

        public static int BinarySearch<T>(this IList<T> list, T value) where T : IComparable<T>
        {
            return list.BinarySearch(value, Comparer<T>.Default);
        }
    }
}
