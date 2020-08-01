﻿using System.Collections.Generic;

namespace Arc.YTSubConverter.Util
{
    internal static class StringExtensions
    {
        public static List<string> Split(this string str, string separator, int? maxItems = null)
        {
            List<string> result = new List<string>();
            int start = 0;
            int end;
            while (start <= str.Length)
            {
                if (start < str.Length && (maxItems == null || result.Count < maxItems.Value - 1))
                {
                    end = str.IndexOf(separator, start);
                    if (end < 0)
                        end = str.Length;
                }
                else
                {
                    end = str.Length;
                }

                result.Add(str.Substring(start, end - start));
                start = end + separator.Length;
            }
            return result;
        }
    }
}
