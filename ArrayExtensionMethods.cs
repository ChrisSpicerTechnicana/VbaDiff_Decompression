using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VbaDiff.Decompression
{
    internal static class ArrayExtensionMethods
    {
        public static T[] SubArray<T>(this T[] data, int index, int length)
        {
            T[] result = new T[length];
            Array.Copy(data, index, result, 0, length);
            return result;
        }
    }
}
