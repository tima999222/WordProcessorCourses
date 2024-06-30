using System;
using System.Collections.Generic;
using System.Linq;

public static class EnumerableExtensions
{
    public static T SecondOrDefault<T>(this IEnumerable<T> source)
    {
        if (source == null)
        {
            throw new ArgumentNullException(nameof(source));
        }

        using (var iterator = source.GetEnumerator())
        {
            if (iterator.MoveNext() && iterator.MoveNext())
            {
                return iterator.Current;
            }
        }

        return default(T);
    }
}