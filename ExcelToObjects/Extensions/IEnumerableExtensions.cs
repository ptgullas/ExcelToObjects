using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelToObjects.Extensions {
    public static class IEnumerableExtensions {
        // returns the item and index as a tuple
        public static IEnumerable<(T item, int index)> WithIndex<T>(this IEnumerable<T> self) 
            => self?.Select((item, index) => (item, index)) ?? new List<(T, int)>();
    }
}
