using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MarkdownRenderer.Helpers;

public static class CollectionExtensions
{
    public static Stack<T> Clone<T>(this Stack<T> original) 
    {
        var arr = new T[original.Count];
        original.CopyTo(arr, 0);
        Array.Reverse(arr);
        return new Stack<T>(arr);
    }
}
