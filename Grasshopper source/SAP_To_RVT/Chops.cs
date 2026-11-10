using System.Collections.Generic;
using System.Linq;


public static class ChunkExtensions
{
    public static List<List<T>> Chunk<T>(this IEnumerable<T> data, int chopSize)
    {
        return data.Select((x, i) => new { Index = i, Value = x })
                   .GroupBy(x => x.Index / chopSize)
                   .Select(x => x.Select(v => v.Value).ToList())
                   .ToList();
    }
}
