using System.Text.Json;
using LWTemplateHelper;

internal static class Program
{
    private static int Main()
    {
        int failures = 0;
        foreach (var (name, test) in new (string, Action)[]
        {
            ("Resolve legacy and explicit selections in master order", ResolveSelections),
            ("JSON preserves missing, null, empty, and explicit selections", RoundTripSelections),
            ("Saving removes deleted master names without selecting additions", SavePrunesSelections)
        })
        {
            try { test(); Console.WriteLine($"PASS: {name}"); }
            catch (Exception ex) { failures++; Console.WriteLine($"FAIL: {name}: {ex.GetBaseException().Message}"); }
        }
        return failures == 0 ? 0 : 1;
    }

    private static void ResolveSelections()
    {
        var type = new TemplateTypeConfig { Bookmarks = ["A", "B", "C"] };
        var legacy = new TemplateSubtypeConfig();
        Equal(type.GetBookmarksForSubtype(legacy), "A", "B", "C");
        Assert(legacy.Bookmarks is null, "Resolving legacy selection must not materialize it.");

        var empty = new TemplateSubtypeConfig { Bookmarks = [] };
        Equal(type.GetBookmarksForSubtype(empty));
        var subset = new TemplateSubtypeConfig { Bookmarks = ["c", "a", "STALE"] };
        Equal(type.GetBookmarksForSubtype(subset), "A", "C");
        Equal(subset.Bookmarks, "c", "a", "STALE");
        Equal(type.Bookmarks, "A", "B", "C");

        type.Bookmarks.Add("D");
        Equal(type.GetBookmarksForSubtype(subset), "A", "C");
        Equal(type.GetBookmarksForSubtype(legacy), "A", "B", "C", "D");
        type.Bookmarks.Clear();
        Equal(type.GetBookmarksForSubtype(legacy));
        Equal(type.GetBookmarksForSubtype(subset));
    }

    private static void RoundTripSelections()
    {
        string directory = Path.Combine(AppContext.BaseDirectory, "fixtures", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        File.WriteAllText(Path.Combine(directory, "Legacy.json"), """
            { "Name": "Legacy", "Bookmarks": ["A", "B"], "Subtypes": [
              { "Name": "Missing" }, { "Name": "Null", "Bookmarks": null },
              { "Name": "Empty", "Bookmarks": [] }, { "Name": "Subset", "Bookmarks": ["A"] }
            ] }
            """);
        var store = new TemplateConfigStore(directory);
        var type = store.LoadAll().Single();
        Assert(type.Subtypes[0].Bookmarks is null && type.Subtypes[1].Bookmarks is null, "Legacy selections changed.");
        Equal(type.Subtypes[2].Bookmarks!);
        Equal(type.Subtypes[3].Bookmarks!, "A");
        store.Save(type);
        var reloaded = new TemplateConfigStore(directory).LoadAll().Single();
        Assert(reloaded.Subtypes[0].Bookmarks is null && reloaded.Subtypes[1].Bookmarks is null, "Legacy selection was materialized on save.");
        Equal(reloaded.GetBookmarksForSubtype(reloaded.Subtypes[0]), "A", "B");
        Equal(reloaded.Subtypes[2].Bookmarks!);
        Equal(reloaded.Subtypes[3].Bookmarks!, "A");
        using var json = JsonDocument.Parse(File.ReadAllText(Path.Combine(directory, "Legacy.json")));
        Assert(json.RootElement.GetProperty("Subtypes")[2].GetProperty("Bookmarks").GetArrayLength() == 0, "Empty selection must serialize as [].");
    }

    private static void SavePrunesSelections()
    {
        string directory = Path.Combine(AppContext.BaseDirectory, "fixtures", Guid.NewGuid().ToString("N"));
        var store = new TemplateConfigStore(directory);
        var type = new TemplateTypeConfig
        {
            Name = "Pruning", Bookmarks = ["A", "B"],
            Subtypes = [new() { Bookmarks = ["a", "B"] }, new() { Bookmarks = [] }, new()]
        };
        type.Bookmarks = ["A", "C"];
        store.Save(type);
        var reloaded = store.LoadAll().Single();
        Equal(reloaded.Subtypes[0].Bookmarks!, "A");
        Equal(reloaded.Subtypes[1].Bookmarks!);
        Assert(reloaded.Subtypes[2].Bookmarks is null, "Saving the master list must preserve legacy state.");
        Equal(reloaded.GetBookmarksForSubtype(reloaded.Subtypes[2]), "A", "C");
        type.Bookmarks = ["RENAMED_A", "C"];
        store.Save(type);
        Equal(store.LoadAll().Single().Subtypes[0].Bookmarks!);
    }

    private static void Equal(IEnumerable<string> actual, params string[] expected) =>
        Assert(actual.SequenceEqual(expected), $"Expected [{string.Join(", ", expected)}], got [{string.Join(", ", actual)}].");

    private static void Assert(bool condition, string message)
    {
        if (!condition) throw new InvalidOperationException(message);
    }
}
