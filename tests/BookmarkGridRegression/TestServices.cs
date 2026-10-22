namespace LWTemplateHelper;

// In-memory services keep the UI regression checks away from Word and user configuration.
internal sealed class TemplateConfigStore
{
    private TemplateTypeConfig? _saved;
    public int SaveCount { get; private set; }
    public string ConfigDirectory => "In-memory regression configuration";
    public List<TemplateTypeConfig> LoadAll() => _saved is null ? [] : [_saved];
    public void Save(TemplateTypeConfig config, string? previousName = null)
    {
        _saved = config;
        SaveCount++;
    }
}

internal sealed record FinalizeResult(string EnglishBackupPath, string FrenchBackupPath);

internal sealed class WordSession : IDisposable
{
    public bool IsOpen { get; private set; }
    public bool IsTestMode => false;
    public string SessionDirectory => "In-memory regression session";
    public HashSet<string> English { get; } = new(StringComparer.OrdinalIgnoreCase);
    public HashSet<string> French { get; } = new(StringComparer.OrdinalIgnoreCase);
    public void Open(string englishSource, string frenchSource) => IsOpen = true;
    public void Close() => IsOpen = false;
    public void Dispose() => Close();
    public (HashSet<string> English, HashSet<string> French) GetBookmarkNames() => (English, French);
    public void AddBookmark(string name)
    {
        English.Add(name);
        French.Add(name);
    }
    public string AddNextInfoBookmark()
    {
        int highest = English.Union(French).Select(name => TryGetInfoNumber(name, out int n) ? n : 0).Max();
        string name = $"INFO_{highest + 1}";
        AddBookmark(name);
        return name;
    }
    public static bool TryGetInfoNumber(string name, out int number)
    {
        number = 0;
        return name.StartsWith("INFO_", StringComparison.OrdinalIgnoreCase)
            && int.TryParse(name[5..], out number) && number >= 1;
    }
    public void LocateEnglish(string name) => throw new NotSupportedException();
    public void LocateFrench(string name) => throw new NotSupportedException();
    public void LocateBoth(string name) => throw new NotSupportedException();
    public void ArrangeSideBySide() => throw new NotSupportedException();
    public void FocusEnglish() => throw new NotSupportedException();
    public void FocusFrench() => throw new NotSupportedException();
    public void ReplaceBookmarkRange(string name) => throw new NotSupportedException();
    public void RenameBookmark(string oldName, string newName) => throw new NotSupportedException();
    public void DeleteBookmarkMarkers(string name) => throw new NotSupportedException();
    public void StartTestMode() => throw new NotSupportedException();
    public void ResetTestMode() => throw new NotSupportedException();
    public void CloseTestMode() => throw new NotSupportedException();
    public (bool English, bool French) TestSetValue(string name, string value) => throw new NotSupportedException();
    public (bool English, bool French) TestDeleteContent(string name) => throw new NotSupportedException();
    public FinalizeResult FinalizeToSources(string english, string french) => throw new NotSupportedException();
}
