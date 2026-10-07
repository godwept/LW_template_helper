using System.Text.Json;

namespace LWTemplateHelper;

internal sealed class TemplateConfigStore
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNameCaseInsensitive = true
    };

    public string ConfigDirectory { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "LWTemplateHelper",
        "TemplateTypes");

    public TemplateConfigStore()
    {
        Directory.CreateDirectory(ConfigDirectory);
    }

    public List<TemplateTypeConfig> LoadAll()
    {
        Directory.CreateDirectory(ConfigDirectory);

        var results = new List<TemplateTypeConfig>();

        foreach (string path in Directory.GetFiles(ConfigDirectory, "*.json").OrderBy(x => x))
        {
            try
            {
                string json = File.ReadAllText(path);
                var config = JsonSerializer.Deserialize<TemplateTypeConfig>(json, JsonOptions);

                if (config is null || string.IsNullOrWhiteSpace(config.Name))
                    continue;

                config.Bookmarks ??= [];
                config.Subtypes ??= [];
                results.Add(config);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Could not read configuration file '{path}'. {ex.Message}", ex);
            }
        }

        return results
            .OrderBy(x => x.Name, StringComparer.CurrentCultureIgnoreCase)
            .ToList();
    }

    public void Save(TemplateTypeConfig config, string? previousName = null)
    {
        Directory.CreateDirectory(ConfigDirectory);

        string destination = GetPath(config.Name);
        string temporary = destination + ".tmp";

        string json = JsonSerializer.Serialize(config, JsonOptions);
        File.WriteAllText(temporary, json);
        File.Move(temporary, destination, overwrite: true);

        if (!string.IsNullOrWhiteSpace(previousName))
        {
            string previousPath = GetPath(previousName);
            if (!string.Equals(previousPath, destination, StringComparison.OrdinalIgnoreCase) && File.Exists(previousPath))
                File.Delete(previousPath);
        }
    }

    private string GetPath(string typeName)
    {
        char[] invalid = Path.GetInvalidFileNameChars();
        string safeName = new(typeName
            .Trim()
            .Select(ch => invalid.Contains(ch) ? '_' : ch)
            .ToArray());

        if (string.IsNullOrWhiteSpace(safeName))
            safeName = "TemplateType";

        return Path.Combine(ConfigDirectory, safeName + ".json");
    }
}
