using System.Text.RegularExpressions;

namespace LWTemplateHelper;

internal sealed class TemplateTypeConfig
{
    public string Name { get; set; } = string.Empty;
    public List<string> Bookmarks { get; set; } = [];
    public List<TemplateSubtypeConfig> Subtypes { get; set; } = [];

    public override string ToString() => Name;
}

internal sealed class TemplateSubtypeConfig
{
    public string Name { get; set; } = string.Empty;
    public string EnglishTemplatePath { get; set; } = string.Empty;
    public string FrenchTemplatePath { get; set; } = string.Empty;

    public override string ToString() => Name;
}

internal static partial class ConfigurationValidation
{
    [GeneratedRegex("^[A-Za-z][A-Za-z0-9_]{0,39}$")]
    private static partial Regex BookmarkNameRegex();

    public static bool TryValidateBookmarkName(string name, out string error)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            error = "Bookmark names cannot be blank.";
            return false;
        }

        if (!BookmarkNameRegex().IsMatch(name))
        {
            error = $"'{name}' is not a valid Word bookmark name. Use 1-40 characters, start with a letter, and use only letters, numbers, or underscores.";
            return false;
        }

        error = string.Empty;
        return true;
    }

    public static bool TryValidateTemplatePath(string path, string language, out string error)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            error = $"Select the {language} .docx template.";
            return false;
        }

        if (!string.Equals(Path.GetExtension(path), ".docx", StringComparison.OrdinalIgnoreCase))
        {
            error = $"The {language} template must be a .docx file.";
            return false;
        }

        if (!File.Exists(path))
        {
            error = $"The {language} template file does not exist.";
            return false;
        }

        error = string.Empty;
        return true;
    }
}
