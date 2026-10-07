namespace LWTemplateHelper;

internal sealed class TemplateTypeDialog : Form
{
    private readonly TextBox _nameBox = new() { Dock = DockStyle.Fill };
    private readonly TextBox _bookmarksBox = new()
    {
        Dock = DockStyle.Fill,
        Multiline = true,
        ScrollBars = ScrollBars.Vertical,
        AcceptsReturn = true
    };

    public string TemplateTypeName { get; private set; } = string.Empty;
    public List<string> BookmarkNames { get; private set; } = [];

    public TemplateTypeDialog(TemplateTypeConfig? existing = null)
    {
        Text = existing is null ? "New Template Type" : "Edit Template Type";
        StartPosition = FormStartPosition.CenterParent;
        MinimizeBox = false;
        MaximizeBox = false;
        Size = new Size(560, 470);
        MinimumSize = new Size(500, 400);

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(12),
            ColumnCount = 1,
            RowCount = 5
        };

        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        root.Controls.Add(new Label { Text = "Template Type name", AutoSize = true });
        root.Controls.Add(_nameBox);
        root.Controls.Add(new Label
        {
            Text = "Bookmark names — one per line",
            AutoSize = true,
            Margin = new Padding(3, 12, 3, 3)
        });
        root.Controls.Add(_bookmarksBox);

        var buttons = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.RightToLeft,
            AutoSize = true
        };

        var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, AutoSize = true };
        var save = new Button { Text = "Save", AutoSize = true };
        save.Click += (_, _) => SaveAndClose();

        buttons.Controls.Add(cancel);
        buttons.Controls.Add(save);
        root.Controls.Add(buttons);

        Controls.Add(root);
        AcceptButton = save;
        CancelButton = cancel;

        if (existing is not null)
        {
            _nameBox.Text = existing.Name;
            _bookmarksBox.Lines = existing.Bookmarks.ToArray();
        }
    }

    private void SaveAndClose()
    {
        string name = _nameBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(name))
        {
            MessageBox.Show(this, "Enter a Template Type name.", "Template Type",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var bookmarks = _bookmarksBox.Lines
            .Select(x => x.Trim())
            .Where(x => x.Length > 0)
            .ToList();

        var duplicate = bookmarks
            .GroupBy(x => x, StringComparer.OrdinalIgnoreCase)
            .FirstOrDefault(g => g.Count() > 1);

        if (duplicate is not null)
        {
            MessageBox.Show(this, $"Bookmark '{duplicate.Key}' is listed more than once.", "Bookmark names",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        foreach (string bookmark in bookmarks)
        {
            if (!ConfigurationValidation.TryValidateBookmarkName(bookmark, out string error))
            {
                MessageBox.Show(this, error, "Bookmark names", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
        }

        TemplateTypeName = name;
        BookmarkNames = bookmarks;
        DialogResult = DialogResult.OK;
        Close();
    }
}

internal sealed class TemplateSubtypeDialog : Form
{
    private readonly TextBox _nameBox = new() { Dock = DockStyle.Fill };
    private readonly TextBox _englishPath = new() { ReadOnly = true, Dock = DockStyle.Fill };
    private readonly TextBox _frenchPath = new() { ReadOnly = true, Dock = DockStyle.Fill };

    public string SubtypeName { get; private set; } = string.Empty;
    public string EnglishTemplatePath { get; private set; } = string.Empty;
    public string FrenchTemplatePath { get; private set; } = string.Empty;

    public TemplateSubtypeDialog(TemplateSubtypeConfig? existing = null)
    {
        Text = existing is null ? "New Template Subtype" : "Edit Template Subtype";
        StartPosition = FormStartPosition.CenterParent;
        MinimizeBox = false;
        MaximizeBox = false;
        Size = new Size(720, 260);
        MinimumSize = new Size(620, 250);

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(12),
            ColumnCount = 1,
            RowCount = 4,
            AutoSize = true
        };

        root.Controls.Add(BuildTextRow("Subtype name", _nameBox, null));
        root.Controls.Add(BuildTextRow("English template", _englishPath, () => Browse(_englishPath)));
        root.Controls.Add(BuildTextRow("French template", _frenchPath, () => Browse(_frenchPath)));

        var buttons = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.RightToLeft,
            AutoSize = true
        };

        var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, AutoSize = true };
        var save = new Button { Text = "Save", AutoSize = true };
        save.Click += (_, _) => SaveAndClose();

        buttons.Controls.Add(cancel);
        buttons.Controls.Add(save);
        root.Controls.Add(buttons);

        Controls.Add(root);
        AcceptButton = save;
        CancelButton = cancel;

        if (existing is not null)
        {
            _nameBox.Text = existing.Name;
            _englishPath.Text = existing.EnglishTemplatePath;
            _frenchPath.Text = existing.FrenchTemplatePath;
        }
    }

    private static Control BuildTextRow(string label, TextBox textBox, Action? browse)
    {
        var row = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            ColumnCount = browse is null ? 2 : 3,
            Margin = new Padding(0, 4, 0, 4)
        };

        row.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
        row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        row.Controls.Add(new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
        row.Controls.Add(textBox, 1, 0);

        if (browse is not null)
        {
            row.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            var button = new Button { Text = "Browse...", AutoSize = true };
            button.Click += (_, _) => browse();
            row.Controls.Add(button, 2, 0);
        }

        return row;
    }

    private void Browse(TextBox target)
    {
        using var dialog = new OpenFileDialog
        {
            Filter = "Word documents (*.docx)|*.docx",
            CheckFileExists = true,
            Multiselect = false,
            Title = "Select Word template"
        };

        if (dialog.ShowDialog(this) == DialogResult.OK)
            target.Text = dialog.FileName;
    }

    private void SaveAndClose()
    {
        string name = _nameBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(name))
        {
            MessageBox.Show(this, "Enter a subtype name.", "Template Subtype",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        if (!ConfigurationValidation.TryValidateTemplatePath(_englishPath.Text, "English", out string error) ||
            !ConfigurationValidation.TryValidateTemplatePath(_frenchPath.Text, "French", out error))
        {
            MessageBox.Show(this, error, "Template files", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        SubtypeName = name;
        EnglishTemplatePath = _englishPath.Text;
        FrenchTemplatePath = _frenchPath.Text;
        DialogResult = DialogResult.OK;
        Close();
    }
}
