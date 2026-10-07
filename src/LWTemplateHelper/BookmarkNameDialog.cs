namespace LWTemplateHelper;

internal sealed class BookmarkNameDialog : Form
{
    private readonly TextBox _nameBox = new() { Dock = DockStyle.Fill };

    public string BookmarkName { get; private set; } = string.Empty;

    public BookmarkNameDialog(string currentName)
    {
        Text = "Rename Bookmark";
        StartPosition = FormStartPosition.CenterParent;
        MinimizeBox = false;
        MaximizeBox = false;
        ShowInTaskbar = false;
        Size = new Size(460, 160);
        MinimumSize = new Size(420, 150);

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(12),
            ColumnCount = 1,
            RowCount = 3,
            AutoSize = true
        };

        root.Controls.Add(new Label
        {
            Text = $"New name for '{currentName}'",
            AutoSize = true
        });

        _nameBox.Text = currentName;
        _nameBox.SelectAll();
        root.Controls.Add(_nameBox);

        var buttons = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            FlowDirection = FlowDirection.RightToLeft
        };

        var cancel = new Button
        {
            Text = "Cancel",
            DialogResult = DialogResult.Cancel,
            AutoSize = true
        };

        var rename = new Button
        {
            Text = "Rename",
            AutoSize = true
        };
        rename.Click += (_, _) => SaveAndClose();

        buttons.Controls.Add(cancel);
        buttons.Controls.Add(rename);
        root.Controls.Add(buttons);

        Controls.Add(root);
        AppTheme.Apply(this);
        AppTheme.StylePrimaryButton(rename);
        AcceptButton = rename;
        CancelButton = cancel;
    }

    protected override void OnShown(EventArgs e)
    {
        base.OnShown(e);
        _nameBox.Focus();
    }

    private void SaveAndClose()
    {
        string name = _nameBox.Text.Trim();

        if (!ConfigurationValidation.TryValidateBookmarkName(name, out string error))
        {
            MessageBox.Show(this, error, "Bookmark name", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        BookmarkName = name;
        DialogResult = DialogResult.OK;
        Close();
    }
}
