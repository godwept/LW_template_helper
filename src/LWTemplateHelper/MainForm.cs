namespace LWTemplateHelper;

internal sealed class MainForm : Form
{
    private readonly TextBox _englishPath = new() { ReadOnly = true, Dock = DockStyle.Fill };
    private readonly TextBox _frenchPath = new() { ReadOnly = true, Dock = DockStyle.Fill };
    private readonly Label _status = new() { AutoSize = true, Text = "Select English and French .docx templates." };
    private readonly Button _openButton = new() { Text = "Open Working Copies", AutoSize = true };
    private readonly Button _closeButton = new() { Text = "Close Word Session", AutoSize = true, Enabled = false };
    private readonly Button _sideBySideButton = new() { Text = "Side by Side", AutoSize = true, Enabled = false };
    private readonly Button _englishFocusButton = new() { Text = "English Focus", AutoSize = true, Enabled = false };
    private readonly Button _frenchFocusButton = new() { Text = "French Focus", AutoSize = true, Enabled = false };
    private readonly WordSession _wordSession = new();

    public MainForm()
    {
        Text = "Letter Wizard Template Bookmark Manager";
        StartPosition = FormStartPosition.CenterScreen;
        MinimumSize = new Size(760, 310);
        Size = new Size(900, 340);

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(12),
            ColumnCount = 1,
            RowCount = 5,
            AutoSize = true
        };

        root.Controls.Add(BuildFileRow("English template", _englishPath, () => SelectTemplate(_englishPath)));
        root.Controls.Add(BuildFileRow("French template", _frenchPath, () => SelectTemplate(_frenchPath)));

        var sessionButtons = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
        sessionButtons.Controls.AddRange([_openButton, _closeButton]);
        root.Controls.Add(sessionButtons);

        var layoutButtons = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
        layoutButtons.Controls.AddRange([_sideBySideButton, _englishFocusButton, _frenchFocusButton]);
        root.Controls.Add(layoutButtons);

        var statusGroup = new GroupBox { Text = "Session status", Dock = DockStyle.Fill, Padding = new Padding(10) };
        statusGroup.Controls.Add(_status);
        root.Controls.Add(statusGroup);

        Controls.Add(root);

        _openButton.Click += (_, _) => OpenWorkingCopies();
        _closeButton.Click += (_, _) => CloseSession();
        _sideBySideButton.Click += (_, _) => RunWordAction(_wordSession.ArrangeSideBySide);
        _englishFocusButton.Click += (_, _) => RunWordAction(_wordSession.FocusEnglish);
        _frenchFocusButton.Click += (_, _) => RunWordAction(_wordSession.FocusFrench);
        FormClosing += (_, _) => _wordSession.Dispose();
    }

    private static Control BuildFileRow(string labelText, TextBox pathBox, Action browse)
    {
        var row = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            ColumnCount = 3
        };
        row.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
        row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        row.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

        row.Controls.Add(new Label { Text = labelText, AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
        row.Controls.Add(pathBox, 1, 0);

        var button = new Button { Text = "Browse...", AutoSize = true };
        button.Click += (_, _) => browse();
        row.Controls.Add(button, 2, 0);
        return row;
    }

    private void SelectTemplate(TextBox target)
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

    private void OpenWorkingCopies()
    {
        if (string.IsNullOrWhiteSpace(_englishPath.Text) || string.IsNullOrWhiteSpace(_frenchPath.Text))
        {
            MessageBox.Show(this, "Select both English and French templates first.", "Templates required",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        try
        {
            _openButton.Enabled = false;
            UseWaitCursor = true;
            _status.Text = "Creating working copies and opening Word...";

            _wordSession.Open(_englishPath.Text, _frenchPath.Text);
            SetSessionControls(true);

            _status.Text =
                $"Word session open. Originals are not open. Working files: {_wordSession.SessionDirectory}";
        }
        catch (Exception ex)
        {
            SetSessionControls(false);
            _status.Text = "Could not open the Word session.";
            MessageBox.Show(this, ex.Message, "Word session error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            UseWaitCursor = false;
            if (!_wordSession.IsOpen)
                _openButton.Enabled = true;
        }
    }

    private void CloseSession()
    {
        try
        {
            _wordSession.Close();
            _status.Text = "Word session closed. Source templates were not modified.";
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Close error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
        finally
        {
            SetSessionControls(false);
        }
    }

    private void RunWordAction(Action action)
    {
        try
        {
            action();
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Word layout error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    private void SetSessionControls(bool open)
    {
        _openButton.Enabled = !open;
        _closeButton.Enabled = open;
        _sideBySideButton.Enabled = open;
        _englishFocusButton.Enabled = open;
        _frenchFocusButton.Enabled = open;
    }
}
