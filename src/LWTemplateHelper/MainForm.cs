namespace LWTemplateHelper;

internal sealed class MainForm : Form
{
    private readonly ComboBox _templateTypeCombo = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 260 };
    private readonly ComboBox _subtypeCombo = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 260 };
    private readonly Button _newTypeButton = new() { Text = "New Type", AutoSize = true };
    private readonly Button _editTypeButton = new() { Text = "Edit Type", AutoSize = true };
    private readonly Button _newSubtypeButton = new() { Text = "New Subtype", AutoSize = true };
    private readonly Button _editSubtypeButton = new() { Text = "Edit Subtype", AutoSize = true };
    private readonly ListBox _bookmarkList = new() { Dock = DockStyle.Fill };
    private readonly TextBox _englishPath = new() { ReadOnly = true, Dock = DockStyle.Fill };
    private readonly TextBox _frenchPath = new() { ReadOnly = true, Dock = DockStyle.Fill };
    private readonly Label _status = new() { AutoSize = true, Text = "Create or select a Template Type and subtype." };
    private readonly Button _openButton = new() { Text = "Open Working Copies", AutoSize = true, Enabled = false };
    private readonly Button _closeButton = new() { Text = "Close Word Session", AutoSize = true, Enabled = false };
    private readonly Button _sideBySideButton = new() { Text = "Side by Side", AutoSize = true, Enabled = false };
    private readonly Button _englishFocusButton = new() { Text = "English Focus", AutoSize = true, Enabled = false };
    private readonly Button _frenchFocusButton = new() { Text = "French Focus", AutoSize = true, Enabled = false };

    private readonly WordSession _wordSession = new();
    private readonly TemplateConfigStore _configStore = new();
    private List<TemplateTypeConfig> _templateTypes = [];

    public MainForm()
    {
        Text = "Letter Wizard Template Bookmark Manager";
        StartPosition = FormStartPosition.CenterScreen;
        MinimumSize = new Size(850, 620);
        Size = new Size(980, 700);

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(12),
            ColumnCount = 1,
            RowCount = 6
        };

        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        root.Controls.Add(BuildConfigurationGroup());
        root.Controls.Add(BuildBookmarkGroup());
        root.Controls.Add(BuildSourceGroup());

        var sessionButtons = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
        sessionButtons.Controls.AddRange([_openButton, _closeButton]);
        root.Controls.Add(sessionButtons);

        var layoutButtons = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
        layoutButtons.Controls.AddRange([_sideBySideButton, _englishFocusButton, _frenchFocusButton]);
        root.Controls.Add(layoutButtons);

        var statusGroup = new GroupBox { Text = "Session status", Dock = DockStyle.Fill, Padding = new Padding(10), AutoSize = true };
        statusGroup.Controls.Add(_status);
        root.Controls.Add(statusGroup);

        Controls.Add(root);

        _templateTypeCombo.SelectedIndexChanged += (_, _) => TemplateTypeChanged();
        _subtypeCombo.SelectedIndexChanged += (_, _) => SubtypeChanged();
        _newTypeButton.Click += (_, _) => CreateTemplateType();
        _editTypeButton.Click += (_, _) => EditTemplateType();
        _newSubtypeButton.Click += (_, _) => CreateSubtype();
        _editSubtypeButton.Click += (_, _) => EditSubtype();

        _openButton.Click += (_, _) => OpenWorkingCopies();
        _closeButton.Click += (_, _) => CloseSession();
        _sideBySideButton.Click += (_, _) => RunWordAction(_wordSession.ArrangeSideBySide);
        _englishFocusButton.Click += (_, _) => RunWordAction(_wordSession.FocusEnglish);
        _frenchFocusButton.Click += (_, _) => RunWordAction(_wordSession.FocusFrench);
        FormClosing += (_, _) => _wordSession.Dispose();

        LoadConfigurations();
    }

    private Control BuildConfigurationGroup()
    {
        var group = new GroupBox
        {
            Text = "Template configuration",
            Dock = DockStyle.Top,
            AutoSize = true,
            Padding = new Padding(10)
        };

        var table = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            ColumnCount = 4,
            RowCount = 2
        };

        table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 105));
        table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

        table.Controls.Add(new Label { Text = "Template Type", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
        table.Controls.Add(_templateTypeCombo, 1, 0);
        table.Controls.Add(_newTypeButton, 2, 0);
        table.Controls.Add(_editTypeButton, 3, 0);

        table.Controls.Add(new Label { Text = "Subtype", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 1);
        table.Controls.Add(_subtypeCombo, 1, 1);
        table.Controls.Add(_newSubtypeButton, 2, 1);
        table.Controls.Add(_editSubtypeButton, 3, 1);

        group.Controls.Add(table);
        return group;
    }

    private Control BuildBookmarkGroup()
    {
        var group = new GroupBox
        {
            Text = "Configured bookmarks",
            Dock = DockStyle.Fill,
            Padding = new Padding(10)
        };

        group.Controls.Add(_bookmarkList);
        return group;
    }

    private Control BuildSourceGroup()
    {
        var group = new GroupBox
        {
            Text = "Selected subtype source templates",
            Dock = DockStyle.Top,
            AutoSize = true,
            Padding = new Padding(10)
        };

        var table = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            ColumnCount = 2,
            RowCount = 2
        };

        table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
        table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        table.Controls.Add(new Label { Text = "English template", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
        table.Controls.Add(_englishPath, 1, 0);
        table.Controls.Add(new Label { Text = "French template", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 1);
        table.Controls.Add(_frenchPath, 1, 1);

        group.Controls.Add(table);
        return group;
    }

    private TemplateTypeConfig? SelectedType => _templateTypeCombo.SelectedItem as TemplateTypeConfig;
    private TemplateSubtypeConfig? SelectedSubtype => _subtypeCombo.SelectedItem as TemplateSubtypeConfig;

    private void LoadConfigurations(string? preferredType = null, string? preferredSubtype = null)
    {
        try
        {
            _templateTypes = _configStore.LoadAll();

            _templateTypeCombo.BeginUpdate();
            _templateTypeCombo.Items.Clear();
            foreach (var type in _templateTypes)
                _templateTypeCombo.Items.Add(type);
            _templateTypeCombo.EndUpdate();

            if (_templateTypes.Count == 0)
            {
                _templateTypeCombo.SelectedIndex = -1;
                TemplateTypeChanged();
                _status.Text = $"No Template Types configured. Configuration folder: {_configStore.ConfigDirectory}";
                return;
            }

            int typeIndex = 0;
            if (!string.IsNullOrWhiteSpace(preferredType))
            {
                int found = _templateTypes.FindIndex(x =>
                    string.Equals(x.Name, preferredType, StringComparison.OrdinalIgnoreCase));
                if (found >= 0)
                    typeIndex = found;
            }

            _templateTypeCombo.SelectedIndex = typeIndex;

            if (!string.IsNullOrWhiteSpace(preferredSubtype))
            {
                for (int i = 0; i < _subtypeCombo.Items.Count; i++)
                {
                    if (_subtypeCombo.Items[i] is TemplateSubtypeConfig subtype &&
                        string.Equals(subtype.Name, preferredSubtype, StringComparison.OrdinalIgnoreCase))
                    {
                        _subtypeCombo.SelectedIndex = i;
                        break;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Configuration error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void TemplateTypeChanged()
    {
        var type = SelectedType;

        _bookmarkList.Items.Clear();
        _subtypeCombo.Items.Clear();

        if (type is null)
        {
            _editTypeButton.Enabled = false;
            _newSubtypeButton.Enabled = false;
            _editSubtypeButton.Enabled = false;
            ClearSubtypeDisplay();
            return;
        }

        _editTypeButton.Enabled = !_wordSession.IsOpen;
        _newSubtypeButton.Enabled = !_wordSession.IsOpen;

        foreach (string bookmark in type.Bookmarks)
            _bookmarkList.Items.Add(bookmark);

        foreach (var subtype in type.Subtypes.OrderBy(x => x.Name, StringComparer.CurrentCultureIgnoreCase))
            _subtypeCombo.Items.Add(subtype);

        if (_subtypeCombo.Items.Count > 0)
            _subtypeCombo.SelectedIndex = 0;
        else
            SubtypeChanged();
    }

    private void SubtypeChanged()
    {
        var subtype = SelectedSubtype;

        if (subtype is null)
        {
            ClearSubtypeDisplay();
            return;
        }

        _englishPath.Text = subtype.EnglishTemplatePath;
        _frenchPath.Text = subtype.FrenchTemplatePath;
        _editSubtypeButton.Enabled = !_wordSession.IsOpen;
        _openButton.Enabled = !_wordSession.IsOpen;
        _status.Text = $"Ready: {SelectedType?.Name} / {subtype.Name}";
    }

    private void ClearSubtypeDisplay()
    {
        _englishPath.Clear();
        _frenchPath.Clear();
        _editSubtypeButton.Enabled = false;
        _openButton.Enabled = false;
    }

    private void CreateTemplateType()
    {
        using var dialog = new TemplateTypeDialog();
        if (dialog.ShowDialog(this) != DialogResult.OK)
            return;

        if (_templateTypes.Any(x => string.Equals(x.Name, dialog.TemplateTypeName, StringComparison.OrdinalIgnoreCase)))
        {
            MessageBox.Show(this, "A Template Type with that name already exists.", "Template Type",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var config = new TemplateTypeConfig
        {
            Name = dialog.TemplateTypeName,
            Bookmarks = dialog.BookmarkNames
        };

        SaveConfiguration(config);
        LoadConfigurations(config.Name);
    }

    private void EditTemplateType()
    {
        var current = SelectedType;
        if (current is null)
            return;

        using var dialog = new TemplateTypeDialog(current);
        if (dialog.ShowDialog(this) != DialogResult.OK)
            return;

        if (_templateTypes.Any(x =>
                !ReferenceEquals(x, current) &&
                string.Equals(x.Name, dialog.TemplateTypeName, StringComparison.OrdinalIgnoreCase)))
        {
            MessageBox.Show(this, "A Template Type with that name already exists.", "Template Type",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        string previousName = current.Name;
        string? selectedSubtype = SelectedSubtype?.Name;

        current.Name = dialog.TemplateTypeName;
        current.Bookmarks = dialog.BookmarkNames;

        SaveConfiguration(current, previousName);
        LoadConfigurations(current.Name, selectedSubtype);
    }

    private void CreateSubtype()
    {
        var type = SelectedType;
        if (type is null)
            return;

        using var dialog = new TemplateSubtypeDialog();
        if (dialog.ShowDialog(this) != DialogResult.OK)
            return;

        if (type.Subtypes.Any(x => string.Equals(x.Name, dialog.SubtypeName, StringComparison.OrdinalIgnoreCase)))
        {
            MessageBox.Show(this, "A subtype with that name already exists for this Template Type.", "Template Subtype",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        type.Subtypes.Add(new TemplateSubtypeConfig
        {
            Name = dialog.SubtypeName,
            EnglishTemplatePath = dialog.EnglishTemplatePath,
            FrenchTemplatePath = dialog.FrenchTemplatePath
        });

        SaveConfiguration(type);
        LoadConfigurations(type.Name, dialog.SubtypeName);
    }

    private void EditSubtype()
    {
        var type = SelectedType;
        var subtype = SelectedSubtype;
        if (type is null || subtype is null)
            return;

        using var dialog = new TemplateSubtypeDialog(subtype);
        if (dialog.ShowDialog(this) != DialogResult.OK)
            return;

        if (type.Subtypes.Any(x =>
                !ReferenceEquals(x, subtype) &&
                string.Equals(x.Name, dialog.SubtypeName, StringComparison.OrdinalIgnoreCase)))
        {
            MessageBox.Show(this, "A subtype with that name already exists for this Template Type.", "Template Subtype",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        subtype.Name = dialog.SubtypeName;
        subtype.EnglishTemplatePath = dialog.EnglishTemplatePath;
        subtype.FrenchTemplatePath = dialog.FrenchTemplatePath;

        SaveConfiguration(type);
        LoadConfigurations(type.Name, subtype.Name);
    }

    private void SaveConfiguration(TemplateTypeConfig config, string? previousName = null)
    {
        try
        {
            _configStore.Save(config, previousName);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Save configuration error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void OpenWorkingCopies()
    {
        var subtype = SelectedSubtype;
        if (subtype is null)
        {
            MessageBox.Show(this, "Select a configured subtype first.", "Template Subtype",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        try
        {
            _openButton.Enabled = false;
            UseWaitCursor = true;
            _status.Text = "Creating working copies and opening Word...";

            _wordSession.Open(subtype.EnglishTemplatePath, subtype.FrenchTemplatePath);
            SetSessionControls(true);

            _status.Text =
                $"Word session open for {SelectedType?.Name} / {subtype.Name}. Working files: {_wordSession.SessionDirectory}";
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
                SetSessionControls(false);
        }
    }

    private void CloseSession()
    {
        try
        {
            _wordSession.Close();
            _status.Text = $"Word session closed. Source templates were not modified. Ready: {SelectedType?.Name} / {SelectedSubtype?.Name}";
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
        _templateTypeCombo.Enabled = !open;
        _subtypeCombo.Enabled = !open;
        _newTypeButton.Enabled = !open;
        _editTypeButton.Enabled = !open && SelectedType is not null;
        _newSubtypeButton.Enabled = !open && SelectedType is not null;
        _editSubtypeButton.Enabled = !open && SelectedSubtype is not null;

        _openButton.Enabled = !open && SelectedSubtype is not null;
        _closeButton.Enabled = open;
        _sideBySideButton.Enabled = open;
        _englishFocusButton.Enabled = open;
        _frenchFocusButton.Enabled = open;
    }
}
