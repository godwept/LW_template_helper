namespace LWTemplateHelper;

internal sealed class MainForm : Form
{
    private readonly ComboBox _templateTypeCombo = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 260 };
    private readonly ComboBox _subtypeCombo = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 260 };
    private readonly Button _newTypeButton = new() { Text = "New Type", AutoSize = true };
    private readonly Button _editTypeButton = new() { Text = "Edit Type", AutoSize = true };
    private readonly Button _newSubtypeButton = new() { Text = "New Subtype", AutoSize = true };
    private readonly Button _editSubtypeButton = new() { Text = "Edit Subtype", AutoSize = true };

    private readonly DataGridView _bookmarkGrid = CreateBookmarkGrid();
    private readonly DataGridView _infoBookmarkGrid = CreateBookmarkGrid();
    private readonly DataGridView _otherBookmarkGrid = CreateBookmarkGrid();
    private readonly Label _bookmarkSummary = new()
    {
        AutoSize = true,
        Text = "Open working copies to compare bookmark status.",
        Anchor = AnchorStyles.Left
    };
    private readonly Button _refreshStatusButton = new() { Text = "Refresh Status", AutoSize = true, Enabled = false };
    private readonly Button _locateEnglishButton = new() { Text = "Locate English", AutoSize = true, Enabled = false };
    private readonly Button _locateFrenchButton = new() { Text = "Locate French", AutoSize = true, Enabled = false };
    private readonly Button _locateBothButton = new() { Text = "Locate Both", AutoSize = true, Enabled = false };
    private readonly Button _addBookmarkButton = new() { Text = "Add Bookmark", AutoSize = true, Enabled = false };
    private readonly Button _replaceRangeButton = new() { Text = "Replace Range", AutoSize = true, Enabled = false };
    private readonly Button _renameBookmarkButton = new() { Text = "Rename", AutoSize = true, Enabled = false };
    private readonly Button _deleteBookmarkButton = new() { Text = "Delete Bookmark", AutoSize = true, Enabled = false };
    private readonly Button _addInfoButton = new() { Text = "Add INFO", AutoSize = true, Enabled = false };
    private readonly Label _infoSummary = new() { AutoSize = true, Text = "INFO bookmarks: none", Anchor = AnchorStyles.Left };

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
    private string? _selectedBookmarkName;
    private bool _selectedBookmarkIsConfigured;
    private bool _selectedBookmarkIsInfo;

    public MainForm()
    {
        Text = "Letter Wizard Template Bookmark Manager";
        StartPosition = FormStartPosition.CenterScreen;
        MinimumSize = new Size(900, 720);
        Size = new Size(1040, 820);

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

        var statusGroup = new GroupBox
        {
            Text = "Session status",
            Dock = DockStyle.Fill,
            Padding = new Padding(10),
            AutoSize = true
        };
        statusGroup.Controls.Add(_status);
        root.Controls.Add(statusGroup);

        Controls.Add(root);

        _templateTypeCombo.SelectedIndexChanged += (_, _) => TemplateTypeChanged();
        _subtypeCombo.SelectedIndexChanged += (_, _) => SubtypeChanged();
        _newTypeButton.Click += (_, _) => CreateTemplateType();
        _editTypeButton.Click += (_, _) => EditTemplateType();
        _newSubtypeButton.Click += (_, _) => CreateSubtype();
        _editSubtypeButton.Click += (_, _) => EditSubtype();

        _bookmarkGrid.CellClick += (_, e) =>
            SelectBookmarkRow(_bookmarkGrid, e.RowIndex, configured: true, info: false);
        _infoBookmarkGrid.CellClick += (_, e) =>
            SelectBookmarkRow(_infoBookmarkGrid, e.RowIndex, configured: false, info: true);
        _otherBookmarkGrid.CellClick += (_, e) =>
            SelectBookmarkRow(_otherBookmarkGrid, e.RowIndex, configured: false, info: false);
        _refreshStatusButton.Click += (_, _) => RefreshBookmarkStatus();
        _locateEnglishButton.Click += (_, _) => LocateSelectedBookmark(_wordSession.LocateEnglish);
        _locateFrenchButton.Click += (_, _) => LocateSelectedBookmark(_wordSession.LocateFrench);
        _locateBothButton.Click += (_, _) => LocateSelectedBookmark(_wordSession.LocateBoth);
        _addBookmarkButton.Click += (_, _) => AddSelectedBookmark();
        _replaceRangeButton.Click += (_, _) => ReplaceSelectedBookmarkRange();
        _renameBookmarkButton.Click += (_, _) => RenameSelectedBookmark();
        _deleteBookmarkButton.Click += (_, _) => DeleteSelectedBookmark();
        _addInfoButton.Click += (_, _) => AddInfoBookmark();
        _openButton.Click += (_, _) => OpenWorkingCopies();
        _closeButton.Click += (_, _) => CloseSession();
        _sideBySideButton.Click += (_, _) => RunWordAction(_wordSession.ArrangeSideBySide);
        _englishFocusButton.Click += (_, _) => RunWordAction(_wordSession.FocusEnglish);
        _frenchFocusButton.Click += (_, _) => RunWordAction(_wordSession.FocusFrench);
        FormClosing += (_, _) => _wordSession.Dispose();

        LoadConfigurations();
    }

    private static DataGridView CreateBookmarkGrid()
    {
        var grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            AllowUserToResizeRows = false,
            RowHeadersVisible = false,
            MultiSelect = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            AutoGenerateColumns = false,
            BackgroundColor = SystemColors.Window
        };

        grid.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "Bookmark",
            HeaderText = "Bookmark",
            AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
            FillWeight = 60
        });

        grid.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "English",
            HeaderText = "EN",
            AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
            FillWeight = 20,
            DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter }
        });

        grid.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "French",
            HeaderText = "FR",
            AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
            FillWeight = 20,
            DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter }
        });

        return grid;
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
            Text = "Bookmark status",
            Dock = DockStyle.Fill,
            Padding = new Padding(10)
        };

        var table = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 9
        };

        table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        table.RowStyles.Add(new RowStyle(SizeType.Percent, 50));
        table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        table.RowStyles.Add(new RowStyle(SizeType.Percent, 25));
        table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        table.RowStyles.Add(new RowStyle(SizeType.Percent, 25));

        var statusToolbar = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            WrapContents = false
        };
        statusToolbar.Controls.Add(_refreshStatusButton);
        statusToolbar.Controls.Add(_bookmarkSummary);

        var locateToolbar = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            WrapContents = false
        };
        locateToolbar.Controls.AddRange([_locateEnglishButton, _locateFrenchButton, _locateBothButton]);

        var editToolbar = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            WrapContents = false
        };
        editToolbar.Controls.AddRange(
            [_addBookmarkButton, _addInfoButton, _replaceRangeButton, _renameBookmarkButton, _deleteBookmarkButton]);

        var infoHeader = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            WrapContents = false
        };
        infoHeader.Controls.Add(new Label
        {
            Text = "INFO bookmarks",
            AutoSize = true,
            Margin = new Padding(3, 8, 8, 3)
        });
        infoHeader.Controls.Add(_infoSummary);

        table.Controls.Add(statusToolbar);
        table.Controls.Add(locateToolbar);
        table.Controls.Add(editToolbar);
        table.Controls.Add(new Label
        {
            Text = "Configured bookmarks — click a row to select it",
            AutoSize = true,
            Margin = new Padding(3, 8, 3, 3)
        });
        table.Controls.Add(_bookmarkGrid);
        table.Controls.Add(infoHeader);
        table.Controls.Add(_infoBookmarkGrid);
        table.Controls.Add(new Label
        {
            Text = "Other Bookmarks",
            AutoSize = true,
            Margin = new Padding(3, 8, 3, 3)
        });
        table.Controls.Add(_otherBookmarkGrid);

        group.Controls.Add(table);
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

        _subtypeCombo.Items.Clear();
        PopulateConfiguredBookmarkRows(type);
        _infoBookmarkGrid.Rows.Clear();
        _infoSummary.Text = "INFO bookmarks: none";
        _otherBookmarkGrid.Rows.Clear();
        ClearSelectedBookmark();

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

        foreach (var subtype in type.Subtypes.OrderBy(x => x.Name, StringComparer.CurrentCultureIgnoreCase))
            _subtypeCombo.Items.Add(subtype);

        if (_subtypeCombo.Items.Count > 0)
            _subtypeCombo.SelectedIndex = 0;
        else
            SubtypeChanged();
    }

    private void PopulateConfiguredBookmarkRows(TemplateTypeConfig? type)
    {
        _bookmarkGrid.Rows.Clear();

        if (type is null)
        {
            _bookmarkSummary.Text = "Create or select a Template Type.";
            return;
        }

        foreach (string bookmark in type.Bookmarks)
        {
            int rowIndex = _bookmarkGrid.Rows.Add(bookmark, "—", "—");
            _bookmarkGrid.Rows[rowIndex].DefaultCellStyle.BackColor = SystemColors.Window;
        }

        _bookmarkGrid.ClearSelection();

        _bookmarkSummary.Text = _wordSession.IsOpen
            ? "Click Refresh Status to compare the open working copies."
            : "Open working copies to compare bookmark status.";
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

    private void SelectBookmarkRow(
        DataGridView selectedGrid,
        int rowIndex,
        bool configured,
        bool info)
    {
        if (rowIndex < 0 || rowIndex >= selectedGrid.Rows.Count)
            return;

        object? value = selectedGrid.Rows[rowIndex].Cells["Bookmark"].Value;
        _selectedBookmarkName = value?.ToString();
        _selectedBookmarkIsConfigured = configured;
        _selectedBookmarkIsInfo = info;

        foreach (var grid in new[] { _bookmarkGrid, _infoBookmarkGrid, _otherBookmarkGrid })
        {
            if (!ReferenceEquals(grid, selectedGrid))
                grid.ClearSelection();
        }

        UpdateBookmarkActionButtons();
    }

    private void ClearSelectedBookmark()
    {
        _selectedBookmarkName = null;
        _selectedBookmarkIsConfigured = false;
        _selectedBookmarkIsInfo = false;
        _bookmarkGrid.ClearSelection();
        _infoBookmarkGrid.ClearSelection();
        _otherBookmarkGrid.ClearSelection();
        UpdateBookmarkActionButtons();
    }

    private void UpdateBookmarkActionButtons()
    {
        bool enabled = _wordSession.IsOpen && !string.IsNullOrWhiteSpace(_selectedBookmarkName);

        _locateEnglishButton.Enabled = enabled;
        _locateFrenchButton.Enabled = enabled;
        _locateBothButton.Enabled = enabled;

        _addBookmarkButton.Enabled = enabled && _selectedBookmarkIsConfigured;
        _replaceRangeButton.Enabled = enabled;
        _renameBookmarkButton.Enabled = enabled && !_selectedBookmarkIsInfo;
        _deleteBookmarkButton.Enabled = enabled;
    }

    private void LocateSelectedBookmark(Action<string> locateAction)
    {
        if (string.IsNullOrWhiteSpace(_selectedBookmarkName))
            return;

        try
        {
            locateAction(_selectedBookmarkName);
            _status.Text = $"Located bookmark: {_selectedBookmarkName}";
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Locate bookmark", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    private void AddSelectedBookmark()
    {
        if (string.IsNullOrWhiteSpace(_selectedBookmarkName) || !_selectedBookmarkIsConfigured)
            return;

        RunBookmarkEdit(
            () => _wordSession.AddBookmark(_selectedBookmarkName),
            $"Added bookmark '{_selectedBookmarkName}' using the current English/French Word selections.");
    }

    private void AddInfoBookmark()
    {
        if (!_wordSession.IsOpen)
            return;

        try
        {
            string bookmarkName = _wordSession.AddNextInfoBookmark();
            RefreshBookmarkStatus();
            _status.Text = $"Added {bookmarkName} using the current English/French Word selections.";
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Add INFO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    private void ReplaceSelectedBookmarkRange()
    {
        if (string.IsNullOrWhiteSpace(_selectedBookmarkName))
            return;

        RunBookmarkEdit(
            () => _wordSession.ReplaceBookmarkRange(_selectedBookmarkName),
            $"Replaced range for bookmark '{_selectedBookmarkName}' using the current English/French Word selections.");
    }

    private void RenameSelectedBookmark()
    {
        if (string.IsNullOrWhiteSpace(_selectedBookmarkName))
            return;

        string oldName = _selectedBookmarkName;

        using var dialog = new BookmarkNameDialog(oldName);
        if (dialog.ShowDialog(this) != DialogResult.OK)
            return;

        string newName = dialog.BookmarkName;

        RunBookmarkEdit(
            () => _wordSession.RenameBookmark(oldName, newName),
            $"Renamed bookmark '{oldName}' to '{newName}'. Template Type configuration was not changed.");
    }

    private void DeleteSelectedBookmark()
    {
        if (string.IsNullOrWhiteSpace(_selectedBookmarkName))
            return;

        string bookmarkName = _selectedBookmarkName;

        var result = MessageBox.Show(
            this,
            $"Remove bookmark markers for '{bookmarkName}' from the working copies?\n\nThe bookmarked text will not be deleted.",
            "Delete bookmark markers",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning,
            MessageBoxDefaultButton.Button2);

        if (result != DialogResult.Yes)
            return;

        RunBookmarkEdit(
            () => _wordSession.DeleteBookmarkMarkers(bookmarkName),
            $"Removed bookmark markers for '{bookmarkName}'. Document text was preserved.");
    }

    private void RunBookmarkEdit(Action action, string successMessage)
    {
        try
        {
            action();
            RefreshBookmarkStatus();
            _status.Text = successMessage;
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Bookmark operation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
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

        if (SaveConfiguration(config))
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

        if (SaveConfiguration(current, previousName))
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

        if (SaveConfiguration(type))
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

        if (SaveConfiguration(type))
            LoadConfigurations(type.Name, subtype.Name);
    }

    private bool SaveConfiguration(TemplateTypeConfig config, string? previousName = null)
    {
        try
        {
            _configStore.Save(config, previousName);
            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Save configuration error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return false;
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
            RefreshBookmarkStatus();

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

    private void RefreshBookmarkStatus()
    {
        var type = SelectedType;
        if (type is null || !_wordSession.IsOpen)
            return;

        try
        {
            var (english, french) = _wordSession.GetBookmarkNames();
            var configured = new HashSet<string>(type.Bookmarks, StringComparer.OrdinalIgnoreCase);

            int both = 0;
            int mismatch = 0;
            int neither = 0;

            _bookmarkGrid.Rows.Clear();

            foreach (string bookmark in type.Bookmarks)
            {
                bool enExists = english.Contains(bookmark);
                bool frExists = french.Contains(bookmark);

                if (enExists && frExists)
                    both++;
                else if (enExists != frExists)
                    mismatch++;
                else
                    neither++;

                int rowIndex = _bookmarkGrid.Rows.Add(
                    bookmark,
                    enExists ? "Exists" : "Missing",
                    frExists ? "Exists" : "Missing");

                ApplyConfiguredStatusStyle(_bookmarkGrid.Rows[rowIndex], enExists, frExists);
            }

            var infoNames = english
                .Union(french, StringComparer.OrdinalIgnoreCase)
                .Select(name => new
                {
                    Name = name,
                    IsInfo = WordSession.TryGetInfoNumber(name, out int number),
                    Number = number
                })
                .Where(x => x.IsInfo)
                .OrderBy(x => x.Number)
                .ThenBy(x => x.Name, StringComparer.OrdinalIgnoreCase)
                .ToList();

            int infoBoth = 0;
            int infoMismatch = 0;

            _infoBookmarkGrid.Rows.Clear();

            foreach (var info in infoNames)
            {
                bool enExists = english.Contains(info.Name);
                bool frExists = french.Contains(info.Name);

                if (enExists && frExists)
                    infoBoth++;
                else
                    infoMismatch++;

                int rowIndex = _infoBookmarkGrid.Rows.Add(
                    info.Name,
                    enExists ? "Exists" : "—",
                    frExists ? "Exists" : "—");

                ApplyInfoStatusStyle(_infoBookmarkGrid.Rows[rowIndex], enExists, frExists);
            }

            int highestInfo = infoNames.Count == 0 ? 0 : infoNames.Max(x => x.Number);
            _infoSummary.Text =
                $"Both: {infoBoth}, mismatch: {infoMismatch} | Next: INFO_{highestInfo + 1}";

            var infoNameSet = new HashSet<string>(
                infoNames.Select(x => x.Name),
                StringComparer.OrdinalIgnoreCase);

            var otherNames = english
                .Union(french, StringComparer.OrdinalIgnoreCase)
                .Where(name => !configured.Contains(name) && !infoNameSet.Contains(name))
                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                .ToList();

            _otherBookmarkGrid.Rows.Clear();

            foreach (string bookmark in otherNames)
            {
                bool enExists = english.Contains(bookmark);
                bool frExists = french.Contains(bookmark);

                int rowIndex = _otherBookmarkGrid.Rows.Add(
                    bookmark,
                    enExists ? "Exists" : "—",
                    frExists ? "Exists" : "—");

                ApplyOtherBookmarkStyle(_otherBookmarkGrid.Rows[rowIndex], enExists, frExists);
            }

            _bookmarkGrid.ClearSelection();
            _infoBookmarkGrid.ClearSelection();
            _otherBookmarkGrid.ClearSelection();
            _selectedBookmarkName = null;
            _selectedBookmarkIsConfigured = false;
            _selectedBookmarkIsInfo = false;
            UpdateBookmarkActionButtons();

            _bookmarkSummary.Text =
                $"Configured: {both} both, {mismatch} mismatch, {neither} unused | INFO: {infoNames.Count} | Other: {otherNames.Count}";

            _status.Text = $"Bookmark status refreshed. EN: {english.Count}, FR: {french.Count}.";
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Bookmark status error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    private static void ApplyConfiguredStatusStyle(DataGridViewRow row, bool english, bool french)
    {
        row.DefaultCellStyle.BackColor = english != french
            ? Color.MistyRose
            : english
                ? Color.Honeydew
                : Color.WhiteSmoke;
    }

    private static void ApplyInfoStatusStyle(DataGridViewRow row, bool english, bool french)
    {
        row.DefaultCellStyle.BackColor = english != french
            ? Color.MistyRose
            : Color.Honeydew;
    }

    private static void ApplyOtherBookmarkStyle(DataGridViewRow row, bool english, bool french)
    {
        row.DefaultCellStyle.BackColor = english != french
            ? Color.MistyRose
            : Color.LemonChiffon;
    }

    private void CloseSession()
    {
        try
        {
            _wordSession.Close();
            _status.Text =
                $"Word session closed. Source templates were not modified. Ready: {SelectedType?.Name} / {SelectedSubtype?.Name}";
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Close error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
        finally
        {
            SetSessionControls(false);
            PopulateConfiguredBookmarkRows(SelectedType);
            _infoBookmarkGrid.Rows.Clear();
            _infoSummary.Text = "INFO bookmarks: none";
            _otherBookmarkGrid.Rows.Clear();
            ClearSelectedBookmark();
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
        _refreshStatusButton.Enabled = open;
        _addInfoButton.Enabled = open;
        _sideBySideButton.Enabled = open;
        _englishFocusButton.Enabled = open;
        _frenchFocusButton.Enabled = open;

        UpdateBookmarkActionButtons();
    }
}
