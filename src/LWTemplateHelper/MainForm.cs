namespace LWTemplateHelper;

internal sealed record ValidationResult(string Summary, bool HasBlockingErrors, bool HasWarnings);

internal sealed class MainForm : Form
{
    private readonly ComboBox _templateTypeCombo = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 260 };
    private readonly ComboBox _subtypeCombo = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 260 };
    private readonly ThemedButton _newTypeButton = new() { Text = "New Type", AutoSize = true };
    private readonly ThemedButton _editTypeButton = new() { Text = "Edit Type", AutoSize = true };
    private readonly ThemedButton _newSubtypeButton = new() { Text = "New Subtype", AutoSize = true };
    private readonly ThemedButton _editSubtypeButton = new() { Text = "Edit Subtype", AutoSize = true };

    private readonly DataGridView _bookmarkGrid = CreateBookmarkGrid();
    private readonly DataGridView _infoBookmarkGrid = CreateBookmarkGrid();
    private readonly DataGridView _otherBookmarkGrid = CreateBookmarkGrid();
    private readonly Label _bookmarkSummary = new()
    {
        AutoSize = true,
        Text = "Open working copies to compare bookmark status.",
        Anchor = AnchorStyles.Left
    };
    private readonly ThemedButton _refreshStatusButton = new()
    {
        Text = "↻",
        AutoSize = false,
        Size = new Size(34, 32),
        Enabled = false,
        AccessibleName = "Refresh Status"
    };
    private readonly ThemedButton _locateEnglishButton = new() { Text = "Locate English", AutoSize = true, Enabled = false };
    private readonly ThemedButton _locateFrenchButton = new() { Text = "Locate French", AutoSize = true, Enabled = false };
    private readonly ThemedButton _locateBothButton = new() { Text = "Locate Both", AutoSize = true, Enabled = false };
    private readonly ThemedButton _addBookmarkButton = new() { Text = "Add Bookmark", AutoSize = true, Enabled = false };
    private readonly ThemedButton _replaceRangeButton = new() { Text = "Replace Range", AutoSize = true, Enabled = false };
    private readonly ThemedButton _renameBookmarkButton = new() { Text = "Rename", AutoSize = true, Enabled = false };
    private readonly ThemedButton _deleteBookmarkButton = new() { Text = "Delete Bookmark", AutoSize = true, Enabled = false };
    private readonly ThemedButton _addInfoButton = new() { Text = "Add INFO", AutoSize = true, Enabled = false };
    private readonly Label _infoSummary = new() { AutoSize = true, Text = "INFO bookmarks: none", Anchor = AnchorStyles.Left };

    private readonly TextBox _englishPath = new() { ReadOnly = true, Dock = DockStyle.Fill };
    private readonly TextBox _frenchPath = new() { ReadOnly = true, Dock = DockStyle.Fill };
    private readonly Label _status = new() { AutoSize = true, Text = "Create or select a Template Type and subtype." };
    private readonly ThemedButton _openButton = new() { Text = "Open Working Copies", AutoSize = true, Enabled = false };
    private readonly ThemedButton _closeButton = new() { Text = "Close Word Session", AutoSize = true, Enabled = false };
    private readonly ThemedButton _validateButton = new() { Text = "Validate", AutoSize = true, Enabled = false };
    private readonly ThemedButton _finalizeButton = new() { Text = "Save / Finalize", AutoSize = true, Enabled = false };
    private readonly ThemedButton _sideBySideButton = new() { Text = "Side by Side", AutoSize = true, Enabled = false };
    private readonly ThemedButton _englishFocusButton = new() { Text = "English Focus", AutoSize = true, Enabled = false };
    private readonly ThemedButton _frenchFocusButton = new() { Text = "French Focus", AutoSize = true, Enabled = false };

    private readonly Label _testModeLabel = new()
    {
        AutoSize = true,
        Text = "Test Mode inactive",
        ForeColor = SystemColors.GrayText,
        Anchor = AnchorStyles.Left
    };
    private readonly ThemedButton _startTestButton = new() { Text = "Start Test Mode", AutoSize = true, Enabled = false };
    private readonly ThemedButton _resetTestButton = new() { Text = "Reset Test", AutoSize = true, Enabled = false };
    private readonly ThemedButton _closeTestButton = new() { Text = "Close Test", AutoSize = true, Enabled = false };
    private readonly TextBox _testValueBox = new() { Text = "TEST VALUE", Width = 260, Enabled = false };
    private readonly ThemedButton _testSetValueButton = new() { Text = "Set Value", AutoSize = true, Enabled = false };
    private readonly ThemedButton _testDeleteContentButton = new() { Text = "Delete Content", AutoSize = true, Enabled = false };

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
        MinimumSize = new Size(1050, 700);
        Size = new Size(1320, 820);

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(12),
            ColumnCount = 1,
            RowCount = 5
        };

        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 155));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        root.Controls.Add(BuildConfigurationGroup());
        root.Controls.Add(BuildBookmarkGroup());
        root.Controls.Add(BuildTestModeGroup());
        root.Controls.Add(BuildSourceGroup());

        _status.Dock = DockStyle.Fill;
        _status.Margin = new Padding(6, 4, 6, 4);
        root.Controls.Add(_status);

        Controls.Add(root);

        AppTheme.Apply(this);
        AppTheme.StylePrimaryButton(_openButton);
        AppTheme.StylePrimaryButton(_finalizeButton);
        AppTheme.StylePrimaryButton(_addBookmarkButton);
        AppTheme.StylePrimaryButton(_addInfoButton);
        AppTheme.StylePrimaryButton(_startTestButton);
        AppTheme.StyleDangerButton(_closeButton);
        AppTheme.StyleDangerButton(_deleteBookmarkButton);
        AppTheme.StyleDangerButton(_closeTestButton);
        _testModeLabel.ForeColor = AppTheme.MutedText;
        _status.ForeColor = AppTheme.MutedText;
        _refreshStatusButton.Font = new Font("Segoe UI Symbol", 12F, FontStyle.Bold);

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
        _startTestButton.Click += (_, _) => StartTestMode();
        _resetTestButton.Click += (_, _) => ResetTestMode();
        _closeTestButton.Click += (_, _) => CloseTestMode();
        _testSetValueButton.Click += (_, _) => TestSetValue();
        _testDeleteContentButton.Click += (_, _) => TestDeleteContent();
        _openButton.Click += (_, _) => OpenWorkingCopies();
        _closeButton.Click += (_, _) => CloseSession();
        _validateButton.Click += (_, _) => ShowValidationSummary();
        _finalizeButton.Click += (_, _) => FinalizeTemplates();
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
            Dock = DockStyle.Fill,
            AutoSize = false,
            Padding = new Padding(12)
        };

        var table = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            ColumnCount = 4,
            RowCount = 3
        };

        table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        table.RowStyles.Add(new RowStyle(SizeType.AutoSize));

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

        var actionRow = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            ColumnCount = 2,
            RowCount = 1,
            Margin = new Padding(0, 10, 0, 0)
        };
        actionRow.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        actionRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        actionRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

        var sessionToolbar = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            WrapContents = true,
            Margin = new Padding(0)
        };
        sessionToolbar.Controls.AddRange([_openButton, _closeButton, _validateButton, _finalizeButton]);

        var layoutToolbar = new FlowLayoutPanel
        {
            AutoSize = true,
            WrapContents = false,
            FlowDirection = FlowDirection.LeftToRight,
            Anchor = AnchorStyles.Top | AnchorStyles.Right,
            Margin = new Padding(12, 0, 0, 0)
        };
        layoutToolbar.Controls.AddRange([_sideBySideButton, _englishFocusButton, _frenchFocusButton]);

        actionRow.Controls.Add(sessionToolbar, 0, 0);
        actionRow.Controls.Add(layoutToolbar, 1, 0);

        table.Controls.Add(actionRow, 0, 2);
        table.SetColumnSpan(actionRow, 4);

        group.Controls.Add(table);
        return group;
    }

    private Control BuildBookmarkGroup()
    {
        var group = new GroupBox
        {
            Text = "Bookmark workspace",
            Dock = DockStyle.Fill,
            Padding = new Padding(12)
        };

        var grids = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 3,
            RowCount = 1,
            Margin = new Padding(0)
        };
        grids.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        grids.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 34));
        grids.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33));
        grids.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33));

        grids.Controls.Add(BuildBookmarkColumn(
            "Configured bookmarks",
            "Click a row to select it",
            _bookmarkGrid), 0, 0);

        grids.Controls.Add(BuildBookmarkColumn(
            "INFO bookmarks",
            null,
            _infoBookmarkGrid,
            _infoSummary), 1, 0);

        grids.Controls.Add(BuildBookmarkColumn(
            "Other Bookmarks",
            "Bookmarks not defined by this Template Type",
            _otherBookmarkGrid), 2, 0);

        var statusPanel = new Panel
        {
            Dock = DockStyle.Top,
            Height = 36,
            Padding = new Padding(0, 0, 0, 4)
        };

        _bookmarkSummary.Dock = DockStyle.Fill;
        _bookmarkSummary.TextAlign = ContentAlignment.MiddleLeft;
        _refreshStatusButton.Dock = DockStyle.Right;

        statusPanel.Controls.Add(_bookmarkSummary);
        statusPanel.Controls.Add(_refreshStatusButton);

        var actionPanel = new Panel
        {
            Dock = DockStyle.Top,
            Height = 42,
            Padding = new Padding(0, 2, 0, 6)
        };

        var editToolbar = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = false,
            WrapContents = false,
            Margin = new Padding(0)
        };
        editToolbar.Controls.AddRange(
            [_addBookmarkButton, _addInfoButton, _replaceRangeButton, _renameBookmarkButton, _deleteBookmarkButton]);

        var locateToolbar = new FlowLayoutPanel
        {
            Dock = DockStyle.Right,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            WrapContents = false,
            FlowDirection = FlowDirection.LeftToRight,
            Margin = new Padding(12, 0, 0, 0)
        };
        locateToolbar.Controls.AddRange([_locateEnglishButton, _locateFrenchButton, _locateBothButton]);

        actionPanel.Controls.Add(editToolbar);
        actionPanel.Controls.Add(locateToolbar);

        // Add Fill first, then top-docked rows. This keeps the grids directly
        // under the toolbars with no TableLayoutPanel sizing surprises.
        group.Controls.Add(grids);
        group.Controls.Add(actionPanel);
        group.Controls.Add(statusPanel);

        return group;
    }

    private static Control BuildBookmarkColumn(
        string title,
        string? subtitle,
        DataGridView grid,
        Label? summary = null)
    {
        var group = new GroupBox
        {
            Text = title,
            Dock = DockStyle.Fill,
            Padding = new Padding(8),
            Margin = new Padding(4)
        };

        var table = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2
        };
        table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        table.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

        Control header;

        if (summary is not null)
        {
            header = summary;
        }
        else
        {
            header = new Label
            {
                Text = subtitle ?? string.Empty,
                AutoSize = true,
                ForeColor = AppTheme.MutedText,
                Margin = new Padding(3, 2, 3, 6)
            };
        }

        table.Controls.Add(header);
        table.Controls.Add(grid);
        group.Controls.Add(table);
        return group;
    }

    private Control BuildTestModeGroup()
    {
        var group = new GroupBox
        {
            Text = "Test Mode",
            Dock = DockStyle.Top,
            AutoSize = true,
            Padding = new Padding(10)
        };

        var table = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            ColumnCount = 1,
            RowCount = 2
        };

        var modeRow = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            WrapContents = false
        };
        modeRow.Controls.AddRange([_startTestButton, _resetTestButton, _closeTestButton, _testModeLabel]);

        var operationRow = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            ColumnCount = 5,
            RowCount = 1,
            Margin = new Padding(0)
        };
        operationRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        operationRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        operationRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        operationRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        operationRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        var testValueLabel = new Label
        {
            Text = "Test value",
            AutoSize = true,
            Anchor = AnchorStyles.Left,
            Margin = new Padding(3, 3, 6, 3)
        };

        _testValueBox.Anchor = AnchorStyles.Left;
        _testValueBox.Margin = new Padding(0, 3, 6, 3);
        _testSetValueButton.Anchor = AnchorStyles.Left;
        _testDeleteContentButton.Anchor = AnchorStyles.Left;

        operationRow.Controls.Add(testValueLabel, 0, 0);
        operationRow.Controls.Add(_testValueBox, 1, 0);
        operationRow.Controls.Add(_testSetValueButton, 2, 0);
        operationRow.Controls.Add(_testDeleteContentButton, 3, 0);

        table.Controls.Add(modeRow);
        table.Controls.Add(operationRow);
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
            _bookmarkGrid.Rows[rowIndex].DefaultCellStyle.BackColor = AppTheme.PanelBack;
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
        bool selected = _wordSession.IsOpen && !string.IsNullOrWhiteSpace(_selectedBookmarkName);
        bool editingEnabled = selected && !_wordSession.IsTestMode;

        _locateEnglishButton.Enabled = editingEnabled;
        _locateFrenchButton.Enabled = editingEnabled;
        _locateBothButton.Enabled = editingEnabled;

        _addBookmarkButton.Enabled = editingEnabled && _selectedBookmarkIsConfigured;
        _replaceRangeButton.Enabled = editingEnabled;
        _renameBookmarkButton.Enabled = editingEnabled && !_selectedBookmarkIsInfo;
        _deleteBookmarkButton.Enabled = editingEnabled;

        _testSetValueButton.Enabled = _wordSession.IsTestMode && selected && !_selectedBookmarkIsInfo;
        _testDeleteContentButton.Enabled = _wordSession.IsTestMode && selected;
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

    private void StartTestMode()
    {
        try
        {
            UseWaitCursor = true;
            _wordSession.StartTestMode();
            UpdateTestModeUi();
            _status.Text = "Test Mode started from the current saved working-copy state.";
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Test Mode", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
        finally
        {
            UseWaitCursor = false;
        }
    }

    private void ResetTestMode()
    {
        try
        {
            UseWaitCursor = true;
            _wordSession.ResetTestMode();
            UpdateTestModeUi();
            _status.Text = "Test copies reset from the current working copies.";
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Reset Test", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
        finally
        {
            UseWaitCursor = false;
        }
    }

    private void CloseTestMode()
    {
        try
        {
            _wordSession.CloseTestMode();
            UpdateTestModeUi();
            _status.Text = "Test Mode closed. Working copies were not changed by test operations.";
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Close Test", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    private void TestSetValue()
    {
        if (string.IsNullOrWhiteSpace(_selectedBookmarkName))
            return;

        try
        {
            var result = _wordSession.TestSetValue(_selectedBookmarkName, _testValueBox.Text);
            _status.Text =
                $"Test Set Value applied to '{_selectedBookmarkName}' ({FormatTestSides(result.English, result.French)}).";
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Test Set Value", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    private void TestDeleteContent()
    {
        if (string.IsNullOrWhiteSpace(_selectedBookmarkName))
            return;

        try
        {
            var result = _wordSession.TestDeleteContent(_selectedBookmarkName);
            _status.Text =
                $"Test Delete Content applied to '{_selectedBookmarkName}' ({FormatTestSides(result.English, result.French)}).";
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Test Delete Content", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    private static string FormatTestSides(bool english, bool french)
    {
        if (english && french)
            return "EN + FR";

        if (english)
            return "EN only";

        return "FR only";
    }

    private void UpdateTestModeUi()
    {
        bool open = _wordSession.IsOpen;
        bool test = _wordSession.IsTestMode;

        _startTestButton.Enabled = open && !test;
        _resetTestButton.Enabled = test;
        _closeTestButton.Enabled = test;
        _testValueBox.Enabled = test;

        _testModeLabel.Text = test
            ? "TEST MODE ACTIVE — edits apply only to Test_E.docx / Test_F.docx"
            : "Test Mode inactive";

        _testModeLabel.ForeColor = test ? AppTheme.DangerAccent : AppTheme.MutedText;

        _refreshStatusButton.Enabled = open && !test;
        _validateButton.Enabled = open && !test;
        _finalizeButton.Enabled = open && !test;
        _addInfoButton.Enabled = open && !test;
        _sideBySideButton.Enabled = open && !test;
        _englishFocusButton.Enabled = open && !test;
        _frenchFocusButton.Enabled = open && !test;

        UpdateBookmarkActionButtons();
    }

    private ValidationResult BuildValidationSummary()
    {
        var type = SelectedType;
        var subtype = SelectedSubtype;

        if (type is null || subtype is null || !_wordSession.IsOpen)
        {
            return new ValidationResult(
                "Open a configured template subtype before validating.",
                HasBlockingErrors: true,
                HasWarnings: false);
        }

        var blocking = new List<string>();
        var warnings = new List<string>();

        if (!ConfigurationValidation.TryValidateTemplatePath(
                subtype.EnglishTemplatePath,
                "English",
                out string englishPathError))
        {
            blocking.Add(englishPathError);
        }

        if (!ConfigurationValidation.TryValidateTemplatePath(
                subtype.FrenchTemplatePath,
                "French",
                out string frenchPathError))
        {
            blocking.Add(frenchPathError);
        }

        if (blocking.Count == 0 &&
            string.Equals(
                Path.GetFullPath(subtype.EnglishTemplatePath),
                Path.GetFullPath(subtype.FrenchTemplatePath),
                StringComparison.OrdinalIgnoreCase))
        {
            blocking.Add("English and French source templates point to the same file.");
        }

        if (_wordSession.IsTestMode)
            blocking.Add("Close Test Mode before validating/finalizing.");

        HashSet<string> english;
        HashSet<string> french;

        try
        {
            (english, french) = _wordSession.GetBookmarkNames();
        }
        catch (Exception ex)
        {
            blocking.Add($"Could not read working-copy bookmarks: {ex.Message}");
            english = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            french = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        }

        var configured = new HashSet<string>(type.Bookmarks, StringComparer.OrdinalIgnoreCase);
        var configuredMismatches = new List<string>();
        int configuredBoth = 0;
        int configuredUnused = 0;

        foreach (string bookmark in type.Bookmarks)
        {
            bool en = english.Contains(bookmark);
            bool fr = french.Contains(bookmark);

            if (en && fr)
                configuredBoth++;
            else if (!en && !fr)
                configuredUnused++;
            else
                configuredMismatches.Add($"{bookmark} ({(en ? "EN only" : "FR only")})");
        }

        if (configuredMismatches.Count > 0)
        {
            warnings.Add(
                $"Configured bookmark mismatches ({configuredMismatches.Count}): {FormatValidationNames(configuredMismatches)}");
        }

        var infoNames = english
            .Union(french, StringComparer.OrdinalIgnoreCase)
            .Where(name => WordSession.TryGetInfoNumber(name, out _))
            .OrderBy(name =>
            {
                WordSession.TryGetInfoNumber(name, out int number);
                return number;
            })
            .ThenBy(name => name, StringComparer.OrdinalIgnoreCase)
            .ToList();

        var infoMismatches = infoNames
            .Where(name => english.Contains(name) != french.Contains(name))
            .Select(name => $"{name} ({(english.Contains(name) ? "EN only" : "FR only")})")
            .ToList();

        if (infoMismatches.Count > 0)
        {
            warnings.Add(
                $"INFO mismatches ({infoMismatches.Count}): {FormatValidationNames(infoMismatches)}");
        }

        var infoSet = new HashSet<string>(infoNames, StringComparer.OrdinalIgnoreCase);
        var otherNames = english
            .Union(french, StringComparer.OrdinalIgnoreCase)
            .Where(name => !configured.Contains(name) && !infoSet.Contains(name))
            .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (otherNames.Count > 0)
        {
            warnings.Add(
                $"Other Bookmarks ({otherNames.Count}): {FormatValidationNames(otherNames)}");
        }

        var summary = new List<string>
        {
            $"{type.Name} / {subtype.Name}",
            string.Empty,
            $"Configured bookmarks: {configuredBoth} in both, {configuredUnused} unused, {configuredMismatches.Count} mismatch",
            $"INFO bookmarks: {infoNames.Count}, {infoMismatches.Count} mismatch",
            $"Other Bookmarks: {otherNames.Count}",
            string.Empty,
            $"English source: {subtype.EnglishTemplatePath}",
            $"French source: {subtype.FrenchTemplatePath}"
        };

        if (blocking.Count > 0)
        {
            summary.Add(string.Empty);
            summary.Add("CANNOT FINALIZE:");
            summary.AddRange(blocking.Select(x => "• " + x));
        }

        if (warnings.Count > 0)
        {
            summary.Add(string.Empty);
            summary.Add("WARNINGS:");
            summary.AddRange(warnings.Select(x => "• " + x));
        }

        if (blocking.Count == 0 && warnings.Count == 0)
        {
            summary.Add(string.Empty);
            summary.Add("No validation warnings found.");
        }

        return new ValidationResult(
            string.Join(Environment.NewLine, summary),
            blocking.Count > 0,
            warnings.Count > 0);
    }

    private static string FormatValidationNames(IReadOnlyList<string> names)
    {
        const int maxShown = 12;

        if (names.Count <= maxShown)
            return string.Join(", ", names);

        return string.Join(", ", names.Take(maxShown)) + $" … +{names.Count - maxShown} more";
    }

    private void ShowValidationSummary()
    {
        var validation = BuildValidationSummary();

        MessageBox.Show(
            this,
            validation.Summary,
            "Template validation",
            MessageBoxButtons.OK,
            validation.HasBlockingErrors
                ? MessageBoxIcon.Error
                : validation.HasWarnings
                    ? MessageBoxIcon.Warning
                    : MessageBoxIcon.Information);
    }

    private void FinalizeTemplates()
    {
        var subtype = SelectedSubtype;
        if (subtype is null || !_wordSession.IsOpen)
            return;

        var validation = BuildValidationSummary();

        if (validation.HasBlockingErrors)
        {
            MessageBox.Show(
                this,
                validation.Summary,
                "Cannot finalize",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            return;
        }

        string prompt =
            validation.Summary +
            Environment.NewLine + Environment.NewLine +
            "Finalize will:" + Environment.NewLine +
            "• save both working documents through Word;" + Environment.NewLine +
            "• create timestamped backups beside both source templates;" + Environment.NewLine +
            "• overwrite the configured English and French source templates." + Environment.NewLine + Environment.NewLine +
            "Continue?";

        var answer = MessageBox.Show(
            this,
            prompt,
            "Save / Finalize templates",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning,
            MessageBoxDefaultButton.Button2);

        if (answer != DialogResult.Yes)
            return;

        try
        {
            UseWaitCursor = true;
            _validateButton.Enabled = false;
            _finalizeButton.Enabled = false;
            _status.Text = "Saving working copies, creating backups, and finalizing source templates...";

            var result = _wordSession.FinalizeToSources(
                subtype.EnglishTemplatePath,
                subtype.FrenchTemplatePath);

            _status.Text = "Finalize complete. Source templates were updated and backups were created.";

            MessageBox.Show(
                this,
                "Templates finalized successfully." +
                Environment.NewLine + Environment.NewLine +
                $"English backup: {result.EnglishBackupPath}" + Environment.NewLine +
                $"French backup: {result.FrenchBackupPath}",
                "Finalize complete",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            _status.Text = "Finalize failed. Review the error before continuing.";
            MessageBox.Show(
                this,
                ex.Message,
                "Finalize error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
        finally
        {
            UseWaitCursor = false;
            SetSessionControls(_wordSession.IsOpen);
            UpdateTestModeUi();
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
            UpdateTestModeUi();
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
                .Where(name => WordSession.TryGetInfoNumber(name, out _))
                .Select(name =>
                {
                    WordSession.TryGetInfoNumber(name, out int number);
                    return new { Name = name, Number = number };
                })
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
            string nextInfo = highestInfo == int.MaxValue ? "unavailable" : $"INFO_{highestInfo + 1}";
            _infoSummary.Text =
                $"Both: {infoBoth}, mismatch: {infoMismatch} | Next: {nextInfo}";

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
            ? AppTheme.ErrorBack
            : english
                ? AppTheme.SuccessBack
                : AppTheme.NeutralBack;
        row.DefaultCellStyle.ForeColor = AppTheme.Text;
    }

    private static void ApplyInfoStatusStyle(DataGridViewRow row, bool english, bool french)
    {
        row.DefaultCellStyle.BackColor = english != french
            ? AppTheme.ErrorBack
            : AppTheme.SuccessBack;
        row.DefaultCellStyle.ForeColor = AppTheme.Text;
    }

    private static void ApplyOtherBookmarkStyle(DataGridViewRow row, bool english, bool french)
    {
        row.DefaultCellStyle.BackColor = english != french
            ? AppTheme.ErrorBack
            : AppTheme.WarningBack;
        row.DefaultCellStyle.ForeColor = AppTheme.Text;
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
            UpdateTestModeUi();
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
        _validateButton.Enabled = open && !_wordSession.IsTestMode;
        _finalizeButton.Enabled = open && !_wordSession.IsTestMode;

        _startTestButton.Enabled = open && !_wordSession.IsTestMode;
        _resetTestButton.Enabled = _wordSession.IsTestMode;
        _closeTestButton.Enabled = _wordSession.IsTestMode;
        _testValueBox.Enabled = _wordSession.IsTestMode;

        _refreshStatusButton.Enabled = open && !_wordSession.IsTestMode;
        _addInfoButton.Enabled = open && !_wordSession.IsTestMode;
        _sideBySideButton.Enabled = open && !_wordSession.IsTestMode;
        _englishFocusButton.Enabled = open && !_wordSession.IsTestMode;
        _frenchFocusButton.Enabled = open && !_wordSession.IsTestMode;

        UpdateBookmarkActionButtons();
    }
}
