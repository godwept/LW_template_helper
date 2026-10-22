using System.Diagnostics;
using System.Reflection;
using LWTemplateHelper;

internal static class Program
{
    [STAThread]
    private static int Main()
    {
        ApplicationConfiguration.Initialize();
        int failures = 0;
        foreach (var (name, test) in new (string, Action)[]
        {
            ("Add Bookmark preserves all grid positions", AddBookmarkPreservesPositions),
            ("Add INFO reveals the last row and preserves the other grids", AddInfoRevealsLastRow),
            ("Refresh handles shortened and empty lists", RefreshHandlesShortenedLists),
            ("Subtype checklist restores new, legacy, and explicit selections", SubtypeChecklistDefaults),
            ("Subtype dialog saves subsets/empty lists and preserves its source", SubtypeDialogSaveAndCancel),
            ("MainForm creates explicit empty/subset selections and respects Cancel", CreateSubtypeSelections),
            ("MainForm edits legacy selections independently and respects Cancel", EditSubtypeSelections),
            ("Subtype switching filters rows before Word opens", SubtypeSwitchRows),
            ("Refresh ignores known unselected bookmarks and preserves Other/INFO", RefreshSubtypeStatus),
            ("Validation filters selections and retains English-only/INFO warnings", ValidateSubtypeStatus),
            ("Filtered lists preserve scrolling and handle shrink/empty refresh", FilteredScrollPositions),
            ("Subtype checklist and Save/Cancel fit at normal/minimum sizes", SubtypeDialogLayout)
        })
        {
            try { test(); Console.WriteLine($"PASS: {name}"); }
            catch (Exception ex) { failures++; Console.WriteLine($"FAIL: {name}: {ex.GetBaseException().Message}"); }
        }
        return failures == 0 ? 0 : 1;
    }

    private static MainForm CreateForm()
    {
        var form = new MainForm { ShowInTaskbar = false, Opacity = 0 };
        var type = new TemplateTypeConfig
        {
            Name = "Regression",
            Bookmarks = Enumerable.Range(1, 100).Select(n => $"BOOKMARK_{n}").ToList(),
            Subtypes = [new() { Name = "Legacy" }]
        };
        var combo = Field<ComboBox>(form, "_templateTypeCombo");
        combo.Items.Add(type);
        combo.SelectedIndex = 0;
        var session = Field<WordSession>(form, "_wordSession");
        session.Open("", "");
        foreach (int n in Enumerable.Range(1, 100))
        {
            session.AddBookmark($"INFO_{n}");
            session.AddBookmark($"OTHER_{n:000}");
        }
        form.Show();
        Application.DoEvents();
        Invoke(form, "RefreshBookmarkStatus");
        return form;
    }

    private static DataGridView[] Grids(MainForm form) =>
        [Field<DataGridView>(form, "_bookmarkGrid"), Field<DataGridView>(form, "_infoBookmarkGrid"),
         Field<DataGridView>(form, "_otherBookmarkGrid")];

    private static int[] ScrollToMiddle(MainForm form)
    {
        var grids = Grids(form);
        for (int i = 0; i < grids.Length; i++) grids[i].FirstDisplayedScrollingRowIndex = 40 + i * 10;
        return grids.Select(g => g.FirstDisplayedScrollingRowIndex).ToArray();
    }

    private static void AddBookmarkPreservesPositions()
    {
        using var form = CreateForm();
        var grids = Grids(form);
        grids[0].CurrentCell = grids[0].Rows[45].Cells[0];
        Invoke(form, "SelectBookmarkRow", grids[0], 45, true, false);
        int[] before = ScrollToMiddle(form);
        Invoke(form, "AddSelectedBookmark");
        Assert(grids[0].Rows[45].Cells["English"].Value?.ToString() == "Exists", "Bookmark status did not refresh.");
        for (int i = 0; i < grids.Length; i++)
            Assert(grids[i].FirstDisplayedScrollingRowIndex == before[i], $"Grid {i} moved from {before[i]} to {grids[i].FirstDisplayedScrollingRowIndex}.");
    }

    private static void AddInfoRevealsLastRow()
    {
        using var form = CreateForm();
        var grids = Grids(form);
        int[] before = ScrollToMiddle(form);
        Invoke(form, "AddInfoBookmark");
        Assert(grids[1].Rows[^1].Cells[0].Value?.ToString() == "INFO_101", "Expected new INFO row at the end.");
        Assert(grids[1].Rows[^1].Displayed, "New INFO row is offscreen.");
        foreach (int i in new[] { 0, 2 })
            Assert(grids[i].FirstDisplayedScrollingRowIndex == before[i], $"Grid {i} moved after Add INFO.");
    }

    private static void RefreshHandlesShortenedLists()
    {
        using var form = CreateForm();
        ScrollToMiddle(form);
        var session = Field<WordSession>(form, "_wordSession");
        session.English.Clear();
        session.French.Clear();
        session.AddBookmark("INFO_1");
        Invoke(form, "RefreshBookmarkStatus");
        var grids = Grids(form);
        Assert(grids[0].FirstDisplayedScrollingRowIndex == 40, "Configured grid moved during refresh.");
        Assert(grids[1].Rows.Count == 1 && grids[1].Rows[0].Displayed, "Short INFO list was not displayed.");
        Assert(grids[2].Rows.Count == 0, "Other grid should be empty.");
        session.English.Clear();
        session.French.Clear();
        Invoke(form, "RefreshBookmarkStatus");
        Assert(grids[1].Rows.Count == 0, "INFO grid should be empty.");
    }

    private static void SubtypeChecklistDefaults()
    {
        var type = new TemplateTypeConfig { Bookmarks = ["A", "B", "C"] };
        using (var dialog = NewSubtypeDialog(type))
            Equal(Field<CheckedListBox>(dialog, "_bookmarksList").CheckedItems.Cast<string>());
        var legacy = new TemplateSubtypeConfig { Name = "Legacy", EnglishTemplatePath = "E.docx", FrenchTemplatePath = "F.docx" };
        using (var dialog = NewSubtypeDialog(type, legacy))
        {
            Equal(Field<CheckedListBox>(dialog, "_bookmarksList").CheckedItems.Cast<string>(), "A", "B", "C");
            Assert(Field<TextBox>(dialog, "_nameBox").Text == "Legacy", "Subtype name was not restored.");
            Assert(Field<TextBox>(dialog, "_englishPath").Text == "E.docx" && Field<TextBox>(dialog, "_frenchPath").Text == "F.docx", "Subtype paths were not restored.");
        }
        using (var dialog = NewSubtypeDialog(type, new() { Bookmarks = ["c", "a", "STALE"] }))
            Equal(Field<CheckedListBox>(dialog, "_bookmarksList").CheckedItems.Cast<string>(), "A", "C");
        using (var dialog = NewSubtypeDialog(type, new() { Bookmarks = [] }))
            Equal(Field<CheckedListBox>(dialog, "_bookmarksList").CheckedItems.Cast<string>());
        using (var dialog = NewSubtypeDialog(new()))
            Assert(Field<CheckedListBox>(dialog, "_bookmarksList").Items.Count == 0, "Empty master should give an empty checklist.");
    }

    private static void SubtypeDialogSaveAndCancel()
    {
        var type = new TemplateTypeConfig { Bookmarks = ["A", "B", "C"] };
        var existing = new TemplateSubtypeConfig { Name = "Original", Bookmarks = ["B"] };
        using (var dialog = NewSubtypeDialog(type, existing))
        {
            FillSubtypeDialog(dialog, "Updated");
            SetCheckedNames(dialog, "A", "C");
            Invoke(dialog, "SaveAndClose");
            Assert(dialog.DialogResult == DialogResult.OK, "Valid dialog did not save.");
            Equal(DialogBookmarkNames(dialog), "A", "C");
            Assert(existing.Name == "Original" && existing.EnglishTemplatePath == "" && existing.FrenchTemplatePath == "", "Dialog changed source name/paths.");
            Equal(existing.Bookmarks!, "B");
        }
        using (var dialog = NewSubtypeDialog(type))
        {
            FillSubtypeDialog(dialog, "Empty");
            Invoke(dialog, "SaveAndClose");
            Assert(dialog.DialogResult == DialogResult.OK, "Empty selection should save.");
            Equal(DialogBookmarkNames(dialog));
        }
        using (var dialog = NewSubtypeDialog(type, existing))
        {
            SetCheckedNames(dialog, "A");
            dialog.DialogResult = DialogResult.Cancel;
        }
        Equal(existing.Bookmarks!, "B");
        Assert(existing.Name == "Original", "Cancel changed the source subtype.");
    }

    private static void CreateSubtypeSelections()
    {
        using var form = new MainForm();
        var type = new TemplateTypeConfig { Name = "Creation", Bookmarks = ["A", "B", "C"] };
        var combo = Field<ComboBox>(form, "_templateTypeCombo");
        combo.Items.Add(type);
        combo.SelectedIndex = 0;
        DriveSubtypeDialog(form, "CreateSubtype", dialog =>
        {
            FillSubtypeDialog(dialog, "Empty");
            Invoke(dialog, "SaveAndClose");
        });
        Assert(type.Subtypes.Single().Bookmarks is not null, "New subtype must have an explicit selection.");
        Equal(type.Subtypes.Single().Bookmarks!);
        DriveSubtypeDialog(form, "CreateSubtype", dialog =>
        {
            FillSubtypeDialog(dialog, "Subset");
            SetCheckedNames(dialog, "A", "C");
            Invoke(dialog, "SaveAndClose");
        });
        Equal(type.Subtypes.Single(s => s.Name == "Subset").Bookmarks!, "A", "C");
        var store = Field<TemplateConfigStore>(form, "_configStore");
        int saves = store.SaveCount;
        DriveSubtypeDialog(form, "CreateSubtype", dialog => dialog.DialogResult = DialogResult.Cancel);
        Assert(type.Subtypes.Count == 2 && store.SaveCount == saves, "Cancel must not create/save a subtype.");
    }

    private static void EditSubtypeSelections()
    {
        using var form = new MainForm();
        var legacy = new TemplateSubtypeConfig { Name = "Legacy" };
        var other = new TemplateSubtypeConfig { Name = "Other", Bookmarks = ["C"] };
        var type = new TemplateTypeConfig { Name = "Editing", Bookmarks = ["A", "B", "C"], Subtypes = [legacy, other] };
        var combo = Field<ComboBox>(form, "_templateTypeCombo");
        combo.Items.Add(type);
        combo.SelectedIndex = 0;
        DriveSubtypeDialog(form, "EditSubtype", dialog =>
        {
            FillSubtypeDialog(dialog, "Legacy");
            SetCheckedNames(dialog, "B");
            Invoke(dialog, "SaveAndClose");
        });
        Assert(legacy.Bookmarks is not null, "Editing must materialize a legacy selection.");
        Equal(legacy.Bookmarks!, "B");
        Equal(other.Bookmarks!, "C");
        DriveSubtypeDialog(form, "EditSubtype", dialog =>
        {
            SetCheckedNames(dialog);
            Invoke(dialog, "SaveAndClose");
        });
        Equal(legacy.Bookmarks!);
        var store = Field<TemplateConfigStore>(form, "_configStore");
        int saves = store.SaveCount;
        DriveSubtypeDialog(form, "EditSubtype", dialog =>
        {
            Field<TextBox>(dialog, "_nameBox").Text = "Cancelled";
            SetCheckedNames(dialog, "A");
            dialog.DialogResult = DialogResult.Cancel;
        });
        Assert(legacy.Name == "Legacy" && store.SaveCount == saves, "Cancel must not update/save the subtype.");
        Equal(legacy.Bookmarks!);
        Equal(other.Bookmarks!, "C");
    }

    private static void SubtypeSwitchRows()
    {
        using var form = new MainForm();
        var subset = new TemplateSubtypeConfig { Name = "Subset", Bookmarks = ["C", "A"] };
        var empty = new TemplateSubtypeConfig { Name = "Empty", Bookmarks = [] };
        var legacy = new TemplateSubtypeConfig { Name = "Legacy" };
        var type = new TemplateTypeConfig { Name = "Switching", Bookmarks = ["A", "B", "C"], Subtypes = [subset, empty, legacy] };
        var combo = Field<ComboBox>(form, "_templateTypeCombo");
        combo.Items.Add(type);
        combo.SelectedIndex = 0;
        var subtypeCombo = Field<ComboBox>(form, "_subtypeCombo");
        var grid = Grids(form)[0];
        subtypeCombo.SelectedItem = subset;
        Equal(GridNames(grid), "A", "C");
        Invoke(form, "SelectBookmarkRow", grid, 0, true, false);
        subtypeCombo.SelectedItem = empty;
        Equal(GridNames(grid));
        Assert(Field<string?>(form, "_selectedBookmarkName") is null, "Switching retained a stale bookmark selection.");
        subtypeCombo.SelectedItem = legacy;
        Equal(GridNames(grid), "A", "B", "C");
        var noSubtypes = new TemplateTypeConfig { Name = "No subtypes", Bookmarks = ["D"] };
        combo.Items.Add(noSubtypes);
        combo.SelectedItem = noSubtypes;
        Equal(GridNames(grid));
        combo.SelectedIndex = -1;
        Equal(GridNames(grid));
    }

    private static MainForm CreateStatusForm()
    {
        var form = new MainForm();
        var (englishPath, frenchPath) = FixturePaths();
        var type = new TemplateTypeConfig
        {
            Name = "Status", Bookmarks = ["A", "B", "C", "salutation"],
            Subtypes = [new() { Name = "Selected", Bookmarks = ["A", "C"], EnglishTemplatePath = englishPath, FrenchTemplatePath = frenchPath }]
        };
        var combo = Field<ComboBox>(form, "_templateTypeCombo");
        combo.Items.Add(type);
        combo.SelectedIndex = 0;
        var session = Field<WordSession>(form, "_wordSession");
        session.Open("", "");
        session.English.UnionWith(["A", "B", "salutation", "INFO_1", "UNKNOWN"]);
        session.French.UnionWith(["A", "INFO_1"]);
        return form;
    }

    private static void RefreshSubtypeStatus()
    {
        using var form = CreateStatusForm();
        Invoke(form, "RefreshBookmarkStatus");
        var grids = Grids(form);
        Equal(GridNames(grids[0]), "A", "C");
        Equal(GridNames(grids[1]), "INFO_1");
        Equal(GridNames(grids[2]), "UNKNOWN");
        Assert(Field<Label>(form, "_bookmarkSummary").Text.Contains("1 both, 0 mismatch, 1 unused"), "Counts include unselected bookmarks.");
        var subtype = (TemplateSubtypeConfig)Field<ComboBox>(form, "_subtypeCombo").SelectedItem!;
        subtype.Bookmarks = [];
        Invoke(form, "RefreshBookmarkStatus");
        Equal(GridNames(grids[0]));
        Equal(GridNames(grids[1]), "INFO_1");
        Equal(GridNames(grids[2]), "UNKNOWN");
        Assert(Field<Label>(form, "_bookmarkSummary").Text.Contains("0 both, 0 mismatch, 0 unused"), "Empty selection has configured counts.");
    }

    private static void ValidateSubtypeStatus()
    {
        using var form = CreateStatusForm();
        var subtype = (TemplateSubtypeConfig)Field<ComboBox>(form, "_subtypeCombo").SelectedItem!;
        subtype.Bookmarks = ["A", "C", "salutation"];
        var result = (ValidationResult)Invoke(form, "BuildValidationSummary")!;
        Assert(!result.HasBlockingErrors, "Valid fixture paths should not block validation.");
        Assert(result.Summary.Contains("1 in both, 1 unused, 1 mismatch"), "Validation includes unselected bookmarks.");
        Assert(result.Summary.Contains("salutation (EN only)"), "English-only salutation should keep its warning.");
        Assert(!result.Summary.Contains("B (EN only)"), "Known unselected B was reported as a mismatch.");
        Assert(result.Summary.Contains("Other Bookmarks (1): UNKNOWN"), "Only the unknown bookmark should be Other.");
        subtype.Bookmarks = [];
        result = (ValidationResult)Invoke(form, "BuildValidationSummary")!;
        Assert(result.Summary.Contains("0 in both, 0 unused, 0 mismatch"), "Empty selection has configured validation counts.");
        Assert(!result.Summary.Contains("Configured bookmark mismatches"), "Empty selection creates mismatch warnings.");
        Assert(result.Summary.Contains("Other Bookmarks (1): UNKNOWN"), "Empty selection altered Other reporting.");
        var session = Field<WordSession>(form, "_wordSession");
        session.French.Remove("INFO_1");
        result = (ValidationResult)Invoke(form, "BuildValidationSummary")!;
        Assert(result.Summary.Contains("INFO_1 (EN only)"), "INFO mismatch reporting must remain independent.");
        session.English.Remove("UNKNOWN");
        session.English.Remove("INFO_1");
        result = (ValidationResult)Invoke(form, "BuildValidationSummary")!;
        Assert(!result.HasWarnings && !result.HasBlockingErrors, "Known unselected bookmarks must produce no warnings.");
    }

    private static void FilteredScrollPositions()
    {
        using var form = CreateForm();
        var type = (TemplateTypeConfig)Field<ComboBox>(form, "_templateTypeCombo").SelectedItem!;
        var subtype = type.Subtypes.Single();
        subtype.Bookmarks = type.Bookmarks.Skip(10).ToList();
        Invoke(form, "RefreshBookmarkStatus");
        var grids = Grids(form);
        Assert(grids[0].Rows.Count == 90, "Expected an explicit filtered list.");
        grids[0].CurrentCell = grids[0].Rows[45].Cells[0];
        Invoke(form, "SelectBookmarkRow", grids[0], 45, true, false);
        int[] before = ScrollToMiddle(form);
        Invoke(form, "AddSelectedBookmark");
        Assert(grids[0].Rows[45].Cells["English"].Value?.ToString() == "Exists", "Filtered bookmark was not added.");
        for (int i = 0; i < grids.Length; i++)
            Assert(grids[i].FirstDisplayedScrollingRowIndex == before[i], $"Filtered grid {i} moved after Add Bookmark.");
        Invoke(form, "AddInfoBookmark");
        Assert(grids[1].Rows[^1].Displayed, "New INFO row should be visible with filtered configured rows.");
        foreach (int i in new[] { 0, 2 })
            Assert(grids[i].FirstDisplayedScrollingRowIndex == before[i], $"Grid {i} moved after Add INFO.");
        subtype.Bookmarks = type.Bookmarks.Take(3).ToList();
        Invoke(form, "RefreshBookmarkStatus");
        // WinForms can move to row zero when every remaining row fits in the viewport.
        Assert(grids[0].Rows.Count == 3 && grids[0].FirstDisplayedScrollingRowIndex is >= 0 and < 3, "Scroll row is outside the shortened list.");
        Assert(grids[0].Rows.Cast<DataGridViewRow>().All(row => row.Displayed), "Shortened list should remain visible.");
        subtype.Bookmarks = [];
        Invoke(form, "RefreshBookmarkStatus");
        Equal(GridNames(grids[0]));
    }

    private static void SubtypeDialogLayout()
    {
        using var dialog = NewSubtypeDialog(new() { Bookmarks = ["A", "B", "C"] });
        dialog.Opacity = 0;
        dialog.ShowInTaskbar = false;
        dialog.Show();
        foreach (var size in new[] { dialog.Size, dialog.MinimumSize })
        {
            dialog.Size = size;
            dialog.PerformLayout();
            var checklist = Field<CheckedListBox>(dialog, "_bookmarksList");
            Assert(checklist.Height >= 100, "Checklist has too little usable height.");
            foreach (var control in new Control[] { checklist, Field<TextBox>(dialog, "_nameBox"),
                Field<TextBox>(dialog, "_englishPath"), Field<TextBox>(dialog, "_frenchPath"),
                (Button)dialog.AcceptButton!, (Button)dialog.CancelButton! })
            {
                var bounds = new Rectangle(dialog.PointToClient(control.PointToScreen(Point.Empty)), control.Size);
                Assert(dialog.ClientRectangle.Contains(bounds), $"{control.GetType().Name} is clipped at {size}.");
            }
        }
    }

    private static IEnumerable<string> GridNames(DataGridView grid) =>
        grid.Rows.Cast<DataGridViewRow>().Select(row => row.Cells["Bookmark"].Value!.ToString()!);

    private static void DriveSubtypeDialog(MainForm form, string action, Action<TemplateSubtypeDialog> edit)
    {
        Exception? callbackError = null;
        bool handled = false;
        var elapsed = Stopwatch.StartNew();
        using var timer = new System.Windows.Forms.Timer { Interval = 20 };
        timer.Tick += (_, _) =>
        {
            var dialog = Application.OpenForms.OfType<TemplateSubtypeDialog>().FirstOrDefault();
            if (elapsed.Elapsed > TimeSpan.FromSeconds(5))
            {
                callbackError = new TimeoutException("Subtype dialog did not complete.");
                if (dialog is not null) dialog.DialogResult = DialogResult.Cancel;
                timer.Stop();
                return;
            }
            if (dialog is null || handled) return;
            handled = true;
            try
            {
                dialog.ShowInTaskbar = false;
                dialog.Opacity = 0;
                edit(dialog);
            }
            catch (Exception ex)
            {
                callbackError = ex;
                dialog.DialogResult = DialogResult.Cancel;
            }
        };
        timer.Start();
        try { Invoke(form, action); }
        finally { timer.Stop(); }
        if (callbackError is not null) throw callbackError;
        Assert(handled, "MainForm did not open a subtype dialog.");
    }

    private static List<string> DialogBookmarkNames(TemplateSubtypeDialog dialog) =>
        (List<string>)(typeof(TemplateSubtypeDialog).GetProperty("BookmarkNames")?.GetValue(dialog)
            ?? throw new InvalidOperationException("Dialog does not return bookmark selections."));

    private static void FillSubtypeDialog(TemplateSubtypeDialog dialog, string name)
    {
        var (english, french) = FixturePaths();
        Field<TextBox>(dialog, "_nameBox").Text = name;
        Field<TextBox>(dialog, "_englishPath").Text = english;
        Field<TextBox>(dialog, "_frenchPath").Text = french;
    }

    private static (string English, string French) FixturePaths()
    {
        string directory = Path.Combine(AppContext.BaseDirectory, "fixtures");
        Directory.CreateDirectory(directory);
        string english = Path.Combine(directory, "E.docx");
        string french = Path.Combine(directory, "F.docx");
        File.WriteAllText(english, "Dialog validation fixture; never opened in Word.");
        File.WriteAllText(french, "Dialog validation fixture; never opened in Word.");
        return (english, french);
    }

    private static void SetCheckedNames(TemplateSubtypeDialog dialog, params string[] names)
    {
        var list = Field<CheckedListBox>(dialog, "_bookmarksList");
        for (int i = 0; i < list.Items.Count; i++)
            list.SetItemChecked(i, names.Contains(list.Items[i].ToString()));
    }

    private static TemplateSubtypeDialog NewSubtypeDialog(TemplateTypeConfig type, TemplateSubtypeConfig? subtype = null) =>
        (TemplateSubtypeDialog)Activator.CreateInstance(typeof(TemplateSubtypeDialog), type, subtype)!;

    private static void Equal(IEnumerable<string> actual, params string[] expected) =>
        Assert(actual.SequenceEqual(expected), $"Expected [{string.Join(", ", expected)}], got [{string.Join(", ", actual)}].");

    private static T Field<T>(object owner, string name) =>
        (T)owner.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(owner)!;

    private static object? Invoke(object owner, string name, params object[] args) =>
        owner.GetType().GetMethod(name, BindingFlags.Instance | BindingFlags.NonPublic)!.Invoke(owner, args);

    private static void Assert(bool condition, string message)
    {
        if (!condition) throw new InvalidOperationException(message);
    }
}
