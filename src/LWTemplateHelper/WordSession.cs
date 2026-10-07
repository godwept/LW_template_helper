using System.ComponentModel;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace LWTemplateHelper;

internal sealed class WordSession : IDisposable
{
    private Word.Application? _word;
    private Word.Document? _englishDocument;
    private Word.Document? _frenchDocument;
    private Word.Document? _englishTestDocument;
    private Word.Document? _frenchTestDocument;

    public bool IsOpen => _word is not null;
    public bool IsTestMode => _englishTestDocument is not null && _frenchTestDocument is not null;
    public string? SessionDirectory { get; private set; }
    public string? EnglishWorkingPath { get; private set; }
    public string? FrenchWorkingPath { get; private set; }
    public string? EnglishTestPath { get; private set; }
    public string? FrenchTestPath { get; private set; }

    public void Open(string englishSource, string frenchSource)
    {
        if (IsOpen)
            throw new InvalidOperationException("A Word session is already open.");

        ValidateSource(englishSource);
        ValidateSource(frenchSource);

        SessionDirectory = Path.Combine(
            Path.GetTempPath(),
            "LWTemplateHelper",
            DateTime.Now.ToString("yyyyMMdd-HHmmss") + "-" + Guid.NewGuid().ToString("N")[..8]);

        Directory.CreateDirectory(SessionDirectory);

        EnglishWorkingPath = Path.Combine(SessionDirectory, "Working_E.docx");
        FrenchWorkingPath = Path.Combine(SessionDirectory, "Working_F.docx");

        File.Copy(englishSource, EnglishWorkingPath, overwrite: false);
        File.Copy(frenchSource, FrenchWorkingPath, overwrite: false);

        try
        {
            _word = new Word.Application
            {
                Visible = true,
                DisplayAlerts = Word.WdAlertLevel.wdAlertsAll
            };

            _englishDocument = _word.Documents.Open(
                FileName: EnglishWorkingPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: true);

            _frenchDocument = _word.Documents.Open(
                FileName: FrenchWorkingPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: true);

            ArrangeSideBySide();
        }
        catch
        {
            Close();
            throw;
        }
    }

    public (HashSet<string> English, HashSet<string> French) GetBookmarkNames()
    {
        EnsureOpen();

        return (
            ReadBookmarkNames(_englishDocument!),
            ReadBookmarkNames(_frenchDocument!));
    }

    public void LocateEnglish(string bookmarkName) =>
        LocateBookmark(_englishDocument, bookmarkName, "English");

    public void LocateFrench(string bookmarkName) =>
        LocateBookmark(_frenchDocument, bookmarkName, "French");

    public void LocateBoth(string bookmarkName)
    {
        EnsureOpen();
        LocateBookmark(_englishDocument, bookmarkName, "English");
        LocateBookmark(_frenchDocument, bookmarkName, "French");
    }

    public void StartTestMode()
    {
        EnsureOpen();

        if (IsTestMode)
            throw new InvalidOperationException("Test Mode is already open.");

        if (SessionDirectory is null || EnglishWorkingPath is null || FrenchWorkingPath is null)
            throw new InvalidOperationException("The working-copy session paths are unavailable.");

        _englishDocument!.Save();
        _frenchDocument!.Save();

        var englishWorkingBounds = GetNativeBounds(_englishDocument);
        var frenchWorkingBounds = GetNativeBounds(_frenchDocument);

        EnglishTestPath = Path.Combine(SessionDirectory, "Test_E.docx");
        FrenchTestPath = Path.Combine(SessionDirectory, "Test_F.docx");

        DeleteFileIfExists(EnglishTestPath);
        DeleteFileIfExists(FrenchTestPath);

        File.Copy(EnglishWorkingPath, EnglishTestPath, overwrite: true);
        File.Copy(FrenchWorkingPath, FrenchTestPath, overwrite: true);

        try
        {
            _englishTestDocument = _word!.Documents.Open(
                FileName: EnglishTestPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: true);

            _frenchTestDocument = _word.Documents.Open(
                FileName: FrenchTestPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: true);

            PositionDocumentOver(_englishTestDocument, englishWorkingBounds);
            PositionDocumentOver(_frenchTestDocument, frenchWorkingBounds);
        }
        catch
        {
            CloseTestMode();
            throw;
        }
    }

    public void ResetTestMode()
    {
        EnsureOpen();

        if (!IsTestMode)
            throw new InvalidOperationException("Test Mode is not open.");

        CloseTestMode();
        StartTestMode();
    }

    public void CloseTestMode()
    {
        CloseDocument(ref _frenchTestDocument);
        CloseDocument(ref _englishTestDocument);

        DeleteFileIfExists(FrenchTestPath);
        DeleteFileIfExists(EnglishTestPath);

        FrenchTestPath = null;
        EnglishTestPath = null;
    }

    public (bool English, bool French) TestSetValue(string bookmarkName, string value)
    {
        EnsureTestMode();

        bool english = SetBookmarkValueIfExists(_englishTestDocument!, bookmarkName, value);
        bool french = SetBookmarkValueIfExists(_frenchTestDocument!, bookmarkName, value);

        if (!english && !french)
            throw new InvalidOperationException($"Bookmark '{bookmarkName}' does not exist in either test document.");

        return (english, french);
    }

    public (bool English, bool French) TestDeleteContent(string bookmarkName)
    {
        EnsureTestMode();

        bool english = DeleteBookmarkContentIfExists(_englishTestDocument!, bookmarkName);
        bool french = DeleteBookmarkContentIfExists(_frenchTestDocument!, bookmarkName);

        if (!english && !french)
            throw new InvalidOperationException($"Bookmark '{bookmarkName}' does not exist in either test document.");

        return (english, french);
    }

    public string AddNextInfoBookmark()
    {
        EnsureOpen();

        var english = ReadBookmarkNames(_englishDocument!);
        var french = ReadBookmarkNames(_frenchDocument!);

        int highest = 0;

        foreach (string name in english.Concat(french))
        {
            if (TryGetInfoNumber(name, out int number) && number > highest)
                highest = number;
        }

        if (highest == int.MaxValue)
            throw new InvalidOperationException("Cannot create another INFO bookmark because the INFO number is too large.");

        string bookmarkName = $"INFO_{highest + 1}";
        AddBookmark(bookmarkName);
        return bookmarkName;
    }

    public void AddBookmark(string bookmarkName)
    {
        EnsureOpen();
        ValidateBookmarkName(bookmarkName);

        Word.Range? englishRange = null;
        Word.Range? frenchRange = null;

        try
        {
            if (BookmarkExists(_englishDocument!, bookmarkName) ||
                BookmarkExists(_frenchDocument!, bookmarkName))
            {
                throw new InvalidOperationException(
                    $"Bookmark '{bookmarkName}' already exists in at least one working copy. Use Replace Range if you intend to move or repair it.");
            }

            englishRange = GetCurrentSelectionRangeDuplicate(_englishDocument!, "English");
            frenchRange = GetCurrentSelectionRangeDuplicate(_frenchDocument!, "French");

            AddBookmarkMarker(_englishDocument!, bookmarkName, englishRange);

            try
            {
                AddBookmarkMarker(_frenchDocument!, bookmarkName, frenchRange);
            }
            catch
            {
                DeleteBookmarkMarkerIfExists(_englishDocument!, bookmarkName);
                throw;
            }
        }
        finally
        {
            ReleaseCom(englishRange);
            ReleaseCom(frenchRange);
        }
    }

    public void ReplaceBookmarkRange(string bookmarkName)
    {
        EnsureOpen();
        ValidateBookmarkName(bookmarkName);

        Word.Range? englishRange = null;
        Word.Range? frenchRange = null;
        Word.Range? oldEnglishRange = null;
        Word.Range? oldFrenchRange = null;

        try
        {
            oldEnglishRange = GetBookmarkRangeDuplicate(_englishDocument!, bookmarkName);
            oldFrenchRange = GetBookmarkRangeDuplicate(_frenchDocument!, bookmarkName);

            if (oldEnglishRange is null && oldFrenchRange is null)
            {
                throw new InvalidOperationException(
                    $"Bookmark '{bookmarkName}' does not exist in either working copy. Use Add Bookmark instead.");
            }

            englishRange = GetCurrentSelectionRangeDuplicate(_englishDocument!, "English");
            frenchRange = GetCurrentSelectionRangeDuplicate(_frenchDocument!, "French");

            DeleteBookmarkMarkerIfExists(_englishDocument!, bookmarkName);
            DeleteBookmarkMarkerIfExists(_frenchDocument!, bookmarkName);

            try
            {
                AddBookmarkMarker(_englishDocument!, bookmarkName, englishRange);
                AddBookmarkMarker(_frenchDocument!, bookmarkName, frenchRange);
            }
            catch
            {
                DeleteBookmarkMarkerIfExists(_englishDocument!, bookmarkName);
                DeleteBookmarkMarkerIfExists(_frenchDocument!, bookmarkName);

                if (oldEnglishRange is not null)
                    AddBookmarkMarker(_englishDocument!, bookmarkName, oldEnglishRange);

                if (oldFrenchRange is not null)
                    AddBookmarkMarker(_frenchDocument!, bookmarkName, oldFrenchRange);

                throw;
            }
        }
        finally
        {
            ReleaseCom(englishRange);
            ReleaseCom(frenchRange);
            ReleaseCom(oldEnglishRange);
            ReleaseCom(oldFrenchRange);
        }
    }

    public void RenameBookmark(string oldName, string newName)
    {
        EnsureOpen();
        ValidateBookmarkName(newName);

        if (string.Equals(oldName, newName, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("Enter a different bookmark name.");

        Word.Range? oldEnglishRange = null;
        Word.Range? oldFrenchRange = null;

        try
        {
            oldEnglishRange = GetBookmarkRangeDuplicate(_englishDocument!, oldName);
            oldFrenchRange = GetBookmarkRangeDuplicate(_frenchDocument!, oldName);

            if (oldEnglishRange is null && oldFrenchRange is null)
                throw new InvalidOperationException($"Bookmark '{oldName}' does not exist in either working copy.");

            if (BookmarkExists(_englishDocument!, newName) ||
                BookmarkExists(_frenchDocument!, newName))
            {
                throw new InvalidOperationException(
                    $"Bookmark '{newName}' already exists in at least one working copy.");
            }

            DeleteBookmarkMarkerIfExists(_englishDocument!, oldName);
            DeleteBookmarkMarkerIfExists(_frenchDocument!, oldName);

            try
            {
                if (oldEnglishRange is not null)
                    AddBookmarkMarker(_englishDocument!, newName, oldEnglishRange);

                if (oldFrenchRange is not null)
                    AddBookmarkMarker(_frenchDocument!, newName, oldFrenchRange);
            }
            catch
            {
                DeleteBookmarkMarkerIfExists(_englishDocument!, newName);
                DeleteBookmarkMarkerIfExists(_frenchDocument!, newName);

                if (oldEnglishRange is not null)
                    AddBookmarkMarker(_englishDocument!, oldName, oldEnglishRange);

                if (oldFrenchRange is not null)
                    AddBookmarkMarker(_frenchDocument!, oldName, oldFrenchRange);

                throw;
            }
        }
        finally
        {
            ReleaseCom(oldEnglishRange);
            ReleaseCom(oldFrenchRange);
        }
    }

    public void DeleteBookmarkMarkers(string bookmarkName)
    {
        EnsureOpen();

        Word.Range? oldEnglishRange = null;
        Word.Range? oldFrenchRange = null;

        try
        {
            oldEnglishRange = GetBookmarkRangeDuplicate(_englishDocument!, bookmarkName);
            oldFrenchRange = GetBookmarkRangeDuplicate(_frenchDocument!, bookmarkName);

            if (oldEnglishRange is null && oldFrenchRange is null)
                throw new InvalidOperationException($"Bookmark '{bookmarkName}' does not exist in either working copy.");

            try
            {
                DeleteBookmarkMarkerIfExists(_englishDocument!, bookmarkName);
                DeleteBookmarkMarkerIfExists(_frenchDocument!, bookmarkName);
            }
            catch
            {
                if (oldEnglishRange is not null && !BookmarkExists(_englishDocument!, bookmarkName))
                    AddBookmarkMarker(_englishDocument!, bookmarkName, oldEnglishRange);

                if (oldFrenchRange is not null && !BookmarkExists(_frenchDocument!, bookmarkName))
                    AddBookmarkMarker(_frenchDocument!, bookmarkName, oldFrenchRange);

                throw;
            }
        }
        finally
        {
            ReleaseCom(oldEnglishRange);
            ReleaseCom(oldFrenchRange);
        }
    }

    public void ArrangeSideBySide()
    {
        EnsureOpen();
        ArrangeDocumentsSideBySide(_englishDocument, _frenchDocument);
    }

    private static void ArrangeDocumentsSideBySide(
        Word.Document? englishDocument,
        Word.Document? frenchDocument)
    {
        var manager = Form.ActiveForm ?? throw new InvalidOperationException("Manager window is unavailable.");
        var area = Screen.FromControl(manager).WorkingArea;
        var englishWindow = GetWindow(englishDocument);
        var frenchWindow = GetWindow(frenchDocument);

        try
        {
            englishWindow.WindowState = Word.WdWindowState.wdWindowStateNormal;
            frenchWindow.WindowState = Word.WdWindowState.wdWindowStateNormal;

            int leftWidth = area.Width / 2;
            int rightWidth = area.Width - leftWidth;

            PositionNative(englishWindow, area.Left, area.Top, leftWidth, area.Height);
            PositionNative(frenchWindow, area.Left + leftWidth, area.Top, rightWidth, area.Height);

            englishWindow.Activate();
        }
        finally
        {
            ReleaseCom(englishWindow);
            ReleaseCom(frenchWindow);
        }
    }

    public void FocusEnglish() => Focus(_englishDocument);

    public void FocusFrench() => Focus(_frenchDocument);

    private static bool SetBookmarkValueIfExists(
        Word.Document document,
        string bookmarkName,
        string value)
    {
        Word.Bookmark? bookmark = null;
        Word.Range? range = null;

        try
        {
            bookmark = FindBookmark(document, bookmarkName);

            if (bookmark is null)
                return false;

            range = bookmark.Range;

            // Mirrors the current Letter Wizard VBA helper:
            // assigning Range.Text replaces the content and consumes the bookmark.
            range.Text = value;
            return true;
        }
        finally
        {
            ReleaseCom(range);
            ReleaseCom(bookmark);
        }
    }

    private static bool DeleteBookmarkContentIfExists(
        Word.Document document,
        string bookmarkName)
    {
        Word.Bookmark? bookmark = null;
        Word.Range? range = null;

        try
        {
            bookmark = FindBookmark(document, bookmarkName);

            if (bookmark is null)
                return false;

            range = bookmark.Range;
            range.Delete();
            return true;
        }
        finally
        {
            ReleaseCom(range);
            ReleaseCom(bookmark);
        }
    }

    private void EnsureTestMode()
    {
        if (!IsTestMode || _englishTestDocument is null || _frenchTestDocument is null)
            throw new InvalidOperationException("Start Test Mode first.");
    }

    private static void DeleteFileIfExists(string? path)
    {
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            return;

        try
        {
            File.Delete(path);
        }
        catch (IOException)
        {
            // A stale test file is disposable. If it is still locked, the
            // subsequent copy/open operation will surface a useful error.
        }
        catch (UnauthorizedAccessException)
        {
            // Same principle as above; let the next operation report the failure.
        }
    }

    internal static bool TryGetInfoNumber(string bookmarkName, out int number)
    {
        number = 0;

        if (!bookmarkName.StartsWith("INFO_", StringComparison.OrdinalIgnoreCase))
            return false;

        return int.TryParse(bookmarkName.AsSpan(5), out number) && number >= 1;
    }

    private static void ValidateBookmarkName(string bookmarkName)
    {
        if (!ConfigurationValidation.TryValidateBookmarkName(bookmarkName, out string error))
            throw new ArgumentException(error, nameof(bookmarkName));
    }

    private static Word.Range GetCurrentSelectionRangeDuplicate(
        Word.Document document,
        string language)
    {
        Word.Window? window = null;
        Word.Selection? selection = null;
        Word.Range? range = null;

        try
        {
            window = GetWindow(document);
            selection = window.Selection;
            range = selection.Range;

            if (range.Start >= range.End)
            {
                throw new InvalidOperationException(
                    $"Select the text to bookmark in the {language} working copy first.");
            }

            return range.Duplicate;
        }
        finally
        {
            ReleaseCom(range);
            ReleaseCom(selection);
            ReleaseCom(window);
        }
    }

    private static bool BookmarkExists(Word.Document document, string bookmarkName)
    {
        Word.Bookmark? bookmark = null;

        try
        {
            bookmark = FindBookmark(document, bookmarkName);
            return bookmark is not null;
        }
        finally
        {
            ReleaseCom(bookmark);
        }
    }

    private static Word.Range? GetBookmarkRangeDuplicate(Word.Document document, string bookmarkName)
    {
        Word.Bookmark? bookmark = null;
        Word.Range? range = null;

        try
        {
            bookmark = FindBookmark(document, bookmarkName);

            if (bookmark is null)
                return null;

            range = bookmark.Range;
            return range.Duplicate;
        }
        finally
        {
            ReleaseCom(range);
            ReleaseCom(bookmark);
        }
    }

    private static void AddBookmarkMarker(Word.Document document, string bookmarkName, Word.Range range)
    {
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? addedBookmark = null;

        try
        {
            bookmarks = document.Bookmarks;
            object rangeObject = range;
            addedBookmark = bookmarks.Add(bookmarkName, ref rangeObject);
        }
        finally
        {
            ReleaseCom(addedBookmark);
            ReleaseCom(bookmarks);
        }
    }

    private static bool DeleteBookmarkMarkerIfExists(Word.Document document, string bookmarkName)
    {
        Word.Bookmark? bookmark = null;

        try
        {
            bookmark = FindBookmark(document, bookmarkName);

            if (bookmark is null)
                return false;

            bookmark.Delete();
            return true;
        }
        finally
        {
            ReleaseCom(bookmark);
        }
    }

    private static HashSet<string> ReadBookmarkNames(Word.Document document)
    {
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        Word.Bookmarks? bookmarks = null;

        try
        {
            bookmarks = document.Bookmarks;

            for (int i = 1; i <= bookmarks.Count; i++)
            {
                object index = i;
                Word.Bookmark? bookmark = null;

                try
                {
                    bookmark = bookmarks.get_Item(ref index);
                    names.Add(bookmark.Name);
                }
                finally
                {
                    ReleaseCom(bookmark);
                }
            }

            return names;
        }
        finally
        {
            ReleaseCom(bookmarks);
        }
    }

    private static void LocateBookmark(Word.Document? document, string bookmarkName, string language)
    {
        if (document is null)
            throw new InvalidOperationException("The Word session is not open.");

        Word.Bookmark? bookmark = null;
        Word.Range? range = null;
        Word.Window? window = null;

        try
        {
            bookmark = FindBookmark(document, bookmarkName);

            if (bookmark is null)
                throw new InvalidOperationException(
                    $"Bookmark '{bookmarkName}' does not exist in the {language} working copy.");

            range = bookmark.Range;
            window = GetWindow(document);
            window.Activate();
            range.Select();
        }
        finally
        {
            ReleaseCom(range);
            ReleaseCom(bookmark);
            ReleaseCom(window);
        }
    }

    private static Word.Bookmark? FindBookmark(Word.Document document, string bookmarkName)
    {
        Word.Bookmarks? bookmarks = null;

        try
        {
            bookmarks = document.Bookmarks;

            for (int i = 1; i <= bookmarks.Count; i++)
            {
                object index = i;
                Word.Bookmark? bookmark = bookmarks.get_Item(ref index);

                if (string.Equals(bookmark.Name, bookmarkName, StringComparison.OrdinalIgnoreCase))
                    return bookmark;

                ReleaseCom(bookmark);
            }

            return null;
        }
        finally
        {
            ReleaseCom(bookmarks);
        }
    }

    private static (int Left, int Top, int Width, int Height) GetNativeBounds(Word.Document? document)
    {
        var window = GetWindow(document);

        try
        {
            var hwnd = new IntPtr(window.Hwnd);

            if (!GetWindowRect(hwnd, out NativeRect rect))
            {
                throw new Win32Exception(
                    Marshal.GetLastWin32Error(),
                    "Windows could not read the Word window position.");
            }

            return (
                rect.Left,
                rect.Top,
                rect.Right - rect.Left,
                rect.Bottom - rect.Top);
        }
        finally
        {
            ReleaseCom(window);
        }
    }

    private static void PositionDocumentOver(
        Word.Document? document,
        (int Left, int Top, int Width, int Height) bounds)
    {
        var window = GetWindow(document);

        try
        {
            window.WindowState = Word.WdWindowState.wdWindowStateNormal;
            PositionNative(window, bounds.Left, bounds.Top, bounds.Width, bounds.Height);

            var hwnd = new IntPtr(window.Hwnd);
            BringWindowToTop(hwnd);
        }
        finally
        {
            ReleaseCom(window);
        }
    }

    private static void PositionNative(Word.Window window, int left, int top, int width, int height)
    {
        var hwnd = new IntPtr(window.Hwnd);

        if (!MoveWindow(hwnd, left, top, width, height, true))
            throw new Win32Exception(Marshal.GetLastWin32Error(), "Windows could not position the Word window.");
    }

    private static void Focus(Word.Document? document)
    {
        if (document is null)
            throw new InvalidOperationException("The Word session is not open.");

        var window = GetWindow(document);
        try
        {
            window.WindowState = Word.WdWindowState.wdWindowStateMaximize;
            window.Activate();
        }
        finally
        {
            ReleaseCom(window);
        }
    }

    private static Word.Window GetWindow(Word.Document? document)
    {
        if (document is null)
            throw new InvalidOperationException("The Word session is not open.");

        return document.ActiveWindow;
    }

    private void EnsureOpen()
    {
        if (!IsOpen || _englishDocument is null || _frenchDocument is null)
            throw new InvalidOperationException("Open an English and French template first.");
    }

    private static void ValidateSource(string path)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException("Template file was not found.", path);

        if (!string.Equals(Path.GetExtension(path), ".docx", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("Only .docx templates are supported.", nameof(path));
    }

    public void Close()
    {
        CloseTestMode();
        CloseDocument(ref _frenchDocument);
        CloseDocument(ref _englishDocument);

        if (_word is not null)
        {
            try
            {
                _word.Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
            }
            catch (COMException)
            {
                // Word may already have been closed manually.
            }
            finally
            {
                ReleaseCom(_word);
                _word = null;
            }
        }

        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }

    private static void CloseDocument(ref Word.Document? document)
    {
        if (document is null)
            return;

        try
        {
            document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
        }
        catch (COMException)
        {
            // The user may have already closed the document/window.
        }
        finally
        {
            ReleaseCom(document);
            document = null;
        }
    }

    private static void ReleaseCom(object? value)
    {
        if (value is not null && Marshal.IsComObject(value))
            Marshal.FinalReleaseComObject(value);
    }

    public void Dispose()
    {
        Close();
        GC.SuppressFinalize(this);
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct NativeRect
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetWindowRect(IntPtr hWnd, out NativeRect rect);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool BringWindowToTop(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool MoveWindow(
        IntPtr hWnd,
        int x,
        int y,
        int width,
        int height,
        [MarshalAs(UnmanagedType.Bool)] bool repaint);
}
