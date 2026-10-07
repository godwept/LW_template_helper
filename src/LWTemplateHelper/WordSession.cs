using System.ComponentModel;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace LWTemplateHelper;

internal sealed record CapturedSelectionInfo(string Preview, int Start, int End);

internal sealed class WordSession : IDisposable
{
    private Word.Application? _word;
    private Word.Document? _englishDocument;
    private Word.Document? _frenchDocument;
    private Word.Range? _englishCapturedRange;
    private Word.Range? _frenchCapturedRange;

    public bool IsOpen => _word is not null;
    public string? SessionDirectory { get; private set; }
    public string? EnglishWorkingPath { get; private set; }
    public string? FrenchWorkingPath { get; private set; }

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
        ArrangeSideBySide();
    }

    public CapturedSelectionInfo CaptureEnglishSelection() =>
        CaptureSelection(_englishDocument, ref _englishCapturedRange, "English");

    public CapturedSelectionInfo CaptureFrenchSelection() =>
        CaptureSelection(_frenchDocument, ref _frenchCapturedRange, "French");

    public CapturedSelectionInfo? GetEnglishCapturedSelection() =>
        ReadCapturedSelection(ref _englishCapturedRange);

    public CapturedSelectionInfo? GetFrenchCapturedSelection() =>
        ReadCapturedSelection(ref _frenchCapturedRange);

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

            englishRange = GetCapturedRangeDuplicate(ref _englishCapturedRange, "English");
            frenchRange = GetCapturedRangeDuplicate(ref _frenchCapturedRange, "French");

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

            englishRange = GetCapturedRangeDuplicate(ref _englishCapturedRange, "English");
            frenchRange = GetCapturedRangeDuplicate(ref _frenchCapturedRange, "French");

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

        var manager = Form.ActiveForm ?? throw new InvalidOperationException("Manager window is unavailable.");
        var area = Screen.FromControl(manager).WorkingArea;
        var englishWindow = GetWindow(_englishDocument);
        var frenchWindow = GetWindow(_frenchDocument);

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

    private static void ValidateBookmarkName(string bookmarkName)
    {
        if (!ConfigurationValidation.TryValidateBookmarkName(bookmarkName, out string error))
            throw new ArgumentException(error, nameof(bookmarkName));
    }

    private static Word.Range GetCapturedRangeDuplicate(
        ref Word.Range? storedRange,
        string language)
    {
        if (ReadCapturedSelection(ref storedRange) is null || storedRange is null)
        {
            throw new InvalidOperationException(
                $"Capture a valid {language} selection before performing this bookmark operation.");
        }

        try
        {
            return storedRange.Duplicate;
        }
        catch (COMException)
        {
            ReleaseCom(storedRange);
            storedRange = null;
            throw new InvalidOperationException(
                $"The captured {language} selection is no longer valid. Capture it again.");
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

    private static CapturedSelectionInfo CaptureSelection(
        Word.Document? document,
        ref Word.Range? storedRange,
        string language)
    {
        if (document is null)
            throw new InvalidOperationException("The Word session is not open.");

        Word.Window? window = null;
        Word.Selection? selection = null;
        Word.Range? currentRange = null;
        Word.Range? duplicate = null;

        try
        {
            window = GetWindow(document);
            selection = window.Selection;
            currentRange = selection.Range;

            if (currentRange.Start >= currentRange.End)
                throw new InvalidOperationException(
                    $"Select the text to capture in the {language} working copy first.");

            duplicate = currentRange.Duplicate;
            var info = BuildSelectionInfo(duplicate);

            ReleaseCom(storedRange);
            storedRange = duplicate;
            duplicate = null;

            return info;
        }
        finally
        {
            ReleaseCom(duplicate);
            ReleaseCom(currentRange);
            ReleaseCom(selection);
            ReleaseCom(window);
        }
    }

    private static CapturedSelectionInfo? ReadCapturedSelection(ref Word.Range? range)
    {
        if (range is null)
            return null;

        try
        {
            if (range.Start >= range.End)
            {
                ReleaseCom(range);
                range = null;
                return null;
            }

            return BuildSelectionInfo(range);
        }
        catch (COMException)
        {
            ReleaseCom(range);
            range = null;
            return null;
        }
    }

    private static CapturedSelectionInfo BuildSelectionInfo(Word.Range range)
    {
        string preview = (range.Text ?? string.Empty)
            .Replace("\r", "¶")
            .Replace("\a", "¤")
            .Replace("\v", "↵")
            .Replace("\t", "→");

        preview = preview.Trim();

        if (preview.Length > 90)
            preview = preview[..87] + "...";

        if (preview.Length == 0)
            preview = "(selected range)";

        return new CapturedSelectionInfo(preview, range.Start, range.End);
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
        ClearCapturedRanges();
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

    private void ClearCapturedRanges()
    {
        ReleaseCom(_englishCapturedRange);
        _englishCapturedRange = null;

        ReleaseCom(_frenchCapturedRange);
        _frenchCapturedRange = null;
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
