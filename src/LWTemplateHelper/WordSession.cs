using System.ComponentModel;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace LWTemplateHelper;

internal sealed class WordSession : IDisposable
{
    private Word.Application? _word;
    private Word.Document? _englishDocument;
    private Word.Document? _frenchDocument;

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
