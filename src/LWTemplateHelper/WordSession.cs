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

            ArrangeTopBottom();
        }
        catch
        {
            Close();
            throw;
        }
    }

    public void ArrangeTopBottom() => ArrangeTiled(topBottom: true);

    public void ArrangeSideBySide() => ArrangeTiled(topBottom: false);

    public void FocusEnglish() => Focus(_englishDocument);

    public void FocusFrench() => Focus(_frenchDocument);

    private void ArrangeTiled(bool topBottom)
    {
        EnsureOpen();

        var area = Screen.FromControl(Form.ActiveForm ?? throw new InvalidOperationException("Manager window is unavailable.")).WorkingArea;
        var englishWindow = GetWindow(_englishDocument);
        var frenchWindow = GetWindow(_frenchDocument);

        try
        {
            if (topBottom)
            {
                int firstHeight = area.Height / 2;
                Position(englishWindow, area.Left, area.Top, area.Width, firstHeight);
                Position(frenchWindow, area.Left, area.Top + firstHeight, area.Width, area.Height - firstHeight);
            }
            else
            {
                int firstWidth = area.Width / 2;
                Position(englishWindow, area.Left, area.Top, firstWidth, area.Height);
                Position(frenchWindow, area.Left + firstWidth, area.Top, area.Width - firstWidth, area.Height);
            }

            englishWindow.Activate();
        }
        finally
        {
            ReleaseCom(englishWindow);
            ReleaseCom(frenchWindow);
        }
    }

    private static void Position(Word.Window window, int left, int top, int width, int height)
    {
        window.WindowState = Word.WdWindowState.wdWindowStateNormal;
        window.Left = left;
        window.Top = top;
        window.Width = Math.Max(width, 300);
        window.Height = Math.Max(height, 250);
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
            throw new ArgumentException("Only .docx templates are supported in Phase 1.", nameof(path));
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
}
