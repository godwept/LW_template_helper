using System.Runtime.InteropServices;

namespace LWTemplateHelper;

internal static class AppTheme
{
    public static readonly Color WindowBack = Color.FromArgb(17, 24, 39);
    public static readonly Color PanelBack = Color.FromArgb(31, 41, 55);
    public static readonly Color PanelBackAlt = Color.FromArgb(36, 47, 62);
    public static readonly Color HeaderBack = Color.FromArgb(30, 58, 95);
    public static readonly Color InputBack = Color.FromArgb(40, 50, 65);
    public static readonly Color Text = Color.FromArgb(241, 245, 249);
    public static readonly Color MutedText = Color.FromArgb(148, 163, 184);
    public static readonly Color Accent = Color.FromArgb(0, 180, 170);
    public static readonly Color AccentDark = Color.FromArgb(0, 122, 116);
    public static readonly Color Danger = Color.FromArgb(153, 55, 66);
    public static readonly Color DangerAccent = Color.FromArgb(239, 99, 111);
    public static readonly Color Border = Color.FromArgb(55, 65, 81);

    public static readonly Color SuccessBack = Color.FromArgb(31, 78, 60);
    public static readonly Color ErrorBack = Color.FromArgb(96, 44, 51);
    public static readonly Color NeutralBack = Color.FromArgb(45, 53, 64);
    public static readonly Color WarningBack = Color.FromArgb(82, 70, 40);

    public static void Apply(Form form)
    {
        form.BackColor = WindowBack;
        form.ForeColor = Text;
        form.Font = new Font("Segoe UI", 9F, FontStyle.Regular);
        ApplyRecursive(form);

        form.HandleCreated += (_, _) => TryEnableDarkTitleBar(form.Handle);
        if (form.IsHandleCreated)
            TryEnableDarkTitleBar(form.Handle);
    }

    public static void StylePrimaryButton(Button button)
    {
        StyleButton(button);
        button.BackColor = AccentDark;
        button.FlatAppearance.BorderColor = Accent;
    }

    public static void StyleDangerButton(Button button)
    {
        StyleButton(button);
        button.BackColor = Danger;
        button.FlatAppearance.BorderColor = DangerAccent;
    }

    public static void StyleButton(Button button)
    {
        button.FlatStyle = FlatStyle.Flat;
        button.FlatAppearance.BorderSize = 1;
        button.FlatAppearance.BorderColor = Accent;
        button.FlatAppearance.MouseOverBackColor = Color.FromArgb(44, 62, 78);
        button.FlatAppearance.MouseDownBackColor = AccentDark;
        button.BackColor = PanelBackAlt;
        button.ForeColor = Text;
        button.Padding = new Padding(8, 2, 8, 2);
        button.UseVisualStyleBackColor = false;
    }

    public static void StyleGrid(DataGridView grid)
    {
        grid.EnableHeadersVisualStyles = false;
        grid.BackgroundColor = PanelBack;
        grid.BorderStyle = BorderStyle.FixedSingle;
        grid.GridColor = Border;
        grid.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
        grid.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
        grid.ColumnHeadersDefaultCellStyle.BackColor = HeaderBack;
        grid.ColumnHeadersDefaultCellStyle.ForeColor = Text;
        grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = HeaderBack;
        grid.ColumnHeadersDefaultCellStyle.SelectionForeColor = Text;
        grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold);
        grid.ColumnHeadersHeight = 32;

        grid.DefaultCellStyle.BackColor = PanelBack;
        grid.DefaultCellStyle.ForeColor = Text;
        grid.DefaultCellStyle.SelectionBackColor = AccentDark;
        grid.DefaultCellStyle.SelectionForeColor = Color.White;
        grid.DefaultCellStyle.Padding = new Padding(3);

        grid.AlternatingRowsDefaultCellStyle.BackColor = PanelBackAlt;
        grid.AlternatingRowsDefaultCellStyle.ForeColor = Text;
        grid.AlternatingRowsDefaultCellStyle.SelectionBackColor = AccentDark;
        grid.AlternatingRowsDefaultCellStyle.SelectionForeColor = Color.White;
        grid.RowTemplate.Height = 28;
    }

    private static void ApplyRecursive(Control parent)
    {
        foreach (Control control in parent.Controls)
        {
            switch (control)
            {
                case Button button:
                    StyleButton(button);
                    break;

                case DataGridView grid:
                    StyleGrid(grid);
                    break;

                case TextBox textBox:
                    textBox.BackColor = textBox.ReadOnly ? PanelBackAlt : InputBack;
                    textBox.ForeColor = Text;
                    textBox.BorderStyle = BorderStyle.FixedSingle;
                    break;

                case ComboBox combo:
                    combo.BackColor = InputBack;
                    combo.ForeColor = Text;
                    combo.FlatStyle = FlatStyle.Flat;
                    break;

                case GroupBox group:
                    group.BackColor = PanelBack;
                    group.ForeColor = Color.FromArgb(147, 197, 253);
                    group.Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold);
                    break;

                case Label label:
                    label.ForeColor = Text;
                    break;

                case Panel:
                case TableLayoutPanel:
                case FlowLayoutPanel:
                    control.BackColor = Color.Transparent;
                    break;
            }

            if (control.HasChildren)
                ApplyRecursive(control);
        }
    }

    private static void TryEnableDarkTitleBar(IntPtr hwnd)
    {
        if (hwnd == IntPtr.Zero)
            return;

        int enabled = 1;

        try
        {
            if (DwmSetWindowAttribute(hwnd, 20, ref enabled, sizeof(int)) != 0)
                DwmSetWindowAttribute(hwnd, 19, ref enabled, sizeof(int));
        }
        catch
        {
            // Cosmetic only. Older Windows builds may not support this attribute.
        }
    }

    [DllImport("dwmapi.dll")]
    private static extern int DwmSetWindowAttribute(
        IntPtr hwnd,
        int attribute,
        ref int attributeValue,
        int attributeSize);
}
