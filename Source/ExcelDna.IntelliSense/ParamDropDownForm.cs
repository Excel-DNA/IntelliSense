using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace ExcelDna.IntelliSense;

public class ParamDropDownForm : Form
{
    FormattedText _text;
    string _lastCurrentArgValue = "";
    int _maxTextHeight;

    TextFormatFlags _textFormatFlags =
        TextFormatFlags.Left
        | TextFormatFlags.Top
        | TextFormatFlags.NoPadding
        | TextFormatFlags.SingleLine
        | TextFormatFlags.ExternalLeading;

    SynchronizationContext _syncContextMain;

    private Color _highlightColor = Color.LightBlue;
    private Font _font = new Font("Segoe UI", 9f);

    Win32Window _owner;

    // Various graphics object cached
    Color _textColor;
    Color _linkColor;
    Pen _borderPen;
    Pen _borderLightPen;
    Pen _selectedItemPen;
    Dictionary<FontStyle, Font> _fonts;

    int _currentLeft;
    int _currentTop;
    int _showLeft;
    int _showTop;
    int _topOffset; // Might be trying to move the tooltip out of the way of Excel's tip - we track this extra offset here
    int? _listLeft;
    int? _listTop;

    bool _mouseCaptured = false;
    bool _kbCaptured = false;

    static Font s_standardFont;

    int _selectedIndex = 0;

    protected override bool ShowWithoutActivation => true;

    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public IntPtr OwnerHandle
    {
        get
        {
            if (_owner == null)
                return IntPtr.Zero;
            return _owner.Handle;
        }
        set
        {
            if (_owner == null || _owner.Handle != value)
            {
                _owner = new Win32Window(value);
                if (Visible)
                {
                    // We want to change the owner.
                    // That's hard, so we hide and re-show.
                    Hide();
                    ShowSelector();
                }
            }
        }
    }

    protected override CreateParams CreateParams
    {
        get
        {
            CreateParams createParams;
            const int CS_DROPSHADOW = 0x00020000;
            //const int WS_CHILD = 0x40000000;
            const int WS_EX_TOOLWINDOW = 0x00000080;
            const int WS_EX_NOACTIVATE = 0x08000000;
            // const int WS_EX_TOPMOST = 0x00000008;
            const int WS_TABSTOP = 0x00010000;
            // NOTE: I've seen exception with invalid handle in the base.CreateParams call here...
            createParams = base.CreateParams;
            createParams.ClassStyle &= ~CS_DROPSHADOW;
            // baseParams.Style |= WS_CHILD;
            createParams.ExStyle |= (WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW);
            createParams.ExStyle &= ~WS_TABSTOP;
            return createParams;
        }
    }

    public ParamDropDownForm(IntPtr hwndOwner, SynchronizationContext syncContextMain)
    {
        InitializeComponents();
        _owner = new Win32Window(hwndOwner);
        _syncContextMain = syncContextMain;
        // CONSIDER: Maybe make a more general solution that lazy-loads as needed
        _fonts = new Dictionary<FontStyle, Font>
        {
            { FontStyle.Regular, new Font("Segoe UI", 9, FontStyle.Regular) },
            { FontStyle.Bold, new Font("Segoe UI", 9, FontStyle.Bold) },
            { FontStyle.Italic, new Font("Segoe UI", 9, FontStyle.Italic) },
            { FontStyle.Underline, new Font("Segoe UI", 9, FontStyle.Underline) },
            { FontStyle.Bold | FontStyle.Italic, new Font("Segoe UI", 9, FontStyle.Bold | FontStyle.Italic) },
        };
        _textColor = Color.FromArgb(72, 72, 72);
        _linkColor = Color.Blue;
        _borderPen = new Pen(Color.FromArgb(195, 195, 195));
        _borderLightPen = new Pen(Color.FromArgb(225, 225, 225));
        _selectedItemPen = new Pen(Color.FromArgb(204, 232, 255), 1.5f);
        SetStyle(ControlStyles.UserMouse | ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint, true);
        Debug.Print($"Created ToolTipForm with owner {hwndOwner}");
    }

    private void InitializeComponents()
    {
        this.SuspendLayout();
        this.StartPosition = FormStartPosition.Manual;
        this.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        this.BackColor = Color.LightGray;
        this.ClientSize = new Size(114, 20);
        this.ControlBox = false;
        this.DoubleBuffered = true;
        this.ForeColor = Color.DimGray;
        this.FormBorderStyle = FormBorderStyle.None;
        this.Name = "ParamDropDownForm";
        this.ShowInTaskbar = false;
        this.ResumeLayout(false);
    }

    internal void OnKeyboardEvent(object sender, KeyboardEventHook.KeyboardEventArgs args)
    {
        if (!Visible)
        {
            return;
        }

        Keys key = (Keys)args.Key;
        var count = _text.Count();

        switch (key)
        {
            case Keys.Up:
                if (args.LParam.IsReleased)
                {
                    return;
                }
                _selectedIndex -= 1;
                if (_selectedIndex < 0)
                {
                    _selectedIndex += count;
                }
                Invalidate();
                return;
            case Keys.Down:
                if (args.LParam.IsReleased)
                {
                    return;
                }
                _selectedIndex += 1;
                if (_selectedIndex > count - 1)
                {
                    _selectedIndex -= count;
                }
                Invalidate();
                return;
            case Keys.Tab:
                if (!args.LParam.IsReleased || _kbCaptured)
                {
                    return;
                }
                _kbCaptured = true;
                _syncContextMain.Post(AcceptSelection, true);
                return;
            case Keys.Escape:
                if (args.LParam.IsReleased)
                {
                    return;
                }
                Hide();
                return;
        }
    }

    protected override void WndProc(ref Message m)
    {
        const int WM_MOUSEACTIVATE = 0x21;
        const int WM_MOUSEMOVE = 0x0200;
        const int WM_LBUTTONDOWN = 0x0201;
        const int WM_LBUTTONUP = 0x0202;
        const int WM_LBUTTONDBLCLK = 0x203;
        const int WM_SETCURSOR = 0x20;
        const int MA_NOACTIVATE = 0x0003;
        // const int WM_KEYDOWN = 0x0100;
        // const int WM_CHAR = 0x0102;

        switch (m.Msg)
        {
            // Prevent activation by mouse interaction
            case WM_MOUSEACTIVATE:
                m.Result = (IntPtr)MA_NOACTIVATE;
                return;
            // We're never active, so we need to do our own mouse handling
            case WM_LBUTTONDOWN:
                MouseButtonDown();
                return;
            case WM_MOUSEMOVE:
                MouseMoved();
                return;
            case WM_LBUTTONUP:
                MouseButtonUp();
                return;
            case WM_SETCURSOR:
            case WM_LBUTTONDBLCLK:
                // We need to handle this message to prevent flicker (possibly because we're not 'active').
                m.Result = new IntPtr(1); //Signify that we dealt with the message.
                return;
            // case WM_KEYDOWN:
            // case WM_CHAR:
            //     Keys key = (Keys)(int)m.WParam;
            //     // OnKeyDown(key);
            //     return;
            default:
                base.WndProc(ref m);
                return;
        }
    }

    void ShowSelector()
    {
        try
        {
            Show(_owner);
        }
        catch (Exception e)
        {
            Debug.Write("ToolTipForm.Show error " + e);
        }
    }

    internal void ShowSelector(
        FormattedText text,
        string functionName,
        string currentArgValue,
        int left,
        int top,
        int topOffset,
        int? listLeft = null,
        int? listTop = null
    )
    {
        // Debug.Print($"@@@ ShowToolTip - Old TopOffset: {_topOffset}, New TopOffset: {topOffset}");

        _text = text;
        _lastCurrentArgValue = currentArgValue;

        if (
            left != _showLeft
            || top != _showTop
            || topOffset != _topOffset
            || listLeft != _listLeft
            || listTop != _listTop
        )
        {
            Win32Helper.TryGetCaretScreenPosition(_owner.Handle, out var p);

            // // Update the start position and the current position
            _currentLeft = p.X; // Don't move off the screen
            _currentTop = top;
            _showLeft = _currentLeft;
            _showTop = _currentTop;
            _topOffset = topOffset;
            _listLeft = listLeft;
            _listTop = listTop;
        }
        if (!Visible)
        {
            // Debug.Print($"ShowToolTip - Showing ToolTipForm: {linePrefix} => {_text.ToString()}");
            // Make sure we're in the right position before we're first shown
            _selectedIndex = 0;
            SetBounds(_currentLeft, _currentTop + _topOffset, 0, 0);
            // _showTimeTicks = DateTime.UtcNow.Ticks;

            ShowSelector();
        }
        else
        {
            // Debug.Print($"ShowToolTip - Invalidating ToolTipForm: {linePrefix} => {_text.ToString()}");
            Invalidate();
        }
    }

    public void MoveToolTip(int left, int top, int topOffset, int? listLeft = null, int? listTop = null)
    {
        Debug.Print($"@@@ MoveToolTip - Old TopOffset: {_topOffset}, New TopOffset: {topOffset}");
        Win32Helper.TryGetCaretScreenPosition(_owner.Handle, out var p);
        left = p.X;
        // We might consider checking the new position against earlier mouse movements
        _currentLeft = left;
        _currentTop = top;
        _showLeft = left;
        _showTop = top;
        _topOffset = topOffset;
        _listLeft = listLeft;
        _listTop = listTop;
        Invalidate();
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);

        int maxTextWidth = 0;
        var numberOfItems = _text.Count();
        for (var index = 0; index < numberOfItems; index++)
        {
            var line = _text.ElementAt(index);
            var str = line.ToString();
            var textSize = TextRenderer.MeasureText(e.Graphics, str, _font);
            maxTextWidth = Math.Max(maxTextWidth, textSize.Width);
            _maxTextHeight = Math.Max(_maxTextHeight, textSize.Height);
            var textRect = new Rectangle(2, 2 + index * textSize.Height, textSize.Width + 2, textSize.Height + 2);

            using var brush = new SolidBrush(_highlightColor);
            if (index == _selectedIndex)
            {
                e.Graphics.FillRectangle(brush, textRect);
                // e.Graphics.DrawRectangle(_selectedItemPen, textRect);
            }
            TextRenderer.DrawText(e.Graphics, str, _font, textRect, _textColor, _textFormatFlags);
        }

        var height = numberOfItems * _maxTextHeight + numberOfItems * 2;
        var width = maxTextWidth;
        SetBounds(_currentLeft, _currentTop + _topOffset, width, height);
        DrawRoundedRectangle(e.Graphics, new RectangleF(0, 0, Width - 1, Height - 1), 2, 2);
    }

    void DrawRoundedRectangle(Graphics g, RectangleF r, float radiusX, float radiusY)
    {
        var oldMode = g.SmoothingMode;
        g.SmoothingMode = SmoothingMode.None;

        g.DrawRectangle(_borderLightPen, new Rectangle((int)r.X, (int)r.Y, 1, 1));
        g.DrawRectangle(_borderLightPen, new Rectangle((int)(r.X + r.Width - 1), (int)r.Y, 1, 1));
        g.DrawRectangle(_borderLightPen, new Rectangle((int)(r.X + r.Width - 1), (int)(r.Y + r.Height - 1), 1, 1));
        g.DrawRectangle(_borderLightPen, new Rectangle((int)(r.X), (int)(r.Y + r.Height - 1), 1, 1));
        g.DrawRectangle(_borderPen, new Rectangle((int)r.X, (int)r.Y, (int)r.Width, (int)r.Height));

        g.SmoothingMode = oldMode;
    }

    void MouseButtonDown()
    {
        _mouseCaptured = true;
        Win32Helper.SetCapture(Handle);
    }

    void MouseMoved()
    {
        if (_maxTextHeight == 0)
            return;

        var relativeP = PointToClient(Cursor.Position);
        var index = relativeP.Y / _maxTextHeight;
        if (index >= 0 && index < _text.Count())
        {
            _selectedIndex = index;
            Invalidate();
        }
    }

    void MouseButtonUp()
    {
        if (!_mouseCaptured)
        {
            return;
        }

        _mouseCaptured = false;
        Win32Helper.ReleaseCapture();

        IDataObject oldClipboard = null;

        try
        {
            oldClipboard = Clipboard.GetDataObject();

            AcceptSelection();
        }
        finally
        {
            // Restore clipboard
            if (oldClipboard != null)
            {
                try
                {
                    Clipboard.SetDataObject(oldClipboard);
                }
                catch
                {
                    // ignore – clipboard might be locked by another process
                }
            }
        }
    }

    private void AcceptSelection(object state = null)
    {
        var hide = state != null && (bool)state;
        var selectedItem = _text.ElementAtOrDefault(_selectedIndex)?.ToString();

        if (selectedItem == null)
        {
            return;
        }

        SendKeys.SendWait($"{{BACKSPACE {_lastCurrentArgValue.Length}}}");

        var handledSelectedItem = $"\"{selectedItem}\"";

        Clipboard.SetText(handledSelectedItem);

        SendKeys.SendWait("^v");

        _kbCaptured = false;

        if (hide)
            Hide();
        else
            Invalidate();
    }

    internal static void SetStandardFont(string standardFontName, double standardFontSize)
    {
        s_standardFont = new Font(standardFontName, (float)standardFontSize);
    }
}
