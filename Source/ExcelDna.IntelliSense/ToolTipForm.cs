using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace ExcelDna.IntelliSense
{
    // CONSIDER: Maybe some ideas from here: http://codereview.stackexchange.com/questions/55916/lightweight-rich-link-label

    // TODO: Drop shadow: http://stackoverflow.com/questions/16493698/drop-shadow-on-a-borderless-winform
    class ToolTipForm  : Form
    {
        FormattedText _text;
        int _linePrefixWidth;
        System.ComponentModel.IContainer components;
        Win32Window _owner;
        // Help Link
        Rectangle _linkClientRect;
        bool _linkActive;
        string _linkAddress;
        long _showTimeTicks; // Track to prevent mouse click-through into help
        // Mouse Capture information for moving        
        bool _captured = false;
        Point _mouseDownScreenLocation;
        Point _mouseDownFormLocation;
        // We keep track of this, else Visibility seems to confuse things...
        int _currentLeft;
        int _currentTop;
        int _showLeft;
        int _showTop;
        int _topOffset; // Might be trying to move the tooltip out of the way of Excel's tip - we track this extra offset here
        int? _listLeft;
        // Various graphics object cached
        Color _textColor;
        Color _linkColor;
        Pen _borderPen;
        Pen _borderLightPen;
        Dictionary<FontStyle, Font> _fonts;
        ToolTip tipDna;

        static Font s_standardFont;

        public ToolTipForm(IntPtr hwndOwner)
        {
            Debug.Assert(hwndOwner != IntPtr.Zero);
            InitializeComponent();
            _owner = new Win32Window(hwndOwner);
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
            SetStyle(ControlStyles.UserMouse | ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint, true);
            Debug.Print($"Created ToolTipForm with owner {hwndOwner}");
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

            switch (m.Msg)
            {
                // Prevent activation by mouse interaction
                case WM_MOUSEACTIVATE:
                    m.Result = (IntPtr)MA_NOACTIVATE;
                    return;
                // We're never active, so we need to do our own mouse handling
                case WM_LBUTTONDOWN:
                    MouseButtonDown(GetMouseLocation(m.LParam));
                    return;
                case WM_MOUSEMOVE:
                    MouseMoved(GetMouseLocation(m.LParam));
                    return;
                case WM_LBUTTONUP:
                    MouseButtonUp(GetMouseLocation(m.LParam));
                    return;
                case WM_SETCURSOR:
                case WM_LBUTTONDBLCLK:
                    // We need to handle this message to prevent flicker (possibly because we're not 'active').
                    m.Result = new IntPtr(1); //Signify that we dealt with the message.
                    return;
                default:
                    base.WndProc(ref m);
                    return;
            }
        }

        void ShowToolTip()
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

        public void ShowToolTip(FormattedText text, string linePrefix, int left, int top, int topOffset, int? listLeft = null)
        {
            Debug.Print($"@@@ ShowToolTip - Old TopOffset: {_topOffset}, New TopOffset: {topOffset}");
            _text = text;
            _linePrefixWidth = MeasureFormulaStringWidth(linePrefix);
            left += _linePrefixWidth;
            if (left != _showLeft || top != _showTop || topOffset != _topOffset || listLeft != _listLeft)
            {
                // Update the start position and the current position
                _currentLeft = Math.Max(left, 0);   // Don't move off the screen
                _currentTop = Math.Max(top, -topOffset);
                _showLeft = _currentLeft;
                _showTop = _currentTop;
                _topOffset = topOffset;
                _listLeft = listLeft;
            }
            if (!Visible)
            {
                Debug.Print($"ShowToolTip - Showing ToolTipForm: {linePrefix} => {_text.ToString()}");
                // Make sure we're in the right position before we're first shown
                SetBounds(_currentLeft, _currentTop + _topOffset, 0, 0);
                _showTimeTicks = DateTime.UtcNow.Ticks;
                ShowToolTip();
            }
            else
            {
                Debug.Print($"ShowToolTip - Invalidating ToolTipForm: {linePrefix} => {_text.ToString()}");
                Invalidate();
            }
        }

        public void MoveToolTip(int left, int top, int topOffset, int? listLeft = null)
        {
            Debug.Print($"@@@ MoveToolTip - Old TopOffset: {_topOffset}, New TopOffset: {topOffset}");
            left += _linePrefixWidth;
            // We might consider checking the new position against earlier mouse movements
            _currentLeft = left;
            _currentTop = top;
            _showLeft = left;
            _showTop = top;
            _topOffset = topOffset;
            _listLeft = listLeft;
            Invalidate();
        }

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
                        ShowToolTip();
                    }
                }
            }
        }

        // TODO: Move or clean up or something...
        int MeasureFormulaStringWidth(string formulaString)
        {
            if (string.IsNullOrEmpty(formulaString))
                return 0;

            var size = TextRenderer.MeasureText(formulaString, s_standardFont);
            return size.Width;
        }

        #region Mouse Handling
        
        void MouseButtonDown(Point screenLocation)
        {
            if (!_linkClientRect.Contains(PointToClient(screenLocation)))
            {
                _captured = true;
                Win32Helper.SetCapture(Handle);
                _mouseDownScreenLocation = screenLocation;
                _mouseDownFormLocation = new Point(_currentLeft, _currentTop);
            }
        }

        void MouseMoved(Point screenLocation)
        {
            if (_captured)
            {
                int dx = screenLocation.X - _mouseDownScreenLocation.X;
                int dy = screenLocation.Y - _mouseDownScreenLocation.Y;
                _currentLeft = _mouseDownFormLocation.X + dx;
                _currentTop = _mouseDownFormLocation.Y + dy;
                Invalidate();
                return;
            }
            var inLink = _linkClientRect.Contains(PointToClient(screenLocation));
            if ((inLink && !_linkActive) ||
                (!inLink && _linkActive))
            {
                _linkActive = !_linkActive;
                Invalidate();
            }
            if (inLink)
                Cursor.Current = Cursors.Hand;
            else
                Cursor.Current = Cursors.SizeAll;
        }

        void MouseButtonUp(Point screenLocation)
        {
            if (_captured)
            {
                _captured = false;
                Win32Helper.ReleaseCapture();
                return;
            }

            // This delay check of 500 ms is inserted to prevent the double-click on the formula list to also be processed 
            // as a click on the toolip, launching the help erroneously.
            var nowTicks = DateTime.UtcNow.Ticks;
            if (nowTicks - _showTimeTicks < 5000000)
                return;

            var inLink = _linkClientRect.Contains(PointToClient(screenLocation));
            if (inLink)
            {
                LaunchLink(_linkAddress);
            }
        }

        void LaunchLink(string address)
        {
            try
            {
                if (address.StartsWith("http", StringComparison.OrdinalIgnoreCase))
                {
                    Process.Start(address);
                }
                else
                {
                    var parts = address.Split('!');
                    if (parts.Length == 2)
                    {
                        // (This is the expected case)
                        // Assume we have a filename!topicid
                        var fileName = parts[0];
                        var topicId = parts[1];
                        if (File.Exists(fileName))
                        {
                            dynamic app = ExcelDna.Integration.ExcelDnaUtil.Application;
                            app.Help(fileName, topicId);
                            // Help.ShowHelp(null, fileName, HelpNavigator.TopicId, topicId);
                        }
                        else
                        {
                            MessageBox.Show($"The help link could not be activated:\r\n\r\nThe file {fileName} could not be located.", "IntelliSense by Excel-DNA", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        // Just show the file ...?
                        if (File.Exists(address))
                        {
                            dynamic app = ExcelDna.Integration.ExcelDnaUtil.Application;
                            app.Help(address); 
                            // Help.ShowHelp(null, address, HelpNavigator.TableOfContents);
                        }
                        else
                        {
                            MessageBox.Show($"The help link could not be activated:\r\n\r\nThe file {address} could not be located.", "IntelliSense by Excel-DNA", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"The help link could not be activated:\r\n\r\n{ex.Message}", "IntelliSense by Excel-DNA", MessageBoxButtons.OK, MessageBoxIcon.Error);
                // NOTE: In this case, the Excel process does not quit after closing Excel...
            }
        }
        
        Point GetMouseLocation(IntPtr lParam)
        {
            int x = (short)(unchecked((int)(long)lParam)  & 0xFFFF);
            int y = (short)((unchecked((int)(long)lParam) >> 16) & 0xFFFF);
            return PointToScreen(new Point(x, y));
        }
        #endregion

        #region Painting
        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);

            const int maxWidth = 1000;  // Should depend on screen width?
            const int maxHeight = 500;
            const int leftPadding = 6;
            const int linePadding = 0;
            const int widthPadding = 12;
            const int heightPadding = 2;
            const int minLineHeight = 16;

            int layoutLeft = ClientRectangle.Location.X + leftPadding;
            int layoutTop = ClientRectangle.Location.Y;

            var textFormatFlags = TextFormatFlags.Left | 
                                  TextFormatFlags.Top | 
                                  TextFormatFlags.NoPadding | 
                                  TextFormatFlags.SingleLine | 
                                  TextFormatFlags.ExternalLeading;

            List<int> lineWidths = new List<int>();
            int currentHeight = 0; // Measured from layoutTop, going down
            foreach (var line in _text)
            {
                currentHeight += linePadding;
                int lineHeight = minLineHeight;
                int lineWidth = 0;
                foreach (var run in line)
                {
                    // We support only a single link, for now
                    Font font;
                    Color color;
                    if (run.IsLink && _linkActive)
                    {
                        font = _fonts[FontStyle.Underline];
                        color = _linkColor;
                    }
                    else
                    {
                        font = _fonts[run.Style];
                        color = _textColor;
                    }

                    foreach (var text in GetRunParts(run.Text))
                    {
                        if (text == "") continue;

                        var location = new Point(layoutLeft + lineWidth, layoutTop + currentHeight);
                        var proposedSize = new Size(maxWidth - lineWidth, maxHeight - currentHeight);
                        var textSize = TextRenderer.MeasureText(e.Graphics, text, font, proposedSize, textFormatFlags);
                        if (textSize.Width <= proposedSize.Width)
                        {
                            // Draw it in this line
                            TextRenderer.DrawText(e.Graphics, text, font, location, color, textFormatFlags);
                        }
                        else
                        {
                            if (lineWidth > 0)  // Check if we aren't on the first line, and already overflowing - might then line-break rather than ellipses...
                            {
                                // Make a new line and definitely draw it there (maybe with ellipses?)
                                lineWidths.Add(lineWidth);
                                currentHeight += lineHeight;
                                currentHeight += linePadding;

                                lineHeight = minLineHeight;
                                lineWidth = 2;  // Start with a little bit of indent on these lines

                                // TODO: Clean up this duplication
                                location = new Point(layoutLeft + lineWidth, layoutTop + currentHeight);
                                proposedSize = new Size(maxWidth - lineWidth, maxHeight - currentHeight);
                                textSize = TextRenderer.MeasureText(e.Graphics, text, font, proposedSize, textFormatFlags);
                                if (textSize.Width <= proposedSize.Width)
                                {
                                    // Draw it in this line (the new one)
                                    TextRenderer.DrawText(e.Graphics, text, font, location, color, textFormatFlags);
                                }
                                else
                                {
                                    // Even too long for a full line - draw truncated
                                    textSize = new Size(proposedSize.Width, textSize.Height);
                                    var bounds = new Rectangle(location, textSize);
                                    TextRenderer.DrawText(e.Graphics, text, font, bounds, color, textFormatFlags | TextFormatFlags.EndEllipsis);
                                }
                            }
                            else
                            {
                                // Draw truncated
                                textSize = new Size(proposedSize.Width, textSize.Height);
                                var bounds = new Rectangle(location, textSize);
                                //  new Rectangle(layoutLeft + lineWidth, layoutTop + currentHeight, maxWidth, maxHeight - currentHeight)
                                TextRenderer.DrawText(e.Graphics, text, font, bounds, color, textFormatFlags | TextFormatFlags.EndEllipsis);
                            }
                        }

                        if (run.IsLink)
                        {
                            _linkClientRect = new Rectangle(layoutLeft + lineWidth, layoutTop + currentHeight, textSize.Width, textSize.Height);
                            _linkAddress = run.LinkAddress;
                        }

                        lineWidth += textSize.Width; // + 1;
                        lineHeight = Math.Max(lineHeight, textSize.Height);
                    }
                }
                lineWidths.Add(lineWidth);
                currentHeight += lineHeight;
            }

            var width = lineWidths.Max() + widthPadding;
            var height = currentHeight + heightPadding;

            UpdateLocation(width, height);
            DrawRoundedRectangle(e.Graphics, new RectangleF(0, 0, Width - 1, Height - 1), 2, 2);
        }

        static IEnumerable<string> GetRunParts(string runText)
        {
            int lastStart = 0;
            for (int i = 0; i < runText.Length; i++)
            {
                if (runText[i] == ',' || runText[i] == ' ')
                {
                    yield return runText.Substring(lastStart, i - lastStart + 1);
                    lastStart = i + 1;
                }
            }
            yield return runText.Substring(lastStart);
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

        void UpdateLocation(int width, int height)
        {
            var workingArea = Screen.GetWorkingArea(new Point(_currentLeft, _currentTop + _topOffset));
            bool tipFits = workingArea.Contains(new Rectangle(_currentLeft, _currentTop + _topOffset, width, height));
            if (!tipFits && (_currentLeft == _showLeft && _currentTop == _showTop))
            {
                // It doesn't fit and it's still where we initially tried to show it 
                // (so it probably hasn't been moved).
                if (_listLeft == null)
                {
                    // Not in list selection mode - probably FormulaEdit
                    _currentLeft -= Math.Max(0, (_currentLeft + width) - workingArea.Right);
                    // CONSIDER: Move up too???
                }
                else
                {
                    const int leftPadding = 4;
                    // Check if it fits on the left
                    if (width < _listLeft.Value - leftPadding)
                    {
                        _currentLeft = _listLeft.Value - width - leftPadding;
                    }
                }

                if (_currentLeft < 0)
                    _currentLeft = 0;
            }
            SetBounds(_currentLeft, _currentTop + _topOffset, width, height);
        }
        #endregion

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams createParams;
                const int CS_DROPSHADOW = 0x00020000;
                //const int WS_CHILD = 0x40000000;
                const int WS_EX_TOOLWINDOW = 0x00000080;
                const int WS_EX_NOACTIVATE = 0x08000000;
                // NOTE: I've seen exception with invalid handle in the base.CreateParams call here...
                createParams = base.CreateParams;
                createParams.ClassStyle |= CS_DROPSHADOW;
                // baseParams.Style |= WS_CHILD;
                createParams.ExStyle |= (WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW);
                return createParams;
            }
        }

        protected override bool ShowWithoutActivation => true;

        void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.tipDna = new System.Windows.Forms.ToolTip(this.components);
            this.SuspendLayout();
            // 
            // tipDna
            // 
            this.tipDna.ShowAlways = true;
            // 
            // ToolTipForm
            // 
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(114, 20);
            this.ControlBox = false;
            this.DoubleBuffered = true;
            this.ForeColor = System.Drawing.Color.DimGray;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "ToolTipForm";
            this.ShowInTaskbar = false;
            this.tipDna.SetToolTip(this, "IntelliSense by Excel-DNA");
            this.ResumeLayout(false);

        }

        internal static void SetStandardFont(string standardFontName, double standardFontSize)
        {
            s_standardFont = new Font(standardFontName, (float)standardFontSize);
        }

        class Win32Window : IWin32Window
        {
            public IntPtr Handle
            {
                get;
                private set;
            }

            public Win32Window(IntPtr handle)
            {
                Handle = handle;
            }
        }
    }
}
