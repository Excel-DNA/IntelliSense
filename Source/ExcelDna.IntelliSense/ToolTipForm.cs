using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Windows.Forms;

namespace ExcelDna.IntelliSense
{
    // CONSIDER: Maybe some ideas from here: http://codereview.stackexchange.com/questions/55916/lightweight-rich-link-label

    // TODO: Drop shadow: http://stackoverflow.com/questions/16493698/drop-shadow-on-a-borderless-winform
    class ToolTipForm  : Form
    {
        FormattedText _text;
        System.ComponentModel.IContainer components;
        Win32Window _owner;
        // Help Link
        Rectangle _linkClientRect;
        bool _linkActive;
        string _linkAddress;
        // Mouse Capture information for moving        
        bool _captured = false;
        Point _mouseDownScreenLocation;
        Point _mouseDownFormLocation;
        // We keep track of this, else Visibility seems to confuse things...
        int _currentLeft;
        int _currentTop;
        int _showLeft;
        int _showTop;
        int? _listLeft;
        // Various graphics object cached
        Brush _textBrush;
        Brush _linkBrush;
        Pen _borderPen;
        Pen _borderLightPen;
        Dictionary<FontStyle, Font> _fonts;
        ToolTip tipDna;

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
            //_textBrush = new SolidBrush(Color.FromArgb(68, 68, 68));  // Best matches Excel's built-in color, but I think a bit too light
            _textBrush = new SolidBrush(Color.FromArgb(52, 52, 52));
            _linkBrush = new SolidBrush(Color.Blue);
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

        public void ShowToolTip(FormattedText text, int left, int top, int? listLeft)
        {
            _text = text;
            if (left != _showLeft || top != _showTop || listLeft != _listLeft)
            {
                // Update the start position and the current position
                _currentLeft = left;
                _currentTop = top;
                _showLeft = left;
                _showTop = top;
                _listLeft = listLeft;
            }
            if (!Visible)
            {
                Debug.Print($"ShowToolTip - Showing ToolTipForm: {_text.ToString()}");
                // Make sure we're in the right position before we're first shown
                SetBounds(_currentLeft, _currentTop, 0, 0);
                ShowToolTip();
            }
            else
            {
                Debug.Print($"ShowToolTip - Invalidating ToolTipForm: {_text.ToString()}");
                Invalidate();
            }
        }


        public void MoveToolTip(int left, int top, int? listLeft)
        {
            // We might consider checking the new position against earlier mouse movements
            _currentLeft = left;
            _currentTop = top;
            _showLeft = left;
            _showTop = top;
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

            var inLink = _linkClientRect.Contains(PointToClient(screenLocation));
            if (inLink)
            {
                LaunchLink(_linkAddress);
            }
        }

        void LaunchLink(string address)
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
                    Help.ShowHelp(null, parts[0], HelpNavigator.TopicId, parts[1]);
                }
                else
                {
                    // Just show the file ...?
                    Help.ShowHelp(null, address);
                }
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
            const int leftPadding = 6;
            const int linePadding = 0;
            const int widthPadding = 12;
            const int heightPadding = 2;

            base.OnPaint(e);
            List<int> lineWidths = new List<int>();
            int totalWidth = 0;
            int totalHeight = 0;

            using (StringFormat format =
                (StringFormat)StringFormat.GenericTypographic.Clone())
            {
                int layoutLeft = ClientRectangle.Location.X + leftPadding;
                int layoutTop = ClientRectangle.Location.Y;
                Rectangle layoutRect = new Rectangle(layoutLeft, layoutTop - 1, 1000, 500);

                format.FormatFlags |= StringFormatFlags.MeasureTrailingSpaces;
                Size textSize;

                foreach (var line in _text)
                {
                    totalHeight += linePadding;
                    int lineHeight = 16;
                    foreach (var run in line)
                    {
                        // We support only a single link, for now

                        Font font;
                        Brush brush;
                        if (run.IsLink && _linkActive)
                        {
                            font = _fonts[FontStyle.Underline];
                            brush = _linkBrush;
                        }
                        else
                        {
                            font = _fonts[run.Style];
                            brush = _textBrush;
                        }

                        // TODO: Empty strings are a problem....
                        var text = run.Text == "" ? " " : run.Text;

                        DrawString(e.Graphics, brush, ref layoutRect, out textSize, format, text, font);

                        if (run.IsLink)
                        {
                            _linkClientRect = new Rectangle(layoutRect.X - textSize.Width, layoutRect.Y, textSize.Width, textSize.Height);
                            _linkAddress = run.LinkAddress;
                        }

                        totalWidth += textSize.Width;
                        lineHeight = Math.Max(lineHeight, textSize.Height);

                        // Pad by one extra pixel between runs, until we figure out kerning between runs
                        layoutRect.X += 1;
                        totalWidth += 1;
                    }
                    lineWidths.Add(totalWidth);
                    totalWidth = 0;
                    totalHeight += lineHeight;
                    layoutRect = new Rectangle(layoutLeft, layoutTop + totalHeight - 1, 1000, 500);
                }
            }
            var width = lineWidths.Max() + widthPadding;
            var height = totalHeight + heightPadding;
            UpdateLocation(width, height);
            DrawRoundedRectangle(e.Graphics, new RectangleF(0,0, Width - 1, Height - 1), 2, 2);
        }

        void DrawString(Graphics g, Brush brush, ref Rectangle rect, out Size used,
                                StringFormat format, string text, Font font)
        {
            using (StringFormat copy = (StringFormat)format.Clone())
            {
                copy.SetMeasurableCharacterRanges(new CharacterRange[]
                    {
                        new CharacterRange(0, text.Length)
                    });
                Region[] regions = g.MeasureCharacterRanges(text, font, rect, copy);

                g.DrawString(text, font, brush, rect, format);

                int height = (int)(regions[0].GetBounds(g).Height);
                int width = (int)(regions[0].GetBounds(g).Width);

                // First just one line...
                used = new Size(width, height);

                rect.X += width;
                rect.Width -= width;
            }
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
            var workingArea = Screen.GetWorkingArea(new Point(_currentLeft, _currentTop));
            bool tipFits = workingArea.Contains(new Rectangle(_currentLeft, _currentTop, width, height));
            if (!tipFits && (_currentLeft == _showLeft && _currentTop == _showTop))
            {
                // It doesn't fit and it's still where we initially tried to show it 
                // (so it probably hasn't been moved).
                if (_listLeft == null)
                {
                    // Not in list selection mode - probably FormulaEdit
                    _currentLeft = Math.Max(0, (_currentLeft + width) - workingArea.Right);
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
            }
            SetBounds(_currentLeft, _currentTop, width, height);
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
