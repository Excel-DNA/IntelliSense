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
        Label label;
        System.ComponentModel.IContainer components;
        Win32Window _owner;
        int _left;
        int _top;
        Brush _textBrush;
        Brush _linkBrush;
        Pen _borderPen;
        Pen _borderLightPen;
        ToolTip tipDna;
        Dictionary<FontStyle, Font> _fonts;
        Rectangle _linkRect;
        bool _linkActive;
        string _linkAddress;

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
            _textBrush = new SolidBrush(Color.FromArgb(68, 68, 68));
            _linkBrush = new SolidBrush(Color.Blue);
            _borderPen = new Pen(Color.FromArgb(195, 195, 195));
            _borderLightPen = new Pen(Color.FromArgb(225, 225, 225));
            //Win32Helper.SetParent(this.Handle, hwndOwner);

            // _owner = new NativeWindow();
            // _owner.AssignHandle(hwndParent); (...with ReleaseHandle in Dispose)
            Debug.Print($"Created ToolTipForm with owner {hwndOwner}");
        }

        protected override void DefWndProc(ref Message m)
        {
            const int WM_MOUSEACTIVATE = 0x21;
            const int MA_NOACTIVATE = 0x0003;

            switch (m.Msg)
            {
                case WM_MOUSEACTIVATE:
                    m.Result = (IntPtr)MA_NOACTIVATE;
                    return;
            }
            base.DefWndProc(ref m);
        }

        public void ShowToolTip()
        {
            try
            {
                Show(_owner);
                //Show();
            }
            catch (Exception e)
            {
                Debug.Write("ToolTipForm.Show error " + e);
            }
        }

        public void ShowToolTip(FormattedText text, int left, int top)
        {
            _text = text;
            _left = left;
            _top = top;
            if (!Visible)
            {
                Debug.Print($"Showing ToolTipForm: {_text.ToString()}");
                Left = _left;
                Top = _top;
                ShowToolTip();
            }
            else
            {
                Debug.Print($"Invalidating ToolTipForm: {_text.ToString()}");
                Invalidate();
            }
        }

        public void MoveToolTip(int left, int top)
        {
            Invalidate();
            _left = left;
            _top = top;
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
                    //Win32Helper.SetParent(this.Handle, value);
                    if (Visible)
                    {
                        // CONSIDER: Rather just change Owner....
                        Hide();
                        ShowToolTip();
                    }
                }
            }
        }

        protected override bool ShowWithoutActivation
        {
            get { return true; }
        }

        //protected override void Dispose(bool disposing)
        //{
        //    base.Dispose(disposing);
        //    if (_owner != null)
        //    {
        //        _owner.ReleaseHandle();
        //        _owner = null;
        //    }
        //}

        
        // Sometimes has Invalid Handle error when calling base.CreateParams (called from Invalidate() for some reason)
        CreateParams _createParams;

        protected override CreateParams CreateParams
        {
            get
            {
                //if (_createParams == null)
                {
                    const int CS_DROPSHADOW = 0x00020000;
                    //const int WS_CHILD = 0x40000000;
                    const int WS_EX_TOOLWINDOW = 0x00000080;
                    const int WS_EX_NOACTIVATE = 0x08000000;
                    // NOTE: I've seen exception with invalid handle in the base.CreateParams call here...
                    _createParams = base.CreateParams;
                    _createParams.ClassStyle |= CS_DROPSHADOW;
                    // baseParams.Style |= WS_CHILD;
                    _createParams.ExStyle |= (WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW);
                }
                return _createParams;
            }
        }

        protected override void OnHandleDestroyed(EventArgs e)
        {
            base.OnHandleDestroyed(e);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Debug.Assert(_left != 0 || _top != 0);
            Logger.Display.Verbose($"ToolTipForm OnPaint: {_text.ToString()} @ ({_left},{_top})");
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
                            _linkRect = new Rectangle(layoutRect.X - textSize.Width, layoutRect.Y, textSize.Width, textSize.Height);
                            _linkAddress = run.LinkAddress;
                        }

                        totalWidth += textSize.Width;
                        lineHeight = Math.Max(lineHeight, textSize.Height);
                    }
                    lineWidths.Add(totalWidth);
                    totalWidth = 0;
                    totalHeight += lineHeight;
                    layoutRect = new Rectangle(layoutLeft, layoutTop + totalHeight - 1, 1000, 500);
                }
            }
            var width = lineWidths.Max() + widthPadding;
            var height = totalHeight + heightPadding;
            SetBounds(_left, _top, width, height);
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

            
        public void DrawRoundedRectangle(Graphics g, RectangleF r, float radiusX, float radiusY)
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

        void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.label = new System.Windows.Forms.Label();
            this.tipDna = new System.Windows.Forms.ToolTip(this.components);
            this.SuspendLayout();
            // 
            // label
            // 
            this.label.AutoSize = true;
            this.label.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.label.Location = new System.Drawing.Point(3, 3);
            this.label.Name = "label";
            this.label.Size = new System.Drawing.Size(105, 13);
            this.label.TabIndex = 0;
            this.label.Text = "Some long label text.";
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
            this.MouseClick += new System.Windows.Forms.MouseEventHandler(this.ToolTipForm_MouseClick);
            this.MouseMove += new System.Windows.Forms.MouseEventHandler(this.ToolTipForm_MouseMove);
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

        void ToolTipForm_MouseMove(object sender, MouseEventArgs e)
        {
            var inLink = _linkRect.Contains(e.Location);
            if ((inLink && !_linkActive) ||
                (!inLink && _linkActive))
            {
                _linkActive = !_linkActive;
                Invalidate();
            }
        }

        void ToolTipForm_MouseClick(object sender, MouseEventArgs e)
        {
            var inLink = _linkRect.Contains(e.Location);
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
    }
}
