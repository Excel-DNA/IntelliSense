using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExcelDna.IntelliSense
{

    // TODO: Drop shadow: http://stackoverflow.com/questions/16493698/drop-shadow-on-a-borderless-winform
    class ToolTipForm  : Form
    {
        FormattedText _text;
        Label label;
        Label labelDna;
        ToolTip tipDna;
        System.ComponentModel.IContainer components;
        Win32Window _owner;
        int _left;
        int _top;

        public ToolTipForm(IntPtr hwndOwner)
        {
            Debug.Assert(hwndOwner != IntPtr.Zero);
            InitializeComponent();
            _owner = new Win32Window(hwndOwner);
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
            if (!Visible)
            {
                Debug.Print($"Showing ToolTipForm: {_text.ToString()}");
                ShowToolTip();
            }
            else
            {
                Debug.Print($"Invalidating ToolTipForm: {_text.ToString()}");
                Invalidate();
            }
            _left = left;
            _top = top;
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
                if (_owner == null) return IntPtr.Zero;
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
                        // Rather just change Owner....
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
            Debug.Print($"Painting ToolTipForm: {_text.ToString()}");
            const int leftPadding = 3;
            const int linePadding = 2;
            const int widthPadding = 10;
            const int heightPadding = 7;

            base.OnPaint(e);
            List<int> lineWidths = new List<int>();
            int totalWidth = 0;
            int totalHeight = 0;

            using (StringFormat format =
                (StringFormat)StringFormat.GenericTypographic.Clone())
            {
                int layoutLeft = ClientRectangle.Location.X + leftPadding;
                int layoutTop = ClientRectangle.Location.Y;
                Rectangle layoutRect = new Rectangle(layoutLeft, layoutTop, 1000, 500);

                format.FormatFlags |= StringFormatFlags.MeasureTrailingSpaces;
                Size textSize;

                foreach (var line in _text)
                {
                    int lineHeight = 0;
                    foreach (var run in line)
                    {
                        using (var font = new Font("Segoe UI", 9, run.Style))   // TODO: Look this up or something ...?
                        {
                            // TODO: Empty strings are a problem....
                            var text = run.Text == "" ? " " : run.Text;

                            // TODO: Find the color SystemBrushes.ControlDarkDark
                            DrawString(e.Graphics, SystemBrushes.WindowFrame, ref layoutRect, out textSize, format, text, font);
                            totalWidth += textSize.Width;
                            lineHeight = Math.Max(lineHeight, textSize.Height);
                        }
                    }
                    lineWidths.Add(totalWidth);
                    totalWidth = 0;
                    totalHeight += lineHeight;
                    layoutRect = new Rectangle(layoutLeft, layoutTop + totalHeight + linePadding, 1000, 500);
                }
            }

            SetBounds(_left, _top, lineWidths.Max() + widthPadding, totalHeight + heightPadding);
//            Size = new Size(lineWidths.Max() + widthPadding, totalHeight + heightPadding);
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

        void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.label = new System.Windows.Forms.Label();
            this.labelDna = new System.Windows.Forms.Label();
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
            // labelDna
            // 
            this.labelDna.BackColor = System.Drawing.Color.CornflowerBlue;
            this.labelDna.Location = new System.Drawing.Point(0, 0);
            this.labelDna.Name = "labelDna";
            this.labelDna.Size = new System.Drawing.Size(2, 2);
            this.labelDna.TabIndex = 0;
            this.tipDna.SetToolTip(this.labelDna, "IntelliSense by Excel-DNA");
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
            this.Controls.Add(this.labelDna);
            this.DoubleBuffered = true;
            this.ForeColor = System.Drawing.Color.DimGray;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "ToolTipForm";
            this.ShowInTaskbar = false;
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
