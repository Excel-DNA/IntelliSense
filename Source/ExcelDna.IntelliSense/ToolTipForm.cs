using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExcelDna.IntelliSense
{
    class ToolTipForm  : Form
    {
        FormattedText _text;
        private Label label;
        private Label label1;
        private ToolTip toolTip1;
        private System.ComponentModel.IContainer components;
        Win32Window _owner;

        public ToolTipForm(IntPtr hwndOwner)
        {
            InitializeComponent();
            _owner = new Win32Window(hwndOwner);
            // _owner = new NativeWindow();
            // _owner.AssignHandle(hwndParent); (...with ReleaseHandle in Dispose)
            Debug.Print($"Created ToolTipForm with owner {hwndOwner}");
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
                ShowToolTip();
            }
            else
            {
                Invalidate();
            }
            Left = left;
            Top = top;
        }

        public void MoveToolTip(int left, int top)
        {
            Invalidate();
            Left = left;
            Top = top;
        }

        public IntPtr OwnerHandle
        {
            get
            {
                if (_owner == null) return IntPtr.Zero;
                return _owner.Handle;
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

        private const int CS_DROPSHADOW = 0x00020000;
        private const int WS_EX_TOOLWINDOW = 0x00000080;
        private const int WS_EX_NOACTIVATE = 0x08000000;
        protected override CreateParams CreateParams
        {
            get
            {
                // NOTE: I've seen exception with invalid handle in the base.CreateParams call here...
                CreateParams baseParams = base.CreateParams;
                baseParams.ClassStyle |= CS_DROPSHADOW;
                baseParams.ExStyle |= ( WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW );
                return baseParams;
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
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

            Size = new Size(lineWidths.Max() + widthPadding, totalHeight + heightPadding);
        }

        private void DrawString(Graphics g, Brush brush, ref Rectangle rect, out Size used,
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

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.label = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
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
            // label1
            // 
            this.label1.BackColor = System.Drawing.Color.CornflowerBlue;
            this.label1.Location = new System.Drawing.Point(0, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(2, 2);
            this.label1.TabIndex = 0;
            this.toolTip1.SetToolTip(this.label1, "IntelliSense by Excel-DNA");
            // 
            // ToolTipForm
            // 
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(114, 20);
            this.ControlBox = false;
            this.Controls.Add(this.label1);
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
