using System;
using System.Runtime.InteropServices;
using System.Text;
using ExcelDna.Integration;

namespace ExcelDna.IntelliSense
{
    static class Win32Helper
    {
        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;

            public POINT(int x, int y)
            {
                this.X = x;
                this.Y = y;
            }

            public static implicit operator System.Drawing.Point(POINT p)
            {
                return new System.Drawing.Point(p.X, p.Y);
            }

            public static implicit operator POINT(System.Drawing.Point p)
            {
                return new POINT(p.X, p.Y);
            }
        }

        enum WM : uint
        {
            GETTEXT = 0x000D,
            GETTEXTLENGTH = 0x000E,
        }

        [DllImport("kernel32.dll")]
        public static extern IntPtr GetModuleHandle(string lpModuleName);
        [DllImport("kernel32.dll")]
        static extern uint GetCurrentProcessId();
        [DllImport("user32.dll")]
        static extern int GetClassNameW(IntPtr hwnd, [MarshalAs(UnmanagedType.LPWStr)] StringBuilder buf, int nMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern IntPtr SendMessage(IntPtr hWnd, WM Msg, IntPtr wParam, [Out] StringBuilder lParam);

        // TODO: Not for 64-bit
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern int SendMessage(IntPtr hWnd, UInt32 Msg, int wParam, int lParam);


        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool GetCursorPos(out POINT lpPoint);
        // Or use System.Drawing.Point (Forms only)

        [DllImport("user32.dll")]
        static extern bool ScreenToClient(IntPtr hWnd, ref POINT lpPoint);

        public static POINT GetClientCursorPos(IntPtr hWnd)
        {
            POINT pt;
            bool ok = GetCursorPos(out pt);
            bool ok2 = ScreenToClient(hWnd, ref pt);
            return pt;
        }

        public static string GetWindowTextRaw(IntPtr hwnd)
        {
            // Allocate correct string length first
            int length = (int)SendMessage(hwnd, WM.GETTEXTLENGTH, IntPtr.Zero, null);
            StringBuilder sb = new StringBuilder(length + 1);
            SendMessage(hwnd, WM.GETTEXT, (IntPtr)sb.Capacity, sb);
            return sb.ToString();
        }

        static StringBuilder _buffer = new StringBuilder(65000);

        public static string GetXllName()
        {
            return (string)typeof(DnaLibrary).GetProperty("XllPath", System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic).GetValue(null, null);
        }

        public static IntPtr GetXllModuleHandle()
        {
            return GetModuleHandle(GetXllName());
        }

        public static uint GetExcelProcessId()
        {
            return GetCurrentProcessId();
        }

        public static string GetClassName(IntPtr hWnd)
        {
            _buffer.Length = 0;
            GetClassNameW(hWnd, _buffer, _buffer.Capacity);
            return _buffer.ToString();
        }

        public static string GetText(IntPtr hWnd)
        {
            // Allocate correct string length first
            int length = GetWindowTextLength(hWnd);
            var sb = new StringBuilder(length + 1);
            GetWindowText(hWnd, sb, sb.Capacity);
            return sb.ToString();
        }

        public static int GetPosFromChar(IntPtr hWnd, int ch)
        {
            return SendMessage(hWnd, 214 /*EM_POSFROMCHAR*/, ch, 0);
        }
    }
}
