using System;
using System.Windows.Forms;

namespace ExcelDna.IntelliSense;

internal class Win32Window : IWin32Window
{
    public IntPtr Handle { get; private set; }

    public Win32Window(IntPtr handle)
    {
        Handle = handle;
    }
}
