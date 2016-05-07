using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Automation;

namespace ExcelDna.IntelliSense
{
    // Makes windows resizable
    class WindowResizer
    {
                // CONSIDER: This will run on our automation thread
        //           Should be OK to call MoveWindow from there - it just posts messages to the window.
        void TestMoveWindow(AutomationElement listWindow, int xOffset, int yOffset)
        {
            var hwndList = (IntPtr)(int)(listWindow.GetCurrentPropertyValue(AutomationElement.NativeWindowHandleProperty));
            var listRect = (Rect)listWindow.GetCurrentPropertyValue(AutomationElement.BoundingRectangleProperty);
            Debug.Print("Moving from {0}, {1}", listRect.X, listRect.Y);
            Win32Helper.MoveWindow(hwndList, (int)listRect.X - xOffset, (int)listRect.Y - yOffset, (int)listRect.Width, 2 * (int)listRect.Height, true);
        }

    }
}
