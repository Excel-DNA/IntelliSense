using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace ExcelDna.IntelliSense
{
    class XlCall
    {
        [DllImport("XLCALL32.DLL")]
        static extern int LPenHelper(int wCode, ref FmlaInfo fmlaInfo); // 'long' return value means 'int' in C (so why different?)

        const int xlSpecial = 0x4000;
        const int xlGetFmlaInfo = (14 | xlSpecial);

        // Edit Modes
        const int xlModeReady = 0;	// not in edit mode
        const int xlModeEnter = 1;	// enter mode
        const int xlModeEdit  = 2;	// edit mode
        const int xlModePoint = 4;	// point mode

        [StructLayout(LayoutKind.Sequential)]
        struct FmlaInfo
        {
            public int wPointMode;    // current edit mode.  0 => rest of struct undefined
            public int cch;           // count of characters in formula
            public IntPtr lpch;       // pointer to formula characters.  READ ONLY!!!
            public int ichFirst;      // char offset to start of selection
            public int ichLast;       // char offset to end of selection (may be > cch)
            public int ichCaret;      // char offset to blinking caret
        }

        // Returns null if not in edit mode
        public static string GetFormulaEditPrefix()
        {
            var fmlaInfo = new FmlaInfo();
            var result = LPenHelper(xlGetFmlaInfo, ref fmlaInfo);
            if (result != 0)
            {
                throw new InvalidOperationException("LPenHelper failed. Result: " + result);
            }
            if (fmlaInfo.wPointMode == xlModeReady) return null; 
            
            Debug.Print("LPenHelper Status: PointMode: {0}, Formula: {1}, First: {2}, Last: {3}, Caret: {4}",
                fmlaInfo.wPointMode, Marshal.PtrToStringUni(fmlaInfo.lpch, fmlaInfo.cch), fmlaInfo.ichFirst, fmlaInfo.ichLast, fmlaInfo.ichCaret);

            var prefixLen = Math.Min(Math.Max(fmlaInfo.ichCaret, fmlaInfo.ichLast), fmlaInfo.cch);  // I've never seen ichLast > cch !?
            return Marshal.PtrToStringUni(fmlaInfo.lpch, prefixLen);
        }
    }
}
