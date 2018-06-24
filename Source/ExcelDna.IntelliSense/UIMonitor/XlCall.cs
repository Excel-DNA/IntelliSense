using System;
using System.Diagnostics;
using System.Runtime.ExceptionServices;
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

        // This give a global mechanism to indicate shutdown as early as possible
        // Maybe helps with debugging...
        static bool _shutdownStarted = false;
        public static void ShutdownStarted()
        { 
            _shutdownStarted = true;
        }

        // Returns null if not in edit mode
        // TODO: What do we know about the threading constraints on this call?

        // NOTE: I've only seen this crash during shutdown
        //       We enable handling of the AccessViolation as best we can
        [HandleProcessCorruptedStateExceptions]
        public static string GetFormulaEditPrefix()
        {
            if (_shutdownStarted)
                return null;

            try
            {
                var fmlaInfo = new FmlaInfo();
                var result = LPenHelper(xlGetFmlaInfo, ref fmlaInfo);
                if (result != 0)
                {
                    Logger.WindowWatcher.Warn($"LPenHelper Failed. Result: {result}");
                    // This exception is poorly handled when it happens in the SyncMacro
                    // We only expect the error to happen during shutdown, in which case we might as well
                    // handle this asif no formula is being edited
                    // throw new InvalidOperationException("LPenHelper Failed. Result: " + result);
                    return null;
                }
                if (fmlaInfo.wPointMode == xlModeReady)
                {
                    Logger.WindowWatcher.Verbose($"LPenHelper PointMode Ready");
                    return null;
                }

                // We seem to mis-track in the click case when wPointMode changed from xlModeEnter to xlModeEdit

                Logger.WindowWatcher.Verbose("LPenHelper Status: PointMode: {0}, Formula: {1}, First: {2}, Last: {3}, Caret: {4}",
                    fmlaInfo.wPointMode, Marshal.PtrToStringUni(fmlaInfo.lpch, fmlaInfo.cch), fmlaInfo.ichFirst, fmlaInfo.ichLast, fmlaInfo.ichCaret);

                var prefixLen = Math.Min(Math.Max(fmlaInfo.ichCaret, fmlaInfo.ichLast), fmlaInfo.cch);  // I've never seen ichLast > cch !?
                return Marshal.PtrToStringUni(fmlaInfo.lpch, prefixLen);
            }
            catch (AccessViolationException)
            {
                Logger.WindowWatcher.Warn("LPenHelper - Access Violation!");
                return null;
            }
            catch (Exception ex)
            {
                // Some unexpected error - for now we log as an error and re-throw
                Logger.WindowWatcher.Error(ex, "LPenHelper - Unexpected Error");
                throw;
            }
        }
    }
}
