using ExcelDna.Integration;
using System;
using System.Diagnostics;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;

namespace ExcelDna.IntelliSense
{
    class XlCallHelper
    {
        // This give a global mechanism to indicate shutdown as early as possible
        // Maybe helps with debugging...
        static bool _shutdownStarted = false;
        public static void ShutdownStarted()
        { 
            _shutdownStarted = true;
        }

        // This call must be made on the main thread
        // Returns null if not in edit mode
        public static string GetFormulaEditPrefix()
        {
            if (_shutdownStarted)
                return null;

            try
            {
                var fmlaInfo = new XlCall.FmlaInfo();
                var result = XlCall.PenHelper(XlCall.xlGetFmlaInfo, ref fmlaInfo);
                if (result != 0)
                {
                    Logger.WindowWatcher.Warn($"LPenHelper Failed. Result: {result}");
                    // This exception is poorly handled when it happens in the SyncMacro
                    // We only expect the error to happen during shutdown, in which case we might as well
                    // handle this asif no formula is being edited
                    // throw new InvalidOperationException("LPenHelper Failed. Result: " + result);
                    return null;
                }
                if (fmlaInfo.wPointMode == XlCall.xlModeReady)
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
            catch (Exception ex)
            {
                // Some unexpected error - for now we log as an error and re-throw
                Logger.WindowWatcher.Error(ex, "LPenHelper - Unexpected Error");
                throw;
            }
        }
    }
}
