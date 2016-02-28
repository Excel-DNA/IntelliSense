using System;
using System.Runtime.InteropServices;

namespace ExcelDna.IntelliSense
{
    class WinEventHook
    {

        internal delegate void WinEventDelegate(
              IntPtr hWinEventHook, WinEventHook.WinEvent eventType,
              IntPtr hWnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime);

        [DllImport("user32.dll")]
        static extern IntPtr SetWinEventHook(
              WinEvent eventMin, WinEvent eventMax, IntPtr hmodWinEventProc, WinEventDelegate lpfnWinEventProc,
              uint idProcess, uint idThread, SetWinEventHookFlags dwFlags);

        [DllImport("user32.dll")]
        static extern bool UnhookWinEvent(IntPtr hWinEventHook);

        [Flags]
        internal enum SetWinEventHookFlags : uint
        {
            WINEVENT_INCONTEXT = 4,
            WINEVENT_OUTOFCONTEXT = 0,
            WINEVENT_SKIPOWNPROCESS = 2,
            WINEVENT_SKIPOWNTHREAD = 1
        }

        internal enum WinEvent : uint
        {
            EVENT_OBJECT_CREATE = 0x8000, // hwnd ID idChild is created item
            EVENT_OBJECT_DESTROY = 0x8001, // hwnd ID idChild is destroyed item
            EVENT_OBJECT_SHOW = 0x8002, // hwnd ID idChild is shown item
            EVENT_OBJECT_HIDE = 0x8003, // hwnd ID idChild is hidden item
            EVENT_OBJECT_REORDER = 0x8004, // hwnd ID idChild is parent of zordering children
            EVENT_OBJECT_FOCUS = 0x8005, // hwnd ID idChild is focused item
            EVENT_OBJECT_SELECTION = 0x8006, // hwnd ID idChild is selected item (if only one), or idChild is OBJID_WINDOW if complex
            EVENT_OBJECT_SELECTIONADD = 0x8007, // hwnd ID idChild is item added
            EVENT_OBJECT_SELECTIONREMOVE = 0x8008, // hwnd ID idChild is item removed
            EVENT_OBJECT_SELECTIONWITHIN = 0x8009, // hwnd ID idChild is parent of changed selected items
            EVENT_OBJECT_STATECHANGE = 0x800A, // hwnd ID idChild is item w/ state change
            EVENT_OBJECT_LOCATIONCHANGE = 0x800B, // hwnd ID idChild is moved/sized item
            EVENT_OBJECT_NAMECHANGE = 0x800C, // hwnd ID idChild is item w/ name change
            EVENT_OBJECT_DESCRIPTIONCHANGE = 0x800D, // hwnd ID idChild is item w/ desc change
            EVENT_OBJECT_VALUECHANGE = 0x800E, // hwnd ID idChild is item w/ value change
            EVENT_OBJECT_PARENTCHANGE = 0x800F, // hwnd ID idChild is item w/ new parent
            EVENT_OBJECT_HELPCHANGE = 0x8010, // hwnd ID idChild is item w/ help change
            EVENT_OBJECT_DEFACTIONCHANGE = 0x8011, // hwnd ID idChild is item w/ def action change
            EVENT_OBJECT_ACCELERATORCHANGE = 0x8012, // hwnd ID idChild is item w/ keybd accel change
            EVENT_OBJECT_INVOKED = 0x8013, // hwnd ID idChild is item invoked
            EVENT_OBJECT_TEXTSELECTIONCHANGED = 0x8014, // hwnd ID idChild is item w? test selection change
            EVENT_OBJECT_CONTENTSCROLLED = 0x8015,
            EVENT_SYSTEM_ARRANGMENTPREVIEW = 0x8016,
            EVENT_OBJECT_END = 0x80FF,
            EVENT_AIA_START = 0xA000,
            EVENT_AIA_END = 0xAFFF,
       }

        readonly WinEventDelegate _procDelegate;    // So that it does not get GC'ed ...?
        readonly IntPtr _hWinEventHook;
        readonly string _xllPath;

        public WinEventHook(WinEventDelegate handler, WinEvent eventMin, WinEvent eventMax, string xllPath)
        {
            // Note : could we use another handle than one of an xll ? 
            // Thus we could avoid carrying the xll path.
            // The events are still hooked after the xll has been unloaded.
            _procDelegate = handler;
            _xllPath = xllPath;
            var xllModuleHandle = Win32Helper.GetModuleHandle(_xllPath);
            var excelProcessId = Win32Helper.GetExcelProcessId();
            _hWinEventHook = SetWinEventHook(eventMin, eventMax, xllModuleHandle, handler, excelProcessId, 0, SetWinEventHookFlags.WINEVENT_INCONTEXT);
            Logger.WinEvents.Info($"SetWinEventHook for {_xllPath}");
        }

        //public WinEventHook(WinEventDelegate handler, WinEvent eventId)
        //    : this(handler, eventId, eventId)
        //{
        //}

        public void Stop()
        {
            Logger.WinEvents.Info($"UnhookWinEvent for {_xllPath}");
            UnhookWinEvent(_hWinEventHook);
        }

        // Usage Example for EVENT_OBJECT_CREATE (http://msdn.microsoft.com/en-us/library/windows/desktop/dd318066%28v=vs.85%29.aspx)
        // var _objectCreateHook = new EventHook(OnObjectCreate, EventHook.EVENT_OBJECT_CREATE);
        // ...
        // static void OnObjectCreate(IntPtr hWinEventHook, uint eventType, IntPtr hWnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime) {
        //    if (!Win32.GetClassName(hWnd).StartsWith("ClassICareAbout"))
        //        return;
        // Note - in Console program, doesn't fire if you have a Console.ReadLine active, so use a Form
    }
}
