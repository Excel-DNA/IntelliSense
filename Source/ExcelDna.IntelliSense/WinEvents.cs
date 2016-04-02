using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Threading;

namespace ExcelDna.IntelliSense
{
    // This class sets up a WinEventHook for the main Excel process - watching for a range or event types specified.
    // Events received (on the main Excel thread) are posted onto the Automation thread (via syncContextAuto)
    class WinEventHook : IDisposable
    {
        public class WinEventArgs : EventArgs
        {
            public WinEvent EventType;
            public IntPtr WindowHandle;
            public int ObjectId;
            public int ChildId;
            public uint EventThreadId;
            public uint EventTimeMs;

            public WinEventArgs(WinEvent eventType, IntPtr hWnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
            {
                EventType = eventType;
                WindowHandle = hWnd;
                ObjectId = idObject;
                ChildId = idChild;
                EventThreadId = dwEventThread;
                EventTimeMs = dwmsEventTime;
            }
        }

        delegate void WinEventDelegate(
              IntPtr hWinEventHook, WinEventHook.WinEvent eventType,
              IntPtr hWnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime);

        [DllImport("user32.dll")]
        static extern IntPtr SetWinEventHook(
              WinEvent eventMin, WinEvent eventMax, IntPtr hmodWinEventProc, WinEventDelegate lpfnWinEventProc,
              uint idProcess, uint idThread, SetWinEventHookFlags dwFlags);

        [DllImport("user32.dll")]
        static extern bool UnhookWinEvent(IntPtr hWinEventHook);

        [Flags]
        enum SetWinEventHookFlags : uint
        {
            WINEVENT_INCONTEXT = 4,
            WINEVENT_OUTOFCONTEXT = 0,
            WINEVENT_SKIPOWNPROCESS = 2,
            WINEVENT_SKIPOWNTHREAD = 1
        }

        public enum WinEvent : uint
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

        public event EventHandler<WinEventArgs> WinEventReceived;

        readonly IntPtr _hWinEventHook;
        readonly SynchronizationContext _syncContextAuto;
        readonly WinEventDelegate _handleWinEventDelegate;  // Ensures delegate that we pass to SetWinEventHook is not GC'd

        public WinEventHook(WinEvent eventMin, WinEvent eventMax, SynchronizationContext syncContextAuto)
        {
            if (syncContextAuto == null)
                throw new ArgumentNullException(nameof(syncContextAuto));
            _syncContextAuto = syncContextAuto;
            var xllModuleHandle = Win32Helper.GetXllModuleHandle();
            var excelProcessId = Win32Helper.GetExcelProcessId();
            _handleWinEventDelegate = HandleWinEvent;
            _hWinEventHook = SetWinEventHook(eventMin, eventMax, xllModuleHandle, _handleWinEventDelegate, excelProcessId, 0, SetWinEventHookFlags.WINEVENT_INCONTEXT);
            if (_hWinEventHook == IntPtr.Zero)
            {
                Logger.WinEvents.Error("SetWinEventHook failed");
                throw new Win32Exception("SetWinEventHook failed");
            }
            Logger.WinEvents.Info($"SetWinEventHook success");
        }

        // This runs on the Excel main thread - get off quickly
        void HandleWinEvent(IntPtr hWinEventHook, WinEvent eventType, IntPtr hWnd, 
                        int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
        {
            // CONSIDER: We might add some filtering here... maybe only interested in some of the window / event combinations
            _syncContextAuto.Post(OnWinEventReceived, new WinEventArgs(eventType, hWnd, idObject, idChild, dwEventThread, dwmsEventTime));
        }

        void OnWinEventReceived(object winEventArgsObj)
        {
            var winEventArgs = (WinEventArgs)winEventArgsObj;
            Logger.WinEvents.Verbose($"{winEventArgs.EventType} - Window {winEventArgs.WindowHandle:X} ({Win32Helper.GetClassName(winEventArgs.WindowHandle)} - Object/Child {winEventArgs.ObjectId} / {winEventArgs.ChildId} - Thread {winEventArgs.EventThreadId} at {winEventArgs.EventTimeMs}");
            WinEventReceived?.Invoke(this, winEventArgs);
        }

        #region IDisposable Support
        bool disposedValue = false; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects).
                }

                Logger.WinEvents.Info($"UnhookWinEvent");
                UnhookWinEvent(_hWinEventHook);

                disposedValue = true;
            }
        }

         ~WinEventHook()
        {
           // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
           Dispose(false);
         }

        // This code added to correctly implement the disposable pattern.
        public void Dispose()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        #endregion
    }
}
