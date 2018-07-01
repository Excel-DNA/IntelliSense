using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;

namespace ExcelDna.IntelliSense
{
    // This class sets up a WinEventHook for the main Excel process - watching for a range or event types specified.
    // Events received (on the main Excel thread) are posted by the handler onto the Automation thread (via syncContextAuto)
    // NOTE: Currently we make the SetWinEventHook call on the main Excel thread, which is pumping Windows messages.
    //       In an alternative implementation we could either create our own thread that pumps Windows messages and use this for WinEvents, 
    //       or we could change our automation thread to also pump windows messages.
    class WinEventHook : IDisposable
    {
        public class WinEventArgs : EventArgs
        {
            public WinEvent EventType;
            public IntPtr WindowHandle;
            public WinEventObjectId ObjectId;
            public int ChildId;
            public uint EventThreadId;
            public uint EventTimeMs;

            public WinEventArgs(WinEvent eventType, IntPtr hWnd, WinEventObjectId idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
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
              IntPtr hWnd, WinEventObjectId idObject, int idChild, uint dwEventThread, uint dwmsEventTime);

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
            EVENT_MIN = 0x00000001,
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
            EVENT_SYSTEM_MOVESIZESTART = 0x000A, 
            EVENT_SYSTEM_MOVESIZEEND = 0x000B,  // The movement or resizing of a window has finished. This event is sent by the system, never by servers.
            EVENT_SYSTEM_MINIMIZESTART = 0x0016,    // A window object is about to be minimized. 
            EVENT_SYSTEM_MINIMIZEEND = 0x0017,      // A window object is about to be restored. 
            EVENT_SYSTEM_END = 0x00FF,
            EVENT_OBJECT_END = 0x80FF,
            EVENT_AIA_START = 0xA000,
            EVENT_AIA_END = 0xAFFF,
        }

        public enum WinEventObjectId : int
        {
            OBJID_SELF = 0,
            OBJID_SYSMENU = -1,
            OBJID_TITLEBAR = -2,
            OBJID_MENU = -3,
            OBJID_CLIENT = -4,
            OBJID_VSCROLL = -5,
            OBJID_HSCROLL = -6,
            OBJID_SIZEGRIP = -7,
            OBJID_CARET = -8,
            OBJID_CURSOR = -9,
            OBJID_ALERT = -10,
            OBJID_SOUND = -11,
            OBJID_QUERYCLASSNAMEIDX = -12,
            OBJID_NATIVEOM = -16
        }

        public event EventHandler<WinEventArgs> WinEventReceived;

        /* readonly */ IntPtr _hWinEventHook;
        readonly SynchronizationContext _syncContextAuto;
        readonly SynchronizationContext _syncContextMain;
        readonly IntPtr _hWndFilterOrZero;    // If non-zero, only these window events are processed
        readonly WinEventDelegate _handleWinEventDelegate;  // Ensures delegate that we pass to SetWinEventHook is not GC'd
        readonly WinEvent _eventMin;
        readonly WinEvent _eventMax;

        // Can be called on any thread, but installed by calling into the main thread, and will only start receiving events then
        public WinEventHook(WinEvent eventMin, WinEvent eventMax, SynchronizationContext syncContextAuto, SynchronizationContext syncContextMain, IntPtr hWndFilterOrZero)
        {
            _syncContextAuto = syncContextAuto ?? throw new ArgumentNullException(nameof(syncContextAuto));
            _syncContextMain = syncContextMain ?? throw new ArgumentNullException(nameof(syncContextMain));
            _hWndFilterOrZero = hWndFilterOrZero;
            _handleWinEventDelegate = HandleWinEvent;
            _eventMin = eventMin;
            _eventMax = eventMax;
            syncContextMain.Post(InstallWinEventHook, null);
        }

        // Must run on the main Excel thread (or another thread where Windows messages are pumped)
        void InstallWinEventHook(object _)
        {
            var excelProcessId = Win32Helper.GetExcelProcessId();
            _hWinEventHook = SetWinEventHook(_eventMin, _eventMax, IntPtr.Zero, _handleWinEventDelegate, excelProcessId, 0, SetWinEventHookFlags.WINEVENT_OUTOFCONTEXT);
            if (_hWinEventHook == IntPtr.Zero)
            {
                Logger.WinEvents.Error("SetWinEventHook failed");
                // Is SetLastError used? - SetWinEventHook documentation does not indicate so
                throw new Win32Exception("SetWinEventHook failed");
            }
            Logger.WinEvents.Info($"SetWinEventHook success on thread {Thread.CurrentThread.ManagedThreadId}");
        }

        // Must run on the same thread that InstallWinEventHook ran on
        void UninstallWinEventHook(object _)
        {
            if (_hWinEventHook == IntPtr.Zero)
            {
                Logger.WinEvents.Warn($"UnhookWinEvent unexpectedly called with no hook installed - thread {Thread.CurrentThread.ManagedThreadId}");
                return;
            }

            try
            {
                Logger.WinEvents.Info($"UnhookWinEvent called on thread {Thread.CurrentThread.ManagedThreadId}");
                bool result = UnhookWinEvent(_hWinEventHook);
                if (!result)
                {
                    // GetLastError?
                    Logger.WinEvents.Info($"UnhookWinEvent failed");
                }
                else
                {
                    Logger.WinEvents.Info("UnhookWinEvent success");
                }
            }
            catch (Exception ex)
            {
                Logger.WinEvents.Warn($"UnhookWinEvent Exception {ex}");
            }
            finally
            {
                _hWinEventHook = IntPtr.Zero;
            }
        }

        // This runs on the Excel main thread (usually, not always) - get off quickly
        void HandleWinEvent(IntPtr hWinEventHook, WinEvent eventType, IntPtr hWnd,
                            WinEventObjectId idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
        {
            // Debug.Print($"++++++++++++++ WinEvent Received: {eventType} on thread {Thread.CurrentThread.ManagedThreadId} from thread {dwEventThread} +++++++++++++++++++++++++++");
            try
            {
                if (_hWndFilterOrZero != IntPtr.Zero && hWnd != _hWndFilterOrZero)
                    return;

                if (!IsSupportedWinEvent(eventType))
                    return;

                // CONSIDER: We might add some filtering here... maybe only interested in some of the window / event combinations
                _syncContextAuto.Post(OnWinEventReceived, new WinEventArgs(eventType, hWnd, idObject, idChild, dwEventThread, dwmsEventTime));
            }
            catch (Exception ex)
            {
                Logger.WinEvents.Warn($"HandleWinEvent Exception {ex}");
            }
        }

        // A quick filter that runs on the Excel main thread (or other thread handling the WinEvent)
        bool IsSupportedWinEvent(WinEvent winEvent)
        {
            return winEvent == WinEvent.EVENT_OBJECT_CREATE ||
                   winEvent == WinEvent.EVENT_OBJECT_DESTROY ||
                   winEvent == WinEvent.EVENT_OBJECT_SHOW ||
                   winEvent == WinEvent.EVENT_OBJECT_HIDE ||
                   winEvent == WinEvent.EVENT_OBJECT_FOCUS ||
                   winEvent == WinEvent.EVENT_OBJECT_LOCATIONCHANGE ||   // Only for the on-demand hook
                   winEvent == WinEvent.EVENT_OBJECT_SELECTION ||           // Only for the PopupList
                   winEvent == WinEvent.EVENT_OBJECT_TEXTSELECTIONCHANGED;
        }

        // Runs on our Automation thread (via SyncContext passed into the constructor)
        // CONSIDER: Performance impact of logging (including GetClassName) here 
        void OnWinEventReceived(object winEventArgsObj)
        {
            var winEventArgs = (WinEventArgs)winEventArgsObj;
#if DEBUG
            if (winEventArgs.ObjectId != WinEventObjectId.OBJID_CURSOR)
                Logger.WinEvents.Verbose($"{winEventArgs.EventType} - Window {winEventArgs.WindowHandle:X} ({Win32Helper.GetClassName(winEventArgs.WindowHandle)} - Object/Child {winEventArgs.ObjectId} / {winEventArgs.ChildId} - Thread {winEventArgs.EventThreadId} at {winEventArgs.EventTimeMs}");
#endif
            WinEventReceived?.Invoke(this, winEventArgs);
        }

        #region IDisposable Support
        // Must be called on the main thread
        public void Dispose()
        {
            Debug.Assert(Thread.CurrentThread.ManagedThreadId == 1);
            Logger.WinEvents.Info($"WinEventHook Dispose on thread {Thread.CurrentThread.ManagedThreadId}");
            UninstallWinEventHook(null);
        }
        #endregion
    }
}
