using System;
using System.Diagnostics;
using System.Threading;

namespace ExcelDna.IntelliSense
{
    public class WindowLocationWatcher : IDisposable
    {
        IntPtr _hWnd;
        SynchronizationContext _syncContextAuto;
        SynchronizationContext _syncContextMain;
        WinEventHook _windowMoveSizeHook;
        WinEventHook _locationChangeEventHook;

        public event EventHandler LocationChanged;

        // NOTE: An earlier attempt was to monitor LOCATIONCHANGE only between EVENT_SYSTEM_MOVESIZESTART and EVENT_SYSTEM_MOVESIZEEND
        //       This nearly worked, and meant we were watching many fewer events ...
        //       ...but we missed some of the resizing events for the window, leaving our tooltip stranded.
        //       So until we can find a workaround for that (perhaps a timer would work fine for this), we watch all the LOCATIONCHANGE events.
        public WindowLocationWatcher(IntPtr hWnd, SynchronizationContext syncContextAuto, SynchronizationContext syncContextMain)
        {
            _hWnd = hWnd;
            _syncContextAuto = syncContextAuto;
            _syncContextMain = syncContextMain;
            _windowMoveSizeHook = new WinEventHook(WinEventHook.WinEvent.EVENT_SYSTEM_MOVESIZESTART, WinEventHook.WinEvent.EVENT_SYSTEM_MOVESIZEEND, _syncContextAuto, syncContextMain, _hWnd);
            _windowMoveSizeHook.WinEventReceived += _windowMoveSizeHook_WinEventReceived;

            SetUpLocationChangeEventListener();
        }

        void SetUpLocationChangeEventListener()
        {
            // NB: Including the next event 'EVENT_OBJECT_LOCATIONCHANGE (0x800B = 32779)' will cause the Excel main window to lag when dragging.
            // This drag issue seems to have been introduced with an Office update around November 2022.
            // To workaround this, we unhook from this event upon encountering EVENT_SYSTEM_MOVESIZESTART and then hook again upon encountering
            // EVENT_SYSTEM_MOVESIZEEND (see UnhookFromLocationChangeUponDraggingExcelMainWindow).
            _locationChangeEventHook = new WinEventHook(WinEventHook.WinEvent.EVENT_OBJECT_LOCATIONCHANGE, WinEventHook.WinEvent.EVENT_OBJECT_LOCATIONCHANGE, _syncContextAuto, _syncContextMain, IntPtr.Zero);
            _locationChangeEventHook.WinEventReceived += _windowMoveSizeHook_WinEventReceived;
        }

        // This allows us to temporarily stop listening to EVENT_OBJECT_LOCATIONCHANGE events when the user is dragging the Excel main window.
        // Otherwise we are going to bump into https://github.com/Excel-DNA/IntelliSense/issues/123. The rest of the time we need to stay
        // hooked to EVENT_OBJECT_LOCATIONCHANGE for IntelliSense to work correctly (see https://github.com/Excel-DNA/IntelliSense/issues/124).
        void UnhookFromLocationChangeUponDraggingExcelMainWindow(WinEventHook.WinEventArgs e)
        {
            if (e.EventType == WinEventHook.WinEvent.EVENT_SYSTEM_MOVESIZESTART)
            {
                _syncContextMain.Post(_ => _locationChangeEventHook?.Dispose(), null);
            }

            if (e.EventType == WinEventHook.WinEvent.EVENT_SYSTEM_MOVESIZEEND)
            {
                _syncContextMain.Post(_ => SetUpLocationChangeEventListener(), null);
            }
        }

        void _windowMoveSizeHook_WinEventReceived(object sender, WinEventHook.WinEventArgs winEventArgs)
        {
#if DEBUG
            Logger.WinEvents.Verbose($"{winEventArgs.EventType} - Window {winEventArgs.WindowHandle:X} ({Win32Helper.GetClassName(winEventArgs.WindowHandle)} - Object/Child {winEventArgs.ObjectId} / {winEventArgs.ChildId} - Thread {winEventArgs.EventThreadId} at {winEventArgs.EventTimeMs}");
#endif

            UnhookFromLocationChangeUponDraggingExcelMainWindow(winEventArgs);

            LocationChanged?.Invoke(this, EventArgs.Empty);
        }

        // Runs on the Main thread, perhaps during shutdown
        public void Dispose()
        {
            Debug.Assert(Thread.CurrentThread.ManagedThreadId == 1);
            if (_windowMoveSizeHook != null)
            {
                _windowMoveSizeHook.Dispose();
                _windowMoveSizeHook = null;
            }

            if (_locationChangeEventHook != null)
            {
                _locationChangeEventHook.Dispose();
                _locationChangeEventHook = null;
            }
        }
    }
}
