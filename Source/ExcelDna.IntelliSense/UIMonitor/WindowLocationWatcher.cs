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
        //       (for the purpose of moving the tooltip to the correct position when the user moves the Excel main window)
        //       This nearly worked, and meant we were watching many fewer events ...
        //       ...but we missed some of the resizing events for the window, leaving our tooltip stranded.
        //       We then started to watch all the LOCATIONCHANGE events, but it caused the Excel main window to lag when dragging.
        //       (This drag issue seems to have been introduced with an Office update around November 2022)
        //       So until we can find a workaround for that (perhaps a timer would work fine for this), we decided not to bother
        //       with tracking the tooltip position (we still update it as soon as the Excel main window moving ends).
        //       We still need to watch the LOCATIONCHANGE events, otherwise the tooltip is not shown at all in some cases.
        //       To workaround the Excel main window lagging, we unhook from LOCATIONCHANGE upon encountering EVENT_SYSTEM_MOVESIZESTART
        //       and then hook again upon encountering EVENT_SYSTEM_MOVESIZEEND (see UnhookFromLocationChangeUponDraggingExcelMainWindow).
        public WindowLocationWatcher(IntPtr hWnd, SynchronizationContext syncContextAuto, SynchronizationContext syncContextMain)
        {
            Debug.Print("===========>>>>>>>>  WindowsLocationWatcher      <<<<<<<<<<<<===========");

            _hWnd = hWnd;
            _syncContextAuto = syncContextAuto;
            _syncContextMain = syncContextMain;

            _syncContextMain.Post(_ => SetUpHooks(), null);
        }

        void SetUpHooks()
        {
            _windowMoveSizeHook = new WinEventHook(WinEventHook.WinEvent.EVENT_SYSTEM_MOVESIZESTART, WinEventHook.WinEvent.EVENT_SYSTEM_MOVESIZEEND, _syncContextAuto, _syncContextMain, _hWnd);
            // _windowMoveSizeHook.WinEventReceived += _windowMoveSizeHook_WinEventReceived;
            _windowMoveSizeHook.DirectCallEvents[WinEventHook.WinEvent.EVENT_SYSTEM_MOVESIZESTART] = _ => 
            {
                // Runs on the main thread
                ClearLocationChangeEventListener();
            };
            _windowMoveSizeHook.DirectCallEvents[WinEventHook.WinEvent.EVENT_SYSTEM_MOVESIZEEND] = _ =>
            {
                // Runs on the main thread
                SetUpLocationChangeEventListener();
                _syncContextAuto.Post(__ => NotifyLocationChanged(), null);
            };

            SetUpLocationChangeEventListener();
        }

        void SetUpLocationChangeEventListener()
        {
            Debug.Assert(Thread.CurrentThread.ManagedThreadId == 1);
            Debug.Assert(_locationChangeEventHook == null);
            _locationChangeEventHook = new WinEventHook(WinEventHook.WinEvent.EVENT_OBJECT_LOCATIONCHANGE, WinEventHook.WinEvent.EVENT_OBJECT_LOCATIONCHANGE, _syncContextAuto, _syncContextMain, IntPtr.Zero);
            _locationChangeEventHook.WinEventReceived += _windowMoveSizeHook_WinEventReceived;
        }

        void ClearLocationChangeEventListener()
        {
            Debug.Assert(Thread.CurrentThread.ManagedThreadId == 1);
            if (_locationChangeEventHook != null)
            {
                _locationChangeEventHook.Dispose();
                _locationChangeEventHook = null;
            }
        }

        // This allows us to temporarily stop listening to EVENT_OBJECT_LOCATIONCHANGE events when the user is dragging the Excel main window.
        // Otherwise we are going to bump into https://github.com/Excel-DNA/IntelliSense/issues/123. The rest of the time we need to stay
        // hooked to EVENT_OBJECT_LOCATIONCHANGE for IntelliSense to work correctly (see https://github.com/Excel-DNA/IntelliSense/issues/124).
        void UnhookFromLocationChangeUponDraggingExcelMainWindow(WinEventHook.WinEventArgs e)
        {
            if (e.EventType == WinEventHook.WinEvent.EVENT_SYSTEM_MOVESIZESTART)
            {
                Debug.WriteLine("===========>>>>>>>>  EVENT_SYSTEM_MOVESIZESTART   <<<<<<<<<<<<===========");
                _syncContextMain.Post(_ => ClearLocationChangeEventListener(), null);
            }

            if (e.EventType == WinEventHook.WinEvent.EVENT_SYSTEM_MOVESIZEEND)
            {
                Debug.WriteLine("===========>>>>>>>>  EVENT_SYSTEM_MOVESIZEEND   <<<<<<<<<<<<===========");
                _syncContextMain.Post(_ => SetUpLocationChangeEventListener(), null);
            }
        }

        // Runs on the automation thread
        void _windowMoveSizeHook_WinEventReceived(object sender, WinEventHook.WinEventArgs winEventArgs)
        {
#if DEBUG
            Logger.WinEvents.Verbose($"{winEventArgs.EventType} - Window {winEventArgs.WindowHandle:X} ({Win32Helper.GetClassName(winEventArgs.WindowHandle)} - Object/Child {winEventArgs.ObjectId} / {winEventArgs.ChildId} - Thread {winEventArgs.EventThreadId} at {winEventArgs.EventTimeMs}");
#endif

            UnhookFromLocationChangeUponDraggingExcelMainWindow(winEventArgs);

            NotifyLocationChanged();
        }

        void NotifyLocationChanged()
        {
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
            Debug.Print("===========>>>>>>>>  WindowsLocationWatcher  Disposed    <<<<<<<<<<<<===========");
        }
    }
}
