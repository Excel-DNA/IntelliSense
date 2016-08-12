using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;

namespace ExcelDna.IntelliSense
{
    public class WindowLocationWatcher : IDisposable
    {
        IntPtr _hWnd;
        SynchronizationContext _syncContextAuto;
        WinEventHook _windowMoveSizeHook;
        WinEventHook _windowLocationChangeHook;

        public event EventHandler LocationChanged;

        public WindowLocationWatcher(IntPtr hWnd, SynchronizationContext syncContextAuto)
        {
            _hWnd = hWnd;
            _syncContextAuto = syncContextAuto;
            _windowMoveSizeHook = new WinEventHook(WinEventHook.WinEvent.EVENT_SYSTEM_MOVESIZESTART, WinEventHook.WinEvent.EVENT_SYSTEM_MOVESIZEEND, _syncContextAuto, _hWnd);
            _windowMoveSizeHook.WinEventReceived += _windowMoveHook_WinEventReceived;
        }

        void _windowMoveHook_WinEventReceived(object sender, WinEventHook.WinEventArgs winEventArgs)
        {
#if DEBUG
            Logger.WinEvents.Verbose($"{winEventArgs.EventType} - Window {winEventArgs.WindowHandle:X} ({Win32Helper.GetClassName(winEventArgs.WindowHandle)} - Object/Child {winEventArgs.ObjectId} / {winEventArgs.ChildId} - Thread {winEventArgs.EventThreadId} at {winEventArgs.EventTimeMs}");
#endif
            if (winEventArgs.EventType == WinEventHook.WinEvent.EVENT_SYSTEM_MOVESIZESTART)
            {
                if (_windowLocationChangeHook != null)
                {
                    Debug.Fail("Unexpected move start without end");
                    _windowLocationChangeHook.Dispose();
                }
                _windowLocationChangeHook = new WinEventHook(WinEventHook.WinEvent.EVENT_OBJECT_LOCATIONCHANGE, WinEventHook.WinEvent.EVENT_OBJECT_LOCATIONCHANGE, _syncContextAuto, _hWnd);
                _windowLocationChangeHook.WinEventReceived += _windowLocationChangeHook_WinEventReceived;
            }
            else if (winEventArgs.EventType == WinEventHook.WinEvent.EVENT_SYSTEM_MOVESIZEEND)
            {
                _windowLocationChangeHook.Dispose();
                _windowLocationChangeHook = null;
            }
        }

        void _windowLocationChangeHook_WinEventReceived(object sender, WinEventHook.WinEventArgs winEventArgs)
        {
#if DEBUG
            Logger.WinEvents.Verbose($"{winEventArgs.EventType} - Window {winEventArgs.WindowHandle:X} ({Win32Helper.GetClassName(winEventArgs.WindowHandle)} - Object/Child {winEventArgs.ObjectId} / {winEventArgs.ChildId} - Thread {winEventArgs.EventThreadId} at {winEventArgs.EventTimeMs}");
#endif
            LocationChanged?.Invoke(this, EventArgs.Empty);
        }

        public void Dispose()
        {
            if (_windowMoveSizeHook != null)
            {
                _windowMoveSizeHook.Dispose();
                _windowMoveSizeHook = null;
            }
            if (_windowLocationChangeHook != null)
            {
                _windowLocationChangeHook.Dispose();
                _windowLocationChangeHook = null;
            }
        }
    }
}
