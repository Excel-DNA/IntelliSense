using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace ExcelDna.IntelliSense
{
    class KeyboardEventHook : IDisposable
    {
        public class KeyboardEventArgs : EventArgs
        {
            public readonly Keys Key;
            public readonly KeyboardLParam LParam;

            public KeyboardEventArgs(Keys key, IntPtr lParam)
            {
                Key = key;
                LParam = new KeyboardLParam() { Raw = lParam };
            }
        }

        public struct KeyboardLParam
        {
            public IntPtr Raw;

            public readonly long Value => Raw.ToInt64();

            public readonly ushort RepeatCount => (ushort)(Value & 0xFFFF);
            public readonly byte ScanCode => (byte)((Value >> 16) & 0xFF);
            public readonly bool IsExtended => (Value & (1 << 24)) != 0;
            public readonly bool ContextCodeAlt => (Value & (1 << 29)) != 0;
            public readonly bool WasDown => (Value & (1 << 30)) != 0; // 1 = was down before this message
            public readonly bool IsReleased => (Value & (1 << 31)) != 0; // 1 = key released (keyup)
        }

        internal delegate IntPtr KeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr SetWindowsHookEx(WH idHook, KeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll")]
        static extern uint GetCurrentThreadId();

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern bool UnhookWindowsHookEx(IntPtr hHook);

        internal enum WH : int
        {
            KEYBOARD = 2,
            KEYBOARD_LL = 13,
        }

        public event EventHandler<KeyboardEventArgs> KeyboardEventReceived;

        public delegate bool ShouldSwallowDelegate();
        public event ShouldSwallowDelegate ShouldSwallow;

        /* readonly */
        IntPtr _hKeyboardEventHook;

        // readonly SynchronizationContext _syncContextAuto;
        readonly SynchronizationContext _syncContextMain;
        readonly KeyboardProc _handleKeyboardEventDelegate; // Ensures delegate that we pass to SetWindowsHookEx is not GC'd

        // Can be called on any thread, but installed by calling into the main thread, and will only start receiving events then
        public KeyboardEventHook(SynchronizationContext syncContextMain)
        {
            // _syncContextAuto = syncContextAuto ?? throw new ArgumentNullException(nameof(syncContextAuto));
            _syncContextMain = syncContextMain ?? throw new ArgumentNullException(nameof(syncContextMain));
            _handleKeyboardEventDelegate = HandleKeyboardEvent;
            syncContextMain.Post(InstallKeyboardEventHook, null);
        }

        // Must run on the main Excel thread (or another thread where Windows messages are pumped)
        void InstallKeyboardEventHook(object _)
        {
            uint tid = GetCurrentThreadId();
            _hKeyboardEventHook = SetWindowsHookEx(WH.KEYBOARD, _handleKeyboardEventDelegate, IntPtr.Zero, tid);
            if (_hKeyboardEventHook == IntPtr.Zero)
            {
                Logger.WinEvents.Error("SetWindowsHookEx failed");
                // Is SetLastError used? - SetWindowsHookEx documentation does not indicate so
                throw new Win32Exception("SetWindowsHookEx failed");
            }
            Logger.WinEvents.Info($"SetWindowsHookEx success on thread {Thread.CurrentThread.ManagedThreadId}");
        }

        // This runs on the Excel main thread (usually, not always) - get off quickly
        IntPtr HandleKeyboardEvent(int nCode, IntPtr wParam, IntPtr lParam)
        {
            Debug.Print(
                $"++++++++++++++ HandleKeyboardEvent Received: {wParam} on thread {Thread.CurrentThread.ManagedThreadId} +++++++++++++++++++++++++++"
            );
            try
            {
                var vkCode = wParam.ToInt32(); // wParam is the virtual-key code here
                Keys key = (Keys)vkCode;
                var handlers = ShouldSwallow?.GetInvocationList();
                if (nCode < 0 || !IsSupportedKeyboardEvent(key) || handlers == null)
                {
                    return CallNextHookEx(_hKeyboardEventHook, nCode, wParam, lParam);
                }

                bool shouldSwallow = false;
                foreach (var d in handlers)
                {
                    var handler = (ShouldSwallowDelegate)d;

                    // Invoke directly
                    shouldSwallow |= handler.Invoke();
                }

                if (!shouldSwallow)
                {
                    return CallNextHookEx(_hKeyboardEventHook, nCode, wParam, lParam);
                }

                // CONSIDER: We might add some filtering here... maybe only interested in some of the window / event combinations
                _syncContextMain.Post(OnKeyboardEventReceived, new KeyboardEventArgs(key, lParam));
                return (IntPtr)1; // swallow event
            }
            catch (Exception ex)
            {
                Logger.WinEvents.Warn($"HandleKeyboardEvent Exception {ex}");
                return (IntPtr)0;
            }
        }

        void UninstallKeyboardEventHook(object _)
        {
            if (_hKeyboardEventHook == IntPtr.Zero)
            {
                Logger.WinEvents.Warn(
                    $"UninstallKeyboardEvent unexpectedly called with no hook installed - thread {Thread.CurrentThread.ManagedThreadId}"
                );
                return;
            }

            try
            {
                Logger.WinEvents.Info(
                    $"UninstallKeyboardEvent called on thread {Thread.CurrentThread.ManagedThreadId}"
                );
                bool result = UnhookWindowsHookEx(_hKeyboardEventHook);
                if (!result)
                {
                    // GetLastError?
                    Logger.WinEvents.Info($"UninstallKeyboardEvent failed");
                }
                else
                {
                    Logger.WinEvents.Info("UninstallKeyboardEvent success");
                }
            }
            catch (Exception ex)
            {
                Logger.WinEvents.Warn($"UninstallKeyboardEvent Exception {ex}");
            }
            finally
            {
                _hKeyboardEventHook = IntPtr.Zero;
            }
        }

        bool IsSupportedKeyboardEvent(Keys key)
        {
            return key is Keys.Down or Keys.Up or Keys.Escape or Keys.Tab;
        }

        void OnKeyboardEventReceived(object kbEventArgsObj)
        {
            var kbEventArgs = (KeyboardEventArgs)kbEventArgsObj;

            KeyboardEventReceived?.Invoke(this, kbEventArgs);
        }

        #region IDisposable Support
        // Must be called on the main thread
        public void Dispose()
        {
            //Debug.Assert(Thread.CurrentThread.ManagedThreadId == 1);
            Logger.WinEvents.Info($"KeyboardEventHook Dispose on thread {Thread.CurrentThread.ManagedThreadId}");
            UninstallKeyboardEventHook(null);
        }
        #endregion
    }
}
