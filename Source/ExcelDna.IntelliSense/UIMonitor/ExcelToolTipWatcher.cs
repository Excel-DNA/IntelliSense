using System;
using System.Collections.Generic;
using System.Threading;
using System.Windows;

namespace ExcelDna.IntelliSense
{
    class ExcelToolTipWatcher : IDisposable
    {
        public enum ToolTipChangeType
        {
            Show,
            Hide,
            // Move
        }

        public class ToolTipChangeEventArgs : EventArgs
        {
            public ToolTipChangeType ChangeType { get; private set; }
            public IntPtr Handle { get; private set; }

            public ToolTipChangeEventArgs(ToolTipChangeType changeType, IntPtr handle)
            {
                ChangeType = changeType;
                Handle = handle;
            }

            public override string ToString() => $"{ChangeType}:0x{Handle:x}";
        }

        // NOTE: Event will always be raised on our automation thread
        public event EventHandler<ToolTipChangeEventArgs> ToolTipChanged;  // Either text or location

        // CONSIDER: What should this look like?
        public IEnumerable<IntPtr> GetToolTips() => _toolTips;
        // CONSIDER: Rather a Stack? Check the assumption that Hide happens in reverse order
        HashSet<IntPtr> _toolTips = new HashSet<IntPtr>();
        SynchronizationContext _syncContextAuto; // Not used... 
        WindowWatcher _windowWatcher;

        public ExcelToolTipWatcher (WindowWatcher windowWatcher, SynchronizationContext syncContextAuto)
        {
            _syncContextAuto = syncContextAuto;
            _windowWatcher = windowWatcher;
            _windowWatcher.ExcelToolTipWindowChanged += _windowWatcher_ExcelToolTipWindowChanged;
        }

        // Runs on our automation thread
        void _windowWatcher_ExcelToolTipWindowChanged(object sender, WindowWatcher.WindowChangedEventArgs e)
        {
            switch (e.Type)
            {
                case WindowWatcher.WindowChangedEventArgs.ChangeType.Show:
                    if (_toolTips.Add(e.WindowHandle))
                        ToolTipChanged?.Invoke(this, new ToolTipChangeEventArgs(ToolTipChangeType.Show, e.WindowHandle));
                    break;
                case WindowWatcher.WindowChangedEventArgs.ChangeType.Hide:
                case WindowWatcher.WindowChangedEventArgs.ChangeType.Destroy:
                    if (_toolTips.Remove(e.WindowHandle))
                        ToolTipChanged?.Invoke(this, new ToolTipChangeEventArgs(ToolTipChangeType.Hide, e.WindowHandle));
                    break;
                case WindowWatcher.WindowChangedEventArgs.ChangeType.Create:
                case WindowWatcher.WindowChangedEventArgs.ChangeType.Focus:
                case WindowWatcher.WindowChangedEventArgs.ChangeType.Unfocus:
                default:
                    // Ignoring these....
                    break;
            }
        }

        public void Dispose()
        {
            Logger.WindowWatcher.Info($"ExcelToolTip Dispose Begin");
            _windowWatcher.ExcelToolTipWindowChanged -= _windowWatcher_ExcelToolTipWindowChanged;
            _windowWatcher = null;

            //_syncContextAuto.Send(delegate
            //{
            //    Debug.Print("Disposing ExcelToolTipWatcher - In Automation context");
            //}, null);
            Logger.WindowWatcher.Info($"ExcelToolTip Dispose End");
        }
    }
}
