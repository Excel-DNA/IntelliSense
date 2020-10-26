using System;
using System.Timers;

namespace ExcelDna.IntelliSense.Util
{
    /// <summary>
    /// Upon a signal, executes the specified action after the specified delay.
    /// If any other signal arrives during the waiting period, the delay interval begins anew.
    /// </summary>
    internal class RenewableDelayExecutor : IDisposable
    {
        private readonly Timer _timer;
        private readonly Action _action;
        private readonly int _debounceIntervalMilliseconds;

        public bool IsDisposed { get; private set; }

        public RenewableDelayExecutor(int debounceIntervalMilliseconds, Action action)
        {
            _action = action;
            _debounceIntervalMilliseconds = debounceIntervalMilliseconds;
            _timer = new Timer
            {
                AutoReset = false,
                Interval = _debounceIntervalMilliseconds,
            };

            _timer.Elapsed += OnTimerElapsed;
        }

        private void OnTimerElapsed(object sender, ElapsedEventArgs e)
        {
            if (IsDisposed)
            {
                return;
            }

            _action();
        }

        public void Signal()
        {
            _timer.Stop();
            _timer.Start();
        }

        private void Dispose(bool isDisposing)
        {
            IsDisposed = true;

            _timer.Elapsed -= OnTimerElapsed;
            _timer.Dispose();
        }

        public void Dispose() => Dispose(true);
    }
}