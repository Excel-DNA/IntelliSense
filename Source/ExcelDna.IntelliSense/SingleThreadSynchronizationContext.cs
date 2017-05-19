using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Threading;

namespace ExcelDna.IntelliSense
{
    // This code is from Stephen Toub's blog post on AsyncPump: http://blogs.msdn.com/b/pfxteam/archive/2012/01/20/10259049.aspx?PageIndex=2#comments

    /// <summary>Provides a SynchronizationContext that's single-threaded.</summary>
    sealed class SingleThreadSynchronizationContext : SynchronizationContext
    {
        /// <summary>The queue of work items.</summary>
        readonly BlockingCollection<KeyValuePair<SendOrPostCallback, object>> _queue =
            new BlockingCollection<KeyValuePair<SendOrPostCallback, object>>();
        int _threadId = 0;

        // /// <summary>The processing thread.</summary>
        // readonly Thread m_thread = Thread.CurrentThread;

        /// <summary>Dispatches an asynchronous message to the synchronization context.</summary>
        /// <param name="d">The System.Threading.SendOrPostCallback delegate to call.</param>
        /// <param name="state">The object passed to the delegate.</param>
        public override void Post(SendOrPostCallback d, object state)
        {
            if (d == null) throw new ArgumentNullException("d");
            _queue.Add(new KeyValuePair<SendOrPostCallback, object>(d, state));
        }

        public override void Send(SendOrPostCallback d, object state)
        {
            if (Thread.CurrentThread.ManagedThreadId == _threadId)
            {
                d(state);
                return;
            }

            // We're being called on another thread...
            AutoResetEvent ev = new AutoResetEvent(false);
            Post(d, state);
            Post((object are) => ((AutoResetEvent)are).Set(), ev);
            ev.WaitOne();
        }

        /// <summary>Runs a loop to process all queued work items.</summary>
        public void RunOnCurrentThread()
        {
            _threadId = Thread.CurrentThread.ManagedThreadId;
            Logger.Monitor.Info($"SingleThreadSynchronizationContext Running (Thread {_threadId})!");
            foreach (var workItem in _queue.GetConsumingEnumerable())
            {
                try
                {
                    workItem.Key(workItem.Value);
                }
                catch (Exception ex)
                {
                    Logger.Monitor.Warn($"SingleThreadSynchronizationContext ### Unhandled exception (Thread {_threadId}) - {ex}");
                }
            }

            Logger.Monitor.Info("SingleThreadSynchronizationContext Complete!");
        }

        /// <summary>Notifies the context that no more work will arrive.</summary>
        public void Complete() { _queue.CompleteAdding(); }
    }
}
