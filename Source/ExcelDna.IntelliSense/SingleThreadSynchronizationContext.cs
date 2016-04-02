using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;

namespace ExcelDna.IntelliSense
{
    // This code is from Stephen Toub's blog post on AsyncPump: http://blogs.msdn.com/b/pfxteam/archive/2012/01/20/10259049.aspx?PageIndex=2#comments

    /// <summary>Provides a SynchronizationContext that's single-threaded.</summary>
    sealed class SingleThreadSynchronizationContext : SynchronizationContext
    {
        /// <summary>The queue of work items.</summary>
        private readonly BlockingCollection<KeyValuePair<SendOrPostCallback, object>> m_queue =
            new BlockingCollection<KeyValuePair<SendOrPostCallback, object>>();

        // /// <summary>The processing thread.</summary>
        // private readonly Thread m_thread = Thread.CurrentThread;

        /// <summary>Dispatches an asynchronous message to the synchronization context.</summary>
        /// <param name="d">The System.Threading.SendOrPostCallback delegate to call.</param>
        /// <param name="state">The object passed to the delegate.</param>
        public override void Post(SendOrPostCallback d, object state)
        {
            if (d == null) throw new ArgumentNullException("d");
            m_queue.Add(new KeyValuePair<SendOrPostCallback, object>(d, state));
        }

        /// <summary>Not supported.</summary>
        public override void Send(SendOrPostCallback d, object state)
        {
            throw new NotSupportedException("Synchronously sending is not supported.");
        }

        /// <summary>Runs a loop to process all queued work items.</summary>
        public void RunOnCurrentThread()
        {
            foreach (var workItem in m_queue.GetConsumingEnumerable())
            {
                try
                {
                    workItem.Key(workItem.Value);
                }
                catch (Exception ex)
                {
                    Debug.Print($"### Unhandled exception on Automation thread - {ex.ToString()}");
                }
            }

            Debug.Print("SingleThreadSynchronizationContext Complete!");
        }

        /// <summary>Notifies the context that no more work will arrive.</summary>
        public void Complete() { m_queue.CompleteAdding(); }
    }
}
