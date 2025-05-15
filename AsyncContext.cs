using System;
using System.Threading;
using System.Threading.Tasks;

namespace SOFSInventoryDownloader
{
    public static class AsyncContext
    {
        public static void Run(Func<Task> func)
        {
            var prevCtx = SynchronizationContext.Current;
            try
            {
                var syncCtx = new SingleThreadSynchronizationContext();
                SynchronizationContext.SetSynchronizationContext(syncCtx);
                var task = func();
                task.ContinueWith(_ => syncCtx.Complete(), TaskScheduler.Default);
                syncCtx.RunOnCurrentThread();
                task.GetAwaiter().GetResult();
            }
            finally
            {
                SynchronizationContext.SetSynchronizationContext(prevCtx);
            }
        }
    }

    internal class SingleThreadSynchronizationContext : SynchronizationContext
    {
        private readonly BlockingQueue<Task> _tasks = new();

        public override void Post(SendOrPostCallback d, object state)
        {
            _tasks.Enqueue(() => d(state));
        }

        public void RunOnCurrentThread()
        {
            while (_tasks.TryDequeue(out var task))
                task.Invoke();
        }

        public void Complete() => _tasks.CompleteAdding();
    }
}
