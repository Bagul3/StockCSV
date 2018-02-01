using System;
using System.Threading;

namespace StockCSV.Mechanism
{
    public abstract class Job
    {
        public void ExecuteJob()
        {
            if (!this.IsRepeatable()) return;
            while (true)
            {
                var now = DateTime.Now.TimeOfDay;

                if (now > GetStartTime() && now < GetEndTime())
                {
                    
                    Console.WriteLine($"The Clean Job thread started successfully.");
                    new LogWriter("The Clean Job thread started successfully");
                    this.DoCleanup();
                    this.DoJob();
                }
                Thread.Sleep(this.GetRepetitionIntervalTime());
            }
        }

        public virtual object GetParameters()
        {
            return null;
        }

        public abstract void DoJob();

        public abstract void DoCleanup();

        public abstract bool IsRepeatable();

        public abstract int GetRepetitionIntervalTime();

        public abstract TimeSpan GetStartTime();

        public abstract TimeSpan GetEndTime();
    }
}
