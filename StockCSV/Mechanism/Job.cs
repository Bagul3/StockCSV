using System;
using System.Threading;

namespace StockCSV.Mechanism
{
    public abstract class Job
    {
        public void ExecuteJob()
        {
            if (this.IsRepeatable())
            {
                while (true)
                {
                    var start = TimeSpan.Parse("06:30");  // 10 PM
                    var end = TimeSpan.Parse("06:45");    // 2 AM
                    var now = DateTime.Now.TimeOfDay;

                    if (now > start && now < end)
                    {
                    
                            Console.WriteLine($"The Clean Job thread started successfully.");
                            new LogWriter("The Clean Job thread started successfully");
                            this.DoCleanup();
                            this.DoJob();
                            Thread.Sleep(this.GetRepetitionIntervalTime());
                    }
                }
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
    }
}
