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
                    this.DoJob();
                    Thread.Sleep(this.GetRepetitionIntervalTime());
                }
            }
            this.DoJob();
        }

        public virtual object GetParameters()
        {
            return null;
        }

        public abstract string GetRef();

        public abstract void DoJob();

        public abstract bool IsRepeatable();

        public abstract int GetRepetitionIntervalTime();
    }
}
