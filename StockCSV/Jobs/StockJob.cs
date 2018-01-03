using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using StockCSV.Mechanism;

namespace StockCSV.Jobs
{
    public class StockJob : Job
    {

        private readonly VulnService service = new VulnService();

        public override string GetName()
        {
            return this.GetType().Name;
        }

        public override void DoJob()
        {
            Console.WriteLine($"The Job \"{this.GetEndpoint()}\" was executed.");
            this.service.InsertVulnerabilities(this.GetEndpoint()).Wait();
        }

        public override bool IsRepeatable()
        {
            return true;
        }

        public override string GetEndpoint()
        {
            return "Apache";
        }

        public override int GetRepetitionIntervalTime()
        {
            return 1000;
        }
    }
}
