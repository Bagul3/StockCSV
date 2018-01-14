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

        //private readonly VulnService service = new VulnService();
        private readonly Database _database = new Database(@"C:\Users\van-d\Downloads\Cordners Data Dump\Cordners Data Dump\");
        private string _ref;

        public StockJob(string _ref)
        {
            this._ref = _ref;
        }

        public override void DoJob()
        {
            Console.WriteLine($"The Job \"{this.GetRef()}\" was executed.");
           // this._database.StockQuery(this.GetRef());
        }

        public override bool IsRepeatable()
        {
            return true;
        }

        public override string GetRef()
        {
            return "Apache";
        }

        public override int GetRepetitionIntervalTime()
        {
            return 1000;
        }
    }
}
