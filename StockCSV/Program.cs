using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using StockCSV.Mechanism;

namespace StockCSV
{
    class Program
    {
        //static void Main(string[] args)
        //{
        //    //var database = new Database();
        //    //database.DoesFileExist();
        //    //var result = database.QueryDescriptionRefs();
        //    ////database = new Database(@"C:\Users\van-d\Downloads\Cordners Data Dump\Cordners Data Dump\");
        //    ////database.CreateDBFFile(result);
        //    //database.StockQuery(result);
        //}

        public static void Main(string[] args)
        {
            Start();
        }

        private static void Start()
        {
            var jobManager = new JobManager();
            jobManager.ExecuteAllJobs();
        }
    }
}
