using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StockCSV
{
    class Program
    {
        static void Main(string[] args)
        {
            Database database = new Database();
            var result = database.QueryDescriptionRefs();
            database = new Database(@"C:\Users\van-d\Downloads\Cordners Data Dump\Cordners Data Dump\");
            //database.CreateDBFFile(result);
            database.StockQuery(result);
        }
    }
}
