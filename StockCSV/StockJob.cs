using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using StockCSV.Mechanism;
using StockCSV.Queries;

namespace StockCSV
{
    public class StockJob : Job
    {
        private readonly LogWriter _logger = new LogWriter();

        public override void DoJob()
        {
            var t2TreFs = QueryDescriptionRefs();
            var csv = new StringBuilder();
            Console.WriteLine("Generating stock.csv: This will take a few minutes, please wait....");
            _logger.LogWrite("Generating stock.csv: This will take a few minutes, please wait....");

            using (var connectionHandler = new OleDbConnection(System.Configuration.ConfigurationManager.AppSettings["AccessConnectionString"]))
            {
                connectionHandler.Open();

                var headers = $"{"sku"},{"qty"},{"is_in_stock"}";
                csv.AppendLine(headers);

                foreach (var reff in t2TreFs)
                {
                    var data = ExecuteOleCommand(connectionHandler, reff);
                    BuildCsv(data, ref csv);
                }
            }
            File.AppendAllText(System.Configuration.ConfigurationManager.AppSettings["OutputPath"], csv.ToString());
            Console.WriteLine("Job Finished");
            _logger.LogWrite("Finished");
        }
        
        private static string BuildChildSku(int i, DataRow dr)
        {
            var append = (1000 + i).ToString();

            var childSku = dr["NewStyle"] + append.Substring(1, 3);
            return childSku;
        }

        public override void DoCleanup()
        {
            Console.WriteLine("Clean up: removing exisiting stock.csv");
            if (File.Exists(System.Configuration.ConfigurationManager.AppSettings["OutputPath"]))
            {
                File.Delete(System.Configuration.ConfigurationManager.AppSettings["OutputPath"]);
            }
        }

        public override bool IsRepeatable()
        {
            return true;
        }

        public override int GetRepetitionIntervalTime()
        {
            return 5000;
        }

        public override TimeSpan GetStartTime()
        {
            return TimeSpan.Parse("23:17");
        }

        public override TimeSpan GetEndTime()
        {
            return TimeSpan.Parse("23:20");
        }

        private IEnumerable<string> QueryDescriptionRefs()
        {
            var dvEmp = new DataView();
            _logger.LogWrite("Getting refs from description file");
            try
            {
                using (var connectionHandler = new OleDbConnection(System.Configuration.ConfigurationManager.AppSettings["ExcelConnectionString"]))
                {
                    connectionHandler.Open();
                    var adp = new OleDbDataAdapter("SELECT * FROM [Sheet1$A:A]", connectionHandler);

                    var dsXls = new DataSet();
                    adp.Fill(dsXls);
                    dvEmp = new DataView(dsXls.Tables[0]);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                _logger.LogWrite("Error occured getting refs from description file: " + e);
            }

            return (from DataRow row in dvEmp.Table.Rows select row.ItemArray[0].ToString()).ToList();
        }

        private static DataSet ExecuteOleCommand(OleDbConnection connectionHandler, string reff)
        {
            var data = new DataSet();
            var myAccessCommand = new OleDbCommand(Query.StockQuery, connectionHandler);
            myAccessCommand.Parameters.AddWithValue("?", reff);

            var myDataAdapter = new OleDbDataAdapter(myAccessCommand);
            myDataAdapter.Fill(data);
            return data;
        }

        private void BuildCsv(DataSet data, ref StringBuilder csv)
        {
            var inStockFlag = false;
            var actualStock = "0";
            var parentSku = "";

            foreach (DataRow dr in data.Tables[0].Rows)
            {
                _logger.LogWrite("Working....");
                var isStock = 0;
                for (var i = 1; i < 14; i++)
                {
                    if (!string.IsNullOrEmpty(dr["QTY" + i].ToString()) && dr["QTY" + i].ToString() != "" &&
                        Convert.ToInt32(dr["QTY" + i]) > 0)
                    {
                        actualStock = String.IsNullOrEmpty(dr["LY" + i].ToString())
                            ? dr["QTY" + i].ToString()
                            : (Convert.ToInt32(dr["QTY" + i]) - Convert.ToInt32(dr["LY" + i])).ToString();
                        isStock = 1;
                        inStockFlag = true;
                    }
                    else
                    {
                        isStock = 0;
                    }
                    parentSku = dr["NewStyle"].ToString();
                    var childSku = BuildChildSku(i, dr);
                    var newLine = $"{childSku},{actualStock},{isStock}";
                    csv.AppendLine(newLine);
                    actualStock = "0";
                }

                isStock = inStockFlag ? 1 : 0;
                if (!string.IsNullOrEmpty(dr["NewStyle"].ToString()))
                {
                    var newLine2 = $"{parentSku},{actualStock},{isStock}";
                    csv.AppendLine(newLine2);
                }
                inStockFlag = false;
                if (data.Tables[0].Rows.Count > 1)
                {
                    break;
                }
            }
        }
    }
}
