using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace StockCSV
{
    public class Database
    {
        string connectionPath = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = ";
        OleDbConnection accessConnection = null;

        public Database(string databaseConnectionString)
        {
            connectionPath += databaseConnectionString;
        }

        public Database()
        {

        }

        private void OpenConnection(string accessConnectionPath)
        {
            try
            {
                accessConnection = new OleDbConnection(accessConnectionPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: Failed to create a database connection. \n{0}", ex.Message);
            }
        }

        public string StockQuery(List<string> desc)
        {
            string con = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\van-d\Downloads\Cordners Data Dump\Cordners Data Dump;Extended Properties=dBASE III;";
            //T2_HEAD INNER JOIN (T2_BRA INNER JOIN Colour ON T2_BRA.COLOUR = Colour.NewCol) ON T2_BRA.REF = T2_HEAD.REF
            using (OleDbConnection connectionHandler = new OleDbConnection(con))
            {
                var stockQuery = @"SELECT ([T2_BRA]![REF] & [F7]) AS NEWSTYLE, T2_HEAD.SHORT, T2_HEAD.DESC, T2_HEAD.GROUP, T2_HEAD.STYPE, T2_HEAD.SIZERANGE, T2_HEAD.SUPPLIER, T2_HEAD.SUPPREF, T2_HEAD.VAT, T2_HEAD.BASESELL, T2_HEAD.SELL, T2_HEAD.SELLB, T2_HEAD.SELL1, Sum(T2_BRA.Q11) AS QTY1, Sum(T2_BRA.Q12) AS QTY2, Sum(T2_BRA.Q13) AS QTY3, Sum(T2_BRA.Q14) AS QTY4, Sum(T2_BRA.Q15) AS QTY5, Sum(T2_BRA.Q16) AS QTY6, Sum(T2_BRA.Q17) AS QTY7, Sum(T2_BRA.Q18) AS QTY8, Sum(T2_BRA.Q19) AS QTY9, Sum(T2_BRA.Q20) AS QTY10, Sum(T2_BRA.Q21) AS QTY11, Sum(T2_BRA.Q22) AS QTY12, Sum(T2_BRA.Q23) AS QTY13,
                    Sum(T2_BRA.LY11) AS LY1, Sum(T2_BRA.LY12) AS LY2, Sum(T2_BRA.LY13) AS LY3, Sum(T2_BRA.LY14) AS LY4, Sum(T2_BRA.LY15) AS LY5, Sum(T2_BRA.LY16) AS LY6, Sum(T2_BRA.LY17) AS LY7, Sum(T2_BRA.LY18) AS LY8, Sum(T2_BRA.LY19) AS LY9, Sum(T2_BRA.LY20) AS LY10, Sum(T2_BRA.LY21) AS LY11, Sum(T2_BRA.LY22) AS LY12, Sum(T2_BRA.LY23) AS LY13 
                        FROM T2_BRA INNER JOIN T2_HEAD ON T2_HEAD.REF = T2_BRA.REF
                            Group By([T2_BRA]![REF] & [F7]), T2_HEAD.Short, T2_HEAD.Desc, T2_HEAD.Group, T2_HEAD.STYPE, T2_HEAD.SIZERANGE, T2_HEAD.SUPPLIER, T2_HEAD.SUPPREF, T2_HEAD.VAT, T2_HEAD.BASESELL, T2_HEAD.SELL, T2_HEAD.SELLB, T2_HEAD.SELL1
                                ORDER BY ([T2_BRA]![REF] & [F7]) DESC";

                DataSet data = new DataSet();
                connectionHandler.Open();
                OleDbCommand myAccessCommand = new OleDbCommand(stockQuery, connectionHandler);
                myAccessCommand.Parameters.Add(new OleDbParameter("@settings", desc[0]));

                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);
                myDataAdapter.Fill(data);
                
                foreach (DataRow dr in data.Tables[0].Rows)
                {
                    //if(Convert.ToInt32(dr["SELL"]) - Convert.ToInt32(dr["SELLB"]) == 0)
                    //{
                    //    Console.WriteLine("Sale price: none");
                    //}
                    //else
                    //{
                    //    Console.WriteLine("Sale price: " + dr["SELLB"]);
                    //}
                    Console.WriteLine(dr["NEWSTYLE"]);
                }
            }
            return null;
        }

        public List<string> QueryDescriptionXLX()
        {
            var query = "Select T2TREF FROM";
            OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\van-d\Downloads\Cordners Data Dump\Cordners Data Dump\descriptions.xls;Extended Properties='Excel 12.0;IMEX=1;'");
            StringBuilder stbQuery = new StringBuilder();
            stbQuery.Append("SELECT * FROM [Sheet1$A:A]");
            OleDbDataAdapter adp = new OleDbDataAdapter(stbQuery.ToString(), con);

            DataSet dsXLS = new DataSet();
            adp.Fill(dsXLS);
            DataView dvEmp = new DataView(dsXLS.Tables[0]);

            List<string> descriptionSkuNumbers = new List<string>();

            foreach(DataRow row in dvEmp.Table.Rows)
            {
                descriptionSkuNumbers.Add(row.ItemArray[0].ToString());
            }

            return descriptionSkuNumbers;

        }

        public string QueryAndBuild(string imageDetails)
        {
            string querySelect = "SELECT * FROM ShopStaff";

            DataSet data = new DataSet();
            OpenConnection(connectionPath);
            try
            {
                OleDbCommand myAccessCommand = new OleDbCommand(querySelect, accessConnection);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);
                myDataAdapter.Fill(data);
                foreach (DataTable table in data.Tables)
                    WriteToCsvFile(table, "D:\\Users\\Conor\\ContractWork\\output-" + imageDetails + ".csv");

                myDataAdapter.Dispose();
            }
            catch (Exception ex)
            {
                return "Error: Failed to retrieve the required data from the DataBase.\n{0}";
            }
            finally
            {
                accessConnection.Close();
            }
            return imageDetails + ".csv CSV file created successfully";
        }

        private void WriteToCsvFile(DataTable dataTable, string filePath)
        {
            StringBuilder fileContent = new StringBuilder();
            foreach (var col in dataTable.Columns)
                fileContent.Append(col.ToString() + ",");

            fileContent.Replace(",", Environment.NewLine, fileContent.Length - 1, 1);
            foreach (DataRow dr in dataTable.Rows)
            {
                foreach (var column in dr.ItemArray)
                    fileContent.Append("\"" + column.ToString() + "\",");
                fileContent.Replace(",", Environment.NewLine, fileContent.Length - 1, 1);
            }
            System.IO.File.WriteAllText(filePath, fileContent.ToString());
        }
    }
}
