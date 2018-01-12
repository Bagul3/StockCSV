﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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

        public void CreateDBFFile(List<string> T2TREFs)
        {
            string connectionString = @"Provider = Microsoft.Jet.OLEDB.4.0; Data Source = C:\Users\van-d\Downloads\Cordners Data Dump\Cordners Data Dump\; Extended Properties = dBase IV";
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();
                OleDbCommand command = connection.CreateCommand();

                command.CommandText = "create table Descriptions(T2TREF int)";
                command.ExecuteNonQuery();
                connection.Close();
            }
            
            using (OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\van-d\Downloads\Cordners Data Dump\Cordners Data Dump\;Extended Properties=dBASE III;"))
            {
                connection.Open();
                OleDbCommand command = connection.CreateCommand();
                foreach (var t2tref in T2TREFs)
                {
                    var sql = "Insert INTO DESCRIPT (T2TREF) VALUES ({0});";                    
                    sql = string.Format(sql, String.Format("{0:00000}", t2tref));
                    command.CommandText = sql;
                    command.ExecuteNonQuery();
                }
            }                
        }

        public StringBuilder StockQuery(string reff)
        {
            string con = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\van-d\Downloads\Cordners Data Dump\Cordners Data Dump\;Extended Properties=dBASE III;";

            var csv = new StringBuilder();

            using (OleDbConnection connectionHandler = new OleDbConnection(con))
            {
                connectionHandler.Open();

                var headers = string.Format("{0},{1},{2}", "sku", "qty", "is_in_stock");
                csv.AppendLine(headers);

                //var stockQuery = @"SELECT T2_HEAD.SHORT, T2_HEAD.[DESC], T2_HEAD.[GROUP], T2_HEAD.STYPE, T2_HEAD.SIZERANGE, 
                //                    T2_HEAD.SUPPLIER, T2_HEAD.SUPPREF, T2_HEAD.VAT, T2_HEAD.BASESELL, T2_HEAD.SELL, T2_HEAD.SELLB, T2_HEAD.SELL1, Sum(T2_BRA.Q11) AS QTY1, Sum(T2_BRA.Q12) AS QTY2, 
                //                        Sum(T2_BRA.Q13) AS QTY3, Sum(T2_BRA.Q14) AS QTY4, Sum(T2_BRA.Q15) AS QTY5, Sum(T2_BRA.Q16) AS QTY6, Sum(T2_BRA.Q17) AS QTY7, Sum(T2_BRA.Q18) AS QTY8, 
                //                            Sum(T2_BRA.Q19) AS QTY9, Sum(T2_BRA.Q20) AS QTY10, Sum(T2_BRA.Q21) AS QTY11, Sum(T2_BRA.Q22) AS QTY12, Sum(T2_BRA.Q23) AS QTY13, T2_HEAD.REF,
                //                                Sum(T2_BRA.LY11) AS LY1, Sum(T2_BRA.LY12) AS LY2, Sum(T2_BRA.LY13) AS LY3, Sum(T2_BRA.LY14) AS LY4, Sum(T2_BRA.LY15) AS LY5, 
                //                                    Sum(T2_BRA.LY16) AS LY6, Sum(T2_BRA.LY17) AS LY7, Sum(T2_BRA.LY18) AS LY8, Sum(T2_BRA.LY19) AS LY9, Sum(T2_BRA.LY20) AS LY10, 
                //                                        Sum(T2_BRA.LY21) AS LY11, Sum(T2_BRA.LY22) AS LY12, Sum(T2_BRA.LY23) AS LY13 
                //                                            FROM DESCRIPT INNER JOIN (T2_BRA INNER JOIN T2_HEAD ON T2_BRA.REF = T2_HEAD.REF) ON DESCRIPT.T2TREF = T2_BRA.REF 
                //                                                Group By T2_HEAD.Short, T2_HEAD.[Desc], T2_HEAD.[Group], T2_HEAD.STYPE, T2_HEAD.SIZERANGE, T2_HEAD.SUPPLIER, T2_HEAD.SUPPREF, T2_HEAD.VAT, T2_HEAD.BASESELL, T2_HEAD.SELL, T2_HEAD.SELLB, T2_HEAD.SELL1, T2_HEAD.REF";

                //var stockQuery = @"SELECT T2_HEAD.SHORT, 
                //T2_HEAD.[DESC], T2_HEAD.[GROUP], T2_HEAD.STYPE, 
                //T2_HEAD.SIZERANGE, T2_HEAD.SUPPLIER, T2_HEAD.SUPPREF, T2_HEAD.VAT, T2_HEAD.BASESELL,
                // T2_HEAD.SELL, T2_HEAD.SELLB, T2_HEAD.SELL1, 
                //Sum(T2_BRA.Q11) AS SumOfQ11, 
                //Sum(T2_BRA.Q12) AS SumOfQ12, 
                //Sum(T2_BRA.Q13) AS SumOfQ13, 
                //Sum(T2_BRA.Q14) AS SumOfQ14, 
                //Sum(T2_BRA.Q15) AS SumOfQ15, 
                //Sum(T2_BRA.Q16) AS SumOfQ16, 
                //Sum(T2_BRA.Q17) AS SumOfQ17, 
                //Sum(T2_BRA.Q18) AS SumOfQ18, 
                //Sum(T2_BRA.Q19) AS SumOfQ19, 
                //Sum(T2_BRA.Q20) AS SumOfQ20,
                //Sum(T2_BRA.Q21) AS SumOfQ21, 
                //Sum(T2_BRA.Q22) AS SumOfQ22, 
                //Sum(T2_BRA.Q23) AS SumOfQ23, 
                //Sum(T2_BRA.LY11) AS SumOfLY11, 
                //Sum(T2_BRA.LY12) AS SumOfLY12, 
                //Sum(T2_BRA.LY13) AS SumOfLY13, 
                //Sum(T2_BRA.LY14) AS SumOfLY14, 
                //Sum(T2_BRA.LY15) AS SumOfLY15, 
                //Sum(T2_BRA.LY16) AS SumOfLY16, 
                //Sum(T2_BRA.LY17) AS SumOfLY17, 
                //Sum(T2_BRA.LY18) AS SumOfLY18, 
                //Sum(T2_BRA.LY19) AS SumOfLY19, 
                //Sum(T2_BRA.LY20) AS SumOfLY20, 
                //Sum(T2_BRA.LY21) AS SumOfLY21, 
                //Sum(T2_BRA.LY22) AS SumOfLY22, 
                //Sum(T2_BRA.LY23) AS SumOfLY23 
                //FROM DESCRIPT INNER JOIN (T2_BRA INNER JOIN T2_HEAD ON T2_BRA.REF = T2_HEAD.REF)
                //ON DESCRIPT.T2TREF = T2_BRA.REF 
                //Group By T2_HEAD.Short, T2_HEAD.[Desc], T2_HEAD.[Group], T2_HEAD.STYPE, T2_HEAD.SIZERANGE, T2_HEAD.SUPPLIER, T2_HEAD.SUPPREF, 
                //T2_HEAD.VAT, T2_HEAD.BASESELL, T2_HEAD.SELL, T2_HEAD.SELLB, T2_HEAD.SELL1";

                var stockQuery = @"SELECT T2_HEAD.SHORT, T2_HEAD.[DESC], T2_HEAD.[GROUP], T2_HEAD.STYPE, T2_HEAD.SIZERANGE, 
                                    T2_HEAD.SUPPLIER, T2_HEAD.SUPPREF, T2_HEAD.VAT, T2_HEAD.BASESELL, T2_HEAD.SELL, T2_HEAD.SELLB, T2_HEAD.SELL1, Sum(T2_BRA.Q11) AS QTY1, Sum(T2_BRA.Q12) AS QTY2, 
                                        Sum(T2_BRA.Q13) AS QTY3, Sum(T2_BRA.Q14) AS QTY4, Sum(T2_BRA.Q15) AS QTY5, Sum(T2_BRA.Q16) AS QTY6, Sum(T2_BRA.Q17) AS QTY7, Sum(T2_BRA.Q18) AS QTY8, 
                                            Sum(T2_BRA.Q19) AS QTY9, Sum(T2_BRA.Q20) AS QTY10, Sum(T2_BRA.Q21) AS QTY11, Sum(T2_BRA.Q22) AS QTY12, Sum(T2_BRA.Q23) AS QTY13, T2_HEAD.REF,
                                                Sum(T2_BRA.LY11) AS LY1, Sum(T2_BRA.LY12) AS LY2, Sum(T2_BRA.LY13) AS LY3, Sum(T2_BRA.LY14) AS LY4, Sum(T2_BRA.LY15) AS LY5, 
                                                    Sum(T2_BRA.LY16) AS LY6, Sum(T2_BRA.LY17) AS LY7, Sum(T2_BRA.LY18) AS LY8, Sum(T2_BRA.LY19) AS LY9, Sum(T2_BRA.LY20) AS LY10, 
                                                        Sum(T2_BRA.LY21) AS LY11, Sum(T2_BRA.LY22) AS LY12, Sum(T2_BRA.LY23) AS LY13 
                                                            FROM DESCRIPT INNER JOIN (T2_BRA INNER JOIN T2_HEAD ON T2_BRA.REF = T2_HEAD.REF) ON DESCRIPT.T2TREF = T2_BRA.REF 
                                                                Group By T2_HEAD.Short, T2_HEAD.[Desc], T2_HEAD.[Group], T2_HEAD.STYPE, T2_HEAD.SIZERANGE, 
                                                                    T2_HEAD.SUPPLIER, T2_HEAD.SUPPREF, T2_HEAD.VAT, T2_HEAD.BASESELL, T2_HEAD.SELL, T2_HEAD.SELLB, T2_HEAD.SELL1, T2_HEAD.REF";

                //var stockQuery = @"SELECT * FROM DESCRIPT ";

                //var stockQuery = @"SELECT ([T2_BRA]![REF] & [F7]) AS NEWSTYLE, T2_HEAD.SHORT, T2_HEAD.DESC, T2_HEAD.GROUP, T2_HEAD.STYPE, T2_HEAD.SIZERANGE, T2_HEAD.SUPPLIER, T2_HEAD.SUPPREF, T2_HEAD.VAT, T2_HEAD.BASESELL, T2_HEAD.SELL, T2_HEAD.SELLB, T2_HEAD.SELL1, Sum(T2_BRA.Q11) AS SumOfQ11, Sum(T2_BRA.Q12) AS SumOfQ12, Sum(T2_BRA.Q13) AS SumOfQ13, Sum(T2_BRA.Q14) AS SumOfQ14, Sum(T2_BRA.Q15) AS SumOfQ15, Sum(T2_BRA.Q16) AS SumOfQ16, Sum(T2_BRA.Q17) AS SumOfQ17, Sum(T2_BRA.Q18) AS SumOfQ18, Sum(T2_BRA.Q19) AS SumOfQ19, Sum(T2_BRA.Q20) AS SumOfQ20, Sum(T2_BRA.Q21) AS SumOfQ21, Sum(T2_BRA.Q22) AS SumOfQ22, Sum(T2_BRA.Q23) AS SumOfQ23, T2_HEAD.REF,
                //                    Sum(T2_BRA.LY11) AS SumOfLY11, Sum(T2_BRA.LY12) AS SumOfLY12, Sum(T2_BRA.LY13) AS SumOfLY13, Sum(T2_BRA.LY14) AS SumOfLY14, Sum(T2_BRA.LY15) AS SumOfLY15, Sum(T2_BRA.LY16) AS SumOfLY16, Sum(T2_BRA.LY17) AS SumOfLY17, Sum(T2_BRA.LY18) AS SumOfLY18, Sum(T2_BRA.LY19) AS SumOfLY19, Sum(T2_BRA.LY20) AS SumOfLY20, Sum(T2_BRA.LY21) AS SumOfLY21, Sum(T2_BRA.LY22) AS SumOfLY22, Sum(T2_BRA.LY23) AS SumOfLY23 
                //                        FROM DESCRIPT INNER JOIN (T2_BRA INNER JOIN T2_HEAD ON T2_BRA.REF = T2_HEAD.REF) ON DESCRIPT.T2TREF = T2_BRA.REF 
                //                            Group By([T2_BRA]![REF] & [F7]), T2_HEAD.Short, T2_HEAD.Desc, T2_HEAD.Group, T2_HEAD.STYPE, T2_HEAD.SIZERANGE, T2_HEAD.SUPPLIER, T2_HEAD.SUPPREF, T2_HEAD.VAT, T2_HEAD.BASESELL, T2_HEAD.SELL, T2_HEAD.SELLB, T2_HEAD.SELL1, T2_HEAD.REF
                //                                ORDER BY ([T2_BRA]![REF] & [F7]) DESC;";

                DataSet data = new DataSet();
                OleDbCommand myAccessCommand = new OleDbCommand(stockQuery, connectionHandler);
                //myAccessCommand.Parameters.AddWithValue("?", reff);

                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);
                myDataAdapter.Fill(data);

                var actualStock = "";
                var is_stock = 0;
                var InStockFlag = false;
                var GroupSKUS = "";
                var salesPrice = "";

              foreach (DataRow dr in data.Tables[0].Rows)
                {
                    for (var i = 1; i < 14; i++)
                    {
                        if (Convert.ToInt32(dr["SELL"]) - Convert.ToInt32(dr["SELLB"]) == 0)
                        {
                            salesPrice = "";
                        }
                        else
                        {
                            salesPrice = dr["SELLB"].ToString();
                        }

                        if (!String.IsNullOrEmpty(dr["QTY" + i].ToString()))
                        {
                            if (Convert.ToInt32(dr["QTY" + i]) > 0)
                            {
                                if (String.IsNullOrEmpty(dr["LY" + i].ToString()))
                                {
                                    actualStock = dr["QTY" + i].ToString();
                                }
                                else
                                {
                                    actualStock = (Convert.ToInt32(dr["QTY" + i]) - Convert.ToInt32(dr["LY" + i])).ToString();
                                }

                                is_stock = 1;
                                InStockFlag = true;
                            }
                            else
                            {
                                is_stock = 0;
                            }
                            string append = (100 + i).ToString();
                            GroupSKUS = dr["REF"].ToString();
                            var GroupSKUS2 = dr["REF"].ToString() + append.Substring(0, 3);
                            var newLine = string.Format("{0},{1},{2}", GroupSKUS2, actualStock, is_stock);
                            csv.AppendLine(newLine);
                        }

                        actualStock = "0";
                    }

                    if (InStockFlag)
                    {
                        is_stock = 1;
                    }
                    else
                    {
                        is_stock = 0;
                    }

                    var newLine2 = string.Format("{0},{1},{2}", GroupSKUS, actualStock, is_stock);
                    csv.AppendLine(newLine2);
                    InStockFlag = false;
                }

            }
            File.AppendAllText(@"C:\Users\van-d\Documents\stocknew.csv", csv.ToString());
            return null;
        }

        public List<string> QueryDescriptionXLX()
        {
            var query = "Select T2TREF FROM";
            OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\van-d\Downloads\Cordners Data Dump\descriptions.xls;Extended Properties='Excel 12.0;IMEX=1;'");
            StringBuilder stbQuery = new StringBuilder();
            stbQuery.Append("SELECT * FROM [Sheet1$A:A]");
            OleDbDataAdapter adp = new OleDbDataAdapter(stbQuery.ToString(), con);

            DataSet dsXLS = new DataSet();
            adp.Fill(dsXLS);
            DataView dvEmp = new DataView(dsXLS.Tables[0]);

            List<string> descriptionSkuNumbers = new List<string>();

            foreach (DataRow row in dvEmp.Table.Rows)
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
