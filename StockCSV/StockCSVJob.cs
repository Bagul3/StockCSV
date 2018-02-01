using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using StockCSV.Mechanism;

namespace StockCSV
{
    public class StockCSVJob : Job
    {
        private LogWriter _logger = new LogWriter();

        public void CreateDbfFile(List<string> T2TREFs)
        {
            using (var connection = new OleDbConnection(System.Configuration.ConfigurationManager.AppSettings["AccessConnectionString"]))
            {
                connection.Open();
                var command = connection.CreateCommand();

                command.CommandText = "create table Descriptions(T2TREF int)";
                command.ExecuteNonQuery();
                connection.Close();
            }
            
            using (var connection = new OleDbConnection())
            {
                connection.Open();
                var command = connection.CreateCommand();
                foreach (var t2Tref in T2TREFs)
                {
                    var sql = "Insert INTO DESCRIPT (T2TREF) VALUES ({0});";                    
                    sql = string.Format(sql, String.Format("{0:00000}", t2Tref));
                    command.CommandText = sql;
                    command.ExecuteNonQuery();
                }
            }                
        }

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
                    const string stockQuery = @"SELECT ([T2_BRA].[REF] + [F7]) AS NewStyle, T2_HEAD.SHORT, T2_HEAD.[DESC], T2_HEAD.[GROUP], 
		T2_HEAD.STYPE, T2_HEAD.SIZERANGE,
					T2_HEAD.SUPPLIER, T2_HEAD.SUPPREF, T2_HEAD.VAT, 
						T2_HEAD.BASESELL, T2_HEAD.SELL, T2_HEAD.SELLB, T2_HEAD.SELL1, Dept.MasterDept, Dept.MasterDept,

			Sum(T2_BRA.Q11) AS QTY1, Sum(T2_BRA.Q12) AS QTY2, 
                Sum(T2_BRA.Q13) AS QTY3, Sum(T2_BRA.Q14) AS QTY4, Sum(T2_BRA.Q15) AS QTY5, Sum(T2_BRA.Q16) AS QTY6, Sum(T2_BRA.Q17) AS QTY7, Sum(T2_BRA.Q18) AS QTY8, 
                    Sum(T2_BRA.Q19) AS QTY9, Sum(T2_BRA.Q20) AS QTY10, Sum(T2_BRA.Q21) AS QTY11, Sum(T2_BRA.Q22) AS QTY12, Sum(T2_BRA.Q23) AS QTY13, T2_HEAD.REF,
                        Sum(T2_BRA.LY11) AS LY1, Sum(T2_BRA.LY12) AS LY2, Sum(T2_BRA.LY13) AS LY3, Sum(T2_BRA.LY14) AS LY4, Sum(T2_BRA.LY15) AS LY5, 
                            Sum(T2_BRA.LY16) AS LY6, Sum(T2_BRA.LY17) AS LY7, Sum(T2_BRA.LY18) AS LY8, Sum(T2_BRA.LY19) AS LY9, Sum(T2_BRA.LY20) AS LY10, 
                                Sum(T2_BRA.LY21) AS LY11, Sum(T2_BRA.LY22) AS LY12, Sum(T2_BRA.LY23) AS LY13 
									FROM ((((((T2_BRA INNER JOIN T2_HEAD ON T2_BRA.REF = T2_HEAD.REF) INNER JOIN (SELECT Right(T2_LOOK.[KEY],3) AS NewCol, T2_LOOK.F1 AS MasterColour, Left(T2_LOOK.[KEY],3) AS Col, T2_LOOK.F7
								FROM T2_LOOK
								WHERE (Left(T2_LOOK.[KEY],3))='COL') as Colour ON T2_BRA.COLOUR = Colour.NewCol) INNER JOIN 

								(SELECT Mid(T2_LOOK.[KEY],4,6) AS SuppCode, T2_LOOK.F1 AS MasterSupplier
									FROM T2_LOOK
										WHERE (((Left(T2_LOOK.[KEY],3))='SUP'))
											) as  Suppliers ON T2_HEAD.SUPPLIER = Suppliers.SuppCode) INNER JOIN

											(SELECT Right([T2_LOOK].[KEY],3) AS DeptCode, T2_LOOK.F1 AS MasterDept
												FROM T2_LOOK
													WHERE (Left([T2_LOOK].[KEY],3))='TYP') As Dept ON T2_HEAD.STYPE = Dept.DeptCode) INNER JOIN
								(SELECT Mid(T2_LOOK.[KEY], 4, 6) AS StkType,
									T2_LOOK.F1 AS MasterStocktype
										FROM T2_LOOK
											WHERE Left(T2_LOOK.[KEY], 3) = 'CAT'
											) as Stocktype
									ON T2_HEAD.[GROUP] = Stocktype.StkType) 	LEFT JOIN

									(SELECT Right(T2_LOOK.[KEY],3) AS SubDeptCode, T2_LOOK.F1 AS MasterSubDept
										FROM T2_LOOK
											WHERE (Left(T2_LOOK.[KEY],3))='US2') AS SubDept ON T2_HEAD.USER2 = SubDept.SubDeptCode)
                                    WHERE [T2_BRA].[REF] = ?
									GROUP BY ([T2_BRA].[REF] + [F7]),
									T2_HEAD.SHORT, T2_HEAD.[DESC], T2_HEAD.[GROUP],  Dept.MasterDept, T2_HEAD.STYPE, T2_HEAD.SIZERANGE, T2_HEAD.SUPPLIER, T2_HEAD.SUPPREF, 
									T2_HEAD.VAT, T2_HEAD.BASESELL, T2_HEAD.SELL, T2_HEAD.SELLB, T2_HEAD.SELL1, T2_HEAD.REF
                                    ORDER BY ([T2_BRA].[REF] + [F7]) DESC";

                    var data = new DataSet();
                    var myAccessCommand = new OleDbCommand(stockQuery, connectionHandler);
                    myAccessCommand.Parameters.AddWithValue("?", reff);

                    var myDataAdapter = new OleDbDataAdapter(myAccessCommand);
                    myDataAdapter.Fill(data);

                    var actualStock = "0";
                    var inStockFlag = false;
                    var groupSkus = "";

                    foreach (DataRow dr in data.Tables[0].Rows)
                    {
                        _logger.LogWrite("Working....");
                        var isStock = 0;
                        for (var i = 1; i < 14; i++)
                        {
                            if (!string.IsNullOrEmpty(dr["QTY" + i].ToString()))
                            {
                                if (dr["QTY" + i].ToString() != "")
                                {
                                    if (Convert.ToInt32(dr["QTY" + i]) > 0)
                                    {
                                        if (String.IsNullOrEmpty(dr["LY" + i].ToString()))
                                        {
                                            actualStock = dr["QTY" + i].ToString();
                                        }
                                        else
                                        {
                                            actualStock =
                                                (Convert.ToInt32(dr["QTY" + i]) - Convert.ToInt32(dr["LY" + i]))
                                                .ToString();
                                        }

                                        isStock = 1;
                                        inStockFlag = true;
                                    }
                                    else
                                    {
                                        isStock = 0;
                                    }
                                    var append = (1000 + i).ToString();
                                    groupSkus = dr["NewStyle"].ToString();
                                    var groupSkus2 = dr["NewStyle"] + append.Substring(1, 3);
                                    var newLine = $"{groupSkus2},{actualStock},{isStock}";
                                    csv.AppendLine(newLine);
                                }
                                actualStock = "0";
                            }
                        }

                        isStock = inStockFlag ? 1 : 0;
                        if (!string.IsNullOrEmpty(dr["NewStyle"].ToString()))
                        {
                            var newLine2 = $"{groupSkus},{actualStock},{isStock}";
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
            File.AppendAllText(System.Configuration.ConfigurationManager.AppSettings["OutputPath"], csv.ToString());
            Console.WriteLine("Job Finished");
            _logger.LogWrite("Finished");
        }

        public override void DoCleanup()
        {
            Console.WriteLine("Clean up: removing exisiting stock.csv");
            if (File.Exists(System.Configuration.ConfigurationManager.AppSettings["OutputPath"]))
            {
                try
                {
                    File.Delete(System.Configuration.ConfigurationManager.AppSettings["OutputPath"]);
                }
                catch (Exception e)
                {
                    _logger.LogWrite("Error occured deleting previous stock.csv file, please ensure the file is not been used by another process :" + e);
                }
                
            }
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
                    var adp = new OleDbDataAdapter("SELECT * FROM [Sheet1$B:B]", connectionHandler);

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

        public override bool IsRepeatable()
        {
            return true;
        }

        public override int GetRepetitionIntervalTime()
        {
            return 1800000;
        }

        public override TimeSpan GetStartTime()
        {
            return TimeSpan.Parse(System.Configuration.ConfigurationManager.AppSettings["StartTime"]);
        }

        public override TimeSpan GetEndTime()
        {
            return TimeSpan.Parse(System.Configuration.ConfigurationManager.AppSettings["EndTime"]);
        }
    }
}
