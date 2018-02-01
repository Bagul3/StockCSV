using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using StockCSV.Mechanism;
using StockCSV.Models;

namespace StockCSV
{
    public class ImportCSVJob : Job
    {
        private LogWriter _logger = new LogWriter();
        private ExcelMapper _mapper = new ExcelMapper();
        
        public override void DoJob()
        {
            var t2TreFs = ReadImageDetails(@"C:\Users\Conor\Desktop\Cordners Data Dump\images\");
            var csv = new StringBuilder();
            Console.WriteLine("Generating stock.csv: This will take a few minutes, please wait....");
            _logger.LogWrite("Generating stock.csv: This will take a few minutes, please wait....");
            var descriptions = _mapper.MapToDescriptions();

            using (var connectionHandler = new OleDbConnection(System.Configuration.ConfigurationManager.AppSettings["AccessConnectionString"]))
            {
                connectionHandler.Open();
                var short_description = "";

                var headers =
                    $"{"store"},{"websites"},{"attribut_set"},{"type"},{"sku"},{"has_options"}, {"name"}, {"page_layout"}, {"options_container"}, {"price"}, {"weight"}, {"status"}, {"visibility"}, {"short_description"}, {"qty"}, {"product_name"}, {"color"}" +
                    $"{"size"},{"tax_class_id"}, {"configurable_attributes"}, {"simples_skus"}, {"manufacturer"}, {"is_in_stock"}, {"categories"}, {"season"}, {"stock_type"}, {"image"}, {"small_image"}, {"thumbnail"}, {"gallery"}, {"condition"}, {"ean"}, {"description"}";

                csv.AppendLine(headers);
                foreach (var refff in t2TreFs)
                {
                    var reff = refff.Substring(0, 6);
                    const string stockQuery = @"SELECT ([T2_BRA].[REF] + [F7]) AS NEWSTYLE, Suppliers.MasterSupplier, Dept.MasterDept, Colour.MasterColour, Colour.F7, T2_HEAD.SHORT, 
	T2_HEAD.[DESC], T2_HEAD.[GROUP], T2_HEAD.STYPE, T2_HEAD.SIZERANGE, T2_HEAD.SUPPLIER, T2_HEAD.SUPPREF, T2_HEAD.VAT, T2_HEAD.BASESELL, T2_HEAD.SELL, 
		T2_HEAD.SELLB, T2_HEAD.SELL1, Sum(T2_BRA.Q11) AS QTY1, Sum(T2_BRA.Q12) AS QTY2, Sum(T2_BRA.Q13) AS QTY3, 
			Sum(T2_BRA.Q14) AS QTY4, Sum(T2_BRA.Q15) AS QTY5, Sum(T2_BRA.Q16) AS QTY6, Sum(T2_BRA.Q17) AS QTY7, Sum(T2_BRA.Q18) AS QTY8, 
				Sum(T2_BRA.Q19) AS QTY9, Sum(T2_BRA.Q20) AS QTY10, Sum(T2_BRA.Q21) AS QTY11, Sum(T2_BRA.Q22) AS QTY12, Sum(T2_BRA.Q23) AS QTY13, 
					T2_HEAD.REF,Stocktype.MasterStocktype,SubDept.MasterSubDept,
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
									GROUP BY ([T2_BRA].[REF] + [F7]), Suppliers.MasterSupplier, Dept.MasterDept, Colour.MasterColour, Colour.F7,
									 T2_HEAD.SHORT, T2_HEAD.[DESC], T2_HEAD.[GROUP], T2_HEAD.STYPE, T2_HEAD.SIZERANGE, T2_HEAD.SUPPLIER, T2_HEAD.SUPPREF,
									  T2_HEAD.VAT, T2_HEAD.BASESELL,
									 T2_HEAD.SELL, T2_HEAD.SELLB, T2_HEAD.SELL1,T2_HEAD.REF,stocktype.MasterStocktype,SubDept.MasterSubDept
									ORDER BY ([T2_BRA].[REF] + [F7]) DESC;";

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
                                    short_description = BuildShortDescription(descriptions.First(x => x.T2TRef == reff));
                                    var descripto = descriptions.Where(x => x.T2TRef == reff)
                                        .Select(y => y.Descriptio).First();
                                    var name = dr["MasterSupplier"] + "  " + descriptions.Where(x => x.T2TRef == reff).Select(y => y.Descriptio).First() + " in " + dr["MasterColour"];
                                    var newLine = $"{"admin"}, {"admin base"}, {"Default"}, {"simple"}, {groupSkus2}, {"1"}, {name},{"No layout updates."},{"Product Info Column"}, {dr["BASESELL"]}, {"0.01"}, {"Enabled"}, {"Not Visible Individually"}, {short_description}, {actualStock}, {descripto}, {dr["MasterColour"]}" +
                                                  $"{"ShoeSizeTODO"}, {"tax class todo"}, {""},{""}, {dr["MasterSupplier"]}, {isStock}, {"TODO categories"}, {""}, {dr["GROUP"]}, {"+/" + reff + ".jpg"}, {"/" + reff + ".jpg"},{"/" + reff + ".jpg"},{"TODO: gallery"}, {"new"}, {""}, {descriptions.Where(x => x.T2TRef == reff).Select(y => y.Description).First()}";
                                    //var newLine = $"{groupSkus2},{actualStock},{isStock}";"
                                    csv.AppendLine(newLine);
                                }
                                actualStock = "0";
                            }
                        }

                        isStock = inStockFlag ? 1 : 0;
                        if (!string.IsNullOrEmpty(dr["NewStyle"].ToString()))
                        {
                            var name = dr["MasterSupplier"] + " in " + dr["MasterColour"];
                            short_description = BuildShortDescription(descriptions.First(x => x.T2TRef == reff));
                            var descripto = descriptions.Where(x => x.T2TRef == reff)
                                .Select(y => y.Descriptio).First();
                            var newLine2 = $"{"admin"}, {"admin base"}, {"Default"}, {"configurable"}, {groupSkus}, {"1"}, {name}, {"No layout updates."},{"Product Info Column"},{dr["BASESELL"]}, {"0.01"}, {"Enabled"}, {"Catalog Search"}, {short_description}, {0}, {descripto}, {dr["MasterColour"]}";
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
            File.AppendAllText(@"C: \Users\Conor\Desktop\import_products.csv", csv.ToString());
            Console.WriteLine("Job Finished");
            _logger.LogWrite("Finished");
        }

        public override void DoCleanup()
        {

        }
        
        public override bool IsRepeatable()
        {
            return true;
        }

        public override int GetRepetitionIntervalTime()
        {
            return 1000;
        }

        public override TimeSpan GetStartTime()
        {
            return TimeSpan.Parse(System.Configuration.ConfigurationManager.AppSettings["StartTime"]);
        }

        public override TimeSpan GetEndTime()
        {
            return TimeSpan.Parse(System.Configuration.ConfigurationManager.AppSettings["EndTime"]);
        }

        private string BuildShortDescription(Descriptions description)
        {
            if (string.IsNullOrEmpty(description.Bullet1) && string.IsNullOrEmpty(description.Bullet2) &&
                string.IsNullOrEmpty(description.Bullet3) && string.IsNullOrEmpty(description.Bullet4) &&
                string.IsNullOrEmpty(description.Bullet5) && string.IsNullOrEmpty(description.Bullet6) &&
                string.IsNullOrEmpty(description.Bullet7))
            {
                return "<ul></ul>";
            }
            else
            {
                var bullet1 = String.IsNullOrEmpty(description.Bullet1) ? "" : "<li>" + description.Bullet1 + "</li>";
                var bullet2 = String.IsNullOrEmpty(description.Bullet2) ? "" : "<li>" + description.Bullet2 + "</li>";
                var bullet3 = String.IsNullOrEmpty(description.Bullet3) ? "" : "<li>" + description.Bullet3 + "</li>";
                var bullet4 = String.IsNullOrEmpty(description.Bullet4) ? "" : "<li>" + description.Bullet4 + "</li>";
                var bullet5 = String.IsNullOrEmpty(description.Bullet5) ? "" : "<li>" + description.Bullet5 + "</li>";
                var bullet6 = String.IsNullOrEmpty(description.Bullet6) ? "" : "<li>" + description.Bullet6 + "</li>";
                var bullet7 = String.IsNullOrEmpty(description.Bullet7) ? "" : "<li>" + description.Bullet7 + "</li>";
                return "<ul>" + bullet1 + bullet2 + bullet3 + bullet4 + bullet5 + bullet6 + bullet7 + "</ul>";
            }

        }

        private IEnumerable<string> ReadImageDetails(string path)
        {
            try
            {
                var imageDetails = Directory.GetFiles(path, "*.jpg*", SearchOption.AllDirectories)
                    .ToList();
                return RemoveUnUsedValues(ParseImageNames(imageDetails));
            }
            catch (Exception e)
            {
                _logger.LogWrite("Error occured reading files: " + e);
                return new List<string>();
            }
        }

        private static IEnumerable<string> ParseImageNames(IEnumerable<string> imageDetails)
        {
            return imageDetails.Select(Path.GetFileNameWithoutExtension).ToList();
        }

        private static IEnumerable<string> RemoveUnUsedValues(IEnumerable<string> imageDetails)
        {
            return imageDetails.Where(image => !image.Contains("_")).ToList();
        }
    }
}
