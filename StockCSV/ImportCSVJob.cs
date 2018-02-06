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
        private readonly LogWriter _logger = new LogWriter();
        private readonly ExcelMapper _mapper = new ExcelMapper();

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
                    $"{"store"},{"websites"},{"attribut_set"},{"type"},{"sku"},{"has_options"},{"name"},{"page_layout"},{"options_container"},{"price"},{"weight"},{"status"},{"visibility"},{"short_description"},{"qty"},{"product_name"},{"color"}," +
                    $"{"size"},{"tax_class_id"},{"configurable_attributes"},{"simples_skus"},{"manufacturer"},{"is_in_stock"},{"categories"},{"season"},{"stock_type"},{"image"},{"small_image"},{"thumbnail"},{"gallery"}," +
                    $"{"condition"},{"ean"},{"description"},{"model"}";

                csv.AppendLine(headers);
                foreach (var refff in t2TreFs.Where(x => !x.Contains("_")))
                {
                    var reff = refff.Substring(0, 6);
                    var reffColour = refff.Substring(0, 9);
                    var data = new DataSet();
                    var myAccessCommand = new OleDbCommand(SqlQuery.ImportProductsQuery, connectionHandler);
                    myAccessCommand.Parameters.AddWithValue("?", reff);

                    var myDataAdapter = new OleDbDataAdapter(myAccessCommand);
                    myDataAdapter.Fill(data);

                    var actualStock = "0";
                    var inStockFlag = false;
                    var groupSkus = "";
                    var simpleSkusList = new List<string>();

                    foreach (DataRow dr in data.Tables[0].Rows)
                    {
                        _logger.LogWrite("Working....");
                        var isStock = 0;
                        simpleSkusList = new List<string>();
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

                                    var size = "";
                                    if (i < 10)
                                    {
                                        size = dr["S0" + i].ToString();
                                    }
                                    else
                                    {
                                        size = dr["S" + i].ToString();
                                    }
                                    if (size.Contains("½"))
                                        size = size.Replace("½", ".5");
                                    if (!string.IsNullOrEmpty(size))
                                    {
                                        simpleSkusList.Add(groupSkus2);
                                        var newLine = BuildChildImportProduct(groupSkus2, dr, descriptions, reff, short_description, actualStock, descripto, size, isStock, reffColour, t2TreFs);
                                        csv.AppendLine(newLine);
                                    }
                                    
                                }
                                actualStock = "0";
                            }
                        }

                        isStock = inStockFlag ? 1 : 0;
                        if (!string.IsNullOrEmpty(dr["NewStyle"].ToString()))
                        {
                            var newLine = ParentImportProduct(groupSkus, descriptions, reff, dr, simpleSkusList, isStock, reffColour, t2TreFs);
                            csv.AppendLine(newLine);
                        }
                        inStockFlag = false;
                        if (data.Tables[0].Rows.Count > 1)
                        {
                            break;
                        }
                    }

                }
            }
            File.AppendAllText(@"C:\Users\Conor\Desktop\import_products_regen.csv", csv.ToString());
            Console.WriteLine("Job Finished");
            _logger.LogWrite("Finished");
        }

        private string ParentImportProduct(string groupSkus, List<Descriptions> descriptions, string reff, DataRow dr, List<string> simpleSkusList,
            int isStock, string reffColour, IEnumerable<string> t2TreFs)
        {
            var store = "\"admin\"";
            var websites = Websites().TrimEnd();
            var attribut_set = "\"Default\"";
            var type = "\"configurable\"";
            var sku = "\"" + groupSkus.TrimEnd() + "\"";
            var hasOption = "\"1\"";
            var name = "\"" + descriptions.Where(x => x.T2TRef == reff).Select(y => y.Descriptio).First() + " in " +
                       dr["MasterColour"] + "\"";
            var pageLayout = "\"No layout updates.\"";
            var optionsContainer = "\"Product Info Column\"";
            var price = "\"" + dr["BASESELL"].ToString().TrimEnd() + "\"";
            var weight = "\"0.01\"";
            var status = "\"Enabled\"";
            var visibility = Visibility().TrimEnd();
            var shortDescription = "\"" + BuildShortDescription(descriptions.First(x => x.T2TRef == reff)) + "\"";
            var gty = "\"0\"";
            var productName = "\"" + descriptions.Where(x => x.T2TRef == reff).Select(y => y.Descriptio).First() + "\"";
            var color = "\"" + dr["MasterColour"].ToString().TrimEnd() + "\"";
            var sizeRange = "\"\"";
            var vat = dr["VAT"].ToString() == "A" ? "TAX" : "None";
            var taxClass = "\"" + vat + "\"";
            var configurableAttribute = "\"size\"";
            var simpleSku = BuildSimpleSku(simpleSkusList, reff);
            var manufactor = "\"" + dr["MasterSupplier"] + "\"";
            var isInStock = "\"" + isStock + "\"";
            var category = "\"" + Category(dr) + "\"";
            var season = "\"\"";
            var stockType = "\"" + dr["GROUP"] + "\"";
            var image = "\"+/" + reffColour + ".jpg\"";
            var smallImage = "\"/" + reffColour + ".jpg\"";
            var thumbnail = "\"/" + reffColour + ".jpg\"";
            var gallery = "\"" + BuildGalleryImages(t2TreFs, reff) + "\"";
            var condition = "\"new\"";
            var ean = "\"\"";
            var description = "\"" + descriptions.Where(x => x.T2TRef == reff).Select(y => y.Description).First().TrimEnd() +
                              "\"";
            var model = "\"" + dr["SHORT"] + "\"";

            var newLine = $"{store}," +
                          $"{websites},{attribut_set},{type},{sku},{hasOption},{name.TrimEnd()},{pageLayout},{optionsContainer},{price},{weight},{status}," +
                          $"{visibility}," +
                          $"{shortDescription},{gty},{productName},{color}," +
                          $"{sizeRange},{taxClass},{configurableAttribute},{simpleSku},{manufactor},{isInStock}," +
                          $"{category},{season},{stockType},{image},{smallImage},{thumbnail},{gallery},{condition},{ean}," +
                          $"{description},{model}";
            return newLine;
        }

        private static string Category(DataRow dr)
        {
            var category = dr["MasterStocktype"] + "/Shop By Department/" + dr["MasterDept"] + ";;";

            if (dr["MasterSubDept"] != "ANY" || dr["MasterSubDept"] != "")
            {
                category = category + dr["MasterStocktype"] + "/Shop By Department/" +
                           dr["MasterDept"] + "/" + dr["MasterSubDept"] + "::1::1::0;;";
            }

            category = category + dr["MasterStocktype"] + "/Shop By Brand/" +
                       dr["MasterSupplier"] + ";;";
            category = category + dr["MasterStocktype"] + "/Shop By Brand/" +
                       dr["MasterSupplier"] + "/" + dr["MasterDept"] + "::1::1::0;;";

            if (dr["MasterSubDept"] != "ANY" || dr["MasterSubDept"] != "")
            {
                category = category + dr["MasterStocktype"] + "/Shop By Brand/" +
                           dr["MasterSupplier"] + "/" + dr["MasterDept"] +
                           "/" + dr["MasterSubDept"] + "::1::1::0;;";
            }
            category = category + "Brands/" + dr["MasterSupplier"];
            return category;
        }

        private static string BuildChildImportProduct(string groupSkus2, DataRow dr, List<Descriptions> descriptions, string reff,
            string short_description, string actualStock, string descripto, string size, int isStock, string reffColour,
            IEnumerable<string> t2TreFs)
        {
            const string store = "\"admin\"";
            var websites = Websites().TrimEnd();
            const string attribut_set = "\"Default\"";
            const string type = "\"simple\"";
            var sku = "\"" + groupSkus2.TrimEnd() + "\"";
            const string hasOption = "\"1\"";
            var name = "\"" + dr["MasterSupplier"] + " " + descriptions.Where(x => x.T2TRef == reff).Select(y => y.Descriptio).First() + " in " + dr["MasterColour"] + "\"";
            const string pageLayout = "\"No layout updates.\"";
            const string optionsContainer = "\"Product Info Column\"";
            var price = "\"" + dr["BASESELL"].ToString().TrimEnd() + "\"";
            const string weight = "\"0.01\"";
            const string status = "\"Enabled\"";
            const string visibility = "\"Not Visible Individually\"";
            var shortDescription = "\"" + short_description.TrimEnd() + "\"";
            var gty = "\"" + actualStock + "\"";
            var productName = "\"" + descripto.TrimEnd() + "\"";
            var color = "\"" + dr["MasterColour"].ToString().TrimEnd() + "\"";
            var sizeRange = "\"" + dr["SIZERANGE"] + size + "\"";
            var vat = dr["VAT"].ToString() == "A" ? "TAX" : "None";
            var taxClass = "\"" + vat + "\"";
            const string configurableAttribute = "\"\"";
            const string simpleSku = "\"\"";
            var manufactor = "\"" + dr["MasterSupplier"] + "\"";
            var isInStock = "\"" + isStock + "\"";
            const string category = "\"\"";
            const string season = "\"\"";
            var stockType = "\"" + dr["GROUP"] + "\"";
            var image = "\"+/" + reffColour + ".jpg\"";
            var smallImage = "\"/" + reffColour + ".jpg\"";
            var thumbnail = "\"/" + reffColour + ".jpg\"";
            var gallery = "\"" + BuildGalleryImages(t2TreFs, reff) + "\"";
            const string condition = "\"new\"";
            const string ean = "\"\"";
            var description = "\"" + descriptions.Where(x => x.T2TRef == reff).Select(y => y.Description).First().TrimEnd() +"\"";
            var model = "\"" + dr["SHORT"] + "\"";

            var newLine = $"{store}," +
                          $"{websites},{attribut_set},{type},{sku},{hasOption},{name.TrimEnd()},{pageLayout},{optionsContainer},{price},{weight},{status},{visibility}," +
                          $"{shortDescription},{gty},{productName},{color}," +
                          $"{sizeRange},{taxClass},{configurableAttribute},{simpleSku},{manufactor},{isInStock}," +
                          $"{category},{season},{stockType},{image},{smallImage},{thumbnail},{gallery},{condition},{ean}," +
                          $"{description},{model}";
            return newLine;
        }


        private static string BuildGalleryImages(IEnumerable<string> t2TreFs, string reff)
        {
            var images = t2TreFs.Where(t2tRef => t2tRef.Contains(reff)).ToList();
            return images.Aggregate("", (current, image) => current + ("/" + image + ".jpg;"));
        }

        private static string BuildSimpleSku(IEnumerable<string> t2TreFs, string reff)
        {
            var output = "\"";
            foreach (var t2tReff in t2TreFs)
            {
                if (t2tReff.Contains(reff))
                {
                    output += t2tReff + ",";
                }
            }
            return output.Remove(output.Length-2) + "\"";
        }

        private static string Websites()
        {
            var output = "\"";
            var fields = new string[]{"admin","base"};
            foreach (var field in fields)
            {
                    output += field + ",";
            }
            return output.Remove(output.Length - 1) + "\"";
        }

        private static string Visibility()
        {
            var output = "\"";
            var fields = new string[] { "Catalog", "Search" };
            foreach (var field in fields)
            {
                output += field + ",";
            }
            return output.Remove(output.Length - 1) + "\"";
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
                var bullet1 = string.IsNullOrEmpty(description.Bullet1) ? "" : "<li>" + description.Bullet1 + "</li>";
                var bullet2 = string.IsNullOrEmpty(description.Bullet2) ? "" : "<li>" + description.Bullet2 + "</li>";
                var bullet3 = string.IsNullOrEmpty(description.Bullet3) ? "" : "<li>" + description.Bullet3 + "</li>";
                var bullet4 = string.IsNullOrEmpty(description.Bullet4) ? "" : "<li>" + description.Bullet4 + "</li>";
                var bullet5 = string.IsNullOrEmpty(description.Bullet5) ? "" : "<li>" + description.Bullet5 + "</li>";
                var bullet6 = string.IsNullOrEmpty(description.Bullet6) ? "" : "<li>" + description.Bullet6 + "</li>";
                var bullet7 = string.IsNullOrEmpty(description.Bullet7) ? "" : "<li>" + description.Bullet7 + "</li>";
                return "<ul>" + bullet1 + bullet2 + bullet3 + bullet4 + bullet5 + bullet6 + bullet7 + "</ul>";
            }

        }

        private IEnumerable<string> ReadImageDetails(string path)
        {
            try
            {
                var imageDetails = Directory.GetFiles(path, "*.jpg*", SearchOption.AllDirectories)
                    .ToList();
                return ParseImageNames(imageDetails);
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

        //        private static IEnumerable<string> RemoveUnUsedValues(IEnumerable<string> imageDetails)
        //        {
        //            return imageDetails.Where(image => !image.Contains("_")).ToList();
        //        }
    }
}

//var newLine = $"{"admin"}," +
//              $"{Websites().Trim()}, {"Default"}, {"simple"}, {groupSkus2.Trim()}, {"1"}, {name.TrimEnd()},{"No layout updates."},{"Product Info Column"}, {dr["BASESELL"]}, {"0.01"}, {"Enabled"}, {"Not Visible Individually"}, " +
//              $"{short_description.TrimEnd()}, {actualStock}, {descripto.TrimEnd()}, {dr["MasterColour"]}, " +
//              $"{dr["SIZERANGE"] + size}, {vat}, {""},{""}, {dr["MasterSupplier"]}, {isStock}, " +
//              $"{""}, {""}, {dr["GROUP"]}, {"+/" + reffColour + ".jpg"}, {"/" + reffColour + ".jpg"},{"/" + reffColour + ".jpg"},{galleryImagesString.TrimEnd()}, {"new"}, {""}, " +
//              $"{descriptions.Where(x => x.T2TRef == reff).Select(y => y.Description).First().TrimEnd()}";

//var newLine2 = $"{"admin"}," +
//               $"{Websites()},{"Default"}, {"configurable"}, {groupSkus}, {"1"}, {name}, " +
//               $"{"No layout updates."},{"Product Info Column"},{dr["BASESELL"]}, {"0.01"}, {"Enabled"}, " +
//               $"{"Catalog Search"}, {short_description}, {0}, {descripto}, {dr["MasterColour"]}, " +
//               $"{""}, {vat}, {"size"},{simpleSkus}, {dr["MasterSupplier"]}, {isStock}, " +
//               $"{category}, {""}, {dr["GROUP"]}, {"+/" + reff + ".jpg"}, {"/" + reff + ".jpg"},{"/" + reff + ".jpg"},{galleryImagesString}, {"new"}, {""}, " +
//               $"{descriptions.Where(x => x.T2TRef == reff).Select(y => y.Description).First()}";