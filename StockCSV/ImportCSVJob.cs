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
                    $"{"store"},{"websites"},{"attribut_set"},{"type"},{"sku"},{"has_options"}, {"name"}, {"page_layout"}, {"options_container"}, {"price"}, {"weight"}, {"status"}, {"visibility"}, {"short_description"}, {"qty"}, {"product_name"}, {"color"}," +
                    $" {"size"},{"tax_class_id"}, {"configurable_attributes"}, {"simples_skus"}, {"manufacturer"}, {"is_in_stock"}, {"categories"}, {"season"}, {"stock_type"}, {"image"}, {"small_image"}, {"thumbnail"}, {"gallery"}," +
                    $" {"condition"}, {"ean"}, {"description"}";

                csv.AppendLine(headers);
                foreach (var refff in t2TreFs.Where(x => !x.Contains("_")))
                {
                    var reff = refff.Substring(0, 6);

                    var data = new DataSet();
                    var myAccessCommand = new OleDbCommand(SqlQuery.ImportProductsQuery, connectionHandler);
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
                                        var galleryImagesString = BuildGalleryImages(t2TreFs, reff);
                                        var vat = dr["VAT"].ToString() == "A" ? "TAX" : "NONE";
                                        var name = dr["MasterSupplier"] + "  " + descriptions.Where(x => x.T2TRef == reff).Select(y => y.Descriptio).First() + " in " + dr["MasterColour"];
                                        var newLine = $"{"admin"}, {"admin base"}, {"Default"}, {"simple"}, {groupSkus2}, {"1"}, {name},{"No layout updates."},{"Product Info Column"}, {dr["BASESELL"]}, {"0.01"}, {"Enabled"}, {"Not Visible Individually"}, " +
                                                      $"{short_description}, {actualStock}, {descripto}, {dr["MasterColour"]}, " +
                                                      $"{dr["SIZERANGE"] + size}, {vat}, {""},{""}, {dr["MasterSupplier"]}, {isStock}, " +
                                                      $"{"TODO categories"}, {""}, {dr["GROUP"]}, {"+/" + reff + ".jpg"}, {"/" + reff + ".jpg"},{"/" + reff + ".jpg"},{galleryImagesString}, {"new"}, {""}, " +
                                                      $"{descriptions.Where(x => x.T2TRef == reff).Select(y => y.Description).First()}";
                                        //var newLine = $"{groupSkus2},{actualStock},{isStock}";"
                                        csv.AppendLine(newLine);
                                    }
                                    
                                }
                                actualStock = "0";
                            }
                        }

                        isStock = inStockFlag ? 1 : 0;
                        if (!string.IsNullOrEmpty(dr["NewStyle"].ToString()))
                        {
                            var name = descriptions.Where(x => x.T2TRef == reff).Select(y => y.Descriptio).First() + " in " + dr["MasterColour"];
                            short_description = BuildShortDescription(descriptions.First(x => x.T2TRef == reff));
                            var descripto = descriptions.Where(x => x.T2TRef == reff)
                                .Select(y => y.Descriptio).First();
                            var galleryImagesString = BuildGalleryImages(t2TreFs, reff);
                            var vat = dr["VAT"].ToString() == "A" ? "TAX" : "NONE";
                            var newLine2 = $"{"admin"}, {"admin base"}, {"Default"}, {"configurable"}, {groupSkus}, {"1"}, {name}, " +
                                           $"{"No layout updates."},{"Product Info Column"},{dr["BASESELL"]}, {"0.01"}, {"Enabled"}, " +
                                           $"{"Catalog Search"}, {short_description}, {0}, {descripto}, {dr["MasterColour"]}, " +
                                           $"{""}, {vat}, {"size"},{"TODO"}, {dr["MasterSupplier"]}, {isStock}, " +
                                           $"{"TODO categories"}, {""}, {dr["GROUP"]}, {"+/" + reff + ".jpg"}, {"/" + reff + ".jpg"},{"/" + reff + ".jpg"},{galleryImagesString}, {"new"}, {""}, " +
                                           $"{descriptions.Where(x => x.T2TRef == reff).Select(y => y.Description).First()}";
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
            File.AppendAllText(@"C:\Users\Conor\Desktop\import_products_regen.csv", csv.ToString());
            Console.WriteLine("Job Finished");
            _logger.LogWrite("Finished");
        }

        private string BuildGalleryImages(IEnumerable<string> t2TreFs, string reff)
        {
            var images = t2TreFs.Where(t2tRef => t2tRef.Contains(reff)).ToList();
            return images.Aggregate("", (current, image) => current + ("/" + image + ".jpg;"));
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