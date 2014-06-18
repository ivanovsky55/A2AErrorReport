using ClosedXML.Excel;
using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using System.Xml.XPath;

namespace ReportLib
{
    public class Report
    {

        public static void GenerateReport(string folderPath, string outFile, string dateTime)
        {
            Console.WriteLine("Starting Report Generation");
            var wb = new XLWorkbook();

            var docDetails = new XDocument(new XElement("DetailsRows"));
            var docOverview = new XDocument(new XElement("OverviewRows"));
            var docTotals = new XDocument(new XElement("TotalRows"));

            Console.WriteLine("Iterating through all markets in " + folderPath);
            foreach (string f in Directory.GetDirectories(folderPath))
            {
                if (!f.Contains("OCMS2DX"))
                    continue;

                //string market = new DirectoryInfo(f).Name.Replace("OCMS2Dx-", "");
                string market = new DirectoryInfo(f).Name.Replace("OCMS2DX-", "");
                string xmlPath = Path.Combine(f, "ConvertAssets.xml");

                //Testing flags, if you only want to run a few markets
                if (IsSkippableMarket(market)) continue;

                var length = new System.IO.FileInfo(xmlPath).Length / 1024 / 1024;
                Console.WriteLine("     [" + DateTime.Now.ToString("T") + "] [Size: " + length + " MB] Working on market: " + market);
                if (!File.Exists(xmlPath))
                {
                    Console.WriteLine("      Unable to find ConvertAssets.xml for this market. Skipping.");
                    continue;
                }

                //Fix German's messed up character
                if (market.ToLower() == "german")
                {
                    Console.WriteLine("        GermanFixStart");
                    string text = File.ReadAllText(xmlPath);
                    text = text.Replace("&#xB;", " ");
                    File.WriteAllText(xmlPath, text);
                    Console.WriteLine("        GermanFixEnd");
                }

                XDocument doc = XDocument.Load(xmlPath);
                foreach (XElement xe in doc.XPathSelectElements("//Asset"))
                {
                    string assetId = xe.XPathSelectElement("AssetId").Value;

                    if (!xe.XPathSelectElement("Errors").IsEmpty)
                    {
                        XElement xeError = xe.XPathSelectElement("Errors");
                        foreach (XElement err in xeError.XPathSelectElements("Error"))
                        {
                            string errorLevel = err.XPathSelectElement("Level").Value;
                            string errorMessage = err.XPathSelectElement("Message").Value;
                            errorMessage = errorMessage.Replace(System.Environment.NewLine, " ");
                            errorMessage = errorMessage.Replace('\n', ' ');
                            errorMessage = errorMessage.Replace('\r', ' ');
                            errorMessage = errorMessage.Replace("\r\n", " ");
                            string baseErrorMessage = GetSubMessage(errorMessage);

                            docDetails.Root.Add(new XElement("Row",
                                new XAttribute("Market", market),
                                new XAttribute("AssetId", assetId),
                                new XAttribute("ErrorLevel", errorLevel),
                                new XAttribute("BaseError", baseErrorMessage),
                                new XAttribute("FullError", errorMessage)));

                            ////OVERVIEW SECTION
                            //<OverviewRows>
                            //    <Row Market='English' Message='Asset Unpublished' Count='1' Type='Error'>
                            //      <AffectedAsset Id='HA1'>
                            //      <AffectedAsset Id='HA2'>
                            //    </Row>
                            //</OverviewRows>
                            string query = "//Row[@Market='" + market + "' and @Message='" + baseErrorMessage +
                                           "' and @Type='" + errorLevel + "']";
                            XElement xrow = docOverview.XPathSelectElement(query);
                            if (xrow != null)
                            {
                                //Already in the Xml, increase the count
                                xrow.Attribute("Count").SetValue((Convert.ToInt32(xrow.Attribute("Count").Value) + 1));

                                if (xrow.XPathSelectElement("AffectedAsset[@Id='" + assetId + "']") == null)
                                    xrow.Add(new XElement("AffectedAsset",
                                                new XAttribute("Id", assetId)));
                            }
                            else
                            {
                                //If not already in the Xml, add it, set count to 1
                                // ReSharper disable once PossibleNullReferenceException
                                docOverview.Root.Add(new XElement("Row",
                                    new XAttribute("Market", market),
                                    new XAttribute("Message", baseErrorMessage),
                                    new XAttribute("Type", errorLevel),
                                    new XAttribute("Count", 1),
                                        new XElement("AffectedAsset",
                                            new XAttribute("Id", assetId))));
                            }

                            ////TOTALS SECTION
                            //<TotalRows>
                            //  <Row Message='Asset Unpublished' Type='Error' TotalErrorsUS='1' TotalErrorsIntl='1' TotalAffectedAssetsUS='1' TotalAffectedAssetsIntl='1' />
                            //</TotalRows>

                            string queryTotal = "//Row[@Message='" + baseErrorMessage +
                                           "' and @Type='" + errorLevel + "']";
                            XElement xrowTotal = docTotals.XPathSelectElement(queryTotal);
                            if (xrowTotal != null)
                            {
                                //Already in the Xml, increase the ERROR count
                                if (market.ToLower() == "english")
                                {
                                    xrowTotal.Attribute("TotalErrorsUS")
                                        .SetValue((Convert.ToInt32(xrowTotal.Attribute("TotalErrorsUS").Value) + 1));
                                }
                                else
                                {
                                    xrowTotal.Attribute("TotalErrorsIntl")
                                        .SetValue((Convert.ToInt32(xrowTotal.Attribute("TotalErrorsIntl").Value) + 1));
                                }

                                if (xrowTotal.XPathSelectElement("AffectedAsset[@Id='" + assetId + "' and @Market='" + market + "']") == null)
                                {
                                    //Asset not in list, add it and update affected assets count
                                    xrowTotal.Add(new XElement("AffectedAsset",
                                        new XAttribute("Id", assetId),
                                            new XAttribute("Market", market)));

                                    if (market.ToLower() == "english")
                                    {
                                        xrowTotal.Attribute("TotalAffectedAssetsUS")
                                            .SetValue((Convert.ToInt32(xrowTotal.Attribute("TotalAffectedAssetsUS").Value) + 1));
                                    }
                                    else
                                    {
                                        xrowTotal.Attribute("TotalAffectedAssetsIntl")
                                            .SetValue((Convert.ToInt32(xrowTotal.Attribute("TotalAffectedAssetsIntl").Value) + 1));
                                    }
                                }
                            }
                            else
                            {
                                //not in the Xml, add it, set count to 1 for US or Intl

                                int usVal = 0;
                                int intlVal = 0;
                                if (market.ToLower() == "english")
                                {
                                    usVal = 1;
                                }
                                else
                                {
                                    intlVal = 1;
                                }


                                // ReSharper disable once PossibleNullReferenceException
                                docTotals.Root.Add(new XElement("Row",
                                    new XAttribute("Message", baseErrorMessage),
                                    new XAttribute("Type", errorLevel),
                                    new XAttribute("TotalErrorsUS", usVal),
                                    new XAttribute("TotalErrorsIntl", intlVal),
                                    new XAttribute("TotalAffectedAssetsUS", usVal),
                                    new XAttribute("TotalAffectedAssetsIntl", intlVal),
                                        new XElement("AffectedAsset",
                                            new XAttribute("Id", assetId),
                                            new XAttribute("Market", market))));
                            }
                        }
                    }
                }
            }
            docDetails.Save(Path.Combine(Path.GetDirectoryName(outFile), dateTime + "_Details.xml"));
            docOverview.Save(Path.Combine(Path.GetDirectoryName(outFile), dateTime + "_Overview.xml"));
            docTotals.Save(Path.Combine(Path.GetDirectoryName(outFile), dateTime + "_Totals.xml"));

            Console.WriteLine("Creating and formatting spreadsheet.");

            //docDetails = XDocument.Load(@"\\WIN-E5U3RQA9QQH\MigrationTransfer\A2A_Reports\2014-05-28-10-46_Details.xml");
            //docOverview = XDocument.Load(@"\\WIN-E5U3RQA9QQH\MigrationTransfer\A2A_Reports\2014-05-28-10-46_Overview.xml");
            //docTotals = XDocument.Load(@"\\WIN-E5U3RQA9QQH\MigrationTransfer\A2A_Reports\2014-05-28-10-46_Totals.xml");

            //Add Details sheet

            IXLWorksheet wsDetails = wb.Worksheets.Add("Details");
            wsDetails.Cell("A1").Value = "Market";
            wsDetails.Cell("B1").Value = "AssetId";
            wsDetails.Cell("C1").Value = "ErrorLevel";
            wsDetails.Cell("D1").Value = "BaseError";
            wsDetails.Cell("E1").Value = "FullError";
            int detailsRow = 2;
            foreach (XElement xeRow in docDetails.XPathSelectElements("//Row"))
            {
                //Not enough memory to add the bazillion detail lines on these items.
                if (xeRow.Attribute("BaseError").Value == "AltText Mismatch") continue;
                if (xeRow.Attribute("BaseError").Value == "Text for Token does not match the current content text. Removing token and injecting content text instead.") continue;
                if (xeRow.Attribute("BaseError").Value == "Header encountered without a bookmark, creating new bookmark") continue;

                wsDetails.Cell("A" + detailsRow).Value = xeRow.Attribute("Market").Value;
                wsDetails.Cell("B" + detailsRow).Value = xeRow.Attribute("AssetId").Value;
                wsDetails.Cell("C" + detailsRow).Value = xeRow.Attribute("ErrorLevel").Value;
                wsDetails.Cell("D" + detailsRow).Value = xeRow.Attribute("BaseError").Value;
                wsDetails.Cell("E" + detailsRow).Value = xeRow.Attribute("FullError").Value;
                detailsRow++;
            }
            docDetails = null;

            //Add Overview sheet
            IXLWorksheet wsOverview = wb.Worksheets.Add("Overview");
            wsOverview.Cell("A1").Value = "Market";
            wsOverview.Cell("B1").Value = "Message";
            wsOverview.Cell("C1").Value = "ErrorCount";
            wsOverview.Cell("D1").Value = "AffectedAssets";
            wsOverview.Cell("E1").Value = "Type";

            int overviewRow = 2;
            foreach (XElement xeRow in docOverview.XPathSelectElements("//Row"))
            {
                var xeAffectedAssets = xeRow.XPathSelectElements("AffectedAsset").Count();

                wsOverview.Cell("A" + overviewRow).Value = xeRow.Attribute("Market").Value;
                wsOverview.Cell("B" + overviewRow).Value = xeRow.Attribute("Message").Value;
                wsOverview.Cell("C" + overviewRow).Value = xeRow.Attribute("Count").Value;
                wsOverview.Cell("D" + overviewRow).Value = xeAffectedAssets;
                wsOverview.Cell("E" + overviewRow).Value = xeRow.Attribute("Type").Value;
                overviewRow++;
            }
            docOverview = null;

            //Add Totals sheet
            IXLWorksheet wsTotals = wb.Worksheets.Add("Totals");
            wsTotals.Cell("A1").Value = "Message";
            wsTotals.Cell("B1").Value = "Type";
            wsTotals.Cell("C1").Value = "TotalErrorsUS";
            wsTotals.Cell("D1").Value = "TotalErrorsIntl";
            wsTotals.Cell("E1").Value = "TotalAffectedAssetsUS";
            wsTotals.Cell("F1").Value = "TotalAffectedAssetsIntl";

            int totalsRow = 2;
            foreach (XElement xeRow in docTotals.XPathSelectElements("//Row"))
            {
                wsTotals.Cell("A" + totalsRow).Value = xeRow.Attribute("Message").Value;
                wsTotals.Cell("B" + totalsRow).Value = xeRow.Attribute("Type").Value;
                wsTotals.Cell("C" + totalsRow).Value = xeRow.Attribute("TotalErrorsUS").Value;
                wsTotals.Cell("D" + totalsRow).Value = xeRow.Attribute("TotalErrorsIntl").Value;
                wsTotals.Cell("E" + totalsRow).Value = xeRow.Attribute("TotalAffectedAssetsUS").Value;
                wsTotals.Cell("F" + totalsRow).Value = xeRow.Attribute("TotalAffectedAssetsIntl").Value;
                totalsRow++;
            }
            docTotals = null;

            IXLRange range = wsDetails.RangeUsed();
            wsDetails.Range("A1:" + wsDetails.RangeUsed().RangeAddress.LastAddress).SetAutoFilter();
            range.SetAutoFilter();

            IXLRange rangeOverview = wsOverview.RangeUsed();
            wsOverview.Range("A1:" + wsOverview.RangeUsed().RangeAddress.LastAddress).SetAutoFilter();
            rangeOverview.SetAutoFilter();

            IXLRange rangeTotals = wsTotals.RangeUsed();
            wsTotals.Range("A1:" + wsTotals.RangeUsed().RangeAddress.LastAddress).SetAutoFilter();
            rangeTotals.SetAutoFilter();

            wsDetails.Column(1).Width = 19;
            wsDetails.Column(2).Width = 12;
            wsDetails.Column(3).Width = 12;
            wsDetails.Column(4).Width = 40;
            wsDetails.Column(5).Width = 50;

            wsOverview.Column(1).Width = 19;
            wsOverview.Column(2).Width = 60;
            wsOverview.Column(3).Width = 13;
            wsOverview.Column(4).Width = 17;
            wsOverview.Column(5).Width = 8;

            wsTotals.Column(1).Width = 60;
            wsTotals.Column(2).Width = 8;
            wsTotals.Column(3).Width = 16;
            wsTotals.Column(4).Width = 16;
            wsTotals.Column(5).Width = 24;
            wsTotals.Column(6).Width = 24;

            wb.SaveAs(outFile);
        }

        private static bool IsSkippableMarket(string market)
        {
            var mk = market.ToLower();

            ////Run some markets
            //if (mk != "english"
            // && mk != "chinese-simplified"
            // && mk != "french"
            // && mk != "portuguese-brazil"
            // && mk != "german"
            // && mk != "japanese"
            // && mk != "russian"
            // && mk != "spanish"
            // )return true;


            //Skip some markets
            //if (mk == "albanian") return true;
            //if (mk == "amharic") return true;
            //if (mk == "arabic") return true;
            //if (mk == "azerbaijani") return true;
            //if (mk == "bangla") return true;
            //if (mk == "basque") return true;
            //if (mk == "belarusian") return true;
            //if (mk == "bulgarian") return true;
            //if (mk == "catalan") return true;
            //if (mk == "chinese-simplified") return true;
            //if (mk == "chinese-traditional") return true;
            //if (mk == "croatian") return true;
            //if (mk == "czech") return true;
            //if (mk == "danish") return true;
            //if (mk == "dutch") return true;
            //if (mk == "english") return true;
            //if (mk == "german") return true;
            //if (mk == "hebrew") return true;
            //if (mk == "slovenian") return true;
            //if (mk == "estonian") return true;
            //if (mk == "filipino") return true;
            //if (mk == "finnish") return true;
            //if (mk == "french") return true;
            //if (mk == "hindi") return true;
            //if (mk == "vietnamese") return true;
            //if (mk == "indonesian") return true;
            //if (mk == "malay") return true;
            return false;
        }

        public static string GetSubMessage(string s)
        {
            string result;
            if (s == null) return "";
            int l = s.IndexOf(":", StringComparison.Ordinal);
            result = l > 0 ? s.Substring(0, l) : s;
            result = Regex.Replace(result, "^Bookmark\\ (.*?)\\ maps\\ to\\ different\\ aliases$", "Bookmark maps to different aliases");
            result = Regex.Replace(result, "^The\\ TOC\\ instruction\\ .*?\\ has\\ excluded\\ the\\ following\\ Headings\\ from\\ being\\ included\\ in\\ the\\ TOC$", "TOC Exclusion");
            result = Regex.Replace(result, "^Token\\ .*?\\ inserted\\ within\\ a\\ .*?\\ construct,\\ which\\ is\\ not\\ allowed\\ in\\ DDUEML\\ schema\\.\\ Replacing\\ with\\ token\\ text\\ .*?\\ instead\\.$", "Token inserted within a construct that is not allowed in DDUEML schema.");
            result = Regex.Replace(result, "^Unable\\ to\\ resolve\\ external\\ .*?$", "Unable to resolve external rId");
            result = Regex.Replace(result, "^Exception$", "Object reference not set to an instance of an object");
            result = Regex.Replace(result, "^Token\\ .*?\\ inserted\\ within\\ TOC\\ construct,\\ which\\ is\\ not\\ allowed\\ in\\ DDUEML\\ schema\\.\\ Replacing\\ with\\ token.*?$", "Token inserted within a construct that is not allowed in DDUEML schema.");
            result = Regex.Replace(result, "^Token\\ .*?\\ encountered\\ inside\\ a\\ .*?,\\ which\\ is\\ not\\ allowed\\ in\\ DDUEML\\ schema\\.\\ Replaced\\ with\\ token\\ text .*?$", "Token inserted within a construct that is not allowed in DDUEML schema.");
            result = Regex.Replace(result, "^Token\\ .*?\\ encountered\\ inside\\ TOC,\\ which\\ is\\ not\\ allowed\\ in\\ DDUEML\\ schema\\.\\ Replacing\\ with\\ token.*?$", "Token inserted within a construct that is not allowed in DDUEML schema.");
            result = Regex.Replace(result, "^Discarded.*?\\ unsupported\\ Content\\ Control\\ .*?$", "Discarded unsupported Content Control");
            result = Regex.Replace(result, "Multiple\\ nodes\\ found\\ for", "Multiple nodes found for Wordml attribute");
            result = Regex.Replace(result, "^Asset\\ contains\\ .*?\\ instance\\(s\\)\\ of\\ the\\ ExcelWorkbook\\ Content\\ Control\\.\\ Please\\ see\\ OM$", "Asset contains instances of the ExcelWorkbook Content Control. Please see OM");
            result = Regex.Replace(result, "^Reference\\ to\\ VA\\ Asset\\ .*?\\ which\\ does\\ not\\ have\\ an\\ MSN\\ UUID\\.$", "Reference to VA Asset which does not have an MSN UUID.");
            result = Regex.Replace(result, "^Reference\\ undefined\\ or\\ deleted\\ VA\\ Asset\\ .*?$", "Reference undefined or deleted VA Asset");
            result = Regex.Replace(result, "^\"?Text\\ for\\ .*?\\ Token\\ .*?$", "Text for Token does not match the current content text. Removing token and injecting content text instead.");
            result = Regex.Replace(result, "^This\\ asset\\ references\\ non-searchable\\ VA\\ asset\\ .*?\\ which\\ will\\ cause\\ this\\ VA\\ asset\\ to\\ be\\ included\\ in\\ the\\ migration\\.$", "This asset references a non-searchable VA asset, which will cause this VA asset to be included in the migration.");
            result = Regex.Replace(result, "^\"?Document\\ .*\\ AltText\\ \\(.*$", "AltText Mismatch");
            result = Regex.Replace(result, "^\\d*\\ Token\\(s\\)\\ in\\ TOC,\\ check\\ TOC\\ for\\ missing\\ or\\ duplicate\\ TOC\\ text\\.$", "Token in TOC");
            result = Regex.Replace(result, "^Alert\\ .*\\ found\\ nested\\ inside\\ .*\\..*$", "Unsupported nested Alert element");
            result = Regex.Replace(result, "^Asset\\ contains\\ \\d*\\ sub-TOC\\(s\\),.*$", "Asset contains sub-TOCs");
            result = Regex.Replace(result, "^Expandos\\ are\\ used\\ \\d*\\ time\\(s\\)\\ in\\ this\\ asset\\.$", "Asset contains expandos");
            result = Regex.Replace(result, "^.*\\ item\\ encountered\\ without\\ a\\ bookmark,\\ creating\\ new\\ bookmark$", "Header encountered without a bookmark, creating new bookmark");
            result = Regex.Replace(result, "^Reference\\ to\\ .*\\ Asset.*which\\ has\\ known\\ errors.$", "Reference to asset that has known errors");
            //result = Regex.Replace(result, "REGEX_HERE", "DESIRED_MESSAGE_HERE");
            //result = Regex.Replace(result, "REGEX_HERE", "DESIRED_MESSAGE_HERE");
            return result.Replace("'", " ");
        }
    }
}
