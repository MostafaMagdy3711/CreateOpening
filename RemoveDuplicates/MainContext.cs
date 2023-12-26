using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using HtmlAgilityPack;
using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Xml;
using System.Xml.Linq;


namespace RemoveDuplicates
{
    //Ahmed Sherif
    [Transaction(TransactionMode.Manual)]
    public class MainContext : IExternalCommand
    {
        public static UIDocument UiDoc;
        public static Document Doc;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UiDoc = commandData.Application.ActiveUIDocument;
            Doc = UiDoc.Document;
            string ExcelFilePath;
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "HTML Files|*.html;*.htm",
                Title = "Select HTML File"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                ExcelFilePath = openFileDialog.FileName;

                StringBuilder sb = new StringBuilder();
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                List<int> Elements_Id_1 = ExtractNumbersFromHtml(ExcelFilePath, 1);
                List<int> Elements_Id_2 = ExtractNumbersFromHtml(ExcelFilePath, 2);




                for (int i = 0; i < Math.Min(Elements_Id_2.Count, Elements_Id_1.Count); i++)
                {
                    ElementId Id_1 = new ElementId(Elements_Id_1[0]);
                    ElementId Id_2 = new ElementId(Elements_Id_2[0]);

                    Element E1 = Doc.GetElement(Id_1);
                    Element E2 = Doc.GetElement(Id_2);

                    double Vol1 = GetElementSolid(E1)?.Volume ?? 0;
                    double Vol2 = GetElementSolid(E2)?.Volume ?? 0;

                    using (Transaction trans = new Transaction(Doc, "Delete Duplicate"))
                    {
                        trans.Start();
                        try
                        {
                            if (Vol1 < Vol2)
                            {
                                Doc.Delete(Id_1);
                            }
                            else
                            {
                                Doc.Delete(Id_2);
                            }
                        }
                        catch
                        {

                        }
                        trans.Commit();

                    }



                }


            }


            return Result.Succeeded;
        }


        static int ExtractNumber(string input)
        {
            // Define a regular expression pattern to match the number
            string pattern = @"\d+";

            // Use Regex.Match to find the first match in the input string
            Match match = Regex.Match(input, pattern);

            // Check if a match was found
            if (match.Success)
            {
                // Get the matched value and convert it to an integer
                if (int.TryParse(match.Value, out int result))
                {
                    return result;
                }
            }

            // Return -1 if no match or parsing fails
            return -1;
        }

        public static Solid GetElementSolid(Element element)
        {


            Solid solid = default;
            Options OP = new Options();
            OP.DetailLevel = ViewDetailLevel.Fine;
            OP.ComputeReferences = true;
            try
            {
                GeometryElement columnGeometry = element.get_Geometry(OP);
                foreach (GeometryObject geomObj in columnGeometry)
                {
                    GeometryInstance GeoIns = geomObj as GeometryInstance;

                    var GeoInstance = GeoIns.GetInstanceGeometry();
                    foreach (var item in GeoInstance)
                    {
                        solid = item as Solid;
                        if (solid != null && solid.Volume != 0)
                        {
                            return item as Solid;
                        }
                    }
                }

                return null;
            }
            catch
            {
                GeometryElement columnGeometry = element.get_Geometry(OP);
                foreach (var GeoEle in columnGeometry)
                {
                    solid = GeoEle as Solid;
                    if (solid != null && solid.Volume != 0)
                    {
                        return solid;
                    }
                }

                return null;

            }






        }


        static List<int> ExtractNumbersFromHtml(string filePath, int ItemNum)
        {
            List<int> numbers = new List<int>();

            if (ItemNum == 1)
            {
                try
                {
                    // Read the entire content of the HTML file
                    string htmlContent = File.ReadAllText(filePath);

                    // Define a regular expression pattern to match lines containing numbers
                    //string pattern = @"<i>Element ID</i>";
                    string pattern = @"<td class=""item1Content""><i>Element ID</i>\s*:\s*(\d+)</td>";

                    // Match the pattern in the entire HTML content
                    MatchCollection matches = Regex.Matches(htmlContent, pattern);

                    // Process each match and extract the number
                    foreach (Match match in matches)
                    {
                        var ss = match.Value.Split('D');
                        int x = ExtractNumber(ss[1]);

                        numbers.Add(x);

                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"An error occurred: {ex.Message}");
                }
            }

            else if (ItemNum == 2)
            {
                try
                {
                    // Read the entire content of the HTML file
                    string htmlContent = File.ReadAllText(filePath);

                    // Define a regular expression pattern to match lines containing numbers
                    //string pattern = @"<i>Element ID</i>";
                    string pattern = @"<td class=""item2Content""><i>Element ID</i>\s*:\s*(\d+)</td>";

                    // Match the pattern in the entire HTML content
                    MatchCollection matches = Regex.Matches(htmlContent, pattern);

                    // Process each match and extract the number
                    foreach (Match match in matches)
                    {
                        var ss = match.Value.Split('D');
                        int x = ExtractNumber(ss[1]);

                        numbers.Add(x);

                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"An error occurred: {ex.Message}");
                }
            }


            return numbers;
        }


    }
}
