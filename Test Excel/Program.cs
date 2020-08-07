using ClosedXML.Excel;
using LinqToExcel;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Test_Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            string pathToExcelFile = @"data.xlsx";
            List<Product> dataList = new List<Product>();

            for (int i = 0; i < 500000; i++)
            {
                dataList.Add(new Product()
                {
                    ProductId = i,
                    CategoryName = "qqqq",
                    ProductName = "qqq",
                    Test = true,
                });
            }

            Console.WriteLine("start " +DateTime.Now);
            EPPlus_.Save(dataList, pathToExcelFile);
            Console.WriteLine("stop "+DateTime.Now);


            Console.WriteLine("start " + DateTime.Now);
            var open = EPPlus_.Open(pathToExcelFile);
            Console.WriteLine("stop " + DateTime.Now);


            Console.ReadKey();
        }

        class LinqToExcel_
        {
            //LinqToExcel
            //https://chrisbitting.com/2015/12/24/reading-excel-files-in-net-using-linqtoexcel/
            //https://github.com/paulyoder/LinqToExcel
            //https://www.c-sharpcorner.com/article/linq-to-excel-in-action/
            //https://www.youtube.com/watch?v=t3BEUP0OTFM

            public static List<Product> Open(string pathToExcelFile = "data.xlsx")
            {
                string sheetName = "data";

                var excelFile = new ExcelQueryFactory(pathToExcelFile);

                var artistAlbums = from a in excelFile.Worksheet<Product>(sheetName) select a;

                return artistAlbums.ToList();
            }
        }

        class ClosedXML_
        {
            //https://github.com/ClosedXML/ClosedXML/issues/619

            public static List<Product> Open(string fileName = @"data.xlsx")
            {
                var workbook = new XLWorkbook(fileName);
                var worksheet = workbook.Worksheet(1);
                var rows = worksheet.RangeUsed().RowsUsed(); // Skip header row

                Console.WriteLine("start test" + DateTime.Now);
                var product = rows.Select(w =>
                {
                    try { int tests = Convert.ToInt32(w.Cell(1).Value); }
                    catch { return null; }

                    return new Product()
                    {
                        ProductId = Convert.ToInt32(w.Cell(1).Value),
                        CategoryName = Convert.ToString(w.Cell(2).Value),
                        ProductName = Convert.ToString(w.Cell(3).Value),
                        Test = Convert.ToBoolean(w.Cell(4).Value),

                    };

                }).ToList();
                return product;
            }

            public static List<Product> Open2(string fileName = @"data.xlsx")
            {
                var workbook = new XLWorkbook(fileName);
                var worksheet = workbook.Worksheet(1);
                var rows = worksheet.RangeUsed().RowsUsed(); // Skip header row

                Console.WriteLine("start tessst" + DateTime.Now);
                var product = new List<Product>();
                foreach (var row in rows)
                {
                    try { int tests = Convert.ToInt32(row.Cell(1).Value); }
                    catch { continue; }

                    product.Add(new Product()
                    {
                        ProductId = Convert.ToInt32(row.Cell(1).Value),
                        CategoryName = Convert.ToString(row.Cell(2).Value),
                        ProductName = Convert.ToString(row.Cell(3).Value),
                        Test = Convert.ToBoolean(row.Cell(4).Value),

                    });
                }
                return product;
            }

            public static void Save(List<Product> dataList, string path = @"data.xlsx")
            {
                using (var workbook = new XLWorkbook())
                {     //creates the workbook
                    var wsDetailedData = workbook.AddWorksheet("data"); //creates the worksheet with sheetname 'data'
                    wsDetailedData.Cell(1, 1).InsertTable(dataList); //inserts the data to cell A1 including default column name
                    workbook.SaveAs(path); //saves the workbook
                    workbook.Dispose();
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }

            public static void SaveStream(List<Product> dataList, string path = @"data.xlsx")
            {
                using (var workbook = new XLWorkbook())
                {     //creates the workbook
                    var wsDetailedData = workbook.AddWorksheet("data"); //creates the worksheet with sheetname 'data'
                    wsDetailedData.Cell(1, 1).InsertTable(dataList); //inserts the data to cell A1 including default column name

                    using (MemoryStream memoryStream = SaveWorkbookToMemoryStream(workbook))
                    {
                        File.WriteAllBytes(path, memoryStream.ToArray());
                        memoryStream.Dispose();
                    }
                    workbook.Dispose();
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }

            private static MemoryStream SaveWorkbookToMemoryStream(XLWorkbook workbook)
            {
                using (MemoryStream stream = new MemoryStream())
                {
                    workbook.SaveAs(stream, new SaveOptions { EvaluateFormulasBeforeSaving = false, GenerateCalculationChain = false, ValidatePackage = false });
                    return stream;
                }
            }
        }

        class EPPlus_
        {
            public static List<Product> Open(string path = @"data.xlsx")
            {
                ExcelPackage.LicenseContext = LicenseContext.Commercial;
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var excel = new ExcelPackage(new FileInfo(path)))
                {
                    ExcelWorksheet sheet = excel.Workbook.Worksheets.FirstOrDefault();
                    var data = sheet.ImportExcelToList<Product>();
                    return data;
                }
            }

            public static void Save(List<Product> dataList, string path = @"data.xlsx")
            {
                ExcelPackage.LicenseContext = LicenseContext.Commercial;
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var excel = new ExcelPackage(new FileInfo(path)))
                {
                    ExcelWorksheet sheet = excel.Workbook.Worksheets.Add("date");
                    var table = sheet.Cells["A1"].LoadFromCollection(dataList, true, TableStyles.None);
                    excel.Save();
                }
            }
        }
    }

    public class Product
    {
        public int ProductId { get; set; } = new Random().Next();
        public string ProductName { get; set; } = "eqwe";
        public string CategoryName { get; set; } = "dasdas";
        public bool Test { get; set; } = false;
    }

    public static class ImportExcelReader
    {
        public static List<T> ImportExcelToList<T>(this ExcelWorksheet worksheet) where T : new()
        {
            //DateTime Conversion
            Func<double, DateTime> convertDateTime = new Func<double, DateTime>(excelDate =>
            {
                if (excelDate < 1)
                {
                    throw new ArgumentException("Excel dates cannot be smaller than 0.");
                }

                DateTime dateOfReference = new DateTime(1900, 1, 1);

                if (excelDate > 60d)
                {
                    excelDate = excelDate - 2;
                }
                else
                {
                    excelDate = excelDate - 1;
                }

                return dateOfReference.AddDays(excelDate);
            });

            ExcelTable table = null;

            if (worksheet.Tables.Any())
            {
                table = worksheet.Tables.FirstOrDefault();
            }
            else
            {
                table = worksheet.Tables.Add(worksheet.Dimension, "tbl" + worksheet.Name);

                ExcelAddressBase newaddy = new ExcelAddressBase(table.Address.Start.Row, table.Address.Start.Column, table.Address.End.Row + 1, table.Address.End.Column);

                //Edit the raw XML by searching for all references to the old address
                table.TableXml.InnerXml = table.TableXml.InnerXml.Replace(table.Address.ToString(), newaddy.ToString());
            }

            //Get the cells based on the table address
            List<IGrouping<int, ExcelRangeBase>> groups = table.WorkSheet.Cells[table.Address.Start.Row, table.Address.Start.Column, table.Address.End.Row, table.Address.End.Column]
                .GroupBy(cell => cell.Start.Row)
                .ToList();

            //Assume the second row represents column data types (big assumption!)
            List<Type> types = groups.Skip(1).FirstOrDefault().Select(rcell => rcell.Value.GetType()).ToList();

            //Get the properties of T
            List<PropertyInfo> modelProperties = new T().GetType().GetProperties().ToList();

            //Assume first row has the column names
            var colnames = groups.FirstOrDefault()
                .Select((hcell, idx) => new
                {
                    Name = hcell.Value.ToString(),
                    index = idx
                })
                .Where(o => modelProperties.Select(p => p.Name).Contains(o.Name))
                .ToList();

            //Everything after the header is data
            List<List<object>> rowvalues = groups
                .Skip(1) //Exclude header
                .Select(cg => cg.Select(c => c.Value).ToList()).ToList();

            //Create the collection container
            List<T> collection = new List<T>();
            foreach (List<object> row in rowvalues)
            {
                T tnew = new T();
                foreach (var colname in colnames)
                {
                    //This is the real wrinkle to using reflection - Excel stores all numbers as double including int
                    object val = row[colname.index];
                    Type type = types[colname.index];
                    PropertyInfo prop = modelProperties.FirstOrDefault(p => p.Name == colname.Name);

                    //If it is numeric it is a double since that is how excel stores all numbers
                    if (type == typeof(double))
                    {
                        //Unbox it
                        double unboxedVal = (double)val;

                        //FAR FROM A COMPLETE LIST!!!
                        if (prop.PropertyType == typeof(int))
                        {
                            prop.SetValue(tnew, (int)unboxedVal);
                        }
                        else if (prop.PropertyType == typeof(double))
                        {
                            prop.SetValue(tnew, unboxedVal);
                        }
                        else if (prop.PropertyType == typeof(DateTime))
                        {
                            prop.SetValue(tnew, convertDateTime(unboxedVal));
                        }
                        else if (prop.PropertyType == typeof(string))
                        {
                            prop.SetValue(tnew, val.ToString());
                        }
                        else
                        {
                            throw new NotImplementedException(string.Format("Type '{0}' not implemented yet!", prop.PropertyType.Name));
                        }
                    }
                    else
                    {
                        //Its a string
                        prop.SetValue(tnew, val);
                    }
                }
                collection.Add(tnew);
            }

            return collection;
        }
    }
}
