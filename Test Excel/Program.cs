using ClosedXML.Excel;
using LinqToExcel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Test_Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            string pathToExcelFile = @"data.xlsx";
            List<Product> dataList = new List<Product>()
            {
                //new Product()
                // {
                //     ProductId = 2,
                //     CategoryName = "wwww",
                //     ProductName ="wwww",
                //     Test = true,
            };

            for (int i = 0; i < 100000; i++)
            {
                dataList.Add(new Product()
                {
                    ProductId = i,
                    CategoryName = "qqqq",
                    ProductName = "qqq",
                    Test = true,
                });
            }

            Console.WriteLine(DateTime.Now);

            ClosedXML_.SaveTableStream(dataList.ToArray(), "datas.xlsx");

            Console.WriteLine(DateTime.Now);

            Console.ReadKey();
        }

        class LinqToExcel_
        {
            //LinqToExcel
            //https://chrisbitting.com/2015/12/24/reading-excel-files-in-net-using-linqtoexcel/
            //https://github.com/paulyoder/LinqToExcel
            //https://www.c-sharpcorner.com/article/linq-to-excel-in-action/
            //https://www.youtube.com/watch?v=t3BEUP0OTFM


            public static List<Product> OpenTable1(string pathToExcelFile = @"data.xlsx")
            {
                ConnexionExcel ConxObject = new ConnexionExcel(pathToExcelFile);
                var test = ConxObject.UrlConnexion.WorksheetNoHeader("data")
                    .ToList()
                    .Select(x =>
                    {
                        try { int tests = Convert.ToInt32(x.First().Value); }
                        catch { return null; }

                        return new Product()
                        {
                            ProductId = Convert.ToInt32(x[0].Value),
                            CategoryName = Convert.ToString(x[1].Value),
                            ProductName = Convert.ToString(x[2].Value),
                            Test = x[3].Value.ToString() == "ИСТИНА" ? true : false,
                        }; ;
                    })
                    .ToList();
                test.RemoveAll(x => x == null);
                return test;
            }

            public static List<Product> OpenTable2(string pathToExcelFile = "data.xlsx")
            {
                string sheetName = "data";

                var excelFile = new ExcelQueryFactory(pathToExcelFile);

                var artistAlbums = from a in excelFile.Worksheet<Product>(sheetName) select a;

                return artistAlbums.ToList();
            }

            public static void SaveTable(string pathToExcelFile = "data.xlsx")
            {
                
            }
        }

        class ClosedXML_
        {
            //https://github.com/ClosedXML/ClosedXML/issues/619
            //

            public static void OpenTable2(string fileName = @"data.xlsx")
            {
                Console.WriteLine("start" + DateTime.Now);
                var workbook = new XLWorkbook(fileName);
                var worksheet = workbook.Worksheet(1);
                // получим все строки в файле
                var rows = worksheet.RangeUsed().RowsUsed(); // Skip header row

                Console.WriteLine("start test" + DateTime.Now);
                var test = rows.Select(w =>
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
                Console.WriteLine("stop test" + DateTime.Now);
            }

            public static void OpenTable1(string fileName = @"data.xlsx")
            {
                Console.WriteLine("start" + DateTime.Now);
                var workbook = new XLWorkbook(fileName);
                var worksheet = workbook.Worksheet(1);
                // получим все строки в файле
                var rows = worksheet.RangeUsed().RowsUsed(); // Skip header row

                Console.WriteLine("start tessst" + DateTime.Now);
                var tessst = new List<Product>();
                foreach (var row in rows)
                {
                    try { int tests = Convert.ToInt32(row.Cell(1).Value); }
                    catch { continue; }

                    tessst.Add(new Product()
                    {
                        ProductId = Convert.ToInt32(row.Cell(1).Value),
                        CategoryName = Convert.ToString(row.Cell(2).Value),
                        ProductName = Convert.ToString(row.Cell(3).Value),
                        Test = Convert.ToBoolean(row.Cell(4).Value),

                    });
                    //string rowNumber = $"val1 {row.Cell(1).Value} val2 {row.Cell(2).Value} val3 {row.Cell(3).Value} val4 {row.Cell(4).Value}";
                    //Console.WriteLine(rowNumber);
                }
                Console.WriteLine("stop tessst" + DateTime.Now);
            }

            public static void SaveTable(Product[] dataList, string path = @"data.xlsx")
            {
                var workbook = new XLWorkbook();     //creates the workbook
                var wsDetailedData = workbook.AddWorksheet("data"); //creates the worksheet with sheetname 'data'
                wsDetailedData.Cell(1, 1).InsertTable(dataList); //inserts the data to cell A1 including default column name
                workbook.SaveAs(path); //saves the workbook
                workbook.Dispose();
            }

            public static void SaveTableStream(Product[] dataList, string path = @"data.xlsx")
            {
                var workbook = new XLWorkbook();     //creates the workbook
                var wsDetailedData = workbook.AddWorksheet("data"); //creates the worksheet with sheetname 'data'
                wsDetailedData.Cell(1, 1).InsertTable(dataList); //inserts the data to cell A1 including default column name

                using (MemoryStream memoryStream = SaveWorkbookToMemoryStream(workbook))
                {
                    File.WriteAllBytes(path, memoryStream.ToArray());
                    memoryStream.Dispose();
                }
                workbook.Dispose();
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

        public class ConnexionExcel
        {
            public string _pathExcelFile;
            public ExcelQueryFactory _urlConnexion;
            public ConnexionExcel(string path)
            {
                this._pathExcelFile = path;
                this._urlConnexion = new ExcelQueryFactory(_pathExcelFile);
            }
            public string PathExcelFile
            {
                get
                {
                    return _pathExcelFile;
                }
            }
            public ExcelQueryFactory UrlConnexion
            {
                get
                {
                    return _urlConnexion;
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
    }
}
