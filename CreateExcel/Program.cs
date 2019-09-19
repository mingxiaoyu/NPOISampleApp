using System;
using NPOI.HSSF.UserModel;
using NPOI.HPSF;
using NPOI.POIFS.FileSystem;
using NPOI.SS.UserModel;
using Bogus;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Collections.Generic;
using System.Data;
using NPOI.XSSF.UserModel;
using Newtonsoft.Json;

namespace CreateExcel
{
    public class Model
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public DateTimeOffset Now { get; set; }
        public bool New { get; set; }
    }
    /// <summary>
    /// HSSF类，只支持2007以前的excel（文件扩展名为xls），而XSSH支持07以后的
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            Randomizer.Seed = new Random(8675309);

            var faker = new Faker<Model>()
                .RuleFor(x => x.Id, f => f.Random.Int(0, 100))
                .RuleFor(x => x.Name, f => f.Name.FullName())
                .RuleFor(x => x.Now, f =>
                {
                    var date = f.Date.Past();
                    return new DateTimeOffset(date);
                })
                .RuleFor(x => x.New, f => f.Random.Bool())
            ;


            var lists = faker.Generate(10);

            WriteExcel_One(lists);
            WriteExcel(lists);
            Console.WriteLine("Hello World!");
        }

        private static void SetCell(IWorkbook hssfworkbook, ICell cell, PropertyInfo propertyInfo, Model model)
        {
            ICellStyle cellStyle = hssfworkbook.CreateCellStyle();
            IDataFormat format = hssfworkbook.CreateDataFormat();
            cellStyle.DataFormat = format.GetFormat("yyyy-MM-dd HH:mm:ss");

            var value = propertyInfo.GetValue(model);
            switch (value)
            {
                case string s:
                    cell.SetCellValue($"{s.Replace("\"", "\"\"")}");
                    break;
                case DateTimeOffset dto:
                    cell.SetCellValue(dto.UtcDateTime);
                    cell.CellStyle = cellStyle;
                    break;
                case DateTime dto:
                    cell.SetCellValue(dto);
                    break;
                case bool b:
                    cell.SetCellValue(b);
                    break;
                default:
                    cell.SetCellValue(Convert.ToString(value));
                    break;
            }
        }

        private static string FormatRow(PropertyInfo propertyInfo, Model model)
        {
            var value = propertyInfo.GetValue(model);
            switch (value)
            {
                case string s: return $"\"{s.Replace("\"", "\"\"")}\"";
                case DateTimeOffset dto: return dto.ToString("yyyy-MM-dd HH:mm:ss");
                case bool b: return b ? "True" : "False";
                default: return Convert.ToString(value);
            }
        }

        static void WriteExcel_One(List<Model> list)
        {
            IWorkbook hssfworkbook;

            // InitializeWorkbook
            hssfworkbook = new HSSFWorkbook();

            ////Create a entry of DocumentSummaryInformation
            //DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            //dsi.Company = "NPOI Team";
            //hssfworkbook.DocumentSummaryInformation = dsi;

            ////Create a entry of SummaryInformation
            //SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            //si.Subject = "NPOI SDK Example";
            //hssfworkbook.SummaryInformation = si;

            ISheet sheet = hssfworkbook.CreateSheet(typeof(Model).Name);

            var properties = typeof(Model).GetProperties();

            //Header
            var header = sheet.CreateRow(0);
            for (var i = 0; i < properties.Length; i++)
            {
                var cell = header.CreateCell(i);
                cell.SetCellValue(properties[i].Name);
            }

            //body

            IRow sheetRow = null;

            for (int i = 0; i < list.Count; i++)
            {
                sheetRow = sheet.CreateRow(i + 1);

                for (var j = 0; j < properties.Length; j++)
                {
                    ICell Row1 = sheetRow.CreateCell(j);
                    //string cellvalue = FormatRow(properties[j], lists[i]);
                    //Row1.SetCellValue(cellvalue);
                    SetCell(hssfworkbook, Row1, properties[j], list[i]);
                }
            }

            FileStream file = new FileStream(@"test.xls", FileMode.Create);
            hssfworkbook.Write(file);
            file.Close();
        }

        static void WriteExcel(List< Model> list)
        {
           
            // Lets converts our object data to Datatable for a simplified logic.
            // Datatable is most easy way to deal with complex datatypes for easy reading and formatting.

            DataTable table = (DataTable)JsonConvert.DeserializeObject(JsonConvert.SerializeObject(list), (typeof(DataTable)));
            var memoryStream = new MemoryStream();

            using (var fs = new FileStream("Result.xlsx", FileMode.Create, FileAccess.Write))
            {
                IWorkbook workbook = new XSSFWorkbook();
                ISheet excelSheet = workbook.CreateSheet("Sheet1");

                List<string> columns = new List<string>();
                IRow row = excelSheet.CreateRow(0);
                int columnIndex = 0;

                foreach (System.Data.DataColumn column in table.Columns)
                {
                    columns.Add(column.ColumnName);
                    row.CreateCell(columnIndex).SetCellValue(column.ColumnName);
                    columnIndex++;
                }

                int rowIndex = 1;
                foreach (DataRow dsrow in table.Rows)
                {
                    row = excelSheet.CreateRow(rowIndex);
                    int cellIndex = 0;
                    foreach (String col in columns)
                    {
                        row.CreateCell(cellIndex).SetCellValue(dsrow[col].ToString());
                        cellIndex++;
                    }

                    rowIndex++;
                }
                workbook.Write(fs);
            }

        }
    }
}
