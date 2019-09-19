using Newtonsoft.Json;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace ReadExcel
{
    public class Model
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public DateTimeOffset Now { get; set; }
        public bool New { get; set; }
    }

    class Program
    {
        /// <summary>
        /// HSSF类，只支持2007以前的excel（文件扩展名为xls），而XSSH支持07以后的
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            DataTable dtTable = new DataTable();
            List<string> rowList = new List<string>();
            ISheet sheet;
            using (var stream = new FileStream("test.xls", FileMode.Open, FileAccess.Read))
            {
                stream.Position = 0;
                IWorkbook hssWorkbook = new HSSFWorkbook(stream);

                sheet = hssWorkbook.GetSheetAt(0);
                IRow headerRow = sheet.GetRow(0);
                int cellCount = headerRow.LastCellNum;
                for (int j = 0; j < cellCount; j++)
                {
                    ICell cell = headerRow.GetCell(j);
                    if (cell == null || string.IsNullOrWhiteSpace(cell.ToString())) continue;
                    {
                        dtTable.Columns.Add(cell.ToString());
                    }
                }
                for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)
                {
                    IRow row = sheet.GetRow(i);
                    if (row == null) continue;
                    if (row.Cells.All(d => d.CellType == CellType.Blank)) continue;
                    for (int j = row.FirstCellNum; j < cellCount; j++)
                    {
                        if (row.GetCell(j) != null)
                        {
                            if (!string.IsNullOrEmpty(row.GetCell(j).ToString()) && !string.IsNullOrWhiteSpace(row.GetCell(j).ToString()))
                            {
                                rowList.Add(row.GetCell(j).ToString());
                            }
                        }
                    }
                    if (rowList.Count > 0)
                        dtTable.Rows.Add(rowList.ToArray());
                    rowList.Clear();
                }
            }
            var str = JsonConvert.SerializeObject(dtTable);
            var list = JsonConvert.DeserializeObject<List<Model>>(str);
            Console.WriteLine("Hello World!");
        }
    }
}
