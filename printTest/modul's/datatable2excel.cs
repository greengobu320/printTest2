using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
namespace printTest.modul_s
{
    internal class datatable2excel
    {
        public datatable2excel(DataTable dataTable, string _path) {
            Excel.Application excel = new Excel.Application();
            excel.Visible = false;
            Excel.Workbook workbook = excel.Workbooks.Add();
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;
            Excel.Range headerRange = worksheet.Range["A1"].Resize[1, dataTable.Columns.Count];
            headerRange.Value = dataTable.Columns.OfType<DataColumn>().Select(c => c.ColumnName).ToArray();
            int colIndex = 0;
            int rowIndex = 0;
            object[,] myArr = new object[dataTable.Rows.Count + 1, dataTable.Columns.Count];           
            foreach (DataRow dr in dataTable.Rows)
            {
                colIndex = 0;
                Parallel.For(0, dataTable.Columns.Count, i =>
                {
                    myArr[rowIndex + 1, i] = dr[i];
                });
                rowIndex++;
            }
            worksheet.Range["A2"].Resize[myArr.GetLength(0), myArr.GetLength(1)].Value2 = myArr;
            worksheet.UsedRange.Columns.AutoFit();
            string filePath = $@"{_path}\ExportedData.xlsx";
            workbook.SaveAs(filePath);
            workbook.Close();
            excel.Quit();
        }    
    }
}
