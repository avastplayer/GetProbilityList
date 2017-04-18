using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Data;
using System.IO;

namespace GetProbilityList
{
    public class ExcelHelper
    {
        public string FilePath { get; set; } //文件名
        public IWorkbook Workbook { get; set; }
        public FileStream Fs { get; set; }
        public bool Disposed { get; set; }

        public ExcelHelper(string filePath)//构造函数
        {
            FilePath = filePath;
            Disposed = false;
        }

        public DataTable ExcelToDataTable(string sheetName, bool isFirstRowColumn)
        {
            DataTable data = new DataTable();
            try
            {
                Fs = new FileStream(FilePath, FileMode.Open, FileAccess.Read);
                if (FilePath.IndexOf(".xlsx", StringComparison.Ordinal) > 0) // 2007版本
                    Workbook = new XSSFWorkbook(Fs);
                else if (FilePath.IndexOf(".xls", StringComparison.Ordinal) > 0) // 2003版本
                    Workbook = new HSSFWorkbook(Fs);

                ISheet sheet;
                if (sheetName != null)
                {
                    sheet = Workbook.GetSheet(sheetName) ?? Workbook.GetSheetAt(0);
                }
                else
                {
                    sheet = Workbook.GetSheetAt(0);
                }
                if (sheet == null) return data;
                IRow firstRow = sheet.GetRow(0);
                int cellCount = firstRow.LastCellNum; //一行最后一个cell的编号 即总的列数

                int startRow;
                if (isFirstRowColumn)
                {
                    for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                    {
                        ICell cell = firstRow.GetCell(i);
                        string cellValue = cell?.StringCellValue;
                        if (cellValue == null) continue;
                        DataColumn column = new DataColumn(cellValue);
                        data.Columns.Add(column);
                    }
                    startRow = sheet.FirstRowNum + 1;
                }
                else
                {
                    startRow = sheet.FirstRowNum;
                }

                //最后一列的标号
                int rowCount = sheet.LastRowNum;
                for (int i = startRow; i <= rowCount; ++i)
                {
                    IRow row = sheet.GetRow(i);
                    if (row == null) continue; //没有数据的行默认是null　　　　　　　

                    DataRow dataRow = data.NewRow();
                    for (int j = row.FirstCellNum; j < cellCount; ++j)
                    {
                        if (row.GetCell(j) != null) //同理，没有数据的单元格都默认是null
                            dataRow[j] = row.GetCell(j).ToString();
                    }
                    data.Rows.Add(dataRow);
                }

                return data;
            }
            catch (Exception ex)
            {
                Console.WriteLine(@"Exception: " + ex.Message);
                return null;
            }
        }

        public int DataTableToExcel(DataTable data, string sheetName, bool isColumnWritten)
        {
            Fs = new FileStream(FilePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            if (FilePath.IndexOf(".xlsx", StringComparison.Ordinal) > 0) // 2007版本
                Workbook = new XSSFWorkbook();
            else if (FilePath.IndexOf(".xls", StringComparison.Ordinal) > 0) // 2003版本
                Workbook = new HSSFWorkbook();

            try
            {
                ISheet sheet;
                if (Workbook != null)
                {
                    sheet = Workbook.CreateSheet(sheetName);
                }
                else
                {
                    return -1;
                }

                int j;
                int count;
                if (isColumnWritten) //写入DataTable的列名
                {
                    IRow row = sheet.CreateRow(0);
                    for (j = 0; j < data.Columns.Count; ++j)
                    {
                        row.CreateCell(j).SetCellValue(data.Columns[j].ColumnName);
                    }
                    count = 1;
                }
                else
                {
                    count = 0;
                }

                int i;
                for (i = 0; i < data.Rows.Count; ++i)
                {
                    IRow row = sheet.CreateRow(count);
                    for (j = 0; j < data.Columns.Count; ++j)
                    {
                        row.CreateCell(j).SetCellValue(data.Rows[i][j].ToString());
                    }
                    ++count;
                }
                Workbook.Write(Fs); //写入到excel
                return count;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                return -1;
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (Disposed) return;
            if (disposing)
            {
                Fs?.Close();
            }

            Fs = null;
            Disposed = true;
        }
    }
}