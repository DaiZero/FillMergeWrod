using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace AutoFillWrodDoc
{
    public class NPOIExcelHepler : IDisposable
    {
        private readonly string _fileName; //文件名
        private IWorkbook _workbook;
        private FileStream _fs;
        private bool _disposed;

        public NPOIExcelHepler(string fileName)
        {
            _fileName = fileName;
            _disposed = false;
        }

        /// <summary>
        /// 将DataTable数据导入到excel中
        /// </summary>
        /// <param name="data">要导入的数据</param>
        /// <param name="isColumnWritten">DataTable的列名是否要导入</param>
        /// <param name="sheetName">要导入的excel的sheet的名称</param>
        /// <returns>导入数据行数(包含列名那一行)</returns>
        public int DataTableToExcel(DataTable data, string sheetName, bool isColumnWritten)
        {
            if (_fileName.IndexOf(".xlsx", StringComparison.Ordinal) > 0) // 2007版本
                _workbook = new XSSFWorkbook();
            else if (_fileName.IndexOf(".xls", StringComparison.Ordinal) > 0) // 2003版本
                _workbook = new HSSFWorkbook();

            try
            {
                ISheet sheet;
                if (_workbook != null)
                {
                    sheet = _workbook.CreateSheet(sheetName);
                }
                else
                {
                    //sheet = _workbook.CreateSheet(sheetName);
                    return -1;
                }

                int count;
                int j;
                if (isColumnWritten) //写入DataTable的列名
                {
                    IRow row = sheet.CreateRow(0);
                    for (j = 0; j < data.Columns.Count; ++j)
                    {
                        ICell cell = row.CreateCell(j);
                        cell.SetCellValue(data.Columns[j].ColumnName);

                        #region ==定义表头样式 add by Grady==

                        ICellStyle style = _workbook.CreateCellStyle();//创建样式对象
                        IFont font = _workbook.CreateFont(); //创建一个字体样式对象
                        font.FontName = "宋体"; //和excel里面的字体对应
                        //font.Color = new HSSFColor.Pink().Indexed;//颜色参考NPOI的颜色对照表(替换掉PINK())
                        font.Color = 8;//颜色参考NPOI的颜色对照表(替换掉PINK())
                        font.IsBold = true;
                        font.FontHeightInPoints = 17;//字体大小
                        font.Boldweight = short.MaxValue;//字体加粗
                        style.SetFont(font); //将字体样式赋给样式对象
                        cell.CellStyle = style; //把样式赋给单元格

                        #endregion

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

                        #region ==列自适应大小 add by Grady ==
                        sheet.AutoSizeColumn(j);//列宽度自适应大小 
                        #endregion
                    }
                    ++count;
                }


                #region MyRegion ==先将workBook写到内存流中，再转换为字节，最后保存到文件中，这样就可以兼容2003和2007版本了 Modify by Grady==

                //直接将workBook写入到excel文件中,2007版本及以上有问题，具体原因不详
                //_fs = new FileStream(_fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                //_workbook.Write(_fs); 


                //写到内存中
                MemoryStream stream = new MemoryStream();
                _workbook.Write(stream);

                //转为字节数组  
                var buf = stream.ToArray();

                //保存为Excel文件  
                using (FileStream fs = new FileStream(_fileName, FileMode.Create, FileAccess.Write))
                {
                    fs.Write(buf, 0, buf.Length);
                    fs.Flush();
                }

                #endregion

                return count;
            }
            catch
            {
                return -1;
            }
        }

        /// <summary>
        /// 将excel中的数据导入到DataTable中
        /// </summary>
        /// <param name="sheetName">excel工作薄sheet的名称</param>
        /// <param name="isFirstRowColumn">第一行是否是DataTable的列名</param>
        /// <returns>返回的DataTable</returns>
        public DataTable ExcelToDataTable(string sheetName, bool isFirstRowColumn = true)
        {
            DataTable data = new DataTable();
            try
            {
                _fs = new FileStream(_fileName, FileMode.Open, FileAccess.Read);
                if (_fileName.IndexOf(".xlsx", StringComparison.Ordinal) > 0) // 2007版本
                    _workbook = new XSSFWorkbook(_fs);
                else if (_fileName.IndexOf(".xls", StringComparison.Ordinal) > 0) // 2003版本
                    _workbook = new HSSFWorkbook(_fs);

                ISheet sheet;
                if (!string.IsNullOrWhiteSpace(sheetName))
                {
                    sheet = _workbook.GetSheet(sheetName) ?? _workbook.GetSheetAt(0);
                }
                else
                {
                    sheet = _workbook.GetSheetAt(0);
                }
                if (sheet != null)
                {
                    IRow firstRow = sheet.GetRow(0);
                    int cellCount = firstRow.LastCellNum; //一行最后一个cell的编号 即总的列数

                    int startRow;
                    if (isFirstRowColumn)
                    {
                        for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                        {
                            ICell cell = firstRow.GetCell(i);
                            string cellValue = cell?.StringCellValue;
                            if (cellValue != null)
                            {
                                DataColumn column = new DataColumn(cellValue);
                                data.Columns.Add(column);
                            }
                        }
                        startRow = sheet.FirstRowNum + 1;
                    }
                    else
                    {
                        for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                        {
                            DataColumn column = new DataColumn(i.ToString());
                            data.Columns.Add(column);
                        }
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
                            //var cellValue =Convert.ToString(row.GetCell(j));
                            var cell = row.GetCell(j);
                            #region ==对excel数据类型处理   add by Grady==
                            if (cell == null)
                            {
                                dataRow[j] = "";
                            }
                            else
                            {
                                switch (cell.CellType)
                                {
                                    case CellType.Blank:
                                        dataRow[j] = "";
                                        break;
                                    case CellType.Numeric:
                                        short format = cell.CellStyle.DataFormat;
                                        //对时间格式（2015.12.5、2015/12/5、2015-12-5等）的处理
                                        if (format == 14 || format == 31 || format == 57 || format == 58)
                                            dataRow[j] = cell.DateCellValue;
                                        else
                                            dataRow[j] = cell.NumericCellValue;
                                        break;
                                    case CellType.String:
                                        dataRow[j] = cell.StringCellValue;
                                        break;
                                }
                            }
                            #endregion
                            //if (cellValue != null) //同理，没有数据的单元格都默认是null
                            //    dataRow[j] = cellValue;
                        }
                        data.Rows.Add(dataRow);
                    }
                }
                return data;
            }
            catch
            {
                return null;
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    _fs?.Close();
                }
                _fs = null;
                _disposed = true;
            }
        }

        /// <summary>
        /// 将excel导入到datatable
        /// </summary>
        /// <param name="filePath">excel路径</param>
        /// <param name="isColumnName">第一行是否是列名</param>
        /// <param name="beginRow">读取开始行（从1开始）</param>
        /// <returns>返回datatable</returns>
        public DataTable ExcelToDataTable1(string filePath, bool isColumnName, int beginRow)
        {
            DataTable dataTable = null;
            FileStream fs = null;
            DataColumn column = null;
            DataRow dataRow = null;
            IWorkbook workbook = null;
            ISheet sheet = null;
            IRow row = null;
            ICell cell = null;
            int startRow = beginRow - 1;
            try
            {
                using (fs = File.OpenRead(filePath))
                {
                    // 2007版本
                    if (filePath.IndexOf(".xlsx") > 0)
                        workbook = new XSSFWorkbook(fs);
                    // 2003版本
                    else if (filePath.IndexOf(".xls") > 0)
                        workbook = new HSSFWorkbook(fs);

                    if (workbook != null)
                    {
                        sheet = workbook.GetSheetAt(0);//读取第一个sheet，当然也可以循环读取每个sheet
                        dataTable = new DataTable();
                        if (sheet != null)
                        {
                            int rowCount = sheet.LastRowNum;//总行数
                            if (rowCount > 0)
                            {
                                IRow firstRow = sheet.GetRow(0);//第一行
                                int cellCount = firstRow.LastCellNum;//列数

                                //构建datatable的列
                                if (isColumnName)
                                {
                                    startRow = startRow + 1;//如果第一行是列名，则从第二行开始读取
                                    for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                                    {
                                        cell = firstRow.GetCell(i);
                                        if (cell != null)
                                        {
                                            if (cell.StringCellValue != null)
                                            {

                                                column = new DataColumn(cell.StringCellValue);
                                                dataTable.Columns.Add(column);
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                                    {
                                        column = new DataColumn("column" + (i + 1));
                                        dataTable.Columns.Add(column);
                                    }
                                }

                                //填充行
                                for (int i = startRow; i <= rowCount; ++i)
                                {
                                    row = sheet.GetRow(i);
                                    if (row == null) continue;

                                    dataRow = dataTable.NewRow();
                                    for (int j = row.FirstCellNum; j < cellCount; ++j)
                                    {
                                        cell = row.GetCell(j);
                                        if (cell == null)
                                        {
                                            dataRow[j] = "";
                                        }
                                        else
                                        {
                                            //CellType(Unknown = -1,Numeric = 0,String = 1,Formula = 2,Blank = 3,Boolean = 4,Error = 5,)
                                            switch (cell.CellType)
                                            {
                                                case CellType.Blank:
                                                    dataRow[j] = "";
                                                    break;
                                                case CellType.Numeric:
                                                    short format = cell.CellStyle.DataFormat;
                                                    //对时间格式（2015.12.5、2015/12/5、2015-12-5等）的处理
                                                    if (format == 14 || format == 31 || format == 57 || format == 58)
                                                        dataRow[j] = cell.DateCellValue;
                                                    else
                                                        dataRow[j] = cell.NumericCellValue;
                                                    break;
                                                case CellType.String:
                                                    dataRow[j] = cell.StringCellValue;
                                                    break;
                                            }
                                        }
                                    }
                                    dataTable.Rows.Add(dataRow);
                                }
                            }
                        }
                    }
                }
                return dataTable;
            }
            catch (Exception)
            {
                if (fs != null)
                {
                    fs.Close();
                }
                return null;
            }
        }

        /// <summary>
        /// 用于将excel表格中列索引转成列号字母，从A对应1开始
        /// </summary>
        /// <param name="index">列索引</param>
        /// <returns>列号</returns>
        private string IndexToColumn(int index)
        {
            if (index <= 0)
            {
                throw new Exception("Invalid parameter");
            }
            index--;
            string column = string.Empty;
            do
            {
                if (column.Length > 0)
                {
                    index--;
                }
                column = ((char)(index % 26 + (int)'A')).ToString() + column;
                index = (int)((index - index % 26) / 26);
            } while (index > 0);
            return column;
        }

        /// <summary>
        /// excel列头导入到datatable
        /// </summary>
        /// <param name="filePath">excel路径</param>
        /// <param name="columNameIndex">第几行是列名</param>
        /// <param name="startColumnIndex">第几列开始</param>
        /// <returns>返回datatable</returns>
        public List<string> ExcelColumnToDataTable(string filePath, int columNameIndex, int startColumnIndex)
        {
            List<string> columNameList = null;
            FileStream fs = null;
            IWorkbook workbook = null;
            ISheet sheet = null;
            ICell cell = null;
            try
            {
                using (fs = File.OpenRead(filePath))
                {
                    // 2007版本
                    if (filePath.IndexOf(".xlsx") > 0)
                        workbook = new XSSFWorkbook(fs);
                    // 2003版本
                    else if (filePath.IndexOf(".xls") > 0)
                        workbook = new HSSFWorkbook(fs);

                    if (workbook != null)
                    {
                        sheet = workbook.GetSheetAt(0);//读取第一个sheet，当然也可以循环读取每个sheet
                        columNameList = new List<string>();
                        if (sheet != null)
                        {
                            int rowCount = sheet.LastRowNum;//总行数
                            if (columNameIndex > rowCount)
                            {
                                return null;
                            }

                            IRow Row = sheet.GetRow(columNameIndex - 1);//第一行
                            int cellCount = Row.LastCellNum;//列数
                            for (int i = startColumnIndex - 1; i < cellCount; ++i)
                            {
                                cell = Row.GetCell(i);
                                if (cell != null)
                                {
                                    if (cell.StringCellValue != null)
                                    {
                                        string item = cell.StringCellValue.Replace(" ", "");
                                        columNameList.Add(item);
                                    }
                                }
                            }
                        }
                    }
                }
                return columNameList;
            }
            catch (Exception)
            {
                if (fs != null)
                {
                    fs.Close();
                }
                return null;
            }

        }

        /// <summary>
        /// 获取Excel标签中的行号和列号
        /// </summary>
        /// <param name="filePath">Excel文件路径</param>
        /// <param name="tag">String类型的标签</param>
        /// <returns>成功获取到标签，则返回标签行号和列号；反之，则返回null</returns>
        public Tuple<int, int> GetTagIndex(string filePath, string tag)
        {
            if (string.IsNullOrWhiteSpace(filePath) || string.IsNullOrWhiteSpace(tag))
            {
                return null;
            }

            IWorkbook workbook = null;
            ISheet sheet = null;
            FileStream fs = null;
            try
            {
                using (fs = File.OpenRead(filePath))
                {
                    // 2007版本
                    if (filePath.IndexOf(".xlsx") > 0)
                        workbook = new XSSFWorkbook(fs);
                    // 2003版本
                    else if (filePath.IndexOf(".xls") > 0)
                        workbook = new HSSFWorkbook(fs);

                    if (workbook != null)
                    {
                        sheet = workbook.GetSheetAt(0);
                        if (sheet != null)
                        {
                            int rowCount = sheet.LastRowNum;
                            int columnCount = sheet.LastRowNum;
                            for (int i = 0; i < rowCount; i++)
                            {
                                for (int j = 0; j < columnCount; j++)
                                {
                                    //判断是否为string类型
                                    if (sheet.GetRow(i).GetCell(j) != null)
                                    {
                                        switch (sheet.GetRow(i).GetCell(j).CellType)
                                        {
                                            case CellType.String:
                                                var str = sheet.GetRow(i).GetCell(j).StringCellValue;
                                                if (tag.Equals(str.Replace(" ", "")))
                                                {
                                                    return new Tuple<int, int>(i, j);
                                                }
                                                break;
                                            default:
                                                break;
                                        }
                                    }


                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (fs != null)
                {
                    fs.Close();
                }
                return null;
            }
            return null;
        }
    }
}
