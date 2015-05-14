using NPOI.SS.Formula.Eval;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI.WebControls;

namespace FileOperate.Excel
{

    /// <summary>
    /// 利用NPOI对Excel进行导入导出操作（dll类库为NPOI 2.1.1）
    /// 导出Excel格式为2007，导入时不区分2003和2007版本
    /// </summary>
    public class ExcelOperate : IExcelOperate
    {

        public IWorkbook DtToWorkBook(DataTable dt)
        {
            IWorkbook IWorkbook = new XSSFWorkbook();
            string _strSheetName = string.IsNullOrEmpty(dt.TableName) ? "Sheet1" : dt.TableName;
            ISheet sheet = IWorkbook.CreateSheet(_strSheetName);
            DtToISheet(ref sheet, dt);
            return IWorkbook;
        }

        public IWorkbook DsToWorkBook(DataSet ds)
        {
            IWorkbook IWorkbook = new XSSFWorkbook();
            ISheet sheet;
            for (int i = 0; i < ds.Tables.Count; i++)
            {
                DataTable dt = ds.Tables[i];


                string _strSheetName = !string.IsNullOrEmpty(dt.TableName) && IWorkbook.GetSheet(dt.TableName) == null ? "Sheet" + i.ToString() : dt.TableName;
                if (string.IsNullOrEmpty(dt.TableName)) _strSheetName = "Sheet" + i.ToString();
                else if (IWorkbook.GetSheet(dt.TableName) == null) _strSheetName = dt.TableName;
                else _strSheetName = "Sheet" + i.ToString() + ":" + dt.TableName;

                sheet = IWorkbook.CreateSheet(_strSheetName);
                DtToISheet(ref sheet, dt);
            }
            return IWorkbook;
        }

        public IWorkbook GvToWorkBook(GridView gv)
        {
            IWorkbook hssfworkbook = new XSSFWorkbook();
            ISheet sheet = hssfworkbook.CreateSheet("sheet1");
            GvToISheet(ref sheet, gv);
            return hssfworkbook;
        }

        #region DataTable导出为ISheet

        public ISheet DtToISheet(ref ISheet sheet, DataTable dt)
        {
            return DtToISheet(ref sheet, dt, false);
        }

        /// <summary>
        /// 将DataTable转为NPOI工作簿
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="flagFormat">是否合并单元格（消耗资源过大，不推荐使用）</param>
        /// <returns></returns>
        public ISheet DtToISheet(ref ISheet sheet, DataTable dt, bool flagFormat)
        {
            List<int> colTypeList = GetColumnsType(dt);

            //IWorkbook hssfworkbook new XSSFWorkbook();
            //ISheet sheet = hssfworkbook.CreateSheet("sheet1");

            ICellStyle cellStyleDecimal = GetCellStyleDecimal(sheet.Workbook);
            ICellStyle cellStyleDateTime = GetCellStyleDateTime(sheet.Workbook);
            ICellStyle cellStyle = GetCellStyleCommon(sheet.Workbook);

            int groups = AddSheetHeader(sheet, dt, cellStyle);//表头行数
            //为表格创建足够多的行
            for (int i = groups; i < groups + dt.Rows.Count; i++) { sheet.CreateRow(i); }
            int maxColumnMerge = -1;// DataTableMergSimpleValueRow(sheet, dt, groups);//DataTable数据合并区列索引最大值
            if (flagFormat)
            {
                maxColumnMerge = DataTableMergSimpleValueRow(sheet, dt, groups);//DataTable数据合并区列索引最大值
            }
            AddSheetBody(sheet, dt, cellStyle, colTypeList, groups, maxColumnMerge);

            //设置列宽
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                sheet.SetColumnWidth(i, 18 * 256);
            }

            return sheet;
        }

        #region 将DataTable转为NPOI工作簿时添加表头

        /// <summary>
        /// 将DataTable转为NPOI工作簿时添加表头
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="dt"></param>
        /// <param name="headerCellStyle"></param>
        /// <returns>表头行数</returns>
        private static int AddSheetHeader(ISheet sheet, DataTable dt, ICellStyle headerCellStyle)
        {
            int groups = 0;//表头分行
            foreach (DataColumn c in dt.Columns)
            {
                int len = c.ColumnName.Split('→').Length;
                if (groups < len) groups = len;
            }
            if (groups < 1) groups = 1;

            //预先创建足够的Excel行
            for (int i = 0; i < groups; i++) { sheet.CreateRow(i); }

            /***************************************/
            /***************添加表头****************/
            /***************************************/

            ICell cell;
            string groupName = String.Empty;//每一个分组的标题
            int groupStartIndex = 0;//每一个分组的起始位置,用于添加单元格合并区域
            for (int i = 0; i < dt.Columns.Count; i++)//i代表列索引
            {
                //如果Caption没有设置，则返回 ColumnName 的值
                string headCellStr = dt.Columns[i].ColumnName;

                //当前列所在分组的标题，不存在分组时标题设为空字符
                string groupNameNow = headCellStr.Substring(0, headCellStr.IndexOf('→') == -1 ? 0 : headCellStr.IndexOf('→'));
                //开始创建单元格的行索引
                int startRowIndex = !string.IsNullOrEmpty(groupName) && groupNameNow.Equals(groupName) ? 1 : 0;
                //当前列的分组数
                int titleGroups = headCellStr.IndexOf('→') != -1 ? headCellStr.Split('→').Length : 1;
                for (int j = startRowIndex; j < titleGroups; j++)
                {
                    cell = sheet.GetRow(j).CreateCell(i);
                    //获得子标题
                    string subTitle = headCellStr.IndexOf('→') != -1 ? headCellStr.Split('→')[j].ToString() : headCellStr;
                    cell.SetCellValue(subTitle);
                    cell.CellStyle = headerCellStyle;
                }
                if (titleGroups < groups)//当前列分组数小于最大分组数，为最后一列添加合并区域
                {
                    sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(titleGroups - 1, groups - 1, i, i));
                }
                //当前分组标题与前一分组标题不同，要开始新的分组
                if (!groupName.Equals(groupNameNow))
                {
                    if (i > 0 && groupStartIndex != i - 1)//剔除第一列出现新列头的情况
                    {
                        sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 0, groupStartIndex, i - 1));
                    }
                    groupStartIndex = i;//重置合并区域开始位置索引
                    groupName = groupNameNow;
                }
                else if (string.IsNullOrEmpty(groupName))
                {
                    groupStartIndex = i;//重置合并区域开始位置索引
                    //groupName = string.Empty;
                }
                if (i == dt.Columns.Count - 1 && groupStartIndex < i)
                {
                    sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 0, groupStartIndex, i));
                }
            }

            ICellStyle _headrStyle = GetCellStyleHead(sheet.Workbook);
            for (int i = 0; i < groups; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    if (sheet.GetRow(i).GetCell(j) != null)
                    {
                        sheet.GetRow(i).GetCell(j).CellStyle = _headrStyle;
                    }
                    else { continue; }
                }
            }
            return groups;
        }

        #endregion

        #region 将DataTable转为NPOI工作簿时添加表数据
        /// <summary>
        /// 将DataTable转为NPOI工作簿时添加表数据
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="dt"></param>
        /// <param name="cellStyle"></param>
        /// <param name="colTypeList">每一列的数据类型</param>
        /// <param name="groups">表头行数（作为数据行第一行的索引）</param>
        /// <param name="maxColumnMerge">Excel所有非表头区域的合并单元格的最大列索引，若不启用则传人-1</param>
        private static void AddSheetBody(ISheet sheet, DataTable dt, ICellStyle cellStyle, List<int> colTypeList, int groups, int maxColumnMerge)
        {
            IRow row;
            ICell cell;
            ICellStyle cellStyleDecimal = GetCellStyleDecimal(sheet.Workbook);
            ICellStyle cellStyleDateTime = GetCellStyleDateTime(sheet.Workbook);

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                row = sheet.GetRow(i + groups);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    if (dt.Rows[i][j] == null || string.IsNullOrEmpty(dt.Rows[i][j].ToString()))
                    {
                        continue;
                    }
                    #region 当列索引小于最大合并列索引时判断单元格是否处于合并区域内

                    //当列索引小于最大合并列索引时判断单元格是否处于合并区域内
                    bool flagSkip = false;
                    if (j <= maxColumnMerge)
                    {
                        for (int m = 0; m < sheet.NumMergedRegions; m++)//遍历所有合并区域
                        {
                            NPOI.SS.Util.CellRangeAddress a = sheet.GetMergedRegion(m);
                            if (a.LastRow < groups)//剔除标题处的合并区域
                            {
                                continue;
                            }
                            //当前单元格是处于合并区域内且不为合并区域第一个单元格时，跳过此单元格
                            if (a.FirstRow < i + groups && a.LastRow > i + groups)
                            {
                                flagSkip = true;
                                //Debug.WriteLine("第" + i.ToString() + "行" + j.ToString() + "列被跳过");
                                break;
                            }
                        }
                    }
                    if (flagSkip) continue;
                    #endregion

                    //创建单元格
                    cell = row.CreateCell(j);
                    if (colTypeList.Count == 0 || colTypeList.Count < j || colTypeList[j] <= 0)//无法获取到该列类型
                    {
                        cell.SetCellValue(dt.Rows[i][j].ToString());
                        cell.CellStyle = cellStyle;
                    }
                    else
                    {
                        string cellText = dt.Rows[i][j].ToString();
                        try
                        {
                            switch (colTypeList[j])
                            {
                                case 1: cell.SetCellValue(int.Parse(cellText));//int类型
                                    cell.CellStyle = cellStyle;
                                    break;
                                case 2: cell.SetCellValue(double.Parse(cellText));//decimal数据类型
                                    cell.CellStyle = cellStyleDecimal;
                                    break;
                                case 3: cell.SetCellValue(DateTime.Parse(cellText));//日期类型
                                    cell.CellStyle = cellStyleDateTime;
                                    break;
                                default: cell.SetCellValue(cellText);
                                    cell.CellStyle = cellStyle;
                                    break;
                            }
                        }
                        catch
                        {
                            cell.SetCellValue("单元格导出失败");
                        }
                    }
                }
            }
        }
        #endregion

        #region 根据DataTable行组中相邻行相同值添加合并单元格区域

        /// <summary>
        /// 将DataTable各列中相邻行值相同的单元格合并显示
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="dt"></param>
        /// <param name="groupCount">NPOI表头行数</param>
        /// <returns>被合并的列的最大索引值</returns>
        private static int DataTableMergSimpleValueRow(ISheet sheet, DataTable dt, int groupCount)
        {
            if (dt.Rows.Count == 0 || dt.Columns.Count == 0) return 0;

            int maxColumnMerge = 0;//被合并的列的最大索引值
            for (int ColumnIndex = 0; ColumnIndex < dt.Columns.Count; ColumnIndex++)
            {
                #region 是否停止添加合并区域
                bool flagSkip = true;

                if (ColumnIndex == 0)//第一列始终需要合并
                {
                    flagSkip = false;
                }
                else
                {
                    for (int m = 0; m < sheet.NumMergedRegions; m++)//遍历所有合并区域
                    {
                        NPOI.SS.Util.CellRangeAddress a = sheet.GetMergedRegion(m);
                        if (a.LastRow < groupCount)//剔除标题处的合并区域
                        {
                            continue;
                        }
                        //当存在最大列包含上一列的合并区域时，当前列仍为可合并状态
                        if (a.LastColumn >= maxColumnMerge)
                        {
                            flagSkip = false;
                            break;
                        }
                    }
                }
                if (flagSkip)//当前列的前一列不包括任何合并行时，停止后续列的行合并
                {
                    return maxColumnMerge - 1 >= 0 ? maxColumnMerge - 1 : 0;
                }
                #endregion

                int rowspan = 0;
                for (int i = dt.Rows.Count - 1; i >= 0; i--)
                {
                    if (DataColumnIsSimple(sheet, dt, i, ColumnIndex, groupCount))
                    {
                        rowspan++;
                    }
                    else if (rowspan > 0)
                    {
                        sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(groupCount + i, groupCount + i + rowspan, ColumnIndex, ColumnIndex));
                        rowspan = 0;
                    }
                }
                if (rowspan > 1)//行数超过一行时第一列完全一致
                {
                    sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(groupCount, groupCount + rowspan + 1, ColumnIndex, ColumnIndex));
                }

                maxColumnMerge = ColumnIndex;

            }
            if (maxColumnMerge == dt.Columns.Count - 1)
            {
                return maxColumnMerge;
            }
            return maxColumnMerge - 1;
        }

        /// <summary>
        /// 判断当前列是否为简单列
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="row">最大行索引</param>
        /// <param name="column">列序号</param>
        /// <returns></returns>
        private static bool DataColumnIsSimple(ISheet sheet, DataTable dt, int row, int column, int groups)
        {
            if (dt.Rows.Count < row || dt.Columns.Count < column || string.IsNullOrEmpty(dt.Rows[row][column].ToString()) || row == 0) return false;

            //for (int i = column; i >= 0; i--)
            //{
            //    if (dt.Rows[row][column].ToString() != dt.Rows[row - 1][column].ToString()) return false;
            //}
            //bool flag = false;
            if (dt.Rows[row][column].ToString() == dt.Rows[row - 1][column].ToString())
            {
                int numMergedRegions = sheet.NumMergedRegions;
                if (numMergedRegions == 0 || column == 0) return true;
                //如果前一列的同一行和上一行处于相同的合并区域内
                for (int m = 0; m < numMergedRegions; m++)
                {
                    NPOI.SS.Util.CellRangeAddress a = sheet.GetMergedRegion(m);
                    int preCol = column - 1 > 0 ? column - 1 : 0;
                    if (a.FirstColumn <= preCol && a.LastColumn >= preCol && a.FirstRow <= groups + row - 1 && a.LastRow >= row + groups)
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        #endregion

        #endregion

        #region 将GridView转为ISheet
        /// <summary>
        /// 将GridView转为NPOI工作簿
        /// </summary>
        /// <param name="gv">需要处理的GridView</param>
        /// <returns></returns>
        public ISheet GvToISheet(ref ISheet sheet, GridView gv)
        {
            List<int> colTypeList = GetColumnsType((DataTable)gv.DataSource);

            //IWorkbook hssfworkbook = new XSSFWorkbook();
            //ISheet sheet = hssfworkbook.CreateSheet("sheet1");

            ICellStyle cellStyle = GetCellStyleCommon(sheet.Workbook);

            int colCount = 0;//记录GridView列数
            //rowInex记录表头的行数
            int rowIndex = AddSheetHeader(sheet, gv.HeaderRow, cellStyle, "</th></tr><tr>", out colCount);//添加表头
            AddSheetBody(sheet, gv, cellStyle, colTypeList, colCount, rowIndex);



            for (int i = 0; i < gv.Columns.Count; i++)
            {
                sheet.SetColumnWidth(i, 18 * 256);
            }

            return sheet;
        }

        /// <summary>
        /// 为Excel添加表头
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="headerRow">GridView的HeaderRow属性</param>
        /// <param name="headerCellStyle">表头格式</param>
        /// <param name="flagNewLine">转行标志</param>
        /// <param name="colCount">Excel表列数</param>
        /// <returns>Excel表格行数</returns>
        private static int AddSheetHeader(ISheet sheet, GridViewRow headerRow, ICellStyle headerCellStyle, string flagNewLine, out int colCount)
        {
            //int 
            colCount = 0;//记录GridView列数
            int rowInex = 0;//记录表头的行数

            IRow row = sheet.CreateRow(0);
            ICell cell;

            int groupCount = 0;//记录分组数
            int colIndex = 0;//记录列索引，并于结束表头遍历后记录总列数
            for (int i = 0; i < headerRow.Cells.Count; i++)
            {
                if (rowInex != groupCount)//新增了标题行时重新创建
                {
                    row = sheet.CreateRow(rowInex);
                    groupCount = rowInex;
                }

                #region 是否跳过当前单元格

                for (int m = 0; m < sheet.NumMergedRegions; m++)//遍历所有合并区域
                {
                    NPOI.SS.Util.CellRangeAddress a = sheet.GetMergedRegion(m);
                    //当前单元格是处于合并区域内
                    if (a.FirstColumn <= colIndex && a.LastColumn >= colIndex
                        && a.FirstRow <= rowInex && a.LastRow >= rowInex)
                    {
                        colIndex++;
                        m = 0;//重新遍历所有合并区域判断新单元格是否位于合并区域
                    }
                }


                #endregion

                cell = row.CreateCell(colIndex);
                cell.CellStyle = headerCellStyle;

                TableCell tablecell = headerRow.Cells[i];

                //跨列属性可能为添加了html属性colspan，也可能是由cell的ColumnSpan属性指定
                int colSpan = 0;
                int rowSpan = 0;

                #region 获取跨行跨列属性值
                //跨列
                if (!string.IsNullOrEmpty(tablecell.Attributes["colspan"]))
                {
                    colSpan = int.Parse(tablecell.Attributes["colspan"].ToString());
                    colSpan--;
                }
                if (tablecell.ColumnSpan > 1)
                {
                    colSpan = tablecell.ColumnSpan;
                    colSpan--;
                }

                //跨行
                if (!string.IsNullOrEmpty(tablecell.Attributes["rowSpan"]))
                {
                    rowSpan = int.Parse(tablecell.Attributes["rowSpan"].ToString());
                    rowSpan--;
                }
                if (tablecell.RowSpan > 1)
                {
                    rowSpan = tablecell.RowSpan;
                    rowSpan--;
                }
                #endregion

                //添加excel合并区域
                if (colSpan > 0 || rowSpan > 0)
                {
                    sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(rowInex, rowInex + rowSpan, colIndex, colIndex + colSpan));
                    colIndex += colSpan + 1;//重新设置列索引
                }
                else
                {
                    colIndex++;
                }
                string strHeader = headerRow.Cells[i].Text;

                if (strHeader.Contains(flagNewLine))//换行标记，当只存在一行标题时不存在</th></tr><tr>，此时colCount无法被赋值
                {
                    rowInex++;
                    colCount = colIndex;
                    colIndex = 0;

                    strHeader = strHeader.Substring(0, strHeader.IndexOf("</th></tr><tr>"));
                }
                cell.SetCellValue(strHeader);
            }
            if (groupCount == 0)//只有一行标题时另外为colCount赋值
            {
                colCount = colIndex;
            }
            rowInex++;//表头结束后另起一行开始记录控件数据行索引

            ICellStyle _headrStyle = GetCellStyleHead(sheet.Workbook);
            for (int i = 0; i < rowInex; i++)
            {
                for (int j = 0; j < colCount; j++)
                {
                    if (sheet.GetRow(i).GetCell(j) != null)
                    {
                        sheet.GetRow(i).GetCell(j).CellStyle = _headrStyle;
                    }
                }
            }

            return rowInex;
        }

        /// <summary>
        /// 为Excel添加表数据
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="gv"></param>
        /// <param name="colTypeList">GridView每一列的数据类型</param>
        /// <param name="colCount">GridView的总列数</param>
        /// <param name="rowInex">添加Excel数据行的起始索引号</param>
        /// <param name="cellStyle">表格基础格式</param>
        /// <returns>Excel表格行数</returns>
        private static int AddSheetBody(ISheet sheet, GridView gv, ICellStyle cellStyle, List<int> colTypeList, int colCount, int rowInex)
        {
            IRow row;
            ICell cell;
            ICellStyle cellStyleDecimal = GetCellStyleDecimal(sheet.Workbook);
            ICellStyle cellStyleDateTime = GetCellStyleDateTime(sheet.Workbook);

            int rowCount = gv.Rows.Count;

            for (int i = 0; i < rowCount; i++)
            {
                row = sheet.CreateRow(rowInex);

                for (int j = 0; j < colCount; j++)
                {
                    if (gv.Rows[i].Cells[j].Visible == false) continue;

                    string cellText = gv.Rows[i].Cells[j].Text.Trim();
                    cellText = cellText.Replace("&nbsp;", "");//替换空字符占位符
                    cellText = cellText.Replace("&gt;", ">");//替换 > 占位符

                    if (string.IsNullOrEmpty(cellText)) continue;//单元格为空跳过

                    cell = row.CreateCell(j);
                    if (colTypeList.Count == 0 || colTypeList.Count < j || colTypeList[j] <= 0)//无法获取到该列类型
                    {
                        cell.SetCellValue(cellText);
                        cell.CellStyle = cellStyle;
                    }
                    else
                    {
                        try
                        {
                            switch (colTypeList[j])
                            {
                                case 1: cell.SetCellValue(int.Parse(cellText));//int类型
                                    cell.CellStyle = cellStyle;
                                    break;
                                case 2: cell.SetCellValue(double.Parse(cellText));//decimal数据类型
                                    cell.CellStyle = cellStyleDecimal;
                                    break;
                                case 3: cell.SetCellValue(DateTime.Parse(cellText));//日期类型
                                    cell.CellStyle = cellStyleDateTime;
                                    break;
                                default: cell.SetCellValue(cellText);
                                    cell.CellStyle = cellStyle;
                                    break;
                            }
                        }
                        catch
                        {
                            cell.SetCellValue("单元格导出失败");
                        }
                    }

                    int MergeAcross = gv.Rows[i].Cells[j].ColumnSpan > 0 ? gv.Rows[i].Cells[j].ColumnSpan - 1 : 0;//跨列，即合并的列数

                    int MergeDown = gv.Rows[i].Cells[j].RowSpan > 0 ? gv.Rows[i].Cells[j].RowSpan - 1 : 0;//跨行，即合并的行数

                    if (MergeAcross > 0 || MergeDown > 0)//存在要合并的行
                    {
                        sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(rowInex, rowInex + MergeDown, j, j + MergeAcross));
                        j += MergeAcross;
                    }
                }
                rowInex++;
            }
            return rowInex;
        }

        #endregion

        #region 辅助方法

        #region 根据DataTable获取列类型
        /// <summary>
        /// 根据DataTable获取列类型
        /// </summary>
        /// <param name="gv"></param>
        /// <returns>1：Int32；2：Decimal；3：DateTime；4：String</returns>
        private static List<int> GetColumnsType(DataTable tb)
        {
            List<int> colTypeList = new List<int>();
            foreach (DataColumn col in tb.Columns)
            {
                int dataType = 0;
                if (col.DataType.FullName == "System.Int32") dataType = 1;
                else if (col.DataType.FullName == "System.Decimal") dataType = 2;
                else if (col.DataType.FullName == "System.DateTime") dataType = 3;
                else dataType = 4;
                colTypeList.Add(dataType);
            }
            return colTypeList;
        }
        #endregion

        #region 获取NPOI单元格类型
        /// <summary>
        /// 单元格居中，作为单元格类型基础方法不在外部调用
        /// </summary>
        /// <param name="cellStyle"></param>
        /// <returns></returns>
        private static ICellStyle CellStyleBasic(ICellStyle cellStyle)
        {
            cellStyle.Alignment = HorizontalAlignment.Center;
            cellStyle.VerticalAlignment = VerticalAlignment.Center;
            return cellStyle;
        }
        /// <summary>
        /// 通用单元格格式
        /// </summary>
        /// <param name="hssfworkbook"></param>
        /// <returns></returns>
        private static ICellStyle GetCellStyleCommon(IWorkbook hssfworkbook)
        {
            ICellStyle cellStyle = hssfworkbook.CreateCellStyle();
            CellStyleBasic(cellStyle);
            return cellStyle;
        }
        private static ICellStyle GetCellStyleDecimal(IWorkbook hssfworkbook)
        {
            ICellStyle cellStyleDecimal = hssfworkbook.CreateCellStyle();
            CellStyleBasic(cellStyleDecimal);
            cellStyleDecimal.DataFormat = NPOI.HSSF.UserModel.HSSFDataFormat.GetBuiltinFormat("0.000");
            return cellStyleDecimal;
        }
        private static ICellStyle GetCellStyleDateTime(IWorkbook hssfworkbook)
        {
            ICellStyle cellStyleDateTime = hssfworkbook.CreateCellStyle();
            CellStyleBasic(cellStyleDateTime);
            cellStyleDateTime.DataFormat = hssfworkbook.CreateDataFormat().GetFormat("yyyy/m/d h:mm:ss");
            return cellStyleDateTime;
        }
        private static ICellStyle GetCellStyleHead(IWorkbook hssfworkbook)
        {
            ICellStyle cellStyle = hssfworkbook.CreateCellStyle();
            CellStyleBasic(cellStyle);
            cellStyle.WrapText = true;//表头自动换行

            NPOI.SS.UserModel.IFont font = hssfworkbook.CreateFont();
            font.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;
            cellStyle.SetFont(font);//字体加粗
            return cellStyle;
        }
        #endregion

        #endregion

        #region 将IWorkBook写入本地xlsx文件或内存流
        /// <summary>
        /// 将IWorkBook写入本地xls文件
        /// </summary>
        /// <param name="workBook"></param>
        /// <param name="fileName">保存的文件路径</param>
        /// <returns>xls文件路径</returns>
        public bool IWorkBookToLocalFile(IWorkbook workBook, string filePath)
        {
            if (workBook == null) return false;
            if (string.IsNullOrEmpty(filePath) || filePath.IndexOf('.') == -1) return false;

            string _foleFolder = filePath.Substring(0, filePath.LastIndexOf("\\"));

            try
            {
                if (!Directory.Exists(_foleFolder)) Directory.CreateDirectory(_foleFolder);
            }
            catch (Exception) { return false; }

            string _suffix = filePath.Substring(filePath.LastIndexOf('.'), filePath.Length - filePath.LastIndexOf('.'));

            if (_suffix != ".xlsx") filePath += ".xlsx";
            try
            {
                FileStream fs = File.Create(filePath);
                workBook.Write(fs);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// 将IWorkBook写入内存流
        /// </summary>
        /// <param name="hssfworkbook"></param>
        /// <returns></returns>
        public MemoryStream IWorkbookToMemoeyStream(IWorkbook hssfworkbook)
        {
            MemoryStream memory = null;

            try
            {
                memory = new MemoryStream();
                hssfworkbook.Write(memory);
            }
            catch (Exception)
            {
                return null;
            }
            return memory;
        }
        #endregion

        /// <summary>
        /// 将指定sheet中的数据导出到datatable中
        /// </summary>
        /// <param name="sheet">需要导出的Sheet表</param>
        /// <param name="HeaderRowIndex">Sheet表中列头所在行号，-1表示没有列头</param>
        /// <param name="needHeader">是否需要表头</param>
        /// <param name="ImportMessage">错误信息，不为空表明导入过程中出现错误</param>
        /// <returns></returns>
        public DataTable DtFromSheet(ISheet sheet, int HeaderRowIndex, bool needHeader, out string ImportMessage)
        {
            ImportMessage = string.Empty;
            DataTable table = new DataTable();
            IRow headerRow;
            int cellCount;
            try
            {
                if (HeaderRowIndex < 0 || !needHeader)//无列名时以列索引作为列名
                {
                    headerRow = sheet.GetRow(0) as IRow;
                    cellCount = headerRow.LastCellNum;

                    for (int i = headerRow.FirstCellNum; i <= cellCount; i++)
                    {
                        DataColumn column = new DataColumn(Convert.ToString(i));
                        table.Columns.Add(column);
                    }
                }
                else
                {
                    headerRow = sheet.GetRow(HeaderRowIndex) as IRow;
                    cellCount = headerRow.LastCellNum;

                    #region 获取Table表列名（列名重复时添加标记“重复列名”，无列名取列索引作为列名）
                    for (int i = headerRow.FirstCellNum; i < cellCount; i++)
                    {
                        string _colName = string.Empty;
                        if (headerRow.GetCell(i) == null) _colName = table.Columns.IndexOf(Convert.ToString(i)) > 0 ? "重复列名" + i.ToString() : i.ToString();//有列明但列名处单元格为空时以行索引作为列名
                        else if (table.Columns.IndexOf(headerRow.GetCell(i).ToString()) > 0) _colName = "重复列名" + i.ToString();
                        else _colName = headerRow.GetCell(i).ToString();

                        DataColumn column = new DataColumn(headerRow.GetCell(i).ToString());
                        table.Columns.Add(column);
                    }
                    #endregion
                }
                int rowCount = sheet.LastRowNum;
                for (int i = (HeaderRowIndex + 1); i <= sheet.LastRowNum; i++)
                {
                    try
                    {
                        IRow row;
                        if (sheet.GetRow(i) == null) row = sheet.CreateRow(i) as IRow;
                        else row = sheet.GetRow(i) as IRow;

                        DataRow dataRow = table.NewRow();

                        #region 导入数据行
                        for (int j = row.FirstCellNum; j <= cellCount; j++)
                        {
                            try
                            {
                                if (row.GetCell(j) != null)
                                {
                                    switch (row.GetCell(j).CellType)
                                    {
                                        case CellType.String:
                                            string str = row.GetCell(j).StringCellValue.Trim();
                                            if (str != null && str.Length > 0)
                                            {
                                                dataRow[j] = str.ToString();
                                            }
                                            else
                                            {
                                                dataRow[j] = null;
                                            }
                                            break;
                                        case CellType.Numeric:
                                            if (DateUtil.IsCellDateFormatted(row.GetCell(j)))
                                            {
                                                dataRow[j] = DateTime.FromOADate(row.GetCell(j).NumericCellValue);
                                            }
                                            else
                                            {
                                                dataRow[j] = Convert.ToDouble(row.GetCell(j).NumericCellValue);
                                            }
                                            break;
                                        case CellType.Boolean:
                                            dataRow[j] = Convert.ToString(row.GetCell(j).BooleanCellValue);
                                            break;
                                        case CellType.Error:
                                            dataRow[j] = ErrorEval.GetText(row.GetCell(j).ErrorCellValue);
                                            break;
                                        case CellType.Formula:
                                            switch (row.GetCell(j).CachedFormulaResultType)
                                            {
                                                case CellType.String:
                                                    string strFORMULA = row.GetCell(j).StringCellValue.Trim();
                                                    if (strFORMULA != null && strFORMULA.Length > 0)
                                                    {
                                                        dataRow[j] = strFORMULA.ToString();
                                                    }
                                                    else
                                                    {
                                                        dataRow[j] = null;
                                                    }
                                                    break;
                                                case CellType.Numeric:
                                                    dataRow[j] = Convert.ToString(row.GetCell(j).NumericCellValue);
                                                    break;
                                                case CellType.Boolean:
                                                    dataRow[j] = Convert.ToString(row.GetCell(j).BooleanCellValue);
                                                    break;
                                                case CellType.Error:
                                                    dataRow[j] = ErrorEval.GetText(row.GetCell(j).ErrorCellValue);
                                                    break;
                                                default:
                                                    dataRow[j] = "";
                                                    break;
                                            }
                                            break;
                                        default:
                                            dataRow[j] = "";
                                            break;
                                    }
                                }
                            }
                            catch (Exception exception)
                            {
                                ImportMessage += string.Format("导入第{0}行第{1}列单元格时转换出错：{2}；/r/n", (i + 1).ToString(), (j + 1).ToString(), exception.Message);
                            }
                        }
                        #endregion
                        table.Rows.Add(dataRow);
                    }
                    catch (Exception exception)
                    {
                        ImportMessage += string.Format("导入第{0}行时出错：{1}；/r/n", (i + 1).ToString(), exception.Message);
                    }
                }
            }
            catch (Exception exception)
            {
                ImportMessage += string.Format("读取当前Sheet表时出错：{0}；/r/n", exception.Message);
            }
            return table;
        }

    }
}