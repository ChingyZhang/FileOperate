using NPOI.SS.UserModel;
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
    /// Excel操作类接口
    /// </summary>
    public interface IExcelOperate
    {
        IWorkbook DtToWorkBook(DataTable dt);
        IWorkbook DsToWorkBook(DataSet ds);
        ISheet DtToISheet(ref ISheet sheet, DataTable dt);
        /// <summary>
        /// 将DataTable转为NPOI工作簿
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="flagFormat">是否合并单元格（消耗资源过大，不推荐使用）</param>
        /// <returns></returns>
        ISheet DtToISheet(ref ISheet sheet, DataTable dt, bool flagFormat);

        IWorkbook GvToWorkBook(GridView gv);
        ISheet GvToISheet(ref ISheet sheet, GridView gv);

        DataTable DtFromSheet(ISheet sheet, int HeaderRowIndex, bool needHeader, out string ImportMessage);

        /// <summary>
        /// 将IWorkBook写入本地xls文件，文件名后缀需为xls
        /// </summary>
        /// <param name="workBook"></param>
        /// <param name="filePath"></param>
        /// <returns></returns>
        bool IWorkBookToLocalFile(IWorkbook workBook, string filePath);

        MemoryStream IWorkbookToMemoeyStream(IWorkbook hssfworkbook);


    }

}