using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;

namespace FileOperate.FileDownLoad
{
    public interface IFileDownLoad
    {
        void DownLoadFile(Page page, MemoryStream stream, string fileName);
        void DownLoadFile(Page page, string filePath);
    }
}