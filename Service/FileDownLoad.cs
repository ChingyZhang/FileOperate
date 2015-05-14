using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;

namespace FileOperate.FileDownLoad
{
    public class FileDownLoad : IFileDownLoad
    {
        #region Excel文件下载

        /// <summary>
        /// 对提供下载的附件名进行编码
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        private string FileNameEncode(Page page, string fileName)
        {
            bool isFireFox = false;
            if (page.Request.ServerVariables["http_user_agent"].ToLower().IndexOf("firefox") != -1)
            {
                isFireFox = true;
            }
            if (isFireFox == true)
            {
                //文件名前后加双引号
                fileName = "\"" + fileName + "\"";
            }
            else
            {
                //非火狐浏览器对中文文件名进行HTML编码
                fileName = HttpUtility.UrlEncode(fileName, System.Text.Encoding.UTF8);
            }
            return fileName;
        }

        /// <summary>
        /// 根据路径获取文件后缀名
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        private string FileSuffix(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || filePath.IndexOf('.') == -1) return string.Empty;
            string _suffix = filePath.Substring(filePath.LastIndexOf('.'), filePath.Length - filePath.LastIndexOf('.'));
            if (_suffix == null) _suffix = string.Empty;
            return _suffix;
        }

        /// <summary>
        /// 获取下载文件的ContentType
        /// </summary>
        /// <param name="suffix"></param>
        /// <returns></returns>
        private string ResponseCntent(string suffix)
        {
            if (string.IsNullOrEmpty(suffix)) suffix = ".*";
            string _contentType = string.Empty;
            switch (suffix)
            {
                case ".*": _contentType = "application/octet-stream"; break;

                case ".doc": _contentType = "application/msword"; break;
                case ".xlsx": _contentType = "application/ms-excel"; break;
                case ".xls": _contentType = "application/ms-excel"; break;
                case ".zip": _contentType = "application/zip"; break;

                case ".exe": _contentType = "application/x-msdownload"; break;

                case ".htm": _contentType = "text/html"; break;
                case ".html": _contentType = "text/html"; break;

                case ".bmp": _contentType = "application/x-bmp"; break;
                case ".gif": _contentType = "image/gif"; break;
                case ".ico": _contentType = "image/x-icon"; break;
                case ".png": _contentType = "image/png"; break;
                case ".jpe": _contentType = "image/jpeg"; break;
                case ".jpg": _contentType = "image/jpeg"; break;
                case ".jpeg": _contentType = "image/jpeg"; break;

                case ".js": _contentType = "application/x-javascript"; break;
                case ".css": _contentType = "text/css"; break;

                default: _contentType = "application/octet-stream"; break;
            }
            return _contentType;
        }

        /// <summary>
        /// 文件下载
        /// </summary>
        /// <param name="stream">内存流</param>
        /// <param name="fileName">下载文件名</param>
        public void DownLoadFile(Page page, MemoryStream stream, string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
            {
                fileName = DateTime.Now.ToString("yyyyMMddHHmmss");
            }
            string _suffix = FileSuffix(fileName);
            fileName = FileNameEncode(page, fileName);
            page.Response.Clear();
            page.Response.AddHeader("Content-Disposition", "attachment; filename=" + fileName);

            string _contentType = ResponseCntent(_suffix);
            page.Response.ContentType = _contentType;

            byte[] bytes = stream.ToArray();
            page.Response.BinaryWrite(bytes);//此方法额外占用二进制字节流的内存空间
            //stream.WriteTo(Response.OutputStream);//通知浏览器下载文件            
            page.Response.Flush();
            page.Response.End();//End操作后不能进行后续打印操作，但不进行End操作会导致导出的xlsx文件信息不完整（单元格导出正常）
        }

        public void DownLoadFile(Page page, string filePath)
        {
            string fileName = "";
            try
            {
                fileName = filePath.Substring(filePath.LastIndexOf("\\") + 1, filePath.Length - filePath.LastIndexOf("\\") - 1);
            }
            catch { fileName = ""; }
            if (string.IsNullOrEmpty(fileName))
            {
                fileName = DateTime.Now.ToString("yyyyMMddHHmmss");
            }

            fileName = FileNameEncode(page, fileName);
            page.Response.Clear();
            //Response.BufferOutput = true;
            page.Response.AddHeader("Content-Disposition", "attachment; filename=" + fileName);
            page.Response.ContentType = "application/ms-excel";
            if (File.Exists(filePath))
            {
                page.Response.WriteFile(filePath);//通知浏览器下载文件
            }
            else
            {
                page.Response.Write("alert('无法找到下载文件')");
            }
            page.Response.Flush();
            page.Response.End();
        }

        #endregion
    }
}