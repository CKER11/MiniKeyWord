using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System.IO;
using System.Data;
using System.Windows;
using NPOI.SS.Util;
using NLog;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;

namespace FindKeyWordsInFile
{
    public class PDFHelper
    {
        public static Logger logger = LogManager.GetCurrentClassLogger();
        private string fileName = null; //文件名

        public PDFHelper(string fileName)
        {
            this.fileName = fileName;
        }

        internal bool HasKeyWord(string key, out string contentPre, out string keyWord, out string contentSuf)
        {
            contentPre = "";
            keyWord = "";
            contentSuf = "";
            try
            {
                string value = ITextSharpHandler.ReadPdf(fileName);
                return WordSubHelper.GetThreeContent(value, key, out contentPre, out keyWord, out contentSuf);
            }
            catch (Exception ex)
            {
                logger.Error(fileName + ":" + ex.Message);
                //Console.WriteLine("Exception: " + ex.Message);
                //MessageBox.Show("Exception: " + ex.Message);
                return false;
            }
        }
    }
    public static class ITextSharpHandler
    {
        public static Logger logger = LogManager.GetCurrentClassLogger();
        /// <summary>
        /// 读取PDF文本内容
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static string ReadPdf(string fileName)
        {
            if (!File.Exists(fileName))
            {
                logger.Error(@"指定的PDF文件不存在：" + fileName);
                return string.Empty;
            }
            //
            string fileContent = string.Empty;
            StringBuilder sbFileContent = new StringBuilder();
            //打开文件
            PdfReader reader = null;
            try
            {
                reader = new PdfReader(fileName);
            }
            catch (Exception ex)
            {
                logger.Error(string.Format(@"加载PDF文件{0}失败,错误:{1}", new string[] { fileName, ex.ToString() }));

                if (reader != null)
                {
                    reader.Close();
                    reader = null;
                }

                return string.Empty;
            }

            try
            {
                //循环各页（索引从1开始）
                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    sbFileContent.AppendLine(PdfTextExtractor.GetTextFromPage(reader, i));

                }

            }
            catch (Exception ex)
            {
                logger.Error(string.Format(@"解析PDF文件{0}失败,错误:{1}", new string[] { fileName, ex.ToString() }));

            }
            finally
            {
                if (reader != null)
                {
                    reader.Close();
                    reader = null;
                }
            }
            //
            fileContent = sbFileContent.ToString();
            return fileContent;
        }
        /// <summary>
        /// 获取PDF页数
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static int GetPdfPageCount(string fileName)
        {
            if (!File.Exists(fileName))
            {
                logger.Error(@"指定的PDF文件不存在：" + fileName);
                return -1;
            }
            //打开文件
            PdfReader reader = null;
            try
            {
                reader = new PdfReader(fileName);
            }
            catch (Exception ex)
            {
                logger.Error(string.Format(@"加载PDF文件{0}失败,错误:{1}", new string[] { fileName, ex.ToString() }));

                if (reader != null)
                {
                    reader.Close();
                    reader = null;
                }

                return -1;
            }
            //
            return reader.NumberOfPages;
        }
    }
}