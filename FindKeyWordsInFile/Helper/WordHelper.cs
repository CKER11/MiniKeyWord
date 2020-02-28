using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data;
using System.Windows;
//using MSWord = Microsoft.Office.Interop.Word;
using MSWord = NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using NetOffice.OfficeApi.Enums;
using NLog;

namespace FindKeyWordsInFile
{
    public class WordHelper : IDisposable
    {
        public static Logger logger = LogManager.GetCurrentClassLogger();
        private string docName = null; //文件名
        private bool disposed;
        MSWord.Document doc = null;
        MSWord.Application wordapp = null;
        public WordHelper(string fileName)
        {
            this.docName = fileName;
        }

        internal bool HasKeyWord(string key, out string contentPre, out string keyWord, out string contentSuf)
        {
            contentPre = "";
            keyWord = "";
            contentSuf = "";
            try
            {
                wordapp = new MSWord.Application();
                wordapp.Visible = false; //所打开的MSWord程序，是否是可见的
                wordapp.WindowState = NetOffice.WordApi.Enums.WdWindowState.wdWindowStateMinimize;
                wordapp.Width = 1;
                wordapp.Height = 1;
                wordapp.DisplayAlerts = NetOffice.WordApi.Enums.WdAlertLevel.wdAlertsNone;
                //doc = wordapp.Documents.OpenNoRepairDialog(docName,false,true,false, "", "", true, "", "", WdOpenFormat.wdOpenFormatAuto, MsoEncoding.msoEncodingAutoDetect, false);
                doc = wordapp.Documents.OpenNoRepairDialog(docName, false, true, false, "", "", true, "", "", WdOpenFormat.wdOpenFormatAuto, MsoEncoding.msoEncodingAutoDetect, false);
                //foreach (MSWord.Paragraph para in doc.Paragraphs)
                //{
                //    string text = para.Range.Text; //获得文本
                //    if (text.IndexOf(key) >= 0)
                //        return true;
                //}
                int paraCount = doc.Paragraphs.Count;//获取文档总的段落数
                MSWord.Paragraph par = doc.Paragraphs[1];//获取第一段落
                string context = par.Range.Text;//获取第一段落内容
                /////////获取多个段落的内容///////////
                MSWord.Range range = doc.Range();
                MSWord.Paragraph para1 = doc.Content.Paragraphs.FirstOrDefault();
                range.Start = para1.Range.Start;
                MSWord.Paragraph para2 = doc.Content.Paragraphs.LastOrDefault();
                range.End = para2.Range.End;
                ThreadExitis("WINWORD", true);
                return WordSubHelper.GetThreeContent(range.Text, key, out contentPre, out keyWord, out contentSuf);
            }
            catch (Exception ex)
            {
                logger.Error(docName + ":" + ex.Message);
                //Console.WriteLine("Exception: " + ex.Message);
                return false;
            }
        }

        public  static bool ThreadExitis(string threadName, bool kill)
        {
            bool bo = false;
            System.Diagnostics.Process[] processList = System.Diagnostics.Process.GetProcesses();
            foreach (System.Diagnostics.Process process in processList)
            {

                if (process.ProcessName.ToLower() == threadName.ToLower())
                {
                    if (kill)
                    {
                        bo = false;
                        process.Kill(); //结束进程 
                    }
                    else
                    {
                        bo = true;
                    }
                }
            }
            return bo;
        }
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!this.disposed)
            {
                if (disposing)
                {
                    if (wordapp != null)
                    {
                        if (doc != null)
                        {
                            MSWord._Document docc = doc as MSWord._Document;
                            docc.Close();
                        }
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(wordapp);
                    }
                    if (doc != null)
                        doc.Close();
                }

                doc = null;
                disposed = true;
            }
        }

    }
}