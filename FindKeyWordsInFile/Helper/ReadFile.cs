using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using NLog;

namespace FindKeyWordsInFile
{
    public  class ReadFile
    {
       public static Logger logger = LogManager.GetCurrentClassLogger();
        internal static bool Read(string fileName, string key, out string contentPre, out string keyWord, out string contentSuf)
        {
            contentPre = "";
            keyWord = "";
            contentSuf = "";
            try
            {
                string ext = System.IO.Path.GetExtension(fileName).ToLower();
                switch (ext)
                {
                    case ".pdf":
                        PDFHelper PDFHelper = new PDFHelper(fileName);
                        return PDFHelper.HasKeyWord(key, out contentPre, out keyWord, out contentSuf);
                    case ".xls":
                    case ".xlsx":
                        ExcelHelper ExcelHelper = new ExcelHelper(fileName);
                        return ExcelHelper.HasKeyWord(key, out contentPre, out keyWord, out contentSuf);
                    case ".doc":
                    case ".docx":
                        WordHelper WordHelper = new WordHelper(fileName);
                        return WordHelper.HasKeyWord(key, out contentPre, out keyWord, out contentSuf);
                    case ".txt":
                    case ".svg":
                    case ".cs":
                    default:
                        string res = File.ReadAllText(fileName);
                        return WordSubHelper.GetThreeContent(res, key,out contentPre,out keyWord,out contentSuf);
                       
                }
            }
            catch (Exception ex)
            {
                logger.Error("读取错误：" + fileName + ":" + ex.Message);
                return false;
            }
        }
    }
}
