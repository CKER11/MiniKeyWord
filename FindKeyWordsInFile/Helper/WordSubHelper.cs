using System;

namespace FindKeyWordsInFile
{
    public class WordSubHelper
    {
        //截取长度
        private static short SubWordLen = 20;
        public static bool GetThreeContent(string res, string key, out string contentPre, out string keyWord, out string contentSuf)
        {
            contentPre = "";
            keyWord = "";
            contentSuf = "";
            int index = res.ToLower().IndexOf(key);
            int keyLen = key.Length;
           
            if (index >= 0)
            {
                //如果关键词长度超过了截取长度SubWordLen
                if (keyLen >= SubWordLen)
                {
                    int startIndex = index;
                    if (startIndex > 0)
                    {
                        contentPre = "...";
                    }
                    keyWord = res.Substring(startIndex, keyLen);
                    if (startIndex +  keyLen < res.Length )
                    {
                        contentSuf = "...";
                    }
                }
                else
                {
                    int contentPreLen = (SubWordLen - keyLen) / 2;
                    int contentSufLen = (SubWordLen - keyLen) - contentPreLen;
                    int startIndex = index - contentPreLen;
                    //前面的文字不足
                    if (startIndex < 0)
                    {
                        startIndex = 0;
                        contentPreLen = index;
                    }
                    int endIndex = index + keyLen + contentSufLen;
                    //后面的文字不足
                    if (endIndex > res.Length)
                    {
                        endIndex = res.Length;
                        contentSufLen = res.Length - index - keyLen;
                    }
                    if (contentPreLen > 0)
                    {
                        contentPre = (startIndex == 0 ? "" : "...") + res.Substring(startIndex, index - startIndex);
                    }
                    keyWord = res.Substring(startIndex + contentPreLen, keyLen);
                    if (contentSufLen > 0)
                    {
                        contentSuf = res.Substring(index + keyLen, contentSufLen) + ((index + keyLen + contentSufLen) == res.Length ? "": "...");
                    }
                }
                return true;
            }
            return false;
        }
    }
}