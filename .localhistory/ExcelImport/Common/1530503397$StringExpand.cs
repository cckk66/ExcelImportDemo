using ExcelImport.Enum;
using Microsoft.International.Converters.PinYinConverter;
using System;
using System.Collections;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelImportDemo
{
    /// <summary>
    /// 扩展类
    /// </summary>
    public static class StringExpand
    {
        #region 全角半角转换
        /// <summary>
        /// 转换全角
        /// </summary>
        /// <param name="Value">待转换文字</param>
        /// <returns>转换后文字</returns>
        public static string ToFullWidth(this string Value)
        {
            char[] c = Value.ToCharArray();
            for (int i = 0; i < c.Length; i++)
            {
                if (c[i] == 32)
                {
                    c[i] = (char)12288;
                    continue;
                }
                if (c[i] < 127)
                {
                    c[i] = (char)(c[i] + 65248);
                }
            }
            return new string(c);
        }
        /// <summary>
        /// 转换半角
        /// </summary>
        /// <param name="Value">待转换字符串</param>
        /// <returns>转换后字符串</returns>
        public static string ToHalfWidth(this string Value)
        {
            char[] c = Value.ToCharArray();
            for (int i = 0; i < c.Length; i++)
            {
                if (c[i] == 12288)
                {
                    c[i] = (char)32;
                    continue;
                }
                if (c[i] > 65280 && c[i] < 65375)
                    c[i] = (char)(c[i] - 65248);
            }
            return new string(c);
        }
        #endregion

        /// <summary>
        /// 去除字符串中的数字
        /// </summary>
        /// <param name="key">字符串</param>
        /// <returns>非数字</returns>
        public static string RemoveNumber(this string key)
        {
            if (string.IsNullOrWhiteSpace(key)) return string.Empty;
            return Regex.Replace(key, @"\d", "");
        }
        /// <summary>
        /// 去除字符串中的非数字
        /// </summary>
        /// <param name="key">字符串</param>
        /// <returns>数字</returns>
        public static string RemoveNotNumber(this string key)
        {
            if (string.IsNullOrWhiteSpace(key)) return string.Empty;
            return Regex.Replace(key, @"[^\d]*", "");
        }

        /// <summary>
        /// 去除字符串中的非数字，保留逗号
        /// </summary>
        /// <param name="key">字符串</param>
        /// <returns>带逗号分隔的数字</returns>
        public static string RemoveNotNumberSplit(this string key)
        {
            if (string.IsNullOrWhiteSpace(key))
                return string.Empty;
            key = key.Trim();
            return Regex.Replace(key, @"[^\d\,]*", "").Trim(',');
        }
        /// <summary>
        /// 字符串数组转换为整数数组
        /// 默认或失败返回空整型数组,不会null
        /// </summary>
        /// <param name="arrs">待转换的string[]</param>
        /// <param name="distinct">是否去重 bool</param>
        /// <param name="sort">排序方式Enum_Sort,默认 未知=不排序</param>
        /// <param name="RemoveZero">是否移除0 bool</param>
        /// <returns>转换后整型数组</returns>
        public static int[] StringArrayToIntArray(this string[] arrs, bool distinct = false, Enum_Sort sort= Enum_Sort.未知,bool RemoveZero=false) {

            if (arrs == null) return new int[] { };
            try
            {
                int[] tempintarray = Array.ConvertAll(arrs, new Converter<string, int>(ToInt));

                if (tempintarray.Length < 1) return tempintarray;

                if (distinct)
                    tempintarray = tempintarray.Distinct().ToArray();
                if(RemoveZero)
                    tempintarray=tempintarray.Where(a => a!=0).ToArray();

                switch (sort) {
                    case Enum_Sort.升序: {
                            return tempintarray.OrderBy(a => a).ToArray();
                        }
                    case Enum_Sort.降序: {
                            return tempintarray.OrderByDescending(a => a).ToArray();
                        }
                    default: {
                            return tempintarray;
                        }
                }
            }
            catch { return new int[] { }; }
        }
        
        #region 正则表达式扩展方法
        /// <summary>
        /// 是否找到了匹配项
        /// </summary>
        /// <param name="s">待检测字符串</param>
        /// <param name="pattern">正则</param>
        /// <returns>是否匹配bool</returns>
        public static bool IsMatch(this string s, string pattern)
        {
            if (s == null) return false;
            else return Regex.IsMatch(s, pattern);
        }
        /// <summary>
        /// 搜索第一个匹配项
        /// </summary>
        /// <param name="s">待检测字符串</param>
        /// <param name="pattern">正则</param>
        /// <returns>返回符合条件的第一项</returns>
        public static string Match(this string s, string pattern)
        {
            if (s == null) return "";
            return Regex.Match(s, pattern).Value;
        }
        /// <summary>
        /// 替换
        /// </summary>
        /// <param name="s">待检测字符串</param>
        /// <param name="pattern">正则</param>
        /// <param name="Value">目标字符串</param>
        /// <returns>替换后字符串</returns>
        public static string regReplace(this string s, string pattern, string Value)
        {
            Regex reg = new Regex(pattern);
            return reg.Replace(s, Value);
        }

        #endregion

        #region 基础类型转换扩展
        /// <summary>
        /// 转换为Int
        /// 失败返回默认值
        /// </summary>
        /// <param name="s">待转换对象</param>
        /// <param name="DefaultValue">默认值</param>
        /// <returns>转换结果</returns>
        public static int ToInt(this object s,int DefaultValue=0) {
            try { return Convert.ToInt32(s); } catch { return DefaultValue; }
        }
        /// <summary>
        /// 转换为Int
        /// 失败返回默认值
        /// </summary>
        /// <param name="s">待转换对象</param>
        /// <param name="DefaultValue">默认值</param>
        /// <returns>转换结果</returns>
        public static int ToIntNoZero(this object s, int DefaultValue = 0)
        {
            try {
                int a= Convert.ToInt32(s);
                if (a == 0) return DefaultValue;
                return a; } catch { return DefaultValue; }
        }
        /// <summary>
        /// 转换为Int
        /// 失败返回0
        /// </summary>
        /// <param name="s">待转换对象</param>
        /// <returns>转换结果</returns>
        public static int ToInt(string s)
        {
            try { return Convert.ToInt32(s); } catch { return 0; }
        }
        /// <summary>
        /// 转换为Double
        /// 失败返回默认值
        /// </summary>
        /// <param name="s">待转换对象</param>
        /// <param name="DefaultValue">默认值</param>
        /// <returns>转换结果</returns>
        public static double ToDouble(this object s, double DefaultValue = 0)
        {
            try { return Convert.ToDouble(s); } catch { return DefaultValue; }
        }
        /// <summary>
        /// 转换为decimal
        /// 失败返回默认值
        /// </summary>
        /// <param name="s">待转换对象</param>
        /// <param name="DefaultValue">默认值</param>
        /// <returns>转换结果</returns>
        public static decimal ToDecimal(this object s, decimal DefaultValue = 0)
        {
            try { return Convert.ToDecimal(s); } catch { return DefaultValue; }
        }
        /// <summary>
        /// 转换为Float
        /// 失败返回默认值
        /// </summary>
        /// <param name="s">待转换对象</param>
        /// <param name="DefaultValue">默认值</param>
        /// <returns>转换结果</returns>
        public static float ToFloat(this object s, float DefaultValue = 0)
        {
            try { return (float)s; } catch { return DefaultValue; }
        }
        /// <summary>
        /// 转换为Long
        /// 失败返回默认值
        /// </summary>
        /// <param name="s">待转换对象</param>
        /// <param name="DefaultValue">默认值</param>
        /// <returns>转换结果</returns>
        public static long ToLong(this object s, long DefaultValue = 0)
        {
            try { return Convert.ToInt64(s); } catch { return DefaultValue; }
        }
        /// <summary>
        /// 转换为String
        /// 失败返回默认值
        /// </summary>
        /// <param name="s">待转换对象</param>
        /// <param name="DefaultValue">默认值</param>
        /// <returns>转换结果</returns>
        public static string ToStr(this object s, string DefaultValue="")
        {
            try { return Convert.ToString(s); } catch { return DefaultValue; }
        }
        /// <summary>
        /// 转换为GUID
        /// 失败返回NULL
        /// </summary>
        /// <param name="s">待转换对象</param>
        /// <param name="errorInfo">转换错误描述</param>
        /// <returns>转换结果，错误时返回ApiException</returns>
        public static Guid ToGuid(this string s,string errorInfo)
        {
            try { return Guid.Parse(s); } catch { throw new Exceptions(errorInfo); }
        }
        /// <summary>
        /// 转换为String
        /// 失败返回默认值
        /// 将检测是否无效字符串（为空或Empty或""）无效会返回默认值
        /// </summary>
        /// <param name="s">待转换对象</param>
        /// <param name="DefaultValue">默认值</param>
        /// <returns>转换结果</returns>
        public static string ToStrNoEmpty(this object s, string DefaultValue = "")
        {
            try {
                string temp= Convert.ToString(s);
                if (string.IsNullOrWhiteSpace(temp)) return DefaultValue;
                return temp;
            } catch { return DefaultValue; }
        }
        /// <summary>
        /// 判断字符串是否有效，无效则返回DBNull，有效返回本体
        /// </summary>
        /// <param name="s">待检测对象</param>
        /// <returns>返回结果</returns>
        public static object DBNULL(this object s)
        {

            if (string.IsNullOrWhiteSpace(s.ToStr()))
                return DBNull.Value;
            return s;
        }
        #endregion

        

        #region 中文转换
        /// <summary>
        /// 字符串转简拼(汉字首字母) 大写
        /// </summary>
        /// <param name="s">需要转换的汉字语句</param>
        /// <param name="DefaultValue">默认值</param>
        /// <returns>转换后简拼(汉字首字母) 大写</returns>
        public static string ToPinYin(this string s, string DefaultValue = "",bool JustFirst=false)
        {
            if (string.IsNullOrWhiteSpace(s))
                return DefaultValue;
            try
            {
                string resultStr = "";
                string resultFirst = "";
                char[] numChars = { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' };
                foreach (char item in s.ToCharArray())
                {
                    if (ChineseChar.IsValidChar(item))
                    {
                        ChineseChar cc = new ChineseChar(item);
                        if (!string.IsNullOrWhiteSpace(cc.Pinyins[0]) && cc.Pinyins[0].Length > 0)
                        {
                            resultStr += cc.Pinyins[0].TrimEnd(numChars);
                            resultFirst += cc.Pinyins[0].Substring(0, 1);//拼音首字母
                        }

                    }
                    else
                    {
                        resultStr += item.ToString();
                    }
                }
                if (JustFirst)
                    return (string.IsNullOrWhiteSpace(resultFirst)? DefaultValue: resultFirst).ToUpper();
                else
                    return string.Format("{0} {1}", resultFirst, resultStr).ToUpper();


            }
            catch { return DefaultValue; }

        }
        #endregion

        #region 加密处理
        /// <summary>
        /// Base64加密
        /// </summary>
        /// <param name="encryptString">要加密的字符串</param>
        /// <returns>返回加密后字符串,异常时返回""</returns>
        public static string EncodeBase64(this string encryptString)
        {

            try
            {
                char[] Base64Code = new char[]{'A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T',
                                            'U','V','W','X','Y','Z','a','b','c','d','e','f','g','h','i','j','k','l','m','n',
                                            'o','p','q','r','s','t','u','v','w','x','y','z','0','1','2','3','4','5','6','7',
                                            '8','9','+','/','='};
                byte empty = (byte)0;
                ArrayList byteMessage = new ArrayList(Encoding.Default.GetBytes(encryptString));
                StringBuilder outmessage;
                int messageLen = byteMessage.Count;
                int page = messageLen / 3;
                int use = 0;
                if ((use = messageLen % 3) > 0)
                {
                    for (int i = 0; i < 3 - use; i++)
                        byteMessage.Add(empty);
                    page++;
                }
                outmessage = new System.Text.StringBuilder(page * 4);
                for (int i = 0; i < page; i++)
                {
                    byte[] instr = new byte[3];
                    instr[0] = (byte)byteMessage[i * 3];
                    instr[1] = (byte)byteMessage[i * 3 + 1];
                    instr[2] = (byte)byteMessage[i * 3 + 2];
                    int[] outstr = new int[4];
                    outstr[0] = instr[0] >> 2;
                    outstr[1] = ((instr[0] & 0x03) << 4) ^ (instr[1] >> 4);
                    if (!instr[1].Equals(empty))
                        outstr[2] = ((instr[1] & 0x0f) << 2) ^ (instr[2] >> 6);
                    else
                        outstr[2] = 64;
                    if (!instr[2].Equals(empty))
                        outstr[3] = (instr[2] & 0x3f);
                    else
                        outstr[3] = 64;
                    outmessage.Append(Base64Code[outstr[0]]);
                    outmessage.Append(Base64Code[outstr[1]]);
                    outmessage.Append(Base64Code[outstr[2]]);
                    outmessage.Append(Base64Code[outstr[3]]);
                }
                return outmessage.ToString();
            }
            catch //(Exception ex)
            {
                return "";
                //throw ex;
            }
        }
        /// <summary>
        /// Base64解密
        /// </summary>
        /// <param name="decryptString">要解密的字符串</param>
        /// <returns>返回解密后字符串,异常时返回异常原因</returns>
        public static string DecodeBase64(this string decryptString)
        {
            //将空格替换为加号
            decryptString = decryptString.Replace(" ", "+");

            try
            {
                if ((decryptString.Length % 4) != 0)
                {
                    return "包含不正确的BASE64编码";
                }
                if (!Regex.IsMatch(decryptString, "^[A-Z0-9/+=]*$", RegexOptions.IgnoreCase))
                {
                    return "包含不正确的BASE64编码";
                }
                string Base64Code = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=";
                int page = decryptString.Length / 4;
                ArrayList outMessage = new ArrayList(page * 3);
                char[] message = decryptString.ToCharArray();
                for (int i = 0; i < page; i++)
                {
                    byte[] instr = new byte[4];
                    instr[0] = (byte)Base64Code.IndexOf(message[i * 4]);
                    instr[1] = (byte)Base64Code.IndexOf(message[i * 4 + 1]);
                    instr[2] = (byte)Base64Code.IndexOf(message[i * 4 + 2]);
                    instr[3] = (byte)Base64Code.IndexOf(message[i * 4 + 3]);
                    byte[] outstr = new byte[3];
                    outstr[0] = (byte)((instr[0] << 2) ^ ((instr[1] & 0x30) >> 4));
                    if (instr[2] != 64)
                    {
                        outstr[1] = (byte)((instr[1] << 4) ^ ((instr[2] & 0x3c) >> 2));
                    }
                    else
                    {
                        outstr[2] = 0;
                    }
                    if (instr[3] != 64)
                    {
                        outstr[2] = (byte)((instr[2] << 6) ^ instr[3]);
                    }
                    else
                    {
                        outstr[2] = 0;
                    }
                    outMessage.Add(outstr[0]);
                    if (outstr[1] != 0)
                        outMessage.Add(outstr[1]);
                    if (outstr[2] != 0)
                        outMessage.Add(outstr[2]);
                }
                byte[] outbyte = (byte[])outMessage.ToArray(Type.GetType("System.Byte"));
                return Encoding.Default.GetString(outbyte);
            }
            catch //(Exception ex)
            {
                return "转换异常";
            }
        }

        /// <summary>
        /// 字符串转MD5编码
        /// </summary>
        /// <param name="value">待转换字符串</param>
        /// <returns>MD5值</returns>
        public static string ToMD5Encode(this string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return string.Empty;
            StringBuilder sb = new StringBuilder();
            byte[] bytevalue = Encoding.Default.GetBytes(value);

            MD5 md5 = MD5.Create();
            byte[] byteHash = md5.ComputeHash(bytevalue);

            foreach (var item in byteHash)
            {
                sb.Append(item.ToString("x2"));
            }

            return sb.ToString();
            //byte[] b = Encoding.Default.GetBytes(value);
            //b = new MD5CryptoServiceProvider().ComputeHash(b);
            //string ret = "";
            //for (int i = 0; i < b.Length; i++)
            //    ret += b[i].ToString("x").PadLeft(2, '0');
            //return ret;
        }
        #endregion

        #region 辖区编码转换
        /// <summary>
        /// 检测辖区编码有效性
        /// </summary>
        /// <param name="CodeId">待检测编码</param>
        /// <returns>是否有效</returns>
        public static bool CheckCode(this string CodeId) {
            //必须存在
            if (string.IsNullOrWhiteSpace(CodeId)) return false;
            //无非数字
            if (CodeId.RemoveNotNumber().Length != CodeId.Length) return false;
            CodeId = CodeId.TrimEndCode();
            if (CodeId.Length < 1) return false;
            //不超过6位时，编码是2的倍数，否则编码是3的倍数
            if (CodeId.Length <= 6) 
                return CodeId.Length % 2 == 0;
            else
                return CodeId.Length % 3 == 0;
        }
        /// <summary>
        /// 移除辖区编码后面的0，不限制层级
        /// </summary>
        /// <param name="CodeId">需要操作的辖区编码</param>
        /// <returns>错误时返回string.empty</returns>
        public static string TrimEndCode(this string CodeId)
        {
            if (string.IsNullOrWhiteSpace(CodeId))
                return string.Empty;
            CodeId = CodeId.TrimEnd('0');
            if (CodeId.Length <= 6)
            {
                if (CodeId.Length % 2 != 0)
                {
                    int cutlen = CodeId.Length % 2;
                    return CodeId.PadRight(CodeId.Length + 2 - cutlen, '0');
                }
                return CodeId;
            }

            if (CodeId.Length % 3 != 0)
            {
                int cutlen = CodeId.Length % 3;
                return CodeId.PadRight(CodeId.Length + 3 - cutlen, '0');
            }


            return CodeId;
        }
        /// <summary>
        /// 获取辖区层级（不限制辖区层级）
        /// 最低层级1=省级
        /// 0表示未知层级
        /// </summary>
        /// <param name="CodeID">辖区编码</param>
        /// <returns>不存在时返回0</returns>
        public static int GetCodeLev(this string CodeID)
        {
            CodeID = TrimEndCode(CodeID);
            if (string.IsNullOrWhiteSpace(CodeID))
                return 0;
            if (CodeID.Length <= 6)
            {
                return CodeID.Length / 2;
            }
            else
            {
                return (CodeID.Length - 6) / 3 + 3;
            }
        }
        /// <summary>
        /// 获取上级辖区
        /// 不限制层级
        /// </summary>
        /// <param name="CodeID">辖区编码</param>
        /// <returns>无效时返回string.Empty</returns>
        public static string GetCodeFather(this string CodeID)
        {
            CodeID = TrimEndCode(CodeID);
            if (string.IsNullOrWhiteSpace(CodeID) || CodeID.Length < 2)
                return string.Empty;

            if (CodeID.Length <= 6)
                return CodeID.Substring(0, CodeID.Length - 2).PadRight(15, '0');
            else
                return CodeID.Substring(0, CodeID.Length - 3).PadRight(15, '0');
        }
        /// <summary>
        /// 获取辖区当前层级的顺序码
        /// 不限制层级
        /// 0表示未知顺序码
        /// </summary>
        /// <param name="CodeID">辖区编码</param>
        /// <returns>顺序码，失败时返回0</returns>
        public static int GetASCId(this string CodeID)
        {
            CodeID = TrimEndCode(CodeID);
            if (string.IsNullOrWhiteSpace(CodeID))
                return 0;

            if (CodeID.Length <= 6)
                return CodeID.Substring(CodeID.Length - 2, 2).ToInt();
            else
                return CodeID.Substring(CodeID.Length - 3, 3).ToInt();
        }
        /// <summary>
        /// 获取辖区完整路径（从区级开始)
        /// 异常返回string.Empty
        /// </summary>
        /// <param name="CodeID">辖区编码</param>
        /// <returns>失败时返回string.Empty</returns>
        public static string GetPath(this string CodeID)
        {
            CodeID = TrimEndCode(CodeID);
            if (string.IsNullOrWhiteSpace(CodeID))
                return string.Empty;

            return GetPathChild(CodeID);
        }
        /// <summary>
        /// 获取辖区路径辅助递归方法
        /// 异常返回string.Empty
        /// </summary>
        /// <param name="CodeID">辖区编码</param>
        /// <returns>失败时返回string.Empty</returns>
        private static string GetPathChild(string CodeID)
        {
            if (string.IsNullOrWhiteSpace(CodeID) || CodeID.Length < 6)
                return string.Empty;
            if (CodeID.Length == 6)
            {
                return CodeID.PadRight(15, '0');
            }
            else
            {
                return string.Format("{0},{1}", GetPathChild(CodeID.Substring(0, CodeID.Length - 3)), CodeID.PadRight(15, '0'));
            }
        }
        #endregion

        #region APPSetting
        /// <summary>
        /// 获取APPSetting配置项值
        /// </summary>
        /// <param name="name">配置键</param>
        /// <param name="defaultValue">异常时的默认值</param>
        /// <returns>配置项值</returns>
        public static string GetAppSetting(string name, string defaultValue = "")
        {
            try
            {
                return System.Configuration.ConfigurationManager.AppSettings[name];
            }
            catch { return defaultValue; }
        }
        #endregion
        #region 验证码
        /// <summary>
        /// 生成指定长度验证编码
        /// </summary>
        /// <returns>返回生成的验证编码</returns>
        public static string GenerateCode(int codelength=4)
        {
            //定义验证码长度
            if (codelength < 1) return "";
            int number;
            string RandomCode = string.Empty;
            Random r = new Random();
            for (int i = 0; i < codelength; i++)
            {
                number = r.Next();
                number = number % 36;  //字符从0-9,A-Z中随机产生，对应的ASCLL码分别为48-57,65-90
                if (number < 10)
                    number += 48;
                else
                    number += 55;

                if (number == 48 || number == 79 || number == 49 || number == 73)//过滤数字0和字母"O",过滤数字1和字母"i"
                    number = number + 2;

                RandomCode += ((char)number).ToString();
            }
            return RandomCode;
        }
        #endregion

        #region HTML
        ///   <summary>
        ///   去除HTML标记
        ///   </summary>
        ///   <param name="Htmlstring">包括HTML的源码</param>
        ///   <returns>已经去除后的文字</returns>
        public static string NoHTML(this string Htmlstring)
        {
            if (string.IsNullOrWhiteSpace(Htmlstring)) return Htmlstring;
            //删除脚本
            Htmlstring = Regex.Replace(Htmlstring, @"<script[^>]*?>.*?</script>", "",RegexOptions.IgnoreCase);
            //删除HTML
            Htmlstring = Regex.Replace(Htmlstring, @"<(.[^>]*)>", "",RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"([\r\n])[\s]+", "",RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"-->", "", RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"<!--.*", "", RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&(quot|#34);", "\"",RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&(amp|#38);", "&",RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&(lt|#60);", "<",RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&(gt|#62);", ">",RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&(nbsp|#160);", "   ",RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&(iexcl|#161);", "\xa1",RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&(cent|#162);", "\xa2",RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&(pound|#163);", "\xa3",RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&(copy|#169);", "\xa9",RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&#(\d+);", "",RegexOptions.IgnoreCase);

            //Htmlstring.Replace("<", "");
            //Htmlstring.Replace(">", "");
            //Htmlstring.Replace("\r\n", "");
            //Htmlstring = HttpContext.Current.Server.HtmlEncode(Htmlstring).Trim();

            return Htmlstring.Trim();
        }
        #endregion
    }
}
