using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace ReportsLibraryOpenXmlSdk.Services
{
    internal static class StringWork
    {
        /// <summary>
        /// Метод удаляет недопустимые символы в названии файла
        /// </summary>
        /// <param name="strInner">Строка названия файла</param>
        /// <returns>Возвращает строку с удаленными недопустимыми символами</returns>
        public static string ReplaceInvalidChar(string strInner)
        {
            string strOut = strInner;

            foreach (char c in Path.GetInvalidFileNameChars())
            {
                strOut = strOut.Replace(c.ToString(), "");
            }

            return strOut;
        }
    }
}
