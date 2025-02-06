using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportsLibraryOpenXmlSdk.Services
{
    internal class FileDirectory
    {
        /// <summary>
        /// Метод проверяет существует ли директория. Если нет то создает.
        /// </summary>
        /// <param name="path">Путь к директории</param>
        /// <param name="directory">Название директории</param>
        public static void CreateDirectory(string path, string directory)
        {
            string fullPath = $@"{path}\{directory}";

            if (!Directory.Exists(fullPath))
                Directory.CreateDirectory(fullPath);
        }
    }
}
