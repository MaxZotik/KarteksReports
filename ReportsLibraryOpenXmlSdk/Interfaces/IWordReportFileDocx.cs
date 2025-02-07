using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportsLibraryOpenXmlSdk.Interfaces
{
    public interface IWordReportFileDocx
    {
        /// <summary>
        /// Путь к файлу отчета с расширением ".docx"
        /// </summary>
        string PathReport { get; set; }

        /// <summary>
        /// Полное название отчета с расширением ".docx"
        /// </summary>
        string NameReport { get; set; }

        /// <summary>
        /// Метод создает файл отчета из шаблона
        /// </summary>
        /// <returns></returns>
        bool CreateWordReportDocx();
    }
}
