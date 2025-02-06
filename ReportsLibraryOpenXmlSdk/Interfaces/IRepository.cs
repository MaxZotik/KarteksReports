using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportsLibraryOpenXmlSdk.Interfaces
{
    public interface IRepository<T>
    {
        /// <summary>
        /// Список объектов таблицы "Gearboxes" из БД MSSQL
        /// </summary>
        List<T> ArrayObjects { get; set; }
    }
}
