namespace ReportsLibraryOpenXmlSdk.Entities
{
    /// <summary>
    /// Объект из таблицы "Gearboxes" БД MSSQL
    /// </summary>
    /// <param name="id">Код типа редуктора</param>
    /// <param name="reductorName">Название типа редуктора</param>
    /// <param name="reductorScheme">Чертёж типа редуктора</param>
    public class Gearboxes(int id, string reductorName, string reductorScheme)
    {
        /// <summary>
        /// Код типа редуктора
        /// </summary>
        public int Id { get; set; } = id;

        /// <summary>
        /// Название типа редуктора
        /// </summary>
        public string ReductorName { get; set; } = reductorName;

        /// <summary>
        /// Чертёж типа редуктора
        /// </summary>
        public string ReductorScheme { get; set; } = reductorScheme;
    }
}
