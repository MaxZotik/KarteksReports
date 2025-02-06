using DocumentFormat.OpenXml.Packaging;

namespace ReportsLibraryOpenXmlSdk.Entities
{
    /// <summary>
    /// Объект из таблицы "Test_Data" БД MSSQL
    /// </summary>
    /// <param name="equipmentTypeCode">Код типа редуктора</param>
    /// <param name="equipmentNumber">Номер редуктора</param>
    /// <param name="fio">ФИО</param>
    /// <param name="equipmentCreationDatetime">Дата и время создания записи в БД</param>
    /// <param name="modeOneStartDatetime">Дата и время начала режима 1 (по часовой)</param>
    /// <param name="modeOneEndDatetime">Дата и время завершения режима 1 (по часовой)</param>
    /// <param name="modeTwoStartDatetime">Дата и время начала режима 2 (против часовой)</param>
    /// <param name="modeTwoEndDatetime">Дата и время завершения режима 2 (против часовой)</param>
    /// <param name="averageTestRotationSpeed">Средняя частота вращения во время испытаний</param>
    /// <param name="averageLoadRotationSpeed">Средняя частота вращения во время испытаний (нагрузка)</param>
    public class TestData(int equipmentTypeCode, string equipmentNumber, string fio, DateTime equipmentCreationDatetime, DateTime modeOneStartDatetime,
            DateTime modeOneEndDatetime, DateTime modeTwoStartDatetime, DateTime modeTwoEndDatetime, float averageTestRotationSpeed, float averageLoadRotationSpeed)
    {
        /// <summary>
        /// Код типа редуктора
        /// </summary>
        public int EquipmentTypeCode { get; set; } = equipmentTypeCode;

        /// <summary>
        /// Номер редуктора
        /// </summary>
        public string EquipmentNumber { get; set; } = equipmentNumber;

        /// <summary>
        /// ФИО
        /// </summary>
        public string Fio {  get; set; } = fio;

        /// <summary>
        /// Дата и время создания записи в БД
        /// </summary>
        public DateTime EquipmentCreationDatetime { get; set; } = equipmentCreationDatetime;

        /// <summary>
        /// Дата и время начала режима 1 (по часовой)
        /// </summary>
        public DateTime ModeOneStartDatetime { get; set; } = modeOneStartDatetime;

        /// <summary>
        /// Дата и время завершения режима 1 (по часовой)
        /// </summary>
        public DateTime ModeOneEndDatetime { get; set; } = modeOneEndDatetime;

        /// <summary>
        /// Дата и время начала режима 2 (против часовой)
        /// </summary>
        public DateTime ModeTwoStartDatetime { get; set; } = modeTwoStartDatetime;

        /// <summary>
        /// Дата и время завершения режима 2 (против часовой)
        /// </summary>
        public DateTime ModeTwoEndDatetime { get; set; } = modeTwoEndDatetime;

        /// <summary>
        /// Средняя частота вращения во время испытаний 
        /// </summary>
        public float AverageTestRotationSpeed { get; set; } = averageTestRotationSpeed;

        /// <summary>
        /// Средняя частота вращения во время испытаний (нагрузка)
        /// </summary>
        public float AverageLoadRotationSpeed { get; set; } = averageLoadRotationSpeed;
    }
}
