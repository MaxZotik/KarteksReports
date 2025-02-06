using ReportsLibraryOpenXmlSdk.Entities;
using ReportsLibraryOpenXmlSdk.Interfaces;

namespace ReportsLibraryOpenXmlSdk.Word
{
    /// <summary>
    /// Класс шаблон для создания отчета "Шаблон испытуемый"
    /// </summary>
    /// <param name="wordReportFileDocx">Интерфейс содержит путь к файлу отчета с расширением '.docx' и полное название файла отчета с расширением</param>
    /// <param name="gearboxes">Объект из таблицы "Gearboxes" БД MSSQL</param>
    /// <param name="fio">ФИО Исполнители</param>
    /// <param name="rotation">Частота вращения во время испытаний</param>
    /// /// <param name="rotationWorkload">Частота вращения во время испытаний (нагрузка)</param>
    /// <param name="dateTime">Дата и время</param>
    public class Template(IWordReportFileDocx wordReportFileDocx, Gearboxes gearboxes, string fio, string rotation, string rotationWorkload, DateTime dateTime)
    {

        /// <summary>
        /// Интерфейс содержит путь к файлу отчета с расширением '.docx' и полное название файла отчета с расширением
        /// </summary>
        public IWordReportFileDocx WordReportFileDocx { get; private set; } = wordReportFileDocx;

        /// <summary>
        /// Редуктор название заголовок
        /// </summary>
        public string ReductorNameTitle { get; private set; } = $@"{gearboxes.ReductorName}, {gearboxes.ReductorScheme}";

        /// <summary>
        /// Дата и время заголовок
        /// </summary>
        public string DateAndTimeHeading { get; private set; } = dateTime.ToString("dd.MM.yyyy");

        /// <summary>
        /// Редуктор название объект испытаний
        /// </summary>
        public string ReductorNameOfTheTestObject { get; private set; } = $@"{gearboxes.ReductorName}, {gearboxes.ReductorScheme}";

        /// <summary>
        /// ФИО Исполнители
        /// </summary>
        public string FioPerformers {  get; private set; } = fio;

        /// <summary>
        /// Редуктор название измерения вибрации
        /// </summary>
        public string ReductorNumberOfVibrationMeasurements { get; private set; } = $@"{gearboxes.ReductorName}, {gearboxes.ReductorScheme}";

        /// <summary>
        /// Дата и время измерения вибрации
        /// </summary>
        public string DateAndTimeVibrationMeasurement { get; private set; } = dateTime.ToString("dd.MM.yyyy");

        /// <summary>
        /// Частота вращения во время испытаний
        /// </summary>
        public string FrequencyOfRotationDuringTesting {  get; private set; } = rotation;

        /// <summary>
        /// Частота вращения во время испытаний (нагрузка)
        /// </summary>
        public string FrequencyOfRotationDuringWorkload { get; private set; } = rotationWorkload;

        /// <summary>
        /// Схема редуктора
        /// </summary>
        public string SchemaReductor {  get; private set; } = $@"{AppDomain.CurrentDomain.BaseDirectory}Resources\MnemonicDiagrams\{gearboxes.Id}.png";

        /// <summary>
        /// Редуктор название критерии оценки
        /// </summary>
        public string ReductorNumberOfAssessmentCriteria { get; private set; } = $@"{gearboxes.ReductorName}, {gearboxes.ReductorScheme}";

        /// <summary>
        /// Редуктор название заключение
        /// </summary>
        public string ReductorNameConclusion { get; private set; } = $@"{gearboxes.ReductorName}, {gearboxes.ReductorScheme}";

        /// <summary>
        /// ФИО Заключение
        /// </summary>
        public string FioConclusion { get; private set; } = fio;
    }
}
