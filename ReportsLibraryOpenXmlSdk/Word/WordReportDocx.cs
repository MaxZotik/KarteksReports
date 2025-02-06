using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ReportsLibraryOpenXmlSdk.Services;
using ReportsLibraryOpenXmlSdk.Interfaces;
using ReportsLibraryOpenXmlSdk.Entities;

namespace ReportsLibraryOpenXmlSdk.Word
{
    public class WordReportDocx : IWordReportFileDocx
    {
        private static readonly string _nameDirectory = "Reports";

        private static readonly string _path = AppDomain.CurrentDomain.BaseDirectory;

        private readonly string _pathTemplate;

        private readonly Logger _logger;

        /// <summary>
        /// Путь к файлу отчета с расширением ".docx"
        /// </summary>
        public string PathReport { get; set; }

        /// <summary>
        /// Полное название отчета с расширением ".docx"
        /// </summary>
        public string NameReport {  get; set; }

        static WordReportDocx()
        {
            FileDirectory.CreateDirectory(_path, _nameDirectory);
        }

        /// <summary>
        /// Класс создания файла отчета Word с расширением .docx из файла шаблона
        /// </summary>
        /// <param name="template">Название шаблона с расширением .dotx</param>
        /// <param name="dateTime">Дата созадания теста зачитывается из БД</param>
        /// <param name="equipName">Название оборудования</param>
        public WordReportDocx(string template, DateTime dateTime, string equipName = Constants.ConstTitles.TITLE_WORKLOAD)
        {
            _logger = new Logger();
            _pathTemplate = @$"{_path}Resources\Templates\{template}";
            NameReport = GetNameReport(equipName, dateTime);
            PathReport = $@"{_path}Reports\{NameReport}";
            CreateWordReportDocx();
        }

        private static string GetNameReport(string equipName, DateTime dateTime)
        {
            return @$"{dateTime.ToString("yyyyMMdd_HHmmss")}_{StringWork.ReplaceInvalidChar(equipName)}.docx";
        }

        private void CreateWordReportDocx()
        {
            try
            {
                if (File.Exists(_pathTemplate))
                {
                    File.Copy(_pathTemplate, PathReport, true);

                    using var newDocWord = WordprocessingDocument.Open(PathReport, true);

                    newDocWord.ChangeDocumentType(WordprocessingDocumentType.Document);

                    var mainPart = newDocWord.MainDocumentPart;

                    newDocWord.Save();
                }
                else
                {
                    _= _logger.LogAsync(@$"Файл шаблона не найден! {_pathTemplate}");
                    StatusReport.SetStatusExaminedError();
                    StatusReport.SetStatusWorkloadError();
                }
            }
            catch(Exception ex)
            {
                _= _logger.LogAsync(ex.Message);
                StatusReport.SetStatusExaminedError();
                StatusReport.SetStatusWorkloadError();
            }
           
        }
    }
}
