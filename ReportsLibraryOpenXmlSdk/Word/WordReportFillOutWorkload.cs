using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DFW = DocumentFormat.OpenXml.Wordprocessing;

namespace ReportsLibraryOpenXmlSdk.Word
{
    public class WordReportFillOutWorkload
    {
        private Template _template;

        public WordReportFillOutWorkload(Template template)
        {
            _template = template;
            InsertTextBookmarks();
        }

        private void InsertTextBookmarks()
        {
            using var docWord = WordprocessingDocument.Open(_template.WordReportFileDocx.PathReport, true);

            foreach (BookmarkStart book in docWord.MainDocumentPart.Document.Body.Descendants<BookmarkStart>())
            {
                foreach (string str in Constants.BookmarksWord.BookmarksTemplateExaminedText)
                {
                    if (book.Name == str)
                    {
                        RunProperties runProperties = new RunProperties();

                        RunFonts runFonts = new RunFonts() { Ascii = "Times New Roman" };
                        FontSize fontSize = new FontSize() { Val = "24" };

                        runProperties.Append(runFonts);
                        runProperties.Append(fontSize);

                        var text = new DFW.Text(GetTextBookmark(str));
                        var runElement = new DFW.Run(runProperties, text);
                        book.InsertAfterSelf(runElement);
                    }
                }
            }

            docWord.Save();
        }

        private string GetTextBookmark(string text)
        {
            string str = " ";

            switch (text)
            {
                case "Редуктор_название_заголовок":
                    str = Constants.ConstTitles.TITLE_WORKLOAD;
                    break;
                case "Дата_и_время_заголовок":
                    str = _template.DateAndTimeHeading;
                    break;
                case "Редуктор_название_объект_испытаний":
                    str = Constants.ConstTitles.TITLE_WORKLOAD;
                    break;
                case "ФИО_Исполнители":
                    str = _template.FioPerformers;
                    break;
                case "Редуктор_название_измерения_вибрации":
                    str = Constants.ConstTitles.TITLE_WORKLOAD;
                    break;
                case "Дата_и_время_измерения_вибрации":
                    str = _template.DateAndTimeVibrationMeasurement;
                    break;
                case "ФИО_Заключение":
                    str = _template.FioConclusion;
                    break;
                case "Частота_вращения_во_время_испытаний":
                    str = _template.FrequencyOfRotationDuringWorkload;
                    break;
                case "Редуктор_название_критерии_оценки":
                    str = Constants.ConstTitles.TITLE_WORKLOAD;
                    break;
                default:
                    str = " ";
                    break;
            }

            return str;
        }
    }
}
