using ReportsLibraryOpenXmlSdk.Entities;
using ReportsLibraryOpenXmlSdk.Interfaces;
using ReportsLibraryOpenXmlSdk.Services;
using System.Xml;

namespace ReportsLibraryOpenXmlSdk.Repository
{
    public class RepositoryDiagramExamined : IRepository<DiagramExamined>
    {
        /// <summary>
        /// Список объектов "DiagramConfig"
        /// </summary>
        public List<DiagramExamined> ArrayObjects { get; set; }

        private readonly Logger _logger;
        private readonly string _path = @$"{AppDomain.CurrentDomain.BaseDirectory}Resources\Settings\_diagramExamined.xml";

        public RepositoryDiagramExamined(int gearboxID) 
        {
            ArrayObjects = new();
            _logger = new();
            InitArrauObject(gearboxID);
        }

        private void InitArrauObject(int _gearboxID)
        {
            try
            {
                if (File.Exists(_path))
                {
                    XmlDocument xDoc = new();
                    xDoc.Load(_path);
                    XmlElement? root = xDoc.DocumentElement;

                    if (root != null)
                    {
                        foreach (XmlElement xElem in root)
                        {
                            if (_gearboxID == int.Parse(xElem.ChildNodes[0].InnerText))
                            {
                                ArrayObjects.Add(new DiagramExamined(
                                int.Parse(xElem.ChildNodes[0].InnerText),
                                int.Parse(xElem.ChildNodes[1].InnerText),
                                int.Parse(xElem.ChildNodes[2].InnerText),
                                int.Parse(xElem.ChildNodes[3].InnerText),
                                int.Parse(xElem.ChildNodes[4].InnerText),
                                int.Parse(xElem.ChildNodes[5].InnerText),
                                int.Parse(xElem.ChildNodes[6].InnerText),
                                int.Parse(xElem.ChildNodes[7].InnerText)));
                            }
                        }
                    }
                }
                else
                {
                    _ = _logger.LogAsync(@$"Отсутствует файл настроек схемы подключния! : {_path}");

                    StatusReport.SetStatusExaminedError();
                }
            }
            catch (Exception ex)
            {
                _ = _logger.LogAsync(ex.Message);

                StatusReport.SetStatusExaminedError();
            }
        }
    }
}
