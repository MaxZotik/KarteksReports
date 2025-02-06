using ReportsLibraryOpenXmlSdk.Entities;
using ReportsLibraryOpenXmlSdk.Interfaces;
using ReportsLibraryOpenXmlSdk.Services;
using System.Xml;

namespace ReportsLibraryOpenXmlSdk.Repository
{
    public class RepositoryDiagramWorkload : IRepository<DiagramWorkload>
    {
        /// <summary>
        /// Список объектов "DiagramWorkload"
        /// </summary>
        public List<DiagramWorkload> ArrayObjects { get; set; }

        private readonly Logger _logger;
        private readonly string _path = @$"{AppDomain.CurrentDomain.BaseDirectory}Resources\Settings\_diagramWorkload.xml";

        public RepositoryDiagramWorkload()
        {
            ArrayObjects = new();
            _logger = new();
            InitArrauObject();
        }

        private void InitArrauObject()
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
                            ArrayObjects.Add(new DiagramWorkload(
                            int.Parse(xElem.ChildNodes[0].InnerText),
                            int.Parse(xElem.ChildNodes[1].InnerText),
                            int.Parse(xElem.ChildNodes[2].InnerText),
                            int.Parse(xElem.ChildNodes[3].InnerText),
                            int.Parse(xElem.ChildNodes[4].InnerText),
                            int.Parse(xElem.ChildNodes[5].InnerText),
                            int.Parse(xElem.ChildNodes[6].InnerText)));
                        }
                    }
                }
                else
                {
                    _ = _logger.LogAsync(@$"Отсутствует файл настроек схемы подключния! : {_path}");

                    StatusReport.SetStatusWorkloadError();
                }
            }
            catch(Exception ex)
            {
                _ = _logger.LogAsync(ex.Message);

                StatusReport.SetStatusWorkloadError();
            }
        }
    }
}
