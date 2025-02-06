using ReportsLibraryOpenXmlSdk.Entities;
using ReportsLibraryOpenXmlSdk.Interfaces;
using ReportsLibraryOpenXmlSdk.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace ReportsLibraryOpenXmlSdk.Repository
{
    public class RepositoryChannelConfig : IRepository<ChannelConfig>
    {
        public List<ChannelConfig> ArrayObjects { get; set; }

        private readonly Logger _logger;
        private readonly string _path = @$"{AppDomain.CurrentDomain.BaseDirectory}Resources\Settings\_configChannel.xml";
        public RepositoryChannelConfig() 
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
                            ArrayObjects.Add(new ChannelConfig(
                                int.Parse(xElem.ChildNodes[0].InnerText),
                                int.Parse(xElem.ChildNodes[1].InnerText),
                                int.Parse(xElem.ChildNodes[2].InnerText),
                                int.Parse(xElem.ChildNodes[3].InnerText),
                                int.Parse(xElem.ChildNodes[4].InnerText),
                                int.Parse(xElem.ChildNodes[5].InnerText)));
                        }
                    }
                }
                else
                {
                    _ = _logger.LogAsync(@$"Отсутствует файл настроек конфигурации каналов! : {_path}");

                    StatusReport.SetStatusExaminedError();
                    StatusReport.SetStatusWorkloadError();
                }
            }
            catch (Exception ex)
            {
                _ = _logger.LogAsync(ex.Message);

                StatusReport.SetStatusExaminedError();
                StatusReport.SetStatusWorkloadError();
            }
        }
    }
}
