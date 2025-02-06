using ReportsLibraryOpenXmlSdk.Entities;
using ReportsLibraryOpenXmlSdk.Interfaces;
using ReportsLibraryOpenXmlSdk.Services;
using System.Xml;

namespace ReportsLibraryOpenXmlSdk.DbData
{
    public class ConnectionStringMssql : IConnectDb
    {
        /// <summary>
        /// Строка подключения к БД
        /// </summary>
        public string? ConnectionString { get; set; }

        private readonly string _path = @$"{AppDomain.CurrentDomain.BaseDirectory}Resources\Settings\_connections.xml";

        private readonly ConnectionConfigMssql? _connectionConfigMssql;

        private readonly Logger _logger;

        public ConnectionStringMssql()
        {
            _logger = new();
            _connectionConfigMssql = GetConnectionConfigMssql();
            InitConnectionString();
        }

        private void InitConnectionString()
        {
            if (_connectionConfigMssql != null)
            {
                ConnectionString = $"Server={_connectionConfigMssql.Server};Database={_connectionConfigMssql.NameDb};User ID={_connectionConfigMssql.User};Password={_connectionConfigMssql.Password};Encrypt=False;TrustServerCertificate=False";
            }
            else
            {
                _= _logger.LogAsync(@$"Не удалось получить строку подключения к БД!");
            }
        }

        private ConnectionConfigMssql? GetConnectionConfigMssql()
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
                        return new ConnectionConfigMssql(
                            xElem.ChildNodes[0].InnerText,
                            xElem.ChildNodes[1].InnerText,
                            xElem.ChildNodes[2].InnerText,
                            xElem.ChildNodes[3].InnerText);
                    }
                }
            }
            else
            {
                _= _logger.LogAsync(@$"Отсутствует файл настроек подключения к БД! : {_path}");
            }

            return null;
        }
    }
}
