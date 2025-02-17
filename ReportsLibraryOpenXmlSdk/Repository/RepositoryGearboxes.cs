using ReportsLibraryOpenXmlSdk.DbData;
using ReportsLibraryOpenXmlSdk.Entities;
using ReportsLibraryOpenXmlSdk.Interfaces;
using ReportsLibraryOpenXmlSdk.Enum;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Data.SqlClient;
using ReportsLibraryOpenXmlSdk.Services;

namespace ReportsLibraryOpenXmlSdk.Repository
{
    public class RepositoryGearboxes : IRepository<Gearboxes>
    {
        /// <summary>
        /// Список объектов таблицы "Gearboxes" из БД MSSQL
        /// </summary>
        public List<Gearboxes> ArrayObjects { get; set; }

        private DataTable? tableDatas;

        private readonly IConnectDb _connectionString;
        private readonly Logger _logger;

        public RepositoryGearboxes(IConnectDb connectDb)
        {
            _logger = new();
            ArrayObjects = new();
            _connectionString = connectDb;
            GetDataDb();
            _= AddObjectAsync();
        }

        private void GetDataDb()
        {
            string selectQuery = $"SELECT * FROM {Table.Gearboxes.ToString()}";

            try
            {
                using (SqlConnection connection = new(_connectionString.ConnectionString))
                {
                    SqlDataAdapter adapter = new(selectQuery, connection);
                    DataSet ds = new();
                    adapter.Fill(ds);

                    tableDatas = ds.Tables[0];
                }              
            }
            catch (Exception ex)
            {
                _= _logger.LogAsync(ex.Message);

                StatusReport.SetStatusExaminedError();
                StatusReport.SetStatusWorkloadError();
            }
        }

        private async Task AddObjectAsync()
        {
            try
            {
                if (tableDatas is not null && tableDatas.Rows.Count != 0)
                {
                    int count = tableDatas.Rows.Count;

                    for (int i = 0; i < count; i++)
                    {
                        int id = int.Parse(tableDatas.Rows[i].ItemArray[0].ToString());
                        string reductorName = tableDatas.Rows[i].ItemArray[1].ToString();
                        string reductorScheme = tableDatas.Rows[i].ItemArray[2].ToString();

                        ArrayObjects.Add(new Gearboxes(id, reductorName, reductorScheme));
                    }
                }
                else
                {
                    _ = _logger.LogAsync($@"Данные по оборудованию отсутствуют!");

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
