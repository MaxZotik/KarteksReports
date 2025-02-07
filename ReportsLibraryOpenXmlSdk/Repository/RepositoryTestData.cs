using ReportsLibraryOpenXmlSdk.Entities;
using ReportsLibraryOpenXmlSdk.Interfaces;
using ReportsLibraryOpenXmlSdk.Enum;
using System.Data;
using Microsoft.Data.SqlClient;
using ReportsLibraryOpenXmlSdk.Services;

namespace ReportsLibraryOpenXmlSdk.Repository
{
    public class RepositoryTestData : IRepository<TestData>
    {
        /// <summary>
        /// Список объектов таблицы "Test_Data" из БД MSSQL содержит объект с данными о последнем тесте
        /// </summary>
        public List<TestData> ArrayObjects { get; set; }

        private readonly IConnectDb _connectionString;
        private readonly Logger _logger;

        public RepositoryTestData(IConnectDb connectDb) 
        {
            _logger = new Logger();
            ArrayObjects = new();
            _connectionString = connectDb;
            GetDataDb();
        }

        private void GetDataDb()
        {
            string selectQuery = $"SELECT top 1 * FROM {Table.Test_Data.ToString()} order by [equipment_creation_datetime] desc";

            try
            {
                using SqlConnection connection = new(_connectionString.ConnectionString);

                SqlDataAdapter adapter = new(selectQuery, connection);
                DataSet ds = new();
                adapter.Fill(ds);

                DataTable? tableDatas = ds.Tables[0];

                if (tableDatas.Rows.Count != 0)
                {
                    int equipmentTypeCode = int.Parse(tableDatas.Rows[0].ItemArray[1].ToString());
                    string? equipmentNumber = tableDatas.Rows[0].ItemArray[2].ToString();
                    string? fio = tableDatas.Rows[0].ItemArray[3].ToString();

                    _ = DateTime.TryParse(tableDatas.Rows[0].ItemArray[4].ToString(), out DateTime equipmentCreationDatetime);
                    _ = DateTime.TryParse(tableDatas.Rows[0].ItemArray[5].ToString(), out DateTime modeOneStartDatetime);
                    _ = DateTime.TryParse(tableDatas.Rows[0].ItemArray[6].ToString(), out DateTime modeOneEndDatetime);
                    _ = DateTime.TryParse(tableDatas.Rows[0].ItemArray[9].ToString(), out DateTime modeTwoStartDatetime);
                    _ = DateTime.TryParse(tableDatas.Rows[0].ItemArray[10].ToString(), out DateTime modeTwoEndDatetime);

                    float averageTestRotationSpeed;

                    if (tableDatas.Rows[0].ItemArray[13].ToString() == "")
                    {
                        averageTestRotationSpeed = 0.0f;
                    }
                    else
                    {
                        averageTestRotationSpeed = float.Parse(tableDatas.Rows[0].ItemArray[13].ToString());
                    }

                    float averageLoadRotationSpeed;

                    if (tableDatas.Rows[0].ItemArray.Length == 15 && tableDatas.Rows[0].ItemArray[14].ToString() != "")
                    {
                        averageLoadRotationSpeed = float.Parse(tableDatas.Rows[0].ItemArray[14].ToString());
                    }
                    else
                    {
                        averageLoadRotationSpeed = 0.0f;
                    }
                    


                    ArrayObjects.Add(new TestData(equipmentTypeCode, equipmentNumber, fio, equipmentCreationDatetime, modeOneStartDatetime, modeOneEndDatetime,
                        modeTwoStartDatetime, modeTwoEndDatetime, averageTestRotationSpeed, averageLoadRotationSpeed));
                }
                else
                {
                    _= _logger.LogAsync(@$"В БД данные о тесте отсутствуют!");

                    StatusReport.SetStatusExaminedError();
                    StatusReport.SetStatusExaminedError();
                }
            }
            catch (Exception ex)
            {
                _= _logger.LogAsync($@"RepositoryTestData : {ex.Message}");

                StatusReport.SetStatusExaminedError();
                StatusReport.SetStatusExaminedError();
            }
        }
    }
}
