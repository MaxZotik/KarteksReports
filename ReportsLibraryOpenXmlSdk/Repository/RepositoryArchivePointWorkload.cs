using Microsoft.Data.SqlClient;
using ReportsLibraryOpenXmlSdk.Constants;
using ReportsLibraryOpenXmlSdk.Entities;
using ReportsLibraryOpenXmlSdk.Interfaces;
using ReportsLibraryOpenXmlSdk.Services;
using Table = ReportsLibraryOpenXmlSdk.Enum.Table;

namespace ReportsLibraryOpenXmlSdk.Repository
{
    public class RepositoryArchivePointWorkload : IRepository<ArchivePoint>
    {
        public List<ArchivePoint> ArrayObjects { get; set; }

        private readonly Logger _logger;
        private readonly IConnectDb _connectionString;
        private readonly RepositoryDiagramWorkload _repositoryDiagram;
        private readonly RepositoryTestData _repositoryTestData;
        private readonly RepositoryChannelConfig _repositoryChannelConfig;

        public RepositoryArchivePointWorkload(IConnectDb connectDb, RepositoryDiagramWorkload repositoryDiagramWorkload, RepositoryTestData repositoryTestData, RepositoryChannelConfig repositoryChannelConfig)
        {
            ArrayObjects = new();
            _logger = new();
            _connectionString = connectDb;
            _repositoryDiagram = repositoryDiagramWorkload;
            _repositoryTestData = repositoryTestData;
            _repositoryChannelConfig = repositoryChannelConfig;

            InitArrauObject();
        }

        private void InitArrauObject()
        {
            int count = _repositoryDiagram.ArrayObjects.Count;
            int gearboxID = _repositoryTestData.ArrayObjects[0].EquipmentTypeCode;

            for (int i = 0; i < count; i++)
            {
                int _point = _repositoryDiagram.ArrayObjects[i].Point;

                string _nameX = Measurements.MeasurementsText[0];
                float _axisXmovePlus = GetDataArchive(gearboxID, _repositoryDiagram.ArrayObjects[i].AxisXmvk, _repositoryDiagram.ArrayObjects[i].AxisXchannel, _repositoryChannelConfig.ArrayObjects[0].MoveParameter, _repositoryChannelConfig.ArrayObjects[0].MoveFrequency, _repositoryTestData.ArrayObjects[0].ModeOneStartDatetime, _repositoryTestData.ArrayObjects[0].ModeOneEndDatetime);
                float _axisXspeedPlus = GetDataArchive(gearboxID, _repositoryDiagram.ArrayObjects[i].AxisXmvk, _repositoryDiagram.ArrayObjects[i].AxisXchannel, _repositoryChannelConfig.ArrayObjects[0].SpeedParameter, _repositoryChannelConfig.ArrayObjects[0].SpeedFrequency, _repositoryTestData.ArrayObjects[0].ModeOneStartDatetime, _repositoryTestData.ArrayObjects[0].ModeOneEndDatetime);
                float _axisXacceleratePlus = GetDataArchive(gearboxID, _repositoryDiagram.ArrayObjects[i].AxisXmvk, _repositoryDiagram.ArrayObjects[i].AxisXchannel, _repositoryChannelConfig.ArrayObjects[0].AccelerateParameter, _repositoryChannelConfig.ArrayObjects[0].AccelerateFrequency, _repositoryTestData.ArrayObjects[0].ModeOneStartDatetime, _repositoryTestData.ArrayObjects[0].ModeOneEndDatetime);
                float _axisXmoveMinus = GetDataArchive(gearboxID, _repositoryDiagram.ArrayObjects[i].AxisXmvk, _repositoryDiagram.ArrayObjects[i].AxisXchannel, _repositoryChannelConfig.ArrayObjects[0].MoveParameter, _repositoryChannelConfig.ArrayObjects[0].MoveFrequency, _repositoryTestData.ArrayObjects[0].ModeTwoStartDatetime, _repositoryTestData.ArrayObjects[0].ModeTwoEndDatetime);
                float _axisXspeedMinus = GetDataArchive(gearboxID, _repositoryDiagram.ArrayObjects[i].AxisXmvk, _repositoryDiagram.ArrayObjects[i].AxisXchannel, _repositoryChannelConfig.ArrayObjects[0].SpeedParameter, _repositoryChannelConfig.ArrayObjects[0].SpeedFrequency, _repositoryTestData.ArrayObjects[0].ModeTwoStartDatetime, _repositoryTestData.ArrayObjects[0].ModeTwoEndDatetime);
                float _axisXaccelerateMinus = GetDataArchive(gearboxID, _repositoryDiagram.ArrayObjects[i].AxisXmvk, _repositoryDiagram.ArrayObjects[i].AxisXchannel, _repositoryChannelConfig.ArrayObjects[0].AccelerateParameter, _repositoryChannelConfig.ArrayObjects[0].AccelerateFrequency, _repositoryTestData.ArrayObjects[0].ModeTwoStartDatetime, _repositoryTestData.ArrayObjects[0].ModeTwoEndDatetime);

                string _nameY = Measurements.MeasurementsText[1];
                float _axisYmovePlus = GetDataArchive(gearboxID, _repositoryDiagram.ArrayObjects[i].AxisYmvk, _repositoryDiagram.ArrayObjects[i].AxisYchannel, _repositoryChannelConfig.ArrayObjects[0].MoveParameter, _repositoryChannelConfig.ArrayObjects[0].MoveFrequency, _repositoryTestData.ArrayObjects[0].ModeOneStartDatetime, _repositoryTestData.ArrayObjects[0].ModeOneEndDatetime);
                float _axisYspeedPlus = GetDataArchive(gearboxID, _repositoryDiagram.ArrayObjects[i].AxisYmvk, _repositoryDiagram.ArrayObjects[i].AxisYchannel, _repositoryChannelConfig.ArrayObjects[0].SpeedParameter, _repositoryChannelConfig.ArrayObjects[0].SpeedFrequency, _repositoryTestData.ArrayObjects[0].ModeOneStartDatetime, _repositoryTestData.ArrayObjects[0].ModeOneEndDatetime);
                float _axisYacceleratePlus = GetDataArchive(gearboxID, _repositoryDiagram.ArrayObjects[i].AxisYmvk, _repositoryDiagram.ArrayObjects[i].AxisYchannel, _repositoryChannelConfig.ArrayObjects[0].AccelerateParameter, _repositoryChannelConfig.ArrayObjects[0].AccelerateFrequency, _repositoryTestData.ArrayObjects[0].ModeOneStartDatetime, _repositoryTestData.ArrayObjects[0].ModeOneEndDatetime);
                float _axisYmoveMinus = GetDataArchive(gearboxID, _repositoryDiagram.ArrayObjects[i].AxisYmvk, _repositoryDiagram.ArrayObjects[i].AxisYchannel, _repositoryChannelConfig.ArrayObjects[0].MoveParameter, _repositoryChannelConfig.ArrayObjects[0].MoveFrequency, _repositoryTestData.ArrayObjects[0].ModeTwoStartDatetime, _repositoryTestData.ArrayObjects[0].ModeTwoEndDatetime);
                float _axisYspeedMinus = GetDataArchive(gearboxID, _repositoryDiagram.ArrayObjects[i].AxisYmvk, _repositoryDiagram.ArrayObjects[i].AxisYchannel, _repositoryChannelConfig.ArrayObjects[0].SpeedParameter, _repositoryChannelConfig.ArrayObjects[0].SpeedFrequency, _repositoryTestData.ArrayObjects[0].ModeTwoStartDatetime, _repositoryTestData.ArrayObjects[0].ModeTwoEndDatetime);
                float _axisYaccelerateMinus = GetDataArchive(gearboxID, _repositoryDiagram.ArrayObjects[i].AxisYmvk, _repositoryDiagram.ArrayObjects[i].AxisYchannel, _repositoryChannelConfig.ArrayObjects[0].AccelerateParameter, _repositoryChannelConfig.ArrayObjects[0].AccelerateFrequency, _repositoryTestData.ArrayObjects[0].ModeTwoStartDatetime, _repositoryTestData.ArrayObjects[0].ModeTwoEndDatetime);

                string _nameZ = Measurements.MeasurementsText[2];
                float _axisZmovePlus = GetDataArchive(gearboxID, _repositoryDiagram.ArrayObjects[i].AxisZmvk, _repositoryDiagram.ArrayObjects[i].AxisZchannel, _repositoryChannelConfig.ArrayObjects[0].MoveParameter, _repositoryChannelConfig.ArrayObjects[0].MoveFrequency, _repositoryTestData.ArrayObjects[0].ModeOneStartDatetime, _repositoryTestData.ArrayObjects[0].ModeOneEndDatetime);
                float _axisZspeedPlus = GetDataArchive(gearboxID, _repositoryDiagram.ArrayObjects[i].AxisZmvk, _repositoryDiagram.ArrayObjects[i].AxisZchannel, _repositoryChannelConfig.ArrayObjects[0].SpeedParameter, _repositoryChannelConfig.ArrayObjects[0].SpeedFrequency, _repositoryTestData.ArrayObjects[0].ModeOneStartDatetime, _repositoryTestData.ArrayObjects[0].ModeOneEndDatetime);
                float _axisZacceleratePlus = GetDataArchive(gearboxID, _repositoryDiagram.ArrayObjects[i].AxisZmvk, _repositoryDiagram.ArrayObjects[i].AxisZchannel, _repositoryChannelConfig.ArrayObjects[0].AccelerateParameter, _repositoryChannelConfig.ArrayObjects[0].AccelerateFrequency, _repositoryTestData.ArrayObjects[0].ModeOneStartDatetime, _repositoryTestData.ArrayObjects[0].ModeOneEndDatetime);
                float _axisZmoveMinus = GetDataArchive(gearboxID, _repositoryDiagram.ArrayObjects[i].AxisZmvk, _repositoryDiagram.ArrayObjects[i].AxisZchannel, _repositoryChannelConfig.ArrayObjects[0].MoveParameter, _repositoryChannelConfig.ArrayObjects[0].MoveFrequency, _repositoryTestData.ArrayObjects[0].ModeTwoStartDatetime, _repositoryTestData.ArrayObjects[0].ModeTwoEndDatetime);
                float _axisZspeedMinus = GetDataArchive(gearboxID, _repositoryDiagram.ArrayObjects[i].AxisZmvk, _repositoryDiagram.ArrayObjects[i].AxisZchannel, _repositoryChannelConfig.ArrayObjects[0].SpeedParameter, _repositoryChannelConfig.ArrayObjects[0].SpeedFrequency, _repositoryTestData.ArrayObjects[0].ModeTwoStartDatetime, _repositoryTestData.ArrayObjects[0].ModeTwoEndDatetime);
                float _axisZaccelerateMinus = GetDataArchive(gearboxID, _repositoryDiagram.ArrayObjects[i].AxisZmvk, _repositoryDiagram.ArrayObjects[i].AxisZchannel, _repositoryChannelConfig.ArrayObjects[0].AccelerateParameter, _repositoryChannelConfig.ArrayObjects[0].AccelerateFrequency, _repositoryTestData.ArrayObjects[0].ModeTwoStartDatetime, _repositoryTestData.ArrayObjects[0].ModeTwoEndDatetime);

                ArrayObjects.Add(new ArchivePoint(_point, _nameX, _axisXmovePlus, _axisXspeedPlus, _axisXacceleratePlus, _axisXmoveMinus, _axisXspeedMinus, _axisXaccelerateMinus,
                    _nameY, _axisYmovePlus, _axisYspeedPlus, _axisYacceleratePlus, _axisYmoveMinus, _axisYspeedMinus, _axisYaccelerateMinus,
                    _nameZ, _axisZmovePlus, _axisZspeedPlus, _axisZacceleratePlus, _axisZmoveMinus, _axisZspeedMinus, _axisZaccelerateMinus));
            }
        }


        private float GetDataArchive(int gearboxID, int mvk, int channel, int parameter, int frequency, DateTime? dateTimeStart, DateTime? dateTimeEnd)
        {
            DateTime dateTime = new DateTime();

            if (dateTimeStart == dateTime || dateTimeEnd == dateTime)
            {
                return 0;
            }
            else
            {
                try
                {
                    DateTime dtStart = (DateTime)dateTimeStart;
                    DateTime dtEnd = (DateTime)dateTimeEnd;

                    string query = $@"SELECT ISNULL ((SELECT MAX([MVK Value Max]) FROM {Table.Archive.ToString()} WHERE [Equipment] = {gearboxID} AND [MVK Number] = {mvk} AND [Chanel] = {channel} AND [ID Parameters] = {parameter} AND [ID Frequency] = {frequency} AND [Time] BETWEEN CONVERT(datetime, '{dtStart.ToString("yyyy-MM-dd HH:mm:ss")}', 20) AND CONVERT(datetime, '{dtEnd.ToString("yyyy-MM-dd HH:mm:ss")}', 20)), 0.0)";

                    using SqlConnection connection = new(_connectionString.ConnectionString);

                    connection.Open();

                    SqlCommand command = new(query, connection);

                    double data = Convert.ToDouble(command.ExecuteScalar());

                    Math.Round(data, 3);

                    return float.Parse(data.ToString());
                }
                catch (Exception ex)
                {
                    _ = _logger.LogAsync($@"RepositoryArchivePointWorkload : {ex.Message}");
                    StatusReport.SetStatusWorkloadError();
                }
            }      

            return 0;
        }
    }
}
