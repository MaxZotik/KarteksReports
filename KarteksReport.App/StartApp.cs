using ReportsLibraryOpenXmlSdk.DbData;
using ReportsLibraryOpenXmlSdk.Entities;
using ReportsLibraryOpenXmlSdk.Interfaces;
using ReportsLibraryOpenXmlSdk.Repository;
using ReportsLibraryOpenXmlSdk.Services;
using ReportsLibraryOpenXmlSdk.Word;
using System.Windows.Controls;

namespace KarteksReport.App
{   
    public class StartApp
    {
        private readonly Logger _logger;
        private readonly IConnectDb _connectDb;
        private readonly RepositoryGearboxes _repositoryGearboxes;
        private readonly RepositoryChannelConfig _repositoryChannelConfig;
        private readonly RepositoryTestData _repositoryTestData;

        public StartApp() 
        {
            _logger = new Logger();
            _connectDb = new ConnectionStringMssql();
            _repositoryGearboxes = new(_connectDb);
            _repositoryChannelConfig = new();
            _repositoryTestData = new(_connectDb);
        }


        public async Task StartExamined(ProgressBar progressBar)
        {
            progressBar.Value += 2;
            await Task.Delay(500);

            try
            {
                WordReportDocx reportDocx = new("Template_examined.dotx", _repositoryTestData.ArrayObjects[0].EquipmentCreationDatetime, _repositoryTestData.ArrayObjects[0].EquipmentNumber);

                if (!reportDocx.CreateWordReportDocx())
                {
                    progressBar.Value += 3;
                    StatusReport.SetStatusExaminedError();
                    return;
                }

                Template template = new(
                    reportDocx,
                    _repositoryGearboxes.ArrayObjects.First(g => g.Id == _repositoryTestData.ArrayObjects[0].EquipmentTypeCode),
                    _repositoryTestData.ArrayObjects[0].Fio,
                    _repositoryTestData.ArrayObjects[0].AverageTestRotationSpeed.ToString(),
                    _repositoryTestData.ArrayObjects[0].AverageLoadRotationSpeed.ToString(),
                    _repositoryTestData.ArrayObjects[0].EquipmentCreationDatetime);

                WordReportFillOutExamined wordReportFillOutExamined = new(template);
                RepositoryDiagramExamined repositoryDiagramExamined = new(_repositoryTestData.ArrayObjects[0].EquipmentTypeCode);           

                RepositoryArchivePointExamined repositoryArchivePointExamined = new(_connectDb, repositoryDiagramExamined, _repositoryTestData, _repositoryChannelConfig);
                WordReportTableExamined wordReportTableExamined = new(template, repositoryArchivePointExamined);

                progressBar.Value += +3;
                await Task.Delay(500);

                StatusReport.SetStatusExaminedReady();
            }
            catch (Exception ex)
            {
                progressBar.Value += 3;
                _ = _logger.LogAsync(ex.Message);
                StatusReport.SetStatusExaminedError();
            }
        }


        public async Task StartWorkload(ProgressBar progressBar)
        {
            progressBar.Value += 2;
            await Task.Delay(500);

            try
            {
                WordReportDocx reportDocxWork = new("Template_workload.dotx", _repositoryTestData.ArrayObjects[0].EquipmentCreationDatetime);

                if (!reportDocxWork.CreateWordReportDocx())
                {
                    progressBar.Value += 3;
                    StatusReport.SetStatusExaminedError();
                    return;
                }

                Template template = new(
                    reportDocxWork,
                    _repositoryGearboxes.ArrayObjects.First(g => g.Id == _repositoryTestData.ArrayObjects[0].EquipmentTypeCode),
                    _repositoryTestData.ArrayObjects[0].Fio,
                    _repositoryTestData.ArrayObjects[0].AverageTestRotationSpeed.ToString(),
                    _repositoryTestData.ArrayObjects[0].AverageLoadRotationSpeed.ToString(),
                    _repositoryTestData.ArrayObjects[0].EquipmentCreationDatetime);

                WordReportFillOutWorkload wordReportFillOutWorkload = new(template);
                RepositoryDiagramWorkload repositoryDiagramWorkload = new();

                RepositoryArchivePointWorkload repositoryArchivePointWorkload = new(_connectDb, repositoryDiagramWorkload, _repositoryTestData, _repositoryChannelConfig);
                WordReportTableWorkload wordReportTableWorkload = new WordReportTableWorkload(template, repositoryArchivePointWorkload);

                progressBar.Value += 3;
                await Task.Delay(500);

                StatusReport.SetStatusWorkloadReady();
            }
            catch (Exception ex)
            {
                progressBar.Value += 3;
                _ = _logger.LogAsync(ex.Message);
                StatusReport.SetStatusWorkloadError();
            }
        }
    }
}
