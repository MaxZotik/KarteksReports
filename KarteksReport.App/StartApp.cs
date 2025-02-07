using ReportsLibraryOpenXmlSdk.DbData;
using ReportsLibraryOpenXmlSdk.Entities;
using ReportsLibraryOpenXmlSdk.Interfaces;
using ReportsLibraryOpenXmlSdk.Repository;
using ReportsLibraryOpenXmlSdk.Services;
using ReportsLibraryOpenXmlSdk.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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




        //public static async Task Start(ProgressBar progressBar, Button button, TextBlock textBlock)
        //{
        //    Logger _logger = new Logger();
        //    textBlock.Text = StatusReport.Status;

        //    try
        //    {
        //        IConnectDb connectDb = new ConnectionStringMssql();

        //        RepositoryGearboxes repositoryGearboxes = new(connectDb);

        //        RepositoryTestData repositoryTestData = new(connectDb);

        //        List<Template> templateExamineds = new();

        //        WordReportDocx reportDocx = new("Template_examined.dotx", repositoryTestData.ArrayObjects[0].EquipmentCreationDatetime, repositoryTestData.ArrayObjects[0].EquipmentNumber);

        //        textBlock.Text = StatusReport.Status;
        //        progressBar.Value += 1;

        //        WordReportDocx reportDocxWork = new("Template_workload.dotx", repositoryTestData.ArrayObjects[0].EquipmentCreationDatetime);

        //        textBlock.Text = StatusReport.Status;
        //        progressBar.Value += 1;
        //        await Task.Delay(500);

        //        templateExamineds.Add(new Template(
        //            reportDocx,
        //            repositoryGearboxes.ArrayObjects.First(g => g.Id == repositoryTestData.ArrayObjects[0].EquipmentTypeCode),
        //            repositoryTestData.ArrayObjects[0].Fio,
        //            repositoryTestData.ArrayObjects[0].AverageTestRotationSpeed.ToString(),
        //            repositoryTestData.ArrayObjects[0].AverageLoadRotationSpeed.ToString(),
        //            repositoryTestData.ArrayObjects[0].EquipmentCreationDatetime));

        //        templateExamineds.Add(new Template(
        //                reportDocxWork,
        //            repositoryGearboxes.ArrayObjects.First(g => g.Id == repositoryTestData.ArrayObjects[0].EquipmentTypeCode),
        //            repositoryTestData.ArrayObjects[0].Fio,
        //            repositoryTestData.ArrayObjects[0].AverageTestRotationSpeed.ToString(),
        //            repositoryTestData.ArrayObjects[0].AverageLoadRotationSpeed.ToString(),
        //            repositoryTestData.ArrayObjects[0].EquipmentCreationDatetime));

        //        textBlock.Text = StatusReport.Status;
        //        progressBar.Value += 2;
        //        await Task.Delay(500);


        //        WordReportFillOutExamined wordReportFillOutExamined = new(templateExamineds[0]);
        //        RepositoryDiagramExamined repositoryDiagramExamined = new(repositoryTestData.ArrayObjects[0].EquipmentTypeCode);

        //        textBlock.Text = StatusReport.Status;
        //        progressBar.Value += 2;
        //        await Task.Delay(500);

        //        WordReportFillOutWorkload wordReportFillOutWorkload = new(templateExamineds[1]);
        //        RepositoryDiagramWorkload repositoryDiagramWorkload = new();

        //        textBlock.Text = StatusReport.Status;
        //        progressBar.Value += 2;
        //        await Task.Delay(500);

                
        //        RepositoryChannelConfig repositoryChannelConfig = new();

        //        RepositoryArchivePointExamined repositoryArchivePointExamined = new(connectDb, repositoryDiagramExamined, repositoryTestData, repositoryChannelConfig);
        //        WordReportTableExamined wordReportTableExamined = new WordReportTableExamined(templateExamineds[0], repositoryArchivePointExamined);

        //        textBlock.Text = StatusReport.Status;
        //        progressBar.Value += 1;

        //        RepositoryArchivePointWorkload repositoryArchivePointWorkload = new(connectDb, repositoryDiagramWorkload, repositoryTestData, repositoryChannelConfig);
        //        WordReportTableWorkload wordReportTableWorkload = new WordReportTableWorkload(templateExamineds[1], repositoryArchivePointWorkload);

        //        textBlock.Text = StatusReport.Status;
        //        progressBar.Value += 1;
        //        await Task.Delay(500);


        //        button.IsEnabled = true;

        //        textBlock.Text = StatusReport.Status;
        //        StatusReport.SetStatusExaminedReady();
        //        StatusReport.SetStatusWorkloadReady();
        //    }
        //    catch (Exception ex)
        //    {
        //        _ = _logger.LogAsync(ex.Message);

        //        textBlock.Text = StatusReport.Status;
        //        StatusReport.SetStatusExaminedError();
        //        StatusReport.SetStatusWorkloadError();
        //    }
        //}



        //public static async Task Start(ProgressBar progressBar, Button button, TextBlock textBlock)
        //{
        //    Logger _logger = new Logger();
        //    textBlock.Text = StatusReport.Status;

        //    try
        //    {
        //        IConnectDb _connectDb = new ConnectionStringMssql();

        //        RepositoryGearboxes _repositoryGearboxes = new(_connectDb);

        //        RepositoryTestData _repositoryTestData = new(_connectDb);

        //        List<Template> templateExamineds = new();

        //        WordReportDocx reportDocx = new("Template_examined.dotx", _repositoryTestData.ArrayObjects[0].EquipmentCreationDatetime, _repositoryTestData.ArrayObjects[0].EquipmentNumber);

        //        textBlock.Text = StatusReport.Status;
        //        progressBar.Value += 1;

        //        WordReportDocx reportDocxWork = new("Template_workload.dotx", _repositoryTestData.ArrayObjects[0].EquipmentCreationDatetime);

        //        textBlock.Text = StatusReport.Status;
        //        progressBar.Value += 1;
        //        await Task.Delay(500);

        //        templateExamineds.Add(new Template(
        //            reportDocx,
        //            _repositoryGearboxes.ArrayObjects.First(g => g.Id == _repositoryTestData.ArrayObjects[0].EquipmentTypeCode),
        //            _repositoryTestData.ArrayObjects[0].Fio,
        //            _repositoryTestData.ArrayObjects[0].AverageTestRotationSpeed.ToString(),
        //            _repositoryTestData.ArrayObjects[0].AverageLoadRotationSpeed.ToString(),
        //            _repositoryTestData.ArrayObjects[0].EquipmentCreationDatetime));

        //        templateExamineds.Add(new Template(
        //                reportDocxWork,
        //            _repositoryGearboxes.ArrayObjects.First(g => g.Id == _repositoryTestData.ArrayObjects[0].EquipmentTypeCode),
        //            _repositoryTestData.ArrayObjects[0].Fio,
        //            _repositoryTestData.ArrayObjects[0].AverageTestRotationSpeed.ToString(),
        //            _repositoryTestData.ArrayObjects[0].AverageLoadRotationSpeed.ToString(),
        //            _repositoryTestData.ArrayObjects[0].EquipmentCreationDatetime));

        //        textBlock.Text = StatusReport.Status;
        //        progressBar.Value += 2;
        //        await Task.Delay(500);


        //        WordReportFillOutExamined wordReportFillOutExamined = new(templateExamineds[0]);
        //        RepositoryDiagramExamined repositoryDiagramExamined = new(_repositoryTestData.ArrayObjects[0].EquipmentTypeCode);

        //        textBlock.Text = StatusReport.Status;
        //        progressBar.Value += 2;
        //        await Task.Delay(500);

        //        WordReportFillOutWorkload wordReportFillOutWorkload = new(templateExamineds[1]);
        //        RepositoryDiagramWorkload repositoryDiagramWorkload = new();

        //        textBlock.Text = StatusReport.Status;
        //        progressBar.Value += 2;
        //        await Task.Delay(500);


        //        RepositoryChannelConfig _repositoryChannelConfig = new();

        //        RepositoryArchivePointExamined repositoryArchivePointExamined = new(_connectDb, repositoryDiagramExamined, _repositoryTestData, _repositoryChannelConfig);
        //        WordReportTableExamined wordReportTableExamined = new WordReportTableExamined(templateExamineds[0], repositoryArchivePointExamined);

        //        textBlock.Text = StatusReport.Status;
        //        progressBar.Value += 1;

        //        RepositoryArchivePointWorkload repositoryArchivePointWorkload = new(_connectDb, repositoryDiagramWorkload, _repositoryTestData, _repositoryChannelConfig);
        //        WordReportTableWorkload wordReportTableWorkload = new WordReportTableWorkload(templateExamineds[1], repositoryArchivePointWorkload);

        //        textBlock.Text = StatusReport.Status;
        //        progressBar.Value += 1;
        //        await Task.Delay(500);


        //        button.IsEnabled = true;

        //        textBlock.Text = StatusReport.Status;
        //        StatusReport.SetStatusExaminedReady();
        //        StatusReport.SetStatusWorkloadReady();
        //    }
        //    catch (Exception ex)
        //    {
        //        _ = _logger.LogAsync(ex.Message);

        //        textBlock.Text = StatusReport.Status;
        //        StatusReport.SetStatusExaminedError();
        //        StatusReport.SetStatusWorkloadError();
        //    }
        //}
    }
}
