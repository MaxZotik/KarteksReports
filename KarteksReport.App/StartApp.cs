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
        public static async Task Start(ProgressBar progressBar, Button button, TextBlock textBlock)
        {
            Logger _logger = new Logger();
            textBlock.Text = StatusReport.Status;

            try
            {
                IConnectDb connectDb = new ConnectionStringMssql();

                RepositoryGearboxes repositoryGearboxes = new(connectDb);

                RepositoryTestData repositoryTestData = new(connectDb);

                List<Template> templateExamineds = new();

                WordReportDocx reportDocx = new("Template_examined.dotx", repositoryTestData.ArrayObjects[0].EquipmentCreationDatetime, repositoryTestData.ArrayObjects[0].EquipmentNumber);

                textBlock.Text = StatusReport.Status;
                progressBar.Value += 1;

                WordReportDocx reportDocxWork = new("Template_workload.dotx", repositoryTestData.ArrayObjects[0].EquipmentCreationDatetime);

                textBlock.Text = StatusReport.Status;
                progressBar.Value += 1;
                await Task.Delay(500);

                templateExamineds.Add(new Template(
                    reportDocx,
                    repositoryGearboxes.ArrayObjects.First(g => g.Id == repositoryTestData.ArrayObjects[0].EquipmentTypeCode),
                    repositoryTestData.ArrayObjects[0].Fio,
                    repositoryTestData.ArrayObjects[0].AverageTestRotationSpeed.ToString(),
                    repositoryTestData.ArrayObjects[0].AverageLoadRotationSpeed.ToString(),
                    repositoryTestData.ArrayObjects[0].EquipmentCreationDatetime));

                templateExamineds.Add(new Template(
                        reportDocxWork,
                    repositoryGearboxes.ArrayObjects.First(g => g.Id == repositoryTestData.ArrayObjects[0].EquipmentTypeCode),
                    repositoryTestData.ArrayObjects[0].Fio,
                    repositoryTestData.ArrayObjects[0].AverageTestRotationSpeed.ToString(),
                    repositoryTestData.ArrayObjects[0].AverageLoadRotationSpeed.ToString(),
                    repositoryTestData.ArrayObjects[0].EquipmentCreationDatetime));

                textBlock.Text = StatusReport.Status;
                progressBar.Value += 2;
                await Task.Delay(500);


                WordReportFillOutExamined wordReportFillOutExamined = new(templateExamineds[0]);
                RepositoryDiagramExamined repositoryDiagramExamined = new(repositoryTestData.ArrayObjects[0].EquipmentTypeCode);

                textBlock.Text = StatusReport.Status;
                progressBar.Value += 2;
                await Task.Delay(500);

                WordReportFillOutWorkload wordReportFillOutWorkload = new(templateExamineds[1]);
                RepositoryDiagramWorkload repositoryDiagramWorkload = new();

                textBlock.Text = StatusReport.Status;
                progressBar.Value += 2;
                await Task.Delay(500);

                
                RepositoryChannelConfig repositoryChannelConfig = new();

                RepositoryArchivePointExamined repositoryArchivePointExamined = new(connectDb, repositoryDiagramExamined, repositoryTestData, repositoryChannelConfig);
                WordReportTableExamined wordReportTableExamined = new WordReportTableExamined(templateExamineds[0], repositoryArchivePointExamined);

                textBlock.Text = StatusReport.Status;
                progressBar.Value += 1;

                RepositoryArchivePointWorkload repositoryArchivePointWorkload = new(connectDb, repositoryDiagramWorkload, repositoryTestData, repositoryChannelConfig);
                WordReportTableWorkload wordReportTableWorkload = new WordReportTableWorkload(templateExamineds[1], repositoryArchivePointWorkload);

                textBlock.Text = StatusReport.Status;
                progressBar.Value += 1;
                await Task.Delay(500);


                button.IsEnabled = true;

                textBlock.Text = StatusReport.Status;
                StatusReport.SetStatusExaminedReady();
                StatusReport.SetStatusWorkloadReady();
            }
            catch (Exception ex)
            {
                _ = _logger.LogAsync(ex.Message);

                textBlock.Text = StatusReport.Status;
                StatusReport.SetStatusExaminedError();
                StatusReport.SetStatusWorkloadError();
            }
        }       
    }
}
