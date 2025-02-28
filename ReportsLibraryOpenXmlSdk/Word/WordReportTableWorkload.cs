using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ReportsLibraryOpenXmlSdk.Entities;
using ReportsLibraryOpenXmlSdk.Repository;

namespace ReportsLibraryOpenXmlSdk.Word
{
    public class WordReportTableWorkload
    {
        Template _template;
        RepositoryArchivePointWorkload _repositoryArchivePointWorkload;

        public WordReportTableWorkload(Template template, RepositoryArchivePointWorkload repositoryArchivePointWorkload)
        {
            _template = template;
            _repositoryArchivePointWorkload = repositoryArchivePointWorkload;

            AddRowsTable();
            EditPointTable();
            EditCellsTable();
        }

        /// <summary>
        /// Метод добавляет строчки в таблицу
        /// </summary>
        private void AddRowsTable()
        {
            using WordprocessingDocument doc = WordprocessingDocument.Open(_template.WordReportFileDocx.PathReport, true);

            Table table = doc.MainDocumentPart.Document.Body.Elements<Table>().Last();

            int count = _repositoryArchivePointWorkload.ArrayObjects.Count - 1;

            for (int i = 0; i < count; i++)
            {
                TableRow row2 = table.Elements<TableRow>().ElementAt(2);
                TableRow row3 = table.Elements<TableRow>().ElementAt(3);
                TableRow row4 = table.Elements<TableRow>().ElementAt(4);


                TableRow rowCopy2 = (TableRow)row2.CloneNode(true);
                TableRow rowCopy3 = (TableRow)row3.CloneNode(true);
                TableRow rowCopy4 = (TableRow)row4.CloneNode(true);

                table.AppendChild(rowCopy2);
                table.AppendChild(rowCopy3);
                table.AppendChild(rowCopy4);
            }
        }

        /// <summary>
        /// Метод заполняет ячейки данными
        /// </summary>
        private void EditCellsTable()
        {
            using WordprocessingDocument doc = WordprocessingDocument.Open(_template.WordReportFileDocx.PathReport, true);

            Table table = doc.MainDocumentPart.Document.Body.Elements<Table>().Last();

            int count = _repositoryArchivePointWorkload.ArrayObjects.Count;

            int numRow = 2;

            for (int i = 0; i < count; i++)
            {
                TableRow rowXmovePlus = table.Elements<TableRow>().ElementAt(numRow);
                TableCell cellXmovePlus = rowXmovePlus.Elements<TableCell>().ElementAt(2);
                Paragraph parXmovePlus = cellXmovePlus.Elements<Paragraph>().First();
                Run runXmovePlus = parXmovePlus.Elements<Run>().First();
                Text textXmovePlus = runXmovePlus.Elements<Text>().First();
                textXmovePlus.Text = @$"{Math.Round(_repositoryArchivePointWorkload.ArrayObjects[i].AxisXmovePlus, 2)}";

                TableRow rowXspeedPlus = table.Elements<TableRow>().ElementAt(numRow);
                TableCell cellXspeedPlus = rowXspeedPlus.Elements<TableCell>().ElementAt(3);
                Paragraph parXspeedPlus = cellXspeedPlus.Elements<Paragraph>().First();
                Run runXspeedPlus = parXspeedPlus.Elements<Run>().First();
                Text textXspeedPlus = runXspeedPlus.Elements<Text>().First();
                textXspeedPlus.Text = @$"{Math.Round(_repositoryArchivePointWorkload.ArrayObjects[i].AxisXspeedPlus, 2)}";

                TableRow rowXacceleratePlus = table.Elements<TableRow>().ElementAt(numRow);
                TableCell cellXacceleratedPlus = rowXacceleratePlus.Elements<TableCell>().ElementAt(4);
                Paragraph parXacceleratePlus = cellXacceleratedPlus.Elements<Paragraph>().First();
                Run runXacceleratePlus = parXacceleratePlus.Elements<Run>().First();
                Text textXacceleratePlus = runXacceleratePlus.Elements<Text>().First();
                textXacceleratePlus.Text = @$"{Math.Round(_repositoryArchivePointWorkload.ArrayObjects[i].AxisXacceleratePlus, 2)}";

                TableRow rowXmoveMinus = table.Elements<TableRow>().ElementAt(numRow);
                TableCell cellXmoveMinus = rowXmoveMinus.Elements<TableCell>().ElementAt(5);
                Paragraph parXmoveMinus = cellXmoveMinus.Elements<Paragraph>().First();
                Run runXmoveMinus = parXmoveMinus.Elements<Run>().First();
                Text textXmoveMinus = runXmoveMinus.Elements<Text>().First();
                textXmoveMinus.Text = @$"{Math.Round(_repositoryArchivePointWorkload.ArrayObjects[i].AxisXmoveMinus, 2)}";

                TableRow rowXspeedMinus = table.Elements<TableRow>().ElementAt(numRow);
                TableCell cellXspeedMinus = rowXspeedMinus.Elements<TableCell>().ElementAt(6);
                Paragraph parXspeedMinus = cellXspeedMinus.Elements<Paragraph>().First();
                Run runXspeedMinus = parXspeedMinus.Elements<Run>().First();
                Text textXspeedMinus = runXspeedMinus.Elements<Text>().First();
                textXspeedMinus.Text = @$"{Math.Round(_repositoryArchivePointWorkload.ArrayObjects[i].AxisXspeedMinus, 2)}";

                TableRow rowXaccelerateMinus = table.Elements<TableRow>().ElementAt(numRow);
                TableCell cellXacceleratedMinus = rowXaccelerateMinus.Elements<TableCell>().ElementAt(7);
                Paragraph parXaccelerateMinus = cellXacceleratedMinus.Elements<Paragraph>().First();
                Run runXaccelerateMinus = parXaccelerateMinus.Elements<Run>().First();
                Text textXaccelerateMinus = runXaccelerateMinus.Elements<Text>().First();
                textXaccelerateMinus.Text = @$"{Math.Round(_repositoryArchivePointWorkload.ArrayObjects[i].AxisXaccelerateMinus, 2)}";


                TableRow rowYmovePlus = table.Elements<TableRow>().ElementAt(numRow + 1);
                TableCell cellYmovePlus = rowYmovePlus.Elements<TableCell>().ElementAt(2);
                Paragraph parYmovePlus = cellYmovePlus.Elements<Paragraph>().First();
                Run runYmovePlus = parYmovePlus.Elements<Run>().First();
                Text textYmovePlus = runYmovePlus.Elements<Text>().First();
                textYmovePlus.Text = @$"{Math.Round(_repositoryArchivePointWorkload.ArrayObjects[i].AxisYmovePlus, 2)}";

                TableRow rowYspeedPlus = table.Elements<TableRow>().ElementAt(numRow + 1);
                TableCell cellYspeedPlus = rowYspeedPlus.Elements<TableCell>().ElementAt(3);
                Paragraph parYspeedPlus = cellYspeedPlus.Elements<Paragraph>().First();
                Run runYspeedPlus = parYspeedPlus.Elements<Run>().First();
                Text textYspeedPlus = runYspeedPlus.Elements<Text>().First();
                textYspeedPlus.Text = @$"{Math.Round(_repositoryArchivePointWorkload.ArrayObjects[i].AxisYspeedPlus, 2)}";

                TableRow rowYacceleratePlus = table.Elements<TableRow>().ElementAt(numRow + 1);
                TableCell cellYacceleratedPlus = rowYacceleratePlus.Elements<TableCell>().ElementAt(4);
                Paragraph parYacceleratePlus = cellYacceleratedPlus.Elements<Paragraph>().First();
                Run runYacceleratePlus = parYacceleratePlus.Elements<Run>().First();
                Text textYacceleratePlus = runYacceleratePlus.Elements<Text>().First();
                textYacceleratePlus.Text = @$"{Math.Round(_repositoryArchivePointWorkload.ArrayObjects[i].AxisYacceleratePlus, 2)}";

                TableRow rowYmoveMinus = table.Elements<TableRow>().ElementAt(numRow + 1);
                TableCell cellYmoveMinus = rowYmoveMinus.Elements<TableCell>().ElementAt(5);
                Paragraph parYmoveMinus = cellYmoveMinus.Elements<Paragraph>().First();
                Run runYmoveMinus = parYmoveMinus.Elements<Run>().First();
                Text textYmoveMinus = runYmoveMinus.Elements<Text>().First();
                textYmoveMinus.Text = @$"{Math.Round(_repositoryArchivePointWorkload.ArrayObjects[i].AxisYmoveMinus, 2)}";

                TableRow rowYspeedMinus = table.Elements<TableRow>().ElementAt(numRow + 1);
                TableCell cellYspeedMinus = rowYspeedMinus.Elements<TableCell>().ElementAt(6);
                Paragraph parYspeedMinus = cellYspeedMinus.Elements<Paragraph>().First();
                Run runYspeedMinus = parYspeedMinus.Elements<Run>().First();
                Text textYspeedMinus = runYspeedMinus.Elements<Text>().First();
                textYspeedMinus.Text = @$"{Math.Round(_repositoryArchivePointWorkload.ArrayObjects[i].AxisYspeedMinus, 2)}";

                TableRow rowYaccelerateMinus = table.Elements<TableRow>().ElementAt(numRow + 1);
                TableCell cellYacceleratedMinus = rowYaccelerateMinus.Elements<TableCell>().ElementAt(7);
                Paragraph parYaccelerateMinus = cellYacceleratedMinus.Elements<Paragraph>().First();
                Run runYaccelerateMinus = parYaccelerateMinus.Elements<Run>().First();
                Text textYaccelerateMinus = runYaccelerateMinus.Elements<Text>().First();
                textYaccelerateMinus.Text = @$"{Math.Round(_repositoryArchivePointWorkload.ArrayObjects[i].AxisYaccelerateMinus, 2)}";


                TableRow rowZmovePlus = table.Elements<TableRow>().ElementAt(numRow + 2);
                TableCell cellZmovePlus = rowZmovePlus.Elements<TableCell>().ElementAt(2);
                Paragraph parZmovePlus = cellZmovePlus.Elements<Paragraph>().First();
                Run runZmovePlus = parZmovePlus.Elements<Run>().First();
                Text textZmovePlus = runZmovePlus.Elements<Text>().First();
                textZmovePlus.Text = @$"{Math.Round(_repositoryArchivePointWorkload.ArrayObjects[i].AxisZmovePlus, 2)}";

                TableRow rowZspeedPlus = table.Elements<TableRow>().ElementAt(numRow + 2);
                TableCell cellZspeedPlus = rowZspeedPlus.Elements<TableCell>().ElementAt(3);
                Paragraph parZspeedPlus = cellZspeedPlus.Elements<Paragraph>().First();
                Run runZspeedPlus = parZspeedPlus.Elements<Run>().First();
                Text textZspeedPlus = runZspeedPlus.Elements<Text>().First();
                textZspeedPlus.Text = @$"{Math.Round(_repositoryArchivePointWorkload.ArrayObjects[i].AxisZspeedPlus, 2)}";

                TableRow rowZacceleratePlus = table.Elements<TableRow>().ElementAt(numRow + 2);
                TableCell cellZacceleratedPlus = rowZacceleratePlus.Elements<TableCell>().ElementAt(4);
                Paragraph parZacceleratePlus = cellZacceleratedPlus.Elements<Paragraph>().First();
                Run runZacceleratePlus = parZacceleratePlus.Elements<Run>().First();
                Text textZacceleratePlus = runZacceleratePlus.Elements<Text>().First();
                textZacceleratePlus.Text = @$"{Math.Round(_repositoryArchivePointWorkload.ArrayObjects[i].AxisZacceleratePlus, 2)}";

                TableRow rowZmoveMinus = table.Elements<TableRow>().ElementAt(numRow + 2);
                TableCell cellZmoveMinus = rowZmoveMinus.Elements<TableCell>().ElementAt(5);
                Paragraph parZmoveMinus = cellZmoveMinus.Elements<Paragraph>().First();
                Run runZmoveMinus = parZmoveMinus.Elements<Run>().First();
                Text textZmoveMinus = runZmoveMinus.Elements<Text>().First();
                textZmoveMinus.Text = @$"{Math.Round(_repositoryArchivePointWorkload.ArrayObjects[i].AxisZmoveMinus, 2)}";

                TableRow rowZspeedMinus = table.Elements<TableRow>().ElementAt(numRow + 2);
                TableCell cellZspeedMinus = rowZspeedMinus.Elements<TableCell>().ElementAt(6);
                Paragraph parZspeedMinus = cellZspeedMinus.Elements<Paragraph>().First();
                Run runZspeedMinus = parZspeedMinus.Elements<Run>().First();
                Text textZspeedMinus = runZspeedMinus.Elements<Text>().First();
                textZspeedMinus.Text = @$"{Math.Round(_repositoryArchivePointWorkload.ArrayObjects[i].AxisZspeedMinus, 2)}";

                TableRow rowZaccelerateMinus = table.Elements<TableRow>().ElementAt(numRow + 2);
                TableCell cellZacceleratedMinus = rowZaccelerateMinus.Elements<TableCell>().ElementAt(7);
                Paragraph parZaccelerateMinus = cellZacceleratedMinus.Elements<Paragraph>().First();
                Run runZaccelerateMinus = parZaccelerateMinus.Elements<Run>().First();
                Text textZaccelerateMinus = runZaccelerateMinus.Elements<Text>().First();
                textZaccelerateMinus.Text = @$"{Math.Round(_repositoryArchivePointWorkload.ArrayObjects[i].AxisZaccelerateMinus, 2)}";

                numRow += 3;
            }

            StatusReport.SetStatusWorkloadReady();
        }

        /// <summary>
        /// Метод добавляет наименование точек в таблице
        /// </summary>
        private void EditPointTable()
        {
            using WordprocessingDocument doc = WordprocessingDocument.Open(_template.WordReportFileDocx.PathReport, true);

            Table table = doc.MainDocumentPart.Document.Body.Elements<Table>().Last();

            int count = _repositoryArchivePointWorkload.ArrayObjects.Count;

            int num = 2;

            for (int i = 0; i < count; i++)
            {
                TableRow rowNum = table.Elements<TableRow>().ElementAt(num);
                TableCell cellNum = rowNum.Elements<TableCell>().ElementAt(0);
                Paragraph par = cellNum.Elements<Paragraph>().First();
                Run run = par.Elements<Run>().First();
                Text text = run.Elements<Text>().First();

                text.Text = @$"{_repositoryArchivePointWorkload.ArrayObjects[i].Point}";

                num += 3;
            }
        }
    }
}
