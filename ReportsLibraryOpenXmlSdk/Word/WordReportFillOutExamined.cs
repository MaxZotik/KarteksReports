using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using DFW = DocumentFormat.OpenXml.Wordprocessing;
using ReportsLibraryOpenXmlSdk.Entities;

namespace ReportsLibraryOpenXmlSdk.Word
{
    public class WordReportFillOutExamined
    {
        private Template _template;

        public WordReportFillOutExamined(Template template)
        {
            _template = template;
            InsertTextBookmarks();
            ReplaceBookmarksWithImage();
        }

        private void InsertTextBookmarks()
        {
            using var docWord = WordprocessingDocument.Open(_template.WordReportFileDocx.PathReport, true);

            foreach (BookmarkStart book in docWord.MainDocumentPart.Document.Body.Descendants<BookmarkStart>())
            {
                foreach (string str in Constants.BookmarksWord.BookmarksTemplateExaminedText)
                {
                    if (book.Name == str)
                    {
                        RunProperties runProperties = new RunProperties();

                        RunFonts runFonts = new RunFonts() { Ascii = "Times New Roman" };
                        FontSize fontSize = new FontSize() { Val = "24" };

                        runProperties.Append(runFonts);
                        runProperties.Append(fontSize);


                        var text = new DFW.Text(GetTextBookmark(str));
                        var runElement = new DFW.Run(runProperties, text);
                        book.InsertAfterSelf(runElement);
                    }
                }
            }

            docWord.Save();
        }

        private string GetTextBookmark(string text)
        {
            string str = " ";

            switch (text)
            {
                case "Редуктор_название_заголовок":
                    str = _template.ReductorNameTitle;
                    break;
                case "Дата_и_время_заголовок":
                    str = _template.DateAndTimeHeading;
                    break;
                case "Редуктор_название_объект_испытаний":
                    str = _template.ReductorNameOfTheTestObject;
                    break;
                case "ФИО_Исполнители":
                    str = _template.FioPerformers;
                    break;
                case "Редуктор_название_измерения_вибрации":
                    str = _template.ReductorNumberOfVibrationMeasurements;
                    break;
                case "Дата_и_время_измерения_вибрации":
                    str = _template.DateAndTimeVibrationMeasurement;
                    break;
                case "Частота_вращения_во_время_испытаний":
                    str = _template.FrequencyOfRotationDuringTesting;
                    break;
                case "Редуктор_название_критерии_оценки":
                    str = _template.ReductorNumberOfAssessmentCriteria;
                    break;
                case "ФИО_Заключение":
                    str = _template.FioConclusion;
                    break;
                default:
                    str = " ";
                    break;
            }

            return str;
        }

        private void ReplaceBookmarksWithImage()
        {
            using WordprocessingDocument doc = WordprocessingDocument.Open(_template.WordReportFileDocx.PathReport, true);

            // Read all bookmarks from the word doc
            foreach (BookmarkStart bookmarkStart in doc.MainDocumentPart.RootElement.Descendants<BookmarkStart>())
            {
                foreach (string str in Constants.BookmarksWord.BookmarksTemplateExaminedPng)
                {
                    if (bookmarkStart.Name == str)
                    {
                        InsertImageIntoBookmark(doc, bookmarkStart);
                        // remove the bookmark
                        bookmarkStart.Remove();
                    }
                        
                }                                   
            }

            doc.Save();
        }

        public void InsertImageIntoBookmark(WordprocessingDocument doc, BookmarkStart bookmarkStart)
        {
            // Remove anything present inside the bookmark
            OpenXmlElement elem = bookmarkStart.NextSibling();
            while (elem != null && !(elem is BookmarkEnd))
            {
                OpenXmlElement nextElem = elem.NextSibling();
                elem.Remove();
                elem = nextElem;
            }

            // Create an imagepart
            var imagePart = AddImagePart(doc.MainDocumentPart, _template.SchemaReductor);

            // insert the image part after the bookmark start
            AddImageToBody(doc.MainDocumentPart.GetIdOfPart(imagePart), bookmarkStart);
        }

        public ImagePart AddImagePart(MainDocumentPart mainPart, string imageFilename)
        {
            ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);

            using (FileStream stream = new FileStream(imageFilename, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }

            return imagePart;
        }

        private void AddImageToBody(string relationshipId, BookmarkStart bookmarkStart)
        {
            SizePng png = new(_template.SchemaReductor);

            var element =
                new Drawing(
                    new DW.Inline(
                        new DW.Extent() { Cx = png.Width, Cy = png.Height },
                        new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                        new DW.DocProperties() { Id = (UInt32Value)1U, Name = "Pictere 2" },
                        new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks() { NoChangeAspect = true }),
                        new A.Graphic(
                            new A.GraphicData(
                                new PIC.Picture(
                                    new PIC.NonVisualPictureProperties(
                                        new PIC.NonVisualDrawingProperties()
                                        {
                                            Id = (UInt32Value)0U,
                                            Name = "New.Png"
                                        },
                                        new PIC.NonVisualPictureDrawingProperties()),
                                    new PIC.BlipFill(
                                        new A.Blip(
                                            new A.BlipExtensionList(
                                                new A.BlipExtension()
                                                {
                                                    Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                }))
                                        {
                                            Embed = relationshipId,
                                            CompressionState = A.BlipCompressionValues.Print
                                        },
                                        new A.Stretch(new A.FillRectangle())),
                                    new PIC.ShapeProperties(
                                        new A.Transform2D(new A.Offset() { X = 0L, Y = 0L }, new A.Extents() { Cx = png.Width, Cy = png.Height }),
                                        new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })))
                            {
                                Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"
                            }))
                    {
                        DistanceFromTop = (UInt32Value)0U,
                        DistanceFromBottom = (UInt32Value)0U,
                        DistanceFromLeft = (UInt32Value)0U,
                        DistanceFromRight = (UInt32Value)0U,
                        EditId = "50D07946"
                    });

            bookmarkStart.Parent.InsertAfter<Run>(new Run(element), bookmarkStart);
        }

    }
}
