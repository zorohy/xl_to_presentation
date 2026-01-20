using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;
using SixLabors.ImageSharp.Processing;
using SixLabors.ImageSharp.Drawing.Processing;
using SixLabors.Fonts;

// Open XML SDK namespaces
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using P = DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;

// Using aliases to resolve ambiguity
using ImageSharpColor = SixLabors.ImageSharp.Color;
using ImageSharpFont = SixLabors.Fonts.Font;

namespace ExcelToPptConverter
{
    class Program
    {
        // --- Configuration ---
        private const int RowsPerSlide = 30;
        private const string ExcelPath = "C:\\vcu\\VCU\\Work\\Personal\\Personal\\Projects\\xl_to_presentation\\xl_to_presentation\\output\\data.xlsx";
        private const string OutputPptxPath = "C:\\vcu\\VCU\\Work\\Personal\\Personal\\Projects\\xl_to_presentation\\xl_to_presentation\\output\\Presentation.pptx";

        // standard 16:9 Slide Dimensions in EMUs (English Metric Units)
        // 1 inch = 914400 EMUs. 10 x 5.625 inches is standard 16:9.
        private const long SlideWidthEmu = 9144000;
        private const long SlideHeightEmu = 5143500;

        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("Reading Excel file...");
                if (!File.Exists(ExcelPath))
                {
                    Console.WriteLine($"Error: {ExcelPath} not found.");
                    return;
                }

                var data = ReadExcel(ExcelPath);

                if (data.Count == 0)
                {
                    Console.WriteLine("No data found in Excel.");
                    return;
                }

                var headers = data[0].ToList();
                var rows = data.Skip(1).ToList();

                Console.WriteLine($"Found {headers.Count} columns and {rows.Count} data rows.");
                var chunks = rows.Chunk(RowsPerSlide).ToList();
                var imagePaths = new List<string>();

                for (int i = 0; i < chunks.Count; i++)
                {
                    string imgPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"table_part_{i}.png");
                    GenerateTableImage(headers, chunks[i], imgPath);
                    imagePaths.Add(imgPath);
                }

                Console.WriteLine("Generating PowerPoint presentation using Open XML SDK (Free/No Watermark)...");
                CreatePowerPoint(imagePaths, OutputPptxPath);

                // Cleanup temporary files
                foreach (var path in imagePaths)
                {
                    if (File.Exists(path)) File.Delete(path);
                }

                Console.WriteLine($"Success! Presentation saved to: {System.IO.Path.GetFullPath(OutputPptxPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Critical Error: {ex.Message}");
                if (ex.InnerException != null) Console.WriteLine($"Inner Exception: {ex.InnerException.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }

        static List<string[]> ReadExcel(string path)
        {
            var result = new List<string[]>();
            using var workbook = new XLWorkbook(path);
            var worksheet = workbook.Worksheets.First();
            var range = worksheet.RangeUsed();

            foreach (var row in range.Rows())
            {
                var rowData = row.Cells().Select(c => c.Value.ToString()).ToArray();
                result.Add(rowData);
            }
            return result;
        }

        static void GenerateTableImage(List<string> headers, string[][] rowChunk, string outputPath)
        {
            const int cellPadding = 15;
            const int fontSize = 20;
            const int headerHeight = 60;
            const int rowHeight = 45;

            if (!SystemFonts.Collection.TryGet("Arial", out var fontFamily))
            {
                fontFamily = SystemFonts.Collection.Families.FirstOrDefault();
                if (fontFamily == default) throw new Exception("No system fonts found.");
            }

            ImageSharpFont font = fontFamily.CreateFont(fontSize, SixLabors.Fonts.FontStyle.Regular);
            ImageSharpFont boldFont = fontFamily.CreateFont(fontSize, SixLabors.Fonts.FontStyle.Bold);

            int colCount = headers.Count;
            int rowCount = rowChunk.Length;

            int[] colWidths = new int[colCount];
            for (int c = 0; c < colCount; c++)
            {
                var measureOptions = new TextOptions(font);
                FontRectangle headerSize = TextMeasurer.MeasureSize(headers[c], measureOptions);
                float maxW = headerSize.Width;

                foreach (var r in rowChunk)
                {
                    string cellValue = (c < r.Length) ? r[c] : "";
                    FontRectangle rowSize = TextMeasurer.MeasureSize(cellValue, measureOptions);
                    if (rowSize.Width > maxW) maxW = rowSize.Width;
                }
                colWidths[c] = (int)maxW + (cellPadding * 2);
            }

            int totalTableWidth = colWidths.Sum();
            int totalTableHeight = headerHeight + (rowCount * rowHeight);

            using (Image<Rgba32> image = new Image<Rgba32>(totalTableWidth, totalTableHeight))
            {
                image.Mutate(ctx =>
                {
                    ctx.Fill(ImageSharpColor.White);
                    ctx.Fill(ImageSharpColor.FromRgb(230, 230, 230), new RectangleF(0, 0, totalTableWidth, headerHeight));

                    int currentY = 0;
                    int currentX = 0;

                    for (int i = 0; i < colCount; i++)
                    {
                        var rect = new RectangleF(currentX, currentY, colWidths[i], headerHeight);
                        ctx.Draw(ImageSharpColor.Black, 1, rect);

                        var richOptions = new RichTextOptions(boldFont)
                        {
                            Origin = new PointF(currentX + cellPadding, currentY + (headerHeight / 4))
                        };
                        ctx.DrawText(richOptions, headers[i], ImageSharpColor.Black);

                        currentX += colWidths[i];
                    }

                    currentY += headerHeight;
                    foreach (var row in rowChunk)
                    {
                        currentX = 0;
                        for (int i = 0; i < colCount; i++)
                        {
                            var rect = new RectangleF(currentX, currentY, colWidths[i], rowHeight);
                            ctx.Draw(ImageSharpColor.Black, 1, rect);

                            string val = i < row.Length ? row[i] : "";
                            var richOptions = new RichTextOptions(font)
                            {
                                Origin = new PointF(currentX + cellPadding, currentY + (rowHeight / 4))
                            };
                            ctx.DrawText(richOptions, val, ImageSharpColor.Black);

                            currentX += colWidths[i];
                        }
                        currentY += rowHeight;
                    }
                });

                image.Save(outputPath);
            }
        }

        static void CreatePowerPoint(List<string> imagePaths, string outputPath)
        {
            using (PresentationDocument presentationDocument = PresentationDocument.Create(outputPath, PresentationDocumentType.Presentation))
            {
                // Create Presentation Part
                PresentationPart presentationPart = presentationDocument.AddPresentationPart();
                presentationPart.Presentation = new Presentation();

                // 1. Setup Slide Master and Layout Parts
                SlideMasterPart slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
                SlideLayoutPart slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();

                // 2. Define Slide Master (Essential Structure)
                slideMasterPart.SlideMaster = new SlideMaster(
                    new CommonSlideData(new ShapeTree(
                        new P.NonVisualGroupShapeProperties(new P.NonVisualDrawingProperties() { Id = 1, Name = "" }, new P.NonVisualGroupShapeDrawingProperties(), new P.ApplicationNonVisualDrawingProperties()),
                        new GroupShapeProperties(new D.TransformGroup()),
                        new P.Shape(new P.NonVisualShapeProperties(new P.NonVisualDrawingProperties() { Id = 2, Name = "Title" }, new P.NonVisualShapeDrawingProperties(), new P.ApplicationNonVisualDrawingProperties()), new P.ShapeProperties(), new P.TextBody(new D.BodyProperties(), new D.ListStyle(), new D.Paragraph()))
                    )),
                    new ColorMap() { Background1 = D.ColorSchemeIndexValues.Light1, Text1 = D.ColorSchemeIndexValues.Dark1, Background2 = D.ColorSchemeIndexValues.Light2, Text2 = D.ColorSchemeIndexValues.Dark2, Accent1 = D.ColorSchemeIndexValues.Accent1, Accent2 = D.ColorSchemeIndexValues.Accent2, Accent3 = D.ColorSchemeIndexValues.Accent3, Accent4 = D.ColorSchemeIndexValues.Accent4, Accent5 = D.ColorSchemeIndexValues.Accent5, Accent6 = D.ColorSchemeIndexValues.Accent6, Hyperlink = D.ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = D.ColorSchemeIndexValues.FollowedHyperlink },
                    new SlideLayoutIdList(new SlideLayoutId() { Id = 2147483648U, RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart) })
                );

                // 3. Define Slide Layout (Inherit Master)
                slideLayoutPart.SlideLayout = new SlideLayout(new CommonSlideData(new ShapeTree(
                    new P.NonVisualGroupShapeProperties(new P.NonVisualDrawingProperties() { Id = 1, Name = "" }, new P.NonVisualGroupShapeDrawingProperties(), new P.ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new D.TransformGroup())
                )));
                slideLayoutPart.AddPart(slideMasterPart);

                // 4. Initialize Presentation Properties in CORRECT XML Order
                presentationPart.Presentation.AppendChild(new SlideMasterIdList(new SlideMasterId() { Id = 2147483648U, RelationshipId = presentationPart.GetIdOfPart(slideMasterPart) }));
                presentationPart.Presentation.AppendChild(new SlideIdList());
                presentationPart.Presentation.AppendChild(new SlideSize() { Cx = (int)SlideWidthEmu, Cy = (int)SlideHeightEmu, Type = SlideSizeValues.Screen16x9 });
                presentationPart.Presentation.AppendChild(new NotesSize() { Cx = 6858000, Cy = 9144000 });
                presentationPart.Presentation.AppendChild(new DefaultTextStyle());

                SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;
                uint slideIdCounter = 256;

                foreach (var imgPath in imagePaths)
                {
                    // Create Slide Part
                    SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();
                    slidePart.Slide = new Slide(new CommonSlideData(new ShapeTree(
                        new P.NonVisualGroupShapeProperties(new P.NonVisualDrawingProperties() { Id = 1, Name = "" }, new P.NonVisualGroupShapeDrawingProperties(), new P.ApplicationNonVisualDrawingProperties()),
                        new GroupShapeProperties(new D.TransformGroup())
                    )));

                    // EVERY Slide must link to a Layout Part
                    slidePart.AddPart(slideLayoutPart);

                    // Image Insertion
                    ImagePart imagePart = slidePart.AddImagePart(ImagePartType.Png);
                    using (FileStream stream = new FileStream(imgPath, FileMode.Open, FileAccess.Read))
                    {
                        imagePart.FeedData(stream);
                    }

                    using (var image = Image.Load(imgPath))
                    {
                        double imgW = image.Width;
                        double imgH = image.Height;
                        double margin = 400000;
                        double availW = SlideWidthEmu - (margin * 2);
                        double availH = SlideHeightEmu - (margin * 2);

                        double scale = 1.0;
                        if ((imgW * 9525) > availW) scale = availW / (imgW * 9525);
                        if ((imgH * 9525 * scale) > availH) scale = availH / (imgH * 9525);

                        long finalW = (long)(imgW * 9525 * scale);
                        long finalH = (long)(imgH * 9525 * scale);
                        long x = (long)((SlideWidthEmu - finalW) / 2);
                        long y = (long)((SlideHeightEmu - finalH) / 2);

                        AddImageToShapeTree(slidePart.Slide.CommonSlideData.ShapeTree, slidePart.GetIdOfPart(imagePart), x, y, finalW, finalH);
                    }

                    // Register slide in Presentation part
                    string relId = presentationPart.GetIdOfPart(slidePart);
                    slideIdList.Append(new SlideId() { Id = slideIdCounter++, RelationshipId = relId });
                }

                // Final save
                presentationPart.Presentation.Save();
            }
        }

        private static void AddImageToShapeTree(ShapeTree shapeTree, string relationshipId, long x, long y, long cx, long cy)
        {
            var picture = new P.Picture();

            picture.NonVisualPictureProperties = new P.NonVisualPictureProperties(
                new P.NonVisualDrawingProperties() { Id = (uint)(new Random().Next(10, 1000)), Name = "TableImage" },
                new P.NonVisualPictureDrawingProperties(new D.PictureLocks() { NoChangeAspect = true }),
                new P.ApplicationNonVisualDrawingProperties());

            picture.BlipFill = new P.BlipFill(
                new D.Blip() { Embed = relationshipId, CompressionState = D.BlipCompressionValues.Print },
                new D.Stretch(new D.FillRectangle()));

            picture.ShapeProperties = new P.ShapeProperties(
                new D.Transform2D(
                    new D.Offset() { X = x, Y = y },
                    new D.Extents() { Cx = cx, Cy = cy }),
                new D.PresetGeometry() { Preset = D.ShapeTypeValues.Rectangle });

            shapeTree.Append(picture);
        }
    }

    public static class LinqExtensions
    {
        public static IEnumerable<T[]> Chunk<T>(this IEnumerable<T> source, int size)
        {
            var list = source.ToList();
            for (int i = 0; i < list.Count; i += size)
            {
                yield return list.GetRange(i, Math.Min(size, list.Count - i)).ToArray();
            }
        }
    }
}