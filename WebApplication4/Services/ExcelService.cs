using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace WebApplication4.Services
{
    public class ExcelService
    {
        private readonly MinioService _minioService;

        public ExcelService(MinioService minioService)
        {
            _minioService = minioService;
        }

        public async Task<int> ProcessExcel([FromForm] IFormFile file, Requisicao requisicao)
        {
            int statusCode = 0;

            if (file == null || file.Length == 0)
            {
                return 500;
            }

            string bucketName = requisicao.destino.BucketName;
            string nomeArquivo = requisicao.destino.Caminho;

            try
            {
                using (var memoryStream = new MemoryStream())
                {
                    await file.CopyToAsync(memoryStream);

                    IDictionary<string, string> wordReplacements = requisicao.dicionarioStrings;
                    IDictionary<string, ArquivoInfo> originalImageReplacements = requisicao.dicionarioImagens;
                    IDictionary<string, byte[]> imageReplacements = new Dictionary<string, byte[]>();

                    foreach (var entry in originalImageReplacements)
                    {
                        var arquivoInfo = entry.Value;
                        var imagem = await _minioService.DownloadFileAsync(arquivoInfo.BucketName, arquivoInfo.Caminho);
                        imageReplacements.Add(entry.Key, imagem);
                    }

                    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(memoryStream, true))
                    {
                        WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                        foreach (Sheet sheet in workbookPart.Workbook.Sheets)
                        {
                            WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                            ReplaceWordsAndImagesInWorksheet(worksheetPart, workbookPart, wordReplacements, imageReplacements);
                        }
                    }

                    memoryStream.Position = 0; // Reset stream position to the beginning
                    await _minioService.UploadFileAsync(bucketName, nomeArquivo, memoryStream, memoryStream.Length);
                }

                return 200;
            }
            catch (Exception)
            {
                return 500;
            }
        }

        private void ReplaceWordsAndImagesInWorksheet(WorksheetPart worksheetPart, WorkbookPart workbookPart, IDictionary<string, string> wordReplacements, IDictionary<string, byte[]> imageReplacements)
        {
            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

            foreach (Row row in sheetData.Elements<Row>())
            {
                foreach (Cell cell in row.Elements<Cell>())
                {
                    string cellValue = GetCellValue(cell, workbookPart);

                    if (cellValue != null)
                    {
                        foreach (var kvp in imageReplacements)
                        {
                            if (cellValue.Contains(kvp.Key))
                            {
                                string tempImagePath = Path.GetTempFileName();
                                File.WriteAllBytes(tempImagePath, kvp.Value);

                                InsertImageIntoWorksheet(worksheetPart, cell, tempImagePath);

                                File.Delete(tempImagePath);
                            }
                        }

                        foreach (var kvp in wordReplacements)
                        {
                            if (cellValue.Contains(kvp.Key))
                            {
                                cellValue = cellValue.Replace(kvp.Key, kvp.Value);
                                cell.CellValue = new CellValue(cellValue);
                                cell.DataType = new EnumValue<CellValues>(CellValues.String);
                            }
                        }
                    }
                }
            }
        }

        private string GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            if (cell.CellValue == null)
                return null;

            string value = cell.CellValue.InnerText;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(int.Parse(value)).InnerText;
            }

            return value;
        }


        private void InsertImageIntoWorksheet(WorksheetPart worksheetPart, Cell cell, string imagePath)
        {
            DrawingsPart drawingsPart;
            WorksheetDrawing worksheetDrawing;

            if (worksheetPart.DrawingsPart == null)
            {
                // Create a new DrawingsPart and WorksheetDrawing
                drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
                worksheetDrawing = new WorksheetDrawing();
                drawingsPart.WorksheetDrawing = worksheetDrawing;
                worksheetPart.Worksheet.Append(new DocumentFormat.OpenXml.Spreadsheet.Drawing() { Id = worksheetPart.GetIdOfPart(drawingsPart) });
                worksheetDrawing.Save();
            }
            else
            {
                // Use the existing DrawingsPart and WorksheetDrawing
                drawingsPart = worksheetPart.DrawingsPart;
                worksheetDrawing = drawingsPart.WorksheetDrawing;
            }

            // Add the image part to the DrawingsPart
            ImagePart imagePart = drawingsPart.AddImagePart(ImagePartType.Png);
            using (FileStream stream = new FileStream(imagePath, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }

            // Calculate position and size
            int column = GetColumnIndex(cell.CellReference);
            int row = GetRowIndex(cell.CellReference);

            // Define the anchor for the image
            uint imageId = (uint)(drawingsPart.ImageParts.Count() + 1);
            var twoCellAnchor = new DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor(
                new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker
                {
                    ColumnId = new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnId((column).ToString()),
                    RowId = new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowId((row - 1).ToString())
                },
                new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker
                {
                    ColumnId = new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnId((column + 1).ToString()),
                    RowId = new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowId(row.ToString())
                },
                new DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture(
                    new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualPictureProperties(
                        new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties()
                        {
                            Id = imageId,
                            Name = "Picture " + imageId
                        },
                        new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualPictureDrawingProperties()),
                    new DocumentFormat.OpenXml.Drawing.Spreadsheet.BlipFill(
                        new DocumentFormat.OpenXml.Drawing.Blip()
                        {
                            Embed = drawingsPart.GetIdOfPart(imagePart),
                            CompressionState = DocumentFormat.OpenXml.Drawing.BlipCompressionValues.Print
                        },
                        new DocumentFormat.OpenXml.Drawing.Stretch(
                            new DocumentFormat.OpenXml.Drawing.FillRectangle())),
                    new DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeProperties(
                        new DocumentFormat.OpenXml.Drawing.Transform2D(
                            new DocumentFormat.OpenXml.Drawing.Offset() { X = 0, Y = 0 },
                            new DocumentFormat.OpenXml.Drawing.Extents() { Cx = 990000L, Cy = 792000L }),
                        new DocumentFormat.OpenXml.Drawing.PresetGeometry(new DocumentFormat.OpenXml.Drawing.AdjustValueList())
                        { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle })));

            // Append the TwoCellAnchor to the WorksheetDrawing
            worksheetDrawing.Append(twoCellAnchor);
            worksheetDrawing.Save();
        }

        private int GetColumnIndex(string cellReference)
        {
            int columnIndex = 0;
            foreach (char ch in cellReference)
            {
                if (Char.IsLetter(ch))
                {
                    columnIndex = (columnIndex * 26) + (ch - 'A' + 1);
                }
                else
                {
                    break;
                }
            }
            return columnIndex - 1;
        }

        private int GetRowIndex(string cellReference)
        {
            string rowPart = new string(cellReference.Where(Char.IsDigit).ToArray());
            return int.Parse(rowPart);
        }
    }
}
