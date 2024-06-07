using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace WebApplication4.Services
{
    public class PowerPointService
    {
        private readonly MinioService _minioService;

        public PowerPointService(MinioService minioService)
        {
            _minioService = minioService;
        }

        public async Task<int> ProcessPowerPointPresentation([FromForm] IFormFile file, Requisicao requisicao)
        {

            int statusCode = 0;

            if (file == null || file.Length == 0)
            {
                return 500;
            }

            string _bucketName = requisicao.destino.BucketName;
            string nomeArquivo = requisicao.destino.Caminho;
            string saveFilePath;

            var tempDirectory = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
            Directory.CreateDirectory(tempDirectory);
            var tempFilePath = Path.Combine(tempDirectory, file.FileName);

            Application powerPointApp = null;
            Presentation presentation = null;

            try
            {
                using (var fileStream = new FileStream(tempFilePath, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    await file.CopyToAsync(fileStream);
                }

                IDictionary<String, String> wordReplacements = requisicao.dicionarioStrings;

                IDictionary<string, ArquivoInfo> originalImageReplacements = requisicao.dicionarioImagens;

                IDictionary<string, byte[]> imageReplacements = new Dictionary<string, byte[]>();

                foreach (var entry in originalImageReplacements)
                {
                    var arquivoInfo = entry.Value;

                    var imagem = await _minioService.DownloadFileAsync(arquivoInfo.BucketName, arquivoInfo.Caminho);

                    imageReplacements.Add(entry.Key, imagem);
                }

                powerPointApp = new Application();
                presentation = powerPointApp.Presentations.Open(tempFilePath, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);

                foreach (Slide slide in presentation.Slides)
                {
                    foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in slide.Shapes)
                    {
                        if (shape.HasTextFrame == MsoTriState.msoTrue)
                        {
                            var textFrame = shape.TextFrame;
                            var textRange = textFrame.TextRange;

                            foreach (var pair in wordReplacements)
                            {
                                if (textRange.Text.Contains(pair.Key))
                                {
                                    textRange.Text = textRange.Text.Replace(pair.Key, pair.Value);
                                }
                            }

                            foreach (var pair in imageReplacements)
                            {
                                if (textRange.Text.IndexOf(pair.Key, StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    string tempImagePath = Path.GetTempFileName();
                                    File.WriteAllBytes(tempImagePath, pair.Value);

                                    float left = shape.Left;
                                    float top = shape.Top;
                                    float width = shape.Width;
                                    float height = shape.Height;

                                    shape.Delete();
                                    slide.Shapes.AddPicture(tempImagePath, MsoTriState.msoFalse, MsoTriState.msoCTrue, left, top, width, height);

                                    File.Delete(tempImagePath);
                                }
                            }
                        }
                    }
                }

                saveFilePath = Path.Combine(tempDirectory, "processed_" + file.FileName);
                presentation.SaveAs(saveFilePath);
                presentation.Close();
                powerPointApp.Quit();
                powerPointApp = null;

                var objectName = Path.GetFileName(saveFilePath);
                //await _minioService.UploadFileAsync(_bucketName, nomeArquivo, saveFilePath);

                return 200;
            }
            catch (Exception ex)
            {
                return 500;

                if (Directory.Exists(tempDirectory))
                {
                    Directory.Delete(tempDirectory, true);
                }
            }
            finally
            {
                if (tempFilePath != null && System.IO.File.Exists(tempFilePath))
                {
                    System.IO.File.Delete(tempFilePath);
                }
                if (powerPointApp != null)
                {
                    powerPointApp.Quit();
                }
            }
        }
    }
}

