using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;


namespace WebApplication4.Services
{
    public class WordService
    {
        private readonly MinioService _minioService;

        public WordService(MinioService minioService)
        {
            _minioService = minioService;
        }

        public async Task<int> ProcessWordDocument([FromForm] IFormFile file, Requisicao requisicao)
        {
            int statusCode = 0;

            if (file == null || file.Length == 0)
            {
                return 500;
            }

            string _bucketName = requisicao.destino.BucketName;
            string nomeArquivo = requisicao.destino.Caminho;

            try
            {
                using (var memoryStream = new MemoryStream())
                {
                    await file.CopyToAsync(memoryStream);
                    memoryStream.Position = 0;

                    using (var doc = WordprocessingDocument.Open(memoryStream, true))
                    {
                        IDictionary<string, string> wordReplacements = requisicao.dicionarioStrings;
                        IDictionary<string, ArquivoInfo> originalImageReplacements = requisicao.dicionarioImagens;

                        IDictionary<string, byte[]> imageReplacements = new Dictionary<string, byte[]>();

                        foreach (var entry in originalImageReplacements)
                        {
                            var arquivoInfo = entry.Value;

                            var imagem = await _minioService.DownloadFileAsync(arquivoInfo.BucketName, arquivoInfo.Caminho);

                            imageReplacements.Add(entry.Key, imagem);
                        }

                        ReplaceWordsAndImagesInElement(doc.MainDocumentPart.Document.Body, wordReplacements, imageReplacements, doc);

                        // Substituir palavras nos cabeçalhos e rodapés
                        ReplaceWordsInHeadersAndFooters(doc, wordReplacements, imageReplacements, doc);

                        // Salvar as alterações
                        doc.Save();
                    }

                    // Enviar o documento modificado para o serviço de armazenamento
                    memoryStream.Position = 0;
                    await _minioService.UploadFileAsync(_bucketName, nomeArquivo, memoryStream, memoryStream.Length);
                }

                return 200;
            }
            catch (Exception ex)
            {
                return 500;
            }
        }

        private void ReplaceWordsAndImagesInElement(OpenXmlElement element, IDictionary<string, string> wordReplacements, IDictionary<string, byte[]> imageReplacements, WordprocessingDocument doc)
        {
            if (element == null) return;

            // Iterar por todos os elementos no elemento pai
            foreach (var childElement in element.Elements())
            {
                if (childElement is DocumentFormat.OpenXml.Wordprocessing.Paragraph)
                {
                    // Se for um parágrafo, substituir palavras e imagens no texto do parágrafo
                    var paragraph = (DocumentFormat.OpenXml.Wordprocessing.Paragraph)childElement;
                    var texts = paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().ToList();
                    var fullText = string.Join("", texts.Select(t => t.Text));


                    // Substituir imagens
                    foreach (var kvp in imageReplacements)
                    {
                        var imageKey = kvp.Key;
                        var imageData = kvp.Value;

                        // Verificar se a palavra-chave está presente no texto do parágrafo
                        if (fullText.Contains(imageKey))
                        {
                            fullText = image;
                        }
                    }
                    // Substituir palavras
                    foreach (var kvp in wordReplacements)
                    {
                        if (fullText.Contains(kvp.Key))
                        {
                            fullText = fullText.Replace(kvp.Key, kvp.Value);
                        }
                    }

                    int startIndex = 0;
                    foreach (var text in texts)
                    {
                        int length = text.Text.Length;
                        if (startIndex < fullText.Length)
                        {
                            text.Text = fullText.Substring(startIndex, Math.Min(length, fullText.Length - startIndex));
                            startIndex += length;
                        }
                        else
                        {
                            text.Text = string.Empty;
                        }
                    }
                }
                else if (childElement is DocumentFormat.OpenXml.Wordprocessing.Table)
                {
                    // Se for uma tabela, substituir palavras e imagens em todas as células da tabela
                    var table = (DocumentFormat.OpenXml.Wordprocessing.Table)childElement;
                    foreach (var row in table.Elements<DocumentFormat.OpenXml.Wordprocessing.TableRow>())
                    {
                        foreach (var cell in row.Elements<DocumentFormat.OpenXml.Wordprocessing.TableCell>())
                        {
                            ReplaceWordsAndImagesInElement(cell, wordReplacements, imageReplacements, doc);
                        }
                    }
                }
                else if (childElement is SdtElement)
                {
                    // Se for um elemento de controle de conteúdo (como um caixa de texto), substituir palavras e imagens dentro dele
                    ReplaceWordsAndImagesInElement(childElement, wordReplacements, imageReplacements, doc);
                }
            }
        }

        private void ReplaceWordsInHeadersAndFooters(WordprocessingDocument doc, IDictionary<string, string> wordReplacements, IDictionary<string, byte[]> imageReplacements, WordprocessingDocument documento)
        {
            var sectionProps = doc.MainDocumentPart.Document.Descendants<SectionProperties>().ToList();

            foreach (var section in sectionProps)
            {
                foreach (var headerRef in section.Descendants<HeaderReference>())
                {
                    var headerPart = (HeaderPart)doc.MainDocumentPart.GetPartById(headerRef.Id);
                    ReplaceWordsAndImagesInElement(headerPart.Header, wordReplacements, imageReplacements, documento);
                }

                foreach (var footerRef in section.Descendants<FooterReference>())
                {
                    var footerPart = (FooterPart)doc.MainDocumentPart.GetPartById(footerRef.Id);
                    ReplaceWordsAndImagesInElement(footerPart.Footer, wordReplacements, imageReplacements, documento);
                }
            }
        }
    }
}


