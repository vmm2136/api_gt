using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http.HttpResults;
using Microsoft.AspNetCore.Mvc;
using MimeMapping;
using MimeTypes;
using System.IO;
using System.Net;
using System.Net.Http;
using WebApplication4.Services;

[ApiController]
[Route("[controller]")]
[Authorize]
public class ValidacaoController : ControllerBase
{

    private readonly WordService _wordService;
    private readonly ExcelService _excelService;
    private readonly PowerPointService _powerPointService;
    private readonly MinioService _minioService;

    public ValidacaoController(WordService wordService, ExcelService excelService, PowerPointService powerPointService, MinioService minioService)
    {
        _wordService = wordService;
        _excelService = excelService;
        _powerPointService = powerPointService;
        _minioService = minioService;

    }

    [HttpPost("validar-e-redirecionar")]
    public async Task<IActionResult> ValidarERedirecionarArquivo([FromBody] Requisicao requisicao)
    {
        if (requisicao == null)
        {
            return BadRequest("Sem dados para continuar o processo!");
        }

        var documento = await _minioService.DownloadFileAsync(requisicao.origem.BucketName, requisicao.origem.Caminho);
        var contentType = await _minioService.GetContentType(requisicao.origem.BucketName, requisicao.origem.Caminho);
        string extensao = MimeTypeMap.GetExtension(contentType) ?? ".tmp";
        string nomeOriginal = Path.GetFileNameWithoutExtension(requisicao.origem.Caminho) + extensao;

        using (var memoryStream = new MemoryStream(documento))
        {
            var file = new FormFile(memoryStream, 0, documento.Length, "file", nomeOriginal)
            {
                Headers = new HeaderDictionary(),
                ContentType = contentType
            };

            switch (extensao)
            {
                case ".docx":
                    int resultDocx = await _wordService.ProcessWordDocument(file, requisicao);
                    memoryStream.Dispose();
                    return GetActionResultForStatusCode(resultDocx);
                case ".xlsx":
                    int resultXlsx = await _excelService.ProcessExcel(file, requisicao);
                    return GetActionResultForStatusCode(resultXlsx);
              //  case ".pptx":
              //      var resultPptx = await _powerPointService.ProcessPowerPointPresentation(file, requisicao);
              //      memoryStream.Dispose();
              //      return GetActionResultForStatusCode(resultPptx);
                default:
                    return BadRequest("Formato de arquivo não suportado.");
            }


            IActionResult GetActionResultForStatusCode(int statusCode)
            {
                switch (statusCode)
                {
                    case 200:
                        return Ok(); // Status code 200 OK
                    case 404:
                        return NotFound(); // Status code 404 Not Found
                    case 500:
                        return StatusCode(500); // Status code 500 Internal Server Error
                    default:
                        return new StatusCodeResult(statusCode); // Outros códigos de status
                }
            }

        }
    }
}

public class ArquivoInfo
{
    public string BucketName{ get; set; }
    public string Caminho { get; set; }
}

public class Requisicao
{
    public ArquivoInfo origem { get; set; }
    public IDictionary<String, String> dicionarioStrings { get; set; }
    public IDictionary<String, ArquivoInfo> dicionarioImagens { get; set; }
    public ArquivoInfo destino { get; set; }
}


