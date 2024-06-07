using Minio;
using Minio.DataModel.Args;
using Minio.Exceptions;
using System;
using System.IO;
using System.Threading.Tasks;

public class MinioService
{
    private readonly IMinioClient _minioClient;

    public MinioService()
    {
        _minioClient = new MinioClient()
            .WithEndpoint("localhost:9000")
            .WithCredentials("minio", "minio123")
            .Build();
    }

    public async Task UploadFileAsync(string bucketName, string objectName, Stream dataStream, long size)
    {
        try
        {
            bool found = await _minioClient.BucketExistsAsync(new BucketExistsArgs().WithBucket(bucketName));
            if (!found)
            {
                await _minioClient.MakeBucketAsync(new MakeBucketArgs().WithBucket(bucketName));
            }
            await _minioClient.PutObjectAsync(new PutObjectArgs()
                .WithBucket(bucketName)
                .WithObject(objectName)
                .WithStreamData(dataStream)
                .WithObjectSize(size));
            Console.WriteLine("Successfully uploaded " + objectName);
        }
        catch (MinioException e)
        {
            Console.WriteLine("File Upload Error: " + e.Message);
        }
    }

    public async Task<byte[]> DownloadFileAsync(string bucketName, string caminho)
    {
        try
        {
            using (MemoryStream ms = new MemoryStream())
            {

                string nomeArquivo = Path.GetFileName(caminho);

                await _minioClient.GetObjectAsync(new GetObjectArgs()
                    .WithBucket(bucketName)
                    .WithObject(caminho)
                    .WithCallbackStream((stream) =>
                    {
                        stream.CopyTo(ms);
                    }));
                return ms.ToArray();
            }
        }
        catch (MinioException e)
        {
            return null;
        }
    }

    public async Task<string> GetContentType(string bucketName, string objectName)
    {
        var statObjectArgs = new StatObjectArgs()
            .WithBucket(bucketName)
            .WithObject(objectName);
        var objectStat = await _minioClient.StatObjectAsync(statObjectArgs);
        return objectStat.ContentType;
    }

}
