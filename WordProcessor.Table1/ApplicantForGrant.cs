using AStepanov.Core.Ex;
using ASTepanov.Docx;
using DocumentFormat.OpenXml.Drawing.Charts;
using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using WordProcessor.Table1.Entities;

namespace WordProcessor.Table1
{
    public static class ApplicantForGrant
    {
        /// <summary>
        /// Создает зип со всеми файлами которые замапил по шаблону
        /// </summary>
        /// <param name="list">Сгруппированные по ИНН данные для вордов</param>
        /// <returns>Массив байтов которые можно сохранить в зип архив</returns>
        public static byte[]? CreateApplicationsForOrder(List<IGrouping<string, DataForWord>> list, string zipPath)
        {
            if (!list.Any()) return null;

            byte[] docTemplate = null;
            var path = Path.Combine(Assembly.GetExecutingAssembly().Directory(), "template.docx");
            using (var s = File.OpenRead(path))
            {
                docTemplate = s.ReadFully(); // считывает шаблон
            }

            using (var fileStream = new FileStream(zipPath, FileMode.OpenOrCreate)) // Открываем или создаем архив
            {
                using (var archive =
                       new ZipArchive(fileStream, ZipArchiveMode.Update))
                {
                    foreach (var d in list.AsParallel().WithDegreeOfParallelism(8))
                    {
                        var docProcessor = new DocumentProcessor(docTemplate);
                        var fileName = $"{d.Key}".Replace("/", "_");
                        docProcessor.Map(d.First());
                        docProcessor.MapItems(d.First().TrainedStudents, 2);
                        docProcessor.MapItems(d.First().Events, 2);
                        
                        //Error Tables
                        docProcessor.MapItems(d.First().Error1, 3);
                        docProcessor.MapItems(d.First().Error2, 3);
                        docProcessor.MapItems(d.First().Error3, 3);
                        
                        docProcessor.MapItemsOther(d.First().Startups, 3);

                        var fileInArchive = archive.CreateEntry(fileName + ".docx", CompressionLevel.Optimal);
                        using (var entryStream = fileInArchive.Open())
                        using (var fileToCompressStream = new MemoryStream(docProcessor.Bytes()))
                        {
                            fileToCompressStream.CopyTo(entryStream);
                        }
                    }
                }
            }
            return File.ReadAllBytes(zipPath);
        }
    }
}