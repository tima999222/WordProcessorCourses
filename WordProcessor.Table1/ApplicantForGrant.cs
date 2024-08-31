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
                        var fileName = $"Таблицы-{d.Key}".Replace("/", "_");
                        docProcessor.Map(d.First());

                        //Counters
                        var firstItem = d.FirstOrDefault();

                        if (firstItem != null && firstItem.TrainedStudents.Any())
                        {
                            docProcessor.MapErrorCounters("GSCOUNT", firstItem.TrainedStudents.Count(s => s.LeaderId != ""));
                        }
                        else
                        {
                            docProcessor.MapErrorCounters("GSCOUNT", 0);
                        }

                        if (firstItem != null && firstItem.Events.Any())
                        {
                            docProcessor.MapErrorCounters("GECOUNT", firstItem.Events.Count(e => e.LeaderId != ""));
                        }
                        else
                        {
                            docProcessor.MapErrorCounters("GECOUNT", 0);
                        }

                        if (firstItem != null && firstItem.Startups.Any())
                        {
                            docProcessor.MapErrorCounters("S1", firstItem.Startups.Count(s => s.HasSign != ""));
                        }
                        else
                        {
                            docProcessor.MapErrorCounters("S1", 0);
                        }
                        
                        //Table 1 (Students)
                        var sErrList = new List<string>()
                        {
                            "Выявлено дублирование ID участников",
                            "Выявлено дублирование ФИО участников",
                            "Отсутствует профиль в ИС «Leader-ID»",
                            "Профиль в ИС «Leader-ID» принадлежит другому человеку",
                            "Участие обучившегося в мероприятиях АП не подтверждается в данных ИС «Leader-ID»",
                            ""
                        };

                        if (firstItem != null && firstItem.Error1.Any())
                        {
                            docProcessor.MapErrorCounters("SDCOUNT",
                                firstItem.Error1.Count(e =>
                                    e.Reason == "Выявлено дублирование ID участников" ||
                                    e.Reason == "Выявлено дублирование ФИО участников"));
                        }
                        else
                        {
                            docProcessor.MapErrorCounters("SDCOUNT", 0);
                        }

                        if (firstItem != null && firstItem.Error1.Any())
                        {
                            docProcessor.MapErrorCounters("SNCOUNT",
                                firstItem.Error1.Count(e =>
                                    e.Reason ==
                                    "Участие обучившегося в мероприятиях АП не подтверждается в данных ИС «Leader-ID»"));
                        }
                        else
                        {
                            docProcessor.MapErrorCounters("SNCOUNT", 0);
                        }

                        if (firstItem != null && firstItem.Error1.Any())
                        {
                            docProcessor.MapErrorCounters("SOCOUNT",
                                firstItem.Error1.Count(e => e.Reason == "Отсутствует профиль в ИС «Leader-ID»"));
                        }
                        else
                        {
                            docProcessor.MapErrorCounters("SOCOUNT", 0);
                        }

                        if (firstItem != null && firstItem.Error1.Any())
                        {
                            var oddErrCount = firstItem.Error1.Count(e => e.Reason.Contains(';') || !sErrList.Contains(e.Reason)
                            );

                            var errList1 = new List<string>();

                            foreach (var e in firstItem.Error1)
                            {
                                if (e.Reason.Contains(';') && !errList1.Contains("Несколько ошибок"))
                                {
                                    errList1.Add("Несколько ошибок");
                                }
                                else
                                {
                                    if (!sErrList.Contains(e.Reason))
                                    {
                                        if (!errList1.Contains(e.Reason) && !e.Reason.Contains(';'))
                                            errList1.Add(e.Reason);
                                    }
                                }
                            }

                            string oddErrors1 = string.Join(";\n", errList1);

                            docProcessor.MapErrorNames("SERRORS", oddErrors1);
                            docProcessor.MapErrorCounters("SRCOUNT", oddErrCount);
                        }
                        else
                        {
                            docProcessor.MapErrorNames("SERRORS", null);
                            docProcessor.MapErrorCounters("SRCOUNT", 0);
                        }
                        
                        //Table 2 (Events)
                        var eErrList = new List<string>()
                        {
                            "Количество участников мероприятия равно \"0\" (по данным ИС \"Leader-ID\", кол-во участников 0)",
                            "Количество участников мероприятия равно \"0\" (по данным ИС \"Leader-ID\", посетители мероприятий не входят в список участников АП)",
                            "Мероприятие дублируется (не подтверждено как уникальное)",
                            "Отсутствует регистрация мероприятия в ИС «Leader-ID»",
                            "Дата проведения мероприятия вне рамок отчетного периода",
                            "Наименование мероприятия в отчетной документации не соответствует наименованию мероприятия в ИС «Leader-ID»",
                            ""
                        };

                        if (firstItem != null && firstItem.Error2.Any())
                        {
                            docProcessor.MapErrorCounters("EDCOUNT",
                                firstItem.Error2.Count(e =>
                                    e.Reason == "Мероприятие дублируется (не подтверждено как уникальное)"));
                        }
                        else
                        {
                            docProcessor.MapErrorCounters("EDCOUNT", 0);
                        }

                        if (firstItem != null && firstItem.Error2.Any())
                        {
                            docProcessor.MapErrorCounters("EOCOUNT",
                                firstItem.Error2.Count(e =>
                                    e.Reason == "Отсутствует регистрация мероприятия в ИС «Leader-ID»"));
                        }
                        else
                        {
                            docProcessor.MapErrorCounters("EOCOUNT", 0);
                        }

                        if (firstItem != null && firstItem.Error2.Any())
                        {
                            docProcessor.MapErrorCounters("EZCOUNT",
                                firstItem.Error2.Count(e => e.Reason == "Количество участников мероприятия равно \"0\" (по данным ИС \"Leader-ID\", кол-во участников 0)" || e.Reason == "Количество участников мероприятия равно \"0\" (по данным ИС \"Leader-ID\", посетители мероприятий не входят в список участников АП)"));
                        }
                        else
                        {
                            docProcessor.MapErrorCounters("EZCOUNT", 0);
                        }
                        
                        if (firstItem != null && firstItem.Error2.Any())
                        {
                            docProcessor.MapErrorCounters("ENCOUNT",
                                firstItem.Error2.Count(e => e.Reason == "Наименование мероприятия в отчетной документации не соответствует наименованию мероприятия в ИС «Leader-ID»"));
                        }
                        else
                        {
                            docProcessor.MapErrorCounters("ENCOUNT", 0);
                        }
                        
                        if (firstItem != null && firstItem.Error2.Any())
                        {
                            docProcessor.MapErrorCounters("EPCOUNT",
                                firstItem.Error2.Count(e => e.Reason == "Дата проведения мероприятия вне рамок отчетного периода"));
                        }
                        else
                        {
                            docProcessor.MapErrorCounters("EPCOUNT", 0);
                        }

                        if (firstItem != null && firstItem.Error2.Any())
                        {
                            var oddErrCount = firstItem.Error2.Count(e => e.Reason.Contains(';') || !eErrList.Contains(e.Reason));

                            var errList2 = new List<string>();

                            foreach (var e in firstItem.Error2)
                            {
                                if (e.Reason.Contains(';') && !errList2.Contains("Несколько ошибок"))
                                {
                                    errList2.Add("Несколько ошибок");
                                }
                                else
                                {
                                    if (!eErrList.Contains(e.Reason) && !e.Reason.Contains(';'))
                                    {
                                        if (!errList2.Contains(e.Reason))
                                            errList2.Add(e.Reason);
                                    }
                                }
                            }

                            string oddErrors2 = string.Join(";\n", errList2);

                            docProcessor.MapErrorNames("EERRORS", oddErrors2);
                            docProcessor.MapErrorCounters("ERCOUNT", oddErrCount);
                        }
                        else
                        {
                            docProcessor.MapErrorNames("EERRORS", null);
                            docProcessor.MapErrorCounters("ERCOUNT", 0);
                        }
                        
                        //Table 3 (Startups)
                        var ssErrList = new List<string>()
                        {
                            "Отсутствие в составе команды стартап-проекта обучившихся по соответствующей АП",
                            "Отсутствие регистрации стартап-проекта в ИС «Projects»",
                            "Выявлено дублирование стартап-проектов",
                            ""
                        };

                        if (firstItem != null && firstItem.Error3.Any())
                        {
                            docProcessor.MapErrorCounters("X",
                                firstItem.Error3.Count(e =>
                                    e.Reason == "Отсутствие в составе команды стартап-проекта обучившихся по соответствующей АП"));
                        }
                        else
                        {
                            docProcessor.MapErrorCounters("X", 0);
                        }

                        if (firstItem != null && firstItem.Error3.Any())
                        {
                            docProcessor.MapErrorCounters("Q",
                                firstItem.Error3.Count(e =>
                                    e.Reason == "Отсутствие регистрации стартап-проекта в ИС «Projects»"));
                        }
                        else
                        {
                            docProcessor.MapErrorCounters("Q", 0);
                        }

                        if (firstItem != null && firstItem.Error3.Any())
                        {
                            docProcessor.MapErrorCounters("~",
                                firstItem.Error3.Count(e => e.Reason == "Выявлено дублирование стартап-проектов"));
                        }
                        else
                        {
                            docProcessor.MapErrorCounters("~", 0);
                        }

                        if (firstItem != null && firstItem.Error3.Any())
                        {
                            var oddErrCount = firstItem.Error3.Count(e => e.Reason.Contains(';') || !ssErrList.Contains(e.Reason));

                            var errList3 = new List<string>();

                            foreach (var e in firstItem.Error3)
                            {
                                if (e.Reason.Contains(';') && !errList3.Contains("Несколько ошибок"))
                                {
                                    errList3.Add("Несколько ошибок");
                                }
                                else
                                {
                                    if (!ssErrList.Contains(e.Reason) && !e.Reason.Contains(';'))
                                    {
                                        if (!errList3.Contains(e.Reason))
                                            errList3.Add(e.Reason);
                                    }
                                }
                            }

                            string oddErrors3 = string.Join(";\n", errList3);

                            docProcessor.MapErrorNames("|", oddErrors3);
                            docProcessor.MapErrorCounters("‘", oddErrCount);
                        }
                        else
                        {
                            docProcessor.MapErrorNames("|", null);
                            docProcessor.MapErrorCounters("‘", 0);
                        }


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