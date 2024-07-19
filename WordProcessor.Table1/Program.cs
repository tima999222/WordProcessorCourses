using AStepanov.Core.Ex;
using System.Reflection;
using DocumentFormat.OpenXml.Math;
using Microsoft.IdentityModel.Tokens;
using WordProcessor.Table1;
using WordProcessor.Table1.Entities;
using Serilog;
using Serilog.Core;


namespace WordProcessor.Table1
{
    internal class Program
    {
        private static readonly object fileLock = new object();
        private static Logger logger = ConfigureLogger();

        static void Main()
        {
            logger.Information("Starting Up");

            IEnumerable<string> contracts =
            [
                #region Important

                //"70-2023-000670",
                //"70-2023-000622",
                //"70-2023-000669",
                //"70-2023-000620",
                //"70-2023-000672",
                //"70-2023-000677",
                //"70-2023-000678",
                //"70-2023-000671",
                //"70-2023-000621",
                //"70-2023-000673",

                //"70-2023-000626",
                //"70-2023-000644",
                //"70-2023-000645",
                //"70-2023-000650",
                //"70-2023-000651",
                //"70-2023-000655",
                //"70-2023-000662",
                //"70-2023-000663",
                //"70-2023-000664",
                //"70-2023-000667",

                //"70-2023-000665",
                //"70-2023-000680",
                //"70-2023-000681",
                //"70-2023-000682",
                //"70-2023-000683",
                //"70-2023-000684",
                //"70-2023-000685",
                //"70-2023-000691",
                //"70-2023-000692",
                //"70-2023-000693",

                //"70-2023-000668",
                //70-2023-000679",
                //"70-2023-000694",
                //"70-2023-000695",
                //"70-2023-000696",
                //"70-2023-000697",
                //"70-2023-000719",
                //"70-2023-000768",


                //"70-2023-000647",
                //"70-2023-000648"

                #endregion

                #region Not Important

                //"70-2023-000618",
                //"70-2023-000619",
                //"70-2023-000623",
                //"70-2023-000624",
                //"70-2023-000625",
                //"70-2023-000627",
                //"70-2023-000628",
                "70-2023-000629",
                //"70-2023-000633",
                //"70-2023-000634",

                //"70-2023-000636",
                //"70-2023-000637",
                //"70-2023-000638",
                //"70-2023-000641",
                //"70-2023-000642",
                //"70-2023-000643",
                //"70-2023-000646",
                //"70-2023-000661",
                //"70-2023-000686",
                //"70-2023-000687",

                //"70-2023-000688",
                //"70-2023-000689",
                //"70-2023-000690",
                //"70-2023-000698",
                //"70-2023-000700",
                //"70-2023-000701",
                //"70-2023-000702",
                //"70-2023-000703",
                //"70-2023-000704",

                //"70-2023-000706",
                //"70-2023-000707",
                //"70-2023-000708",
                //"70-2023-000710",
                //"70-2023-000711",
                //"70-2023-000712",
                //"70-2023-000716",

                //"70-2023-000720",
                //"70-2023-000722",
                //"70-2023-000723",
                //"70-2023-000726",
                "70-2023-000728",
                "70-2023-000729",
                "70-2023-000730",
                //"70-2023-000732",
                //"70-2023-000733",
                "70-2023-000734",

                //"70-2023-000735",
                //"70-2023-000737",
                //"70-2023-000738",
                //"70-2023-000741",
                //"70-2023-000742",
                //"70-2023-000743",
                //"70-2023-000744",
                //"70-2023-000746",
                //"70-2023-000747",
                //"70-2023-000748",

                //"70-2023-000621"

                //"70-2023-000752",
                //"70-2023-000753",
                //"70-2023-000757",
                //"70-2023-000758",
                //"70-2023-000761",
                //"70-2023-000762",
                "70-2023-000765",
                "70-2023-000774",
                "70-2023-000775",

                "70-2023-000736",
                "70-2023-000770",
                "70-2023-000771"

                #endregion
            ];

            /*foreach (var contractsChunk in SplitIntoChunks(contracts, 3))
            {
                Parallel.ForEach(contractsChunk, contract =>
                {
                    fillContract(contract, logger);
                    Thread.Sleep(3000);
                });
            }*/

            fillContract("70-2023-000667", logger);
        }

        static Logger ConfigureLogger()
        {
            return new LoggerConfiguration()
                .MinimumLevel.Debug()
                .WriteTo.Console()
                .CreateLogger();
        }

        static IEnumerable<IEnumerable<T>> SplitIntoChunks<T>(IEnumerable<T> source, int chunkSize)
        {
            var list = new List<T>(source);
            for (int i = 0; i < list.Count; i += chunkSize)
            {
                yield return list.GetRange(i, Math.Min(chunkSize, list.Count - i));
            }
        }

        static void fillContract(string contractNumber, Logger logger)
        {
            logger.Information("Trying to get data from database for contract with number [{contractNumber}]",
                contractNumber);
            List<DataForWord> dataFromDB = GetDataFromDatabase(contractNumber, logger);


            List<DataForWord> testData = null;

            var groupedData = new List<IGrouping<string, DataForWord>>();


            if (dataFromDB == null || !dataFromDB.Any())
            {
                logger.Information("Could not get data for contract with number [{contractNumber}]", contractNumber);
                testData = GenerateTestData(logger);
                groupedData = GroupData(testData);
            }
            else
            {
                logger.Information("Got data for contract with number [{contractNumber}]", contractNumber);
                groupedData = GroupData(dataFromDB);
            }

            if (groupedData != null && groupedData.Any())
            {
                logger.Information("Creating file(s)...");
                var zipPath = Assembly.GetExecutingAssembly().Directory() + "/documents.zip";
                lock (fileLock) // Locking the critical section
                {
                    var result = ApplicantForGrant.CreateApplicationsForOrder(groupedData, zipPath);
                    if (result != null)
                    {
                        foreach (var data in groupedData)
                            logger.Information("File [{contractNumber}] added to archive",
                                "Таблицы-" + data.Key + ".docx");
                    }
                }
            }
        }

        static List<DataForWord> GetDataFromDatabase(string contractNumber, Logger logger)
        {
            var dataList = new List<DataForWord>();
            var trainedStudents = new List<TrainedStudent>();
            var events = new List<Event>();
            var startups = new List<Startup>();

            var errors1 = new List<ErrorTable1>();
            var errors2 = new List<ErrorTable2>();
            var errors3 = new List<ErrorTable3>();


            logger.Information("Getting participants...");
            //trainedStudents = Connection.GetParticipantsForContract(contractNumber);
            trainedStudents = Connection.GetNewTable1ForContract(contractNumber);
            logger.Information("Got data about [{count}] participants for contract [{c}]", trainedStudents.Count,
                contractNumber);

            logger.Information("Getting errors in Table 1...");
            //errors1 = Connection.GetErrors1ForContract(contractNumber);
            errors1 = Connection.GetErrorsForContract(contractNumber);

            foreach (var p in errors1)
            {
                if (p.Name.IsNullOrEmpty())
                {
                    var name = Connection.GetStudentByID(p.LeaderLink.Replace("https://leader-id.ru/users/", ""),
                        contractNumber);

                    if (name != "#Н/Д")
                    {
                        p.Name = name;
                    }
                }
            }

            foreach (var e in errors1)
            {
                if (e.Reason.Contains("Выявлено дублирование"))
                {
                    e.Reason = e.Reason.Replace("Выявлено дублирование",
                        "Выявлено дублирование ID участников");
                }
                
                if (e.Reason.Contains("фамилия"))
                {
                    e.Reason = e.Reason.Replace("фамилия",
                        "Выявлено дублирование ФИО участников");
                }
            }

            foreach (var e in errors1)
            {
                var docList = new List<string>();

                if (e.Reason.Contains(
                        "Участие обучившегося в мероприятиях АП не подтверждается в данных ИС «Leader-ID»"))
                {
                    docList.Add(
                        "Необходимо предоставить документы, подтверждающие участие обучившегося в мероприятиях АП, проведение которых подтверждено. Такими документами может являться, в том числе, протокол мероприятия за подписью уполномоченного лица с поименным списком участников мероприятия, поименный список регистрации участников мероприятия с оригинальными подписями, в котором указан в том числе данный обучившийся. В дополнение к указанными документам могут прикладываться новости, подтверждающие проведение мероприятия с участием данного обучившегося (с указанием ссылки на источник и предоставлением скриншота публикации), скриншоты, подтверждающие проведение мероприятия в формате онлайн с участием данного обучившегося, фотоотчеты с мероприятия, иные документы и материалы, которые позволяют подтвердить факт участия обучившегося в мероприятиях АП.\nНаправляемые документы должны сопровождаться письмом за подписью ректора, либо уполномоченного лица с подтверждением полномочий подписанта, в котором подтверждается участие обучившегося в мероприятиях АП.");
                }

                if (e.Reason.Contains("Выявлено дублирование ID участников") || e.Reason.Contains("Выявлено дублирование ФИО участников"))
                {
                    docList.Add(
                        "Необходимо исключить дублирование в доработанной отчетной документации. Направляемый доработанный отчет о ходе реализации акселерационной программы (по форме приложения \u2116 6 к договору о предоставлении гранта) должен сопровождаться письмом за подписью ректора, либо уполномоченного лица с подтверждением полномочий подписанта, в котором даются пояснения по выявленному замечанию, с приложением дополнительных подтверждающих материалов, в случае, если дублирование отсутствует, либо подтверждается исключение дублирования из отчетных материалов.");
                }

                if (e.Reason.Contains("Отсутствует профиль в ИС «Leader-ID»"))
                {
                    docList.Add(
                        "Необходимо предоставить документы, подтверждающие участие обучившегося, профиль которого отсутствует в ИС «Leader-ID», в мероприятиях АП, проведение которых подтверждено. Такими документами может являться, в том числе, протокол мероприятия за подписью уполномоченного лица с поименным списком участников мероприятия, поименный список регистрации участников мероприятия с оригинальными подписями, в котором указан, в том числе, данный обучившийся, профиль которого отсутствует в ИС «Leader-ID». В дополнение к указанными документам могут прикладываться новости, подтверждающие проведение мероприятия с участием данного обучившегося, профиль которого отсутствует в ИС «Leader-ID» (с указанием ссылки на источник и предоставлением скриншота публикации), скриншоты, подтверждающие проведение мероприятия в формате онлайн с участием данного обучившегося, профиль которого отсутствует в ИС «Leader-ID», иные документы и материалы, которые позволяют подтвердить факт участия обучившегося, профиль которого отсутствует в ИС «Leader-ID», в мероприятиях АП.\nНаправляемые документы должны сопровождаться письмом за подписью ректора, либо уполномоченного лица с подтверждением полномочий подписанта, в котором подтверждается участие обучившегося, профиль которого отсутствует в ИС «Leader-ID», в мероприятиях АП.");
                }

                if (e.Reason.Contains("Профиль в ИС «Leader-ID» принадлежит другому человеку"))
                {
                    docList.Add(
                        "Необходимо скорректировать отчет о ходе реализации акселерационной программы (по форме приложения \u2116 6 к договору о предоставлении гранта)  и указать корректный профиль данного обучившегося в ИС «Leader-ID», либо причину его отсутствия. В случае отсутствия профиля в ИС «Leader-ID», необходимо предоставить документы, подтверждающие участие обучившегося, профиль которого отсутствует в ИС «Leader-ID», в мероприятиях АП, проведение которых подтверждено). Такими документами может являться, в том числе, протокол мероприятия за подписью уполномоченного лица с поименным списком участников мероприятия, поименный список регистрации участников мероприятия с оригинальными подписями, в котором указан, в том числе, данный обучившийся, профиль которого отсутствует в ИС «Leader-ID». В дополнение к указанными документам могут прикладываться новости, подтверждающие проведение мероприятия с участием данного обучившегося, профиль которого отсутствует в ИС «Leader-ID» (с указанием ссылки на источник и предоставлением скриншота публикации), скриншоты, подтверждающие проведение мероприятия в формате онлайн с участием данного обучившегося, профиль которого отсутствует в ИС «Leader-ID», иные документы и материалы, которые позволяют подтвердить факт участия обучившегося, профиль которого отсутствует в ИС «Leader-ID», в мероприятиях АП.\nНаправляемый доработанный отчет должен сопровождаться письмом за подписью ректора, либо уполномоченного лица с подтверждением полномочий подписанта, в котором подтверждается указание корректного профиля в ИС «Leader-ID» данного обучившегося.");
                }

                string result = string.Join(";\n", docList);

                e.Documents = result;
            }

            logger.Information("Got [{count}] errors in table 1 for contract [{c}]", errors1.Count, contractNumber);

            logger.Information("Getting events...");
            //events = Connection.GetEventsForContract(contractNumber);
            events = Connection.GetNewTable2ForContract(contractNumber);

            var badEvents = events.Where(e => e.DateStart.Contains("2019"));

            var badEvents0 = new List<Event>();
            
            var eventsToError = new List<ErrorTable2>();

            var studentIDs = trainedStudents.Select(s => s.LeaderId).ToList();

            foreach (var e in events)
            {
                var badCount = 0;
                var eIDs = e.LeaderIdNumber.Replace(" ", "").Split(',').ToList();

                if (eIDs.Any())
                {

                    foreach (var sid in eIDs)
                    {
                        if (!studentIDs.Contains(sid))
                        {
                            badCount++;
                        }
                    }

                    if (badCount == eIDs.Count())
                    {
                        badEvents0.Add(e);
                        eventsToError.Add(new ErrorTable2
                        {
                            Number = 0.ToString(),
                            Name = e.Name,
                            Link = e.Link,
                            Reason =
                                "Количество участников мероприятия равно \"0\" (по данным ИС \"Leader-ID\", посетители мероприятий не входят в список участников АП)",
                            Documents = "-",
                            Remark = "-",
                            Comment = "-"
                        });
                    }
                }
            }

            events = events.Except(badEvents).ToList();
            events = events.Except(badEvents0).ToList();

            var eventID = 1;

            foreach (var e in events)
            {
                e.Number = eventID.ToString();
                eventID++;
            }

            logger.Information("Got data about [{count}] events for contract [{c}]", events.Count, contractNumber);

            logger.Information("Getting errors in Table 2...");
            errors2 = Connection.GetErrors2ForContract(contractNumber);

            var err2ID = 1;

            foreach (var e in badEvents)
            {
                errors2.Add(new ErrorTable2
                {
                    Number = 0.ToString(),
                    Name = e.Name,
                    Link = e.Link,
                    Reason = "Дата проведения мероприятия вне рамок отчетного периода",
                    Documents = "-",
                    Remark = "-",
                    Comment = "-"
                });
            }

            foreach (var e in eventsToError)
            {
                errors2.Add(e);
            }

            foreach (var e in errors2)
            {
                e.Number = err2ID.ToString();
                err2ID++;
            }

            foreach (var e in errors2)
            {
                if (e.Reason.Contains("Выявлено дублирование"))
                {
                    e.Reason = e.Reason.Replace("Выявлено дублирование",
                        "Мероприятие дублируется (не подтверждено как уникальное)");
                }
                
                if (e.Reason.Contains("Количество участников мероприятия равно 0"))
                {
                    e.Reason = e.Reason.Replace("Количество участников мероприятия равно 0",
                        "Количество участников мероприятия равно \"0\" (по данным ИС \"Leader-ID\", кол-во участников 0)");
                }
            }

            foreach (var e in errors2)
            {
                var docList = new List<string>();

                if (e.Reason.Contains("Количество участников мероприятия равно \"0\""))
                {
                    docList.Add(
                        "Необходимо предоставить документы, подтверждающие проведение мероприятия АП, на котором присутствовали обучившиеся по данной АП, с указанием темы, места и времени проведения, перечня присутствовавших на мероприятии участников.\nТакими документами может являться, в том числе, протокол мероприятия за подписью уполномоченного лица с поименным списком участников мероприятия, поименный список регистрации участников мероприятия с   подписями. В дополнение к указанными документам могут прикладываться новости, подтверждающие проведение мероприятия (с указанием ссылки на источник и предоставлением скриншота публикации), скриншоты, подтверждающие проведение мероприятия в формате онлайн, фотоотчеты с мероприятия, иные документы и материалы, которые позволяют подтвердить факт проведения мероприятия.\nНаправляемые документы должны сопровождаться письмом за подписью ректора, либо уполномоченного лица с подтверждением полномочий подписанта, в котором подтверждается проведение мероприятий в рамках Плана реализации АП, на котором присутствовали обучившиеся по данной АП, с указанием темы, места и времени проведения.");
                }

                if (e.Reason.Contains("Мероприятие дублируется (не подтверждено как уникальное)"))
                {
                    docList.Add(
                        "Необходимо исключить дублирование в доработанной отчетной документации. Направляемый доработанный отчет о ходе реализации акселерационной программы (по форме приложения \u2116 6 к договору о предоставлении гранта) должен сопровождаться письмом за подписью ректора, либо уполномоченного лица с подтверждением полномочий подписанта, в котором подтверждается исключение дублирования из отчетных материалов.");
                }

                if (e.Reason.Contains("Отсутствует регистрация мероприятия в ИС «Leader-ID»"))
                {
                    docList.Add(
                        "Необходимо предоставить документы, подтверждающие проведение мероприятия АП, не зарегистрированного в ИС «Leader-ID», с указанием темы, места и времени проведения, перечня присутствовавших на мероприятии участников.\nТакими документами может являться, в том числе, протокол мероприятия за подписью уполномоченного лица с поименным списком участников мероприятия, поименный список регистрации участников мероприятия с   подписями. В дополнение к указанными документам могут прикладываться новости, подтверждающие проведение мероприятия (с указанием ссылки на источник и предоставлением скриншота публикации), скриншоты, подтверждающие проведение мероприятия в формате онлайн, фотоотчеты с мероприятия, иные документы и материалы, которые позволяют подтвердить факт проведения мероприятия.\n\nНаправляемые документы должны сопровождаться письмом за подписью ректора, либо уполномоченного лица с подтверждением полномочий подписанта, в котором подтверждается проведение мероприятий в рамках Плана реализации АП, не зарегистрированного в ИС «Leader-ID», с указанием темы, места и времени проведения.");
                }

                if (e.Reason.Contains("Дата проведения мероприятия вне рамок отчетного периода"))
                {
                    docList.Add(
                        "Необходимо исключить из доработанной отчетной документации мероприятия, проведенные вне рамок отчетного периода. Направляемый доработанный отчет о ходе реализации акселерационной программы (по форме приложения \u2116 6 к договору о предоставлении гранта) должен сопровождаться письмом за подписью ректора, либо уполномоченного лица с подтверждением полномочий подписанта, в котором даются пояснения по выявленному замечанию.");
                }

                if (e.Reason.Contains(
                        "Наименование мероприятия в отчетной документации не соответствует наименованию мероприятия в ИС «Leader-ID»"))
                {
                    docList.Add(
                        "Необходимо внести соответствующие изменения в доработанную отчетную документацию и указать корректное наименование мероприятия или корректный номер мероприятия в ИС «Leader-ID». Направляемый доработанный отчет о ходе реализации акселерационной программы (по форме приложения \u2116 6 к договору о предоставлении гранта) должен сопровождаться письмом за подписью ректора, либо уполномоченного лица с подтверждением полномочий подписанта, в котором подтверждается внесение соответствующих изменений и указание корректных наименований мероприятий.");
                }

                string result = string.Join(";\n", docList);

                e.Documents = result;
            }

            foreach (var e in errors2)
            {
                if (e.Name.IsNullOrEmpty())
                {
                    e.Name = Connection.GetEventByID(e.Link.Replace("https://leader-id.ru/events/", ""),
                        contractNumber);
                }
            }

            logger.Information("Got [{count}] errors in table 2 for contract [{c}]", errors2.Count, contractNumber);

            logger.Information("Getting startups...");
            startups = Connection.GetStartupsForContract(contractNumber);

            var badStartups = startups.Where(s =>
                s.Participants.Count() == s.Participants.Count(p => p.EventIDs == '-'.ToString())).ToList();

            var dupeStartups = startups.Where(s => Convert.ToInt32(s.DupeCount) > 1).ToList();

            var dupesToRemove = new List<Startup>();

            foreach (var s in dupeStartups)
            {
                if (badStartups.Contains(s))
                {
                    dupesToRemove.Add(s);
                    badStartups.Add(s);
                }
            }

            dupeStartups = dupeStartups.Except(dupesToRemove).ToList();

            badStartups = badStartups.OrderBy(s => s.Link).ToList();

            var startupsToErrors = new List<Startup>();
            
            foreach (var s in startups)
            {
                var badCount = 0;
                var sIDs = s.Participants.Select(p => p.LeaderID).ToList();
                    
                if (sIDs.Any())
                {

                    foreach (var sid in sIDs)
                    {
                        if (!studentIDs.Contains(sid))
                        {
                            badCount++;
                        }
                    }

                    if (badCount == sIDs.Count())
                    {
                        startupsToErrors.Add(s);
                    }
                }
            }

            startups = startups.Except(badStartups).ToList();
            startups = startups.Except(startupsToErrors).ToList();

            var startupNum = 1;
            foreach (var startup in startups)
            {
                startup.Number = startupNum.ToString();
                startupNum++;
            }

            foreach (var s in startups)
            {
                foreach (var p in s.Participants)
                {
                    if (p.Name.IsNullOrEmpty())
                    {
                        var name = Connection.GetStudentByID(p.LeaderID, contractNumber);

                        if (name != "#Н/Д")
                        {
                            p.Name = name;
                        }
                    }
                }
            }

            logger.Information("Got data about [{count}] startups for contract [{c}]", startups.Count, contractNumber);

            logger.Information("Getting errors in Table 3...");
            errors3 = Connection.GetErrors3ForContract(contractNumber);

            var errorsToRemove = new List<ErrorTable3>();
            var badStartupsToRemove = new List<Startup>();

            foreach (var s in badStartups)
            {
                if (errors3.Where(e => e.Link == s.Link).ToList().Any())
                {
                    badStartupsToRemove.Add(s);
                }
            }

            badStartups = badStartups.Except(badStartupsToRemove).ToList();

            foreach (var startup in badStartups)
            {
                if (Convert.ToInt32(startup.DupeCount) > 1)
                {
                    for (var i = 0; i < Convert.ToInt32(startup.DupeCount); i++)
                    {
                        errors3.Add(new ErrorTable3
                        {
                            Number = 0.ToString(),
                            Name = startup.Name,
                            Link = startup.Link,
                            Reason =
                                "Отсутствие в составе команды стартап-проекта обучившихся по соответствующей АП; Выявлено дублирование",
                            Documents = "",
                            Remark = "",
                            Comment = ""
                        });
                    }
                }
                else
                {
                    errors3.Add(new ErrorTable3
                    {
                        Number = 0.ToString(),
                        Name = startup.Name,
                        Link = startup.Link,
                        Reason = "Отсутствие в составе команды стартап-проекта обучившихся по соответствующей АП",
                        Documents = "",
                        Remark = "",
                        Comment = ""
                    });
                }
            }

            foreach (var startup in dupeStartups)
            {
                errors3.Add(new ErrorTable3
                {
                    Number = 0.ToString(),
                    Name = startup.Name,
                    Link = startup.Link,
                    Reason = "Выявлено дублирование",
                    Documents = "",
                    Remark = "",
                    Comment = ""
                });
            }

            foreach (var s in startupsToErrors)
            {
                var links = errors3.Select(e => e.Link).ToList();
                if (!links.Contains(s.Link))
                {
                    errors3.Add(new ErrorTable3
                    {
                        Number = 0.ToString(),
                        Name = s.Name,
                        Link = s.Link,
                        Reason = "Отсутствие в составе команды стартап-проекта обучившихся по соответствующей АП",
                        Documents = "",
                        Remark = "",
                        Comment = ""
                    });
                }
            }

            foreach (var err in errors3)
            {
                var isGood = Connection.CheckIfHasGoodParticipants(contractNumber, err.Link);

                if (!isGood && !err.Reason.Contains(
                        "Отсутствие в составе команды стартап-проекта обучившихся по соответствующей АП") &&
                    !err.Reason.IsNullOrEmpty())
                {
                    err.Reason += "; Отсутствие в составе команды стартап-проекта обучившихся по соответствующей АП";
                }

                if (err.Reason.Contains("; Выявлено дублирование; Выявлено дублирование"))
                {
                    err.Reason = err.Reason.Replace("; Выявлено дублирование; Выявлено дублирование",
                        "; Выявлено дублирование");
                }

                foreach (var startup in startups)
                {
                    if (startup.Link == err.Link
                        && err.Reason ==
                        "Статус проекта в акселераторе по данным в ИС «Projects»: «Проходит акселерацию» (или иной статус отличный от \"Завершено успешно\"); Выявлено дублирование")
                    {
                        if (!errorsToRemove.Contains(err))
                            errorsToRemove.Add(err);
                    }

                    if (startup.Link == err.Link && !err.Reason.Contains("Выявлено дублирование"))
                    {
                        errorsToRemove.Add(err);
                    }
                }
            }

            errors3 = errors3.Except(errorsToRemove).ToList();

            foreach (var e in errors3)
            {
                if (e.Reason.Contains("Выявлено дублирование"))
                {
                    e.Reason = e.Reason.Replace("Выявлено дублирование", "Выявлено дублирование стартап-проектов");
                }
            }

            foreach (var e in errors3)
            {
                var docList = new List<string>();

                if (e.Reason.Contains(
                        "Статус проекта в акселераторе по данным в ИС «Projects»: «Проходит акселерацию» (или иной статус отличный от \"Завершено успешно\")"))
                {
                    docList.Add(
                        "Необходимо изменить статус стартап-проектов, успешно завершивших прохождение акселерационной программы, на «Завершено успешно» в ИС «Projects». Необходимо предоставить документы, подтверждающие успешное завершение стартап-проектами соответствующей акселерационной программы.\nТакими документами могут являться, в том числе, скриншоты ИС «Projects», с указанием статуса каждого зарегистрированного стартап-проекта «Завершено успешно», иные документы и материалы, которые позволяют подтвердить факт успешного завершения стартап-проектами прохождения акселерационной программы на момент завершения АП.\nНаправляемые документы должны сопровождаться письмом за подписью ректора, либо уполномоченного лица с подтверждением полномочий подписанта, в котором подтверждается успешное завершение акселерационной программы всеми стартап-проектами на момент завершения АП.");
                }

                if (e.Reason.Contains("Отсутствие в составе команды стартап-проекта обучившихся по соответствующей АП"))
                {
                    docList.Add(
                        "Необходимо предоставить документы, подтверждающие участие в составе команды стартап-проекта обучившихся по соответствующей АП, участие которых в мероприятиях акселерационной программы подтверждено, а также скорректированный паспорт стартап-проекта (по форме приложения \u2116 15 к договору о предоставлении гранта).\nТакими документами может являться, в том числе, поименный список участников команды стартап-проекта, среди которых присутствуют подтвержденные участники мероприятий акселерационной программы, в соответствии с отчетом о ходе реализации акселерационной программы. \nНаправляемые документы должны сопровождаться письмом за подписью ректора, либо уполномоченного лица с подтверждением полномочий подписанта, в котором подтверждается участие в команде стартап-проекта обучившихся в рамках соответствующей акселерационной программы.");
                }

                if (e.Reason.Contains(
                        "Стартап-проект не принадлежит заявленной акселерационной программе"))
                {
                    docList.Add(
                        "Необходимо предоставить письмо (пояснительную записку) за подписью ректора, либо уполномоченного лица с подтверждением полномочий подписанта, с указанием полного перечня стартап-проектов, успешно завершивших заявленную акселерационную программу с указанием ссылок на них в ИС «Projects», в котором даются пояснения по выявленному замечанию и подтверждается принадлежность стартап-проекта к заявленной акселерационной программе с приложением дополнительных подтверждающих материалов, либо подтверждается исключение из отчетной документации стартап-проектов, которые не принадлежат заявленной АП.");
                }

                if (e.Reason.Contains("Отсутствие регистрации стартап-проекта в ИС «Projects»"))
                {
                    docList.Add(
                        "Необходимо предоставить документы, подтверждающие регистрацию в ИС «Projects» и успешное прохождение акселерационной программы стартап-проектом.», а также скорректированный паспорт стартап-проекта (по форме приложения \u2116 15 к договору о предоставлении гранта).\nТакими документами может являться, в том числе, скриншот страницы (с указанием гиперссылки) стартап-проекта, успешно завершившего заявленную акселерационную программу, в ИС «Projects», иные документы и материалы, которые позволяют подтвердить факт регистрации стартап-проекта в ИС «Projects» и его успешное завершение заявленной акселерационной программы.\nНаправляемые документы должны сопровождаться письмом за подписью ректора, либо уполномоченного лица с подтверждением полномочий подписанта, в котором даются пояснения по выявленному замечанию, подтверждается регистрация всех стартап-проектов в ИС «Projects», с указанием полного перечня стартап-проектов, успешно завершивших заявленную акселерационную программу с указанием ссылок на них в ИС «Projects».");
                }

                if (e.Reason.Contains("Выявлено дублирование стартап-проектов"))
                {
                    docList.Add(
                        "Необходимо исключить дублирование в доработанной отчетной документации. Предоставить письмо (пояснительную записку) за подписью ректора, либо уполномоченного лица с подтверждением полномочий подписанта, в котором даются пояснения по выявленному замечанию, с приложением дополнительных подтверждающих материалов, в случае, если дублирование отсутствует, либо подтверждается исключение из отчетной документации дублирующихся стартап-проектов, с указанием полного перечня стартап-проектов, успешно завершивших заявленную акселерационную программу с указанием ссылок на них в ИС «Projects», и скорректированные паспорта стартап-проектов (по форме приложения \u2116 15 к договору о предоставлении гранта), в случае необходимости.");
                }

                string result = string.Join(";\n", docList);

                e.Documents = result;
            }

            errors3 = errors3.OrderBy(e => e.Link).ThenBy(e => e.Reason).ToList();
            
            var errNum = 0;

            foreach (var err in errors3)
            {
                errNum++;
                err.Number = errNum.ToString();
            }

            logger.Information("Got [{count}] errors in table 3 for contract [{c}]", errors3.Count, contractNumber);
            if (!errors1.Any() && errors1.Count() == 0)
            {
                errors1.Add(new ErrorTable1
                {
                    Number = "",
                    Name = "",
                    LeaderLink = "",
                    Reason = "",
                    Documents = "",
                    Remark = "",
                    Comment = ""
                });
            }

            if (!errors2.Any() && errors2.Count() == 0)
            {
                errors2.Add(new ErrorTable2
                {
                    Number = "",
                    Name = "",
                    Link = "",
                    Reason = "",
                    Documents = "",
                    Remark = "",
                    Comment = ""
                });
            }

            if (!errors3.Any() && errors3.Count() == 0)
            {
                errors3.Add(new ErrorTable3
                {
                    Number = "",
                    Name = "",
                    Link = "",
                    Reason = "",
                    Documents = "",
                    Remark = "",
                    Comment = ""
                });
            }

            if (!trainedStudents.Any() && trainedStudents.Count() == 0)
            {
                trainedStudents.Add(new TrainedStudent
                {
                    Number = "",
                    LeaderId = "",
                    FIO = "",
                    EventsId = "",
                    Count = "",
                    StartUp = "",
                    Link = ""
                });
            }

            if (!events.Any() && events.Count() == 0)
            {
                events.Add(new Event
                {
                    Number = "",
                    Name = "",
                    LeaderId = "",
                    Link = "",
                    DateStart = "",
                    Format = "",
                    CountOfParticipants = "",
                    LeaderIdNumber = ""
                });
            }

            if (!startups.Any() && startups.Count() == 0)
            {
                startups.Add(new Startup
                {
                    Number = "",
                    Name = "",
                    Link = "",
                    HasSign = "",
                    Category = ""
                });
            }

            dataList.Add(new DataForWord(contractNumber, trainedStudents, events, startups, errors1, errors2, errors3));

            return dataList;
        }


// для тестовых данных
        static List<DataForWord> GenerateTestData(Logger logger)
        {
            logger.Information("Generating test data...");
            var random = new Random();
            var dataList = new List<DataForWord>();

            for (int i = 0; i < 5; i++)
            {
                var contractNumber = $"CN-{random.Next(1000, 9999)}";
                logger.Information("Generated contract number [{contract}]", contractNumber);

                var trainedStudents = new List<TrainedStudent>();
                var events = new List<Event>();

                for (int j = 0; j < random.Next(1, 5); j++)
                {
                    var eventNumber = random.Next(1000, 9999);
                    events.Add(new Event
                    {
                        Number = (j + 1).ToString(),
                        Name = $"Event-{eventNumber}",
                        LeaderId = $"LID-{random.Next(1000, 9999)}",
                        Link = $"http://eventlink.com/{eventNumber}",
                        DateStart = DateTime.Now.AddDays(random.Next(-30, 30)).ToString("yyyy-MM-dd HH:mm"),
                        Format = "Online",
                        CountOfParticipants = random.Next(10, 100).ToString(),
                        LeaderIdNumber = $"LIDN-{random.Next(1000, 9999)}"
                    });
                }

                logger.Information("Generated [{count}] events", events.Count);


                for (int k = 0; k < random.Next(1, 10); k++)
                {
                    var studentNumber = random.Next(1000, 9999);
                    trainedStudents.Add(new TrainedStudent
                    {
                        Number = (k + 1).ToString(),
                        LeaderId = random.Next(1000, 9999).ToString(),
                        FIO = $"Student-{studentNumber}",
                        EventsId = $"EID-{random.Next(1000, 9999)}",
                        Count = random.Next(1, 10).ToString(),
                        StartUp = $"Startup-{random.Next(1000, 9999)}",
                        Link = $"http://startuplink.com/{studentNumber}"
                    });
                }

                logger.Information("Generated [{count}] participants", trainedStudents.Count);

                var startups = new List<Startup>();
                for (int l = 0; l < random.Next(1, 10); l++)
                {
                    var participants = new List<Participant>();
                    for (int m = 0; m < random.Next(1, 10); m++)
                    {
                        participants.Add(new Participant
                        {
                            Name = $"Participant-{m + 1}",
                            LeaderID = $"LID-{random.Next(1000, 9999)}",
                            EventIDs = String.Join(", ",
                                Enumerable.Range(0, random.Next(1, 5)).Select(x => random.Next(1000, 9999).ToString()))
                        });
                    }

                    startups.Add(new Startup
                    {
                        Number = (l + 1).ToString(),
                        Name = $"Startup-{l + 1}",
                        Link = $"http://startuplink.com/{i + 1}",
                        HasSign = String.Empty,
                        Category = String.Empty,
                        Participants = participants
                    });
                }

                logger.Information("Generated [{count}] startups", startups.Count);

                var errors1 = new List<ErrorTable1>();
                for (int n = 0; n < random.Next(1, 10); n++)
                {
                    var leaderIDNumber = random.Next(1000, 9999);
                    errors1.Add(new ErrorTable1()
                    {
                        Number = (n + 1).ToString(),
                        Name = $"Student-{leaderIDNumber}",
                        LeaderLink = $"https://leader-id.ru/users/{leaderIDNumber}",
                        Reason = GenerateRandomString(random.Next(1, 100)),
                        Documents = GenerateRandomString(random.Next(1, 100)),
                        Remark = GenerateRandomString(random.Next(1, 100)),
                        Comment = GenerateRandomString(random.Next(1, 100))
                    });
                }

                var errors2 = new List<ErrorTable2>();
                for (int n = 0; n < random.Next(1, 10); n++)
                {
                    var leaderIDNumber = random.Next(1000, 9999);
                    errors2.Add(new ErrorTable2()
                    {
                        Number = (n + 1).ToString(),
                        Name = $"Student-{leaderIDNumber}",
                        Link = $"https://leader-id.ru/users/{leaderIDNumber}",
                        Reason = GenerateRandomString(random.Next(1, 100)),
                        Documents = GenerateRandomString(random.Next(1, 100)),
                        Remark = GenerateRandomString(random.Next(1, 100)),
                        Comment = GenerateRandomString(random.Next(1, 100))
                    });
                }

                var errors3 = new List<ErrorTable3>();
                for (int n = 0; n < random.Next(1, 10); n++)
                {
                    var leaderIDNumber = random.Next(1000, 9999);
                    errors2.Add(new ErrorTable2()
                    {
                        Number = (n + 1).ToString(),
                        Name = $"Student-{leaderIDNumber}",
                        Link = $"https://leader-id.ru/users/{leaderIDNumber}",
                        Reason = GenerateRandomString(random.Next(1, 100)),
                        Documents = GenerateRandomString(random.Next(1, 100)),
                        Remark = GenerateRandomString(random.Next(1, 100)),
                        Comment = GenerateRandomString(random.Next(1, 100))
                    });
                }

                logger.Information("Generated [{count}] errors for table 1", errors1.Count);

                dataList.Add(new DataForWord(contractNumber, trainedStudents, events, startups, errors1, errors2,
                    errors3));
            }

            return dataList;
        }

        static List<IGrouping<string, DataForWord>> GroupData(List<DataForWord> testData)
        {
            return testData.GroupBy(data => data.ContractNumber).ToList();
        }

        static string GenerateRandomString(int length)
        {
            var random = new Random();
            var chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
            return new string(Enumerable.Repeat(chars, length).Select(s => s[random.Next(s.Length)]).ToArray());
        }
    }
}