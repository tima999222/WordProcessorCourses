using AStepanov.Core.Ex;
using System.Reflection;
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
                "70-2023-000670",
                "70-2023-000622",
                "70-2023-000669",
                "70-2023-000620",
                "70-2023-000672",
                "70-2023-000677",
                "70-2023-000678",
                "70-2023-000671",
                "70-2023-000621",
                "70-2023-000673",

                "70-2023-000626",
                "70-2023-000644",
                "70-2023-000645",
                "70-2023-000650",
                "70-2023-000651",
                "70-2023-000655",
                "70-2023-000662",
                "70-2023-000663",
                "70-2023-000664",
                "70-2023-000667",

                "70-2023-000665",
                "70-2023-000680",
                "70-2023-000681",
                "70-2023-000682",
                "70-2023-000683",
                "70-2023-000684",
                "70-2023-000685",
                "70-2023-000691",
                "70-2023-000692",
                "70-2023-000693",

                "70-2023-000668",
                "70-2023-000679",
                "70-2023-000694",
                "70-2023-000695",
                "70-2023-000696",
                "70-2023-000697",
                "70-2023-000719",
                "70-2023-000768",
            
                
                "70-2023-000647",
                "70-2023-000648"
                #endregion
                #region Not Important
                /*"70-2023-000618",
                "70-2023-000619",
                "70-2023-000623",
                "70-2023-000624",
                "70-2023-000625",
                "70-2023-000627",
                "70-2023-000628",
                "70-2023-000629",
                "70-2023-000633",
                "70-2023-000634",
                
                "70-2023-000636",
                "70-2023-000637",
                "70-2023-000638",
                "70-2023-000641",
                "70-2023-000642",
                "70-2023-000643",
                "70-2023-000646",
                "70-2023-000661",
                "70-2023-000686",
                "70-2023-000687",
                
                "70-2023-000688",
                "70-2023-000689",
                "70-2023-000690",
                "70-2023-000698",
                "70-2023-000700",
                "70-2023-000701",
                "70-2023-000702",
                "70-2023-000703",
                "70-2023-000704",
                
                "70-2023-000706",
                "70-2023-000707",
                "70-2023-000708",
                "70-2023-000710",
                "70-2023-000711",
                "70-2023-000712",
                "70-2023-000716",
                
                "70-2023-000720",
                "70-2023-000722",
                "70-2023-000723",
                "70-2023-000726",
                "70-2023-000728",
                "70-2023-000729",
                "70-2023-000730",
                "70-2023-000732",
                "70-2023-000733",
                "70-2023-000734",
                
                "70-2023-000735",
                "70-2023-000737",
                "70-2023-000738",
                "70-2023-000741",
                "70-2023-000742",
                "70-2023-000743",
                "70-2023-000744",
                "70-2023-000746",
                "70-2023-000747",
                "70-2023-000748",*/
                
                
                #endregion
            ];

            foreach (var contractsChunk in SplitIntoChunks(contracts, 3))
            {
                Parallel.ForEach(contractsChunk, contract =>
                {
                    fillContract(contract, logger);
                    Thread.Sleep(5000);
                });
            }
            
           // fillContract("70-2023-000628", logger);
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
                            logger.Information("File [{contractNumber}] added to archive", "Таблицы-" + data.Key + ".docx");
                    }
                }
            }
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
                        Number = j + 1,
                        Name = $"Event-{eventNumber}",
                        LeaderId = $"LID-{random.Next(1000, 9999)}",
                        Link = $"http://eventlink.com/{eventNumber}",
                        DateStart = DateTime.Now.AddDays(random.Next(-30, 30)).ToString("yyyy-MM-dd HH:mm"),
                        Format = "Online",
                        CountOfParticipants = random.Next(10, 100),
                        LeaderIdNumber = $"LIDN-{random.Next(1000, 9999)}"
                    });
                }

                logger.Information("Generated [{count}] events", events.Count);


                for (int k = 0; k < random.Next(1, 10); k++)
                {
                    var studentNumber = random.Next(1000, 9999);
                    trainedStudents.Add(new TrainedStudent
                    {
                        Number = k + 1,
                        LeaderId = random.Next(1000, 9999),
                        FIO = $"Student-{studentNumber}",
                        EventsId = $"EID-{random.Next(1000, 9999)}",
                        Count = random.Next(1, 10),
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
                        Number = l + 1,
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

                logger.Information("Generated [{count}] errors for table 1", errors1.Count);

                dataList.Add(new DataForWord(contractNumber, trainedStudents, events, startups, errors1, errors2));
            }

            return dataList;
        }

        static List<DataForWord> GetDataFromDatabase(string contractNumber, Logger logger)
        {
            var dataList = new List<DataForWord>();
            var trainedStudents = new List<TrainedStudent>();
            var events = new List<Event>();
            var startups = new List<Startup>();

            var errors1 = new List<ErrorTable1>();
            var errors2 = new List<ErrorTable2>();


            logger.Information("Getting participants...");
            //trainedStudents = Connection.GetParticipantsForContract(contractNumber);
            trainedStudents = Connection.GetNewTable1ForContract(contractNumber);
            logger.Information("Got data about [{count}] participants for contract [{c}]", trainedStudents.Count,
                contractNumber);

            logger.Information("Getting errors in Table 1...");
            //errors1 = Connection.GetErrors1ForContract(contractNumber);
            errors1 = Connection.GetErrorsForContract(contractNumber);
            logger.Information("Got [{count}] errors in table 1 for contract [{c}]", errors1.Count, contractNumber);

            logger.Information("Getting events...");
            //events = Connection.GetEventsForContract(contractNumber);
            events = Connection.GetNewTable2ForContract(contractNumber);
            logger.Information("Got data about [{count}] events for contract [{c}]", events.Count, contractNumber);
            
            logger.Information("Getting errors in Table 2...");
            errors2 = Connection.GetErrors2ForContract(contractNumber);
            logger.Information("Got [{count}] errors in table 2 for contract [{c}]", errors2.Count, contractNumber);

            logger.Information("Getting startups...");
            startups = Connection.GetStartupsForContract(contractNumber);
            logger.Information("Got data about [{count}] startups for contract [{c}]", startups.Count, contractNumber);

            if (!errors1.Any() && errors1.Count() == 0)
            {
                errors1.Add(new ErrorTable1());
            }
            
            if (!errors2.Any() && errors2.Count() == 0)
            {
                errors2.Add(new ErrorTable2());
            }

            if (!trainedStudents.Any() && trainedStudents.Count() == 0)
            {
                trainedStudents.Add(new TrainedStudent());
            }

            if (!events.Any() && events.Count() == 0)
            {
                events.Add(new Event());
            }

            if (!startups.Any() && startups.Count() == 0)
            {
                startups.Add(new Startup());
            }

            dataList.Add(new DataForWord(contractNumber, trainedStudents, events, startups, errors1, errors2));

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