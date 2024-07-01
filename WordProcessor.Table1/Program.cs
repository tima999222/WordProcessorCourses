using AStepanov.Core.Ex;
using System.Reflection;
using WordProcessor.Table1;
using WordProcessor.Table1.Entities;
using Serilog;
using Serilog.Core;

var logger = ConfigureLogger();

logger.Information("Starting Up");

var contractNumber = "70-2023-000622";

logger.Information("Trying to get data from database for contract with number [{contractNumber}]", contractNumber);
List<DataForWord> dataFromDB = null;//GetDataFromDatabase(contractNumber, logger);


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
    var result = ApplicantForGrant.CreateApplicationsForOrder(groupedData, zipPath);
    if (result != null)
    {
        foreach (var data in groupedData)
            logger.Information("File [{contractNumber}] added to archive", data.Key + ".docx");
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
                DateStart = DateTime.Now.AddDays(random.Next(-30, 30)),
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
        for (int l = 0; l < 10; l++)
        {
            var participants = new List<Participant>();
            for (int m = 0; m < random.Next(1, 10); m++)
            {
                participants.Add(new Participant
                {
                    Number = m + 1,
                    Name = $"Participant-{m + 1}",
                    LeaderID = $"LID-{random.Next(1000, 9999)}",
                    EventIDs = String.Join(", ", Enumerable.Range(0, random.Next(1, 5)).Select(x => random.Next(1000, 9999).ToString()))
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
        
        dataList.Add(new DataForWord(contractNumber, trainedStudents, events, startups));
    }

    return dataList;
}

static List<DataForWord> GetDataFromDatabase(string contractNumber, Logger logger)
{
    var dataList = new List<DataForWord>();
    var trainedStudents = new List<TrainedStudent>();
    var events = new List<Event>();
    var startups = new List<Startup>();
    

    logger.Information("Getting participants...");
    trainedStudents = Connection.GetParticipantsForContract(contractNumber);
    logger.Information("Got data about [{count}] participants", trainedStudents.Count);
    
    logger.Information("Getting events...");
    events = Connection.GetEventsForContract(contractNumber);
    logger.Information("Got data about [{count}] events", events.Count);
    
    logger.Information("Getting startups...");
    startups = Connection.GetStartupsForContract(contractNumber);
    logger.Information("Got data about [{count}] startups", startups.Count);
    
    dataList.Add(new DataForWord(contractNumber, trainedStudents, events, startups));
    
    return dataList;
}

static List<IGrouping<string, DataForWord>> GroupData(List<DataForWord> testData)
{
    return testData.GroupBy(data => data.ContractNumber).ToList();
}

static Logger ConfigureLogger()
{
    return new LoggerConfiguration()
        .MinimumLevel.Debug()
        .WriteTo.Console()
        .CreateLogger();
}