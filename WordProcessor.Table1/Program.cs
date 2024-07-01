using AStepanov.Core.Ex;
using System.Reflection;
using WordProcessor.Table1;
using WordProcessor.Table1.Entities;
using Serilog;
using Serilog.Core;
using Serilog.Sinks.SystemConsole;

var logger = ConfigureLogger();

logger.Information("Starting Up");

var contractNumber = "70-2023-000622";

logger.Information("Trying to get data from database for contract with number [{contractNumber}]", contractNumber);
var dataFromDB = GetDataFromDatabase(contractNumber, logger);
logger.Information("Got data for contract with number [{contractNumber}]", contractNumber);


var testData = GenerateTestData();

var groupedData = new List<IGrouping<string, DataForWord>>();

logger.Information("Grouping data...");
if (!dataFromDB.Any())
{
    groupedData = GroupData(testData);
}
else
{
   
    groupedData = GroupData(dataFromDB);
}
logger.Information("Data grouped successfully");

logger.Information("Creating file...");
var result = ApplicantForGrant.CreateApplicationsForOrder(groupedData);
if (result != null)
{
    File.WriteAllBytes(Assembly.GetExecutingAssembly().Directory() + "/documents.zip",
        result); //zip спавнится в папке bin
    logger.Information("File [{contractNumber}.docx] added to archive", contractNumber);
}


// для тестовых данных
static List<DataForWord> GenerateTestData()
{
    var random = new Random();
    var dataList = new List<DataForWord>();

    for (int i = 0; i < 5; i++)
    {
        var contractNumber = $"CN-{random.Next(1000, 9999)}";
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

        dataList.Add(new DataForWord(contractNumber, trainedStudents, events));
    }

    return dataList;
}

static List<DataForWord> GetDataFromDatabase(string contractNumber, Logger logger)
{
    var dataList = new List<DataForWord>();
    var trainedStudents = new List<TrainedStudent>();
    var events = new List<Event>();
    

    logger.Information("Getting participants...");
    trainedStudents = Connection.GetParticipantsForContract(contractNumber);
    logger.Information("Got data about [{count}] participants", trainedStudents.Count);
    
    logger.Information("Getting events...");
    events = Connection.GetEventsForContract(contractNumber);
    logger.Information("Got data about [{count}] events", events.Count);
    
    dataList.Add(new DataForWord(contractNumber, trainedStudents, events));
    
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