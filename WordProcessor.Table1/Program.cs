using AStepanov.Core.Ex;
using System.Reflection;
using WordProcessor.Table1;
using WordProcessor.Table1.Entities;


var dataFromDB = GetDataFromDatabase();
var testData = GenerateTestData();

var groupedData = new List<IGrouping<string, DataForWord>>();

if (!dataFromDB.Any())
{
    groupedData = GroupData(testData);
}
else
{
    groupedData = GroupData(dataFromDB);
}


var result = ApplicantForGrant.CreateApplicationsForOrder(groupedData);
if (result != null)
{
    File.WriteAllBytes(Assembly.GetExecutingAssembly().Directory() + "/documents.zip",
        result); //zip спавнится в папке bin
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

//TODO: Сделать, чтобы все договоры по очереди перебирались
static List<DataForWord> GetDataFromDatabase()
{
    var dataList = new List<DataForWord>();
    
    var contractNumber = "70-2023-000622";
    var trainedStudents = new List<TrainedStudent>();
    var events = new List<Event>();


    trainedStudents = Connection.GetParticipantsForContract(contractNumber);
    events = Connection.GetEventsForContract(contractNumber);
    
    dataList.Add(new DataForWord(contractNumber, trainedStudents, events));
    
    return dataList;
}

static List<IGrouping<string, DataForWord>> GroupData(List<DataForWord> testData)
{
    return testData.GroupBy(data => data.ContractNumber).ToList();
}