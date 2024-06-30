using AStepanov.Core.Ex;
using WordProcessor.Data.Entities;

namespace WordProcessor.Tests
{
    public class Tests
    {
        private DataLine _testData;

        [SetUp]
        public void Setup()
        {
            _testData = new DataLine
            {
                Number = 1,
                LeaderId = 123,
                FIO = "Иван Иванов",
                EventsId = "E1",
                Count = 15,
                StartUp = "Project1",
                Link = "http://example.com/1"
            };
        }

        [Test]
        public void Test1()
        {
            var res = _testData.GetProperties();

            foreach(var prop in res)
            {
                Console.WriteLine(prop.Key + " = " + prop.Value);
            }
        }
    }
}