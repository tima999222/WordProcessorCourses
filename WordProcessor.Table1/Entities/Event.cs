using System;

namespace WordProcessor.Table1.Entities
{
    public class Event
    {
        public long Number { get; set; }   
        public string Name { get; set; }
        public string LeaderId { get; set; }
        public string Link { get; set; }
        public string DateStart { get; set; }
        public string Format { get; set; }
        public long CountOfParticipants { get; set; }
        public string LeaderIdNumber { get; set; }

    }
}
