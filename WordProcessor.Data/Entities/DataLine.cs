﻿namespace WordProcessor.Data.Entities
{
    public class DataLine
    {
        public long Number { get; set; }
        public long LeaderId { get; set; }
        public string FIO { get; set; }
        public string EventsId { get; set; }
        public int Count { get; set; }
        public string StartUp { get; set; }
        public string Link { get; set; }
    }
}
