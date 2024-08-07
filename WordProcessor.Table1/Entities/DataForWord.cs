﻿using DocumentFormat.OpenXml.Math;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordProcessor.Table1.Entities
{
    public class DataForWord
    {
        public DataForWord()
        {
            
        }

        public DataForWord(string contractNumber, List<TrainedStudent> trainedStudents, List<Event> events, List<Startup> startups, List<ErrorTable1> error1, List<ErrorTable2> error2, List<ErrorTable3> error3)
        {
            ContractNumber = contractNumber;
            TrainedStudents = trainedStudents;
            Events = events;
            Startups = startups;
            Error1 = error1;
            Error2 = error2;
            Error3 = error3;
        }

        public string ContractNumber { get; set; }
        public List<TrainedStudent> TrainedStudents { get; set; }
        public List<Event> Events { get; set; }
        public List<Startup> Startups { get; set; }
        
        public List<ErrorTable1> Error1 { get; set; }
        public List<ErrorTable2> Error2 { get; set; }
        
        public List<ErrorTable3> Error3 { get; set; }
    }
}
