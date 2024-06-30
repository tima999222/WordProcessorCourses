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

        public DataForWord(string contractNumber, List<TrainedStudent> trainedStudents, List<Event> events)
        {
            ContractNumber = contractNumber;
            TrainedStudents = trainedStudents;
            Events = events;
        }

        public string ContractNumber { get; set; }
        public List<TrainedStudent> TrainedStudents { get; set; }
        public List<Event> Events { get; set; }
    }
}