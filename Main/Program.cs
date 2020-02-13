//******************************************************
// File: Program.cs
//
// Purpose: Contains the class definition for Program.
//          Program will read from input files, 
//          write to separate output files, and 
//          displays a string that contains 
//          descriptive text and data to 
//          the screen.
//
// Written By: Natalie Wong
//
// Compiler: Visual Studio 2019
//
//******************************************************

using Payroll;
using System;
using System.Collections.Generic;
using System.IO;

namespace Main
{
    internal class Program
    {
        #region Methods
        //**********************************************************
        // Method: Main
        //
        // Purpose: Calls the Read, Write, and ToString methods.
        //**********************************************************
        private static void Main(string[] args)
        {
            Worker worker = ReadWorkerData("Worker.txt");
            WriteWorkerData(worker, "WriteWorkerData.txt");

            Shift shift = ReadShiftData("Shift.txt");
            WriteShiftData(shift, "WriteShiftData.txt");

            Console.WriteLine(worker.ToString());
            Console.WriteLine(shift.ToString());

            Console.Read();
        }

        public static List<Worker> ReadWorkersData(string fileName)
        {
            var workerList = new List<Worker>();
            var fileStream = new FileStream(fileName, FileMode.Open);
            var streamReader = new StreamReader(fileStream);

            string data;
            while (!streamReader.EndOfStream)
            {
                var worker = new Worker();
                data = streamReader.ReadLine();
                worker.Name = data;

                data = streamReader.ReadLine();
                if (int.TryParse(data, out int potentialInt))
                {
                    worker.Id = potentialInt;
                }

                data = streamReader.ReadLine();
                if (double.TryParse(data, out double potentialDouble))
                {
                    worker.PayRate = potentialDouble;
                }
                workerList.Add(worker);
            }

            return workerList;
        }

        //**********************************************************
        // Method: ReadWorkerData
        //
        // Purpose: To read in worker data from a specified file.
        //**********************************************************
        public static Worker ReadWorkerData(string fileName)
        {
            var worker = new Worker();
            var fileStream = new FileStream(fileName, FileMode.Open);
            var streamReader = new StreamReader(fileStream);

            string data;

            data = streamReader.ReadLine();
            worker.Name = data;

            data = streamReader.ReadLine();
            // Converts the string read, into an int
            if (int.TryParse(data, out int potentialInt))
            {
                worker.Id = potentialInt;
            }

            data = streamReader.ReadLine();
            // Converts the string read, into a double
            if (double.TryParse(data, out double potentialDouble))
            {
                worker.PayRate = potentialDouble;
            }

            return worker;
        }

        //**********************************************************
        // Method: ReadShiftData
        //
        // Purpose: To read in shift data from a specified file.
        //**********************************************************
        public static Shift ReadShiftData(string fileName)
        {
            var shift = new Shift();

            var fileStream = new FileStream(fileName, FileMode.Open);
            var streamReader = new StreamReader(fileStream);

            string data;
            data = streamReader.ReadLine();
            shift.WorkerID = data;

            data = streamReader.ReadLine();
            // Converts the string read, into a double
            if (double.TryParse(data, out double potentialDouble))
            {
                shift.HoursWorked = potentialDouble;
            }

            try
            {
                data = streamReader.ReadLine();
                // Convers the string read, into an int
                int potentialYear = int.Parse(data);

                data = streamReader.ReadLine();
                // Convers the string read, into an int
                int potentialMonth = int.Parse(data);

                data = streamReader.ReadLine();
                // Convers the string read, into an int
                int potentialDay = int.Parse(data);

                var datetime = new DateTime(potentialYear, potentialMonth, potentialDay);
                shift.Date = datetime;
            }
            catch (Exception exception)
            {
                Console.WriteLine($"Error reading in the date: {exception.ToString()}");
            }

            return shift;
        }

        //*****************************************************************
        // Method: WriteWorkerData
        //
        // Purpose: To write worker data to a specified output file.
        //*****************************************************************
        public static void WriteWorkerData(Worker worker, string fileName)
        {
            using (var streamWriter = new StreamWriter(fileName))
            {
                streamWriter.WriteLine(worker.Name);
                streamWriter.WriteLine(worker.Id);
                streamWriter.WriteLine(worker.PayRate);
            }
        }

        public static void WriteWorkersData(List<Worker> workers, string fileName)
        {
            using (var streamWriter = new StreamWriter(fileName))
            {
                foreach (Worker worker in workers)
                {
                    streamWriter.WriteLine(worker.Name);
                    streamWriter.WriteLine(worker.Id);
                    streamWriter.WriteLine(worker.PayRate);
                }
            }
        }

        //*****************************************************************
        // Method: WriteShiftData
        //
        // Purpose: To write shift data to a specified output file.
        //*****************************************************************
        public static void WriteShiftData(Shift shift, string fileName)
        {
            using (var streamWriter = new StreamWriter(fileName))
            {
                streamWriter.WriteLine(shift.WorkerID);
                streamWriter.WriteLine(shift.HoursWorked);

                int month = shift.Date.Month;
                int year = shift.Date.Year;
                int day = shift.Date.Day;

                streamWriter.WriteLine(month);
                streamWriter.WriteLine(year);
                streamWriter.WriteLine(day);
            }
        }
        #endregion
    }
}
