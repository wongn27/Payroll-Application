using Payroll;
using System;
using System.Collections.Generic;
using System.IO;

namespace Main
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            //List<Worker> workerList = ReadWorkersData();

            //foreach (Worker worker in workerList)
            //{
            //    Console.WriteLine($"{worker.Name} {worker.Id} {worker.PayRate.ToString("n02")}");
            //}

            Worker worker = ReadWorkerData("Worker.txt");
            WriteWorkerData(worker, "WriteWorkerData.txt");
            Shift shift = ReadShiftData();
            WriteShiftData(shift, "WriteShiftData.txt");
            //Console.WriteLine($"{shift.WorkerID} {shift.HoursWorked} {shift.Date}");

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

        public static Worker ReadWorkerData(string fileName)
        {
            var worker = new Worker();
            var fileStream = new FileStream(fileName, FileMode.Open);
            var streamReader = new StreamReader(fileStream);

            string data;

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

            return worker;
        }

        public static Shift ReadShiftData()
        {
            var shift = new Shift();

            var fileStream = new FileStream("Shift.txt", FileMode.Open);
            var streamReader = new StreamReader(fileStream);

            string data;
            data = streamReader.ReadLine();
            shift.WorkerID = data;

            data = streamReader.ReadLine();
            if (double.TryParse(data, out double potentialDouble))
            {
                shift.HoursWorked = potentialDouble;
            }


            try
            {
                data = streamReader.ReadLine();
                int potentialYear = int.Parse(data);

                data = streamReader.ReadLine();
                int potentialMonth = int.Parse(data);

                data = streamReader.ReadLine();
                int potentialDay = int.Parse(data);

                var datetime = new DateTime(potentialYear, potentialMonth, potentialDay);
                shift.Date = datetime;
            }
            catch (Exception e)
            {
                Console.WriteLine($"Error reading in the date: {e.ToString()}");
            }

            return shift;
        }

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

        public static void WriteShiftData(Shift shift, string fileName)
        {
            using (var streamWriter = new StreamWriter(fileName))
            {
                streamWriter.WriteLine(shift.WorkerID);
                streamWriter.WriteLine(shift.HoursWorked);
                streamWriter.WriteLine(shift.Date);
            }
        }
    }
}
