//******************************************************
// File: Program.cs
//
// Purpose: Contains the class definition for Program.
//          Program will read the Worker and Shift from 
//          a JSON and XML file, write Worker and Shift
//          to a JSON and XML file, write both Worker
//          and Shift to an Excel file, and display
//          Worker and Shift data on the screen.      
//
// Written By: Natalie Wong
//
// Compiler: Visual Studio 2019
//
//******************************************************

using Microsoft.Office.Interop.Excel;
using Payroll;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;

namespace Main2
{
    public class Program
    {
        /// <summary>
        /// Presents a menu to the user and performs an action depending on what
        /// the user chooses to do.
        /// </summary>
        /// <returns></returns>
        private static void Main(string[] args)
        {
            var worker = new Worker();
            var shift = new Shift();
            int choice;

            do
            {
                Console.WriteLine("Payroll Menu");
                Console.WriteLine("------------");
                Console.WriteLine("1 - Read Worker from JSON file");
                Console.WriteLine("2 - Read Worker from XML file");
                Console.WriteLine("3 - Write Worker to JSON file");
                Console.WriteLine("4 - Write Worker to XML file");
                Console.WriteLine("5 - Write Worker to Excel file");
                Console.WriteLine("6 - Display Worker data on screen");
                Console.WriteLine("7 - Read Shift from JSON file");
                Console.WriteLine("8 - Read Shift from XML file");
                Console.WriteLine("9 - Write Shift to JSON file");
                Console.WriteLine("10 - Write Shift to XML file");
                Console.WriteLine("11 - Write Shift to Excel file");
                Console.WriteLine("12 - Display Shift data on screen");
                Console.WriteLine("13 - Exit");
                Console.Write("Enter Choice: ");

                int userInput = Convert.ToInt32(Console.ReadLine());

                choice = userInput;

                switch (choice)
                {
                    case 1:
                        try
                        {
                            string fileName = UserEnteringFileName();
                            worker = DeserializeWorkerJSON(fileName);
                        }
                        catch (Exception exception)
                        {
                            Console.WriteLine($"Error: {exception.Message}");
                        }
                        break;

                    case 2:
                        try
                        {
                            string fileName = UserEnteringFileName();
                            worker = DeserializeWorkerXML(fileName);
                        }
                        catch (Exception exception)
                        {
                            Console.WriteLine($"Error: {exception.Message}");
                        }
                        break;

                    case 3:
                        try
                        {
                            string fileName = UserEnteringFileName();
                            SerializeWorkerJSON(fileName, worker);
                        }
                        catch (Exception exception)
                        {
                            Console.WriteLine($"Error: {exception.Message}");
                        }
                        break;

                    case 4:
                        try
                        {
                            string fileName = UserEnteringFileName();
                            SerializeWorkerXML(fileName, worker);
                        }
                        catch (Exception exception)
                        {
                            Console.WriteLine($"Error: {exception.Message}");
                        }
                        break;

                    case 5:
                        try
                        {
                            string fileName = UserEnteringFileName();
                            WriteWorkerToExcel(fileName, worker);
                        }
                        catch (Exception exception)
                        {
                            Console.WriteLine($"Error: {exception.Message}");
                        }
                        break;
                    case 6:
                        Console.WriteLine(worker.ToString());
                        break;

                    case 7:
                        try
                        {
                            string fileName = UserEnteringFileName();
                            shift = DeserializeShiftJSON(fileName);
                        }
                        catch (Exception exception)
                        {
                            Console.WriteLine($"Error: {exception.Message}");
                        }
                        break;

                    case 8:
                        try
                        {
                            string fileName = UserEnteringFileName();
                            shift = DeserializeShiftXML(fileName);
                        }
                        catch (Exception exception)
                        {
                            Console.WriteLine($"Error: {exception.Message}");
                        }
                        break;

                    case 9:
                        try
                        {
                            string fileName = UserEnteringFileName();
                            SerializeShiftJSON(fileName, shift);      
                        }
                        catch (Exception exception)
                        {
                            Console.WriteLine($"Error: {exception.Message}");
                        }
                        break;

                    case 10:
                        try
                        {
                            string fileName = UserEnteringFileName();
                            SerializeShiftXML(fileName, shift);
                        }
                        catch (Exception exception)
                        {
                            Console.WriteLine($"Error: {exception.Message}");
                        }
                        break;

                    case 11:
                        try
                        {
                            string fileName = UserEnteringFileName();
                            WriteShiftToExcel(fileName, shift);
                        }
                        catch (Exception exception)
                        {
                            Console.WriteLine($"Error: {exception.Message}");
                        }
                        break;
                    case 12:
                        Console.WriteLine(shift.ToString());
                        break;
                }
            } while (choice != 13);
        }

        /// <summary>
        /// Asks the user to enter a file name.
        /// </summary>
        /// <returns></returns>
        public static string UserEnteringFileName()
        {
            Console.Write("Enter filename: ");

            string fileName = Console.ReadLine();

            return fileName;
        }

        #region Worker Methods
        /// <summary>
        /// Saves Worker data in JSON file format <see cref="Worker"/>
        /// </summary>
        /// <returns></returns>
        public static void SerializeWorkerJSON(string fileName, Worker worker)
        {
            var writer = new FileStream(fileName, FileMode.Create, FileAccess.Write);
            var dataContractJsonSerializer = new DataContractJsonSerializer(typeof(Worker));

            dataContractJsonSerializer.WriteObject(writer, worker);
            writer.Close();
        }

        /// <summary>
        /// Reads Worker from JSON file <see cref="Worker"/>
        /// </summary>
        /// <returns></returns>
        public static Worker DeserializeWorkerJSON(string fileName)
        {
            Worker worker;

            var reader = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            var inputSerializer = new DataContractJsonSerializer(typeof(Worker));

            worker = (Worker)inputSerializer.ReadObject(reader);
            reader.Close();

            return worker;
        }

        /// <summary>
        /// Reads Worker from XML file <see cref="Worker"/>
        /// </summary>
        /// <returns></returns>
        public static Worker DeserializeWorkerXML(string fileName)
        {
            Worker worker;

            var reader = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            var inputSerializer = new DataContractSerializer(typeof(Worker));

            worker = (Worker)inputSerializer.ReadObject(reader);
            reader.Close();

            return worker;
        }

        /// <summary>
        /// Saves Worker data in XML file format <see cref="Worker"/>
        /// </summary>
        /// <returns></returns>
        public static void SerializeWorkerXML(string fileName, Worker worker)
        {
            var writer = new FileStream(fileName, FileMode.Create, FileAccess.Write);
            var serXml = new DataContractSerializer(typeof(Worker)); // For serializing to XML

            serXml.WriteObject(writer, worker);
            writer.Close();
        }

        /// <summary>
        /// Writes Worker to Excel file <see cref="Worker"/>
        /// </summary>
        /// <returns></returns>
        public static void WriteWorkerToExcel(string fileName, Worker worker)
        {
            Application excelApp;
            Workbooks excelWorkBooks;
            _Workbook excelWorkBook;
            _Worksheet excelWorkSheet;

            // Start Excel and get Application object.
            excelApp = new Application
            {
                Visible = false
            };

            // Get a new workbook and worksheet
            excelWorkBooks = excelApp.Workbooks;
            excelWorkBook = excelWorkBooks.Add();
            excelWorkSheet = (_Worksheet)excelWorkBook.ActiveSheet;

            excelWorkSheet.Cells[1, 1] = "Name";
            excelWorkSheet.Cells[1, 2] = "Id";
            excelWorkSheet.Cells[1, 3] = "Pay Rate";

            excelWorkSheet.Cells[2, 1] = worker.Name;
            excelWorkSheet.Cells[2, 2] = worker.Id;
            excelWorkSheet.Cells[2, 3] = worker.PayRate;

            Range excelRange = excelWorkSheet.Range["A1", "C1"];
            excelRange.Font.Bold = true;
            excelRange.Font.Underline = true;

            excelWorkBook.SaveAs(fileName);
            excelWorkBook.Close();
            excelApp.Quit();

            // Need to release COM objects.These are hidden resources that get allocated
            if (excelWorkSheet != null)
            {
                Marshal.ReleaseComObject(excelWorkSheet);
            }

            if (excelWorkBook != null)
            {
                Marshal.ReleaseComObject(excelWorkBook);
            }

            if (excelWorkBooks != null)
            {
                Marshal.ReleaseComObject(excelWorkBooks);
            }

            if (excelApp != null)
            {
                Marshal.ReleaseComObject(excelApp);
            }

            if (excelRange != null)
            {
                Marshal.ReleaseComObject(excelRange);
            }
        }
        #endregion

        #region Shift Methods
        /// <summary>
        /// Saves Shift data in JSON file format <see cref="Shift"/>
        /// </summary>
        /// <returns></returns>
        public static void SerializeShiftJSON(string fileName, Shift shift)                                                             
        {
            var writer = new FileStream(fileName, FileMode.Create, FileAccess.Write);
            var dataContractJsonSerializer = new DataContractJsonSerializer(typeof(Shift));

            dataContractJsonSerializer.WriteObject(writer, shift);
            writer.Close();
        }

        /// <summary>
        /// Reads Shift from JSON file <see cref="Shift"/>
        /// </summary>
        /// <returns></returns>
        public static Shift DeserializeShiftJSON(string fileName)
        {
            Shift shift;

            var reader = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            var inputSerializer = new DataContractJsonSerializer(typeof(Shift));

            shift = (Shift)inputSerializer.ReadObject(reader);
            reader.Close();

            return shift;
        }

        /// <summary>
        /// Reads Shift from XML file <see cref="Shift"/>
        /// </summary>
        /// <returns></returns>
        public static Shift DeserializeShiftXML(string fileName)
        {
            Shift shift;

            var reader = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            var inputSerializer = new DataContractSerializer(typeof(Shift));

            shift = (Shift)inputSerializer.ReadObject(reader);
            reader.Close();

            return shift;
        }

        /// <summary>
        /// Saves Shift data in XML file format <see cref="Shift"/>
        /// </summary>
        /// <returns></returns>
        public static void SerializeShiftXML(string fileName, Shift shift)
        {
            var writer = new FileStream(fileName, FileMode.Create, FileAccess.Write);
            var serXml = new DataContractSerializer(typeof(Shift)); // For serializing to XML

            serXml.WriteObject(writer, shift);
            writer.Close();
        }

        /// <summary>
        /// Writes Shift to Excel file <see cref="Shift"/>
        /// </summary>
        /// <returns></returns>
        public static void WriteShiftToExcel(string fileName, Shift shift)
        {
            int month = shift.Date.Month;
            int day = shift.Date.Day;
            int year = shift.Date.Year;

            Application excelApp;
            Workbooks excelWorkBooks;
            _Workbook excelWorkBook;
            _Worksheet excelWorkSheet;

            // Start Excel and get Application object.
            excelApp = new Application
            {
                Visible = false
            };

            // Get a new workbook and worksheet
            excelWorkBooks = excelApp.Workbooks;
            excelWorkBook = excelWorkBooks.Add();
            excelWorkSheet = (_Worksheet)excelWorkBook.ActiveSheet;

            excelWorkSheet.Cells[1, 1] = "Worker Id";
            excelWorkSheet.Cells[1, 2] = "Hours Worked";
            excelWorkSheet.Cells[1, 3] = "Date";

            excelWorkSheet.Cells[2, 1] = shift.WorkerID;
            excelWorkSheet.Cells[2, 2] = shift.HoursWorked;
            excelWorkSheet.Cells[2, 3] = month + "/" + day + "/" + year;

            Range excelRange = excelWorkSheet.Range["A1", "C1"];
            excelRange.Font.Bold = true;
            excelRange.Font.Underline = true;

            excelWorkBook.SaveAs(fileName);
            excelWorkBook.Close();
            excelApp.Quit();

            // Need to release COM objects.These are hidden resources that get allocated
            if (excelWorkSheet != null)
            {
                Marshal.ReleaseComObject(excelWorkSheet);
            }

            if (excelWorkBook != null)
            {
                Marshal.ReleaseComObject(excelWorkBook);
            }

            if (excelWorkBooks != null)
            {
                Marshal.ReleaseComObject(excelWorkBooks);
            }

            if (excelApp != null)
            {
                Marshal.ReleaseComObject(excelApp);
            }

            if (excelRange != null)
            {
                Marshal.ReleaseComObject(excelRange);
            }
        }
    }
    #endregion
}