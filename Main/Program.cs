//******************************************************
// File: Program.cs
//
// Purpose: Contains the class definition for Program.
//          Program will read the Department from 
//          a JSON and XML file, write Department
//          to a JSON and XML file, write Department
//          to an Excel file, display Department
//          data on the screen, and find a worker.
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

namespace Main
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
            var department = new Department();
            int choice;
            
            do
            {
                Console.WriteLine("Department Menu");
                Console.WriteLine("---------------");
                Console.WriteLine("1 - Read department from JSON file");
                Console.WriteLine("2 - Read department from XML file");
                Console.WriteLine("3 - Write department to JSON file");
                Console.WriteLine("4 - Write department to XML file");
                Console.WriteLine("5 - Write department to Excel file");
                Console.WriteLine("6 - Display all department data on screen");
                Console.WriteLine("7 - Find worker");
                Console.WriteLine("8 - Exit");
                Console.Write("Enter Choice: ");

                int userInput = Convert.ToInt32(Console.ReadLine());

                choice = userInput;

                switch (choice)
                {
                    case 1:
                        try
                        {
                            string fileName = UserEnteringFileName();
                            department = DeserializeDepartmentJSON(fileName);
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
                            SerializeDepartmentJSON(fileName, department);
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
                            WriteDepartmentToExcel(fileName, department);
                        }
                        catch (Exception exception)
                        {
                            Console.WriteLine($"Error: {exception.Message}");
                        }
                        break;

                    case 6:
                        Console.WriteLine(department.ToString());
                        break;

                    case 7:
                        try
                        {
                            Console.WriteLine("Enter worker id: ");
                            int userEntered = Convert.ToInt32(Console.ReadLine());
                        }
                        catch (Exception exception)
                        {
                            Console.WriteLine($"Error: {exception.Message}");
                        }
                        break;

                }
            } while (choice != 8);
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

        /// <summary>
        /// Reads Department from JSON file <see cref="Department"/>
        /// </summary>
        /// <returns></returns>
        public static Department DeserializeDepartmentJSON(string fileName)
        {
            Department department;

            var reader = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            var inputSerializer = new DataContractJsonSerializer(typeof(Department));

            department = (Department)inputSerializer.ReadObject(reader);
            reader.Close();

            return department;
        }

        /// <summary>
        /// Saves Department data in JSON file format <see cref="Department"/>
        /// </summary>
        /// <returns></returns>
        public static void SerializeDepartmentJSON(string fileName, Department department)
        {
            var writer = new FileStream(fileName, FileMode.Create, FileAccess.Write);
            var dataContractJsonSerializer = new DataContractJsonSerializer(typeof(Department));

            dataContractJsonSerializer.WriteObject(writer, department);
            writer.Close();
        }





        /// <summary>
        /// Writes Worker to Excel file <see cref="Department"/>
        /// </summary>
        /// <returns></returns>
        public static void WriteDepartmentToExcel(string fileName, Department department)
        {
            Application excelApp;
            Workbooks excelWorkBooks;
            _Workbook excelWorkBook;
            _Worksheet excelWorkSheet;

            //Start Excel and get Application object.
            excelApp = new Application
            {
                Visible = false
            };

            //Get a new workbook and worksheet
            excelWorkBooks = excelApp.Workbooks;
            excelWorkBook = excelWorkBooks.Add();
            excelWorkSheet = (_Worksheet)excelWorkBook.ActiveSheet;

            // nested for loop for inputing data in to cells
            excelWorkSheet.Cells[1, 1] = "Department Name";
            excelWorkSheet.Cells[1, 2] = "Technical Support";
            excelWorkSheet.Cells[3, 1] = "Workers";
            excelWorkSheet.Cells[3, 6] = "Shifts";
            excelWorkSheet.Cells[5, 1] = "Name";
            excelWorkSheet.Cells[5, 2] = "Id";
            excelWorkSheet.Cells[5, 3] = "Pay Rate";
            excelWorkSheet.Cells[5, 6] = "Worker Id";
            excelWorkSheet.Cells[5, 7] = "Hours Worked";
            excelWorkSheet.Cells[5, 8] = "Date";

            int workerRow = 6;
            foreach (Worker worker in department.Workers)
            {
                excelWorkSheet.Cells[workerRow, 1] = worker.Name;
                excelWorkSheet.Cells[workerRow, 2] = worker.Id;
                excelWorkSheet.Cells[workerRow, 3] = worker.PayRate;
                ++workerRow;
            }
            int shiftRow = 6;
            foreach (Shift shift in department.Shifts)
            {
                excelWorkSheet.Cells[shiftRow, 6] = shift.WorkerID;
                excelWorkSheet.Cells[shiftRow, 7] = shift.HoursWorked;
                excelWorkSheet.Cells[shiftRow, 8] = shift.Date;
                ++shiftRow;
            }

            ////Range excelRange = excelWorkSheet.Range["A1"];
            //excelRange.Font.Bold = true;
            //excelRange.Font.Underline = true;

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

            //if (excelRange != null)
            //{
            //    Marshal.ReleaseComObject(excelRange);
            //}
        }

    }
}