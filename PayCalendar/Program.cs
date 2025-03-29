using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Globalization;
using System.Windows.Media;
using Color = System.Drawing.Color;

namespace PayCalendar
{
    public class Program
    {
        public static string[] Holidays;
        public static void GetHolidays()
        {
            string Alerts = @"D:\PayCalendar Automation\Holidays.txt";
            if (File.Exists(Alerts))
            {
                Holidays = File.ReadAllLines(Alerts);
                for(int i=0;i<=Holidays.Count()-1;i++) 
                {
                    Holidays[i] = Holidays[i] + " 00:00:00";
                    Holidays[i] = Holidays[i].Replace("/", "-");
                }
            }
        }
        public static void Main(string[] args)
        {
            int currentYear = DateTime.Now.Year;
            int financialYearStart = DateTime.Now.Month >= 4 ? currentYear+1 : currentYear;
            int financialYearEnd = financialYearStart + 1;
            string outputFilePath= @"D:\PayCalendar Automation\";
            using (var outputPackage = new ExcelPackage())
            {
                Console.WriteLine("Enter client name:");
                string Client=Console.ReadLine();
                var outputWorksheet = outputPackage.Workbook.Worksheets.Add(Client+" Pay calendar");
                outputWorksheet.Cells[1, 1].Value = "Pay Period";
                outputWorksheet.Cells[1, 2].Value = "Customer Provides Payroll Inputs";
                outputWorksheet.Cells[1, 3].Value = "Payroll Reports to QC";
                outputWorksheet.Cells[1, 4].Value = "Payroll Reports to Customer";
                outputWorksheet.Cells[1, 5].Value = "Customer Approves the Payroll Reports";
                outputWorksheet.Cells[1, 6].Value = "Bank File to Customer";
                outputWorksheet.Cells[1, 7].Value = "Pay Date";
                outputWorksheet.Cells[1, 8].Value = "Statutory Filing Due Date";
                outputWorksheet.Cells[1, 9].Value = "Profession Tax Payment Confirmation";
                outputWorksheet.Cells[1, 10].Value = "Profession Tax return";
                outputWorksheet.Cells[1, 11].Value = "TDS Payment Confirmation";
                outputWorksheet.Cells[1, 12].Value = "TDS Return";
                outputWorksheet.Cells[1, 13].Value = "PF Payment Confirmation";
                outputWorksheet.Cells[1, 14].Value = "ESIC Payment Confirmation";
                outputWorksheet.Cells[1, 15].Value = "ESIC return";
                Console.WriteLine("Last working date of every month in the current financial year:");
                int row = 2;
                for (int month = 4; month <= 12; month++) // April to December of current year
                {
                    DateTime lastWorkingDay = GetLastWorkingDay(financialYearStart, month);
                    Console.WriteLine($"{lastWorkingDay:MMMM yyyy}: {lastWorkingDay:dddd, dd-MM-yyyy}");
                    outputWorksheet.Cells[row, 7].Value = $"{lastWorkingDay:dd-MM-yyyy}";
                    row++;
                }
                for (int month = 1; month <= 3; month++) // January to March of next year
                {
                    DateTime lastWorkingDay = GetLastWorkingDay(financialYearEnd, month);
                    Console.WriteLine($"{lastWorkingDay:MMMM yyyy}: {lastWorkingDay:dddd, dd-MM-yyyy}");
                    outputWorksheet.Cells[row, 7].Value = $"{lastWorkingDay:dd-MM-yyyy}";
                    row++;
                }
                Console.WriteLine("\nEnter Expected Input Provide Date");
                int clientDate = Convert.ToInt32(Console.ReadLine());

                Console.WriteLine("\n\nClient Input Dates for the current financial year are:\n");
                row = 2;
                for (int month = 4; month <= 12; month++) // April to December of current year
                {
                    DateTime lastWorkingDay = InputDay(financialYearStart, month, clientDate);
                    Console.WriteLine($"{lastWorkingDay:MMMM yyyy}: {lastWorkingDay:dddd, dd-MM-yyyy}");
                    outputWorksheet.Cells[row, 1].Value = $"{lastWorkingDay:MMMM yyyy}";
                    outputWorksheet.Cells[row,2].Value= $"{lastWorkingDay:dd-MM-yyyy}";
                    row++;
                }
                for (int month = 1; month <= 3; month++) // January to March of next year
                {
                    DateTime lastWorkingDay = InputDay(financialYearEnd, month, clientDate);
                    Console.WriteLine($"{lastWorkingDay:MMMM yyyy}: {lastWorkingDay:dddd, dd-MM-yyyy}");
                    outputWorksheet.Cells[row, 1].Value = $"{lastWorkingDay:MMMM yyyy}";
                    outputWorksheet.Cells[row, 2].Value = $"{lastWorkingDay:dd-MM-yyyy}";
                    row++;
                }
                row = 2;
                Console.WriteLine("\n\nDates for Reports sent to QC are:\n");
                for (int month = 4; month <= 12; month++) // April to December of current year
                {
                    DateTime lastWorkingDay = GetNextWorkingDay(financialYearStart, month,DateTime.ParseExact(outputWorksheet.Cells[row, 2].Text, "dd-MM-yyyy", CultureInfo.InvariantCulture).Day, 2);
                    Console.WriteLine($"{lastWorkingDay:MMMM yyyy}: {lastWorkingDay:dddd, dd-MM-yyyy}");
                    outputWorksheet.Cells[row, 3].Value = $"{lastWorkingDay:dd-MM-yyyy}";
                    row++;
                }
                for (int month = 1; month <= 3; month++) // January to March of next year
                {
                    DateTime lastWorkingDay = GetNextWorkingDay(financialYearEnd, month, DateTime.ParseExact(outputWorksheet.Cells[row, 2].Text, "dd-MM-yyyy", CultureInfo.InvariantCulture).Day ,2);
                    Console.WriteLine($"{lastWorkingDay:MMMM yyyy}: {lastWorkingDay:dddd, dd-MM-yyyy}");
                    outputWorksheet.Cells[row, 3].Value = $"{lastWorkingDay:dd-MM-yyyy}";
                    row++;
                }
                row = 2;
                Console.WriteLine("\n\nDates for Reports sent to Client after QC are:\n");
                for (int month = 4; month <= 12; month++) // April to December of current year
                {
                    DateTime lastWorkingDay = GetNextWorkingDay(financialYearStart, month, DateTime.ParseExact(outputWorksheet.Cells[row, 3].GetValue<string>(), "dd-MM-yyyy", CultureInfo.InvariantCulture).Day ,1);
                    Console.WriteLine($"{lastWorkingDay:MMMM yyyy}: {lastWorkingDay:dddd, dd-MM-yyyy}");
                    outputWorksheet.Cells[row, 4].Value = $"{lastWorkingDay:dd-MM-yyyy}";
                    row++;
                }
                for (int month = 1; month <= 3; month++) // January to March of next year
                {
                    DateTime lastWorkingDay = GetPreviousWorkingDay(financialYearEnd, month, DateTime.ParseExact(outputWorksheet.Cells[row, 3].Text, "dd-MM-yyyy", CultureInfo.InvariantCulture).Day ,1);
                    Console.WriteLine($"{lastWorkingDay:MMMM yyyy}: {lastWorkingDay:dddd, dd-MM-yyyy}");
                    outputWorksheet.Cells[row, 4].Value = $"{lastWorkingDay:dd-MM-yyyy}";
                    row++;
                }
                row = 2;
                Console.WriteLine("\n\nDates Bank File to client:\n");
                for (int month = 4; month <= 12; month++) // April to December of current year
                {
                    DateTime lastWorkingDay = GetPreviousWorkingDay(financialYearStart, month, DateTime.ParseExact(outputWorksheet.Cells[row, 7].Text, "dd-MM-yyyy", CultureInfo.InvariantCulture).Day , -2);
                    Console.WriteLine($"{lastWorkingDay:MMMM yyyy}: {lastWorkingDay:dddd, dd-MM-yyyy}");
                    outputWorksheet.Cells[row, 6].Value = $"{lastWorkingDay:dd-MM-yyyy}";
                    row++;
                }
                for (int month = 1; month <= 3; month++) // January to March of next year
                {
                    DateTime lastWorkingDay = GetPreviousWorking5Day(financialYearEnd, month, DateTime.ParseExact(outputWorksheet.Cells[row, 7].Text, "dd-MM-yyyy", CultureInfo.InvariantCulture).Day , -2);
                    Console.WriteLine($"{lastWorkingDay:MMMM yyyy}: {lastWorkingDay:dddd, dd-MM-yyyy}");
                    outputWorksheet.Cells[row, 6].Value = $"{lastWorkingDay:dd-MM-yyyy}";
                    row++;
                }
                row = 2;
                Console.WriteLine("\n\nDates for Client Approves Reports:\n");
                for (int month = 4; month <= 12; month++) // April to December of current year
                {
                    DateTime lastWorkingDay = GetPreviousWorkingDay(financialYearStart, month, DateTime.ParseExact(outputWorksheet.Cells[row, 6].Text, "dd-MM-yyyy", CultureInfo.InvariantCulture).Day , -1);
                    Console.WriteLine($"{lastWorkingDay:MMMM yyyy}: {lastWorkingDay:dddd, dd-MM-yyyy}");
                    outputWorksheet.Cells[row, 5].Value = $"{lastWorkingDay:dd-MM-yyyy} "+ $"{lastWorkingDay:ddd}";
                    row++;
                }
                for (int month = 1; month <= 3; month++) // January to March of next year
                {
                    DateTime lastWorkingDay = GetPreviousWorkingDay(financialYearEnd, month, DateTime.ParseExact(outputWorksheet.Cells[row, 6].Text, "dd-MM-yyyy", CultureInfo.InvariantCulture).Day , -1);
                    Console.WriteLine($"{lastWorkingDay:MMMM yyyy}: {lastWorkingDay:dddd, dd-MM-yyyy}");
                    outputWorksheet.Cells[row, 5].Value = $"{lastWorkingDay:dd-MM-yyyy}";
                    row++;
                }
                row = 2;
                Console.WriteLine("\n\nDates for Statutory filling are:\n");
                for (int month = 4; month <= 12; month++) // April to December of current year
                {
                    DateTime lastWorkingDay = GetNextWorkingDay(financialYearStart, month, DateTime.ParseExact(outputWorksheet.Cells[row, 7].Text, "dd-MM-yyyy", CultureInfo.InvariantCulture).Day , 2);
                    Console.WriteLine($"{lastWorkingDay:MMMM yyyy}: {lastWorkingDay:dddd, dd-MM-yyyy}");
                    outputWorksheet.Cells[row, 8].Value = $"{lastWorkingDay:dd-MM-yyyy}";
                    row++;
                }
                for (int month = 1; month <= 3; month++) // January to March of next year
                {
                    DateTime lastWorkingDay = GetNextWorkingDay(financialYearEnd, month, DateTime.ParseExact(outputWorksheet.Cells[row, 7].Text, "dd-MM-yyyy", CultureInfo.InvariantCulture).Day , 2);
                    Console.WriteLine($"{lastWorkingDay:MMMM yyyy}: {lastWorkingDay:dddd, dd-MM-yyyy}");
                    outputWorksheet.Cells[row, 8].Value = $"{lastWorkingDay:dd-MM-yyyy}";
                    row++;
                }
                row = 2;
                Console.WriteLine("\n\nProfession Tax Payment Confirmation are:\n");
                for (int month = 4; month <= 12; month++) // April to December of current year
                {
                    DateTime lastWorkingDay = GetPreviousWorkingDay(financialYearStart, month+1, 10,0);
                    Console.WriteLine($"{lastWorkingDay:MMMM yyyy}: {lastWorkingDay:dddd, dd-MM-yyyy}");
                    outputWorksheet.Cells[row, 9].Value = $"{lastWorkingDay:dd-MM-yyyy}";
                    row++;
                }
                for (int month = 1; month <= 3; month++) // January to March of next year
                {
                    DateTime lastWorkingDay = GetPreviousWorkingDay(financialYearEnd, month + 1, 10,0);
                    Console.WriteLine($"{lastWorkingDay:MMMM yyyy}: {lastWorkingDay:dddd, dd-MM-yyyy}");
                    outputWorksheet.Cells[row, 9].Value = $"{lastWorkingDay:dd-MM-yyyy}";
                    row++;
                }
                row = 2;
                Console.WriteLine("\n\nDates for Profession Tax return are:\n");
                for (int month = 4; month <= 12; month++) // April to December of current year
                {
                    DateTime lastWorkingDay = GetPreviousWorkingDay(financialYearStart, month+1, 15, 0);
                    Console.WriteLine($"{lastWorkingDay:MMMM yyyy}: {lastWorkingDay:dddd, dd-MM-yyyy}");
                    outputWorksheet.Cells[row, 10].Value = $"{lastWorkingDay:dd-MM-yyyy}";
                    row++;
                }
                for (int month = 1; month <= 3; month++) // January to March of next year
                {
                    DateTime lastWorkingDay = GetPreviousWorkingDay(financialYearEnd, month + 1, 15, 0);
                    Console.WriteLine($"{lastWorkingDay:MMMM yyyy}: {lastWorkingDay:dddd, dd-MM-yyyy}");
                    outputWorksheet.Cells[row, 10].Value = $"{lastWorkingDay:dd-MM-yyyy}";
                    row++;
                }
                row = 2;
                Console.WriteLine("\n\nDates for TDS Payment Confirmation are:\n");
                for (int month = 4; month <= 12; month++) // April to December of current year
                {
                    DateTime lastWorkingDay = GetPreviousWorkingDay(financialYearStart, month+1, 10, 0);
                    Console.WriteLine($"{lastWorkingDay:MMMM yyyy}: {lastWorkingDay:dddd, dd-MM-yyyy}");
                    outputWorksheet.Cells[row, 11].Value = $"{lastWorkingDay:dd-MM-yyyy}";
                    row++;
                }
                for (int month = 1; month <= 3; month++) // January to March of next year
                {
                    DateTime lastWorkingDay = GetPreviousWorkingDay(financialYearEnd, month+1, 10, 0);
                    Console.WriteLine($"{lastWorkingDay:MMMM yyyy}: {lastWorkingDay:dddd, dd-MM-yyyy}");
                    outputWorksheet.Cells[row, 11].Value = $"{lastWorkingDay:dd-MM-yyyy}";
                    row++;
                }
                row = 2;
                Console.WriteLine("\n\nDates for PF Payment Confirmation are:\n");
                for (int month = 4; month <= 12; month++) // April to December of current year
                {
                    DateTime lastWorkingDay = GetPreviousWorkingDay(financialYearStart, month+1, 12, 0);
                    Console.WriteLine($"{lastWorkingDay:MMMM yyyy}: {lastWorkingDay:dddd, dd-MM-yyyy}");
                    outputWorksheet.Cells[row, 13].Value = $"{lastWorkingDay:dd-MM-yyyy}";
                    row++;
                }
                for (int month = 1; month <= 3; month++) // January to March of next year
                {
                    DateTime lastWorkingDay = GetPreviousWorkingDay(financialYearEnd, month+1, 12, 0);
                    Console.WriteLine($"{lastWorkingDay:MMMM yyyy}: {lastWorkingDay:dddd, dd-MM-yyyy}");
                    outputWorksheet.Cells[row, 13].Value = $"{lastWorkingDay:dd-MM-yyyy}";
                    row++;
                }
                row = 2;
                Console.WriteLine("\n\nDates for ESIC Payment Confirmation are:\n");
                for (int month = 4; month <= 12; month++) // April to December of current year
                {
                    DateTime lastWorkingDay = GetPreviousWorkingDay(financialYearStart, month + 1, 12, 0);
                    Console.WriteLine($"{lastWorkingDay:MMMM yyyy}: {lastWorkingDay:dddd, dd-MM-yyyy}");
                    outputWorksheet.Cells[row, 14].Value = $"{lastWorkingDay:dd-MM-yyyy}";
                    row++;
                }
                for (int month = 1; month <= 3; month++) // January to March of next year
                {
                    DateTime lastWorkingDay = GetPreviousWorkingDay(financialYearEnd, month + 1, 12, 0);
                    Console.WriteLine($"{lastWorkingDay:MMMM yyyy}: {lastWorkingDay:dddd, dd-MM-yyyy}");
                    outputWorksheet.Cells[row, 14].Value = $"{lastWorkingDay:dd-MM-yyyy}";
                    row++;
                }
                row = 2;
                Console.WriteLine("\n\nDates for ESIC return are:\n");
                for (int month = 4; month <= 12; month++) // April to December of current year
                {
                    DateTime lastWorkingDay = GetPreviousWorkingDay(financialYearStart, month + 1, 12, 0);
                    Console.WriteLine($"{lastWorkingDay:MMMM yyyy}: {lastWorkingDay:dddd, dd-MM-yyyy}");
                    outputWorksheet.Cells[row, 15].Value = $"{lastWorkingDay:dd-MM-yyyy}";
                    row++;
                }
                for (int month = 1; month <= 3; month++) // January to March of next year
                {
                    DateTime lastWorkingDay = GetPreviousWorkingDay(financialYearEnd, month + 1, 12, 0);
                    Console.WriteLine($"{lastWorkingDay:MMMM yyyy}: {lastWorkingDay:dddd, dd-MM-yyyy}");
                    outputWorksheet.Cells[row, 15].Value = $"{lastWorkingDay:dd-MM-yyyy}";
                    row++;
                }
                //Console.WriteLine("\n\nEnter Pay Date:\n");
                //int PayDate = Convert.ToInt32(Console.ReadLine());
                //Console.WriteLine("\n\nStatutory Filing Due Date:\n");
                //for (int month = 4; month <= 12; month++) // April to December of current year
                //{
                //    DateTime lastWorkingDay = GetPreviousWorkingDay(financialYearStart+1, month, GetLastWorkingDay(financialYearStart+1, month).Day+2);
                //    Console.WriteLine($"{lastWorkingDay:MMMM yyyy}: {lastWorkingDay:dddd, dd-MM-yyyy}");
                //}
                //for (int month = 1; month <= 3; month++) // January to March of next year
                //{
                //    DateTime lastWorkingDay = GetPreviousWorkingDay(financialYearStart+1, month, GetLastWorkingDay(financialYearStart+1, month).Day + 2);
                //    Console.WriteLine($"{lastWorkingDay:MMMM yyyy}: {lastWorkingDay:dddd, dd-MM-yyyy}");
                //}

                //Console.WriteLine("\n\nStatutory Filing Due Date:\n");
                //for (int month = 4; month <= 12; month++) // April to December of current year
                //{
                //    DateTime lastWorkingDay = GetPreviousWorkingDay(financialYearStart+1, month, GetLastWorkingDay(financialYearStart+1, month).Day+2);
                //    Console.WriteLine($"{lastWorkingDay:MMMM yyyy}: {lastWorkingDay:dddd, dd-MM-yyyy}");
                //}

                //for (int month = 1; month <= 3; month++) // January to March of next year
                //{
                //    DateTime lastWorkingDay = GetPreviousWorkingDay(financialYearStart+1, month, GetLastWorkingDay(financialYearStart+1, month).Day + 2);
                //    Console.WriteLine($"{lastWorkingDay:MMMM yyyy}: {lastWorkingDay:dddd, dd-MM-yyyy}");
                //}
                //int fixedDate = 30; // Change this to test with different dates

                //for (int month = 1; month <= 3; month++) // January to March of next year
                //{
                //    DateTime lastWorkingDay = GetPreviousWorkingDay(financialYearEnd+1, month, fixedDate);
                //    Console.WriteLine($"{lastWorkingDay:MMMM yyyy}: {lastWorkingDay:dddd, dd-MM-yyyy}");
                //}

                //Console.WriteLine("\nLast working date of every month in the current financial year (moving to Monday if weekend):");
                //for (int month = 4; month <= 12; month++)
                //{
                //    DateTime lastWorkingDay = GetNextWorkingDay(financialYearStart+1, month, fixedDate);
                //    Console.WriteLine($"{lastWorkingDay:MMMM yyyy}: {lastWorkingDay:dddd, dd-MM-yyyy}");
                //}

                //for (int month = 1; month <= 3; month++)
                //{
                //    DateTime lastWorkingDay = GetNextWorkingDay(financialYearEnd+1, month, fixedDate);
                //    Console.WriteLine($"{lastWorkingDay:MMMM yyyy}: {lastWorkingDay:dddd, dd-MM-yyyy}");
                //}
                
                // Apply Formatting
                
                for (int i = 1; i <= 100; i++) // Assuming 100 rows
                {
                    outputWorksheet.Row(i).Height = 27; // Set row height
                    outputWorksheet.Column(i).Width = 20;
                }

                // Increase column width for all columns (adjust as needed)
                for (int i = 1; i <= 50; i++) // Assuming 50 columns (A to AX)
                {
                    outputWorksheet.Column(i).Width = 20; // Set column width
                    outputWorksheet.Cells[1, i].Style.WrapText = true;
                }
                DateTime d;
                d= GetPreviousWorkingDay(currentYear, 4, 15, 0); 
                outputWorksheet.Cells[2, 12].Value = $"{d:dd-MM-yyyy}";
                d = GetPreviousWorkingDay(currentYear, 7, 15, 0);
                outputWorksheet.Cells[5, 12].Value = $"{d:dd-MM-yyyy}";
                d = GetPreviousWorkingDay(currentYear, 10, 15, 0);
                outputWorksheet.Cells[8, 12].Value = $"{d:dd-MM-yyyy}";
                d = GetPreviousWorkingDay(currentYear + 1, 1, 15, 0);
                outputWorksheet.Cells[11, 12].Value = $"{d:dd-MM-yyyy}";
                outputWorksheet.Cells[outputWorksheet.Dimension.Address].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                outputWorksheet.Cells[outputWorksheet.Dimension.Address].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                int endRow = outputWorksheet.Dimension.End.Row;
                int endCol = outputWorksheet.Dimension.End.Column;
                //for(int i=2;i<=endRow;i++) 
                //{
                //    for (int j = 2; j <= endCol; j++)
                //    {
                //        if (outputWorksheet.Cells[i, j].Text != "") { 
                //        DateTime lastWorkingDay = outputWorksheet.Cells[i,j].GetValue<DateTime>();
                //        //outputWorksheet.Cells[i,j].Value=$"{lastWorkingDay:dd-MM-yyyy} "+ $"{lastWorkingDay:ddd}";
                //        }
                //    }
                //}
                for (int j = endCol; j >= 2; j--)
                {
                    outputWorksheet.InsertColumn(j,1);
                    for (int i = 2; i <= endRow; i++)
                    {
                        if (outputWorksheet.Cells[i, j-1].Text != "")
                        {
                            outputWorksheet.Cells[1, j].Value="Day";
                            DateTime lastWorkingDay = outputWorksheet.Cells[i, j-1].GetValue<DateTime>();
                            outputWorksheet.Cells[i, j].Value = $"{lastWorkingDay:ddd}";
                        }
                    }
                }
                int oColumn = outputWorksheet.Dimension.End.Column;
                int oRow = outputWorksheet.Dimension.End.Row;
                using (var headerRange = outputWorksheet.Cells[1, 1, 1, oColumn])
                {
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    headerRange.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                    headerRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    headerRange.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                }

                using (var dataRange = outputWorksheet.Cells[1, 1, oRow, oColumn])
                {
                    dataRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    dataRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    dataRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    dataRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                }
                outputWorksheet.Cells[outputWorksheet.Dimension.Address].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                outputWorksheet.Cells[outputWorksheet.Dimension.Address].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                 endRow = outputWorksheet.Dimension.End.Row;
                 endCol = outputWorksheet.Dimension.End.Column;
                for (int j = endCol; j >= 2; j--)
                {
                    if (outputWorksheet.Cells[1, j].Text == "")
                    {
                       outputWorksheet.DeleteColumn(j);
                    }
                }
                outputPackage.SaveAs(new FileInfo(outputFilePath + Client+ " Pay Calendar.xlsx"));
                Console.WriteLine("Operation successful press Enter to exit. ");
                Console.ReadLine();
            }
        }
        static DateTime GetPreviousWorkingDay(int year, int month, int day, int diff)
        {
            DateTime givenDay;

            try
            {
                givenDay = new DateTime(year, month, day);
                while (Holidays.Contains(givenDay.ToString()))
                {
                    Console.WriteLine("Holiday detected.");
                    givenDay = givenDay.AddDays(-1);
                }
            }
            catch
            {
                // If day exceeds the month's end, move to the next month
                month++;
                if (month > 12)
                {
                    month = 1;
                    year++;
                }
                givenDay = new DateTime(year, month, day);
            }
            // Apply the difference (add or subtract days)
            givenDay = givenDay.AddDays(diff);

            //if (givenDay.DayOfWeek == DayOfWeek.Saturday)
            //    givenDay = givenDay.AddDays(-1);
            if (givenDay.DayOfWeek == DayOfWeek.Sunday)
                givenDay = givenDay.AddDays(-1);

            return givenDay;
            }
        static DateTime GetPreviousWorking5Day(int year, int month, int day, int diff)
        {
            DateTime givenDay;

            try
            {
                givenDay = new DateTime(year, month, day);
            }
            catch
            {
                // If day exceeds the month's end, move to the next month
                month++;
                if (month > 12)
                {
                    month = 1;
                    year++;
                }
                givenDay = new DateTime(year, month, 1);
            }
            // Apply the difference (add or subtract days)
            givenDay = givenDay.AddDays(diff);

            if (givenDay.DayOfWeek == DayOfWeek.Saturday)
                givenDay = givenDay.AddDays(-1);
            if (givenDay.DayOfWeek == DayOfWeek.Sunday)
                givenDay = givenDay.AddDays(-2);
            return givenDay;
        }
        // Method to return last working day (move to Monday if weekend)
        static DateTime GetNextWorkingDay(int year, int month, int day, int diff)
        {
            GetHolidays();
            DateTime givenDay;
            try
            {
                givenDay = new DateTime(year, month, day);
                while (Holidays.Contains(givenDay.ToString())) 
                {
                    Console.WriteLine("Holiday detected.");
                    givenDay = givenDay.AddDays(1);
                }
            }
            catch
            {
                // If day exceeds the month's end, move to the next month
                month++;
                if (month > 12)
                {
                    month = 1;
                    year++;
                }
                givenDay = new DateTime(year, month, 1);
            }
            // Apply the difference (add or subtract days)
            givenDay = givenDay.AddDays(diff);

            //if (givenDay.DayOfWeek == DayOfWeek.Saturday)
            //    givenDay = givenDay.AddDays(2);
            if (givenDay.DayOfWeek == DayOfWeek.Sunday)
                givenDay = givenDay.AddDays(1);
            if (givenDay.DayOfWeek-1 == DayOfWeek.Sunday)
                givenDay = givenDay.AddDays(1);
            return givenDay;
        }
        static DateTime GetNextWorkingDayBackwards(int year, int month, int day)
        {
            DateTime givenDay;
            try
            {
                givenDay = new DateTime(year, month, day);
            }
            catch
            {
                givenDay = new DateTime(year, month, DateTime.DaysInMonth(year, month));
            }

            //if (givenDay.DayOfWeek == DayOfWeek.Saturday)
            //    givenDay = givenDay.AddDays(2);
            if (givenDay.DayOfWeek == DayOfWeek.Sunday)
                givenDay = givenDay.AddDays(1);
            return givenDay;
        }
        static DateTime InputDay(int year, int month, int day)
        {
            GetHolidays();
            DateTime givenDay;
            try
            {
                givenDay = new DateTime(year, month, day);
                while (Holidays.Contains(givenDay.ToString()))
                {
                    Console.WriteLine("Holiday detected.");
                    givenDay = givenDay.AddDays(1);
                }
            }
            catch
            {
                givenDay = new DateTime(year, month, DateTime.DaysInMonth(year, month));
            }

            if (givenDay.DayOfWeek == DayOfWeek.Saturday)
                givenDay = givenDay.AddDays(2);
            if (givenDay.DayOfWeek == DayOfWeek.Sunday)
                givenDay = givenDay.AddDays(1);
            return givenDay;
        }
        static DateTime GetLastWorkingDay(int year, int month)
        {
            GetHolidays();
            DateTime lastDay = new DateTime(year, month, DateTime.DaysInMonth(year, month));
            while (Holidays.Contains(lastDay.ToString()))
            {
                Console.WriteLine("Holiday detected.");
                lastDay = lastDay.AddDays(-1);
            }
            if (lastDay.DayOfWeek == DayOfWeek.Saturday)
            {
                lastDay = lastDay.AddDays(-1);
                return lastDay;
            }
            if (lastDay.DayOfWeek == DayOfWeek.Sunday)
            {
                lastDay = lastDay.AddDays(-2);
                return lastDay;
            }
            return lastDay;
            //while (lastDay.DayOfWeek == DayOfWeek.Saturday || lastDay.DayOfWeek == DayOfWeek.Sunday)
            //{
            //    lastDay = lastDay.AddDays(-1);
            //}

            //return lastDay;
        }
    }
}
