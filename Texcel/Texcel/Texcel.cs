using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Reflection;
using Microsoft.Office.Interop.Excel;

namespace Texcel
{
    public class rawParser
    {
        static void Main()
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Console.Write("Making excel object...\n");
            if (excel == null)
            {
                Console.Write("Excel could not be started.");
                return;
            }
            else
            {
                Console.Write("Complete\n");
            }
            Workbook wb = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            Console.Write("Making workbook...\n");
            if (wb == null)
            {
                Console.Write("Worksheet could not be created.");
            }
            else
            {
                Console.Write("Complete\n");
            }
            Console.Write("Select mode: <S>ingle file or <M>ultiple files\n: ");
            Boolean val = false;
            int num = 2;
            while (!val)
            {
                char op = char.ToLower(Convert.ToChar(Console.ReadLine()));
                if (op == 's')
                {
                    val = true;
                    Console.Write("Enter Subject Number to process\n: ");
                    int run = int.Parse(Console.ReadLine());
                    excelWrite(run, 0, excel, wb);
                }
                else if (op == 'm')
                {
                    val = true;
                    Console.Write("Select sub-mode: <R>ange of files or <L>ist of files\n: ");
                    char subop = char.ToLower(Convert.ToChar(Console.ReadLine()));

                    if (subop == 'r')
                    {
                        Boolean valid = false;
                        int low, high;
                        while (!valid)
                        {
                            Console.Write("To process multiple files, please enter a range of\nnumbers to process.\n");
                            Console.Write("Enter lower number to process\n: ");
                            low = int.Parse(Console.ReadLine());
                            Console.Write("Enter upper number to process\n: ");
                            high = int.Parse(Console.ReadLine());
                            num = high + 1;
                            if (low <= high)
                            {
                                valid = true;
                                for (int x = high; x >= low; x--)
                                {
                                    excelWrite(x, 0, excel, wb);
                                }
                            }
                            else { Console.Write("Error: Low must be smaller than or equal to High\n"); }
                        }
                    }
                    else if (subop == 'l')
                    {
                        Boolean reading = true;
                        List<int> nums = new List<int>();
                        Console.Write("Enter subject numbers one at a time.\nWhen finished, enter -1.\n");
                        while (reading)
                        {
                            Console.Write(": ");
                            int temp = int.Parse(Console.ReadLine());
                            if (temp == -1) { reading = false; }
                            else if (temp != -1)
                            {
                                nums.Add(temp);
                            }
                        }
                        num--;
                        foreach (int x in nums)
                        {
                            num++;
                            excelWrite(x, 0, excel, wb);
                        }
                    }
                }
                else { Console.Write("Sorry, that isn't a valid command.");  }
            }
            Console.Write("Press <enter> to complete");
            int finish = Console.Read();
            wb.Sheets[num].Name = "Data";
            excel.Visible = true;
        }

        public static void excelWrite(int subj, int row, Microsoft.Office.Interop.Excel.Application excel, Workbook wb)
        {
            string subject = "subject" + subj + ".txt";
            Console.Write("Finding " + subject + "...\n");

            StreamReader sr = new StreamReader(subject);
            Console.Write("File found!\n");
            
            string raw = sr.ReadToEnd();
            string[] scans = raw.Split(' ');
            int x = 0;
            foreach (string s in scans)
            {
                string scan = s.Trim();
                scans[x] = scan;
                x++;
            }

            Worksheet ws = wb.Worksheets.Add();
            ws.Name = "Subject " + subj;
            Console.Write("Making worksheet "+subj+"...\n");
            if (ws == null)
            {
                Console.Write("Worksheet could not be created.");
            }
            else
            {
                Console.Write("Complete\n");
            }
            if (row <= 0) { row = 1; }
            int col = 1;
            if (row == 1)
            {
                ws.Cells[row, 1] = "Heading";
                ws.Cells[row, 2] = "Pitch";
                ws.Cells[row, 3] = "Roll";
                ws.Cells[row, 4] = "Pos-X";
                ws.Cells[row, 5] = "Pos-Y";
                ws.Cells[row, 6] = "Pos-Z";
                ws.Cells[row, 7] = "Vel-X";
                ws.Cells[row, 8] = "Vel-Y";
                ws.Cells[row, 9] = "Vel-Z";
                ws.Cells[row, 10] = "Hour";
                ws.Cells[row, 11] = "Minute";
                ws.Cells[row, 12] = "Second";
                ws.Cells[row, 13] = "Milli";
                row++;
            }
            Console.Write("Beginning File Write\n");
            foreach (string a in scans)
            {
                if (a != scans[0] && a != scans[1] && a != scans[2] && a != scans[scans.Length - 1] && a != scans[scans.Length - 2] && a != scans[scans.Length - 3] && a != scans[scans.Length - 4])
                {
                    if (a != ":" && a != ";" && a != ",")
                    {
                        ws.Cells[row, col] = Convert.ToInt64(a);
                        col++;
                    }
                    if (a == ";")
                    {
                        col = 1;
                        row++;
                    }
                }
            }
            Console.Write("File Write "+subj+" Complete\n");
        }
    }
}
