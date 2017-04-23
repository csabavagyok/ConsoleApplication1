using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Globalization;

namespace ConsoleApplication1
{
    public abstract class TryOpenXlsx
    {
        public static void TryOpenFile()
        { 
            try
            {
                //welcome messagee
                Console.WriteLine("Üdvözlöm.\nA munkaóra elszámolás ellenőrzése azonnal kezdődik...");

                //get and print last access time
                string filePath = @"C:\Downloaded Torrents\Kispest.xlsx";
                var lastAccessDate = File.GetLastWriteTime(filePath);
                Console.WriteLine("A jelenleg használt Excel fájl utolsó módosításának dátuma:" + lastAccessDate);

                //declare variables
                Excel.Application xlApp = new Excel.Application();
                xlApp.Visible = false;
                Excel.Workbook xlWB;
                Excel.Worksheet xlWS;
                Excel.Range range;
                int counter = 0;
                byte errorCounter = 0;
                byte goodCounter = 0;
                decimal difference = 0;
                List<string> errorNames = new List<string>();

                //open file and worksheet #1
                xlWB = xlApp.Workbooks.Open(filePath, 0, true, 5, null, null, true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", true, false, 0, true, 1, 0);
                Console.WriteLine("Az állomány megnyitása sikeres.");
                xlWS = (Excel.Worksheet)xlWB.Sheets.Item[1];
                Console.WriteLine("A munkalap megnyitása sikeres.");
                Console.WriteLine("Az elemek átvizsgálása.");

                //iterate through selected range items
                for (int i = 4; i < 61; i++)
                {
                    string str = "A" + i;
                    range = xlWS.Range[str];
                    var name = range.Text;
                    if (name != "")
                    {
                        //display name to screen
                        Console.WriteLine(name + " munkaóra elszámolásának kiértékelése...");
                        counter++;
                        

                        //display if pre-calculated pay-off was correct
                        //declare variables
                        var hoursOfMonth = xlWS.Range["D2"].Text;    //havi óraszám
                        var hours = decimal.Parse(xlWS.Range["D" + i].Text);        //szerződés szerinti óraszám
                        var workHours = decimal.Parse(xlWS.Range["E" + i].Text);    //teljesített óraszám
                        var sick = 0;
                        var sickCheck = xlWS.Range["G" + i].Text;         //betegség napok
                        if ( sickCheck != "")
                        { sick = int.Parse(sickCheck); }
                        var dayOff = 0;
                        var dayOffCheck = xlWS.Range["I" + i].Text;
                        if ( dayOffCheck != "")
                        { dayOff = int.Parse(dayOffCheck); }
                        var lecture = 0;
                        var lectureCheck = xlWS.Range["K" + i].Text;
                        if( lectureCheck != "")
                        { lecture = int.Parse(lectureCheck); }
                        var overTime = decimal.Parse(xlWS.Range["O" + i].Text);     //túlóra
                        var afterNoon = decimal.Parse(xlWS.Range["P" + i].Text);    //délutáni pótlék
                        var nightTime = decimal.Parse(xlWS.Range["Q" + i].Text);    //éjszakai pótlék
                        decimal extra = 0;
                        var extraCheck = xlWS.Range["R" + i].Text;
                        if ( extraCheck != "" )
                        { extra = decimal.Parse(extraCheck); }
                        var numberValue = hours / decimal.Parse(hoursOfMonth);

                        //Console.WriteLine("A dolgozó havi kötelező óraszáma: " + hours);
                        //Console.WriteLine("A dolgozó beteg órái: " + sick * numberValue * 8);
                        Console.WriteLine("Létszámérték: " + numberValue);
                        //Console.WriteLine("Beteg óra: " + sick*numberValue*8);
                        //Console.WriteLine("Szabadság óra: " + dayOff*numberValue*8);
                        //Console.WriteLine("Tanulmány óra: " + lecture);
                        Console.WriteLine("Túlóra: " + overTime);
                        //Console.WriteLine("200% :" + extra);
                        //Console.WriteLine("Telj.óra: " + workHours);
                        var determineOverTime = 
                            workHours -
                            hours +
                            ((sick * numberValue * 8) +
                            (dayOff * numberValue * 8) +
                            lecture -
                            extra);
                        Console.WriteLine("SZÁMOLÁS: " + determineOverTime);
                        //determine if pay-off was correct
                        if (determineOverTime == 
                            overTime)
                        {
                            goodCounter++;
                            Console.ForegroundColor = ConsoleColor.DarkGreen;
                            Console.WriteLine("Elszámolás RENDBEN.");
                            Console.ForegroundColor = ConsoleColor.White;
                        }
                        else
                        {
                            errorNames.Add(name);
                            errorCounter++;
                            difference = overTime - determineOverTime;
                            Console.WriteLine("AZ ALÁBBI ÓRASZÁM ELTÉRÉST ÉSZLELTEM: " + difference);
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("HIBA AZ ELSZÁMOLÁSBAN!");
                            Console.ForegroundColor = ConsoleColor.White;
                        }

                        //TODO use another class to create new elements in the combobox
                    }
                }
                Console.WriteLine("====================================");
                Console.WriteLine("Elszámolt dolgozók száma: " + counter + " fő.");
                double goodPercent = (double)goodCounter / (double)counter;
                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine("Helyesen elszámolt dolgozók száma: " + goodCounter + " fő. " + goodPercent.ToString("#0.##%", CultureInfo.InvariantCulture));
                Console.ForegroundColor = ConsoleColor.White;
                double errorPercent = (double)errorCounter / (double)counter;
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Helytelenül elszámolt dolgozók száma: " + errorCounter + " fő. " + errorPercent.ToString("#0.##%", CultureInfo.InvariantCulture));
                Console.WriteLine("Helytelenül elszámolt dolgozók nevei:");
                foreach (var item in errorNames)
                {
                    Console.WriteLine(item);
                }
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("A munkaóra elszámolás kiértékelése véget ért.");
                //close file and dispose of com objects
                xlWB.Close(false, null, null);
                xlApp.Quit();
                Console.WriteLine("Az állomány terminálva.");
                Console.Beep();
                Marshal.ReleaseComObject(xlWS);
                Marshal.ReleaseComObject(xlWB);
                Marshal.ReleaseComObject(xlApp);

                Console.Read();
                //close file
                /*file.Close();
                Console.WriteLine("Specified file was closed successfully.");*/
            }
            catch (FileNotFoundException)
            {
                //write error to screen
                Console.WriteLine("Error. File was not found.");
            }
            
        }
    }
}
