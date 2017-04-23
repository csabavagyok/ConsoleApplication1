using System;
using System.Linq;
using System.IO;
using System.Collections;
using LinqToExcel;

namespace ConsoleApplication1
{
    /// <summary>
    /// Class to read XLSX contents and display them on the Console.
    /// </summary>
    class Program
    {
        public static void Main(string[] args)
        {
            TryOpenXlsx.TryOpenFile();
        }
    }
}