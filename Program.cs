using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;

namespace excel1
{
    class Program
    {
        static void Main(string[] args)
        {
            string dir = @"C:\FilePath\Target\";
            char delim = ',';

            foreach (string f in GetFilesInDirectory(dir))
            {
                string fileOut = dir + Path.GetFileNameWithoutExtension(f) +  ".xlsx";
                WriteToExcel(fileOut, ReadCSVFIle(",", f), delim);
            }
            
        }

        public static string[] GetFilesInDirectory(string dir)
        {
            string[] files = Directory.GetFiles(dir);
            return files;
        }

        public static string[] ReadCSVFIle(string delim, string fileName)
        {
            string[] allLines = File.ReadAllLines(fileName);

            return allLines;
        }

        public static void WriteToExcel(string filePath, string[] fileLines, char delim)
        {
            var app = new Excel.Application();
            app.Visible = false;
            try
            {
                var book = app.Workbooks.Add();
                try
                {
                    var sheet = book.Sheets[1];
                    

                    try
                    {
                        var cell = sheet.Range["A1"] as Excel.Range;
                        try
                        {
                            int counter = 1;
                            foreach(string s in fileLines)
                            {

                                string[] wrds = s.Split(delim);
                                for (int i = 0; i < wrds.Length; i++)
                                {
                                    var cell2 = sheet.Cells(counter, i+1);
                                    cell2.Value = "'" + wrds[i];
                                }

                                counter++;
                            }                          
                        }
                        finally
                        {
                            while (0 < Marshal.ReleaseComObject(cell)) { }
                        }
                    }
                    finally
                    {
                        while (0 < Marshal.ReleaseComObject(sheet)) { }
                    }
                }
                finally
                {
                    book.SaveAs(filePath);
                    while (0 < Marshal.ReleaseComObject(book)) { }
                }
            }
            finally
            {                
                app.Visible = false;
                app.Quit();
            }
        }
    }
}
