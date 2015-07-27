using System;

using System.Reflection;
using System.Windows.Forms;
using System.IO; //Path
using System.Diagnostics; //Process
using Excel = Microsoft.Office.Interop.Excel;

namespace GameNotificationTool
{
    class Program
    {
        static void Main(string[] args)
        {
            //get current directory of application
            string folder = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName);
            string file = folder + @"\Games Watch List.xlsx";
            
            Excel.Application excel = null;
            Excel.Workbook book = null;

            try
            {
                excel = new Excel.Application();
                book = excel.Workbooks.Open(file);
                
                Excel.Worksheet sheet = (Excel.Worksheet)book.Worksheets[1];

                Excel.Range range = null;

                if (sheet != null)
                {
                    range = sheet.UsedRange;
                }

                object[,] valueArray = (object[,])range.get_Value(
                    Excel.XlRangeValueDataType.xlRangeValueDefault);

                //title, release date, console, pre-ordered

                for (int row = 2; row <= sheet.UsedRange.Rows.Count; row++)
                {
                    try
                    {
                        //check release date
                        DateTime release = (DateTime)valueArray[row, 2];
                        DateTime current = DateTime.Now;

                        TimeSpan diff = release - current;
                        int datediff = (int)diff.TotalDays;

                        if ((datediff < 7) && !(datediff < 0))
                        {
                            string title = valueArray[row, 1].ToString();
                            string console = valueArray[row, 3].ToString();
                            string preorder = valueArray[row, 4].ToString();

                            string output;

                            if (datediff == 0)
                                output = title + " (" + console + ") is out now!";
                            else
                                output = datediff + " days until " + title + " (" + console + ") is released!";

                            if (preorder == "Y")
                                output = output + "\nYou have pre-ordered this game.";
                            else
                                output = output + "\nYou have not pre-ordered this game.";

                            MessageBox.Show(output,
                                "Game Notification Tool",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Asterisk);
                        }
                    }
                    catch (Exception)
                    {
                        //skip
                    }
                }
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("An error has occurred while running Game Notification Tool.\nError: " + e.Message);
            }

            book.Close();
        }
    }
}
