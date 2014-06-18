using System;
using System.Globalization;
using A2A = ReportLib;

namespace ErrorReportCmd
{
    class Program
    {
        private static void Main()
        {
            string dateTime = DateTime.Now.ToString("yyyy-MM-dd-HH-mm", CultureInfo.InvariantCulture);

            //const string folder = @"\\WIN-E5U3RQA9QQH\MigrationTransfer\A2A_Out";
            const string folder = @"\\hyunor01srv\share\SDSeeding\A2AMigrationOutput";

            //const string folder = @"C:\Outs";

            string outFile = @"\\WIN-E5U3RQA9QQH\MigrationTransfer\A2A_Reports\ErrorReport_" + dateTime + ".xlsx";
            //string outFile = @"C:\Outs\ErrorReport_" + dateTime + ".xlsx";

            A2A.Report.GenerateReport(folder, outFile, dateTime);

            Console.WriteLine("Done! Press any key to open the report location and exit.");
            Console.ReadLine();

            string argument = @"/select, " + outFile;
            System.Diagnostics.Process.Start("explorer.exe", argument);
        }
    }
}
