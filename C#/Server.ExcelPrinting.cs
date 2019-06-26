// Copyright (c) 2019 ActivePDF, Inc.
// ActivePDF WebGrabber

using System;

// Make sure to add the ActivePDF product .NET DLL(s) to your application.
// .NET DLL(s) are typically found in the products 'bin' folder.

namespace ServerExamples
{
    class Program
    {
        static void Main(string[] args)
        {
            string strPath =
                System.AppDomain.CurrentDomain.BaseDirectory.Replace("\\", "/");

            // Instantiate Object
            APServer.Server server = new APServer.Server();

            // Path and filename of output
            server.NewDocumentName = "Server.ExcelPrinting.pdf";
            server.OutputDirectory = strPath;

            // Start the print job
            ServerDK.Results.ServerResult result = server.BeginPrintToPDF();
            if (result.ServerStatus == ServerDK.Results.ServerStatus.Success)
            {
                // Automate Excel to print a document to activePDF Server
                // NOTE: You must add a reference to the
                // Microsoft.Office.Interop.Excel library found in the
                // reference manager under Assemblies -> Extensions
                Microsoft.Office.Interop.Excel._Application oXLS =
                    new Microsoft.Office.Interop.Excel.Application();
                oXLS.DisplayAlerts = false;
                oXLS.Visible = false;
                object m = System.Type.Missing;
                Microsoft.Office.Interop.Excel._Workbook oWB =
                    oXLS.Workbooks.Open(
                        $"{strPath}Server.Excel.Input.xlsx", m, true, m, m, m,
                        true, m, m, false, false, m, false);
                oWB.Activate();
                oWB.PrintOut(1, 999, 1, false, server.NewPrinterName, false, false);
                oWB.Close(0);
                oXLS.Quit();

                // Wait(seconds) for job to complete
                result = server.EndPrintToPDF(waitTime: 30);
            }

            // Output result
            WriteResult(result);

            // Process Complete
            Console.WriteLine("Done!");
            Console.WriteLine("Press any key to exit.");
            Console.ReadKey();
        }

        public static void WriteResult(ServerDK.Results.ServerResult Result)
        {
            Console.WriteLine($"Server Status: {Result.ServerStatus}");
            if (Result.ServerStatus != ServerDK.Results.ServerStatus.Success)
            {
                Console.WriteLine($"Result Origin: {Result.Origin.Class}.{Result.Origin.Function}");
                if (!String.IsNullOrEmpty(Result.Details))
                {
                    Console.WriteLine($"Result Details: {Result.Details}");
                }
                if (Result.ResultException != null)
                {
                    Console.WriteLine("Exception caught during conversion.");
                    Console.WriteLine($"Excpetion Details: {Result.ResultException.Message}");
                    Console.WriteLine($"Exception Stack Trace: {Result.ResultException.StackTrace}");
                }
            }
        }
    }
}
