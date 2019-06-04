// Copyright (c) 2019 ActivePDF, Inc.
// ActivePDF Server

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
            server.NewDocumentName = "Server.WordPrinting.pdf";
            server.OutputDirectory = strPath;

            // Start the print job
            ServerDK.Results.ServerResult result = server.BeginPrintToPDF();
            if (result.ServerStatus == ServerDK.Results.ServerStatus.Success)
            {
                // Automate Word to print a document to Server
                // NOTE: You must add the 'Microsoft.Office.Interop.Word'
                // reference
                Microsoft.Office.Interop.Word._Application oWORD =
                    new Microsoft.Office.Interop.Word.Application();
                oWORD.ActivePrinter = server.NewPrinterName;
                oWORD.DisplayAlerts =
                    Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone;
                oWORD.Visible = false;
                Microsoft.Office.Interop.Word.Document oDOC =
                    oWORD.Documents.Open($"{strPath}Server.Word.Input.doc");
                oDOC.Activate();
                oWORD.PrintOut();
                oWORD.Documents.Close();
                oWORD.Quit();

                // Wait(seconds) for job to complete
                result = server.EndPrintToPDF(30);
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
