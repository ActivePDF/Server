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
            server.NewDocumentName = "Server.PowerPointPrinting.pdf";
            server.OutputDirectory = strPath;

            // Start the print job
            ServerDK.Results.ServerResult result = server.BeginPrintToPDF();
            if (result.ServerStatus == ServerDK.Results.ServerStatus.Success)
            {
                // Automate PowerPoint to print a document to activePDF Server
                // NOTE: You must add the 'Microsoft PowerPoint
                // <<version number>> Object Library' COM object as a reference
                // to your .NET application to access the PowerPoint Object.
                Microsoft.Office.Interop.PowerPoint._Application oPPT =
                    new Microsoft.Office.Interop.PowerPoint.Application();
                Microsoft.Office.Interop.PowerPoint.Presentation oPRES =
                    oPPT.Presentations.Open(
                        $"{strPath}Server.PowerPoint.Input.pptx",
                        Microsoft.Office.Core.MsoTriState.msoTrue,
                        Microsoft.Office.Core.MsoTriState.msoFalse,
                        Microsoft.Office.Core.MsoTriState.msoFalse);
                Microsoft.Office.Interop.PowerPoint.PrintOptions objOptions =
                    oPRES.PrintOptions;
                objOptions.ActivePrinter = server.NewPrinterName;
                objOptions.PrintInBackground = 0;
                oPRES.PrintOut(1, 9999, "", 1, 0);
                oPRES.Saved = Microsoft.Office.Core.MsoTriState.msoTrue;
                oPRES.Close();
                oPPT.Quit();

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