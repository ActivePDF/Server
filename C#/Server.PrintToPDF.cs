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
            server.NewDocumentName = "Server.PrintToPDF.pdf";
            server.OutputDirectory = strPath;

            // Start the print job
            ServerDK.Results.ServerResult result = server.BeginPrintToPDF();
            if (result.ServerStatus == ServerDK.Results.ServerStatus.Success)
            {

                // Here is where you can print to ActivePDF Server to create a
                // PDF from any print job, set your application to print to a
                //static activePDF Server printer or call server.NewPrinterName
                // to dynamically create a new printer on the fly. This example
                // simply calls server.TestPrintToPDF for testing purposes
                result = server.TestPrintToPDF("Hello World!");
                if (result.ServerStatus ==
                    ServerDK.Results.ServerStatus.Success)
                {
                    // Wait(seconds) for job to complete
                    result = server.EndPrintToPDF(30);
                }
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
