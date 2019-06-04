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
            server.OutputDirectory = strPath;

            // Add bookmarks to pages in the PDF
            server.AddPageBookmark("Page 1", 0, 1, "Fit");
            server.AddPageBookmark("Page 2", 0, 2, "Fit");

            // Convert the PostScript file into PDF
            ServerDK.Results.ServerResult result =
                server.ConvertPSToPDF(
                    $"{strPath}Server.Input.ps",
                    $"{strPath}Server.AddPageBookmark.pdf");

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
