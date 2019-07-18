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

            // Add bookmarks to pages in the PDF
            server.AddPageBookmark(
                Title: "Parent",
                subCount: 1,
                PageNbr: 1,
                View: "Fit");
            server.AddPageBookmark(
                Title: "Child 1",
                subCount: 0,
                PageNbr: 2,
                View: "Fit");

            // Add bookmarks to URLs
            server.AddURLBookmark(
                Title: "Parent URL Bookmark",
                subCount: 1,
                URL: "http://www.activepdf.com");
            server.AddURLBookmark(
                Title: "Child URL Bookmark",
                subCount: 0,
                URL: "https://www.activepdf.com/products/server");

            // Add bookmarks pointing to pages in external PDF
            // Both Local and UNC file paths are accepted
            server.AddLinkedPDFBookmark(
                Title: "Parent PDF Bookmark",
                subCount: 1,
                PDFFilename: $"{strPath}Server.Sample.pdf",
                PageNbr: 1,
                View: "Fit");
            server.AddLinkedPDFBookmark(
                Title: "Child PDF Bookmark",
                subCount: 0,
                PDFFilename: $"{strPath}Server.Sample.pdf",
                PageNbr: 2,
                View: "Fit");

            // Add bookmarks pointing to any external file
            // Both Local and UNC file paths are accepted
            server.AddFileBookmark(
                Title: "Parent File Bookmark",
                subCount: 1,
                Filename: $"{strPath}Server.PowerPoint.Input.pptx");
            server.AddFileBookmark(
                Title: "Child  File Bookmark",
                subCount: 0,
                Filename: $"{strPath}Server.Word.Input.doc");

            // Convert the PostScript file into PDF
            ServerDK.Results.ServerResult result =
                server.ConvertPSToPDF(
                    PSFile: $"{strPath}Server.Input.ps",
                    PDF: $"{strPath}Server.AddBookmarks.pdf");

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