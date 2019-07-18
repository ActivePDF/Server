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

            // Add PDF marks to be used in the converted PDF
            // PDF Mark Reference:
            // http://www.adobe.com/content/dam/Adobe/en/devnet/acrobat/pdfs/pdfmark_reference.pdf

            // Notes (Page 20 - PDF Marks Reference)
            server.AddPDFMark(PDFMark: "[ /SrcPg 1 /Rect [32 32 216 144] /Open false /Title (ActivePDF Comment) /Contents (Note Type Comment Example.) /Color [1 0 0] /Subtype /Text /ANN pdfmark");

            // Free Text (Page 17 - PDF Marks Reference)
            // Used for the text in the link created below
            server.AddPDFMark(PDFMark: "[ /SrcPg 1 /Rect [262 26 350 46] /Contents (ActivePDF.com) /DA ([0 0 1] rg /Helv 12 Tf) /BS << /W 0 >> /Q 1 /Subtype /FreeText /ANN pdfmark");

            // Links (Page  - PDF Marks Reference)
            // Add a link around the activePDF.com text created above
            server.AddPDFMark(PDFMark: "[ /SrcPg 1 /Rect [262 26 350 46] /Contents (ActivePDF.com) /BS << /W 0 >> /Action << /Subtype /URI /URI (http://ActivePDF.com) >> /Subtype /Link /ANN pdfmark");

            // Bookmarks (Page 26 - PDF Marks Reference)
            server.AddPDFMark(PDFMark: "[ /Count -5 /Title (ActivePDF Server - AddPDFMark) /Page 1 /F 2 /OUT pdfmark");
            server.AddPDFMark(PDFMark: "[ /Page 1 /View [/Fit] /Title (AddPDFMark - Page 1) /C [0 0 0] /F 2 /OUT pdfmark");
            server.AddPDFMark(PDFMark: "[ /Page 2 /View [/Fit] /Title (AddPDFMark - Page 2) /C [0 0 0] /F 2 /OUT pdfmark");

            // Document Information (Page 28 - PDF Marks Reference)
            server.AddPDFMark(PDFMark: "[ /Title (PDF Marks) /Author (ActivePDF Server) /Subject (PDF Marks) /Keywords (pdfmark, server, example) /DOCINFO pdfmark");

            // Document View Options (Page 29 - PDF Marks Reference)
            server.AddPDFMark(PDFMark: "[ /PageMode /UseOutlines /Page 1 /View [/Fit] /DOCVIEW pdfmark");

            // Document Open Options
            server.AddPDFMark(PDFMark: "[ {Catalog} << /ViewerPreferences << /HideToolbar true  /HideMenubar true >> >> /PUT pdfmark");

            // Convert the PostScript file into PDF
            ServerDK.Results.ServerResult result =
                server.ConvertPSToPDF(
                    PSFile: $"{strPath}Server.Input.ps",
                    PDF: $"{strPath}Server.AddPDFMark.pdf");

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