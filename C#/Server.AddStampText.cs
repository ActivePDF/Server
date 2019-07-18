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

            // Stamp Images and Text onto the output PDF
            server.AddStampCollection("TXTinternal");
            server.StampFont = "Helvetica";
            server.StampFontSize = 108;
            server.StampFontTransparency = 0.3f;
            server.StampRotation = 45.0f;

            server.StampFillMode = ADK.PDF.FontFillMode.FillThenStroke;
            server.StampColorNET =
                new ADK.PDF.Color()
                {
                    Red = 255,
                    Green = 0,
                    Blue = 0,
                    Gray = 0
                };
            server.StampStrokeColorNET =
                new ADK.PDF.Color()
                {
                    Red = 100,
                    Green = 0,
                    Blue = 0,
                    Gray = 0
                };

            server.AddStampText(
                x: 116.0f,
                y: 156.0f,
                stampText: "Internal Only");

            // Set whether the stamp collection(s) appears in the background or
            // foreground
            server.StampBackground = 0;

            // Convert the PostScript file into PDF
            ServerDK.Results.ServerResult result =
                server.ConvertPSToPDF(
                    PSFile: $"{strPath}Server.Input.ps",
                    PDF: $"{strPath}Server.AddStampText.pdf");

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