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

            // The below font options only work with conversions that go
            // through the printer or postscript file conversions. Set the
            // quality options for the created PDF For custom settings to take
            // effect set the configuration to custom
            server.PredefinedSetting =
                ADK.PostScript.PredefinedConfiguration.Custom;

            // Specifies if ASCII85 encoding should be applied to binary streams
            server.ASCIIEncode = true;

            // Automatically control the page orientation based on text flow
            server.AutoRotate = true;

            // Color Image Quality Settings
            server.ColorImageDownsampleThreshold = 1;
            server.ColorImageDownsampleType =
                ADK.PostScript.Images.DownsampleOption.None;
            server.ColorImageFilter =
                ADK.PostScript.Images.CompressionOption.FlateEncode;
            server.ColorImageResolution = 72;

            // Specifies if CMYK colors should be converted to RGB
            server.ConvertCMYKToRGB = true;

            // Gray Image Quality Settings
            server.GrayImageDownsampleThreshold = 1;
            server.GrayImageDownsampleType =
                ADK.PostScript.Images.DownsampleOption.None;
            server.GrayImageFilter =
                ADK.PostScript.Images.CompressionOption.FlateEncode;
            server.GrayImageResolution = 72;

            // Monochrome Image Quality Settings
            server.MonoImageDownsampleThreshold = 1;
            server.MonoImageDownsampleType =
                ADK.PostScript.Images.DownsampleOption.None;
            server.MonoImageFilter =
                ADK.PostScript.Images.MonochromeCompression.FlateEncode;
            server.MonoImageResolution = 72;

            // Set whether existing halftone settings should be preserved
            server.PreserveHalftone =
                ADK.PostScript.PreserveSettingOption.Preserve;

            // Set whether existing overprint settings should be preserved
            server.PreserveOverprint =
                ADK.PostScript.PreserveSettingOption.Preserve;

            // Set how transfer functions from the input file are handled
            server.PreserveTransferFunction =
                ADK.PostScript.PreserveTransferSettings.Preserve;

            // Set the DPI for the created PDF
            server.Resolution = 300.0f;

            // Set whether the UCRandBGInfo, from the input file, should be
            // preserved
            server.UCRandBGInfo =
                ADK.PostScript.PreserveSettingOption.Preserve;

            // Convert the PostScript file into PDF
            ServerDK.Results.ServerResult result =
                server.ConvertPSToPDF(
                    $"{strPath}Server.Input.ps",
                    $"{strPath}Server.OutputQuality.pdf");

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
