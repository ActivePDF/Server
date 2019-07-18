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

            // Add an email
            server.AddEMail();

            // Set server information
            // UPDATE THIS LINE WITH YOUR SERRVER INFORMATION
            server.SetSMTPInfo(server: "#.#.#.#", port: 0);
            server.SetSMTPCredentials(
                user: "john.doe",
                domain: "activePDF",
                password: "asdfasdf");

            // Set email addresses
            server.SetSenderInfo(
                friendlyName: "John Doe",
                address: "john.doe@asdidlwenra.com");
            server.SetReplyToInfo(
                friendlyName: "John Doe",
                address: "john.doe@asdidlwenra.com");
            server.SetRecipientInfo(
                friendlyName: "Jane Doe",
                address: "jane.doe@asdidlwenra.com");
            server.AddToCC(
                friendlyName: "Jim Doe",
                address: "jim.doe@asdidlwenra.com");
            server.AddToBcc(
                friendlyName: "Janice Doe",
                address: "janice.doe@asdidlwenra.com");

            // Subject and Body
            server.EMailSubject = "PDF Delivery from activePDF";
            server.SetEMailBody(bodyText: "<html><body style='background-color: #EEE; padding: 4px;'>Here is your PDF!</body></html>", isHtml: true);

            // Attachments - Binary attachments can be added with
            // AddEMailBinaryAttachment
            server.AddEMailAttachment($"{strPath}Server.Input.ps");

            // Other email options
            server.EMailReadReceipt = false;
            server.EMailAttachOutput = true;

            // Convert the PostScript file into PDF
            ServerDK.Results.ServerResult result =
                server.ConvertPSToPDF(
                    PSFile: $"{strPath}Server.Input.ps",
                    PDF: $"{strPath}Server.SendEmail.pdf");

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