using System;
using System.IO;
using System.Threading;
using System.Linq;

namespace BiosDownloader
{
    class Program
    {
        static void Main(string[] args)
        {
            string biosDownloadLink = Methods.GetDownloadLink(); //Get BIOS download link for current machine            
            string fileName = biosDownloadLink.Split('/').Last(); //Get filename of BIOS to name downloaded file
            string downloadDir = Path.Combine(Path.GetTempPath() + @"BIOS\");
            string theLatestBios = fileName.Split('_').Last().Substring(0, fileName.Split('_').Last().Length - 4); //Get the BIOS verion number from Dell's website
            string currentBIOS = Methods.GetCurrentBIOS(); //Get current machine's BIOS version
            Methods.LogActionToFile("Retrieved machine specific BIOS link");

            Methods.LogActionToFile($"The latest BIOS is {theLatestBios}, the current version is {currentBIOS}");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"The latest BIOS is {theLatestBios}, the current version is {currentBIOS}");

            if (currentBIOS == theLatestBios) //If the current version is the latest then there is no need to update
            {
                Console.ForegroundColor = ConsoleColor.Magenta;
                Methods.LogActionToFile("BIOS Version is already up to date.");
                Console.WriteLine("BIOS Version is already up to date. Closing.");
                Thread.Sleep(2000);
                Environment.Exit(0);
            }
            else //Current version is not latest version, update needed
            {
                Methods.DownloadFile(biosDownloadLink, fileName); //Download the file and name it accordingly
                Methods.StartUpgrade(downloadDir, fileName); //Start process executable
            }
        }     
    }
}