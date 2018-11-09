using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Diagnostics;
using System.IO;
using System.Management;
using System.Net.NetworkInformation;
using System.Threading;
using System.Linq;

namespace BiosDownloader
{
    class Program
    {
        static void Main(string[] args)
        {
            string biosDownloadLink = GetDownloadLink(); //Get BIOS download link for current machine            
            string fileName = biosDownloadLink.Split('/').Last(); //Get filename of BIOS to name downloaded file
            string downloadDir = Path.Combine(Path.GetTempPath() + @"BIOS\");
            string theLatestBios = fileName.Split('_').Last().Substring(0, fileName.Split('_').Last().Length - 4); //Get the BIOS verion number from Dell's website
            string currentBIOS = GetCurrentBIOS(); //Get current machine's BIOS version
            LogActionToFile("Retrieved machine specific BIOS link");

            LogActionToFile($"The latest BIOS is {theLatestBios}, the current version is {currentBIOS}");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"The latest BIOS is {theLatestBios}, the current version is {currentBIOS}");

            if (currentBIOS == theLatestBios) //If the current version is the latest then there is no need to update
            {
                Console.ForegroundColor = ConsoleColor.Magenta;
                LogActionToFile("BIOS Version is already up to date.");
                Console.WriteLine("BIOS Version is already up to date. Closing.");
                Thread.Sleep(2000);
                Environment.Exit(0);
            }
            else //Current version is not latest version, update needed
            {
                DownloadFile(biosDownloadLink, fileName); //Download the file and name it accordingly
                StartUpgrade(downloadDir, fileName); //Start process executable
            }
        }

        public static bool CheckForInternetConnection()
        {
            Ping myPing = new Ping();
            String host = "downloads.dell.com";
            byte[] buffer = new byte[32];
            int timeout = 1000;
            PingOptions pingOptions = new PingOptions();
            try
            {
                PingReply reply = myPing.Send(host, timeout, buffer, pingOptions);
                if (reply.Status == IPStatus.Success)
                {
                    LogActionToFile("Internet Check passed");
                    return true;
                }
            }
            catch (Exception)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                LogActionToFile("Internet Connection check failed");
                Console.WriteLine("Unable to access the Dell support web page.");
                Console.WriteLine();
                Console.WriteLine("Press any key to close the program");
                Console.ReadKey();
                Environment.Exit(0);
            }
            return false;
        }
        public static string GetServiceTag()
        {
            string serviceTag = null;
            try //Use WMI for finding BIOS information
            {
                ManagementObjectSearcher searcher =
                    new ManagementObjectSearcher("root\\CIMV2",
                    "SELECT * FROM Win32_BIOS");
                foreach (ManagementObject queryObj in searcher.Get())
                {
                    serviceTag = queryObj["SerialNumber"].ToString();
                }
            }
            catch (ManagementException e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                LogActionToFile($"An error occurred while querying for WMI data: { e.Message}");         
                Console.WriteLine($"An error occurred while querying for WMI data: {e.Message}");
                Console.ReadKey();
                Environment.Exit(0);
            }
            return serviceTag;
        }
        public static string GetCurrentBIOS()
        {
            string currentBIOSVersion = null;
            try //Use WMI for finding BIOS information
            {
                ManagementObjectSearcher searcher =
                    new ManagementObjectSearcher("root\\CIMV2",
                    "SELECT * FROM Win32_BIOS");
                foreach (ManagementObject queryObj in searcher.Get())
                {
                    currentBIOSVersion = queryObj["SMBIOSBIOSVersion"].ToString();
                }
            }
            catch (ManagementException e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                LogActionToFile($"An error occurred while querying for WMI data: { e.Message}");
                Console.WriteLine($"An error occurred while querying for WMI data: {e.Message}");
                Console.WriteLine("Press any key to close the program");
                Console.ReadKey();
                Environment.Exit(0);
            }
            return currentBIOSVersion;
        }
        public static string GetDownloadLink()
        {
            int retryAttempts = 1;
            CheckForInternetConnection(); //Try to ping download.dell.com. If it does not reply, close program
            //string testSerialNumber = "4CHV0Q2";
            //string brandonSerialNumber = "7GH1MR2";
            //string e5550SerialNumber = "7RVSVL1";
            ChromeOptions options = new ChromeOptions();
            options.AddArguments("--window-size=1280,720");
            options.AddArguments("--disable-gpu");
            options.AddArguments("--disable-extensions");
            options.AddArguments("--start-maximized");
            options.AddArgument("--log-level=3");
            options.AddArguments("--headless");
            options.AddArgument("--silent");
            using (ChromeDriver driver = new ChromeDriver(options))
            {
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
                try
                {
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("Loading Support Page. This will take a few moments."); //Load machine specific download page
                    driver.Navigate().GoToUrl("https://www.dell.com/support/home/us/en/04/product-support/servicetag/" + GetServiceTag());
                    Console.WriteLine();
                    if (driver.FindElement(By.CssSelector(".alert.alert-warning.alert-dismissable.ng-scope")).Displayed)
                    //Dell's page for missing Service Tags
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        LogActionToFile($"Service Tag {GetServiceTag()} does not exist on Dell's page");
                        Console.WriteLine();
                        Console.WriteLine();
                        Console.WriteLine("This machine does not have a working Service Tag page at Dell.com");
                        Console.WriteLine("Due to this page being missing, it is not possible to download the BIOS.");
                        Console.WriteLine("Press any key to close");
                        Console.ReadKey();
                        Console.WriteLine();
                        Console.ForegroundColor = ConsoleColor.DarkCyan;
                        Console.WriteLine("Exiting program....");
                        Thread.Sleep(1000);
                        driver.Dispose();
                        Environment.Exit(0);
                    }
                }
                catch (Exception e)
                {
                    if (!e.Message.Contains(".alert.alert-warning.alert-dismissable.ng-scope"))//Error is already caught
                    {
                        //the error is not expected, print error message
                        if (retryAttempts <= 3)
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine($"Error: {e.Message}");
                            LogActionToFile($"Error: {e.Message}");
                            Console.WriteLine($"Attempting to reload support page. Attempt #{retryAttempts}.");
                            retryAttempts++;
                            GetDownloadLink();
                        }
                        else
                        {
                            LogActionToFile($"Repeated issues loading page elements, closing. Attempt #{retryAttempts}.");
                            Console.WriteLine($"Repeated issues loading page elements, closing. Attempt #{retryAttempts}.");
                            Thread.Sleep(1000);
                            Environment.Exit(0);
                        }
                    }
                }

                Console.WriteLine("Loading Drivers Tab");
                try //Click on driver's tab
                {
                    driver.FindElement(By.Id("tab-drivers")).Click();
                }
                catch (Exception e)
                {                   
                    if (retryAttempts <= 3)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("Error: Failed to open Drivers tab");
                        LogActionToFile($"Error: {e.Message}");
                        Console.WriteLine($"Attempting to reload support page. Attempt #{retryAttempts}.");
                        retryAttempts++;
                        GetDownloadLink();
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        LogActionToFile($"Repeated issues loading page elements, closing. Attempts #{retryAttempts}.");
                        Console.WriteLine($"Repeated issues loading page elements, closing. Attempt #{retryAttempts}.");
                        Thread.Sleep(1000);
                        Environment.Exit(0);
                    }
                }

                LogActionToFile("Successfully loaded Drivers tab");
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Loading BIOS Downloads");

                try //Enter "BIOS" to filter downloads
                {
                    driver.FindElement(By.XPath("//*[@id=\"ddlcategoryFilter\"]")).SendKeys("BIOS");
                }
                catch (Exception e)
                {
                    if (retryAttempts <= 3)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("Error: Failed to filter to BIOS");
                        LogActionToFile($"Error: {e.Message}");
                        Console.WriteLine($"Attempting to reload support page. Attempt #{retryAttempts}.");
                        retryAttempts++;
                        GetDownloadLink();
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        LogActionToFile($"Repeated issues loading page elements, closing. Attempts: {retryAttempts}");
                        Console.WriteLine($"Repeated issues loading page elements, closing. Attempts: {retryAttempts}");
                        Thread.Sleep(1000);
                        Environment.Exit(0);
                    }
                }
                LogActionToFile("Successfully loaded BIOS section");
                Console.WriteLine("Retrieving BIOS download link");
                string biosDownloadLink = null;

                try //Retrieve BIOS download link from first item in list after filtering to BIOS
                {
                    biosDownloadLink = driver.FindElement(By.CssSelector(".pointerCursor.text-blue.dellmetrics-driverdownloads.dld0")).GetAttribute("href");
                }
                catch (Exception e)
                {
                    if (retryAttempts <= 3)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"Error: {e.Message}");
                        LogActionToFile($"Error: {e.Message}");
                        Console.WriteLine("Reloading");
                        retryAttempts++;
                        GetDownloadLink();
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        LogActionToFile($"Repeated issues loading page elements, closing. Attempts #{retryAttempts}.");
                        Console.WriteLine($"Repeated issues loading page elements, closing. Attempts #{retryAttempts}.");
                        Thread.Sleep(1000);
                        Environment.Exit(0);
                    }
                }
                driver.Dispose(); //Dispose of webdriver as BIOS download link has been retrieved
                return biosDownloadLink;
            }
        }
        public static void DownloadFile(string link, string name)
        {
            using (var client = new System.Net.WebClient())
            {
                Console.ForegroundColor = ConsoleColor.Green;
                string downloadDir = Path.Combine(Path.GetTempPath() + @"BIOS\");
                DirectoryInfo di = new DirectoryInfo(downloadDir);
                foreach (FileInfo file in di.GetFiles()) /*If the directory exists and contains an .exe, delete it */
                {
                    if (file.Extension == ".exe")
                    {
                        LogActionToFile("Potentially outdated file exists in BIOS directory, deleting.");
                        file.Delete();
                    }
                }

                Console.WriteLine("Downloading BIOS...");
                client.DownloadFile(link, downloadDir + name); //Download latest version
                LogActionToFile("Successfully downloaded BIOS update");
                Console.WriteLine("File Downloaded");
            }
        }
        public static void StartUpgrade(string path, string name)
        {
            Console.WriteLine("Starting BIOS Upgrade...");
            try
            {                
                ProcessStartInfo pi = new ProcessStartInfo
                {
                    Verb = "runas",
                    FileName = path + name
                };
                Process.Start(pi); //Start BIOS upgrade prompting for user to run as
                LogActionToFile("Correctly opened BIOS upgrade executable");
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                LogActionToFile($"Executable not opened correctly: {e.Message}");
                Console.WriteLine($"Upgrade failed! Error Code: {e.Message}");
                Thread.Sleep(5000);
                Environment.Exit(0);
            }
        }
        public static void LogActionToFile(string msg)
        {
            Directory.CreateDirectory(Path.GetTempPath() + "BIOS\\");
            StreamWriter streamWriter = File.AppendText(Path.Combine(Path.GetTempPath() + "BIOS\\") + "log.txt");
            try //Format text entered to prepend date and time
            {
                string message = string.Format("{0:G}: {1}.", DateTime.Now, msg);
                streamWriter.WriteLine(message);
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Error writing log to file. Error code: {e.Message}");
            }
            finally
            {
                streamWriter.Close();
            }
        }
    }
}