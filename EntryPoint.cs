using System;
using System.Collections.Generic;
using System.Text;
using System.Management;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.Win32;

using MsiRcw = MyCompany.Ops.Utils.Products.MsiRcw;
using System.Diagnostics;

namespace Products
{
    /// <summary>
    /// Summary description for Installer.
    /// </summary>
    /// <remarks>see readme1.htm/readme2.htm for the details on why this is needed</remarks>
    [ComImport, Guid("000C1090-0000-0000-C000-000000000046")]
    public class Installer { }

    /// <summary>
    /// Summary description for EntryPoint.
    /// </summary>
    public class EntryPoint
    {
        private const string Status = "in progress";
        private const string Version = "02jul12";

        private const string RegExKbPattern = "(KB\\d{7})|(KB\\d{6})|(KB\\d{5})";
        private static Regex RegExKb = new Regex(RegExKbPattern, RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private const string DateTimeFormatString = "yyyyMMdd h:MM:ss";
        
        [Flags]
        public enum ArgsSupported
        {
            None = 0x1,
            ShowUsage = 0x2,
            ProductType = 0x4,
            ListProducts = 0x8 ,
            SearchProducts = 0x10,
            SearchProductsByVersion = 0x20,
            SearchProductsBySource = 0x40,
            RemoveProducts = 0x80,
            Verbose = 0x100,
            SingleLine = 0x200,
            VerboseSingleLine = Verbose | SingleLine
        }

        public enum ProductTypes
        {
            None,
            All,
            Msi,
            NonMsi,
            Mu,
            Wu
        }

        public enum SearchType
        {
            None,
            Name,
            Source,
            Version
        }

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        private static void Main(string[] args)
        {
            //String[] argsTest = Environment.GetCommandLineArgs();

            ArgsSupported argsFound = ArgsSupported.None;
            ProductTypes productType = ProductTypes.Msi;  // default setting | None | All | NonMsi
            string productName = String.Empty;
            string productSource = String.Empty;
            string productVersion = String.Empty;

            //foreach (string arg in args)
            for (int i = 0; (argsFound & ArgsSupported.ShowUsage) != ArgsSupported.ShowUsage && i < args.Length; i++)
            {
                //switch (args.GetValue(i).ToString().ToLower()) 
                switch (args[i].ToLower())
                {
                    case "/l":
                        argsFound |= ArgsSupported.ListProducts;
                        break;

                    case "/s":
                        if (args.Length > ++i)
                        {
                            productName = args.GetValue(i).ToString();
                            argsFound |= ArgsSupported.SearchProducts;
                        }
                        else  // switch value not provided
                        {
                            argsFound |= ArgsSupported.ShowUsage;
                        }
                        break;

                    case "/ss":
                        if (args.Length > ++i)
                        {
                            productSource = args.GetValue(i).ToString();
                            argsFound |= ArgsSupported.SearchProductsBySource;
                        }
                        else  // switch value not provided
                        {
                            argsFound |= ArgsSupported.ShowUsage;
                        }
                        break;

                    case "/sv":
                        if (args.Length > ++i)
                        {
                            productVersion = args.GetValue(i).ToString();
                            argsFound |= ArgsSupported.SearchProductsByVersion;
                        }
                        else  // switch value not provided
                        {
                            argsFound |= ArgsSupported.ShowUsage;
                        }
                        break;

                    case "/r":
                        if (args.Length > ++i)
                        {
                            productName = args.GetValue(i).ToString();
                            argsFound |= ArgsSupported.SearchProducts;
                            argsFound |= ArgsSupported.RemoveProducts;
                        }
                        else  // switch value not provided
                        {
                            argsFound |= ArgsSupported.ShowUsage;
                        }
                        break;

                    case "/rs":
                        if (args.Length > ++i)
                        {
                            productSource = args.GetValue(i).ToString();
                            argsFound |= ArgsSupported.SearchProductsBySource;
                            argsFound |= ArgsSupported.RemoveProducts;
                        }
                        else  // switch value not provided
                        {
                            argsFound |= ArgsSupported.ShowUsage;
                        }
                        break;

                    case "/rv":
                        if (args.Length > ++i)
                        {
                            productVersion = args.GetValue(i).ToString();
                            argsFound |= ArgsSupported.SearchProductsByVersion;
                            argsFound |= ArgsSupported.RemoveProducts;
                        }
                        else  // switch value not provided
                        {
                            argsFound |= ArgsSupported.ShowUsage;
                        }
                        break;

                    case "/t":
                        if (args.Length > ++i)
                        {
                            string productTypeValue = args.GetValue(i).ToString();
                            if (Regex.IsMatch(productTypeValue, @"(all)|(msi)|(nonmsi)|(mu)|(wu)") == true)
                            {
                                switch (productTypeValue)
                                {
                                    case "all":
                                        productType = ProductTypes.All;
                                        break;
                                    case "msi":
                                        productType = ProductTypes.Msi;
                                        break;
                                    case "nonmsi":
                                        productType = ProductTypes.NonMsi;
                                        break;
                                    case "mu":
                                        productType = ProductTypes.Mu;
                                        break;
                                    case "wu":
                                        productType = ProductTypes.Wu;
                                        break;
                                    default:
                                        productType = ProductTypes.None;
                                        break;
                                }
                                argsFound |= ArgsSupported.ProductType;
                            }
                            else  // unsupported switchParam provided
                            {
                                argsFound |= ArgsSupported.ShowUsage;
                            }
                        }
                        else  // switch value not provided
                        {
                            argsFound |= ArgsSupported.ShowUsage;
                        }
                        break;

                    case "/v":
                        //if (args.Length > ++i)
                        //{
                        //  productName = args.GetValue(i).ToString();
                        //  argsFound |= ArgsSupported.GetProductVersion;
                            argsFound |= ArgsSupported.Verbose;
                        //}
                        //else  // switch value not provided
                        //{
                        //  argsFound |= ArgsSupported.ShowUsage;
                        //}
                        break;

                    case "/vsl":
                        //if (args.Length > ++i)
                        //{
                        //  productName = args.GetValue(i).ToString();
                        //  argsFound |= ArgsSupported.GetProductVersion;
                            argsFound |= ArgsSupported.VerboseSingleLine;
                        //}
                        //else  // switch value not provided
                        //{
                        //  argsFound |= ArgsSupported.ShowUsage;
                        //}
                        break;

                    default:  // /? | /h | <unsupportedArg>
                        argsFound |= ArgsSupported.ShowUsage;
                        break;
                }
            }

            if (argsFound == ArgsSupported.None || (argsFound & ArgsSupported.ShowUsage) == ArgsSupported.ShowUsage)
            {
                ShowUsage();
            }
            else
            {
                if ((argsFound & ArgsSupported.ListProducts) == ArgsSupported.ListProducts)
                {
                    ListProducts(productType, argsFound);
                }
                else if ((argsFound & ArgsSupported.SearchProducts) == ArgsSupported.SearchProducts)
                {
                    SearchProducts(productType, productName, SearchType.Name, argsFound);
                }
                else if ((argsFound & ArgsSupported.SearchProductsBySource) == ArgsSupported.SearchProductsBySource)
                {
                    SearchProducts(productType, productSource, SearchType.Source, argsFound);
                }
                else if ((argsFound & ArgsSupported.SearchProductsByVersion) == ArgsSupported.SearchProductsByVersion)
                {
                    SearchProducts(productType, productVersion, SearchType.Version, argsFound);
                }
            }
        }

        private static void ShowUsage()
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string asmVersion = asm.GetName().Version.ToString();
            string asmExe = Regex.Match(Environment.CommandLine, @"[^\\]+\.exe", RegexOptions.None).Value;

            Console.WriteLine("\nstatus = " + Status + ", version = " + Version + "\n");  
            // or Console.WriteLine("\nversion = " + Version + "\n");
            Console.WriteLine("description");
            Console.WriteLine("  list and search for installed products from a command line interface\n");
            Console.WriteLine("usage");
            Console.WriteLine("  " + asmExe + " [/l | /s <searchText>] [/t <all | msi | nonmsi | mu | wu>] [/v]\n");
            //Console.WriteLine("  " + asmExe + " [/l | /s|r[sv] <searchText>] [/t <all | msi | nonmsi | mu | wu>] [/v]\n");
            Console.WriteLine("where");
            Console.WriteLine("  /l = list products");
            //Console.WriteLine("  /c = check if product is installed");
            Console.WriteLine("  /s = search product names containing search text");
            Console.WriteLine("  /ss = search product install source containing search text");
            Console.WriteLine("  /sv = search product versions containing search text");
            //Console.WriteLine("  /r = search product names containing search text and remove");
            //Console.WriteLine("  /rs = search product install source containing search text and remove");
            //Console.WriteLine("  /rv = search product versions containing search text and remove");
            Console.WriteLine("  /t = product type options are \"all\", \"msi\" (default), \"nonmsi\", \"mu\" or \"wu\"");
            Console.WriteLine("  /v = verbose product info output");
            Console.WriteLine("  /vsl = verbose product info output on single line for sort and diff purposes");
            Console.WriteLine("  /? = or /h or no arguments shows usage info\n");
            Console.WriteLine("examples");
            Console.WriteLine("  " + asmExe + " /l /t all /v");
            Console.WriteLine("  " + asmExe + " /s office /t msi");
            //Console.WriteLine("  " + asmExe + " /s \"sql server\" /t nonmsi");
            Console.WriteLine("  " + asmExe + " /s \"silverlight \\d \\w* ?sdk\" /vsl");
            Console.WriteLine("  " + asmExe + " /ss \"src\\\\vs12\\\\x86\" /vsl");
            Console.WriteLine("  " + asmExe + " /ss \"v11.0.5\\d{4}\\packages\" /vsl");
            Console.WriteLine("  " + asmExe + " /sv \"11.0.5\\d{4}\" /vsl");
            //Console.WriteLine("  " + asmExe + " /s \"Microsoft .NET Framework (?:4.5|4.5 Developer Preview) Multi-Targeting Pack\"");
            //Console.WriteLine("  " + asmExe + " /r \"Blend 5\" /vsl");
            Console.WriteLine("  " + asmExe + " /s kb2345678 /t mu");
            Console.WriteLine("");
        }

        private static void ListProducts(ProductTypes productType, ArgsSupported argsFound)
        {
            bool verbose = (argsFound & ArgsSupported.Verbose) == ArgsSupported.Verbose;
            bool verboseSingleLine = (argsFound & ArgsSupported.VerboseSingleLine) == ArgsSupported.VerboseSingleLine;

            if (productType == ProductTypes.All || productType == ProductTypes.Msi)
            {
                Console.WriteLine("------------ begin ListProducts Msi -------------");

#if DEBUG
                //Type typeFromProgID = Type.GetTypeFromProgID("WindowsInstaller.Installer");
                //Type typeFromInterop = typeof(MsiRcw.Installer);
#endif
                // initially the wmi Windows Installer provider was available by default on x86 os installs and as an optional component on x64 os installs.
                // it now appears that it is there by default in both cases and so we could optional switch back to using it and get away from our dependency 
                // on the WindowsInstaller.Installer COM runtime callable wrapper

                //ManagementClass win32Products = new ManagementClass("Win32_Product");
                //foreach (ManagementObject win32Product in win32Products.GetInstances())
                //{
                //    Console.WriteLine("Id = {0} " + "Name = {1} Version = {2}", win32Product["IdentifyingNumber"], win32Product["Name"], win32Product["Version"]);
                //}
                
                // or

                //SelectQuery selectQuery = new SelectQuery("Win32_Product");
                //ManagementObjectSearcher searcher = new ManagementObjectSearcher(selectQuery);
                //foreach (ManagementObject win32Product in searcher.Get())
                //{
                //    Console.WriteLine("Id = {0} " + "Name = {1} Version = {2}",
                //        win32Product["IdentifyingNumber"], win32Product["Name"], win32Product["Version"]);
                //}

                // or

                MsiRcw.Installer installer = (MsiRcw.Installer)new Installer();
                //foreach (string product in installer.Products)
                for (int i = 0; i < installer.Products.Count; i++)
                {
                    //string currentItemId = product;
                    string currentItemId = installer.Products[i];
                    string currentItemName = string.Empty;
                    string currentItemInstallSource = string.Empty;
                    string currentItemVersion = string.Empty;
                    try
                    {
                        currentItemName = installer.get_ProductInfo(currentItemId, "InstalledProductName");
                        currentItemInstallSource = installer.get_ProductInfo(currentItemId, "InstallSource");
                        if (currentItemInstallSource.Length == 0) currentItemInstallSource = " ";  // "<value not found>";
                        currentItemVersion = installer.get_ProductInfo(currentItemId, "VersionString");
                    }
                    catch //(COMException ex)
                    {
                        Console.WriteLine("Exception on currentItemId {0}", currentItemId);
                        //Console.WriteLine("Exception message {0}", ex.Message);
                        //Console.WriteLine("Exception innerException {0}", ex.InnerException);
                        //Console.WriteLine("Exception stackTrace {0}", ex.StackTrace);
                    }
                    if (currentItemVersion.Length == 0) currentItemVersion = " ";  // "<value not found>";
                    if (verbose)
                    {
                        string currentItemProductId = installer.get_ProductInfo(currentItemId, "ProductID");
                        if (currentItemProductId.Length == 0) currentItemProductId = " ";  // "<value not found>";
                        string currentItemInstallLocation = installer.get_ProductInfo(currentItemId, "InstallLocation");
                        if (currentItemInstallLocation.Length == 0) currentItemInstallLocation = " ";  // "<value not found>";
                        string format = "{0}, {1},\n{2}, {3}, {4}";
                        if (verboseSingleLine) format = "{0}, {1}, {2}, {3}, {4}";
                        Console.WriteLine(format, currentItemName, currentItemVersion, /* currentItemProductId */ currentItemId, 
                            currentItemInstallLocation, currentItemInstallSource);
                    }
                    else
                    {
                        Console.WriteLine("{0}, {1}", currentItemName, currentItemVersion);
                    }
                }
                Console.WriteLine("Total Msi Hits = {0}", installer.Products.Count);
                Console.WriteLine("------------- end ListProducts Msi --------------");
            }

            if (productType == ProductTypes.All || productType == ProductTypes.NonMsi)
            {
                RegistryKey rk = Registry.LocalMachine.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\UnInstall");
                try
                {
                    Console.WriteLine("----------- begin ListProducts NonMsi -----------");
                    int totalNonMsiHits = 0;
                    foreach (string subkey in rk.GetSubKeyNames())
                    //for (int i = 0; i < rk.SubKeyCount; i++)
                    {
                        //string subkey = rk.GetSubKeyNames().GetValue(i);
                        RegistryKey sk = rk.OpenSubKey(subkey);
                        if ((sk.GetValue("WindowsInstaller") != null && Convert.ToInt32(sk.GetValue("WindowsInstaller", 1)) == 0) ||
                            (sk.GetValue("WindowsInstaller") == null && sk.GetValue("DisplayName") != null && 
                             sk.GetValue("ParentKeyName") == null && RegExKb.Match(sk.GetValue("", "").ToString()).Success == false))
                        {
                            totalNonMsiHits++;
                            string currentItemId = /* sk.Name */ subkey;
                            string currentItemName = Convert.ToString(sk.GetValue("DisplayName"));
                            string currentItemInstallSource = Convert.ToString(sk.GetValue("InstallSource"));
                            if (currentItemInstallSource.Length == 0) currentItemInstallSource = " ";  // "<value not found>";
                            string currentItemVersion = Convert.ToString(sk.GetValue("DisplayVersion"));
                            if (currentItemVersion.Length == 0) currentItemVersion = " ";  // "<value not found>";

                            if (verbose)
                            {
                                string currentItemUninstallString = Convert.ToString(sk.GetValue("UninstallString"));
                                if (currentItemUninstallString.Length == 0) currentItemUninstallString = " ";  // "<value not found>";
                                string currentItemInstallLocation = Convert.ToString(sk.GetValue("InstallLocation"));
                                if (currentItemInstallLocation.Length == 0) currentItemInstallLocation = " ";  // "<value not found>";
                                string format = "{0}, {1}, {2},\n{3}, {4}, {5}";
                                if (verboseSingleLine) format = "{0}, {1}, {2}, {3}, {4}, {5}";
                                Console.WriteLine(format, currentItemId, currentItemName, currentItemVersion, currentItemUninstallString,
                                    currentItemInstallLocation, currentItemInstallSource);
                            }
                            else
                            {
                                Console.WriteLine("{0}, {1}, {2}", currentItemId, currentItemName, currentItemVersion);
                            }                            
                        }
                        sk.Close();
                    }
                    Console.WriteLine("Total NonMsi Hits = {0}", totalNonMsiHits++);
                    Console.WriteLine("------------- end ListProducts NonMsi --------------");
                }
                finally
                {
                    rk.Close();
                }

                rk = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\UnInstall");
                try
                {
                    Console.WriteLine("----------- begin ListProducts NonMsiCu -----------");
                    int totalNonMsiHits = 0;
                    foreach (string subkey in rk.GetSubKeyNames())
                    //for (int i = 0; i < rk.SubKeyCount; i++)
                    {
                        //string subkey = rk.GetSubKeyNames().GetValue(i);
                        RegistryKey sk = rk.OpenSubKey(subkey);
                        if ((sk.GetValue("WindowsInstaller") != null && Convert.ToInt32(sk.GetValue("WindowsInstaller", 1)) == 0) ||
                            (sk.GetValue("WindowsInstaller") == null && sk.GetValue("DisplayName") != null && 
                             sk.GetValue("ParentKeyName") == null && RegExKb.Match(sk.GetValue("", "").ToString()).Success == false))
                        {
                            totalNonMsiHits++;
                            string currentItemId = /* sk.Name */ subkey;
                            string currentItemName = Convert.ToString(sk.GetValue("DisplayName"));
                            string currentItemInstallSource = Convert.ToString(sk.GetValue("InstallSource"));
                            if (currentItemInstallSource.Length == 0) currentItemInstallSource = " ";  // "<value not found>";
                            string currentItemVersion = Convert.ToString(sk.GetValue("DisplayVersion"));
                            if (currentItemVersion.Length == 0) currentItemVersion = " ";  // "<value not found>";

                            if (verbose)
                            {
                                string currentItemUninstallString = Convert.ToString(sk.GetValue("UninstallString"));
                                if (currentItemUninstallString.Length == 0) currentItemUninstallString = " ";  // "<value not found>";
                                string currentItemInstallLocation = Convert.ToString(sk.GetValue("InstallLocation"));
                                if (currentItemInstallLocation.Length == 0) currentItemInstallLocation = " ";  // "<value not found>";
                                string format = "{0}, {1}, {2},\n{3}, {4}, {5}";
                                if (verboseSingleLine) format = "{0}, {1}, {2}, {3}, {4}, {5}";
                                Console.WriteLine(format, currentItemId, currentItemName, currentItemVersion, currentItemUninstallString,
                                    currentItemInstallLocation, currentItemInstallSource);
                            }
                            else
                            {
                                Console.WriteLine("{0}, {1}, {2}", currentItemId, currentItemName, currentItemVersion);
                            }                            
                        }
                        sk.Close();
                    }
                    Console.WriteLine("Total NonMsiCu Hits = {0}", totalNonMsiHits++);
                    Console.WriteLine("------------- end ListProducts NonMsiCu --------------");
                }
                finally
                {
                    rk.Close();
                }

                if (false == Environment.GetEnvironmentVariable("Processor_Architecture").Equals("x86"))
                {
                    rk = Registry.LocalMachine.OpenSubKey("Software\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\UnInstall");
                    try
                    {
                        Console.WriteLine("----------- begin ListProducts NonMsiX86 -----------");                        
                        int totalNonMsiX86Hits = 0;
                        foreach (string subkey in rk.GetSubKeyNames())
                        //for (int i = 0; i < rk.SubKeyCount; i++)
                        {
                            //string subkey = rk.GetSubKeyNames().GetValue(i);
                            RegistryKey sk = rk.OpenSubKey(subkey);
                            if ((sk.GetValue("WindowsInstaller") != null && Convert.ToInt32(sk.GetValue("WindowsInstaller", 1)) == 0) ||
                                (sk.GetValue("WindowsInstaller") == null && sk.GetValue("DisplayName") != null && 
                                 sk.GetValue("ParentKeyName") == null && RegExKb.Match(sk.GetValue("", "").ToString()).Success == false))
                            {
                                totalNonMsiX86Hits++;
                                string currentItemId = /* sk.Name */ subkey;
                                string currentItemName = Convert.ToString(sk.GetValue("DisplayName"));
                                string currentItemInstallSource = Convert.ToString(sk.GetValue("InstallSource"));
                                if (currentItemInstallSource.Length == 0) currentItemInstallSource = " ";  // "<value not found>";
                                string currentItemVersion = Convert.ToString(sk.GetValue("DisplayVersion"));
                                if (currentItemVersion.Length == 0) currentItemVersion = " ";  // "<value not found>";

                                if (verbose)
                                {
                                    string currentItemUninstallString = Convert.ToString(sk.GetValue("UninstallString"));
                                    if (currentItemUninstallString.Length == 0) currentItemUninstallString = " ";  // "<value not found>";
                                    string currentItemInstallLocation = Convert.ToString(sk.GetValue("InstallLocation"));
                                    if (currentItemInstallLocation.Length == 0) currentItemInstallLocation = " ";  // "<value not found>";
                                    string format = "{0}, {1}, {2},\n{3}, {4}, {5}";
                                    if (verboseSingleLine) format = "{0}, {1}, {2}, {3}, {4}, {5}";
                                    Console.WriteLine(format, currentItemId, currentItemName, currentItemVersion, currentItemUninstallString,
                                        currentItemInstallLocation, currentItemInstallSource);
                                }
                                else
                                {
                                    Console.WriteLine("{0}, {1}, {2}", currentItemId, currentItemName, currentItemVersion);
                                }
                            }
                            sk.Close();
                        }
                        Console.WriteLine("Total NonMsiX86 Hits = {0}", totalNonMsiX86Hits);
                        Console.WriteLine("------------- end ListProducts NonMsiX86 --------------");
                    }
                    finally
                    {
                        rk.Close();
                    }

                    //rk = Registry.CurrentUser.OpenSubKey("Software\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\UnInstall");
                    //try
                    //{
                    //    Console.WriteLine("----------- begin ListProducts NonMsiCuX86 -----------");                        
                    //    int totalNonMsiX86Hits = 0;
                    //    foreach (string subkey in rk.GetSubKeyNames())
                    //    //for (int i = 0; i < rk.SubKeyCount; i++)
                    //    {
                    //        //string subkey = rk.GetSubKeyNames().GetValue(i);
                    //        RegistryKey sk = rk.OpenSubKey(subkey);
                    //        if ((sk.GetValue("WindowsInstaller") != null && Convert.ToInt32(sk.GetValue("WindowsInstaller", 1)) == 0) ||
                    //            (sk.GetValue("WindowsInstaller") == null && sk.GetValue("DisplayName") != null && 
                    //             sk.GetValue("ParentKeyName") == null && RegExKb.Match(sk.GetValue("", "").ToString()).Success == false))
                    //        {
                    //            totalNonMsiX86Hits++;
                    //            string currentItemId = /* sk.Name */ subkey;
                    //            string currentItemName = Convert.ToString(sk.GetValue("DisplayName"));
                    //            string currentItemInstallSource = Convert.ToString(sk.GetValue("InstallSource"));
                    //            if (currentItemInstallSource.Length == 0) currentItemInstallSource = " ";  // "<value not found>";
                    //            string currentItemVersion = Convert.ToString(sk.GetValue("DisplayVersion"));
                    //            if (currentItemVersion.Length == 0) currentItemVersion = " ";  // "<value not found>";

                    //            if (verbose)
                    //            {
                    //                string currentItemUninstallString = Convert.ToString(sk.GetValue("UninstallString"));
                    //                if (currentItemUninstallString.Length == 0) currentItemUninstallString = " ";  // "<value not found>";
                    //                string currentItemInstallLocation = Convert.ToString(sk.GetValue("InstallLocation"));
                    //                if (currentItemInstallLocation.Length == 0) currentItemInstallLocation = " ";  // "<value not found>";
                    //                string format = "{0}, {1}, {2},\n{3}, {4}, {5}";
                    //                if (verboseSingleLine) format = "{0}, {1}, {2}, {3}, {4}, {5}";
                    //                Console.WriteLine(format, currentItemId, currentItemName, currentItemVersion, currentItemUninstallString,
                    //                    currentItemInstallLocation, currentItemInstallSource);
                    //            }
                    //            else
                    //            {
                    //                Console.WriteLine("{0}, {1}, {2}", currentItemId, currentItemName, currentItemVersion);
                    //            }
                    //        }
                    //        sk.Close();
                    //    }
                    //    Console.WriteLine("Total NonMsiCuX86 Hits = {0}", totalNonMsiX86Hits);
                    //    Console.WriteLine("------------- end ListProducts NonMsiCuX86 --------------");
                    //}
                    //finally
                    //{
                    //    rk.Close();
                    //}
                }
            }

            if (productType == ProductTypes.All || productType == ProductTypes.Mu)
            {
                RegistryKey rk = Registry.LocalMachine.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Products");
// then interate to find every Installer\UserData\S-1-5-18\Products\<guid>\Patches\<guid>\DisplayName entry
                try
                {
                    Console.WriteLine("----------- begin ListProducts Microsoft Update -----------");

                    int totalMicrosoftUpdateHits = 0;
                    foreach (string subkey in rk.GetSubKeyNames())
                    {
                        RegistryKey skPatches = rk.OpenSubKey(subkey + "\\Patches");
                        foreach (string subkeyPatchesGuid in skPatches.GetSubKeyNames())
                        {
                            RegistryKey skPatchesGuid = skPatches.OpenSubKey(subkeyPatchesGuid);
                            if (skPatchesGuid.GetValue("DisplayName") != null /* && skPatchesGuid.GetValue("State") != null */ &&
                                skPatchesGuid.GetValue("State").ToString() == "1")
                            {
                                totalMicrosoftUpdateHits++;
                                string currentItemId = /* skPatchesGuid.Name */ subkeyPatchesGuid;
                                string currentItemName = Convert.ToString(skPatchesGuid.GetValue("DisplayName"));
                                string currentItemInstallSource = Convert.ToString(skPatchesGuid.GetValue("MoreInfoURL"));
                                if (currentItemInstallSource.Length == 0) currentItemInstallSource = " ";  // "<value not found>";
                                string currentItemVersion = Convert.ToString(skPatchesGuid.GetValue("Installed"));
                                if (currentItemVersion.Length == 0) currentItemVersion = " ";  // "<value not found>";

                                if (verbose)
                                {
                                    string /* currentItemUninstallString = Convert.ToString(skPatchesGuid.GetValue("MoreInfoURL"));
                                    if (currentItemUninstallString.Length == 0) */ currentItemUninstallString = " ";  // "<value not found>";
                                    string /* currentItemInstallLocation = Convert.ToString(skPatchesGuid.GetValue("MoreInfoURL"));
                                    if (currentItemInstallLocation.Length == 0) */ currentItemInstallLocation = " ";  // "<value not found>";
                                    string format = "{0}, {1},\n{2}, {3}, {4}";
                                    if (verboseSingleLine) format = "{0}, {1}, {2}, {3}, {4}";
                                    Console.WriteLine(format, currentItemName, currentItemVersion, currentItemUninstallString,
                                        currentItemInstallLocation, currentItemInstallSource);
                                }
                                else
                                {
                                    Console.WriteLine("{0}, {1}", currentItemName, currentItemVersion);
                                }
                            }
                            skPatchesGuid.Close();
                        }
                        skPatches.Close();
                    }
                    Console.WriteLine("Total Microsoft Update Hits = {0}", totalMicrosoftUpdateHits);
                    Console.WriteLine("------------- end ListProducts Microsoft Update --------------");
                }
                finally
                {
                    rk.Close();
                }
            }

            if (productType == ProductTypes.All || productType == ProductTypes.Wu)
            {
                RegistryKey rk = Registry.LocalMachine.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Component Based Servicing\\Packages");
// then interate to find every Component Based Servicing\Packages\Package_*\Visibility = 1/visible versus 2/hidden entry
                try
                {
                    Console.WriteLine("----------- begin ListProducts Windows Update -----------");

                    int totalWindowsUpdateHits = 0;
                    foreach (string subkey in rk.GetSubKeyNames())
                    {
                        if (subkey.StartsWith("Package_"))
                        {
                            RegistryKey sk = rk.OpenSubKey(subkey);
                            if (sk.GetValue("Visibility") != null && sk.GetValue("Visibility").ToString() == "1")
                            {
                                totalWindowsUpdateHits++;
                                //string currentItemId = sk.GetValue("InstallName").ToString();
                                string currentItemId = sk.Name;  // InstallName is "update.mum" on w08r2sp1/w7 installs so use key name
                                string currentItemName = RegExKb.Match(currentItemId).Groups[0].Value;
                                if (string.IsNullOrEmpty(currentItemName)) currentItemName = currentItemId;
                                string currentItemInstallSource = Convert.ToString(sk.GetValue("InstallLocation"));
                                if (currentItemInstallSource.Length == 0) currentItemInstallSource = " ";  // "<value not found>";
                                int installTimeHigh = Convert.ToInt32(sk.GetValue("InstallTimeHigh"));
                                int installTimeLow = Convert.ToInt32(sk.GetValue("InstallTimeLow"));
                                Int64 fileTime = ((Int64)installTimeHigh << 32) + installTimeLow;
                                string currentItemVersion = DateTime.FromFileTime(fileTime).ToString(DateTimeFormatString);  
                                if (currentItemVersion.Length == 0) currentItemVersion = " ";  // "<value not found>";

                                if (verbose)
                                {
                                    string /* currentItemUninstallString = Convert.ToString(skPatchesGuid.GetValue("InstallLocation"));
                                    if (currentItemUninstallString.Length == 0) */ currentItemUninstallString = " ";  // "<value not found>";
                                    string /* currentItemInstallLocation = Convert.ToString(skPatchesGuid.GetValue("InstallLocation"));
                                    if (currentItemInstallLocation.Length == 0) */ currentItemInstallLocation = " ";  // "<value not found>";
                                    string format = "{0}, {1},\n{2}, {3}, {4}";
                                    if (verboseSingleLine) format = "{0}, {1}, {2}, {3}, {4}";
                                    Console.WriteLine(format, currentItemName, currentItemVersion, currentItemUninstallString,
                                        currentItemInstallLocation, currentItemInstallSource);
                                }
                                else
                                {
                                    Console.WriteLine("{0}, {1}", currentItemName, currentItemVersion);
                                }
                            }
                            sk.Close();
                        }
                    }
                    Console.WriteLine("Total Windows Update Hits = {0}", totalWindowsUpdateHits);
                    Console.WriteLine("------------- end ListProducts Windows Update --------------");
                }
                finally
                {
                    rk.Close();
                }
            }
        }

        private static void SearchProducts(ProductTypes productType, string searchText, SearchType searchType, ArgsSupported argsFound)
        {
            bool verbose = (argsFound & ArgsSupported.Verbose) == ArgsSupported.Verbose;
            bool verboseSingleLine = (argsFound & ArgsSupported.VerboseSingleLine) == ArgsSupported.VerboseSingleLine;

            if (productType == ProductTypes.All || productType == ProductTypes.Msi)
            {
                Console.WriteLine("----------- begin SearchProducts Msi ------------");

#if DEBUG
                //Type typeFromProgID = Type.GetTypeFromProgID("WindowsInstaller.Installer");
                //Type typeFromInterop = typeof(MsiRcw.Installer);
#endif
                MsiRcw.Installer installer = (MsiRcw.Installer)new Installer();

                int totalMsiHits = 0;
                //foreach (string product in installer.Products)
                for (int i = 0; i < installer.Products.Count; i++)
                {
                    //string currentItemId = product;
                    string currentItemId = installer.Products[i];
                    string currentItemName = string.Empty;
                    string currentItemInstallSource = string.Empty;
                    string currentItemVersion = string.Empty;
                    string currentItemProductId = string.Empty;
                    try
                    {
                        currentItemName = installer.get_ProductInfo(currentItemId, "InstalledProductName");
                        currentItemInstallSource = installer.get_ProductInfo(currentItemId, "InstallSource");
                        if (currentItemInstallSource.Length == 0) currentItemInstallSource = " ";  // "<value not found>";
                        currentItemVersion = installer.get_ProductInfo(currentItemId, "VersionString");
                        if (currentItemVersion.Length == 0) currentItemVersion = " ";  // "<value not found>";

                        currentItemProductId = installer.get_ProductInfo(currentItemId, "ProductID");
                        if (currentItemProductId.Length == 0) currentItemProductId = " ";  // "<value not found>";
                    }
                    catch //(Exception ex)
                    {
                        //Console.WriteLine("Exception on currentItemId {0} with message {1}", currentItemId, ex.Message);
                        //Console.WriteLine("Exception on currentItemId {0} with stackTrace {1}", currentItemId, ex.StackTrace);
                        Console.WriteLine("Exception on currentItemId {0}", currentItemId);
                    }

#if DEBUG
                    //if (currentItemName.Contains("Silverlight")) System.Diagnostics.Debugger.Break();
                    //if (currentItemInstallSource.Contains("v11.0.40927\\packages")) System.Diagnostics.Debugger.Break();
                    //if (currentItemName.Contains("Multi-Targeting Pack")) System.Diagnostics.Debugger.Break();
                    //if (currentItemName.Contains("Blend")) System.Diagnostics.Debugger.Break();
#endif

                    if ((searchType == SearchType.Name && Regex.Match(currentItemName, searchText, RegexOptions.IgnoreCase).Success == true) ||
                        (searchType == SearchType.Source && Regex.Match(currentItemInstallSource, searchText, RegexOptions.IgnoreCase).Success == true) ||
                        (searchType == SearchType.Version && Regex.Match(currentItemVersion, searchText, RegexOptions.IgnoreCase).Success == true))
                    {
                        totalMsiHits++;
                        if (verbose)
                        {
                            string currentItemInstallLocation = installer.get_ProductInfo(currentItemId, "InstallLocation");
                            if (currentItemInstallLocation.Length == 0) currentItemInstallLocation = " ";  // "<value not found>";
                            string format = "{0}, {1},\n{2}, {3}, {4}";
                            if (verboseSingleLine) format = "{0}, {1}, {2}, {3}, {4}";
                            Console.WriteLine(format, currentItemName, currentItemVersion, /* currentItemProductId */ currentItemId,
                                currentItemInstallLocation, currentItemInstallSource);
                        }
                        else
                        {
                            Console.WriteLine("{0}, {1}", currentItemName, currentItemVersion);
                        }
                        if ((argsFound & ArgsSupported.RemoveProducts) == ArgsSupported.RemoveProducts)
                        {
                            MsiUninstall(currentItemId);
                        }
                    }
                }
                Console.WriteLine("Total Msi Hits = {0}", totalMsiHits);
                Console.WriteLine("------------ end SearchProducts Msi -------------");
            }

            if (productType == ProductTypes.All || productType == ProductTypes.NonMsi)
            {
                RegistryKey rk = Registry.LocalMachine.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\UnInstall");
                try
                {
                    Console.WriteLine("---------- begin SearchProducts NonMsi ----------");

                    int totalNonMsiHits = 0;
                    foreach (string subkey in rk.GetSubKeyNames())
                    //for (int i = 0; i < rk.SubKeyCount; i++)
                    {
                        //string subkey = rk.GetSubKeyNames().GetValue(i);
                        RegistryKey sk = rk.OpenSubKey(subkey);
                        if ((sk.GetValue("WindowsInstaller") != null && Convert.ToInt32(sk.GetValue("WindowsInstaller", 1)) == 0) ||
                            (sk.GetValue("WindowsInstaller") == null && sk.GetValue("DisplayName") != null && 
                             sk.GetValue("ParentKeyName") == null && RegExKb.Match(sk.GetValue("", "").ToString()).Success == false))
                        {
                            string currentItemId = /* sk.Name */ subkey;
                            string currentItemName = Convert.ToString(sk.GetValue("DisplayName"));
                            string currentItemInstallSource = Convert.ToString(sk.GetValue("InstallSource"));
                            if (currentItemInstallSource.Length == 0) currentItemInstallSource = " ";  // "<value not found>";
                            string currentItemVersion = Convert.ToString(sk.GetValue("DisplayVersion"));
                            if (currentItemVersion.Length == 0) currentItemVersion = " ";  // "<value not found>";

                            if ((searchType == SearchType.Name && Regex.Match(currentItemId, searchText, RegexOptions.IgnoreCase).Success == true) ||
                                (searchType == SearchType.Source && Regex.Match(currentItemInstallSource, searchText, RegexOptions.IgnoreCase).Success == true) ||
                                (searchType == SearchType.Version && Regex.Match(currentItemVersion, searchText, RegexOptions.IgnoreCase).Success == true))
                            {
                                totalNonMsiHits++;
                                if (verbose)
                                {
                                    string currentItemUninstallString = Convert.ToString(sk.GetValue("UninstallString"));
                                    if (currentItemUninstallString.Length == 0) currentItemUninstallString = " ";  // "<value not found>";
                                    string currentItemInstallLocation = Convert.ToString(sk.GetValue("InstallLocation"));
                                    if (currentItemInstallLocation.Length == 0) currentItemInstallLocation = " ";  // "<value not found>";                        
                                    string format = "{0}, {1}, {2},\n{3}, {4}, {5}";
                                    if (verboseSingleLine) format = "{0}, {1}, {2}, {3}, {4}, {5}";
                                    Console.WriteLine(format, currentItemId, currentItemName, currentItemVersion, currentItemUninstallString,
                                        currentItemInstallLocation, currentItemInstallSource);
                                }
                                else
                                {
                                    Console.WriteLine("{0}, {1}, {2}", currentItemId, currentItemName, currentItemVersion);
                                }
                            }
                        }
                        sk.Close();
                    }
                    Console.WriteLine("Total NonMsi Hits = {0}", totalNonMsiHits);
                    Console.WriteLine("----------- end SearchProducts NonMsi -----------");
                }
                finally
                {
                    rk.Close();
                }

                rk = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\UnInstall");
                try
                {
                    Console.WriteLine("---------- begin SearchProducts NonMsiCu ----------");

                    int totalNonMsiHits = 0;
                    foreach (string subkey in rk.GetSubKeyNames())
                    //for (int i = 0; i < rk.SubKeyCount; i++)
                    {
                        //string subkey = rk.GetSubKeyNames().GetValue(i);
                        RegistryKey sk = rk.OpenSubKey(subkey);
                        if ((sk.GetValue("WindowsInstaller") != null && Convert.ToInt32(sk.GetValue("WindowsInstaller", 1)) == 0) ||
                            (sk.GetValue("WindowsInstaller") == null && sk.GetValue("DisplayName") != null && 
                             sk.GetValue("ParentKeyName") == null && RegExKb.Match(sk.GetValue("", "").ToString()).Success == false))
                        {
                            string currentItemId = /* sk.Name */ subkey;
                            string currentItemName = Convert.ToString(sk.GetValue("DisplayName"));
                            string currentItemInstallSource = Convert.ToString(sk.GetValue("InstallSource"));
                            if (currentItemInstallSource.Length == 0) currentItemInstallSource = " ";  // "<value not found>";
                            string currentItemVersion = Convert.ToString(sk.GetValue("DisplayVersion"));
                            if (currentItemVersion.Length == 0) currentItemVersion = " ";  // "<value not found>";

                            if ((searchType == SearchType.Name && Regex.Match(currentItemId, searchText, RegexOptions.IgnoreCase).Success == true) ||
                                (searchType == SearchType.Source && Regex.Match(currentItemInstallSource, searchText, RegexOptions.IgnoreCase).Success == true) ||
                                (searchType == SearchType.Version && Regex.Match(currentItemVersion, searchText, RegexOptions.IgnoreCase).Success == true))
                            {
                                totalNonMsiHits++;
                                if (verbose)
                                {
                                    string currentItemUninstallString = Convert.ToString(sk.GetValue("UninstallString"));
                                    if (currentItemUninstallString.Length == 0) currentItemUninstallString = " ";  // "<value not found>";
                                    string currentItemInstallLocation = Convert.ToString(sk.GetValue("InstallLocation"));
                                    if (currentItemInstallLocation.Length == 0) currentItemInstallLocation = " ";  // "<value not found>";                        
                                    string format = "{0}, {1}, {2},\n{3}, {4}, {5}";
                                    if (verboseSingleLine) format = "{0}, {1}, {2}, {3}, {4}, {5}";
                                    Console.WriteLine(format, currentItemId, currentItemName, currentItemVersion, currentItemUninstallString,
                                        currentItemInstallLocation, currentItemInstallSource);
                                }
                                else
                                {
                                    Console.WriteLine("{0}, {1}, {2}", currentItemId, currentItemName, currentItemVersion);
                                }
                            }
                        }
                        sk.Close();
                    }
                    Console.WriteLine("Total NonMsiCu Hits = {0}", totalNonMsiHits);
                    Console.WriteLine("----------- end SearchProducts NonMsiCu -----------");
                }
                finally
                {
                    rk.Close();
                }

                if (false == Environment.GetEnvironmentVariable("Processor_Architecture").Equals("x86"))
                {
                    rk = Registry.LocalMachine.OpenSubKey("Software\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\UnInstall");
                    try
                    {
                        Console.WriteLine("---------- begin SearchProducts NonMsiX86 ----------");

                        int totalNonMsiX86Hits = 0;
                        foreach (string subkey in rk.GetSubKeyNames())
                        //for (int i = 0; i < rk.SubKeyCount; i++)
                        {
                            //string subkey = rk.GetSubKeyNames().GetValue(i);
                            RegistryKey sk = rk.OpenSubKey(subkey);
                            if ((sk.GetValue("WindowsInstaller") != null && Convert.ToInt32(sk.GetValue("WindowsInstaller", 1)) == 0) ||
                                (sk.GetValue("WindowsInstaller") == null && sk.GetValue("DisplayName") != null && 
                                 sk.GetValue("ParentKeyName") == null && RegExKb.Match(sk.GetValue("", "").ToString()).Success == false))
                            {
                                string currentItemId = /* sk.Name */ subkey;
                                string currentItemName = Convert.ToString(sk.GetValue("DisplayName"));
                                string currentItemInstallSource = Convert.ToString(sk.GetValue("InstallSource"));
                                if (currentItemInstallSource.Length == 0) currentItemInstallSource = " ";  // "<value not found>";
                                string currentItemVersion = Convert.ToString(sk.GetValue("DisplayVersion"));
                                if (currentItemVersion.Length == 0) currentItemVersion = " ";  // "<value not found>";

                                if ((searchType == SearchType.Name && Regex.Match(currentItemId, searchText, RegexOptions.IgnoreCase).Success == true) ||
                                    (searchType == SearchType.Source && Regex.Match(currentItemInstallSource, searchText, RegexOptions.IgnoreCase).Success == true) ||
                                    (searchType == SearchType.Version && Regex.Match(currentItemVersion, searchText, RegexOptions.IgnoreCase).Success == true))
                                {
                                    totalNonMsiX86Hits++;
                                    if (verbose)
                                    {
                                        string currentItemUninstallString = Convert.ToString(sk.GetValue("UninstallString"));
                                        if (currentItemUninstallString.Length == 0) currentItemUninstallString = " ";  // "<value not found>";
                                        string currentItemInstallLocation = Convert.ToString(sk.GetValue("InstallLocation"));
                                        if (currentItemInstallLocation.Length == 0) currentItemInstallLocation = " ";  // "<value not found>";
                                        string format = "{0}, {1}, {2},\n{3}, {4}, {5}";
                                        if (verboseSingleLine) format = "{0}, {1}, {2}, {3}, {4}, {5}";
                                        Console.WriteLine(format, currentItemId, currentItemName, currentItemVersion, currentItemUninstallString,
                                            currentItemInstallLocation, currentItemInstallSource);
                                    }
                                    else
                                    {
                                        Console.WriteLine("{0}, {1}, {2}", currentItemId, currentItemName, currentItemVersion);
                                    }
                                }
                            }
                            sk.Close();
                        }
                        Console.WriteLine("Total NonMsiX86 Hits = {0}", totalNonMsiX86Hits);
                        Console.WriteLine("----------- end SearchProducts NonMsiX86 -----------");
                    }
                    finally
                    {
                        rk.Close();
                    }

                    //rk = Registry.CurrentUser.OpenSubKey("Software\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\UnInstall");
                    //try
                    //{
                    //    Console.WriteLine("---------- begin SearchProducts NonMsiCuX86 ----------");

                    //    int totalNonMsiX86Hits = 0;
                    //    foreach (string subkey in rk.GetSubKeyNames())
                    //    //for (int i = 0; i < rk.SubKeyCount; i++)
                    //    {
                    //        //string subkey = rk.GetSubKeyNames().GetValue(i);
                    //        RegistryKey sk = rk.OpenSubKey(subkey);
                    //        if ((sk.GetValue("WindowsInstaller") != null && Convert.ToInt32(sk.GetValue("WindowsInstaller", 1)) == 0) ||
                    //            (sk.GetValue("WindowsInstaller") == null && sk.GetValue("DisplayName") != null && 
                    //             sk.GetValue("ParentKeyName") == null && RegExKb.Match(sk.GetValue("", "").ToString()).Success == false))
                    //        {
                    //            string currentItemId = /* sk.Name */ subkey;
                    //            string currentItemName = Convert.ToString(sk.GetValue("DisplayName"));
                    //            string currentItemInstallSource = Convert.ToString(sk.GetValue("InstallSource"));
                    //            if (currentItemInstallSource.Length == 0) currentItemInstallSource = " ";  // "<value not found>";
                    //            string currentItemVersion = Convert.ToString(sk.GetValue("DisplayVersion"));
                    //            if (currentItemVersion.Length == 0) currentItemVersion = " ";  // "<value not found>";

                    //            if ((searchType == SearchType.Name && Regex.Match(currentItemId, searchText, RegexOptions.IgnoreCase).Success == true) ||
                    //                (searchType == SearchType.Source && Regex.Match(currentItemInstallSource, searchText, RegexOptions.IgnoreCase).Success == true) ||
                    //                (searchType == SearchType.Version && Regex.Match(currentItemVersion, searchText, RegexOptions.IgnoreCase).Success == true))
                    //            {
                    //                totalNonMsiX86Hits++;
                    //                if (verbose)
                    //                {
                    //                    string currentItemUninstallString = Convert.ToString(sk.GetValue("UninstallString"));
                    //                    if (currentItemUninstallString.Length == 0) currentItemUninstallString = " ";  // "<value not found>";
                    //                    string currentItemInstallLocation = Convert.ToString(sk.GetValue("InstallLocation"));
                    //                    if (currentItemInstallLocation.Length == 0) currentItemInstallLocation = " ";  // "<value not found>";
                    //                    string format = "{0}, {1}, {2},\n{3}, {4}, {5}";
                    //                    if (verboseSingleLine) format = "{0}, {1}, {2}, {3}, {4}, {5}";
                    //                    Console.WriteLine(format, currentItemId, currentItemName, currentItemVersion, currentItemUninstallString,
                    //                        currentItemInstallLocation, currentItemInstallSource);
                    //                }
                    //                else
                    //                {
                    //                    Console.WriteLine("{0}, {1}, {2}", currentItemId, currentItemName, currentItemVersion);
                    //                }
                    //            }
                    //        }
                    //        sk.Close();
                    //    }
                    //    Console.WriteLine("Total NonMsiCuX86 Hits = {0}", totalNonMsiX86Hits);
                    //    Console.WriteLine("----------- end SearchProducts NonMsiCuX86 -----------");
                    //}
                    //finally
                    //{
                    //    rk.Close();
                    //}
                }
            }

            if (productType == ProductTypes.All || productType == ProductTypes.Mu)
            {
                RegistryKey rk = Registry.LocalMachine.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Products");
// then interate to find every Installer\UserData\S-1-5-18\Products\<guid>\Patches\<guid>\DisplayName entry
                try
                {
                    Console.WriteLine("---------- begin SearchProducts Microsoft Update ----------");

                    int totalMicrosoftUpdateHits = 0;
                    foreach (string subkey in rk.GetSubKeyNames())
                    {
                        RegistryKey skPatches = rk.OpenSubKey(subkey + "\\Patches");
                        foreach (string subkeyPatchesGuid in skPatches.GetSubKeyNames())
                        {
                            RegistryKey skPatchesGuid = skPatches.OpenSubKey(subkeyPatchesGuid);
                            if (skPatchesGuid.GetValue("DisplayName") != null /* && skPatchesGuid.GetValue("State") != null */ &&
                                skPatchesGuid.GetValue("State").ToString() == "1")
                            {
                                string currentItemId = /* skPatchesGuid.Name */ subkeyPatchesGuid;
                                string currentItemName = Convert.ToString(skPatchesGuid.GetValue("DisplayName"));
                                string currentItemInstallSource = Convert.ToString(skPatchesGuid.GetValue("MoreInfoURL"));
                                if (currentItemInstallSource.Length == 0) currentItemInstallSource = " ";  // "<value not found>";
                                string currentItemVersion = Convert.ToString(skPatchesGuid.GetValue("Installed"));
                                if (currentItemVersion.Length == 0) currentItemVersion = " ";  // "<value not found>";

                                if ((searchType == SearchType.Name && Regex.Match(currentItemName, searchText, RegexOptions.IgnoreCase).Success == true) ||
                                    (searchType == SearchType.Source && Regex.Match(currentItemInstallSource, searchText, RegexOptions.IgnoreCase).Success == true) ||
                                    (searchType == SearchType.Version && Regex.Match(currentItemVersion, searchText, RegexOptions.IgnoreCase).Success == true))
                                {
                                    totalMicrosoftUpdateHits++;
                                    if (verbose)
                                    {
                                        string /* currentItemUninstallString = Convert.ToString(skPatchesGuid.GetValue("MoreInfoURL"));
                                        if (currentItemUninstallString.Length == 0) */ currentItemUninstallString = " ";  // "<value not found>";
                                        string /* currentItemInstallLocation = Convert.ToString(skPatchesGuid.GetValue("MoreInfoURL"));
                                        if (currentItemInstallLocation.Length == 0) */ currentItemInstallLocation = " ";  // "<value not found>";
                                        string format = "{0}, {1},\n{2}, {3}, {4}";
                                        if (verboseSingleLine) format = "{0}, {1}, {2}, {3}, {4}";
                                        Console.WriteLine(format, currentItemName /* currentItemId */, currentItemVersion, currentItemUninstallString,
                                            currentItemInstallLocation, currentItemInstallSource);
                                    }
                                    else
                                    {
                                        Console.WriteLine("{0}, {1}", currentItemName /* currentItemId */, currentItemVersion);
                                    }                                    
                                }
                            }
                            skPatchesGuid.Close();
                        }
                        skPatches.Close();
                    }
                    Console.WriteLine("Total Microsoft Update Hits = {0}", totalMicrosoftUpdateHits);
                    Console.WriteLine("----------- end SearchProducts Microsoft Update -----------");
                }
                finally
                {
                    rk.Close();
                }
            }

            if (productType == ProductTypes.All || productType == ProductTypes.Wu)
            {
                RegistryKey rk = Registry.LocalMachine.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Component Based Servicing\\Packages");
// then interate to find every Component Based Servicing\Packages\Package_*\Visibility = 1/visible versus 2/hidden entry
                try
                {
                    Console.WriteLine("----------- begin SearchProducts Windows Update -----------");

                    int totalWindowsUpdateHits = 0;
                    foreach (string subkey in rk.GetSubKeyNames())
                    {
                        if (subkey.StartsWith("Package_"))
                        {
                            RegistryKey sk = rk.OpenSubKey(subkey);
                            if (sk.GetValue("Visibility") != null && sk.GetValue("Visibility").ToString() == "1")
                            {
                                //string currentItemId = sk.GetValue("InstallName").ToString();  
                                string currentItemId = sk.Name;  // InstallName is "update.mum" on w08r2sp1/w7 installs so use key name
                                string currentItemName = RegExKb.Match(currentItemId).Groups[0].Value;
                                if (string.IsNullOrEmpty(currentItemName)) currentItemName = currentItemId;
                                string currentItemInstallSource = Convert.ToString(sk.GetValue("InstallLocation"));
                                if (currentItemInstallSource.Length == 0) currentItemInstallSource = " ";  // "<value not found>";
                                int installTimeHigh = Convert.ToInt32(sk.GetValue("InstallTimeHigh"));
                                int installTimeLow = Convert.ToInt32(sk.GetValue("InstallTimeLow"));
                                Int64 fileTime = ((Int64)installTimeHigh << 32) + installTimeLow;
                                string currentItemVersion = DateTime.FromFileTime(fileTime).ToString(DateTimeFormatString);
                                if (currentItemVersion.Length == 0) currentItemVersion = " ";  // "<value not found>";                                

                                if ((searchType == SearchType.Name && Regex.Match(currentItemName, searchText, RegexOptions.IgnoreCase).Success == true) ||
                                    (searchType == SearchType.Source && Regex.Match(currentItemInstallSource, searchText, RegexOptions.IgnoreCase).Success == true) ||
                                    (searchType == SearchType.Version && Regex.Match(currentItemVersion, searchText, RegexOptions.IgnoreCase).Success == true))
                                {
                                    totalWindowsUpdateHits++;
                                    if (verbose)
                                    {
                                        string /* currentItemUninstallString = Convert.ToString(skPatchesGuid.GetValue("InstallLocation"));
                                        if (currentItemUninstallString.Length == 0) */ currentItemUninstallString = " ";  // "<value not found>";
                                        string /* currentItemInstallLocation = Convert.ToString(skPatchesGuid.GetValue("InstallLocation"));
                                        if (currentItemInstallLocation.Length == 0) */ currentItemInstallLocation = " ";  // "<value not found>";
                                        string format = "{0}, {1},\n{2}, {3}, {4}";
                                        if (verboseSingleLine) format = "{0}, {1}, {2}, {3}, {4}";
                                        Console.WriteLine(format, currentItemName, currentItemVersion, currentItemUninstallString,
                                            currentItemInstallLocation, currentItemInstallSource);
                                    }
                                    else
                                    {
                                        Console.WriteLine("{0}, {1}", currentItemName, currentItemVersion);
                                    }                                    
                                }
                            }
                            sk.Close();
                        }
                    }
                    Console.WriteLine("Total Windows Update Hits = {0}", totalWindowsUpdateHits);
                    Console.WriteLine("------------- end SearchProducts Windows Update --------------");
                }
                finally
                {
                    rk.Close();
                }
            }
        }

        private static void MsiUninstall(string pid)
        {
            if (string.IsNullOrWhiteSpace(pid)) return;

            //var arguments = string.Format(" /x {0} /passive reboot=reallySuppress ignoredependencies=all /l* {1}\\{0}Uninst.log", pid,
            var arguments = string.Format(" /x {0} /qb reboot=reallySuppress ignoredependencies=all /l* {1}\\{0}Uninst.log", pid,
                Environment.GetEnvironmentVariable("temp"));
// including ignoredependencies=all property causes silent uninstalls when msi has chained ref counts, use at your own risk

            ProcessStartInfo processStartInfo = new ProcessStartInfo(Environment.GetEnvironmentVariable("windir") + 
                "\\system32\\msiexec.exe", arguments);
// redirect standard output to the Process.StandardOutput StreamReader
            processStartInfo.RedirectStandardOutput = true; processStartInfo.UseShellExecute = false; 
// do not create the black window
            processStartInfo.CreateNoWindow = true;
            Process processs = Process.Start(processStartInfo);
            processs.WaitForExit();
// optionally get the standard output result
            //string standardOutputResult = processs.StandardOutput.ReadToEnd();
        }

        private static bool IsProductInstalled(ProductTypes productType, string productName)
        {
            bool rv = false;

            if (productType == ProductTypes.All || productType == ProductTypes.Msi)
            {
                MsiRcw.Installer installer = (MsiRcw.Installer)new Installer();

                //foreach (string product in installer.Products)
                for (int i = 0; i < installer.Products.Count && rv == false; i++)
                {
                    //string currentItemId = product;
                    string currentItemId = installer.Products[i];
                    string currentItemName = installer.get_ProductInfo(currentItemId, "InstalledProductName");
                    if (productName.ToLower() == currentItemName.ToLower())
                    {
                        MsiRcw.MsiInstallState state = installer.get_ProductState(currentItemId);
                        if (state == MsiRcw.MsiInstallState.msiInstallStateDefault)
                        {
                            rv = true;
                        }
                    }
                }
            }

            if (productType == ProductTypes.All || productType == ProductTypes.NonMsi)
            {
                RegistryKey rk = Registry.LocalMachine.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\UnInstall");
                try
                {
                    //foreach (string subkey in rk.GetSubKeyNames())
                    for (int i = 0; i < rk.SubKeyCount && rv == false; i++)
                    {
                        string subkey = rk.GetSubKeyNames().GetValue(i).ToString();
                        RegistryKey sk = rk.OpenSubKey(subkey);
                        if (sk.GetValue("WindowsInstaller") == null)
                        {
                            string currentItemId = /* sk.Name */ subkey;
                            string currentItemName = Convert.ToString(sk.GetValue("DisplayName"));
                            if (productName.ToLower() == /* currentItemName */ currentItemId.ToLower())
                            {
                                rv = true;
                            }
                        }
                    }
                }
                finally
                {
                    rk.Close();
                }

                if (false == Environment.GetEnvironmentVariable("Processor_Architecture").Equals("x86"))
                {
                    rk = Registry.LocalMachine.OpenSubKey("Software\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\UnInstall");

                    try
                    {
                        //foreach (string subkey in rk.GetSubKeyNames())
                        for (int i = 0; i < rk.SubKeyCount && rv == false; i++)
                        {
                            string subkey = rk.GetSubKeyNames().GetValue(i).ToString();
                            RegistryKey sk = rk.OpenSubKey(subkey);
                            if (sk.GetValue("WindowsInstaller") == null)
                            {
                                string currentItemId = /* sk.Name */ subkey;
                                string currentItemName = Convert.ToString(sk.GetValue("DisplayName"));
                                if (productName.ToLower() == /* currentItemName */ currentItemId.ToLower())
                                {
                                    rv = true;
                                }
                            }
                        }
                    }
                    finally
                    {
                        rk.Close();
                    }
                }
            }

            return rv;
        }

        private static string GetProductVersion(ProductTypes productType, string productName)
        {
            string rv = String.Empty;

            if (productType == ProductTypes.All || productType == ProductTypes.Msi)
            {
                MsiRcw.Installer installer = (MsiRcw.Installer)new Installer();

                //foreach (string product in installer.Products)
                for (int i = 0; i < installer.Products.Count && rv.Length == 0; i++)
                {
                    //string currentItemId = product;
                    string currentItemId = installer.Products[i];
                    string currentItemName = installer.get_ProductInfo(currentItemId, "InstalledProductName");
                    if (productName.ToLower() == currentItemName.ToLower())
                    {
                        MsiRcw.MsiInstallState state = installer.get_ProductState(currentItemId);
                        if (state == MsiRcw.MsiInstallState.msiInstallStateDefault)
                        {
                            string version = installer.get_ProductInfo(currentItemId, "VersionString");
                            if (version.Length > 0) rv = version;
                            else rv = "<value not found>";
                        }
                    }
                }
            }

            if (productType == ProductTypes.All || productType == ProductTypes.NonMsi)
            {
                RegistryKey rk = Registry.LocalMachine.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\UnInstall");

                try
                {
                    //foreach (string subkey in rk.GetSubKeyNames())
                    for (int i = 0; i < rk.SubKeyCount && rv.Length == 0; i++)
                    {
                        string subkey = rk.GetSubKeyNames().GetValue(i).ToString();
                        RegistryKey sk = rk.OpenSubKey(subkey);
                        if (sk.GetValue("WindowsInstaller") == null)
                        {
                            string currentItemId = /* sk.Name */ subkey;
                            string currentItemName = Convert.ToString(sk.GetValue("DisplayName"));
                            if (productName.ToLower() == /* currentItemName */ currentItemId.ToLower())
                            {
                                string version = Convert.ToString(sk.GetValue("DisplayVersion"));
                                if (version.Length > 0) rv = version;
                                else rv = "<value not found>";
                            }
                        }
                    }
                }
                finally
                {
                    rk.Close();
                }

                if (false == Environment.GetEnvironmentVariable("Processor_Architecture").Equals("x86"))
                {
                    rk = Registry.LocalMachine.OpenSubKey("Software\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\UnInstall");

                    try
                    {
                        //foreach (string subkey in rk.GetSubKeyNames())
                        for (int i = 0; i < rk.SubKeyCount && rv.Length == 0; i++)
                        {
                            string subkey = rk.GetSubKeyNames().GetValue(i).ToString();
                            RegistryKey sk = rk.OpenSubKey(subkey);
                            if (sk.GetValue("WindowsInstaller") == null)
                            {
                                string currentItemId = /* sk.Name */ subkey;
                                string currentItemName = Convert.ToString(sk.GetValue("DisplayName"));
                                if (productName.ToLower() == /* currentItemName */ currentItemId.ToLower())
                                {
                                    string version = Convert.ToString(sk.GetValue("DisplayVersion"));
                                    if (version.Length > 0)
                                    {
                                        rv = version;
                                    }
                                    else
                                    {
                                        rv = "<value not found>";
                                    }
                                }
                            }
                        }
                    }
                    finally
                    {
                        rk.Close();
                    }
                }
            }

            return rv;
        }
    }
}
