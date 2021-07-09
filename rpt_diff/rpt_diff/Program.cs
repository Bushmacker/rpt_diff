using System;
using System.Diagnostics;
using System.IO;
using Gnu.Getopt;

namespace rpt_diff
{
    static class Program
    {
        /*
         * All return codes
         */
        enum ExitCode : int
        {
            Success = 0,
            WrongDiffApp = 1,
            WrongRptFile = 2,
            WrongArgs = 3,
            WrongModel = 4,
            ConvertError = 5
        }
        /*
         * The main entry point for the application.
         */
        [STAThread]
        static int Main(string[] args)
        {
            string diffPath = "", xmlFilePath1 = "", xmlFilePath2 = "", rptFile1 = "", rptFile2 = "", rptFolder = "";
            int ModelNumber = 1; // vychozi je ReportClientDocument

            Getopt opt = new Getopt("", args, "mp:f:s:d:");
            int c;
            while ((c = opt.getopt()) != (-1))
            {
                switch ((char)c)
                {
                    case 'p':
                        if (!File.Exists(opt.Optarg))
                        {
                            Console.Error.WriteLine("Error: Unknown diff application - Bad DiffUtilPath");
                            WriteUsage();
                            return (int)ExitCode.WrongDiffApp;
                        }
                        else
                        {
                            diffPath = opt.Optarg;
                        }
                        break;
                    case 'f':
                        if (!File.Exists(opt.Optarg))
                        {
                            Console.Error.WriteLine("Error: Can't find RPT file - Bad RPTPath1");
                            WriteUsage();
                            return (int)ExitCode.WrongRptFile;
                        }
                        else
                        {
                            rptFile1 = opt.Optarg;
                        }
                        break;
                    case 's':
                        if (!File.Exists(opt.Optarg))
                        {
                            Console.Error.WriteLine("Error: Can't find RPT file - Bad RPTPath2");
                            WriteUsage();
                            return (int)ExitCode.WrongRptFile;
                        }
                        else
                        {
                            rptFile2 = opt.Optarg;
                        }
                        break;
                    case 'm':
                        // pokud je zadan prepinac m tak pouzijem model ReportDocument
                        ModelNumber = 0;
                        break;
                    case 'd':
                        if (!Directory.Exists(opt.Optarg))
                        {
                            Console.Error.WriteLine("Error: Can't find RPT directory - Bad RPTDirectory");
                            WriteUsage();
                            return (int)ExitCode.WrongRptFile;
                        }
                        else
                        {
                            rptFolder = opt.Optarg;
                        }
                        break;
                    default:
                        WriteUsage();
                        return (int)ExitCode.WrongArgs;
                }
            }

            try
            {
                if (File.Exists(rptFile1))
                {
                    Console.WriteLine("Using object model: \"" + ((ModelNumber == 0) ? "ReportDocument" : "ReportClientDocument") + "\"");
                    Console.WriteLine("Converting file: \"" + rptFile1 + "\"");
                    xmlFilePath1 = RptToXml.ConvertRptToXml(rptFile1, ModelNumber);
                    Console.WriteLine("File \"" + rptFile1 + "\" converted to \"" + xmlFilePath1 + "\"");
                }
                if (File.Exists(rptFile2))
                {
                    Console.WriteLine("Using object model: \"" + ((ModelNumber == 0) ? "ReportDocument" : "ReportClientDocument") + "\"");
                    Console.WriteLine("Converting file: \"" + rptFile2 + "\"");
                    xmlFilePath2 = RptToXml.ConvertRptToXml(rptFile2, ModelNumber);
                    Console.WriteLine("File \"" + rptFile2 + "\" converted to \"" + xmlFilePath2 + "\"");
                }
                if (Directory.Exists(rptFolder))
                {
                    string[] files = Directory.GetFiles(rptFolder, "*.rpt", SearchOption.AllDirectories);
                    string xmlFilePath;
                    foreach (string file in files)
                    {
                        Console.WriteLine("Using object model: \"" + ((ModelNumber == 0) ? "ReportDocument" : "ReportClientDocument") + "\"");
                        Console.WriteLine("Converting file: \"" + file + "\"");
                        xmlFilePath = RptToXml.ConvertRptToXml(file, ModelNumber);
                        Console.WriteLine("File \"" + file + "\" converted to \"" + xmlFilePath + "\"");
                    }

                }
            }
            catch (Exception e)
            {
                Console.Error.WriteLine("Error: Convert to XML error");
                Console.Error.WriteLine(e);
                return (int)ExitCode.ConvertError;
            }
            if (File.Exists(diffPath) && (File.Exists(xmlFilePath1) || File.Exists(xmlFilePath2)))
            {
                Console.WriteLine("Starting diff application: \"" + diffPath + "\"");
                Process diffProc = Process.Start(diffPath, String.Format("\"{0}\" \"{1}\"", xmlFilePath1, xmlFilePath2));
                diffProc.WaitForExit();
                Console.WriteLine("Diff application exited");
                // delete xml files after closing diff application
                if (File.Exists(xmlFilePath1))
                {
                    Console.WriteLine("Deleting file: \"" + xmlFilePath1 + "\"");
                    File.Delete(xmlFilePath1);
                }
                if (File.Exists(xmlFilePath2))
                {
                    Console.WriteLine("Deleting file: \"" + xmlFilePath2 + "\"");
                    File.Delete(xmlFilePath2);
                }
            }

            return (int)ExitCode.Success;
        }
        static void WriteUsage()
        {
            Console.WriteLine("Usage: rpt_diff.exe -m -p DiffUtilPath -f RPTPath1 -s RPTPath2 -d RPTDirectory");
            Console.WriteLine("       -m - Select object model ReportDocument | ReportClientDocument is default if -m not used ");
            Console.WriteLine("       -p DiffUtilPath - Full path to external diff application .exe file that can compare two xml files (for example KDiff)");
            Console.WriteLine("       -f RPTPath1 - Full path to first .rpt file to be converted to xml");
            Console.WriteLine("       -s RPTPath2 - Full path to second .rpt file to be converted to xml and compared with first file");
            Console.WriteLine("       -d RPTDirectory - Directory to be processed, all of *.rpt files in diretory and its subdirectories will be converted to XML");
        }

    }
}
