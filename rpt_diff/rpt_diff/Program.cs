using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Drawing;
using System.Diagnostics;
using System.IO;

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
            WrongArgs    = 3,
            WrongModel   = 4,
            ConvertError = 5
        }
        /*
         * The main entry point for the application.
         */
        [STAThread]
        static int Main(string[] args)
        {
            string diffPath = "", xmlFilePath1 ="", xmlFilePath2 = "";
            int ModelNumber = 0;
            switch (args.Length)
            {
                case 3:
                    {
                        if (!File.Exists(args[1]))
                        {
                            Console.Error.WriteLine("Error: Can't find RPT file - Bad RPTPath1");
                            WriteUsage();
                            return (int)ExitCode.WrongRptFile;
                        }
                        if(!Int32.TryParse(args[2], out ModelNumber)||(ModelNumber!= 0&&ModelNumber!=1))
                        {   
                            Console.Error.WriteLine("Error: Wrong ModelNumber select - must be 0 or 1");
                            WriteUsage();
                            return (int)ExitCode.WrongModel;
                        }
                        Console.WriteLine("Using object model: \"" + ((ModelNumber==0)?"ReportDocument":"ReportClientDocument") + "\"");
                        Console.WriteLine("Converting file: \"" + args[1]+"\"");
                        try
                        {
                            xmlFilePath1 = RptToXml.ConvertRptToXml(args[1], ModelNumber);
                        }
                        catch(Exception e)
                        {
                            Console.Error.WriteLine("Error: Convert to XML error");
                            Console.Error.WriteLine(e);
                            return (int)ExitCode.ConvertError;
                        }
                        Console.WriteLine("File \"" + args[1] + "\" converted to \"" + xmlFilePath1 + "\"");
                        break;
                    }
                case 4:
                    {
                        if (!File.Exists(args[1]))
                        {
                            Console.Error.WriteLine("Error: Can't find RPT file - Bad RPTPath1");
                            WriteUsage();
                            return (int)ExitCode.WrongRptFile;
                        }
                        if(!Int32.TryParse(args[3], out ModelNumber)||(ModelNumber!= 0&&ModelNumber!=1))
                        {   
                            Console.Error.WriteLine("Error: Wrong ModelNumber select - must be 0 or 1");
                            WriteUsage();
                            return (int)ExitCode.WrongModel;
                        }
                        Console.WriteLine("Using object model: \"" + ((ModelNumber == 0) ? "ReportDocument" : "ReportClientDocument") + "\"");
                        Console.WriteLine("Converting file: \"" + args[1] + "\"");
                        try
                        {
                            xmlFilePath1 = RptToXml.ConvertRptToXml(args[1], ModelNumber);
                        }
                        catch (Exception e)
                        {
                            Console.Error.WriteLine("Error: Convert to XML error");
                            Console.Error.WriteLine(e);
                            return (int)ExitCode.ConvertError;
                        }
                        Console.WriteLine("File \"" + args[1] + "\" converted to \"" + xmlFilePath1+"\"");
                        if (!File.Exists(args[2]))
                        {
                            Console.Error.WriteLine("Error: Can't find RPT file - Bad RPTPath2");
                            WriteUsage();
                            return (int)ExitCode.WrongRptFile;
                        }
                        Console.WriteLine("Converting file: \"" + args[2] + "\"");
                        try
                        {
                            xmlFilePath2 = RptToXml.ConvertRptToXml(args[2], ModelNumber);
                        }
                        catch (Exception e)
                        {
                            Console.Error.WriteLine("Error: Convert to XML error");
                            Console.Error.WriteLine(e);
                            return (int)ExitCode.ConvertError;
                        }
                        Console.WriteLine("File \"" + args[2] + "\" converted to \"" + xmlFilePath2 + "\"");
                        break;
                    }
                default:
                    {
                        Console.Error.WriteLine("Error: Wrong number of parameters");
                        WriteUsage();
                        return (int)ExitCode.WrongArgs;
                    }
            }
            if (!File.Exists(args[0]))
            {
                Console.Error.WriteLine("Error: Unknown diff application - Bad DiffUtilPath");
                WriteUsage();
                return (int)ExitCode.WrongDiffApp;
            }
            diffPath = args[0];
            Console.WriteLine("Starting diff application: \"" + diffPath+"\"");
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
            
            return (int)ExitCode.Success;            
        }
        static void WriteUsage()
        {
            Console.WriteLine("Usage: rpt_diff.exe DiffUtilPath RPTPath1 [RPTPath2] ModelNumber");
            Console.WriteLine("       DiffUtilPath - Full path to external diff application .exe file that can compare two xml files (for example KDiff)");
            Console.WriteLine("       RPTPath1 - Full path to first .rpt file to be converted to xml");
            Console.WriteLine("       RPTPath2 - Full path to second .rpt file to be converted to xml and compared with first file");
            Console.WriteLine("       ModelNumber - Select between object model 0 - ReportDocument or 1 (recomended)- ReportClientDocument ");
        }
    }
}
