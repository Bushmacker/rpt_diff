using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

using CrystalDecisions.ReportAppServer.Controllers;

using ExtensionMethods;

namespace rpt_diff
{
    class Controllers
    {
        public static void ProcessCustomFunctionController(CustomFunctionController cfc, XmlWriter xmlw)
        {
            DataDefModel.ProcessCustomFunctions(cfc.GetCustomFunctions(), xmlw); 
        }
        public static void ProcessDatabaseController(DatabaseController dc, XmlWriter xmlw)
        {
            DataDefModel.ProcessDatabase(dc.Database, xmlw);
        }

        public static void ProcessDataDefController(DataDefController ddc, XmlWriter xmlw)
        {
            DataDefModel.ProcessDataDefinition(ddc.DataDefinition, xmlw);
        }


        public static void ProcessPrintOutputController(PrintOutputController poc, XmlWriter xmlw)
        {
            ReportDefModel.ProcessPrintOptions(poc.GetPrintOptions(), xmlw);
            ReportDefModel.ProcessSavedXMLExportFormats(poc.GetSavedXMLExportFormats(), xmlw);
        }

        public static void ProcessReportDefController(ReportDefController2 rdc, XmlWriter xmlw)
        {
            ReportDefModel.ProcessReportDefinition(rdc.ReportDefinition, xmlw);
        }

        public static void ProcessSubreportController(SubreportController sc, XmlWriter xmlw)
        {
            foreach (string Subreport in sc.GetSubreportNames())
            {
                xmlw.WriteStartElement("Subreport");            
                ProcessSubreportClientDocument(sc.GetSubreport(Subreport),xmlw);
                ReportDefModel.ProcessSubreportLinks(sc.GetSubreportLinks(Subreport), xmlw);
                xmlw.WriteEndElement();
            }
        }

        private static void ProcessSubreportClientDocument(SubreportClientDocument scd, XmlWriter xmlw)
        {
            xmlw.WriteAttributeString("EnableOnDemand", scd.EnableOnDemand.ToStringSafe());
            xmlw.WriteAttributeString("EnableReimport", scd.EnableReimport.ToStringSafe());
            xmlw.WriteAttributeString("IsImported", scd.IsImported.ToStringSafe());
            xmlw.WriteAttributeString("Name", scd.Name);
            xmlw.WriteAttributeString("SubreportLocation", scd.SubreportLocation);
            ProcessDatabaseController(scd.DatabaseController, xmlw);
            ProcessDataDefController(scd.DataDefController, xmlw);
            ProcessReportDefController(scd.ReportDefController, xmlw);
            ReportDefModel.ProcessReportOptions(scd.ReportOptions, xmlw);
        }

    }
}
