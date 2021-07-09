using System.Xml;

using CrystalDecisions.ReportAppServer.ClientDoc;

using ExtensionMethods;

namespace rpt_diff
{
    /*
    *  ReportClientDocumentModel
    *  - newer
    *  - contains all atributes availible from rpt file
    *  - can be used in future to merge changes or back convert from xml to rpt
    */
    class ReportClientDocumentModel
    {
        public static void ProcessReport(ISCDReportClientDocument report, XmlWriter xmlw, string reportDoc = "ReportClientDocument")
        {
            xmlw.WriteStartElement(reportDoc);
            xmlw.WriteAttributeString("DisplayName", report.DisplayName);
            xmlw.WriteAttributeString("IsModified", report.IsModified.ToStringSafe());
            xmlw.WriteAttributeString("IsOpen", report.IsOpen.ToStringSafe());
            xmlw.WriteAttributeString("IsReadOnly", report.IsReadOnly.ToStringSafe());
            xmlw.WriteAttributeString("LocaleID", report.LocaleID.ToStringSafe());
            xmlw.WriteAttributeString("MajorVersion", report.MajorVersion.ToStringSafe());
            xmlw.WriteAttributeString("MinorVersion", report.MinorVersion.ToStringSafe());
            xmlw.WriteAttributeString("Path", report.Path);
            xmlw.WriteAttributeString("PreferredViewingLocaleID", report.PreferredViewingLocaleID.ToStringSafe());
            xmlw.WriteAttributeString("ProductLocaleID", report.ProductLocaleID.ToStringSafe());
            xmlw.WriteAttributeString("ReportAppServer", report.ReportAppServer);
            Controllers.ProcessCustomFunctionController(report.CustomFunctionController, xmlw);
            Controllers.ProcessDatabaseController(report.DatabaseController, xmlw);
            Controllers.ProcessDataDefController(report.DataDefController, xmlw);
            Controllers.ProcessPrintOutputController(report.PrintOutputController, xmlw);
            Controllers.ProcessReportDefController(report.ReportDefController, xmlw);
            ReportDefModel.ProcessReportOptions(report.ReportOptions, xmlw);
            Controllers.ProcessSubreportController(report.SubreportController, xmlw);
            DataDefModel.ProcessSummaryInfo(report.SummaryInfo, xmlw);
            xmlw.WriteEndElement();
        }
    }
}
