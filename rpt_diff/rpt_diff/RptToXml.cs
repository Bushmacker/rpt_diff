using System;
using System.Text;
using System.Xml;
using System.IO;

using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;

namespace rpt_diff
{
    class RptToXml
    {
        /*
         * ConvertRptToXml
         * Opens report file found in rptPath and converts it to xml using model specified by parameter model.
         * params:
         *  rptPath - full path to rpt file to be converted to xml
         *  model   - specifies which object model use to convert 
         *          - 0 = ReportDocumentModel
         *          - 1 = ReportClientDocumentModel (RAS)
         */
        public static string ConvertRptToXml(string rptPath, int model)
        {
            ReportDocument report = new ReportDocument();
            report.Load(rptPath, OpenReportMethod.OpenReportByTempCopy);
            string xmlPath = Path.ChangeExtension(rptPath, "xml");
            using (XmlTextWriter xmlw = new XmlTextWriter(xmlPath, Encoding.UTF8) { Formatting = Formatting.Indented })
            {
                xmlw.WriteStartDocument();
                if (model == 0)
                {
                    ReportDocumentModel.ProcessReport(report, xmlw);
                }
                else
                {
                    ReportClientDocumentModel.ProcessReport(report.ReportClientDocument, xmlw);
                }

                xmlw.WriteEndDocument();
                xmlw.Flush();
                xmlw.Close();
            }

            report.Close();
            report.Dispose();
            GC.Collect();
            return xmlPath;
        }


    }
}
