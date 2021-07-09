using System;
using System.Xml;

using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;

using ExtensionMethods;

namespace rpt_diff
{
    /*
    *  ReportDocumentModel
    *  - older
    *  - not include custom formulas 
    */
    class ReportDocumentModel
    {
        public static void ProcessReport(ReportDocument report, XmlWriter xmlw, string reportDoc = "ReportDocument")
        {
            xmlw.WriteStartElement(reportDoc);
            if (!report.IsSubreport) // not supported in subreportt
            {
                xmlw.WriteAttributeString("DefaultXmlExportSelection", report.DefaultXmlExportSelection.ToStringSafe());
                xmlw.WriteAttributeString("FileName", report.FileName);
                xmlw.WriteAttributeString("FilePath", report.FilePath);
                xmlw.WriteAttributeString("HasSavedData", report.HasSavedData.ToStringSafe());
                xmlw.WriteAttributeString("IsLoaded", report.IsLoaded.ToStringSafe());
                xmlw.WriteAttributeString("IsRPTR", report.IsRPTR.ToStringSafe());
                xmlw.WriteAttributeString("ReportAppServer", report.ReportAppServer);// <EMBEDDED_REPORT_ENGINE> 
            }
            xmlw.WriteAttributeString("IsSubreport", report.IsSubreport.ToStringSafe());
            xmlw.WriteAttributeString("Name", report.Name);
            xmlw.WriteAttributeString("RecordSelectionFormula", report.RecordSelectionFormula);

            ProcessDatabase(report.Database, xmlw);
            ProcessDataDefinition(report.DataDefinition, xmlw);
            ProcessDataSourceConnections(report.DataSourceConnections, xmlw);
            ProcessReportDefinition(report.ReportDefinition, xmlw);
            if (!report.IsSubreport) // not supported in subreportt
            {
                ProcessPrintOptions(report.PrintOptions, xmlw);
                ProcessReportOptions(report.ReportOptions, xmlw);
                ProcessSubreports(report.Subreports, xmlw);
                ProcessSummaryInfo(report.SummaryInfo, xmlw);
                ProcessSavedXmlExportFormats(report.SavedXmlExportFormats, xmlw);
            }
            xmlw.WriteEndElement();
        }
        private static void ProcessDatabase(Database db, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("Database");
            ProcessTables(db.Tables, xmlw);
            ProcessLinks(db.Links, xmlw);
            xmlw.WriteEndElement();
        }
        private static void ProcessTables(Tables tbls, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("Tables");
            xmlw.WriteAttributeString("Count", tbls.Count.ToStringSafe());
            foreach (Table tbl in tbls)
            {
                ProcessTable(tbl, xmlw);
                tbl.Dispose();
            }
            xmlw.WriteEndElement();
        }
        private static void ProcessTable(Table tbl, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("Table");
            xmlw.WriteAttributeString("Location", tbl.Location);
            xmlw.WriteAttributeString("Name", tbl.Name);
            ProcessDatabaseFieldDefinitions(tbl.Fields, xmlw);
            ProcessTableLogOnInfo(tbl.LogOnInfo, xmlw);
            xmlw.WriteEndElement();
        }
        private static void ProcessDatabaseFieldDefinitions(DatabaseFieldDefinitions flds, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("DatabaseFieldDefinitions");
            xmlw.WriteAttributeString("Count", flds.Count.ToStringSafe());
            foreach (DatabaseFieldDefinition fld in flds)
            {
                ProcessDatabaseFieldDefinition(fld, xmlw);
                fld.Dispose();
            }
            xmlw.WriteEndElement();
        }
        private static void ProcessDatabaseFieldDefinition(DatabaseFieldDefinition fld, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("DatabaseFieldDefinition");
            xmlw.WriteAttributeString("FormulaName", fld.FormulaName);
            xmlw.WriteAttributeString("Kind", fld.Kind.ToStringSafe());
            xmlw.WriteAttributeString("Name", fld.Name);
            xmlw.WriteAttributeString("NumberOfBytes", fld.NumberOfBytes.ToStringSafe());
            xmlw.WriteAttributeString("TableName", fld.TableName);
            xmlw.WriteAttributeString("UseCount", fld.UseCount.ToStringSafe());
            xmlw.WriteAttributeString("ValueType", fld.ValueType.ToStringSafe());
            xmlw.WriteEndElement();
        }
        private static void ProcessTableLogOnInfo(TableLogOnInfo tloi, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("TableLogOnInfo");
            xmlw.WriteAttributeString("ReportName", tloi.ReportName);
            xmlw.WriteAttributeString("TableName", tloi.TableName);
            ProcessConnectionInfo(tloi.ConnectionInfo, xmlw);
            xmlw.WriteEndElement();
        }
        private static void ProcessConnectionInfo(ConnectionInfo ci, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("ConnectionInfo");
            xmlw.WriteAttributeString("AllowCustomConnection", ci.AllowCustomConnection.ToStringSafe());
            xmlw.WriteAttributeString("DatabaseName", ci.DatabaseName);
            xmlw.WriteAttributeString("DBConnHandler", ci.DBConnHandler.ToStringSafe());
            xmlw.WriteAttributeString("IntegratedSecurity", ci.IntegratedSecurity.ToStringSafe());
            xmlw.WriteAttributeString("Password", ci.Password);
            xmlw.WriteAttributeString("ServerName", ci.ServerName);
            xmlw.WriteAttributeString("Type", ci.Type.ToStringSafe());
            xmlw.WriteAttributeString("UserID", ci.UserID);
            xmlw.WriteEndElement();
        }
        private static void ProcessLinks(TableLinks links, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("TableLinks");
            xmlw.WriteAttributeString("Count", links.Count.ToStringSafe());
            foreach (TableLink link in links)
            {
                ProcessTableLink(link, xmlw);
                link.Dispose();
            }

            xmlw.WriteEndElement();
        }
        private static void ProcessTableLink(TableLink link, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("TableLink");
            xmlw.WriteAttributeString("JoinType", link.JoinType.ToStringSafe());
            xmlw.WriteAttributeString("SourceTable", link.SourceTable.Name);
            xmlw.WriteAttributeString("DestinationTable", link.DestinationTable.Name);
            foreach (DatabaseFieldDefinition sfld in link.SourceFields)
            {
                xmlw.WriteStartElement("SourceField");
                xmlw.WriteAttributeString("FormulaName", sfld.FormulaName);
                xmlw.WriteEndElement();
                sfld.Dispose();
            }
            foreach (DatabaseFieldDefinition dfld in link.DestinationFields)
            {
                xmlw.WriteStartElement("DestinationField");
                xmlw.WriteAttributeString("FormulaName", dfld.FormulaName);
                xmlw.WriteEndElement();
                dfld.Dispose();
            }
            xmlw.WriteEndElement();
        }
        private static void ProcessDataDefinition(DataDefinition dd, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("DataDefinition");
            xmlw.WriteAttributeString("GroupSelectionFormula", dd.GroupSelectionFormula);
            xmlw.WriteAttributeString("GroupSelectionFormulaRaw", dd.GroupSelectionFormulaRaw);
            xmlw.WriteAttributeString("RecordSelectionFormula", dd.RecordSelectionFormula);
            xmlw.WriteAttributeString("RecordSelectionFormulaRaw", dd.RecordSelectionFormulaRaw);
            xmlw.WriteAttributeString("SavedDataSelectionFormula", dd.SavedDataSelectionFormula);
            xmlw.WriteAttributeString("SavedDataSelectionFormulaRaw", dd.SavedDataSelectionFormulaRaw);
            ProcessFormulaFieldDefinitions(dd.FormulaFields, xmlw);
            ProcessGroupNameFieldDefinitions(dd.GroupNameFields, xmlw);
            ProcessGroups(dd.Groups, xmlw);
            ProcessParameterFieldDefinitions(dd.ParameterFields, xmlw);
            ProcessRunningTotalFieldDefinitions(dd.RunningTotalFields, xmlw);
            ProcessSortFields(dd.SortFields, xmlw);
            ProcessSQLExpressionFields(dd.SQLExpressionFields, xmlw);
            ProcessSummaryFieldDefinitions(dd.SummaryFields, xmlw);

            xmlw.WriteEndElement();
        }
        private static void ProcessFormulaFieldDefinitions(FormulaFieldDefinitions ffds, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("FormulaFieldDefinitions");
            xmlw.WriteAttributeString("Count", ffds.Count.ToStringSafe());
            foreach (FormulaFieldDefinition ffd in ffds)
            {
                ProcessFormulaFieldDefinition(ffd, xmlw);
                ffd.Dispose();
            }
            xmlw.WriteEndElement();
        }
        private static void ProcessFormulaFieldDefinition(FormulaFieldDefinition ffd, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("FormulaFieldDefinition");
            xmlw.WriteAttributeString("FormulaName", ffd.FormulaName);
            xmlw.WriteAttributeString("Kind", ffd.Kind.ToStringSafe());
            xmlw.WriteAttributeString("Name", ffd.Name);
            xmlw.WriteAttributeString("NumberOfBytes", ffd.NumberOfBytes.ToStringSafe());
            xmlw.WriteAttributeString("Text", ffd.Text);
            xmlw.WriteAttributeString("UseCount", ffd.UseCount.ToStringSafe());
            xmlw.WriteAttributeString("ValueType", ffd.ValueType.ToStringSafe());
            xmlw.WriteEndElement();
        }
        private static void ProcessGroupNameFieldDefinitions(GroupNameFieldDefinitions gnfds, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("GroupNameFieldDefinitions");
            xmlw.WriteAttributeString("Count", gnfds.Count.ToStringSafe());
            foreach (GroupNameFieldDefinition gnfd in gnfds)
            {
                ProcessGroupNameFieldDefinition(gnfd, xmlw);
                gnfd.Dispose();
            }
            xmlw.WriteEndElement();
        }
        private static void ProcessGroupNameFieldDefinition(GroupNameFieldDefinition gnfd, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("GroupNameFieldDefinition");
            xmlw.WriteAttributeString("FormulaName", gnfd.FormulaName);
            xmlw.WriteAttributeString("Group", gnfd.Group.ConditionField.FormulaName);
            xmlw.WriteAttributeString("GroupNameFieldName", gnfd.GroupNameFieldName);
            xmlw.WriteAttributeString("Kind", gnfd.Kind.ToStringSafe());
            xmlw.WriteAttributeString("Name", gnfd.Name);
            xmlw.WriteAttributeString("NumberOfBytes", gnfd.NumberOfBytes.ToStringSafe());
            //xmlw.WriteAttributeString("UseCount", gnfd.UseCount.ToStringSafe()); //This member is now obsolete.
            xmlw.WriteAttributeString("ValueType", gnfd.ValueType.ToStringSafe());
            xmlw.WriteEndElement();
        }
        private static void ProcessGroups(Groups groups, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("Groups");
            xmlw.WriteAttributeString("Count", groups.Count.ToStringSafe());
            foreach (Group group in groups)
            {
                ProcessGroup(group, xmlw);
                group.Dispose();
            }
            xmlw.WriteEndElement();
        }
        private static void ProcessGroup(Group group, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("Group");
            xmlw.WriteAttributeString("ConditionField", group.ConditionField.FormulaName);
            xmlw.WriteAttributeString("GroupOptionsCondition", group.GroupOptions.Condition.ToStringSafe());
            xmlw.WriteEndElement();
        }
        private static void ProcessFieldDefinition(FieldDefinition cf, XmlWriter xmlw, string field)
        {
            xmlw.WriteStartElement(field);
            xmlw.WriteAttributeString("FormulaName", cf.FormulaName);
            xmlw.WriteAttributeString("Kind", cf.Kind.ToStringSafe());
            xmlw.WriteAttributeString("Name", cf.Name);
            xmlw.WriteAttributeString("NumberOfBytes", cf.NumberOfBytes.ToStringSafe());
            //xmlw.WriteAttributeString("UseCount", cf.UseCount.ToStringSafe());
            xmlw.WriteAttributeString("ValueType", cf.ValueType.ToStringSafe());
            xmlw.WriteEndElement();
        }

        private static void ProcessParameterFieldDefinitions(ParameterFieldDefinitions pfds, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("ParameterFieldDefinitions");
            xmlw.WriteAttributeString("Count", pfds.Count.ToStringSafe());
            foreach (ParameterFieldDefinition pfd in pfds)
            {
                ProcessParameterFieldDefinition(pfd, xmlw);
                pfd.Dispose();
            }
            xmlw.WriteEndElement();
        }
        private static void ProcessParameterFieldDefinition(ParameterFieldDefinition pfd, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("ParameterFieldDefinition");
            xmlw.WriteAttributeString("DefaultValueDisplayType", pfd.DefaultValueDisplayType.ToStringSafe());
            xmlw.WriteAttributeString("DefaultValueSortMethod", pfd.DefaultValueSortMethod.ToStringSafe());
            xmlw.WriteAttributeString("DefaultValueSortOrder", pfd.DefaultValueSortOrder.ToStringSafe());
            xmlw.WriteAttributeString("DiscreteOrRangeKind", pfd.DiscreteOrRangeKind.ToStringSafe());
            xmlw.WriteAttributeString("EditMask", pfd.EditMask);
            xmlw.WriteAttributeString("EnableAllowEditingDefaultValue", pfd.EnableAllowEditingDefaultValue.ToStringSafe());
            xmlw.WriteAttributeString("EnableAllowMultipleValue", pfd.EnableAllowMultipleValue.ToStringSafe());
            xmlw.WriteAttributeString("EnableNullValue", pfd.EnableNullValue.ToStringSafe());
            xmlw.WriteAttributeString("FormulaName", pfd.FormulaName);
            xmlw.WriteAttributeString("HasCurrentValue", pfd.HasCurrentValue.ToStringSafe());
            xmlw.WriteAttributeString("IsOptionalPrompt", pfd.IsOptionalPrompt.ToStringSafe());
            try
            {
                xmlw.WriteAttributeString("IsLinked", pfd.IsLinked().ToStringSafe());
            }
            catch (NotSupportedException) //IsLinked not supported in subreport
            { }

            xmlw.WriteAttributeString("Kind", pfd.Kind.ToStringSafe());
            xmlw.WriteAttributeString("MaximumValue", pfd.MaximumValue.ToStringSafe());
            xmlw.WriteAttributeString("MinimumValue", pfd.MinimumValue.ToStringSafe());
            xmlw.WriteAttributeString("Name", pfd.Name);
            xmlw.WriteAttributeString("NumberOfBytes", pfd.NumberOfBytes.ToStringSafe());
            xmlw.WriteAttributeString("ParameterFieldName", pfd.ParameterFieldName);
            xmlw.WriteAttributeString("ParameterFieldUsage2", pfd.ParameterFieldUsage2.ToStringSafe());
            xmlw.WriteAttributeString("ParameterType", pfd.ParameterType.ToStringSafe());
            xmlw.WriteAttributeString("ParameterValueKind", pfd.ParameterValueKind.ToStringSafe());
            xmlw.WriteAttributeString("PromptText", pfd.PromptText);
            xmlw.WriteAttributeString("ReportName", pfd.ReportName);
            xmlw.WriteAttributeString("UseCount", pfd.UseCount.ToStringSafe());
            xmlw.WriteAttributeString("ValueType", pfd.ValueType.ToStringSafe());
            ProcessParameterValues(pfd.CurrentValues, xmlw, "CurrentValues");
            ProcessParameterValues(pfd.DefaultValues, xmlw, "DefaultValues");
            xmlw.WriteEndElement();
        }
        private static void ProcessParameterValues(ParameterValues pvs, XmlWriter xmlw, string values)
        {
            xmlw.WriteStartElement(values);
            xmlw.WriteAttributeString("Count", pvs.Count.ToStringSafe());
            xmlw.WriteAttributeString("Capacity", pvs.Capacity.ToStringSafe());
            xmlw.WriteAttributeString("IsNoValue", pvs.IsNoValue.ToStringSafe());
            foreach (ParameterValue pv in pvs)
            {
                ProcessParameterValue(pv, xmlw, values.Remove(values.Length - 1, 1));
            }
            xmlw.WriteEndElement();
        }
        private static void ProcessParameterValue(ParameterValue pv, XmlWriter xmlw, string value)
        {
            xmlw.WriteStartElement(value);
            xmlw.WriteAttributeString("Description", pv.Description);
            xmlw.WriteAttributeString("IsRange", pv.IsRange.ToStringSafe());
            xmlw.WriteAttributeString("Kind", pv.Kind.ToStringSafe());
            if (pv.IsRange)
            {
                ParameterRangeValue prv = (ParameterRangeValue)pv;
                xmlw.WriteAttributeString("EndValue", prv.EndValue.ToStringSafe());
                xmlw.WriteAttributeString("LowerBoundType", prv.LowerBoundType.ToStringSafe());
                xmlw.WriteAttributeString("StartValue", prv.StartValue.ToStringSafe());
                xmlw.WriteAttributeString("UpperBoundType", prv.UpperBoundType.ToStringSafe());
            }
            else
            {
                ParameterDiscreteValue pdv = (ParameterDiscreteValue)pv;
                xmlw.WriteAttributeString("Value", pdv.Value.ToStringSafe());
            }
            xmlw.WriteEndElement();
        }
        private static void ProcessRunningTotalFieldDefinitions(RunningTotalFieldDefinitions rtfds, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("RunningTotalFieldDefinitions");
            xmlw.WriteAttributeString("Count", rtfds.Count.ToStringSafe());
            foreach (RunningTotalFieldDefinition rtfd in rtfds)
            {
                ProcessRunningTotalFieldDefinition(rtfd, xmlw);
                rtfd.Dispose();
            }
            xmlw.WriteEndElement();
        }
        private static void ProcessRunningTotalFieldDefinition(RunningTotalFieldDefinition rtfd, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("RunningTotalFieldDefinition");
            xmlw.WriteAttributeString("EvaluationCondition", ProcessCondition(rtfd.EvaluationCondition));
            xmlw.WriteAttributeString("EvaluationConditionType", rtfd.EvaluationConditionType.ToStringSafe());
            xmlw.WriteAttributeString("IsPercentageSummary", rtfd.IsPercentageSummary.ToStringSafe());
            xmlw.WriteAttributeString("FormulaName", rtfd.FormulaName);
            xmlw.WriteAttributeString("Kind", rtfd.Kind.ToStringSafe());
            xmlw.WriteAttributeString("Name", rtfd.Name);
            xmlw.WriteAttributeString("NumberOfBytes", rtfd.NumberOfBytes.ToStringSafe());
            xmlw.WriteAttributeString("Operation", rtfd.Operation.ToStringSafe());
            xmlw.WriteAttributeString("OperationParameter", rtfd.OperationParameter.ToStringSafe());
            xmlw.WriteAttributeString("ResetCondition", ProcessCondition(rtfd.ResetCondition));
            xmlw.WriteAttributeString("ResetConditionType", rtfd.ResetConditionType.ToStringSafe());
            xmlw.WriteAttributeString("SummarizedField", rtfd.SummarizedField.FormulaName);
            xmlw.WriteAttributeString("UseCount", rtfd.UseCount.ToStringSafe());
            xmlw.WriteAttributeString("ValueType", rtfd.ValueType.ToStringSafe());
            if (rtfd.Group != null)
            {
                xmlw.WriteAttributeString("Group", rtfd.Group.ConditionField.FormulaName);
            }
            if (rtfd.SecondarySummarizedField != null)
            {
                xmlw.WriteAttributeString("SecondarySummarizedField", rtfd.SecondarySummarizedField.FormulaName);
            }
            xmlw.WriteEndElement();
        }
        private static string ProcessCondition(object condition)
        {
            FieldDefinition dfdCondition = condition as FieldDefinition;
            Group gCondition = condition as Group;
            string formName;
            if (dfdCondition != null) // Field
            {
                formName = dfdCondition.FormulaName;
            }
            else if (gCondition != null) // Group
            {
                formName = gCondition.ConditionField.FormulaName;
            }
            else // Custom formula
            {
                formName = condition.ToStringSafe();
            }
            dfdCondition.Dispose();
            gCondition.Dispose();
            return formName;
        }

        private static void ProcessSortFields(SortFields sfs, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("SortFields");
            xmlw.WriteAttributeString("Count", sfs.Count.ToStringSafe());
            foreach (SortField sf in sfs)
            {
                ProcessSortField(sf, xmlw);
                sf.Dispose();
            }
            xmlw.WriteEndElement();
        }
        private static void ProcessSortField(SortField sf, XmlWriter xmlw)
        {
            TopBottomNSortField tbnsf = sf as TopBottomNSortField;
            if (tbnsf != null)
            {
                xmlw.WriteStartElement("TopBottomNSortField");
                xmlw.WriteAttributeString("EnableDiscardOtherGroups", tbnsf.EnableDiscardOtherGroups.ToStringSafe());
                xmlw.WriteAttributeString("Field", tbnsf.Field.FormulaName);
                xmlw.WriteAttributeString("NotInTopBottomNName", tbnsf.NotInTopBottomNName);
                xmlw.WriteAttributeString("NumberOfTopOrBottomNGroups", tbnsf.NumberOfTopOrBottomNGroups.ToStringSafe());
                xmlw.WriteAttributeString("SortDirection", tbnsf.SortDirection.ToStringSafe());
                xmlw.WriteAttributeString("SortType", tbnsf.SortType.ToStringSafe());
                tbnsf.Dispose();
            }
            else
            {
                xmlw.WriteStartElement("SortField");
                xmlw.WriteAttributeString("Field", sf.Field.FormulaName);
                xmlw.WriteAttributeString("SortDirection", sf.SortDirection.ToStringSafe());
                xmlw.WriteAttributeString("SortType", sf.SortType.ToStringSafe());
            }
            xmlw.WriteEndElement();
        }
        private static void ProcessSQLExpressionFields(SQLExpressionFieldDefinitions sexfds, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("SQLExpressionFieldDefinitions");
            xmlw.WriteAttributeString("Count", sexfds.Count.ToStringSafe());
            foreach (SQLExpressionFieldDefinition sefd in sexfds)
            {
                ProcessSQLExpressionFieldDefinition(sefd, xmlw);
                sefd.Dispose();
            }
            xmlw.WriteEndElement();
        }
        private static void ProcessSQLExpressionFieldDefinition(SQLExpressionFieldDefinition sefd, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("SQLExpressionFieldDefinition");
            xmlw.WriteAttributeString("FormulaName", sefd.FormulaName);
            xmlw.WriteAttributeString("Kind", sefd.Kind.ToStringSafe());
            xmlw.WriteAttributeString("Name", sefd.Name);
            xmlw.WriteAttributeString("NumberOfBytes", sefd.NumberOfBytes.ToStringSafe());
            xmlw.WriteAttributeString("Text", sefd.Text);
            xmlw.WriteAttributeString("UseCount", sefd.UseCount.ToStringSafe());
            xmlw.WriteAttributeString("ValueType", sefd.ValueType.ToStringSafe());
            xmlw.WriteEndElement();
        }
        private static void ProcessSummaryFieldDefinitions(SummaryFieldDefinitions sfds, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("SummaryFieldDefinitions");
            xmlw.WriteAttributeString("Count", sfds.Count.ToStringSafe());
            foreach (SummaryFieldDefinition sfd in sfds)
            {
                ProcessSummaryFieldDefinition(sfd, xmlw);
                sfd.Dispose();
            }
            xmlw.WriteEndElement();
        }
        private static void ProcessSummaryFieldDefinition(SummaryFieldDefinition sfd, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("SummaryFieldDefinition");
            xmlw.WriteAttributeString("FormulaName", sfd.FormulaName);
            xmlw.WriteAttributeString("Kind", sfd.Kind.ToStringSafe());
            xmlw.WriteAttributeString("Name", sfd.Name);
            xmlw.WriteAttributeString("NumberOfBytes", sfd.NumberOfBytes.ToStringSafe());
            xmlw.WriteAttributeString("Operation", sfd.Operation.ToStringSafe());
            xmlw.WriteAttributeString("OperationParameter", sfd.OperationParameter.ToStringSafe());
            xmlw.WriteAttributeString("SummarizedField", sfd.SummarizedField.FormulaName);
            xmlw.WriteAttributeString("UseCount", sfd.UseCount.ToStringSafe());
            xmlw.WriteAttributeString("ValueType", sfd.ValueType.ToStringSafe());
            if (sfd.Group != null)
            {
                xmlw.WriteAttributeString("Group", sfd.Group.ConditionField.FormulaName);
            }
            if (sfd.SecondarySummarizedField != null)
            {
                xmlw.WriteAttributeString("SecondarySummarizedField", sfd.SecondarySummarizedField.FormulaName);
            }
            xmlw.WriteEndElement();
        }
        private static void ProcessDataSourceConnections(DataSourceConnections dscs, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("DataSourceConnections");
            foreach (IConnectionInfo ici in dscs)
            {
                ProcessIConnectionInfo(ici, xmlw);
            }
            xmlw.WriteEndElement();
        }
        private static void ProcessIConnectionInfo(IConnectionInfo ici, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("DataSourceConnections");
            xmlw.WriteAttributeString("DBConnHandler", ici.DBConnHandler.ToStringSafe());
            xmlw.WriteAttributeString("DatabaseName", ici.DatabaseName);
            xmlw.WriteAttributeString("IntegratedSecurity", ici.IntegratedSecurity.ToStringSafe());
            xmlw.WriteAttributeString("Password", ici.Password);
            xmlw.WriteAttributeString("ServerName", ici.ServerName);
            xmlw.WriteAttributeString("Type", ici.Type.ToStringSafe());
            xmlw.WriteAttributeString("UserID", ici.UserID);
            xmlw.WriteEndElement();
        }
        private static void ProcessPrintOptions(PrintOptions po, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("PrintOptions");
            xmlw.WriteAttributeString("CustomPaperSource", po.CustomPaperSource.ToStringSafe());
            xmlw.WriteAttributeString("PageContentHeight", po.PageContentHeight.ToStringSafe());
            xmlw.WriteAttributeString("PageContentWidth", po.PageContentWidth.ToStringSafe());
            xmlw.WriteAttributeString("PaperOrientation", po.PaperOrientation.ToStringSafe());
            xmlw.WriteAttributeString("PaperSize", po.PaperSize.ToStringSafe());
            xmlw.WriteAttributeString("PaperSource", po.PaperSource.ToStringSafe());
            xmlw.WriteAttributeString("PrinterDuplex", po.PrinterDuplex.ToStringSafe());
            xmlw.WriteAttributeString("PrinterName", po.PrinterName);
            ProcessPageMargins(po.PageMargins, xmlw);
            xmlw.WriteEndElement();
        }
        private static void ProcessPageMargins(PageMargins pm, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("PageMargins");
            xmlw.WriteAttributeString("bottomMargin", pm.bottomMargin.ToStringSafe());
            xmlw.WriteAttributeString("leftMargin", pm.leftMargin.ToStringSafe());
            xmlw.WriteAttributeString("rightMargin", pm.rightMargin.ToStringSafe());
            xmlw.WriteAttributeString("topMargin", pm.topMargin.ToStringSafe());
            xmlw.WriteEndElement();
        }

        private static void ProcessReportDefinition(ReportDefinition rd, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("ReportDefinition");
            ProcessAreas(rd.Areas, xmlw);
            xmlw.WriteEndElement();
        }

        private static void ProcessAreas(Areas areas, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("Areas");
            xmlw.WriteAttributeString("Count", areas.Count.ToStringSafe());
            foreach (Area area in areas)
            {
                ProcessArea(area, xmlw);
                area.Dispose();
            }
            xmlw.WriteEndElement();
        }

        private static void ProcessArea(Area area, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("Area");
            xmlw.WriteAttributeString("Kind", area.Kind.ToStringSafe());
            xmlw.WriteAttributeString("Name", area.Name);
            ProcessAreaFormat(area.AreaFormat, xmlw);
            ProcessSections(area.Sections, xmlw);
            xmlw.WriteEndElement();
        }

        private static void ProcessAreaFormat(AreaFormat af, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("AreaFormat");
            xmlw.WriteAttributeString("EnableHideForDrillDown", af.EnableHideForDrillDown.ToStringSafe());
            xmlw.WriteAttributeString("EnableKeepTogether", af.EnableKeepTogether.ToStringSafe());
            xmlw.WriteAttributeString("EnableNewPageAfter", af.EnableNewPageAfter.ToStringSafe());
            xmlw.WriteAttributeString("EnableNewPageBefore", af.EnableNewPageBefore.ToStringSafe());
            xmlw.WriteAttributeString("EnablePrintAtBottomOfPage", af.EnablePrintAtBottomOfPage.ToStringSafe());
            xmlw.WriteAttributeString("EnableResetPageNumberAfter", af.EnableResetPageNumberAfter.ToStringSafe());
            xmlw.WriteAttributeString("EnableSuppress", af.EnableSuppress.ToStringSafe());
            xmlw.WriteEndElement();
        }

        private static void ProcessSections(Sections sections, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("Sections");
            xmlw.WriteAttributeString("Count", sections.Count.ToStringSafe());
            foreach (Section section in sections)
            {
                ProcessSection(section, xmlw);
                section.Dispose();
            }
            xmlw.WriteEndElement();
        }

        private static void ProcessSection(Section section, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("Section");
            xmlw.WriteAttributeString("Height", section.Height.ToStringSafe());
            xmlw.WriteAttributeString("Kind", section.Kind.ToStringSafe());
            xmlw.WriteAttributeString("Name", section.Name);
            ProcessReportObjects(section.ReportObjects, xmlw);
            ProcessSectionFormat(section.SectionFormat, xmlw);
            xmlw.WriteEndElement();
        }
        private static void ProcessReportObjects(ReportObjects ros, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("ReportObjects");
            xmlw.WriteAttributeString("Count", ros.Count.ToStringSafe());
            foreach (ReportObject ro in ros)
            {
                ProcessReportObjects(ro, xmlw);
                ro.Dispose();
            }
            xmlw.WriteEndElement();
        }
        private static void ProcessReportObjects(ReportObject ro, XmlWriter xmlw)
        {
            xmlw.WriteStartElement(ro.Kind.ToStringSafe());
            // for all types the same properties
            xmlw.WriteAttributeString("Height", ro.Height.ToStringSafe());
            xmlw.WriteAttributeString("Kind", ro.Kind.ToStringSafe());
            xmlw.WriteAttributeString("Left", ro.Left.ToStringSafe());
            xmlw.WriteAttributeString("Name", ro.Name);
            xmlw.WriteAttributeString("Top", ro.Top.ToStringSafe());
            xmlw.WriteAttributeString("Width", ro.Width.ToStringSafe());
            // kind specific properties
            switch (ro.Kind)
            {
                case ReportObjectKind.BlobFieldObject:
                    {
                        BlobFieldObject bfo = (BlobFieldObject)ro;
                        ProcessDatabaseFieldDefinition(bfo.DataSource, xmlw);
                        bfo.Dispose();
                        break;
                    }
                case ReportObjectKind.BoxObject:
                    {
                        BoxObject bo = (BoxObject)ro;
                        xmlw.WriteAttributeString("Bottom", bo.Bottom.ToStringSafe());
                        xmlw.WriteAttributeString("EnableExtendToBottomOfSection", bo.EnableExtendToBottomOfSection.ToStringSafe());
                        xmlw.WriteAttributeString("EndSectionName", bo.EndSectionName);
                        xmlw.WriteAttributeString("FillColor", bo.FillColor.ToStringSafe());
                        xmlw.WriteAttributeString("LineColor", bo.LineColor.ToStringSafe());
                        xmlw.WriteAttributeString("LineStyle", bo.LineStyle.ToStringSafe());
                        xmlw.WriteAttributeString("LineThickness", bo.LineThickness.ToStringSafe());
                        xmlw.WriteAttributeString("Right", bo.Right.ToStringSafe());
                        bo.Dispose();
                        break;
                    }
                case ReportObjectKind.FieldHeadingObject:
                    {
                        FieldHeadingObject fho = (FieldHeadingObject)ro;
                        xmlw.WriteAttributeString("Color", fho.Color.ToStringSafe());
                        xmlw.WriteAttributeString("Font", fho.Font.ToStringSafe());
                        xmlw.WriteAttributeString("Text", fho.Text);
                        xmlw.WriteAttributeString("FieldObjectName", fho.FieldObjectName);
                        fho.Dispose();
                        break;
                    }
                case ReportObjectKind.FieldObject:
                    {
                        FieldObject fo = (FieldObject)ro;
                        xmlw.WriteAttributeString("Color", fo.Color.ToStringSafe());
                        xmlw.WriteAttributeString("Font", fo.Font.ToStringSafe());
                        if (fo.DataSource != null)
                        {
                            ProcessFieldDefinition(fo.DataSource, xmlw, "DataSource");
                        }
                        ProcessFieldFormat(fo.FieldFormat, xmlw);
                        fo.Dispose();
                        break;
                    }
                case ReportObjectKind.LineObject:
                    {
                        LineObject lo = (LineObject)ro;
                        xmlw.WriteAttributeString("Bottom", lo.Bottom.ToStringSafe());
                        xmlw.WriteAttributeString("EnableExtendToBottomOfSection", lo.EnableExtendToBottomOfSection.ToStringSafe());
                        xmlw.WriteAttributeString("EndSectionName", lo.EndSectionName);
                        xmlw.WriteAttributeString("LineColor", lo.LineColor.ToStringSafe());
                        xmlw.WriteAttributeString("LineStyle", lo.LineStyle.ToStringSafe());
                        xmlw.WriteAttributeString("LineThickness", lo.LineThickness.ToStringSafe());
                        xmlw.WriteAttributeString("Right", lo.Right.ToStringSafe());
                        lo.Dispose();
                        break;
                    }
                case ReportObjectKind.SubreportObject:
                    {
                        SubreportObject so = (SubreportObject)ro;
                        xmlw.WriteAttributeString("EnableOnDemand", so.EnableOnDemand.ToStringSafe());
                        xmlw.WriteAttributeString("SubreportName", so.SubreportName);
                        ro.Dispose();
                        break;
                    }
                case ReportObjectKind.TextObject:
                    {
                        TextObject to = (TextObject)ro;
                        xmlw.WriteAttributeString("Color", to.Color.ToStringSafe());
                        xmlw.WriteAttributeString("Font", to.Font.ToStringSafe());
                        xmlw.WriteAttributeString("Text", to.Text);
                        to.Dispose();
                        break;
                    }
                default:
                    {
                        break;
                    }
            }
            ProcessBorder(ro.Border, xmlw);
            ProcessObjectFormat(ro.ObjectFormat, xmlw);
            xmlw.WriteEndElement();
        }

        private static void ProcessFieldFormat(FieldFormat ff, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("FieldFormat");
            //if (ff.BooleanFormat != null)
            //{
            ProcessBooleanFormat(ff.BooleanFormat, xmlw);
            //}
            //else if(ff.CommonFormat!=null)
            //{
            ProcessCommonFormat(ff.CommonFormat, xmlw);
            //}
            //else if (ff.DateTimeFormat != null)//datetime before date and time 
            //{
            ProcessDateTimeFormat(ff.DateTimeFormat, xmlw);
            //}
            //else if (ff.DateFormat != null)
            //{
            ProcessDateFormat(ff.DateFormat, xmlw);
            //}
            //else if (ff.NumericFormat != null)
            //{
            ProcessNumericFormat(ff.NumericFormat, xmlw);
            //}
            //else if (ff.TimeFormat != null)
            //{
            ProcessTimeFormat(ff.TimeFormat, xmlw);
            //}
            //else
            //{
            //    ((Action)(() => { }))();
            //}
            xmlw.WriteEndElement();
        }

        private static void ProcessBooleanFormat(BooleanFieldFormat bff, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("BooleanFieldFormat");
            xmlw.WriteAttributeString("OutputType", bff.OutputType.ToStringSafe());
            xmlw.WriteEndElement();
        }

        private static void ProcessCommonFormat(CommonFieldFormat cff, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("CommonFieldFormat");
            xmlw.WriteAttributeString("EnableSuppressIfDuplicated", cff.EnableSuppressIfDuplicated.ToStringSafe());
            xmlw.WriteAttributeString("EnableUseSystemDefaults", cff.EnableUseSystemDefaults.ToStringSafe());
            xmlw.WriteEndElement();
        }

        private static void ProcessDateTimeFormat(DateTimeFieldFormat dtff, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("DateTimeFieldFormat");
            xmlw.WriteAttributeString("DateTimeSeparator", dtff.DateTimeSeparator);
            ProcessDateFormat(dtff.DateFormat, xmlw);
            ProcessTimeFormat(dtff.TimeFormat, xmlw);
            xmlw.WriteEndElement();
        }

        private static void ProcessDateFormat(DateFieldFormat dff, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("DateFieldFormat");
            xmlw.WriteAttributeString("DayFormat", dff.DayFormat.ToStringSafe());
            xmlw.WriteAttributeString("MonthFormat", dff.MonthFormat.ToStringSafe());
            xmlw.WriteAttributeString("YearFormat", dff.YearFormat.ToStringSafe());
            xmlw.WriteEndElement();
        }

        private static void ProcessNumericFormat(NumericFieldFormat nff, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("NumericFieldFormat");
            xmlw.WriteAttributeString("CurrencySymbolFormat", nff.CurrencySymbolFormat.ToStringSafe());
            xmlw.WriteAttributeString("DecimalPlaces", nff.DecimalPlaces.ToStringSafe());
            xmlw.WriteAttributeString("EnableUseLeadingZero", nff.EnableUseLeadingZero.ToStringSafe());
            xmlw.WriteAttributeString("NegativeFormat", nff.NegativeFormat.ToStringSafe());
            xmlw.WriteAttributeString("RoundingFormat", nff.RoundingFormat.ToStringSafe());
            xmlw.WriteEndElement();
        }

        private static void ProcessTimeFormat(TimeFieldFormat tff, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("TimeFieldFormat");
            xmlw.WriteAttributeString("AMPMFormat", tff.AMPMFormat.ToStringSafe());
            xmlw.WriteAttributeString("AMString", tff.AMString);
            xmlw.WriteAttributeString("HourFormat", tff.HourFormat.ToStringSafe());
            xmlw.WriteAttributeString("HourMinuteSeparator", tff.HourMinuteSeparator);
            xmlw.WriteAttributeString("MinuteFormat", tff.MinuteFormat.ToStringSafe());
            xmlw.WriteAttributeString("MinuteSecondSeparator", tff.MinuteSecondSeparator);
            xmlw.WriteAttributeString("PMString", tff.PMString);
            xmlw.WriteAttributeString("SecondFormat", tff.SecondFormat.ToStringSafe());
            xmlw.WriteAttributeString("TimeBase", tff.TimeBase.ToStringSafe());
            xmlw.WriteEndElement();
        }

        private static void ProcessBorder(Border border, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("Border");
            xmlw.WriteAttributeString("BackgroundColor", border.BackgroundColor.ToStringSafe());
            xmlw.WriteAttributeString("BorderColor", border.BorderColor.ToStringSafe());
            xmlw.WriteAttributeString("BottomLineStyle", border.BottomLineStyle.ToStringSafe());
            xmlw.WriteAttributeString("HasDropShadow", border.HasDropShadow.ToStringSafe());
            xmlw.WriteAttributeString("LeftLineStyle", border.LeftLineStyle.ToStringSafe());
            xmlw.WriteAttributeString("RightLineStyle", border.RightLineStyle.ToStringSafe());
            xmlw.WriteAttributeString("TopLineStyle", border.TopLineStyle.ToStringSafe());
            xmlw.WriteEndElement();
        }

        private static void ProcessObjectFormat(ObjectFormat of, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("ObjectFormat");
            xmlw.WriteAttributeString("CssClass", of.CssClass);
            xmlw.WriteAttributeString("EnableCanGrow", of.EnableCanGrow.ToStringSafe());
            xmlw.WriteAttributeString("EnableCloseAtPageBreak", of.EnableCloseAtPageBreak.ToStringSafe());
            xmlw.WriteAttributeString("EnableKeepTogether", of.EnableKeepTogether.ToStringSafe());
            xmlw.WriteAttributeString("EnableSuppress", of.EnableSuppress.ToStringSafe());
            xmlw.WriteAttributeString("HorizontalAlignment", of.HorizontalAlignment.ToStringSafe());
            xmlw.WriteEndElement();
        }

        private static void ProcessSectionFormat(SectionFormat sf, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("SectionFormat");
            xmlw.WriteAttributeString("BackgroundColor", sf.BackgroundColor.ToStringSafe());
            xmlw.WriteAttributeString("CssClass", sf.CssClass);
            xmlw.WriteAttributeString("EnableKeepTogether", sf.EnableKeepTogether.ToStringSafe());
            xmlw.WriteAttributeString("EnableNewPageAfter", sf.EnableNewPageAfter.ToStringSafe());
            xmlw.WriteAttributeString("EnableNewPageBefore", sf.EnableNewPageBefore.ToStringSafe());
            xmlw.WriteAttributeString("EnablePrintAtBottomOfPage", sf.EnablePrintAtBottomOfPage.ToStringSafe());
            xmlw.WriteAttributeString("EnableResetPageNumberAfter", sf.EnableResetPageNumberAfter.ToStringSafe());
            xmlw.WriteAttributeString("EnableSuppress", sf.EnableSuppress.ToStringSafe());
            xmlw.WriteAttributeString("EnableSuppressIfBlank", sf.EnableSuppressIfBlank.ToStringSafe());
            xmlw.WriteAttributeString("EnableUnderlaySection", sf.EnableUnderlaySection.ToStringSafe());
            xmlw.WriteEndElement();
        }
        private static void ProcessReportOptions(ReportOptions ro, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("ReportOptions");
            xmlw.WriteAttributeString("EnableSaveDataWithReport", ro.EnableSaveDataWithReport.ToStringSafe());
            xmlw.WriteAttributeString("EnableSavePreviewPicture", ro.EnableSavePreviewPicture.ToStringSafe());
            xmlw.WriteAttributeString("EnableSaveSummariesWithReport", ro.EnableSaveSummariesWithReport.ToStringSafe());
            xmlw.WriteAttributeString("EnableUseDummyData", ro.EnableUseDummyData.ToStringSafe());
            xmlw.WriteAttributeString("InitialDataContext", ro.InitialDataContext);
            xmlw.WriteAttributeString("InitialReportPartName", ro.InitialReportPartName);
            xmlw.WriteEndElement();
        }
        private static void ProcessSubreports(Subreports subs, XmlWriter xmlw)
        {
            foreach (ReportDocument sub in subs)
            {
                ProcessReport(sub, xmlw, "Subreport");
                sub.Dispose();
            }
        }
        private static void ProcessSummaryInfo(SummaryInfo si, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("SummaryInfo");
            xmlw.WriteAttributeString("KeywordsInReport", si.KeywordsInReport);
            xmlw.WriteAttributeString("ReportAuthor", si.ReportAuthor);
            xmlw.WriteAttributeString("ReportComments", si.ReportComments);
            xmlw.WriteAttributeString("ReportSubject", si.ReportSubject);
            xmlw.WriteAttributeString("ReportTitle", si.ReportTitle);
            xmlw.WriteEndElement();
        }
        private static void ProcessSavedXmlExportFormats(XmlExportFormats xefs, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("XmlExportFormats");
            xmlw.WriteAttributeString("Count", xefs.Count.ToStringSafe());
            foreach (XmlExportFormat xef in xefs)
            {
                ProcessSavedXmlExportFormat(xef, xmlw);
                xef.Dispose();
            }
            xmlw.WriteEndElement();
        }
        private static void ProcessSavedXmlExportFormat(XmlExportFormat xef, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("XmlExportFormat");
            xmlw.WriteAttributeString("Description", xef.Description);
            xmlw.WriteAttributeString("ExportBlobField", xef.ExportBlobField.ToStringSafe());
            xmlw.WriteAttributeString("FileExtension", xef.FileExtension);
            xmlw.WriteAttributeString("Name", xef.Name);
            xmlw.WriteEndElement();
        }
    }
}
