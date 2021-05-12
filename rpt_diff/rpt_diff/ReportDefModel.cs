using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

using CrystalDecisions.ReportAppServer.ReportDefModel;
using CrystalDecisions.ReportAppServer.DataDefModel;

using ExtensionMethods;

namespace rpt_diff
{
    class ReportDefModel
    {
        public static void ProcessReportOptions(ReportOptions ro, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("ReportOptions");
            xmlw.WriteAttributeString("CanSelectDistinctRecords", ro.CanSelectDistinctRecords.ToStringSafe());
            xmlw.WriteAttributeString("ConvertDateTimeType", ro.ConvertDateTimeType.ToStringSafe());
            xmlw.WriteAttributeString("ConvertNullFieldToDefault", ro.ConvertNullFieldToDefault.ToStringSafe());
            xmlw.WriteAttributeString("ConvertOtherNullsToDefault", ro.ConvertOtherNullsToDefault.ToStringSafe());
            xmlw.WriteAttributeString("EnableAsyncQuery", ro.EnableAsyncQuery.ToStringSafe());
            xmlw.WriteAttributeString("EnablePushDownGroupBy", ro.EnablePushDownGroupBy.ToStringSafe());
            xmlw.WriteAttributeString("EnableSaveDataWithReport", ro.EnableSaveDataWithReport.ToStringSafe());
            xmlw.WriteAttributeString("EnableSelectDistinctRecords", ro.EnableSelectDistinctRecords.ToStringSafe());
            xmlw.WriteAttributeString("EnableUseIndexForSpeed", ro.EnableUseIndexForSpeed.ToStringSafe());
            xmlw.WriteAttributeString("ErrorOnMaxNumOfRecords", ro.ErrorOnMaxNumOfRecords.ToStringSafe());
            xmlw.WriteAttributeString("InitialDataContext", ro.InitialDataContext);
            xmlw.WriteAttributeString("InitialReportPartName", ro.InitialReportPartName);
            xmlw.WriteAttributeString("MaxNumOfRecords", ro.MaxNumOfRecords.ToStringSafe());
            xmlw.WriteAttributeString("RefreshCEProperties", ro.RefreshCEProperties.ToStringSafe());
            xmlw.WriteEndElement();
        }

        public static void ProcessPrintOptions(PrintOptions po, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("PrintOptions");
            xmlw.WriteAttributeString("DissociatePageSizeAndPrinterPaperSize", po.DissociatePageSizeAndPrinterPaperSize.ToStringSafe());
            xmlw.WriteAttributeString("DriverName", po.DriverName);
            xmlw.WriteAttributeString("PageContentHeight", po.PageContentHeight.ToStringSafe());
            xmlw.WriteAttributeString("PageContentWidth", po.PageContentWidth.ToStringSafe());
            xmlw.WriteAttributeString("PaperOrientation", po.PaperOrientation.ToStringSafe());
            xmlw.WriteAttributeString("PaperSize", po.PaperSize.ToStringSafe());
            xmlw.WriteAttributeString("PaperSource", po.PaperSource.ToStringSafe());
            xmlw.WriteAttributeString("PortName", po.PortName);
            xmlw.WriteAttributeString("PrinterDuplex", po.PrinterDuplex.ToStringSafe());
            xmlw.WriteAttributeString("PrinterName", po.PrinterName);
            xmlw.WriteAttributeString("SavedDriverName", po.SavedDriverName);
            xmlw.WriteAttributeString("SavedPaperName", po.SavedPaperName);
            xmlw.WriteAttributeString("SavedPortName", po.SavedPortName);
            xmlw.WriteAttributeString("SavedPrinterName", po.SavedPrinterName);
            ProcessPageMargins(po.PageMargins, xmlw);
            xmlw.WriteEndElement();
        }
        private static void ProcessPageMargins(PageMargins pm, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("PageMargins");
            xmlw.WriteAttributeString("Bottom", pm.Bottom.ToStringSafe());
            xmlw.WriteAttributeString("Left", pm.Left.ToStringSafe());
            xmlw.WriteAttributeString("Right", pm.Right.ToStringSafe());
            xmlw.WriteAttributeString("Top", pm.Top.ToStringSafe());
            xmlw.WriteAttributeString("BottomFormula", pm.PageMarginConditionFormulas[CrPageMarginConditionFormulaTypeEnum.crPageMarginConditionFormulaTypeBottom].Text);
            xmlw.WriteAttributeString("LeftFormula", pm.PageMarginConditionFormulas[CrPageMarginConditionFormulaTypeEnum.crPageMarginConditionFormulaTypeLeft].Text);
            xmlw.WriteAttributeString("RightFormula", pm.PageMarginConditionFormulas[CrPageMarginConditionFormulaTypeEnum.crPageMarginConditionFormulaTypeRight].Text);
            xmlw.WriteAttributeString("TopFormula", pm.PageMarginConditionFormulas[CrPageMarginConditionFormulaTypeEnum.crPageMarginConditionFormulaTypeTop].Text);
            xmlw.WriteEndElement();
        }
        public static void ProcessSavedXMLExportFormats(XMLExportFormats xefs, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("XmlExportFormats");
            xmlw.WriteAttributeString("Count", xefs.Count.ToStringSafe());
            xmlw.WriteAttributeString("DefaultExportSelection", xefs.DefaultExportSelection.ToStringSafe());
            foreach (XMLExportFormat xef in xefs)
            {
                ProcessSavedXMLExportFormat(xef, xmlw);
            }
            xmlw.WriteEndElement();
        }
        private static void ProcessSavedXMLExportFormat(XMLExportFormat xef, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("XmlExportFormat");
            xmlw.WriteAttributeString("Description", xef.Description);
            xmlw.WriteAttributeString("ExportBlobField", xef.ExportBlobField.ToStringSafe());
            xmlw.WriteAttributeString("FileExtension", xef.FileExtension);
            xmlw.WriteAttributeString("Name", xef.Name);
            xmlw.WriteEndElement();
        }

        public static void ProcessSubreportLinks(SubreportLinks sls, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("SubreportLinks");
            xmlw.WriteAttributeString("Count", sls.Count.ToStringSafe());
            foreach (SubreportLink sl in sls)
            {
                ProcessSubreportLink(sl, xmlw);
            }
            xmlw.WriteEndElement();
        }

        private static void ProcessSubreportLink(SubreportLink sl, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("SubreportLink");
            xmlw.WriteAttributeString("LinkedParameterName", sl.LinkedParameterName);
            xmlw.WriteAttributeString("MainReportFieldName", sl.MainReportFieldName);
            xmlw.WriteAttributeString("SubreportFieldName", sl.SubreportFieldName);
            xmlw.WriteEndElement();
        }

        public static void ProcessReportDefinition(ReportDefinition rd, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("ReportDefinition");
            xmlw.WriteAttributeString("ReportKind", rd.ReportKind.ToStringSafe());
            xmlw.WriteAttributeString("RptrReportFlag", rd.RptrReportFlag.ToStringSafe());
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
            }
            xmlw.WriteEndElement();
        }

        private static void ProcessArea(Area area, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("Area");
            xmlw.WriteAttributeString("Kind", area.Kind.ToStringSafe());
            xmlw.WriteAttributeString("Name", area.Name);
            ProcessAreaFormat(area.Format, xmlw);
            ProcessSections(area.Sections, xmlw);
            xmlw.WriteEndElement();
        }

        private static void ProcessAreaFormat(ISCRAreaFormat af, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("AreaFormat");
            xmlw.WriteAttributeString("EnableClampPageFooter", af.EnableClampPageFooter.ToStringSafe());
            xmlw.WriteAttributeString("EnableHideForDrillDown", af.EnableHideForDrillDown.ToStringSafe());
            xmlw.WriteAttributeString("EnableKeepTogether", af.EnableKeepTogether.ToStringSafe());
            xmlw.WriteAttributeString("EnableNewPageAfter", af.EnableNewPageAfter.ToStringSafe());
            xmlw.WriteAttributeString("EnableNewPageBefore", af.EnableNewPageBefore.ToStringSafe());
            xmlw.WriteAttributeString("EnablePrintAtBottomOfPage", af.EnablePrintAtBottomOfPage.ToStringSafe());
            xmlw.WriteAttributeString("EnableResetPageNumberAfter", af.EnableResetPageNumberAfter.ToStringSafe());
            xmlw.WriteAttributeString("EnableSuppress", af.EnableSuppress.ToStringSafe());
            xmlw.WriteAttributeString("VisibleRecordNumberPerPage", af.VisibleRecordNumberPerPage.ToStringSafe());

            xmlw.WriteAttributeString("BackgroundColor", af.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeBackgroundColor].Text);
            xmlw.WriteAttributeString("EnableClampPageFooterFormula", af.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableClampPageFooter].Text);
            xmlw.WriteAttributeString("EnableHideForDrillDownFormula", af.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableHideForDrillDown].Text);
            xmlw.WriteAttributeString("EnableKeepTogetherFormula", af.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableKeepTogether].Text);
            xmlw.WriteAttributeString("EnableNewPageAfterFormula", af.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableNewPageAfter].Text);
            xmlw.WriteAttributeString("EnableNewPageBeforeFormula", af.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableNewPageBefore].Text);
            xmlw.WriteAttributeString("EnablePrintAtBottomOfPageFormula", af.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnablePrintAtBottomOfPage].Text);
            xmlw.WriteAttributeString("EnableResetPageNumberAfterFormula", af.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableResetPageNumberAfter].Text);
            xmlw.WriteAttributeString("EnableSuppressFormula", af.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableSuppress].Text);
            xmlw.WriteAttributeString("EnableSuppressIfBlankFormula", af.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableSuppressIfBlank].Text);
            xmlw.WriteAttributeString("EnableUnderlaySectionFormula", af.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableUnderlaySection].Text);
            xmlw.WriteAttributeString("GroupNumberPerPageFormula", af.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeGroupNumberPerPage].Text);
            xmlw.WriteAttributeString("RecordNumberPerPageFormula", af.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeRecordNumberPerPage].Text);
            xmlw.WriteEndElement();
        }

        private static void ProcessSections(Sections sections, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("Sections");
            xmlw.WriteAttributeString("Count", sections.Count.ToStringSafe());
            foreach (Section section in sections)
            {
                ProcessSection(section, xmlw);
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
            ProcessSectionFormat(section.Format, xmlw);
            xmlw.WriteEndElement();
        }
        private static void ProcessReportObjects(ReportObjects ros, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("ReportObjects");
            xmlw.WriteAttributeString("Count", ros.Count.ToStringSafe());
            foreach (ReportObject ro in ros)
            {
                ProcessReportObject(ro, xmlw);
            }
            xmlw.WriteEndElement();
        }
        private static void ProcessReportObject(ReportObject ro, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("ReportObject");
            // for all types the same properties

            xmlw.WriteAttributeString("Height", ro.Height.ToStringSafe());
            xmlw.WriteAttributeString("Kind", ro.Kind.ToStringSafe());
            xmlw.WriteAttributeString("Left", ro.Left.ToStringSafe());
            xmlw.WriteAttributeString("LinkedURI", ro.LinkedURI);
            xmlw.WriteAttributeString("Name", ro.Name);
            //xmlw.WriteAttributeString("ReportPartIDDataContext", ro.ReportPartBookmark.ReportPartID.DataContext);
            //xmlw.WriteAttributeString("ReportPartIDName", ro.ReportPartBookmark.ReportPartID.Name);
            //xmlw.WriteAttributeString("ReportURI", ro.ReportPartBookmark.ReportURI);
            xmlw.WriteAttributeString("Top", ro.Top.ToStringSafe());
            xmlw.WriteAttributeString("Width", ro.Width.ToStringSafe());
            // kind specific properties
            switch (ro.Kind)
            {
                case CrReportObjectKindEnum.crReportObjectKindBlobField:
                    {
                        BlobFieldObject bfo = (BlobFieldObject)ro;
                        xmlw.WriteAttributeString("DataSourceName", bfo.DataSourceName);
                        xmlw.WriteAttributeString("OriginalHeight", bfo.OriginalHeight.ToStringSafe());
                        xmlw.WriteAttributeString("OriginalWidth", bfo.OriginalWidth.ToStringSafe());
                        xmlw.WriteAttributeString("XScaling", bfo.XScaling.ToStringSafe());
                        xmlw.WriteAttributeString("YScaling", bfo.YScaling.ToStringSafe());
                        ProcessPictureFormat(bfo.PictureFormat,xmlw);
                        break;
                    }
                case CrReportObjectKindEnum.crReportObjectKindBox: 
                    {
                        BoxObject bo = (BoxObject)ro;
                        xmlw.WriteAttributeString("Bottom", bo.Bottom.ToStringSafe());
                        xmlw.WriteAttributeString("CornerEllipseHeight", bo.CornerEllipseHeight.ToStringSafe());
                        xmlw.WriteAttributeString("CornerEllipseWidth", bo.CornerEllipseWidth.ToStringSafe());
                        xmlw.WriteAttributeString("EnableExtendToBottomOfSection", bo.EnableExtendToBottomOfSection.ToStringSafe());
                        xmlw.WriteAttributeString("EndSectionName", bo.EndSectionName);
                        xmlw.WriteAttributeString("FillColor", bo.FillColor.ToStringSafe());
                        xmlw.WriteAttributeString("LineColor", bo.LineColor.ToStringSafe());
                        xmlw.WriteAttributeString("LineStyle", bo.LineStyle.ToStringSafe());
                        xmlw.WriteAttributeString("LineThickness", bo.LineThickness.ToStringSafe());
                        xmlw.WriteAttributeString("Right", bo.Right.ToStringSafe());
                        break;
                    }
                case CrReportObjectKindEnum.crReportObjectKindFieldHeading:
                    {
                        FieldHeadingObject fho = (FieldHeadingObject)ro;
                        xmlw.WriteAttributeString("FieldObjectName", fho.FieldObjectName);
                        xmlw.WriteAttributeString("MaxNumberOfLines", fho.MaxNumberOfLines.ToStringSafe());
                        xmlw.WriteAttributeString("ReadingOrder", fho.ReadingOrder.ToStringSafe());
                        xmlw.WriteAttributeString("Text", fho.Text);
                        ProcessFontColor(fho.FontColor, xmlw);
                        ProcessParagraphs(fho.Paragraphs, xmlw);
                        break;
                    }
                case CrReportObjectKindEnum.crReportObjectKindField:
                    {
                        FieldObject fo = (FieldObject)ro;
                        xmlw.WriteAttributeString("DataSourceName", fo.DataSourceName);
                        xmlw.WriteAttributeString("FieldValueType", fo.FieldValueType.ToStringSafe());
                        ProcessFieldFormat(fo.FieldFormat, xmlw, fo.FieldValueType);
                        ProcessFontColor(fo.FontColor, xmlw);
                        break;
                    }
                case CrReportObjectKindEnum.crReportObjectKindLine:
                    {
                        LineObject lo = (LineObject)ro;
                        xmlw.WriteAttributeString("Bottom", lo.Bottom.ToStringSafe());
                        xmlw.WriteAttributeString("EnableExtendToBottomOfSection", lo.EnableExtendToBottomOfSection.ToStringSafe());
                        xmlw.WriteAttributeString("EndSectionName", lo.EndSectionName);
                        xmlw.WriteAttributeString("LineColor", lo.LineColor.ToStringSafe());
                        xmlw.WriteAttributeString("LineStyle", lo.LineStyle.ToStringSafe());
                        xmlw.WriteAttributeString("LineThickness", lo.LineThickness.ToStringSafe());
                        xmlw.WriteAttributeString("Right", lo.Right.ToStringSafe());
                        break;
                    }
                case CrReportObjectKindEnum.crReportObjectKindSubreport:
                    {
                        SubreportObject so = (SubreportObject)ro;
                        xmlw.WriteAttributeString("EnableOnDemand", so.EnableOnDemand.ToStringSafe());
                        xmlw.WriteAttributeString("IsImported", so.IsImported.ToStringSafe());
                        xmlw.WriteAttributeString("SubreportLocation", so.SubreportLocation);
                        xmlw.WriteAttributeString("SubreportName", so.SubreportName);
                        break;
                    }
                case CrReportObjectKindEnum.crReportObjectKindText:
                    {
                        TextObject to = (TextObject)ro;
                        xmlw.WriteAttributeString("MaxNumberOfLines", to.MaxNumberOfLines.ToStringSafe());
                        xmlw.WriteAttributeString("ReadingOrder", to.ReadingOrder.ToStringSafe());
                        xmlw.WriteAttributeString("Text", to.Text);
                        ProcessFontColor(to.FontColor, xmlw);
                        ProcessParagraphs(to.Paragraphs, xmlw);
                        break;
                    }
                default:
                    {
                        break;
                    }
            }
            ProcessBorder(ro.Border, xmlw);
            ProcessObjectFormat(ro.Format, xmlw);
            xmlw.WriteEndElement();
        }

        private static void ProcessParagraphs(Paragraphs paragraphs, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("Paragraphs");
            xmlw.WriteAttributeString("Count", paragraphs.Count.ToStringSafe());
            foreach (Paragraph p in paragraphs)
            {
                ProcessParagraph(p, xmlw);
            }
            xmlw.WriteEndElement();
        }

        private static void ProcessParagraph(Paragraph p, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("Paragraph");
            xmlw.WriteAttributeString("Alignment", p.Alignment.ToStringSafe());
            xmlw.WriteAttributeString("ReadingOrder", p.ReadingOrder.ToStringSafe());
            ProcessFontColor(p.FontColor, xmlw);
            ProcessIndentAndSpacingFormat(p.IndentAndSpacingFormat, xmlw);
            ProcessParagraphElements(p.ParagraphElements, xmlw);
            ProcessTabStops(p.TabStops, xmlw);
            xmlw.WriteEndElement();
        }

        private static void ProcessParagraphElements(ParagraphElements pe, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("ParagraphElements");
            xmlw.WriteAttributeString("Count", pe.Count.ToStringSafe());
            foreach (ParagraphElement p in pe)
            {
                ProcessParagraphElement(p, xmlw);
            }
            xmlw.WriteEndElement();
        }

        private static void ProcessParagraphElement(ParagraphElement p, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("ParagraphElement");
            xmlw.WriteAttributeString("Kind", p.Kind.ToStringSafe());
            ProcessFontColor(p.FontColor, xmlw);
            switch (p.Kind)
            {
                case CrParagraphElementKindEnum.crParagraphElementKindText:
                    if (p is ParagraphTextElement pText)
                    {
                        xmlw.WriteElementString("Text", pText.Text);
                    }
                    break;
                case CrParagraphElementKindEnum.crParagraphElementKindField:
                    if (p is ParagraphFieldElement pField)
                    {
                        if (pField is FieldObject pFieldObject)
                        {
                            ProcessFieldFormat(pField.FieldFormat, xmlw, pFieldObject.FieldValueType);
                        }
                        xmlw.WriteElementString("FieldDataSource", pField.DataSource);
                    }
                    break;
                case CrParagraphElementKindEnum.crParagraphElementKindTab:
                default:
                    Console.WriteLine($"Unhandled ParagraphElement Kind {p.Kind} {p.ClassName}");
                    break;
            }
            xmlw.WriteEndElement();
        }

        private static void ProcessFontColor(FontColor fc, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("FontColor");
            xmlw.WriteAttributeString("Color", fc.Color.ToStringSafe());
            xmlw.WriteAttributeString("Bold", fc.Font.Bold.ToStringSafe());
            xmlw.WriteAttributeString("Charset", fc.Font.Charset.ToStringSafe());
            xmlw.WriteAttributeString("Italic", fc.Font.Italic.ToStringSafe());
            xmlw.WriteAttributeString("Name", fc.Font.Name);
            xmlw.WriteAttributeString("Size", fc.Font.Size.ToStringSafe());
            xmlw.WriteAttributeString("Strikethrough", fc.Font.Strikethrough.ToStringSafe());
            xmlw.WriteAttributeString("Underline", fc.Font.Underline.ToStringSafe());
            xmlw.WriteAttributeString("Weight", fc.Font.Weight.ToStringSafe());
            xmlw.WriteAttributeString("ColorFormula", fc.ConditionFormulas[CrFontColorConditionFormulaTypeEnum.crFontColorConditionFormulaTypeColor].Text);
            xmlw.WriteAttributeString("NameFormula", fc.ConditionFormulas[CrFontColorConditionFormulaTypeEnum.crFontColorConditionFormulaTypeName].Text);
            xmlw.WriteAttributeString("SizeFormula", fc.ConditionFormulas[CrFontColorConditionFormulaTypeEnum.crFontColorConditionFormulaTypeSize].Text);
            xmlw.WriteAttributeString("StrikeoutFormula", fc.ConditionFormulas[CrFontColorConditionFormulaTypeEnum.crFontColorConditionFormulaTypeStrikeout].Text);
            xmlw.WriteAttributeString("StyleFormula", fc.ConditionFormulas[CrFontColorConditionFormulaTypeEnum.crFontColorConditionFormulaTypeStyle].Text);
            xmlw.WriteAttributeString("UnderlineFormula", fc.ConditionFormulas[CrFontColorConditionFormulaTypeEnum.crFontColorConditionFormulaTypeUnderline].Text);
            xmlw.WriteEndElement();
        }

        private static void ProcessTabStops(TabStops ts, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("TabStops");
            xmlw.WriteAttributeString("Count", ts.Count.ToStringSafe());
            foreach (TabStop t in ts)
            {
                ProcessTabStop(t, xmlw);
            }
            xmlw.WriteEndElement();
        }

        private static void ProcessTabStop(TabStop t, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("TabStop");
            xmlw.WriteAttributeString("Alignment", t.Alignment.ToStringSafe());
            xmlw.WriteAttributeString("XOffset", t.XOffset.ToStringSafe());
            xmlw.WriteEndElement();
        }

        private static void ProcessPictureFormat(PictureFormat pf, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("PictureFormat");
            xmlw.WriteAttributeString("BottomCropping", pf.BottomCropping.ToStringSafe());
            xmlw.WriteAttributeString("LeftCropping", pf.LeftCropping.ToStringSafe());
            xmlw.WriteAttributeString("RightCropping", pf.RightCropping.ToStringSafe());
            xmlw.WriteAttributeString("TopCropping", pf.TopCropping.ToStringSafe());
            xmlw.WriteEndElement();
        }

        private static void ProcessFieldFormat(FieldFormat ff, XmlWriter xmlw, CrFieldValueTypeEnum vt)
        {
            switch (vt)
            {
                case CrFieldValueTypeEnum.crFieldValueTypeBooleanField:
                    {
                        ProcessBooleanFormat(ff.BooleanFormat, xmlw);
                        break;
                    }
                case CrFieldValueTypeEnum.crFieldValueTypeDateField:
                    {
                        ProcessDateFormat(ff.DateFormat, xmlw); 
                        break;
                    }
                case CrFieldValueTypeEnum.crFieldValueTypeDateTimeField:
                    {
                        ProcessDateTimeFormat(ff.DateTimeFormat, xmlw);    
                        break;
                    }
                case CrFieldValueTypeEnum.crFieldValueTypeCurrencyField:
                case CrFieldValueTypeEnum.crFieldValueTypeInt16sField:
                case CrFieldValueTypeEnum.crFieldValueTypeInt16uField:
                case CrFieldValueTypeEnum.crFieldValueTypeInt32sField:
                case CrFieldValueTypeEnum.crFieldValueTypeInt32uField:
                case CrFieldValueTypeEnum.crFieldValueTypeInt8sField:
                case CrFieldValueTypeEnum.crFieldValueTypeInt8uField:
                case CrFieldValueTypeEnum.crFieldValueTypeNumberField:
                    {
                        ProcessNumericFormat(ff.NumericFormat, xmlw);
                        break;
                    }
                case CrFieldValueTypeEnum.crFieldValueTypeStringField:
                    {
                        ProcessStringFormat(ff.StringFormat, xmlw);
                        break;
                    }
                case CrFieldValueTypeEnum.crFieldValueTypeTimeField:
                    {
                        ProcessTimeFormat(ff.TimeFormat, xmlw);
                        break;
                    }
                default:
                    {
                        
                        break;
                    }
            }
            ProcessCommonFormat(ff.CommonFormat, xmlw);
        }

        private static void ProcessBooleanFormat(BooleanFieldFormat bff, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("BooleanFieldFormat");
            xmlw.WriteAttributeString("OutputFormat", bff.OutputFormat.ToStringSafe());
            xmlw.WriteAttributeString("OutputFormatFormula", bff.ConditionFormulas[CrBooleanFieldFormatConditionFormulaTypeEnum.crBooleanFieldFormatConditionFormulaTypeOutputFormat].Text);
            xmlw.WriteEndElement();
        }

        private static void ProcessCommonFormat(CommonFieldFormat cff, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("CommonFieldFormat");
            xmlw.WriteAttributeString("EnableSuppressIfDuplicated", cff.EnableSuppressIfDuplicated.ToStringSafe());
            xmlw.WriteAttributeString("EnableSystemDefault", cff.EnableSystemDefault.ToStringSafe());
            xmlw.WriteAttributeString("EnableSuppressIfDuplicatedFormula", cff.ConditionFormulas[CrCommonFieldFormatConditionFormulaTypeEnum.crCommonFieldFormatConditionFormulaTypeSuppressIfDuplicated].Text);
            xmlw.WriteAttributeString("EnableSystemDefaultFormula", cff.ConditionFormulas[CrCommonFieldFormatConditionFormulaTypeEnum.crCommonFieldFormatConditionFormulaTypeUseSystemDefault].Text);
            xmlw.WriteEndElement();
        }

        private static void ProcessDateTimeFormat(DateTimeFieldFormat dtff, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("DateTimeFieldFormat");
            xmlw.WriteAttributeString("DateTimeSeparator", dtff.DateTimeSeparator);
            xmlw.WriteAttributeString("DateTimeOrder", dtff.DateTimeOrder.ToStringSafe());
            xmlw.WriteAttributeString("DateTimeSeparatorFormula", dtff.ConditionFormulas[CrDateTimeFieldFormatConditionFormulaTypeEnum.crDateTimeFieldFormatConditionFormulaTypeDateTimeOrder].Text);
            xmlw.WriteAttributeString("DateTimeOrderFormula", dtff.ConditionFormulas[CrDateTimeFieldFormatConditionFormulaTypeEnum.crDateTimeFieldFormatConditionFormulaTypeDateTimeSeparator].Text);
            xmlw.WriteEndElement();
        }

        private static void ProcessDateFormat(DateFieldFormat dff, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("DateFieldFormat");
            xmlw.WriteAttributeString("CalendarType", dff.CalendarType.ToStringSafe());
            xmlw.WriteAttributeString("DateFirstSeparator", dff.DateFirstSeparator.ToStringSafe());
            xmlw.WriteAttributeString("DateOrder", dff.DateOrder.ToStringSafe());
            xmlw.WriteAttributeString("DatePrefixSeparator", dff.DatePrefixSeparator.ToStringSafe());
            xmlw.WriteAttributeString("DateSecondSeparator", dff.DateSecondSeparator.ToStringSafe());
            xmlw.WriteAttributeString("DateSuffixSeparator", dff.DateSuffixSeparator.ToStringSafe());
            xmlw.WriteAttributeString("DayFormat", dff.DayFormat.ToStringSafe());
            xmlw.WriteAttributeString("DayOfWeekPosition", dff.DayOfWeekPosition.ToStringSafe());
            xmlw.WriteAttributeString("DayOfWeekSeparator", dff.DayOfWeekSeparator.ToStringSafe());
            xmlw.WriteAttributeString("DayOfWeekType", dff.DayOfWeekType.ToStringSafe());            
            xmlw.WriteAttributeString("EraType", dff.EraType.ToStringSafe());
            xmlw.WriteAttributeString("MonthFormat", dff.MonthFormat.ToStringSafe());
            xmlw.WriteAttributeString("SystemDefaultType", dff.SystemDefaultType.ToStringSafe());
            xmlw.WriteAttributeString("YearFormat", dff.YearFormat.ToStringSafe());
            xmlw.WriteAttributeString("CalendarTypeFormula", dff.ConditionFormulas[CrDateFieldFormatConditionFormulaTypeEnum.crDateFieldFormatConditionFormulaTypeCalendarType].Text);
            xmlw.WriteAttributeString("DateFirstSeparatorFormula", dff.ConditionFormulas[CrDateFieldFormatConditionFormulaTypeEnum.crDateFieldFormatConditionFormulaTypeDateFirstSeparator].Text);
            xmlw.WriteAttributeString("DateOrderFormula", dff.ConditionFormulas[CrDateFieldFormatConditionFormulaTypeEnum.crDateFieldFormatConditionFormulaTypeDateOrder].Text);
            xmlw.WriteAttributeString("DatePrefixSeparatorFormula", dff.ConditionFormulas[CrDateFieldFormatConditionFormulaTypeEnum.crDateFieldFormatConditionFormulaTypeDatePrefixSeparator].Text);
            xmlw.WriteAttributeString("DateSecondSeparatorFormula", dff.ConditionFormulas[CrDateFieldFormatConditionFormulaTypeEnum.crDateFieldFormatConditionFormulaTypeDateSecondSeparator].Text);
            xmlw.WriteAttributeString("DateSuffixSeparatorFormula", dff.ConditionFormulas[CrDateFieldFormatConditionFormulaTypeEnum.crDateFieldFormatConditionFormulaTypeDateSuffixSeparator].Text);
            xmlw.WriteAttributeString("DayFormatFormula", dff.ConditionFormulas[CrDateFieldFormatConditionFormulaTypeEnum.crDateFieldFormatConditionFormulaTypeDayFormat].Text);
            xmlw.WriteAttributeString("DayOfWeekPositionFormula", dff.ConditionFormulas[CrDateFieldFormatConditionFormulaTypeEnum.crDateFieldFormatConditionFormulaTypeDayOfWeekPosition].Text);
            xmlw.WriteAttributeString("DayOfWeekSeparatorFormula", dff.ConditionFormulas[CrDateFieldFormatConditionFormulaTypeEnum.crDateFieldFormatConditionFormulaTypeDayOfWeekSeparator].Text);
            xmlw.WriteAttributeString("DayOfWeekTypeFormula", dff.ConditionFormulas[CrDateFieldFormatConditionFormulaTypeEnum.crDateFieldFormatConditionFormulaTypeDayOfWeekType].Text);
            xmlw.WriteAttributeString("EraTypeFormula", dff.ConditionFormulas[CrDateFieldFormatConditionFormulaTypeEnum.crDateFieldFormatConditionFormulaTypeEraType].Text);
            xmlw.WriteAttributeString("MonthFormatFormula", dff.ConditionFormulas[CrDateFieldFormatConditionFormulaTypeEnum.crDateFieldFormatConditionFormulaTypeMonthFormat].Text);
            xmlw.WriteAttributeString("SystemDefaultTypeFormula", dff.ConditionFormulas[CrDateFieldFormatConditionFormulaTypeEnum.crDateFieldFormatConditionFormulaTypeSystemDefaultType].Text);
            xmlw.WriteAttributeString("YearFormatFormula", dff.ConditionFormulas[CrDateFieldFormatConditionFormulaTypeEnum.crDateFieldFormatConditionFormulaTypeYearFormat].Text);          
            xmlw.WriteEndElement();
        }
        
        private static void ProcessNumericFormat(NumericFieldFormat nff, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("NumericFieldFormat");
            xmlw.WriteAttributeString("CurrencyPosition", nff.CurrencyPosition.ToStringSafe());
            xmlw.WriteAttributeString("CurrencySymbol", nff.CurrencySymbol);
            xmlw.WriteAttributeString("CurrencySymbolFormat", nff.CurrencySymbolFormat.ToStringSafe());
            xmlw.WriteAttributeString("DecimalSymbol", nff.DecimalSymbol);
            xmlw.WriteAttributeString("DisplayReverseSign", nff.DisplayReverseSign.ToStringSafe());
            xmlw.WriteAttributeString("EnableSuppressIfZero", nff.EnableSuppressIfZero.ToStringSafe());
            xmlw.WriteAttributeString("EnableUseLeadZero", nff.EnableUseLeadZero.ToStringSafe());
            xmlw.WriteAttributeString("NDecimalPlaces", nff.NDecimalPlaces.ToStringSafe());
            xmlw.WriteAttributeString("NegativeFormat", nff.NegativeFormat.ToStringSafe());
            xmlw.WriteAttributeString("OneCurrencySymbolPerPage", nff.OneCurrencySymbolPerPage.ToStringSafe());
            xmlw.WriteAttributeString("RoundingFormat", nff.RoundingFormat.ToStringSafe());
            xmlw.WriteAttributeString("ThousandsSeparator", nff.ThousandsSeparator.ToStringSafe());
            xmlw.WriteAttributeString("ThousandSymbol", nff.ThousandSymbol);
            xmlw.WriteAttributeString("ZeroValueString", nff.ZeroValueString);
            xmlw.WriteAttributeString("CurrencyPositionFormula", nff.ConditionFormulas[CrNumericFieldFormatConditionFormulaTypeEnum.crNumericFieldFormatConditionFormulaTypeCurrencyPosition].Text);
            xmlw.WriteAttributeString("CurrencySymbolFormula", nff.ConditionFormulas[CrNumericFieldFormatConditionFormulaTypeEnum.crNumericFieldFormatConditionFormulaTypeCurrencySymbol].Text);
            xmlw.WriteAttributeString("CurrencySymbolFormatFormula", nff.ConditionFormulas[CrNumericFieldFormatConditionFormulaTypeEnum.crNumericFieldFormatConditionFormulaTypeCurrencySymbolFormat].Text);
            xmlw.WriteAttributeString("DecimalSymbolFormula", nff.ConditionFormulas[CrNumericFieldFormatConditionFormulaTypeEnum.crNumericFieldFormatConditionFormulaTypeDecimalSymbol].Text);
            xmlw.WriteAttributeString("DisplayReverseSignFormula", nff.ConditionFormulas[CrNumericFieldFormatConditionFormulaTypeEnum.crNumericFieldFormatConditionFormulaTypeDisplayReverseSign].Text);
            xmlw.WriteAttributeString("EnableSuppressIfZeroFormula", nff.ConditionFormulas[CrNumericFieldFormatConditionFormulaTypeEnum.crNumericFieldFormatConditionFormulaTypeEnableSuppressIfZero].Text);
            xmlw.WriteAttributeString("EnableUseLeadZeroFormula", nff.ConditionFormulas[CrNumericFieldFormatConditionFormulaTypeEnum.crNumericFieldFormatConditionFormulaTypeEnableUseLeadZero].Text);
            xmlw.WriteAttributeString("NDecimalPlacesFormula", nff.ConditionFormulas[CrNumericFieldFormatConditionFormulaTypeEnum.crNumericFieldFormatConditionFormulaTypeNDecimalPlaces].Text);
            xmlw.WriteAttributeString("NegativeFormatFormula", nff.ConditionFormulas[CrNumericFieldFormatConditionFormulaTypeEnum.crNumericFieldFormatConditionFormulaTypeNegativeFormat].Text);
            xmlw.WriteAttributeString("OneCurrencySymbolPerPageFormula", nff.ConditionFormulas[CrNumericFieldFormatConditionFormulaTypeEnum.crNumericFieldFormatConditionFormulaTypeOneCurrencySymbolPerPage].Text);
            xmlw.WriteAttributeString("RoundingFormatFormula", nff.ConditionFormulas[CrNumericFieldFormatConditionFormulaTypeEnum.crNumericFieldFormatConditionFormulaTypeRoundingFormat].Text);
            xmlw.WriteAttributeString("ThousandsSeparatorFormula", nff.ConditionFormulas[CrNumericFieldFormatConditionFormulaTypeEnum.crNumericFieldFormatConditionFormulaTypeThousandsSeparator].Text);
            xmlw.WriteAttributeString("ThousandSymbolFormula", nff.ConditionFormulas[CrNumericFieldFormatConditionFormulaTypeEnum.crNumericFieldFormatConditionFormulaTypeThousandSymbol].Text);
            xmlw.WriteAttributeString("ZeroValueStringFormula", nff.ConditionFormulas[CrNumericFieldFormatConditionFormulaTypeEnum.crNumericFieldFormatConditionFormulaTypeZeroValueString].Text);           
            xmlw.WriteEndElement();
        }

        private static void ProcessStringFormat(StringFieldFormat sff, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("StringFieldFormat");
            xmlw.WriteAttributeString("EnableWordWrap", sff.EnableWordWrap.ToStringSafe());
            xmlw.WriteAttributeString("CharacterSpacing", sff.CharacterSpacing.ToStringSafe());
            xmlw.WriteAttributeString("MaxNumberOfLines", sff.MaxNumberOfLines.ToStringSafe());
            xmlw.WriteAttributeString("ReadingOrder", sff.ReadingOrder.ToStringSafe());
            xmlw.WriteAttributeString("TextFormat", sff.TextFormat.ToStringSafe());
            ProcessIndentAndSpacingFormat(sff.IndentAndSpacingFormat,xmlw);
            xmlw.WriteEndElement();
        }

        private static void ProcessIndentAndSpacingFormat(IndentAndSpacingFormat iasf, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("IndentAndSpacingFormat");
            xmlw.WriteAttributeString("FirstLineIndent", iasf.FirstLineIndent.ToStringSafe());
            xmlw.WriteAttributeString("LeftIndent", iasf.LeftIndent.ToStringSafe());
            xmlw.WriteAttributeString("LineSpacing", iasf.LineSpacing.ToStringSafe());
            xmlw.WriteAttributeString("LineSpacingType", iasf.LineSpacingType.ToStringSafe());
            xmlw.WriteAttributeString("RightIndent", iasf.RightIndent.ToStringSafe());
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
            xmlw.WriteAttributeString("AMPMFormatFormula", tff.ConditionFormulas[CrTimeFieldFormatConditionFormulaTypeEnum.crTimeFieldFormatConditionFormulaTypeAMPMFormat].Text);
            xmlw.WriteAttributeString("AMStringFormula", tff.ConditionFormulas[CrTimeFieldFormatConditionFormulaTypeEnum.crTimeFieldFormatConditionFormulaTypeAMString].Text);
            xmlw.WriteAttributeString("HourFormatFormula", tff.ConditionFormulas[CrTimeFieldFormatConditionFormulaTypeEnum.crTimeFieldFormatConditionFormulaTypeHourFormat].Text);
            xmlw.WriteAttributeString("HourMinuteSeparatorFormula", tff.ConditionFormulas[CrTimeFieldFormatConditionFormulaTypeEnum.crTimeFieldFormatConditionFormulaTypeHourMinuteSeparator].Text);
            xmlw.WriteAttributeString("MinuteFormatFormula", tff.ConditionFormulas[CrTimeFieldFormatConditionFormulaTypeEnum.crTimeFieldFormatConditionFormulaTypeMinuteFormat].Text);
            xmlw.WriteAttributeString("MinuteSecondSeparatorFormula", tff.ConditionFormulas[CrTimeFieldFormatConditionFormulaTypeEnum.crTimeFieldFormatConditionFormulaTypeMinuteSecondSeparator].Text);
            xmlw.WriteAttributeString("PMStringFormula", tff.ConditionFormulas[CrTimeFieldFormatConditionFormulaTypeEnum.crTimeFieldFormatConditionFormulaTypePMString].Text);
            xmlw.WriteAttributeString("SecondFormatFormula", tff.ConditionFormulas[CrTimeFieldFormatConditionFormulaTypeEnum.crTimeFieldFormatConditionFormulaTypeSecondFormat].Text);
            xmlw.WriteAttributeString("TimeBaseFormula", tff.ConditionFormulas[CrTimeFieldFormatConditionFormulaTypeEnum.crTimeFieldFormatConditionFormulaTypeTimeBase].Text);
            xmlw.WriteEndElement();
        }

        private static void ProcessBorder(Border border, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("Border");
            xmlw.WriteAttributeString("BackgroundColor", border.BackgroundColor.ToStringSafe());
            xmlw.WriteAttributeString("BorderColor", border.BorderColor.ToStringSafe());
            xmlw.WriteAttributeString("BottomLineStyle", border.BottomLineStyle.ToStringSafe());
            xmlw.WriteAttributeString("EnableTightHorizontal", border.EnableTightHorizontal.ToStringSafe());
            xmlw.WriteAttributeString("HasDropShadow", border.HasDropShadow.ToStringSafe());
            xmlw.WriteAttributeString("LeftLineStyle", border.LeftLineStyle.ToStringSafe());
            xmlw.WriteAttributeString("RightLineStyle", border.RightLineStyle.ToStringSafe());
            xmlw.WriteAttributeString("TopLineStyle", border.TopLineStyle.ToStringSafe());
            xmlw.WriteAttributeString("BackgroundColorFormula", border.ConditionFormulas[CrBorderConditionFormulaTypeEnum.crBorderConditionFormulaTypeBackgroundColor].Text);
            xmlw.WriteAttributeString("BorderColorFormula", border.ConditionFormulas[CrBorderConditionFormulaTypeEnum.crBorderConditionFormulaTypeBorderColor].Text);
            xmlw.WriteAttributeString("BottomLineStyleFormula", border.ConditionFormulas[CrBorderConditionFormulaTypeEnum.crBorderConditionFormulaTypeBottomLineStyle].Text);
            xmlw.WriteAttributeString("HasDropShadowFormula", border.ConditionFormulas[CrBorderConditionFormulaTypeEnum.crBorderConditionFormulaTypeHasDropShadow].Text);
            xmlw.WriteAttributeString("LeftLineStyleFormula", border.ConditionFormulas[CrBorderConditionFormulaTypeEnum.crBorderConditionFormulaTypeLeftLineStyle].Text);
            xmlw.WriteAttributeString("RightLineStyleFormula", border.ConditionFormulas[CrBorderConditionFormulaTypeEnum.crBorderConditionFormulaTypeRightLineStyle].Text);
            xmlw.WriteAttributeString("TightHorizontalFormula", border.ConditionFormulas[CrBorderConditionFormulaTypeEnum.crBorderConditionFormulaTypeTightHorizontal].Text);
            xmlw.WriteAttributeString("TightVerticalFormula", border.ConditionFormulas[CrBorderConditionFormulaTypeEnum.crBorderConditionFormulaTypeTightVertical].Text);
            xmlw.WriteAttributeString("TopLineStyleFormula", border.ConditionFormulas[CrBorderConditionFormulaTypeEnum.crBorderConditionFormulaTypeTopLineStyle].Text);
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
            xmlw.WriteAttributeString("HyperlinkText", of.HyperlinkText);
            xmlw.WriteAttributeString("HyperlinkType", of.HyperlinkType.ToStringSafe());
            xmlw.WriteAttributeString("TextRotationAngle", of.TextRotationAngle.ToStringSafe());
            xmlw.WriteAttributeString("ToolTipText", of.ToolTipText);
            xmlw.WriteAttributeString("VerticalAlignment", of.VerticalAlignment.ToStringSafe());
            xmlw.WriteAttributeString("CssClassFormula", of.ConditionFormulas[CrObjectFormatConditionFormulaTypeEnum.crObjectFormatConditionFormulaTypeCssClass].Text);
            xmlw.WriteAttributeString("DeltaWidthFormula", of.ConditionFormulas[CrObjectFormatConditionFormulaTypeEnum.crObjectFormatConditionFormulaTypeDeltaWidth].Text);
            xmlw.WriteAttributeString("DeltaXFormula", of.ConditionFormulas[CrObjectFormatConditionFormulaTypeEnum.crObjectFormatConditionFormulaTypeDeltaX].Text);
            xmlw.WriteAttributeString("DisplayStringFormula", of.ConditionFormulas[CrObjectFormatConditionFormulaTypeEnum.crObjectFormatConditionFormulaTypeDisplayString].Text);
            xmlw.WriteAttributeString("EnableCanGrowFormula", of.ConditionFormulas[CrObjectFormatConditionFormulaTypeEnum.crObjectFormatConditionFormulaTypeEnableCanGrow].Text);
            xmlw.WriteAttributeString("EnableCloseAtPageBreakFormula", of.ConditionFormulas[CrObjectFormatConditionFormulaTypeEnum.crObjectFormatConditionFormulaTypeEnableCloseAtPageBreak].Text);
            xmlw.WriteAttributeString("EnableKeepTogetherFormula", of.ConditionFormulas[CrObjectFormatConditionFormulaTypeEnum.crObjectFormatConditionFormulaTypeEnableKeepTogether].Text);
            xmlw.WriteAttributeString("EnableSuppressFormula", of.ConditionFormulas[CrObjectFormatConditionFormulaTypeEnum.crObjectFormatConditionFormulaTypeEnableSuppress].Text);
            xmlw.WriteAttributeString("HorizontalAlignmentFormula", of.ConditionFormulas[CrObjectFormatConditionFormulaTypeEnum.crObjectFormatConditionFormulaTypeHorizontalAlignment].Text);
            xmlw.WriteAttributeString("HyperlinkFormula", of.ConditionFormulas[CrObjectFormatConditionFormulaTypeEnum.crObjectFormatConditionFormulaTypeHyperlink].Text);
            xmlw.WriteAttributeString("TextRotationAngleFormula", of.ConditionFormulas[CrObjectFormatConditionFormulaTypeEnum.crObjectFormatConditionFormulaTypeRotation].Text);
            xmlw.WriteAttributeString("ToolTipTextFormula", of.ConditionFormulas[CrObjectFormatConditionFormulaTypeEnum.crObjectFormatConditionFormulaTypeToolTipText].Text);
            xmlw.WriteAttributeString("VerticalAlignmentFormula", of.ConditionFormulas[CrObjectFormatConditionFormulaTypeEnum.crObjectFormatConditionFormulaTypeVerticalAlignment].Text);
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
            xmlw.WriteAttributeString("PageOrientation", sf.PageOrientation.ToStringSafe());
            xmlw.WriteAttributeString("BackgroundColorFormula", sf.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeBackgroundColor].Text);
            xmlw.WriteAttributeString("EnableClampPageFooterFormula", sf.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableClampPageFooter].Text);
            xmlw.WriteAttributeString("EnableHideForDrillDownFormula", sf.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableHideForDrillDown].Text);
            xmlw.WriteAttributeString("EnableKeepTogetherFormula", sf.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableKeepTogether].Text);
            xmlw.WriteAttributeString("EnableNewPageAfterFormula", sf.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableNewPageAfter].Text);
            xmlw.WriteAttributeString("EnableNewPageBeforeFormula", sf.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableNewPageBefore].Text);
            xmlw.WriteAttributeString("EnablePrintAtBottomOfPageFormula", sf.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnablePrintAtBottomOfPage].Text);
            xmlw.WriteAttributeString("EnableResetPageNumberAfterFormula", sf.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableResetPageNumberAfter].Text);
            xmlw.WriteAttributeString("EnableSuppressFormula", sf.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableSuppress].Text);
            xmlw.WriteAttributeString("EnableSuppressIfBlankFormula", sf.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableSuppressIfBlank].Text);
            xmlw.WriteAttributeString("EnableUnderlaySectionFormula", sf.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableUnderlaySection].Text);
            xmlw.WriteAttributeString("GroupNumberPerPageFormula", sf.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeGroupNumberPerPage].Text);
            xmlw.WriteAttributeString("RecordNumberPerPageFormula", sf.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeRecordNumberPerPage].Text);
            xmlw.WriteEndElement();
        }
        
    }
}
