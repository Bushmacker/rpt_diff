using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

using CrystalDecisions.ReportAppServer.DataDefModel;

using ExtensionMethods;

namespace rpt_diff
{
    class DataDefModel
    {
        public static void ProcessCustomFunctions(CustomFunctions cfs, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("CustomFunctions");
            xmlw.WriteAttributeString("Count", cfs.Count.ToStringSafe());
            foreach (CustomFunction cf in cfs)
            {
                ProcessCustomFunction(cf, xmlw);
            }
            xmlw.WriteEndElement();
        }

        private static void ProcessCustomFunction(CustomFunction cf, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("CustomFunction");
            xmlw.WriteAttributeString("Name", cf.Name);
            xmlw.WriteAttributeString("Syntax", cf.Syntax.ToStringSafe());
            xmlw.WriteString(cf.Text);
            xmlw.WriteEndElement();
        }

        public static void ProcessDatabase(Database database, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("Database");
            ProcessTables(database.Tables, xmlw);
            ProcessTableLinks(database.TableLinks, xmlw);
            xmlw.WriteEndElement();
        }

        private static void ProcessTables(Tables tbls, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("Tables");
            xmlw.WriteAttributeString("Count", tbls.Count.ToStringSafe());
            foreach (Table tbl in tbls)
            {
                ProcessTable(tbl, xmlw);
            }
            xmlw.WriteEndElement();
        }

        private static void ProcessTable(Table tbl, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("Table");
            xmlw.WriteAttributeString("Alias", tbl.Alias);
            xmlw.WriteAttributeString("Description", tbl.Description);
            xmlw.WriteAttributeString("Name", tbl.Name);
            xmlw.WriteAttributeString("QualifiedName", tbl.QualifiedName);

            ProcessFields(tbl.DataFields, xmlw, "Data");
            ProcessConnectionInfo(tbl.ConnectionInfo, xmlw);
            xmlw.WriteEndElement();
        }

        private static void ProcessFields(Fields flds, XmlWriter xmlw, string type)
        {
            xmlw.WriteStartElement(type+"Fields");
            xmlw.WriteAttributeString("Count", flds.Count.ToStringSafe());
            foreach (Field fld in flds)
            {
                ProcessField(fld, xmlw,type);
            }
            xmlw.WriteEndElement();
        }

        private static void ProcessField(Field fld, XmlWriter xmlw,string type)
        {
            xmlw.WriteStartElement(type+"Field");
            xmlw.WriteAttributeString("Description", fld.Description);
            xmlw.WriteAttributeString("FormulaForm", fld.FormulaForm);
            xmlw.WriteAttributeString("HeadingText", fld.HeadingText);
            xmlw.WriteAttributeString("IsRecurring", fld.IsRecurring.ToStringSafe());
            xmlw.WriteAttributeString("Kind", fld.Kind.ToStringSafe());
            xmlw.WriteAttributeString("Length", fld.Length.ToStringSafe());
            xmlw.WriteAttributeString("LongName", fld.LongName);
            xmlw.WriteAttributeString("Name", fld.Name);
            xmlw.WriteAttributeString("ShortName", fld.ShortName);
            xmlw.WriteAttributeString("Type", fld.Type.ToStringSafe());
            xmlw.WriteAttributeString("UseCount", fld.UseCount.ToStringSafe());
            switch (fld.Kind)
            {
                case CrFieldKindEnum.crFieldKindDBField:
                    {
                        DBField dbf = (DBField)fld;
                        xmlw.WriteAttributeString("TableAlias", dbf.TableAlias);
                        break;
                    }
                case CrFieldKindEnum.crFieldKindFormulaField:
                    {
                        FormulaField ff = (FormulaField)fld;
                        xmlw.WriteAttributeString("FormulaNullTreatment", ff.FormulaNullTreatment.ToStringSafe());
                        xmlw.WriteAttributeString("Text", ff.Text.ToStringSafe());
                        xmlw.WriteAttributeString("Syntax", ff.Syntax.ToStringSafe());
                        
                        break;
                    }
                case CrFieldKindEnum.crFieldKindGroupNameField:
                    {
                        GroupNameField gnf = (GroupNameField)fld;
                        xmlw.WriteAttributeString("Group", gnf.Group.ConditionField.FormulaForm);
                        break;
                    }
                case CrFieldKindEnum.crFieldKindParameterField:
                    {
                        ParameterField pf = (ParameterField)fld;
                        xmlw.WriteAttributeString("AllowCustomCurrentValues", pf.AllowCustomCurrentValues.ToStringSafe());
                        xmlw.WriteAttributeString("AllowMultiValue", pf.AllowMultiValue.ToStringSafe());
                        xmlw.WriteAttributeString("AllowNullValue", pf.AllowNullValue.ToStringSafe());
                        if (pf.BrowseField != null)
                        {
                            xmlw.WriteAttributeString("BrowseField", pf.BrowseField.FormulaForm);
                        }
                        xmlw.WriteAttributeString("DefaultValueSortMethod", pf.DefaultValueSortMethod.ToStringSafe());
                        xmlw.WriteAttributeString("DefaultValueSortOrder", pf.DefaultValueSortOrder.ToStringSafe());
                        xmlw.WriteAttributeString("EditMask", pf.EditMask);
                        xmlw.WriteAttributeString("IsEditableOnPanel", pf.IsEditableOnPanel.ToStringSafe());
                        xmlw.WriteAttributeString("IsOptionalPrompt", pf.IsOptionalPrompt.ToStringSafe());  
                        xmlw.WriteAttributeString("IsShownOnPanel", pf.IsShownOnPanel.ToStringSafe());

                        xmlw.WriteAttributeString("ParameterType", pf.ParameterType.ToStringSafe());
                        xmlw.WriteAttributeString("ReportName", pf.ReportName);
                        xmlw.WriteAttributeString("ValueRangeKind", pf.ValueRangeKind.ToStringSafe());

                       
                        if (pf.CurrentValues != null)
                        {
                            ProcessValues(pf.CurrentValues, xmlw, "Current");
                        }
                        if (pf.DefaultValues != null)
                        {
                            ProcessValues(pf.DefaultValues, xmlw, "Default");
                        }
                        if (pf.InitialValues != null)
                        {
                            ProcessValues(pf.InitialValues, xmlw, "Initial");
                        }
                        if (pf.Values != null)
                        {
                            ProcessValues(pf.Values, xmlw, "");
                        }
                        ProcessValue(pf.MaximumValue as Value, xmlw, "Maximum");
                        ProcessValue(pf.MinimumValue as Value, xmlw, "Minimum");
                        
                        
                        break;   
                    }
                case CrFieldKindEnum.crFieldKindRunningTotalField:
                    {
                        RunningTotalField rtf = (RunningTotalField)fld;
                        xmlw.WriteAttributeString("EvaluateCondition", ProcessCondition(rtf.EvaluateCondition));
                        xmlw.WriteAttributeString("EvaluateConditionType", rtf.EvaluateConditionType.ToStringSafe());
                        xmlw.WriteAttributeString("Operation", rtf.Operation.ToStringSafe());
                        xmlw.WriteAttributeString("ResetCondition", ProcessCondition(rtf.ResetCondition));
                        xmlw.WriteAttributeString("ResetConditionType", rtf.ResetConditionType.ToStringSafe());
                        xmlw.WriteAttributeString("SummarizedField", rtf.SummarizedField.FormulaForm);
                        break;
                    }
                case CrFieldKindEnum.crFieldKindSpecialField:
                    {
                        SpecialField sf = (SpecialField)fld;
                        xmlw.WriteAttributeString("SpecialType", sf.SpecialType.ToStringSafe());
                        break;
                    }
                case CrFieldKindEnum.crFieldKindSummaryField: 
                    {
                        SummaryField smf = (SummaryField)fld;
                        if (smf.Group != null)
                        {
                            xmlw.WriteAttributeString("Group", smf.Group.ConditionField.FormulaForm);
                        }
                        xmlw.WriteAttributeString("Operation", smf.Operation.ToStringSafe());
                        xmlw.WriteAttributeString("SummarizedField", smf.SummarizedField.FormulaForm);
                        break;
                    }
                case CrFieldKindEnum.crFieldKindUnknownField:
                    {
                        break;
                    }
            }
            xmlw.WriteEndElement();
        }

        private static void ProcessValues(Values values, XmlWriter xmlw, string type)
        {
            xmlw.WriteStartElement(type+"Values");
            xmlw.WriteAttributeString("Count", values.Count.ToStringSafe());
            foreach (Value val in values)
            {
                ProcessValue(val, xmlw);
            } 
            xmlw.WriteEndElement();
            
        }

        private static void ProcessValue(Value val, XmlWriter xmlw, string type="")
        {
            xmlw.WriteStartElement(type+"Value");
            ConstantValue cvc = val as ConstantValue;
            ExpressionValue ev = val as ExpressionValue;
            ParameterFieldDiscreteValue pfdv = val as ParameterFieldDiscreteValue;
            ParameterFieldRangeValue pfrv = val as ParameterFieldRangeValue;
            ParameterFieldValue pfv = val as ParameterFieldValue;
            if (cvc != null)
            {
                xmlw.WriteAttributeString("Value", cvc.Value.ToStringSafe());
            }
            else if (ev != null)
            {
                xmlw.WriteAttributeString("Expression", ev.Expression);
            }
            else if (pfdv != null)
            {
                xmlw.WriteAttributeString("Description", pfdv.Description);
                xmlw.WriteAttributeString("Value", Convert.ToString(pfdv.Value));
            }
            else if (pfrv != null)
            {
                xmlw.WriteAttributeString("BeginValue", Convert.ToString(pfrv.BeginValue));
                xmlw.WriteAttributeString("Description", pfrv.Description);
                xmlw.WriteAttributeString("EndValue", Convert.ToString(pfrv.EndValue));
                xmlw.WriteAttributeString("LowerBoundType", Convert.ToString(pfrv.LowerBoundType));
                xmlw.WriteAttributeString("UpperBoundType", Convert.ToString(pfrv.UpperBoundType));
            }
            else if (pfv != null)
            {
                xmlw.WriteAttributeString("Description", pfv.Description);
            }
            xmlw.WriteEndElement();
        }
        private static string ProcessCondition(object condition)
        {
            Field dfdCondition = condition as Field;
            Group gCondition = condition as Group;
            if (dfdCondition != null) // Field
            {
                return dfdCondition.FormulaForm;
            }
            else if (gCondition != null) // Group
            {
                return gCondition.ConditionField.FormulaForm;
            }
            else // Custom formula
            {
                return condition.ToStringSafe();
            }
        }

        private static void ProcessConnectionInfo(ConnectionInfo ci, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("ConnectionInfo");
            xmlw.WriteAttributeString("Kind", ci.Kind.ToStringSafe());
            xmlw.WriteAttributeString("Password", ci.Password);
            xmlw.WriteAttributeString("UserName", ci.UserName);
            ProcessPropertyBag(ci.Attributes, xmlw);
            xmlw.WriteEndElement();
        }

        private static void ProcessPropertyBag(PropertyBag pb, XmlWriter xmlw)
        {
            foreach (string pid in pb.PropertyIDs)
            {
                xmlw.WriteAttributeString(pid.Replace(" ", string.Empty), pb.StringValue[pid]);
            }
        }

        private static void ProcessTableLinks(TableLinks tls, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("TableLinks");
            xmlw.WriteAttributeString("Count", tls.Count.ToStringSafe());
            foreach (TableLink tl in tls)
            {
                ProcessTableLink(tl, xmlw);
            }
            xmlw.WriteEndElement();
        }

        private static void ProcessTableLink(TableLink tl, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("TableLink");
            xmlw.WriteAttributeString("JoinType", tl.JoinType.ToStringSafe());
            xmlw.WriteAttributeString("SourceTableAlias", tl.SourceTableAlias);
            xmlw.WriteAttributeString("TargetTableAlias", tl.TargetTableAlias);
            foreach (string sfn in tl.SourceFieldNames)
            {
                xmlw.WriteStartElement("SourceField");
                xmlw.WriteAttributeString("Name", sfn);
                xmlw.WriteEndElement();
            }
            foreach (string tfn in tl.TargetFieldNames)
            {
                xmlw.WriteStartElement("TargetField");
                xmlw.WriteAttributeString("Name", tfn);
                xmlw.WriteEndElement();
            }
            xmlw.WriteEndElement();
        }

        public static void ProcessDataDefinition(DataDefinition dd, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("DataDefinition");
            ProcessFields(dd.FormulaFields, xmlw, "Formula");
            ProcessFilter(dd.GroupFilter, xmlw, "GroupFilter");
            ProcessGroups(dd.Groups, xmlw);
            ProcessFields(dd.ParameterFields, xmlw,"Parameter");
            ProcessFilter(dd.RecordFilter, xmlw, "RecordFilter");
            //ProcessFields(dd.ResultFields, xmlw, "Result"); // redundant only shows fields that are used in the report 
            ProcessFields(dd.RunningTotalFields, xmlw, "RunningTotal");
            ProcessFilter(dd.SavedDataFilter, xmlw, "SavedDataFilter");
            ProcessSorts(dd.Sorts, xmlw);
            ProcessFields(dd.SummaryFields, xmlw, "Summary");
            xmlw.WriteEndElement();
        }


        private static void ProcessFilter(Filter filter, XmlWriter xmlw,string type)
        {
            xmlw.WriteStartElement(type);
            xmlw.WriteAttributeString("FormulaNullTreatment", filter.FormulaNullTreatment.ToStringSafe());
            xmlw.WriteAttributeString("FreeEditingText", filter.FreeEditingText);
            xmlw.WriteAttributeString("Name", filter.Name);
            ProcessFilterItems(filter.FilterItems, xmlw);
            xmlw.WriteEndElement();
        }

        private static void ProcessFilterItems(FilterItems fis, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("FilterItems");
            xmlw.WriteAttributeString("Count", fis.Count.ToStringSafe());
            foreach (FilterItem fi in fis) 
            {
                ProcessFilterItem(fi, xmlw);
            }
            xmlw.WriteEndElement();
        }

        private static void ProcessFilterItem(FilterItem fi, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("FilterItem");
            xmlw.WriteAttributeString("ComputeText", fi.ComputeText());
            xmlw.WriteEndElement();
        }

        private static void ProcessGroups(Groups groups, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("Groups");
            xmlw.WriteAttributeString("Count", groups.Count.ToStringSafe());
            foreach (Group group in groups)
            {
                ProcessGroup(group, xmlw);
            }
            xmlw.WriteEndElement();
        }
        private static void ProcessGroup(Group group, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("Group");
            xmlw.WriteAttributeString("ConditionField", group.ConditionField.FormulaForm);
            ProcessGroupOptions(group.Options, xmlw);
            xmlw.WriteEndElement();
        }

        private static void ProcessGroupOptions(ISCRGroupOptions go, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("GroupOptions");

            xmlw.WriteAttributeString("GroupNameFormula", go.ConditionFormulas[CrGroupOptionsConditionFormulaTypeEnum.crGroupName].Text);
            xmlw.WriteAttributeString("SortDirectionFormula", go.ConditionFormulas[CrGroupOptionsConditionFormulaTypeEnum.crSortDirection].Text);
            DateGroupOptions dgo = go as DateGroupOptions;
            SpecifiedGroupOptions sgo = go as SpecifiedGroupOptions;
            if (dgo != null)
            {
                xmlw.WriteAttributeString("DateCondition", dgo.DateCondition.ToStringSafe());
            }
            if (sgo != null)
            {
                xmlw.WriteAttributeString("SpecifiedValueFilters", sgo.SpecifiedValueFilters.ToStringSafe());
                xmlw.WriteAttributeString("UnspecifiedValuesName", sgo.UnspecifiedValuesName.ToStringSafe());
                xmlw.WriteAttributeString("UnspecifiedValuesType", sgo.UnspecifiedValuesType.ToStringSafe());
            }
            xmlw.WriteEndElement();
        }

        private static void ProcessSorts(Sorts sorts, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("Sorts");
            xmlw.WriteAttributeString("Count", sorts.Count.ToStringSafe());
            foreach (Sort sort in sorts)
            {
                ProcessSort(sort, xmlw);
            }
            xmlw.WriteEndElement();
        }

        private static void ProcessSort(Sort sort, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("Sort");
            xmlw.WriteAttributeString("Direction", sort.Direction.ToStringSafe());
            xmlw.WriteAttributeString("SortField", sort.SortField.FormulaForm);
            xmlw.WriteEndElement();
        }

        public static void ProcessSummaryInfo(SummaryInfo si, XmlWriter xmlw)
        {
            xmlw.WriteStartElement("SummaryInfo");
            xmlw.WriteAttributeString("Author", si.Author);
            xmlw.WriteAttributeString("Comments", si.Comments);
            xmlw.WriteAttributeString("IsSavingWithPreview", si.IsSavingWithPreview.ToStringSafe());
            xmlw.WriteAttributeString("Keywords", si.Keywords);
            xmlw.WriteAttributeString("LastSavedBy", si.LastSavedBy);
            xmlw.WriteAttributeString("RevisionNumber", si.RevisionNumber);
            xmlw.WriteAttributeString("Subject", si.Subject);
            xmlw.WriteAttributeString("Title", si.Title);
            xmlw.WriteEndElement();
        }

    }
}
