using System;
using System.Globalization;
using System.Linq;
using Microsoft.SharePoint;
using Microsoft.SharePoint.WebControls;
using Microsoft.SharePoint.Taxonomy;
using System.IO;
using System.Web.UI;
using System.Web;
using System.Web.UI.HtmlControls;

namespace Rjesh.Solutions
{
    public partial class ExportToExcel : LayoutsPageBase
    {
        HttpResponse _exportResponse;
        bool _isErrorOccured = false;

        protected void Page_Init(object sender, EventArgs e)
        {
            _exportResponse = Response;
            // Get view and list guids from the request.
            Guid listGUID = new Guid(Request["ListGuid"].ToString(CultureInfo.InvariantCulture));
            Guid viewGUID = new Guid(Request["ViewGuid"].ToString(CultureInfo.InvariantCulture));

            // Get the list item ids in csv format.
            var listItemsID = Request["IDDict"].ToString(CultureInfo.InvariantCulture).Split(new Char[] { ',' });
            ExportSelectedItemsToExcel(listGUID, viewGUID, listItemsID);

        }
        protected void Page_Load(object sender, EventArgs e)
        {

        }
        public void ExportSelectedItemsToExcel(Guid listID, Guid listViewID, String[] listIdsCsv)
        {

            try
            {
                SPList list = SPContext.Current.Web.Lists[listID];
                SPView listView = list.Views[listViewID];

                var exportListTable = new HtmlTable {Border = 1, CellPadding = 3, CellSpacing = 3};
                // Set the table's formatting-related properties.

                // Start adding content to the table.
                HtmlTableCell htmlcell;

                // Add header row in HTML table
                var htmlrow = new HtmlTableRow();
                SPViewFieldCollection viewHeaderFields = listView.ViewFields;
                for (var index = 0; index < viewHeaderFields.Count; index++)
                    foreach (
                        SPField field in
                            listView.ParentList.Fields.Cast<SPField>().Where(
                                field => field.InternalName == viewHeaderFields[index]))
                    {
                        if (!field.Hidden)
                        {
                            htmlcell = new HtmlTableCell
                                           {
                                               BgColor = "#0099FF",
                                               InnerHtml = field.Title.ToString(CultureInfo.InvariantCulture)
                                           };
                            htmlrow.Cells.Add(htmlcell);
                        }
                        break;
                    }
                exportListTable.Rows.Add(htmlrow);

                // Add rows in HTML table based on the fields in view.
                foreach (string id in listIdsCsv.Where(id => !String.IsNullOrEmpty(id)))
                {
                    htmlrow = new HtmlTableRow();

                    SPListItem item = list.GetItemById(Convert.ToInt32(id));
                    SPViewFieldCollection viewFields = listView.ViewFields;
                    for (var i = 0; i < viewFields.Count; i++)
                    {
                        foreach (SPField field in listView.ParentList.Fields.Cast<SPField>().Where(field => field.InternalName == viewFields[i]))
                        {
                            if (!field.Hidden)
                            {
                                htmlcell = new HtmlTableCell();
                                if (item[field.InternalName] != null)
                                {
                                            
                                    if((string.CompareOrdinal(field.TypeAsString, "TaxonomyFieldType") == 0)||(string.CompareOrdinal(field.TypeAsString, "TaxonomyFieldTypeMulti") == 0))
                                    {
                                        htmlcell.InnerHtml = GetValuesAsCsvFromTaxonomyField(field, item);
                                    }
                                    else if (field.Type == SPFieldType.Lookup)
                                    {
                                        // call the method to get lookup field as csv
                                        htmlcell.InnerHtml = GetValuesAsCsvFromLookupField(field, item);
                                    }
                                    else if (field.Type == SPFieldType.User)
                                    {
                                        htmlcell.InnerHtml = GetValuesAsCsvFromUserField(field, item);
                                    }
                                    else if (field.Type == SPFieldType.Invalid)
                                    {
                                        htmlcell.InnerHtml = GetValuesAsCSVFromInvalidTypeField(field, item);
                                    }
                                    else if (field.Type == SPFieldType.Calculated)
                                    {
                                        var cf = (SPFieldCalculated)field;
                                        htmlcell.InnerHtml = cf.GetFieldValueAsText(item[field.InternalName]);
                                    }
                                    else
                                    {
                                        htmlcell.InnerHtml = item[field.InternalName].ToString();
                                    }
                                }
                                else
                                {
                                    htmlcell.InnerHtml = String.Empty;
                                }
                                htmlrow.Cells.Add(htmlcell);
                            }
                            break;
                        }
                    }
                    exportListTable.Rows.Add(htmlrow);
                }

                // Write the HTML table contents to response as excel file
                using (var sw = new StringWriter())
                {
                    using (var htw = new HtmlTextWriter(sw))
                    {
                        exportListTable.RenderControl(htw);
                        _exportResponse.Clear();
                        _exportResponse.ContentType = "application/vnd.ms-excel";
                        _exportResponse.AddHeader("content-disposition", string.Format("attachment; filename={0}", list.Title + ".xls"));
                        _exportResponse.Cache.SetCacheability(HttpCacheability.NoCache);
                        _exportResponse.ContentEncoding = System.Text.Encoding.Unicode;
                        _exportResponse.BinaryWrite(System.Text.Encoding.Unicode.GetPreamble());
                        _exportResponse.Write(sw.ToString());
                        _exportResponse.End();
                    }
                }
            }
            catch (System.Threading.ThreadAbortException exception)
            {
                // Do nothing on thread abort exception.
            }
            catch (Exception ex)
            {
                _isErrorOccured = true;
                SPHelper.LogError("Rjesh.Solutions.Exportoexcel", ex.Message.ToString(CultureInfo.InvariantCulture));
            }
            finally
            {
                if (_isErrorOccured)
                {
                    SPHelper.LogError("Rjesh.Solutions.Exportoexcel", "Export completed with errors");
                }               

            }

        }
        private string GetValuesAsCSVFromInvalidTypeField(SPField field, SPListItem item)
        {
            string csvLookupString = string.Empty;
            var lookupField = (SPFieldLookup)field;

            if (lookupField.AllowMultipleValues)
            {
                var values = item[field.InternalName] as SPFieldLookupValueCollection;
                if (values != null)
                    csvLookupString = values.Aggregate(csvLookupString, (current, value) => current + value.LookupValue.ToString(CultureInfo.InvariantCulture) + "; ");
            }
            else
            {
                var fieldValue = new SPFieldLookupValue(item[field.InternalName].ToString());
                csvLookupString = fieldValue.LookupValue.ToString(CultureInfo.InvariantCulture);
            }
            return csvLookupString;
        }
        public String GetValuesAsCsvFromLookupField(SPField field, SPItem item)
        {
            string csvLookupString = string.Empty;
            var lookupField = (SPFieldLookup)field;

            if (lookupField.AllowMultipleValues)
            {
                var values = item[field.InternalName] as SPFieldLookupValueCollection;
                if (values != null)
                    csvLookupString = values.Aggregate(csvLookupString, (current, value) => current + value.LookupValue.ToString(CultureInfo.InvariantCulture) + "; ");
            }
            else
            {
                var fieldValue = new SPFieldLookupValue(item[field.InternalName].ToString());
                csvLookupString = fieldValue.LookupValue.ToString(CultureInfo.InvariantCulture);
            }
            return csvLookupString;
        }
        public String GetValuesAsCsvFromUserField(SPField field, SPItem item)
        {
            var csvLookupString = string.Empty;            
            var userFieldValueCol = new SPFieldUserValueCollection(SPContext.Current.Web, item[field.InternalName].ToString());
            return userFieldValueCol.Select(singlevalue => singlevalue.ToString().Split('#')).Aggregate(csvLookupString, (current, userValue) => current + userValue[1] + "; ");
        }
        public String GetValuesAsCsvFromTaxonomyField(SPField field, SPItem item)
        {
            var txField = field as TaxonomyField;
            var csvLookupString = string.Empty;
            if (txField != null)
            {
                if (txField.AllowMultipleValues)
                {
                    var taxfieldValColl = (item[field.InternalName] as TaxonomyFieldValueCollection);
                    if (taxfieldValColl != null)
                        csvLookupString = taxfieldValColl.Aggregate(csvLookupString, (current, singlevalue) => current + singlevalue.Label.ToString(CultureInfo.InvariantCulture) + "; ");
                }
                else
                {
                    var singlevalue = item[field.InternalName] as TaxonomyFieldValue;
                    if (singlevalue != null) csvLookupString = csvLookupString + singlevalue.Label.ToString(CultureInfo.InvariantCulture);
                }
            }
            return csvLookupString;
        }
    }
}
