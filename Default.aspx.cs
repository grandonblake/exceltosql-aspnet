using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using OfficeOpenXml;

namespace exceltosql
{
    public partial class Default : System.Web.UI.Page
    {
        string connectionString = @"Server=IOT-LT98\SQLEXPRESS;Database=test;Trusted_Connection=True;";
        protected void Page_Load(object sender, EventArgs e)
        {
            AsyncPostBackTrigger trigger = new AsyncPostBackTrigger();
            trigger.ControlID = UploadButton.UniqueID;
            trigger.EventName = "Click";
            UpdatePanel1.Triggers.Add(trigger);

            // Generate form if Excel data is available (after an upload)
            if (ViewState["excelColumns"] != null)
            {
                GenerateMappingForm();
            }
        }
        protected void UploadButton_Click(object sender, EventArgs e)
        {
            if (ExcelFileUpload.HasFile)
            {
                string filePath = Server.MapPath("~/Files/") + ExcelFileUpload.FileName;
                ExcelFileUpload.SaveAs(filePath); // saves on the local directory of the project under "Files" folder
                string FileExtension = Path.GetExtension(filePath); // gets the extension of the uploaded file

                if (File.Exists(filePath))
                {
                    if (FileExtension == ".xls" || FileExtension == ".xlsx" || FileExtension == ".csv") // only accepts .xls, .xlsx, and .csv files
                    {
                        ExcelPackage.LicenseContext = LicenseContext.Commercial;
                        ExcelPackage package = new ExcelPackage(new FileInfo(filePath));
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Assuming data is in the first sheet

                        int rowCount = worksheet.Dimension.Rows;
                        int colCount = worksheet.Dimension.Columns;

                        List<string> excelColumns = new List<string>();
                        for (int col = 1; col <= colCount; col++)
                        {
                            excelColumns.Add(worksheet.Cells[1, col].Value?.ToString() ?? "");
                        }

                        // Calls the function to get the table columns of the SQL database
                        List<string> sqlColumns = GetSqlTableColumns();

                        DataTable excelData = new DataTable();
                        for (int i = 0; i < excelColumns.Count; i++)
                        {
                            excelData.Columns.Add(excelColumns[i]); // add columns of excel file to excelData
                        }

                        for (int row = 1; row <= rowCount; row++)
                        {
                            DataRow dataRow = excelData.NewRow();
                            for (int col = 1; col <= colCount; col++)
                            {
                                dataRow[col - 1] = worksheet.Cells[row, col].Value?.ToString() ?? ""; // add rows of excel file to excelData
                            }
                            excelData.Rows.Add(dataRow);
                        }

                        ViewState["excelColumns"] = excelColumns; // excel column names only
                        ViewState["sqlColumns"] = sqlColumns; // sql column names only
                        ViewState["excelData"] = excelData; // content of excel cells (with column names)

                        // Call the function to regenerate the form with dynamic elements
                        GenerateMappingForm();

                        package.Dispose(); // Dispose the package after use

                        SubmitButton.Visible = true;
                    }
                    else
                    {
                        Response.Write("Wrong file extension");
                    }
                }
            }
            else
            {
                Response.Write("No file is selected");
            }
        }

        private List<string> GetSqlTableColumns()
        {
            List<string> columnNames = new List<string>();
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                string tableName = "cars"; // the table name
                string query = $"SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '{tableName}'";

                SqlCommand command = new SqlCommand(query, connection);
                connection.Open();

                using (SqlDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        columnNames.Add(reader.GetString(0));
                    }
                }
            }
            return columnNames;
        }

        private void GenerateMappingForm()
        {
            mappingContainer.Controls.Clear();

            // Fetch the data you stored in ViewState
            List<string> excelColumns = (List<string>)ViewState["excelColumns"]; // excel column names only
            List<string> sqlColumns = (List<string>)ViewState["sqlColumns"]; // sql column names only

            // Create DropDownLists dynamically
            for (int i = 0; i < excelColumns.Count; i++)
            {
                Label excelLabel = new Label();
                excelLabel.Text = excelColumns[i]; // the label for the excel column name

                DropDownList sqlDropDown = new DropDownList(); // the dropdown list for the sql column name
                sqlDropDown.EnableViewState = true;
                sqlDropDown.ID = "sqlColumnMapping_" + i;

                sqlDropDown.Items.Add("None"); //the "None" option
                // Populate SQL Columns
                foreach (string sqlCol in sqlColumns)
                {
                    sqlDropDown.Items.Add(sqlCol);
                }

                mappingContainer.Controls.Add(excelLabel);
                mappingContainer.Controls.Add(sqlDropDown);
                mappingContainer.Controls.Add(new LiteralControl("<br />"));
            }
        }

        protected void SubmitButton_Click(object sender, EventArgs e)
        {
            // Mappings
            Dictionary<string, string> mappings = GetColumnMappings();

            System.Diagnostics.Debug.WriteLine("mappings");
            foreach (KeyValuePair<string, string> kvp in mappings)
            {
                System.Diagnostics.Debug.WriteLine($"Key = {kvp.Key}, Value = {kvp.Value}");
            }

            DataTable excelData = (DataTable)ViewState["excelData"]; // gets content of cells of excel file

            System.Diagnostics.Debug.WriteLine("excelData");
            foreach (DataRow row in excelData.Rows)
            {
                List<string> cellValues = new List<string>();
                foreach (DataColumn column in excelData.Columns)
                {
                    cellValues.Add(row[column].ToString());
                }
                Debug.WriteLine(string.Join(", ", cellValues));
            }

            Dictionary<string, List<string>> mappedData = new Dictionary<string, List<string>>();

            foreach (KeyValuePair<string, string> mapping in mappings)
            {
                mappedData[mapping.Value] = new List<string>();
            }

            for (int i = 1; i < excelData.Rows.Count; i++)
            {
                DataRow row = excelData.Rows[i];
                foreach (DataColumn column in excelData.Columns)
                {
                    if (mappings.ContainsKey(column.ColumnName))
                    {
                        mappedData[mappings[column.ColumnName]].Add(row[column].ToString());
                    }
                }
            }

            System.Diagnostics.Debug.WriteLine("mappedData");
            foreach (KeyValuePair<string, List<string>> kvp in mappedData)
            {
                System.Diagnostics.Debug.WriteLine($"{kvp.Key}: {string.Join(", ", kvp.Value)}");
            }

            List<Dictionary<string, string>> dataToBeInserted = new List<Dictionary<string, string>>();

            for (int i = 0; i < mappedData.First().Value.Count; i++)
            {
                Dictionary<string, string> row = new Dictionary<string, string>();
                foreach (KeyValuePair<string, List<string>> kvp in mappedData)
                {
                    row[kvp.Key] = kvp.Value[i];
                }
                dataToBeInserted.Add(row);
            }

            System.Diagnostics.Debug.WriteLine("dataToBeInserted");
            foreach (Dictionary<string, string> row in dataToBeInserted)
            {
                List<string> cellValues = new List<string>();
                foreach (KeyValuePair<string, string> kvp in row)
                {
                    cellValues.Add($"{kvp.Key}: {kvp.Value}");
                }
                Debug.WriteLine("{" + string.Join(", ", cellValues) + "}");
            }


        }

        private Dictionary<string, string> GetColumnMappings()
        {
            Dictionary<string, string> mappings = new Dictionary<string, string>();

            // Iterate through your mapping container's controls
            foreach (Control control in mappingContainer.Controls)
            {
                if (control is DropDownList)
                {
                    DropDownList sqlDropDown = (DropDownList)control;

                    // Get the corresponding Excel column
                    Control previousControl = sqlDropDown.Parent.Controls[sqlDropDown.Parent.Controls.IndexOf(sqlDropDown) - 1];
                    string excelColumn = null;
                    if (previousControl is Label)
                    {
                        excelColumn = ((Label)previousControl).Text;
                    }

                    // Get the selected SQL column 
                    string selectedSqlColumn = sqlDropDown.SelectedValue;

                    if (selectedSqlColumn != "None" && excelColumn != null) // Ensure a SQL mapping is chosen and Excel column is found
                    {
                        mappings.Add(excelColumn, selectedSqlColumn);
                    }
                }
            }

            return mappings;
        }

    }
}