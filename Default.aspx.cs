using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
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
                ViewState["filePath"] = filePath;
                ExcelFileUpload.SaveAs(filePath); // saves on the local directory of the project under "Files" folder
                System.Diagnostics.Debug.WriteLine(filePath);
                string FileExtension = Path.GetExtension(filePath); // gets the extension of the uploaded file
                System.Diagnostics.Debug.WriteLine(FileExtension);

                if (File.Exists(filePath))
                {
                    if (FileExtension == ".xls" || FileExtension == ".xlsx" || FileExtension == ".csv") // only accepts .xls, .xlsx, and .csv files
                    {
                        ExcelPackage.LicenseContext = LicenseContext.Commercial;
                        ExcelPackage package = new ExcelPackage(new FileInfo(filePath));
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Assuming data is in the first sheet

                        int rowCount = worksheet.Dimension.Rows;
                        int colCount = worksheet.Dimension.Columns;

                        ViewState["rowCount"] = rowCount;
                        ViewState["colCount"] = colCount;

                        List<string> excelColumns = new List<string>();
                        for (int col = 1; col <= colCount; col++)
                        {
                            excelColumns.Add(worksheet.Cells[1, col].Value?.ToString() ?? "");
                        }

                        List<string> sqlColumns = GetSqlTableColumns();

                        List<List<string>> cellData = new List<List<string>>();
                        for (int row = 1; row <= rowCount; row++)
                        {
                            List<string> rowData = new List<string>();
                            for (int col = 1; col <= colCount; col++)
                            {
                                rowData.Add(worksheet.Cells[row, col].Value?.ToString() ?? "");
                            }
                            cellData.Add(rowData);
                        }
                        
                        // Pass data to the view
                        ViewState["excelColumns"] = excelColumns;
                        ViewState["sqlColumns"] = sqlColumns;
                        ViewState["cellData"] = cellData;

                        // Call a function to regenerate the form with dynamic elements
                        GenerateMappingForm();

                        package.Dispose(); // Dispose the package after use

                        SubmitButton.Visible = true;
                    }
                    else
                    {
                        Response.Write("Wrong file extension");
                    }
                }
                else
                {
                    // Handle case where the file does not exist
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
                string tableName = "cars"; // Your table name
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
            mappingContainer.Controls.Clear(); // Clear any existing controls

            // Fetch the data you stored in ViewState
            List<string> excelColumns = (List<string>)ViewState["excelColumns"];
            List<string> sqlColumns = (List<string>)ViewState["sqlColumns"];

            // Create DropDownLists dynamically
            for (int i = 0; i < excelColumns.Count; i++)
            {
                Label excelLabel = new Label();
                excelLabel.Text = excelColumns[i];

                DropDownList sqlDropDown = new DropDownList();
                sqlDropDown.EnableViewState = true;
                sqlDropDown.ID = "sqlColumnMapping_" + i;

                sqlDropDown.Items.Add("None");
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
            // Fetch data from ViewState
            List<string> excelColumns = (List<string>)ViewState["excelColumns"];
            List<List<string>> cellData = (List<List<string>>)ViewState["cellData"];

            // Create DataTable
            DataTable excelData = new DataTable();
            for (int i = 0; i < excelColumns.Count; i++)
            {
                excelData.Columns.Add(excelColumns[i]);
            }

            // Populate DataTable from cellData
            foreach (List<string> rowData in cellData)
            {
                DataRow dataRow = excelData.NewRow();
                for (int i = 0; i < rowData.Count; i++)
                {
                    dataRow[i] = rowData[i];
                }
                excelData.Rows.Add(dataRow);
            }

            // Mappings
            Dictionary<string, string> mappings = GetColumnMappings();

            ViewState["excelData"] = excelData;
            ViewState["columnMapping"] = mappings;

            // *** SQL IMPORT LOGIC WOULD GO HERE ***  
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

                    // Get the corresponding Excel column (you might need to adjust how to retrieve it based on your UI structure)
                    string excelColumn = FindCorrespondingExcelColumn(sqlDropDown);

                    // Get the selected SQL column 
                    string selectedSqlColumn = sqlDropDown.SelectedValue;

                    if (selectedSqlColumn != "None") // Ensure a SQL mapping is chosen
                    {
                        System.Diagnostics.Debug.WriteLine(excelColumn);
                        System.Diagnostics.Debug.WriteLine(selectedSqlColumn);
                        mappings.Add(excelColumn, selectedSqlColumn);
                    }
                }
            }

            return mappings;
        }
        private string FindCorrespondingExcelColumn(DropDownList sqlDropDown)
        {
            Control previousControl = sqlDropDown.Parent.Controls[sqlDropDown.Parent.Controls.IndexOf(sqlDropDown) - 1];
            if (previousControl is Label)
            {
                return ((Label)previousControl).Text;
            }
            else
            {
                // Handle the case where no previous Label is found (potential error)
                return null;
            }
        }
    }
}