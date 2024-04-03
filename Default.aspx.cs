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
using System.Windows.Media;
using OfficeOpenXml;

namespace exceltosql
{
    public partial class Default : System.Web.UI.Page
    {

        // string connectionString = @"Server=IOT-LT98\SQLEXPRESS;Database=test;Trusted_Connection=True;";
        protected void Page_Load(object sender, EventArgs e)
        {
            // Set up asynchronous trigger for upload button
            AsyncPostBackTrigger trigger = new AsyncPostBackTrigger();
            trigger.ControlID = UploadButton.UniqueID;
            trigger.EventName = "Click";
            UpdatePanel1.Triggers.Add(trigger);

            // Generate mapping form if both Excel and SQL columns are available
            if (ViewState["excelColumns"] != null && ViewState["sqlColumns"] != null)
            {
                GenerateMappingForm();
            }
        }

        // Uploads the Excel file
        protected void UploadButton_Click(object sender, EventArgs e)
        {
            if (ExcelFileUpload.HasFile) // If the FileUpload has a file
            {
                string filePath = Server.MapPath("~/Files/") + ExcelFileUpload.FileName;
                ExcelFileUpload.SaveAs(filePath); // Save uploaded file in local directory
                ViewState["filePath"] = filePath;
                string FileExtension = Path.GetExtension(filePath); // Get file extension

                if (File.Exists(filePath)) // If the file exists in the file path
                {
                    if (FileExtension == ".xls" || FileExtension == ".xlsx" || FileExtension == ".csv") // Check file extension
                    {
                        ExcelPackage.LicenseContext = LicenseContext.Commercial;
                        ExcelPackage package = new ExcelPackage(new FileInfo(filePath));

                        WorksheetList.Items.Clear();

                        // Add worksheets to dropdown list
                        foreach (var worksheet in package.Workbook.Worksheets)
                        {
                            WorksheetList.Items.Add(new ListItem(worksheet.Name));
                        }

                        // shows the select worksheet step
                        WorksheetList.Visible = true;
                        SelectWorksheetButton.Visible = true;

                        // hides the excel upload step
                        ExcelFileUpload.Visible = false;
                        UploadButton.Visible = false;

                        // hides the error label if successful
                        ErrorLabel.Text = string.Empty;
                        ErrorLabel.Visible = false;

                        package.Dispose(); // Dispose Excel package
                    }
                    else
                    {
                        ErrorLabel.Text = "Error: Wrong file extension";
                        ErrorLabel.Visible = true;
                    }
                }
            }
            else
            {
                ErrorLabel.Text = "Error: No file is selected";
                ErrorLabel.Visible = true;
            }
        }

        // Selects the worksheet in the uploaded Excel file
        protected void SelectWorksheetButton_Click(object sender, EventArgs e)
        {
            string filePath = ViewState["filePath"].ToString(); // gets the saved file path

            if (File.Exists(filePath))
            {
                ExcelPackage.LicenseContext = LicenseContext.Commercial;
                ExcelPackage package = new ExcelPackage(new FileInfo(filePath));
                ExcelWorksheet worksheet = package.Workbook.Worksheets[WorksheetList.SelectedItem.Text]; // selects the chosen worksheet

                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                List<string> excelColumns = new List<string>(); // variable to store the uploaded excel's column names
                for (int col = 1; col <= colCount; col++)
                {
                    excelColumns.Add(worksheet.Cells[1, col].Value?.ToString() ?? ""); // Extract excel column names
                }

                DataTable excelData = new DataTable(); // variable to store the contents of the uploaded excel (column + row)
                for (int i = 0; i < excelColumns.Count; i++)
                {
                    excelData.Columns.Add(excelColumns[i]); // Add columns to data table
                }

                for (int row = 1; row <= rowCount; row++)
                {
                    DataRow dataRow = excelData.NewRow(); // variable to store the rows of the uploaded excel
                    for (int col = 1; col <= colCount; col++)
                    {
                        // Add data (rows) to data table (excelData)
                        dataRow[col - 1] = worksheet.Cells[row, col].Value?.ToString() ?? "";
                    }
                    excelData.Rows.Add(dataRow);
                }

                // Inside your SelectWorksheetButton_Click method, after the excelData is populated:

                foreach (DataRow row in excelData.Rows)
                {
                    System.Diagnostics.Debug.WriteLine("----- Row -----");
                    foreach (var item in row.ItemArray)
                    {
                        System.Diagnostics.Debug.WriteLine(item);
                    }
                }

                ViewState["excelColumns"] = excelColumns; // Store excel column names
                ViewState["excelData"] = excelData; // Store excel data

                package.Dispose(); // Dispose Excel package

                SqlServerNameLabel.Visible = true;
                SqlServerName.Visible = true;

                DatabaseNameLabel.Visible = true;
                DatabaseName.Visible = true;
                ConnectDatabase.Visible = true;


                WorksheetList.Visible = false;
                SelectWorksheetButton.Visible = false;
            }
        }

        protected void ConnectDatabaseButton_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(SqlServerName.Text) || string.IsNullOrEmpty(DatabaseName.Text))
            {
                // Show an error message
                ErrorLabel.Text = "Please fill in both the SQL Server Name and Database Name.";
                ErrorLabel.Visible = true;
            }
            else
            {
                string sqlServerName = SqlServerName.Text.Trim();
                string databaseName = DatabaseName.Text.Trim();

                string connectionString = $"Server={sqlServerName};Database={databaseName};Trusted_Connection=True;";

                ErrorLabel.Text = string.Empty;
                ErrorLabel.Visible = false;

                try
                {
                    using (SqlConnection connection = new SqlConnection(connectionString))
                    {
                        connection.Open();
                        // If the connection is successful, store the connection string in ViewState
                        ViewState["connectionString"] = connectionString;

                        // Hide the error label if the connection is successful
                        ErrorLabel.Text = string.Empty;
                        ErrorLabel.Visible = false;

                        TableList.Items.Clear();
                        List<string> tableNames = GetSqlTableNames();
                        foreach (string tableName in tableNames)
                        {
                            TableList.Items.Add(new ListItem(tableName));
                        }

                        TableList.Visible = true;
                        SelectTableButton.Visible = true;


                        SqlServerNameLabel.Visible = false;
                        SqlServerName.Visible = false;

                        DatabaseNameLabel.Visible = false;
                        DatabaseName.Visible = false;
                        ConnectDatabase.Visible = false;
                    }
                }
                catch (Exception ex)
                {
                    ErrorLabel.Text = "Error: " + ex.Message;
                    ErrorLabel.Visible = true;
                }
            }
            
        }

        // Retrieves a list of SQL Server table names
        private List<string> GetSqlTableNames()
            {
                string connectionString = ViewState["connectionString"].ToString();

                List<string> tableNames = new List<string>();
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    string query = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE'";

                    SqlCommand command = new SqlCommand(query, connection);
                    connection.Open();

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            tableNames.Add(reader.GetString(0));
                        }
                    }
                }
                return tableNames;
            }

        // Selects the table in the SQL database
        protected void SelectTableButton_Click(object sender, EventArgs e)
        {
            string selectedTable = TableList.SelectedItem.Text;
            List<string> sqlColumns = GetSqlTableColumns(selectedTable); // Get SQL table columns

            ViewState["selectedTable"] = selectedTable;
            ViewState["sqlColumns"] = sqlColumns;

            GenerateMappingForm(); // Generate mapping form

            InsertDataButton.Visible = true;


            TableList.Visible = false;
            SelectTableButton.Visible = false;
        }

            // Retrieves a list of column names for a specific SQL Server table
            private List<string> GetSqlTableColumns(string tableName)
            {
                string connectionString = ViewState["connectionString"].ToString();

                List<string> columnNames = new List<string>();
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
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

            // Dynamically creates a form for mapping Excel columns to SQL Server columns
            private void GenerateMappingForm()
            {
                mappingContainer.Controls.Clear();

                List<string> excelColumns = (List<string>)ViewState["excelColumns"];
                List<string> sqlColumns = (List<string>)ViewState["sqlColumns"];

                // Create DropDownLists for mapping
                for (int i = 0; i < excelColumns.Count; i++)
                {
                    Label excelLabel = new Label();
                    excelLabel.Text = excelColumns[i]; // Label for Excel column

                    DropDownList sqlDropDown = new DropDownList(); // Dropdown for SQL column mapping
                    sqlDropDown.EnableViewState = true;
                    sqlDropDown.ID = "sqlColumnMapping_" + i;

                    sqlDropDown.Items.Add("None"); // Add "None" option
                    foreach (string sqlCol in sqlColumns)
                    {
                        sqlDropDown.Items.Add(sqlCol); // Add SQL columns to dropdown
                    }

                    mappingContainer.Controls.Add(excelLabel);
                    mappingContainer.Controls.Add(sqlDropDown);
                    mappingContainer.Controls.Add(new LiteralControl("<br />"));
                }
            }

        protected void InsertDataButton_Click(object sender, EventArgs e)
        {
            // Get column mappings from form
            Dictionary<string, string> mappings = GetColumnMappings();

            DataTable excelData = (DataTable)ViewState["excelData"]; // Get excel data

            // Create dictionary for mapped data
            Dictionary<string, List<string>> mappedData = new Dictionary<string, List<string>>();
            foreach (KeyValuePair<string, string> mapping in mappings)
            {
                mappedData[mapping.Value] = new List<string>();
            }

            // Map excel data to SQL columns
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

            // Prepare data for insertion
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

            try
            {
                InsertData(dataToBeInserted); // Insert data into SQL Server

                // Show the success label
                SuccessLabel.Visible = true;
                InsertAgainButton.Visible = true;

                mappingContainer.Controls.Clear();
                ViewState.Clear();
                InsertDataButton.Visible = false;

                ErrorLabel.Text = string.Empty;
                ErrorLabel.Visible = false;
            }
            catch (Exception ex)
            {
                // Display the error message in the ErrorLabel
                ErrorLabel.Text = "Error: " + ex.Message;
                ErrorLabel.Visible = true;
            }
        }

            // Extracts user-defined column mappings from the form
            private Dictionary<string, string> GetColumnMappings()
            {
                Dictionary<string, string> mappings = new Dictionary<string, string>();

                foreach (Control control in mappingContainer.Controls)
                {
                    if (control is DropDownList)
                    {
                        DropDownList sqlDropDown = (DropDownList)control;

                        // Get corresponding Excel column name
                        Control previousControl = sqlDropDown.Parent.Controls[sqlDropDown.Parent.Controls.IndexOf(sqlDropDown) - 1];
                        string excelColumn = null;
                        if (previousControl is Label)
                        {
                            excelColumn = ((Label)previousControl).Text;
                        }

                        // Get selected SQL column
                        string selectedSqlColumn = sqlDropDown.SelectedValue;

                        if (selectedSqlColumn != "None" && excelColumn != null) // Check for valid mapping
                        {
                            mappings.Add(excelColumn, selectedSqlColumn);
                        }
                    }
                }

                return mappings;
            }

        // Inserts data into the specified SQL Server table
            private void InsertData(List<Dictionary<string, string>> dataToBeInserted)
            {
                string selectedTable = ViewState["selectedTable"].ToString();
                string connectionString = ViewState["connectionString"].ToString();

                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    try
                    {
                        foreach (var row in dataToBeInserted)
                        {
                            // Build SQL query for insertion
                            string columns = string.Join(", ", row.Keys);
                            string values = string.Join(", ", row.Values.Select(v => $"'{v}'"));

                            string query = $"INSERT INTO {selectedTable} ({columns}) VALUES ({values})";

                            using (SqlCommand command = new SqlCommand(query, connection))
                            {
                                command.ExecuteNonQuery(); // Execute insert query
                            }
                        }
                    }
                    catch
                    {
                        // Rethrow the exception to be handled in the calling method
                        throw;
                    }
                }
            }

        protected void InsertAgainButton_Click(object sender, EventArgs e)
        {
            SuccessLabel.Visible = false;
            InsertAgainButton.Visible = false;

            ExcelFileUpload.Visible = true;
            UploadButton.Visible = true;
        }

        //bigint
        //binary(50)
        //bit
        //char(10)
        //date
        //datetime
        //datetime2(7)
        //datetimeoffset(7)
        //decimal(18, 0)
        //float
        //geography
        //geometry
        //hierarchyid
        //image
        //int
        //money
        //nchar(10)
        //ntext
        //numeric(18, 0)
        //nvarchar(50)
        //nvarchar(MAX)
        //real
        //smalldatetime
        //smallint
        //smallmoney
        //sql_variant
        //text
        //time(7)
        //timestamp
        //tinyint
        //uniqueidentifier
        //varbinary(50)
        //varbinary(MAX)
        //varchar(50)
        //varchar(MAX)
        //xml
    }
}