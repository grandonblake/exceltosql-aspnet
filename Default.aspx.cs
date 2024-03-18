using System;
using System.Collections.Generic;
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
        }
        protected void UploadButton_Click(object sender, EventArgs e)
        {
            if (ExcelFileUpload.HasFile)
            {
                string filePath = Server.MapPath("~/Files/") + ExcelFileUpload.FileName;
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

                        List<string> excelColumns = new List<string>();
                        for (int col = 1; col <= colCount; col++)
                        {
                            excelColumns.Add(worksheet.Cells[1, col].Value?.ToString() ?? "");
                        }

                        // Database Logic
                        List<string> sqlColumns = GetSqlTableColumns();

                        // Pass data to the view
                        ViewState["excelColumns"] = excelColumns;
                        ViewState["sqlColumns"] = sqlColumns;

                        // Call a function to regenerate the form with dynamic elements
                        GenerateMappingForm();

                        /*for (int row = 1; row <= rowCount; row++)
                        {
                            for (int col = 1; col <= colCount; col++)
                            {
                                Response.Write(worksheet.Cells[row, col].Value.ToString() + " ");
                            }
                            Response.Write("<br/>");
                        }*/

                        package.Dispose(); // Dispose the package after use
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
                sqlDropDown.ID = "sqlColumnMapping_" + i;

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

        protected void Button1_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("Asdadsadadada");
            

        }
    }
}