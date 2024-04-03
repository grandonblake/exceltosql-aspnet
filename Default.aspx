<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="exceltosql.Default" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <style>
        body {
            font-family: "Times New Roman", Times, serif;
            text-align: center;
        }
        h1 {
            font-size: 36px;
            font-weight: bold;
            margin-bottom: 10px;
        }
        h2 {
            font-size: 14px;
            margin-bottom: 50px;
        }
        .upload-area {
            display: flex;
            align-items: center;
            margin-bottom: 10px;
            justify-content: center;
        }
        
    </style>
</head>
<body>
    <h1>Excel to SQL</h1>
    <h2>Insert your Excel data to your SQL database by mapping your Excel columns and SQL database's table's columns</h2>
    <form id="form1" runat="server">
        <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePartialRendering="false" /> 
        <%---Setting EnablePartialRendering="false" in your ScriptManager control forces a full page 
            postback instead of an AJAX partial update within the UpdatePanel. 
            This bypasses the parsing step that ASP.NET's AJAX framework would typically perform ---%>

        <asp:UpdatePanel ID="UpdatePanel1" runat="server">
             <ContentTemplate>
                 <div>
                    <div class="upload-area">
                        <asp:FileUpload ID="ExcelFileUpload" runat="server" accept=".xls,.xlsx,.csv" />
                        <asp:DropDownList ID="WorksheetList" runat="server" Visible="false"/>
                        <label for="SqlServerName" runat="server" id="SqlServerNameLabel" visible="false">Input SQL Server Name: </label>
                        <asp:TextBox ID="SqlServerName" runat="server" Visible="false"></asp:TextBox>
                        <asp:DropDownList ID="TableList" runat="server" Visible="false"/>
                        <div id="mappingContainer" runat="server"></div>
                        <asp:Label ID="SuccessLabel" runat="server" Text="Insert Successful!" Visible="false" ForeColor="Green"></asp:Label>
                    </div>
                     <div class="upload-area">
                        <label for="DatabaseName" runat="server" id="DatabaseNameLabel" visible="false">Input Database Name: </label>
                        <asp:TextBox ID="DatabaseName" runat="server" Visible="false"></asp:TextBox>
                    </div>
                    <div class="upload-area">
                        <asp:Button ID="UploadButton" runat="server" Text="Upload" OnClick="UploadButton_Click" />
                        <asp:Button ID="SelectWorksheetButton" runat="server" Text="Select Excel Worksheet" OnClick="SelectWorksheetButton_Click" Visible="false"/>               
                        <asp:Button ID="ConnectDatabase" runat="server" Text="Connect Database" OnClick="ConnectDatabaseButton_Click" Visible="false"/>
                        <asp:Button ID="SelectTableButton" runat="server" Text="Select SQL Table" OnClick="SelectTableButton_Click" Visible="false"/>
                        <asp:Button ID="InsertDataButton" runat="server" Text="Insert Data" OnClick="InsertDataButton_Click" Visible="false" />
                        <asp:Button ID="InsertAgainButton" runat="server" Text="Insert Again?" OnClick="InsertAgainButton_Click" Visible="false" />
                    </div>
                    <br />
                    <asp:Label ID="ErrorLabel" runat="server" Visible="false" ForeColor="Red"></asp:Label>
                    
                    <br />
                 </div>
                 
              </ContentTemplate>
        </asp:UpdatePanel>
        <br />        
    </form>
</body>
</html>
