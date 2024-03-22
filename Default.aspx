<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="exceltosql.Default" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePartialRendering="false" /> 
        <%---Setting EnablePartialRendering="false" in your ScriptManager control forces a full page 
            postback instead of an AJAX partial update within the UpdatePanel. 
            This bypasses the parsing step that ASP.NET's AJAX framework would typically perform ---%>

        <asp:UpdatePanel ID="UpdatePanel1" runat="server">
             <ContentTemplate>
                 <div>
                    <asp:FileUpload ID="ExcelFileUpload" runat="server" accept=".xls,.xlsx,.csv" />
                    <asp:Button ID="UploadButton" runat="server" Text="Upload" OnClick="UploadButton_Click" />
                    <br />
                    <br />
                    <asp:DropDownList ID="WorksheetList" runat="server" Visible="false"/>
                    <asp:Button ID="SelectWorksheetButton" runat="server" Text="Select Excel Worksheet" OnClick="SelectWorksheetButton_Click" Visible="false"/>
                    <br />
                    <br />
                    <asp:DropDownList ID="TableList" runat="server" Visible="false"/>
                    <asp:Button ID="SelectTableButton" runat="server" Text="Select SQL Table" OnClick="SelectTableButton_Click" Visible="false"/>
                    <br />
                    <br />
                 </div>
                 <div id="mappingContainer" runat="server"></div>
              </ContentTemplate>
        </asp:UpdatePanel>
        <br />
        <asp:Button ID="SubmitButton" runat="server" Text="Submit" OnClick="SubmitButton_Click" Visible="false" />

    </form>
</body>
</html>
