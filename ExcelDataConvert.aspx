<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="ExcelDataConvert.aspx.cs" Inherits="ExcelFileToTableConvert.ExcelDataConvert" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>ExcelFileToTableConvert</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC" crossorigin="anonymous" />
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
    <!-- Include jQuery -->
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <!-- Include Footable CSS -->
    <link href="https://cdn.jsdelivr.net/npm/footable/css/footable.core.min.css" rel="stylesheet">
    <!-- Include Footable JavaScript -->
    <script src="https://cdn.jsdelivr.net/npm/footable/dist/footable.all.min.js"></script>
</head>
<body>
    <form id="form1" runat="server">
        <%--Section one header--%>
    <section class="py-3 header bg-dark">
        <div class="container text-white">
            <div class="row">
                <p class="text-center fs-4 fw-bold">Excel File To Table Convert</p>
            </div>
        </div>
    </section>
        <%--section two content--%>
    <section class="py-5 body">
        <div class="container">
            <div class="row fs-5 mx-5 text-dark justify-content-center border border-1 border-dark p-5">
                <div class="col-3 d-flex justify-content-center align-items-center">
                    <%--text--%>
                    <p class="fs-5 fw-normal text-dark m-0">Upload Excel File</p>
                </div>
                <div class="col-5 d-flex justify-content-center align-items-center">
                    <%--file uploder--%>
                    <asp:FileUpload ID="uploadFile" runat="server"/>
                </div>
                <div class="col-1 d-flex justify-content-center align-items-center">
                    <%--upload button--%>
                    <asp:Button CssClass="btn btn-primary" ID="btnUpload" runat="server" Text="Upload" OnClick="btnUpload_Click"/>
                </div>
                <div class="col-1 d-flex justify-content-center align-items-center">
                    <%--view button--%>
                    <asp:Button CssClass="btn btn-outline-success" ID="btnView" runat="server" Text="View" OnClick="btnView_Click" />
                </div>
            </div>
            <div class="row mx-5 justify-content-end">
                <%--note msg--%>
                <p class="fs-6 fw-light text-end">*This page is under Development Try Uploading Only Excel Files</p>
            </div>
        </div>
    </section>
        <%--section two grid--%>
    <section class="py-2 pb-5 body-2">
        <div class="container mb-3">
            <%--repeater--%>
            <asp:Repeater ID="repeaterView" runat="server">
                <%--header--%>
                <HeaderTemplate>
                    <%--table--%> 
                <table id="tableView" border="1" class="table table-dark table-striped">
                    <thead>
                         <tr>
                            <th>Serial No</th>
                            <th>User ID</th>
                            <th>Name</th>
                            <th data-breakpoints="xs sm">Email</th>
                            <th data-breakpoints="xs sm">Phone number</th>
                         </tr>
                        </thead>
                    <tbody>
                     </HeaderTemplate>
                <%--Table Column--%> 
                     <ItemTemplate>
                         <tr>
                            <td><%# Eval("srno") %></td>
                            <td><%# Eval("userid") %></td>
                            <td><%# Eval("name") %></td>
                            <td><%# Eval("email") %></td>
                            <td><%# Eval("phone") %></td>
                         </tr>
                     </ItemTemplate>
                <%--footer template--%> 
                    <FooterTemplate>
                        </tbody>
                        <tfoot>
                            <tr>
                                <td colspan="4">
                                    <div class="pagination pagination-centered d-flex"></div>
                                </td>
                            </tr>
                        </tfoot>
                        </table>
                    </FooterTemplate>
            </asp:Repeater>
        </div>
    </section>
        <%--section for footer--%>
    <section class="py-2 bg-dark footer fixed-bottom">
        <div class="container text-white">
            <div class="row">
                <p class="text-center fs-5 fw-normal">Sweathkumar Project 2024</p>
            </div>
        </div>
    </section>
    </form>
    <%--script--%> 
    <script>
        <%--footables
        $(document).ready(function () {
            $('#tableView').footable();
        });--%>
    </script>
</body>
</html>
