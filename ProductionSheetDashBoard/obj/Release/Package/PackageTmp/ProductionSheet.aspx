<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="ProductionSheet.aspx.cs" Inherits="ProductionSheetDashBoard.ProductionSheet" %>


<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Production Sheet</title>
    <link href="https://maxcdn.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css" rel="stylesheet">
</head>

<body>
    <form id="form1" runat="server">
        <div class="container">
            <div class="row mt-3">
                <div class="col-md-12">
                    <h2>Production Sheet Generation</h2>
                </div>
            </div>
            <div class="row">
                <div class="col-md-6">
                    <label for="ddlStateName" class="form-label">Select State</label>
                    <asp:DropDownList ID="ddlStateName" runat="server" CssClass="form-control" DataTextField="HSRPStateName" DataValueField="HSRP_StateID"></asp:DropDownList>
                </div>
                <div class="col-md-6">
                    <label for="txtDate" class="form-label">EC Name</label>
                    <asp:TextBox ID="txtDate" runat="server" CssClass="form-control" autocomplete="off" EnableViewState="False" Visible="true"></asp:TextBox>
                </div>
            </div>
            
            <div class="row mt-3">
               <%-- <div class="col-md-6">
                    <span style="margin-bottom:20px;font-weight:bold;display:block;" >OEM Production Sheet</span>
                    <asp:Button ID="btnPreview" runat="server" CssClass="btn btn-success btn-block" Text="Generate Production Sheet" OnClick="Search_Click" Width="100%" />
                    <asp:Button ID="btnHero" runat="server" CssClass="btn btn-success btn-block" Text="Hero Production Sheet" OnClick="btnHero_Click" Width="100%" />
                    <asp:Button ID="BtnMulti" runat="server" CssClass="btn btn-success btn-block" Text="Multi Brand" OnClick="BtnMulti_Click" Width="100%" />
                    <asp:Button ID="btnHRPS" runat="server" CssClass="btn btn-success btn-block" Text="HR MultiBrand" Visible="true" OnClick="btnHRMultiBrand_Click" Width="100%" />
                    <asp:Button ID="btnOLA" runat="server" CssClass="btn btn-success btn-block" Text="OLA Production Sheet" OnClick="btnOLA_Click" Width="100%" />
                    <asp:Button ID="btnTwenty" runat="server" CssClass="btn btn-success btn-block" Text="Twenty Two Motors" OnClick="btnTwenty_Click" Width="100%" />
                    <asp:Button ID="btnPCA" runat="server" CssClass="btn btn-success btn-block" Text="PCA InterState" Visible="true" OnClick="btnPCA_Click" Width="100%" />
                </div>--%>
                <div class="col-md-9">
                    <%--<span style="margin-bottom:20px;font-weight:bold;display:block;"> BMHSRP Production Sheet</span>--%>
                    <asp:Button ID="btnBMHSRPPSSheet" runat="server" CssClass="btn btn-success btn-block" Text="BMHSRP Production Sheet" OnClick="btnBMHSRPPSSheet_Click" Width="50%" />
                    <asp:Button ID="btnDelivery" runat="server" CssClass="btn btn-success btn-block" Text="BookMYHSRP Home Delivery PS With Frames" OnClick="btnDelivery_Click" Width="50%" />
                    <asp:Button ID="btnDeliveryWFrame" runat="server" CssClass="btn btn-success btn-block" Text="BookMYHSRP Home Delivery PS Without Frames" OnClick="btnDeliveryWFrame_Click" Width="50%" />
                </div>
            </div>
            <div class="row">
                <div class="col-md-12">
                    <asp:Label ID="Label1" runat="server" Text=""></asp:Label>
                    <asp:Label ID="lblErrMess" ForeColor="Red" Font-Size="18px" runat="server"></asp:Label>
                </div>
            </div>
        </div>
    </form>
</body>

</html>
