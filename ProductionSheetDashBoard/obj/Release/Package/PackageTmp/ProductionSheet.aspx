<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="ProductionSheet.aspx.cs" Inherits="ProductionSheetDashBoard.ProductionSheet" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<%--<head runat="server">
    <title></title>

 <script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css"/>
  <link rel="stylesheet" href="/resources/demos/style.css"/>

  <script src="https://code.jquery.com/jquery-1.12.4.js"></script>
    <script src="Scripts/jquery-3.4.1.js"></script>
    <script src="Scripts/jquery-ui-1.12.1.js"></script>
   
   <script>
       $(function () {
           $("#txtDate").datepicker({ dateFormat: 'dd/mm/yy' }).val();
       });
   </script>
    <style type="text/css">
        .button {}
        .auto-style1 {
            height: 29px;
        }
    </style>
</head>--%>
<body>
    <form id="form1" runat="server">
    <div>
       &nbsp;&nbsp;&nbsp;
       <table>
                    <tr>
                        <td><h2 style="color:#000000">Production Sheet Generation</h2></td>
                    </tr>
                    <tr>
                       <%-- <td>
                             <asp:Label ID="Label3" runat="server" Font-Names="Arial" Font-Size="10pt" Text="Select OEM"   style="margin-right:20px;"/>
                             <asp:DropDownList runat="server" Font-Names="Arial" Font-Size="10pt"  TabIndex="4" ID="ddlOemType" >
                                 <asp:ListItem Value="All">All </asp:ListItem>
                                 <asp:ListItem Value="TVS">TVS</asp:ListItem>
                                 <asp:ListItem Value="JCB">JCB</asp:ListItem>
                             </asp:DropDownList>
                                
                      
                        </td> --%>
                        <td class="auto-style1">
                             <asp:Label ID="Label2" runat="server" Font-Names="Arial" Font-Size="10pt" Text="Select State"   style="margin-right:20px;"/>
                             <asp:DropDownList TabIndex="5" ID="ddlStateName"  
                                runat="server" DataTextField="HSRPStateName" DataValueField="HSRP_StateID" >
                            <%--   <asp:ListItem Value="37">DL</asp:ListItem>--%>
                                
                            </asp:DropDownList>
                        </td> 
                        <td class="auto-style1">
                         
                             
                        <%--    Appointment Date--%>

                            EC Name:

                        </td> 
                        <td class="auto-style1">
                          
<asp:TextBox ID="txtDate" runat="server"  autocomplete="off" EnableViewState="False" Visible="true"></asp:TextBox>
                            </td>
                    </td>

                      
                        <td class="auto-style1">
                          
                        </td>
                         <td class="auto-style1">
                          
                        </td>
                    </tr>

                    <tr>
                       
                        <td>
                            
                            &nbsp;</td> 
                        <td>
                         
                             
                        </td> 
                        <td style="padding-left:40px;" >
                      
                             
                            
                            
                        </td>  
                      
                        <td>
                          
                        </td>
                    </tr>

            <tr>
                       
                        <td>
                            
                        </td> 
                        <td>
                         
                             
                        </td> 
                        <td style="padding-left:40px;" >
                      
                             
                             

                        
                            
                        &nbsp;
                      
                             
                              

                        
                            
                        </td>  
                      
                        <td>
                          
                        </td>
                    </tr>
             </table>
        <table>
                    <tr>

                        <td colspan="8" style="padding-left:40px;"">
                           &nbsp;
<asp:Button ID="btnPreview" runat="server" CssClass="button" BackColor="Green" ForeColor="#ffffff"
                                Text="Generate Production Sheet" OnClick="Search_Click" Width="180px" />
                             <asp:Button ID="btnHero" runat="server" CssClass="button" BackColor="Green" ForeColor="#ffffff"
                                Text="Hero Production Sheet" Width="230px" OnClick="btnHero_Click" />

     


                             <asp:Button ID="btnBMHSRPPSSheet" runat="server" CssClass="button" BackColor="Green" ForeColor="#ffffff"
                                Text="BMHSRP Production Sheet" Width="230px" OnClick="btnBMHSRPPSSheet_Click" />

                                  <asp:Button ID="btnDelivery" runat="server" CssClass="button" BackColor="Green" ForeColor="#ffffff"
                                Text=" BookMYHSRP Home Delivery PS With Frames "   Width="300px" OnClick="btnDelivery_Click" />

                            <asp:Button ID="BtnMulti" runat="server" CssClass="button" BackColor="Green" ForeColor="#ffffff"
                                Text="Multi Brand" OnClick="BtnMulti_Click" />

                            <asp:Button ID="btnHRPS" runat="server" CssClass="button" BackColor="Green" ForeColor="#ffffff"
                                Text="HR MultiBrand"  Visible="true" OnClick="btnHRMultiBrand_Click"   />

                        <asp:Button ID="btnDeliveryWFrame" runat="server" CssClass="button" BackColor="Green" ForeColor="#ffffff"
                                Text=" BookMYHSRP Home Delivery PS Without Frames "   Width="320px" OnClick="btnDeliveryWFrame_Click" />


                            <asp:Button ID="btnOLA" runat="server" CssClass="button" BackColor="Green" ForeColor="#ffffff"
                                Text="OLA Production Sheet"  Width="220px" OnClick="btnOLA_Click"  />

                                  <asp:Button ID="btnTwenty" runat="server" CssClass="button" BackColor="Green" ForeColor="#ffffff"
                                Text="Twenty Two Motors"  Width="220px" OnClick="btnTwenty_Click"   />

                                <asp:Button ID="btnPCA" runat="server" CssClass="button" BackColor="Green" ForeColor="#ffffff"
                                Text="PCA InterState"  Width="220px" Visible="true" OnClick="btnPCA_Click"   />
                             
                             
                           <%--   <asp:Button ID="btnRenault" runat="server" CssClass="button" BackColor="Green" ForeColor="#ffffff"
                                Text="Generate Production Sheet Renault Kerala" OnClick="btnRenault_Click" Width="270px" />
                            <asp:Button ID="btnTVSProductionsheet" runat="server" CssClass="button" BackColor="Green" ForeColor="#ffffff"
                                Text="TVS Production Sheet" OnClick="btnTVS_Click"  Width="180px" />


                               <asp:Button ID="btnJCBProductionSheet" runat="server" CssClass="button" BackColor="Green" ForeColor="#ffffff"
                                Text="JCB Production Sheet" OnClick="btnJCB_Click"  Width="220px" />

                           
                            

                            
                             <asp:Button ID="btnExternal" runat="server" CssClass="button" BackColor="Green" ForeColor="#ffffff"
                                Text="External Production Sheet"  Width="220px" OnClick="btnExternal_Click" />
                              <asp:Button ID="btnHero" runat="server" CssClass="button" BackColor="Green" ForeColor="#ffffff"
                                Text="Hero Production Sheet" OnClick="btnHero_Click"  Width="220px" />--%>

                        
 
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <asp:Label ID="Label1" runat="server" Text=""></asp:Label>
                            <asp:Label ID="lblErrMess" ForeColor="Red" Font-Size="18px" runat="server"></asp:Label>
                        </td> 
                    </tr>
                </table>
    </div>
    </form>
</body>
</html>
