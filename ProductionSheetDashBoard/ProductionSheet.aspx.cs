using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using OpenHtmlToPdf;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Text;


namespace ProductionSheetDashBoard
{
    public partial class ProductionSheet : System.Web.UI.Page
    {
        //Test
        string filePrefix = string.Empty;
        string strCompanyName = string.Empty;
        string SQLString = string.Empty;
        string userType = string.Empty;
        string CnnString = System.Configuration.ConfigurationManager.ConnectionStrings["ConnectionString"].ToString();
        string dirPath = ConfigurationSettings.AppSettings["PdfProduction"].ToString();
        string Hsrp_stateid, Navembid = "";
        string EmbCenterName = "";
        string userId = "";
        protected void Page_Load(object sender, EventArgs e)
        {

            //Remove MH  and Back Button in MHHSRP Production sheet

            if (Request.QueryString["UID"] != null)
            {
                userId = Request.QueryString["UID"];
            }

            //if (Request.QueryString["UID"] == null)
            //{
            //    userId = "32576";


            //}


            //if (Request.QueryString["UID"] == null)
            //{
            //    userId = "31971";
            //}

            //if (Request.QueryString["UID"] == null)
            //{
            //    userId = "66552";//Request.QueryString["UID"];
            //}

            //Hero vida interstate Production sheet Changes
            if (!IsPostBack)
            {

                FillUserDetails();

                FillDDLState();


            }

            lblErrMess.Text = "";



        }

        //Changes in Victory production sheet changes
         private void FillUserDetails()

        {
           //string sqlQuery = "select top 1 U.HSRP_stateid,Navembid ,EmbCenterName from users u  join RTOlocation r on U.RTOlocationid=r.RTOlocationid where u.UserID  = '26127'";
            string sqlQuery = "select top 1 U.HSRP_stateid,Navembid ,EmbCenterName,U.UserType from users u  join RTOlocation r on U.RTOlocationid=r.RTOlocationid where u.UserID  = '" + userId+"'";
            //string sqlQuery = "select top 1 U.HSRP_stateid,Navembid ,EmbCenterName from users u  join RTOlocation r on U.RTOlocationid=r.RTOlocationid where u.UserID  = '53164' ";



            DataTable dt = Utils.Utils.GetDataTable(sqlQuery, CnnString);
            if (dt.Rows.Count > 0)
            {
                Hsrp_stateid = dt.Rows[0]["HSRP_stateid"].ToString();
                Navembid = dt.Rows[0]["Navembid"].ToString();
                EmbCenterName = dt.Rows[0]["EmbCenterName"].ToString();
                txtDate.Text = dt.Rows[0]["EmbCenterName"].ToString();
                userType= dt.Rows[0]["UserType"].ToString();

            }

        }
        private void FillDDLState()
        {

            //if (userType == "0")
            //{



            //    string sqlQuery = "select HSRP_StateID, HSRPStateName from hsrpstate where HSRP_StateID ='" + Hsrp_stateid + "' and ActiveStatus='Y'  order by HSRPStateName";
            //    DataTable dtState = Utils.Utils.GetDataTable(sqlQuery, CnnString);
            //    ddlStateName.DataSource = dtState;
            //    ddlStateName.DataBind();
            //    ddlStateName.Items.Insert(0, new ListItem("--Select State--", "0"));

            //}

            //else if (userType == "4")
            //{

            //    // string sqlQuery = "select HSRP_StateID, (select  HSRPStateName  from hsrpstate h where h.HSRP_StateID =hsrp_stateid)  and ActiveStatus='Y'  order by HSRPStateName";


            //    string sqlQuery = "select distinct H.HSRP_StateID,  H.HSRPStateName from  UserStateMapping U  join HSRPstate H on U.hsrp_stateid = H.hsrp_stateid  where mapstatus = 'Y'  and USerid = '" + userId + "' order by HSRPStateName";
            //     DataTable dtState = Utils.Utils.GetDataTable(sqlQuery, CnnString);
            //    ddlStateName.DataSource = dtState;
            //    ddlStateName.DataBind();
            //    ddlStateName.Items.Insert(0, new ListItem("--Select State--", "0"));
            //}

            //else
            //{

                string sqlQuery = "select HSRP_StateID, HSRPStateName from hsrpstate where HSRP_StateID ='" + Hsrp_stateid + "' and ActiveStatus='Y'  order by HSRPStateName";
                DataTable dtState = Utils.Utils.GetDataTable(sqlQuery, CnnString);
                ddlStateName.DataSource = dtState;
                ddlStateName.DataBind();
                ddlStateName.Items.Insert(0, new ListItem("--Select State--", "0"));


            //}



        }

        //private void FilldropDownListOrganization()
        //{
        //    if (UserType == "0")
        //    {
        //        SQLString = "select HSRPStateName,HSRP_StateID from HSRPState  where ActiveStatus='Y' Order by HSRPStateName";
        //        Utils.PopulateDropDownList(DropDownListStateName, SQLString.ToString(), CnnString, "--Select State--");
        //        // DropDownListStateName.SelectedIndex = DropDownListStateName.Items.Count - 1;
        //    }

        //    else if (UserType == "4")
        //    {
        //        SQLString = "select HSRPStateName,a.HSRP_StateID from HSRPState a , UserStateMapping b  where a.hsrp_stateid=b.hsrp_stateid and b.userid='" + strUserID.ToString() + "' and ActiveStatus='Y' Order by HSRPStateName";
        //        Utils.PopulateDropDownList(DropDownListStateName, SQLString.ToString(), CnnString, "--Select State--");
        //        // DropDownListStateName.SelectedIndex = DropDownListStateName.Items.Count - 1;

        //    }
        //    else
        //    {
        //        SQLString = "select HSRPStateName,HSRP_StateID from HSRPState  where HSRP_StateID=" + HSRPStateID + " and ActiveStatus='Y' Order by HSRPStateName";
        //        DataTable dts = Utils.GetDataTable(SQLString, CnnString);
        //        DropDownListStateName.DataSource = dts;
        //        DropDownListStateName.DataBind();
        //    }
        //}

        protected void Search_Click(object sender, EventArgs e)
        {
            //New Format production sheet Changes for All state 
            //  string Nave    Session["Navembid"]

            if (ddlStateName.SelectedItem.Text == "--Select State--")
            {
                lblErrMess.Text = "Please select  State";
                return;
            }
           

            string sqlPrefixQuery = "select top 1 CONVERT(INT, isnull(OrderNo,'0')) + 1 as orderNo from ProductionSheetAutoGenerated_List with(nolock) order by id desc";
            DataTable dtPrefix = Utils.Utils.GetDataTable(sqlPrefixQuery, CnnString);

            if (dtPrefix.Rows.Count > 0)
            {
                filePrefix = dtPrefix.Rows[0]["orderNo"].ToString();
            }
            else
            {
                filePrefix = "1";
            }

            //txtFilePrefix.Text = filePrefix;

            if (filePrefix.Length > 0)
            {

                SheetGeneration();

                btnRenault_Click(sender, e);
                btnSuzu_Click(sender, e);
                VICTORY_ELECTRIC(sender, e);
                btnTVS_Click(sender, e);
                btnJCB_Click(sender, e);
                btnExternal_Click(sender, e);




                //SheetGeneration();
                //SheetGenerationHero();
                //SheetGenerationRenault();
                //SheetGenerationTVS();
                //SheetGenerationJCB();
                //SheetGenerationExternal();


            }



        }

        protected void VICTORY_ELECTRIC(object sender, EventArgs e)
        {

            if (ddlStateName.SelectedItem.Text == "--Select State--")
            {
                lblErrMess.Text = "Please select  State";
                return;
            }


            string sqlPrefixQuery = "select top 1 CONVERT(INT, isnull(OrderNo,'0')) + 1 as orderNo from ProductionSheetAutoGenerated_List order by id desc";
            DataTable dtPrefix = Utils.Utils.GetDataTable(sqlPrefixQuery, CnnString);

            if (dtPrefix.Rows.Count > 0)
            {
                filePrefix = dtPrefix.Rows[0]["orderNo"].ToString();
            }
            else
            {
                filePrefix = "1";
            }

            //txtFilePrefix.Text = filePrefix;

            if (filePrefix.Length > 0)
            {

                VICTORY_ELECTRIC();


            }

        }

        /// <summary>
        /// Author:  Rabindra Pradhan
        /// Date:    09-08-2023
        /// Add Victory Electric Part to generate Production sheet
        /// </summary>
        private void VICTORY_ELECTRIC()
        {
            string stateECQuery = string.Empty;
            string ReqNum = string.Empty;
            //string Navembid = Session["Navembid"].ToString();
            FillUserDetails();


            // stateECQuery = "select distinct a.HSRP_StateID, (select HSRPStateName from hsrpstate c where c.hsrp_stateid=a.hsrp_stateid) as HSRPStateName,(select HSRPStateShortName from hsrpstate d where d.hsrp_stateid=a.hsrp_stateid) as HSRPStateShortName, navembcode from hsrprecords a,rtolocation b where a.HSRP_StateID ='" + ddlStateName.SelectedValue + "' and   b.Navembcode not like '%CODO%'    and a.rtolocationid=b.rtolocationid and isnull(NewPdfRunningNo,'') = '' and isnull(erpassigndate,'') != ''  order by  a.HSRP_StateID";

            stateECQuery = "select distinct a.HSRP_StateID, (select HSRPStateName from hsrpstate c where c.hsrp_stateid=a.hsrp_stateid) as HSRPStateName,(select HSRPStateShortName from hsrpstate d where d.hsrp_stateid=a.hsrp_stateid) as HSRPStateShortName, d.navembcode from hsrprecords a  with(nolock) join DealerAffixation d with(nolock)  on  a.Affix_Id=d.SubDealerId and NewPdfRunningNo is null and  a.HSRP_StateID ='" + ddlStateName.SelectedValue + "' and d.Navembcode not like '%CODO%'  and erpassigndate is not null   AND d.navembcode='" + Navembid + "' order by  a.HSRP_StateID";


            DataTable dtSE = Utils.Utils.GetDataTable(stateECQuery, CnnString);
            if (dtSE.Rows.Count > 0)
            {
                foreach (DataRow dr in dtSE.Rows)
                {
                    #region

                    string HSRP_StateID = dr["HSRP_StateID"].ToString().Trim();
                    string HSRPStateName = dr["HSRPStateName"].ToString().Trim();
                    string HSRPStateShortName = dr["HSRPStateShortName"].ToString().Trim();
                    string Navembcode = dr["Navembcode"].ToString().Trim();

                    string dir = dirPath + DateTime.Now.ToString("yyyy-MM-dd") + "\\" + HSRPStateShortName + "\\";

                    string fileName = "VictoryACEEULERSAHNIANAND" + "-" + filePrefix + "-" + Navembcode + ".pdf";
                    string filePath = dir + fileName;



                    StringBuilder html = new StringBuilder();

                    Boolean findRecord = false;
                    string strProductionSheetNo = string.Empty;

                    /*
                     *  Start body & HTMl Tag
                     */
                    #region
                    html.Append(
                        "<!DOCTYPE html>" +
                        "<html>" +
                        "<head>" +
                            "<meta charset='UTF-8'><title>Title</title>" +
                            "<style>" +
                                "@page {" +
                                    /* headers*/
                                    "@top-left {" +
                                        "content: 'Left header';" +
                                    "}" +
                                    "@top-right {" +
                                        "content: 'Right header';" +
                                    "}" +

                                    /* footers */
                                    "@bottom-left {" +
                                        "content: 'Lorem ipsum';" +
                                    "} " +
                                    "@bottom-right {" +
                                        "content: 'Page ' counter(page) ' of ' counter(pages);" +
                                    "}" +
                                    "@bottom-center  {" +
                                        "content:element(footer);" +
                                    "}" +
                                "}" +
                                 "#footer {" +
                                    "position: running(footer);" +
                                "}" +
                                "table {" +
                                  "border-collapse: collapse;" +
                                "}" +

                                "table, th, td {" +
                                    "border: 1px solid black;" +
                                    "text-align: left;" +
                                    "vertical-align: top;" +
                                    "padding:5px;" +
                                "}" +
                            "</style>" +
                        "</head>" +
                        "<body>");
                    #endregion

                    string oemDealerQuery = string.Empty;




                    oemDealerQuery = "select distinct om.oemid as oemid, om.name as oemname, dm.dealerid, d.Subdealername as Dealername, dm.dealercode, d.Address,d.SubDealerId, dm.HSRP_StateID, " +
                        "dm.RTOLocationID from oemmaster om with(nolock) " +
                        "left join dealermaster dm with(nolock) on dm.oemid = om.oemid join DealerAffixation d with(nolock) on d.DealerID=dm.DealerId where D.STATE_ID =" + HSRP_StateID + " and   d.Navembcode='" + Navembcode + "' and " +

                        "dm.dealerid in (select distinct dealerid from hsrprecords where NewPdfRunningNo is null and erpassigndate is not null and   affix_id is not null ) and   (( OM.OEMid=1367) or (d.dealerid=45012) or (d.dealerid=92724) or (d.dealerid=94991))";


                    #region
                    DataTable dtOD = Utils.Utils.GetDataTable(oemDealerQuery, CnnString);
                    if (dtOD.Rows.Count > 0)
                    {
                        foreach (DataRow drOD in dtOD.Rows)
                        {

                            string oemid = drOD["oemid"].ToString().Trim();
                            string dealerid = drOD["dealerid"].ToString().Trim();
                            string oemname = drOD["oemname"].ToString().Trim();
                            string dealername = drOD["dealername"].ToString().Trim();
                            string Address = drOD["Address"].ToString().Trim();
                            string strsubaffixid = drOD["SubDealerId"].ToString();

                            //start sql query
                            #region
                            string productionQuery = string.Empty;


                            DataTable dtProduction = new DataTable();




                            #endregion
                            //end sql query

                            #region

                            SqlConnection con = new SqlConnection(CnnString);
                            SqlCommand cmd = new SqlCommand("USP_VictoryProductionSheet", con);
                            cmd.CommandType = CommandType.StoredProcedure;
                            con.Open();
                            cmd.Parameters.AddWithValue("@navembid", Navembcode);
                            cmd.Parameters.AddWithValue("@HSRP_StateID", HSRP_StateID);
                            cmd.Parameters.AddWithValue("@Dealerid", dealerid);
                            cmd.Parameters.AddWithValue("@Affix_Id", strsubaffixid);
                            SqlDataAdapter da = new SqlDataAdapter(cmd);
                            // dtProduction = new DataTable();
                            da.Fill(dtProduction);
                            con.Close();


                            // DataTable dtProduction = Utils.Utils.GetDataTable(productionQuery, CnnString);
                            if (dtProduction.Rows.Count > 0)
                            {

                                findRecord = true;
                                string strRunningNo = string.Empty;
                                //string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) as maxSheetNo from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + NAVEMBID + "'";
                                string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) from EmbossingCentersNew with(nolock) where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                                string strCom = Utils.Utils.getScalarValue(strSel, CnnString);
                                //DataTable dtSheetNo = fillDataTable(strSel, CnnString);
                                if (strCom.Equals(0) || strCom.Length == 0)
                                {
                                    strRunningNo = "0000001";
                                }
                                else
                                {
                                    strRunningNo = string.Format("{0:0000000}", Convert.ToInt32(strCom) + 1);
                                }
                                string strRequeNo = "select (prefixtext+right('00000'+ convert(varchar,Lastreqno+1),5)) as Reqno from prefix_Requisition with(nolock)  where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                                ReqNum = Utils.Utils.getScalarValue(strRequeNo, CnnString);
                                string strPRFIX = "select PrefixText from EmbossingCentersNew with(nolock) where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                                string strPRFIXCom = Utils.Utils.getScalarValue(strPRFIX, CnnString);
                                strProductionSheetNo = strPRFIXCom + strRunningNo;

                                string RTOLocationName = dtProduction.Rows[0]["RTOLocationName"].ToString();

                                #region

                                html.Append("<div style='page-break-before: avoid;page-break-inside: avoid;page-break-after: always;'>" +
                                        "<table style='width:100%;border: 0px;'>" +
                                            "<tr style='border: 0px;'>" +
                                                "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>Report Generation Date:</b> " + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</div>" +
                                                "</td>" +

                                                 "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>ROSMERTA SAFETY SYSTEMS LIMITED</b> " + "</div>" +
                                                "</td>" +


                                                 "<td colspan='5' style='border: 0px;'>" +
                                                    "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Production Sheet No:</b> " + strRunningNo + " </div>" +
                                                "</td>" +

                                            "</tr>" +

                                              "<tr style='border: 0px;'>" +
                                                "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>State:</b> " + HSRPStateName + "</div>" +
                                                "</td>" +

                                                 "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>Production Sheet</b> " + "</div>" +
                                                "</td>" +


                                                 "<td colspan='5' style='border: 0px;'>" +
                                                    "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Dealer ID/Name:</b> " + dealerid + "/" + dealername + "/" + strsubaffixid + " </div>" +
                                                "</td>" +

                                            "</tr>" +

                                                "<tr style='border: 0px;'>" +
                                                "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>EC Location: </b> " + RTOLocationName + "</div>" +
                                                "</td>" +

                                                 "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='style='text-align:left'><b>Oem :</b> " + oemname + "</div>" +
                                                "</td>" +


                                                 "<td colspan='5' style='border: 0px;'>" +
                                                     "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Dealer Address:</b> " + Address + " </div>" +
                                                //"<div style='text-align:left'><b>Dealer Address:</b> " + Address + " </div>" +
                                                "</td>" +

                                            "</tr>" +


  "<tr>" +
                                                "<td style='text-align:center'>Sr. No.</td>" +
                                                  "<td>Vehicle No.</td>" +
                                                   "<td>Front Plate Size</td>" +
                                                "<td>Front Laser No.</td>" +

                                                "<td>Rear Plate Size</td>" +
                                                "<td>Rear Laser No.</td>" +

                                                "<td>H. S. Foil </td>" +
                                                "<td>Caution Sticker</td>" +
                                                "<td>Fuel Type</td>" +
                                                "<td >VT</td>" +
                                                "<td >VC</ td>" +

                                            "</tr>");

                                #endregion

                                #region
                                foreach (DataRow drProduction in dtProduction.Rows)
                                {
                                    string HsrprecordID = drProduction["hsrprecordID"].ToString().Trim();
                                    string SRNo = drProduction["SRNo"].ToString().Trim();
                                    string ORD = "";// drProduction["ORD"].ToString().Trim();
                                    string VC = drProduction["VehicleClass"].ToString().Trim();
                                    string VehicleNo = drProduction["VehicleRegNo"].ToString().Trim();
                                    // string VehicleNo = "<td style='text-align:center;font-size:20px;'>" + "<b>" + drProduction["VehicleRegNo"].ToString().Trim() + "</b> </td>" +
                                    string VT = drProduction["VehicleType"].ToString().Trim();
                                    //string ChassisNo = drProduction["ChassisNo"].ToString().Trim();
                                    //string EngineNo = drProduction["EngineNo"].ToString().Trim();
                                    string FuelType = drProduction["FuelType"].ToString().Trim();
                                    string FrontPSize = drProduction["FrontProductCode"].ToString().Trim();
                                    // string FrontLaserNo = "<td style='text-align:center;font-size:20px;'>" + "<b>" + drProduction["HSRP_Front_LaserCode"].ToString().Trim() + "</b> </td>" +
                                    string FrontLaserNo = drProduction["HSRP_Front_LaserCode"].ToString().Trim();
                                    string RearPSize = drProduction["RearProductCode"].ToString().Trim();
                                    string RearLaserNo = drProduction["HSRP_Rear_LaserCode"].ToString().Trim();
                                    //string RearLaserNo = "<td style='text-align:center;font-size:20px;'>" + "<b>" + drProduction["HSRP_Rear_LaserCode"].ToString().Trim() + "</b> </td>" +
                                    string Amount = drProduction["roundoff_netamount"].ToString().Trim();
                                    //string OrderStatus = drProduction["OrderStatus"].ToString().Trim();
                                    string stickerColor = drProduction["stickerColor"].ToString().Trim();
                                    string HotStampingFoilColour = drProduction["HotStampingFoilColour"].ToString().Trim();

                                    html.Append("<tr>" +
                                   "<td style='text-align:center'>" + SRNo + "</td>" +
                                       "<td style='text-align:center;font-size:20px;'>" + "<b>" + VehicleNo + "</b> </td>" +

                                        "<td>" + FrontPSize + "</td>" +
                                         "<td style='text-align:center;font-size:20px;'>" + "<b>" + FrontLaserNo + "</b> </td>" +
                                           "<td>" + RearPSize + "</td>" +
                                         "<td style='text-align:center;font-size:20px;'>" + "<b>" + RearLaserNo + "</b> </td>" +

                                          "<td>" + HotStampingFoilColour + "</td>" +
                                           "<td>" + stickerColor + "</td>" +
                                            "<td>" + FuelType + "</td>" +
                                            "<td>" + VT + "</td>" +
                                            "<td>" + VC + "</td>" +





                               "</tr>");


                                    try
                                    {
                                        //start updating hsrprecords 
                                        string sqlUpdateHSRPRecords = "update hsrprecords set sendtoProductionStatus='Y', NAVPDFFlag='1', NewPdfRunningNo='" + strProductionSheetNo + "', Requisitionsheetno='" + ReqNum + "',  " +
                                              "PdfDownloadDate=GetDate(), pdfFileName='" + fileName + "', PDFDownloadUserID='1' where hsrprecordID='" + HsrprecordID + "' ";

                                        Utils.Utils.ExecNonQuery(sqlUpdateHSRPRecords, CnnString);   // uncomment after testing
                                    }
                                    catch (Exception ev)
                                    {
                                        Label1.Text = "hsrprecords update error: " + ev.Message;
                                    }
                                    //end 

                                }
                                #endregion
                                html.Append("</table>");

                                html.Append("</div>");

                                try
                                {
                                    string StrSqlUpdateECQuery = "update EmbossingCentersNew set NewProductionSheetRunningNo='" + strProductionSheetNo + "' " +
                                     "where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                                    Utils.Utils.ExecNonQuery(StrSqlUpdateECQuery, CnnString); // uncomment after testing
                                }
                                catch (Exception ev)
                                {
                                    Label1.Text = "EmbossingCentersNew update error: " + ev.Message;
                                }
                            }

                            #endregion

                        }
                    }// close oemDealerQuery
                    #endregion

                    #region "Req Generate"
                    string strComNew = string.Empty;
                    string strReqNumber = string.Empty;
                    string strReqNo = string.Empty;



                    string strSqlQuery1 = "select CompanyName from hsrpstate with(nolock) where hsrp_stateid='" + HSRP_StateID + "'";
                    strCompanyName = Utils.Utils.getScalarValue(strSqlQuery1, CnnString);

                    string strEMBName = " select EmbCenterName from EmbossingCentersNew with(nolock) where  State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                    strComNew = Utils.Utils.getScalarValue(strEMBName, CnnString);

                    if (strComNew != "")
                    {



                        strReqNo = "select (prefixtext+right('00000'+ convert(varchar,lastno+1),5)) as Reqno from prefix_Requisition with(nolock)  where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                        strReqNumber = Utils.Utils.getScalarValue(strReqNo, CnnString);

                        SQLString = "Exec [laserreqSlip1DashBoard]  '" + HSRP_StateID + "','" + Navembcode + "' ,  '" + ReqNum + "'";
                        DataTable dtResult = Utils.Utils.GetDataTable(SQLString, CnnString);

                        string strQuery = string.Empty;
                        string strRtoLocationName = string.Empty;
                        int Itotal = 0;

                        html.Append("<div style='width:100%;height:100%;'>" +
                                            "<table style='width:100%'>" +

                                                "<tr>" +
                                                    "<td colspan='12'>" +
                                                        "<div style='text-align:center;padding:8px;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + strCompanyName + "</b>" +

                                                            "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                             "<b style='font-size:20px;'>" + "MATERIAL REQUSITION NOTE" + "</b>" +

                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +

                                                "<tr>" +
                                                    "<td colspan='12'>" +
                                                        "<div style='text-align:center;padding:8px;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Production Sheet Date :" + DateTime.Now.ToString("dd-MM-yyyy") + "</b>" +
                                                            "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                             "<b style='font-size:20px;'>" + "" + "</b>" +
                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +
                                                "<tr>" +
                                                    "<td colspan='6'>" +
                                                        "<div style='text-align:left;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>REQ.NO:-</b>" +
                                                            "" + strReqNumber + "" +
                                                        "</div>" +
                                                    "</td>" +
                                                    "<td colspan='6'>" +
                                                        "<div style='text-align:left;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>Embossing Center:</b>" +
                                                            "" + strComNew + "" +
                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +
                                                "<tr>" +
                                                    "<td colspan='12'>" +
                                                        "<div style='text-align:left;padding:2px;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "State:" + HSRPStateName + "</b>" +
                                                            "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                             "<b style='font-size:20px;'>" + "" + "</b>" +
                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +
                                                 "<tr>" +
                                                    "<td colspan='12'>" +
                                                        "<div style='text-align:left;padding:8px;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + " " + "</b>" +
                                                            "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                             "<b style='font-size:20px;'>" + "" + "</b>" +
                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +

                                                "<tr>" +
                                                    "<td colspan='1' style='text-align:center'>SR.N.</td>" +
                                                    "<td colspan='3'>Product Size</td>" +
                                                    "<td colspan='1'>Laser Count</td>" +
                                                    "<td colspan='1'>Start Laser No</td>" +
                                                    "<td colspan='1'>End Laser No</td>" +

                                                "</tr>");


                        if (dtResult.Rows.Count > 0)
                        {
                            for (int i = 0; i < dtResult.Rows.Count; i++)
                            {
                                string ID = dtResult.Rows[i]["ID"].ToString();
                                string productcode = dtResult.Rows[i]["productcode"].ToString();
                                string LaserCount = dtResult.Rows[i]["LaserCount"].ToString();
                                Itotal = Convert.ToInt32(dtResult.Rows[i]["Total"].ToString());
                                string BeginLaser = dtResult.Rows[i]["BeginLaser"].ToString();
                                string EndLaser = dtResult.Rows[i]["EndLaser"].ToString();




                                html.Append("<tr>" +
                                   "<td colspan='1' style='text-align:left'>" + ID + "</td>" +
                                   "<td colspan='3'>" + productcode + "</td>" +

                                   "<td colspan='1'>" + LaserCount + "</td>" +
                                   "<td colspan='1'>" + BeginLaser + "</td>" +
                                   "<td colspan='1'>" + EndLaser + "</td>" +

                               "</tr>");
                            }
                        }
                        html.Append("<tr>" +
                                 "<td colspan='1' style='text-align:center' > " + "<b>Grand Total:</b>" + "</td>" +
                                 "<td colspan='3'>" + "" + "</td>" +

                                 "<td colspan='1'>" + Itotal + "</td>" +
                                 "<td colspan='1'>" + " " + "</td>" +
                                 "<td colspan='1'>" + " " + "</td>" +

                             "</tr>");




                        html.Append("<tr>" +
                         "<td colspan='2' > " + "<b>REQUESTED BY </b>" + "</td>" +
                         "<td colspan='2'>" + "<b>AUTHORIZED BY </b>" + "</td>" +

                         "<td colspan='2'>" + "<b>ISSUED BY </b>" + "</td>" +
                         "<td colspan='2'>" + "<b>RECEIVED BY</b>" + "</td>" +


                     "</tr>");

                        html.Append("<tr>" +
                                                  "<td colspan='12'>" +
                                                      "<div style='text-align:left;padding:2px;'>" +
                                                          "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Name" + "</b>" +

                                                      "</div>" +
                                                  "</td>" +
                                              "</tr>");

                        html.Append("<tr>" +
                                                 "<td colspan='12'>" +
                                                     "<div style='text-align:left;padding:2px;'>" +
                                                         "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Designation" + "</b>" +

                                                     "</div>" +
                                                 "</td>" +
                                             "</tr>");

                        html.Append("<tr>" +
                                               "<td colspan='12'>" +
                                                   "<div style='text-align:right;padding:8px;'>" +
                                                       "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Download By: Admin " + "</b>" +

                                                   "</div>" +
                                               "</td>" +
                                           "</tr>");

                        html.Append("<tr>" +
                                              "<td colspan='12'>" +
                                                  "<div style='text-align:right;padding:8px;'>" +
                                                      "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Sheet Generated By :" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</b>" +

                                                  "</div>" +
                                              "</td>" +
                                          "</tr>");







                        html.Append("</table>");

                        html.Append("</div>");

                        try
                        {
                            //start updating hsrprecords 
                            string Query = "update prefix_Requisition set lastno=lastno+1,Lastreqno=Lastreqno+1 where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                            Utils.Utils.ExecNonQuery(Query, CnnString);
                        }
                        catch (Exception ev)
                        {
                            Label1.Text = "prefix Requisition update error: " + ev.Message;
                        }
                    }


                    #endregion

                    /*
                     * Close body & HTMl Tag
                     */
                    html.Append("</body>" +
                        "</html>");


                    if (findRecord)
                    {

                        string strSaveSQlQuery = "insert into ProductionSheetAutoGenerated_List (HSRP_StateID, State_Code, RTO_ID, Emb_Center_Id, FileName, Productiondate, ProductionTime, OrderNo) values " +
                                    "('" + HSRP_StateID + "', '" + HSRPStateShortName + "', '', '" + Navembcode + "', '" + fileName + "', '" + DateTime.Now.ToString("yyyy-MM-dd") + "', '" + DateTime.Now.ToString("HH:mm:ss") + "', '" + filePrefix + "') ";
                        Utils.Utils.ExecNonQuery(strSaveSQlQuery, CnnString); // uncomment after testing
                        Label1.Text = Label1.Text + "successfully created production sheet of State: " + HSRPStateShortName + ", , Emb_Center_Id: " + Navembcode + Environment.NewLine;

                        //lblLog.Text = lblLog.Text + Environment.NewLine + "Start Production Query at:" + DateTime.Now;
                        #region
                        try
                        {
                            if (!Directory.Exists(dir))
                            {


                                Directory.CreateDirectory(dir);

                            }

                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        try
                        {
                            var pdf = Pdf
                                    .From(html.ToString())
                                    .OfSize(PaperSize.A4)
                                    .WithTitle("Title")
                                    .WithoutOutline()
                                    .WithMargins(1.25.Centimeters())
                                     .Landscape()
                                    .Comressed()
                                    .Content();

                            FileStream readStream = File.Create(filePath);
                            BinaryWriter binaryWriter = new BinaryWriter(readStream);

                            // Write the binary data to the file
                            binaryWriter.Write(pdf);
                            binaryWriter.Close();
                            readStream.Close();
                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        #endregion
                    }

                    #endregion

                }//close foreach stateEcQuery
            }
        }


        private void SheetGeneration()
        {
            string stateECQuery = string.Empty;

            //string Navembid = Session["Navembid"].ToString();
            FillUserDetails();
            string ReqNum = string.Empty;

            //if ((ddlStateName.SelectedValue == "20") || (ddlStateName.SelectedValue == "10"))
            //if ((ddlStateName.SelectedValue == "20") || (ddlStateName.SelectedValue == "10") || (ddlStateName.SelectedValue == "17") || (ddlStateName.SelectedValue == "14"))
            //{


                stateECQuery = "select distinct a.HSRP_StateID, (select HSRPStateName from hsrpstate c where c.hsrp_stateid=a.hsrp_stateid) as HSRPStateName,(select HSRPStateShortName from hsrpstate d where d.hsrp_stateid=a.hsrp_stateid) as HSRPStateShortName, navembcode from hsrprecords a with(nolock),rtolocation b where a.HSRP_StateID ='" + ddlStateName.SelectedValue + "' and   b.Navembcode not like '%CODO%'    and a.rtolocationid=b.rtolocationid and NewPdfRunningNo is null and erpassigndate is not null and OrderStatus='New Order' AND b.NAVEMBID='" + Navembid + "'   order by  a.HSRP_StateID";
            //}
            //else
            //{
            //    stateECQuery = "select distinct a.HSRP_StateID, (select HSRPStateName from hsrpstate c where c.hsrp_stateid=a.hsrp_stateid) as HSRPStateName,(select HSRPStateShortName from hsrpstate d where d.hsrp_stateid=a.hsrp_stateid) as HSRPStateShortName, navembcode from hsrprecords a with(nolock),rtolocation b where a.HSRP_StateID ='" + ddlStateName.SelectedValue + "' and   b.Navembcode not like '%CODO%'    and a.rtolocationid=b.rtolocationid and NewPdfRunningNo is null and erpassigndate is not null and OrderStatus='New Order' and   b.navembcode like  '%'+(select HSRPStateShortName from hsrpstate d where d.hsrp_stateid=a.hsrp_stateid)+'%'  and b.NAVEMBID='" + Navembid + "'    order by  a.HSRP_StateID";

            //   // stateECQuery = "select distinct a.HSRP_StateID, (select HSRPStateName from hsrpstate c where c.hsrp_stateid=a.hsrp_stateid) as HSRPStateName,(select HSRPStateShortName from hsrpstate d where d.hsrp_stateid=a.hsrp_stateid) as HSRPStateShortName, navembcode from hsrprecords a with(nolock),rtolocation b where a.HSRP_StateID ='" + ddlStateName.SelectedValue + "' and   b.Navembcode  like '%CODO%'    and a.rtolocationid=b.rtolocationid and NewPdfRunningNo is null and erpassigndate is not null and OrderStatus='New Order' and   b.navembcode like  '%'+(select HSRPStateShortName from hsrpstate d where d.hsrp_stateid=a.hsrp_stateid)+'%'  and b.NAVEMBID='" + Navembid + "'    order by  a.HSRP_StateID";

            //}


            DataTable dtSE = Utils.Utils.GetDataTable(stateECQuery, CnnString);
            if (dtSE.Rows.Count > 0)
            {
                foreach (DataRow dr in dtSE.Rows)
                {
                    #region

                    string HSRP_StateID = dr["HSRP_StateID"].ToString().Trim();
                    string HSRPStateName = dr["HSRPStateName"].ToString().Trim();
                    string HSRPStateShortName = dr["HSRPStateShortName"].ToString().Trim();
                    string Navembcode = dr["Navembcode"].ToString().Trim();

                    string dir = dirPath + DateTime.Now.ToString("yyyy-MM-dd") + "\\" + HSRPStateShortName + "\\";
                    //string dir = dirPath + "/" + DateTime.Now.ToString("yyyy-MM-dd") + "/" + HSRPStateShortName + "/";
                    // string fileName = filePrefix + "-" + Navembcode + ".pdf";
                    string fileName = "AllOEM" + "-" + filePrefix + "-" + Navembcode + ".pdf";
                    string filePath = dir + fileName;

                    //string folderpath = ConfigurationManager.AppSettings["InvoiceFolder"].ToString() + "/" + FinYear + "/" + oemid + "/" + HSRPStateID + "/";

                    StringBuilder html = new StringBuilder();

                    Boolean findRecord = false;
                    string strProductionSheetNo = string.Empty;

                    /*
                     *  Start body & HTMl Tag
                     */
                    #region
                    html.Append(
                        "<!DOCTYPE html>" +
                        "<html>" +
                        "<head>" +
                            "<meta charset='UTF-8'><title>Title</title>" +
                            "<style>" +
                                "@page {" +
                                    /* headers*/
                                    "@top-left {" +
                                        "content: 'Left header';" +
                                    "}" +
                                    "@top-right {" +
                                        "content: 'Right header';" +
                                    "}" +

                                    /* footers */
                                    "@bottom-left {" +
                                        "content: 'Lorem ipsum';" +
                                    "} " +
                                    "@bottom-right {" +
                                        "content: 'Page ' counter(page) ' of ' counter(pages);" +
                                    "}" +
                                    "@bottom-center  {" +
                                        "content:element(footer);" +
                                    "}" +
                                "}" +
                                 "#footer {" +
                                    "position: running(footer);" +
                                "}" +
                                "table {" +
                                  "border-collapse: collapse;" +
                                "}" +

                                "table, th, td {" +
                                    "border: 1px solid black;" +
                                    "text-align: left;" +
                                    "vertical-align: top;" +
                                    "padding:5px;" +
                                "}" +
                            "</style>" +
                        "</head>" +
                        "<body>");
                    #endregion

                    string oemDealerQuery = string.Empty;

                    DataTable dtOD = new DataTable();

                    
                    SqlConnection con1 = new SqlConnection(CnnString);
                    SqlCommand cmd1 = new SqlCommand("USP_BindDealerForProduction", con1);
                    cmd1.CommandType = CommandType.StoredProcedure;
                    con1.Open();
                    
                    cmd1.Parameters.AddWithValue("@HSRP_StateID", HSRP_StateID);
                    cmd1.Parameters.AddWithValue("@navembid", Navembcode);


                    SqlDataAdapter da = new SqlDataAdapter(cmd1);
                    // dtOD = new DataTable();
                    cmd1.CommandTimeout = 400;
                    da.Fill(dtOD);
                    con1.Close();

                    #region
                    //DataTable dtOD = Utils.Utils.GetDataTable(oemDealerQuery, CnnString);
                    if (dtOD.Rows.Count > 0)
                    {
                        foreach (DataRow drOD in dtOD.Rows)
                        {

                            string oemid = drOD["oemid"].ToString().Trim();
                            string dealerid = drOD["dealerid"].ToString().Trim();
                            string oemname = drOD["oemname"].ToString().Trim();
                            string dealername = drOD["dealername"].ToString().Trim();
                            string Address = drOD["Address"].ToString().Trim();

                            //start sql query
                            #region
                            string productionQuery = string.Empty;

                            DataTable dtProduction = new DataTable(); ;




                            SqlConnection con = new SqlConnection(CnnString);
                            SqlCommand cmd = new SqlCommand("USP_AllOEMProductionSheet", con);
                            cmd.CommandType = CommandType.StoredProcedure;
                            con.Open();
                            cmd.Parameters.AddWithValue("@navembid", Navembcode);
                            cmd.Parameters.AddWithValue("@HSRP_StateID", HSRP_StateID);
                            cmd.Parameters.AddWithValue("@Dealerid", dealerid);
                            //cmd.Parameters.AddWithValue("@Affix_Id", strsubaffixid);
                            SqlDataAdapter da1 = new SqlDataAdapter(cmd);
                            // dtProduction = new DataTable();
                            da1.Fill(dtProduction);
                            con.Close();

                            #endregion
                            //end sql query

                            #region
                            //DataTable dtProduction = Utils.Utils.GetDataTable(productionQuery, CnnString);
                            if (dtProduction.Rows.Count > 0)
                            {

                                findRecord = true;
                                string strRunningNo = string.Empty;
                                //string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) as maxSheetNo from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + NAVEMBID + "'";
                                //string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                                string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                                string strCom = Utils.Utils.getScalarValue(strSel, CnnString);
                                //DataTable dtSheetNo = fillDataTable(strSel, CnnString);
                                if (strCom.Equals(0) || strCom.Length == 0)
                                {
                                    strRunningNo = "0000001";
                                }
                                else
                                {
                                    strRunningNo = string.Format("{0:0000000}", Convert.ToInt32(strCom) + 1);
                                }
                                string strRequeNo = "select (prefixtext+right('00000'+ convert(varchar,Lastreqno+1),5)) as Reqno from prefix_Requisition  with(nolock)  where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                                ReqNum = Utils.Utils.getScalarValue(strRequeNo, CnnString);
                                string strPRFIX = "select PrefixText from EmbossingCentersNew  with(nolock) where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                                string strPRFIXCom = Utils.Utils.getScalarValue(strPRFIX, CnnString);
                                strProductionSheetNo = strPRFIXCom + strRunningNo;

                                string RTOLocationName = dtProduction.Rows[0]["RTOLocationName"].ToString();

                                #region
                                html.Append("<div style='page-break-before: avoid;page-break-inside: avoid;page-break-after: always;'>" +
                                        "<table style='width:100%;border: 0px;'>" +
                                            "<tr style='border: 0px;'>" +
                                                "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>Report Generation Date:</b> " + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</div>" +
                                                "</td>" +

                                                 "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>ROSMERTA SAFETY SYSTEMS LIMITED</b> " + "</div>" +
                                                "</td>" +


                                                 "<td colspan='5' style='border: 0px;'>" +
                                                    "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Production Sheet No:</b> " + strProductionSheetNo +  " </div>" +
                                                "</td>" +

                                            "</tr>" +

                                              "<tr style='border: 0px;'>" +
                                                "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>State:</b> " + HSRPStateName + "</div>" +
                                                "</td>" +

                                                 "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>Production Sheet</b> " + "</div>" +
                                                "</td>" +


                                                 "<td colspan='5' style='border: 0px;'>" +
                                                    "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Dealer ID/Name:</b> " + dealerid + "/" + dealername + " </div>" +
                                                "</td>" +

                                            "</tr>" +

                                                "<tr style='border: 0px;'>" +
                                                "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>EC Location: </b> " + RTOLocationName + "</div>" +
                                                "</td>" +

                                                 "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='style='text-align:left'><b>Oem :</b> " + oemname +  "</div>" +
                                                "</td>" +


                                                 "<td colspan='5' style='border: 0px;'>" +
                                                     "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Dealer Address:</b> " + Address + " </div>" +
                                                    //"<div style='text-align:left'><b>Dealer Address:</b> " + Address + " </div>" +
                                                "</td>" +

                                            "</tr>" +

                                            
                                            // Sr.No.Vehicle No.Front Plate Size    Front Laser No.Rear Plate Size Rear Laser No.H.S.Foil  Caution Sticker Fuel Type   VT  VC

                                            "<tr>" +
                                                "<td style='text-align:center'>Sr. No.</td>" +
                                                  "<td>Vehicle No.</td>" +
                                                   "<td>Front Plate Size</td>" +
                                                "<td>Front Laser No.</td>" +

                                                "<td>Rear Plate Size</td>" +
                                                "<td>Rear Laser No.</td>" +

                                                "<td>H. S. Foil </td>" +
                                                "<td>Caution Sticker</td>" +
                                                "<td>Fuel Type</td>" +
                                                "<td >VT</td>" +
                                                "<td >VC</ td>" +
                                           
                                            "</tr>");
                                #endregion

                                #region
                                foreach (DataRow drProduction in dtProduction.Rows)
                                {
                                    string HsrprecordID = drProduction["hsrprecordID"].ToString().Trim();
                                    string SRNo = drProduction["SRNo"].ToString().Trim();
                                    string ORD = "";// drProduction["ORD"].ToString().Trim();
                                    string VC = drProduction["VehicleClass"].ToString().Trim();
                                    string VehicleNo = drProduction["VehicleRegNo"].ToString().Trim();
                                    string VT = drProduction["VehicleType"].ToString().Trim();
                                    //string ChassisNo = drProduction["ChassisNo"].ToString().Trim();
                                    //string EngineNo = drProduction["EngineNo"].ToString().Trim();
                                    string FuelType = drProduction["FuelType"].ToString().Trim();
                                    string FrontPSize = drProduction["FrontProductCode"].ToString().Trim();
                                    string FrontLaserNo = drProduction["HSRP_Front_LaserCode"].ToString().Trim();
                                    string RearPSize = drProduction["RearProductCode"].ToString().Trim();
                                    string RearLaserNo = drProduction["HSRP_Rear_LaserCode"].ToString().Trim();
                                    string Amount = drProduction["roundoff_netamount"].ToString().Trim();
                                    //string OrderStatus = drProduction["OrderStatus"].ToString().Trim();
                                    string stickerColor = drProduction["stickerColor"].ToString().Trim();
                                    string HotStampingFoilColour = drProduction["HotStampingFoilColour"].ToString().Trim();

                                    html.Append("<tr>" +
                                       "<td style='text-align:center'>" + SRNo + "</td>" +
                                           "<td style='text-align:center;font-size:20px;'>" + "<b>" + VehicleNo + "</b> </td>" +
                                           
                                            "<td>" + FrontPSize + "</td>" +
                                             "<td style='text-align:center;font-size:20px;'>" + "<b>" + FrontLaserNo + "</b> </td>" +
                                               "<td>" + RearPSize + "</td>" +
                                             "<td style='text-align:center;font-size:20px;'>" + "<b>" + RearLaserNo + "</b> </td>" +
                                              
                                              "<td>" + HotStampingFoilColour + "</td>" +
                                               "<td>" + stickerColor + "</td>" +
                                                "<td>" + FuelType + "</td>" +
                                                "<td>" + VT + "</td>" +
                                                "<td>" + VC + "</td>" +
                                   

                                     
                                      
                                       
                                   "</tr>");

                                    try
                                    {
                                        //start updating hsrprecords 
                                        string sqlUpdateHSRPRecords = "update hsrprecords set sendtoProductionStatus='Y', NAVPDFFlag='1', NewPdfRunningNo='" + strProductionSheetNo + "', Requisitionsheetno='" + ReqNum + "',  " +
                                              "PdfDownloadDate=GetDate(), pdfFileName='" + fileName + "', PDFDownloadUserID='1' where hsrprecordID='" + HsrprecordID + "' ";

                                        Utils.Utils.ExecNonQuery(sqlUpdateHSRPRecords, CnnString);   // uncomment after testing
                                    }
                                    catch (Exception ev)
                                    {
                                        Label1.Text = "hsrprecords update error: " + ev.Message;
                                    }
                                    //end 

                                }
                                #endregion
                                html.Append("</table>");

                                html.Append("</div>");

                                try
                                {
                                    string StrSqlUpdateECQuery = "update EmbossingCentersNew set NewProductionSheetRunningNo='" + strProductionSheetNo + "' " +
                                     "where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                                    Utils.Utils.ExecNonQuery(StrSqlUpdateECQuery, CnnString); // uncomment after testing
                                }
                                catch (Exception ev)
                                {
                                    Label1.Text = "EmbossingCentersNew update error: " + ev.Message;
                                }
                            }

                            #endregion

                        }
                    }// close oemDealerQuery
                    #endregion

                    #region "Req Generate"
                    string strComNew = string.Empty;
                    string strReqNumber = string.Empty;
                    string strReqNo = string.Empty;
                    string strSqlQuery1 = "select CompanyName from hsrpstate  with(nolock) where hsrp_stateid='" + HSRP_StateID + "'";
                    strCompanyName = Utils.Utils.getScalarValue(strSqlQuery1, CnnString);

                    string strEMBName = " select EmbCenterName from EmbossingCentersNew  with(nolock) where  State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                    strComNew = Utils.Utils.getScalarValue(strEMBName, CnnString);
                    if (strComNew != string.Empty)
                    {
                        //strReqNo = "select (prefixtext+right('00000'+ convert(varchar,lastno+1),5)) as Reqno from prefix_Requisition  with(nolock)  where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                        //strReqNumber = Utils.Utils.getScalarValue(strReqNo, CnnString);

                        SQLString = "Exec [laserreqSlip1DashBoard]  '" + HSRP_StateID + "','" + Navembcode + "' ,  '" + ReqNum + "'";
                        DataTable dtResult = Utils.Utils.GetDataTable(SQLString, CnnString);
                        string strQuery = string.Empty;
                        string strRtoLocationName = string.Empty;
                        int Itotal = 0;

                        html.Append("<div style='width:100%;height:100%;'>" +
                                            "<table style='width:100%'>" +

                                                "<tr>" +
                                                    "<td colspan='12'>" +
                                                        "<div style='text-align:center;padding:8px;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + strCompanyName + "</b>" +

                                                            "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                             "<b style='font-size:20px;'>" + "MATERIAL REQUSITION NOTE" + "</b>" +

                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +

                                                "<tr>" +
                                                    "<td colspan='12'>" +
                                                        "<div style='text-align:center;padding:8px;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Production Sheet Date :" + DateTime.Now.ToString("dd-MM-yyyy") + "</b>" +
                                                            "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                             "<b style='font-size:20px;'>" + "" + "</b>" +
                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +
                                                "<tr>" +
                                                    "<td colspan='6'>" +
                                                        "<div style='text-align:left;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>REQ.NO:-</b>" +
                                                            "" + ReqNum + "" +
                                                        "</div>" +
                                                    "</td>" +
                                                    "<td colspan='6'>" +
                                                        "<div style='text-align:left;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>Embossing Center:</b>" +
                                                            "" + strComNew + "" +
                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +
                                                "<tr>" +
                                                    "<td colspan='12'>" +
                                                        "<div style='text-align:left;padding:2px;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "State:" + HSRPStateName + "</b>" +
                                                            "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                             "<b style='font-size:20px;'>" + "" + "</b>" +
                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +
                                                 "<tr>" +
                                                    "<td colspan='12'>" +
                                                        "<div style='text-align:left;padding:8px;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + " " + "</b>" +
                                                            "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                             "<b style='font-size:20px;'>" + "" + "</b>" +
                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +

                                                "<tr>" +
                                                    "<td colspan='1' style='text-align:center'>SR.N.</td>" +
                                                    "<td colspan='3'>Product Size</td>" +
                                                    "<td colspan='1'>Laser Count</td>" +
                                                    "<td colspan='1'>Start Laser No</td>" +
                                                    "<td colspan='1'>End Laser No</td>" +

                                                "</tr>");


                        if (dtResult.Rows.Count > 0)
                        {
                            for (int i = 0; i < dtResult.Rows.Count; i++)
                            {
                                string ID = dtResult.Rows[i]["ID"].ToString();
                                string productcode = dtResult.Rows[i]["productcode"].ToString();
                                string LaserCount = dtResult.Rows[i]["LaserCount"].ToString();
                                Itotal = Convert.ToInt32(dtResult.Rows[i]["Total"].ToString());
                                string BeginLaser = dtResult.Rows[i]["BeginLaser"].ToString();
                                string EndLaser = dtResult.Rows[i]["EndLaser"].ToString();




                                html.Append("<tr>" +
                                   "<td colspan='1' style='text-align:left'>" + ID + "</td>" +
                                   "<td colspan='3'>" + productcode + "</td>" +

                                   "<td colspan='1'>" + LaserCount + "</td>" +
                                   "<td colspan='1'>" + BeginLaser + "</td>" +
                                   "<td colspan='1'>" + EndLaser + "</td>" +

                               "</tr>");
                            }
                        }
                        html.Append("<tr>" +
                                 "<td colspan='1' style='text-align:center' > " + "<b>Grand Total:</b>" + "</td>" +
                                 "<td colspan='3'>" + "" + "</td>" +

                                 "<td colspan='1'>" + Itotal + "</td>" +
                                 "<td colspan='1'>" + " " + "</td>" +
                                 "<td colspan='1'>" + " " + "</td>" +

                             "</tr>");




                        html.Append("<tr>" +
                         "<td colspan='2' > " + "<b>REQUESTED BY </b>" + "</td>" +
                         "<td colspan='2'>" + "<b>AUTHORIZED BY </b>" + "</td>" +

                         "<td colspan='2'>" + "<b>ISSUED BY </b>" + "</td>" +
                         "<td colspan='2'>" + "<b>RECEIVED BY</b>" + "</td>" +


                     "</tr>");

                        html.Append("<tr>" +
                                                  "<td colspan='12'>" +
                                                      "<div style='text-align:left;padding:2px;'>" +
                                                          "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Name" + "</b>" +

                                                      "</div>" +
                                                  "</td>" +
                                              "</tr>");

                        html.Append("<tr>" +
                                                 "<td colspan='12'>" +
                                                     "<div style='text-align:left;padding:2px;'>" +
                                                         "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Designation" + "</b>" +

                                                     "</div>" +
                                                 "</td>" +
                                             "</tr>");

                        html.Append("<tr>" +
                                               "<td colspan='12'>" +
                                                   "<div style='text-align:right;padding:8px;'>" +
                                                       "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Download By: Admin " + "</b>" +

                                                   "</div>" +
                                               "</td>" +
                                           "</tr>");

                        html.Append("<tr>" +
                                              "<td colspan='12'>" +
                                                  "<div style='text-align:right;padding:8px;'>" +
                                                      "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Sheet Generated By :" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</b>" +

                                                  "</div>" +
                                              "</td>" +
                                          "</tr>");






                        html.Append("</table>");

                        html.Append("</div>");

                        try
                        {
                            //start updating hsrprecords 
                            string Query = "update prefix_Requisition set lastno=lastno+1,Lastreqno=Lastreqno+1 where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                            Utils.Utils.ExecNonQuery(Query, CnnString);
                        }
                        catch (Exception ev)
                        {
                            Label1.Text = "prefix Requisition update error: " + ev.Message;
                        }
                    }


                    #endregion

                    /*
                     * Close body & HTMl Tag
                     */
                    html.Append("</body>" +
                        "</html>");


                    if (findRecord)
                    {

                        string strSaveSQlQuery = "insert into ProductionSheetAutoGenerated_List (HSRP_StateID, State_Code, RTO_ID, Emb_Center_Id, FileName, Productiondate, ProductionTime, OrderNo) values " +
                                    "('" + HSRP_StateID + "', '" + HSRPStateShortName + "', '', '" + Navembcode + "', '" + fileName + "', '" + DateTime.Now.ToString("yyyy-MM-dd") + "', '" + DateTime.Now.ToString("HH:mm:ss") + "', '" + filePrefix + "') ";
                        Utils.Utils.ExecNonQuery(strSaveSQlQuery, CnnString); // uncomment after testing
                        Label1.Text = Label1.Text + "successfully created production sheet of State: " + HSRPStateShortName + ", , Emb_Center_Id: " + Navembcode + Environment.NewLine;

                        //lblLog.Text = lblLog.Text + Environment.NewLine + "Start Production Query at:" + DateTime.Now;
                        #region
                        try
                        {
                            if (!Directory.Exists(dir))
                            {


                                Directory.CreateDirectory(dir);

                            }

                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        try
                        {
                            var pdf = Pdf
                                    .From(html.ToString())
                                    .OfSize(PaperSize.A4)
                                    .WithTitle("Title")
                                    .WithoutOutline()
                                    .WithMargins(1.25.Centimeters())
                                   .Landscape()
                                    .Comressed()
                                    .Content();

                            FileStream readStream = File.Create(filePath);
                            BinaryWriter binaryWriter = new BinaryWriter(readStream);

                            // Write the binary data to the file
                            binaryWriter.Write(pdf);
                            binaryWriter.Close();
                            readStream.Close();
                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        #endregion
                    }

                    #endregion

                }//close foreach stateEcQuery
            }
        }

        protected void btnSuzu_Click(object sender, EventArgs e)
        {

            if (ddlStateName.SelectedItem.Text == "--Select State--")
            {
                lblErrMess.Text = "Please select  State";
                return;
            }


            string sqlPrefixQuery = "select top 1 CONVERT(INT, isnull(OrderNo,'0')) + 1 as orderNo from ProductionSheetAutoGenerated_List order by id desc";
            DataTable dtPrefix = Utils.Utils.GetDataTable(sqlPrefixQuery, CnnString);

            if (dtPrefix.Rows.Count > 0)
            {
                filePrefix = dtPrefix.Rows[0]["orderNo"].ToString();
            }
            else
            {
                filePrefix = "1";
            }

            //txtFilePrefix.Text = filePrefix;

            if (filePrefix.Length > 0)
            {

                SheetGenerationSuzu();


            }

        }

        protected void btnRenault_Click(object sender, EventArgs e)
        {

            if (ddlStateName.SelectedItem.Text == "--Select State--")
            {
                lblErrMess.Text = "Please select  State";
                return;
            }
           

            string sqlPrefixQuery = "select top 1 CONVERT(INT, isnull(OrderNo,'0')) + 1 as orderNo from ProductionSheetAutoGenerated_List order by id desc";
            DataTable dtPrefix = Utils.Utils.GetDataTable(sqlPrefixQuery, CnnString);

            if (dtPrefix.Rows.Count > 0)
            {
                filePrefix = dtPrefix.Rows[0]["orderNo"].ToString();
            }
            else
            {
                filePrefix = "1";
            }

            //txtFilePrefix.Text = filePrefix;

            if (filePrefix.Length > 0)
            {

                SheetGenerationRenault();


            }

        }

        private void SheetGenerationSuzu()
        {
            string stateECQuery = string.Empty;
            string ReqNum = string.Empty;
           
            FillUserDetails();

            stateECQuery = "select distinct a.HSRP_StateID, (select HSRPStateName from hsrpstate c where c.hsrp_stateid=a.hsrp_stateid) as HSRPStateName,(select HSRPStateShortName from hsrpstate d where d.hsrp_stateid=a.hsrp_stateid) as HSRPStateShortName, d.navembcode from hsrprecords a join DealerAffixation d  on  a.Affix_Id=d.SubDealerId and NewPdfRunningNo is null and  a.HSRP_StateID ='15' and d.Navembcode not like '%CODO%'  and erpassigndate is not null   AND d.navembcode='" + Navembid + "' order by  a.HSRP_StateID";


            DataTable dtSE = Utils.Utils.GetDataTable(stateECQuery, CnnString);
            if (dtSE.Rows.Count > 0)
            {
                foreach (DataRow dr in dtSE.Rows)
                {
                    #region

                    string HSRP_StateID = dr["HSRP_StateID"].ToString().Trim();
                    string HSRPStateName = dr["HSRPStateName"].ToString().Trim();
                    string HSRPStateShortName = dr["HSRPStateShortName"].ToString().Trim();
                    string Navembcode = dr["Navembcode"].ToString().Trim();

                    string dir = dirPath + DateTime.Now.ToString("yyyy-MM-dd") + "\\" + HSRPStateShortName + "\\";
                    
                    string fileName = "Isuzu" + "-" + filePrefix + "-" + Navembcode + ".pdf";
                    string filePath = dir + fileName;
                    StringBuilder html = new StringBuilder();

                    Boolean findRecord = false;
                    string strProductionSheetNo = string.Empty;

                    /*
                     *  Start body & HTMl Tag
                     */
                    #region
                    html.Append(
                        "<!DOCTYPE html>" +
                        "<html>" +
                        "<head>" +
                            "<meta charset='UTF-8'><title>Title</title>" +
                            "<style>" +
                                "@page {" +
                                    /* headers*/
                                    "@top-left {" +
                                        "content: 'Left header';" +
                                    "}" +
                                    "@top-right {" +
                                        "content: 'Right header';" +
                                    "}" +

                                    /* footers */
                                    "@bottom-left {" +
                                        "content: 'Lorem ipsum';" +
                                    "} " +
                                    "@bottom-right {" +
                                        "content: 'Page ' counter(page) ' of ' counter(pages);" +
                                    "}" +
                                    "@bottom-center  {" +
                                        "content:element(footer);" +
                                    "}" +
                                "}" +
                                 "#footer {" +
                                    "position: running(footer);" +
                                "}" +
                                "table {" +
                                  "border-collapse: collapse;" +
                                "}" +

                                "table, th, td {" +
                                    "border: 1px solid black;" +
                                    "text-align: left;" +
                                    "vertical-align: top;" +
                                    "padding:5px;" +
                                "}" +
                            "</style>" +
                        "</head>" +
                        "<body>");
                    #endregion

                    string oemDealerQuery = string.Empty;




                    oemDealerQuery = "select distinct om.oemid as oemid, om.name as oemname, dm.dealerid, d.Subdealername as Dealername, dm.dealercode, d.Address,d.SubDealerId, dm.HSRP_StateID, " +
                        "dm.RTOLocationID from oemmaster om " +
                        "left join dealermaster dm on dm.oemid = om.oemid join DealerAffixation d on d.DealerID=dm.DealerId where dm.HSRP_StateID =" + HSRP_StateID + " and   d.Navembcode='" + Navembcode + "' and " +

                        "dm.dealerid in (select distinct dealerid from hsrprecords where NewPdfRunningNo is null and erpassigndate is not null and   affix_id is not null ) and    OM.OEMid in('8')";


                    #region
                    DataTable dtOD = Utils.Utils.GetDataTable(oemDealerQuery, CnnString);
                    if (dtOD.Rows.Count > 0)
                    {
                        foreach (DataRow drOD in dtOD.Rows)
                        {

                            string oemid = drOD["oemid"].ToString().Trim();
                            string dealerid = drOD["dealerid"].ToString().Trim();
                            string oemname = drOD["oemname"].ToString().Trim();
                            string dealername = drOD["dealername"].ToString().Trim();
                            string Address = drOD["Address"].ToString().Trim();
                            string strsubaffixid = drOD["SubDealerId"].ToString();

                            //start sql query
                            #region
                            string productionQuery = string.Empty;


                            DataTable dtProduction = new DataTable();




                            #endregion
                            //end sql query

                            #region

                            SqlConnection con = new SqlConnection(CnnString);
                            SqlCommand cmd = new SqlCommand("USP_RenaultProductionSheet", con);
                            cmd.CommandType = CommandType.StoredProcedure;
                            con.Open();
                            cmd.Parameters.AddWithValue("@navembid", Navembcode);
                            cmd.Parameters.AddWithValue("@HSRP_StateID", HSRP_StateID);
                            cmd.Parameters.AddWithValue("@Dealerid", dealerid);
                            cmd.Parameters.AddWithValue("@Affix_Id", strsubaffixid);
                            SqlDataAdapter da = new SqlDataAdapter(cmd);
                            // dtProduction = new DataTable();
                            da.Fill(dtProduction);
                            con.Close();


                            // DataTable dtProduction = Utils.Utils.GetDataTable(productionQuery, CnnString);
                            if (dtProduction.Rows.Count > 0)
                            {

                                findRecord = true;
                                string strRunningNo = string.Empty;
                                //string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) as maxSheetNo from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + NAVEMBID + "'";
                                string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                                string strCom = Utils.Utils.getScalarValue(strSel, CnnString);
                                //DataTable dtSheetNo = fillDataTable(strSel, CnnString);
                                if (strCom.Equals(0) || strCom.Length == 0)
                                {
                                    strRunningNo = "0000001";
                                }
                                else
                                {
                                    strRunningNo = string.Format("{0:0000000}", Convert.ToInt32(strCom) + 1);
                                }
                                string strRequeNo = "select (prefixtext+right('00000'+ convert(varchar,Lastreqno+1),5)) as Reqno from prefix_Requisition  where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                                ReqNum = Utils.Utils.getScalarValue(strRequeNo, CnnString);
                                string strPRFIX = "select PrefixText from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                                string strPRFIXCom = Utils.Utils.getScalarValue(strPRFIX, CnnString);
                                strProductionSheetNo = strPRFIXCom + strRunningNo;

                                string RTOLocationName = dtProduction.Rows[0]["RTOLocationName"].ToString();

                                #region

                                html.Append("<div style='page-break-before: avoid;page-break-inside: avoid;page-break-after: always;'>" +
                                        "<table style='width:100%;border: 0px;'>" +
                                            "<tr style='border: 0px;'>" +
                                                "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>Report Generation Date:</b> " + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</div>" +
                                                "</td>" +

                                                 "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>ROSMERTA SAFETY SYSTEMS LIMITED</b> " + "</div>" +
                                                "</td>" +


                                                 "<td colspan='5' style='border: 0px;'>" +
                                                    "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Production Sheet No:</b> " + strRunningNo + " </div>" +
                                                "</td>" +

                                            "</tr>" +

                                              "<tr style='border: 0px;'>" +
                                                "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>State:</b> " + HSRPStateName + "</div>" +
                                                "</td>" +

                                                 "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>Production Sheet</b> " + "</div>" +
                                                "</td>" +


                                                 "<td colspan='5' style='border: 0px;'>" +
                                                    "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Dealer ID/Name:</b> " + dealerid + "/" + dealername + "/" + strsubaffixid + " </div>" +
                                                "</td>" +

                                            "</tr>" +

                                                "<tr style='border: 0px;'>" +
                                                "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>EC Location: </b> " + RTOLocationName + "</div>" +
                                                "</td>" +

                                                 "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='style='text-align:left'><b>Oem :</b> " + oemname + "</div>" +
                                                "</td>" +


                                                 "<td colspan='5' style='border: 0px;'>" +
                                                     "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Dealer Address:</b> " + Address + " </div>" +
                                                //"<div style='text-align:left'><b>Dealer Address:</b> " + Address + " </div>" +
                                                "</td>" +

                                            "</tr>" +


  "<tr>" +
                                                "<td style='text-align:center'>Sr. No.</td>" +
                                                  "<td>Vehicle No.</td>" +
                                                   "<td>Front Plate Size</td>" +
                                                "<td>Front Laser No.</td>" +

                                                "<td>Rear Plate Size</td>" +
                                                "<td>Rear Laser No.</td>" +

                                                "<td>H. S. Foil </td>" +
                                                "<td>Caution Sticker</td>" +
                                                "<td>Fuel Type</td>" +
                                                "<td >VT</td>" +
                                                "<td >VC</ td>" +

                                            "</tr>");

                                #endregion

                                #region
                                foreach (DataRow drProduction in dtProduction.Rows)
                                {
                                    string HsrprecordID = drProduction["hsrprecordID"].ToString().Trim();
                                    string SRNo = drProduction["SRNo"].ToString().Trim();
                                    string ORD = "";// drProduction["ORD"].ToString().Trim();
                                    string VC = drProduction["VehicleClass"].ToString().Trim();
                                    string VehicleNo = drProduction["VehicleRegNo"].ToString().Trim();
                                    // string VehicleNo = "<td style='text-align:center;font-size:20px;'>" + "<b>" + drProduction["VehicleRegNo"].ToString().Trim() + "</b> </td>" +
                                    string VT = drProduction["VehicleType"].ToString().Trim();
                                    //string ChassisNo = drProduction["ChassisNo"].ToString().Trim();
                                    //string EngineNo = drProduction["EngineNo"].ToString().Trim();
                                    string FuelType = drProduction["FuelType"].ToString().Trim();
                                    string FrontPSize = drProduction["FrontProductCode"].ToString().Trim();
                                    // string FrontLaserNo = "<td style='text-align:center;font-size:20px;'>" + "<b>" + drProduction["HSRP_Front_LaserCode"].ToString().Trim() + "</b> </td>" +
                                    string FrontLaserNo = drProduction["HSRP_Front_LaserCode"].ToString().Trim();
                                    string RearPSize = drProduction["RearProductCode"].ToString().Trim();
                                    string RearLaserNo = drProduction["HSRP_Rear_LaserCode"].ToString().Trim();
                                    //string RearLaserNo = "<td style='text-align:center;font-size:20px;'>" + "<b>" + drProduction["HSRP_Rear_LaserCode"].ToString().Trim() + "</b> </td>" +
                                    string Amount = drProduction["roundoff_netamount"].ToString().Trim();
                                    //string OrderStatus = drProduction["OrderStatus"].ToString().Trim();
                                    string stickerColor = drProduction["stickerColor"].ToString().Trim();
                                    string HotStampingFoilColour = drProduction["HotStampingFoilColour"].ToString().Trim();

                                    html.Append("<tr>" +
                                   "<td style='text-align:center'>" + SRNo + "</td>" +
                                       "<td style='text-align:center;font-size:20px;'>" + "<b>" + VehicleNo + "</b> </td>" +

                                        "<td>" + FrontPSize + "</td>" +
                                         "<td style='text-align:center;font-size:20px;'>" + "<b>" + FrontLaserNo + "</b> </td>" +
                                           "<td>" + RearPSize + "</td>" +
                                         "<td style='text-align:center;font-size:20px;'>" + "<b>" + RearLaserNo + "</b> </td>" +

                                          "<td>" + HotStampingFoilColour + "</td>" +
                                           "<td>" + stickerColor + "</td>" +
                                            "<td>" + FuelType + "</td>" +
                                            "<td>" + VT + "</td>" +
                                            "<td>" + VC + "</td>" +





                               "</tr>");


                                    try
                                    {
                                        //start updating hsrprecords 
                                        string sqlUpdateHSRPRecords = "update hsrprecords set sendtoProductionStatus='Y', NAVPDFFlag='1', NewPdfRunningNo='" + strProductionSheetNo + "', Requisitionsheetno='" + ReqNum + "',  " +
                                              "PdfDownloadDate=GetDate(), pdfFileName='" + fileName + "', PDFDownloadUserID='1' where hsrprecordID='" + HsrprecordID + "' ";

                                        Utils.Utils.ExecNonQuery(sqlUpdateHSRPRecords, CnnString);   // uncomment after testing
                                    }
                                    catch (Exception ev)
                                    {
                                        Label1.Text = "hsrprecords update error: " + ev.Message;
                                    }
                                    //end 

                                }
                                #endregion
                                html.Append("</table>");

                                html.Append("</div>");

                                try
                                {
                                    string StrSqlUpdateECQuery = "update EmbossingCentersNew set NewProductionSheetRunningNo='" + strProductionSheetNo + "' " +
                                     "where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                                    Utils.Utils.ExecNonQuery(StrSqlUpdateECQuery, CnnString); // uncomment after testing
                                }
                                catch (Exception ev)
                                {
                                    Label1.Text = "EmbossingCentersNew update error: " + ev.Message;
                                }
                            }

                            #endregion

                        }
                    }// close oemDealerQuery
                    #endregion

                    #region "Req Generate"
                    string strComNew = string.Empty;
                    string strReqNumber = string.Empty;
                    string strReqNo = string.Empty;



                    string strSqlQuery1 = "select CompanyName from hsrpstate where hsrp_stateid='" + HSRP_StateID + "'";
                    strCompanyName = Utils.Utils.getScalarValue(strSqlQuery1, CnnString);

                    string strEMBName = " select EmbCenterName from EmbossingCentersNew where  State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                    strComNew = Utils.Utils.getScalarValue(strEMBName, CnnString);

                    if (strComNew != "")
                    {



                        strReqNo = "select (prefixtext+right('00000'+ convert(varchar,lastno+1),5)) as Reqno from prefix_Requisition  where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                        strReqNumber = Utils.Utils.getScalarValue(strReqNo, CnnString);

                        SQLString = "Exec [laserreqSlip1DashBoard]  '" + HSRP_StateID + "','" + Navembcode + "' ,  '" + ReqNum + "'";
                        DataTable dtResult = Utils.Utils.GetDataTable(SQLString, CnnString);

                        string strQuery = string.Empty;
                        string strRtoLocationName = string.Empty;
                        int Itotal = 0;

                        html.Append("<div style='width:100%;height:100%;'>" +
                                            "<table style='width:100%'>" +

                                                "<tr>" +
                                                    "<td colspan='12'>" +
                                                        "<div style='text-align:center;padding:8px;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + strCompanyName + "</b>" +

                                                            "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                             "<b style='font-size:20px;'>" + "MATERIAL REQUSITION NOTE" + "</b>" +

                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +

                                                "<tr>" +
                                                    "<td colspan='12'>" +
                                                        "<div style='text-align:center;padding:8px;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Production Sheet Date :" + DateTime.Now.ToString("dd-MM-yyyy") + "</b>" +
                                                            "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                             "<b style='font-size:20px;'>" + "" + "</b>" +
                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +
                                                "<tr>" +
                                                    "<td colspan='6'>" +
                                                        "<div style='text-align:left;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>REQ.NO:-</b>" +
                                                            "" + strReqNumber + "" +
                                                        "</div>" +
                                                    "</td>" +
                                                    "<td colspan='6'>" +
                                                        "<div style='text-align:left;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>Embossing Center:</b>" +
                                                            "" + strComNew + "" +
                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +
                                                "<tr>" +
                                                    "<td colspan='12'>" +
                                                        "<div style='text-align:left;padding:2px;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "State:" + HSRPStateName + "</b>" +
                                                            "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                             "<b style='font-size:20px;'>" + "" + "</b>" +
                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +
                                                 "<tr>" +
                                                    "<td colspan='12'>" +
                                                        "<div style='text-align:left;padding:8px;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + " " + "</b>" +
                                                            "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                             "<b style='font-size:20px;'>" + "" + "</b>" +
                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +

                                                "<tr>" +
                                                    "<td colspan='1' style='text-align:center'>SR.N.</td>" +
                                                    "<td colspan='3'>Product Size</td>" +
                                                    "<td colspan='1'>Laser Count</td>" +
                                                    "<td colspan='1'>Start Laser No</td>" +
                                                    "<td colspan='1'>End Laser No</td>" +

                                                "</tr>");


                        if (dtResult.Rows.Count > 0)
                        {
                            for (int i = 0; i < dtResult.Rows.Count; i++)
                            {
                                string ID = dtResult.Rows[i]["ID"].ToString();
                                string productcode = dtResult.Rows[i]["productcode"].ToString();
                                string LaserCount = dtResult.Rows[i]["LaserCount"].ToString();
                                Itotal = Convert.ToInt32(dtResult.Rows[i]["Total"].ToString());
                                string BeginLaser = dtResult.Rows[i]["BeginLaser"].ToString();
                                string EndLaser = dtResult.Rows[i]["EndLaser"].ToString();




                                html.Append("<tr>" +
                                   "<td colspan='1' style='text-align:left'>" + ID + "</td>" +
                                   "<td colspan='3'>" + productcode + "</td>" +

                                   "<td colspan='1'>" + LaserCount + "</td>" +
                                   "<td colspan='1'>" + BeginLaser + "</td>" +
                                   "<td colspan='1'>" + EndLaser + "</td>" +

                               "</tr>");
                            }
                        }
                        html.Append("<tr>" +
                                 "<td colspan='1' style='text-align:center' > " + "<b>Grand Total:</b>" + "</td>" +
                                 "<td colspan='3'>" + "" + "</td>" +

                                 "<td colspan='1'>" + Itotal + "</td>" +
                                 "<td colspan='1'>" + " " + "</td>" +
                                 "<td colspan='1'>" + " " + "</td>" +

                             "</tr>");




                        html.Append("<tr>" +
                         "<td colspan='2' > " + "<b>REQUESTED BY </b>" + "</td>" +
                         "<td colspan='2'>" + "<b>AUTHORIZED BY </b>" + "</td>" +

                         "<td colspan='2'>" + "<b>ISSUED BY </b>" + "</td>" +
                         "<td colspan='2'>" + "<b>RECEIVED BY</b>" + "</td>" +


                     "</tr>");

                        html.Append("<tr>" +
                                                  "<td colspan='12'>" +
                                                      "<div style='text-align:left;padding:2px;'>" +
                                                          "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Name" + "</b>" +

                                                      "</div>" +
                                                  "</td>" +
                                              "</tr>");

                        html.Append("<tr>" +
                                                 "<td colspan='12'>" +
                                                     "<div style='text-align:left;padding:2px;'>" +
                                                         "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Designation" + "</b>" +

                                                     "</div>" +
                                                 "</td>" +
                                             "</tr>");

                        html.Append("<tr>" +
                                               "<td colspan='12'>" +
                                                   "<div style='text-align:right;padding:8px;'>" +
                                                       "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Download By: Admin " + "</b>" +

                                                   "</div>" +
                                               "</td>" +
                                           "</tr>");

                        html.Append("<tr>" +
                                              "<td colspan='12'>" +
                                                  "<div style='text-align:right;padding:8px;'>" +
                                                      "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Sheet Generated By :" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</b>" +

                                                  "</div>" +
                                              "</td>" +
                                          "</tr>");







                        html.Append("</table>");

                        html.Append("</div>");

                        try
                        {
                            //start updating hsrprecords 
                            string Query = "update prefix_Requisition set lastno=lastno+1,Lastreqno=Lastreqno+1 where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                            Utils.Utils.ExecNonQuery(Query, CnnString);
                        }
                        catch (Exception ev)
                        {
                            Label1.Text = "prefix Requisition update error: " + ev.Message;
                        }
                    }


                    #endregion

                    /*
                     * Close body & HTMl Tag
                     */
                    html.Append("</body>" +
                        "</html>");


                    if (findRecord)
                    {

                        string strSaveSQlQuery = "insert into ProductionSheetAutoGenerated_List (HSRP_StateID, State_Code, RTO_ID, Emb_Center_Id, FileName, Productiondate, ProductionTime, OrderNo) values " +
                                    "('" + HSRP_StateID + "', '" + HSRPStateShortName + "', '', '" + Navembcode + "', '" + fileName + "', '" + DateTime.Now.ToString("yyyy-MM-dd") + "', '" + DateTime.Now.ToString("HH:mm:ss") + "', '" + filePrefix + "') ";
                        Utils.Utils.ExecNonQuery(strSaveSQlQuery, CnnString); // uncomment after testing
                        Label1.Text = Label1.Text + "successfully created production sheet of State: " + HSRPStateShortName + ", , Emb_Center_Id: " + Navembcode + Environment.NewLine;

                        //lblLog.Text = lblLog.Text + Environment.NewLine + "Start Production Query at:" + DateTime.Now;
                        #region
                        try
                        {
                            if (!Directory.Exists(dir))
                            {


                                Directory.CreateDirectory(dir);

                            }

                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        try
                        {
                            var pdf = Pdf
                                    .From(html.ToString())
                                    .OfSize(PaperSize.A4)
                                    .WithTitle("Title")
                                    .WithoutOutline()
                                    .WithMargins(1.25.Centimeters())
                                     .Landscape()
                                    .Comressed()
                                    .Content();

                            FileStream readStream = File.Create(filePath);
                            BinaryWriter binaryWriter = new BinaryWriter(readStream);

                            // Write the binary data to the file
                            binaryWriter.Write(pdf);
                            binaryWriter.Close();
                            readStream.Close();
                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        #endregion
                    }

                    #endregion

                }//close foreach stateEcQuery
            }
        }

        private void SheetGenerationRenault()
        {
            string stateECQuery = string.Empty;
            string ReqNum = string.Empty;
            //string Navembid = Session["Navembid"].ToString();
            FillUserDetails();


            // stateECQuery = "select distinct a.HSRP_StateID, (select HSRPStateName from hsrpstate c where c.hsrp_stateid=a.hsrp_stateid) as HSRPStateName,(select HSRPStateShortName from hsrpstate d where d.hsrp_stateid=a.hsrp_stateid) as HSRPStateShortName, navembcode from hsrprecords a,rtolocation b where a.HSRP_StateID ='" + ddlStateName.SelectedValue + "' and   b.Navembcode not like '%CODO%'    and a.rtolocationid=b.rtolocationid and isnull(NewPdfRunningNo,'') = '' and isnull(erpassigndate,'') != ''  order by  a.HSRP_StateID";

            stateECQuery = "select distinct a.HSRP_StateID, (select HSRPStateName from hsrpstate c where c.hsrp_stateid=a.hsrp_stateid) as HSRPStateName,(select HSRPStateShortName from hsrpstate d where d.hsrp_stateid=a.hsrp_stateid) as HSRPStateShortName, d.navembcode from hsrprecords a join DealerAffixation d  on  a.Affix_Id=d.SubDealerId and isnull(NewPdfRunningNo,'') = '' and  a.HSRP_StateID ='19' and d.Navembcode not like '%CODO%'  and isnull(erpassigndate,'') != ''   AND d.navembcode='" + Navembid + "' order by  a.HSRP_StateID";


            DataTable dtSE = Utils.Utils.GetDataTable(stateECQuery, CnnString);
            if (dtSE.Rows.Count > 0)
            {
                foreach (DataRow dr in dtSE.Rows)
                {
                    #region

                    string HSRP_StateID = dr["HSRP_StateID"].ToString().Trim();
                    string HSRPStateName = dr["HSRPStateName"].ToString().Trim();
                    string HSRPStateShortName = dr["HSRPStateShortName"].ToString().Trim();
                    string Navembcode = dr["Navembcode"].ToString().Trim();

                    string dir = dirPath + DateTime.Now.ToString("yyyy-MM-dd") + "\\" + HSRPStateShortName + "\\";
                    //string dir = dirPath + "/" + DateTime.Now.ToString("yyyy-MM-dd") + "/" + HSRPStateShortName + "/";
                    //string fileName = filePrefix + "-" + Navembcode + ".pdf";
                    string fileName = "Renault" + "-" + filePrefix + "-" + Navembcode + ".pdf";
                    string filePath = dir + fileName;

                    //string folderpath = ConfigurationManager.AppSettings["InvoiceFolder"].ToString() + "/" + FinYear + "/" + oemid + "/" + HSRPStateID + "/";

                    StringBuilder html = new StringBuilder();

                    Boolean findRecord = false;
                    string strProductionSheetNo = string.Empty;

                    /*
                     *  Start body & HTMl Tag
                     */
                    #region
                    html.Append(
                        "<!DOCTYPE html>" +
                        "<html>" +
                        "<head>" +
                            "<meta charset='UTF-8'><title>Title</title>" +
                            "<style>" +
                                "@page {" +
                                    /* headers*/
                                    "@top-left {" +
                                        "content: 'Left header';" +
                                    "}" +
                                    "@top-right {" +
                                        "content: 'Right header';" +
                                    "}" +

                                    /* footers */
                                    "@bottom-left {" +
                                        "content: 'Lorem ipsum';" +
                                    "} " +
                                    "@bottom-right {" +
                                        "content: 'Page ' counter(page) ' of ' counter(pages);" +
                                    "}" +
                                    "@bottom-center  {" +
                                        "content:element(footer);" +
                                    "}" +
                                "}" +
                                 "#footer {" +
                                    "position: running(footer);" +
                                "}" +
                                "table {" +
                                  "border-collapse: collapse;" +
                                "}" +

                                "table, th, td {" +
                                    "border: 1px solid black;" +
                                    "text-align: left;" +
                                    "vertical-align: top;" +
                                    "padding:5px;" +
                                "}" +
                            "</style>" +
                        "</head>" +
                        "<body>");
                    #endregion

                    string oemDealerQuery = string.Empty;



                    //oemDealerQuery = "select distinct om.oemid as oemid, om.name as oemname, dm.dealerid, dm.dealername, dm.dealercode, dm.Address, dm.HSRP_StateID, " +
                    //    "dm.RTOLocationID from oemmaster om " +
                    //    "left join dealermaster dm on dm.oemid = om.oemid where dm.HSRP_StateID =" + HSRP_StateID + " and " +
                    //    "dm.RTOLocationID in (select RTOLocationID from rtolocation where Navembcode='" + Navembcode + "' ) and " +
                    //    "dm.dealerid in (select distinct dealerid from hsrprecords where isnull(NewPdfRunningNo,'') = '' and isnull(erpassigndate,'') != '' and HSRP_StateID =" + HSRP_StateID + ") and Om.OEMID   in('21','43')";


                    oemDealerQuery = "select distinct om.oemid as oemid, om.name as oemname, dm.dealerid, d.Subdealername as Dealername, dm.dealercode, d.Address,d.SubDealerId, dm.HSRP_StateID, " +
                        "dm.RTOLocationID from oemmaster om " +
                        "left join dealermaster dm on dm.oemid = om.oemid join DealerAffixation d on d.DealerID=dm.DealerId where dm.HSRP_StateID =" + HSRP_StateID + " and   d.Navembcode='" + Navembcode + "' and " +

                        "dm.dealerid in (select distinct dealerid from hsrprecords where isnull(NewPdfRunningNo,'') = '' and isnull(erpassigndate,'') != '' and   isnull(affix_id,'')!='' ) and    OM.OEMid in('43')";


                    #region
                    DataTable dtOD = Utils.Utils.GetDataTable(oemDealerQuery, CnnString);
                    if (dtOD.Rows.Count > 0)
                    {
                        foreach (DataRow drOD in dtOD.Rows)
                        {

                            string oemid = drOD["oemid"].ToString().Trim();
                            string dealerid = drOD["dealerid"].ToString().Trim();
                            string oemname = drOD["oemname"].ToString().Trim();
                            string dealername = drOD["dealername"].ToString().Trim();
                            string Address = drOD["Address"].ToString().Trim();
                            string strsubaffixid = drOD["SubDealerId"].ToString();

                            //start sql query
                            #region
                            string productionQuery = string.Empty;


                            DataTable dtProduction = new DataTable();

                          


                            #endregion
                            //end sql query

                            #region

                            SqlConnection con = new SqlConnection(CnnString);
                            SqlCommand cmd = new SqlCommand("USP_RenaultProductionSheet", con);
                            cmd.CommandType = CommandType.StoredProcedure;
                            con.Open();
                            cmd.Parameters.AddWithValue("@navembid", Navembcode);
                            cmd.Parameters.AddWithValue("@HSRP_StateID", HSRP_StateID);
                            cmd.Parameters.AddWithValue("@Dealerid", dealerid);
                            cmd.Parameters.AddWithValue("@Affix_Id", strsubaffixid);
                            SqlDataAdapter da = new SqlDataAdapter(cmd);
                            // dtProduction = new DataTable();
                            da.Fill(dtProduction);
                            con.Close();


                            // DataTable dtProduction = Utils.Utils.GetDataTable(productionQuery, CnnString);
                            if (dtProduction.Rows.Count > 0)
                            {

                                findRecord = true;
                                string strRunningNo = string.Empty;
                                //string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) as maxSheetNo from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + NAVEMBID + "'";
                                string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                                string strCom = Utils.Utils.getScalarValue(strSel, CnnString);
                                //DataTable dtSheetNo = fillDataTable(strSel, CnnString);
                                if (strCom.Equals(0) || strCom.Length == 0)
                                {
                                    strRunningNo = "0000001";
                                }
                                else
                                {
                                    strRunningNo = string.Format("{0:0000000}", Convert.ToInt32(strCom) + 1);
                                }
                                string strRequeNo = "select (prefixtext+right('00000'+ convert(varchar,Lastreqno+1),5)) as Reqno from prefix_Requisition  where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                                ReqNum = Utils.Utils.getScalarValue(strRequeNo, CnnString);
                                string strPRFIX = "select PrefixText from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                                string strPRFIXCom = Utils.Utils.getScalarValue(strPRFIX, CnnString);
                                strProductionSheetNo = strPRFIXCom + strRunningNo;

                                string RTOLocationName = dtProduction.Rows[0]["RTOLocationName"].ToString();

                                #region

                                html.Append("<div style='page-break-before: avoid;page-break-inside: avoid;page-break-after: always;'>" +
                                        "<table style='width:100%;border: 0px;'>" +
                                            "<tr style='border: 0px;'>" +
                                                "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>Report Generation Date:</b> " + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</div>" +
                                                "</td>" +

                                                 "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>ROSMERTA SAFETY SYSTEMS LIMITED</b> " + "</div>" +
                                                "</td>" +


                                                 "<td colspan='5' style='border: 0px;'>" +
                                                    "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Production Sheet No:</b> " + strRunningNo + " </div>" +
                                                "</td>" +

                                            "</tr>" +

                                              "<tr style='border: 0px;'>" +
                                                "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>State:</b> " + HSRPStateName + "</div>" +
                                                "</td>" +

                                                 "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>Production Sheet</b> " + "</div>" +
                                                "</td>" +


                                                 "<td colspan='5' style='border: 0px;'>" +
                                                    "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Dealer ID/Name:</b> " + dealerid + "/" + dealername + "/" + strsubaffixid + " </div>" +
                                                "</td>" +

                                            "</tr>" +

                                                "<tr style='border: 0px;'>" +
                                                "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>EC Location: </b> " + RTOLocationName + "</div>" +
                                                "</td>" +

                                                 "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='style='text-align:left'><b>Oem :</b> " + oemname + "</div>" +
                                                "</td>" +


                                                 "<td colspan='5' style='border: 0px;'>" +
                                                     "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Dealer Address:</b> " + Address + " </div>" +
                                                //"<div style='text-align:left'><b>Dealer Address:</b> " + Address + " </div>" +
                                                "</td>" +

                                            "</tr>" +


  "<tr>" +
                                                "<td style='text-align:center'>Sr. No.</td>" +
                                                  "<td>Vehicle No.</td>" +
                                                   "<td>Front Plate Size</td>" +
                                                "<td>Front Laser No.</td>" +

                                                "<td>Rear Plate Size</td>" +
                                                "<td>Rear Laser No.</td>" +

                                                "<td>H. S. Foil </td>" +
                                                "<td>Caution Sticker</td>" +
                                                "<td>Fuel Type</td>" +
                                                "<td >VT</td>" +
                                                "<td >VC</ td>" +

                                            "</tr>");

                                #endregion

                                #region
                                foreach (DataRow drProduction in dtProduction.Rows)
                                {
                                    string HsrprecordID = drProduction["hsrprecordID"].ToString().Trim();
                                    string SRNo = drProduction["SRNo"].ToString().Trim();
                                    string ORD = "";// drProduction["ORD"].ToString().Trim();
                                    string VC = drProduction["VehicleClass"].ToString().Trim();
                                    string VehicleNo = drProduction["VehicleRegNo"].ToString().Trim();
                                   // string VehicleNo = "<td style='text-align:center;font-size:20px;'>" + "<b>" + drProduction["VehicleRegNo"].ToString().Trim() + "</b> </td>" +
                                    string VT = drProduction["VehicleType"].ToString().Trim();
                                    //string ChassisNo = drProduction["ChassisNo"].ToString().Trim();
                                    //string EngineNo = drProduction["EngineNo"].ToString().Trim();
                                    string FuelType = drProduction["FuelType"].ToString().Trim();
                                    string FrontPSize = drProduction["FrontProductCode"].ToString().Trim();
                                   // string FrontLaserNo = "<td style='text-align:center;font-size:20px;'>" + "<b>" + drProduction["HSRP_Front_LaserCode"].ToString().Trim() + "</b> </td>" +
                                   string FrontLaserNo = drProduction["HSRP_Front_LaserCode"].ToString().Trim();
                                    string RearPSize = drProduction["RearProductCode"].ToString().Trim();
                                    string RearLaserNo = drProduction["HSRP_Rear_LaserCode"].ToString().Trim();
                                    //string RearLaserNo = "<td style='text-align:center;font-size:20px;'>" + "<b>" + drProduction["HSRP_Rear_LaserCode"].ToString().Trim() + "</b> </td>" +
                                    string Amount = drProduction["roundoff_netamount"].ToString().Trim();
                                    //string OrderStatus = drProduction["OrderStatus"].ToString().Trim();
                                    string stickerColor = drProduction["stickerColor"].ToString().Trim();
                                    string HotStampingFoilColour = drProduction["HotStampingFoilColour"].ToString().Trim();

                                    html.Append("<tr>" +
                                   "<td style='text-align:center'>" + SRNo + "</td>" +
                                       "<td style='text-align:center;font-size:20px;'>" + "<b>" + VehicleNo + "</b> </td>" +

                                        "<td>" + FrontPSize + "</td>" +
                                         "<td style='text-align:center;font-size:20px;'>" + "<b>" + FrontLaserNo + "</b> </td>" +
                                           "<td>" + RearPSize + "</td>" +
                                         "<td style='text-align:center;font-size:20px;'>" + "<b>" + RearLaserNo + "</b> </td>" +

                                          "<td>" + HotStampingFoilColour + "</td>" +
                                           "<td>" + stickerColor + "</td>" +
                                            "<td>" + FuelType + "</td>" +
                                            "<td>" + VT + "</td>" +
                                            "<td>" + VC + "</td>" +





                               "</tr>");


                                    try
                                    {
                                        //start updating hsrprecords 
                                        string sqlUpdateHSRPRecords = "update hsrprecords set sendtoProductionStatus='Y', NAVPDFFlag='1', NewPdfRunningNo='" + strProductionSheetNo + "', Requisitionsheetno='" + ReqNum + "',  " +
                                              "PdfDownloadDate=GetDate(), pdfFileName='" + fileName + "', PDFDownloadUserID='1' where hsrprecordID='" + HsrprecordID + "' ";

                                        Utils.Utils.ExecNonQuery(sqlUpdateHSRPRecords, CnnString);   // uncomment after testing
                                    }
                                    catch (Exception ev)
                                    {
                                        Label1.Text = "hsrprecords update error: " + ev.Message;
                                    }
                                    //end 

                                }
                                #endregion
                                html.Append("</table>");

                                html.Append("</div>");

                                try
                                {
                                    string StrSqlUpdateECQuery = "update EmbossingCentersNew set NewProductionSheetRunningNo='" + strProductionSheetNo + "' " +
                                     "where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                                    Utils.Utils.ExecNonQuery(StrSqlUpdateECQuery, CnnString); // uncomment after testing
                                }
                                catch (Exception ev)
                                {
                                    Label1.Text = "EmbossingCentersNew update error: " + ev.Message;
                                }
                            }

                            #endregion

                        }
                    }// close oemDealerQuery
                    #endregion

                    #region "Req Generate"
                    string strComNew = string.Empty;
                    string strReqNumber = string.Empty;
                    string strReqNo = string.Empty;



                    string strSqlQuery1 = "select CompanyName from hsrpstate where hsrp_stateid='" + HSRP_StateID + "'";
                    strCompanyName = Utils.Utils.getScalarValue(strSqlQuery1, CnnString);

                    string strEMBName = " select EmbCenterName from EmbossingCentersNew where  State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                    strComNew = Utils.Utils.getScalarValue(strEMBName, CnnString);

                    if (strComNew != "")
                    {



                        strReqNo = "select (prefixtext+right('00000'+ convert(varchar,lastno+1),5)) as Reqno from prefix_Requisition  where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                        strReqNumber = Utils.Utils.getScalarValue(strReqNo, CnnString);

                        SQLString = "Exec [laserreqSlip1DashBoard]  '" + HSRP_StateID + "','" + Navembcode + "' ,  '" + ReqNum + "'";
                        DataTable dtResult = Utils.Utils.GetDataTable(SQLString, CnnString);

                        string strQuery = string.Empty;
                        string strRtoLocationName = string.Empty;
                        int Itotal = 0;

                        html.Append("<div style='width:100%;height:100%;'>" +
                                            "<table style='width:100%'>" +

                                                "<tr>" +
                                                    "<td colspan='12'>" +
                                                        "<div style='text-align:center;padding:8px;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + strCompanyName + "</b>" +

                                                            "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                             "<b style='font-size:20px;'>" + "MATERIAL REQUSITION NOTE" + "</b>" +

                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +

                                                "<tr>" +
                                                    "<td colspan='12'>" +
                                                        "<div style='text-align:center;padding:8px;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Production Sheet Date :" + DateTime.Now.ToString("dd-MM-yyyy") + "</b>" +
                                                            "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                             "<b style='font-size:20px;'>" + "" + "</b>" +
                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +
                                                "<tr>" +
                                                    "<td colspan='6'>" +
                                                        "<div style='text-align:left;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>REQ.NO:-</b>" +
                                                            "" + strReqNumber + "" +
                                                        "</div>" +
                                                    "</td>" +
                                                    "<td colspan='6'>" +
                                                        "<div style='text-align:left;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>Embossing Center:</b>" +
                                                            "" + strComNew + "" +
                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +
                                                "<tr>" +
                                                    "<td colspan='12'>" +
                                                        "<div style='text-align:left;padding:2px;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "State:" + HSRPStateName + "</b>" +
                                                            "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                             "<b style='font-size:20px;'>" + "" + "</b>" +
                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +
                                                 "<tr>" +
                                                    "<td colspan='12'>" +
                                                        "<div style='text-align:left;padding:8px;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + " " + "</b>" +
                                                            "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                             "<b style='font-size:20px;'>" + "" + "</b>" +
                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +

                                                "<tr>" +
                                                    "<td colspan='1' style='text-align:center'>SR.N.</td>" +
                                                    "<td colspan='3'>Product Size</td>" +
                                                    "<td colspan='1'>Laser Count</td>" +
                                                    "<td colspan='1'>Start Laser No</td>" +
                                                    "<td colspan='1'>End Laser No</td>" +

                                                "</tr>");


                        if (dtResult.Rows.Count > 0)
                        {
                            for (int i = 0; i < dtResult.Rows.Count; i++)
                            {
                                string ID = dtResult.Rows[i]["ID"].ToString();
                                string productcode = dtResult.Rows[i]["productcode"].ToString();
                                string LaserCount = dtResult.Rows[i]["LaserCount"].ToString();
                                Itotal = Convert.ToInt32(dtResult.Rows[i]["Total"].ToString());
                                string BeginLaser = dtResult.Rows[i]["BeginLaser"].ToString();
                                string EndLaser = dtResult.Rows[i]["EndLaser"].ToString();




                                html.Append("<tr>" +
                                   "<td colspan='1' style='text-align:left'>" + ID + "</td>" +
                                   "<td colspan='3'>" + productcode + "</td>" +

                                   "<td colspan='1'>" + LaserCount + "</td>" +
                                   "<td colspan='1'>" + BeginLaser + "</td>" +
                                   "<td colspan='1'>" + EndLaser + "</td>" +

                               "</tr>");
                            }
                        }
                        html.Append("<tr>" +
                                 "<td colspan='1' style='text-align:center' > " + "<b>Grand Total:</b>" + "</td>" +
                                 "<td colspan='3'>" + "" + "</td>" +

                                 "<td colspan='1'>" + Itotal + "</td>" +
                                 "<td colspan='1'>" + " " + "</td>" +
                                 "<td colspan='1'>" + " " + "</td>" +

                             "</tr>");




                        html.Append("<tr>" +
                         "<td colspan='2' > " + "<b>REQUESTED BY </b>" + "</td>" +
                         "<td colspan='2'>" + "<b>AUTHORIZED BY </b>" + "</td>" +

                         "<td colspan='2'>" + "<b>ISSUED BY </b>" + "</td>" +
                         "<td colspan='2'>" + "<b>RECEIVED BY</b>" + "</td>" +


                     "</tr>");

                        html.Append("<tr>" +
                                                  "<td colspan='12'>" +
                                                      "<div style='text-align:left;padding:2px;'>" +
                                                          "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Name" + "</b>" +

                                                      "</div>" +
                                                  "</td>" +
                                              "</tr>");

                        html.Append("<tr>" +
                                                 "<td colspan='12'>" +
                                                     "<div style='text-align:left;padding:2px;'>" +
                                                         "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Designation" + "</b>" +

                                                     "</div>" +
                                                 "</td>" +
                                             "</tr>");

                        html.Append("<tr>" +
                                               "<td colspan='12'>" +
                                                   "<div style='text-align:right;padding:8px;'>" +
                                                       "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Download By: Admin " + "</b>" +

                                                   "</div>" +
                                               "</td>" +
                                           "</tr>");

                        html.Append("<tr>" +
                                              "<td colspan='12'>" +
                                                  "<div style='text-align:right;padding:8px;'>" +
                                                      "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Sheet Generated By :" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</b>" +

                                                  "</div>" +
                                              "</td>" +
                                          "</tr>");







                        html.Append("</table>");

                        html.Append("</div>");

                        try
                        {
                            //start updating hsrprecords 
                            string Query = "update prefix_Requisition set lastno=lastno+1,Lastreqno=Lastreqno+1 where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                            Utils.Utils.ExecNonQuery(Query, CnnString);
                        }
                        catch (Exception ev)
                        {
                            Label1.Text = "prefix Requisition update error: " + ev.Message;
                        }
                    }


                    #endregion

                    /*
                     * Close body & HTMl Tag
                     */
                    html.Append("</body>" +
                        "</html>");


                    if (findRecord)
                    {

                        string strSaveSQlQuery = "insert into ProductionSheetAutoGenerated_List (HSRP_StateID, State_Code, RTO_ID, Emb_Center_Id, FileName, Productiondate, ProductionTime, OrderNo) values " +
                                    "('" + HSRP_StateID + "', '" + HSRPStateShortName + "', '', '" + Navembcode + "', '" + fileName + "', '" + DateTime.Now.ToString("yyyy-MM-dd") + "', '" + DateTime.Now.ToString("HH:mm:ss") + "', '" + filePrefix + "') ";
                        Utils.Utils.ExecNonQuery(strSaveSQlQuery, CnnString); // uncomment after testing
                        Label1.Text = Label1.Text + "successfully created production sheet of State: " + HSRPStateShortName + ", , Emb_Center_Id: " + Navembcode + Environment.NewLine;

                        //lblLog.Text = lblLog.Text + Environment.NewLine + "Start Production Query at:" + DateTime.Now;
                        #region
                        try
                        {
                            if (!Directory.Exists(dir))
                            {


                                Directory.CreateDirectory(dir);

                            }

                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        try
                        {
                            var pdf = Pdf
                                    .From(html.ToString())
                                    .OfSize(PaperSize.A4)
                                    .WithTitle("Title")
                                    .WithoutOutline()
                                    .WithMargins(1.25.Centimeters())
                                     .Landscape()
                                    .Comressed()
                                    .Content();

                            FileStream readStream = File.Create(filePath);
                            BinaryWriter binaryWriter = new BinaryWriter(readStream);

                            // Write the binary data to the file
                            binaryWriter.Write(pdf);
                            binaryWriter.Close();
                            readStream.Close();
                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        #endregion
                    }

                    #endregion

                }//close foreach stateEcQuery
            }
        }

        protected void btnTVS_Click(object sender, EventArgs e)
        {

            if (ddlStateName.SelectedItem.Text == "--Select State--")
            {
                lblErrMess.Text = "Please select  State";
                return;
            }
            //if (ddlStateName.SelectedValue == "31")
            //{

            //    lblErrMess.Text = String.Empty;
            //    lblErrMess.Text = "In DashBoard Uttar Pradesh state  is not allowed to generate Production sheet";
            //    return;
            //}

            string sqlPrefixQuery = "select top 1 CONVERT(INT, isnull(OrderNo,'0')) + 1 as orderNo from ProductionSheetAutoGenerated_List order by id desc";
            DataTable dtPrefix = Utils.Utils.GetDataTable(sqlPrefixQuery, CnnString);

            if (dtPrefix.Rows.Count > 0)
            {
                filePrefix = dtPrefix.Rows[0]["orderNo"].ToString();
            }
            else
            {
                filePrefix = "1";
            }

            //txtFilePrefix.Text = filePrefix;

            if (filePrefix.Length > 0)
            {
                SheetGenerationTVS();
            }
        }



        //string fileName = "TVS" + "-" + filePrefix + "-" + Navembcode + ".pdf";
        private void SheetGenerationTVS()
        {
            string stateECQuery = string.Empty;
            string ReqNum = string.Empty;
            //string Navembid = Session["Navembid"].ToString();
            FillUserDetails();


            stateECQuery = "select distinct a.HSRP_StateID, (select HSRPStateName from hsrpstate c where c.hsrp_stateid=a.hsrp_stateid) as HSRPStateName,(select HSRPStateShortName from hsrpstate d where d.hsrp_stateid=a.hsrp_stateid) as HSRPStateShortName, navembcode from hsrprecords a,rtolocation b where a.HSRP_StateID ='" + ddlStateName.SelectedValue + "'  and a.rtolocationid=b.rtolocationid and  NewPdfRunningNo is null and erpassigndate is not null  and Dealerid in(select dealerid from dealermaster where oemid='40') and OrderStatus='New Order' and    b.Navembcode not like '%CODO%' and b.NAVEMBID='" + Navembid + "'  order by  a.HSRP_StateID";


            DataTable dtSE = Utils.Utils.GetDataTable(stateECQuery, CnnString);
            if (dtSE.Rows.Count > 0)
            {
                foreach (DataRow dr in dtSE.Rows)
                {
                    #region

                    string HSRP_StateID = dr["HSRP_StateID"].ToString().Trim();
                    string HSRPStateName = dr["HSRPStateName"].ToString().Trim();
                    string HSRPStateShortName = dr["HSRPStateShortName"].ToString().Trim();
                    string Navembcode = dr["Navembcode"].ToString().Trim();

                    string dir = dirPath + DateTime.Now.ToString("yyyy-MM-dd") + "\\" + HSRPStateShortName + "\\";
                    //string dir = dirPath + "/" + DateTime.Now.ToString("yyyy-MM-dd") + "/" + HSRPStateShortName + "/";
                    string fileName = "FAL" + "-" + filePrefix + "-" + Navembcode + ".pdf";
                    string filePath = dir + fileName;

                    //string folderpath = ConfigurationManager.AppSettings["InvoiceFolder"].ToString() + "/" + FinYear + "/" + oemid + "/" + HSRPStateID + "/";

                    StringBuilder html = new StringBuilder();

                    Boolean findRecord = false;
                    string strProductionSheetNo = string.Empty;

                    /*
                     *  Start body & HTMl Tag
                     */
                    #region
                    html.Append(
                        "<!DOCTYPE html>" +
                        "<html>" +
                        "<head>" +
                            "<meta charset='UTF-8'><title>Title</title>" +
                            "<style>" +
                                "@page {" +
                                    /* headers*/
                                    "@top-left {" +
                                        "content: 'Left header';" +
                                    "}" +
                                    "@top-right {" +
                                        "content: 'Right header';" +
                                    "}" +

                                    /* footers */
                                    "@bottom-left {" +
                                        "content: 'Lorem ipsum';" +
                                    "} " +
                                    "@bottom-right {" +
                                        "content: 'Page ' counter(page) ' of ' counter(pages);" +
                                    "}" +
                                    "@bottom-center  {" +
                                        "content:element(footer);" +
                                    "}" +
                                "}" +
                                 "#footer {" +
                                    "position: running(footer);" +
                                "}" +
                                "table {" +
                                  "border-collapse: collapse;" +
                                "}" +

                                "table, th, td {" +
                                    "border: 1px solid black;" +
                                    "text-align: left;" +
                                    "vertical-align: top;" +
                                    "padding:5px;" +
                                "}" +
                            "</style>" +
                        "</head>" +
                        "<body>");
                    #endregion

                    string oemDealerQuery = string.Empty;



                    oemDealerQuery = "select distinct om.oemid as oemid, om.name as oemname, dm.dealerid, dm.dealername, dm.dealercode, dm.Address, dm.HSRP_StateID, " +
                        "dm.RTOLocationID from oemmaster om " +
                        "left join dealermaster dm on dm.oemid = om.oemid where dm.HSRP_StateID =" + HSRP_StateID + " and " +
                        "dm.RTOLocationID in (select RTOLocationID from rtolocation where Navembcode='" + Navembcode + "' ) and " +
                        "dm.dealerid in (select distinct dealerid from hsrprecords where NewPdfRunningNo is null and erpassigndate is not null and OrderStatus='New Order' and HSRP_StateID =" + HSRP_StateID + " and  Dealerid in(select dealerid from Dealermaster where oemid='40')) and Om.OEMID   in('40')";


                    #region
                    DataTable dtOD = Utils.Utils.GetDataTable(oemDealerQuery, CnnString);
                    if (dtOD.Rows.Count > 0)
                    {
                        foreach (DataRow drOD in dtOD.Rows)
                        {

                            string oemid = drOD["oemid"].ToString().Trim();
                            string dealerid = drOD["dealerid"].ToString().Trim();
                            string oemname = drOD["oemname"].ToString().Trim();
                            string dealername = drOD["dealername"].ToString().Trim();
                            string Address = drOD["Address"].ToString().Trim();

                            //start sql query
                            #region
                            string productionQuery = string.Empty;
                            DataTable dtProduction = new DataTable();




                            SqlConnection con = new SqlConnection(CnnString);
                            SqlCommand cmd = new SqlCommand("USP_TVSProductionSheet", con);
                            cmd.CommandType = CommandType.StoredProcedure;
                            con.Open();
                            cmd.Parameters.AddWithValue("@navembid", Navembcode);
                            cmd.Parameters.AddWithValue("@HSRP_StateID", HSRP_StateID);
                            cmd.Parameters.AddWithValue("@Dealerid", dealerid);
                            //cmd.Parameters.AddWithValue("@Affix_Id", strsubaffixid);
                            SqlDataAdapter da = new SqlDataAdapter(cmd);
                            // dtProduction = new DataTable();
                            da.Fill(dtProduction);
                            con.Close();


                            #endregion
                            //end sql query

                            #region
                            //DataTable dtProduction = Utils.Utils.GetDataTable(productionQuery, CnnString);
                            if (dtProduction.Rows.Count > 0)
                            {

                                findRecord = true;
                                string strRunningNo = string.Empty;
                                //string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) as maxSheetNo from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + NAVEMBID + "'";
                                string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "' ";
                                string strCom = Utils.Utils.getScalarValue(strSel, CnnString);
                                //DataTable dtSheetNo = fillDataTable(strSel, CnnString);
                                if (strCom.Equals(0) || strCom.Length == 0)
                                {
                                    strRunningNo = "0000001";
                                }
                                else
                                {
                                    strRunningNo = string.Format("{0:0000000}", Convert.ToInt32(strCom) + 1);
                                }

                                string strRequeNo = "select (prefixtext+right('00000'+ convert(varchar,Lastreqno+1),5)) as Reqno from prefix_Requisition  where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                                ReqNum = Utils.Utils.getScalarValue(strRequeNo, CnnString);

                                string strPRFIX = "select PrefixText from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                                string strPRFIXCom = Utils.Utils.getScalarValue(strPRFIX, CnnString);
                                strProductionSheetNo = strPRFIXCom + strRunningNo;

                                string RTOLocationName = dtProduction.Rows[0]["RTOLocationName"].ToString();

                                #region
                                html.Append("<div style='page-break-before: avoid;page-break-inside: avoid;page-break-after: always;'>" +
                                        "<table style='width:100%;border: 0px;'>" +
                                            "<tr style='border: 0px;'>" +
                                                "<td colspan='9' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>Report Generation Date:</b> " + DateTime.Now.ToString("dd-MM-yyyy") + "</div>" +
                                                "</td>" +
                                                "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='float:right;width: 500px;word-wrap: break-word;'>" +
                                                        "<b>Production Sheet No:</b> " + strRunningNo + "<br />" +
                                                        "<b>Dealer ID/Name:</b> " + dealerid + "/" + dealername + "<br />" +
                                                        "<b>Dealer Address:</b> " + Address + "<br />" +
                                                    "</div>" +
                                                "</td>" +
                                            "</tr>" +
                                            "<tr style='border: 0px;'>" +
                                                "<td colspan='12' style='border: 0px;'>" +
                                                    "<div style='text-align:center;font-size:26px;'><b>Future Accessories LLP Production Sheet : -</b> ROSMERTA SAFETY SYSTEMS LIMITED</div>" +
                                                "</td>" +
                                            "</tr>" +
                                            "<tr style='border: 0px;'>" +
                                                "<td colspan='9' style='border: 0px;'>" +
                                                    "<table style='border:0px;width:100%;'>" +
                                                        "<tr style='border:0px;'>" +
                                                            "<td style='border:0px;'><b>State Name :</b> " + HSRPStateName + "</td>" +
                                                            "<td style='border:0px;'><b>Oem Name :</b> " + oemname + "</td>" +
                                                            "<td style='border:0px;'><b>Report Date:</b> " + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</td>" +
                                                        "</tr>" +
                                                    "</table>" +
                                                "</td>" +
                                                "<td colspan='2' style='border: 0px;'>" +
                                                    "<div style='float:right'>" +
                                                        "ORD:Order Open Date<br />" +
                                                        "VC:Vehicle Class<br />" +
                                                        "VT:Vehicle Type<br />" +
                                                        "Front PS:Front Plate Size<br />" +
                                                        "Rear PS:Rear Plate Size<br />" +
                                                        "OS: Order Satus(New Order/Embossing Done/Closed)" +
                                                    "</div>" +
                                                "</td>" +
                                                "<td style='border: 0px;'></td>" +
                                            "</tr>" +
                                            "<tr style='border: 0px;'>" +
                                                 "<td colspan='14' style='border: 0px;'>" +
                                                    "<div style='text-align:left'>Location Name : " + RTOLocationName + " TVSPONO :  " + dtProduction.Rows[0]["TVSMPONO"].ToString() + "</div>" +
                                                "</td>" +
                                            "</tr>" +
                                            "<tr>" +
                                                "<td style='text-align:center'>SR.No</td>" +
                                                  "<td>TVSPONO</td>" +
                                                  "<td>VC</td>" +
                                                "<td>Vehicle No</td>" +
                                                "<td>VT</td>" +
                                                "<td>Chassis No</td>" +
                                                "<td>EngineNo</td>" +
                                                "<td>Fuel Type</td>" +
                                                "<td>Front PS</td>" +
                                                "<td>Front Laser No</td>" +
                                                "<td>Rear PS</td>" +
                                                "<td>Rear Laser No.</td>" +


                                                    "<td>PartNo</td>" +
                                              
                                                   "<td style='text-align:center'>OS</td>" +

                                                 "<td style='text-align:center'>Sticker Color</td>" +
                                            "</tr>");
                                #endregion

                                #region
                                string strtvspono = "";
                                int j = 0;
                                for (int i = 0; i <= dtProduction.Rows.Count - 1; i++)
                                {

                                    string HsrprecordID = dtProduction.Rows[i]["hsrprecordID"].ToString().Trim();
                                    string SRNo = dtProduction.Rows[i]["SRNo"].ToString().Trim();
                                    //if (strtvspono == "")
                                    //{
                                    //    html.Append("<tr style='border:0px;'><td colspan='14' style='border:0px;'><br/><b>TVSPONO.: " + dtProduction.Rows[0]["TVSMPONO"].ToString() + "</b></td></tr>");
                                    //}

                                    if (strtvspono.ToString() == "")
                                    {
                                        strtvspono = dtProduction.Rows[i]["TVSMPONO"].ToString();
                                        //html.Append("<tr style='border:0px;'><td colspan='14' style='border:0px;'><br/><b>TVSPONO.: " + strtvspono + "</b></td></tr>");
                                    }



                                    j = j + 1;
                                    if (strtvspono.ToString() != dtProduction.Rows[i]["TVSMPONO"].ToString())
                                    {


                                        string TVSMPONO = dtProduction.Rows[i]["TVSMPONO"].ToString().Trim();
                                        string ORD = "";// drProduction["ORD"].ToString().Trim();
                                        string VC = dtProduction.Rows[i]["VehicleClass"].ToString().Trim();
                                        string VehicleNo = dtProduction.Rows[i]["VehicleRegNo"].ToString().Trim();
                                        string VT = dtProduction.Rows[i]["VehicleType"].ToString().Trim();
                                        string ChassisNo = dtProduction.Rows[i]["ChassisNo"].ToString().Trim();
                                        string EngineNo = dtProduction.Rows[i]["EngineNo"].ToString().Trim();
                                        string FuelType = dtProduction.Rows[i]["FuelType"].ToString().Trim();
                                        string FrontPSize = dtProduction.Rows[i]["FrontProductCode"].ToString().Trim();
                                        string FrontLaserNo = dtProduction.Rows[i]["HSRP_Front_LaserCode"].ToString().Trim();
                                        string RearPSize = dtProduction.Rows[i]["RearProductCode"].ToString().Trim();
                                        string RearLaserNo = dtProduction.Rows[i]["HSRP_Rear_LaserCode"].ToString().Trim();
                                        // string Amount = drProduction["roundoff_netamount"].ToString().Trim();
                                        string PartNo = dtProduction.Rows[i]["PartNo"].ToString().Trim();

                                        string OrderStatus = dtProduction.Rows[i]["OrderStatus"].ToString().Trim();
                                        string stickerColor = dtProduction.Rows[i]["stickerColor"].ToString().Trim();
                                        j = 1;
                                        strtvspono = dtProduction.Rows[i]["TVSMPONO"].ToString();


                                        html.Append("<tr style='border:0px;'><td colspan='14' style='border:0px;'><br/><b>TVSPONO.: " + strtvspono + "</b></td></tr>");



                                        html.Append("<tr>" +
                                                "<td style='text-align:center'>SR.No</td>" +
                                                  "<td>TVSPONO</td>" +
                                                  "<td>VC</td>" +
                                                "<td>Vehicle No</td>" +
                                                "<td>VT</td>" +
                                                "<td>Chassis No</td>" +
                                                "<td>EngineNo</td>" +
                                                "<td>Fuel Type</td>" +
                                                "<td>Front PS</td>" +
                                                "<td>Front Laser No</td>" +
                                                "<td>Rear PS</td>" +
                                                "<td>Rear Laser No.</td>" +
                                                "<td>PartNo</td>" +
                                                "<td style='text-align:center'>OS</td>" +
                                                 "<td style='text-align:center'>stickerColor</td>" +

                                            "</tr>");

                                    }


                                    html.Append("<tr>" +
                                       "<td style='text-align:center'>" + j + "</td>" +
                                       "<td>" + dtProduction.Rows[i]["TVSMPONO"].ToString().Trim() + "</td>" +
                                       "<td>" + dtProduction.Rows[i]["VehicleClass"].ToString().Trim() + "</td>" +
                                                    // "<td>" + dtProduction.Rows[i]["VehicleRegNo"].ToString().Trim() + "</td>" +
                                                    "<td style='text-align:left;font-size:20px;'>" + "<b>" + dtProduction.Rows[i]["VehicleRegNo"].ToString().Trim() + "</b> </td>" +
                                       "<td>" + dtProduction.Rows[i]["VehicleType"].ToString().Trim() + "</td>" +
                                       "<td>" + dtProduction.Rows[i]["ChassisNo"].ToString().Trim() + "</td>" +
                                       "<td>" + dtProduction.Rows[i]["EngineNo"].ToString().Trim() + "</td>" +
                                       "<td>" + dtProduction.Rows[i]["FuelType"].ToString().Trim() + "</td>" +
                                       "<td>" + dtProduction.Rows[i]["FrontProductCode"].ToString().Trim() + "</td>" +
                                         //"<td>" + dtProduction.Rows[i]["HSRP_Front_LaserCode"] + "</td>" +
                                         "<td style='text-align:left;font-size:20px;'>" + "<b>" + dtProduction.Rows[i]["HSRP_Front_LaserCode"].ToString().Trim() + "</b> </td>" +
                                       "<td>" + dtProduction.Rows[i]["RearProductCode"].ToString().Trim() + "</td>" +
                                        // "<td>" + dtProduction.Rows[i]["HSRP_Rear_LaserCode"].ToString().Trim() + "</td>" +
                                        "<td style='text-align:left;font-size:20px;'>" + "<b>" + dtProduction.Rows[i]["HSRP_Rear_LaserCode"].ToString().Trim() + "</b> </td>" +

                                        "<td>" + dtProduction.Rows[i]["PartNo"].ToString().Trim() + "</td>" +
                                       "<td>" + dtProduction.Rows[i]["OrderStatus"].ToString().Trim() + "</td>" +
                                       "<td>" + dtProduction.Rows[i]["stickerColor"].ToString().Trim() + "</td>" +
                                   "</tr>");

                                    //start updating hsrprecords 
                                    string sqlUpdateHSRPRecords = "update hsrprecords set sendtoProductionStatus='Y', NAVPDFFlag='1', NewPdfRunningNo='" + strProductionSheetNo + "',  Requisitionsheetno='" + ReqNum + "',  " +
                                           "PdfDownloadDate=GetDate(), pdfFileName='" + fileName + "', PDFDownloadUserID='1' where hsrprecordID='" + HsrprecordID + "' ";
                                    Utils.Utils.ExecNonQuery(sqlUpdateHSRPRecords, CnnString);   // uncomment after testing
                                    //end 
                                }


                                #endregion
                                html.Append("</table>");

                                html.Append("</div>");

                                string StrSqlUpdateECQuery = "update EmbossingCentersNew set NewProductionSheetRunningNo='" + strProductionSheetNo + "' " +
                                 "where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                                Utils.Utils.ExecNonQuery(StrSqlUpdateECQuery, CnnString); // uncomment after testing
                            }

                            #endregion

                        }
                    }// close oemDealerQuery
                    #endregion

                    #region "Req Generate"
                    string strSqlQuery1 = "select CompanyName from hsrpstate where hsrp_stateid='" + HSRP_StateID + "'";
                    strCompanyName = Utils.Utils.getScalarValue(strSqlQuery1, CnnString);

                    string strEMBName = " select EmbCenterName from EmbossingCentersNew where  State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                    string strComNew = Utils.Utils.getScalarValue(strEMBName, CnnString);

                    string strReqNo = "select (prefixtext+right('00000'+ convert(varchar,lastno+1),5)) as Reqno from prefix_Requisition  where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                    string strReqNumber = Utils.Utils.getScalarValue(strReqNo, CnnString);

                    SQLString = "Exec [laserreqSlip1DashBoard]  '" + HSRP_StateID + "','" + Navembcode + "' ,  '" + ReqNum + "'";
                    DataTable dtResult = Utils.Utils.GetDataTable(SQLString, CnnString);
                    string strQuery = string.Empty;
                    string strRtoLocationName = string.Empty;
                    int Itotal = 0;

                    html.Append("<div style='width:100%;height:100%;'>" +
                                        "<table style='width:100%'>" +

                                            "<tr>" +
                                                "<td colspan='12'>" +
                                                    "<div style='text-align:center;padding:8px;'>" +
                                                        "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + strCompanyName + "</b>" +

                                                        "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                         "<b style='font-size:20px;'>" + "MATERIAL REQUSITION NOTE" + "</b>" +

                                                    "</div>" +
                                                "</td>" +
                                            "</tr>" +

                                            "<tr>" +
                                                "<td colspan='12'>" +
                                                    "<div style='text-align:center;padding:8px;'>" +
                                                        "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Production Sheet Date :" + DateTime.Now.ToString("dd-MM-yyyy") + "</b>" +
                                                        "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                         "<b style='font-size:20px;'>" + "" + "</b>" +
                                                    "</div>" +
                                                "</td>" +
                                            "</tr>" +
                                            "<tr>" +
                                                "<td colspan='6'>" +
                                                    "<div style='text-align:left;'>" +
                                                        "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>REQ.NO:-</b>" +
                                                        "" + strReqNumber + "" +
                                                    "</div>" +
                                                "</td>" +
                                                "<td colspan='6'>" +
                                                    "<div style='text-align:left;'>" +
                                                        "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>Embossing Center:</b>" +
                                                        "" + strComNew + "" +
                                                    "</div>" +
                                                "</td>" +
                                            "</tr>" +
                                            "<tr>" +
                                                "<td colspan='12'>" +
                                                    "<div style='text-align:left;padding:2px;'>" +
                                                        "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "State:" + HSRPStateName + "</b>" +
                                                        "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                         "<b style='font-size:20px;'>" + "" + "</b>" +
                                                    "</div>" +
                                                "</td>" +
                                            "</tr>" +
                                             "<tr>" +
                                                "<td colspan='12'>" +
                                                    "<div style='text-align:left;padding:8px;'>" +
                                                        "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + " " + "</b>" +
                                                        "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                         "<b style='font-size:20px;'>" + "" + "</b>" +
                                                    "</div>" +
                                                "</td>" +
                                            "</tr>" +

                                            "<tr>" +
                                                "<td colspan='1' style='text-align:center'>SR.N.</td>" +
                                                "<td colspan='3'>Product Size</td>" +
                                                "<td colspan='1'>Laser Count</td>" +
                                                "<td colspan='1'>Start Laser No</td>" +
                                                "<td colspan='1'>End Laser No</td>" +

                                            "</tr>");


                    if (dtResult.Rows.Count > 0)
                    {
                        for (int i = 0; i < dtResult.Rows.Count; i++)
                        {
                            string ID = dtResult.Rows[i]["ID"].ToString();
                            string productcode = dtResult.Rows[i]["productcode"].ToString();
                            string LaserCount = dtResult.Rows[i]["LaserCount"].ToString();
                            Itotal = Convert.ToInt32(dtResult.Rows[i]["Total"].ToString());
                            string BeginLaser = dtResult.Rows[i]["BeginLaser"].ToString();
                            string EndLaser = dtResult.Rows[i]["EndLaser"].ToString();




                            html.Append("<tr>" +
                               "<td colspan='1' style='text-align:left'>" + ID + "</td>" +
                               "<td colspan='3'>" + productcode + "</td>" +

                               "<td colspan='1'>" + LaserCount + "</td>" +
                               "<td colspan='1'>" + BeginLaser + "</td>" +
                               "<td colspan='1'>" + EndLaser + "</td>" +

                           "</tr>");
                        }
                    }
                    html.Append("<tr>" +
                             "<td colspan='1' style='text-align:center' > " + "<b>Grand Total:</b>" + "</td>" +
                             "<td colspan='3'>" + "" + "</td>" +

                             "<td colspan='1'>" + Itotal + "</td>" +
                             "<td colspan='1'>" + " " + "</td>" +
                             "<td colspan='1'>" + " " + "</td>" +

                         "</tr>");




                    html.Append("<tr>" +
                     "<td colspan='2' > " + "<b>REQUESTED BY </b>" + "</td>" +
                     "<td colspan='2'>" + "<b>AUTHORIZED BY </b>" + "</td>" +

                     "<td colspan='2'>" + "<b>ISSUED BY </b>" + "</td>" +
                     "<td colspan='2'>" + "<b>RECEIVED BY</b>" + "</td>" +


                 "</tr>");

                    html.Append("<tr>" +
                                              "<td colspan='12'>" +
                                                  "<div style='text-align:left;padding:2px;'>" +
                                                      "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Name" + "</b>" +

                                                  "</div>" +
                                              "</td>" +
                                          "</tr>");

                    html.Append("<tr>" +
                                             "<td colspan='12'>" +
                                                 "<div style='text-align:left;padding:2px;'>" +
                                                     "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Designation" + "</b>" +

                                                 "</div>" +
                                             "</td>" +
                                         "</tr>");

                    html.Append("<tr>" +
                                           "<td colspan='12'>" +
                                               "<div style='text-align:right;padding:8px;'>" +
                                                   "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Download By: Admin " + "</b>" +

                                               "</div>" +
                                           "</td>" +
                                       "</tr>");

                    html.Append("<tr>" +
                                          "<td colspan='12'>" +
                                              "<div style='text-align:right;padding:8px;'>" +
                                                  "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Sheet Generated By :" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</b>" +

                                              "</div>" +
                                          "</td>" +
                                      "</tr>");






                    html.Append("</table>");

                    html.Append("</div>");


                    try
                    {
                        //start updating hsrprecords 
                        string Query = "update prefix_Requisition set lastno=lastno+1,Lastreqno=Lastreqno+1 where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                        Utils.Utils.ExecNonQuery(Query, CnnString);
                    }
                    catch (Exception ev)
                    {
                        Label1.Text = "prefix Requisition update error: " + ev.Message;
                    }


                    #endregion

                    /*
                     * Close body & HTMl Tag
                     */
                    html.Append("</body>" +
                        "</html>");


                    if (findRecord)
                    {


                        string strSaveSQlQuery = "insert into ProductionSheetAutoGenerated_List (HSRP_StateID, State_Code, RTO_ID, Emb_Center_Id, FileName, Productiondate, ProductionTime, OrderNo) values " +
                                    "('" + HSRP_StateID + "', '" + HSRPStateShortName + "', '', '" + Navembcode + "', '" + fileName + "', '" + DateTime.Now.ToString("yyyy-MM-dd") + "', '" + DateTime.Now.ToString("HH:mm:ss") + "', '" + filePrefix + "') ";
                        Utils.Utils.ExecNonQuery(strSaveSQlQuery, CnnString); // uncomment after testing
                        Label1.Text = Label1.Text + "successfully created production sheet of State: " + HSRPStateShortName + ", , Emb_Center_Id: " + Navembcode + Environment.NewLine;

                        //lblLog.Text = lblLog.Text + Environment.NewLine + "Start Production Query at:" + DateTime.Now;



                        #region
                        try
                        {
                            if (!Directory.Exists(dir))
                            {


                                Directory.CreateDirectory(dir);

                            }

                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        try
                        {
                            var pdf = Pdf
                                    .From(html.ToString())
                                    .OfSize(PaperSize.A4)
                                    .WithTitle("Title")
                                    .WithoutOutline()
                                    .WithMargins(1.25.Centimeters())
                                      .Landscape()
                                    .Comressed()
                                    .Content();

                            FileStream readStream = File.Create(filePath);
                            BinaryWriter binaryWriter = new BinaryWriter(readStream);

                            // Write the binary data to the file
                            binaryWriter.Write(pdf);
                            binaryWriter.Close();
                            readStream.Close();
                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        #endregion
                    }

                    #endregion

                }//close foreach stateEcQuery
            }
        }


        protected void btnJCB_Click(object sender, EventArgs e)
        {

            if (ddlStateName.SelectedItem.Text == "--Select State--")
            {
                lblErrMess.Text = "Please select  State";
                return;
            }

            //if (ddlStateName.SelectedValue == "31")
            //{

            //    lblErrMess.Text = String.Empty;
            //    lblErrMess.Text = "In DashBoard Uttar Pradesh state  is not allowed to generate Production sheet";
            //    return;
            //}

            string sqlPrefixQuery = "select top 1 CONVERT(INT, isnull(OrderNo,'0')) + 1 as orderNo from ProductionSheetAutoGenerated_List order by id desc";
            DataTable dtPrefix = Utils.Utils.GetDataTable(sqlPrefixQuery, CnnString);

            if (dtPrefix.Rows.Count > 0)
            {
                filePrefix = dtPrefix.Rows[0]["orderNo"].ToString();
            }
            else
            {
                filePrefix = "1";
            }

            //txtFilePrefix.Text = filePrefix;

            if (filePrefix.Length > 0)
            {
                SheetGenerationJCB();
            }

        }
        // string fileName = "JCB" + "-" + filePrefix + "-" + Navembcode + ".pdf";
        private void SheetGenerationJCB()
        {
            string stateECQuery = string.Empty;
            string ReqNum = string.Empty;

            //string Navembid = Session["Navembid"].ToString();
            FillUserDetails();
            stateECQuery = "select distinct a.HSRP_StateID, (select HSRPStateName from hsrpstate c where c.hsrp_stateid=a.hsrp_stateid) as HSRPStateName,(select HSRPStateShortName from hsrpstate d where d.hsrp_stateid=a.hsrp_stateid) as HSRPStateShortName, navembcode from hsrprecords a with(nolock) join Dealeraffixation b with(nolock) on a.affix_id=b.subdealerid where a.HSRP_StateID ='" + ddlStateName.SelectedValue + "'  and  NewPdfRunningNo is null and erpassigndate is not null  and a.Dealerid in(select dealerid from dealermaster where oemid='21') and OrderStatus='New Order' and    b.Navembcode not like '%CODO%' and b.Navembcode='" + Navembid + "'  order by  a.HSRP_StateID";
            DataTable dtSE = Utils.Utils.GetDataTable(stateECQuery, CnnString);

            if (dtSE.Rows.Count > 0)
            {
                foreach (DataRow dr in dtSE.Rows)
                {
                    #region

                    string HSRP_StateID = dr["HSRP_StateID"].ToString().Trim();
                    string HSRPStateName = dr["HSRPStateName"].ToString().Trim();
                    string HSRPStateShortName = dr["HSRPStateShortName"].ToString().Trim();
                    //string RTOLocationID = dr["RTOLocationID"].ToString().Trim();
                    //string RTOLocationName = dr["RTOLocationName"].ToString().Trim();
                    //string NAVEMBID = dr["NAVEMBID"].ToString().Trim();
                    string Navembcode = dr["Navembcode"].ToString().Trim();

                    string dir = dirPath + DateTime.Now.ToString("yyyy-MM-dd") + "\\" + HSRPStateShortName + "\\";
                    // string fileName = filePrefix + "-" + Navembcode + ".pdf";
                    string fileName = "JCB" + "-" + filePrefix + "-" + Navembcode + ".pdf";
                    string filePath = dir + fileName;

                    StringBuilder html = new StringBuilder();

                    Boolean findRecord = false;
                    string strProductionSheetNo = string.Empty;

                    /*
                    *  Start body & HTMl Tag
                    */
                    #region
                    html.Append(
                        "<!DOCTYPE html>" +
                        "<html>" +
                        "<head>" +
                            "<meta charset='UTF-8'><title>Title</title>" +
                            "<style>" +
                                "@page {" +
                                    /* headers*/
                                    "@top-left {" +
                                        "content: 'Left header';" +
                                    "}" +
                                    "@top-right {" +
                                        "content: 'Right header';" +
                                    "}" +

                                    /* footers */
                                    "@bottom-left {" +
                                        "content: 'Lorem ipsum';" +
                                    "} " +
                                    "@bottom-right {" +
                                        "content: 'Page ' counter(page) ' of ' counter(pages);" +
                                    "}" +
                                    "@bottom-center  {" +
                                        "content:element(footer);" +
                                    "}" +
                                "}" +
                                 "#footer {" +
                                    "position: running(footer);" +
                                "}" +
                                "table {" +
                                  "border-collapse: collapse;" +
                                "}" +

                                "table, th, td {" +
                                    "border: 1px solid black;" +
                                    "text-align: left;" +
                                    "vertical-align: top;" +
                                    "padding:5px;" +
                                "}" +
                            "</style>" +
                        "</head>" +
                        "<body>");
                    #endregion

                    string oemDealerQuery = string.Empty;

                    oemDealerQuery = "select distinct om.oemid as oemid, om.name as oemname, dm.dealerid, d.Subdealername as Dealername, dm.dealercode, d.Address,d.SubDealerId, dm.HSRP_StateID, " +
                     "dm.RTOLocationID from oemmaster om " +
                     "left join dealermaster dm on dm.oemid = om.oemid join DealerAffixation d on d.DealerID=dm.DealerId where dm.HSRP_StateID =" + HSRP_StateID + " and " +
                      "d.RTOLocationID in (select RTOLocationID from rtolocation where Navembcode='" + Navembcode + "' ) and " +

                     "dm.dealerid in (select distinct dealerid from hsrprecords where NewPdfRunningNo is null and erpassigndate is not null and OrderStatus='New Order' and  affix_id is not NULL ) and    OM.OEMid='21'";


                    #region
                    DataTable dtOD = Utils.Utils.GetDataTable(oemDealerQuery, CnnString);

                    if (dtOD.Rows.Count > 0)
                    {
                        foreach (DataRow drOD in dtOD.Rows)
                        {

                            string oemid = drOD["oemid"].ToString().Trim();
                            string dealerid = drOD["dealerid"].ToString().Trim();
                            string oemname = drOD["oemname"].ToString().Trim();
                            string dealername = drOD["Dealername"].ToString().Trim();
                            string Address = drOD["Address"].ToString().Trim();



                            string strsubaffixid = drOD["SubDealerId"].ToString();

                            //start sql query
                            #region
                            string productionQuery = string.Empty;

                            DataTable dtProduction = new DataTable();

                           

                            SqlConnection con = new SqlConnection(CnnString);
                            SqlCommand cmd = new SqlCommand("USP_JCBProductionSheet", con);
                            cmd.CommandType = CommandType.StoredProcedure;
                            con.Open();
                            cmd.Parameters.AddWithValue("@navembid", Navembcode);
                            cmd.Parameters.AddWithValue("@HSRP_StateID", HSRP_StateID);
                            cmd.Parameters.AddWithValue("@Dealerid", dealerid);
                            cmd.Parameters.AddWithValue("@Affix_Id", strsubaffixid);
                            SqlDataAdapter da = new SqlDataAdapter(cmd);
                            // dtProduction = new DataTable();
                            da.Fill(dtProduction);
                            con.Close();


                            #endregion
                            //end sql query

                            #region
                            // DataTable dtProduction = Utils.Utils.GetDataTable(productionQuery, CnnString);
                            if (dtProduction.Rows.Count > 0)
                            {

                                findRecord = true;
                                string strRunningNo = string.Empty;
                                //string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) as maxSheetNo from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + NAVEMBID + "'";
                                string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "' ";
                                string strCom = Utils.Utils.getScalarValue(strSel, CnnString);
                                //DataTable dtSheetNo = fillDataTable(strSel, CnnString);
                                if (strCom.Equals(0) || strCom.Length == 0)
                                {
                                    strRunningNo = "0000001";
                                }
                                else
                                {
                                    strRunningNo = string.Format("{0:0000000}", Convert.ToInt32(strCom) + 1);
                                }

                                string strRequeNo = "select (prefixtext+right('00000'+ convert(varchar,Lastreqno+1),5)) as Reqno from prefix_Requisition  where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                                ReqNum = Utils.Utils.getScalarValue(strRequeNo, CnnString);

                                string strPRFIX = "select PrefixText from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                                string strPRFIXCom = Utils.Utils.getScalarValue(strPRFIX, CnnString);
                                strProductionSheetNo = strPRFIXCom + strRunningNo;

                                string RTOLocationName = dtProduction.Rows[0]["RTOLocationName"].ToString();


                                #region

                                html.Append("<div style='page-break-before: avoid;page-break-inside: avoid;page-break-after: always;'>" +
                                        "<table style='width:100%;border: 0px;'>" +
                                            "<tr style='border: 0px;'>" +
                                                "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>Report Generation Date:</b> " + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</div>" +
                                                "</td>" +

                                                 "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>ROSMERTA SAFETY SYSTEMS LIMITED</b> " + "</div>" +
                                                "</td>" +


                                                 "<td colspan='5' style='border: 0px;'>" +
                                                    "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Production Sheet No:</b> " + strRunningNo + " </div>" +
                                                "</td>" +

                                            "</tr>" +

                                              "<tr style='border: 0px;'>" +
                                                "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>State:</b> " + HSRPStateName + "</div>" +
                                                "</td>" +

                                                 "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>Production Sheet</b> " + "</div>" +
                                                "</td>" +


                                                 "<td colspan='5' style='border: 0px;'>" +
                                                    "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Dealer ID/Name:</b> " + dealerid + "/" + dealername + "/" + strsubaffixid + " </div>" +
                                                "</td>" +

                                            "</tr>" +

                                                "<tr style='border: 0px;'>" +
                                                "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>EC Location: </b> " + RTOLocationName + "</div>" +
                                                "</td>" +

                                                 "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='style='text-align:left'><b>Oem :</b> " + oemname + "</div>" +
                                                "</td>" +


                                                 "<td colspan='5' style='border: 0px;'>" +
                                                     "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Dealer Address:</b> " + Address + " </div>" +
                                                //"<div style='text-align:left'><b>Dealer Address:</b> " + Address + " </div>" +
                                                "</td>" +

                                            "</tr>" +


                                             "<tr>" +
                                                "<td style='text-align:center'>Sr. No.</td>" +
                                                  "<td>Vehicle No.</td>" +
                                                   "<td>Front Plate Size</td>" +
                                                "<td>Front Laser No.</td>" +

                                                "<td>Rear Plate Size</td>" +
                                                "<td>Rear Laser No.</td>" +

                                                "<td>H. S. Foil </td>" +
                                                "<td>Caution Sticker</td>" +
                                                "<td>Fuel Type</td>" +
                                                "<td >VT</td>" +
                                                "<td >VC</ td>" +

                                            "</tr>");




                                #endregion

                                #region
                                string strtvspono = "";
                                foreach (DataRow drProduction in dtProduction.Rows)
                                {
                                    string HsrprecordID = drProduction["hsrprecordID"].ToString().Trim();
                                    string SRNo = drProduction["SRNo"].ToString().Trim();
                                    string ORD = "";// drProduction["ORD"].ToString().Trim();
                                    string VC = drProduction["VehicleClass"].ToString().Trim();
                                    string VehicleNo = drProduction["VehicleRegNo"].ToString().Trim();
                                    string VT = drProduction["VehicleType"].ToString().Trim();
                                    //string ChassisNo = drProduction["ChassisNo"].ToString().Trim();
                                    //string EngineNo = drProduction["EngineNo"].ToString().Trim();
                                    string FuelType = drProduction["FuelType"].ToString().Trim();
                                    string FrontPSize = drProduction["FrontProductCode"].ToString().Trim();
                                    string FrontLaserNo = drProduction["HSRP_Front_LaserCode"].ToString().Trim();
                                    string RearPSize = drProduction["RearProductCode"].ToString().Trim();
                                    string RearLaserNo = drProduction["HSRP_Rear_LaserCode"].ToString().Trim();
                                    string Amount = drProduction["roundoff_netamount"].ToString().Trim();
                                    //string OrderStatus = drProduction["OrderStatus"].ToString().Trim();
                                    string stickerColor = drProduction["stickerColor"].ToString().Trim();
                                    string HotStampingFoilColour = drProduction["HotStampingFoilColour"].ToString().Trim();

                                    html.Append("<tr>" +
                                       "<td style='text-align:center'>" + SRNo + "</td>" +
                                           "<td style='text-align:center;font-size:20px;'>" + "<b>" + VehicleNo + "</b> </td>" +

                                            "<td>" + FrontPSize + "</td>" +
                                             "<td style='text-align:center;font-size:20px;'>" + "<b>" + FrontLaserNo + "</b> </td>" +
                                               "<td>" + RearPSize + "</td>" +
                                             "<td style='text-align:center;font-size:20px;'>" + "<b>" + RearLaserNo + "</b> </td>" +

                                              "<td>" + HotStampingFoilColour + "</td>" +
                                               "<td>" + stickerColor + "</td>" +
                                                "<td>" + FuelType + "</td>" +
                                                "<td>" + VT + "</td>" +
                                                "<td>" + VC + "</td>" +

                                   "</tr>");


                                    //start updating hsrprecords 
                                    string sqlUpdateHSRPRecords = "update hsrprecords set sendtoProductionStatus='Y', NAVPDFFlag='1', NewPdfRunningNo='" + strProductionSheetNo + "',  " +
                                           "PdfDownloadDate=GetDate(), pdfFileName='" + fileName + "', PDFDownloadUserID='1' where hsrprecordID='" + HsrprecordID + "' ";
                                    Utils.Utils.ExecNonQuery(sqlUpdateHSRPRecords, CnnString);   // uncomment after testing
                                    //end 

                                }



                                #endregion
                                html.Append("</table>");

                                html.Append("</div>");

                                string StrSqlUpdateECQuery = "update EmbossingCentersNew set NewProductionSheetRunningNo='" + strProductionSheetNo + "' " +
                                 "where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                                Utils.Utils.ExecNonQuery(StrSqlUpdateECQuery, CnnString); // uncomment after testing
                            }

                            #endregion

                        }
                    }// close oemDealerQuery
                    #endregion

                    //   #region "Req Generate"
                    //   string strSqlQuery1 = "select CompanyName from hsrpstate where hsrp_stateid='" + HSRP_StateID + "'";
                    //   strCompanyName = Utils.Utils.getScalarValue(strSqlQuery1, CnnString);

                    //   string strEMBName = " select EmbCenterName from EmbossingCentersNew where  State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                    //   string strComNew = Utils.Utils.getScalarValue(strEMBName, CnnString);

                    //   string strReqNo = "select (prefixtext+right('00000'+ convert(varchar,lastno+1),5)) as Reqno from prefix_Requisition  where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                    //   string strReqNumber = Utils.Utils.getScalarValue(strReqNo, CnnString);

                    //   SQLString = "Exec [laserreqSlip1DashBoard]  '" + HSRP_StateID + "','" + Navembcode + "' ,  '" + ReqNum + "'";
                    //   DataTable dtResult = Utils.Utils.GetDataTable(SQLString, CnnString);
                    //   string strQuery = string.Empty;
                    //   string strRtoLocationName = string.Empty;
                    //   int Itotal = 0;

                    //   html.Append("<div style='width:100%;height:100%;'>" +
                    //                       "<table style='width:100%'>" +

                    //                           "<tr>" +
                    //                               "<td colspan='12'>" +
                    //                                   "<div style='text-align:center;padding:8px;'>" +
                    //                                       "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + strCompanyName + "</b>" +

                    //                                       "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                    //                                        "<b style='font-size:20px;'>" + "MATERIAL REQUSITION NOTE" + "</b>" +

                    //                                   "</div>" +
                    //                               "</td>" +
                    //                           "</tr>" +

                    //                           "<tr>" +
                    //                               "<td colspan='12'>" +
                    //                                   "<div style='text-align:center;padding:8px;'>" +
                    //                                       "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Production Sheet Date :" + DateTime.Now.ToString("dd-MM-yyyy") + "</b>" +
                    //                                       "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                    //                                        "<b style='font-size:20px;'>" + "" + "</b>" +
                    //                                   "</div>" +
                    //                               "</td>" +
                    //                           "</tr>" +
                    //                           "<tr>" +
                    //                               "<td colspan='6'>" +
                    //                                   "<div style='text-align:left;'>" +
                    //                                       "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>REQ.NO:-</b>" +
                    //                                       "" + strReqNumber + "" +
                    //                                   "</div>" +
                    //                               "</td>" +
                    //                               "<td colspan='6'>" +
                    //                                   "<div style='text-align:left;'>" +
                    //                                       "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>Embossing Center:</b>" +
                    //                                       "" + strComNew + "" +
                    //                                   "</div>" +
                    //                               "</td>" +
                    //                           "</tr>" +
                    //                           "<tr>" +
                    //                               "<td colspan='12'>" +
                    //                                   "<div style='text-align:left;padding:2px;'>" +
                    //                                       "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "State:" + HSRPStateName + "</b>" +
                    //                                       "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                    //                                        "<b style='font-size:20px;'>" + "" + "</b>" +
                    //                                   "</div>" +
                    //                               "</td>" +
                    //                           "</tr>" +
                    //                            "<tr>" +
                    //                               "<td colspan='12'>" +
                    //                                   "<div style='text-align:left;padding:8px;'>" +
                    //                                       "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + " " + "</b>" +
                    //                                       "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                    //                                        "<b style='font-size:20px;'>" + "" + "</b>" +
                    //                                   "</div>" +
                    //                               "</td>" +
                    //                           "</tr>" +

                    //                           "<tr>" +
                    //                               "<td colspan='1' style='text-align:center'>SR.N.</td>" +
                    //                               "<td colspan='3'>Product Size</td>" +
                    //                               "<td colspan='1'>Laser Count</td>" +
                    //                               "<td colspan='1'>Start Laser No</td>" +
                    //                               "<td colspan='1'>End Laser No</td>" +

                    //                           "</tr>");


                    //   if (dtResult.Rows.Count > 0)
                    //   {
                    //       for (int i = 0; i < dtResult.Rows.Count; i++)
                    //       {
                    //           string ID = dtResult.Rows[i]["ID"].ToString();
                    //           string productcode = dtResult.Rows[i]["productcode"].ToString();
                    //           string LaserCount = dtResult.Rows[i]["LaserCount"].ToString();
                    //           Itotal = Convert.ToInt32(dtResult.Rows[i]["Total"].ToString());
                    //           string BeginLaser = dtResult.Rows[i]["BeginLaser"].ToString();
                    //           string EndLaser = dtResult.Rows[i]["EndLaser"].ToString();




                    //           html.Append("<tr>" +
                    //              "<td colspan='1' style='text-align:left'>" + ID + "</td>" +
                    //              "<td colspan='3'>" + productcode + "</td>" +

                    //              "<td colspan='1'>" + LaserCount + "</td>" +
                    //              "<td colspan='1'>" + BeginLaser + "</td>" +
                    //              "<td colspan='1'>" + EndLaser + "</td>" +

                    //          "</tr>");
                    //       }
                    //   }
                    //   html.Append("<tr>" +
                    //            "<td colspan='1' style='text-align:center' > " + "<b>Grand Total:</b>" + "</td>" +
                    //            "<td colspan='3'>" + "" + "</td>" +

                    //            "<td colspan='1'>" + Itotal + "</td>" +
                    //            "<td colspan='1'>" + " " + "</td>" +
                    //            "<td colspan='1'>" + " " + "</td>" +

                    //        "</tr>");




                    //   html.Append("<tr>" +
                    //    "<td colspan='2' > " + "<b>REQUESTED BY </b>" + "</td>" +
                    //    "<td colspan='2'>" + "<b>AUTHORIZED BY </b>" + "</td>" +

                    //    "<td colspan='2'>" + "<b>ISSUED BY </b>" + "</td>" +
                    //    "<td colspan='2'>" + "<b>RECEIVED BY</b>" + "</td>" +


                    //"</tr>");

                    //   html.Append("<tr>" +
                    //                             "<td colspan='12'>" +
                    //                                 "<div style='text-align:left;padding:2px;'>" +
                    //                                     "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Name" + "</b>" +

                    //                                 "</div>" +
                    //                             "</td>" +
                    //                         "</tr>");

                    //   html.Append("<tr>" +
                    //                            "<td colspan='12'>" +
                    //                                "<div style='text-align:left;padding:2px;'>" +
                    //                                    "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Designation" + "</b>" +

                    //                                "</div>" +
                    //                            "</td>" +
                    //                        "</tr>");

                    //   html.Append("<tr>" +
                    //                          "<td colspan='12'>" +
                    //                              "<div style='text-align:right;padding:8px;'>" +
                    //                                  "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Download By: Admin " + "</b>" +

                    //                              "</div>" +
                    //                          "</td>" +
                    //                      "</tr>");

                    //   html.Append("<tr>" +
                    //                         "<td colspan='12'>" +
                    //                             "<div style='text-align:right;padding:8px;'>" +
                    //                                 "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Sheet Generated By :" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</b>" +

                    //                             "</div>" +
                    //                         "</td>" +
                    //                     "</tr>");






                    //   html.Append("</table>");

                    //   html.Append("</div>");


                    //   try
                    //   {
                    //       //start updating hsrprecords 
                    //       string Query = "update prefix_Requisition set lastno=lastno+1,Lastreqno=Lastreqno+1 where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                    //       Utils.Utils.ExecNonQuery(Query, CnnString);
                    //   }
                    //   catch (Exception ev)
                    //   {
                    //       Label1.Text = "prefix Requisition update error: " + ev.Message;
                    //   }


                    //   #endregion

                    /*
                     * Close body & HTMl Tag
                     */
                    html.Append("</body>" +
                        "</html>");


                    if (findRecord)
                    {


                        string strSaveSQlQuery = "insert into ProductionSheetAutoGenerated_List (HSRP_StateID, State_Code, RTO_ID, Emb_Center_Id, FileName, Productiondate, ProductionTime, OrderNo) values " +
                                    "('" + HSRP_StateID + "', '" + HSRPStateShortName + "', '', '" + Navembcode + "', '" + fileName + "', '" + DateTime.Now.ToString("yyyy-MM-dd") + "', '" + DateTime.Now.ToString("HH:mm:ss") + "', '" + filePrefix + "') ";
                        Utils.Utils.ExecNonQuery(strSaveSQlQuery, CnnString); // uncomment after testing
                        Label1.Text = Label1.Text + "successfully created production sheet of State: " + HSRPStateShortName + ", , Emb_Center_Id: " + Navembcode + Environment.NewLine;

                        //lblLog.Text = lblLog.Text + Environment.NewLine + "Start Production Query at:" + DateTime.Now;



                        #region
                        try
                        {
                            if (!Directory.Exists(dir))
                            {


                                Directory.CreateDirectory(dir);

                            }

                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        try
                        {
                            var pdf = Pdf
                                    .From(html.ToString())
                                    .OfSize(PaperSize.A4)
                                    .WithTitle("Title")
                                    .WithoutOutline()
                                    .WithMargins(1.25.Centimeters())
                                    .Landscape()
                                    .Comressed()
                                    .Content();

                            FileStream readStream = File.Create(filePath);
                            BinaryWriter binaryWriter = new BinaryWriter(readStream);

                            // Write the binary data to the file
                            binaryWriter.Write(pdf);
                            binaryWriter.Close();
                            readStream.Close();
                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        #endregion
                    }

                    #endregion

                }//close foreach stateEcQuery
            }
        }


        protected void btnBookMyHSRP_Click(object sender, EventArgs e)
        {

            if (ddlStateName.SelectedItem.Text == "--Select State--")
            {
                lblErrMess.Text = "Please select  State";
                return;
            }
            string sqlPrefixQuery = "select top 1 CONVERT(INT, isnull(OrderNo,'0')) + 1 as orderNo from ProductionSheetAutoGenerated_List order by id desc";
            DataTable dtPrefix = Utils.Utils.GetDataTable(sqlPrefixQuery, CnnString);

            if (dtPrefix.Rows.Count > 0)
            {
                filePrefix = dtPrefix.Rows[0]["orderNo"].ToString();
            }
            else
            {
                filePrefix = "1";
            }

            //txtFilePrefix.Text = filePrefix;

            if (filePrefix.Length > 0)
            {
                SheetGenerationBookMyHSRP();
            }

        }
        //string fileName = "BookMyHSRP" + "-" + filePrefix + "-" + Navembcode + ".pdf";HSRPStateShortName
        private void SheetGenerationBookMyHSRP()
        {
            string stateECQuery = string.Empty;
            string ReqNum = string.Empty;

            FillUserDetails();

            string stateECShortName = "select distinct HSRP_StateID, 'EC'+HSRPStateShortName as NewHSRPStateShortName, HSRPStateShortName   from  HSRPState  where  HSRP_STateId='" + ddlStateName.SelectedValue + "'";

            DataTable dtECName = Utils.Utils.GetDataTable(stateECShortName, CnnString);

            string ShortECname = dtECName.Rows[0]["NewHSRPStateShortName"].ToString();
         

                stateECQuery = "select distinct a.HSRP_StateID, (select HSRPStateName from hsrpstate c where c.hsrp_stateid=a.hsrp_stateid) as HSRPStateName,(select HSRPStateShortName from hsrpstate d with(nolock) where d.hsrp_stateid='" + ddlStateName.SelectedValue + "') as HSRPStateShortName,   a.Navembid AS  navembcode from hsrprecords a with(nolock) where  NewPdfRunningNo is null and erpassigndate is not null  and (IsBookMyHsrpRecord='Y')  and OrderStatus='New Order' and a.NAVEMBID='" + Navembid + "'   order by  a.HSRP_StateID";

            DataTable dtSE = Utils.Utils.GetDataTable(stateECQuery, CnnString);


            if (dtSE.Rows.Count > 0)
            {
                foreach (DataRow dr in dtSE.Rows)
                {
                    #region

                    string HSRP_StateID = dr["HSRP_StateID"].ToString().Trim();
                    string HSRPStateName = dr["HSRPStateName"].ToString().Trim();
                    string HSRPStateShortName = dr["HSRPStateShortName"].ToString().Trim();

                    string Navembcode = dr["Navembcode"].ToString().Trim();

                    string dir = dirPath + DateTime.Now.ToString("yyyy-MM-dd") + "\\" + HSRPStateShortName + "\\";

                    string fileName = string.Empty;
                    string filePath = string.Empty;



                    StringBuilder html = new StringBuilder();

                    Boolean findRecord = false;
                    string strProductionSheetNo = string.Empty;




                    html.Append(
                    "<!DOCTYPE html>" +
                    "<html>" +
                    "<head>" +
                        "<meta charset='UTF-8'><title>Title</title>" +
                        "<style>" +
                            "@page {" +
                                /* headers*/
                                "@top-left {" +
                                    "content: 'Left header';" +
                                "}" +
                                "@top-right {" +
                                    "content: 'Right header';" +
                                "}" +

                                /* footers */
                                "@bottom-left {" +
                                    "content: 'Lorem ipsum';" +
                                "} " +
                                "@bottom-right {" +
                                    "content: 'Page ' counter(page) ' of ' counter(pages);" +
                                "}" +
                                "@bottom-center  {" +
                                    "content:element(footer);" +
                                "}" +
                            "}" +
                           
                            "#footer {" +
                                "position: running(footer);" +
                            "}" +
                            "table {" +
                              "border-collapse: collapse;" +
                            "}" +
                            "table, th, td {" +
                                "border: 1px solid black;" +
                                "text-align: left;" +
                                "vertical-align: top;" +
                                "padding-left:10px;" +
                                "padding-bottom:6px;" +
                                "padding-right:10px;" +
                                "padding-top:5px;" +

                            "}" +
                            "#main-table table,#main-table th,#main-table td{" +
                             "white-space: nowrap;}" +
                        "</style>" +
                    "</head>" +
                    "<body>");

                    #endregion

                    string fileAppoinmentDate = string.Empty;
                    string maxAppointmentdate = "select distinct  Convert(varchar(10),max(SlotBookingDate),105) as MaxAppointmentdate  " +
                    " from	HSRPrecords H with(nolock) join DealerAffixationCenter d  with(nolock) on H.HSRP_StateID=d.StateID and H.Affix_Id=d.DealerAffixationID  join BookMYHSRPappointment B on H.orderno=B.orderno  " +
                   " where   Navembid='" + Navembcode + "'      and  NewPdfRunningNo is null and erpassigndate is not null and  h.OrderStatus='New Order' and  IsBookMyHsrpRecord='Y' and  h.affix_id is not NULL  and  ((d.TypeofDelivery is null) or (d.TypeofDelivery='Dealer') or(d.TypeofDelivery='RWA')) ";
                    DataTable dtmax = Utils.Utils.GetDataTable(maxAppointmentdate, CnnString);

                    if (dtmax.Rows.Count > 0)
                    {
                        fileAppoinmentDate = dtmax.Rows[0]["MaxAppointmentdate"].ToString();
                    }
                    fileName = "MHHSRP" + "-" + fileAppoinmentDate + "-" + filePrefix + "-" + Navembcode + ".pdf";
                    filePath = dir + fileName;

                    string oemDealerQuery = string.Empty;


                    oemDealerQuery = "select distinct d.oemid as oemid, (select name  from oemmaster where oemid=d.oemid) as oemname,d.dealerid as dealerid,d.Dealername as Dealername,  d.DealerAffixationCenterAddress as Address, " +
                     "d.DealerAffixationID as SubDealerId, H.dealerid AS ParentDealerId,SlotBookingDate from	HSRPrecords H with(nolock) join DealerAffixationCenter d  with(nolock) on H.HSRP_StateID=d.StateID and H.Affix_Id=d.DealerAffixationID  join BookMYHSRPappointment B on H.orderno=B.orderno  " +
                    " where   Navembid='" + Navembcode + "'     and  NewPdfRunningNo is null and erpassigndate is not null and  h.OrderStatus='New Order' and  IsBookMyHsrpRecord='Y' and  h.affix_id is not NULL    and  ((d.TypeofDelivery is null) or (d.TypeofDelivery='Dealer') or(d.TypeofDelivery='RWA')) ";



                    #region
                    DataTable dtOD = Utils.Utils.GetDataTable(oemDealerQuery, CnnString);

                    if (dtOD.Rows.Count > 0)
                    {
                        string allStrProductionSheetNo = string.Empty;
                        Session["ECWiseProductionsheet="] = null;
                        foreach (DataRow drOD in dtOD.Rows)
                        {

                            string oemid = drOD["oemid"].ToString().Trim();
                            string dealerid = drOD["dealerid"].ToString().Trim();
                            string oemname = drOD["oemname"].ToString().Trim();
                            string dealername = drOD["Dealername"].ToString().Trim();
                            string Address = drOD["Address"].ToString().Trim();

                            DateTime AppointmentDate = Convert.ToDateTime(drOD["SlotBookingDate"].ToString().Trim());



                            string strsubaffixid = drOD["SubDealerId"].ToString();
                            string strParentDealerId = drOD["ParentDealerId"].ToString();

                            //start sql query
                            #region
                            string productionQuery = string.Empty;

                            DataTable dtProduction = new DataTable();


                            SqlConnection con = new SqlConnection(CnnString);
                            SqlCommand cmd = new SqlCommand("USP_BookMYHSRPProductionSheetNew", con);
                            cmd.CommandType = CommandType.StoredProcedure;
                            con.Open();

                            cmd.Parameters.AddWithValue("@HSRP_StateID", HSRP_StateID);
                            cmd.Parameters.AddWithValue("@navembid", Navembcode);
                            cmd.Parameters.AddWithValue("@Dealerid", strParentDealerId);
                            cmd.Parameters.AddWithValue("@Affix_Id", strsubaffixid);
                            cmd.Parameters.AddWithValue("@AppointmentDate", AppointmentDate);
                            SqlDataAdapter da = new SqlDataAdapter(cmd);
                           
                            da.Fill(dtProduction);
                            con.Close();

                         



                            #endregion
                            //end sql query

                            #region
                            //DataTable dtProduction = Utils.Utils.GetDataTable(productionQuery, CnnString);
                            if (dtProduction.Rows.Count > 0)
                            {

                                findRecord = true;
                                string strRunningNo = string.Empty;
                                //string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) as maxSheetNo from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + NAVEMBID + "'";
                                string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) from EmbossingCentersNew where    Emb_Center_Id='" + Navembcode + "' ";
                                string strCom = Utils.Utils.getScalarValue(strSel, CnnString);
                                //DataTable dtSheetNo = fillDataTable(strSel, CnnString);
                                if (strCom.Equals(0) || strCom.Length == 0)
                                {
                                    strRunningNo = "0000001";
                                }
                                else
                                {
                                    strRunningNo = string.Format("{0:0000000}", Convert.ToInt32(strCom) + 1);
                                }

                                string strRequeNo = "select (prefixtext+right('00000'+ convert(varchar,Lastreqno+1),5)) as Reqno from prefix_Requisition  where hsrp_stateid='" + ddlStateName.SelectedValue + "' and prefixfor='Req No'";
                                ReqNum = Utils.Utils.getScalarValue(strRequeNo, CnnString);

                                string strPRFIX = "select PrefixText from EmbossingCentersNew where  Emb_Center_Id='" + Navembcode + "'";
                                string strPRFIXCom = Utils.Utils.getScalarValue(strPRFIX, CnnString);
                                strProductionSheetNo = strPRFIXCom + strRunningNo;

                                string RTOLocationName = dtProduction.Rows[0]["RTOLocationName"].ToString();
                               
                                #region



                                html.Append("<div style='page-break-before: avoid;page-break-inside: avoid;page-break-after: always;'>" +

                                     "<table style='width:100%;border: 0px;'>" +
                                         "<tr  style='border: 0px;'>" +
                                             "<td colspan='4'   style='border: 0px; '>" +
                                                 "<div style='text-align:left'><b>Sheet Generation Date:</b> " + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</div>" +
                                             "</td>" +

                                              "<td colspan='3'  style='border: 0px; '>" +
                                                 "<div style='text-align:left'><b>ROSMERTA SAFETY SYSTEMS LIMITED</b> " + "</div>" +
                                             "</td>" +


                                              "<td  colspan='6' style='border: 0px;'>" +
                                                 "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Production Sheet No:</b> " + strProductionSheetNo + " </div>" +
                                             "</td>" +

                                         "</tr>" +

                                           "<tr style='border: 0px;'>" +
                                             "<td colspan='4' style='border: 0px;'>" +
                                                 "<div style='text-align:left'><b>State:</b> " + HSRPStateName + "</div>" +
                                             "</td>" +

                                              "<td colspan='3' style='border: 0px; '>" +
                                                 "<div style='text-align:left;font-size:22px;''><b>MHHSRP (Dealer Delivery) </b> " + "</div>" +
                                             "</td>" +


                                              "<td colspan='6' style='border: 0px;'>" +
                                                // "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Hub/Dealer (ID - Name):</b> " + dealerid + "/" + dealername + "/" + strsubaffixid + " </div>" +
                                                    "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Hub/Dealer (ID - Name):</b> " + dealerid + "/" + dealername +  " </div>" +
                                             "</td>" +

                                         "</tr>" +

                                             "<tr style='border: 0px;'>" +
                                             "<td colspan='4' style='border: 0px;'>" +
                                                 "<div style='text-align:left'><b>EC Location: </b> " + RTOLocationName + "</div>" +
                                             "</td>" +

                                              "<td colspan='3' style='border: 0px;'>" +
                                                "<div style='text-align:left;font-size:22px;'><b>Appointment Date: " + dtProduction.Rows[0]["AppointmentDate"].ToString() + "</ b >" + "</div>" +
                                             "</td>" +


                                              "<td  colspan='6' style='border: 0px;'>" +
                                                                "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Hub/Dealer Address:</b> " + Address + " </div>" +
                                             "</td>" +

                                         "</tr>" +




                                         "<tr>" +
                                              "<td style='text-align:center;white-space: nowrap'>Sr. No.</td>" +
                                                 "<td style='width:15%;white-space: nowrap'>Vehicle No.</td>" +
                                                 "<td>Front Plate Size</td>" +

                                                  "<td style='width:15%;white-space: nowrap'>Front Laser No.</td>" +

                                              "<td>Rear Plate Size</td>" +


                                                  "<td style='width:15%;white-space: nowrap'>Rear Laser No.</td>" +


                                              "<td style='white-space: nowrap'>H. S. Foil </td>" +
                                              "<td style='white-space: nowrap'>Caution Sticker</td>" +
                                              "<td style='white-space: nowrap'>Fuel Type</td>" +
                                              "<td style='white-space: nowrap'>VT</td>" +
                                                "<td style='white-space: nowrap'>VC</td>" +
                                              "<td style='white-space: nowrap'>Frame</ td>" +
                                                "<td style='white-space: nowrap'>Pin Code</ td>" +

                                          "</tr>");


                                #endregion

                                #region
                                string strtvspono = "";

                                int j = 0;
                                int total = 0;
                                StringBuilder UpdateSQL = new StringBuilder();
                                for (int i = 0; i <= dtProduction.Rows.Count - 1; i++)
                                {

                                    j = j + 1;

                                    if (total == 22)
                                    {
                                        total = 0;

                                        html.Append("<div style='page-break-before: avoid;page-break-inside: avoid;page-break-after: always;'>" +
                                   "<table style='width:100%;border: 0px;'>" +
                                       "<tr style='border: 0px;'>" +
                                           "<td colspan='4'  style='border: 0px;'>" +
                                               "<div style='text-align:left'><b>Sheet Generation Date:</b> " + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</div>" +
                                           "</td>" +

                                            "<td colspan='3'  style='border: 0px;'>" +
                                               "<div style='text-align:left'><b>ROSMERTA SAFETY SYSTEMS LIMITED</b> " + "</div>" +
                                           "</td>" +


                                            "<td colspan='6' style='border: 0px; '>" +
                                               "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Production Sheet No:</b> " + strProductionSheetNo + " </div>" +
                                           "</td>" +

                                       "</tr>" +

                                         "<tr style='border: 0px;'>" +
                                           "<td colspan='4' style='border: 0px;'>" +
                                               "<div style='text-align:left'><b>State:</b> " + HSRPStateName + "</div>" +
                                           "</td>" +

                                            "<td colspan='3' style='border: 0px; '>" +
                                              "<div style='text-align:left;font-size:22px;''><b> MHHSRP (Dealer Delivery) </b> " + "</div>" + "</td>" +


                                            "<td colspan='6' style='border: 0px;'>" +

                               "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Hub/Dealer (ID - Name):</b> " + dealerid + "/" + dealername + " </div>" +

                                           // "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Hub/Dealer (ID - Name):</b> " + dealerid + "/" + dealername + "/" + strsubaffixid + " </div>" +
                                           "</td>" +

                                       "</tr>" +

                                           "<tr style='border: 0px;'>" +
                                           "<td colspan='4'  style='border: 0px;'>" +
                                               "<div style='text-align:left'><b>EC Location: </b> " + RTOLocationName + "</div>" +
                                           "</td>" +

                                            "<td colspan='3' style='border: 0px;'>" +
                                              "<div style='text-align:left;font-size:22px;'><b>Appointment Date: " + dtProduction.Rows[0]["AppointmentDate"].ToString() + "</ b >" + "</div>" +
                                           "</td>" +


                                            "<td colspan='6'  style='border: 0px;'>" +
                                                              "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Hub/Dealer Address:</b> " + Address + " </div>" +
                                           "</td>" +

                                       "</tr>" +




                                          "<tr>" +
                                              "<td style='text-align:center;white-space: nowrap'>Sr. No.</td>" +
                                                 "<td style='width:15%;white-space: nowrap'>Vehicle No.</td>" +
                                                 "<td>Front Plate Size</td>" +

                                                  "<td style='width:15%;white-space: nowrap'>Front Laser No.</td>" +

                                              "<td>Rear Plate Size</td>" +


                                                  "<td style='width:15%;white-space: nowrap'>Rear Laser No.</td>" +


                                              "<td style='white-space: nowrap'>H. S. Foil </td>" +
                                              "<td style='white-space: nowrap'>Caution Sticker</td>" +
                                              "<td style='white-space: nowrap'>Fuel Type</td>" +
                                              "<td style='white-space: nowrap'>VT</td>" +
                                                "<td style='white-space: nowrap'>VC</td>" +
                                              "<td style='white-space: nowrap'>Frame</ td>" +
                                                "<td style='white-space: nowrap'>Pin Code</ td>" +

                                          "</tr>");
                                    }
                                    string FS1 = string.Empty;
                                    string FS2 = string.Empty;

                                    string RS1 = string.Empty;
                                    string RS2 = string.Empty;
                                    total = total + 1;

                                    string HsrprecordID = dtProduction.Rows[i]["hsrprecordID"].ToString().Trim();
                                    string SRNo = dtProduction.Rows[i]["SRNo"].ToString().Trim();

                                    string VC = dtProduction.Rows[i]["VehicleClass"].ToString().Trim();
                                    string VT = dtProduction.Rows[i]["VehicleType"].ToString().Trim();
                                    string VehicleNo = dtProduction.Rows[i]["VehicleRegNo"].ToString().Trim();

                                    string FuelType = dtProduction.Rows[i]["FuelType"].ToString().Trim();
                                    string FrontPSize = dtProduction.Rows[i]["FrontProductCode"].ToString().Trim();

                                    string FrontLaserNo = dtProduction.Rows[i]["HSRP_Front_LaserCode"].ToString().Trim();
                                    if (FrontLaserNo != "")
                                    {
                                        FS1 = dtProduction.Rows[i]["HSRP_Front_LaserCode"].ToString().Substring(0, 7);
                                        FS2 = dtProduction.Rows[i]["HSRP_Front_LaserCode"].ToString().Substring(7, 5);
                                    }

                                    string RearPSize = dtProduction.Rows[i]["RearProductCode"].ToString().Trim();
                                    string RearLaserNo = dtProduction.Rows[i]["HSRP_Rear_LaserCode"].ToString().Trim();
                                    if (RearLaserNo != "")
                                    {
                                        RS1 = dtProduction.Rows[i]["HSRP_Rear_LaserCode"].ToString().Substring(0, 7);
                                        RS2 = dtProduction.Rows[i]["HSRP_Rear_LaserCode"].ToString().Substring(7, 5);
                                    }



                                    string StickerColor = dtProduction.Rows[i]["stickerColor"].ToString().Trim();
                                    string HotStampingFoilColour = dtProduction.Rows[i]["HotStampingFoilColour"].ToString().Trim();
                                    string Frame = dtProduction.Rows[i]["Frame"].ToString().Trim();

                                    string Pincode = dtProduction.Rows[i]["Pincode"].ToString().Trim();

                                    html.Append("<tr>" +
                                        "<td style='text-align:center;white-space: nowrap'>" + SRNo + "</td>" +
                                        "<td style='font-size:20px;white-space: nowrap' >" + "<b>" + VehicleNo + "</td>" +
                                        "<td style='white-space: nowrap'> " + FrontPSize + "  </td> " +
                                        "<td style='font-size:20px;white-space: nowrap'>" + "<b>" + FS1 + "<b>" + FS2 + "</b> </td>" +


                                        "<td style='white-space: nowrap'>" + RearPSize + "</td>" +
                                        "<td style='font-size:20px;white-space: nowrap'>" + "<b>" + RS1 + "<b>" + RS2 + "</b> </td>" +

                                          "<td style='white-space: nowrap'>" + HotStampingFoilColour + "</td>" +
                                         "<td style='white-space: nowrap'>" + StickerColor + "</td>" +
                                         "<td style='white-space: nowrap'>" + FuelType + "</td>" +
                                         "<td style='white-space: nowrap'>" + VT + "</td>" +
                                         "<td style='white-space: nowrap'>" + VC + "</td>" +


                                        "<td style='white-space: nowrap'>" + Frame + "</td>" +

                                          "<td style='white-space: nowrap'>" + Pincode + "</td>" +
                                    "</tr>");



                                    UpdateSQL.Append("update hsrprecords set sendtoProductionStatus='Y', NAVPDFFlag='1', NewPdfRunningNo='" + strProductionSheetNo + "', Requisitionsheetno='" + ReqNum + "',PdfDownloadDate=GetDate(), pdfFileName='" + fileName + "', PDFDownloadUserID='1' where hsrprecordID='" + HsrprecordID + "';");


                                }

                                allStrProductionSheetNo += "," + strProductionSheetNo;

                                allStrProductionSheetNo = allStrProductionSheetNo.TrimStart(',');
                                Session["ECWiseProductionsheet="] = allStrProductionSheetNo;
                                #endregion
                                html.Append("</table>");

                                html.Append("</div>");
                                if (UpdateSQL.ToString().Length > 0)
                                {
                                    Utils.Utils.ExecNonQuery(UpdateSQL.ToString(), CnnString);
                                    string StrSqlUpdateECQuery = "update EmbossingCentersNew set NewProductionSheetRunningNo='" + strProductionSheetNo + "' " +
                                     "where  Emb_Center_Id='" + Navembcode + "'";
                                    Utils.Utils.ExecNonQuery(StrSqlUpdateECQuery, CnnString); // uncomment after testing
                                }
                            }

                            #endregion

                        }
                    }                                                                                                                                                                                                                                                                                                       // close oemDealerQuery
                    #endregion
                    #region "Req Generate"
                    string strSqlQuery1 = "select CompanyName from hsrpstate where hsrp_stateid='" + ddlStateName.SelectedValue + "'";
                    strCompanyName = Utils.Utils.getScalarValue(strSqlQuery1, CnnString);

                    string strEMBName = " select EmbCenterName from EmbossingCentersNew where   Emb_Center_Id='" + Navembcode + "'";
                    string strComNew = Utils.Utils.getScalarValue(strEMBName, CnnString);

                    string strReqNo = "select (prefixtext+right('00000'+ convert(varchar,Lastreqno+1),5)) as Reqno from prefix_Requisition  where hsrp_stateid='" + ddlStateName.SelectedValue + "' and prefixfor='Req No'";
                    string strReqNumber = Utils.Utils.getScalarValue(strReqNo, CnnString);

                    SQLString = "Exec [laserreqSlip1DashBoardBookMYHSRP]  '" + Navembcode + "' ,  '" + ReqNum + "'";
                    DataTable dtResult = Utils.Utils.GetDataTable(SQLString, CnnString);
                    string strQuery = string.Empty;
                    string strRtoLocationName = string.Empty;
                    int Itotal = 0;
                    // <div style="break-after:page"></div>

                    // style='float:right;width: 500px;word-wrap: break-word;'
                    //html.Append("<div style='page-break-before: avoid;page-break-inside: avoid;page-break-after: always;'>" +
                    html.Append("<div style='width:100%;'>" +
                                        "<table style='width:100%'>" +

                                            "<tr>" +
                                                "<td colspan='12'>" +
                                                    "<div style='text-align:center;padding:8px;'>" +
                                                        "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + strCompanyName + "</b>" +

                                                        "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                         "<b style='font-size:20px;'>" + "MATERIAL REQUSITION NOTE" + "</b>" +

                                                    "</div>" +
                                                "</td>" +
                                            "</tr>" +

                                            "<tr>" +
                                                "<td colspan='12'>" +
                                                    "<div style='text-align:center;padding:8px;'>" +
                                                        "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Production Sheet Date :" + DateTime.Now.ToString("dd-MM-yyyy") + "</b>" +
                                                        "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                         "<b style='font-size:20px;'>" + "" + "</b>" +
                                                    "</div>" +
                                                "</td>" +
                                            "</tr>" +
                                            "<tr>" +
                                                "<td colspan='6'>" +
                                                    "<div style='text-align:left;'>" +
                                                        "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>REQ.NO:-</b>" +
                                                        "" + strReqNumber + "" +
                                                    "</div>" +
                                                "</td>" +
                                                "<td colspan='6'>" +
                                                    "<div style='text-align:left;'>" +
                                                        "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>Embossing Center:</b>" +
                                                        "" + strComNew + "" +
                                                    "</div>" +
                                                "</td>" +
                                            "</tr>" +
                                            "<tr>" +
                                                "<td colspan='12'>" +
                                                    "<div style='text-align:left;padding:2px;'>" +
                                                        "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "State:" + ddlStateName.SelectedItem.Text + "</b>" +
                                                        "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                         "<b style='font-size:20px;'>" + "" + "</b>" +
                                                    "</div>" +
                                                "</td>" +
                                            "</tr>" +
                                             "<tr>" +
                                                "<td colspan='12'>" +
                                                    "<div style='text-align:left;padding:8px;'>" +
                                                        "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + " " + "</b>" +
                                                        "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                         "<b style='font-size:20px;'>" + "" + "</b>" +
                                                    "</div>" +
                                                "</td>" +
                                            "</tr>" +

                                            "<tr>" +
                                                "<td colspan='1' style='text-align:center'>SR.N.</td>" +
                                                "<td colspan='3'>Product Size</td>" +
                                                "<td colspan='1'>Laser Count</td>" +
                                                "<td colspan='1'>Start Laser No</td>" +
                                                "<td colspan='1'>End Laser No</td>" +

                                            "</tr>");


                    if (dtResult.Rows.Count > 0)
                    {
                        for (int i = 0; i < dtResult.Rows.Count; i++)
                        {
                            string ID = dtResult.Rows[i]["ID"].ToString();
                            string productcode = dtResult.Rows[i]["productcode"].ToString();
                            string LaserCount = dtResult.Rows[i]["LaserCount"].ToString();
                            Itotal = Convert.ToInt32(dtResult.Rows[i]["Total"].ToString());
                            string BeginLaser = dtResult.Rows[i]["BeginLaser"].ToString();
                            string EndLaser = dtResult.Rows[i]["EndLaser"].ToString();




                            html.Append("<tr>" +
                               "<td colspan='1' style='text-align:left'>" + ID + "</td>" +
                               "<td colspan='3'>" + productcode + "</td>" +

                               "<td colspan='1'>" + LaserCount + "</td>" +
                               "<td colspan='1'>" + BeginLaser + "</td>" +
                               "<td colspan='1'>" + EndLaser + "</td>" +

                           "</tr>");
                        }


                        try
                        {
                            //start updating hsrprecords 
                            string Query = "update prefix_Requisition set lastno=lastno+1,Lastreqno=Lastreqno+1 where hsrp_stateid='" + ddlStateName.SelectedValue + "' and prefixfor='Req No'";
                            Utils.Utils.ExecNonQuery(Query, CnnString);
                        }
                        catch (Exception ev)
                        {
                            Label1.Text = "prefix Requisition update error: " + ev.Message;
                        }

                    }
                    html.Append("<tr>" +
                             "<td colspan='1' style='text-align:center' > " + "<b>Grand Total:</b>" + "</td>" +
                             "<td colspan='3'>" + "" + "</td>" +

                             "<td colspan='1'>" + Itotal + "</td>" +
                             "<td colspan='1'>" + " " + "</td>" +
                             "<td colspan='1'>" + " " + "</td>" +

                         "</tr>");




                    html.Append("<tr>" +
                     "<td colspan='2' > " + "<b>REQUESTED BY </b>" + "</td>" +
                     "<td colspan='2'>" + "<b>AUTHORIZED BY </b>" + "</td>" +

                     "<td colspan='2'>" + "<b>ISSUED BY </b>" + "</td>" +
                     "<td colspan='2'>" + "<b>RECEIVED BY</b>" + "</td>" +


                 "</tr>");

                    html.Append("<tr>" +
                                              "<td colspan='12'>" +
                                                  "<div style='text-align:left;padding:2px;'>" +
                                                      "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Name" + "</b>" +

                                                  "</div>" +
                                              "</td>" +
                                          "</tr>");

                    html.Append("<tr>" +
                                             "<td colspan='12'>" +
                                                 "<div style='text-align:left;padding:2px;'>" +
                                                     "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Designation" + "</b>" +

                                                 "</div>" +
                                             "</td>" +
                                         "</tr>");

                    html.Append("<tr>" +
                                           "<td colspan='12'>" +
                                               "<div style='text-align:right;padding:8px;'>" +
                                                   "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Download By: Admin " + "</b>" +

                                               "</div>" +
                                           "</td>" +
                                       "</tr>");

                    html.Append("<tr>" +
                                          "<td colspan='12'>" +
                                              "<div style='text-align:right;padding:8px;'>" +
                                                  "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Sheet Generated By :" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</b>" +

                                              "</div>" +
                                          "</td>" +
                                      "</tr>");

                    html.Append("</table>");

                    html.Append("</div>");



                    html.Append("<tr style='visibility: hidden;'>" + "<td  ><p style=\"page-break-after:always\"/></td></tr>");







                    #region "PS Summary Report"
                    string PS = string.Empty;
                    if (Session["ECWiseProductionsheet="] != null)
                    {
                        PS = Session["ECWiseProductionsheet="].ToString();
                    }
                    /*
                     * Close body & HTMl Tag
                     */

                    SQLString = "Exec [USP_BMPSSummary]  '" + PS + "' ";
                    DataTable dtResultsummary = Utils.Utils.GetDataTable(SQLString, CnnString);

                    html.Append("<div style='width:100%;height:100%;'>" +
                                 "<table style='width:100%'>" +

                                     "<tr>" +
                                         "<td colspan='12'>" +
                                             "<div style='text-align:center;padding:8px;'>" +
                                                 "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Production Summary Report" + "</b>" +

                                             "</div>" +
                                         "</td>" +
                                     "</tr>" +

                                      "<tr>" +
                                                "<td colspan='1' style='text-align:center'>PS NO</td>" +
                                                "<td colspan='3'>Affixation Dealer Name</td>" +
                                                "<td colspan='1'>2W</td>" +
                                                "<td colspan='1'>3W</td>" +
                                                "<td colspan='1'>4W</td>" +
                                                "<td colspan='1'>OTH</td>" +
                                                  "<td colspan='1'>Total</td>" +

                                            "</tr>");

                    if (dtResultsummary.Rows.Count > 0)
                    {
                        int j = 0;
                        int total = 0;
                        StringBuilder UpdateSQL = new StringBuilder();
                        for (int i = 0; i <= dtResultsummary.Rows.Count - 1; i++)
                        {

                            j = j + 1;

                            if (total == 26)
                            {
                                total = 0;

                                html.Append("<div style='width:100%;height:100%;'>" +
                       "<table style='width:100%'>" +

                           "<tr>" +
                               "<td colspan='12'>" +
                                   "<div style='text-align:center;padding:8px;'>" +
                                       "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Production Summary Report" + "</b>" +

                                   "</div>" +
                               "</td>" +
                           "</tr>" +

                            "<tr>" +
                                      "<td colspan='1' style='text-align:center'>PS NO</td>" +
                                      "<td colspan='3'>Affixation Dealer Name</td>" +
                                      "<td colspan='1'>2W</td>" +
                                      "<td colspan='1'>3W</td>" +
                                      "<td colspan='1'>4W</td>" +
                                      "<td colspan='1'>OTH</td>" +
                                        "<td colspan='1'>Total</td>" +

                                  "</tr>");

                            }

                            total = total + 1;

                            //for (int i = 0; i < dtResultsummary.Rows.Count; i++)
                            //{
                            string PSNo = dtResultsummary.Rows[i]["PSNo"].ToString();
                            string AffixationdealerName = dtResultsummary.Rows[i]["AffixationdealerName"].ToString();
                            string W2 = dtResultsummary.Rows[i]["2W"].ToString();
                            string W3 = dtResultsummary.Rows[i]["3W"].ToString();
                            string W4 = dtResultsummary.Rows[i]["4W"].ToString();
                            string OTH = dtResultsummary.Rows[i]["OTH"].ToString();
                            string Total = dtResultsummary.Rows[i]["Total"].ToString();




                            html.Append("<tr>" +
                               "<td colspan='1' style='text-align:left'>" + PSNo + "</td>" +
                               "<td colspan='3'>" + AffixationdealerName + "</td>" +

                               "<td colspan='1'>" + W2 + "</td>" +
                               "<td colspan='1'>" + W3 + "</td>" +
                               "<td colspan='1'>" + W4 + "</td>" +
                                "<td colspan='1'>" + OTH + "</td>" +
                                 "<td colspan='1'>" + Total + "</td>" +

                           "</tr>");




                        }
                    }

                    html.Append("</table>");





                    html.Append("</div>");
                    html.Append("<div style='text-align:left;padding:8px;'>" +
                                            "<b style='font-size:25px;margin-top:2px;margin-bottom:2px;'>" + "<p>&#x25CF The Owner has to furnish the original FIR/SDE copy to the Fitment center at the time of fitment of HSRP.</p><p>&#x25CF The Owner has to deposit the damaged plate at the fitment center at the time of fitment of HSRP.</p><p>&#x25CF The fitment center has to retail the old TV/NTV plate in case of fitment due to conversion,re-assignment.</p>" + "</b>" +

                                        "</div>"
);
                    #endregion

                    html.Append("</body>" +
                        "</html>");



                    #endregion



                    if (findRecord)
                    {




                        string strSaveSQlQuery = "insert into ProductionSheetAutoGenerated_List (HSRP_StateID, State_Code, RTO_ID, Emb_Center_Id, FileName, Productiondate, ProductionTime, OrderNo) values " +
                                    "('" + ddlStateName.SelectedValue + "', '" + HSRPStateShortName + "', '', '" + Navembcode + "', '" + fileName + "', '" + DateTime.Now.ToString("yyyy-MM-dd") + "', '" + DateTime.Now.ToString("HH:mm:ss") + "', '" + filePrefix + "') ";
                        Utils.Utils.ExecNonQuery(strSaveSQlQuery, CnnString); // uncomment after testing
                        Label1.Text = Label1.Text + "successfully created production sheet of State: " + HSRPStateShortName + ", , Emb_Center_Id: " + Navembcode + Environment.NewLine;

                        //lblLog.Text = lblLog.Text + Environment.NewLine + "Start Production Query at:" + DateTime.Now;



                        #region
                        try
                        {
                            if (!Directory.Exists(dir))
                            {


                                Directory.CreateDirectory(dir);

                            }

                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        try
                        {
                            var pdf = Pdf
                                    .From(html.ToString())
                                    .OfSize(PaperSize.A4)
                                    .WithTitle("Title")
                                    .WithoutOutline()
                                    .WithMargins(1.25.Centimeters())
                                    .Landscape()
                                    .Comressed()
                                    .Content();

                            FileStream readStream = File.Create(filePath);
                            BinaryWriter binaryWriter = new BinaryWriter(readStream);

                            // Write the binary data to the file
                            binaryWriter.Write(pdf);
                            binaryWriter.Close();
                            readStream.Close();
                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        #endregion
                    }

                    //#endregion

                }//close foreach stateEcQuery
            }
        }


        protected void btnExternal_Click(object sender, EventArgs e)
        {
            if (ddlStateName.SelectedItem.Text == "--Select State--")
            {
                lblErrMess.Text = "Please select  State";
                return;
            }
            //if (ddlStateName.SelectedValue == "31")
            //{

            //    lblErrMess.Text = String.Empty;
            //    lblErrMess.Text = "In DashBoard Uttar Pradesh state  is not allowed to generate Production sheet";
            //    return;
            //}

            string sqlPrefixQuery = "select top 1 CONVERT(INT, isnull(OrderNo,'0')) + 1 as orderNo from ProductionSheetAutoGenerated_List order by id desc";
            DataTable dtPrefix = Utils.Utils.GetDataTable(sqlPrefixQuery, CnnString);

            if (dtPrefix.Rows.Count > 0)
            {
                filePrefix = dtPrefix.Rows[0]["orderNo"].ToString();
            }
            else
            {
                filePrefix = "1";
            }

            //txtFilePrefix.Text = filePrefix;

            if (filePrefix.Length > 0)
            {

                SheetGenerationExternal();

            }


        }
        private void SheetGenerationExternal()
        {
            string stateECQuery = string.Empty;
            string ReqNum = string.Empty;

            //string Navembid = Session["Navembid"].ToString();
            FillUserDetails();
            stateECQuery = "select distinct a.HSRP_StateID, (select HSRPStateName from hsrpstate c where c.hsrp_stateid=a.hsrp_stateid) as HSRPStateName,(select HSRPStateShortName from hsrpstate d where d.hsrp_stateid=a.hsrp_stateid) as HSRPStateShortName, navembcode from hsrprecords a,rtolocation b where a.HSRP_StateID ='" + ddlStateName.SelectedValue + "' and   b.Navembcode not like '%CODO%' and a.dealerid in(select dealerid from dealermaster where oemid='433')    and a.rtolocationid=b.rtolocationid and NewPdfRunningNo is null and erpassigndate is not null and OrderStatus='New Order'  AND b.NAVEMBID='" + Navembid + "'   order by  a.HSRP_StateID";


            DataTable dtSE = Utils.Utils.GetDataTable(stateECQuery, CnnString);
            if (dtSE.Rows.Count > 0)
            {
                foreach (DataRow dr in dtSE.Rows)
                {
                    #region

                    string HSRP_StateID = dr["HSRP_StateID"].ToString().Trim();
                    string HSRPStateName = dr["HSRPStateName"].ToString().Trim();
                    string HSRPStateShortName = dr["HSRPStateShortName"].ToString().Trim();
                    string Navembcode = dr["Navembcode"].ToString().Trim();

                    string dir = dirPath + DateTime.Now.ToString("yyyy-MM-dd") + "\\" + HSRPStateShortName + "\\";
                    //string dir = dirPath + "/" + DateTime.Now.ToString("yyyy-MM-dd") + "/" + HSRPStateShortName + "/";
                    string fileName = "External" + "-" + filePrefix + "-" + Navembcode + ".pdf";
                    string filePath = dir + fileName;

                    //string folderpath = ConfigurationManager.AppSettings["InvoiceFolder"].ToString() + "/" + FinYear + "/" + oemid + "/" + HSRPStateID + "/";

                    StringBuilder html = new StringBuilder();

                    Boolean findRecord = false;
                    string strProductionSheetNo = string.Empty;

                    /*
                     *  Start body & HTMl Tag
                     */
                    #region
                    html.Append(
                        "<!DOCTYPE html>" +
                        "<html>" +
                        "<head>" +
                            "<meta charset='UTF-8'><title>Title</title>" +
                            "<style>" +
                                "@page {" +
                                    /* headers*/
                                    "@top-left {" +
                                        "content: 'Left header';" +
                                    "}" +
                                    "@top-right {" +
                                        "content: 'Right header';" +
                                    "}" +

                                    /* footers */
                                    "@bottom-left {" +
                                        "content: 'Lorem ipsum';" +
                                    "} " +
                                    "@bottom-right {" +
                                        "content: 'Page ' counter(page) ' of ' counter(pages);" +
                                    "}" +
                                    "@bottom-center  {" +
                                        "content:element(footer);" +
                                    "}" +
                                "}" +
                                 "#footer {" +
                                    "position: running(footer);" +
                                "}" +
                                "table {" +
                                  "border-collapse: collapse;" +
                                "}" +

                                "table, th, td {" +
                                    "border: 1px solid black;" +
                                    "text-align: left;" +
                                    "vertical-align: top;" +
                                    "padding:5px;" +
                                "}" +
                            "</style>" +
                        "</head>" +
                        "<body>");
                    #endregion

                    string oemDealerQuery = string.Empty;



                    //oemDealerQuery = "select distinct om.oemid as oemid, om.name as oemname, dm.dealerid, dm.dealername, dm.dealercode, dm.Address, dm.HSRP_StateID, " +
                    //    "dm.RTOLocationID from oemmaster om " +
                    //    "left join dealermaster dm on dm.oemid = om.oemid where dm.HSRP_StateID =" + HSRP_StateID + " and " +
                    //    "dm.RTOLocationID in (select RTOLocationID from rtolocation where Navembcode='" + Navembcode + "' ) and " +
                    //    "dm.dealerid in (select distinct dealerid from hsrprecords where NewPdfRunningNo is null and erpassigndate is not null and OrderStatus='New Order'and HSRP_StateID =" + HSRP_StateID + ") and Om.OEMID not  in('21','40','12','20')";
                    //"dm.dealerid in (select distinct dealerid from hsrprecords where isnull(NewPdfRunningNo,'') = '' and isnull(erpassigndate,'') != '' and HSRP_StateID =" + HSRP_StateID + ") and Om.OEMID not  in('21','40','12','20')";

                    #region
                    //DataTable dtOD = Utils.Utils.GetDataTable(oemDealerQuery, CnnString);
                    DataTable dtOD;
                    SqlConnection con = new SqlConnection(CnnString);
                    SqlCommand cmd = new SqlCommand("USP_BindExternalDataDashBoard", con);
                    cmd.CommandType = CommandType.StoredProcedure;
                    con.Open();
                    cmd.Parameters.AddWithValue("@navembid", Navembcode);
                    cmd.Parameters.AddWithValue("@HSRP_StateID", HSRP_StateID);
                    //cmd.Parameters.AddWithValue("@Orderdate", orderDate);
                    //cmd.Parameters.AddWithValue("@AuthProd", Authdate);
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    dtOD = new DataTable();
                    da.Fill(dtOD);
                    con.Close();
                    if (dtOD.Rows.Count > 0)
                    {
                        foreach (DataRow drOD in dtOD.Rows)
                        {

                            //string oemid = drOD["oemid"].ToString().Trim();
                            string Contact = drOD["Contact"].ToString().Trim();
                            string oemname = drOD["OemName"].ToString().Trim();
                            string dealername = drOD["DealerName"].ToString().Trim();
                            string Address = drOD["Address"].ToString().Trim();
                            string deliveryaddress = drOD["DeliveryAddress"].ToString();

                            //start sql query
                            #region
                            string productionQuery = string.Empty;


                            DataTable dtProduction;
                            SqlConnection con1 = new SqlConnection(CnnString);
                            SqlCommand cmd1 = new SqlCommand("USP_BindExternalDataProductionsheetDashBoard", con1);
                            cmd1.CommandType = CommandType.StoredProcedure;
                            con1.Open();
                            cmd1.Parameters.AddWithValue("@navembid", Navembcode);
                            cmd1.Parameters.AddWithValue("@HSRP_StateID", HSRP_StateID);
                            // cmd1.Parameters.AddWithValue("@Orderdate", orderDate);
                            //cmd1.Parameters.AddWithValue("@AuthProd", Authdate);
                            cmd1.Parameters.AddWithValue("@DeliveryAddress", deliveryaddress);
                            SqlDataAdapter sda = new SqlDataAdapter(cmd1);
                            dtProduction = new DataTable();
                            sda.Fill(dtProduction);
                            con1.Close();

                           


                            #endregion
                            //end sql query

                            #region
                            //  dtProduction = Utils.Utils.GetDataTable(productionQuery, CnnString);
                            if (dtProduction.Rows.Count > 0)
                            {

                                findRecord = true;
                                string strRunningNo = string.Empty;
                                //string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) as maxSheetNo from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + NAVEMBID + "'";
                                string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                                string strCom = Utils.Utils.getScalarValue(strSel, CnnString);
                                //DataTable dtSheetNo = fillDataTable(strSel, CnnString);
                                if (strCom.Equals(0) || strCom.Length == 0)
                                {
                                    strRunningNo = "0000001";
                                }
                                else
                                {
                                    strRunningNo = string.Format("{0:0000000}", Convert.ToInt32(strCom) + 1);
                                }
                                string strRequeNo = "select (prefixtext+right('00000'+ convert(varchar,Lastreqno+1),5)) as Reqno from prefix_Requisition  where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                                ReqNum = Utils.Utils.getScalarValue(strRequeNo, CnnString);
                                string strPRFIX = "select PrefixText from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                                string strPRFIXCom = Utils.Utils.getScalarValue(strPRFIX, CnnString);
                                strProductionSheetNo = strPRFIXCom + strRunningNo;

                                string RTOLocationName = dtProduction.Rows[0]["RTOLocationName"].ToString();

                                #region

                                html.Append("<div style='page-break-before: avoid;page-break-inside: avoid;page-break-after: always;'>" +
                                        "<table style='width:100%;border: 0px;'>" +
                                            "<tr style='border: 0px;'>" +
                                                "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>Report Generation Date:</b> " + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</div>" +
                                                "</td>" +

                                                 "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>ROSMERTA SAFETY SYSTEMS LIMITED</b> " + "</div>" +
                                                "</td>" +


                                                 "<td colspan='5' style='border: 0px;'>" +
                                                    "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Production Sheet No:</b> " + strRunningNo + " </div>" +
                                                "</td>" +

                                            "</tr>" +

                                              "<tr style='border: 0px;'>" +
                                                "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>State:</b> " + HSRPStateName + "</div>" +
                                                "</td>" +

                                                 "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>Production Sheet</b> " + "</div>" +
                                                "</td>" +


                                                 "<td colspan='5' style='border: 0px;'>" +
                                                    "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Contact Details:</b> " + Contact +  " </div>" +
                                                "</td>" +

                                            "</tr>" +

                                                "<tr style='border: 0px;'>" +
                                                "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>EC Location: </b> " + RTOLocationName + "</div>" +
                                                "</td>" +

                                                 "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='style='text-align:left'><b>Oem :</b> " + oemname + "</div>" +
                                                "</td>" +


                                                 "<td colspan='5' style='border: 0px;'>" +
                                                     "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Dealer Address:</b> " + Address + " </div>" +
                                                //"<div style='text-align:left'><b>Dealer Address:</b> " + Address + " </div>" +
                                                "</td>" +

                                            "</tr>" +


  "<tr>" +
                                                "<td style='text-align:center'>Sr. No.</td>" +
                                                  "<td>Vehicle No.</td>" +
                                                   "<td>Front Plate Size</td>" +
                                                "<td>Front Laser No.</td>" +

                                                "<td>Rear Plate Size</td>" +
                                                "<td>Rear Laser No.</td>" +

                                                "<td>H. S. Foil </td>" +
                                                "<td>Caution Sticker</td>" +
                                                "<td>Fuel Type</td>" +
                                                "<td >VT</td>" +
                                                "<td >VC</ td>" +

                                            "</tr>");

                                #endregion

                                #region
                                foreach (DataRow drProduction in dtProduction.Rows)
                                {
                                    string HsrprecordID = drProduction["hsrprecordID"].ToString().Trim();
                                    string SRNo = drProduction["SRNo"].ToString().Trim();
                                    string ORD = "";// drProduction["ORD"].ToString().Trim();
                                    string VC = drProduction["VehicleClass"].ToString().Trim();
                                    string VehicleNo = drProduction["VehicleRegNo"].ToString().Trim();
                                    string VT = drProduction["VehicleType"].ToString().Trim();
                                    //string ChassisNo = drProduction["ChassisNo"].ToString().Trim();
                                    //string EngineNo = drProduction["EngineNo"].ToString().Trim();
                                    string FuelType = drProduction["FuelType"].ToString().Trim();
                                    string FrontPSize = drProduction["FrontProductCode"].ToString().Trim();
                                    string FrontLaserNo = drProduction["HSRP_Front_LaserCode"].ToString().Trim();
                                    string RearPSize = drProduction["RearProductCode"].ToString().Trim();
                                    string RearLaserNo = drProduction["HSRP_Rear_LaserCode"].ToString().Trim();
                                    string Amount = drProduction["roundoff_netamount"].ToString().Trim();
                                   // string OrderStatus = drProduction["OrderStatus"].ToString().Trim();
                                    string stickerColor = drProduction["stickerColor"].ToString().Trim();
                                    string HotStampingFoilColour = drProduction["HotStampingFoilColour"].ToString().Trim();

                                    html.Append("<tr>" +
                                        "<td style='text-align:center'>" + SRNo + "</td>" +
                                            "<td style='text-align:center;font-size:20px;'>" + "<b>" + VehicleNo + "</b> </td>" +

                                             "<td>" + FrontPSize + "</td>" +
                                              "<td style='text-align:center;font-size:20px;'>" + "<b>" + FrontLaserNo + "</b> </td>" +
                                                "<td>" + RearPSize + "</td>" +
                                              "<td style='text-align:center;font-size:20px;'>" + "<b>" + RearLaserNo + "</b> </td>" +

                                               "<td>" + HotStampingFoilColour + "</td>" +
                                                "<td>" + stickerColor + "</td>" +
                                                 "<td>" + FuelType + "</td>" +
                                                 "<td>" + VT + "</td>" +
                                                 "<td>" + VC + "</td>" +
                                    "</tr>");

                                    try
                                    {
                                        //start updating hsrprecords 
                                        string sqlUpdateHSRPRecords = "update hsrprecords set sendtoProductionStatus='Y', NAVPDFFlag='1', NewPdfRunningNo='" + strProductionSheetNo + "', Requisitionsheetno='" + ReqNum + "',  " +
                                              "PdfDownloadDate=GetDate(), pdfFileName='" + fileName + "', PDFDownloadUserID='1' where hsrprecordID='" + HsrprecordID + "' ";

                                        Utils.Utils.ExecNonQuery(sqlUpdateHSRPRecords, CnnString);   // uncomment after testing
                                    }
                                    catch (Exception ev)
                                    {
                                        Label1.Text = "hsrprecords update error: " + ev.Message;
                                    }
                                    //end 

                                }
                                #endregion
                                html.Append("</table>");

                                html.Append("</div>");

                                try
                                {
                                    string StrSqlUpdateECQuery = "update EmbossingCentersNew set NewProductionSheetRunningNo='" + strProductionSheetNo + "' " +
                                     "where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                                    Utils.Utils.ExecNonQuery(StrSqlUpdateECQuery, CnnString); // uncomment after testing
                                }
                                catch (Exception ev)
                                {
                                    Label1.Text = "EmbossingCentersNew update error: " + ev.Message;
                                }
                            }

                            #endregion

                        }
                    }// close oemDealerQuery
                    #endregion

                    #region "Req Generate"
                    string strSqlQuery1 = "select CompanyName from hsrpstate where hsrp_stateid='" + HSRP_StateID + "'";
                    strCompanyName = Utils.Utils.getScalarValue(strSqlQuery1, CnnString);

                    string strEMBName = " select EmbCenterName from EmbossingCentersNew where  State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                    string strComNew = Utils.Utils.getScalarValue(strEMBName, CnnString);

                    string strReqNo = "select (prefixtext+right('00000'+ convert(varchar,lastno+1),5)) as Reqno from prefix_Requisition  where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                    string strReqNumber = Utils.Utils.getScalarValue(strReqNo, CnnString);

                    SQLString = "Exec [laserreqSlip1DashBoard]  '" + HSRP_StateID + "','" + Navembcode + "' ,  '" + ReqNum + "'";
                    DataTable dtResult = Utils.Utils.GetDataTable(SQLString, CnnString);
                    string strQuery = string.Empty;
                    string strRtoLocationName = string.Empty;
                    int Itotal = 0;

                    html.Append("<div style='width:100%;height:100%;'>" +
                                        "<table style='width:100%'>" +

                                            "<tr>" +
                                                "<td colspan='12'>" +
                                                    "<div style='text-align:center;padding:8px;'>" +
                                                        "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + strCompanyName + "</b>" +

                                                        "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                         "<b style='font-size:20px;'>" + "MATERIAL REQUSITION NOTE" + "</b>" +

                                                    "</div>" +
                                                "</td>" +
                                            "</tr>" +

                                            "<tr>" +
                                                "<td colspan='12'>" +
                                                    "<div style='text-align:center;padding:8px;'>" +
                                                        "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Production Sheet Date :" + DateTime.Now.ToString("dd-MM-yyyy") + "</b>" +
                                                        "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                         "<b style='font-size:20px;'>" + "" + "</b>" +
                                                    "</div>" +
                                                "</td>" +
                                            "</tr>" +
                                            "<tr>" +
                                                "<td colspan='6'>" +
                                                    "<div style='text-align:left;'>" +
                                                        "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>REQ.NO:-</b>" +
                                                        "" + strReqNumber + "" +
                                                    "</div>" +
                                                "</td>" +
                                                "<td colspan='6'>" +
                                                    "<div style='text-align:left;'>" +
                                                        "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>Embossing Center:</b>" +
                                                        "" + strComNew + "" +
                                                    "</div>" +
                                                "</td>" +
                                            "</tr>" +
                                            "<tr>" +
                                                "<td colspan='12'>" +
                                                    "<div style='text-align:left;padding:2px;'>" +
                                                        "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "State:" + HSRPStateName + "</b>" +
                                                        "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                         "<b style='font-size:20px;'>" + "" + "</b>" +
                                                    "</div>" +
                                                "</td>" +
                                            "</tr>" +
                                             "<tr>" +
                                                "<td colspan='12'>" +
                                                    "<div style='text-align:left;padding:8px;'>" +
                                                        "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + " " + "</b>" +
                                                        "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                         "<b style='font-size:20px;'>" + "" + "</b>" +
                                                    "</div>" +
                                                "</td>" +
                                            "</tr>" +

                                            "<tr>" +
                                                "<td colspan='1' style='text-align:center'>SR.N.</td>" +
                                                "<td colspan='3'>Product Size</td>" +
                                                "<td colspan='1'>Laser Count</td>" +
                                                "<td colspan='1'>Start Laser No</td>" +
                                                "<td colspan='1'>End Laser No</td>" +

                                            "</tr>");


                    if (dtResult.Rows.Count > 0)
                    {
                        for (int i = 0; i < dtResult.Rows.Count; i++)
                        {
                            string ID = dtResult.Rows[i]["ID"].ToString();
                            string productcode = dtResult.Rows[i]["productcode"].ToString();
                            string LaserCount = dtResult.Rows[i]["LaserCount"].ToString();
                            Itotal = Convert.ToInt32(dtResult.Rows[i]["Total"].ToString());
                            string BeginLaser = dtResult.Rows[i]["BeginLaser"].ToString();
                            string EndLaser = dtResult.Rows[i]["EndLaser"].ToString();




                            html.Append("<tr>" +
                               "<td colspan='1' style='text-align:left'>" + ID + "</td>" +
                               "<td colspan='3'>" + productcode + "</td>" +

                               "<td colspan='1'>" + LaserCount + "</td>" +
                               "<td colspan='1'>" + BeginLaser + "</td>" +
                               "<td colspan='1'>" + EndLaser + "</td>" +

                           "</tr>");
                        }
                    }
                    html.Append("<tr>" +
                             "<td colspan='1' style='text-align:center' > " + "<b>Grand Total:</b>" + "</td>" +
                             "<td colspan='3'>" + "" + "</td>" +

                             "<td colspan='1'>" + Itotal + "</td>" +
                             "<td colspan='1'>" + " " + "</td>" +
                             "<td colspan='1'>" + " " + "</td>" +

                         "</tr>");




                    html.Append("<tr>" +
                     "<td colspan='2' > " + "<b>REQUESTED BY </b>" + "</td>" +
                     "<td colspan='2'>" + "<b>AUTHORIZED BY </b>" + "</td>" +

                     "<td colspan='2'>" + "<b>ISSUED BY </b>" + "</td>" +
                     "<td colspan='2'>" + "<b>RECEIVED BY</b>" + "</td>" +


                 "</tr>");

                    html.Append("<tr>" +
                                              "<td colspan='12'>" +
                                                  "<div style='text-align:left;padding:2px;'>" +
                                                      "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Name" + "</b>" +

                                                  "</div>" +
                                              "</td>" +
                                          "</tr>");

                    html.Append("<tr>" +
                                             "<td colspan='12'>" +
                                                 "<div style='text-align:left;padding:2px;'>" +
                                                     "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Designation" + "</b>" +

                                                 "</div>" +
                                             "</td>" +
                                         "</tr>");

                    html.Append("<tr>" +
                                           "<td colspan='12'>" +
                                               "<div style='text-align:right;padding:8px;'>" +
                                                   "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Download By: Admin " + "</b>" +

                                               "</div>" +
                                           "</td>" +
                                       "</tr>");

                    html.Append("<tr>" +
                                          "<td colspan='12'>" +
                                              "<div style='text-align:right;padding:8px;'>" +
                                                  "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Sheet Generated By :" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</b>" +

                                              "</div>" +
                                          "</td>" +
                                      "</tr>");






                    html.Append("</table>");

                    html.Append("</div>");

                    try
                    {
                        //start updating hsrprecords 
                        string Query = "update prefix_Requisition set lastno=lastno+1,Lastreqno=Lastreqno+1 where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                        Utils.Utils.ExecNonQuery(Query, CnnString);
                    }
                    catch (Exception ev)
                    {
                        Label1.Text = "prefix Requisition update error: " + ev.Message;
                    }


                    #endregion

                    /*
                     * Close body & HTMl Tag
                     */
                    html.Append("</body>" +
                        "</html>");


                    if (findRecord)
                    {

                        string strSaveSQlQuery = "insert into ProductionSheetAutoGenerated_List (HSRP_StateID, State_Code, RTO_ID, Emb_Center_Id, FileName, Productiondate, ProductionTime, OrderNo) values " +
                                    "('" + HSRP_StateID + "', '" + HSRPStateShortName + "', '', '" + Navembcode + "', '" + fileName + "', '" + DateTime.Now.ToString("yyyy-MM-dd") + "', '" + DateTime.Now.ToString("HH:mm:ss") + "', '" + filePrefix + "') ";
                        Utils.Utils.ExecNonQuery(strSaveSQlQuery, CnnString); // uncomment after testing
                        Label1.Text = Label1.Text + "successfully created production sheet of State: " + HSRPStateShortName + ", , Emb_Center_Id: " + Navembcode + Environment.NewLine;

                        //lblLog.Text = lblLog.Text + Environment.NewLine + "Start Production Query at:" + DateTime.Now;
                        #region
                        try
                        {
                            if (!Directory.Exists(dir))
                            {


                                Directory.CreateDirectory(dir);

                            }

                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        try
                        {
                            var pdf = Pdf
                                    .From(html.ToString())
                                    .OfSize(PaperSize.A4)
                                    .WithTitle("Title")
                                    .WithoutOutline()
                                    .WithMargins(1.25.Centimeters())
                                     .Landscape()
                                    .Comressed()
                                    .Content();

                            FileStream readStream = File.Create(filePath);
                            BinaryWriter binaryWriter = new BinaryWriter(readStream);

                            // Write the binary data to the file
                            binaryWriter.Write(pdf);
                            binaryWriter.Close();
                            readStream.Close();
                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        #endregion
                    }

                    #endregion

                }//close foreach stateEcQuery
            }
        }

        protected void btnHero_Click(object sender, EventArgs e)
        {

            if (ddlStateName.SelectedItem.Text == "--Select State--")
            {
                lblErrMess.Text = "Please select  State";
                return;
            }
          

            string sqlPrefixQuery = "select top 1 CONVERT(INT, isnull(OrderNo,'0')) + 1 as orderNo from ProductionSheetAutoGenerated_List  with(nolock) order by id desc";
            DataTable dtPrefix = Utils.Utils.GetDataTable(sqlPrefixQuery, CnnString);

            if (dtPrefix.Rows.Count > 0)
            {
                filePrefix = dtPrefix.Rows[0]["orderNo"].ToString();
            }
            else
            {
                filePrefix = "1";
            }

            //txtFilePrefix.Text = filePrefix;

            if (filePrefix.Length > 0)
            {
                SheetGenerationHero();
            }

        }

        private void SheetGenerationHero()
        {
            string stateECQuery = string.Empty;
            string ReqNum = string.Empty;
            //string Navembid = Session["Navembid"].ToString();
            FillUserDetails();

            stateECQuery = "select distinct a.HSRP_StateID, (select HSRPStateName from hsrpstate c where c.hsrp_stateid=a.hsrp_stateid) as HSRPStateName,(select HSRPStateShortName from hsrpstate d where d.hsrp_stateid=a.hsrp_stateid) as HSRPStateShortName, navembcode from hsrprecords a with(nolock),rtolocation b where a.HSRP_StateID ='" + ddlStateName.SelectedValue + "'  and a.rtolocationid=b.rtolocationid and  NewPdfRunningNo is null and erpassigndate is not null  and Dealerid in(select dealerid from dealermaster where oemid='20') and OrderStatus='New Order' and    b.Navembcode not like '%CODO%'   AND b.NAVEMBID='" + Navembid + "' order by  a.HSRP_StateID";


            DataTable dtSE = Utils.Utils.GetDataTable(stateECQuery, CnnString);
            if (dtSE.Rows.Count > 0)
            {
                foreach (DataRow dr in dtSE.Rows)
                {
                    #region

                    string HSRP_StateID = dr["HSRP_StateID"].ToString().Trim();
                    string HSRPStateName = dr["HSRPStateName"].ToString().Trim();
                    string HSRPStateShortName = dr["HSRPStateShortName"].ToString().Trim();
                    string Navembcode = dr["Navembcode"].ToString().Trim();

                    string dir = dirPath + DateTime.Now.ToString("yyyy-MM-dd") + "\\" + HSRPStateShortName + "\\";
                    //string dir = dirPath + "/" + DateTime.Now.ToString("yyyy-MM-dd") + "/" + HSRPStateShortName + "/";
                    string fileName = "Hero" + "-" + filePrefix + "-" + Navembcode + ".pdf";
                    string filePath = dir + fileName;

                    //string folderpath = ConfigurationManager.AppSettings["InvoiceFolder"].ToString() + "/" + FinYear + "/" + oemid + "/" + HSRPStateID + "/";

                    StringBuilder html = new StringBuilder();

                    Boolean findRecord = false;
                    string strProductionSheetNo = string.Empty;

                    /*
                     *  Start body & HTMl Tag
                     */
                    #region
                    html.Append(
                        "<!DOCTYPE html>" +
                        "<html>" +
                        "<head>" +
                            "<meta charset='UTF-8'><title>Title</title>" +
                            "<style>" +
                                "@page {" +
                                    /* headers*/
                                    "@top-left {" +
                                        "content: 'Left header';" +
                                    "}" +
                                    "@top-right {" +
                                        "content: 'Right header';" +
                                    "}" +

                                    /* footers */
                                    "@bottom-left {" +
                                        "content: 'Lorem ipsum';" +
                                    "} " +
                                    "@bottom-right {" +
                                        "content: 'Page ' counter(page) ' of ' counter(pages);" +
                                    "}" +
                                    "@bottom-center  {" +
                                        "content:element(footer);" +
                                    "}" +
                                "}" +
                                //"#main-table td:nth-child(1){ width:5%; } #main-table td:nth-child(8),#main-table td:nth-child(9),#main-table td:nth-child(7){ width:6%; } #main-table td:nth-child(10){ width:8%; }" +
                                //  "#main-table td:nth-child(3),#main-table td:nth-child(5){ width:12%;white-space: nowrap; }" +
                                "#footer {" +
                                    "position: running(footer);" +
                                "}" +
                                "table {" +
                                  "border-collapse: collapse;" +
                                "}" +
                                "table, th, td {" +
                                    "border: 1px solid black;" +
                                    "text-align: left;" +
                                    "vertical-align: top;" +
                                    "padding-left:10px;" +
                                    "padding-bottom:5px;" +
                                    "padding-right:10px;" +
                                    "padding-top:5px;" +

                                "}" +
                                "#main-table table,#main-table th,#main-table td{" +
                                 "white-space: nowrap;}" +
                            "</style>" +
                        "</head>" +
                        "<body>");
                    #endregion

                    string oemDealerQuery = string.Empty;



                    oemDealerQuery = "select distinct om.oemid as oemid, om.name as oemname, dm.dealerid, dm.dealername, dm.dealercode, dm.Address, dm.HSRP_StateID, " +
                        "dm.RTOLocationID from oemmaster   om  with(nolock)" +
                        "left join dealermaster dm   with(nolock) on dm.oemid = om.oemid where dm.HSRP_StateID =" + HSRP_StateID + " and " +
                        "dm.RTOLocationID in (select RTOLocationID from rtolocation where Navembcode='" + Navembcode + "' ) and " +
                        "dm.dealerid in (select distinct dealerid from hsrprecords with(nolock) where NewPdfRunningNo is null and VahanStatus='Y' and  erpassigndate is not null and OrderStatus='New Order' and HSRP_StateID =" + HSRP_StateID + " and  Dealerid in(select dealerid from Dealermaster where oemid='20')) and Om.OEMID   in('20')";


                    #region
                    DataTable dtOD = Utils.Utils.GetDataTable(oemDealerQuery, CnnString);
                    if (dtOD.Rows.Count > 0)
                    {
                        foreach (DataRow drOD in dtOD.Rows)
                        {

                            string oemid = drOD["oemid"].ToString().Trim();
                            string dealerid = drOD["dealerid"].ToString().Trim();
                            string oemname = drOD["oemname"].ToString().Trim();
                            string dealername = drOD["dealername"].ToString().Trim();
                            string Address = drOD["Address"].ToString().Trim();

                            //start sql query
                            #region
                            string productionQuery = string.Empty;

                            DataTable dtProduction = new DataTable(); ;

                            SqlConnection con = new SqlConnection(CnnString);
                            SqlCommand cmd = new SqlCommand("USP_HeroProductionSheet", con);
                            cmd.CommandType = CommandType.StoredProcedure;
                            con.Open();
                            cmd.Parameters.AddWithValue("@navembid", Navembcode);
                            cmd.Parameters.AddWithValue("@HSRP_StateID", HSRP_StateID);
                            cmd.Parameters.AddWithValue("@Dealerid", dealerid);

                            SqlDataAdapter da = new SqlDataAdapter(cmd);
                            // dtProduction = new DataTable();
                            da.Fill(dtProduction);
                            con.Close();


                            //productionQuery = "Select ROW_NUMBER() Over (Order by a.HSRP_Front_LaserCode, a.HSRP_Rear_LaserCode) As [SRNo],a.heroorderno, a.hsrprecordID,a.roundoff_netamount,a.OrderStatus, " +
                            //   "a.HSRPRecord_AuthorizationNo,convert(varchar, OrderClosedDate, 105) as OrderClosedDate, " +
                            //   "CONVERT(varchar(20),orderdate ,103) AS OrderBookDate, " +
                            //   "CONVERT(varchar(20),OrderEmbossingDate ,103) AS OrderEmbossingDate, " +
                            //   "a.dealerid as ID, left(a.OwnerName,19) as OwnerName, " +
                            //   "a.TypeOfApplication as FuelType, a.MobileNo, " +
                            //   "(select AffixCenterDesc from AffixationCenters where Affix_id= a.affix_id ) as AffixCenterDesc, " +
                            //   "a.HSRPRecordID, CONVERT(varchar(20), HSRPRecord_AuthorizationDate,103) AS OrderDateAuth, " +
                            //   "a.OrderDate, a.EngineNo, a.ChassisNo, " +
                            //   "(select rtolocationname from rtolocation where rtolocationid =a.rtolocationid) as RTOLocationName, " +
                            //   "a.VehicleRegNo, " +
                            //   "case a.VehicleType when 'MCV/HCV/TRAILERS' then 'Trailers' when 'THREE WHEELER' then 'T.Whe.' " +
                            //   "when 'SCOOTER' then 'SCOO' when 'TRACTOR' then 'TRAC' when 'LMV(CLASS)' then 'L.CL' " +
                            //   "when 'LMV' then 'LMV' when 'MOTOR CYCLE' then 'MO.C' end as VehicleType, " +
                            //   "case a.VehicleClass when 'Transport' then 'T' else 'N.T.' end as VehicleClass, " +
                            //   "a.HSRP_StateID, (select HSRPStateName from hsrpstate where HSRP_StateID=a.HSRP_StateID) as 'State Name', " +
                            //   "(select Distinct Oemname from Dealermaster where dealerid=a.dealerid and Oemid='20' ) as OemName, " +




                            //   "(select replace(ProductCode,'MM-','') from Product where productid= a.RearPlateSize) AS RearProductCode, " +
                            //   "(select replace(ProductCode,'MM-','') from Product where productid= a.FrontPlateSize) AS FrontProductCode, " +
                            //   "a.FrontPlateSize, a.HSRP_Front_LaserCode, a.HSRP_Rear_LaserCode, a.RearPlateSize,HotStampingFoilColour FROM HSRPRecords AS a with(nolock) " +
                            //   " left join  HotStampingFoilColourMaster H  with(nolock) on a.vehicletype=h.Vehicletype and a.vehicleclass=h.Vehicleclass where IsBookMyHsrpRecord='N' and  sendtoProductionStatus ='N'  and heroorderno is not null and a.hsrp_StateID='" + HSRP_StateID + "' " +
                            //   "and a.RTOLocationID in (select rtolocationid from rtolocation where navembid='" + Navembcode + "') " +
                            //   "and  VahanStatus='Y' and newpdfrunningno is  null and ([HSRP_Front_LaserCode] is not null or [HSRP_Rear_LaserCode] is not null) " +
                            //   "and ([HSRP_Front_LaserCode] !='' or [HSRP_Rear_LaserCode] !='') " +
                            //   "and a.orderstatus='New Order' AND a.dealerid = '" + dealerid + "'  " +
                            //   "order by [HSRP_Front_LaserCode] , [HSRP_Rear_LaserCode] asc";





                            #endregion
                            //end sql query

                            #region
                            //DataTable dtProduction = Utils.Utils.GetDataTable(productionQuery, CnnString);
                            if (dtProduction.Rows.Count > 0)
                            {

                                findRecord = true;
                                string strRunningNo = string.Empty;
                                //string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) as maxSheetNo from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + NAVEMBID + "'";
                                string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) from EmbossingCentersNew with(nolock) where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "' ";
                                string strCom = Utils.Utils.getScalarValue(strSel, CnnString);
                                //DataTable dtSheetNo = fillDataTable(strSel, CnnString);
                                if (strCom.Equals(0) || strCom.Length == 0)
                                {
                                    strRunningNo = "0000001";
                                }
                                else
                                {
                                    strRunningNo = string.Format("{0:0000000}", Convert.ToInt32(strCom) + 1);
                                }

                                string strRequeNo = "select (prefixtext+right('00000'+ convert(varchar,Lastreqno+1),5)) as Reqno from prefix_Requisition with(nolock)  where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                                ReqNum = Utils.Utils.getScalarValue(strRequeNo, CnnString);

                                string strPRFIX = "select PrefixText from EmbossingCentersNew with(nolock) where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                                string strPRFIXCom = Utils.Utils.getScalarValue(strPRFIX, CnnString);
                                strProductionSheetNo = strPRFIXCom + strRunningNo;

                                string RTOLocationName = dtProduction.Rows[0]["RTOLocationName"].ToString();



                                #region


                                html.Append("<div style='page-break-before: avoid;page-break-inside: avoid;page-break-after: always;'>" +
                                      "<table style='width:100%;border: 0px;'>" +
                                          "<tr style='border: 0px;'>" +
                                              "<td  style='border: 0px; width:36%;'>" +
                                                  "<div style='text-align:left'><b>Report Generation Date:</b> " + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</div>" +
                                              "</td>" +

                                               "<td  style='border: 0px; width:30%;'>" +
                                                  "<div style='text-align:left'><b>ROSMERTA SAFETY SYSTEMS LIMITED</b> " + "</div>" +
                                              "</td>" +


                                               "<td  style='border: 0px; width:33%;'>" +
                                                  "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Production Sheet No:</b> " + strProductionSheetNo + " </div>" +
                                              "</td>" +

                                          "</tr>" +

                                            "<tr style='border: 0px;'>" +
                                              "<td  style='border: 0px;width:36%;'>" +
                                                  "<div style='text-align:left'><b>State:</b> " + HSRPStateName + "</div>" +
                                              "</td>" +

                                               "<td style='border: 0px; width:30%;'>" +
                                                  "<div style='text-align:left'><b>Production Sheet</b> " + "</div>" +
                                              "</td>" +


                                               "<td style='border: 0px;width:33%;'>" +
                                                  "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Dealer ID/Name:</b> " + dealerid + "/" + dealername + " </div>" +
                                              "</td>" +

                                          "</tr>" +

                                              "<tr style='border: 0px;'>" +
                                              "<td  style='border: 0px;width:36%;'>" +
                                                  "<div style='text-align:left'><b>EC Location: </b> " + RTOLocationName + "</div>" +
                                              "</td>" +

                                               "<td  style='border: 0px;width:30%;'>" +
                                                  "<div style='text-align:left'><b>Oem :" + oemname + " </ b > </div>" +
                                              "</td>" +


                                               "<td  style='border: 0px;width:33%;'>" +
                                                   "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Dealer Address:</b> " + Address + " </div>" +
                                              //"<div style='text-align:left'><b>Dealer Address:</b> " + Address + " </div>" +
                                              "</td>" +

                                          "</tr>" +




                                                                          "</table>" +



                                // Sr.No.Vehicle No.Front Plate Size    Front Laser No.Rear Plate Size Rear Laser No.H.S.Foil  Caution Sticker Fuel Type   VT  VC
                                "<table id='main-table' style='width:100%;border: 0px;'>" +

"<tr>" +
                                                "<td >Sr. No.</td>" +
                                                  "<td style='width:15%;'>Vehicle No.</td>" +
                                                   "<td  >Front Plate Size</td>" +
                                                "<td style='width:15%;'>Front Laser No.</td>" +

                                                "<td>Rear Plate Size</td>" +
                                                "<td style='width:15%;'>Rear Laser No.</td>" +

                                                "<td>H. S. Foil </td>" +

                                                "<td>VT</td>" +
                                                "<td>VC</ td>" +
                                                 "<td>Order No</td>" +

                                            "</tr>");
                                #endregion



                                #region
                                string strheroorderno = "";
                                int j = 0;
                                for (int i = 0; i <= dtProduction.Rows.Count - 1; i++)
                                {

                                    string HsrprecordID = dtProduction.Rows[i]["hsrprecordID"].ToString().Trim();
                                    string SRNo = dtProduction.Rows[i]["SRNo"].ToString().Trim();

                                    string VehicleNo = dtProduction.Rows[i]["VehicleRegNo"].ToString().Trim();
                                    string VC = dtProduction.Rows[i]["VehicleClass"].ToString().Trim();
                                    string VT = dtProduction.Rows[i]["VehicleType"].ToString().Trim();
                                    string FrontPSize = dtProduction.Rows[i]["FrontProductCode"].ToString().Trim();
                                    string FrontLaserNo = dtProduction.Rows[i]["HSRP_Front_LaserCode"].ToString().Trim();
                                    string RearPSize = dtProduction.Rows[i]["RearProductCode"].ToString().Trim();
                                    string RearLaserNo = dtProduction.Rows[i]["HSRP_Rear_LaserCode"].ToString().Trim();
                                    string HotStampingFoilColour = dtProduction.Rows[i]["HotStampingFoilColour"].ToString().Trim();
                                    strheroorderno = dtProduction.Rows[i]["heroorderno"].ToString();



                                    html.Append("<tr>" +


                                             "<td style='text-align:center'>" + SRNo + "</td>" +
                                    "<td style='text-align:left;font-size:20px;'>" + "<b>" + VehicleNo + "</b> </td>" +

                                     "<td >" + FrontPSize + "</td>" +
                                      "<td style='text-align:left;font-size:20px;'>" + "<b>" + FrontLaserNo + "</b> </td>" +
                                        "<td >" + RearPSize + "</td>" +
                                      "<td style='text-align:left;font-size:20px;'>" + "<b>" + RearLaserNo + "</b> </td>" +

                                       "<td style='text-align:center;'>" + HotStampingFoilColour + "</td>" +


                                         "<td style='text-align:center;'>" + VT + "</td>" +
                                         "<td style='text-align:center;'>" + VC + "</td>" +
                                         "<td style='text-align:center;' >" + strheroorderno + "</td>" +








                                    "</tr>");





                                    //start updating hsrprecords 
                                    string sqlUpdateHSRPRecords = "update hsrprecords set sendtoProductionStatus='Y', NAVPDFFlag='1', NewPdfRunningNo='" + strProductionSheetNo + "',  Requisitionsheetno='" + ReqNum + "',  " +
                                           "PdfDownloadDate=GetDate(), pdfFileName='" + fileName + "', PDFDownloadUserID='1' where hsrprecordID='" + HsrprecordID + "' ";
                                    Utils.Utils.ExecNonQuery(sqlUpdateHSRPRecords, CnnString);   // uncomment after testing
                                    //end 
                                }


                                #endregion
                                html.Append("</table>");

                                html.Append("</div>");

                                string StrSqlUpdateECQuery = "update EmbossingCentersNew set NewProductionSheetRunningNo='" + strProductionSheetNo + "' " +
                                 "where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                                Utils.Utils.ExecNonQuery(StrSqlUpdateECQuery, CnnString); // uncomment after testing
                            }

                            #endregion

                        }
                    }// close oemDealerQuery
                    #endregion

                    #region "Req Generate"
                    string strSqlQuery1 = "select CompanyName from hsrpstate with(nolock) where hsrp_stateid='" + HSRP_StateID + "'";
                    strCompanyName = Utils.Utils.getScalarValue(strSqlQuery1, CnnString);

                    string strEMBName = " select EmbCenterName from EmbossingCentersNew with(nolock) where  State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                    string strComNew = Utils.Utils.getScalarValue(strEMBName, CnnString);

                    //string strReqNo = "select (prefixtext+right('00000'+ convert(varchar,lastno+1),5)) as Reqno from prefix_Requisition with(nolock)  where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                    //string strReqNumber = Utils.Utils.getScalarValue(strReqNo, CnnString);

                    SQLString = "Exec [laserreqSlip1DashBoard]  '" + HSRP_StateID + "','" + Navembcode + "' ,  '" + ReqNum + "'";
                    DataTable dtResult = Utils.Utils.GetDataTable(SQLString, CnnString);
                    string strQuery = string.Empty;
                    string strRtoLocationName = string.Empty;
                    int Itotal = 0;

                    html.Append("<div style='width:100%;height:100%;'>" +
                                        "<table  style='width:100%'>" +

                                            "<tr>" +
                                                "<td colspan='12'>" +
                                                    "<div style='text-align:center;padding:8px;'>" +
                                                        "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + strCompanyName + "</b>" +

                                                        "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                         "<b style='font-size:20px;'>" + "MATERIAL REQUSITION NOTE" + "</b>" +

                                                    "</div>" +
                                                "</td>" +
                                            "</tr>" +

                                            "<tr>" +
                                                "<td colspan='12'>" +
                                                    "<div style='text-align:center;padding:8px;'>" +
                                                        "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Production Sheet Date :" + DateTime.Now.ToString("dd-MM-yyyy") + "</b>" +
                                                        "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                         "<b style='font-size:20px;'>" + "" + "</b>" +
                                                    "</div>" +
                                                "</td>" +
                                            "</tr>" +
                                            "<tr>" +
                                                "<td colspan='6'>" +
                                                    "<div style='text-align:left;'>" +
                                                        "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>REQ.NO:-</b>" +
                                                        "" + ReqNum + "" +
                                                    "</div>" +
                                                "</td>" +
                                                "<td colspan='6'>" +
                                                    "<div style='text-align:left;'>" +
                                                        "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>Embossing Center:</b>" +
                                                        "" + strComNew + "" +
                                                    "</div>" +
                                                "</td>" +
                                            "</tr>" +
                                            "<tr>" +
                                                "<td colspan='12'>" +
                                                    "<div style='text-align:left;padding:2px;'>" +
                                                        "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "State:" + HSRPStateName + "</b>" +
                                                        "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                         "<b style='font-size:20px;'>" + "" + "</b>" +
                                                    "</div>" +
                                                "</td>" +
                                            "</tr>" +
                                             "<tr>" +
                                                "<td colspan='12'>" +
                                                    "<div style='text-align:left;padding:8px;'>" +
                                                        "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + " " + "</b>" +
                                                        "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                         "<b style='font-size:20px;'>" + "" + "</b>" +
                                                    "</div>" +
                                                "</td>" +
                                            "</tr>" +

                                            "<tr>" +
                                                "<td colspan='1' style='text-align:center'>SR.N.</td>" +
                                                "<td colspan='3'>Product Size</td>" +
                                                "<td colspan='1'>Laser Count</td>" +
                                                "<td colspan='1'>Start Laser No</td>" +
                                                "<td colspan='1'>End Laser No</td>" +

                                            "</tr>");


                    if (dtResult.Rows.Count > 0)
                    {
                        for (int i = 0; i < dtResult.Rows.Count; i++)
                        {
                            string ID = dtResult.Rows[i]["ID"].ToString();
                            string productcode = dtResult.Rows[i]["productcode"].ToString();
                            string LaserCount = dtResult.Rows[i]["LaserCount"].ToString();
                            Itotal = Convert.ToInt32(dtResult.Rows[i]["Total"].ToString());
                            string BeginLaser = dtResult.Rows[i]["BeginLaser"].ToString();
                            string EndLaser = dtResult.Rows[i]["EndLaser"].ToString();




                            html.Append("<tr>" +
                               "<td colspan='1' style='text-align:left'>" + ID + "</td>" +
                               "<td colspan='3'>" + productcode + "</td>" +

                               "<td colspan='1'>" + LaserCount + "</td>" +
                               "<td colspan='1'>" + BeginLaser + "</td>" +
                               "<td colspan='1'>" + EndLaser + "</td>" +

                           "</tr>");
                        }
                    }
                    html.Append("<tr>" +
                             "<td colspan='1' style='text-align:center' > " + "<b>Grand Total:</b>" + "</td>" +
                             "<td colspan='3'>" + "" + "</td>" +

                             "<td colspan='1'>" + Itotal + "</td>" +
                             "<td colspan='1'>" + " " + "</td>" +
                             "<td colspan='1'>" + " " + "</td>" +

                         "</tr>");




                    html.Append("<tr>" +
                     "<td colspan='2' > " + "<b>REQUESTED BY </b>" + "</td>" +
                     "<td colspan='2'>" + "<b>AUTHORIZED BY </b>" + "</td>" +

                     "<td colspan='2'>" + "<b>ISSUED BY </b>" + "</td>" +
                     "<td colspan='2'>" + "<b>RECEIVED BY</b>" + "</td>" +


                 "</tr>");

                    html.Append("<tr>" +
                                              "<td colspan='12'>" +
                                                  "<div style='text-align:left;padding:2px;'>" +
                                                      "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Name" + "</b>" +

                                                  "</div>" +
                                              "</td>" +
                                          "</tr>");

                    html.Append("<tr>" +
                                             "<td colspan='12'>" +
                                                 "<div style='text-align:left;padding:2px;'>" +
                                                     "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Designation" + "</b>" +

                                                 "</div>" +
                                             "</td>" +
                                         "</tr>");

                    html.Append("<tr>" +
                                           "<td colspan='12'>" +
                                               "<div style='text-align:right;padding:8px;'>" +
                                                   "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Download By: Admin " + "</b>" +

                                               "</div>" +
                                           "</td>" +
                                       "</tr>");

                    html.Append("<tr>" +
                                          "<td colspan='12'>" +
                                              "<div style='text-align:right;padding:8px;'>" +
                                                  "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Sheet Generated By :" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</b>" +

                                              "</div>" +
                                          "</td>" +
                                      "</tr>");






                    html.Append("</table>");

                    html.Append("</div>");


                    try
                    {
                        //start updating hsrprecords 
                        string Query = "update prefix_Requisition set lastno=lastno+1,Lastreqno=Lastreqno+1 where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                        Utils.Utils.ExecNonQuery(Query, CnnString);
                    }
                    catch (Exception ev)
                    {
                        Label1.Text = "prefix Requisition update error: " + ev.Message;
                    }


                    #endregion

                    /*
                     * Close body & HTMl Tag
                     */
                    html.Append("</body>" +
                        "</html>");


                    if (findRecord)
                    {


                        string strSaveSQlQuery = "insert into ProductionSheetAutoGenerated_List (HSRP_StateID, State_Code, RTO_ID, Emb_Center_Id, FileName, Productiondate, ProductionTime, OrderNo) values " +
                                    "('" + HSRP_StateID + "', '" + HSRPStateShortName + "', '', '" + Navembcode + "', '" + fileName + "', '" + DateTime.Now.ToString("yyyy-MM-dd") + "', '" + DateTime.Now.ToString("HH:mm:ss") + "', '" + filePrefix + "') ";
                        Utils.Utils.ExecNonQuery(strSaveSQlQuery, CnnString); // uncomment after testing
                        Label1.Text = Label1.Text + "successfully created production sheet of State: " + HSRPStateShortName + ", , Emb_Center_Id: " + Navembcode + Environment.NewLine;

                        //lblLog.Text = lblLog.Text + Environment.NewLine + "Start Production Query at:" + DateTime.Now;



                        #region
                        try
                        {
                            if (!Directory.Exists(dir))
                            {


                                Directory.CreateDirectory(dir);

                            }

                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        try
                        {
                            var pdf = Pdf
                                    .From(html.ToString())
                                    .OfSize(PaperSize.A4)
                                    .WithTitle("Title")
                                    .WithoutOutline()
                                    .WithMargins(1.25.Centimeters())
                                     .Landscape()
                                    .Comressed()
                                    .Content();

                            FileStream readStream = File.Create(filePath);
                            BinaryWriter binaryWriter = new BinaryWriter(readStream);

                            // Write the binary data to the file
                            binaryWriter.Write(pdf);
                            binaryWriter.Close();
                            readStream.Close();
                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        #endregion
                    }

                    #endregion

                }//close foreach stateEcQuery
            }
        }

        protected void btnDelivery_Click(object sender, EventArgs e)
        {

            if (ddlStateName.SelectedItem.Text == "--Select State--")
            {
                lblErrMess.Text = "Please select  State";
                return;
            }
            string sqlPrefixQuery = "select top 1 CONVERT(INT, isnull(OrderNo,'0')) + 1 as orderNo from ProductionSheetAutoGenerated_List order by id desc";
            DataTable dtPrefix = Utils.Utils.GetDataTable(sqlPrefixQuery, CnnString);

            if (dtPrefix.Rows.Count > 0)
            {
                filePrefix = dtPrefix.Rows[0]["orderNo"].ToString();
            }
            else
            {
                filePrefix = "1";
            }

            //txtFilePrefix.Text = filePrefix;

            if (filePrefix.Length > 0)
            {
                SheetGenerationBookMyHSRPHD();
            }
        }

        private void SheetGenerationBookMyHSRPHD()
        {
            string stateECQuery = string.Empty;
            string ReqNum = string.Empty;

            // string Navembid = Session["Navembid"].ToString();
            FillUserDetails();

            string stateECShortName = "select distinct HSRP_StateID, 'EC'+HSRPStateShortName as NewHSRPStateShortName, HSRPStateShortName   from  HSRPState  where  HSRP_STateId='" + ddlStateName.SelectedValue + "'  ";

            DataTable dtECName = Utils.Utils.GetDataTable(stateECShortName, CnnString);

            string ShortECname = dtECName.Rows[0]["NewHSRPStateShortName"].ToString();
         

            //else
            //{
                //stateECQuery = "select distinct a.HSRP_StateID, (select HSRPStateName from hsrpstate c where c.hsrp_stateid=a.hsrp_stateid) as HSRPStateName,(select HSRPStateShortName from hsrpstate d with(nolock) where d.hsrp_stateid='" + ddlStateName.SelectedValue + "') as HSRPStateShortName,   a.Navembid AS  navembcode from hsrprecords a with(nolock) where  left(a.Navembid,4) = '" + ShortECname + "'  and  NewPdfRunningNo is null and erpassigndate is not null  and (IsBookMyHsrpRecord='Y')  and OrderStatus='New Order'   and a.NAVEMBID='" + Navembid + "' and   Navembid not like '%CODO%'   order by  a.HSRP_StateID";
            stateECQuery = "select distinct a.HSRP_StateID, (select HSRPStateName from hsrpstate c where c.hsrp_stateid=a.hsrp_stateid) as HSRPStateName,(select HSRPStateShortName from hsrpstate d with(nolock) where d.hsrp_stateid='" + ddlStateName.SelectedValue + "') as HSRPStateShortName,   a.Navembid AS  navembcode from hsrprecords a with(nolock) where   NewPdfRunningNo is null and erpassigndate is not null  and (IsBookMyHsrpRecord='Y')  and OrderStatus='New Order'   and a.NAVEMBID='" + Navembid + "'    order by  a.HSRP_StateID";

            DataTable dtSE = Utils.Utils.GetDataTable(stateECQuery, CnnString);
            //}
            //string allStrProductionSheetNo = string.Empty;

            if (dtSE.Rows.Count > 0)
            {
                foreach (DataRow dr in dtSE.Rows)
                {
                    #region

                    string HSRP_StateID = dr["HSRP_StateID"].ToString().Trim();
                    string HSRPStateName = dr["HSRPStateName"].ToString().Trim();
                    string HSRPStateShortName = dr["HSRPStateShortName"].ToString().Trim();
                    
                    string Navembcode = dr["Navembcode"].ToString().Trim();
                    string dir = dirPath + DateTime.Now.ToString("yyyy-MM-dd") + "\\" + HSRPStateShortName + "\\";

                    string fileName = string.Empty;
                    string filePath = string.Empty;
                    // string fileName = filePrefix + "-" + Navembcode + ".pdf";



                    StringBuilder html = new StringBuilder();

                    Boolean findRecord = false;
                    string strProductionSheetNo = string.Empty;

                    /*
                    *  Start body & HTMl Tag
                    */
                    #region
                    html.Append(
                       "<!DOCTYPE html>" +
                       "<html>" +
                       "<head>" +
                           "<meta charset='UTF-8'><title>Title</title>" +
                           "<style>" +
                               "@page {" +
                                   /* headers*/
                                   "@top-left {" +
                                       "content: 'Left header';" +
                                   "}" +
                                   "@top-right {" +
                                       "content: 'Right header';" +
                                   "}" +

                                   /* footers */
                                   "@bottom-left {" +
                                       "content: 'Lorem ipsum';" +
                                   "} " +
                                   "@bottom-right {" +
                                       "content: 'Page ' counter(page) ' of ' counter(pages);" +
                                   "}" +
                                   "@bottom-center  {" +
                                       "content:element(footer);" +
                                   "}" +
                               "}" +
                           
                               "#footer {" +
                                   "position: running(footer);" +
                               "}" +
                               "table {" +
                                 "border-collapse: collapse;" +
                               "}" +
                               "table, th, td {" +
                                   "border: 1px solid black;" +
                                   "text-align: left;" +
                                   "vertical-align: top;" +
                                   "padding-left:10px;" +
                                   "padding-bottom:6px;" +
                                   "padding-right:10px;" +
                                   "padding-top:5px;" +

                               "}" +
                               "#main-table table,#main-table th,#main-table td{" +
                                "white-space: nowrap;}" +
                           "</style>" +
                       "</head>" +
                       "<body>");

                    #endregion

                    string oemDealerQuery = string.Empty;



                    string fileAppoinmentDate = string.Empty;
                    string maxAppointmentdate = "select distinct  Convert(varchar(10),max(SlotBookingDate),105) as MaxAppointmentdate  " +
                    " from	HSRPrecords H with(nolock) join DealerAffixationCenter d  with(nolock) on H.HSRP_StateID=d.StateID and H.Affix_Id=d.DealerAffixationID  join BookMyHSRPAppointment B on H.orderno=B.orderno  " +
                   " where   Navembid='" + Navembcode + "'     and  NewPdfRunningNo is null and erpassigndate is not null and  h.OrderStatus='New Order' and  IsBookMyHsrpRecord='Y' and  h.affix_id is not NULL  and   d.TypeofDelivery in('Home','RWA')  and B.IsFrame='Y'";
                    DataTable dtmax = Utils.Utils.GetDataTable(maxAppointmentdate, CnnString);

                    if (dtmax.Rows.Count > 0)
                    {
                        fileAppoinmentDate = dtmax.Rows[0]["MaxAppointmentdate"].ToString();
                    }
                    fileName = "HomeDeliveryPlasticPacking" + "-" + fileAppoinmentDate + "-" + filePrefix + "-" + Navembcode + ".pdf";
                    filePath = dir + fileName;



                    oemDealerQuery = "select distinct d.oemid as oemid, (select name  from oemmaster where oemid=d.oemid) as oemname,d.DispatchHub,d.dealerid as dealerid,d.DealerAffixationcenterName as Dealername,  d.DealerAffixationCenterAddress as Address, " +
                     "d.DealerAffixationID as SubDealerId, H.dealerid AS ParentDealerId,SlotBookingDate from	HSRPrecords H with(nolock) join DealerAffixationCenter d  with(nolock) on H.HSRP_StateID=d.StateID and H.Affix_Id=d.DealerAffixationID  join BookMyHSRPAppointment B on H.orderno=B.orderno  " +
                    " where   Navembid='" + Navembcode + "'  and   NewPdfRunningNo is null and erpassigndate is not null and  h.OrderStatus='New Order' and  IsBookMyHsrpRecord='Y' and  B.IsFrame='Y' and  h.affix_id is not NULL   and  d.TypeofDelivery in('Home','RWA')";

                    #region
                    DataTable dtOD = Utils.Utils.GetDataTable(oemDealerQuery, CnnString);

                    if (dtOD.Rows.Count > 0)
                    {
                        string allStrProductionSheetNo = string.Empty;
                        Session["ECWiseProductionsheet="] = null;
                        foreach (DataRow drOD in dtOD.Rows)
                        {

                            string oemid = drOD["oemid"].ToString().Trim();
                            string dealerid = drOD["dealerid"].ToString().Trim();
                            string oemname = drOD["oemname"].ToString().Trim();
                            string dealername = drOD["Dealername"].ToString().Trim();
                            string Address = drOD["Address"].ToString().Trim();
                            string dispatchHub = drOD["DispatchHub"].ToString().Trim();

                            DateTime AppointmentDate = Convert.ToDateTime(drOD["SlotBookingDate"].ToString().Trim());


                            string strsubaffixid = drOD["SubDealerId"].ToString();
                            string strParentDealerId = drOD["ParentDealerId"].ToString();

                            //start sql query
                            #region
                            string productionQuery = string.Empty;

                            DataTable dtProduction = new DataTable();


                            SqlConnection con = new SqlConnection(CnnString);
                            SqlCommand cmd = new SqlCommand("USP_BookMYHSRPProductionSheetWIthFramesNew", con);
                            cmd.CommandType = CommandType.StoredProcedure;
                            con.Open();

                            cmd.Parameters.AddWithValue("@HSRP_StateID", HSRP_StateID);
                            cmd.Parameters.AddWithValue("@navembid", Navembcode);
                            cmd.Parameters.AddWithValue("@Dealerid", strParentDealerId);
                            cmd.Parameters.AddWithValue("@Affix_Id", strsubaffixid);
                            cmd.Parameters.AddWithValue("@AppointmentDate", AppointmentDate);
                            SqlDataAdapter da = new SqlDataAdapter(cmd);
                            // dtProduction = new DataTable();
                            da.Fill(dtProduction);
                            con.Close();

                            // string fileAppoinmentDate = AppointmentDate.ToString("dd/MM/yyyy");

                            #endregion
                            //end sql query

                            #region
                            //DataTable dtProduction = Utils.Utils.GetDataTable(productionQuery, CnnString);
                            if (dtProduction.Rows.Count > 0)
                            {

                                findRecord = true;
                                string strRunningNo = string.Empty;
                                //string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) as maxSheetNo from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + NAVEMBID + "'";
                                string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) from EmbossingCentersNew where    Emb_Center_Id='" + Navembcode + "' ";
                                string strCom = Utils.Utils.getScalarValue(strSel, CnnString);
                                //DataTable dtSheetNo = fillDataTable(strSel, CnnString);
                                if (strCom.Equals(0) || strCom.Length == 0)
                                {
                                    strRunningNo = "0000001";
                                }
                                else
                                {
                                    strRunningNo = string.Format("{0:0000000}", Convert.ToInt32(strCom) + 1);
                                }

                                string strRequeNo = "select (prefixtext+right('00000'+ convert(varchar,Lastreqno+1),5)) as Reqno from prefix_Requisition  where hsrp_stateid='" + ddlStateName.SelectedValue + "' and prefixfor='Req No'";
                                ReqNum = Utils.Utils.getScalarValue(strRequeNo, CnnString);

                                string strPRFIX = "select PrefixText from EmbossingCentersNew where  Emb_Center_Id='" + Navembcode + "'";
                                string strPRFIXCom = Utils.Utils.getScalarValue(strPRFIX, CnnString);
                                strProductionSheetNo = strPRFIXCom + strRunningNo;

                                string RTOLocationName = dtProduction.Rows[0]["RTOLocationName"].ToString();

                                string FS1 = string.Empty;
                                string FS2 = string.Empty;



                                #region


                                html.Append("<div style='page-break-before: avoid;page-break-inside: avoid;page-break-after: always;'>" +

                                   "<table style='width:100%;border: 0px;'>" +
                                       "<tr  style='border: 0px;'>" +
                                           "<td colspan='4'   style='border: 0px; '>" +
                                               "<div style='text-align:left'><b>Sheet Generation Date:</b> " + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</div>" +
                                           "</td>" +

                                            "<td colspan='3'  style='border: 0px; '>" +
                                               "<div style='text-align:left'><b>ROSMERTA SAFETY SYSTEMS LIMITED</b> " + "</div>" +
                                           "</td>" +


                                            "<td  colspan='6' style='border: 0px;'>" +
                                               "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Production Sheet No:</b> " + strProductionSheetNo + " </div>" +
                                           "</td>" +

                                       "</tr>" +

                                         "<tr style='border: 0px;'>" +
                                           "<td colspan='4' style='border: 0px;'>" +
                                               "<div style='text-align:left'><b>State:</b> " + HSRPStateName + "</div>" +
                                           "</td>" +

                                            "<td colspan='3' style='border: 0px; '>" +
                                               //"<div style='text-align:left'><b>Book My HSRP (Home Delivery) </b> " + "</div>" +
                                               "<div style='text-align:left;font-size:22px;''><b> MHHSRP (Home Delivery) </b> " + "</div>" +
                                           "</td>" +


                                            "<td colspan='6' style='border: 0px;'>" +
                                               "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Hub/Dealer (ID - Name):</b> " + dealerid + "/" + dispatchHub  + " </div>" +
                                           "</td>" +

                                       "</tr>" +

                                           "<tr style='border: 0px;'>" +
                                           "<td colspan='4' style='border: 0px;'>" +
                                               "<div style='text-align:left'><b>EC Location: </b> " + RTOLocationName + "</div>" +
                                           "</td>" +

                                            "<td colspan='3' style='border: 0px;'>" +
                                              "<div style='text-align:left;font-size:22px;'><b>Appointment Date: " + dtProduction.Rows[0]["AppointmentDate"].ToString() + "</ b >" + "</div>" +
                                           "</td>" +


                                            "<td  colspan='6' style='border: 0px;'>" +
                                                              "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Hub/Dealer Address:</b> " + Address + " </div>" +
                                           "</td>" +

                                       "</tr>" +




                                       "<tr>" +
                                            "<td style='text-align:center;white-space: nowrap'>Sr. No.</td>" +
                                               "<td style='width:15%;white-space: nowrap'>Vehicle No.</td>" +
                                               "<td>Front Plate Size</td>" +

                                                "<td style='width:15%;white-space: nowrap'>Front Laser No.</td>" +

                                            "<td>Rear Plate Size</td>" +


                                                "<td style='width:15%;white-space: nowrap'>Rear Laser No.</td>" +


                                            "<td style='white-space: nowrap'>H. S. Foil </td>" +
                                            "<td style='white-space: nowrap'>Caution Sticker</td>" +
                                            "<td style='white-space: nowrap'>Fuel Type</td>" +
                                            "<td style='white-space: nowrap'>VT</td>" +
                                              "<td style='white-space: nowrap'>VC</td>" +
                                            "<td style='white-space: nowrap'>Frame</ td>" +
                                              "<td style='white-space: nowrap'>Pin Code</ td>" +

                                        "</tr>");





                                #endregion

                                #region
                                string strtvspono = "";

                                int j = 0;
                                int total = 0;
                                StringBuilder UpdateSQL = new StringBuilder();
                                for (int i = 0; i <= dtProduction.Rows.Count - 1; i++)
                                {

                                    j = j + 1;

                                    if (total == 22)
                                    {
                                        total = 0;

                                        html.Append("<div style='page-break-before: avoid;page-break-inside: avoid;page-break-after: always;'>" +
                                 "<table style='width:100%;border: 0px;'>" +
                                     "<tr style='border: 0px;'>" +
                                         "<td colspan='4'  style='border: 0px;'>" +
                                             "<div style='text-align:left'><b>Sheet Generation Date:</b> " + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</div>" +
                                         "</td>" +

                                          "<td colspan='3'  style='border: 0px;'>" +
                                             "<div style='text-align:left'><b>ROSMERTA SAFETY SYSTEMS LIMITED</b> " + "</div>" +
                                         "</td>" +


                                          "<td colspan='6' style='border: 0px; '>" +
                                             "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Production Sheet No:</b> " + strProductionSheetNo + " </div>" +
                                         "</td>" +

                                     "</tr>" +

                                       "<tr style='border: 0px;'>" +
                                         "<td colspan='4' style='border: 0px;'>" +
                                             "<div style='text-align:left'><b>State:</b> " + HSRPStateName + "</div>" +
                                         "</td>" +

                                          "<td colspan='3' style='border: 0px; '>" +
                                               "<div style='text-align:left;font-size:22px;''><b> MHHSRP (Home Delivery) </b> " + "</div>" +
                                         "</td>" +


                                          "<td colspan='6' style='border: 0px;'>" +
                                          "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Hub/Dealer (ID - Name):</b> " + dealerid + "/" + dispatchHub + " </div>" +
                                      //"<div style='float:left;width: 500px;word-wrap: break-word;'><b>Hub/Dealer (ID - Name):</b> " + dealerid + "/" + dealername + "/" + strsubaffixid + " </div>" +
                                         "</td>" +

                                     "</tr>" +

                                         "<tr style='border: 0px;'>" +
                                         "<td colspan='4'  style='border: 0px;'>" +
                                             "<div style='text-align:left'><b>EC Location: </b> " + RTOLocationName + "</div>" +
                                         "</td>" +

                                          "<td colspan='3' style='border: 0px;'>" +
                                            "<div style='text-align:left;font-size:22px;'><b>Appointment Date: " + dtProduction.Rows[0]["AppointmentDate"].ToString() + "</ b >" + "</div>" +
                                         "</td>" +


                                          "<td colspan='6'  style='border: 0px;'>" +
                                                            "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Hub/Dealer Address:</b> " + Address + " </div>" +
                                         "</td>" +

                                     "</tr>" +




                                        "<tr>" +
                                            "<td style='text-align:center;white-space: nowrap'>Sr. No.</td>" +
                                               "<td style='width:15%;white-space: nowrap'>Vehicle No.</td>" +
                                               "<td>Front Plate Size</td>" +

                                                "<td style='width:15%;white-space: nowrap'>Front Laser No.</td>" +

                                            "<td>Rear Plate Size</td>" +


                                                "<td style='width:15%;white-space: nowrap'>Rear Laser No.</td>" +


                                            "<td style='white-space: nowrap'>H. S. Foil </td>" +
                                            "<td style='white-space: nowrap'>Caution Sticker</td>" +
                                            "<td style='white-space: nowrap'>Fuel Type</td>" +
                                            "<td style='white-space: nowrap'>VT</td>" +
                                              "<td style='white-space: nowrap'>VC</td>" +
                                            "<td style='white-space: nowrap'>Frame</ td>" +
                                              "<td style='white-space: nowrap'>Pin Code</ td>" +

                                        "</tr>");


                                    }

                                    string RS1 = string.Empty;
                                    string RS2 = string.Empty;
                                    total = total + 1;
                                    //foreach (DataRow drProduction in dtProduction.Rows)
                                    //{
                                    string HsrprecordID = dtProduction.Rows[i]["hsrprecordID"].ToString().Trim();
                                    string SRNo = dtProduction.Rows[i]["SRNo"].ToString().Trim();
                                    //string ORD = "";// drProduction["ORD"].ToString().Trim();
                                    string VC = dtProduction.Rows[i]["VehicleClass"].ToString().Trim();
                                    string VT = dtProduction.Rows[i]["VehicleType"].ToString().Trim();
                                    string VehicleNo = dtProduction.Rows[i]["VehicleRegNo"].ToString().Trim();
                                    //string VT = drProduction["VehicleType"].ToString().Trim();
                                    //string ChassisNo = drProduction["ChassisNo"].ToString().Trim();
                                    //string EngineNo = drProduction["EngineNo"].ToString().Trim();
                                    string FuelType = dtProduction.Rows[i]["FuelType"].ToString().Trim();
                                    string FrontPSize = dtProduction.Rows[i]["FrontProductCode"].ToString().Trim();

                                    string FrontLaserNo = dtProduction.Rows[i]["HSRP_Front_LaserCode"].ToString().Trim();
                                    if (FrontLaserNo != "")
                                    {
                                        FS1 = dtProduction.Rows[i]["HSRP_Front_LaserCode"].ToString().Substring(0, 7);
                                        FS2 = dtProduction.Rows[i]["HSRP_Front_LaserCode"].ToString().Substring(7, 5);
                                    }

                                    string RearPSize = dtProduction.Rows[i]["RearProductCode"].ToString().Trim();
                                    string RearLaserNo = dtProduction.Rows[i]["HSRP_Rear_LaserCode"].ToString().Trim();
                                    if (RearLaserNo != "")
                                    {
                                        RS1 = dtProduction.Rows[i]["HSRP_Rear_LaserCode"].ToString().Substring(0, 7);
                                        RS2 = dtProduction.Rows[i]["HSRP_Rear_LaserCode"].ToString().Substring(7, 5);
                                    }



                                    string StickerColor = dtProduction.Rows[i]["stickerColor"].ToString().Trim();
                                    string HotStampingFoilColour = dtProduction.Rows[i]["HotStampingFoilColour"].ToString().Trim();
                                    string Frame = dtProduction.Rows[i]["Frame"].ToString().Trim();
                                    //string AppointmentDate = drProduction["AppointmentDate"].ToString().Trim();
                                    string Pincode = dtProduction.Rows[i]["Pincode"].ToString().Trim();
                                    // "<b>Production Sheet No: " + strProductionSheetNo + "<br /> </b>" +
                                    html.Append("<tr>" +
                                          "<td style='text-align:center;white-space: nowrap'>" + SRNo + "</td>" +
                                          "<td style='font-size:20px;white-space: nowrap' >" + "<b>" + VehicleNo + "</td>" +
                                          "<td style='white-space: nowrap'> " + FrontPSize + "  </td> " +
                                          "<td style='font-size:20px;white-space: nowrap'>" + "<b>" + FS1 + "<b>" + FS2 + "</b> </td>" +


                                          "<td style='white-space: nowrap'>" + RearPSize + "</td>" +
                                          "<td style='font-size:20px;white-space: nowrap'>" + "<b>" + RS1 + "<b>" + RS2 + "</b> </td>" +

                                            "<td style='white-space: nowrap'>" + HotStampingFoilColour + "</td>" +
                                           "<td style='white-space: nowrap'>" + StickerColor + "</td>" +
                                           "<td style='white-space: nowrap'>" + FuelType + "</td>" +
                                           "<td style='white-space: nowrap'>" + VT + "</td>" +
                                           "<td style='white-space: nowrap'>" + VC + "</td>" +


                                          "<td style='white-space: nowrap'>" + Frame + "</td>" +

                                            "<td style='white-space: nowrap'>" + Pincode + "</td>" +
                                      "</tr>");




                                    UpdateSQL.Append("update hsrprecords set sendtoProductionStatus='Y', NAVPDFFlag='1', NewPdfRunningNo='" + strProductionSheetNo + "', Requisitionsheetno='" + ReqNum + "',PdfDownloadDate=GetDate(), pdfFileName='" + fileName + "', PDFDownloadUserID='1' where hsrprecordID='" + HsrprecordID + "';");


                                }

                                allStrProductionSheetNo += "," + strProductionSheetNo;

                                allStrProductionSheetNo = allStrProductionSheetNo.TrimStart(',');
                                Session["ECWiseProductionsheet="] = allStrProductionSheetNo;
                                #endregion
                                html.Append("</table>");

                                html.Append("</div>");

                                if (UpdateSQL.ToString().Length > 0)
                                {
                                    Utils.Utils.ExecNonQuery(UpdateSQL.ToString(), CnnString);
                                    string StrSqlUpdateECQuery = "update EmbossingCentersNew set NewProductionSheetRunningNo='" + strProductionSheetNo + "' " +
                                     "where  Emb_Center_Id='" + Navembcode + "'";
                                    Utils.Utils.ExecNonQuery(StrSqlUpdateECQuery, CnnString); // uncomment after testing
                                }

                            }

                            #endregion

                        }
                    }// close oemDealerQuery
                    #endregion
                    #region "Req Generate"
                    string strSqlQuery1 = "select CompanyName from hsrpstate where hsrp_stateid='" + ddlStateName.SelectedValue + "'";
                    strCompanyName = Utils.Utils.getScalarValue(strSqlQuery1, CnnString);

                    string strEMBName = " select EmbCenterName from EmbossingCentersNew where   Emb_Center_Id='" + Navembcode + "'";
                    string strComNew = Utils.Utils.getScalarValue(strEMBName, CnnString);

                    string strReqNo = "select (prefixtext+right('00000'+ convert(varchar,Lastreqno+1),5)) as Reqno from prefix_Requisition  where hsrp_stateid='" + ddlStateName.SelectedValue + "' and prefixfor='Req No'";
                    string strReqNumber = Utils.Utils.getScalarValue(strReqNo, CnnString);

                    SQLString = "Exec [laserreqSlip1DashBoardBookMYHSRP]  '" + Navembcode + "' ,  '" + ReqNum + "'";
                    DataTable dtResult = Utils.Utils.GetDataTable(SQLString, CnnString);
                    string strQuery = string.Empty;
                    string strRtoLocationName = string.Empty;
                    int Itotal = 0;

                    html.Append("<div style='width:100%;height:100%;'>" +
                                        "<table style='width:100%'>" +

                                            "<tr>" +
                                                "<td colspan='12'>" +
                                                    "<div style='text-align:center;padding:8px;'>" +
                                                        "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + strCompanyName + "</b>" +

                                                        "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                         "<b style='font-size:20px;'>" + "MATERIAL REQUSITION NOTE" + "</b>" +

                                                    "</div>" +
                                                "</td>" +
                                            "</tr>" +

                                            "<tr>" +
                                                "<td colspan='12'>" +
                                                    "<div style='text-align:center;padding:8px;'>" +
                                                        "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Production Sheet Date :" + DateTime.Now.ToString("dd-MM-yyyy") + "</b>" +
                                                        "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                         "<b style='font-size:20px;'>" + "" + "</b>" +
                                                    "</div>" +
                                                "</td>" +
                                            "</tr>" +
                                            "<tr>" +
                                                "<td colspan='6'>" +
                                                    "<div style='text-align:left;'>" +
                                                        "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>REQ.NO:-</b>" +
                                                        "" + strReqNumber + "" +
                                                    "</div>" +
                                                "</td>" +
                                                "<td colspan='6'>" +
                                                    "<div style='text-align:left;'>" +
                                                        "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>Embossing Center:</b>" +
                                                        "" + strComNew + "" +
                                                    "</div>" +
                                                "</td>" +
                                            "</tr>" +
                                            "<tr>" +
                                                "<td colspan='12'>" +
                                                    "<div style='text-align:left;padding:2px;'>" +
                                                        "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "State:" + ddlStateName.SelectedItem.Text + "</b>" +
                                                        "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                         "<b style='font-size:20px;'>" + "" + "</b>" +
                                                    "</div>" +
                                                "</td>" +
                                            "</tr>" +
                                             "<tr>" +
                                                "<td colspan='12'>" +
                                                    "<div style='text-align:left;padding:8px;'>" +
                                                        "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + " " + "</b>" +
                                                        "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                         "<b style='font-size:20px;'>" + "" + "</b>" +
                                                    "</div>" +
                                                "</td>" +
                                            "</tr>" +

                                            "<tr>" +
                                                "<td colspan='1' style='text-align:center'>SR.N.</td>" +
                                                "<td colspan='3'>Product Size</td>" +
                                                "<td colspan='1'>Laser Count</td>" +
                                                "<td colspan='1'>Start Laser No</td>" +
                                                "<td colspan='1'>End Laser No</td>" +

                                            "</tr>");


                    if (dtResult.Rows.Count > 0)
                    {
                        for (int i = 0; i < dtResult.Rows.Count; i++)
                        {
                            string ID = dtResult.Rows[i]["ID"].ToString();
                            string productcode = dtResult.Rows[i]["productcode"].ToString();
                            string LaserCount = dtResult.Rows[i]["LaserCount"].ToString();
                            Itotal = Convert.ToInt32(dtResult.Rows[i]["Total"].ToString());
                            string BeginLaser = dtResult.Rows[i]["BeginLaser"].ToString();
                            string EndLaser = dtResult.Rows[i]["EndLaser"].ToString();




                            html.Append("<tr>" +
                               "<td colspan='1' style='text-align:left'>" + ID + "</td>" +
                               "<td colspan='3'>" + productcode + "</td>" +

                               "<td colspan='1'>" + LaserCount + "</td>" +
                               "<td colspan='1'>" + BeginLaser + "</td>" +
                               "<td colspan='1'>" + EndLaser + "</td>" +

                           "</tr>");
                        }

                        try
                        {
                            //start updating hsrprecords 
                            string Query = "update prefix_Requisition set lastno=lastno+1,Lastreqno=Lastreqno+1 where hsrp_stateid='" + ddlStateName.SelectedValue + "' and prefixfor='Req No'";
                            Utils.Utils.ExecNonQuery(Query, CnnString);
                        }
                        catch (Exception ev)
                        {
                            Label1.Text = "prefix Requisition update error: " + ev.Message;
                        }
                    }
                    html.Append("<tr>" +
                             "<td colspan='1' style='text-align:center' > " + "<b>Grand Total:</b>" + "</td>" +
                             "<td colspan='3'>" + "" + "</td>" +

                             "<td colspan='1'>" + Itotal + "</td>" +
                             "<td colspan='1'>" + " " + "</td>" +
                             "<td colspan='1'>" + " " + "</td>" +

                         "</tr>");




                    html.Append("<tr>" +
                     "<td colspan='2' > " + "<b>REQUESTED BY </b>" + "</td>" +
                     "<td colspan='2'>" + "<b>AUTHORIZED BY </b>" + "</td>" +

                     "<td colspan='2'>" + "<b>ISSUED BY </b>" + "</td>" +
                     "<td colspan='2'>" + "<b>RECEIVED BY</b>" + "</td>" +


                 "</tr>");

                    html.Append("<tr>" +
                                              "<td colspan='12'>" +
                                                  "<div style='text-align:left;padding:2px;'>" +
                                                      "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Name" + "</b>" +

                                                  "</div>" +
                                              "</td>" +
                                          "</tr>");

                    html.Append("<tr>" +
                                             "<td colspan='12'>" +
                                                 "<div style='text-align:left;padding:2px;'>" +
                                                     "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Designation" + "</b>" +

                                                 "</div>" +
                                             "</td>" +
                                         "</tr>");

                    html.Append("<tr>" +
                                           "<td colspan='12'>" +
                                               "<div style='text-align:right;padding:8px;'>" +
                                                   "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Download By: Admin " + "</b>" +

                                               "</div>" +
                                           "</td>" +
                                       "</tr>");

                    html.Append("<tr>" +
                                          "<td colspan='12'>" +
                                              "<div style='text-align:right;padding:8px;'>" +
                                                  "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Sheet Generated By :" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</b>" +

                                              "</div>" +
                                          "</td>" +
                                      "</tr>");


                    html.Append("<tr style='visibility: hidden;'>" + "<td  ><p style=\"page-break-after:always\"/></td></tr>");



                    //html.Append("</table>");

                    //html.Append("</div>");

                    #region "PS Summary Report"
                    string PS = string.Empty;
                    if (Session["ECWiseProductionsheet="] != null)
                    {
                        PS = Session["ECWiseProductionsheet="].ToString();
                    }
                    /*
                     * Close body & HTMl Tag
                     */

                    SQLString = "Exec [USP_BMPSSummaryHomedelivery]  '" + PS + "' ";
                    DataTable dtResultsummary = Utils.Utils.GetDataTable(SQLString, CnnString);

                    html.Append("<div style='width:100%;height:100%;'>" +
                                 "<table style='width:100%'>" +

                                     "<tr>" +
                                         "<td colspan='12'>" +
                                             "<div style='text-align:center;padding:8px;'>" +
                                                 "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Home Delivery Summary Report" + "</b>" +

                                             "</div>" +
                                         "</td>" +
                                     "</tr>" +

                                      "<tr>" +
                                                "<td colspan='1' style='text-align:center'>PS NO</td>" +
                                                "<td colspan='3'>Affixation Pin Code</td>" +
                                                "<td colspan='1'>2W</td>" +
                                                "<td colspan='1'>3W</td>" +
                                                "<td colspan='1'>4W</td>" +
                                                "<td colspan='1'>OTH</td>" +
                                                  "<td colspan='1'>Total</td>" +

                                            "</tr>");

                    if (dtResultsummary.Rows.Count > 0)
                    {
                        for (int i = 0; i < dtResultsummary.Rows.Count; i++)
                        {
                            string PSNo = dtResultsummary.Rows[i]["PSNo"].ToString();
                            string AffixationdealerName = dtResultsummary.Rows[i]["Affixation Pin Code"].ToString();
                            string W2 = dtResultsummary.Rows[i]["2W"].ToString();
                            string W3 = dtResultsummary.Rows[i]["3W"].ToString();
                            string W4 = dtResultsummary.Rows[i]["4W"].ToString();
                            string OTH = dtResultsummary.Rows[i]["OTH"].ToString();
                            string Total = dtResultsummary.Rows[i]["Total"].ToString();




                            html.Append("<tr>" +
                               "<td colspan='1' style='text-align:left'>" + PSNo + "</td>" +
                               "<td colspan='3'>" + AffixationdealerName + "</td>" +

                               "<td colspan='1'>" + W2 + "</td>" +
                               "<td colspan='1'>" + W3 + "</td>" +
                               "<td colspan='1'>" + W4 + "</td>" +
                                "<td colspan='1'>" + OTH + "</td>" +
                                 "<td colspan='1'>" + Total + "</td>" +

                           "</tr>");
                        }
                    }

                    html.Append("</table>");

                    html.Append("</div>");
                    html.Append("<div style='text-align:left;padding:8px;'>" +
                                            "<b style='font-size:25px;margin-top:2px;margin-bottom:2px;'>" + "<p>&#x25CF The Owner has to furnish the original FIR/SDE copy to the Fitment center at the time of fitment of HSRP.</p><p>&#x25CF The Owner has to deposit the damaged plate at the fitment center at the time of fitment of HSRP.</p><p>&#x25CF The fitment center has to retail the old TV/NTV plate in case of fitment due to conversion,re-assignment.</p>" + "</b>" +

                                        "</div>"
);

                    html.Append("</div>");

                    #endregion



                    #endregion

                    /*
                     * Close body & HTMl Tag
                     */


                    if (findRecord)
                    {






                        string strSaveSQlQuery = "insert into ProductionSheetAutoGenerated_List (HSRP_StateID, State_Code, RTO_ID, Emb_Center_Id, FileName, Productiondate, ProductionTime, OrderNo) values " +
                                    "('" + ddlStateName.SelectedValue + "', '" + HSRPStateShortName + "', '', '" + Navembcode + "', '" + fileName + "', '" + DateTime.Now.ToString("yyyy-MM-dd") + "', '" + DateTime.Now.ToString("HH:mm:ss") + "', '" + filePrefix + "') ";
                        Utils.Utils.ExecNonQuery(strSaveSQlQuery, CnnString); // uncomment after testing
                        Label1.Text = Label1.Text + "successfully created production sheet of State: " + HSRPStateShortName + ", , Emb_Center_Id: " + Navembcode + Environment.NewLine;

                        //lblLog.Text = lblLog.Text + Environment.NewLine + "Start Production Query at:" + DateTime.Now;



                        #region
                        try
                        {
                            if (!Directory.Exists(dir))
                            {


                                Directory.CreateDirectory(dir);

                            }

                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        try
                        {
                            var pdf = Pdf
                                    .From(html.ToString())
                                    .OfSize(PaperSize.A4)
                                    .WithTitle("Title")
                                    .WithoutOutline()
                                    .WithMargins(1.25.Centimeters())
                                    .Landscape()
                                    .Comressed()
                                    .Content();

                            FileStream readStream = File.Create(filePath);
                            BinaryWriter binaryWriter = new BinaryWriter(readStream);

                            // Write the binary data to the file
                            binaryWriter.Write(pdf);
                            binaryWriter.Close();
                            readStream.Close();
                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        #endregion
                    }

                    #endregion

                }//close foreach stateEcQuery
            }
        }




        protected void btnonlysticker_Click(object sender, EventArgs e)
        {

            if (ddlStateName.SelectedItem.Text == "--Select State--")
            {
                lblErrMess.Text = "Please select  State";
                return;
            }

            if (ddlStateName.SelectedValue == "0")
            {
                lblErrMess.Text = String.Empty;
                lblErrMess.Text = "Please select State";
                return;
            }


            if (txtDate.Text == string.Empty)
            {
                lblErrMess.Text = String.Empty;
                lblErrMess.Text = "Please select Appointment Date";
                return;
            }
            string sqlPrefixQuery = "select top 1 CONVERT(INT, isnull(OrderNo,'0')) + 1 as orderNo from ProductionSheetAutoGenerated_List order by id desc";
            DataTable dtPrefix = Utils.Utils.GetDataTable(sqlPrefixQuery, CnnString);

            if (dtPrefix.Rows.Count > 0)
            {
                filePrefix = dtPrefix.Rows[0]["orderNo"].ToString();
            }
            else
            {
                filePrefix = "1";
            }

            //txtFilePrefix.Text = filePrefix;

            if (filePrefix.Length > 0)
            {
                SheetGenerationBookMyHSRPOSNew();
            }
        }

        private void SheetGenerationBookMyHSRPOS()
        {
            string stateECQuery = string.Empty;
            string ReqNum = string.Empty;



            string stateECShortName = "select distinct HSRP_StateID, 'EC'+HSRPStateShortName as NewHSRPStateShortName, HSRPStateShortName   from  HSRPState  where  HSRP_STateId='" + ddlStateName.SelectedValue + "'";

            DataTable dtECName = Utils.Utils.GetDataTable(stateECShortName, CnnString);

            string ShortECname = dtECName.Rows[0]["NewHSRPStateShortName"].ToString();
            // string HSRPStateShortName = dtECName.Rows[0]["HSRPStateShortName"].ToString();

            stateECQuery = "select distinct a.HSRP_StateID, (select HSRPStateName from hsrpstate c where c.hsrp_stateid=a.hsrp_stateid) as HSRPStateName,(select HSRPStateShortName from hsrpstate d with(nolock) where d.hsrp_stateid='" + ddlStateName.SelectedValue + "') as HSRPStateShortName,   (select navembid from rtolocation where RTOLocationID in (select rtolocationid from DealerAffixationCenter where DealerAffixationcenter.dealeraffixationid = a.affix_id)) AS  navembcode from hsrprecordsonlystricker a where  (select navembid from rtolocation where RTOLocationID in (select rtolocationid from DealerAffixationCenter where DealerAffixationcenter.dealeraffixationid=a.affix_id)) like '%" + ShortECname + "%'  and  NewPdfRunningNo is null   order by  a.HSRP_StateID";
            DataTable dtSE = Utils.Utils.GetDataTable(stateECQuery, CnnString);
            //string allStrProductionSheetNo = string.Empty;

            if (dtSE.Rows.Count > 0)
            {
                foreach (DataRow dr in dtSE.Rows)
                {
                    #region

                    string HSRP_StateID = dr["HSRP_StateID"].ToString().Trim();
                    string HSRPStateName = dr["HSRPStateName"].ToString().Trim();
                    string HSRPStateShortName = dr["HSRPStateShortName"].ToString().Trim();
                    //string RTOLocationID = dr["RTOLocationID"].ToString().Trim();
                    //string RTOLocationName = dr["RTOLocationName"].ToString().Trim();
                    //string NAVEMBID = dr["NAVEMBID"].ToString().Trim();
                    string Navembcode = dr["Navembcode"].ToString().Trim();
                    string dir = dirPath + DateTime.Now.ToString("yyyy-MM-dd") + "\\" + HSRPStateShortName + "\\";

                    string fileName = string.Empty;
                    string filePath = string.Empty;
                    // string fileName = filePrefix + "-" + Navembcode + ".pdf";



                    StringBuilder html = new StringBuilder();

                    Boolean findRecord = false;
                    string strProductionSheetNo = string.Empty;

                    /*
                    *  Start body & HTMl Tag
                    */
                    #region
                    html.Append(
                        "<!DOCTYPE html>" +
                        "<html>" +
                        "<head>" +
                            "<meta charset='UTF-8'><title>Title</title>" +
                            "<style>" +
                                "@page {" +
                                    /* headers*/
                                    "@top-left {" +
                                        "content: 'Left header';" +
                                    "}" +
                                    "@top-right {" +
                                        "content: 'Right header';" +
                                    "}" +

                                    /* footers */
                                    "@bottom-left {" +
                                        "content: 'Lorem ipsum';" +
                                    "} " +
                                    "@bottom-right {" +
                                        "content: 'Page ' counter(page) ' of ' counter(pages);" +
                                    "}" +
                                    "@bottom-center  {" +
                                        "content:element(footer);" +
                                    "}" +
                                "}" +
                                 "#footer {" +
                                    "position: running(footer);" +
                                "}" +
                                "table {" +
                                  "border-collapse: collapse;" +
                                "}" +

                                "table, th, td {" +
                                    "border: 1px solid black;" +
                                    "text-align: left;" +
                                    "vertical-align: top;" +
                                    "padding:5px;" +
                                "}" +
                            "</style>" +
                        "</head>" +
                        "<body>");
                    #endregion

                    string oemDealerQuery = string.Empty;



                    string fileAppoinmentDate = string.Empty;
                    string maxAppointmentdate = "select distinct  Convert(varchar(10),max(SlotDate),105) as MaxAppointmentdate from	HSRPrecordsonlystricker H with(nolock) join DealerAffixationCenter d  with(nolock) on H.HSRP_StateID=d.StateID and H.Affix_Id=d.DealerAffixationID  where  affix_id in (select DealerAffixationid from dealeraffixationcenter where rtolocationid in (select rtolocationid from rtolocation where navembid='" + Navembcode + "')) and  NewPdfRunningNo is null and  h.affix_id is not NULL";

                    DataTable dtmax = Utils.Utils.GetDataTable(maxAppointmentdate, CnnString);

                    if (dtmax.Rows.Count > 0)
                    {
                        fileAppoinmentDate = dtmax.Rows[0]["MaxAppointmentdate"].ToString();
                    }
                    fileName = "HomeDelivery" + "-" + fileAppoinmentDate + "-" + filePrefix + "-" + Navembcode + ".pdf";
                    filePath = dir + fileName;



                    oemDealerQuery = "select distinct d.oemid as oemid, (select name  from oemmaster where oemid=d.oemid) as oemname,d.dealerid as dealerid,d.DealerAffixationcenterName as Dealername,d.DealerAffixationCenterAddress as Address,d.DealerAffixationID as SubDealerId, H.dealerid AS ParentDealerId,SlotDate from	hsrprecordsonlystricker H with(nolock) join DealerAffixationCenter d  with(nolock) on H.HSRP_StateID=d.StateID and H.Affix_Id=d.DealerAffixationID where   NewPdfRunningNo is null and  isnull(BookMyHsrporderno,'')!='' and  h.affix_id is not NULL   and d.TypeofDelivery='Home' and h.created_date>'01-Nov-2020 00:00:00' and affix_id in (select DealerAffixationid from dealeraffixationcenter where rtolocationid in (select rtolocationid from rtolocation where navembid='" + Navembcode + "'))and h.slotdate<=dateadd(day,5,getdate())";

                    #region
                    DataTable dtOD = Utils.Utils.GetDataTable(oemDealerQuery, CnnString);

                    if (dtOD.Rows.Count > 0)
                    {
                        string allStrProductionSheetNo = string.Empty;
                        Session["ECWiseProductionsheet="] = null;
                        foreach (DataRow drOD in dtOD.Rows)
                        {

                            string oemid = drOD["oemid"].ToString().Trim();
                            string dealerid = drOD["dealerid"].ToString().Trim();
                            string oemname = drOD["oemname"].ToString().Trim();
                            string dealername = drOD["Dealername"].ToString().Trim();
                            string Address = drOD["Address"].ToString().Trim();

                            DateTime AppointmentDate = Convert.ToDateTime(drOD["SlotDate"].ToString().Trim());


                            string strsubaffixid = drOD["SubDealerId"].ToString();
                            string strParentDealerId = drOD["ParentDealerId"].ToString();

                            //start sql query
                            #region
                            string productionQuery = string.Empty;

                            DataTable dtProduction = new DataTable();


                            SqlConnection con = new SqlConnection(CnnString);
                            SqlCommand cmd = new SqlCommand("USP_BookMYHSRPProductionSheet_OS", con);
                            cmd.CommandType = CommandType.StoredProcedure;
                            con.Open();

                            cmd.Parameters.AddWithValue("@HSRP_StateID", HSRP_StateID);
                            cmd.Parameters.AddWithValue("@navembid", Navembcode);
                            cmd.Parameters.AddWithValue("@Dealerid", strParentDealerId);
                            cmd.Parameters.AddWithValue("@Affix_Id", strsubaffixid);
                            cmd.Parameters.AddWithValue("@AppointmentDate", AppointmentDate);
                            SqlDataAdapter da = new SqlDataAdapter(cmd);
                            // dtProduction = new DataTable();
                            da.Fill(dtProduction);
                            con.Close();

                            // string fileAppoinmentDate = AppointmentDate.ToString("dd/MM/yyyy");

                            #endregion
                            //end sql query

                            #region
                            //DataTable dtProduction = Utils.Utils.GetDataTable(productionQuery, CnnString);
                            if (dtProduction.Rows.Count > 0)
                            {

                                findRecord = true;
                                string strRunningNo = string.Empty;
                                //string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) as maxSheetNo from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + NAVEMBID + "'";
                                string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) from EmbossingCentersNew where    Emb_Center_Id='" + Navembcode + "' ";
                                string strCom = Utils.Utils.getScalarValue(strSel, CnnString);
                                //DataTable dtSheetNo = fillDataTable(strSel, CnnString);
                                if (strCom.Equals(0) || strCom.Length == 0)
                                {
                                    strRunningNo = "0000001";
                                }
                                else
                                {
                                    strRunningNo = string.Format("{0:0000000}", Convert.ToInt32(strCom) + 1);
                                }

                                string strRequeNo = "select (prefixtext+right('00000'+ convert(varchar,Lastreqno+1),5)) as Reqno from prefix_Requisition  where hsrp_stateid='" + ddlStateName.SelectedValue + "' and prefixfor='Req No'";
                                ReqNum = Utils.Utils.getScalarValue(strRequeNo, CnnString);

                                string strPRFIX = "select PrefixText from EmbossingCentersNew where  Emb_Center_Id='" + Navembcode + "'";
                                string strPRFIXCom = Utils.Utils.getScalarValue(strPRFIX, CnnString);
                                strProductionSheetNo = strPRFIXCom + strRunningNo;

                                string RTOLocationName = dtProduction.Rows[0]["RTOLocationName"].ToString();




                                #region


                                html.Append("<div style='page-break-before: avoid;page-break-inside: avoid;page-break-after: always;'>" +
                                        "<table style='width:100%;border: 0px;'>" +
                                            "<tr style='border: 0px;'>" +
                                                "<td colspan='7' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>Report Generation Date:</b> " + DateTime.Now.ToString("dd-MM-yyyy") + "</div>" +
                                                "</td>" +
                                                "<td colspan='7' style='border: 0px;'>" +
                                                    "<div style='float:right;width: 500px;word-wrap: break-word;'>" +
                                                        "<b>Production Sheet No: " + strProductionSheetNo + "<br /> </b>" +
                                                        "<b>Pin Code:" + dealername + "<br />  </b>" +
                                                        "<b>Pin Code Location:" + Address + "<br /> </b>" +

                                                    //"<b>Appointment Date:-:" + dtProduction.Rows[0]["AppointmentDate"].ToString() + "<br />  </b>" +
                                                    "</div>" +
                                                "</td>" +
                                            "</tr>" +
                                            "<tr style='border: 0px;'>" +
                                                "<td colspan='14' style='border: 0px;'>" +
                                                    "<div style='text-align:center;font-size:26px;'><b>Home Delivery Production Sheet : -</b> ROSMERTA SAFETY SYSTEMS LIMITED</div>" +
                                                "</td>" +
                                            "</tr>" +
                                            "<tr style='border: 0px;'>" +
                                                "<td colspan='9' style='border: 0px;'>" +
                                                    "<table style='border:0px;width:100%;'>" +
                                                        "<tr style='border:0px;'>" +
                                                            "<td style='border:0px;'><b>State Name :</b> " + HSRPStateName + "</td>" +

                                                            "<td style='border:0px; font-size:22px;'><b>Report Date: " + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</td> </b>" +

                                                              "<td style='border:0px; font-size:22px;'><b>Appointment Date:-: " + dtProduction.Rows[0]["AppointmentDate"].ToString() + "</td> </b>" +
                                                        "</tr>" +
                                                    "</table>" +
                                                "</td>" +
                                                "<td colspan='4' style='border: 0px;'>" +
                                                    "<div style='float:right'>" +

                                                        "VC:Vehicle Class<br />" +
                                                        //"VT:Vehicle Type<br />" +
                                                        "Front PS:Front Plate Size<br />" +
                                                        "Rear PS:Rear Plate Size<br />" +

                                                    "</div>" +
                                                "</td>" +
                                                "<td style='border: 0px;'></td>" +
                                            "</tr>" +
                                            "<tr style='border: 0px;'>" +
                                                 "<td colspan='14' style='border: 0px;'>" +
                                                    "<div style='text-align:left'>Location Name : " + RTOLocationName + "</div>" +
                                                "</td>" +
                                            "</tr>" +
                                            "<tr>" +
                                                "<td style='text-align:center'>SR.No</td>" +
                                                "<td>VC</td>" +
                                                "<td>Vehicle No</td>" +
                                                "<td>Fuel Type</td>" +
                                                 "<td>Front PS</td>" +
                                                "<td>Front Laser No</td>" +
                                                "<td>Rear PS</td>" +
                                                "<td>Rear Laser No.</td>" +

                                                "<td style='text-align:center'>Sticker Color</td>" +
                                                 "<td style='text-align:center'>Customer Name</td>" +

                                                   "<td style='width:8%;' style='text-align:center'>Appointment Time</td>" +
                                            "</tr>");
                                #endregion

                                #region
                                string strtvspono = "";

                                int j = 0;
                                int total = 0;
                                StringBuilder UpdateSQL = new StringBuilder();
                                for (int i = 0; i <= dtProduction.Rows.Count - 1; i++)
                                {

                                    j = j + 1;

                                    if (total == 12)
                                    {
                                        total = 0;




                                        html.Append("<div style='page-break-before: avoid;page-break-inside: avoid;page-break-after: always;'>" +
                                        "<table style='width:100%;border: 0px;'>" +
                                            "<tr style='border: 0px;'>" +
                                                "<td colspan='7' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>Report Generation Date:</b> " + DateTime.Now.ToString("dd-MM-yyyy") + "</div>" +
                                                "</td>" +
                                                "<td colspan='7' style='border: 0px;'>" +
                                                    "<div style='float:right;width: 500px;word-wrap: break-word;'>" +
                                                        "<b>Production Sheet No: " + strProductionSheetNo + "<br /> </b>" +
                                                        "<b>Pin Code:" + dealername + "<br />  </b>" +
                                                        "<b>Pin Code Location:" + Address + "<br /> </b>" +

                                                    //"<b>Appointment Date:-:" + dtProduction.Rows[0]["AppointmentDate"].ToString() + "<br />  </b>" +
                                                    "</div>" +
                                                "</td>" +
                                            "</tr>" +
                                            "<tr style='border: 0px;'>" +
                                                "<td colspan='14' style='border: 0px;'>" +
                                                    "<div style='text-align:center;font-size:26px;'><b>Home Delivery Production Sheet : -</b> ROSMERTA SAFETY SYSTEMS LIMITED</div>" +
                                                "</td>" +
                                            "</tr>" +
                                            "<tr style='border: 0px;'>" +
                                                "<td colspan='9' style='border: 0px;'>" +
                                                    "<table style='border:0px;width:100%;'>" +
                                                        "<tr style='border:0px;'>" +
                                                            "<td style='border:0px;'><b>State Name :</b> " + HSRPStateName + "</td>" +
                                                              "<td style='border:0px; font-size:22px;'><b>Report Date: " + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</td> </b>" +

                                                                "<td style='border:0px; font-size:22px;'><b>Appointment Date:-: " + dtProduction.Rows[0]["AppointmentDate"].ToString() + "</td> </b>" +
                                                        "</tr>" +
                                                    "</table>" +
                                                "</td>" +
                                                "<td colspan='4' style='border: 0px;'>" +
                                                    "<div style='float:right'>" +

                                                        "VC:Vehicle Class<br />" +
                                                        //"VT:Vehicle Type<br />" +
                                                        "Front PS:Front Plate Size<br />" +
                                                        "Rear PS:Rear Plate Size<br />" +

                                                    "</div>" +
                                                "</td>" +
                                                "<td style='border: 0px;'></td>" +
                                            "</tr>" +
                                            "<tr style='border: 0px;'>" +
                                                 "<td colspan='14' style='border: 0px;'>" +
                                                    "<div style='text-align:left'>Location Name : " + RTOLocationName + "</div>" +
                                                "</td>" +
                                            "</tr>" +
                                            "<tr>" +
                                                "<td style='text-align:center'>SR.No</td>" +
                                                //"<td>VC</td>" +
                                                "<td>Vehicle No</td>" +
                                                "<td>Fuel Type</td>" +
                                                // "<td>Front PS</td>" +
                                                //"<td>Front Laser No</td>" +
                                                //"<td>Rear PS</td>" +
                                                //"<td>Rear Laser No.</td>" +

                                                "<td style='text-align:center'>Sticker Color</td>" +
                                                 "<td style='text-align:center'>Customer Name</td>" +

                                                   "<td style='width:8%;' style='text-align:center'>Appointment Time</td>" +
                                            "</tr>");


                                    }

                                    total = total + 1;
                                    //foreach (DataRow drProduction in dtProduction.Rows)
                                    //{
                                    string HsrprecordID = dtProduction.Rows[i]["hsrprecordID"].ToString().Trim();
                                    string SRNo = dtProduction.Rows[i]["SRNo"].ToString().Trim();
                                    //string ORD = "";// drProduction["ORD"].ToString().Trim();
                                    string VC = dtProduction.Rows[i]["VehicleClass"].ToString().Trim();
                                    string VehicleNo = dtProduction.Rows[i]["VehRegNo"].ToString().Trim();
                                    //string VT = drProduction["VehicleType"].ToString().Trim();
                                    //string ChassisNo = drProduction["ChassisNo"].ToString().Trim();
                                    //string EngineNo = drProduction["EngineNo"].ToString().Trim();
                                    string FuelType = dtProduction.Rows[i]["FuelType"].ToString().Trim();
                                    string FrontPSize = dtProduction.Rows[i]["FrontProductCode"].ToString().Trim();
                                    //string FrontLaserNo = dtProduction.Rows[i]["HSRP_Front_LaserCode"].ToString().Trim();
                                    //string FS1 = dtProduction.Rows[i]["HSRP_Front_LaserCode"].ToString().Substring(0, 7);
                                    //string FS2 = dtProduction.Rows[i]["HSRP_Front_LaserCode"].ToString().Substring(7, 5);
                                    //string RearPSize = dtProduction.Rows[i]["RearProductCode"].ToString().Trim();
                                    //string RS1 = dtProduction.Rows[i]["HSRP_Rear_LaserCode"].ToString().Substring(0, 7);
                                    //string RS2 = dtProduction.Rows[i]["HSRP_Rear_LaserCode"].ToString().Substring(7, 5);
                                    //string RearLaserNo = dtProduction.Rows[i]["HSRP_Rear_LaserCode"].ToString().Trim();
                                    //string Amount = drProduction["roundoff_netamount"].ToString().Trim();
                                    //string OrderStatus = drProduction["OrderStatus"].ToString().Trim();
                                    string StickerColor = dtProduction.Rows[i]["stickerColor"].ToString().Trim();
                                    string CustomerName = dtProduction.Rows[i]["CustomerName"].ToString().Trim();
                                    //string AppointmentDate = drProduction["AppointmentDate"].ToString().Trim();
                                    string AppointmentTime = dtProduction.Rows[i]["AppointmentTime"].ToString().Trim();

                                    html.Append("<tr>" +
                                      "<td style='text-align:center'>" + SRNo + "</td>" +
                                       //"<td>" + VC + "</td>" +
                                       "<td > <b>" + VehicleNo + " </b> </td> " +
                                        "<td>" + FuelType + "</td>" +


                                        //"<td>" + FrontPSize + "</td>" +
                                        //"<td>" + FS1 + "<b>" + FS2 + "</b> </td>" +
                                        //"<td>" + RearPSize + "</td>" +
                                        // //"<td>" + RearLaserNo + "</td>" +
                                        //"<td>" + RS1 + "<b>" + RS2 + "</b> </td>" +

                                        "<td>" + StickerColor + "</td>" +
                                       "<td>" + CustomerName + "</td>" +

                                         "<td>" + AppointmentTime + "</td>" +
                                   "</tr>");

                                    //start updating hsrprecords 
                                    //string sqlUpdateHSRPRecords = "update hsrprecords set sendtoProductionStatus='Y', NAVPDFFlag='1', NewPdfRunningNo='" + strProductionSheetNo + "', Requisitionsheetno='" + ReqNum + "', " +
                                    //       "PdfDownloadDate=GetDate(), pdfFileName='" + fileName + "', PDFDownloadUserID='1' where hsrprecordID='" + HsrprecordID + "' ";
                                    //Utils.Utils.ExecNonQuery(sqlUpdateHSRPRecords, CnnString);   // uncomment after testing

                                    UpdateSQL.Append("update hsrprecordsonlystricker set  NewPdfRunningNo='" + strProductionSheetNo + "' where hsrprecordID='" + HsrprecordID + "';");

                                    // total = total + 1;

                                    //if(total>11)
                                    //{
                                    //    html.Append("<tr>" +

                                    //            "<td style='text-align:center'>SR.No</td>" +
                                    //            "<td>VC</td>" +
                                    //            "<td>Vehicle No</td>" +
                                    //            "<td>Fuel Type</td>" +
                                    //             "<td>Front PS</td>" +
                                    //            "<td>Front Laser No</td>" +
                                    //            "<td>Rear PS</td>" +
                                    //            "<td>Rear Laser No.</td>" +

                                    //            "<td style='text-align:center'>Sticker Color</td>" +
                                    //             "<td style='text-align:center'>Customer Name</td>" +

                                    //               "<td style='text-align:center'>Appointment Time</td>" +
                                    //        "</tr>");

                                    //}
                                    //end 

                                }

                                allStrProductionSheetNo += "," + strProductionSheetNo;

                                allStrProductionSheetNo = allStrProductionSheetNo.TrimStart(',');
                                Session["ECWiseProductionsheet="] = allStrProductionSheetNo;
                                #endregion
                                html.Append("</table>");

                                html.Append("</div>");

                                if (UpdateSQL.ToString().Length > 0)
                                {
                                    Utils.Utils.ExecNonQuery(UpdateSQL.ToString(), CnnString);
                                    string StrSqlUpdateECQuery = "update EmbossingCentersNew set NewProductionSheetRunningNo='" + strProductionSheetNo + "' " +
                                     "where  Emb_Center_Id='" + Navembcode + "'";
                                    Utils.Utils.ExecNonQuery(StrSqlUpdateECQuery, CnnString); // uncomment after testing
                                }

                                //string StrSqlUpdateECQuery = "update EmbossingCentersNew set NewProductionSheetRunningNo='" + strProductionSheetNo + "' " +
                                // "where  Emb_Center_Id='" + Navembcode + "'";
                                //Utils.Utils.ExecNonQuery(StrSqlUpdateECQuery, CnnString); // uncomment after testing
                            }

                            #endregion

                        }
                    }// close oemDealerQuery
                    #endregion



                    /*
                     * Close body & HTMl Tag
                     */


                    if (findRecord)
                    {






                        string strSaveSQlQuery = "insert into ProductionSheetAutoGenerated_List (HSRP_StateID, State_Code, RTO_ID, Emb_Center_Id, FileName, Productiondate, ProductionTime, OrderNo) values " +
                                    "('" + ddlStateName.SelectedValue + "', '" + HSRPStateShortName + "', '', '" + Navembcode + "', '" + fileName + "', '" + DateTime.Now.ToString("yyyy-MM-dd") + "', '" + DateTime.Now.ToString("HH:mm:ss") + "', '" + filePrefix + "') ";
                        Utils.Utils.ExecNonQuery(strSaveSQlQuery, CnnString); // uncomment after testing
                        Label1.Text = Label1.Text + "successfully created production sheet of State: " + HSRPStateShortName + ", , Emb_Center_Id: " + Navembcode + Environment.NewLine;

                        //lblLog.Text = lblLog.Text + Environment.NewLine + "Start Production Query at:" + DateTime.Now;



                        #region
                        try
                        {
                            if (!Directory.Exists(dir))
                            {


                                Directory.CreateDirectory(dir);

                            }

                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        try
                        {
                            var pdf = Pdf
                                    .From(html.ToString())
                                    .OfSize(PaperSize.A4)
                                    .WithTitle("Title")
                                    .WithoutOutline()
                                    .WithMargins(1.25.Centimeters())
                                    .Landscape()
                                    .Comressed()
                                    .Content();

                            FileStream readStream = File.Create(filePath);
                            BinaryWriter binaryWriter = new BinaryWriter(readStream);

                            // Write the binary data to the file
                            binaryWriter.Write(pdf);
                            binaryWriter.Close();
                            readStream.Close();
                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        #endregion
                    }

                    #endregion

                }//close foreach stateEcQuery
            }
        }

        protected void Button1_Click(object sender, EventArgs e)
        {


            if (ddlStateName.SelectedValue == "0")
            {
                lblErrMess.Text = String.Empty;
                lblErrMess.Text = "Please select State";
                return;
            }


            if (txtDate.Text == string.Empty)
            {
                lblErrMess.Text = String.Empty;
                lblErrMess.Text = "Please select Appointment Date";
                return;
            }
            string sqlPrefixQuery = "select top 1 CONVERT(INT, isnull(OrderNo,'0')) + 1 as orderNo from ProductionSheetAutoGenerated_List order by id desc";
            DataTable dtPrefix = Utils.Utils.GetDataTable(sqlPrefixQuery, CnnString);

            if (dtPrefix.Rows.Count > 0)
            {
                filePrefix = dtPrefix.Rows[0]["orderNo"].ToString();
            }
            else
            {
                filePrefix = "1";
            }

            //txtFilePrefix.Text = filePrefix;

            if (filePrefix.Length > 0)
            {
                SheetGenerationBookMyHSRPHDOSNEW();
            }
        }



        private void SheetGenerationBookMyHSRPHDOSNEW()
        {
            string stateECQuery = string.Empty;
            string ReqNum = string.Empty;

            string[] strArray = null;
            string Appdate = txtDate.Text;
            strArray = Appdate.Split('/');

            string selectedDate = string.Empty;

            string Year, month, day = string.Empty;
            day = strArray[0];
            month = strArray[1];
            Year = strArray[2];
            selectedDate = Year + "-" + month + "-" + day;

            string stateECShortName = "select distinct HSRP_StateID, 'EC'+HSRPStateShortName as NewHSRPStateShortName, HSRPStateShortName   from  HSRPState  where  HSRP_STateId='" + ddlStateName.SelectedValue + "'";

            DataTable dtECName = Utils.Utils.GetDataTable(stateECShortName, CnnString);

            string ShortECname = dtECName.Rows[0]["NewHSRPStateShortName"].ToString();
            // string HSRPStateShortName = dtECName.Rows[0]["HSRPStateShortName"].ToString();

            //stateECQuery = "select distinct a.HSRP_StateID, (select HSRPStateName from hsrpstate c where c.hsrp_stateid=a.hsrp_stateid) as HSRPStateName,(select HSRPStateShortName from hsrpstate d with(nolock) where d.hsrp_stateid='" + ddlStateName.SelectedValue + "') as HSRPStateShortName,   a.Navembid AS  navembcode from hsrprecords a where  left(a.Navembid,4) = '" + ShortECname + "'  and  NewPdfRunningNo is null and erpassigndate is not null  and (IsBookMyHsrpRecord='Y')  and OrderStatus='New Order' and    Navembid not like '%CODO%'  order by  a.HSRP_StateID";
            stateECQuery = "select distinct a.HSRP_StateID, (select HSRPStateName from hsrpstate c where c.hsrp_stateid=a.hsrp_stateid) as HSRPStateName,(select HSRPStateShortName from hsrpstate d with(nolock) where d.hsrp_stateid='" + ddlStateName.SelectedValue + "') as HSRPStateShortName,   (select navembid from rtolocation where RTOLocationID in (select rtolocationid from DealerAffixationCenter where DealerAffixationcenter.dealeraffixationid = a.affix_id)) AS  navembcode from hsrprecordsonlystricker a where convert(date,slotdate)='" + selectedDate + "' and   (select navembid from rtolocation where RTOLocationID in (select rtolocationid from DealerAffixationCenter where DealerAffixationcenter.dealeraffixationid=a.affix_id and Typeofdelivery='Home')) like '%" + ShortECname + "%'  and  NewPdfRunningNo is null   order by  a.HSRP_StateID";
            DataTable dtSE = Utils.Utils.GetDataTable(stateECQuery, CnnString);
            //string allStrProductionSheetNo = string.Empty;

            if (dtSE.Rows.Count > 0)
            {
                foreach (DataRow dr in dtSE.Rows)
                {
                    #region

                    string HSRP_StateID = dr["HSRP_StateID"].ToString().Trim();
                    string HSRPStateName = dr["HSRPStateName"].ToString().Trim();
                    string HSRPStateShortName = dr["HSRPStateShortName"].ToString().Trim();
                    //string RTOLocationID = dr["RTOLocationID"].ToString().Trim();
                    //string RTOLocationName = dr["RTOLocationName"].ToString().Trim();
                    //string NAVEMBID = dr["NAVEMBID"].ToString().Trim();
                    string Navembcode = dr["Navembcode"].ToString().Trim();
                    string dir = dirPath + DateTime.Now.ToString("yyyy-MM-dd") + "\\" + HSRPStateShortName + "\\";

                    string fileName = string.Empty;
                    string filePath = string.Empty;
                    // string fileName = filePrefix + "-" + Navembcode + ".pdf";



                    StringBuilder html = new StringBuilder();

                    Boolean findRecord = false;
                    string strProductionSheetNo = string.Empty;

                    /*
                    *  Start body & HTMl Tag
                    */
                    #region
                    html.Append(
                        "<!DOCTYPE html>" +
                        "<html>" +
                        "<head>" +
                            "<meta charset='UTF-8'><title>Title</title>" +
                            "<style>" +
                                "@page {" +
                                    /* headers*/
                                    "@top-left {" +
                                        "content: 'Left header';" +
                                    "}" +
                                    "@top-right {" +
                                        "content: 'Right header';" +
                                    "}" +

                                    /* footers */
                                    "@bottom-left {" +
                                        "content: 'Lorem ipsum';" +
                                    "} " +
                                    "@bottom-right {" +
                                        "content: 'Page ' counter(page) ' of ' counter(pages);" +
                                    "}" +
                                    "@bottom-center  {" +
                                        "content:element(footer);" +
                                    "}" +
                                "}" +
                                 "#footer {" +
                                    "position: running(footer);" +
                                "}" +
                                "table {" +
                                  "border-collapse: collapse;" +
                                "}" +

                                "table, th, td {" +
                                    "border: 1px solid black;" +
                                    "text-align: left;" +
                                    "vertical-align: top;" +
                                    "padding:5px;" +
                                "}" +
                            "</style>" +
                        "</head>" +
                        "<body>");
                    #endregion

                    string oemDealerQuery = string.Empty;



                    //string fileAppoinmentDate = string.Empty;
                    // string maxAppointmentdate = "select distinct  Convert(varchar(10),max(SlotBookingDate),105) as MaxAppointmentdate  " +
                    // " from	HSRPrecords H with(nolock) join DealerAffixationCenter d  with(nolock) on H.HSRP_StateID=d.StateID and H.Affix_Id=d.DealerAffixationID  join BookMYHSRPappointment B on H.orderno=B.orderno  " +
                    //" where   Navembid='" + Navembcode + "'  and Navembid not like '%CODO%'     and  NewPdfRunningNo is null and erpassigndate is not null and  h.OrderStatus='New Order' and  IsBookMyHsrpRecord='Y' and  h.affix_id is not NULL  and   d.TypeofDelivery='Home' ";


                    string fileAppoinmentDate = string.Empty;
                    fileAppoinmentDate = selectedDate;
                    //string maxAppointmentdate = "select distinct  Convert(varchar(10),max(SlotDate),105) as MaxAppointmentdate from	HSRPrecordsonlystricker H with(nolock) join DealerAffixationCenter d  with(nolock) on H.HSRP_StateID=d.StateID and H.Affix_Id=d.DealerAffixationID  where  affix_id in (select DealerAffixationid from dealeraffixationcenter where rtolocationid in (select rtolocationid from rtolocation where navembid='" + Navembcode + "')) and  NewPdfRunningNo is null and  h.affix_id is not NULL";
                    //DataTable dtmax = Utils.Utils.GetDataTable(maxAppointmentdate, CnnString);

                    //if (dtmax.Rows.Count > 0)
                    //{
                    //    fileAppoinmentDate = dtmax.Rows[0]["MaxAppointmentdate"].ToString();
                    //}
                    fileName = "Home Delivery Sticker" + "-" + fileAppoinmentDate + "-" + filePrefix + "-" + Navembcode + ".pdf";
                    filePath = dir + fileName;



                    //oemDealerQuery = "select distinct d.oemid as oemid, (select name  from oemmaster where oemid=d.oemid) as oemname,d.dealerid as dealerid,d.DealerAffixationcenterName as Dealername,  d.DealerAffixationCenterAddress as Address, " +
                    // "d.DealerAffixationID as SubDealerId, H.dealerid AS ParentDealerId,SlotBookingDate from	HSRPrecords H with(nolock) join DealerAffixationCenter d  with(nolock) on H.HSRP_StateID=d.StateID and H.Affix_Id=d.DealerAffixationID  join BookMYHSRPappointment B on H.orderno=B.orderno  " +
                    //" where   Navembid='" + Navembcode + "'  and Navembid not like '%CODO%'     and  NewPdfRunningNo is null and erpassigndate is not null and  h.OrderStatus='New Order' and  IsBookMyHsrpRecord='Y' and  h.affix_id is not NULL   and d.TypeofDelivery='Home' ";

                    oemDealerQuery = "select distinct d.oemid as oemid, (select name  from oemmaster where oemid=d.oemid) as oemname,d.dealerid as dealerid,d.DealerAffixationcenterName as Dealername,d.DealerAffixationCenterAddress as Address,d.DealerAffixationID as SubDealerId, H.dealerid AS ParentDealerId,slotdate as SlotBookingDate from	hsrprecordsonlystricker H with(nolock) join DealerAffixationCenter d  with(nolock) on H.HSRP_StateID=d.StateID and H.Affix_Id=d.DealerAffixationID where   NewPdfRunningNo is null and  isnull(BookMyHsrporderno,'')!='' and  h.affix_id is not NULL   and d.TypeofDelivery='Home' and h.created_date>'01-Nov-2020 00:00:00' and affix_id in (select DealerAffixationid from dealeraffixationcenter where rtolocationid in (select rtolocationid from rtolocation where navembid='" + Navembcode + "'))and convert(date,h.slotdate)='" + selectedDate + "'";


                    #region
                    DataTable dtOD = Utils.Utils.GetDataTable(oemDealerQuery, CnnString);

                    if (dtOD.Rows.Count > 0)
                    {
                        string allStrProductionSheetNo = string.Empty;
                        Session["ECWiseProductionsheet="] = null;
                        foreach (DataRow drOD in dtOD.Rows)
                        {

                            string oemid = drOD["oemid"].ToString().Trim();
                            string dealerid = drOD["dealerid"].ToString().Trim();
                            string oemname = drOD["oemname"].ToString().Trim();
                            string dealername = drOD["Dealername"].ToString().Trim();
                            string Address = drOD["Address"].ToString().Trim();

                            DateTime AppointmentDate = Convert.ToDateTime(drOD["SlotBookingDate"].ToString().Trim());

                            string displayAppointmentDate = AppointmentDate.ToString("dd/MM/yyyy");
                            string strsubaffixid = drOD["SubDealerId"].ToString();
                            string strParentDealerId = drOD["ParentDealerId"].ToString();

                            //start sql query
                            #region
                            string productionQuery = string.Empty;

                            DataTable dtProduction = new DataTable();


                            SqlConnection con = new SqlConnection(CnnString);
                            SqlCommand cmd = new SqlCommand("USP_BookMYHSRPProductionSheet_OS", con);
                            cmd.CommandType = CommandType.StoredProcedure;
                            con.Open();

                            cmd.Parameters.AddWithValue("@HSRP_StateID", HSRP_StateID);
                            cmd.Parameters.AddWithValue("@navembid", Navembcode);
                            cmd.Parameters.AddWithValue("@Dealerid", strParentDealerId);
                            cmd.Parameters.AddWithValue("@Affix_Id", strsubaffixid);
                            cmd.Parameters.AddWithValue("@AppointmentDate", AppointmentDate);
                            SqlDataAdapter da = new SqlDataAdapter(cmd);
                            // dtProduction = new DataTable();
                            da.Fill(dtProduction);
                            con.Close();

                            // string fileAppoinmentDate = AppointmentDate.ToString("dd/MM/yyyy");

                            #endregion
                            //end sql query

                            #region
                            //DataTable dtProduction = Utils.Utils.GetDataTable(productionQuery, CnnString);
                            if (dtProduction.Rows.Count > 0)
                            {

                                findRecord = true;
                                string strRunningNo = string.Empty;
                                //string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) as maxSheetNo from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + NAVEMBID + "'";
                                string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) from EmbossingCentersNew where    Emb_Center_Id='" + Navembcode + "' ";
                                string strCom = Utils.Utils.getScalarValue(strSel, CnnString);
                                //DataTable dtSheetNo = fillDataTable(strSel, CnnString);
                                if (strCom.Equals(0) || strCom.Length == 0)
                                {
                                    strRunningNo = "0000001";
                                }
                                else
                                {
                                    strRunningNo = string.Format("{0:0000000}", Convert.ToInt32(strCom) + 1);
                                }

                                string strRequeNo = "select (prefixtext+right('00000'+ convert(varchar,Lastreqno+1),5)) as Reqno from prefix_Requisition  where hsrp_stateid='" + ddlStateName.SelectedValue + "' and prefixfor='Req No'";
                                ReqNum = Utils.Utils.getScalarValue(strRequeNo, CnnString);

                                string strPRFIX = "select PrefixText from EmbossingCentersNew where  Emb_Center_Id='" + Navembcode + "'";
                                string strPRFIXCom = Utils.Utils.getScalarValue(strPRFIX, CnnString);
                                strProductionSheetNo = strPRFIXCom + strRunningNo;

                                //string RTOLocationName = dtProduction.Rows[0]["RTOLocationName"].ToString();




                                #region



                                html.Append("<div style='page-break-before: avoid;page-break-inside: avoid;page-break-after: always;'>" +
                                        "<table style='width:100%;border: 0px;'>" +
                                            "<tr style='border: 0px;'>" +
                                                "<td colspan='7' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>Report Generation Date:</b> " + DateTime.Now.ToString("dd-MM-yyyy") + "</div>" +
                                                "</td>" +
                                                "<td colspan='7' style='border: 0px;'>" +
                                                    "<div style='float:right;width: 500px;word-wrap: break-word;'>" +
                                                        "<b>Production Sheet No: " + strProductionSheetNo + "<br /> </b>" +
                                                        "<b>Pin Code:" + dealername + "<br />  </b>" +
                                                        "<b>Pin Code Location:" + Address + "<br /> </b>" +

                                                    // "<b>Appointment Date:-:" + dtProduction.Rows[0]["AppointmentDate"].ToString() + "<br />  </b>" +
                                                    "</div>" +
                                                "</td>" +
                                            "</tr>" +
                                            "<tr style='border: 0px;'>" +
                                                "<td colspan='6' style='border: 0px;'>" +
                                                    "<div style='text-align:center;font-size:24px;'><b>Home Delivery Sticker Production Sheet : -</b> ROSMERTA SAFETY SYSTEMS LIMITED</div>" +
                                                "</td>" +
                                            "</tr>" +
                                            "<tr style='border: 0px;'>" +
                                                "<td colspan='13' style='border: 0px;'>" +
                                                    "<table style='border:0px;width:100%;'>" +
                                                        "<tr style='border:0px;'>" +
                                                            "<td style='border:0px;'><b>State Name :</b> " + HSRPStateName + "</td>" +

                                                            "<td style='border:0px; font-size:22px;'><b>Report Date: " + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</td> </b>" +

                                                              "<td style='border:0px; font-size:22px;'><b>Appointment Date:-: " + dtProduction.Rows[0]["AppointmentDate"].ToString() + "</td> </b>" +
                                                        "</tr>" +
                                                    "</table>" +
                                                "</td>" +
                                            //            "<td colspan='4' style='border: 0px;'>" +
                                            //                "<div style='float:right'>" +

                                            ////                    "VC:Vehicle Class<br />" +
                                            //////"VT:Vehicle Type<br />" +
                                            ////                    "Front PS:Front Plate Size<br />" +
                                            ////                    "Rear PS:Rear Plate Size<br />" +

                                            //                "</div>" +
                                            //"</td>" +
                                            //"<td style='border: 0px;'></td>" +
                                            "</tr>" +
                                            //"<tr style='border: 0px;'>" +
                                            //     "<td colspan='14' style='border: 0px;'>" +
                                            //        "<div style='text-align:left'>Location Name : " + RTOLocationName + "</div>" +
                                            //    "</td>" +
                                            //"</tr>" +
                                            "<tr>" +
                                                "<td style='text-align:center'>SR.No</td>" +
                                                //"<td>VC</td>" +
                                                "<td>Vehicle No</td>" +
                                                "<td>Fuel Type</td>" +
                                                // "<td>Front PS</td>" +
                                                //"<td>Front Laser No</td>" +
                                                //"<td>Rear PS</td>" +
                                                //"<td>Rear Laser No.</td>" +

                                                "<td style='text-align:center'>Sticker Color</td>" +
                                                 "<td style='text-align:center'>Customer Name</td>" +

                                                   "<td style='width:8%;' style='text-align:center'>Appointment Time</td>" +
                                            "</tr>");
                                #endregion

                                #region
                                string strtvspono = "";

                                int j = 0;
                                int total = 0;
                                StringBuilder UpdateSQL = new StringBuilder();
                                for (int i = 0; i <= dtProduction.Rows.Count - 1; i++)
                                {

                                    j = j + 1;

                                    if (total == 13)
                                    {
                                        total = 0;




                                        html.Append("<div style='page-break-before: avoid;page-break-inside: avoid;page-break-after: always;'>" +
                                        "<table style='width:100%;border: 0px;'>" +
                                            "<tr style='border: 0px;'>" +
                                                "<td colspan='7' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>Report Generation Date:</b> " + DateTime.Now.ToString("dd-MM-yyyy") + "</div>" +
                                                "</td>" +
                                                "<td colspan='7' style='border: 0px;'>" +
                                                    "<div style='float:right;width: 500px;word-wrap: break-word;'>" +
                                                        "<b>Production Sheet No: " + strProductionSheetNo + "<br /> </b>" +
                                                        "<b>Pin Code:" + dealername + "<br />  </b>" +
                                                        "<b>Pin Code Location:" + Address + "<br /> </b>" +

                                                    // "<b>Appointment Date:-:" + dtProduction.Rows[0]["AppointmentDate"].ToString() + "<br />  </b>" +
                                                    "</div>" +
                                                "</td>" +
                                            "</tr>" +
                                            "<tr style='border: 0px;'>" +
                                                "<td colspan='14' style='border: 0px;'>" +
                                                    "<div style='text-align:center;font-size:24px;'><b>Book My HSRP Sticker Production Sheet : -</b> ROSMERTA SAFETY SYSTEMS LIMITED</div>" +
                                                "</td>" +
                                            "</tr>" +
                                            "<tr style='border: 0px;'>" +
                                                "<td colspan='13' style='border: 0px;'>" +
                                                    "<table style='border:0px;width:100%;'>" +
                                                        "<tr style='border:0px;'>" +
                                                            "<td style='border:0px;'><b>State Name :</b> " + HSRPStateName + "</td>" +
                                                              "<td style='border:0px; font-size:22px;'><b>Report Date: " + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</td> </b>" +

                                                                "<td style='border:0px; font-size:22px;'><b>Appointment Date:-: " + dtProduction.Rows[0]["AppointmentDate"].ToString() + "</td> </b>" +
                                                        "</tr>" +
                                                    "</table>" +
                                                "</td>" +
                                                //    "<td colspan='4' style='border: 0px;'>" +
                                                ////        "<div style='float:right'>" +

                                                ////            //"VC:Vehicle Class<br />" +
                                                //////"VT:Vehicle Type<br />" +
                                                ////            //"Front PS:Front Plate Size<br />" +
                                                ////            //"Rear PS:Rear Plate Size<br />" +

                                                ////        "</div>" +
                                                //    "</td>" +
                                                "<td style='border: 0px;'></td>" +
                                            "</tr>" +
                                            //"<tr style='border: 0px;'>" +
                                            //     "<td colspan='14' style='border: 0px;'>" +
                                            //        "<div style='text-align:left'>Location Name : " + RTOLocationName + "</div>" +
                                            //    "</td>" +
                                            //"</tr>" +
                                            "<tr>" +
                                                "<td style='text-align:center'>SR.No</td>" +
                                                //"<td>VC</td>" +
                                                "<td>Vehicle No</td>" +
                                                "<td>Fuel Type</td>" +


                                                "<td style='text-align:center'>Sticker Color</td>" +
                                                 "<td style='text-align:center'>Customer Name</td>" +

                                                   "<td style='width:8%;' style='text-align:center'>Appointment Time</td>" +
                                            "</tr>");


                                    }

                                    total = total + 1;
                                    //foreach (DataRow drProduction in dtProduction.Rows)
                                    //{
                                    string HsrprecordID = dtProduction.Rows[i]["hsrprecordID"].ToString().Trim();
                                    string SRNo = dtProduction.Rows[i]["SRNo"].ToString().Trim();

                                    string VehicleNo = dtProduction.Rows[i]["VehRegNo"].ToString().Trim();

                                    string FuelType = dtProduction.Rows[i]["FuelType"].ToString().Trim();

                                    string StickerColor = dtProduction.Rows[i]["stickerColor"].ToString().Trim();
                                    string CustomerName = dtProduction.Rows[i]["CustomerName"].ToString().Trim();
                                    //string AppointmentDate = drProduction["AppointmentDate"].ToString().Trim();
                                    string AppointmentTime = dtProduction.Rows[i]["AppointmentTime"].ToString().Trim();

                                    html.Append("<tr>" +
                                      "<td style='text-align:center'>" + SRNo + "</td>" +
                                       //"<td>" + VC + "</td>" +
                                       "<td > <b>" + VehicleNo + " </b> </td> " +
                                        "<td>" + FuelType + "</td>" +


                                        //"<td>" + FrontPSize + "</td>" +
                                        //"<td>" + FS1 + "<b>" + FS2 + "</b> </td>" +
                                        //"<td>" + RearPSize + "</td>" +
                                        // //"<td>" + RearLaserNo + "</td>" +
                                        //"<td>" + RS1 + "<b>" + RS2 + "</b> </td>" +

                                        "<td>" + StickerColor + "</td>" +
                                       "<td>" + CustomerName + "</td>" +

                                         "<td>" + AppointmentTime + "</td>" +
                                   "</tr>");

                                    //start updating hsrprecords 
                                    //string sqlUpdateHSRPRecords = "update hsrprecords set sendtoProductionStatus='Y', NAVPDFFlag='1', NewPdfRunningNo='" + strProductionSheetNo + "', Requisitionsheetno='" + ReqNum + "', " +
                                    //       "PdfDownloadDate=GetDate(), pdfFileName='" + fileName + "', PDFDownloadUserID='1' where hsrprecordID='" + HsrprecordID + "' ";
                                    //Utils.Utils.ExecNonQuery(sqlUpdateHSRPRecords, CnnString);   // uncomment after testing

                                    UpdateSQL.Append("update hsrprecordsonlystricker set  NewPdfRunningNo='" + strProductionSheetNo + "',stickerPSDownloaddate=GetDate(), stickerPSPDffilename='" + fileName + "' where ID='" + HsrprecordID + "';");

                                }

                                allStrProductionSheetNo += "," + strProductionSheetNo;

                                allStrProductionSheetNo = allStrProductionSheetNo.TrimStart(',');
                                Session["ECWiseProductionsheet="] = allStrProductionSheetNo;
                                #endregion
                                html.Append("</table>");

                                html.Append("</div>");

                                if (UpdateSQL.ToString().Length > 0)
                                {
                                    Utils.Utils.ExecNonQuery(UpdateSQL.ToString(), CnnString);
                                    string StrSqlUpdateECQuery = "update EmbossingCentersNew set NewProductionSheetRunningNo='" + strProductionSheetNo + "' " +
                                     "where  Emb_Center_Id='" + Navembcode + "'";
                                    Utils.Utils.ExecNonQuery(StrSqlUpdateECQuery, CnnString); // uncomment after testing
                                }

                                //string StrSqlUpdateECQuery = "update EmbossingCentersNew set NewProductionSheetRunningNo='" + strProductionSheetNo + "' " +
                                // "where  Emb_Center_Id='" + Navembcode + "'";
                                //Utils.Utils.ExecNonQuery(StrSqlUpdateECQuery, CnnString); // uncomment after testing
                            }

                            #endregion

                        }
                    }// close oemDealerQuery
                    #endregion
                    //   #region "Req Generate"
                    //   string strSqlQuery1 = "select CompanyName from hsrpstate where hsrp_stateid='" + ddlStateName.SelectedValue + "'";
                    //   strCompanyName = Utils.Utils.getScalarValue(strSqlQuery1, CnnString);

                    //   string strEMBName = " select EmbCenterName from EmbossingCentersNew where   Emb_Center_Id='" + Navembcode + "'";
                    //   string strComNew = Utils.Utils.getScalarValue(strEMBName, CnnString);

                    //   string strReqNo = "select (prefixtext+right('00000'+ convert(varchar,Lastreqno+1),5)) as Reqno from prefix_Requisition  where hsrp_stateid='" + ddlStateName.SelectedValue + "' and prefixfor='Req No'";
                    //   string strReqNumber = Utils.Utils.getScalarValue(strReqNo, CnnString);

                    //   SQLString = "Exec [laserreqSlip1DashBoardBookMYHSRP]  '" + Navembcode + "' ,  '" + ReqNum + "'";
                    //   DataTable dtResult = Utils.Utils.GetDataTable(SQLString, CnnString);
                    //   string strQuery = string.Empty;
                    //   string strRtoLocationName = string.Empty;
                    //   int Itotal = 0;

                    //   html.Append("<div style='width:100%;height:100%;'>" +
                    //                       "<table style='width:100%'>" +

                    //                           "<tr>" +
                    //                               "<td colspan='12'>" +
                    //                                   "<div style='text-align:center;padding:8px;'>" +
                    //                                       "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + strCompanyName + "</b>" +

                    //                                       "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                    //                                        "<b style='font-size:20px;'>" + "MATERIAL REQUSITION NOTE" + "</b>" +

                    //                                   "</div>" +
                    //                               "</td>" +
                    //                           "</tr>" +

                    //                           "<tr>" +
                    //                               "<td colspan='12'>" +
                    //                                   "<div style='text-align:center;padding:8px;'>" +
                    //                                       "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Production Sheet Date :" + DateTime.Now.ToString("dd-MM-yyyy") + "</b>" +
                    //                                       "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                    //                                        "<b style='font-size:20px;'>" + "" + "</b>" +
                    //                                   "</div>" +
                    //                               "</td>" +
                    //                           "</tr>" +
                    //                           "<tr>" +
                    //                               "<td colspan='6'>" +
                    //                                   "<div style='text-align:left;'>" +
                    //                                       "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>REQ.NO:-</b>" +
                    //                                       "" + strReqNumber + "" +
                    //                                   "</div>" +
                    //                               "</td>" +
                    //                               "<td colspan='6'>" +
                    //                                   "<div style='text-align:left;'>" +
                    //                                       "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>Embossing Center:</b>" +
                    //                                       "" + strComNew + "" +
                    //                                   "</div>" +
                    //                               "</td>" +
                    //                           "</tr>" +
                    //                           "<tr>" +
                    //                               "<td colspan='12'>" +
                    //                                   "<div style='text-align:left;padding:2px;'>" +
                    //                                       "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "State:" + ddlStateName.SelectedItem.Text + "</b>" +
                    //                                       "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                    //                                        "<b style='font-size:20px;'>" + "" + "</b>" +
                    //                                   "</div>" +
                    //                               "</td>" +
                    //                           "</tr>" +
                    //                            "<tr>" +
                    //                               "<td colspan='12'>" +
                    //                                   "<div style='text-align:left;padding:8px;'>" +
                    //                                       "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + " " + "</b>" +
                    //                                       "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                    //                                        "<b style='font-size:20px;'>" + "" + "</b>" +
                    //                                   "</div>" +
                    //                               "</td>" +
                    //                           "</tr>" +

                    //                           "<tr>" +
                    //                               "<td colspan='1' style='text-align:center'>SR.N.</td>" +
                    //                               "<td colspan='3'>Product Size</td>" +
                    //                               "<td colspan='1'>Laser Count</td>" +
                    //                               "<td colspan='1'>Start Laser No</td>" +
                    //                               "<td colspan='1'>End Laser No</td>" +

                    //                           "</tr>");


                    //   if (dtResult.Rows.Count > 0)
                    //   {
                    //       for (int i = 0; i < dtResult.Rows.Count; i++)
                    //       {
                    //           string ID = dtResult.Rows[i]["ID"].ToString();
                    //           string productcode = dtResult.Rows[i]["productcode"].ToString();
                    //           string LaserCount = dtResult.Rows[i]["LaserCount"].ToString();
                    //           Itotal = Convert.ToInt32(dtResult.Rows[i]["Total"].ToString());
                    //           string BeginLaser = dtResult.Rows[i]["BeginLaser"].ToString();
                    //           string EndLaser = dtResult.Rows[i]["EndLaser"].ToString();




                    //           html.Append("<tr>" +
                    //              "<td colspan='1' style='text-align:left'>" + ID + "</td>" +
                    //              "<td colspan='3'>" + productcode + "</td>" +

                    //              "<td colspan='1'>" + LaserCount + "</td>" +
                    //              "<td colspan='1'>" + BeginLaser + "</td>" +
                    //              "<td colspan='1'>" + EndLaser + "</td>" +

                    //          "</tr>");
                    //       }

                    //       try
                    //       {
                    //           //start updating hsrprecords 
                    //           string Query = "update prefix_Requisition set lastno=lastno+1,Lastreqno=Lastreqno+1 where hsrp_stateid='" + ddlStateName.SelectedValue + "' and prefixfor='Req No'";
                    //           Utils.Utils.ExecNonQuery(Query, CnnString);
                    //       }
                    //       catch (Exception ev)
                    //       {
                    //           Label1.Text = "prefix Requisition update error: " + ev.Message;
                    //       }
                    //   }
                    //   html.Append("<tr>" +
                    //            "<td colspan='1' style='text-align:center' > " + "<b>Grand Total:</b>" + "</td>" +
                    //            "<td colspan='3'>" + "" + "</td>" +

                    //            "<td colspan='1'>" + Itotal + "</td>" +
                    //            "<td colspan='1'>" + " " + "</td>" +
                    //            "<td colspan='1'>" + " " + "</td>" +

                    //        "</tr>");




                    //   html.Append("<tr>" +
                    //    "<td colspan='2' > " + "<b>REQUESTED BY </b>" + "</td>" +
                    //    "<td colspan='2'>" + "<b>AUTHORIZED BY </b>" + "</td>" +

                    //    "<td colspan='2'>" + "<b>ISSUED BY </b>" + "</td>" +
                    //    "<td colspan='2'>" + "<b>RECEIVED BY</b>" + "</td>" +


                    //"</tr>");

                    //   html.Append("<tr>" +
                    //                             "<td colspan='12'>" +
                    //                                 "<div style='text-align:left;padding:2px;'>" +
                    //                                     "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Name" + "</b>" +

                    //                                 "</div>" +
                    //                             "</td>" +
                    //                         "</tr>");

                    //   html.Append("<tr>" +
                    //                            "<td colspan='12'>" +
                    //                                "<div style='text-align:left;padding:2px;'>" +
                    //                                    "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Designation" + "</b>" +

                    //                                "</div>" +
                    //                            "</td>" +
                    //                        "</tr>");

                    //   html.Append("<tr>" +
                    //                          "<td colspan='12'>" +
                    //                              "<div style='text-align:right;padding:8px;'>" +
                    //                                  "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Download By: Admin " + "</b>" +

                    //                              "</div>" +
                    //                          "</td>" +
                    //                      "</tr>");

                    //   html.Append("<tr>" +
                    //                         "<td colspan='12'>" +
                    //                             "<div style='text-align:right;padding:8px;'>" +
                    //                                 "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Sheet Generated By :" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</b>" +

                    //                             "</div>" +
                    //                         "</td>" +
                    //                     "</tr>");


                    //   html.Append("<tr style='visibility: hidden;'>" + "<td  ><p style=\"page-break-after:always\"/></td></tr>");



                    //   //html.Append("</table>");

                    //   //html.Append("</div>");

                    //   #region "PS Summary Report"
                    //   string PS = string.Empty;
                    //   if (Session["ECWiseProductionsheet="] != null)
                    //   {
                    //       PS = Session["ECWiseProductionsheet="].ToString();
                    //   }
                    //   /*
                    //    * Close body & HTMl Tag
                    //    */

                    //   SQLString = "Exec [USP_BMPSSummaryHomedelivery]  '" + PS + "' ";
                    //   DataTable dtResultsummary = Utils.Utils.GetDataTable(SQLString, CnnString);

                    //   html.Append("<div style='width:100%;height:100%;'>" +
                    //                "<table style='width:100%'>" +

                    //                    "<tr>" +
                    //                        "<td colspan='12'>" +
                    //                            "<div style='text-align:center;padding:8px;'>" +
                    //                                "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Home Delivery Summary Report" + "</b>" +

                    //                            "</div>" +
                    //                        "</td>" +
                    //                    "</tr>" +

                    //                     "<tr>" +
                    //                               "<td colspan='1' style='text-align:center'>PS NO</td>" +
                    //                               "<td colspan='3'>Affixation Pin Code</td>" +
                    //                               "<td colspan='1'>2W</td>" +
                    //                               "<td colspan='1'>3W</td>" +
                    //                               "<td colspan='1'>4W</td>" +
                    //                               "<td colspan='1'>OTH</td>" +
                    //                                 "<td colspan='1'>Total</td>" +

                    //                           "</tr>");

                    //   if (dtResultsummary.Rows.Count > 0)
                    //   {
                    //       for (int i = 0; i < dtResultsummary.Rows.Count; i++)
                    //       {
                    //           string PSNo = dtResultsummary.Rows[i]["PSNo"].ToString();
                    //           string AffixationdealerName = dtResultsummary.Rows[i]["Affixation Pin Code"].ToString();
                    //           string W2 = dtResultsummary.Rows[i]["2W"].ToString();
                    //           string W3 = dtResultsummary.Rows[i]["3W"].ToString();
                    //           string W4 = dtResultsummary.Rows[i]["4W"].ToString();
                    //           string OTH = dtResultsummary.Rows[i]["OTH"].ToString();
                    //           string Total = dtResultsummary.Rows[i]["Total"].ToString();




                    //           html.Append("<tr>" +
                    //              "<td colspan='1' style='text-align:left'>" + PSNo + "</td>" +
                    //              "<td colspan='3'>" + AffixationdealerName + "</td>" +

                    //              "<td colspan='1'>" + W2 + "</td>" +
                    //              "<td colspan='1'>" + W3 + "</td>" +
                    //              "<td colspan='1'>" + W4 + "</td>" +
                    //               "<td colspan='1'>" + OTH + "</td>" +
                    //                "<td colspan='1'>" + Total + "</td>" +

                    //          "</tr>");
                    //       }
                    //   }

                    //   html.Append("</table>");

                    //   html.Append("</div>");

                    //   #endregion



                    //   #endregion

                    /*
                     * Close body & HTMl Tag
                     */


                    if (findRecord)
                    {






                        string strSaveSQlQuery = "insert into ProductionSheetAutoGenerated_List (HSRP_StateID, State_Code, RTO_ID, Emb_Center_Id, FileName, Productiondate, ProductionTime, OrderNo) values " +
                                    "('" + ddlStateName.SelectedValue + "', '" + HSRPStateShortName + "', '', '" + Navembcode + "', '" + fileName + "', '" + DateTime.Now.ToString("yyyy-MM-dd") + "', '" + DateTime.Now.ToString("HH:mm:ss") + "', '" + filePrefix + "') ";
                        Utils.Utils.ExecNonQuery(strSaveSQlQuery, CnnString); // uncomment after testing
                        Label1.Text = Label1.Text + "successfully created Home Delivery Sticker production sheet of State: " + HSRPStateShortName + ", , Emb_Center_Id: " + Navembcode + Environment.NewLine;

                        //lblLog.Text = lblLog.Text + Environment.NewLine + "Start Production Query at:" + DateTime.Now;



                        #region
                        try
                        {
                            if (!Directory.Exists(dir))
                            {


                                Directory.CreateDirectory(dir);

                            }

                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        try
                        {
                            var pdf = Pdf
                                    .From(html.ToString())
                                    .OfSize(PaperSize.A4)
                                    .WithTitle("Title")
                                    .WithoutOutline()
                                    .WithMargins(1.25.Centimeters())
                                    .Landscape()
                                    .Comressed()
                                    .Content();

                            FileStream readStream = File.Create(filePath);
                            BinaryWriter binaryWriter = new BinaryWriter(readStream);

                            // Write the binary data to the file
                            binaryWriter.Write(pdf);
                            binaryWriter.Close();
                            readStream.Close();
                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        #endregion
                    }

                    #endregion

                }//close foreach stateEcQuery
            }
        }


        private void SheetGenerationBookMyHSRPOSNew()
        {
            string stateECQuery = string.Empty;
            string ReqNum = string.Empty;

            string[] strArray = null;
            string Appdate = txtDate.Text;
            strArray = Appdate.Split('/');

            string selectedDate = string.Empty;

            string Year, month, day = string.Empty;
            day = strArray[0];
            month = strArray[1];
            Year = strArray[2];
            selectedDate = Year + "-" + month + "-" + day;

            string stateECShortName = "select distinct HSRP_StateID, 'EC'+HSRPStateShortName as NewHSRPStateShortName, HSRPStateShortName   from  HSRPState  where  HSRP_STateId='" + ddlStateName.SelectedValue + "'";

            DataTable dtECName = Utils.Utils.GetDataTable(stateECShortName, CnnString);

            string ShortECname = dtECName.Rows[0]["NewHSRPStateShortName"].ToString();
            // string HSRPStateShortName = dtECName.Rows[0]["HSRPStateShortName"].ToString();

            //stateECQuery = "select distinct a.HSRP_StateID, (select HSRPStateName from hsrpstate c where c.hsrp_stateid=a.hsrp_stateid) as HSRPStateName,(select HSRPStateShortName from hsrpstate d with(nolock) where d.hsrp_stateid='" + ddlStateName.SelectedValue + "') as HSRPStateShortName,   a.Navembid AS  navembcode from hsrprecords a where  left(a.Navembid,4) = '" + ShortECname + "'  and  NewPdfRunningNo is null and erpassigndate is not null  and (IsBookMyHsrpRecord='Y')  and OrderStatus='New Order' and    Navembid not like '%CODO%'  order by  a.HSRP_StateID";
            stateECQuery = "select distinct a.HSRP_StateID, (select HSRPStateName from hsrpstate c where c.hsrp_stateid=a.hsrp_stateid) as HSRPStateName,(select HSRPStateShortName from hsrpstate d with(nolock) where d.hsrp_stateid='" + ddlStateName.SelectedValue + "') as HSRPStateShortName,   (select navembid from rtolocation where RTOLocationID in (select rtolocationid from DealerAffixationCenter where DealerAffixationcenter.dealeraffixationid = a.affix_id)) AS  navembcode from hsrprecordsonlystricker a where  BookMyHSRPOrderNo is not  null and (select navembid from rtolocation where RTOLocationID in (select rtolocationid from DealerAffixationCenter where DealerAffixationcenter.dealeraffixationid=a.affix_id and ((Typeofdelivery is null) or (Typeofdelivery='Dealer')) )) like '%" + ShortECname + "%'  and  NewPdfRunningNo is null and  convert(date,slotdate)='" + selectedDate + "'   order by  a.HSRP_StateID";
            DataTable dtSE = Utils.Utils.GetDataTable(stateECQuery, CnnString);
            //string allStrProductionSheetNo = string.Empty;

            if (dtSE.Rows.Count > 0)
            {
                foreach (DataRow dr in dtSE.Rows)
                {
                    #region

                    string HSRP_StateID = dr["HSRP_StateID"].ToString().Trim();
                    string HSRPStateName = dr["HSRPStateName"].ToString().Trim();
                    string HSRPStateShortName = dr["HSRPStateShortName"].ToString().Trim();
                    //string RTOLocationID = dr["RTOLocationID"].ToString().Trim();
                    //string RTOLocationName = dr["RTOLocationName"].ToString().Trim();
                    //string NAVEMBID = dr["NAVEMBID"].ToString().Trim();
                    string Navembcode = dr["Navembcode"].ToString().Trim();
                    string dir = dirPath + DateTime.Now.ToString("yyyy-MM-dd") + "\\" + HSRPStateShortName + "\\";

                    string fileName = string.Empty;
                    string filePath = string.Empty;
                    // string fileName = filePrefix + "-" + Navembcode + ".pdf";



                    StringBuilder html = new StringBuilder();

                    Boolean findRecord = false;
                    string strProductionSheetNo = string.Empty;

                    /*
                    *  Start body & HTMl Tag
                    */
                    #region
                    html.Append(
                        "<!DOCTYPE html>" +
                        "<html>" +
                        "<head>" +
                            "<meta charset='UTF-8'><title>Title</title>" +
                            "<style>" +
                                "@page {" +
                                    /* headers*/
                                    "@top-left {" +
                                        "content: 'Left header';" +
                                    "}" +
                                    "@top-right {" +
                                        "content: 'Right header';" +
                                    "}" +

                                    /* footers */
                                    "@bottom-left {" +
                                        "content: 'Lorem ipsum';" +
                                    "} " +
                                    "@bottom-right {" +
                                        "content: 'Page ' counter(page) ' of ' counter(pages);" +
                                    "}" +
                                    "@bottom-center  {" +
                                        "content:element(footer);" +
                                    "}" +
                                "}" +
                                 "#footer {" +
                                    "position: running(footer);" +
                                "}" +
                                "table {" +
                                  "border-collapse: collapse;" +
                                "}" +

                                "table, th, td {" +
                                    "border: 1px solid black;" +
                                    "text-align: left;" +
                                    "vertical-align: top;" +
                                    "padding:5px;" +
                                "}" +
                            "</style>" +
                        "</head>" +
                        "<body>");
                    #endregion

                    string oemDealerQuery = string.Empty;



                    //string fileAppoinmentDate = string.Empty;
                    // string maxAppointmentdate = "select distinct  Convert(varchar(10),max(SlotBookingDate),105) as MaxAppointmentdate  " +
                    // " from	HSRPrecords H with(nolock) join DealerAffixationCenter d  with(nolock) on H.HSRP_StateID=d.StateID and H.Affix_Id=d.DealerAffixationID  join BookMYHSRPappointment B on H.orderno=B.orderno  " +
                    //" where   Navembid='" + Navembcode + "'  and Navembid not like '%CODO%'     and  NewPdfRunningNo is null and erpassigndate is not null and  h.OrderStatus='New Order' and  IsBookMyHsrpRecord='Y' and  h.affix_id is not NULL  and   d.TypeofDelivery='Home' ";


                    string fileAppoinmentDate = string.Empty;

                    fileAppoinmentDate = selectedDate;
                    //string maxAppointmentdate = "select distinct  Convert(varchar(10),max(SlotDate),105) as MaxAppointmentdate from	HSRPrecordsonlystricker H with(nolock) join DealerAffixationCenter d  with(nolock) on H.HSRP_StateID=d.StateID and H.Affix_Id=d.DealerAffixationID  where BookMyHSRPOrderNo is not  null and   affix_id in (select DealerAffixationid from dealeraffixationcenter where rtolocationid in (select rtolocationid from rtolocation where navembid='" + Navembcode + "')) and  NewPdfRunningNo is null and  h.affix_id is not NULL";
                    //DataTable dtmax = Utils.Utils.GetDataTable(maxAppointmentdate, CnnString);

                    //if (dtmax.Rows.Count > 0)
                    //{
                    //    fileAppoinmentDate = dtmax.Rows[0]["MaxAppointmentdate"].ToString();
                    //}
                    fileName = "BookMyHSRPSticker" + "-" + fileAppoinmentDate + "-" + filePrefix + "-" + Navembcode + ".pdf";
                    filePath = dir + fileName;



                    //oemDealerQuery = "select distinct d.oemid as oemid, (select name  from oemmaster where oemid=d.oemid) as oemname,d.dealerid as dealerid,d.DealerAffixationcenterName as Dealername,  d.DealerAffixationCenterAddress as Address, " +
                    // "d.DealerAffixationID as SubDealerId, H.dealerid AS ParentDealerId,SlotBookingDate from	HSRPrecords H with(nolock) join DealerAffixationCenter d  with(nolock) on H.HSRP_StateID=d.StateID and H.Affix_Id=d.DealerAffixationID  join BookMYHSRPappointment B on H.orderno=B.orderno  " +
                    //" where   Navembid='" + Navembcode + "'  and Navembid not like '%CODO%'     and  NewPdfRunningNo is null and erpassigndate is not null and  h.OrderStatus='New Order' and  IsBookMyHsrpRecord='Y' and  h.affix_id is not NULL   and d.TypeofDelivery='Home' ";

                    oemDealerQuery = "select distinct d.oemid as oemid, (select name  from oemmaster where oemid=d.oemid) as oemname,d.dealerid as dealerid,d.DealerAffixationcenterName as Dealername,d.DealerAffixationCenterAddress as Address,d.DealerAffixationID as SubDealerId, H.dealerid AS ParentDealerId,slotdate as SlotBookingDate from	hsrprecordsonlystricker H with(nolock) join DealerAffixationCenter d  with(nolock) on H.HSRP_StateID=d.StateID and H.Affix_Id=d.DealerAffixationID where   NewPdfRunningNo is null and  isnull(BookMyHsrporderno,'')!='' and  h.affix_id is not NULL   and ((d.TypeofDelivery is null)  or (d.typeofdelivery='Dealer')) and h.created_date>'01-Nov-2020 00:00:00' and affix_id in (select DealerAffixationid from dealeraffixationcenter where rtolocationid in (select rtolocationid from rtolocation where navembid='" + Navembcode + "'))and convert(date,slotdate)='" + selectedDate + "'";


                    #region
                    DataTable dtOD = Utils.Utils.GetDataTable(oemDealerQuery, CnnString);

                    if (dtOD.Rows.Count > 0)
                    {
                        string allStrProductionSheetNo = string.Empty;
                        Session["ECWiseProductionsheet="] = null;
                        foreach (DataRow drOD in dtOD.Rows)
                        {

                            string oemid = drOD["oemid"].ToString().Trim();
                            string dealerid = drOD["dealerid"].ToString().Trim();
                            string oemname = drOD["oemname"].ToString().Trim();
                            string dealername = drOD["Dealername"].ToString().Trim();
                            string Address = drOD["Address"].ToString().Trim();

                            DateTime AppointmentDate = Convert.ToDateTime(drOD["SlotBookingDate"].ToString().Trim());

                            string displayAppointmentDate = AppointmentDate.ToString("dd/MM/yyyy");
                            string strsubaffixid = drOD["SubDealerId"].ToString();
                            string strParentDealerId = drOD["ParentDealerId"].ToString();

                            //start sql query
                            #region
                            string productionQuery = string.Empty;

                            DataTable dtProduction = new DataTable();


                            SqlConnection con = new SqlConnection(CnnString);
                            SqlCommand cmd = new SqlCommand("[USP_BookMYHSRPProductionSheet_OSDealer]", con);
                            cmd.CommandType = CommandType.StoredProcedure;
                            con.Open();

                            cmd.Parameters.AddWithValue("@HSRP_StateID", HSRP_StateID);
                            cmd.Parameters.AddWithValue("@navembid", Navembcode);
                            cmd.Parameters.AddWithValue("@Dealerid", strParentDealerId);
                            cmd.Parameters.AddWithValue("@Affix_Id", strsubaffixid);
                            cmd.Parameters.AddWithValue("@AppointmentDate", AppointmentDate);
                            SqlDataAdapter da = new SqlDataAdapter(cmd);
                            // dtProduction = new DataTable();
                            da.Fill(dtProduction);
                            con.Close();

                            // string fileAppoinmentDate = AppointmentDate.ToString("dd/MM/yyyy");

                            #endregion
                            //end sql query

                            #region
                            //DataTable dtProduction = Utils.Utils.GetDataTable(productionQuery, CnnString);
                            if (dtProduction.Rows.Count > 0)
                            {

                                findRecord = true;
                                string strRunningNo = string.Empty;
                                //string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) as maxSheetNo from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + NAVEMBID + "'";
                                string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) from EmbossingCentersNew where    Emb_Center_Id='" + Navembcode + "' ";
                                string strCom = Utils.Utils.getScalarValue(strSel, CnnString);
                                //DataTable dtSheetNo = fillDataTable(strSel, CnnString);
                                if (strCom.Equals(0) || strCom.Length == 0)
                                {
                                    strRunningNo = "0000001";
                                }
                                else
                                {
                                    strRunningNo = string.Format("{0:0000000}", Convert.ToInt32(strCom) + 1);
                                }

                                string strRequeNo = "select (prefixtext+right('00000'+ convert(varchar,Lastreqno+1),5)) as Reqno from prefix_Requisition  where hsrp_stateid='" + ddlStateName.SelectedValue + "' and prefixfor='Req No'";
                                ReqNum = Utils.Utils.getScalarValue(strRequeNo, CnnString);

                                string strPRFIX = "select PrefixText from EmbossingCentersNew where  Emb_Center_Id='" + Navembcode + "'";
                                string strPRFIXCom = Utils.Utils.getScalarValue(strPRFIX, CnnString);
                                strProductionSheetNo = strPRFIXCom + strRunningNo;

                                //string RTOLocationName = dtProduction.Rows[0]["RTOLocationName"].ToString();




                                #region



                                html.Append("<div style='page-break-before: avoid;page-break-inside: avoid;page-break-after: always;'>" +
                                        "<table style='width:100%;border: 0px;'>" +
                                            "<tr style='border: 0px;'>" +
                                                "<td colspan='7' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>Report Generation Date:</b> " + DateTime.Now.ToString("dd-MM-yyyy") + "</div>" +
                                                "</td>" +
                                                "<td colspan='7' style='border: 0px;'>" +
                                                    "<div style='float:right;width: 500px;word-wrap: break-word;'>" +
                                                        //"<b>Production Sheet No: " + strProductionSheetNo + "<br /> </b>" +
                                                        //"<b>Pin Code:" + dealername + "<br />  </b>" +
                                                        //"<b>Pin Code Location:" + Address + "<br /> </b>" +
                                                        //"<div style='float:right;width: 500px;word-wrap: break-word;'>" +
                                                        "<b>Production Sheet No:</b> " + strProductionSheetNo + "<br />" +
                                                        "<b>Dealer ID/Name:</b> " + dealerid + "/" + dealername + "/" + strsubaffixid + "<br />" +
                                                        "<b>Dealer Address:</b> " + Address + "<br />" +

                                                    // "<b>Appointment Date:-:" + dtProduction.Rows[0]["AppointmentDate"].ToString() + "<br />  </b>" +
                                                    "</div>" +
                                                "</td>" +
                                            "</tr>" +
                                            "<tr style='border: 0px;'>" +
                                                "<td colspan='6' style='border: 0px;'>" +
                                                    "<div style='text-align:center;font-size:24px;'><b>Book My HSRP Sticker Production Sheet : -</b> ROSMERTA SAFETY SYSTEMS LIMITED</div>" +
                                                "</td>" +
                                            "</tr>" +
                                            "<tr style='border: 0px;'>" +
                                                "<td colspan='13' style='border: 0px;'>" +
                                                    "<table style='border:0px;width:100%;'>" +
                                                        "<tr style='border:0px;'>" +
                                                            "<td style='border:0px;'><b>State Name :</b> " + HSRPStateName + "</td>" +

                                                            "<td style='border:0px; font-size:22px;'><b>Report Date: " + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</td> </b>" +

                                                              "<td style='border:0px; font-size:22px;'><b>Appointment Date:-: " + dtProduction.Rows[0]["AppointmentDate"].ToString() + "</td> </b>" +
                                                        "</tr>" +
                                                    "</table>" +
                                                "</td>" +
                                            //            "<td colspan='4' style='border: 0px;'>" +
                                            //                "<div style='float:right'>" +

                                            ////                    "VC:Vehicle Class<br />" +
                                            //////"VT:Vehicle Type<br />" +
                                            ////                    "Front PS:Front Plate Size<br />" +
                                            ////                    "Rear PS:Rear Plate Size<br />" +

                                            //                "</div>" +
                                            //"</td>" +
                                            //"<td style='border: 0px;'></td>" +
                                            "</tr>" +
                                            //"<tr style='border: 0px;'>" +
                                            //     "<td colspan='14' style='border: 0px;'>" +
                                            //        "<div style='text-align:left'>Location Name : " + RTOLocationName + "</div>" +
                                            //    "</td>" +
                                            //"</tr>" +
                                            "<tr>" +
                                                "<td style='text-align:center'>SR.No</td>" +
                                                //"<td>VC</td>" +
                                                "<td>Vehicle No</td>" +
                                                "<td>Fuel Type</td>" +
                                                // "<td>Front PS</td>" +
                                                //"<td>Front Laser No</td>" +
                                                //"<td>Rear PS</td>" +
                                                //"<td>Rear Laser No.</td>" +

                                                "<td style='text-align:center'>Sticker Color</td>" +
                                                 "<td style='text-align:center'>Customer Name</td>" +

                                                   "<td style='width:8%;' style='text-align:center'>Appointment Time</td>" +
                                            "</tr>");
                                #endregion

                                #region
                                string strtvspono = "";

                                int j = 0;
                                int total = 0;
                                StringBuilder UpdateSQL = new StringBuilder();
                                for (int i = 0; i <= dtProduction.Rows.Count - 1; i++)
                                {

                                    j = j + 1;

                                    if (total == 13)
                                    {
                                        total = 0;




                                        html.Append("<div style='page-break-before: avoid;page-break-inside: avoid;page-break-after: always;'>" +
                                        "<table style='width:100%;border: 0px;'>" +
                                            "<tr style='border: 0px;'>" +
                                                "<td colspan='7' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>Report Generation Date:</b> " + DateTime.Now.ToString("dd-MM-yyyy") + "</div>" +
                                                "</td>" +
                                                "<td colspan='7' style='border: 0px;'>" +
                                                    "<div style='float:right;width: 500px;word-wrap: break-word;'>" +
                                                         //"<b>Production Sheet No: " + strProductionSheetNo + "<br /> </b>" +
                                                         //"<b>Pin Code:" + dealername + "<br />  </b>" +
                                                         //"<b>Pin Code Location:" + Address + "<br /> </b>" +
                                                         "<b>Production Sheet No:</b> " + strProductionSheetNo + "<br />" +
                                                        "<b>Dealer ID/Name:</b> " + dealerid + "/" + dealername + "/" + strsubaffixid + "<br />" +
                                                        "<b>Dealer Address:</b> " + Address + "<br />" +

                                                    // "<b>Appointment Date:-:" + dtProduction.Rows[0]["AppointmentDate"].ToString() + "<br />  </b>" +
                                                    "</div>" +
                                                "</td>" +
                                            "</tr>" +
                                            "<tr style='border: 0px;'>" +
                                                "<td colspan='14' style='border: 0px;'>" +
                                                    "<div style='text-align:center;font-size:24px;'><b>Book My HSRP Sticker Production Sheet : -</b> ROSMERTA SAFETY SYSTEMS LIMITED</div>" +
                                                "</td>" +
                                            "</tr>" +
                                            "<tr style='border: 0px;'>" +
                                                "<td colspan='13' style='border: 0px;'>" +
                                                    "<table style='border:0px;width:100%;'>" +
                                                        "<tr style='border:0px;'>" +
                                                            "<td style='border:0px;'><b>State Name :</b> " + HSRPStateName + "</td>" +
                                                              "<td style='border:0px; font-size:22px;'><b>Report Date: " + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</td> </b>" +

                                                                "<td style='border:0px; font-size:22px;'><b>Appointment Date:-: " + dtProduction.Rows[0]["AppointmentDate"].ToString() + "</td> </b>" +
                                                        "</tr>" +
                                                    "</table>" +
                                                "</td>" +
                                                //    "<td colspan='4' style='border: 0px;'>" +
                                                ////        "<div style='float:right'>" +

                                                ////            //"VC:Vehicle Class<br />" +
                                                //////"VT:Vehicle Type<br />" +
                                                ////            //"Front PS:Front Plate Size<br />" +
                                                ////            //"Rear PS:Rear Plate Size<br />" +

                                                ////        "</div>" +
                                                //    "</td>" +
                                                "<td style='border: 0px;'></td>" +
                                            "</tr>" +
                                            //"<tr style='border: 0px;'>" +
                                            //     "<td colspan='14' style='border: 0px;'>" +
                                            //        "<div style='text-align:left'>Location Name : " + RTOLocationName + "</div>" +
                                            //    "</td>" +
                                            //"</tr>" +
                                            "<tr>" +
                                                "<td style='text-align:center'>SR.No</td>" +
                                                //"<td>VC</td>" +
                                                "<td>Vehicle No</td>" +
                                                "<td>Fuel Type</td>" +


                                                "<td style='text-align:center'>Sticker Color</td>" +
                                                 "<td style='text-align:center'>Customer Name</td>" +

                                                   "<td style='width:8%;' style='text-align:center'>Appointment Time</td>" +
                                            "</tr>");


                                    }

                                    total = total + 1;
                                    //foreach (DataRow drProduction in dtProduction.Rows)
                                    //{
                                    string HsrprecordID = dtProduction.Rows[i]["hsrprecordID"].ToString().Trim();
                                    string SRNo = dtProduction.Rows[i]["SRNo"].ToString().Trim();

                                    string VehicleNo = dtProduction.Rows[i]["VehRegNo"].ToString().Trim();

                                    string FuelType = dtProduction.Rows[i]["FuelType"].ToString().Trim();

                                    string StickerColor = dtProduction.Rows[i]["stickerColor"].ToString().Trim();
                                    string CustomerName = dtProduction.Rows[i]["CustomerName"].ToString().Trim();
                                    //string AppointmentDate = drProduction["AppointmentDate"].ToString().Trim();
                                    string AppointmentTime = dtProduction.Rows[i]["AppointmentTime"].ToString().Trim();

                                    html.Append("<tr>" +
                                      "<td style='text-align:center'>" + SRNo + "</td>" +
                                       //"<td>" + VC + "</td>" +
                                       "<td > <b>" + VehicleNo + " </b> </td> " +
                                        "<td>" + FuelType + "</td>" +


                                        //"<td>" + FrontPSize + "</td>" +
                                        //"<td>" + FS1 + "<b>" + FS2 + "</b> </td>" +
                                        //"<td>" + RearPSize + "</td>" +
                                        // //"<td>" + RearLaserNo + "</td>" +
                                        //"<td>" + RS1 + "<b>" + RS2 + "</b> </td>" +

                                        "<td>" + StickerColor + "</td>" +
                                       "<td>" + CustomerName + "</td>" +

                                         "<td>" + AppointmentTime + "</td>" +
                                   "</tr>");

                                    //start updating hsrprecords 
                                    //string sqlUpdateHSRPRecords = "update hsrprecords set sendtoProductionStatus='Y', NAVPDFFlag='1', NewPdfRunningNo='" + strProductionSheetNo + "', Requisitionsheetno='" + ReqNum + "', " +
                                    //       "PdfDownloadDate=GetDate(), pdfFileName='" + fileName + "', PDFDownloadUserID='1' where hsrprecordID='" + HsrprecordID + "' ";
                                    //Utils.Utils.ExecNonQuery(sqlUpdateHSRPRecords, CnnString);   // uncomment after testing

                                    UpdateSQL.Append("update hsrprecordsonlystricker set  NewPdfRunningNo='" + strProductionSheetNo + "',stickerPSDownloaddate=GetDate(), stickerPSPDffilename='" + fileName + "' where ID='" + HsrprecordID + "';");

                                }

                                allStrProductionSheetNo += "," + strProductionSheetNo;

                                allStrProductionSheetNo = allStrProductionSheetNo.TrimStart(',');
                                Session["ECWiseProductionsheet="] = allStrProductionSheetNo;
                                #endregion
                                html.Append("</table>");

                                html.Append("</div>");

                                if (UpdateSQL.ToString().Length > 0)
                                {
                                    Utils.Utils.ExecNonQuery(UpdateSQL.ToString(), CnnString);
                                    string StrSqlUpdateECQuery = "update EmbossingCentersNew set NewProductionSheetRunningNo='" + strProductionSheetNo + "' " +
                                     "where  Emb_Center_Id='" + Navembcode + "'";
                                    Utils.Utils.ExecNonQuery(StrSqlUpdateECQuery, CnnString); // uncomment after testing
                                }

                                //string StrSqlUpdateECQuery = "update EmbossingCentersNew set NewProductionSheetRunningNo='" + strProductionSheetNo + "' " +
                                // "where  Emb_Center_Id='" + Navembcode + "'";
                                //Utils.Utils.ExecNonQuery(StrSqlUpdateECQuery, CnnString); // uncomment after testing
                            }

                            #endregion

                        }
                    }// close oemDealerQuery
                    #endregion
                    //   #region "Req Generate"
                    //   string strSqlQuery1 = "select CompanyName from hsrpstate where hsrp_stateid='" + ddlStateName.SelectedValue + "'";
                    //   strCompanyName = Utils.Utils.getScalarValue(strSqlQuery1, CnnString);

                    //   string strEMBName = " select EmbCenterName from EmbossingCentersNew where   Emb_Center_Id='" + Navembcode + "'";
                    //   string strComNew = Utils.Utils.getScalarValue(strEMBName, CnnString);

                    //   string strReqNo = "select (prefixtext+right('00000'+ convert(varchar,Lastreqno+1),5)) as Reqno from prefix_Requisition  where hsrp_stateid='" + ddlStateName.SelectedValue + "' and prefixfor='Req No'";
                    //   string strReqNumber = Utils.Utils.getScalarValue(strReqNo, CnnString);

                    //   SQLString = "Exec [laserreqSlip1DashBoardBookMYHSRP]  '" + Navembcode + "' ,  '" + ReqNum + "'";
                    //   DataTable dtResult = Utils.Utils.GetDataTable(SQLString, CnnString);
                    //   string strQuery = string.Empty;
                    //   string strRtoLocationName = string.Empty;
                    //   int Itotal = 0;

                    //   html.Append("<div style='width:100%;height:100%;'>" +
                    //                       "<table style='width:100%'>" +

                    //                           "<tr>" +
                    //                               "<td colspan='12'>" +
                    //                                   "<div style='text-align:center;padding:8px;'>" +
                    //                                       "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + strCompanyName + "</b>" +

                    //                                       "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                    //                                        "<b style='font-size:20px;'>" + "MATERIAL REQUSITION NOTE" + "</b>" +

                    //                                   "</div>" +
                    //                               "</td>" +
                    //                           "</tr>" +

                    //                           "<tr>" +
                    //                               "<td colspan='12'>" +
                    //                                   "<div style='text-align:center;padding:8px;'>" +
                    //                                       "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Production Sheet Date :" + DateTime.Now.ToString("dd-MM-yyyy") + "</b>" +
                    //                                       "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                    //                                        "<b style='font-size:20px;'>" + "" + "</b>" +
                    //                                   "</div>" +
                    //                               "</td>" +
                    //                           "</tr>" +
                    //                           "<tr>" +
                    //                               "<td colspan='6'>" +
                    //                                   "<div style='text-align:left;'>" +
                    //                                       "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>REQ.NO:-</b>" +
                    //                                       "" + strReqNumber + "" +
                    //                                   "</div>" +
                    //                               "</td>" +
                    //                               "<td colspan='6'>" +
                    //                                   "<div style='text-align:left;'>" +
                    //                                       "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>Embossing Center:</b>" +
                    //                                       "" + strComNew + "" +
                    //                                   "</div>" +
                    //                               "</td>" +
                    //                           "</tr>" +
                    //                           "<tr>" +
                    //                               "<td colspan='12'>" +
                    //                                   "<div style='text-align:left;padding:2px;'>" +
                    //                                       "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "State:" + ddlStateName.SelectedItem.Text + "</b>" +
                    //                                       "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                    //                                        "<b style='font-size:20px;'>" + "" + "</b>" +
                    //                                   "</div>" +
                    //                               "</td>" +
                    //                           "</tr>" +
                    //                            "<tr>" +
                    //                               "<td colspan='12'>" +
                    //                                   "<div style='text-align:left;padding:8px;'>" +
                    //                                       "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + " " + "</b>" +
                    //                                       "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                    //                                        "<b style='font-size:20px;'>" + "" + "</b>" +
                    //                                   "</div>" +
                    //                               "</td>" +
                    //                           "</tr>" +

                    //                           "<tr>" +
                    //                               "<td colspan='1' style='text-align:center'>SR.N.</td>" +
                    //                               "<td colspan='3'>Product Size</td>" +
                    //                               "<td colspan='1'>Laser Count</td>" +
                    //                               "<td colspan='1'>Start Laser No</td>" +
                    //                               "<td colspan='1'>End Laser No</td>" +

                    //                           "</tr>");


                    //   if (dtResult.Rows.Count > 0)
                    //   {
                    //       for (int i = 0; i < dtResult.Rows.Count; i++)
                    //       {
                    //           string ID = dtResult.Rows[i]["ID"].ToString();
                    //           string productcode = dtResult.Rows[i]["productcode"].ToString();
                    //           string LaserCount = dtResult.Rows[i]["LaserCount"].ToString();
                    //           Itotal = Convert.ToInt32(dtResult.Rows[i]["Total"].ToString());
                    //           string BeginLaser = dtResult.Rows[i]["BeginLaser"].ToString();
                    //           string EndLaser = dtResult.Rows[i]["EndLaser"].ToString();




                    //           html.Append("<tr>" +
                    //              "<td colspan='1' style='text-align:left'>" + ID + "</td>" +
                    //              "<td colspan='3'>" + productcode + "</td>" +

                    //              "<td colspan='1'>" + LaserCount + "</td>" +
                    //              "<td colspan='1'>" + BeginLaser + "</td>" +
                    //              "<td colspan='1'>" + EndLaser + "</td>" +

                    //          "</tr>");
                    //       }

                    //       try
                    //       {
                    //           //start updating hsrprecords 
                    //           string Query = "update prefix_Requisition set lastno=lastno+1,Lastreqno=Lastreqno+1 where hsrp_stateid='" + ddlStateName.SelectedValue + "' and prefixfor='Req No'";
                    //           Utils.Utils.ExecNonQuery(Query, CnnString);
                    //       }
                    //       catch (Exception ev)
                    //       {
                    //           Label1.Text = "prefix Requisition update error: " + ev.Message;
                    //       }
                    //   }
                    //   html.Append("<tr>" +
                    //            "<td colspan='1' style='text-align:center' > " + "<b>Grand Total:</b>" + "</td>" +
                    //            "<td colspan='3'>" + "" + "</td>" +

                    //            "<td colspan='1'>" + Itotal + "</td>" +
                    //            "<td colspan='1'>" + " " + "</td>" +
                    //            "<td colspan='1'>" + " " + "</td>" +

                    //        "</tr>");




                    //   html.Append("<tr>" +
                    //    "<td colspan='2' > " + "<b>REQUESTED BY </b>" + "</td>" +
                    //    "<td colspan='2'>" + "<b>AUTHORIZED BY </b>" + "</td>" +

                    //    "<td colspan='2'>" + "<b>ISSUED BY </b>" + "</td>" +
                    //    "<td colspan='2'>" + "<b>RECEIVED BY</b>" + "</td>" +


                    //"</tr>");

                    //   html.Append("<tr>" +
                    //                             "<td colspan='12'>" +
                    //                                 "<div style='text-align:left;padding:2px;'>" +
                    //                                     "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Name" + "</b>" +

                    //                                 "</div>" +
                    //                             "</td>" +
                    //                         "</tr>");

                    //   html.Append("<tr>" +
                    //                            "<td colspan='12'>" +
                    //                                "<div style='text-align:left;padding:2px;'>" +
                    //                                    "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Designation" + "</b>" +

                    //                                "</div>" +
                    //                            "</td>" +
                    //                        "</tr>");

                    //   html.Append("<tr>" +
                    //                          "<td colspan='12'>" +
                    //                              "<div style='text-align:right;padding:8px;'>" +
                    //                                  "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Download By: Admin " + "</b>" +

                    //                              "</div>" +
                    //                          "</td>" +
                    //                      "</tr>");

                    //   html.Append("<tr>" +
                    //                         "<td colspan='12'>" +
                    //                             "<div style='text-align:right;padding:8px;'>" +
                    //                                 "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Sheet Generated By :" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</b>" +

                    //                             "</div>" +
                    //                         "</td>" +
                    //                     "</tr>");


                    //   html.Append("<tr style='visibility: hidden;'>" + "<td  ><p style=\"page-break-after:always\"/></td></tr>");



                    //   //html.Append("</table>");

                    //   //html.Append("</div>");

                    //   #region "PS Summary Report"
                    //   string PS = string.Empty;
                    //   if (Session["ECWiseProductionsheet="] != null)
                    //   {
                    //       PS = Session["ECWiseProductionsheet="].ToString();
                    //   }
                    //   /*
                    //    * Close body & HTMl Tag
                    //    */

                    //   SQLString = "Exec [USP_BMPSSummaryHomedelivery]  '" + PS + "' ";
                    //   DataTable dtResultsummary = Utils.Utils.GetDataTable(SQLString, CnnString);

                    //   html.Append("<div style='width:100%;height:100%;'>" +
                    //                "<table style='width:100%'>" +

                    //                    "<tr>" +
                    //                        "<td colspan='12'>" +
                    //                            "<div style='text-align:center;padding:8px;'>" +
                    //                                "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Home Delivery Summary Report" + "</b>" +

                    //                            "</div>" +
                    //                        "</td>" +
                    //                    "</tr>" +

                    //                     "<tr>" +
                    //                               "<td colspan='1' style='text-align:center'>PS NO</td>" +
                    //                               "<td colspan='3'>Affixation Pin Code</td>" +
                    //                               "<td colspan='1'>2W</td>" +
                    //                               "<td colspan='1'>3W</td>" +
                    //                               "<td colspan='1'>4W</td>" +
                    //                               "<td colspan='1'>OTH</td>" +
                    //                                 "<td colspan='1'>Total</td>" +

                    //                           "</tr>");

                    //   if (dtResultsummary.Rows.Count > 0)
                    //   {
                    //       for (int i = 0; i < dtResultsummary.Rows.Count; i++)
                    //       {
                    //           string PSNo = dtResultsummary.Rows[i]["PSNo"].ToString();
                    //           string AffixationdealerName = dtResultsummary.Rows[i]["Affixation Pin Code"].ToString();
                    //           string W2 = dtResultsummary.Rows[i]["2W"].ToString();
                    //           string W3 = dtResultsummary.Rows[i]["3W"].ToString();
                    //           string W4 = dtResultsummary.Rows[i]["4W"].ToString();
                    //           string OTH = dtResultsummary.Rows[i]["OTH"].ToString();
                    //           string Total = dtResultsummary.Rows[i]["Total"].ToString();




                    //           html.Append("<tr>" +
                    //              "<td colspan='1' style='text-align:left'>" + PSNo + "</td>" +
                    //              "<td colspan='3'>" + AffixationdealerName + "</td>" +

                    //              "<td colspan='1'>" + W2 + "</td>" +
                    //              "<td colspan='1'>" + W3 + "</td>" +
                    //              "<td colspan='1'>" + W4 + "</td>" +
                    //               "<td colspan='1'>" + OTH + "</td>" +
                    //                "<td colspan='1'>" + Total + "</td>" +

                    //          "</tr>");
                    //       }
                    //   }

                    //   html.Append("</table>");

                    //   html.Append("</div>");

                    //   #endregion



                    //   #endregion

                    /*
                     * Close body & HTMl Tag
                     */


                    if (findRecord)
                    {






                        string strSaveSQlQuery = "insert into ProductionSheetAutoGenerated_List (HSRP_StateID, State_Code, RTO_ID, Emb_Center_Id, FileName, Productiondate, ProductionTime, OrderNo) values " +
                                    "('" + ddlStateName.SelectedValue + "', '" + HSRPStateShortName + "', '', '" + Navembcode + "', '" + fileName + "', '" + DateTime.Now.ToString("yyyy-MM-dd") + "', '" + DateTime.Now.ToString("HH:mm:ss") + "', '" + filePrefix + "') ";
                        Utils.Utils.ExecNonQuery(strSaveSQlQuery, CnnString); // uncomment after testing
                        Label1.Text = Label1.Text + "successfully created Book My HSRP Sticker production sheet of State: " + HSRPStateShortName + ", , Emb_Center_Id: " + Navembcode + Environment.NewLine;

                        //lblLog.Text = lblLog.Text + Environment.NewLine + "Start Production Query at:" + DateTime.Now;



                        #region
                        try
                        {
                            if (!Directory.Exists(dir))
                            {


                                Directory.CreateDirectory(dir);

                            }

                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        try
                        {
                            var pdf = Pdf
                                    .From(html.ToString())
                                    .OfSize(PaperSize.A4)
                                    .WithTitle("Title")
                                    .WithoutOutline()
                                    .WithMargins(1.25.Centimeters())
                                    .Landscape()
                                    .Comressed()
                                    .Content();

                            FileStream readStream = File.Create(filePath);
                            BinaryWriter binaryWriter = new BinaryWriter(readStream);

                            // Write the binary data to the file
                            binaryWriter.Write(pdf);
                            binaryWriter.Close();
                            readStream.Close();
                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        #endregion
                    }

                    #endregion

                }//close foreach stateEcQuery
            }
        }

        protected void btnThirdSticker_Click(object sender, EventArgs e)
        {
            if (ddlStateName.SelectedValue == "0")
            {
                lblErrMess.Text = String.Empty;
                lblErrMess.Text = "Please select State";
                return;
            }


            if (txtDate.Text == string.Empty)
            {
                lblErrMess.Text = String.Empty;
                lblErrMess.Text = "Please Select Sticker Order Date";
                return;
            }
            string sqlPrefixQuery = "select top 1 CONVERT(INT, isnull(OrderNo,'0')) + 1 as orderNo from ProductionSheetAutoGenerated_List order by id desc";
            DataTable dtPrefix = Utils.Utils.GetDataTable(sqlPrefixQuery, CnnString);

            if (dtPrefix.Rows.Count > 0)
            {
                filePrefix = dtPrefix.Rows[0]["orderNo"].ToString();
            }
            else
            {
                filePrefix = "1";
            }

            //txtFilePrefix.Text = filePrefix;

            if (filePrefix.Length > 0)
            {
                SheetGenerationThirdSticker();
            }
        }



        private void SheetGenerationThirdSticker()
        {
            string stateECQuery = string.Empty;
            string ReqNum = string.Empty;

            string[] strArray = null;
            string Appdate = txtDate.Text;
            strArray = Appdate.Split('/');

            string selectedDate = string.Empty;

            string Year, month, day = string.Empty;
            day = strArray[0];
            month = strArray[1];
            Year = strArray[2];
            selectedDate = Year + "-" + month + "-" + day;

            string stateECShortName = "select distinct HSRP_StateID, 'EC'+HSRPStateShortName as NewHSRPStateShortName, HSRPStateShortName   from  HSRPState  where  HSRP_STateId='" + ddlStateName.SelectedValue + "'";

            DataTable dtECName = Utils.Utils.GetDataTable(stateECShortName, CnnString);

            string ShortECname = dtECName.Rows[0]["NewHSRPStateShortName"].ToString();
            // string HSRPStateShortName = dtECName.Rows[0]["HSRPStateShortName"].ToString();

            //stateECQuery = "select distinct a.HSRP_StateID, (select HSRPStateName from hsrpstate c where c.hsrp_stateid=a.hsrp_stateid) as HSRPStateName,(select HSRPStateShortName from hsrpstate d with(nolock) where d.hsrp_stateid='" + ddlStateName.SelectedValue + "') as HSRPStateShortName,   a.Navembid AS  navembcode from hsrprecords a where  left(a.Navembid,4) = '" + ShortECname + "'  and  NewPdfRunningNo is null and erpassigndate is not null  and (IsBookMyHsrpRecord='Y')  and OrderStatus='New Order' and    Navembid not like '%CODO%'  order by  a.HSRP_StateID";
            //stateECQuery = "select distinct a.HSRP_StateID, (select HSRPStateName from hsrpstate c where c.hsrp_stateid=a.hsrp_stateid) as HSRPStateName,(select HSRPStateShortName from hsrpstate d with(nolock) where d.hsrp_stateid='" + ddlStateName.SelectedValue + "') as HSRPStateShortName,   (select distinct navembid from rtolocation where RTOLocationID in (select rtolocationid from hsrprecordsonlystricker )) AS  navembcode from hsrprecordsonlystricker a where  BookMyHSRPOrderNo is   null   and  NewPdfRunningNo is null and  convert(date,created_date)='" + selectedDate + "'   order by  a.HSRP_StateID";

            stateECQuery = "select distinct a.HSRP_StateID, (select HSRPStateName from hsrpstate c where c.hsrp_stateid=a.hsrp_stateid) as HSRPStateName,(select HSRPStateShortName from hsrpstate d with(nolock) where d.hsrp_stateid='" + ddlStateName.SelectedValue + "') as HSRPStateShortName,R.navembid as  navembcode from hsrprecordsonlystricker a,rtolocation  r  where     Navembid not like '%CODO%' and a.hsrp_stateid='" + ddlStateName.SelectedValue + "' and   a.rtolocationid =r.rtolocationid and   BookMyHSRPOrderNo is   null   and  NewPdfRunningNo is null  and  convert(date,created_date)='" + selectedDate + "'   order by  a.HSRP_StateID";

            DataTable dtSE = Utils.Utils.GetDataTable(stateECQuery, CnnString);
            //string allStrProductionSheetNo = string.Empty;

            if (dtSE.Rows.Count > 0)
            {
                foreach (DataRow dr in dtSE.Rows)
                {
                    #region

                    string HSRP_StateID = dr["HSRP_StateID"].ToString().Trim();
                    string HSRPStateName = dr["HSRPStateName"].ToString().Trim();
                    string HSRPStateShortName = dr["HSRPStateShortName"].ToString().Trim();
                    //string RTOLocationID = dr["RTOLocationID"].ToString().Trim();
                    //string RTOLocationName = dr["RTOLocationName"].ToString().Trim();
                    //string NAVEMBID = dr["NAVEMBID"].ToString().Trim();
                    string Navembcode = dr["Navembcode"].ToString().Trim();
                    string dir = dirPath + DateTime.Now.ToString("yyyy-MM-dd") + "\\" + HSRPStateShortName + "\\";

                    string fileName = string.Empty;
                    string filePath = string.Empty;
                    // string fileName = filePrefix + "-" + Navembcode + ".pdf";



                    StringBuilder html = new StringBuilder();

                    Boolean findRecord = false;
                    string strProductionSheetNo = string.Empty;

                    /*
                    *  Start body & HTMl Tag
                    */
                    #region
                    html.Append(
                        "<!DOCTYPE html>" +
                        "<html>" +
                        "<head>" +
                            "<meta charset='UTF-8'><title>Title</title>" +
                            "<style>" +
                                "@page {" +
                                    /* headers*/
                                    "@top-left {" +
                                        "content: 'Left header';" +
                                    "}" +
                                    "@top-right {" +
                                        "content: 'Right header';" +
                                    "}" +

                                    /* footers */
                                    "@bottom-left {" +
                                        "content: 'Lorem ipsum';" +
                                    "} " +
                                    "@bottom-right {" +
                                        "content: 'Page ' counter(page) ' of ' counter(pages);" +
                                    "}" +
                                    "@bottom-center  {" +
                                        "content:element(footer);" +
                                    "}" +
                                "}" +
                                 "#footer {" +
                                    "position: running(footer);" +
                                "}" +
                                "table {" +
                                  "border-collapse: collapse;" +
                                "}" +

                                "table, th, td {" +
                                    "border: 1px solid black;" +
                                    "text-align: left;" +
                                    "vertical-align: top;" +
                                    "padding:5px;" +
                                "}" +
                            "</style>" +
                        "</head>" +
                        "<body>");
                    #endregion

                    string oemDealerQuery = string.Empty;






                    string fileAppoinmentDate = string.Empty;

                    fileAppoinmentDate = selectedDate;

                    fileName = "BookMyHSRPThirdSticker" + "-" + fileAppoinmentDate + "-" + filePrefix + "-" + Navembcode + ".pdf";
                    filePath = dir + fileName;


                    oemDealerQuery = "select distinct '' as oemid,''as oemname, H.dealerid as dealerid,(select Dealername from dealermaster where dealerid=h.dealerid) as DealerName,(select Address from dealermaster where dealerid=h.dealerid)  as Address,Created_Date as SlotBookingDate from	hsrprecordsonlystricker H with(nolock)   where   NewPdfRunningNo is null and  BookMyHsrporderno is null   and h.created_date>'01-Dec-2020 00:00:00' and H.Rtolocationid in(select rtolocationid  from  rtolocation where navembid='" + Navembcode + "')and convert(date,Created_Date)='" + selectedDate + "'";


                    #region
                    DataTable dtOD = Utils.Utils.GetDataTable(oemDealerQuery, CnnString);

                    if (dtOD.Rows.Count > 0)
                    {
                        string allStrProductionSheetNo = string.Empty;
                        Session["ECWiseProductionsheet="] = null;
                        foreach (DataRow drOD in dtOD.Rows)
                        {

                            string oemid = drOD["oemid"].ToString().Trim();
                            string dealerid = drOD["dealerid"].ToString().Trim();
                            string oemname = drOD["oemname"].ToString().Trim();
                            string dealername = drOD["Dealername"].ToString().Trim();
                            string Address = drOD["Address"].ToString().Trim();

                            DateTime AppointmentDate = Convert.ToDateTime(drOD["SlotBookingDate"].ToString().Trim());

                            string displayAppointmentDate = AppointmentDate.ToString("dd/MM/yyyy");
                            //string strsubaffixid = drOD["SubDealerId"].ToString();
                            string strParentDealerId = drOD["dealerid"].ToString();

                            //start sql query
                            #region
                            string productionQuery = string.Empty;

                            DataTable dtProduction = new DataTable();


                            SqlConnection con = new SqlConnection(CnnString);
                            SqlCommand cmd = new SqlCommand("[USP_BookMYHSRPProductionSheet_OTSDealer]", con);
                            cmd.CommandType = CommandType.StoredProcedure;
                            con.Open();

                            cmd.Parameters.AddWithValue("@HSRP_StateID", HSRP_StateID);
                            cmd.Parameters.AddWithValue("@navembid", Navembcode);
                            cmd.Parameters.AddWithValue("@Dealerid", strParentDealerId);
                            //cmd.Parameters.AddWithValue("@Affix_Id", strsubaffixid);
                            cmd.Parameters.AddWithValue("@AppointmentDate", AppointmentDate);
                            SqlDataAdapter da = new SqlDataAdapter(cmd);
                            // dtProduction = new DataTable();
                            da.Fill(dtProduction);
                            con.Close();

                            // string fileAppoinmentDate = AppointmentDate.ToString("dd/MM/yyyy");

                            #endregion
                            //end sql query

                            #region
                            //DataTable dtProduction = Utils.Utils.GetDataTable(productionQuery, CnnString);
                            if (dtProduction.Rows.Count > 0)
                            {

                                findRecord = true;
                                string strRunningNo = string.Empty;
                                //string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) as maxSheetNo from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + NAVEMBID + "'";
                                string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) from EmbossingCentersNew where    Emb_Center_Id='" + Navembcode + "' ";
                                string strCom = Utils.Utils.getScalarValue(strSel, CnnString);
                                //DataTable dtSheetNo = fillDataTable(strSel, CnnString);
                                if (strCom.Equals(0) || strCom.Length == 0)
                                {
                                    strRunningNo = "0000001";
                                }
                                else
                                {
                                    strRunningNo = string.Format("{0:0000000}", Convert.ToInt32(strCom) + 1);
                                }

                                string strRequeNo = "select (prefixtext+right('00000'+ convert(varchar,Lastreqno+1),5)) as Reqno from prefix_Requisition  where hsrp_stateid='" + ddlStateName.SelectedValue + "' and prefixfor='Req No'";
                                ReqNum = Utils.Utils.getScalarValue(strRequeNo, CnnString);

                                string strPRFIX = "select PrefixText from EmbossingCentersNew where  Emb_Center_Id='" + Navembcode + "'";
                                string strPRFIXCom = Utils.Utils.getScalarValue(strPRFIX, CnnString);
                                strProductionSheetNo = strPRFIXCom + strRunningNo;

                                //string RTOLocationName = dtProduction.Rows[0]["RTOLocationName"].ToString();




                                #region



                                html.Append("<div style='page-break-before: avoid;page-break-inside: avoid;page-break-after: always;'>" +
                                        "<table style='width:100%;border: 0px;'>" +
                                            "<tr style='border: 0px;'>" +
                                                "<td colspan='7' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>Report Generation Date:</b> " + DateTime.Now.ToString("dd-MM-yyyy") + "</div>" +
                                                "</td>" +
                                                "<td colspan='7' style='border: 0px;'>" +
                                                    "<div style='float:right;width: 500px;word-wrap: break-word;'>" +
                                                        //"<b>Production Sheet No: " + strProductionSheetNo + "<br /> </b>" +
                                                        //"<b>Pin Code:" + dealername + "<br />  </b>" +
                                                        //"<b>Pin Code Location:" + Address + "<br /> </b>" +
                                                        //"<div style='float:right;width: 500px;word-wrap: break-word;'>" +
                                                        "<b>Production Sheet No:</b> " + strProductionSheetNo + "<br />" +
                                                        "<b>Dealer ID/Name:</b> " + dealerid + "/" + dealername + "<br />" +
                                                        "<b>Dealer Address:</b> " + Address + "<br />" +

                                                    // "<b>Appointment Date:-:" + dtProduction.Rows[0]["AppointmentDate"].ToString() + "<br />  </b>" +
                                                    "</div>" +
                                                "</td>" +
                                            "</tr>" +
                                            "<tr style='border: 0px;'>" +
                                                "<td colspan='6' style='border: 0px;'>" +
                                                    "<div style='text-align:center;font-size:24px;'><b>Third Sticker Production Sheet : -</b> ROSMERTA SAFETY SYSTEMS LIMITED</div>" +
                                                "</td>" +
                                            "</tr>" +
                                            "<tr style='border: 0px;'>" +
                                                "<td colspan='13' style='border: 0px;'>" +
                                                    "<table style='border:0px;width:100%;'>" +
                                                        "<tr style='border:0px;'>" +
                                                            "<td style='border:0px;'><b>State Name :</b> " + HSRPStateName + "</td>" +

                                                            "<td style='border:0px; font-size:22px;'><b>Report Date: " + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</td> </b>" +

                                                              "<td style='border:0px; font-size:22px;'><b>Order Date: " + dtProduction.Rows[0]["AppointmentDate"].ToString() + "</td> </b>" +
                                                        "</tr>" +
                                                    "</table>" +
                                                "</td>" +

                                            "<tr>" +
                                                "<td style='text-align:center'>SR.No</td>" +
                                                //"<td>VC</td>" +
                                                "<td>Vehicle No</td>" +
                                                "<td>Fuel Type</td>" +
                                                // "<td>Front PS</td>" +
                                                //"<td>Front Laser No</td>" +
                                                //"<td>Rear PS</td>" +
                                                //"<td>Rear Laser No.</td>" +

                                                "<td style='text-align:center'>Sticker Color</td>" +
                                                 "<td style='text-align:center'>Customer Name</td>" +

                                            //"<td style='width:8%;' style='text-align:center'>Appointment Time</td>" +
                                            "</tr>");
                                #endregion

                                #region
                                string strtvspono = "";

                                int j = 0;
                                int total = 0;
                                StringBuilder UpdateSQL = new StringBuilder();
                                for (int i = 0; i <= dtProduction.Rows.Count - 1; i++)
                                {

                                    j = j + 1;

                                    if (total == 13)
                                    {
                                        total = 0;




                                        html.Append("<div style='page-break-before: avoid;page-break-inside: avoid;page-break-after: always;'>" +
                                        "<table style='width:100%;border: 0px;'>" +
                                            "<tr style='border: 0px;'>" +
                                                "<td colspan='7' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>Report Generation Date:</b> " + DateTime.Now.ToString("dd-MM-yyyy") + "</div>" +
                                                "</td>" +
                                                "<td colspan='7' style='border: 0px;'>" +
                                                    "<div style='float:right;width: 500px;word-wrap: break-word;'>" +
                                                         //"<b>Production Sheet No: " + strProductionSheetNo + "<br /> </b>" +
                                                         //"<b>Pin Code:" + dealername + "<br />  </b>" +
                                                         //"<b>Pin Code Location:" + Address + "<br /> </b>" +
                                                         "<b>Production Sheet No:</b> " + strProductionSheetNo + "<br />" +
                                                        "<b>Dealer ID/Name:</b> " + dealerid + "/" + dealername + "/" + "<br />" +
                                                        "<b>Dealer Address:</b> " + Address + "<br />" +

                                                    // "<b>Appointment Date:-:" + dtProduction.Rows[0]["AppointmentDate"].ToString() + "<br />  </b>" +
                                                    "</div>" +
                                                "</td>" +
                                            "</tr>" +
                                            "<tr style='border: 0px;'>" +
                                                "<td colspan='14' style='border: 0px;'>" +
                                                    "<div style='text-align:center;font-size:24px;'><bThird Sticker Production Sheet : -</b> ROSMERTA SAFETY SYSTEMS LIMITED</div>" +
                                                "</td>" +
                                            "</tr>" +
                                            "<tr style='border: 0px;'>" +
                                                "<td colspan='13' style='border: 0px;'>" +
                                                    "<table style='border:0px;width:100%;'>" +
                                                        "<tr style='border:0px;'>" +
                                                            "<td style='border:0px;'><b>State Name :</b> " + HSRPStateName + "</td>" +
                                                              "<td style='border:0px; font-size:22px;'><b>Report Date: " + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</td> </b>" +

                                                                "<td style='border:0px; font-size:22px;'><b>Order Date: " + dtProduction.Rows[0]["AppointmentDate"].ToString() + "</td> </b>" +
                                                        "</tr>" +
                                                    "</table>" +
                                                "</td>" +

                                                "<td style='border: 0px;'></td>" +
                                            "</tr>" +

                                            "<tr>" +
                                                "<td style='text-align:center'>SR.No</td>" +
                                                //"<td>VC</td>" +
                                                "<td>Vehicle No</td>" +
                                                "<td>Fuel Type</td>" +


                                                "<td style='text-align:center'>Sticker Color</td>" +
                                                 "<td style='text-align:center'>Customer Name</td>" +

                                            //"<td style='width:8%;' style='text-align:center'>Appointment Time</td>" +
                                            "</tr>");


                                    }

                                    total = total + 1;

                                    string HsrprecordID = dtProduction.Rows[i]["hsrprecordID"].ToString().Trim();
                                    string SRNo = dtProduction.Rows[i]["SRNo"].ToString().Trim();

                                    string VehicleNo = dtProduction.Rows[i]["VehRegNo"].ToString().Trim();

                                    string FuelType = dtProduction.Rows[i]["FuelType"].ToString().Trim();

                                    string StickerColor = dtProduction.Rows[i]["stickerColor"].ToString().Trim();
                                    string CustomerName = dtProduction.Rows[i]["CustomerName"].ToString().Trim();

                                    // string AppointmentTime = dtProduction.Rows[i]["AppointmentTime"].ToString().Trim();

                                    html.Append("<tr>" +
                                      "<td style='text-align:center'>" + SRNo + "</td>" +
                                       //"<td>" + VC + "</td>" +
                                       "<td > <b>" + VehicleNo + " </b> </td> " +
                                        "<td>" + FuelType + "</td>" +




                                        "<td>" + StickerColor + "</td>" +
                                       "<td>" + CustomerName + "</td>" +

                                   //"<td>" + AppointmentTime + "</td>" +
                                   "</tr>");



                                    UpdateSQL.Append("update hsrprecordsonlystricker set  NewPdfRunningNo='" + strProductionSheetNo + "',stickerPSDownloaddate=GetDate(), stickerPSPDffilename='" + fileName + "' where ID='" + HsrprecordID + "';");

                                }

                                allStrProductionSheetNo += "," + strProductionSheetNo;

                                allStrProductionSheetNo = allStrProductionSheetNo.TrimStart(',');
                                Session["ECWiseProductionsheet="] = allStrProductionSheetNo;
                                #endregion
                                html.Append("</table>");

                                html.Append("</div>");

                                if (UpdateSQL.ToString().Length > 0)
                                {
                                    Utils.Utils.ExecNonQuery(UpdateSQL.ToString(), CnnString);
                                    string StrSqlUpdateECQuery = "update EmbossingCentersNew set NewProductionSheetRunningNo='" + strProductionSheetNo + "' " +
                                     "where  Emb_Center_Id='" + Navembcode + "'";
                                    Utils.Utils.ExecNonQuery(StrSqlUpdateECQuery, CnnString); // uncomment after testing
                                }

                                //string StrSqlUpdateECQuery = "update EmbossingCentersNew set NewProductionSheetRunningNo='" + strProductionSheetNo + "' " +
                                // "where  Emb_Center_Id='" + Navembcode + "'";
                                //Utils.Utils.ExecNonQuery(StrSqlUpdateECQuery, CnnString); // uncomment after testing
                            }

                            #endregion

                        }
                    }// close oemDealerQuery
                    #endregion


                    if (findRecord)
                    {






                        string strSaveSQlQuery = "insert into ProductionSheetAutoGenerated_List (HSRP_StateID, State_Code, RTO_ID, Emb_Center_Id, FileName, Productiondate, ProductionTime, OrderNo) values " +
                                    "('" + ddlStateName.SelectedValue + "', '" + HSRPStateShortName + "', '', 'ECDL001', '" + fileName + "', '" + DateTime.Now.ToString("yyyy-MM-dd") + "', '" + DateTime.Now.ToString("HH:mm:ss") + "', '" + filePrefix + "') ";
                        Utils.Utils.ExecNonQuery(strSaveSQlQuery, CnnString); // uncomment after testing
                        Label1.Text = Label1.Text + "successfully created Book My HSRP Third Sticker production sheet of State: " + HSRPStateShortName + ", , Emb_Center_Id: " + Navembcode + Environment.NewLine;

                        //lblLog.Text = lblLog.Text + Environment.NewLine + "Start Production Query at:" + DateTime.Now;



                        #region
                        try
                        {
                            if (!Directory.Exists(dir))
                            {


                                Directory.CreateDirectory(dir);

                            }

                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        try
                        {
                            var pdf = Pdf
                                    .From(html.ToString())
                                    .OfSize(PaperSize.A4)
                                    .WithTitle("Title")
                                    .WithoutOutline()
                                    .WithMargins(1.25.Centimeters())
                                    .Landscape()
                                    .Comressed()
                                    .Content();

                            FileStream readStream = File.Create(filePath);
                            BinaryWriter binaryWriter = new BinaryWriter(readStream);

                            // Write the binary data to the file
                            binaryWriter.Write(pdf);
                            binaryWriter.Close();
                            readStream.Close();
                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        #endregion
                    }

                    #endregion

                }//close foreach stateEcQuery
            }
        }

        protected void btnDeliveryWFrame_Click(object sender, EventArgs e)
        {
            if (ddlStateName.SelectedItem.Text == "--Select State--")
            {
                lblErrMess.Text = "Please select  State";
                return;
            }

            string sqlPrefixQuery = "select top 1 CONVERT(INT, isnull(OrderNo,'0')) + 1 as orderNo from ProductionSheetAutoGenerated_List order by id desc";
            DataTable dtPrefix = Utils.Utils.GetDataTable(sqlPrefixQuery, CnnString);

            if (dtPrefix.Rows.Count > 0)
            {
                filePrefix = dtPrefix.Rows[0]["orderNo"].ToString();
            }
            else
            {
                filePrefix = "1";
            }

            //txtFilePrefix.Text = filePrefix;

            if (filePrefix.Length > 0)
            {
                SheetGenerationBookMyHSRPHDWithOutFrames();
            }
        }

        private void SheetGenerationBookMyHSRPHDWithOutFrames()
        {
            string stateECQuery = string.Empty;
            string ReqNum = string.Empty;

            //string Navembid = Session["Navembid"].ToString();
            FillUserDetails();

            string stateECShortName = "select distinct HSRP_StateID, 'EC'+HSRPStateShortName as NewHSRPStateShortName, HSRPStateShortName   from  HSRPState  where  HSRP_STateId='" + ddlStateName.SelectedValue + "'";

            DataTable dtECName = Utils.Utils.GetDataTable(stateECShortName, CnnString);

            string ShortECname = dtECName.Rows[0]["NewHSRPStateShortName"].ToString();
            // string HSRPStateShortName = dtECName.Rows[0]["HSRPStateShortName"].ToString();

            //stateECQuery = "select distinct a.HSRP_StateID, (select HSRPStateName from hsrpstate c where c.hsrp_stateid=a.hsrp_stateid) as HSRPStateName,(select HSRPStateShortName from hsrpstate d with(nolock) where d.hsrp_stateid='" + ddlStateName.SelectedValue + "') as HSRPStateShortName,   a.Navembid AS  navembcode from hsrprecords a with(nolock) where  left(a.Navembid,4) = '" + ShortECname + "'  and  a.NAVEMBID='" + Navembid + "' and  NewPdfRunningNo is null and erpassigndate is not null  and (IsBookMyHsrpRecord='Y')  and OrderStatus='New Order' and   a.NAVEMBID='" + Navembid + "' and   Navembid not like '%CODO%'  order by  a.HSRP_StateID";
            stateECQuery = "select distinct a.HSRP_StateID, (select HSRPStateName from hsrpstate c where c.hsrp_stateid=a.hsrp_stateid) as HSRPStateName,(select HSRPStateShortName from hsrpstate d with(nolock) where d.hsrp_stateid='" + ddlStateName.SelectedValue + "') as HSRPStateShortName,   a.Navembid AS  navembcode from hsrprecords a with(nolock) where    a.NAVEMBID='" + Navembid + "' and  NewPdfRunningNo is null and erpassigndate is not null  and (IsBookMyHsrpRecord='Y')  and OrderStatus='New Order' and   a.NAVEMBID='" + Navembid + "'   order by  a.HSRP_StateID";

            DataTable dtSE = Utils.Utils.GetDataTable(stateECQuery, CnnString);
            //string allStrProductionSheetNo = string.Empty;

            if (dtSE.Rows.Count > 0)
            {
                foreach (DataRow dr in dtSE.Rows)
                {
                    #region

                    string HSRP_StateID = dr["HSRP_StateID"].ToString().Trim();
                    string HSRPStateName = dr["HSRPStateName"].ToString().Trim();
                    string HSRPStateShortName = dr["HSRPStateShortName"].ToString().Trim();
                   
                    string Navembcode = dr["Navembcode"].ToString().Trim();
                    string dir = dirPath + DateTime.Now.ToString("yyyy-MM-dd") + "\\" + HSRPStateShortName + "\\";

                    string fileName = string.Empty;
                    string filePath = string.Empty;
                    // string fileName = filePrefix + "-" + Navembcode + ".pdf";



                    StringBuilder html = new StringBuilder();

                    Boolean findRecord = false;
                    string strProductionSheetNo = string.Empty;

                    /*
                    *  Start body & HTMl Tag
                    */
                    #region
                    html.Append(
                      "<!DOCTYPE html>" +
                      "<html>" +
                      "<head>" +
                          "<meta charset='UTF-8'><title>Title</title>" +
                          "<style>" +
                              "@page {" +
                                  /* headers*/
                                  "@top-left {" +
                                      "content: 'Left header';" +
                                  "}" +
                                  "@top-right {" +
                                      "content: 'Right header';" +
                                  "}" +

                                  /* footers */
                                  "@bottom-left {" +
                                      "content: 'Lorem ipsum';" +
                                  "} " +
                                  "@bottom-right {" +
                                      "content: 'Page ' counter(page) ' of ' counter(pages);" +
                                  "}" +
                                  "@bottom-center  {" +
                                      "content:element(footer);" +
                                  "}" +
                              "}" +
                              //"#main-table td:nth-child(1){ width:5%; } #main-table td:nth-child(8),#main-table td:nth-child(9),#main-table td:nth-child(7){ width:6%; } #main-table td:nth-child(10){ width:8%; }" +
                              //  "#main-table td:nth-child(3),#main-table td:nth-child(5){ width:12%;white-space: nowrap; }" +
                              "#footer {" +
                                  "position: running(footer);" +
                              "}" +
                              "table {" +
                                "border-collapse: collapse;" +
                              "}" +
                              "table, th, td {" +
                                  "border: 1px solid black;" +
                                  "text-align: left;" +
                                  "vertical-align: top;" +
                                  "padding-left:10px;" +
                                  "padding-bottom:6px;" +
                                  "padding-right:10px;" +
                                  "padding-top:5px;" +

                              "}" +
                              "#main-table table,#main-table th,#main-table td{" +
                               "white-space: nowrap;}" +
                          "</style>" +
                      "</head>" +
                      "<body>");
                    #endregion

                    string oemDealerQuery = string.Empty;



                    string fileAppoinmentDate = string.Empty;
                    string maxAppointmentdate = "select distinct  Convert(varchar(10),max(SlotBookingDate),105) as MaxAppointmentdate  " +
                    " from	HSRPrecords H with(nolock) join DealerAffixationCenter d  with(nolock) on H.HSRP_StateID=d.StateID and H.Affix_Id=d.DealerAffixationID  join BookMyHSRPAppointment B on H.orderno=B.orderno  " +
                   " where   Navembid='" + Navembcode + "'      and  NewPdfRunningNo is null and erpassigndate is not null and  h.OrderStatus='New Order' and  IsBookMyHsrpRecord='Y' and  h.affix_id is not NULL  and   d.TypeofDelivery in('Home','RWA') and  ((b.IsFrame is null) or (b.IsFrame='N')) ";
                    DataTable dtmax = Utils.Utils.GetDataTable(maxAppointmentdate, CnnString);

                    if (dtmax.Rows.Count > 0)
                    {
                        fileAppoinmentDate = dtmax.Rows[0]["MaxAppointmentdate"].ToString();
                    }
                    fileName = "HomeDeliveryWithOutPlastickPacking" + "-" + fileAppoinmentDate + "-" + filePrefix + "-" + Navembcode + ".pdf";
                    filePath = dir + fileName;



                    oemDealerQuery = "select distinct d.oemid as oemid, (select name  from oemmaster where oemid=d.oemid) as oemname,d.dealerid as dealerid,d.DispatchHub,d.DealerAffixationcenterName as Dealername,  d.DealerAffixationCenterAddress as Address, " +
                     "d.DealerAffixationID as SubDealerId, H.dealerid AS ParentDealerId,SlotBookingDate from	HSRPrecords H with(nolock) join DealerAffixationCenter d  with(nolock) on H.HSRP_StateID=d.StateID and H.Affix_Id=d.DealerAffixationID  join BookMyHSRPAppointment B on H.orderno=B.orderno  " +
                    " where   Navembid='" + Navembcode + "'     and  NewPdfRunningNo is null and erpassigndate is not null and  h.OrderStatus='New Order' and  IsBookMyHsrpRecord='Y' and  h.affix_id is not NULL   and  d.TypeofDelivery in('Home','RWA') and  ((b.IsFrame is null) or (b.IsFrame='N')) ";

                    #region
                    DataTable dtOD = Utils.Utils.GetDataTable(oemDealerQuery, CnnString);

                    if (dtOD.Rows.Count > 0)
                    {
                        string allStrProductionSheetNo = string.Empty;
                        Session["ECWiseProductionsheet="] = null;
                        foreach (DataRow drOD in dtOD.Rows)
                        {

                            string oemid = drOD["oemid"].ToString().Trim();
                            string dealerid = drOD["dealerid"].ToString().Trim();
                            string oemname = drOD["oemname"].ToString().Trim();
                            string dealername = drOD["Dealername"].ToString().Trim();
                            string Address = drOD["Address"].ToString().Trim();
                            string dispatchHub= drOD["DispatchHub"].ToString().Trim();
                            DateTime AppointmentDate = Convert.ToDateTime(drOD["SlotBookingDate"].ToString().Trim());


                            string strsubaffixid = drOD["SubDealerId"].ToString();
                            string strParentDealerId = drOD["ParentDealerId"].ToString();

                            //start sql query
                            #region
                            string productionQuery = string.Empty;

                            DataTable dtProduction = new DataTable();


                            SqlConnection con = new SqlConnection(CnnString);
                            SqlCommand cmd = new SqlCommand("USP_BookMYHSRPProductionSheetWithOutFramesNEW", con);
                            cmd.CommandType = CommandType.StoredProcedure;
                            con.Open();

                            cmd.Parameters.AddWithValue("@HSRP_StateID", HSRP_StateID);
                            cmd.Parameters.AddWithValue("@navembid", Navembcode);
                            cmd.Parameters.AddWithValue("@Dealerid", strParentDealerId);
                            cmd.Parameters.AddWithValue("@Affix_Id", strsubaffixid);
                            cmd.Parameters.AddWithValue("@AppointmentDate", AppointmentDate);
                            SqlDataAdapter da = new SqlDataAdapter(cmd);
                            // dtProduction = new DataTable();
                            da.Fill(dtProduction);
                            con.Close();

                            // string fileAppoinmentDate = AppointmentDate.ToString("dd/MM/yyyy");

                            #endregion
                            //end sql query

                            #region
                            //DataTable dtProduction = Utils.Utils.GetDataTable(productionQuery, CnnString);
                            if (dtProduction.Rows.Count > 0)
                            {

                                findRecord = true;
                                string strRunningNo = string.Empty;
                                //string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) as maxSheetNo from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + NAVEMBID + "'";
                                string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) from EmbossingCentersNew where    Emb_Center_Id='" + Navembcode + "' ";
                                string strCom = Utils.Utils.getScalarValue(strSel, CnnString);
                                //DataTable dtSheetNo = fillDataTable(strSel, CnnString);
                                if (strCom.Equals(0) || strCom.Length == 0)
                                {
                                    strRunningNo = "0000001";
                                }
                                else
                                {
                                    strRunningNo = string.Format("{0:0000000}", Convert.ToInt32(strCom) + 1);
                                }

                                string strRequeNo = "select (prefixtext+right('00000'+ convert(varchar,Lastreqno+1),5)) as Reqno from prefix_Requisition  where hsrp_stateid='" + ddlStateName.SelectedValue + "' and prefixfor='Req No'";
                                ReqNum = Utils.Utils.getScalarValue(strRequeNo, CnnString);

                                string strPRFIX = "select PrefixText from EmbossingCentersNew where  Emb_Center_Id='" + Navembcode + "'";
                                string strPRFIXCom = Utils.Utils.getScalarValue(strPRFIX, CnnString);
                                strProductionSheetNo = strPRFIXCom + strRunningNo;

                                string RTOLocationName = dtProduction.Rows[0]["RTOLocationName"].ToString();

                                string FS1 = string.Empty;
                                string FS2 = string.Empty;



                                #region


                                html.Append("<div style='page-break-before: avoid;page-break-inside: avoid;page-break-after: always;'>" +

                                    "<table style='width:100%;border: 0px;'>" +
                                        "<tr  style='border: 0px;'>" +
                                            "<td colspan='4'   style='border: 0px; '>" +
                                                "<div style='text-align:left'><b>Sheet Generation Date:</b> " + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</div>" +
                                            "</td>" +

                                             "<td colspan='3'  style='border: 0px; '>" +
                                                "<div style='text-align:left'><b>ROSMERTA SAFETY SYSTEMS LIMITED</b> " + "</div>" +
                                            "</td>" +


                                             "<td  colspan='6' style='border: 0px;'>" +
                                                "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Production Sheet No:</b> " + strProductionSheetNo + " </div>" +
                                            "</td>" +

                                        "</tr>" +

                                          "<tr style='border: 0px;'>" +
                                            "<td colspan='4' style='border: 0px;'>" +
                                                "<div style='text-align:left'><b>State:</b> " + HSRPStateName + "</div>" +
                                            "</td>" +

                                             "<td colspan='3' style='border: 0px; '>" +
                                                                                               //"<div style='text-align:left'><b>Book My HSRP (Home Delivery) </b> " + "</div>" +
                                   "<div style='text-align:left;font-size:22px;''><b> MHHSRP (Home Delivery) </b> " + "</div>" +
                             "</td>" +


                                             "<td colspan='6' style='border: 0px;'>" +
                                        "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Hub/Dealer (ID - Name):</b> " + dealerid + "/" + dispatchHub + " </div>" +
                              // "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Hub/Dealer (ID - Name):</b> " + dealerid + "/" + dealername + "/" + strsubaffixid + " </div>" +
                                            "</td>" +

                                        "</tr>" +

                                            "<tr style='border: 0px;'>" +
                                            "<td colspan='4' style='border: 0px;'>" +
                                                "<div style='text-align:left'><b>EC Location: </b> " + RTOLocationName + "</div>" +
                                            "</td>" +

                                             "<td colspan='3' style='border: 0px;'>" +
                                               "<div style='text-align:left;font-size:22px;'><b>Appointment Date: " + dtProduction.Rows[0]["AppointmentDate"].ToString() + "</ b >" + "</div>" +
                                            "</td>" +


                                             "<td  colspan='6' style='border: 0px;'>" +
                                                               "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Hub/Dealer Address:</b> " + Address + " </div>" +
                                            "</td>" +

                                        "</tr>" +




                                        "<tr>" +
                                             "<td style='text-align:center;white-space: nowrap'>Sr. No.</td>" +
                                                "<td style='width:15%;white-space: nowrap'>Vehicle No.</td>" +
                                                "<td>Front Plate Size</td>" +

                                                 "<td style='width:15%;white-space: nowrap'>Front Laser No.</td>" +

                                             "<td>Rear Plate Size</td>" +


                                                 "<td style='width:15%;white-space: nowrap'>Rear Laser No.</td>" +


                                             "<td style='white-space: nowrap'>H. S. Foil </td>" +
                                             "<td style='white-space: nowrap'>Caution Sticker</td>" +
                                             "<td style='white-space: nowrap'>Fuel Type</td>" +
                                             "<td style='white-space: nowrap'>VT</td>" +
                                               "<td style='white-space: nowrap'>VC</td>" +
                                             "<td style='white-space: nowrap'>Frame</ td>" +
                                               "<td style='white-space: nowrap'>Pin Code</ td>" +

                                         "</tr>");

                                #endregion

                                #region
                                string strtvspono = "";

                                int j = 0;
                                int total = 0;
                                StringBuilder UpdateSQL = new StringBuilder();
                                for (int i = 0; i <= dtProduction.Rows.Count - 1; i++)
                                {

                                    j = j + 1;

                                    if (total == 22)
                                    {
                                        total = 0;

                                        html.Append("<div style='page-break-before: avoid;page-break-inside: avoid;page-break-after: always;'>" +
                              "<table style='width:100%;border: 0px;'>" +
                                  "<tr style='border: 0px;'>" +
                                      "<td colspan='4'  style='border: 0px;'>" +
                                          "<div style='text-align:left'><b>Sheet Generation Date:</b> " + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</div>" +
                                      "</td>" +

                                       "<td colspan='3'  style='border: 0px;'>" +
                                          "<div style='text-align:left'><b>ROSMERTA SAFETY SYSTEMS LIMITED</b> " + "</div>" +
                                      "</td>" +


                                       "<td colspan='6' style='border: 0px; '>" +
                                          "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Production Sheet No:</b> " + strProductionSheetNo + " </div>" +
                                      "</td>" +

                                  "</tr>" +

                                    "<tr style='border: 0px;'>" +
                                      "<td colspan='4' style='border: 0px;'>" +
                                          "<div style='text-align:left'><b>State:</b> " + HSRPStateName + "</div>" +
                                      "</td>" +

                                       "<td colspan='3' style='border: 0px; '>" +
                                                                                         //"<div style='text-align:left'><b>Book My HSRP (Home Delivery) </b> " + "</div>" +
                       "<div style='text-align:left;font-size:22px;''><b> MHHSRP (Home Delivery) </b> " + "</div>" +
                     "</td>" +


                                       "<td colspan='6' style='border: 0px;'>" +
                                                                               //   "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Hub/Dealer (ID - Name):</b> " + dealerid + "/" + dealername + "/" + strsubaffixid + " </div>" +
                              "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Hub/Dealer (ID - Name):</b> " + dealerid + "/" + dispatchHub + " </div>" +
                                 "</td>" +

                                  "</tr>" +

                                      "<tr style='border: 0px;'>" +
                                      "<td colspan='4'  style='border: 0px;'>" +
                                          "<div style='text-align:left'><b>EC Location: </b> " + RTOLocationName + "</div>" +
                                      "</td>" +

                                       "<td colspan='3' style='border: 0px;'>" +
                                         "<div style='text-align:left;font-size:22px;'><b>Appointment Date: " + dtProduction.Rows[0]["AppointmentDate"].ToString() + "</ b >" + "</div>" +
                                      "</td>" +


                                       "<td colspan='6'  style='border: 0px;'>" +
                                        "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Hub/Dealer Address:</b> " + Address + " </div>" +
                                      "</td>" +

                                  "</tr>" +




                                     "<tr>" +
                                         "<td style='text-align:center;white-space: nowrap'>Sr. No.</td>" +
                                            "<td style='width:15%;white-space: nowrap'>Vehicle No.</td>" +
                                            "<td>Front Plate Size</td>" +

                                             "<td style='width:15%;white-space: nowrap'>Front Laser No.</td>" +

                                         "<td>Rear Plate Size</td>" +


                                             "<td style='width:15%;white-space: nowrap'>Rear Laser No.</td>" +


                                         "<td style='white-space: nowrap'>H. S. Foil </td>" +
                                         "<td style='white-space: nowrap'>Caution Sticker</td>" +
                                         "<td style='white-space: nowrap'>Fuel Type</td>" +
                                         "<td style='white-space: nowrap'>VT</td>" +
                                           "<td style='white-space: nowrap'>VC</td>" +
                                         "<td style='white-space: nowrap'>Frame</ td>" +
                                           "<td style='white-space: nowrap'>Pin Code</ td>" +

                                     "</tr>");


                                    }


                                    string RS1 = string.Empty;
                                    string RS2 = string.Empty;
                                    total = total + 1;
                                    //foreach (DataRow drProduction in dtProduction.Rows)
                                    //{
                                    string HsrprecordID = dtProduction.Rows[i]["hsrprecordID"].ToString().Trim();
                                    string SRNo = dtProduction.Rows[i]["SRNo"].ToString().Trim();
                                    //string ORD = "";// drProduction["ORD"].ToString().Trim();
                                    string VC = dtProduction.Rows[i]["VehicleClass"].ToString().Trim();
                                    string VT = dtProduction.Rows[i]["VehicleType"].ToString().Trim();
                                    string VehicleNo = dtProduction.Rows[i]["VehicleRegNo"].ToString().Trim();
                                    //string VT = drProduction["VehicleType"].ToString().Trim();
                                    //string ChassisNo = drProduction["ChassisNo"].ToString().Trim();
                                    //string EngineNo = drProduction["EngineNo"].ToString().Trim();
                                    string FuelType = dtProduction.Rows[i]["FuelType"].ToString().Trim();
                                    string FrontPSize = dtProduction.Rows[i]["FrontProductCode"].ToString().Trim();

                                    string FrontLaserNo = dtProduction.Rows[i]["HSRP_Front_LaserCode"].ToString().Trim();
                                    if (FrontLaserNo != "")
                                    {
                                        FS1 = dtProduction.Rows[i]["HSRP_Front_LaserCode"].ToString().Substring(0, 7);
                                        FS2 = dtProduction.Rows[i]["HSRP_Front_LaserCode"].ToString().Substring(7, 5);
                                    }

                                    string RearPSize = dtProduction.Rows[i]["RearProductCode"].ToString().Trim();
                                    string RearLaserNo = dtProduction.Rows[i]["HSRP_Rear_LaserCode"].ToString().Trim();
                                    if (RearLaserNo != "")
                                    {
                                        RS1 = dtProduction.Rows[i]["HSRP_Rear_LaserCode"].ToString().Substring(0, 7);
                                        RS2 = dtProduction.Rows[i]["HSRP_Rear_LaserCode"].ToString().Substring(7, 5);
                                    }



                                    string StickerColor = dtProduction.Rows[i]["stickerColor"].ToString().Trim();
                                    string HotStampingFoilColour = dtProduction.Rows[i]["HotStampingFoilColour"].ToString().Trim();
                                    string Frame = dtProduction.Rows[i]["Frame"].ToString().Trim();
                                    //string AppointmentDate = drProduction["AppointmentDate"].ToString().Trim();
                                    string Pincode = dtProduction.Rows[i]["Pincode"].ToString().Trim();
                                    // "<b>Production Sheet No: " + strProductionSheetNo + "<br /> </b>" +
                                    html.Append("<tr>" +
                                       "<td style='text-align:center;white-space: nowrap'>" + SRNo + "</td>" +
                                       "<td style='font-size:20px;white-space: nowrap' >" + "<b>" + VehicleNo + "</td>" +
                                       "<td style='white-space: nowrap'> " + FrontPSize + "  </td> " +
                                       "<td style='font-size:20px;white-space: nowrap'>" + "<b>" + FS1 + "<b>" + FS2 + "</b> </td>" +


                                       "<td style='white-space: nowrap'>" + RearPSize + "</td>" +
                                       "<td style='font-size:20px;white-space: nowrap'>" + "<b>" + RS1 + "<b>" + RS2 + "</b> </td>" +

                                         "<td style='white-space: nowrap'>" + HotStampingFoilColour + "</td>" +
                                        "<td style='white-space: nowrap'>" + StickerColor + "</td>" +
                                        "<td style='white-space: nowrap'>" + FuelType + "</td>" +
                                        "<td style='white-space: nowrap'>" + VT + "</td>" +
                                        "<td style='white-space: nowrap'>" + VC + "</td>" +


                                       "<td style='white-space: nowrap'>" + Frame + "</td>" +

                                         "<td style='white-space: nowrap'>" + Pincode + "</td>" +
                                   "</tr>");



                                    UpdateSQL.Append("update hsrprecords set sendtoProductionStatus='Y', NAVPDFFlag='1', NewPdfRunningNo='" + strProductionSheetNo + "', Requisitionsheetno='" + ReqNum + "',PdfDownloadDate=GetDate(), pdfFileName='" + fileName + "', PDFDownloadUserID='1' where hsrprecordID='" + HsrprecordID + "';");




                                }

                                allStrProductionSheetNo += "," + strProductionSheetNo;

                                allStrProductionSheetNo = allStrProductionSheetNo.TrimStart(',');
                                Session["ECWiseProductionsheet="] = allStrProductionSheetNo;
                                #endregion
                                html.Append("</table>");

                                html.Append("</div>");

                                if (UpdateSQL.ToString().Length > 0)
                                {
                                    Utils.Utils.ExecNonQuery(UpdateSQL.ToString(), CnnString);
                                    string StrSqlUpdateECQuery = "update EmbossingCentersNew set NewProductionSheetRunningNo='" + strProductionSheetNo + "' " +
                                     "where  Emb_Center_Id='" + Navembcode + "'";
                                    Utils.Utils.ExecNonQuery(StrSqlUpdateECQuery, CnnString); // uncomment after testing
                                }

                                //string StrSqlUpdateECQuery = "update EmbossingCentersNew set NewProductionSheetRunningNo='" + strProductionSheetNo + "' " +
                                // "where  Emb_Center_Id='" + Navembcode + "'";
                                //Utils.Utils.ExecNonQuery(StrSqlUpdateECQuery, CnnString); // uncomment after testing
                            }

                            #endregion

                        }
                    }// close oemDealerQuery
                    #endregion
                    #region "Req Generate"
                    string strSqlQuery1 = "select CompanyName from hsrpstate where hsrp_stateid='" + ddlStateName.SelectedValue + "'";
                    strCompanyName = Utils.Utils.getScalarValue(strSqlQuery1, CnnString);

                    string strEMBName = " select EmbCenterName from EmbossingCentersNew where   Emb_Center_Id='" + Navembcode + "'";
                    string strComNew = Utils.Utils.getScalarValue(strEMBName, CnnString);

                    string strReqNo = "select (prefixtext+right('00000'+ convert(varchar,Lastreqno+1),5)) as Reqno from prefix_Requisition  where hsrp_stateid='" + ddlStateName.SelectedValue + "' and prefixfor='Req No'";
                    string strReqNumber = Utils.Utils.getScalarValue(strReqNo, CnnString);

                    SQLString = "Exec [laserreqSlip1DashBoardBookMYHSRP]  '" + Navembcode + "' ,  '" + ReqNum + "'";
                    DataTable dtResult = Utils.Utils.GetDataTable(SQLString, CnnString);
                    string strQuery = string.Empty;
                    string strRtoLocationName = string.Empty;
                    int Itotal = 0;

                    html.Append("<div style='width:100%;height:100%;'>" +
                                        "<table style='width:100%'>" +

                                            "<tr>" +
                                                "<td colspan='12'>" +
                                                    "<div style='text-align:center;padding:8px;'>" +
                                                        "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + strCompanyName + "</b>" +

                                                        "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                         "<b style='font-size:20px;'>" + "MATERIAL REQUSITION NOTE" + "</b>" +

                                                    "</div>" +
                                                "</td>" +
                                            "</tr>" +

                                            "<tr>" +
                                                "<td colspan='12'>" +
                                                    "<div style='text-align:center;padding:8px;'>" +
                                                        "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Production Sheet Date :" + DateTime.Now.ToString("dd-MM-yyyy") + "</b>" +
                                                        "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                         "<b style='font-size:20px;'>" + "" + "</b>" +
                                                    "</div>" +
                                                "</td>" +
                                            "</tr>" +
                                            "<tr>" +
                                                "<td colspan='6'>" +
                                                    "<div style='text-align:left;'>" +
                                                        "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>REQ.NO:-</b>" +
                                                        "" + strReqNumber + "" +
                                                    "</div>" +
                                                "</td>" +
                                                "<td colspan='6'>" +
                                                    "<div style='text-align:left;'>" +
                                                        "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>Embossing Center:</b>" +
                                                        "" + strComNew + "" +
                                                    "</div>" +
                                                "</td>" +
                                            "</tr>" +
                                            "<tr>" +
                                                "<td colspan='12'>" +
                                                    "<div style='text-align:left;padding:2px;'>" +
                                                        "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "State:" + ddlStateName.SelectedItem.Text + "</b>" +
                                                        "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                         "<b style='font-size:20px;'>" + "" + "</b>" +
                                                    "</div>" +
                                                "</td>" +
                                            "</tr>" +
                                             "<tr>" +
                                                "<td colspan='12'>" +
                                                    "<div style='text-align:left;padding:8px;'>" +
                                                        "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + " " + "</b>" +
                                                        "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                         "<b style='font-size:20px;'>" + "" + "</b>" +
                                                    "</div>" +
                                                "</td>" +
                                            "</tr>" +

                                            "<tr>" +
                                                "<td colspan='1' style='text-align:center'>SR.N.</td>" +
                                                "<td colspan='3'>Product Size</td>" +
                                                "<td colspan='1'>Laser Count</td>" +
                                                "<td colspan='1'>Start Laser No</td>" +
                                                "<td colspan='1'>End Laser No</td>" +

                                            "</tr>");


                    if (dtResult.Rows.Count > 0)
                    {
                        for (int i = 0; i < dtResult.Rows.Count; i++)
                        {
                            string ID = dtResult.Rows[i]["ID"].ToString();
                            string productcode = dtResult.Rows[i]["productcode"].ToString();
                            string LaserCount = dtResult.Rows[i]["LaserCount"].ToString();
                            Itotal = Convert.ToInt32(dtResult.Rows[i]["Total"].ToString());
                            string BeginLaser = dtResult.Rows[i]["BeginLaser"].ToString();
                            string EndLaser = dtResult.Rows[i]["EndLaser"].ToString();




                            html.Append("<tr>" +
                               "<td colspan='1' style='text-align:left'>" + ID + "</td>" +
                               "<td colspan='3'>" + productcode + "</td>" +

                               "<td colspan='1'>" + LaserCount + "</td>" +
                               "<td colspan='1'>" + BeginLaser + "</td>" +
                               "<td colspan='1'>" + EndLaser + "</td>" +

                           "</tr>");
                        }

                        try
                        {
                            //start updating hsrprecords 
                            string Query = "update prefix_Requisition set lastno=lastno+1,Lastreqno=Lastreqno+1 where hsrp_stateid='" + ddlStateName.SelectedValue + "' and prefixfor='Req No'";
                            Utils.Utils.ExecNonQuery(Query, CnnString);
                        }
                        catch (Exception ev)
                        {
                            Label1.Text = "prefix Requisition update error: " + ev.Message;
                        }
                    }
                    html.Append("<tr>" +
                             "<td colspan='1' style='text-align:center' > " + "<b>Grand Total:</b>" + "</td>" +
                             "<td colspan='3'>" + "" + "</td>" +

                             "<td colspan='1'>" + Itotal + "</td>" +
                             "<td colspan='1'>" + " " + "</td>" +
                             "<td colspan='1'>" + " " + "</td>" +

                         "</tr>");




                    html.Append("<tr>" +
                     "<td colspan='2' > " + "<b>REQUESTED BY </b>" + "</td>" +
                     "<td colspan='2'>" + "<b>AUTHORIZED BY </b>" + "</td>" +

                     "<td colspan='2'>" + "<b>ISSUED BY </b>" + "</td>" +
                     "<td colspan='2'>" + "<b>RECEIVED BY</b>" + "</td>" +


                 "</tr>");

                    html.Append("<tr>" +
                                              "<td colspan='12'>" +
                                                  "<div style='text-align:left;padding:2px;'>" +
                                                      "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Name" + "</b>" +

                                                  "</div>" +
                                              "</td>" +
                                          "</tr>");

                    html.Append("<tr>" +
                                             "<td colspan='12'>" +
                                                 "<div style='text-align:left;padding:2px;'>" +
                                                     "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Designation" + "</b>" +

                                                 "</div>" +
                                             "</td>" +
                                         "</tr>");

                    html.Append("<tr>" +
                                           "<td colspan='12'>" +
                                               "<div style='text-align:right;padding:8px;'>" +
                                                   "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Download By: Admin " + "</b>" +

                                               "</div>" +
                                           "</td>" +
                                       "</tr>");

                    html.Append("<tr>" +
                                          "<td colspan='12'>" +
                                              "<div style='text-align:right;padding:8px;'>" +
                                                  "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Sheet Generated By :" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</b>" +

                                              "</div>" +
                                          "</td>" +
                                      "</tr>");


                    html.Append("<tr style='visibility: hidden;'>" + "<td  ><p style=\"page-break-after:always\"/></td></tr>");



                    //html.Append("</table>");

                    //html.Append("</div>");

                    #region "PS Summary Report"
                    string PS = string.Empty;
                    if (Session["ECWiseProductionsheet="] != null)
                    {
                        PS = Session["ECWiseProductionsheet="].ToString();
                    }
                    /*
                     * Close body & HTMl Tag
                     */

                    SQLString = "Exec [USP_BMPSSummaryHomedelivery]  '" + PS + "' ";
                    DataTable dtResultsummary = Utils.Utils.GetDataTable(SQLString, CnnString);

                    html.Append("<div style='width:100%;height:100%;'>" +
                                 "<table style='width:100%'>" +

                                     "<tr>" +
                                         "<td colspan='12'>" +
                                             "<div style='text-align:center;padding:8px;'>" +
                                                 "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Home Delivery Summary Report" + "</b>" +

                                             "</div>" +
                                         "</td>" +
                                     "</tr>" +

                                      "<tr>" +
                                                "<td colspan='1' style='text-align:center'>PS NO</td>" +
                                                "<td colspan='3'>Affixation Pin Code</td>" +
                                                "<td colspan='1'>2W</td>" +
                                                "<td colspan='1'>3W</td>" +
                                                "<td colspan='1'>4W</td>" +
                                                "<td colspan='1'>OTH</td>" +
                                                  "<td colspan='1'>Total</td>" +

                                            "</tr>");

                    if (dtResultsummary.Rows.Count > 0)
                    {
                        for (int i = 0; i < dtResultsummary.Rows.Count; i++)
                        {
                            string PSNo = dtResultsummary.Rows[i]["PSNo"].ToString();
                            string AffixationdealerName = dtResultsummary.Rows[i]["Affixation Pin Code"].ToString();
                            string W2 = dtResultsummary.Rows[i]["2W"].ToString();
                            string W3 = dtResultsummary.Rows[i]["3W"].ToString();
                            string W4 = dtResultsummary.Rows[i]["4W"].ToString();
                            string OTH = dtResultsummary.Rows[i]["OTH"].ToString();
                            string Total = dtResultsummary.Rows[i]["Total"].ToString();




                            html.Append("<tr>" +
                               "<td colspan='1' style='text-align:left'>" + PSNo + "</td>" +
                               "<td colspan='3'>" + AffixationdealerName + "</td>" +

                               "<td colspan='1'>" + W2 + "</td>" +
                               "<td colspan='1'>" + W3 + "</td>" +
                               "<td colspan='1'>" + W4 + "</td>" +
                                "<td colspan='1'>" + OTH + "</td>" +
                                 "<td colspan='1'>" + Total + "</td>" +

                           "</tr>");
                        }
                    }

                    html.Append("</table>");

                    html.Append("</div>");
                    html.Append("<div style='text-align:left;padding:8px;'>" +
                                            "<b style='font-size:25px;margin-top:2px;margin-bottom:2px;'>" + "<p>&#x25CF The Owner has to furnish the original FIR/SDE copy to the Fitment center at the time of fitment of HSRP.</p><p>&#x25CF The Owner has to deposit the damaged plate at the fitment center at the time of fitment of HSRP.</p><p>&#x25CF The fitment center has to retail the old TV/NTV plate in case of fitment due to conversion,re-assignment.</p>" + "</b>" +

                                        "</div>"
);


                    html.Append("</div>");

                    #endregion



                    #endregion

                    /*
                     * Close body & HTMl Tag
                     */


                    if (findRecord)
                    {






                        string strSaveSQlQuery = "insert into ProductionSheetAutoGenerated_List (HSRP_StateID, State_Code, RTO_ID, Emb_Center_Id, FileName, Productiondate, ProductionTime, OrderNo) values " +
                                    "('" + ddlStateName.SelectedValue + "', '" + HSRPStateShortName + "', '', '" + Navembcode + "', '" + fileName + "', '" + DateTime.Now.ToString("yyyy-MM-dd") + "', '" + DateTime.Now.ToString("HH:mm:ss") + "', '" + filePrefix + "') ";
                        Utils.Utils.ExecNonQuery(strSaveSQlQuery, CnnString); // uncomment after testing
                        Label1.Text = Label1.Text + "successfully created production sheet of State: " + HSRPStateShortName + ", , Emb_Center_Id: " + Navembcode + Environment.NewLine;

                        //lblLog.Text = lblLog.Text + Environment.NewLine + "Start Production Query at:" + DateTime.Now;



                        #region
                        try
                        {
                            if (!Directory.Exists(dir))
                            {


                                Directory.CreateDirectory(dir);

                            }

                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        try
                        {
                            var pdf = Pdf
                                    .From(html.ToString())
                                    .OfSize(PaperSize.A4)
                                    .WithTitle("Title")
                                    .WithoutOutline()
                                    .WithMargins(1.25.Centimeters())
                                    .Landscape()
                                    .Comressed()
                                    .Content();

                            FileStream readStream = File.Create(filePath);
                            BinaryWriter binaryWriter = new BinaryWriter(readStream);

                            // Write the binary data to the file
                            binaryWriter.Write(pdf);
                            binaryWriter.Close();
                            readStream.Close();
                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        #endregion
                    }

                    #endregion

                }//close foreach stateEcQuery
            }
        }


        protected void btnBMHSRPPSSheet_Click(object sender, EventArgs e)
        {

           // string Navembid = Session["Navembid"].ToString();
            if (ddlStateName.SelectedItem.Text == "--Select State--")
            {
                lblErrMess.Text = "Please select  State";
                return;
            }
            string sqlPrefixQuery = "select top 1 CONVERT(INT, isnull(OrderNo,'0')) + 1 as orderNo from ProductionSheetAutoGenerated_List order by id desc";
            DataTable dtPrefix = Utils.Utils.GetDataTable(sqlPrefixQuery, CnnString);

            if (dtPrefix.Rows.Count > 0)
            {
                filePrefix = dtPrefix.Rows[0]["orderNo"].ToString();
            }
            else
            {
                filePrefix = "1";
            }

            //txtFilePrefix.Text = filePrefix;

            if (filePrefix.Length > 0)
            {
                SheetGenerationBookMyHSRP();
         
            }
        }

        protected void btnOLA_Click(object sender, EventArgs e)
        {

            if (ddlStateName.SelectedItem.Text == "--Select State--")
            {
                lblErrMess.Text = "Please select  State";
                return;
            }
            string sqlPrefixQuery = "select top 1 CONVERT(INT, isnull(OrderNo,'0')) + 1 as orderNo from ProductionSheetAutoGenerated_List order by id desc";
            DataTable dtPrefix = Utils.Utils.GetDataTable(sqlPrefixQuery, CnnString);

            if (dtPrefix.Rows.Count > 0)
            {
                filePrefix = dtPrefix.Rows[0]["orderNo"].ToString();
            }
            else
            {
                filePrefix = "1";
            }

            //txtFilePrefix.Text = filePrefix;

            if (filePrefix.Length > 0)
            {
                SheetGenerationOLAHubWise();
            }

        }

        protected void btnTwenty_Click(object sender, EventArgs e)
        {
            if (ddlStateName.SelectedItem.Text == "--Select State--")
            {
                lblErrMess.Text = "Please select  State";
                return;
            }


            string sqlPrefixQuery = "select top 1 CONVERT(INT, isnull(OrderNo,'0')) + 1 as orderNo from ProductionSheetAutoGenerated_List order by id desc";
            DataTable dtPrefix = Utils.Utils.GetDataTable(sqlPrefixQuery, CnnString);

            if (dtPrefix.Rows.Count > 0)
            {
                filePrefix = dtPrefix.Rows[0]["orderNo"].ToString();
            }
            else
            {
                filePrefix = "1";
            }

            SheetGenerationTwentyTwo();

        }

        private void SheetGenerationTwentyTwo()
        {
            string stateECQuery = string.Empty;

            //string Navembid = Session["Navembid"].ToString();
            FillUserDetails();
            string ReqNum = string.Empty;

            if ((ddlStateName.SelectedValue == "20") || (ddlStateName.SelectedValue == "10"))
            {


                stateECQuery = "select distinct a.HSRP_StateID, (select HSRPStateName from hsrpstate c where c.hsrp_stateid=a.hsrp_stateid) as HSRPStateName,(select HSRPStateShortName from hsrpstate d where d.hsrp_stateid=a.hsrp_stateid) as HSRPStateShortName, navembcode from hsrprecords a with(nolock),rtolocation b where a.HSRP_StateID ='" + ddlStateName.SelectedValue + "' and   b.Navembcode not like '%CODO%'    and a.rtolocationid=b.rtolocationid and NewPdfRunningNo is null and erpassigndate is not null and OrderStatus='New Order' AND b.NAVEMBID='" + Navembid + "'   order by  a.HSRP_StateID";
            }
            else
            {
                stateECQuery = "select distinct a.HSRP_StateID, (select HSRPStateName from hsrpstate c where c.hsrp_stateid=a.hsrp_stateid) as HSRPStateName,(select HSRPStateShortName from hsrpstate d where d.hsrp_stateid=a.hsrp_stateid) as HSRPStateShortName, navembcode from hsrprecords a with(nolock),rtolocation b where a.HSRP_StateID ='" + ddlStateName.SelectedValue + "' and   b.Navembcode not like '%CODO%'    and a.rtolocationid=b.rtolocationid and NewPdfRunningNo is null and erpassigndate is not null and OrderStatus='New Order' and   b.navembcode like  '%'+(select HSRPStateShortName from hsrpstate d where d.hsrp_stateid=a.hsrp_stateid)+'%'  and b.NAVEMBID='" + Navembid + "'    order by  a.HSRP_StateID";

            }


            DataTable dtSE = Utils.Utils.GetDataTable(stateECQuery, CnnString);
            if (dtSE.Rows.Count > 0)
            {
                foreach (DataRow dr in dtSE.Rows)
                {
                    #region

                    string HSRP_StateID = dr["HSRP_StateID"].ToString().Trim();
                    string HSRPStateName = dr["HSRPStateName"].ToString().Trim();
                    string HSRPStateShortName = dr["HSRPStateShortName"].ToString().Trim();
                    string Navembcode = dr["Navembcode"].ToString().Trim();

                    string dir = dirPath + DateTime.Now.ToString("yyyy-MM-dd") + "\\" + HSRPStateShortName + "\\";
                    //string dir = dirPath + "/" + DateTime.Now.ToString("yyyy-MM-dd") + "/" + HSRPStateShortName + "/";
                    // string fileName = filePrefix + "-" + Navembcode + ".pdf";
                    string fileName = "AllOEM" + "-" + filePrefix + "-" + Navembcode + ".pdf";
                    string filePath = dir + fileName;

                    //string folderpath = ConfigurationManager.AppSettings["InvoiceFolder"].ToString() + "/" + FinYear + "/" + oemid + "/" + HSRPStateID + "/";

                    StringBuilder html = new StringBuilder();

                    Boolean findRecord = false;
                    string strProductionSheetNo = string.Empty;

                    /*
                     *  Start body & HTMl Tag
                     */
                    #region
                    html.Append(
                        "<!DOCTYPE html>" +
                        "<html>" +
                        "<head>" +
                            "<meta charset='UTF-8'><title>Title</title>" +
                            "<style>" +
                                "@page {" +
                                    /* headers*/
                                    "@top-left {" +
                                        "content: 'Left header';" +
                                    "}" +
                                    "@top-right {" +
                                        "content: 'Right header';" +
                                    "}" +

                                    /* footers */
                                    "@bottom-left {" +
                                        "content: 'Lorem ipsum';" +
                                    "} " +
                                    "@bottom-right {" +
                                        "content: 'Page ' counter(page) ' of ' counter(pages);" +
                                    "}" +
                                    "@bottom-center  {" +
                                        "content:element(footer);" +
                                    "}" +
                                "}" +
                                 "#footer {" +
                                    "position: running(footer);" +
                                "}" +
                                "table {" +
                                  "border-collapse: collapse;" +
                                "}" +

                                "table, th, td {" +
                                    "border: 1px solid black;" +
                                    "text-align: left;" +
                                    "vertical-align: top;" +
                                    "padding:5px;" +
                                "}" +
                            "</style>" +
                        "</head>" +
                        "<body>");
                    #endregion

                    string oemDealerQuery = string.Empty;



                    oemDealerQuery = "select distinct om.oemid as oemid, om.name as oemname, dm.dealerid, dm.dealername, dm.dealercode, dm.Address, dm.HSRP_StateID, " +
                        "dm.RTOLocationID from oemmaster om " +
                        "left join dealermaster dm on dm.oemid = om.oemid where dm.HSRP_StateID =" + HSRP_StateID + " and " +
                        //"dm.RTOLocationID in (select RTOLocationID from rtolocation where Navembcode='" + Navembcode + "' ) and " +
                        "dm.dealerid in (select distinct dealerid from hsrprecords  with(nolock) where NewPdfRunningNo is null and erpassigndate is not null and OrderStatus='New Order'and HSRP_StateID =" + HSRP_StateID + " and Navembid='" + Navembcode + "') and Om.OEMID   in('1303')";
                    //"dm.dealerid in (select distinct dealerid from hsrprecords where isnull(NewPdfRunningNo,'') = '' and isnull(erpassigndate,'') != '' and HSRP_StateID =" + HSRP_StateID + ") and Om.OEMID not  in('21','40','12','20')";

                    #region
                    DataTable dtOD = Utils.Utils.GetDataTable(oemDealerQuery, CnnString);
                    if (dtOD.Rows.Count > 0)
                    {
                        foreach (DataRow drOD in dtOD.Rows)
                        {

                            string oemid = drOD["oemid"].ToString().Trim();
                            string dealerid = drOD["dealerid"].ToString().Trim();
                            string oemname = drOD["oemname"].ToString().Trim();
                            string dealername = drOD["dealername"].ToString().Trim();
                            string Address = drOD["Address"].ToString().Trim();

                            //start sql query
                            #region
                            string productionQuery = string.Empty;

                            DataTable dtProduction = new DataTable(); ;




                            SqlConnection con = new SqlConnection(CnnString);
                            SqlCommand cmd = new SqlCommand("USP_AllOEMProductionSheetTwentyTwo", con);
                            cmd.CommandType = CommandType.StoredProcedure;
                            con.Open();
                            cmd.Parameters.AddWithValue("@navembid", Navembcode);
                            cmd.Parameters.AddWithValue("@HSRP_StateID", HSRP_StateID);
                            cmd.Parameters.AddWithValue("@Dealerid", dealerid);
                            //cmd.Parameters.AddWithValue("@Affix_Id", strsubaffixid);
                            SqlDataAdapter da = new SqlDataAdapter(cmd);
                            // dtProduction = new DataTable();
                            da.Fill(dtProduction);
                            con.Close();




                            #endregion
                            //end sql query

                            #region
                            //DataTable dtProduction = Utils.Utils.GetDataTable(productionQuery, CnnString);
                            if (dtProduction.Rows.Count > 0)
                            {

                                findRecord = true;
                                string strRunningNo = string.Empty;
                                //string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) as maxSheetNo from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + NAVEMBID + "'";
                                string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                                string strCom = Utils.Utils.getScalarValue(strSel, CnnString);
                                //DataTable dtSheetNo = fillDataTable(strSel, CnnString);
                                if (strCom.Equals(0) || strCom.Length == 0)
                                {
                                    strRunningNo = "0000001";
                                }
                                else
                                {
                                    strRunningNo = string.Format("{0:0000000}", Convert.ToInt32(strCom) + 1);
                                }
                                string strRequeNo = "select (prefixtext+right('00000'+ convert(varchar,Lastreqno+1),5)) as Reqno from prefix_Requisition  where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                                ReqNum = Utils.Utils.getScalarValue(strRequeNo, CnnString);
                                string strPRFIX = "select PrefixText from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                                string strPRFIXCom = Utils.Utils.getScalarValue(strPRFIX, CnnString);
                                strProductionSheetNo = strPRFIXCom + strRunningNo;

                                string RTOLocationName = dtProduction.Rows[0]["RTOLocationName"].ToString();

                                #region
                               

                                html.Append("<div style='page-break-before: avoid;page-break-inside: avoid;page-break-after: always;'>" +
                                        "<table style='width:100%;border: 0px;'>" +
                                            "<tr style='border: 0px;'>" +
                                                "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>Report Generation Date:</b> " + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</div>" +
                                                "</td>" +

                                                 "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>ROSMERTA SAFETY SYSTEMS LIMITED</b> " + "</div>" +
                                                "</td>" +


                                                 "<td colspan='5' style='border: 0px;'>" +
                                                    "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Production Sheet No:</b> " + strRunningNo + " </div>" +
                                                "</td>" +

                                            "</tr>" +

                                              "<tr style='border: 0px;'>" +
                                                "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>State:</b> " + HSRPStateName + "</div>" +
                                                "</td>" +

                                                 "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>Production Sheet</b> " + "</div>" +
                                                "</td>" +


                                                 "<td colspan='5' style='border: 0px;'>" +
                                                    "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Dealer ID/Name:</b> " + dealerid + "/" + dealername + " </div>" +
                                                "</td>" +

                                            "</tr>" +

                                                "<tr style='border: 0px;'>" +
                                                "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>EC Location: </b> " + RTOLocationName + "</div>" +
                                                "</td>" +

                                                 "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='style='text-align:left'><b>Oem :</b> " + oemname + "</div>" +
                                                "</td>" +


                                                 "<td colspan='5' style='border: 0px;'>" +
                                                     "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Dealer Address:</b> " + Address + " </div>" +
                                                //"<div style='text-align:left'><b>Dealer Address:</b> " + Address + " </div>" +
                                                "</td>" +

                                            "</tr>" +


  "<tr>" +
                                                "<td style='text-align:center'>Sr. No.</td>" +
                                                  "<td>Vehicle No.</td>" +
                                                   "<td>Front Plate Size</td>" +
                                                "<td>Front Laser No.</td>" +

                                                "<td>Rear Plate Size</td>" +
                                                "<td>Rear Laser No.</td>" +

                                                "<td>H. S. Foil </td>" +
                                                "<td>Caution Sticker</td>" +
                                                "<td>Fuel Type</td>" +
                                                "<td >VT</td>" +
                                                "<td >VC</ td>" +

                                            "</tr>");

                                #endregion

                                #region
                                foreach (DataRow drProduction in dtProduction.Rows)
                                {
                                    string HsrprecordID = drProduction["hsrprecordID"].ToString().Trim();
                                    string SRNo = drProduction["SRNo"].ToString().Trim();
                                    string ORD = "";// drProduction["ORD"].ToString().Trim();
                                    string VC = drProduction["VehicleClass"].ToString().Trim();
                                    string VehicleNo = drProduction["VehicleRegNo"].ToString().Trim();
                                    string VT = drProduction["VehicleType"].ToString().Trim();
                                   // string ChassisNo = drProduction["ChassisNo"].ToString().Trim();
                                    //string EngineNo = drProduction["EngineNo"].ToString().Trim();
                                    string FuelType = drProduction["FuelType"].ToString().Trim();
                                    string FrontPSize = drProduction["FrontProductCode"].ToString().Trim();
                                    string FrontLaserNo = drProduction["HSRP_Front_LaserCode"].ToString().Trim();
                                    string RearPSize = drProduction["RearProductCode"].ToString().Trim();
                                    string RearLaserNo = drProduction["HSRP_Rear_LaserCode"].ToString().Trim();
                                    string Amount = drProduction["roundoff_netamount"].ToString().Trim();
                                    //string OrderStatus = drProduction["OrderStatus"].ToString().Trim();
                                    string stickerColor = drProduction["stickerColor"].ToString().Trim();

                                    string HotStampingFoilColour = drProduction["HotStampingFoilColour"].ToString().Trim();

                                    html.Append("<tr>" +
                                      "<td style='text-align:center'>" + SRNo + "</td>" +
                                          "<td style='text-align:center;font-size:20px;'>" + "<b>" + VehicleNo + "</b> </td>" +

                                           "<td>" + FrontPSize + "</td>" +
                                            "<td style='text-align:center;font-size:20px;'>" + "<b>" + FrontLaserNo + "</b> </td>" +
                                              "<td>" + RearPSize + "</td>" +
                                            "<td style='text-align:center;font-size:20px;'>" + "<b>" + RearLaserNo + "</b> </td>" +

                                             "<td>" + HotStampingFoilColour + "</td>" +
                                              "<td>" + stickerColor + "</td>" +
                                               "<td>" + FuelType + "</td>" +
                                               "<td>" + VT + "</td>" +
                                               "<td>" + VC + "</td>" +
                                             "</tr>");


                                    try
                                    {
                                        //start updating hsrprecords 
                                        string sqlUpdateHSRPRecords = "update hsrprecords set sendtoProductionStatus='Y', NAVPDFFlag='1', NewPdfRunningNo='" + strProductionSheetNo + "', Requisitionsheetno='" + ReqNum + "',  " +
                                              "PdfDownloadDate=GetDate(), pdfFileName='" + fileName + "', PDFDownloadUserID='1' where hsrprecordID='" + HsrprecordID + "' ";

                                        Utils.Utils.ExecNonQuery(sqlUpdateHSRPRecords, CnnString);   // uncomment after testing
                                    }
                                    catch (Exception ev)
                                    {
                                        Label1.Text = "hsrprecords update error: " + ev.Message;
                                    }
                                    //end 

                                }
                                #endregion
                                html.Append("</table>");

                                html.Append("</div>");

                                try
                                {
                                    string StrSqlUpdateECQuery = "update EmbossingCentersNew set NewProductionSheetRunningNo='" + strProductionSheetNo + "' " +
                                     "where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                                    Utils.Utils.ExecNonQuery(StrSqlUpdateECQuery, CnnString); // uncomment after testing
                                }
                                catch (Exception ev)
                                {
                                    Label1.Text = "EmbossingCentersNew update error: " + ev.Message;
                                }
                            }

                            #endregion

                        }
                    }// close oemDealerQuery
                    #endregion

                    #region "Req Generate"
                    string strComNew = string.Empty;
                    string strReqNumber = string.Empty;
                    string strReqNo = string.Empty;
                    string strSqlQuery1 = "select CompanyName from hsrpstate where hsrp_stateid='" + HSRP_StateID + "'";
                    strCompanyName = Utils.Utils.getScalarValue(strSqlQuery1, CnnString);

                    string strEMBName = " select EmbCenterName from EmbossingCentersNew where  State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                    strComNew = Utils.Utils.getScalarValue(strEMBName, CnnString);
                    if (strComNew != string.Empty)
                    {
                        strReqNo = "select (prefixtext+right('00000'+ convert(varchar,lastno+1),5)) as Reqno from prefix_Requisition  where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                        strReqNumber = Utils.Utils.getScalarValue(strReqNo, CnnString);

                        SQLString = "Exec [laserreqSlip1DashBoard]  '" + HSRP_StateID + "','" + Navembcode + "' ,  '" + ReqNum + "'";
                        DataTable dtResult = Utils.Utils.GetDataTable(SQLString, CnnString);
                        string strQuery = string.Empty;
                        string strRtoLocationName = string.Empty;
                        int Itotal = 0;

                        html.Append("<div style='width:100%;height:100%;'>" +
                                            "<table style='width:100%'>" +

                                                "<tr>" +
                                                    "<td colspan='12'>" +
                                                        "<div style='text-align:center;padding:8px;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + strCompanyName + "</b>" +

                                                            "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                             "<b style='font-size:20px;'>" + "MATERIAL REQUSITION NOTE" + "</b>" +

                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +

                                                "<tr>" +
                                                    "<td colspan='12'>" +
                                                        "<div style='text-align:center;padding:8px;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Production Sheet Date :" + DateTime.Now.ToString("dd-MM-yyyy") + "</b>" +
                                                            "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                             "<b style='font-size:20px;'>" + "" + "</b>" +
                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +
                                                "<tr>" +
                                                    "<td colspan='6'>" +
                                                        "<div style='text-align:left;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>REQ.NO:-</b>" +
                                                            "" + strReqNumber + "" +
                                                        "</div>" +
                                                    "</td>" +
                                                    "<td colspan='6'>" +
                                                        "<div style='text-align:left;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>Embossing Center:</b>" +
                                                            "" + strComNew + "" +
                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +
                                                "<tr>" +
                                                    "<td colspan='12'>" +
                                                        "<div style='text-align:left;padding:2px;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "State:" + HSRPStateName + "</b>" +
                                                            "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                             "<b style='font-size:20px;'>" + "" + "</b>" +
                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +
                                                 "<tr>" +
                                                    "<td colspan='12'>" +
                                                        "<div style='text-align:left;padding:8px;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + " " + "</b>" +
                                                            "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                             "<b style='font-size:20px;'>" + "" + "</b>" +
                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +

                                                "<tr>" +
                                                    "<td colspan='1' style='text-align:center'>SR.N.</td>" +
                                                    "<td colspan='3'>Product Size</td>" +
                                                    "<td colspan='1'>Laser Count</td>" +
                                                    "<td colspan='1'>Start Laser No</td>" +
                                                    "<td colspan='1'>End Laser No</td>" +

                                                "</tr>");


                        if (dtResult.Rows.Count > 0)
                        {
                            for (int i = 0; i < dtResult.Rows.Count; i++)
                            {
                                string ID = dtResult.Rows[i]["ID"].ToString();
                                string productcode = dtResult.Rows[i]["productcode"].ToString();
                                string LaserCount = dtResult.Rows[i]["LaserCount"].ToString();
                                Itotal = Convert.ToInt32(dtResult.Rows[i]["Total"].ToString());
                                string BeginLaser = dtResult.Rows[i]["BeginLaser"].ToString();
                                string EndLaser = dtResult.Rows[i]["EndLaser"].ToString();




                                html.Append("<tr>" +
                                   "<td colspan='1' style='text-align:left'>" + ID + "</td>" +
                                   "<td colspan='3'>" + productcode + "</td>" +

                                   "<td colspan='1'>" + LaserCount + "</td>" +
                                   "<td colspan='1'>" + BeginLaser + "</td>" +
                                   "<td colspan='1'>" + EndLaser + "</td>" +

                               "</tr>");
                            }
                        }
                        html.Append("<tr>" +
                                 "<td colspan='1' style='text-align:center' > " + "<b>Grand Total:</b>" + "</td>" +
                                 "<td colspan='3'>" + "" + "</td>" +

                                 "<td colspan='1'>" + Itotal + "</td>" +
                                 "<td colspan='1'>" + " " + "</td>" +
                                 "<td colspan='1'>" + " " + "</td>" +

                             "</tr>");




                        html.Append("<tr>" +
                         "<td colspan='2' > " + "<b>REQUESTED BY </b>" + "</td>" +
                         "<td colspan='2'>" + "<b>AUTHORIZED BY </b>" + "</td>" +

                         "<td colspan='2'>" + "<b>ISSUED BY </b>" + "</td>" +
                         "<td colspan='2'>" + "<b>RECEIVED BY</b>" + "</td>" +


                     "</tr>");

                        html.Append("<tr>" +
                                                  "<td colspan='12'>" +
                                                      "<div style='text-align:left;padding:2px;'>" +
                                                          "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Name" + "</b>" +

                                                      "</div>" +
                                                  "</td>" +
                                              "</tr>");

                        html.Append("<tr>" +
                                                 "<td colspan='12'>" +
                                                     "<div style='text-align:left;padding:2px;'>" +
                                                         "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Designation" + "</b>" +

                                                     "</div>" +
                                                 "</td>" +
                                             "</tr>");

                        html.Append("<tr>" +
                                               "<td colspan='12'>" +
                                                   "<div style='text-align:right;padding:8px;'>" +
                                                       "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Download By: Admin " + "</b>" +

                                                   "</div>" +
                                               "</td>" +
                                           "</tr>");

                        html.Append("<tr>" +
                                              "<td colspan='12'>" +
                                                  "<div style='text-align:right;padding:8px;'>" +
                                                      "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Sheet Generated By :" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</b>" +

                                                  "</div>" +
                                              "</td>" +
                                          "</tr>");






                        html.Append("</table>");

                        html.Append("</div>");

                        try
                        {
                            //start updating hsrprecords 
                            string Query = "update prefix_Requisition set lastno=lastno+1,Lastreqno=Lastreqno+1 where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                            Utils.Utils.ExecNonQuery(Query, CnnString);
                        }
                        catch (Exception ev)
                        {
                            Label1.Text = "prefix Requisition update error: " + ev.Message;
                        }
                    }


                    #endregion

                    /*
                     * Close body & HTMl Tag
                     */
                    html.Append("</body>" +
                        "</html>");


                    if (findRecord)
                    {

                        string strSaveSQlQuery = "insert into ProductionSheetAutoGenerated_List (HSRP_StateID, State_Code, RTO_ID, Emb_Center_Id, FileName, Productiondate, ProductionTime, OrderNo) values " +
                                    "('" + HSRP_StateID + "', '" + HSRPStateShortName + "', '', '" + Navembcode + "', '" + fileName + "', '" + DateTime.Now.ToString("yyyy-MM-dd") + "', '" + DateTime.Now.ToString("HH:mm:ss") + "', '" + filePrefix + "') ";
                        Utils.Utils.ExecNonQuery(strSaveSQlQuery, CnnString); // uncomment after testing
                        Label1.Text = Label1.Text + "successfully created production sheet of State: " + HSRPStateShortName + ", , Emb_Center_Id: " + Navembcode + Environment.NewLine;

                        //lblLog.Text = lblLog.Text + Environment.NewLine + "Start Production Query at:" + DateTime.Now;
                        #region
                        try
                        {
                            if (!Directory.Exists(dir))
                            {


                                Directory.CreateDirectory(dir);

                            }

                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        try
                        {
                            var pdf = Pdf
                                    .From(html.ToString())
                                    .OfSize(PaperSize.A4)
                                    .WithTitle("Title")
                                    .WithoutOutline()
                                    .WithMargins(1.25.Centimeters())
                                     .Landscape()
                                    .Comressed()
                                    .Content();

                            FileStream readStream = File.Create(filePath);
                            BinaryWriter binaryWriter = new BinaryWriter(readStream);

                            // Write the binary data to the file
                            binaryWriter.Write(pdf);
                            binaryWriter.Close();
                            readStream.Close();
                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        #endregion
                    }

                    #endregion

                }//close foreach stateEcQuery
            }
        }

        protected void btnPCA_Click(object sender, EventArgs e)
        {
            if (ddlStateName.SelectedItem.Text == "--Select State--")
            {
                lblErrMess.Text = "Please select  State";
                return;
            }


            string sqlPrefixQuery = "select top 1 CONVERT(INT, isnull(OrderNo,'0')) + 1 as orderNo from ProductionSheetAutoGenerated_List order by id desc";
            DataTable dtPrefix = Utils.Utils.GetDataTable(sqlPrefixQuery, CnnString);

            if (dtPrefix.Rows.Count > 0)
            {
                filePrefix = dtPrefix.Rows[0]["orderNo"].ToString();
            }
            else
            {
                filePrefix = "1";
            }
            SheetGenerationPCAinterSTATE();
        }

        private void SheetGenerationPCAinterSTATE()
        {
            string stateECQuery = string.Empty;

            //string Navembid = Session["Navembid"].ToString();
            FillUserDetails();
            string ReqNum = string.Empty;

            if ((ddlStateName.SelectedValue == "20") || (ddlStateName.SelectedValue == "10"))
            {


                stateECQuery = "select distinct a.HSRP_StateID, (select HSRPStateName from hsrpstate c where c.hsrp_stateid=a.hsrp_stateid) as HSRPStateName,(select HSRPStateShortName from hsrpstate d where d.hsrp_stateid=a.hsrp_stateid) as HSRPStateShortName, navembcode from hsrprecords a with(nolock),rtolocation b where a.HSRP_StateID ='" + ddlStateName.SelectedValue + "' and   b.Navembcode not like '%CODO%'    and a.rtolocationid=b.rtolocationid and NewPdfRunningNo is null and erpassigndate is not null and OrderStatus='New Order' AND b.NAVEMBID='" + Navembid + "'   order by  a.HSRP_StateID";
            }
            else
            {
                stateECQuery = "select distinct a.HSRP_StateID, (select HSRPStateName from hsrpstate c where c.hsrp_stateid=a.hsrp_stateid) as HSRPStateName,(select HSRPStateShortName from hsrpstate d where d.hsrp_stateid=a.hsrp_stateid) as HSRPStateShortName, navembcode from hsrprecords a with(nolock),rtolocation b where   b.Navembcode not like '%CODO%'    and a.rtolocationid=b.rtolocationid and NewPdfRunningNo is null and erpassigndate is not null and OrderStatus='New Order' and   b.navembcode like  '%'+(select HSRPStateShortName from hsrpstate d where d.hsrp_stateid=a.hsrp_stateid)+'%'  and b.NAVEMBID='" + Navembid + "'    order by  a.HSRP_StateID";

            }


            DataTable dtSE = Utils.Utils.GetDataTable(stateECQuery, CnnString);
            if (dtSE.Rows.Count > 0)
            {
                foreach (DataRow dr in dtSE.Rows)
                {
                    #region

                    string HSRP_StateID = dr["HSRP_StateID"].ToString().Trim();
                    string HSRPStateName = dr["HSRPStateName"].ToString().Trim();
                    string HSRPStateShortName = dr["HSRPStateShortName"].ToString().Trim();
                    string Navembcode = dr["Navembcode"].ToString().Trim();

                    string dir = dirPath + DateTime.Now.ToString("yyyy-MM-dd") + "\\" + HSRPStateShortName + "\\";
                    //string dir = dirPath + "/" + DateTime.Now.ToString("yyyy-MM-dd") + "/" + HSRPStateShortName + "/";
                    // string fileName = filePrefix + "-" + Navembcode + ".pdf";
                    string fileName = "AllOEM" + "-" + filePrefix + "-" + Navembcode + ".pdf";
                    string filePath = dir + fileName;

                    //string folderpath = ConfigurationManager.AppSettings["InvoiceFolder"].ToString() + "/" + FinYear + "/" + oemid + "/" + HSRPStateID + "/";

                    StringBuilder html = new StringBuilder();

                    Boolean findRecord = false;
                    string strProductionSheetNo = string.Empty;

                    /*
                     *  Start body & HTMl Tag
                     */
                    #region
                    html.Append(
                        "<!DOCTYPE html>" +
                        "<html>" +
                        "<head>" +
                            "<meta charset='UTF-8'><title>Title</title>" +
                            "<style>" +
                                "@page {" +
                                    /* headers*/
                                    "@top-left {" +
                                        "content: 'Left header';" +
                                    "}" +
                                    "@top-right {" +
                                        "content: 'Right header';" +
                                    "}" +

                                    /* footers */
                                    "@bottom-left {" +
                                        "content: 'Lorem ipsum';" +
                                    "} " +
                                    "@bottom-right {" +
                                        "content: 'Page ' counter(page) ' of ' counter(pages);" +
                                    "}" +
                                    "@bottom-center  {" +
                                        "content:element(footer);" +
                                    "}" +
                                "}" +
                                 "#footer {" +
                                    "position: running(footer);" +
                                "}" +
                                "table {" +
                                  "border-collapse: collapse;" +
                                "}" +

                                "table, th, td {" +
                                    "border: 1px solid black;" +
                                    "text-align: left;" +
                                    "vertical-align: top;" +
                                    "padding:5px;" +
                                "}" +
                            "</style>" +
                        "</head>" +
                        "<body>");
                    #endregion

                    string oemDealerQuery = string.Empty;



                    oemDealerQuery = "select distinct om.oemid as oemid, om.name as oemname, dm.dealerid, dm.dealername, dm.dealercode, dm.Address, dm.HSRP_StateID, " +
                        "dm.RTOLocationID from oemmaster om " +
                        "left join dealermaster dm on dm.oemid = om.oemid where dm.HSRP_StateID =" + HSRP_StateID + " and " +
                        "dm.RTOLocationID in (select RTOLocationID from rtolocation where Navembcode='" + Navembcode + "' ) and " +
                        "dm.dealerid in (select distinct dealerid from hsrprecords  with(nolock) where NewPdfRunningNo is null and erpassigndate is not null and OrderStatus='New Order') and Om.OEMID   in('1231')";

                    //oemDealerQuery = "select distinct om.oemid as oemid, om.name as oemname, dm.dealerid, dm.dealername, dm.dealercode, dm.Address, h.HSRP_StateID, h.RTOLocationID from oemmaster om  join dealermaster dm on dm.oemid = om.oemid  join hsrprecords h with(nolock) on h.dealerid = dm.dealerid where h.RTOLocationID in (select RTOLocationID from rtolocation where Navembcode = '" + Navembcode + "' )  and NewPdfRunningNo is null and erpassigndate is not null and OrderStatus = 'New Order'  and Om.OEMID = 1231";
                    #region
                    DataTable dtOD = Utils.Utils.GetDataTable(oemDealerQuery, CnnString);
                    if (dtOD.Rows.Count > 0)
                    {
                        foreach (DataRow drOD in dtOD.Rows)
                        {

                            string oemid = drOD["oemid"].ToString().Trim();
                            string dealerid = drOD["dealerid"].ToString().Trim();
                            string oemname = drOD["oemname"].ToString().Trim();
                            string dealername = drOD["dealername"].ToString().Trim();
                            string Address = drOD["Address"].ToString().Trim();

                            //start sql query
                            #region
                            string productionQuery = string.Empty;

                            DataTable dtProduction = new DataTable(); ;




                            SqlConnection con = new SqlConnection(CnnString);
                            SqlCommand cmd = new SqlCommand("USP_AllOEMProductionSheet", con);
                            cmd.CommandType = CommandType.StoredProcedure;
                            con.Open();
                            cmd.Parameters.AddWithValue("@navembid", Navembcode);
                            cmd.Parameters.AddWithValue("@HSRP_StateID", HSRP_StateID);
                            cmd.Parameters.AddWithValue("@Dealerid", dealerid);
                            //cmd.Parameters.AddWithValue("@Affix_Id", strsubaffixid);
                            SqlDataAdapter da = new SqlDataAdapter(cmd);
                            // dtProduction = new DataTable();
                            da.Fill(dtProduction);
                            con.Close();




                            #endregion
                            //end sql query

                            #region
                            //DataTable dtProduction = Utils.Utils.GetDataTable(productionQuery, CnnString);
                            if (dtProduction.Rows.Count > 0)
                            {

                                findRecord = true;
                                string strRunningNo = string.Empty;
                                //string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) as maxSheetNo from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + NAVEMBID + "'";
                                //string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                                string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) from EmbossingCentersNew where  Emb_Center_Id='" + Navembcode + "'";
                                string strCom = Utils.Utils.getScalarValue(strSel, CnnString);
                                //DataTable dtSheetNo = fillDataTable(strSel, CnnString);
                                if (strCom.Equals(0) || strCom.Length == 0)
                                {
                                    strRunningNo = "0000001";
                                }
                                else
                                {
                                    strRunningNo = string.Format("{0:0000000}", Convert.ToInt32(strCom) + 1);
                                }
                                string strRequeNo = "select (prefixtext+right('00000'+ convert(varchar,Lastreqno+1),5)) as Reqno from prefix_Requisition  where  prefixfor='Req No'";
                                ReqNum = Utils.Utils.getScalarValue(strRequeNo, CnnString);
                                string strPRFIX = "select PrefixText from EmbossingCentersNew where Emb_Center_Id='" + Navembcode + "'";
                                string strPRFIXCom = Utils.Utils.getScalarValue(strPRFIX, CnnString);
                                strProductionSheetNo = strPRFIXCom + strRunningNo;

                                string RTOLocationName = dtProduction.Rows[0]["RTOLocationName"].ToString();

                                #region
                                html.Append("<div style='page-break-before: avoid;page-break-inside: avoid;page-break-after: always;'>" +
                                        "<table style='width:100%;border: 0px;'>" +
                                            "<tr style='border: 0px;'>" +
                                                "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>Report Generation Date:</b> " + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</div>" +
                                                "</td>" +

                                                 "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>ROSMERTA SAFETY SYSTEMS LIMITED</b> " + "</div>" +
                                                "</td>" +


                                                 "<td colspan='5' style='border: 0px;'>" +
                                                    "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Production Sheet No:</b> " + strRunningNo + " </div>" +
                                                "</td>" +

                                            "</tr>" +

                                              "<tr style='border: 0px;'>" +
                                                "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>State:</b> " + HSRPStateName + "</div>" +
                                                "</td>" +

                                                 "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>Production Sheet</b> " + "</div>" +
                                                "</td>" +


                                                 "<td colspan='5' style='border: 0px;'>" +
                                                    "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Dealer ID/Name:</b> " + dealerid + "/" + dealername + " </div>" +
                                                "</td>" +

                                            "</tr>" +

                                                "<tr style='border: 0px;'>" +
                                                "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>EC Location: </b> " + RTOLocationName + "</div>" +
                                                "</td>" +

                                                 "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='style='text-align:left'><b>Oem :</b> " + oemname + "</div>" +
                                                "</td>" +


                                                 "<td colspan='5' style='border: 0px;'>" +
                                                     "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Dealer Address:</b> " + Address + " </div>" +
                                                //"<div style='text-align:left'><b>Dealer Address:</b> " + Address + " </div>" +
                                                "</td>" +

                                            "</tr>" +


  "<tr>" +
                                                "<td style='text-align:center'>Sr. No.</td>" +
                                                  "<td>Vehicle No.</td>" +
                                                   "<td>Front Plate Size</td>" +
                                                "<td>Front Laser No.</td>" +

                                                "<td>Rear Plate Size</td>" +
                                                "<td>Rear Laser No.</td>" +

                                                "<td>H. S. Foil </td>" +
                                                "<td>Caution Sticker</td>" +
                                                "<td>Fuel Type</td>" +
                                                "<td >VT</td>" +
                                                "<td >VC</ td>" +

                                            "</tr>");

                                #endregion

                                #region
                                foreach (DataRow drProduction in dtProduction.Rows)
                                {
                                    string HsrprecordID = drProduction["hsrprecordID"].ToString().Trim();
                                    string SRNo = drProduction["SRNo"].ToString().Trim();
                                    string ORD = "";// drProduction["ORD"].ToString().Trim();
                                    string VC = drProduction["VehicleClass"].ToString().Trim();
                                    string VehicleNo = drProduction["VehicleRegNo"].ToString().Trim();
                                    string VT = drProduction["VehicleType"].ToString().Trim();
                                   // string ChassisNo = drProduction["ChassisNo"].ToString().Trim();
                                    //string EngineNo = drProduction["EngineNo"].ToString().Trim();
                                    string FuelType = drProduction["FuelType"].ToString().Trim();
                                    string FrontPSize = drProduction["FrontProductCode"].ToString().Trim();
                                    string FrontLaserNo = drProduction["HSRP_Front_LaserCode"].ToString().Trim();
                                    string RearPSize = drProduction["RearProductCode"].ToString().Trim();
                                    string RearLaserNo = drProduction["HSRP_Rear_LaserCode"].ToString().Trim();
                                    string Amount = drProduction["roundoff_netamount"].ToString().Trim();
                                    string OrderStatus = drProduction["OrderStatus"].ToString().Trim();
                                    string stickerColor = drProduction["stickerColor"].ToString().Trim();
                                    string HotStampingFoilColour = drProduction["HotStampingFoilColour"].ToString().Trim();

                                    html.Append("<tr>" +
                                       "<td style='text-align:center'>" + SRNo + "</td>" +
                                           "<td style='text-align:center;font-size:20px;'>" + "<b>" + VehicleNo + "</b> </td>" +

                                            "<td>" + FrontPSize + "</td>" +
                                             "<td style='text-align:center;font-size:20px;'>" + "<b>" + FrontLaserNo + "</b> </td>" +
                                               "<td>" + RearPSize + "</td>" +
                                             "<td style='text-align:center;font-size:20px;'>" + "<b>" + RearLaserNo + "</b> </td>" +

                                              "<td>" + HotStampingFoilColour + "</td>" +
                                               "<td>" + stickerColor + "</td>" +
                                                "<td>" + FuelType + "</td>" +
                                                "<td>" + VT + "</td>" +
                                                "<td>" + VC + "</td>" +
                                              "</tr>"); ;

                                    try
                                    {
                                        //start updating hsrprecords 
                                        string sqlUpdateHSRPRecords = "update hsrprecords set sendtoProductionStatus='Y', NAVPDFFlag='1', NewPdfRunningNo='" + strProductionSheetNo + "', Requisitionsheetno='" + ReqNum + "',  " +
                                              "PdfDownloadDate=GetDate(), pdfFileName='" + fileName + "', PDFDownloadUserID='1' where hsrprecordID='" + HsrprecordID + "' ";

                                        Utils.Utils.ExecNonQuery(sqlUpdateHSRPRecords, CnnString);   // uncomment after testing
                                    }
                                    catch (Exception ev)
                                    {
                                        Label1.Text = "hsrprecords update error: " + ev.Message;
                                    }
                                    //end 

                                }
                                #endregion
                                html.Append("</table>");

                                html.Append("</div>");

                                try
                                {
                                    string StrSqlUpdateECQuery = "update EmbossingCentersNew set NewProductionSheetRunningNo='" + strProductionSheetNo + "' " +
                                     "where  Emb_Center_Id='" + Navembcode + "'";
                                    Utils.Utils.ExecNonQuery(StrSqlUpdateECQuery, CnnString); // uncomment after testing
                                }
                                catch (Exception ev)
                                {
                                    Label1.Text = "EmbossingCentersNew update error: " + ev.Message;
                                }
                            }

                            #endregion

                        }
                    }// close oemDealerQuery
                    #endregion

                    //#region "Req Generate"
                    //string strComNew = string.Empty;
                    //string strReqNumber = string.Empty;
                    //string strReqNo = string.Empty;
                    //string strSqlQuery1 = "select CompanyName from hsrpstate where hsrp_stateid='" + HSRP_StateID + "'";
                    //strCompanyName = Utils.Utils.getScalarValue(strSqlQuery1, CnnString);

                    //string strEMBName = " select EmbCenterName from EmbossingCentersNew where   Emb_Center_Id='" + Navembcode + "'";
                    //strComNew = Utils.Utils.getScalarValue(strEMBName, CnnString);
                    //if (strComNew != string.Empty)
                    //{
                    //    strReqNo = "select (prefixtext+right('00000'+ convert(varchar,lastno+1),5)) as Reqno from prefix_Requisition  where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                    //    strReqNumber = Utils.Utils.getScalarValue(strReqNo, CnnString);

                    //    SQLString = "Exec [laserreqSlip1DashBoard]  '" + HSRP_StateID + "','" + Navembcode + "' ,  '" + ReqNum + "'";
                    //    DataTable dtResult = Utils.Utils.GetDataTable(SQLString, CnnString);
                    //    string strQuery = string.Empty;
                    //    string strRtoLocationName = string.Empty;
                    //    int Itotal = 0;

                    //    html.Append("<div style='width:100%;height:100%;'>" +
                    //                        "<table style='width:100%'>" +

                    //                            "<tr>" +
                    //                                "<td colspan='12'>" +
                    //                                    "<div style='text-align:center;padding:8px;'>" +
                    //                                        "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + strCompanyName + "</b>" +

                    //                                        "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                    //                                         "<b style='font-size:20px;'>" + "MATERIAL REQUSITION NOTE" + "</b>" +

                    //                                    "</div>" +
                    //                                "</td>" +
                    //                            "</tr>" +

                    //                            "<tr>" +
                    //                                "<td colspan='12'>" +
                    //                                    "<div style='text-align:center;padding:8px;'>" +
                    //                                        "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Production Sheet Date :" + DateTime.Now.ToString("dd-MM-yyyy") + "</b>" +
                    //                                        "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                    //                                         "<b style='font-size:20px;'>" + "" + "</b>" +
                    //                                    "</div>" +
                    //                                "</td>" +
                    //                            "</tr>" +
                    //                            "<tr>" +
                    //                                "<td colspan='6'>" +
                    //                                    "<div style='text-align:left;'>" +
                    //                                        "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>REQ.NO:-</b>" +
                    //                                        "" + strReqNumber + "" +
                    //                                    "</div>" +
                    //                                "</td>" +
                    //                                "<td colspan='6'>" +
                    //                                    "<div style='text-align:left;'>" +
                    //                                        "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>Embossing Center:</b>" +
                    //                                        "" + strComNew + "" +
                    //                                    "</div>" +
                    //                                "</td>" +
                    //                            "</tr>" +
                    //                            "<tr>" +
                    //                                "<td colspan='12'>" +
                    //                                    "<div style='text-align:left;padding:2px;'>" +
                    //                                        "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "State:" + HSRPStateName + "</b>" +
                    //                                        "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                    //                                         "<b style='font-size:20px;'>" + "" + "</b>" +
                    //                                    "</div>" +
                    //                                "</td>" +
                    //                            "</tr>" +
                    //                             "<tr>" +
                    //                                "<td colspan='12'>" +
                    //                                    "<div style='text-align:left;padding:8px;'>" +
                    //                                        "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + " " + "</b>" +
                    //                                        "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                    //                                         "<b style='font-size:20px;'>" + "" + "</b>" +
                    //                                    "</div>" +
                    //                                "</td>" +
                    //                            "</tr>" +

                    //                            "<tr>" +
                    //                                "<td colspan='1' style='text-align:center'>SR.N.</td>" +
                    //                                "<td colspan='3'>Product Size</td>" +
                    //                                "<td colspan='1'>Laser Count</td>" +
                    //                                "<td colspan='1'>Start Laser No</td>" +
                    //                                "<td colspan='1'>End Laser No</td>" +

                    //                            "</tr>");


                    //    if (dtResult.Rows.Count > 0)
                    //    {
                    //        for (int i = 0; i < dtResult.Rows.Count; i++)
                    //        {
                    //            string ID = dtResult.Rows[i]["ID"].ToString();
                    //            string productcode = dtResult.Rows[i]["productcode"].ToString();
                    //            string LaserCount = dtResult.Rows[i]["LaserCount"].ToString();
                    //            Itotal = Convert.ToInt32(dtResult.Rows[i]["Total"].ToString());
                    //            string BeginLaser = dtResult.Rows[i]["BeginLaser"].ToString();
                    //            string EndLaser = dtResult.Rows[i]["EndLaser"].ToString();




                    //            html.Append("<tr>" +
                    //               "<td colspan='1' style='text-align:left'>" + ID + "</td>" +
                    //               "<td colspan='3'>" + productcode + "</td>" +

                    //               "<td colspan='1'>" + LaserCount + "</td>" +
                    //               "<td colspan='1'>" + BeginLaser + "</td>" +
                    //               "<td colspan='1'>" + EndLaser + "</td>" +

                    //           "</tr>");
                    //        }
                    //    }
                    //    html.Append("<tr>" +
                    //             "<td colspan='1' style='text-align:center' > " + "<b>Grand Total:</b>" + "</td>" +
                    //             "<td colspan='3'>" + "" + "</td>" +

                    //             "<td colspan='1'>" + Itotal + "</td>" +
                    //             "<td colspan='1'>" + " " + "</td>" +
                    //             "<td colspan='1'>" + " " + "</td>" +

                    //         "</tr>");




                    //    html.Append("<tr>" +
                    //     "<td colspan='2' > " + "<b>REQUESTED BY </b>" + "</td>" +
                    //     "<td colspan='2'>" + "<b>AUTHORIZED BY </b>" + "</td>" +

                    //     "<td colspan='2'>" + "<b>ISSUED BY </b>" + "</td>" +
                    //     "<td colspan='2'>" + "<b>RECEIVED BY</b>" + "</td>" +


                    // "</tr>");

                    //    html.Append("<tr>" +
                    //                              "<td colspan='12'>" +
                    //                                  "<div style='text-align:left;padding:2px;'>" +
                    //                                      "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Name" + "</b>" +

                    //                                  "</div>" +
                    //                              "</td>" +
                    //                          "</tr>");

                    //    html.Append("<tr>" +
                    //                             "<td colspan='12'>" +
                    //                                 "<div style='text-align:left;padding:2px;'>" +
                    //                                     "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Designation" + "</b>" +

                    //                                 "</div>" +
                    //                             "</td>" +
                    //                         "</tr>");

                    //    html.Append("<tr>" +
                    //                           "<td colspan='12'>" +
                    //                               "<div style='text-align:right;padding:8px;'>" +
                    //                                   "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Download By: Admin " + "</b>" +

                    //                               "</div>" +
                    //                           "</td>" +
                    //                       "</tr>");

                    //    html.Append("<tr>" +
                    //                          "<td colspan='12'>" +
                    //                              "<div style='text-align:right;padding:8px;'>" +
                    //                                  "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Sheet Generated By :" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</b>" +

                    //                              "</div>" +
                    //                          "</td>" +
                    //                      "</tr>");






                    //    html.Append("</table>");

                    //    html.Append("</div>");

                    //    try
                    //    {
                    //        //start updating hsrprecords 
                    //        string Query = "update prefix_Requisition set lastno=lastno+1,Lastreqno=Lastreqno+1 where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                    //        Utils.Utils.ExecNonQuery(Query, CnnString);
                    //    }
                    //    catch (Exception ev)
                    //    {
                    //        Label1.Text = "prefix Requisition update error: " + ev.Message;
                    //    }
                    //}


                    //#endregion

                    /*
                     * Close body & HTMl Tag
                     */
                    html.Append("</body>" +
                        "</html>");


                    if (findRecord)
                    {

                        string strSaveSQlQuery = "insert into ProductionSheetAutoGenerated_List (HSRP_StateID, State_Code, RTO_ID, Emb_Center_Id, FileName, Productiondate, ProductionTime, OrderNo) values " +
                                    "('" + HSRP_StateID + "', '" + HSRPStateShortName + "', '', '" + Navembcode + "', '" + fileName + "', '" + DateTime.Now.ToString("yyyy-MM-dd") + "', '" + DateTime.Now.ToString("HH:mm:ss") + "', '" + filePrefix + "') ";
                        Utils.Utils.ExecNonQuery(strSaveSQlQuery, CnnString); // uncomment after testing
                        Label1.Text = Label1.Text + "successfully created production sheet of State: " + HSRPStateShortName + ", , Emb_Center_Id: " + Navembcode + Environment.NewLine;

                        //lblLog.Text = lblLog.Text + Environment.NewLine + "Start Production Query at:" + DateTime.Now;
                        #region
                        try
                        {
                            if (!Directory.Exists(dir))
                            {


                                Directory.CreateDirectory(dir);

                            }

                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        try
                        {
                            var pdf = Pdf
                                    .From(html.ToString())
                                    .OfSize(PaperSize.A4)
                                    .WithTitle("Title")
                                    .WithoutOutline()
                                    .WithMargins(1.25.Centimeters())
                                     .Landscape()
                                    .Comressed()
                                    .Content();

                            FileStream readStream = File.Create(filePath);
                            BinaryWriter binaryWriter = new BinaryWriter(readStream);

                            // Write the binary data to the file
                            binaryWriter.Write(pdf);
                            binaryWriter.Close();
                            readStream.Close();
                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        #endregion
                    }

                    #endregion

                }//close foreach stateEcQuery
            }
        }

            private void SheetGenerationOLAHubWise()
        {
            string stateECQuery = string.Empty;
            string ReqNum = string.Empty;



            string stateECShortName = "select distinct HSRP_StateID, 'EC'+HSRPStateShortName as NewHSRPStateShortName, HSRPStateShortName   from  HSRPState  where  HSRP_STateId='" + ddlStateName.SelectedValue + "'";

            DataTable dtECName = Utils.Utils.GetDataTable(stateECShortName, CnnString);

            string ShortECname = dtECName.Rows[0]["NewHSRPStateShortName"].ToString();
            FillUserDetails();
         
            if ((ddlStateName.SelectedValue == "20") || (ddlStateName.SelectedValue == "10") || (ddlStateName.SelectedValue == "17"))
            {

                stateECQuery = "select distinct a.HSRP_StateID, (select HSRPStateName from hsrpstate c where c.hsrp_stateid=a.hsrp_stateid) as HSRPStateName,(select HSRPStateShortName from hsrpstate d with(nolock) where d.hsrp_stateid='" + ddlStateName.SelectedValue + "') as HSRPStateShortName,   a.Navembid AS  navembcode from hsrprecords a with(nolock) join OlaHubsMaster  d on a.dealerid=d.HSRPDealerID  where  a.NAVEMBID='" + Navembid + "' and     NewPdfRunningNo is null and erpassigndate is not null    and OrderStatus='New Order' and    Navembid not like '%CODO%'  order by  a.HSRP_StateID";
            }
            else
            {
               
                stateECQuery = "select distinct a.HSRP_StateID, (select HSRPStateName from hsrpstate c where c.hsrp_stateid=a.hsrp_stateid) as HSRPStateName,(select HSRPStateShortName from hsrpstate d with(nolock) where d.hsrp_stateid='" + ddlStateName.SelectedValue + "') as HSRPStateShortName,   a.Navembid AS  navembcode from hsrprecords a with(nolock) join OlaHubsMaster  d on a.dealerid=d.HSRPDealerID  where  a.NAVEMBID='" + Navembid + "' and   left(a.Navembid,4) = '" + ShortECname + "'  and  NewPdfRunningNo is null and erpassigndate is not null    and OrderStatus='New Order' and    Navembid not like '%CODO%'  order by  a.HSRP_StateID";

            }

            //stateECQuery = "select distinct a.HSRP_StateID, (select HSRPStateName from hsrpstate c where c.hsrp_stateid=a.hsrp_stateid) as HSRPStateName,(select HSRPStateShortName from hsrpstate d with(nolock) where d.hsrp_stateid='" + ddlStateName.SelectedValue + "') as HSRPStateShortName,   a.Navembid AS  navembcode from hsrprecords a with(nolock) join OlaHubsMaster  d on a.dealerid=d.HSRPDealerID  where   left(a.Navembid,4) = '" + ShortECname + "'  and  NewPdfRunningNo is null and erpassigndate is not null    and OrderStatus='New Order' and    Navembid not like '%CODO%'  order by  a.HSRP_StateID";
            DataTable dtSE = Utils.Utils.GetDataTable(stateECQuery, CnnString);


            if (dtSE.Rows.Count > 0)
            {
                foreach (DataRow dr in dtSE.Rows)
                {
                    #region

                    string HSRP_StateID = dr["HSRP_StateID"].ToString().Trim();
                    string HSRPStateName = dr["HSRPStateName"].ToString().Trim();
                    string HSRPStateShortName = dr["HSRPStateShortName"].ToString().Trim();
                    //string RTOLocationID = dr["RTOLocationID"].ToString().Trim();
                    //string RTOLocationName = dr["RTOLocationName"].ToString().Trim();
                    //string NAVEMBID = dr["NAVEMBID"].ToString().Trim();
                    string Navembcode = dr["Navembcode"].ToString().Trim();

                    string dir = dirPath + DateTime.Now.ToString("yyyy-MM-dd") + "\\" + HSRPStateShortName + "\\";

                    string fileName = string.Empty;
                    string filePath = string.Empty;
                    // string fileName = filePrefix + "-" + Navembcode + ".pdf";


                    StringBuilder html = new StringBuilder();

                    Boolean findRecord = false;
                    string strProductionSheetNo = string.Empty;




                    /*
                    *  Start body & HTMl Tag
                    */
                    #region
                    html.Append(
                        "<!DOCTYPE html>" +
                        "<html>" +
                        "<head>" +
                            "<meta charset='UTF-8'><title>Title</title>" +
                            "<style>" +
                                "@page {" +
                                    /* headers*/
                                    "@top-left {" +
                                        "content: 'Left header';" +
                                    "}" +
                                    "@top-right {" +
                                        "content: 'Right header';" +
                                    "}" +

                                    /* footers */
                                    "@bottom-left {" +
                                        "content: 'Lorem ipsum';" +
                                    "} " +
                                    "@bottom-right {" +
                                        "content: 'Page ' counter(page) ' of ' counter(pages);" +
                                    "}" +
                                    "@bottom-center  {" +
                                        "content:element(footer);" +
                                    "}" +
                                "}" +
                                 "#footer {" +
                                    "position: running(footer);" +
                                "}" +
                                "table {" +
                                  "border-collapse: collapse;" +
                                "}" +

                                "table, th, td {" +
                                    "border: 1px solid black;" +
                                    "text-align: left;" +
                                    "vertical-align: top;" +
                                    "padding:5px;" +
                                "}" +
                            "</style>" +
                        "</head>" +
                        "<body>");
                    #endregion

                   
                    fileName = "OLA" + "-" + filePrefix + "-" + Navembcode + ".pdf";
                    filePath = dir + fileName;

                    string oemDealerQuery = string.Empty;


                    #region
                    html.Append(
                        "<!DOCTYPE html>" +
                        "<html>" +
                        "<head>" +
                            "<meta charset='UTF-8'><title>Title</title>" +
                            "<style>" +
                                "@page {" +
                                    /* headers*/
                                    "@top-left {" +
                                        "content: 'Left header';" +
                                    "}" +
                                    "@top-right {" +
                                        "content: 'Right header';" +
                                    "}" +

                                    /* footers */
                                    "@bottom-left {" +
                                        "content: 'Lorem ipsum';" +
                                    "} " +
                                    "@bottom-right {" +
                                        "content: 'Page ' counter(page) ' of ' counter(pages);" +
                                    "}" +
                                    "@bottom-center  {" +
                                        "content:element(footer);" +
                                    "}" +
                                "}" +
                                 "#footer {" +
                                    "position: running(footer);" +
                                "}" +
                                "table {" +
                                  "border-collapse: collapse;" +
                                "}" +

                                "table, th, td {" +
                                    "border: 1px solid black;" +
                                    "text-align: left;" +
                                    "vertical-align: top;" +
                                    "padding:5px;" +
                                "}" +
                            "</style>" +
                        "</head>" +
                        "<body>");
                    #endregion

                    //string oemDealerQuery = string.Empty;



                    oemDealerQuery = "select distinct om.oemid as oemid, om.name as oemname, dm.dealerid, dm.dealername, dm.dealercode, d.[OLA HUB Address]+' '+ d.City + ' '+Convert(Varchar(10),[Pin Code] ) as Address, dm.HSRP_StateID, " +
                        "dm.RTOLocationID from oemmaster om " +
                        "left join dealermaster dm on dm.oemid = om.oemid JOIN  OlaHubsMaster d ON dm.dealerid=d.HSRPDealerID      where dm.HSRP_StateID =" + HSRP_StateID + " and " +

                        "dm.dealerid in (select distinct dealerid from hsrprecords  with(nolock) where NavembId='" + Navembcode + "' and  NewPdfRunningNo is null and erpassigndate is not null and  Vahanstatus='Y' and  OrderStatus='New Order'and HSRP_StateID =" + HSRP_StateID + ") and Om.OEMID   in('1275')";
                    //"dm.dealerid in (select distinct dealerid from hsrprecords where isnull(NewPdfRunningNo,'') = '' and isnull(erpassigndate,'') != '' and HSRP_StateID =" + HSRP_StateID + ") and Om.OEMID not  in('21','40','12','20')";

                    #region
                    DataTable dtOD = Utils.Utils.GetDataTable(oemDealerQuery, CnnString);
                    if (dtOD.Rows.Count > 0)
                    {
                        foreach (DataRow drOD in dtOD.Rows)
                        {

                            string oemid = drOD["oemid"].ToString().Trim();
                            string dealerid = drOD["dealerid"].ToString().Trim();
                            string oemname = drOD["oemname"].ToString().Trim();
                            string dealername = drOD["dealername"].ToString().Trim();
                            string Address = drOD["Address"].ToString().Trim();

                            //start sql query
                            #region
                            string productionQuery = string.Empty;

                            DataTable dtProduction = new DataTable(); ;



                            SqlConnection con = new SqlConnection(CnnString);
                            SqlCommand cmd = new SqlCommand("[USP_AllOEMProductionSheetOLA]", con);
                            cmd.CommandType = CommandType.StoredProcedure;
                            con.Open();
                            cmd.Parameters.AddWithValue("@navembid", Navembcode);
                            cmd.Parameters.AddWithValue("@HSRP_StateID", HSRP_StateID);
                            cmd.Parameters.AddWithValue("@Dealerid", dealerid);
                            //cmd.Parameters.AddWithValue("@Affix_Id", strsubaffixid);
                            SqlDataAdapter da = new SqlDataAdapter(cmd);
                            // dtProduction = new DataTable();
                            da.Fill(dtProduction);
                            con.Close();




                            #endregion
                            //end sql query

                            #region
                            //DataTable dtProduction = Utils.Utils.GetDataTable(productionQuery, CnnString);
                            if (dtProduction.Rows.Count > 0)
                            {

                                findRecord = true;
                                string strRunningNo = string.Empty;
                                //string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) as maxSheetNo from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + NAVEMBID + "'";
                                string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                                string strCom = Utils.Utils.getScalarValue(strSel, CnnString);
                                //DataTable dtSheetNo = fillDataTable(strSel, CnnString);
                                if (strCom.Equals(0) || strCom.Length == 0)
                                {
                                    strRunningNo = "0000001";
                                }
                                else
                                {
                                    strRunningNo = string.Format("{0:0000000}", Convert.ToInt32(strCom) + 1);
                                }
                                string strRequeNo = "select (prefixtext+right('00000'+ convert(varchar,Lastreqno+1),5)) as Reqno from prefix_Requisition  where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                                ReqNum = Utils.Utils.getScalarValue(strRequeNo, CnnString);
                                string strPRFIX = "select PrefixText from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                                string strPRFIXCom = Utils.Utils.getScalarValue(strPRFIX, CnnString);
                                strProductionSheetNo = strPRFIXCom + strRunningNo;

                                string RTOLocationName = dtProduction.Rows[0]["RTOLocationName"].ToString();

                                #region

                                html.Append("<div style='page-break-before: avoid;page-break-inside: avoid;page-break-after: always;'>" +
                                        "<table style='width:100%;border: 0px;'>" +
                                            "<tr style='border: 0px;'>" +
                                                "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>Report Generation Date:</b> " + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</div>" +
                                                "</td>" +

                                                 "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>ROSMERTA SAFETY SYSTEMS LIMITED</b> " + "</div>" +
                                                "</td>" +


                                                 "<td colspan='5' style='border: 0px;'>" +
                                                    "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Production Sheet No:</b> " + strRunningNo + " </div>" +
                                                "</td>" +

                                            "</tr>" +

                                              "<tr style='border: 0px;'>" +
                                                "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>State:</b> " + HSRPStateName + "</div>" +
                                                "</td>" +

                                                 "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>Production Sheet</b> " + "</div>" +
                                                "</td>" +


                                                 "<td colspan='5' style='border: 0px;'>" +
                                                    "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Dealer ID/Name:</b> " + dealerid + "/" + dealername + " </div>" +
                                                "</td>" +

                                            "</tr>" +

                                                "<tr style='border: 0px;'>" +
                                                "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='text-align:left'><b>EC Location: </b> " + RTOLocationName + "</div>" +
                                                "</td>" +

                                                 "<td colspan='3' style='border: 0px;'>" +
                                                    "<div style='style='text-align:left'><b>Oem :</b> " + oemname + "</div>" +
                                                "</td>" +


                                                 "<td colspan='5' style='border: 0px;'>" +
                                                     "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Dealer Address:</b> " + Address + " </div>" +
                                                //"<div style='text-align:left'><b>Dealer Address:</b> " + Address + " </div>" +
                                                "</td>" +

                                            "</tr>" +


  "<tr>" +
                                                "<td style='text-align:center'>Sr. No.</td>" +
                                                  "<td>Vehicle No.</td>" +
                                                   "<td>Front Plate Size</td>" +
                                                "<td>Front Laser No.</td>" +

                                                "<td>Rear Plate Size</td>" +
                                                "<td>Rear Laser No.</td>" +

                                                "<td>H. S. Foil </td>" +
                                                "<td>Caution Sticker</td>" +
                                                "<td>Fuel Type</td>" +
                                                "<td >VT</td>" +
                                                "<td >VC</ td>" +

                                            "</tr>");
                                
                                #endregion

                                #region
                                foreach (DataRow drProduction in dtProduction.Rows)
                                {
                                    string HsrprecordID = drProduction["hsrprecordID"].ToString().Trim();
                                    string SRNo = drProduction["SRNo"].ToString().Trim();
                                    string ORD = "";// drProduction["ORD"].ToString().Trim();
                                    string VC = drProduction["VehicleClass"].ToString().Trim();
                                    string VehicleNo = drProduction["VehicleRegNo"].ToString().Trim();
                                    string VT = drProduction["VehicleType"].ToString().Trim();
                                   // string ChassisNo = drProduction["ChassisNo"].ToString().Trim();
                                    //string EngineNo = drProduction["EngineNo"].ToString().Trim();
                                    string FuelType = drProduction["FuelType"].ToString().Trim();
                                    string FrontPSize = drProduction["FrontProductCode"].ToString().Trim();
                                    string FrontLaserNo = drProduction["HSRP_Front_LaserCode"].ToString().Trim();
                                    string RearPSize = drProduction["RearProductCode"].ToString().Trim();
                                    string RearLaserNo = drProduction["HSRP_Rear_LaserCode"].ToString().Trim();
                                    string Amount = drProduction["roundoff_netamount"].ToString().Trim();
                                    string OrderStatus = drProduction["OrderStatus"].ToString().Trim();
                                    string stickerColor = drProduction["stickerColor"].ToString().Trim();
                                    string HotStampingFoilColour = drProduction["HotStampingFoilColour"].ToString().Trim();

                                    html.Append("<tr>" +
                                        "<td style='text-align:center'>" + SRNo + "</td>" +
                                            "<td style='text-align:center;font-size:20px;'>" + "<b>" + VehicleNo + "</b> </td>" +

                                             "<td>" + FrontPSize + "</td>" +
                                              "<td style='text-align:center;font-size:20px;'>" + "<b>" + FrontLaserNo + "</b> </td>" +
                                                "<td>" + RearPSize + "</td>" +
                                              "<td style='text-align:center;font-size:20px;'>" + "<b>" + RearLaserNo + "</b> </td>" +

                                               "<td>" + HotStampingFoilColour + "</td>" +
                                                "<td>" + stickerColor + "</td>" +
                                                 "<td>" + FuelType + "</td>" +
                                                 "<td>" + VT + "</td>" +
                                                 "<td>" + VC + "</td>" +
                                               "</tr>");

                                    try
                                    {
                                        //start updating hsrprecords 
                                        string sqlUpdateHSRPRecords = "update hsrprecords set sendtoProductionStatus='Y', NAVPDFFlag='1', NewPdfRunningNo='" + strProductionSheetNo + "', Requisitionsheetno='" + ReqNum + "',  " +
                                              "PdfDownloadDate=GetDate(), pdfFileName='" + fileName + "', PDFDownloadUserID='1' where hsrprecordID='" + HsrprecordID + "' ";

                                        Utils.Utils.ExecNonQuery(sqlUpdateHSRPRecords, CnnString);   // uncomment after testing
                                    }
                                    catch (Exception ev)
                                    {
                                        Label1.Text = "hsrprecords update error: " + ev.Message;
                                    }
                                    //end 

                                }
                                #endregion
                                html.Append("</table>");

                                html.Append("</div>");

                                try
                                {
                                    string StrSqlUpdateECQuery = "update EmbossingCentersNew set NewProductionSheetRunningNo='" + strProductionSheetNo + "' " +
                                     "where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                                    Utils.Utils.ExecNonQuery(StrSqlUpdateECQuery, CnnString); // uncomment after testing
                                }
                                catch (Exception ev)
                                {
                                    Label1.Text = "EmbossingCentersNew update error: " + ev.Message;
                                }
                            }

                            #endregion

                        }
                    }// close oemDealerQuery
                    #endregion

                    #region "Req Generate"
                    string strComNew = string.Empty;
                    string strReqNumber = string.Empty;
                    string strReqNo = string.Empty;
                    string strSqlQuery1 = "select CompanyName from hsrpstate where hsrp_stateid='" + HSRP_StateID + "'";
                    strCompanyName = Utils.Utils.getScalarValue(strSqlQuery1, CnnString);

                    string strEMBName = " select EmbCenterName from EmbossingCentersNew where  State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                    strComNew = Utils.Utils.getScalarValue(strEMBName, CnnString);
                    if (strComNew != string.Empty)
                    {
                        strReqNo = "select (prefixtext+right('00000'+ convert(varchar,lastno+1),5)) as Reqno from prefix_Requisition  where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                        strReqNumber = Utils.Utils.getScalarValue(strReqNo, CnnString);

                        SQLString = "Exec [laserreqSlip1DashBoard]  '" + HSRP_StateID + "','" + Navembcode + "' ,  '" + ReqNum + "'";
                        DataTable dtResult = Utils.Utils.GetDataTable(SQLString, CnnString);
                        string strQuery = string.Empty;
                        string strRtoLocationName = string.Empty;
                        int Itotal = 0;

                        html.Append("<div style='width:100%;height:100%;'>" +
                                            "<table style='width:100%'>" +

                                                "<tr>" +
                                                    "<td colspan='12'>" +
                                                        "<div style='text-align:center;padding:8px;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + strCompanyName + "</b>" +

                                                            "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                             "<b style='font-size:20px;'>" + "MATERIAL REQUSITION NOTE" + "</b>" +

                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +

                                                "<tr>" +
                                                    "<td colspan='12'>" +
                                                        "<div style='text-align:center;padding:8px;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Production Sheet Date :" + DateTime.Now.ToString("dd-MM-yyyy") + "</b>" +
                                                            "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                             "<b style='font-size:20px;'>" + "" + "</b>" +
                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +
                                                "<tr>" +
                                                    "<td colspan='6'>" +
                                                        "<div style='text-align:left;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>REQ.NO:-</b>" +
                                                            "" + strReqNumber + "" +
                                                        "</div>" +
                                                    "</td>" +
                                                    "<td colspan='6'>" +
                                                        "<div style='text-align:left;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>Embossing Center:</b>" +
                                                            "" + strComNew + "" +
                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +
                                                "<tr>" +
                                                    "<td colspan='12'>" +
                                                        "<div style='text-align:left;padding:2px;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "State:" + HSRPStateName + "</b>" +
                                                            "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                             "<b style='font-size:20px;'>" + "" + "</b>" +
                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +
                                                 "<tr>" +
                                                    "<td colspan='12'>" +
                                                        "<div style='text-align:left;padding:8px;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + " " + "</b>" +
                                                            "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                             "<b style='font-size:20px;'>" + "" + "</b>" +
                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +

                                                "<tr>" +
                                                    "<td colspan='1' style='text-align:center'>SR.N.</td>" +
                                                    "<td colspan='3'>Product Size</td>" +
                                                    "<td colspan='1'>Laser Count</td>" +
                                                    "<td colspan='1'>Start Laser No</td>" +
                                                    "<td colspan='1'>End Laser No</td>" +

                                                "</tr>");


                        if (dtResult.Rows.Count > 0)
                        {
                            for (int i = 0; i < dtResult.Rows.Count; i++)
                            {
                                string ID = dtResult.Rows[i]["ID"].ToString();
                                string productcode = dtResult.Rows[i]["productcode"].ToString();
                                string LaserCount = dtResult.Rows[i]["LaserCount"].ToString();
                                Itotal = Convert.ToInt32(dtResult.Rows[i]["Total"].ToString());
                                string BeginLaser = dtResult.Rows[i]["BeginLaser"].ToString();
                                string EndLaser = dtResult.Rows[i]["EndLaser"].ToString();




                                html.Append("<tr>" +
                                   "<td colspan='1' style='text-align:left'>" + ID + "</td>" +
                                   "<td colspan='3'>" + productcode + "</td>" +

                                   "<td colspan='1'>" + LaserCount + "</td>" +
                                   "<td colspan='1'>" + BeginLaser + "</td>" +
                                   "<td colspan='1'>" + EndLaser + "</td>" +

                               "</tr>");
                            }
                        }
                        html.Append("<tr>" +
                                 "<td colspan='1' style='text-align:center' > " + "<b>Grand Total:</b>" + "</td>" +
                                 "<td colspan='3'>" + "" + "</td>" +

                                 "<td colspan='1'>" + Itotal + "</td>" +
                                 "<td colspan='1'>" + " " + "</td>" +
                                 "<td colspan='1'>" + " " + "</td>" +

                             "</tr>");




                        html.Append("<tr>" +
                         "<td colspan='2' > " + "<b>REQUESTED BY </b>" + "</td>" +
                         "<td colspan='2'>" + "<b>AUTHORIZED BY </b>" + "</td>" +

                         "<td colspan='2'>" + "<b>ISSUED BY </b>" + "</td>" +
                         "<td colspan='2'>" + "<b>RECEIVED BY</b>" + "</td>" +


                     "</tr>");

                        html.Append("<tr>" +
                                                  "<td colspan='12'>" +
                                                      "<div style='text-align:left;padding:2px;'>" +
                                                          "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Name" + "</b>" +

                                                      "</div>" +
                                                  "</td>" +
                                              "</tr>");

                        html.Append("<tr>" +
                                                 "<td colspan='12'>" +
                                                     "<div style='text-align:left;padding:2px;'>" +
                                                         "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Designation" + "</b>" +

                                                     "</div>" +
                                                 "</td>" +
                                             "</tr>");

                        html.Append("<tr>" +
                                               "<td colspan='12'>" +
                                                   "<div style='text-align:right;padding:8px;'>" +
                                                       "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Download By: Admin " + "</b>" +

                                                   "</div>" +
                                               "</td>" +
                                           "</tr>");

                        html.Append("<tr>" +
                                              "<td colspan='12'>" +
                                                  "<div style='text-align:right;padding:8px;'>" +
                                                      "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Sheet Generated By :" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</b>" +

                                                  "</div>" +
                                              "</td>" +
                                          "</tr>");






                        html.Append("</table>");

                        html.Append("</div>");

                        try
                        {
                            //start updating hsrprecords 
                            string Query = "update prefix_Requisition set lastno=lastno+1,Lastreqno=Lastreqno+1 where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                            Utils.Utils.ExecNonQuery(Query, CnnString);
                        }
                        catch (Exception ev)
                        {
                            Label1.Text = "prefix Requisition update error: " + ev.Message;
                        }
                    }


                    #endregion

                    /*
                     * Close body & HTMl Tag
                     */
                    html.Append("</body>" +
                        "</html>");


                    if (findRecord)
                    {

                        string strSaveSQlQuery = "insert into ProductionSheetAutoGenerated_List (HSRP_StateID, State_Code, RTO_ID, Emb_Center_Id, FileName, Productiondate, ProductionTime, OrderNo) values " +
                                    "('" + HSRP_StateID + "', '" + HSRPStateShortName + "', '', '" + Navembcode + "', '" + fileName + "', '" + DateTime.Now.ToString("yyyy-MM-dd") + "', '" + DateTime.Now.ToString("HH:mm:ss") + "', '" + filePrefix + "') ";
                        Utils.Utils.ExecNonQuery(strSaveSQlQuery, CnnString); // uncomment after testing
                        Label1.Text = Label1.Text + "successfully created production sheet of State: " + HSRPStateShortName + ", , Emb_Center_Id: " + Navembcode + Environment.NewLine;

                        //lblLog.Text = lblLog.Text + Environment.NewLine + "Start Production Query at:" + DateTime.Now;
                        #region
                        try
                        {
                            if (!Directory.Exists(dir))
                            {


                                Directory.CreateDirectory(dir);

                            }

                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        try
                        {
                            var pdf = Pdf
                                    .From(html.ToString())
                                    .OfSize(PaperSize.A4)
                                    .WithTitle("Title")
                                    .WithoutOutline()
                                    .WithMargins(1.25.Centimeters())
                                    .Landscape()
                                    .Comressed()
                                    .Content();

                            FileStream readStream = File.Create(filePath);
                            BinaryWriter binaryWriter = new BinaryWriter(readStream);

                            // Write the binary data to the file
                            binaryWriter.Write(pdf);
                            binaryWriter.Close();
                            readStream.Close();
                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        #endregion
                    }

                    #endregion

                }//close foreach stateEcQuery
            }
        }


        protected void btnmultiBrandDealer_Click(object sender, EventArgs e)
        {

            if (ddlStateName.SelectedItem.Text == "--Select State--")
            {
                lblErrMess.Text = "Please select  State";
                return;
            }


            string sqlPrefixQuery = "select top 1 CONVERT(INT, isnull(OrderNo,'0')) + 1 as orderNo from ProductionSheetAutoGenerated_List order by id desc";
            DataTable dtPrefix = Utils.Utils.GetDataTable(sqlPrefixQuery, CnnString);

            if (dtPrefix.Rows.Count > 0)
            {
                filePrefix = dtPrefix.Rows[0]["orderNo"].ToString();
            }
            else
            {
                filePrefix = "1";
            }

          

            if (filePrefix.Length > 0)
            {

                SheetGenerationmultiBrandDealer();


            }

        }

        private void SheetGenerationmultiBrandDealer()
        {
            string stateECQuery = string.Empty;
            string ReqNum = string.Empty;
            //string Navembid = Session["Navembid"].ToString();
            FillUserDetails();
            stateECQuery = "USP_FetchStateEC '" + Navembid + "'";
            //select distinct a.HSRP_StateID, (select HSRPStateName from hsrpstate c where c.hsrp_stateid=a.hsrp_stateid) as HSRPStateName,(select HSRPStateShortName from hsrpstate d where d.hsrp_stateid=a.hsrp_stateid) as HSRPStateShortName, d.navembcode from hsrprecords a join DealerAffixation d  on  a.Affix_Id=d.SubDealerId and isnull(NewPdfRunningNo,'') = '' and  a.HSRP_StateID =='" + ddlStateName.SelectedValue + "' and d.Navembcode not like '%CODO%'  and isnull(erpassigndate,'') != ''   AND d.navembcode='" + Navembid + "' order by  a.HSRP_StateID";
            DataTable dtSE = Utils.Utils.GetDataTable(stateECQuery, CnnString);
            if (dtSE.Rows.Count > 0)
            {
                foreach (DataRow dr in dtSE.Rows)
                {
                    #region

                    string HSRP_StateID = dr["HSRP_StateID"].ToString().Trim();
                    string HSRPStateName = dr["HSRPStateName"].ToString().Trim();
                    string HSRPStateShortName = dr["HSRPStateShortName"].ToString().Trim();
                    string Navembcode = dr["Navembcode"].ToString().Trim();

                    string dir = dirPath + DateTime.Now.ToString("yyyy-MM-dd") + "\\" + HSRPStateShortName + "\\";
                    //string dir = dirPath + "/" + DateTime.Now.ToString("yyyy-MM-dd") + "/" + HSRPStateShortName + "/";
                    //string fileName = filePrefix + "-" + Navembcode + ".pdf";
                    string fileName = "MultiBrand" + "-" + filePrefix + "-" + Navembcode + ".pdf";
                    string filePath = dir + fileName;

                    //string folderpath = ConfigurationManager.AppSettings["InvoiceFolder"].ToString() + "/" + FinYear + "/" + oemid + "/" + HSRPStateID + "/";

                    StringBuilder html = new StringBuilder();

                    Boolean findRecord = false;
                    string strProductionSheetNo = string.Empty;

                    /*
                     *  Start body & HTMl Tag
                     */
                    #region
                    html.Append(
                        "<!DOCTYPE html>" +
                        "<html>" +
                        "<head>" +
                            "<meta charset='UTF-8'><title>Title</title>" +
                            "<style>" +
                                "@page {" +
                                    /* headers*/
                                    "@top-left {" +
                                        "content: 'Left header';" +
                                    "}" +
                                    "@top-right {" +
                                        "content: 'Right header';" +
                                    "}" +

                                    /* footers */
                                    "@bottom-left {" +
                                        "content: 'Lorem ipsum';" +
                                    "} " +
                                    "@bottom-right {" +
                                        "content: 'Page ' counter(page) ' of ' counter(pages);" +
                                    "}" +
                                    "@bottom-center  {" +
                                        "content:element(footer);" +
                                    "}" +
                                "}" +
                                 "#footer {" +
                                    "position: running(footer);" +
                                "}" +
                                "table {" +
                                  "border-collapse: collapse;" +
                                "}" +

                                "table, th, td {" +
                                    "border: 1px solid black;" +
                                    "text-align: left;" +
                                    "vertical-align: top;" +
                                    "padding:5px;" +
                                "}" +
                            "</style>" +
                        "</head>" +
                        "<body>");
                    #endregion

                    string oemDealerQuery = string.Empty;

                    oemDealerQuery = "USP_FetchMultiBrandDealerid '" + Navembid + "'";

                    #region
                    DataTable dtOD = Utils.Utils.GetDataTable(oemDealerQuery, CnnString);
                    if (dtOD.Rows.Count > 0)
                    {
                        foreach (DataRow drOD in dtOD.Rows)
                        {

                            string oemid = drOD["oemid"].ToString().Trim();
                            string dealerid = drOD["dealerid"].ToString().Trim();
                            string oemname = drOD["oemname"].ToString().Trim();
                            string dealername = drOD["dealername"].ToString().Trim();
                            string Address = drOD["Address"].ToString().Trim();
                            string strsubaffixid = drOD["SubDealerId"].ToString();

                            //start sql query
                            #region
                            string productionQuery = string.Empty;


                            DataTable dtProduction = new DataTable();

                            #endregion
                            //end sql query

                            #region

                            SqlConnection con = new SqlConnection(CnnString);
                            SqlCommand cmd = new SqlCommand("USP_MultiBrandDealerProductionSheet", con);
                            cmd.CommandType = CommandType.StoredProcedure;
                            con.Open();
                            cmd.Parameters.AddWithValue("@navembid", Navembcode);
                            cmd.Parameters.AddWithValue("@HSRP_StateID", HSRP_StateID);
                            cmd.Parameters.AddWithValue("@Dealerid", dealerid);
                            cmd.Parameters.AddWithValue("@Affix_Id", strsubaffixid);
                            SqlDataAdapter da = new SqlDataAdapter(cmd);
                            // dtProduction = new DataTable();
                            da.Fill(dtProduction);
                            con.Close();


                            // DataTable dtProduction = Utils.Utils.GetDataTable(productionQuery, CnnString);
                            if (dtProduction.Rows.Count > 0)
                            {

                                findRecord = true;
                                string strRunningNo = string.Empty;
                                //string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) as maxSheetNo from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + NAVEMBID + "'";
                                string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                                string strCom = Utils.Utils.getScalarValue(strSel, CnnString);
                                //DataTable dtSheetNo = fillDataTable(strSel, CnnString);
                                if (strCom.Equals(0) || strCom.Length == 0)
                                {
                                    strRunningNo = "0000001";
                                }
                                else
                                {
                                    strRunningNo = string.Format("{0:0000000}", Convert.ToInt32(strCom) + 1);
                                }
                                string strRequeNo = "select (prefixtext+right('00000'+ convert(varchar,Lastreqno+1),5)) as Reqno from prefix_Requisition  where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                                ReqNum = Utils.Utils.getScalarValue(strRequeNo, CnnString);
                                string strPRFIX = "select PrefixText from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                                string strPRFIXCom = Utils.Utils.getScalarValue(strPRFIX, CnnString);
                                strProductionSheetNo = strPRFIXCom + strRunningNo;

                                string RTOLocationName = dtProduction.Rows[0]["RTOLocationName"].ToString();

                                #region
                                html.Append("<div style='page-break-before: avoid;page-break-inside: avoid;page-break-after: always;'>" +
                                       "<table style='width:100%;border: 0px;'>" +
                                           "<tr style='border: 0px;'>" +
                                               "<td colspan='3' style='border: 0px;'>" +
                                                   "<div style='text-align:left'><b>Report Generation Date:</b> " + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</div>" +
                                               "</td>" +

                                                "<td colspan='3' style='border: 0px;'>" +
                                                   "<div style='text-align:left'><b>ROSMERTA SAFETY SYSTEMS LIMITED</b> " + "</div>" +
                                               "</td>" +


                                                "<td colspan='5' style='border: 0px;'>" +
                                                   "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Production Sheet No Dealer:</b> " + strRunningNo + " </div>" +
                                               "</td>" +

                                           "</tr>" +

                                             "<tr style='border: 0px;'>" +
                                               "<td colspan='3' style='border: 0px;'>" +
                                                   "<div style='text-align:left'><b>State:</b> " + HSRPStateName + "</div>" +
                                               "</td>" +

                                                "<td colspan='3' style='border: 0px;'>" +
                                                   "<div style='text-align:left'><b>Production Sheet</b> " + "</div>" +
                                               "</td>" +


                                                "<td colspan='5' style='border: 0px;'>" +
                                                   "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Dealer ID/Name:</b> " + dealerid + "/" + dealername + "/" + strsubaffixid + " </div>" +
                                               "</td>" +

                                           "</tr>" +

                                               "<tr style='border: 0px;'>" +
                                               "<td colspan='3' style='border: 0px;'>" +
                                                   "<div style='text-align:left'><b>EC Location: </b> " + RTOLocationName + "</div>" +
                                               "</td>" +

                                                "<td colspan='3' style='border: 0px;'>" +
                                                   "<div style='style='text-align:left'><b>Oem :</b> " + oemname + "</div>" +
                                               "</td>" +


                                                "<td colspan='5' style='border: 0px;'>" +
                                                    "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Dealer Address:</b> " + Address + " </div>" +
                                               //"<div style='text-align:left'><b>Dealer Address:</b> " + Address + " </div>" +
                                               "</td>" +

                                           "</tr>" +


 "<tr>" +
                                               "<td style='text-align:center'>Sr. No.</td>" +
                                                 "<td>Vehicle No.</td>" +
                                                  "<td>Front Plate Size</td>" +
                                               "<td>Front Laser No.</td>" +

                                               "<td>Rear Plate Size</td>" +
                                               "<td>Rear Laser No.</td>" +

                                               "<td>H. S. Foil </td>" +
                                               "<td>Caution Sticker</td>" +
                                               "<td>Fuel Type</td>" +
                                               "<td >VT</td>" +
                                               "<td >VC</ td>" +

                                           "</tr>");
                                #endregion

                                #region
                                foreach (DataRow drProduction in dtProduction.Rows)
                                {
                                    string HsrprecordID = drProduction["hsrprecordID"].ToString().Trim();
                                    string SRNo = drProduction["SRNo"].ToString().Trim();
                                    string ORD = "";// drProduction["ORD"].ToString().Trim();
                                    string VC = drProduction["VehicleClass"].ToString().Trim();
                                    string VehicleNo = drProduction["VehicleRegNo"].ToString().Trim();
                                    // string VehicleNo = "<td style='text-align:center;font-size:20px;'>" + "<b>" + drProduction["VehicleRegNo"].ToString().Trim() + "</b> </td>" +
                                    string VT = drProduction["VehicleType"].ToString().Trim();
                                    //string ChassisNo = drProduction["ChassisNo"].ToString().Trim();
                                    //string EngineNo = drProduction["EngineNo"].ToString().Trim();
                                    string FuelType = drProduction["FuelType"].ToString().Trim();
                                    string FrontPSize = drProduction["FrontProductCode"].ToString().Trim();
                                    // string FrontLaserNo = "<td style='text-align:center;font-size:20px;'>" + "<b>" + drProduction["HSRP_Front_LaserCode"].ToString().Trim() + "</b> </td>" +
                                    string FrontLaserNo = drProduction["HSRP_Front_LaserCode"].ToString().Trim();
                                    string RearPSize = drProduction["RearProductCode"].ToString().Trim();
                                    string RearLaserNo = drProduction["HSRP_Rear_LaserCode"].ToString().Trim();
                                    //string RearLaserNo = "<td style='text-align:center;font-size:20px;'>" + "<b>" + drProduction["HSRP_Rear_LaserCode"].ToString().Trim() + "</b> </td>" +
                                    string Amount = drProduction["roundoff_netamount"].ToString().Trim();
                                    //string OrderStatus = drProduction["OrderStatus"].ToString().Trim();
                                    string stickerColor = drProduction["stickerColor"].ToString().Trim();
                                    string HotStampingFoilColour = drProduction["HotStampingFoilColour"].ToString().Trim();

                                    html.Append("<tr>" +
                                   "<td style='text-align:center'>" + SRNo + "</td>" +
                                       "<td style='text-align:center;font-size:20px;'>" + "<b>" + VehicleNo + "</b> </td>" +

                                        "<td>" + FrontPSize + "</td>" +
                                         "<td style='text-align:center;font-size:20px;'>" + "<b>" + FrontLaserNo + "</b> </td>" +
                                           "<td>" + RearPSize + "</td>" +
                                         "<td style='text-align:center;font-size:20px;'>" + "<b>" + RearLaserNo + "</b> </td>" +

                                          "<td>" + HotStampingFoilColour + "</td>" +
                                           "<td>" + stickerColor + "</td>" +
                                            "<td>" + FuelType + "</td>" +
                                            "<td>" + VT + "</td>" +
                                            "<td>" + VC + "</td>" +





                               "</tr>");

                                    try
                                    {
                                        //start updating hsrprecords 
                                        string sqlUpdateHSRPRecords = "update hsrprecords set sendtoProductionStatus='Y', NAVPDFFlag='1', NewPdfRunningNo='" + strProductionSheetNo + "', Requisitionsheetno='" + ReqNum + "',  " +
                                              "PdfDownloadDate=GetDate(), pdfFileName='" + fileName + "', PDFDownloadUserID='1' where hsrprecordID='" + HsrprecordID + "' ";

                                        Utils.Utils.ExecNonQuery(sqlUpdateHSRPRecords, CnnString);   // uncomment after testing
                                    }
                                    catch (Exception ev)
                                    {
                                        Label1.Text = "hsrprecords update error: " + ev.Message;
                                    }
                                    //end 

                                }
                                #endregion
                                html.Append("</table>");

                                html.Append("</div>");

                                try
                                {
                                    string StrSqlUpdateECQuery = "update EmbossingCentersNew set NewProductionSheetRunningNo='" + strProductionSheetNo + "' " +
                                     "where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                                    Utils.Utils.ExecNonQuery(StrSqlUpdateECQuery, CnnString); // uncomment after testing
                                }
                                catch (Exception ev)
                                {
                                    Label1.Text = "EmbossingCentersNew update error: " + ev.Message;
                                }
                            }

                            #endregion

                        }
                    }// close oemDealerQuery
                    #endregion

                    #region "Req Generate"
                    string strComNew = string.Empty;
                    string strReqNumber = string.Empty;
                    string strReqNo = string.Empty;



                    string strSqlQuery1 = "select CompanyName from hsrpstate where hsrp_stateid='" + HSRP_StateID + "'";
                    strCompanyName = Utils.Utils.getScalarValue(strSqlQuery1, CnnString);

                    string strEMBName = " select EmbCenterName from EmbossingCentersNew where  State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                    strComNew = Utils.Utils.getScalarValue(strEMBName, CnnString);

                    if (strComNew != "")
                    {



                        strReqNo = "select (prefixtext+right('00000'+ convert(varchar,lastno+1),5)) as Reqno from prefix_Requisition  where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                        strReqNumber = Utils.Utils.getScalarValue(strReqNo, CnnString);

                        SQLString = "Exec [laserreqSlip1DashBoard]  '" + HSRP_StateID + "','" + Navembcode + "' ,  '" + ReqNum + "'";
                        DataTable dtResult = Utils.Utils.GetDataTable(SQLString, CnnString);

                        string strQuery = string.Empty;
                        string strRtoLocationName = string.Empty;
                        int Itotal = 0;

                        html.Append("<div style='width:100%;height:100%;'>" +
                                            "<table style='width:100%'>" +

                                                "<tr>" +
                                                    "<td colspan='12'>" +
                                                        "<div style='text-align:center;padding:8px;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + strCompanyName + "</b>" +

                                                            "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                             "<b style='font-size:20px;'>" + "MATERIAL REQUSITION NOTE" + "</b>" +

                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +

                                                "<tr>" +
                                                    "<td colspan='12'>" +
                                                        "<div style='text-align:center;padding:8px;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Production Sheet Date :" + DateTime.Now.ToString("dd-MM-yyyy") + "</b>" +
                                                            "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                             "<b style='font-size:20px;'>" + "" + "</b>" +
                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +
                                                "<tr>" +
                                                    "<td colspan='6'>" +
                                                        "<div style='text-align:left;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>REQ.NO:-</b>" +
                                                            "" + strReqNumber + "" +
                                                        "</div>" +
                                                    "</td>" +
                                                    "<td colspan='6'>" +
                                                        "<div style='text-align:left;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>Embossing Center:</b>" +
                                                            "" + strComNew + "" +
                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +
                                                "<tr>" +
                                                    "<td colspan='12'>" +
                                                        "<div style='text-align:left;padding:2px;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "State:" + HSRPStateName + "</b>" +
                                                            "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                             "<b style='font-size:20px;'>" + "" + "</b>" +
                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +
                                                 "<tr>" +
                                                    "<td colspan='12'>" +
                                                        "<div style='text-align:left;padding:8px;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + " " + "</b>" +
                                                            "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                             "<b style='font-size:20px;'>" + "" + "</b>" +
                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +

                                                "<tr>" +
                                                    "<td colspan='1' style='text-align:center'>SR.N.</td>" +
                                                    "<td colspan='3'>Product Size</td>" +
                                                    "<td colspan='1'>Laser Count</td>" +
                                                    "<td colspan='1'>Start Laser No</td>" +
                                                    "<td colspan='1'>End Laser No</td>" +

                                                "</tr>");


                        if (dtResult.Rows.Count > 0)
                        {
                            for (int i = 0; i < dtResult.Rows.Count; i++)
                            {
                                string ID = dtResult.Rows[i]["ID"].ToString();
                                string productcode = dtResult.Rows[i]["productcode"].ToString();
                                string LaserCount = dtResult.Rows[i]["LaserCount"].ToString();
                                Itotal = Convert.ToInt32(dtResult.Rows[i]["Total"].ToString());
                                string BeginLaser = dtResult.Rows[i]["BeginLaser"].ToString();
                                string EndLaser = dtResult.Rows[i]["EndLaser"].ToString();




                                html.Append("<tr>" +
                                   "<td colspan='1' style='text-align:left'>" + ID + "</td>" +
                                   "<td colspan='3'>" + productcode + "</td>" +

                                   "<td colspan='1'>" + LaserCount + "</td>" +
                                   "<td colspan='1'>" + BeginLaser + "</td>" +
                                   "<td colspan='1'>" + EndLaser + "</td>" +

                               "</tr>");
                            }
                        }
                        html.Append("<tr>" +
                                 "<td colspan='1' style='text-align:center' > " + "<b>Grand Total:</b>" + "</td>" +
                                 "<td colspan='3'>" + "" + "</td>" +

                                 "<td colspan='1'>" + Itotal + "</td>" +
                                 "<td colspan='1'>" + " " + "</td>" +
                                 "<td colspan='1'>" + " " + "</td>" +

                             "</tr>");




                        html.Append("<tr>" +
                         "<td colspan='2' > " + "<b>REQUESTED BY </b>" + "</td>" +
                         "<td colspan='2'>" + "<b>AUTHORIZED BY </b>" + "</td>" +

                         "<td colspan='2'>" + "<b>ISSUED BY </b>" + "</td>" +
                         "<td colspan='2'>" + "<b>RECEIVED BY</b>" + "</td>" +


                     "</tr>");

                        html.Append("<tr>" +
                                                  "<td colspan='12'>" +
                                                      "<div style='text-align:left;padding:2px;'>" +
                                                          "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Name" + "</b>" +

                                                      "</div>" +
                                                  "</td>" +
                                              "</tr>");

                        html.Append("<tr>" +
                                                 "<td colspan='12'>" +
                                                     "<div style='text-align:left;padding:2px;'>" +
                                                         "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Designation" + "</b>" +

                                                     "</div>" +
                                                 "</td>" +
                                             "</tr>");

                        html.Append("<tr>" +
                                               "<td colspan='12'>" +
                                                   "<div style='text-align:right;padding:8px;'>" +
                                                       "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Download By: Admin " + "</b>" +

                                                   "</div>" +
                                               "</td>" +
                                           "</tr>");

                        html.Append("<tr>" +
                                              "<td colspan='12'>" +
                                                  "<div style='text-align:right;padding:8px;'>" +
                                                      "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Sheet Generated By :" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</b>" +

                                                  "</div>" +
                                              "</td>" +
                                          "</tr>");







                        html.Append("</table>");

                        html.Append("</div>");

                        try
                        {
                            //start updating hsrprecords 
                            string Query = "update prefix_Requisition set lastno=lastno+1,Lastreqno=Lastreqno+1 where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                            Utils.Utils.ExecNonQuery(Query, CnnString);
                        }
                        catch (Exception ev)
                        {
                            Label1.Text = "prefix Requisition update error: " + ev.Message;
                        }
                    }


                    #endregion

                    /*
                     * Close body & HTMl Tag
                     */
                    html.Append("</body>" +
                        "</html>");


                    if (findRecord)
                    {

                        string strSaveSQlQuery = "insert into ProductionSheetAutoGenerated_List (HSRP_StateID, State_Code, RTO_ID, Emb_Center_Id, FileName, Productiondate, ProductionTime, OrderNo) values " +
                                    "('" + HSRP_StateID + "', '" + HSRPStateShortName + "', '', '" + Navembcode + "', '" + fileName + "', '" + DateTime.Now.ToString("yyyy-MM-dd") + "', '" + DateTime.Now.ToString("HH:mm:ss") + "', '" + filePrefix + "') ";
                        Utils.Utils.ExecNonQuery(strSaveSQlQuery, CnnString); // uncomment after testing
                        Label1.Text = Label1.Text + "successfully created production sheet of State: " + HSRPStateShortName + ", , Emb_Center_Id: " + Navembcode + Environment.NewLine;

                        //lblLog.Text = lblLog.Text + Environment.NewLine + "Start Production Query at:" + DateTime.Now;
                        #region
                        try
                        {
                            if (!Directory.Exists(dir))
                            {


                                Directory.CreateDirectory(dir);

                            }

                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        try
                        {
                            var pdf = Pdf
                                    .From(html.ToString())
                                    .OfSize(PaperSize.A4)
                                    .WithTitle("Title")
                                    .WithoutOutline()
                                    .WithMargins(1.25.Centimeters())
                                    .Landscape()
                                    .Comressed()
                                    .Content();

                            FileStream readStream = File.Create(filePath);
                            BinaryWriter binaryWriter = new BinaryWriter(readStream);

                            // Write the binary data to the file
                            binaryWriter.Write(pdf);
                            binaryWriter.Close();
                            readStream.Close();
                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        #endregion
                    }

                    #endregion

                }//close foreach stateEcQuery
            }
        }

        protected void btnmultiBrandHome_Click(object sender, EventArgs e)
        {

            if (ddlStateName.SelectedItem.Text == "--Select State--")
            {
                lblErrMess.Text = "Please select  State";
                return;
            }


            string sqlPrefixQuery = "select top 1 CONVERT(INT, isnull(OrderNo,'0')) + 1 as orderNo from ProductionSheetAutoGenerated_List order by id desc";
            DataTable dtPrefix = Utils.Utils.GetDataTable(sqlPrefixQuery, CnnString);

            if (dtPrefix.Rows.Count > 0)
            {
                filePrefix = dtPrefix.Rows[0]["orderNo"].ToString();
            }
            else
            {
                filePrefix = "1";
            }

            //txtFilePrefix.Text = filePrefix;

            if (filePrefix.Length > 0)
            {

                SheetGenerationmultiBrandHome();


            }

        }

        private void SheetGenerationmultiBrandHome()
        {
            string stateECQuery = string.Empty;
            string ReqNum = string.Empty;
            //string Navembid = Session["Navembid"].ToString();
            FillUserDetails();

            stateECQuery = "select distinct a.HSRP_StateID, (select HSRPStateName from hsrpstate c where c.hsrp_stateid=a.hsrp_stateid) as HSRPStateName,(select HSRPStateShortName from hsrpstate d where d.hsrp_stateid=a.hsrp_stateid) as HSRPStateShortName, navembcode from hsrprecords a with(nolock),rtolocation b where a.HSRP_StateID ='" + ddlStateName.SelectedValue + "'     and a.rtolocationid=b.rtolocationid and NewPdfRunningNo is null and erpassigndate is not null and OrderStatus='New Order' AND b.NAVEMBID='" + Navembid + "' and  OLDOrderid=2   order by  a.HSRP_StateID ";
            //stateECQuery = "select distinct a.HSRP_StateID, (select HSRPStateName from hsrpstate c where c.hsrp_stateid=a.hsrp_stateid) as HSRPStateName,(select HSRPStateShortName from hsrpstate d where d.hsrp_stateid=a.hsrp_stateid) as HSRPStateShortName, navembcode from hsrprecords a with(nolock)  join Rtolocation  with(lock)where a.HSRP_StateID ='" + ddlStateName.SelectedValue + "' and   a.Navembcode not like '%CODO%'    and NewPdfRunningNo is null and erpassigndate is not null and OrderStatus='New Order' AND a.NAVEMBID='" + Navembid + "'   order by  a.HSRP_StateID";

            DataTable dtSE = Utils.Utils.GetDataTable(stateECQuery, CnnString);
            if (dtSE.Rows.Count > 0)
            {
                foreach (DataRow dr in dtSE.Rows)
                {
                    #region

                    string HSRP_StateID = dr["HSRP_StateID"].ToString().Trim();
                    string HSRPStateName = dr["HSRPStateName"].ToString().Trim();
                    string HSRPStateShortName = dr["HSRPStateShortName"].ToString().Trim();
                    string Navembcode = dr["Navembcode"].ToString().Trim();

                    string dir = dirPath + DateTime.Now.ToString("yyyy-MM-dd") + "\\" + HSRPStateShortName + "\\";
                    //string dir = dirPath + "/" + DateTime.Now.ToString("yyyy-MM-dd") + "/" + HSRPStateShortName + "/";
                    //string fileName = filePrefix + "-" + Navembcode + ".pdf";
                    string fileName = "MultiBrandHome" + "-" + filePrefix + "-" + Navembcode + ".pdf";
                    string filePath = dir + fileName;

                    //string folderpath = ConfigurationManager.AppSettings["InvoiceFolder"].ToString() + "/" + FinYear + "/" + oemid + "/" + HSRPStateID + "/";

                    StringBuilder html = new StringBuilder();

                    Boolean findRecord = false;
                    string strProductionSheetNo = string.Empty;

                    /*
                     *  Start body & HTMl Tag
                     */
                    #region
                    html.Append(
                        "<!DOCTYPE html>" +
                        "<html>" +
                        "<head>" +
                            "<meta charset='UTF-8'><title>Title</title>" +
                            "<style>" +
                                "@page {" +
                                    /* headers*/
                                    "@top-left {" +
                                        "content: 'Left header';" +
                                    "}" +
                                    "@top-right {" +
                                        "content: 'Right header';" +
                                    "}" +

                                    /* footers */
                                    "@bottom-left {" +
                                        "content: 'Lorem ipsum';" +
                                    "} " +
                                    "@bottom-right {" +
                                        "content: 'Page ' counter(page) ' of ' counter(pages);" +
                                    "}" +
                                    "@bottom-center  {" +
                                        "content:element(footer);" +
                                    "}" +
                                "}" +
                                 "#footer {" +
                                    "position: running(footer);" +
                                "}" +
                                "table {" +
                                  "border-collapse: collapse;" +
                                "}" +

                                "table, th, td {" +
                                    "border: 1px solid black;" +
                                    "text-align: left;" +
                                    "vertical-align: top;" +
                                    "padding:5px;" +
                                "}" +
                            "</style>" +
                        "</head>" +
                        "<body>");
                    #endregion

                    string oemDealerQuery = string.Empty;


                    oemDealerQuery = "USP_FetchMultiBrandHome '" + ddlStateName.SelectedValue + "','" + Navembid + "'";


                    #region
                    DataTable dtOD = Utils.Utils.GetDataTable(oemDealerQuery, CnnString);
                    if (dtOD.Rows.Count > 0)
                    {
                        foreach (DataRow drOD in dtOD.Rows)
                        {

                            string oemid = drOD["oemid"].ToString().Trim();
                            string dealerid = drOD["dealerid"].ToString().Trim();
                            string oemname = drOD["oemname"].ToString().Trim();
                            string dealername = drOD["dealername"].ToString().Trim();
                            string Address = drOD["Address"].ToString().Trim();
                            //string strsubaffixid = drOD["SubDealerId"].ToString();

                            //start sql query
                            #region
                            string productionQuery = string.Empty;


                            DataTable dtProduction = new DataTable();

                            #endregion
                            //end sql query

                            #region

                            SqlConnection con = new SqlConnection(CnnString);
                            SqlCommand cmd = new SqlCommand("USP_MultiBrandHomeProductionSheet", con);
                            cmd.CommandType = CommandType.StoredProcedure;
                            con.Open();
                            cmd.Parameters.AddWithValue("@navembid", Navembcode);
                            cmd.Parameters.AddWithValue("@HSRP_StateID", HSRP_StateID);
                            cmd.Parameters.AddWithValue("@Dealerid", dealerid);
                            //cmd.Parameters.AddWithValue("@Affix_Id", strsubaffixid);
                            SqlDataAdapter da = new SqlDataAdapter(cmd);
                            // dtProduction = new DataTable();
                            da.Fill(dtProduction);
                            con.Close();


                            // DataTable dtProduction = Utils.Utils.GetDataTable(productionQuery, CnnString);
                            if (dtProduction.Rows.Count > 0)
                            {

                                findRecord = true;
                                string strRunningNo = string.Empty;
                                //string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) as maxSheetNo from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + NAVEMBID + "'";
                                string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                                string strCom = Utils.Utils.getScalarValue(strSel, CnnString);
                                //DataTable dtSheetNo = fillDataTable(strSel, CnnString);
                                if (strCom.Equals(0) || strCom.Length == 0)
                                {
                                    strRunningNo = "0000001";
                                }
                                else
                                {
                                    strRunningNo = string.Format("{0:0000000}", Convert.ToInt32(strCom) + 1);
                                }
                                string strRequeNo = "select (prefixtext+right('00000'+ convert(varchar,Lastreqno+1),5)) as Reqno from prefix_Requisition  where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                                ReqNum = Utils.Utils.getScalarValue(strRequeNo, CnnString);
                                string strPRFIX = "select PrefixText from EmbossingCentersNew where Emb_Center_Id='" + Navembcode + "'";
                                string strPRFIXCom = Utils.Utils.getScalarValue(strPRFIX, CnnString);
                                strProductionSheetNo = strPRFIXCom + strRunningNo;

                                string RTOLocationName = dtProduction.Rows[0]["RTOLocationName"].ToString();

                                //#region
                                //html.Append("<div style='page-break-before: avoid;page-break-inside: avoid;page-break-after: always;'>" +
                                //        "<table style='width:100%;border: 0px;'>" +
                                //            "<tr style='border: 0px;'>" +
                                //                "<td colspan='7' style='border: 0px;'>" +
                                //                    "<div style='text-align:left'><b>Report Generation Date:</b> " + DateTime.Now.ToString("dd-MM-yyyy") + "</div>" +
                                //                "</td>" +
                                //                "<td colspan='7' style='border: 0px;'>" +
                                //                    "<div style='float:right;width: 500px;word-wrap: break-word;'>" +
                                //                        "<b>Production Sheet No:</b> " + strRunningNo + "<br />" +
                                //                     "<b>Dealer ID/Name:</b> " + dealerid + "/" + dealername + "/"  + "<br />" +
                                //                         "<b>Dealer Address:</b> " + Address + "<br />" +
                                //                    "</div>" +

                                //                "</td>" +
                                //            "</tr>" +
                                //            "<tr style='border: 0px;'>" +
                                //                "<td colspan='14' style='border: 0px;'>" +
                                //                    "<div style='text-align:center;font-size:26px;'><b>Production Sheet : -</b> Rosmerta Safety System</div>" +
                                //                "</td>" +
                                //            "</tr>" +
                                //            "<tr style='border: 0px;'>" +
                                //                "<td colspan='9' style='border: 0px;'>" +
                                //                    "<table style='border:0px;width:100%;'>" +
                                //                        "<tr style='border:0px;'>" +
                                //                            "<td style='border:0px;'><b>State Name :</b> " + HSRPStateName + "</td>" +
                                //                            "<td style='border:0px;'><b>Oem Name :</b> " + oemname + "</td>" +
                                //                            "<td style='border:0px;'><b>Report Date:</b> " + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</td>" +
                                //                        "</tr>" +
                                //                    "</table>" +
                                //                "</td>" +
                                //                "<td colspan='4' style='border: 0px;'>" +
                                //                    "<div style='float:right'>" +
                                //                        "ORD:Order Open Date<br />" +
                                //                        "VC:Vehicle Class<br />" +
                                //                        "VT:Vehicle Type<br />" +
                                //                        "Front PS:Front Plate Size<br />" +
                                //                        "Rear PS:Rear Plate Size<br />" +
                                //                        "OS: Order Satus(New Order/Embossing Done/Closed)" +
                                //                    "</div>" +
                                //                "</td>" +
                                //                "<td style='border: 0px;'></td>" +
                                //            "</tr>" +
                                //            "<tr style='border: 0px;'>" +
                                //                 "<td colspan='14' style='border: 0px;'>" +
                                //                    "<div style='text-align:left'>Location Name : " + RTOLocationName + "</div>" +
                                //                "</td>" +
                                //            "</tr>" +
                                //            "<tr>" +
                                //                "<td style='text-align:center'>SR.No</td>" +
                                //                "<td>VC</td>" +
                                //                "<td>Vehicle No</td>" +
                                //                "<td>VT</td>" +
                                //                "<td>Chassis No</td>" +
                                //                "<td>EngineNo</td>" +
                                //                "<td>Fuel Type</td>" +
                                //                "<td>Front PS</td>" +
                                //                "<td>Front Laser No</td>" +
                                //                "<td>Rear PS</td>" +
                                //                "<td>Rear Laser No.</td>" +
                                //                "<td style='text-align:center'>OS</td>" +
                                //                 "<td style='text-align:center'>sticker Color</td>" +
                                //            "</tr>");
                                //#endregion

                                //#region
                                //foreach (DataRow drProduction in dtProduction.Rows)
                                //{
                                //    string HsrprecordID = drProduction["hsrprecordID"].ToString().Trim();
                                //    string SRNo = drProduction["SRNo"].ToString().Trim();
                                //    string ORD = "";// drProduction["ORD"].ToString().Trim();
                                //    string VC = drProduction["VehicleClass"].ToString().Trim();
                                //    string VehicleNo = drProduction["VehicleRegNo"].ToString().Trim();
                                //    string VT = drProduction["VehicleType"].ToString().Trim();
                                //    string ChassisNo = drProduction["ChassisNo"].ToString().Trim();
                                //    string EngineNo = drProduction["EngineNo"].ToString().Trim();
                                //    string FuelType = drProduction["FuelType"].ToString().Trim();
                                //    string FrontPSize = drProduction["FrontProductCode"].ToString().Trim();
                                //    string FrontLaserNo = drProduction["HSRP_Front_LaserCode"].ToString().Trim();
                                //    string RearPSize = drProduction["RearProductCode"].ToString().Trim();
                                //    string RearLaserNo = drProduction["HSRP_Rear_LaserCode"].ToString().Trim();
                                //    string Amount = drProduction["roundoff_netamount"].ToString().Trim();
                                //    string OrderStatus = drProduction["OrderStatus"].ToString().Trim();
                                //    string stickerColor = drProduction["stickerColor"].ToString().Trim();


                                //    html.Append("<tr>" +
                                //       "<td style='text-align:center'>" + SRNo + "</td>" +
                                //       "<td>" + VC + "</td>" +
                                //       "<td>" + VehicleNo + "</td>" +
                                //       "<td>" + VT + "</td>" +
                                //       "<td>" + ChassisNo + "</td>" +
                                //       "<td>" + EngineNo + "</td>" +
                                //       "<td>" + FuelType + "</td>" +
                                //       "<td>" + FrontPSize + "</td>" +
                                //       "<td>" + FrontLaserNo + "</td>" +
                                //       "<td>" + RearPSize + "</td>" +
                                //       "<td>" + RearLaserNo + "</td>" +
                                //       "<td>" + OrderStatus + "</td>" +
                                //        "<td>" + stickerColor + "</td>" +
                                //   "</tr>");

                                //    try
                                //    {
                                //        //start updating hsrprecords 
                                //        string sqlUpdateHSRPRecords = "update hsrprecords set sendtoProductionStatus='Y', NAVPDFFlag='1', NewPdfRunningNo='" + strProductionSheetNo + "', Requisitionsheetno='" + ReqNum + "',  " +
                                //              "PdfDownloadDate=GetDate(), pdfFileName='" + fileName + "', PDFDownloadUserID='1' where hsrprecordID='" + HsrprecordID + "' ";

                                //        Utils.Utils.ExecNonQuery(sqlUpdateHSRPRecords, CnnString);   // uncomment after testing
                                //    }
                                //    catch (Exception ev)
                                //    {
                                //        Label1.Text = "hsrprecords update error: " + ev.Message;
                                //    }
                                //    //end 

                                //}
                                //#endregion


                                #region
                                html.Append("<div style='page-break-before: avoid;page-break-inside: avoid;page-break-after: always;'>" +
                                       "<table style='width:100%;border: 0px;'>" +
                                           "<tr style='border: 0px;'>" +
                                               "<td colspan='3' style='border: 0px;'>" +
                                                   "<div style='text-align:left'><b>Report Generation Date:</b> " + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</div>" +
                                               "</td>" +

                                                "<td colspan='3' style='border: 0px;'>" +
                                                   "<div style='text-align:left'><b>ROSMERTA SAFETY SYSTEMS LIMITED</b> " + "</div>" +
                                               "</td>" +


                                                "<td colspan='5' style='border: 0px;'>" +
                                                   "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Production Sheet No:</b> " + strRunningNo + " </div>" +
                                               "</td>" +

                                           "</tr>" +

                                             "<tr style='border: 0px;'>" +
                                               "<td colspan='3' style='border: 0px;'>" +
                                                   "<div style='text-align:left'><b>State:</b> " + HSRPStateName + "</div>" +
                                               "</td>" +

                                                "<td colspan='3' style='border: 0px;'>" +
                                                   "<div style='text-align:left'><b>Production Sheet Home</b> " + "</div>" +
                                               "</td>" +


                                                "<td colspan='5' style='border: 0px;'>" +
                                                   "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Dealer ID/Name:</b> " + dealerid + "/" + dealername + "/" + " </div>" +
                                               "</td>" +

                                           "</tr>" +

                                               "<tr style='border: 0px;'>" +
                                               "<td colspan='3' style='border: 0px;'>" +
                                                   "<div style='text-align:left'><b>EC Location: </b> " + RTOLocationName + "</div>" +
                                               "</td>" +

                                                "<td colspan='3' style='border: 0px;'>" +
                                                   "<div style='style='text-align:left'><b>Oem :</b> " + oemname + "</div>" +
                                               "</td>" +


                                                "<td colspan='5' style='border: 0px;'>" +
                                                    "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Dealer Address:</b> " + Address + " </div>" +
                                               //"<div style='text-align:left'><b>Dealer Address:</b> " + Address + " </div>" +
                                               "</td>" +

                                           "</tr>" +


 "<tr>" +
                                               "<td style='text-align:center'>Sr. No.</td>" +
                                                 "<td>Vehicle No.</td>" +
                                                  "<td>Front Plate Size</td>" +
                                               "<td>Front Laser No.</td>" +

                                               "<td>Rear Plate Size</td>" +
                                               "<td>Rear Laser No.</td>" +

                                               "<td>H. S. Foil </td>" +
                                               "<td>Caution Sticker</td>" +
                                               "<td>Fuel Type</td>" +
                                               "<td >VT</td>" +
                                               "<td >VC</ td>" +

                                           "</tr>");
                                #endregion

                                #region
                                foreach (DataRow drProduction in dtProduction.Rows)
                                {
                                    string HsrprecordID = drProduction["hsrprecordID"].ToString().Trim();
                                    string SRNo = drProduction["SRNo"].ToString().Trim();
                                    string ORD = "";// drProduction["ORD"].ToString().Trim();
                                    string VC = drProduction["VehicleClass"].ToString().Trim();
                                    string VehicleNo = drProduction["VehicleRegNo"].ToString().Trim();
                                    // string VehicleNo = "<td style='text-align:center;font-size:20px;'>" + "<b>" + drProduction["VehicleRegNo"].ToString().Trim() + "</b> </td>" +
                                    string VT = drProduction["VehicleType"].ToString().Trim();
                                    //string ChassisNo = drProduction["ChassisNo"].ToString().Trim();
                                    //string EngineNo = drProduction["EngineNo"].ToString().Trim();
                                    string FuelType = drProduction["FuelType"].ToString().Trim();
                                    string FrontPSize = drProduction["FrontProductCode"].ToString().Trim();
                                    // string FrontLaserNo = "<td style='text-align:center;font-size:20px;'>" + "<b>" + drProduction["HSRP_Front_LaserCode"].ToString().Trim() + "</b> </td>" +
                                    string FrontLaserNo = drProduction["HSRP_Front_LaserCode"].ToString().Trim();
                                    string RearPSize = drProduction["RearProductCode"].ToString().Trim();
                                    string RearLaserNo = drProduction["HSRP_Rear_LaserCode"].ToString().Trim();
                                    //string RearLaserNo = "<td style='text-align:center;font-size:20px;'>" + "<b>" + drProduction["HSRP_Rear_LaserCode"].ToString().Trim() + "</b> </td>" +
                                    string Amount = drProduction["roundoff_netamount"].ToString().Trim();
                                    //string OrderStatus = drProduction["OrderStatus"].ToString().Trim();
                                    string stickerColor = drProduction["stickerColor"].ToString().Trim();
                                    string HotStampingFoilColour = drProduction["HotStampingFoilColour"].ToString().Trim();

                                    html.Append("<tr>" +
                                   "<td style='text-align:center'>" + SRNo + "</td>" +
                                       "<td style='text-align:center;font-size:20px;'>" + "<b>" + VehicleNo + "</b> </td>" +

                                        "<td>" + FrontPSize + "</td>" +
                                         "<td style='text-align:center;font-size:20px;'>" + "<b>" + FrontLaserNo + "</b> </td>" +
                                           "<td>" + RearPSize + "</td>" +
                                         "<td style='text-align:center;font-size:20px;'>" + "<b>" + RearLaserNo + "</b> </td>" +

                                          "<td>" + HotStampingFoilColour + "</td>" +
                                           "<td>" + stickerColor + "</td>" +
                                            "<td>" + FuelType + "</td>" +
                                            "<td>" + VT + "</td>" +
                                            "<td>" + VC + "</td>" +





                               "</tr>");

                                    try
                                    {
                                        //start updating hsrprecords 
                                        string sqlUpdateHSRPRecords = "update hsrprecords set sendtoProductionStatus='Y', NAVPDFFlag='1', NewPdfRunningNo='" + strProductionSheetNo + "', Requisitionsheetno='" + ReqNum + "',  " +
                                              "PdfDownloadDate=GetDate(), pdfFileName='" + fileName + "', PDFDownloadUserID='1' where hsrprecordID='" + HsrprecordID + "' ";

                                        Utils.Utils.ExecNonQuery(sqlUpdateHSRPRecords, CnnString);   // uncomment after testing
                                    }
                                    catch (Exception ev)
                                    {
                                        Label1.Text = "hsrprecords update error: " + ev.Message;
                                    }
                                    //end 

                                }
                                #endregion
                                html.Append("</table>");

                                html.Append("</div>");

                                try
                                {
                                    string StrSqlUpdateECQuery = "update EmbossingCentersNew set NewProductionSheetRunningNo='" + strProductionSheetNo + "' " +
                                     "where  Emb_Center_Id='" + Navembcode + "'";
                                    Utils.Utils.ExecNonQuery(StrSqlUpdateECQuery, CnnString); // uncomment after testing
                                }
                                catch (Exception ev)
                                {
                                    Label1.Text = "EmbossingCentersNew update error: " + ev.Message;
                                }
                            }

                            #endregion

                        }
                    }// close oemDealerQuery
                    #endregion

                    #region "Req Generate"
                    string strComNew = string.Empty;
                    string strReqNumber = string.Empty;
                    string strReqNo = string.Empty;



                    string strSqlQuery1 = "select CompanyName from hsrpstate where hsrp_stateid='" + HSRP_StateID + "'";
                    strCompanyName = Utils.Utils.getScalarValue(strSqlQuery1, CnnString);

                    string strEMBName = " select EmbCenterName from EmbossingCentersNew where   Emb_Center_Id='" + Navembcode + "'";
                    strComNew = Utils.Utils.getScalarValue(strEMBName, CnnString);

                    if (strComNew != "")
                    {



                        strReqNo = "select (prefixtext+right('00000'+ convert(varchar,lastno+1),5)) as Reqno from prefix_Requisition  where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                        strReqNumber = Utils.Utils.getScalarValue(strReqNo, CnnString);

                        SQLString = "Exec [laserreqSlip1DashBoard]  '" + HSRP_StateID + "','" + Navembcode + "' ,  '" + ReqNum + "'";
                        DataTable dtResult = Utils.Utils.GetDataTable(SQLString, CnnString);

                        string strQuery = string.Empty;
                        string strRtoLocationName = string.Empty;
                        int Itotal = 0;

                        html.Append("<div style='width:100%;height:100%;'>" +
                                            "<table style='width:100%'>" +

                                                "<tr>" +
                                                    "<td colspan='12'>" +
                                                        "<div style='text-align:center;padding:8px;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + strCompanyName + "</b>" +

                                                            "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                             "<b style='font-size:20px;'>" + "MATERIAL REQUSITION NOTE" + "</b>" +

                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +

                                                "<tr>" +
                                                    "<td colspan='12'>" +
                                                        "<div style='text-align:center;padding:8px;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Production Sheet Date :" + DateTime.Now.ToString("dd-MM-yyyy") + "</b>" +
                                                            "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                             "<b style='font-size:20px;'>" + "" + "</b>" +
                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +
                                                "<tr>" +
                                                    "<td colspan='6'>" +
                                                        "<div style='text-align:left;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>REQ.NO:-</b>" +
                                                            "" + strReqNumber + "" +
                                                        "</div>" +
                                                    "</td>" +
                                                    "<td colspan='6'>" +
                                                        "<div style='text-align:left;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>Embossing Center:</b>" +
                                                            "" + strComNew + "" +
                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +
                                                "<tr>" +
                                                    "<td colspan='12'>" +
                                                        "<div style='text-align:left;padding:2px;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "State:" + HSRPStateName + "</b>" +
                                                            "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                             "<b style='font-size:20px;'>" + "" + "</b>" +
                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +
                                                 "<tr>" +
                                                    "<td colspan='12'>" +
                                                        "<div style='text-align:left;padding:8px;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + " " + "</b>" +
                                                            "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                             "<b style='font-size:20px;'>" + "" + "</b>" +
                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +

                                                "<tr>" +
                                                    "<td colspan='1' style='text-align:center'>SR.N.</td>" +
                                                    "<td colspan='3'>Product Size</td>" +
                                                    "<td colspan='1'>Laser Count</td>" +
                                                    "<td colspan='1'>Start Laser No</td>" +
                                                    "<td colspan='1'>End Laser No</td>" +

                                                "</tr>");


                        if (dtResult.Rows.Count > 0)
                        {
                            for (int i = 0; i < dtResult.Rows.Count; i++)
                            {
                                string ID = dtResult.Rows[i]["ID"].ToString();
                                string productcode = dtResult.Rows[i]["productcode"].ToString();
                                string LaserCount = dtResult.Rows[i]["LaserCount"].ToString();
                                Itotal = Convert.ToInt32(dtResult.Rows[i]["Total"].ToString());
                                string BeginLaser = dtResult.Rows[i]["BeginLaser"].ToString();
                                string EndLaser = dtResult.Rows[i]["EndLaser"].ToString();




                                html.Append("<tr>" +
                                   "<td colspan='1' style='text-align:left'>" + ID + "</td>" +
                                   "<td colspan='3'>" + productcode + "</td>" +

                                   "<td colspan='1'>" + LaserCount + "</td>" +
                                   "<td colspan='1'>" + BeginLaser + "</td>" +
                                   "<td colspan='1'>" + EndLaser + "</td>" +

                               "</tr>");
                            }
                        }
                        html.Append("<tr>" +
                                 "<td colspan='1' style='text-align:center' > " + "<b>Grand Total:</b>" + "</td>" +
                                 "<td colspan='3'>" + "" + "</td>" +

                                 "<td colspan='1'>" + Itotal + "</td>" +
                                 "<td colspan='1'>" + " " + "</td>" +
                                 "<td colspan='1'>" + " " + "</td>" +

                             "</tr>");




                        html.Append("<tr>" +
                         "<td colspan='2' > " + "<b>REQUESTED BY </b>" + "</td>" +
                         "<td colspan='2'>" + "<b>AUTHORIZED BY </b>" + "</td>" +

                         "<td colspan='2'>" + "<b>ISSUED BY </b>" + "</td>" +
                         "<td colspan='2'>" + "<b>RECEIVED BY</b>" + "</td>" +


                     "</tr>");

                        html.Append("<tr>" +
                                                  "<td colspan='12'>" +
                                                      "<div style='text-align:left;padding:2px;'>" +
                                                          "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Name" + "</b>" +

                                                      "</div>" +
                                                  "</td>" +
                                              "</tr>");

                        html.Append("<tr>" +
                                                 "<td colspan='12'>" +
                                                     "<div style='text-align:left;padding:2px;'>" +
                                                         "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Designation" + "</b>" +

                                                     "</div>" +
                                                 "</td>" +
                                             "</tr>");

                        html.Append("<tr>" +
                                               "<td colspan='12'>" +
                                                   "<div style='text-align:right;padding:8px;'>" +
                                                       "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Download By: Admin " + "</b>" +

                                                   "</div>" +
                                               "</td>" +
                                           "</tr>");

                        html.Append("<tr>" +
                                              "<td colspan='12'>" +
                                                  "<div style='text-align:right;padding:8px;'>" +
                                                      "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Sheet Generated By :" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</b>" +

                                                  "</div>" +
                                              "</td>" +
                                          "</tr>");







                        html.Append("</table>");

                        html.Append("</div>");

                        try
                        {
                            //start updating hsrprecords 
                            string Query = "update prefix_Requisition set lastno=lastno+1,Lastreqno=Lastreqno+1 where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                            Utils.Utils.ExecNonQuery(Query, CnnString);
                        }
                        catch (Exception ev)
                        {
                            Label1.Text = "prefix Requisition update error: " + ev.Message;
                        }
                    }


                    #endregion

                    /*
                     * Close body & HTMl Tag
                     */
                    html.Append("</body>" +
                        "</html>");


                    if (findRecord)
                    {

                        string strSaveSQlQuery = "insert into ProductionSheetAutoGenerated_List (HSRP_StateID, State_Code, RTO_ID, Emb_Center_Id, FileName, Productiondate, ProductionTime, OrderNo) values " +
                                    "('" + HSRP_StateID + "', '" + HSRPStateShortName + "', '', '" + Navembcode + "', '" + fileName + "', '" + DateTime.Now.ToString("yyyy-MM-dd") + "', '" + DateTime.Now.ToString("HH:mm:ss") + "', '" + filePrefix + "') ";
                        Utils.Utils.ExecNonQuery(strSaveSQlQuery, CnnString); // uncomment after testing
                        Label1.Text = Label1.Text + "successfully created production sheet of State: " + HSRPStateShortName + ",  Emb_Center_Id: " + Navembcode + Environment.NewLine;

                        //lblLog.Text = lblLog.Text + Environment.NewLine + "Start Production Query at:" + DateTime.Now;
                        #region
                        try
                        {
                            if (!Directory.Exists(dir))
                            {


                                Directory.CreateDirectory(dir);

                            }

                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        try
                        {
                            var pdf = Pdf
                                    .From(html.ToString())
                                    .OfSize(PaperSize.A4)
                                    .WithTitle("Title")
                                    .WithoutOutline()
                                    .WithMargins(1.25.Centimeters())
                                    .Landscape()
                                    .Comressed()
                                    .Content();

                            FileStream readStream = File.Create(filePath);
                            BinaryWriter binaryWriter = new BinaryWriter(readStream);

                            // Write the binary data to the file
                            binaryWriter.Write(pdf);
                            binaryWriter.Close();
                            readStream.Close();
                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        #endregion
                    }

                    #endregion

                }//close foreach stateEcQuery
            }
        }



        protected void btnHRMultiBrand_Click(object sender, EventArgs e)
        {
            if (ddlStateName.SelectedItem.Text == "--Select State--")
            {
                lblErrMess.Text = "Please select  State";
                return;
            }


            string sqlPrefixQuery = "select top 1 CONVERT(INT, isnull(OrderNo,'0')) + 1 as orderNo from ProductionSheetAutoGenerated_List order by id desc";
            DataTable dtPrefix = Utils.Utils.GetDataTable(sqlPrefixQuery, CnnString);

            if (dtPrefix.Rows.Count > 0)
            {
                filePrefix = dtPrefix.Rows[0]["orderNo"].ToString();
            }
            else
            {
                filePrefix = "1";
            }

            if (filePrefix.Length > 0)
            {

                HRMultiBrandDealer();
                HRMultiBrandHome();
            }
        }

        private void HRMultiBrandDealer()
        {
            string stateECQuery = string.Empty;
            string ReqNum = string.Empty;
            //string Navembid = Session["Navembid"].ToString();
            FillUserDetails();
            stateECQuery = "USP_FetchStateECHR '" + Navembid + "'";
            //select distinct a.HSRP_StateID, (select HSRPStateName from hsrpstate c where c.hsrp_stateid=a.hsrp_stateid) as HSRPStateName,(select HSRPStateShortName from hsrpstate d where d.hsrp_stateid=a.hsrp_stateid) as HSRPStateShortName, d.navembcode from hsrprecords a join DealerAffixation d  on  a.Affix_Id=d.SubDealerId and isnull(NewPdfRunningNo,'') = '' and  a.HSRP_StateID =='" + ddlStateName.SelectedValue + "' and d.Navembcode not like '%CODO%'  and isnull(erpassigndate,'') != ''   AND d.navembcode='" + Navembid + "' order by  a.HSRP_StateID";
            DataTable dtSE = Utils.Utils.GetDataTable(stateECQuery, CnnString);
            if (dtSE.Rows.Count > 0)
            {
                foreach (DataRow dr in dtSE.Rows)
                {
                    #region

                    string HSRP_StateID = dr["HSRP_StateID"].ToString().Trim();
                    string HSRPStateName = dr["HSRPStateName"].ToString().Trim();
                    string HSRPStateShortName = dr["HSRPStateShortName"].ToString().Trim();
                    string Navembcode = dr["Navembcode"].ToString().Trim();

                    string dir = dirPath + DateTime.Now.ToString("yyyy-MM-dd") + "\\" + HSRPStateShortName + "\\";
                    //string dir = dirPath + "/" + DateTime.Now.ToString("yyyy-MM-dd") + "/" + HSRPStateShortName + "/";
                    //string fileName = filePrefix + "-" + Navembcode + ".pdf";
                    string fileName = "MultiBrandHR" + "-" + filePrefix + "-" + Navembcode + ".pdf";
                    string filePath = dir + fileName;

                    //string folderpath = ConfigurationManager.AppSettings["InvoiceFolder"].ToString() + "/" + FinYear + "/" + oemid + "/" + HSRPStateID + "/";

                    StringBuilder html = new StringBuilder();

                    Boolean findRecord = false;
                    string strProductionSheetNo = string.Empty;

                    /*
                     *  Start body & HTMl Tag
                     */
                    #region
                    html.Append(
                        "<!DOCTYPE html>" +
                        "<html>" +
                        "<head>" +
                            "<meta charset='UTF-8'><title>Title</title>" +
                            "<style>" +
                                "@page {" +
                                    /* headers*/
                                    "@top-left {" +
                                        "content: 'Left header';" +
                                    "}" +
                                    "@top-right {" +
                                        "content: 'Right header';" +
                                    "}" +

                                    /* footers */
                                    "@bottom-left {" +
                                        "content: 'Lorem ipsum';" +
                                    "} " +
                                    "@bottom-right {" +
                                        "content: 'Page ' counter(page) ' of ' counter(pages);" +
                                    "}" +
                                    "@bottom-center  {" +
                                        "content:element(footer);" +
                                    "}" +
                                "}" +
                                 "#footer {" +
                                    "position: running(footer);" +
                                "}" +
                                "table {" +
                                  "border-collapse: collapse;" +
                                "}" +

                                "table, th, td {" +
                                    "border: 1px solid black;" +
                                    "text-align: left;" +
                                    "vertical-align: top;" +
                                    "padding:5px;" +
                                "}" +
                            "</style>" +
                        "</head>" +
                        "<body>");
                    #endregion

                    string oemDealerQuery = string.Empty;

                    oemDealerQuery = "USP_FetchMultiBrandDealeridHR '" + Navembid + "'";

                    #region
                    DataTable dtOD = Utils.Utils.GetDataTable(oemDealerQuery, CnnString);
                    if (dtOD.Rows.Count > 0)
                    {
                        foreach (DataRow drOD in dtOD.Rows)
                        {

                            string oemid = drOD["oemid"].ToString().Trim();
                            string dealerid = drOD["dealerid"].ToString().Trim();
                            string oemname = drOD["oemname"].ToString().Trim();
                            string dealername = drOD["dealername"].ToString().Trim();
                            string Address = drOD["Address"].ToString().Trim();
                            string strsubaffixid = drOD["SubDealerId"].ToString();

                            //start sql query
                            #region
                            string productionQuery = string.Empty;


                            DataTable dtProduction = new DataTable();

                            #endregion
                            //end sql query

                            #region

                            SqlConnection con = new SqlConnection(CnnString);
                            SqlCommand cmd = new SqlCommand("USP_MultiBrandDealerProductionSheetHR", con);
                            cmd.CommandType = CommandType.StoredProcedure;
                            con.Open();
                            cmd.Parameters.AddWithValue("@navembid", Navembcode);
                            cmd.Parameters.AddWithValue("@HSRP_StateID", HSRP_StateID);
                            cmd.Parameters.AddWithValue("@Dealerid", dealerid);
                            cmd.Parameters.AddWithValue("@Affix_Id", strsubaffixid);
                            SqlDataAdapter da = new SqlDataAdapter(cmd);
                            // dtProduction = new DataTable();
                            da.Fill(dtProduction);
                            con.Close();


                            // DataTable dtProduction = Utils.Utils.GetDataTable(productionQuery, CnnString);
                            if (dtProduction.Rows.Count > 0)
                            {

                                findRecord = true;
                                string strRunningNo = string.Empty;
                                //string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) as maxSheetNo from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + NAVEMBID + "'";
                                string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                                string strCom = Utils.Utils.getScalarValue(strSel, CnnString);
                                //DataTable dtSheetNo = fillDataTable(strSel, CnnString);
                                if (strCom.Equals(0) || strCom.Length == 0)
                                {
                                    strRunningNo = "0000001";
                                }
                                else
                                {
                                    strRunningNo = string.Format("{0:0000000}", Convert.ToInt32(strCom) + 1);
                                }
                                string strRequeNo = "select (prefixtext+right('00000'+ convert(varchar,Lastreqno+1),5)) as Reqno from prefix_Requisition  where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                                ReqNum = Utils.Utils.getScalarValue(strRequeNo, CnnString);
                                string strPRFIX = "select PrefixText from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                                string strPRFIXCom = Utils.Utils.getScalarValue(strPRFIX, CnnString);
                                strProductionSheetNo = strPRFIXCom + strRunningNo;

                                string RTOLocationName = dtProduction.Rows[0]["RTOLocationName"].ToString();

                                #region
                                html.Append("<div style='page-break-before: avoid;page-break-inside: avoid;page-break-after: always;'>" +
                                       "<table style='width:100%;border: 0px;'>" +
                                           "<tr style='border: 0px;'>" +
                                               "<td colspan='3' style='border: 0px;'>" +
                                                   "<div style='text-align:left'><b>Report Generation Date:</b> " + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</div>" +
                                               "</td>" +

                                                "<td colspan='3' style='border: 0px;'>" +
                                                   "<div style='text-align:left'><b>ROSMERTA SAFETY SYSTEMS LIMITED</b> " + "</div>" +
                                               "</td>" +


                                                "<td colspan='5' style='border: 0px;'>" +
                                                   "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Production Sheet No:</b> " + strRunningNo + " </div>" +
                                               "</td>" +

                                           "</tr>" +

                                             "<tr style='border: 0px;'>" +
                                               "<td colspan='3' style='border: 0px;'>" +
                                                   "<div style='text-align:left'><b>State:</b> " + HSRPStateName + "</div>" +
                                               "</td>" +

                                                "<td colspan='3' style='border: 0px;'>" +
                                                   "<div style='text-align:left'><b>Production Sheet</b> " + "</div>" +
                                               "</td>" +


                                                "<td colspan='5' style='border: 0px;'>" +
                                                   "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Dealer ID/Name:</b> " + dealerid + "/" + dealername + "/" + strsubaffixid + " </div>" +
                                               "</td>" +

                                           "</tr>" +

                                               "<tr style='border: 0px;'>" +
                                               "<td colspan='3' style='border: 0px;'>" +
                                                   "<div style='text-align:left'><b>EC Location: </b> " + RTOLocationName + "</div>" +
                                               "</td>" +

                                                "<td colspan='3' style='border: 0px;'>" +
                                                   "<div style='style='text-align:left'><b>Oem :</b> " + oemname + "</div>" +
                                               "</td>" +


                                                "<td colspan='5' style='border: 0px;'>" +
                                                    "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Dealer Address:</b> " + Address + " </div>" +
                                               //"<div style='text-align:left'><b>Dealer Address:</b> " + Address + " </div>" +
                                               "</td>" +

                                           "</tr>" +


 "<tr>" +
                                               "<td style='text-align:center'>Sr. No.</td>" +
                                                 "<td>Vehicle No.</td>" +
                                                  "<td>Front Plate Size</td>" +
                                               "<td>Front Laser No.</td>" +

                                               "<td>Rear Plate Size</td>" +
                                               "<td>Rear Laser No.</td>" +

                                               "<td>H. S. Foil </td>" +
                                               "<td>Caution Sticker</td>" +
                                               "<td>Fuel Type</td>" +
                                               "<td >VT</td>" +
                                               "<td >VC</ td>" +

                                           "</tr>");
                                #endregion

                                #region
                                foreach (DataRow drProduction in dtProduction.Rows)
                                {
                                    string HsrprecordID = drProduction["hsrprecordID"].ToString().Trim();
                                    string SRNo = drProduction["SRNo"].ToString().Trim();
                                    string ORD = "";// drProduction["ORD"].ToString().Trim();
                                    string VC = drProduction["VehicleClass"].ToString().Trim();
                                    string VehicleNo = drProduction["VehicleRegNo"].ToString().Trim();
                                    // string VehicleNo = "<td style='text-align:center;font-size:20px;'>" + "<b>" + drProduction["VehicleRegNo"].ToString().Trim() + "</b> </td>" +
                                    string VT = drProduction["VehicleType"].ToString().Trim();
                                    //string ChassisNo = drProduction["ChassisNo"].ToString().Trim();
                                    //string EngineNo = drProduction["EngineNo"].ToString().Trim();
                                    string FuelType = drProduction["FuelType"].ToString().Trim();
                                    string FrontPSize = drProduction["FrontProductCode"].ToString().Trim();
                                    // string FrontLaserNo = "<td style='text-align:center;font-size:20px;'>" + "<b>" + drProduction["HSRP_Front_LaserCode"].ToString().Trim() + "</b> </td>" +
                                    string FrontLaserNo = drProduction["HSRP_Front_LaserCode"].ToString().Trim();
                                    string RearPSize = drProduction["RearProductCode"].ToString().Trim();
                                    string RearLaserNo = drProduction["HSRP_Rear_LaserCode"].ToString().Trim();
                                    //string RearLaserNo = "<td style='text-align:center;font-size:20px;'>" + "<b>" + drProduction["HSRP_Rear_LaserCode"].ToString().Trim() + "</b> </td>" +
                                    string Amount = drProduction["roundoff_netamount"].ToString().Trim();
                                    //string OrderStatus = drProduction["OrderStatus"].ToString().Trim();
                                    string stickerColor = drProduction["stickerColor"].ToString().Trim();
                                    string HotStampingFoilColour = drProduction["HotStampingFoilColour"].ToString().Trim();

                                    html.Append("<tr>" +
                                   "<td style='text-align:center'>" + SRNo + "</td>" +
                                       "<td style='text-align:center;font-size:20px;'>" + "<b>" + VehicleNo + "</b> </td>" +

                                        "<td>" + FrontPSize + "</td>" +
                                         "<td style='text-align:center;font-size:20px;'>" + "<b>" + FrontLaserNo + "</b> </td>" +
                                           "<td>" + RearPSize + "</td>" +
                                         "<td style='text-align:center;font-size:20px;'>" + "<b>" + RearLaserNo + "</b> </td>" +

                                          "<td>" + HotStampingFoilColour + "</td>" +
                                           "<td>" + stickerColor + "</td>" +
                                            "<td>" + FuelType + "</td>" +
                                            "<td>" + VT + "</td>" +
                                            "<td>" + VC + "</td>" +





                               "</tr>");

                                    try
                                    {
                                        //start updating hsrprecords 
                                        string sqlUpdateHSRPRecords = "update HSRPRecords_HR set sendtoProductionStatus='Y', NAVPDFFlag='1', NewPdfRunningNo='" + strProductionSheetNo + "', Requisitionsheetno='" + ReqNum + "',  " +
                                              "PdfDownloadDate=GetDate(), pdfFileName='" + fileName + "', PDFDownloadUserID='1' where hsrprecordID='" + HsrprecordID + "' ";

                                        Utils.Utils.ExecNonQuery(sqlUpdateHSRPRecords, CnnString);   // uncomment after testing
                                    }
                                    catch (Exception ev)
                                    {
                                        Label1.Text = "hsrprecords update error: " + ev.Message;
                                    }
                                    //end 

                                }
                                #endregion
                                html.Append("</table>");

                                html.Append("</div>");

                                try
                                {
                                    string StrSqlUpdateECQuery = "update EmbossingCentersNew set NewProductionSheetRunningNo='" + strProductionSheetNo + "' " +
                                     "where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                                    Utils.Utils.ExecNonQuery(StrSqlUpdateECQuery, CnnString); // uncomment after testing
                                }
                                catch (Exception ev)
                                {
                                    Label1.Text = "EmbossingCentersNew update error: " + ev.Message;
                                }
                            }

                            #endregion

                        }
                    }// close oemDealerQuery
                    #endregion

                    #region "Req Generate"
                    string strComNew = string.Empty;
                    string strReqNumber = string.Empty;
                    string strReqNo = string.Empty;



                    string strSqlQuery1 = "select CompanyName from hsrpstate where hsrp_stateid='" + HSRP_StateID + "'";
                    strCompanyName = Utils.Utils.getScalarValue(strSqlQuery1, CnnString);

                    string strEMBName = " select EmbCenterName from EmbossingCentersNew where  State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                    strComNew = Utils.Utils.getScalarValue(strEMBName, CnnString);

                    if (strComNew != "")
                    {



                        strReqNo = "select (prefixtext+right('00000'+ convert(varchar,lastno+1),5)) as Reqno from prefix_Requisition  where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                        strReqNumber = Utils.Utils.getScalarValue(strReqNo, CnnString);

                        SQLString = "Exec [laserreqSlip1DashBoard]  '" + HSRP_StateID + "','" + Navembcode + "' ,  '" + ReqNum + "'";
                        DataTable dtResult = Utils.Utils.GetDataTable(SQLString, CnnString);

                        string strQuery = string.Empty;
                        string strRtoLocationName = string.Empty;
                        int Itotal = 0;

                        html.Append("<div style='width:100%;height:100%;'>" +
                                            "<table style='width:100%'>" +

                                                "<tr>" +
                                                    "<td colspan='12'>" +
                                                        "<div style='text-align:center;padding:8px;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + strCompanyName + "</b>" +

                                                            "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                             "<b style='font-size:20px;'>" + "MATERIAL REQUSITION NOTE" + "</b>" +

                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +

                                                "<tr>" +
                                                    "<td colspan='12'>" +
                                                        "<div style='text-align:center;padding:8px;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Production Sheet Date :" + DateTime.Now.ToString("dd-MM-yyyy") + "</b>" +
                                                            "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                             "<b style='font-size:20px;'>" + "" + "</b>" +
                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +
                                                "<tr>" +
                                                    "<td colspan='6'>" +
                                                        "<div style='text-align:left;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>REQ.NO:-</b>" +
                                                            "" + strReqNumber + "" +
                                                        "</div>" +
                                                    "</td>" +
                                                    "<td colspan='6'>" +
                                                        "<div style='text-align:left;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>Embossing Center:</b>" +
                                                            "" + strComNew + "" +
                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +
                                                "<tr>" +
                                                    "<td colspan='12'>" +
                                                        "<div style='text-align:left;padding:2px;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "State:" + HSRPStateName + "</b>" +
                                                            "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                             "<b style='font-size:20px;'>" + "" + "</b>" +
                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +
                                                 "<tr>" +
                                                    "<td colspan='12'>" +
                                                        "<div style='text-align:left;padding:8px;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + " " + "</b>" +
                                                            "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                             "<b style='font-size:20px;'>" + "" + "</b>" +
                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +

                                                "<tr>" +
                                                    "<td colspan='1' style='text-align:center'>SR.N.</td>" +
                                                    "<td colspan='3'>Product Size</td>" +
                                                    "<td colspan='1'>Laser Count</td>" +
                                                    "<td colspan='1'>Start Laser No</td>" +
                                                    "<td colspan='1'>End Laser No</td>" +

                                                "</tr>");


                        if (dtResult.Rows.Count > 0)
                        {
                            for (int i = 0; i < dtResult.Rows.Count; i++)
                            {
                                string ID = dtResult.Rows[i]["ID"].ToString();
                                string productcode = dtResult.Rows[i]["productcode"].ToString();
                                string LaserCount = dtResult.Rows[i]["LaserCount"].ToString();
                                Itotal = Convert.ToInt32(dtResult.Rows[i]["Total"].ToString());
                                string BeginLaser = dtResult.Rows[i]["BeginLaser"].ToString();
                                string EndLaser = dtResult.Rows[i]["EndLaser"].ToString();




                                html.Append("<tr>" +
                                   "<td colspan='1' style='text-align:left'>" + ID + "</td>" +
                                   "<td colspan='3'>" + productcode + "</td>" +

                                   "<td colspan='1'>" + LaserCount + "</td>" +
                                   "<td colspan='1'>" + BeginLaser + "</td>" +
                                   "<td colspan='1'>" + EndLaser + "</td>" +

                               "</tr>");
                            }
                        }
                        html.Append("<tr>" +
                                 "<td colspan='1' style='text-align:center' > " + "<b>Grand Total:</b>" + "</td>" +
                                 "<td colspan='3'>" + "" + "</td>" +

                                 "<td colspan='1'>" + Itotal + "</td>" +
                                 "<td colspan='1'>" + " " + "</td>" +
                                 "<td colspan='1'>" + " " + "</td>" +

                             "</tr>");




                        html.Append("<tr>" +
                         "<td colspan='2' > " + "<b>REQUESTED BY </b>" + "</td>" +
                         "<td colspan='2'>" + "<b>AUTHORIZED BY </b>" + "</td>" +

                         "<td colspan='2'>" + "<b>ISSUED BY </b>" + "</td>" +
                         "<td colspan='2'>" + "<b>RECEIVED BY</b>" + "</td>" +


                     "</tr>");

                        html.Append("<tr>" +
                                                  "<td colspan='12'>" +
                                                      "<div style='text-align:left;padding:2px;'>" +
                                                          "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Name" + "</b>" +

                                                      "</div>" +
                                                  "</td>" +
                                              "</tr>");

                        html.Append("<tr>" +
                                                 "<td colspan='12'>" +
                                                     "<div style='text-align:left;padding:2px;'>" +
                                                         "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Designation" + "</b>" +

                                                     "</div>" +
                                                 "</td>" +
                                             "</tr>");

                        html.Append("<tr>" +
                                               "<td colspan='12'>" +
                                                   "<div style='text-align:right;padding:8px;'>" +
                                                       "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Download By: Admin " + "</b>" +

                                                   "</div>" +
                                               "</td>" +
                                           "</tr>");

                        html.Append("<tr>" +
                                              "<td colspan='12'>" +
                                                  "<div style='text-align:right;padding:8px;'>" +
                                                      "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Sheet Generated By :" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</b>" +

                                                  "</div>" +
                                              "</td>" +
                                          "</tr>");







                        html.Append("</table>");

                        html.Append("</div>");

                        try
                        {
                            //start updating hsrprecords 
                            string Query = "update prefix_Requisition set lastno=lastno+1,Lastreqno=Lastreqno+1 where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                            Utils.Utils.ExecNonQuery(Query, CnnString);
                        }
                        catch (Exception ev)
                        {
                            Label1.Text = "prefix Requisition update error: " + ev.Message;
                        }
                    }


                    #endregion

                    /*
                     * Close body & HTMl Tag
                     */
                    html.Append("</body>" +
                        "</html>");


                    if (findRecord)
                    {

                        string strSaveSQlQuery = "insert into ProductionSheetAutoGenerated_List (HSRP_StateID, State_Code, RTO_ID, Emb_Center_Id, FileName, Productiondate, ProductionTime, OrderNo) values " +
                                    "('" + HSRP_StateID + "', '" + HSRPStateShortName + "', '', '" + Navembcode + "', '" + fileName + "', '" + DateTime.Now.ToString("yyyy-MM-dd") + "', '" + DateTime.Now.ToString("HH:mm:ss") + "', '" + filePrefix + "') ";
                        Utils.Utils.ExecNonQuery(strSaveSQlQuery, CnnString); // uncomment after testing
                        Label1.Text = Label1.Text + "successfully created production sheet of State: " + HSRPStateShortName + ", , Emb_Center_Id: " + Navembcode + Environment.NewLine;

                        //lblLog.Text = lblLog.Text + Environment.NewLine + "Start Production Query at:" + DateTime.Now;
                        #region
                        try
                        {
                            if (!Directory.Exists(dir))
                            {


                                Directory.CreateDirectory(dir);

                            }

                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        try
                        {
                            var pdf = Pdf
                                    .From(html.ToString())
                                    .OfSize(PaperSize.A4)
                                    .WithTitle("Title")
                                    .WithoutOutline()
                                    .WithMargins(1.25.Centimeters())
                                    .Landscape()
                                    .Comressed()
                                    .Content();

                            FileStream readStream = File.Create(filePath);
                            BinaryWriter binaryWriter = new BinaryWriter(readStream);

                            // Write the binary data to the file
                            binaryWriter.Write(pdf);
                            binaryWriter.Close();
                            readStream.Close();
                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        #endregion
                    }

                    #endregion

                }//close foreach stateEcQuery
            }
        }

        protected void BtnMulti_Click(object sender, EventArgs e)
        {
            btnmultiBrandDealer_Click(sender, e);
            btnmultiBrandHome_Click(sender, e);
        }

        private void HRMultiBrandHome()
        {
            string stateECQuery = string.Empty;
            string ReqNum = string.Empty;
            //string Navembid = Session["Navembid"].ToString();
            FillUserDetails();
            //stateECQuery = "USP_FetchStateECHRHome '" + Navembid + "'";
            stateECQuery = "select distinct a.HSRP_StateID, (select HSRPStateName from hsrpstate c where c.hsrp_stateid=a.hsrp_stateid) as HSRPStateName,(select HSRPStateShortName from hsrpstate d where d.hsrp_stateid=a.hsrp_stateid) as HSRPStateShortName, navembcode from hsrprecords_HR a with(nolock),rtolocation b where a.HSRP_StateID ='" + ddlStateName.SelectedValue + "'     and a.rtolocationid=b.rtolocationid and NewPdfRunningNo is null and erpassigndate is not null and OrderStatus='New Order' AND b.NAVEMBID='" + Navembid + "' and  OLDOrderid=2   order by  a.HSRP_StateID ";
            //select distinct a.HSRP_StateID, (select HSRPStateName from hsrpstate c where c.hsrp_stateid=a.hsrp_stateid) as HSRPStateName,(select HSRPStateShortName from hsrpstate d where d.hsrp_stateid=a.hsrp_stateid) as HSRPStateShortName, d.navembcode from hsrprecords a join DealerAffixation d  on  a.Affix_Id=d.SubDealerId and isnull(NewPdfRunningNo,'') = '' and  a.HSRP_StateID =='" + ddlStateName.SelectedValue + "' and d.Navembcode not like '%CODO%'  and isnull(erpassigndate,'') != ''   AND d.navembcode='" + Navembid + "' order by  a.HSRP_StateID";
            DataTable dtSE = Utils.Utils.GetDataTable(stateECQuery, CnnString);
            if (dtSE.Rows.Count > 0)
            {
                foreach (DataRow dr in dtSE.Rows)
                {
                    #region

                    string HSRP_StateID = dr["HSRP_StateID"].ToString().Trim();
                    string HSRPStateName = dr["HSRPStateName"].ToString().Trim();
                    string HSRPStateShortName = dr["HSRPStateShortName"].ToString().Trim();
                    string Navembcode = dr["Navembcode"].ToString().Trim();

                    string dir = dirPath + DateTime.Now.ToString("yyyy-MM-dd") + "\\" + HSRPStateShortName + "\\";
                    //string dir = dirPath + "/" + DateTime.Now.ToString("yyyy-MM-dd") + "/" + HSRPStateShortName + "/";
                    //string fileName = filePrefix + "-" + Navembcode + ".pdf";
                    string fileName = "MultiBrandHomeHR" + "-" + filePrefix + "-" + Navembcode + ".pdf";
                    string filePath = dir + fileName;

                    //string folderpath = ConfigurationManager.AppSettings["InvoiceFolder"].ToString() + "/" + FinYear + "/" + oemid + "/" + HSRPStateID + "/";

                    StringBuilder html = new StringBuilder();

                    Boolean findRecord = false;
                    string strProductionSheetNo = string.Empty;

                    /*
                     *  Start body & HTMl Tag
                     */
                    #region
                    html.Append(
                        "<!DOCTYPE html>" +
                        "<html>" +
                        "<head>" +
                            "<meta charset='UTF-8'><title>Title</title>" +
                            "<style>" +
                                "@page {" +
                                    /* headers*/
                                    "@top-left {" +
                                        "content: 'Left header';" +
                                    "}" +
                                    "@top-right {" +
                                        "content: 'Right header';" +
                                    "}" +

                                    /* footers */
                                    "@bottom-left {" +
                                        "content: 'Lorem ipsum';" +
                                    "} " +
                                    "@bottom-right {" +
                                        "content: 'Page ' counter(page) ' of ' counter(pages);" +
                                    "}" +
                                    "@bottom-center  {" +
                                        "content:element(footer);" +
                                    "}" +
                                "}" +
                                 "#footer {" +
                                    "position: running(footer);" +
                                "}" +
                                "table {" +
                                  "border-collapse: collapse;" +
                                "}" +

                                "table, th, td {" +
                                    "border: 1px solid black;" +
                                    "text-align: left;" +
                                    "vertical-align: top;" +
                                    "padding:5px;" +
                                "}" +
                            "</style>" +
                        "</head>" +
                        "<body>");
                    #endregion

                    string oemDealerQuery = string.Empty;

                    // oemDealerQuery = "USP_FetchMultiBrandDealeridHR '" + Navembid + "'";

                    oemDealerQuery = "USP_FetchMultiBrandHomeHR '" + ddlStateName.SelectedValue + "','" + Navembid + "'";

                    #region
                    DataTable dtOD = Utils.Utils.GetDataTable(oemDealerQuery, CnnString);
                    if (dtOD.Rows.Count > 0)
                    {
                        foreach (DataRow drOD in dtOD.Rows)
                        {

                            string oemid = drOD["oemid"].ToString().Trim();
                            string dealerid = drOD["dealerid"].ToString().Trim();
                            string oemname = drOD["oemname"].ToString().Trim();
                            string dealername = drOD["dealername"].ToString().Trim();
                            string Address = drOD["Address"].ToString().Trim();
                            //string strsubaffixid = drOD["SubDealerId"].ToString();

                            //start sql query
                            #region
                            string productionQuery = string.Empty;


                            DataTable dtProduction = new DataTable();

                            #endregion
                            //end sql query

                            #region

                            SqlConnection con = new SqlConnection(CnnString);
                            SqlCommand cmd = new SqlCommand("USP_HRMultiBrandHomeProductionSheet", con);
                            cmd.CommandType = CommandType.StoredProcedure;
                            con.Open();
                            cmd.Parameters.AddWithValue("@navembid", Navembcode);
                            cmd.Parameters.AddWithValue("@HSRP_StateID", HSRP_StateID);
                            cmd.Parameters.AddWithValue("@Dealerid", dealerid);
                            //cmd.Parameters.AddWithValue("@Affix_Id", strsubaffixid);
                            SqlDataAdapter da = new SqlDataAdapter(cmd);
                            // dtProduction = new DataTable();
                            da.Fill(dtProduction);
                            con.Close();


                            // DataTable dtProduction = Utils.Utils.GetDataTable(productionQuery, CnnString);
                            if (dtProduction.Rows.Count > 0)
                            {

                                findRecord = true;
                                string strRunningNo = string.Empty;
                                //string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) as maxSheetNo from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + NAVEMBID + "'";
                                string strSel = "select isnull(max(right(newProductionSheetRunningNo,7)),0000000) from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                                string strCom = Utils.Utils.getScalarValue(strSel, CnnString);
                                //DataTable dtSheetNo = fillDataTable(strSel, CnnString);
                                if (strCom.Equals(0) || strCom.Length == 0)
                                {
                                    strRunningNo = "0000001";
                                }
                                else
                                {
                                    strRunningNo = string.Format("{0:0000000}", Convert.ToInt32(strCom) + 1);
                                }
                                string strRequeNo = "select (prefixtext+right('00000'+ convert(varchar,Lastreqno+1),5)) as Reqno from prefix_Requisition  where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                                ReqNum = Utils.Utils.getScalarValue(strRequeNo, CnnString);
                                string strPRFIX = "select PrefixText from EmbossingCentersNew where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                                string strPRFIXCom = Utils.Utils.getScalarValue(strPRFIX, CnnString);
                                strProductionSheetNo = strPRFIXCom + strRunningNo;

                                string RTOLocationName = dtProduction.Rows[0]["RTOLocationName"].ToString();

                                #region
                                html.Append("<div style='page-break-before: avoid;page-break-inside: avoid;page-break-after: always;'>" +
                                       "<table style='width:100%;border: 0px;'>" +
                                           "<tr style='border: 0px;'>" +
                                               "<td colspan='3' style='border: 0px;'>" +
                                                   "<div style='text-align:left'><b>Report Generation Date:</b> " + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</div>" +
                                               "</td>" +

                                                "<td colspan='3' style='border: 0px;'>" +
                                                   "<div style='text-align:left'><b>ROSMERTA SAFETY SYSTEMS LIMITED</b> " + "</div>" +
                                               "</td>" +


                                                "<td colspan='5' style='border: 0px;'>" +
                                                   "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Production Sheet No:</b> " + strRunningNo + " </div>" +
                                               "</td>" +

                                           "</tr>" +

                                             "<tr style='border: 0px;'>" +
                                               "<td colspan='3' style='border: 0px;'>" +
                                                   "<div style='text-align:left'><b>State:</b> " + HSRPStateName + "</div>" +
                                               "</td>" +

                                                "<td colspan='3' style='border: 0px;'>" +
                                                   "<div style='text-align:left'><b>Production Sheet</b> " + "</div>" +
                                               "</td>" +


                                                "<td colspan='5' style='border: 0px;'>" +
                                                   "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Dealer ID/Name:</b> " + dealerid + "/" + dealername + "/" + " </div>" +
                                               "</td>" +

                                           "</tr>" +

                                               "<tr style='border: 0px;'>" +
                                               "<td colspan='3' style='border: 0px;'>" +
                                                   "<div style='text-align:left'><b>EC Location: </b> " + RTOLocationName + "</div>" +
                                               "</td>" +

                                                "<td colspan='3' style='border: 0px;'>" +
                                                   "<div style='style='text-align:left'><b>Oem :</b> " + oemname + "</div>" +
                                               "</td>" +


                                                "<td colspan='5' style='border: 0px;'>" +
                                                    "<div style='float:left;width: 500px;word-wrap: break-word;'><b>Dealer Address:</b> " + Address + " </div>" +
                                               //"<div style='text-align:left'><b>Dealer Address:</b> " + Address + " </div>" +
                                               "</td>" +

                                           "</tr>" +


 "<tr>" +
                                               "<td style='text-align:center'>Sr. No.</td>" +
                                                 "<td>Vehicle No.</td>" +
                                                  "<td>Front Plate Size</td>" +
                                               "<td>Front Laser No.</td>" +

                                               "<td>Rear Plate Size</td>" +
                                               "<td>Rear Laser No.</td>" +

                                               "<td>H. S. Foil </td>" +
                                               "<td>Caution Sticker</td>" +
                                               "<td>Fuel Type</td>" +
                                               "<td >VT</td>" +
                                               "<td >VC</ td>" +

                                           "</tr>");
                                #endregion

                                #region
                                foreach (DataRow drProduction in dtProduction.Rows)
                                {
                                    string HsrprecordID = drProduction["hsrprecordID"].ToString().Trim();
                                    string SRNo = drProduction["SRNo"].ToString().Trim();
                                    string ORD = "";// drProduction["ORD"].ToString().Trim();
                                    string VC = drProduction["VehicleClass"].ToString().Trim();
                                    string VehicleNo = drProduction["VehicleRegNo"].ToString().Trim();
                                    // string VehicleNo = "<td style='text-align:center;font-size:20px;'>" + "<b>" + drProduction["VehicleRegNo"].ToString().Trim() + "</b> </td>" +
                                    string VT = drProduction["VehicleType"].ToString().Trim();
                                    //string ChassisNo = drProduction["ChassisNo"].ToString().Trim();
                                    //string EngineNo = drProduction["EngineNo"].ToString().Trim();
                                    string FuelType = drProduction["FuelType"].ToString().Trim();
                                    string FrontPSize = drProduction["FrontProductCode"].ToString().Trim();
                                    // string FrontLaserNo = "<td style='text-align:center;font-size:20px;'>" + "<b>" + drProduction["HSRP_Front_LaserCode"].ToString().Trim() + "</b> </td>" +
                                    string FrontLaserNo = drProduction["HSRP_Front_LaserCode"].ToString().Trim();
                                    string RearPSize = drProduction["RearProductCode"].ToString().Trim();
                                    string RearLaserNo = drProduction["HSRP_Rear_LaserCode"].ToString().Trim();
                                    //string RearLaserNo = "<td style='text-align:center;font-size:20px;'>" + "<b>" + drProduction["HSRP_Rear_LaserCode"].ToString().Trim() + "</b> </td>" +
                                    string Amount = drProduction["roundoff_netamount"].ToString().Trim();
                                    //string OrderStatus = drProduction["OrderStatus"].ToString().Trim();
                                    string stickerColor = drProduction["stickerColor"].ToString().Trim();
                                    string HotStampingFoilColour = drProduction["HotStampingFoilColour"].ToString().Trim();

                                    html.Append("<tr>" +
                                   "<td style='text-align:center'>" + SRNo + "</td>" +
                                       "<td style='text-align:center;font-size:20px;'>" + "<b>" + VehicleNo + "</b> </td>" +

                                        "<td>" + FrontPSize + "</td>" +
                                         "<td style='text-align:center;font-size:20px;'>" + "<b>" + FrontLaserNo + "</b> </td>" +
                                           "<td>" + RearPSize + "</td>" +
                                         "<td style='text-align:center;font-size:20px;'>" + "<b>" + RearLaserNo + "</b> </td>" +

                                          "<td>" + HotStampingFoilColour + "</td>" +
                                           "<td>" + stickerColor + "</td>" +
                                            "<td>" + FuelType + "</td>" +
                                            "<td>" + VT + "</td>" +
                                            "<td>" + VC + "</td>" +





                               "</tr>");

                                    try
                                    {
                                        //start updating hsrprecords 
                                        string sqlUpdateHSRPRecords = "update HSRPRecords_HR set sendtoProductionStatus='Y', NAVPDFFlag='1', NewPdfRunningNo='" + strProductionSheetNo + "', Requisitionsheetno='" + ReqNum + "',  " +
                                              "PdfDownloadDate=GetDate(), pdfFileName='" + fileName + "', PDFDownloadUserID='1' where hsrprecordID='" + HsrprecordID + "' ";

                                        Utils.Utils.ExecNonQuery(sqlUpdateHSRPRecords, CnnString);   // uncomment after testing
                                    }
                                    catch (Exception ev)
                                    {
                                        Label1.Text = "hsrprecords update error: " + ev.Message;
                                    }
                                    //end 

                                }
                                #endregion
                                html.Append("</table>");

                                html.Append("</div>");

                                try
                                {
                                    string StrSqlUpdateECQuery = "update EmbossingCentersNew set NewProductionSheetRunningNo='" + strProductionSheetNo + "' " +
                                     "where State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                                    Utils.Utils.ExecNonQuery(StrSqlUpdateECQuery, CnnString); // uncomment after testing
                                }
                                catch (Exception ev)
                                {
                                    Label1.Text = "EmbossingCentersNew update error: " + ev.Message;
                                }
                            }

                            #endregion

                        }
                    }// close oemDealerQuery
                    #endregion

                    #region "Req Generate"
                    string strComNew = string.Empty;
                    string strReqNumber = string.Empty;
                    string strReqNo = string.Empty;



                    string strSqlQuery1 = "select CompanyName from hsrpstate where hsrp_stateid='" + HSRP_StateID + "'";
                    strCompanyName = Utils.Utils.getScalarValue(strSqlQuery1, CnnString);

                    string strEMBName = " select EmbCenterName from EmbossingCentersNew where  State_Id='" + HSRP_StateID + "' and Emb_Center_Id='" + Navembcode + "'";
                    strComNew = Utils.Utils.getScalarValue(strEMBName, CnnString);

                    if (strComNew != "")
                    {



                        strReqNo = "select (prefixtext+right('00000'+ convert(varchar,lastno+1),5)) as Reqno from prefix_Requisition  where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                        strReqNumber = Utils.Utils.getScalarValue(strReqNo, CnnString);

                        SQLString = "Exec [laserreqSlip1DashBoard]  '" + HSRP_StateID + "','" + Navembcode + "' ,  '" + ReqNum + "'";
                        DataTable dtResult = Utils.Utils.GetDataTable(SQLString, CnnString);

                        string strQuery = string.Empty;
                        string strRtoLocationName = string.Empty;
                        int Itotal = 0;

                        html.Append("<div style='width:100%;height:100%;'>" +
                                            "<table style='width:100%'>" +

                                                "<tr>" +
                                                    "<td colspan='12'>" +
                                                        "<div style='text-align:center;padding:8px;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + strCompanyName + "</b>" +

                                                            "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                             "<b style='font-size:20px;'>" + "MATERIAL REQUSITION NOTE" + "</b>" +

                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +

                                                "<tr>" +
                                                    "<td colspan='12'>" +
                                                        "<div style='text-align:center;padding:8px;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Production Sheet Date :" + DateTime.Now.ToString("dd-MM-yyyy") + "</b>" +
                                                            "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                             "<b style='font-size:20px;'>" + "" + "</b>" +
                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +
                                                "<tr>" +
                                                    "<td colspan='6'>" +
                                                        "<div style='text-align:left;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>REQ.NO:-</b>" +
                                                            "" + strReqNumber + "" +
                                                        "</div>" +
                                                    "</td>" +
                                                    "<td colspan='6'>" +
                                                        "<div style='text-align:left;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>Embossing Center:</b>" +
                                                            "" + strComNew + "" +
                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +
                                                "<tr>" +
                                                    "<td colspan='12'>" +
                                                        "<div style='text-align:left;padding:2px;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "State:" + HSRPStateName + "</b>" +
                                                            "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                             "<b style='font-size:20px;'>" + "" + "</b>" +
                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +
                                                 "<tr>" +
                                                    "<td colspan='12'>" +
                                                        "<div style='text-align:left;padding:8px;'>" +
                                                            "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + " " + "</b>" +
                                                            "<p style='margin-top:2px;margin-bottom:2px;'>" + " " + "</p>" +
                                                             "<b style='font-size:20px;'>" + "" + "</b>" +
                                                        "</div>" +
                                                    "</td>" +
                                                "</tr>" +

                                                "<tr>" +
                                                    "<td colspan='1' style='text-align:center'>SR.N.</td>" +
                                                    "<td colspan='3'>Product Size</td>" +
                                                    "<td colspan='1'>Laser Count</td>" +
                                                    "<td colspan='1'>Start Laser No</td>" +
                                                    "<td colspan='1'>End Laser No</td>" +

                                                "</tr>");


                        if (dtResult.Rows.Count > 0)
                        {
                            for (int i = 0; i < dtResult.Rows.Count; i++)
                            {
                                string ID = dtResult.Rows[i]["ID"].ToString();
                                string productcode = dtResult.Rows[i]["productcode"].ToString();
                                string LaserCount = dtResult.Rows[i]["LaserCount"].ToString();
                                Itotal = Convert.ToInt32(dtResult.Rows[i]["Total"].ToString());
                                string BeginLaser = dtResult.Rows[i]["BeginLaser"].ToString();
                                string EndLaser = dtResult.Rows[i]["EndLaser"].ToString();




                                html.Append("<tr>" +
                                   "<td colspan='1' style='text-align:left'>" + ID + "</td>" +
                                   "<td colspan='3'>" + productcode + "</td>" +

                                   "<td colspan='1'>" + LaserCount + "</td>" +
                                   "<td colspan='1'>" + BeginLaser + "</td>" +
                                   "<td colspan='1'>" + EndLaser + "</td>" +

                               "</tr>");
                            }
                        }
                        html.Append("<tr>" +
                                 "<td colspan='1' style='text-align:center' > " + "<b>Grand Total:</b>" + "</td>" +
                                 "<td colspan='3'>" + "" + "</td>" +

                                 "<td colspan='1'>" + Itotal + "</td>" +
                                 "<td colspan='1'>" + " " + "</td>" +
                                 "<td colspan='1'>" + " " + "</td>" +

                             "</tr>");




                        html.Append("<tr>" +
                         "<td colspan='2' > " + "<b>REQUESTED BY </b>" + "</td>" +
                         "<td colspan='2'>" + "<b>AUTHORIZED BY </b>" + "</td>" +

                         "<td colspan='2'>" + "<b>ISSUED BY </b>" + "</td>" +
                         "<td colspan='2'>" + "<b>RECEIVED BY</b>" + "</td>" +


                     "</tr>");

                        html.Append("<tr>" +
                                                  "<td colspan='12'>" +
                                                      "<div style='text-align:left;padding:2px;'>" +
                                                          "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Name" + "</b>" +

                                                      "</div>" +
                                                  "</td>" +
                                              "</tr>");

                        html.Append("<tr>" +
                                                 "<td colspan='12'>" +
                                                     "<div style='text-align:left;padding:2px;'>" +
                                                         "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Designation" + "</b>" +

                                                     "</div>" +
                                                 "</td>" +
                                             "</tr>");

                        html.Append("<tr>" +
                                               "<td colspan='12'>" +
                                                   "<div style='text-align:right;padding:8px;'>" +
                                                       "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Download By: Admin " + "</b>" +

                                                   "</div>" +
                                               "</td>" +
                                           "</tr>");

                        html.Append("<tr>" +
                                              "<td colspan='12'>" +
                                                  "<div style='text-align:right;padding:8px;'>" +
                                                      "<b style='font-size:20px;margin-top:2px;margin-bottom:2px;'>" + "Sheet Generated By :" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "</b>" +

                                                  "</div>" +
                                              "</td>" +
                                          "</tr>");







                        html.Append("</table>");

                        html.Append("</div>");

                        try
                        {
                            //start updating hsrprecords 
                            string Query = "update prefix_Requisition set lastno=lastno+1,Lastreqno=Lastreqno+1 where hsrp_stateid='" + HSRP_StateID + "' and prefixfor='Req No'";
                            Utils.Utils.ExecNonQuery(Query, CnnString);
                        }
                        catch (Exception ev)
                        {
                            Label1.Text = "prefix Requisition update error: " + ev.Message;
                        }
                    }


                    #endregion

                    /*
                     * Close body & HTMl Tag
                     */
                    html.Append("</body>" +
                        "</html>");


                    if (findRecord)
                    {

                        string strSaveSQlQuery = "insert into ProductionSheetAutoGenerated_List (HSRP_StateID, State_Code, RTO_ID, Emb_Center_Id, FileName, Productiondate, ProductionTime, OrderNo) values " +
                                    "('" + HSRP_StateID + "', '" + HSRPStateShortName + "', '', '" + Navembcode + "', '" + fileName + "', '" + DateTime.Now.ToString("yyyy-MM-dd") + "', '" + DateTime.Now.ToString("HH:mm:ss") + "', '" + filePrefix + "') ";
                        Utils.Utils.ExecNonQuery(strSaveSQlQuery, CnnString); // uncomment after testing
                        Label1.Text = Label1.Text + "successfully created production sheet of State: " + HSRPStateShortName + ", , Emb_Center_Id: " + Navembcode + Environment.NewLine;

                        //lblLog.Text = lblLog.Text + Environment.NewLine + "Start Production Query at:" + DateTime.Now;
                        #region
                        try
                        {
                            if (!Directory.Exists(dir))
                            {


                                Directory.CreateDirectory(dir);

                            }

                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        try
                        {
                            var pdf = Pdf
                                    .From(html.ToString())
                                    .OfSize(PaperSize.A4)
                                    .WithTitle("Title")
                                    .WithoutOutline()
                                    .WithMargins(1.25.Centimeters())
                                    .Landscape()
                                    .Comressed()
                                    .Content();

                            FileStream readStream = File.Create(filePath);
                            BinaryWriter binaryWriter = new BinaryWriter(readStream);

                            // Write the binary data to the file
                            binaryWriter.Write(pdf);
                            binaryWriter.Close();
                            readStream.Close();
                        }
                        catch (Exception ev)
                        {
                            // Fail silently
                            Label1.Text = ev.Message;
                        }

                        #endregion
                    }

                    #endregion

                }//close foreach stateEcQuery
            }
        }



    }
}