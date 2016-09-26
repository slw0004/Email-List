# Email-List
this code updates the database at a button click and sends reminder emails to clients

/*********************************************************************************************************
*   File:           NoSolicitAutomate.ascx
*   Description:    Enable Updates to No Solicit database and send reminder emails
*   Date:           2016-03-03
*   Developer:     Samantha Maze
*   Revisions:      
*   Notes:          
*********************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using DotNetNuke;
using Telerik.Web.UI;
using Microsoft.Practices.EnterpriseLibrary.ExceptionHandling;
using System.Data;
using System.Text;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Collections;
using Nashville.Web;
using System.IO;
using System.Xml.Linq;
using System.Transactions;


using System.Net.Mail;
public partial class DesktopModules_NV___Forms_Global_NoSolicitUpdate_NoSolicitUpdate : DotNetNuke.Entities.Modules.PortalModuleBase
{
    public const string _FROMEMAIL = "noreply@nashville.gov";
    public const string _FORMCAT = "GLOBAL-No Solicit";
    public const string _SUBJECT = "No solicitations registration renewal";

    private string _ContentType;
    private string _ContentAge;
    public string ipaddress;

    #region Properties

    #endregion

    #region Event Handlers


    protected void btnUpdate_Click(object sender, EventArgs e)
    {
        UpdateDatabase(0);
        EmailTable();
    }
    protected void btnSendEmails_Click(object sender, EventArgs e)
    {
        SendEmail();
        UpdateEmailSent();
        UpdateDatabase(0);
        EmailTable();
    }

    #endregion

    #region Methods

    //This method updates all of the entries in Global_NoSolicitLookup to be set to inactive if they are over a year old.

    private bool UpdateDatabase(int form_Id)
    {
        bool isValid = true;
        try
        {
            using (SqlConnection conn = new SqlConnection(Common.forms_conn_string))
            {
                conn.Open();
                using (SqlCommand cmd =
                    new SqlCommand("UPDATE	Global_NoSolicitLookup SET form_IsActive = 0, form_UpdatedDate = getdate()" +
                                    " WHERE form_IsActive = 1 AND form_UpdatedDate < DATEADD(year,-1,getdate())", conn))
                {
                    int rows = cmd.ExecuteNonQuery();

                }
                conn.Close();

            }
        }
        catch (Exception ex)
        {
            isValid = false;
            throw ex;
        }
        return isValid;

    }
    //This method creates a table of all of the entries from Global_NoSolicit Lookup that need to have reminder emails sent
    private void EmailTable()
    {
        try
        {

            StringBuilder sbOutput = new StringBuilder();
            SqlDataReader sdrEmails;


            using (SqlConnection conn = new SqlConnection(Common.forms_conn_string))
            {
                conn.Open();
                using (SqlCommand cmd =
                    new SqlCommand("Select form_ID, form_Addr1, form_Email, form_AddedDate, form_EmailSent, DATEADD(year, 1, form_UpdatedDate) as 'ExpirationDate' FROM Global_NoSolicitLookup WHERE form_IsActive = 1" +
                    " and form_EmailSent < DATEADD(month, -11, getdate()) AND form_UpdatedDate < DATEADD(month,-11,getdate()) ", conn))
                {
                    int rows = cmd.ExecuteNonQuery();


                    sdrEmails = cmd.ExecuteReader();
                    sbOutput.Append("<h2>" + "Emails to be sent" + "</h2>\n");
                }
                
                if (sdrEmails.HasRows)  
                {
                    sbOutput.Append("<table class='datatable' style='width:100%'>\n");
                    sbOutput.Append("<tr>");
                    sbOutput.Append("<th>Address</th>");
                    sbOutput.Append("<th>Email</th>");
                    sbOutput.Append("<th>AddedDate</th>");
                    sbOutput.Append("<th>EmailSent</th>");
                    sbOutput.Append("<th>ExpirationDate</th>");

                    sbOutput.Append("</tr>");
                    while (sdrEmails.Read())
                    {
                        String idn = ((Int32)sdrEmails.GetSqlInt32(0)).ToString();
                        String eml = (String)sdrEmails.GetSqlString(2);
                        DateTime dte = (DateTime)sdrEmails.GetSqlDateTime(3);
                        DateTime dte2 = (DateTime)sdrEmails.GetSqlDateTime(4);
                        DateTime dte3 = (DateTime)sdrEmails.GetSqlDateTime(5);
                        
                      
                        sbOutput.Append("<tr>");
                        sbOutput.Append("<td>" + sdrEmails.GetSqlString(1) + "</td>");
                        sbOutput.Append("<td>" + eml + "</td>");
                        sbOutput.Append("<td>" + dte.ToString("MM/dd/yyyy HH:mm") + "</td>");
                        sbOutput.Append("<td>" + dte2.ToString("MM/dd/yyyy HH:mm") + "</td>");
                        sbOutput.Append("<td>" + dte3.ToString("MM/dd/yyyy HH:mm") + "</td>");
                        sbOutput.Append("<tr>");
                    }
                  
                    sbOutput.Append("</table>\n");
                    litEmails.Text = sbOutput.ToString();
                    
                }
                else
                {
                    sbOutput.Append("<p>There are no items to display.</p>");
                    litEmails.Text = sbOutput.ToString();
                }
                if (sdrEmails.HasRows)

                    btnSendEmails.Visible = true;

                else
                {
                    btnSendEmails.Visible = false;
                    conn.Close();
                    sdrEmails.Close();

                }
            }
        }
    
        catch (Exception ex){
            String strMessage = "<p><strong>Error in the recordset render.</strong></p>";
            strMessage += "<p>" + ex.Message + "</p>";
            litEmails.Text = strMessage;
        }
    
    }


    private bool SendEmail()
    {
        bool isValid = true;
         
            SqlDataReader sdrEmails;


            using (SqlConnection conn = new SqlConnection(Common.forms_conn_string))
            {
                conn.Open();
                using (SqlCommand cmd =
                    new SqlCommand("Select form_ID, form_Email, form_Addr1, form_Zip FROM Global_NoSolicitLookup WHERE form_IsActive = 1" +
                    " and form_EmailSent < DATEADD(month, -11, getdate()) AND form_UpdatedDate < DATEADD(month,-11,getdate()) ", conn))
                    {
                    int rows = cmd.ExecuteNonQuery();
                    sdrEmails = cmd.ExecuteReader();

                     while (sdrEmails.Read())
                    {
                string param1 = ((Int32)sdrEmails.GetSqlInt32(0)).ToString();
                string param2 = (String)sdrEmails.GetSqlString(1);
                string param3 = (String)sdrEmails.GetSqlString(2);
                string param4 = (String)sdrEmails.GetSqlString(3);
                DateTime dt1 = DateTime.Now;
                String dt2 = (dt1.AddMonths(1)).ToString("d");


                string emailbody = "Thank you for registering on the No Solicitation list for Nashville and Davidson County, Tennessee.  " +
                                   "Your registration will expire on " + dt2 + ". \n\nThe address information we have on record for this email address is:" +
                                   "\n\n" + param3 + "\n" + param4 + "\n\nTo automatically renew your registration for the upcoming" +
                                   " year, please visit https://www.nashville.gov/Metro-Clerk/Solicitation-Permits/No-Solicit-Confirm.aspx?ID=" + param1 +
                                   "\n\nIf you no longer want this address on the No Solicitation list, you do not need to take any action.  It will be"+
                                   " removed automatically after the expiration date listed above.\n\nIf you would like to submit another address, please visit " +
                                   "https://www.nashville.gov/Metro-Clerk/Solicitation-Permits/Address-Registry.aspx" + " .\n\nMetropolitan Clerk's Office" +
                                   "\nMetropolitan Courthouse\nSuite205\n1 Public Square\nNashville, TN 37201\n\nPhone: (615) 862-6770\nFax: (615)880-3733\nE-mail: metroclerk@nashville.gov" +
                                   "\nHours: 8:00 AM - 4:30 PM";
                long admin_queue_Id = 0;
                admin_queue_Id = Notifications.Enqueue(_FORMCAT, param2.ToString(), _FROMEMAIL, "", "", _SUBJECT, emailbody, false);
                }
                    conn.Close();
                }
        }
       return isValid;
    }

    private void UpdateEmailSent()
    {

                    using (SqlConnection conn = new SqlConnection(Common.forms_conn_string))
                    {
                        conn.Open();
                        using (SqlCommand cmd =
                            new SqlCommand("UPDATE	Global_NoSolicitLookup SET form_EmailSent = getdate() WHERE form_IsActive = 1" +
                            "and form_EmailSent < DATEADD(month, -11, getdate())  AND form_UpdatedDate < DATEADD(month,-11,getdate())  ", conn))
                        {
                            int rows = cmd.ExecuteNonQuery();
                        }
                        conn.Close();
   
    }
}
    #endregion
}


