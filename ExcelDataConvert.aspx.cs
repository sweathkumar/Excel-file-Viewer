using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.SqlClient;
using System.Configuration;

namespace ExcelFileToTableConvert
{
    public partial class ExcelDataConvert : System.Web.UI.Page
    {
        // connection
        OleDbConnection con;
        string constring;
        string constr = System.Configuration.ConfigurationManager.ConnectionStrings["myNewDatabase"].ConnectionString;

        protected void Page_Load(object sender, EventArgs e)
        {

        }
        //excel connection
        private void ExcelConn(string constring)
        {
            //connection = string.Format(@"Provider=Microsoft.ACE.OLEDB.8.0;Data Source={'"+ Path +"'};Extended Properties='Excel 8.0 Xml; HDR = No;'");
            con = new OleDbConnection(constring);
        }

        protected void btnUpload_Click(object sender, EventArgs e)
        {
            //checking if the file has been uploaded
            if (uploadFile.HasFile)
            {
                //try catch to handle exceptions
                try
                {
                    string path;
                    //creating upload path
                    path = Server.MapPath("~/Uploads/");
                    if (!Directory.Exists(path))
                    {
                        //creating a directory
                        Directory.CreateDirectory(path);
                    }
                    //getting file name
                    string filename = Path.GetFileName(uploadFile.FileName);
                    //replacing file with new name
                    //string newfile = filename.Replace(filename, "UpdatedEmandatefile.xls");
                    //saving file in uploads folder
                    uploadFile.SaveAs(Server.MapPath("~/Uploads/") + filename);
                    path = Server.MapPath("~/Uploads/" + filename);
                    //getting file extension
                    string extension = Path.GetExtension(uploadFile.PostedFile.FileName);
                    //emptying the connection string container
                    constring = string.Empty;

                    switch (extension)
                    {
                        case ".xls": // Excel 97-03
                            constring = $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={path};Extended Properties='Excel 8.0;HDR=YES';";
                            proceed();
                            break;
                        case ".xlsx": // Excel 07 or higher only
                            constring = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={path};Extended Properties='Excel 12.0 Xml;HDR=YES'";
                            proceed();
                            break;
                        default:
                            // This block executes if extension does not match .xls or .xlsx
                            ClientScript.RegisterStartupScript(GetType(), "alert", "alert('File must be in Xls or Xlsx format only!');", true);
                            break;
                    }

                }
                catch(Exception ex)
                {
                    //alert if failed to insert
                    ClientScript.RegisterStartupScript(GetType(), "alert", "alert('" + ex.Message.ToString() + "');", true);
                }
            }
        }

        protected void proceed()
        {
            //excel connection 
            ExcelConn(constring);
            //reads excel data
            DataTable dt2 = ReadExcelRecords();
            //inserting excel data
            insertDataIntoDatabase(dt2);
            //alert if sucessfully inserted
            ClientScript.RegisterStartupScript(GetType(), "alert", "alert('Excel Record entered into database Sucessfully!');", true);
        }

        //reads excel data method
        protected DataTable ReadExcelRecords()
        {
            //try catch to handle exceptions
            try
            {
                //opening connection
                con.Open();
                //selecting a sheet in excel
                string sheet1 = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null).Rows[0]["TABLE_NAME"].ToString();
                //creating a new datatable
                DataTable dt = new DataTable();
                //creating columns for data table
                dt.Columns.AddRange(new DataColumn[5] {
                    new DataColumn("userid", typeof(int)),
                    new DataColumn("Name", typeof(string)),
                    new DataColumn("email", typeof(string)),
                    new DataColumn("date", typeof(DateTime)),
                    new DataColumn("phone", typeof(string))
                });

                //retriving data from excel
                OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [" + sheet1 + "]", con);

                //filling table into data table
                da.Fill(dt);
                //ittrating each row
                foreach (DataRow row in dt.Rows)
                {
                    // getting date&time
                    DateTime orgDateTime = Convert.ToDateTime(row["date"]);

                    // format date and time 
                    string formatedDate = orgDateTime.ToString("yyyy-MM-dd hh:mm:ss tt");

                    // Assign the formatted date string back to the "date" column in the DataRow
                    row["date"] = formatedDate;
                }

                //returning table
                return dt;

            }
            //exception
            catch(Exception ex)
            {
                //alert 
                ClientScript.RegisterStartupScript(GetType(), "alert", "alert('" + ex.Message.ToString() + "');", true);
                return null;
            }
        }

        protected void insertDataIntoDatabase(DataTable dt)
        {
            try
            {
                //creating connection
                SqlConnection con = new SqlConnection(constr);
                //bulk upload creating
                SqlBulkCopy bc = new SqlBulkCopy(con);
                //bul upload table calling  
                bc.DestinationTableName = "dbo.DataEntry";
                //bulk upload column maping
                bc.ColumnMappings.Add("userid", "userid");
                bc.ColumnMappings.Add("name", "name");
                bc.ColumnMappings.Add("email", "email");
                bc.ColumnMappings.Add("date", "date");
                bc.ColumnMappings.Add("phone", "phone");
                //opening sql connection
                con.Open();
                //writing to server
                bc.WriteToServer(dt);
                //closing sql connection
                con.Close();
            }
            catch(Exception ex)
            {
                //alert 
                ClientScript.RegisterStartupScript(GetType(), "alert", "alert('" + ex.Message.ToString() + "');", true);
            }
        }

        protected void btnView_Click(object sender, EventArgs e)
        {
            try {
                string qry = "SELECT * FROM DATAENTRY";
                SqlConnection con = new SqlConnection(constr);
                con.Open();
                SqlCommand cmd = new SqlCommand(qry, con);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                
                da.Fill(dt);
                con.Close();

                repeaterView.DataSource = dt;
                repeaterView.DataBind();
            }
            catch(Exception ex)
            {
                //alert 
                ClientScript.RegisterStartupScript(GetType(), "alert", "alert('" + ex.Message.ToString() + "');", true);
            }
        }
    }
}