using System;
using System.Windows.Forms;
using NovaNet.Utils;
using System.IO;
using System.Globalization;
using System.Threading;

namespace Beleghata_ID_PDF_creator
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            string yr;
            string mn;
            string dd;
            string qry = string.Empty;
            NovaNet.Utils.dbCon dbcon;

            //txtLogger txLog = new txtLogger(Path.GetDirectoryName(Application.ExecutablePath), LogLevel.Beta);
            try
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                ///For changing regional settings

                Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US", false);
                Microsoft.Win32.Registry.SetValue(@"HKEY_CURRENT_USER\Control Panel\International", "sShortDate", "dd/MM/yyyy");
                Microsoft.Win32.Registry.SetValue(@"HKEY_CURRENT_USER\Control Panel\International", "sLongDate", "dd/MM/yyyy");
                ///

                string path = Path.GetDirectoryName(Application.ExecutablePath);
                if (File.Exists(path + "/prKey.snk"))
                {
                    string lic = Utils.Crypto.Decrypt((path + "/prKey.snk"), (path + "/EDMSLIC.ini"));
                    lic = lic.Substring(lic.Length - 16, 16);
                    if (lic != string.Empty)
                    {
                        string stDateTime = lic.Substring(0, 8);
                        string endDateTime = lic.Substring(8, 8);
                        string currDt = string.Empty;

                        yr = stDateTime.Substring(0, 4);
                        mn = stDateTime.Substring(4, 2);
                        dd = stDateTime.Substring(6, 2);
                        stDateTime = dd + "/" + mn + "/" + yr;

                        yr = endDateTime.Substring(0, 4);
                        mn = endDateTime.Substring(4, 2);
                        dd = endDateTime.Substring(6, 2);
                        endDateTime = dd + "/" + mn + "/" + yr;

                        IFormatProvider culture = new CultureInfo("fr-Fr", true);
                        DateTime stDt = DateTime.ParseExact(stDateTime, "dd/MM/yyyy", culture, DateTimeStyles.NoCurrentDateDefault);

                        DateTime endDt = DateTime.ParseExact(endDateTime, "dd/MM/yyyy", culture, DateTimeStyles.NoCurrentDateDefault);

                        //dbcon = new NovaNet.Utils.dbCon();
                        string sqlCon = DateTime.Now.ToString("dd/MM/yyyy");

                        //    //changed on 26/10/2009
                        //    //Get current date time from database server.
                        DateTime curDate = DateTime.ParseExact(sqlCon, "dd/MM/yyyy", culture, DateTimeStyles.NoCurrentDateDefault);
                        //qry = "create table if not EXISTS `deed_details_exception` (District_code varchar(10) NOT NULL ,RO_Code varchar(10),Book varchar(10),Deed_year varchar(10),Deed_no varchar(10),srl_no int(10),Exception varchar(25),PRIMARY KEY  (District_code,RO_Code,Book,Deed_year,Deed_no,srl_no) )";
                        //cmd = new OdbcCommand(qry, sqlCon);
                        //cmd.ExecuteNonQuery();
                        //cmd.Dispose();

                        //qry = "create table if not EXISTS `index_of_name_exception` (District_code varchar(10) NOT NULL ,RO_Code varchar(10),Book varchar(10),Deed_year varchar(10),Deed_no varchar(10),item_no int(10),srl_no int(10),Exception varchar(25),PRIMARY KEY  (District_code,RO_Code,Book,Deed_year,Deed_no,item_no,srl_no) )";
                        //cmd = new OdbcCommand(qry, sqlCon);
                        //cmd.ExecuteNonQuery();
                        //cmd.Dispose();

                        //qry = "create table if not EXISTS `index_of_property_exception` (District_code varchar(10) NOT NULL ,RO_Code varchar(10),Book varchar(10),Deed_year varchar(10),Deed_no varchar(10),item_no int(10),srl_no int(10),Exception varchar(25),PRIMARY KEY  (District_code,RO_Code,Book,Deed_year,Deed_no,item_no,srl_no) )";
                        //cmd = new OdbcCommand(qry, sqlCon);
                        //cmd.ExecuteNonQuery();
                        //cmd.Dispose();
                        //DataSet ds1 = new DataSet();
                        //qry = "SELECT * FROM information_schema.COLUMNS WHERE COLUMN_NAME='exception' AND TABLE_NAME='index_of_property'";
                        //sqlAdap = new OdbcDataAdapter(qry, sqlCon);
                        //sqlAdap.Fill(ds1);
                        //if (ds1.Tables[0].Rows.Count == 0)
                        //{
                        //    qry = "alter table index_of_property add column exception varchar(10)";
                        //    cmd = new OdbcCommand(qry, sqlCon);
                        //    cmd.ExecuteNonQuery();
                        //    cmd.Dispose();
                        //}

                        if ((stDt <= curDate) && (endDt >= curDate)) // check with license date time, chnaged on 26/10/2009
                        {

                            Application.Run(new Form1());
                        }
                        else
                        {
                            MessageBox.Show("License has been expired. Contact with nevaeh Technology");
                            Application.Exit();
                        }
                    }
                    else
                    {
                        MessageBox.Show("Invalid license. Contact with nevaeh Technology");
                        Application.Exit();
                    }
                }
                else
                {
                    MessageBox.Show("Invalid license. Contact with nevaeh Technology");
                    Application.Exit();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error while doing the operation...." + ex.Message);
            }
        }
    }
}
