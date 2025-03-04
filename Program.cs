using Newtonsoft.Json;
using SAPbobsCOM;
using Serilog;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CT_AutoRecon
{

    class Program
    {
       

        static void Main(string[] args)
        {

            try
            {
                #region Logger Setup
                //This will give us the full name path of the executable file:
                //i.e. C:\Program Files\MyApplication\MyApplication.exe
                string strExeFilePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                //This will strip just the working path name:
                //C:\Program Files\MyApplication
                string strWorkPath = System.IO.Path.GetDirectoryName(strExeFilePath);

                Log.Logger = new LoggerConfiguration()
                                // add console as logging target
                                .WriteTo.Console()

                                .WriteTo.File(@"" + strWorkPath + "\\logs\\" + DateTime.Now.Year + "\\" + DateTime.Now.Month + "\\" + "sap-.logs",
                                              rollingInterval: RollingInterval.Day)

                                // set default minimum level
                                .MinimumLevel.Debug()
                                .CreateLogger();

                #endregion

                Log.Information("[==================== STARTING APPLICATION " + Assembly.GetExecutingAssembly().GetCustomAttribute<AssemblyTitleAttribute>().Title + " " + Assembly.GetExecutingAssembly().GetCustomAttribute<AssemblyDescriptionAttribute>().Description + " Version " + Assembly.GetExecutingAssembly().GetName().Version.ToString() + "   " + DateTime.Now + "==================]");
                Log.Information("Loading DB Details JSON File");
                using (StreamReader r = new StreamReader("DBConfigurations.json"))
                {
                    Log.Information("DB Details File Found!!");
                    string jsonFileData = r.ReadToEnd();
                    if (jsonFileData != "")
                    {

                        List<DBDetails> dbItems = JsonConvert.DeserializeObject<List<DBDetails>>(jsonFileData);
                        foreach (DBDetails dbData in dbItems)
                        {
                            Log.Information($"Sync Started for the database : {dbData.CompanyDB} !!");
                            SAPbobsCOM.Company oCompany = new SAPbobsCOM.Company();
                            int ret = SetCompanyConnection(oCompany,dbData);

                            #region CUSTOMIZATION TOOLS

                            Log.Information("=================== CUSTOMIZATION TOOLS STARTED================");

                            try
                            {
                                if (ret == 0 && ConfigurationManager.AppSettings["SwitchUserDefined"] == "1")
                                {
                                    // GlobalFunctions.CustomizationTools(oCompany);

                                }

                            }
                            catch (Exception ex)
                            {
                                Log.Error(ex.Message);

                            }


                            Log.Information("=================== CUSTOMIZATION TOOLS ENDED================");

                            #endregion

                            #region Auto Cancel Recon

                            Log.Information("=================== Auto Cancel of Reconcilation STARTED================");

                            try
                            {
                                if (ret == 0)
                                {
                                    GlobalFunctions.AutoCancelRecon(oCompany);

                                }

                            }
                            catch (Exception ex)
                            {
                                Log.Error(ex.Message);

                            }


                            Log.Information("=================== Auto Cancel of Reconcilation ENDED================");

                            #endregion

                            #region Auto Post Recon

                            Log.Information("=================== Auto Reconcilation STARTED================");

                            try
                            {
                                if (ret == 0)
                                {
                                    GlobalFunctions.AutoPostRecon(oCompany, dbData);

                                }

                            }
                            catch (Exception ex)
                            {
                                Log.Error(ex.Message);

                            }


                            Log.Information("=================== Auto Reconcilation ENDED================");

                            #endregion

                            #region Auto Adjust Balance

                            Log.Information("=================== Auto Balance Adjustment STARTED================");

                            try
                            {
                                if (ret == 0)
                                {
                                    GlobalFunctions.AutoBalAdjust(oCompany,dbData);

                                }

                            }
                            catch (Exception ex)
                            {
                                Log.Error(ex.Message);

                            }


                            Log.Information("=================== Auto Balance Adjustment ENDED================");

                            #endregion

                            #region Auto Documents Reconcile

                            Log.Information("=================== Auto Document Reconciliation STARTED================");

                            try
                            {
                                if (ret == 0)
                                {
                                    GlobalFunctions.AutoReconCustAcct(oCompany, dbData);

                                }

                            }
                            catch (Exception ex)
                            {
                                Log.Error(ex.Message);

                            }


                            Log.Information("=================== Auto Document Reconciliation ENDED================");

                            #endregion

                            Log.Information($"Sync Ended for the database : {dbData.CompanyDB} !!");
                        }
                    }
                    else
                        Log.Error("JSON File is empty!!!");
                }
            }
            catch(Exception ex)
            {
                Log.Error($"Error In Program Class : {ex.Message}");
            }

        }

        public static int SetCompanyConnection(SAPbobsCOM.Company oCompany,DBDetails dbData)
        {
            try
            {
                Log.Information("Starting to connect SAP company: ");

                int ret = -1;
                //oCompany.CompanyDB = ConfigurationManager.AppSettings["CompanyDB"];
                oCompany.CompanyDB = dbData.CompanyDB;
               // oCompany.DbUserName = ConfigurationManager.AppSettings["DbUserName"];
                oCompany.DbUserName = dbData.DbUserName;
                //oCompany.DbPassword = ConfigurationManager.AppSettings["DbPassword"];
                oCompany.DbPassword = dbData.DbPassword;

               // string Server = ConfigurationManager.AppSettings["Server"]; //"2017";
                string Server = dbData.Server; //"2017";
                switch (Server)
                {
                    case "2005":
                        oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2005;
                        break;
                    case "2008":
                        oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008;
                        break;
                    case "2012":
                        oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012;
                        break;
                    case "2014":
                        oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014;
                        break;
                    case "2016":
                        oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2016;
                        break;
                    case "2017":
                        oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2017;
                        break;
                    case "HANA":
                        oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB;
                        break;
                }
                //oCompany.Server = ConfigurationManager.AppSettings["B1Server"];
                oCompany.Server = dbData.B1Server;
                //oCompany.UserName = ConfigurationManager.AppSettings["UserName"];
                oCompany.UserName = dbData.UserName;
               // oCompany.Password = ConfigurationManager.AppSettings["Password"];
                oCompany.Password = dbData.Password;

                if (!oCompany.Connected)
                {
                    Log.Information("Conmany not connected already ");
                    ret = oCompany.Connect();

                }
                if (ret == 0)
                {
                    Log.Information("Company connected! " + oCompany.CompanyName + " " + oCompany.UserName);
                }
                else
                {
                    string Error = oCompany.GetLastErrorDescription();
                    Log.Error("Could not connect company:   " + Error);
                }
                return ret;
            }
            catch (Exception ex)
            {
                Log.Error("Could not connect company: " + ex.Message);
                return -1;
            }
        }
    }

    public class DBDetails
    {
        public string CompanyDB { get; set; }
        public string DbUserName { get; set; }
        public string DbPassword { get; set; }
        public string Server { get; set; }
        public string B1Server { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }
        public string AutoAdjustMentAccount { get; set; }
        public int OutgoingPaymentSeries { get; set; }
        public int IncomingPaymentSeries { get; set; }
        public int BranchID { get; set; }
    }

}
