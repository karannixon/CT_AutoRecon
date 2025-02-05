using Serilog;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CT_Export
{
    
    class Program
    {
        #region Variables
       
        public static List<string> FilesList = new List<string>();
        #endregion


        static void Main(string[] args)
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

            SAPbobsCOM.Company oCompany = new SAPbobsCOM.Company();
            int ret = SetCompanyConnection(oCompany);

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
                if (ret == 0 )
                {
                    GlobalFunctions.AutoPostRecon(oCompany);

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
                    GlobalFunctions.AutoBalAdjust(oCompany);

                }

            }
            catch (Exception ex)
            {
                Log.Error(ex.Message);

            }


            Log.Information("=================== Auto Balance Adjustment ENDED================");

            #endregion

        }

        public static int SetCompanyConnection(SAPbobsCOM.Company oCompany)
        {
            try
            {
                Log.Information("Starting to connect SAP company: ");

                int ret =-1 ;
                oCompany.CompanyDB = ConfigurationManager.AppSettings["CompanyDB"]; 
                oCompany.DbUserName = ConfigurationManager.AppSettings["DbUserName"];
                oCompany.DbPassword = ConfigurationManager.AppSettings["DbPassword"];
               
                string Server = ConfigurationManager.AppSettings["Server"]; //"2017";
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
                oCompany.Server = ConfigurationManager.AppSettings["B1Server"]; 
                oCompany.UserName = ConfigurationManager.AppSettings["UserName"];
                oCompany.Password = ConfigurationManager.AppSettings["Password"]; 
              
                if (!oCompany.Connected)
                {
                    Log.Information("Conmany not connected already ");
                    ret = oCompany.Connect();

                }
                if (ret == 0)
                {
                   // Log.Information("Company connected! " + oCompany.CompanyName + " " + oCompany.UserName);
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
}
