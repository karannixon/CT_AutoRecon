using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using System.Xml;
using Serilog;
using CsvHelper;
using CsvHelper.Configuration;
using System.Globalization;
using SAPbobsCOM;
using System.Text.RegularExpressions;
using System.Configuration;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Collections;
using System.Runtime.InteropServices;

namespace CT_Export
{
    class GlobalFunctions
    {
        

        #region PDFExport
        public static void AutoVatRecon(SAPbobsCOM.Company oCompany, string table)
        {
            
            SAPbobsCOM.Recordset oRecordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = "";
            if (oCompany.DbServerType == BoDataServerTypes.dst_HANADB)
            {
                query = "Select * from ORCT where cast(ifnull(\"U_ext_entry\",'') as varchar(254))='Y' \r\nand cast(\"U_recondata\" as varchar(254)) != ''  \r\nand cast(ifnull(\"U_recon_num\",'') as varchar(254)) = ''";
            }
            else
            {
                query = "Select * from ORCT where cast(ifnull(\"U_ext_entry\",'') as varchar(254))='Y' \r\nand cast(\"U_recondata\" as varchar(254)) != ''  \r\nand cast(ifnull(\"U_recon_num\",'') as varchar(254)) = ''";

            }

            oRecordset.DoQuery(query);
            
            if (oRecordset.RecordCount != 0)
            {
                

            }
        }
        #endregion

        #region Create Customization Tools
        public static void CustomizationTools(SAPbobsCOM.Company oCompany)
        {
            ////Log.Information("USER DEFINED OBJECTS STARTED");

            ////// Create UDTs, UDOs, and UDFs
            ////Log.Information("Staging Creation Started");


            ////Log.Information("CT ExportConfig Creation Started");

            ////CreateUDTs(oCompany, "CT_ExportConfigs", "CT PDFExport");
            ////CreateUDFs(null, null, oCompany, "Document", "Document", "@CT_ExportConfigs", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 100, "");
            ////CreateUDFs(null, null, oCompany, "Layout_Name", "Layout Name", "@CT_ExportConfigs", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 100, "");
            ////CreateUDFs(null, null, oCompany, "Layout_Path", "Layout Path", "@CT_ExportConfigs", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 100, "");
            ////CreateUDFs(null, null, oCompany, "Active", "Active", "@CT_ExportConfigs", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 10, "");

            ////Log.Information("CT ExportConfig Creation Ended");

            ////Log.Information("Market DOcs Header UDFs Creation Started");


            ////ArrayList validValues1 = new ArrayList();

            ////validValues1.Add(new ValidValues("success", "success"));
            ////validValues1.Add(new ValidValues("failure", "failure"));
            ////string defaultValue1 = "";
            ////CreateUDFs(validValues1, defaultValue1, oCompany, "Drive_Upload_Status", "Status", "OINV", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 10, "");

            ////CreateUDFs(null, null, oCompany, "CT_Exported_Date", "Exported Date", "OINV", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 10, "");
            ////CreateUDFs(null, null, oCompany, "CT_Exported_Reason", "Drive failure Reason", "OINV", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 100, "");

            ////Log.Information("Market DOcs Head UDF Creation Ended");

            ////Log.Information("USER DEFINED OBJECTS COMPLETED");

        }
        #endregion

        #region valid value
        struct ValidValues
        {
            public string sValue;
            public string sDescription;
            public ValidValues(string Value, string Description)
            {
                sValue = Value;
                sDescription = Description;
            }

            public string Description
            {
                get
                {
                    return sDescription;
                }
                set
                {
                    sDescription = value;
                }
            }

            public string Value
            {
                get
                {
                    return sValue;
                }
                set
                {
                    sValue = value;
                }
            }

        }
        #endregion

        #region UDT Functions
        public static void CreateUDTs(SAPbobsCOM.Company oCompany, string UDTCode, string UDTDescription, BoUTBTableType tableType = BoUTBTableType.bott_NoObjectAutoIncrement)
        {
            // Check if UDT already exists
            UserTablesMD userTable = (UserTablesMD)oCompany.GetBusinessObject(BoObjectTypes.oUserTables);



            if (userTable.GetByKey(UDTCode) == false)
            {
                userTable.TableName = UDTCode;
                userTable.TableDescription = UDTDescription;
                userTable.TableType = tableType;
                int result = userTable.Add();

                if (result != 0)
                {
                    Log.Error("UDT " + UDTCode + " " + UDTDescription + " failed to be created." + oCompany.GetLastErrorCode() + "  " + oCompany.GetLastErrorDescription());

                }
                else
                {
                    Log.Information("UDT " + UDTCode + " " + UDTDescription + " created successfully.");
                }
            }
            else
            {
                Log.Information("UDT " + UDTCode + " " + UDTDescription + " already created .");

            }

            // Release the UserFieldsMD object
            System.Runtime.InteropServices.Marshal.ReleaseComObject(userTable);
            GC.Collect(); // Force garbage collection to release the object

        }
        private static bool UDTExists(Company company, string tableName)
        {
            UserTablesMD userTable = (UserTablesMD)company.GetBusinessObject(BoObjectTypes.oUserTables);
            

            System.Runtime.InteropServices.Marshal.ReleaseComObject(userTable);
            GC.Collect();

            return userTable.GetByKey(tableName);
        }
        #endregion

        #region UDF Functions
        public static void CreateUDFs(ArrayList validVal, string defaultValue, SAPbobsCOM.Company oCompany, string UDFCode, string UDFDescription, string TableName, BoFieldTypes type = BoFieldTypes.db_Alpha, BoFldSubTypes subtype = BoFldSubTypes.st_None, int editSize = 20, string linkedTable = "")
        {
            // Check if UDF already exists
            UserFieldsMD userField = (UserFieldsMD)oCompany.GetBusinessObject(BoObjectTypes.oUserFields);
            SAPbobsCOM.Recordset oRecordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = "SELECT FieldID FROM CUFD WHERE TableID = '" + TableName + "' AND AliasID = '" + UDFCode + "'";
            oRecordset.DoQuery(query);

            if (oRecordset.RecordCount == 0)
            {
                // Release the UserFieldsMD object
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
                GC.Collect(); // Force garbage collection to release the object

                userField.TableName = TableName;
                userField.Name = UDFCode;
                userField.Description = UDFDescription;
                userField.Type = type;
                userField.SubType = subtype;
                userField.EditSize = editSize;
                userField.Size = editSize;
                //userField.LinkedTable = "OUSR"; // Linked table for UDF

                //valid_value 
                if (validVal != null && validVal.Count > 0)
                {                    

                    foreach (ValidValues item in validVal)
                    {
                        string Key = item.sValue;
                        string Desc = item.sDescription;

                        userField.ValidValues.Value = item.sValue; ;
                        userField.ValidValues.Description = item.sDescription;
                        userField.ValidValues.Add();
                    }
                }

                int result = userField.Add();

                if (result != 0)
                {
                    string erromessage = oCompany.GetLastErrorDescription();
                    Log.Error("UDF " + UDFCode + " " + UDFDescription + " failed to be created." + oCompany.GetLastErrorCode() + "  " + erromessage);

                }
                else
                {
                    Log.Information("UDF " + UDFCode + " " + UDFDescription + " created successfully.");
                }


            }
            else
            {
                Log.Information("UDF " + UDFCode + " " + UDFDescription + " already created .");

            }
            // Release the UserFieldsMD object
            System.Runtime.InteropServices.Marshal.ReleaseComObject(userField);
            GC.Collect(); // Force garbage collection to release the object

        }
        #endregion
    }

}

