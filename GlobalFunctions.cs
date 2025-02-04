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
using System.Runtime.Remoting.Contexts;
using Newtonsoft.Json;
using System.Security.Cryptography;

namespace CT_Export
{
    class GlobalFunctions
    {


        #region AutoReconcilation
        public static void AutoPostRecon(SAPbobsCOM.Company oCompany)
        {

            SAPbobsCOM.Recordset oRecordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = "";
            if (oCompany.DbServerType == BoDataServerTypes.dst_HANADB)
            {
                query = "Select \r\nT0.\"DocEntry\",\r\nT0.\"U_recondata\",\r\nT0.\"DocDate\",\r\nT0.\"DocNum\",\r\nT0.\"DocTotal\",\r\nT0.\"TransId\",\r\nT0.\"CardCode\",\r\n(select \"Line_ID\" from jdt1 where \"TransId\"=T0.\"TransId\" and \"ShortName\"=T0.\"CardCode\") as \"TransRow\"\r\n from ORCT T0 where cast(ifnull(T0.\"U_ext_entry\",'') as varchar(254))='Y' \r\nand cast(T0.\"U_recondata\" as varchar(254)) != ''  \r\nand cast(ifnull(T0.\"U_recon_num\",'') as varchar(254)) = '' and T0.\"DocEntry\"='430636'";
            }
            else
            {
                query = "Select \r\nT0.\"DocEntry\",\r\nT0.\"DocNum\",\r\nT0.\"DocTotal\",\r\nT0.\"TransId\",\r\nT0.\"CardCode\",\r\n(select \"Line_ID\" from jdt1 where \"TransId\"=T0.\"TransId\" and \"ShortName\"=T0.\"CardCode\") as \"TransRow\"\r\n from ORCT T0 where cast(ifnull(T0.\"U_ext_entry\",'') as varchar(254))='Y' \r\nand cast(T0.\"U_recondata\" as varchar(254)) != ''  \r\nand cast(ifnull(T0.\"U_recon_num\",'') as varchar(254)) = '' and T0.\"DocEntry\"='430636'";

            }

            oRecordset.DoQuery(query);

            if (oRecordset.RecordCount != 0)
            {
                Log.Information("Records Found for Reconcilation!");
                for (int i = 0; i < oRecordset.RecordCount; i++)
                {
                    string errorMessage = "";
                    string reconNum = "";
                    try
                    {
                        Log.Information($"Reconcilation started for Incoming Payment {oRecordset.Fields.Item("DocEntry").Value.ToString()}");
                        List<ReconData> reconData = JsonConvert.DeserializeObject<List<ReconData>>(oRecordset.Fields.Item("U_recondata").Value.ToString());
                        InternalReconciliationsService service = (InternalReconciliationsService)oCompany.GetCompanyService().GetBusinessService(ServiceTypes.InternalReconciliationsService);
                        InternalReconciliationOpenTrans openTrans = (InternalReconciliationOpenTrans)service.GetDataInterface(InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationOpenTrans);
                        InternalReconciliationParams reconciliationParams = (InternalReconciliationParams)service.GetDataInterface(InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationParams);
                        openTrans.CardOrAccount = CardOrAccountEnum.coaCard;
                        //string docDate = oRecordset.Fields.Item("DocDate").Value.ToString();
                        DateTime docDate = DateTime.Parse(oRecordset.Fields.Item("DocDate").Value.ToString());
                        openTrans.ReconDate = docDate;

                        openTrans.InternalReconciliationOpenTransRows.Add();
                        openTrans.InternalReconciliationOpenTransRows.Item(0).Selected = BoYesNoEnum.tYES;

                        //openTrans.InternalReconciliationOpenTransRows.Item(0). = "AACAC6164G";

                        openTrans.InternalReconciliationOpenTransRows.Item(0).TransId = Convert.ToInt32(oRecordset.Fields.Item("TransId").Value.ToString());

                        openTrans.InternalReconciliationOpenTransRows.Item(0).TransRowId = Convert.ToInt32(oRecordset.Fields.Item("TransRow").Value.ToString());

                        openTrans.InternalReconciliationOpenTransRows.Item(0).ReconcileAmount = Convert.ToDouble((oRecordset.Fields.Item("DocTotal").Value.ToString()));


                        SAPbobsCOM.Recordset reconRecordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        string reconQuery = $"select \r\nT0.\"TransId\",\r\nT0.\"DocEntry\",\r\n(select \"Line_ID\" from jdt1 where \"TransId\"=T0.\"TransId\" and \"ShortName\"=T0.\"CardCode\") as \"TransRow\"\r\nfrom oinv T0 where \"DocEntry\" in ({string.Join(",", reconData.Select(x => x.DocEntry.ToString()).ToArray())})";
                        reconRecordset.DoQuery(reconQuery);
                        if (reconRecordset.RecordCount > 0)
                        {
                            for (int j = 0; j < reconRecordset.RecordCount; j++)
                            {
                                openTrans.InternalReconciliationOpenTransRows.Add();

                                openTrans.InternalReconciliationOpenTransRows.Item(j + 1).Selected = BoYesNoEnum.tYES;

                                //openTrans.InternalReconciliationOpenTransRows.Item(0). = "AACAC6164G";

                                openTrans.InternalReconciliationOpenTransRows.Item(j + 1).TransId = Convert.ToInt32(reconRecordset.Fields.Item("TransId").Value.ToString());

                                openTrans.InternalReconciliationOpenTransRows.Item(j + 1).TransRowId = Convert.ToInt32(reconRecordset.Fields.Item("TransRow").Value.ToString());

                                double docEntry = Convert.ToDouble(reconRecordset.Fields.Item("DocEntry").Value.ToString());
                                double appliedamt = 0;
                                var amount = reconData.Where(x => x.DocEntry == docEntry);
                                foreach (var amt in amount)
                                    appliedamt = amt.AppliedAmt;

                                openTrans.InternalReconciliationOpenTransRows.Item(j + 1).ReconcileAmount = appliedamt;
                                reconRecordset.MoveNext();
                            }
                        }
                        reconciliationParams = service.Add(openTrans);
                        reconNum=reconciliationParams.ReconNum.ToString();
                        Log.Information($"Reconciliation successfull for the document {oRecordset.Fields.Item("DocEntry").Value.ToString()} internal recon number is {reconNum}");
                        //reconciliationParams.ReconNum = 212773;
                        //service.Cancel(reconciliationParams);
                    }
                    catch (Exception ex)
                    {
                        Log.Error($"Error in Reconciliation {ex.Message}");
                        errorMessage = ex.Message;
                    }
                    finally {
                        SAPbobsCOM.Payments PayUpdate = (SAPbobsCOM.Payments)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
                        try
                        {
                            if (PayUpdate.GetByKey(Convert.ToInt32(oRecordset.Fields.Item("DocEntry").Value.ToString())))
                            {
                                PayUpdate.UserFields.Fields.Item("U_recon_num").Value = reconNum;
                                PayUpdate.UserFields.Fields.Item("U_recon_error").Value = errorMessage;
                               
                                oCompany.StartTransaction();
                                int updateVal = PayUpdate.Update();
                                if (updateVal == 0)
                                {
                                    if (oCompany.InTransaction)
                                    {
                                        oCompany.EndTransaction(BoWfTransOpt.wf_Commit);
                                    }
                                    Log.Information($"Updation of incoming payment successfull : {PayUpdate.DocEntry.ToString()}");
                                }
                                else
                                    Log.Error($"Unable to update Incoming Payment : {oCompany.GetLastErrorDescription()}");
                            }
                        }
                        catch (Exception e)
                        {
                            Log.Error($"Error in Updating Payment : {e.Message}");
                        }
                    }
                    oRecordset.MoveNext();

                }

            }
        }
        #endregion

        #region AutoBalanceAdjustment
        public static void AutoBalAdjust(SAPbobsCOM.Company oCompany)
        {

            SAPbobsCOM.Recordset oRecordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = "";
            if (oCompany.DbServerType == BoDataServerTypes.dst_HANADB)
            {
                query = "select \r\ntop 1 \r\n\"CardCode\",\r\n\"Balance\"\r\nfrom OCRD where \"Balance\" between -10 and 10 and \"Balance\"!=0 and \"CardType\"='C' and \"CardCode\"='AAIFH2466D'\r\n";
            }
            else
            {
                query = "select \r\ntop 1 \r\n\"CardCode\",\r\n\"Balance\"\r\nfrom OCRD where \"Balance\" between -10 and 10 and \"Balance\"!=0 and \"CardType\"='C'\r\n";

            }

            oRecordset.DoQuery(query);

            if (oRecordset.RecordCount != 0)
            {
                Log.Information("Records Found for Auto Balance Adjustment!");
                for (int i = 0; i < oRecordset.RecordCount; i++)
                {
                    
                    try
                    {
                        Log.Information($"Auto Adjustment Started for BusinessPartner {oRecordset.Fields.Item("CardCode").Value.ToString()} with the Balance of {oRecordset.Fields.Item("Balance").Value.ToString()}");
                        double balanceAmt = Convert.ToDouble(oRecordset.Fields.Item("Balance").Value.ToString());
                        SAPbobsCOM.Payments newPay = (SAPbobsCOM.Payments)oCompany.GetBusinessObject(balanceAmt > 0 ? SAPbobsCOM.BoObjectTypes.oIncomingPayments : SAPbobsCOM.BoObjectTypes.oVendorPayments);
                        if (balanceAmt < 0)
                            newPay.Series =551;
                        newPay.CardCode = oRecordset.Fields.Item("CardCode").Value.ToString();
                        newPay.BPLID = 1;
                        newPay.DocDate = new DateTime(2024, 12, 20);
                        newPay.DueDate = new DateTime(2024, 12, 20); 
                        newPay.TaxDate = new DateTime(2024, 12, 20); 
                        newPay.DocType = BoRcptTypes.rCustomer;
                        newPay.Remarks = "Auto Adjustment Posting";
                        newPay.CashAccount = "_SYS00000001014";
                        newPay.CashSum = Convert.ToDouble(oRecordset.Fields.Item("Balance").Value.ToString()) >0 ? Convert.ToDouble(oRecordset.Fields.Item("Balance").Value.ToString()) : Convert.ToDouble(oRecordset.Fields.Item("Balance").Value.ToString()) *(-1);
                        oCompany.StartTransaction();
                        int ret = newPay.Add();
                        if (ret == 0)
                        {
                            Log.Information($"Auto Adjustment Document Posted with DocEntry {oCompany.GetNewObjectKey()} and Document Type  {oCompany.GetNewObjectType()}");
                            if (oCompany.InTransaction)
                                oCompany.EndTransaction(BoWfTransOpt.wf_Commit);
                        }
                        else
                        {
                            string error = oCompany.GetLastErrorDescription();
                            Log.Error($"Unable to Post the adjustment document Error : {oCompany.GetLastErrorDescription()}");
                            if (oCompany.InTransaction)
                                oCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
                        }
                    }
                    catch (Exception ex)
                    {
                        Log.Error($"Error in Adjustment Document Posting {ex.Message}");
                        if (oCompany.InTransaction)
                            oCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
                    }
                    finally
                    {

                    }
                    oRecordset.MoveNext();

                }

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


    public class ReconData
    {
        public int DocEntry { get; set; }
        public double AppliedAmt { get; set; }
    }
}

