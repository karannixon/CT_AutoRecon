﻿using System;
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

namespace CT_AutoRecon
{
    class GlobalFunctions
    {

        #region AutoCancelReconcilation
        public static void AutoCancelRecon(SAPbobsCOM.Company oCompany)
        {

            SAPbobsCOM.Recordset oRecordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = "";
            if (oCompany.DbServerType == BoDataServerTypes.dst_HANADB)
            {
                query = "select \r\n*\r\n,(select \"U_recon_num\" from ORCT where \"DocEntry\"=T0.\"DocEntry\") as \"ReconNum\"\r\n from \"ReconCancellations\" T0 where T0.\"CancellationAvailable\"='Y'";
            }
            else
            {
                query = "select \r\n*\r\n,(select \"U_recon_num\" from ORCT where \"DocEntry\"=T0.\"DocEntry\") as \"ReconNum\"\r\n from \"ReconCancellations\" T0 where T0.\"CancellationAvailable\"='Y'";

            }

            oRecordset.DoQuery(query);

            if (oRecordset.RecordCount != 0)
            {
                Log.Information("Records Found for Reconcilation Cancellation!");
                for (int i = 0; i < oRecordset.RecordCount; i++)
                {
                    string errorMessage = "";
                    string reconNum = "";
                    try
                    {
                        Log.Information($"Reconcilation cancellation started for Incoming Payment {oRecordset.Fields.Item("DocEntry").Value.ToString()} and Reconciliation Num : {oRecordset.Fields.Item("ReconNum").Value.ToString()}");
                        InternalReconciliationsService service = (InternalReconciliationsService)oCompany.GetCompanyService().GetBusinessService(ServiceTypes.InternalReconciliationsService);
                        InternalReconciliationParams reconciliationParams = (InternalReconciliationParams)service.GetDataInterface(InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationParams);

                        reconciliationParams.ReconNum = Convert.ToInt32(oRecordset.Fields.Item("ReconNum").Value.ToString()); ;
                        service.Cancel(reconciliationParams);
                        Log.Information($"Reconciliation cancellation successfull for the document {oRecordset.Fields.Item("DocEntry").Value.ToString()} internal recon number is {reconNum}");
                    }
                    catch (Exception ex)
                    {
                        Log.Error($"Error in Reconciliation {ex.Message}");
                        errorMessage = ex.Message;
                    }
                    finally
                    {
                        SAPbobsCOM.Payments PayUpdate = (SAPbobsCOM.Payments)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
                        try
                        {
                            if (PayUpdate.GetByKey(Convert.ToInt32(oRecordset.Fields.Item("DocEntry").Value.ToString())))
                            {
                                PayUpdate.UserFields.Fields.Item("U_recon_num").Value = "";
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
                            string status = errorMessage == "" ? "N" : "Y";
                            string UpdateQuery = $"Update \"ReconCancellations\" set \"Error_desc\"='{errorMessage}' , \"CancellationAvailable\"='{status}' where \"DocEntry\"='{oRecordset.Fields.Item("DocEntry").Value.ToString()}'";
                            SAPbobsCOM.Recordset oRecordsetUpdate = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            oRecordsetUpdate.DoQuery(UpdateQuery);
                        }
                        catch (Exception e)
                        {
                            Log.Error($"Error in Updating Payment and table : {e.Message}");
                        }
                    }
                    oRecordset.MoveNext();

                }

            }
        }
        #endregion

        #region AutoReconcilation
        public static void AutoPostRecon(SAPbobsCOM.Company oCompany, DBDetails dbData)
        {

            SAPbobsCOM.Recordset oRecordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = "";
            if (oCompany.DbServerType == BoDataServerTypes.dst_HANADB)
            {
                query = "Select \r\nT0.\"DocEntry\",\r\nT0.\"U_recondata\",\r\nT0.\"DocDate\",\r\nT0.\"DocNum\",\r\nT0.\"DocTotal\",\r\nT0.\"TransId\",\r\nT0.\"CardCode\",\r\n(select \"Line_ID\" from jdt1 where \"TransId\"=T0.\"TransId\" and \"ShortName\"=T0.\"CardCode\") as \"TransRow\"\r\n from ORCT T0 where T0.\"Canceled\"='N' and cast(ifnull(T0.\"U_ext_entry\",'') as varchar(254))='Y' \r\nand cast(T0.\"U_recondata\" as varchar(254)) != ''  \r\nand cast(ifnull(T0.\"U_recon_num\",'') as varchar(254)) = '' \r\nand cast(ifnull(T0.\"U_recon_error\",'') as varchar(254)) = '' ";
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
                        DateTime nextMonth = docDate.AddMonths(1);
                        DateTime comparedDate = new DateTime(nextMonth.Year, nextMonth.Month, 7);
                        openTrans.ReconDate = DateTime.Now >= comparedDate ? DateTime.Now : docDate;
                        openTrans.BPLID = dbData.BranchID;
                        openTrans.InternalReconciliationOpenTransRows.Add();
                        openTrans.InternalReconciliationOpenTransRows.Item(0).Selected = BoYesNoEnum.tYES;

                        //openTrans.InternalReconciliationOpenTransRows.Item(0). = "AACAC6164G";

                        openTrans.InternalReconciliationOpenTransRows.Item(0).TransId = Convert.ToInt32(oRecordset.Fields.Item("TransId").Value.ToString());

                        openTrans.InternalReconciliationOpenTransRows.Item(0).TransRowId = Convert.ToInt32(oRecordset.Fields.Item("TransRow").Value.ToString());

                        //openTrans.InternalReconciliationOpenTransRows.Item(0).ReconcileAmount = Convert.ToDouble((oRecordset.Fields.Item("DocTotal").Value.ToString()));
                        openTrans.InternalReconciliationOpenTransRows.Item(0).ReconcileAmount = reconData.Sum(x => x.AppliedAmt);


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
                        reconNum = reconciliationParams.ReconNum.ToString();
                        Log.Information($"Reconciliation successfull for the document {oRecordset.Fields.Item("DocEntry").Value.ToString()} internal recon number is {reconNum}");
                        //reconciliationParams.ReconNum = 212773;
                        //service.Cancel(reconciliationParams);
                    }
                    catch (Exception ex)
                    {
                        Log.Error($"Error in Reconciliation {ex.Message}");
                        errorMessage = ex.Message;
                    }
                    finally
                    {
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

        #region AutoReconcilation of Accounts with 0 Bal
        public static void AutoReconCustAcct(SAPbobsCOM.Company oCompany, DBDetails dbData)
        {

            SAPbobsCOM.Recordset oRecordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = "";
            if (oCompany.DbServerType == BoDataServerTypes.dst_HANADB)
            {
                query = "select \"CardCode\" from OCRD T0 where T0.\"Balance\"=0 and (select count(*) from JDT1 TX where TX.\"ShortName\"=T0.\"CardCode\" and (TX.\"BalDueDeb\"+TX.\"BalDueCred\")!=0) !=0 ;\r\n";
            }
            else
            {
                query = "select \"CardCode\" from OCRD where \"Balance\"=0";

            }

            oRecordset.DoQuery(query);

            if (oRecordset.RecordCount != 0)
            {

                for (int i = 0; i < oRecordset.RecordCount; i++)
                {
                    SAPbobsCOM.Recordset oRecordsetRecon = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    string reconDocuments = "";
                    if (oCompany.DbServerType == BoDataServerTypes.dst_HANADB)
                    {
                        reconDocuments = $"select \r\nT0.\"TransId\",\r\nT0.\"Line_ID\" as \"TransRowID\",\r\nT0.\"ShortName\" as \"CardCode\",\r\nT0.\"BalDueDeb\"+T0.\"BalDueCred\" as \"AppliedAmount\"\r\nfrom \"JDT1\" T0 \r\nwhere T0.\"ShortName\"='{oRecordset.Fields.Item("CardCode").Value.ToString()}' and (T0.\"BalDueDeb\"+T0.\"BalDueCred\")!=0";
                    }
                    else
                    {
                        reconDocuments = $"select \r\nT0.\"TransId\",\r\nT0.\"Line_ID\" as \"TransRowID\",\r\nT0.\"ShortName\" as \"CardCode\",\r\nT0.\"BalDueDeb\"+T0.\"BalDueCred\" as \"AppliedAmount\"\r\nfrom \"JDT1\" T0 \r\nwhere T0.\"ShortName\"='{oRecordset.Fields.Item("CardCode").Value.ToString()}' and (T0.\"BalDueDeb\"+T0.\"BalDueCred\")!=0";

                    }
                    oRecordsetRecon.DoQuery(reconDocuments);
                    if (oRecordsetRecon.RecordCount > 0)
                    {
                        string errorMessage = "";
                        string reconNum = "";
                        try
                        {
                            Log.Information($"Auto Reconciliation of the documents started for Business Partner : {oRecordset.Fields.Item("CardCode").Value.ToString()}");
                            InternalReconciliationsService service = (InternalReconciliationsService)oCompany.GetCompanyService().GetBusinessService(ServiceTypes.InternalReconciliationsService);
                            InternalReconciliationOpenTrans openTrans = (InternalReconciliationOpenTrans)service.GetDataInterface(InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationOpenTrans);
                            InternalReconciliationParams reconciliationParams = (InternalReconciliationParams)service.GetDataInterface(InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationParams);
                            openTrans.CardOrAccount = CardOrAccountEnum.coaCard;
                            //string docDate = oRecordset.Fields.Item("DocDate").Value.ToString();
                            //DateTime docDate = DateTime.Parse(oRecordset.Fields.Item("DocDate").Value.ToString());
                            openTrans.ReconDate = DateTime.Now;
                            openTrans.BPLID = dbData.BranchID;

                            for (int j = 0; j < oRecordsetRecon.RecordCount; j++)
                            {
                                openTrans.InternalReconciliationOpenTransRows.Add();

                                openTrans.InternalReconciliationOpenTransRows.Item(j).Selected = BoYesNoEnum.tYES;

                                //openTrans.InternalReconciliationOpenTransRows.Item(0). = "AACAC6164G";

                                openTrans.InternalReconciliationOpenTransRows.Item(j).TransId = Convert.ToInt32(oRecordsetRecon.Fields.Item("TransId").Value.ToString());

                                openTrans.InternalReconciliationOpenTransRows.Item(j).TransRowId = Convert.ToInt32(oRecordsetRecon.Fields.Item("TransRowID").Value.ToString());



                                openTrans.InternalReconciliationOpenTransRows.Item(j).ReconcileAmount = Convert.ToDouble(oRecordsetRecon.Fields.Item("AppliedAmount").Value.ToString());
                                oRecordsetRecon.MoveNext();
                            }

                            reconciliationParams = service.Add(openTrans);
                            reconNum = reconciliationParams.ReconNum.ToString();
                            Log.Information($"Reconciliation successfull for the Business Partner :{oRecordset.Fields.Item("CardCode").Value.ToString()} with Reconciliation Number : {reconNum}");
                            //reconciliationParams.ReconNum = 212773;
                            //service.Cancel(reconciliationParams);
                        }
                        catch (Exception ex)
                        {
                            Log.Error($"Error in Documentation sReconciliation {ex.Message}");
                            errorMessage = ex.Message;
                        }
                    }
                    oRecordset.MoveNext();

                }

            }
        }
        #endregion

        #region AutoBalanceAdjustment
        public static void AutoBalAdjust(SAPbobsCOM.Company oCompany, DBDetails dBDetails)
        {

            SAPbobsCOM.Recordset oRecordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = "";
            if (oCompany.DbServerType == BoDataServerTypes.dst_HANADB)
            {
                query = "select \r\n \r\n\"CardCode\",\r\n\"CardType\",\r\n\"Balance\"\r\nfrom OCRD where \"Balance\" between -10 and 10 and \"Balance\"!=0  \r\n";
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
                        SAPbobsCOM.Payments newPay;
                        if (oRecordset.Fields.Item("CardType").Value.ToString() == "C")
                            newPay = (SAPbobsCOM.Payments)oCompany.GetBusinessObject(balanceAmt > 0 ? SAPbobsCOM.BoObjectTypes.oIncomingPayments : SAPbobsCOM.BoObjectTypes.oVendorPayments);
                        else
                            newPay = (SAPbobsCOM.Payments)oCompany.GetBusinessObject(balanceAmt > 0 ? SAPbobsCOM.BoObjectTypes.oIncomingPayments : SAPbobsCOM.BoObjectTypes.oVendorPayments);

                        if (balanceAmt < 0 && dBDetails.OutgoingPaymentSeries > 0)
                            newPay.Series = dBDetails.OutgoingPaymentSeries ;
                        else if (balanceAmt > 0 && dBDetails.IncomingPaymentSeries > 0)
                            newPay.Series =  dBDetails.IncomingPaymentSeries ;

                        newPay.CardCode = oRecordset.Fields.Item("CardCode").Value.ToString();
                        newPay.BPLID = dBDetails.BranchID;
                        newPay.DocDate = DateTime.Now;
                        newPay.DueDate = DateTime.Now;
                        newPay.TaxDate = DateTime.Now;
                        //newPay.DocType = BoRcptTypes.rCustomer;
                        newPay.DocType = (oRecordset.Fields.Item("CardType").Value.ToString() == "C") ? BoRcptTypes.rCustomer : BoRcptTypes.rSupplier;
                        newPay.Remarks = "Auto Adjustment Posting";
                        newPay.CashAccount = dBDetails.AutoAdjustMentAccount;
                        newPay.CashSum = Convert.ToDouble(oRecordset.Fields.Item("Balance").Value.ToString()) > 0 ? Convert.ToDouble(oRecordset.Fields.Item("Balance").Value.ToString()) : Convert.ToDouble(oRecordset.Fields.Item("Balance").Value.ToString()) * (-1);
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

