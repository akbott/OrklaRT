using SAP.Middleware.Connector;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using ERPConnect;
using ERPConnect.BW;
using System.Collections.Specialized;

namespace BPL
{
    public class RfcFunctions
    {
        public static DataTable dataTable;       
        public static void GetUserSettingsFromSAP()
        {
            RfcDestination destination = RfcConnection.GetDestination("0");

            IRfcFunction fieldList = destination.Repository.CreateFunction("RFC_READ_TABLE");
            fieldList.SetValue("QUERY_TABLE", "USR01");
            fieldList.SetValue("DELIMITER", ";");
            fieldList.SetValue("NO_DATA", String.Empty);

            IRfcTable options = fieldList.GetTable("OPTIONS");
            options.Append();
            options.SetValue("TEXT", "BNAME ='" +  Environment.UserName.ToUpper() + "'");

            IRfcTable fields = fieldList.GetTable("FIELDS");
            fields.Append();
            fields.SetValue("FIELDNAME", "DCPFM");
            fields.Append();
            fields.SetValue("FIELDNAME", "DATFM");
            fields.Append();
            fields.SetValue("FIELDNAME", "LANGU");
            
            RfcSessionManager.BeginContext(destination);
            try
            {
                fieldList.Invoke(destination);
            }
            catch { Exception ex; }
            RfcSessionManager.EndContext(destination);

            IRfcTable resultTable = fieldList.GetTable("DATA");
            using (var entities = new DAL.SAPExlEntities())
            {
                int Id = entities.vwCurrentUser.SingleOrDefault().ID;
                var currentUser = entities.CurrentUsers.SingleOrDefault(c => c.ID == Id);
                for (int j = 0; j < resultTable.RowCount; j++)
                {
                    resultTable.CurrentIndex = j;
                    switch (resultTable[j][0].ToString().Split(new Char[] { '=',';'}).GetValue(1).ToString())
                    {
                        case " ":
                            currentUser.DecimalSeparator = ",";
                            currentUser.ThousandSeparator = ".";
                            break;
                        case "X":
                            currentUser.DecimalSeparator = ".";
                            currentUser.ThousandSeparator = ",";
                            break;
                        case "Y":
                            currentUser.DecimalSeparator = ".";
                            currentUser.ThousandSeparator = " ";
                            break;
                        default:
                            currentUser.DecimalSeparator = ",";
                            currentUser.ThousandSeparator = ".";
                            break;
                    }

                    switch (resultTable[j][0].ToString().Split(new Char[] { '=', ';' }).GetValue(2).ToString())
                    {
                        case "1":
                            currentUser.DateFormat = "dd.mm.yyyy";
                            //currentUser.TextFileDateType = 3;
                            break;
                        case "2":
                            currentUser.DateFormat = "mm/dd/yyyy";
                            //currentUser.TextFileDateType = 3;
                            break;
                        case "3":
                            currentUser.DateFormat = "mm-dd-yyyy";
                            //currentUser.TextFileDateType = 3;
                            break;
                        case "4":
                            currentUser.DateFormat = "yyyy.mm.dd";
                            //currentUser.TextFileDateType = 5;
                            break;
                        case "5":
                            currentUser.DateFormat = "yyyy/mm/dd";
                            //currentUser.TextFileDateType = 5;
                            break;
                        case "6":
                            currentUser.DateFormat = "yyyy-mm-dd";
                            //currentUser.TextFileDateType = 5;
                            break;
                        default:
                            currentUser.DateFormat = "yyyy/mm/dd";
                            //currentUser.TextFileDateType = 5;
                            break;
                    }

                    switch (resultTable[j][0].ToString().Split(new Char[] { '=', ';' }).GetValue(3).ToString())
                    {
                        case "C":
                            currentUser.Language = "CS";
                            break;
                        case "D":
                            currentUser.Language = "DE";
                            break;
                        case "E":
                            currentUser.Language = "EN";
                            break;
                        case "L":
                            currentUser.Language = "PL";
                            break;
                        case "O":
                            currentUser.Language = "NO";
                            break;
                        case "Q":
                            currentUser.Language = "SK";
                            break;
                        default:
                            currentUser.Language = "NO";
                            break;
                    }
                }
                entities.SaveChanges();
            }
        }
        public static DataTable BWFunctionCall(string queryName, int reportID, int userID, int variantID, string variantName,string lang)
        {
            StringCollection columnNames = new StringCollection();
            DataTable returnTable = new DataTable();           

            ERPConnect.LIC.SetLic("D609DPPP0C");

            using (R3Connection con = RfcConnection.GetBHPConnection(lang))
            {
                con.Open();
                BWCube query = new BWCube(con);
                query = con.CreateBWCube(queryName);                  

                for (int i = 0; i < query.Dimensions.Count; i++)
                {
                    query.Dimensions[i].SelectForFlatMDX = true;
                    columnNames.Add(query.Dimensions[i].Name);
                    if (query.Dimensions[i].Name.Equals("0PLANT"))
                    {
                        query.Dimensions[i].Properties["20PLANT"].SelectForFlatMDX = true;
                        columnNames.Add(query.Dimensions[i].Properties["20PLANT"].Name);
                    }
                    else if (query.Dimensions[i].Name.Equals("0MATERIAL"))
                    {
                        query.Dimensions[i].Properties["20MATERIAL"].SelectForFlatMDX = true;
                        columnNames.Add(query.Dimensions[i].Properties["20MATERIAL"].Name);
                    }
                    else if (query.Dimensions[i].Name.Equals("0MATERIAL__0G_CWW020"))
                    {
                        query.Dimensions[i].Properties["20MATERIAL__0G_CWW020"].SelectForFlatMDX = true;
                        columnNames.Add(query.Dimensions[i].Properties["20MATERIAL__0G_CWW020"].Name);
                    }
                    else if (query.Dimensions[i].Name.Equals("0MATERIAL__0G_CWW021"))
                    {
                        query.Dimensions[i].Properties["20MATERIAL__0G_CWW021"].SelectForFlatMDX = true;
                        columnNames.Add(query.Dimensions[i].Properties["20MATERIAL__0G_CWW021"].Name);
                    }
                    else if (query.Dimensions[i].Name.Equals("0LANGU"))
                    {
                        query.Dimensions[i].Properties["20LANGU"].SelectForFlatMDX = true;
                        columnNames.Add(query.Dimensions[i].Properties["20LANGU"].Name);
                    }
                    else if (query.Dimensions[i].Name.Equals("0DOC_CURRCY"))
                    {
                        query.Dimensions[i].Properties["20DOC_CURRCY"].SelectForFlatMDX = true;
                        columnNames.Add(query.Dimensions[i].Properties["20DOC_CURRCY"].Name);
                    }
                    //else if (query.Dimensions[i].Name.Equals("0UNIT"))
                    //{
                    //    query.Dimensions[i].Properties["20UNIT"].SelectForFlatMDX = true;
                    //    columnNames.Add(query.Dimensions[i].Properties["20UNIT"].Name);
                    //}
                    //else if (query.Dimensions[i].Name.Equals("0SALES_UNIT"))
                    //{
                    //    query.Dimensions[i].Properties["20SALES_UNIT"].SelectForFlatMDX = true;
                    //    columnNames.Add(query.Dimensions[i].Properties["20SALES_UNIT"].Name);
                    //}
                    //else if (query.Dimensions[i].Name.Equals("1CUDIM"))
                    //{
                    //    query.Dimensions[i].Properties["21CUDIM"].SelectForFlatMDX = true;
                    //    columnNames.Add(query.Dimensions[i].Properties["21CUDIM"].Name);
                    //}                       
                    else if (query.Dimensions[i].Name.Equals("0MATL_TYPE"))
                    {
                        query.Dimensions[i].Properties["20MATL_TYPE"].SelectForFlatMDX = true;
                        columnNames.Add(query.Dimensions[i].Properties["20MATL_TYPE"].Name);
                    }
                    else if (query.Dimensions[i].Name.Equals("0PROD_HIER"))
                    {
                        query.Dimensions[i].Properties["20PROD_HIER"].SelectForFlatMDX = true;
                        columnNames.Add(query.Dimensions[i].Properties["20PROD_HIER"].Name);
                    }
                    else if (query.Dimensions[i].Name.Equals("0SOLD_TO__0CUST_GROUP"))
                    {
                        query.Dimensions[i].Properties["20SOLD_TO__0CUST_GROUP"].SelectForFlatMDX = true;
                        columnNames.Add(query.Dimensions[i].Properties["20SOLD_TO__0CUST_GROUP"].Name);
                    }
                    else if (query.Dimensions[i].Name.Equals("0SOLD_TO"))
                    {
                        query.Dimensions[i].Properties["20SOLD_TO"].SelectForFlatMDX = true;
                        columnNames.Add(query.Dimensions[i].Properties["20SOLD_TO"].Name);
                    }
                    else if (query.Dimensions[i].Name.Equals("0VENDOR"))
                    {
                        query.Dimensions[i].Properties["20VENDOR"].SelectForFlatMDX = true;
                        columnNames.Add(query.Dimensions[i].Properties["20VENDOR"].Name);
                    }                  
                    else if (query.Dimensions[i].Name.Equals("0WORKCENTER"))
                    {
                        query.Dimensions[i].Properties["20WORKCENTER"].SelectForFlatMDX = true;
                        columnNames.Add(query.Dimensions[i].Properties["20WORKCENTER"].Name);
                    }   
                    else if (query.Dimensions[i].Name.Equals("ZIMATNR"))
                    {
                        if (reportID.Equals(7) || reportID.Equals(63))
                        {
                            query.Dimensions[i].Properties["2ZIMATNR"].SelectForFlatMDX = true;
                            columnNames.Add(query.Dimensions[i].Properties["2ZIMATNR"].Name);
                        }                                                
                    }
                    else if (query.Dimensions[i].Name.Equals("ZIMATMREA"))
                    {
                        if(reportID.Equals(33))
                        {
                            query.Dimensions[i].Properties["2ZIMATMREA"].SelectForFlatMDX = true;
                            columnNames.Add(query.Dimensions[i].Properties["2ZIMATMREA"].Name);
                        }                                                
                    }
                    else if (query.Dimensions[i].Name.Equals("0PRODSCHED"))
                    {
                        if (reportID.Equals(12))
                        {
                            query.Dimensions[i].Properties["20PRODSCHED"].SelectForFlatMDX = true;
                            columnNames.Add(query.Dimensions[i].Properties["20PRODSCHED"].Name);
                        }
                    }
                    else if (query.Dimensions[i].Name.Equals("0MPN_MATNR"))
                    {
                        if (reportID.Equals(2) || reportID.Equals(8))
                        {
                            query.Dimensions[i].Properties["20MPN_MATNR"].SelectForFlatMDX = true;
                            columnNames.Add(query.Dimensions[i].Properties["20MPN_MATNR"].Name);
                        }
                    }
                    else if (query.Dimensions[i].Name.Equals("0STOR_LOC"))
                    {
                        if (reportID.Equals(39))
                        {
                            query.Dimensions[i].Properties["20STOR_LOC"].SelectForFlatMDX = true;
                            columnNames.Add(query.Dimensions[i].Properties["20STOR_LOC"].Name);
                        }
                    }
                    else if (query.Dimensions[i].Name.Equals("ZMATERIA1"))
                    {
                        if (reportID.Equals(3))
                        {
                            query.Dimensions[i].Properties["2ZMATERIA1"].SelectForFlatMDX = true;
                            columnNames.Add(query.Dimensions[i].Properties["2ZMATERIA1"].Name);
                        }
                    }
                    else if (query.Dimensions[i].Name.Equals("ZMATERIA2"))
                    {
                        if (reportID.Equals(3))
                        {
                            query.Dimensions[i].Properties["2ZMATERIA2"].SelectForFlatMDX = true;
                            columnNames.Add(query.Dimensions[i].Properties["2ZMATERIA2"].Name);
                        }
                    }
                    else if (query.Dimensions[i].Name.Equals("ZMATL_TY1"))
                    {
                        if (reportID.Equals(3))
                        {
                            query.Dimensions[i].Properties["2ZMATL_TY1"].SelectForFlatMDX = true;
                            columnNames.Add(query.Dimensions[i].Properties["2ZMATL_TY1"].Name);
                        }
                    }
                    else if (query.Dimensions[i].Name.Equals("ZMATL_TY2"))
                    {
                        if (reportID.Equals(3))
                        {
                            query.Dimensions[i].Properties["2ZMATL_TY2"].SelectForFlatMDX = true;
                            columnNames.Add(query.Dimensions[i].Properties["2ZMATL_TY2"].Name);
                        }
                    }
                    else if (query.Dimensions[i].Name.Equals("0MATL_GROUP"))
                    {
                        if (reportID.Equals(3))
                        {
                            query.Dimensions[i].Properties["20MATL_GROUP"].SelectForFlatMDX = true;
                            columnNames.Add(query.Dimensions[i].Properties["20MATL_GROUP"].Name);
                        }
                    }
                    else if (query.Dimensions[i].Name.Equals("0EXECPLANT"))
                    {
                        if (reportID.Equals(3))
                        {
                            query.Dimensions[i].Properties["20EXECPLANT"].SelectForFlatMDX = true;
                            columnNames.Add(query.Dimensions[i].Properties["20EXECPLANT"].Name);
                        }
                    }
                    else if (query.Dimensions[i].Name.Equals("0PLANTTO"))
                    {
                        if (reportID.Equals(3))
                        {
                            query.Dimensions[i].Properties["20PLANTTO"].SelectForFlatMDX = true;
                            columnNames.Add(query.Dimensions[i].Properties["20PLANTTO"].Name);
                        }
                    }
                    else if (query.Dimensions[i].Name.Equals("0COORDER"))
                    {
                        if (reportID.Equals(7) || reportID.Equals(8) || reportID.Equals(12) || reportID.Equals(13) || reportID.Equals(20) || reportID.Equals(62))
                        {
                            query.Dimensions[i].Properties["20COORDER"].SelectForFlatMDX = true;
                            columnNames.Add(query.Dimensions[i].Properties["20COORDER"].Name);
                        }
                    }
                }              

                for (int i = 0; i < query.Measures.Count; i++)
                {
                    query.Measures[i].SelectForFlatMDX = true;
                    columnNames.Add(query.Measures[i].Description);
                }

                foreach (DataRow dataRow in GetUserReportSelections(reportID, userID, variantID).Tables[0].Rows)
                {
                    if (query.Variables[dataRow["ScreenID"].ToString()].SelectionType.Equals(BWVariableSelectionType.Complex) || query.Variables[dataRow["ScreenID"].ToString()].SelectionType.Equals(BWVariableSelectionType.MultipleSingle))
                    {
                        if (dataRow["ScreenID"].ToString().Equals("ZVR013"))
                        {
                            if (!dataRow["LowValue"].ToString().Equals("000000000000000000"))
                            {
                                query.Variables[dataRow["ScreenID"].ToString()].ComplexRanges.Add(dataRow["LowValue"].ToString());
                            }
                        }
                        else
                        { 
                            query.Variables[dataRow["ScreenID"].ToString()].ComplexRanges.Add(dataRow["LowValue"].ToString());
                        }
                        //query.Variables[dataRow["ScreenID"].ToString()].SingleRange.LowValue = dataRow["LowValue"].ToString(); 
                    }
                    else if (query.Variables[dataRow["ScreenID"].ToString()].SelectionType.Equals(BWVariableSelectionType.Interval))
                    {
                        query.Variables[dataRow["ScreenID"].ToString()].SingleRange.LowValue = dataRow["LowValue"].ToString();                   
                        query.Variables[dataRow["ScreenID"].ToString()].SingleRange.HighValue = dataRow["HighValue"].ToString();                        
                    }
                    else if (query.Variables[dataRow["ScreenID"].ToString()].SelectionType.Equals(BWVariableSelectionType.Single))
                    {
                        query.Variables[dataRow["ScreenID"].ToString()].SingleRange.LowValue = dataRow["LowValue"].ToString();                        
                    }
                    
                }                

                try
                {
                    string str1 = query.GetFlatMDX();
                    returnTable = query.Execute();     
                }
                catch { Exception ex; }

                for (int i = 0; i <= returnTable.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < returnTable.Columns.Count; j++)
                    {
                        returnTable.Rows[i][j] = returnTable.Rows[i][j].ToString().Replace("#", String.Empty);
                    }
                }
                
               con.Close();
               con.Dispose();                             
               returnTable.Dispose();               
               GC.SuppressFinalize(query);
               return returnTable;
                
            }
        }
        //public static IEnumerable<string> SplitByLength(string str, int maxLength)
        //{
        //    for (int index = 0; index < str.Length; index += maxLength)
        //    {
        //        yield return str.Substring(index, Convert.ToInt32(str.Substring(index, Math.Min(maxLength, str.Length - index))));
        //    }
        //}

        public static IEnumerable<string> SplitByLength(string str)
        {
            int index = 0;
            while (true)
            {
                if (index >= str.Length)
                {
                    yield break;
                }
                yield return str.Substring(index + 4, Convert.ToInt32(str.Substring(index, 3)));
                index += Convert.ToInt32(str.Substring(index, 3)) + 5;
            }
        }

        public static void RfcTransactionCallUsing(dynamic dataTable, string transCode)
        {
            RfcDestination destination = RfcConnection.GetDestination("2");

            IRfcFunction rfcCallTransUsing = destination.Repository.CreateFunction("RFC_CALL_TRANSACTION_USING");

            rfcCallTransUsing.SetValue("MODE", "E");
            rfcCallTransUsing.SetValue("TCODE", transCode);

            IRfcTable bdcInput = rfcCallTransUsing["BT_DATA"].GetTable();

            using (var entities = new DAL.SAPExlEntities())
            {
                var bdcInputRows = entities.BDCInput.Where(bdc => bdc.TransactionCode == transCode);
                foreach (var bdcRow in bdcInputRows)
                {
                    RfcStructureMetadata bdData = destination.Repository.GetStructureMetadata("BDCDATA");
                    IRfcStructure row = bdData.CreateStructure();
                    row.SetValue("PROGRAM", bdcRow.Program == null ? string.Empty : bdcRow.Program);
                    row.SetValue("DYNPRO", bdcRow.Dynpro);
                    row.SetValue("DYNBEGIN", bdcRow.ID == null ? string.Empty : bdcRow.ID);
                    row.SetValue("FNAM", bdcRow.FieldName == null ? string.Empty : bdcRow.FieldName);

                    if (bdcRow.ValueFieldName != null)
                    {
                        foreach (DataRow dataRow in dataTable.Rows)
                        {
                            if (dataRow["ValueFieldName"].Equals(bdcRow.ValueFieldName))
                            {
                                row.SetValue("FVAL", dataRow["FieldValue"].ToString());
                            }
                        }
                    }
                    else
                    {
                        row.SetValue("FVAL", bdcRow.FieldValue == null ? string.Empty : bdcRow.FieldValue);
                    }

                    bdcInput.Append(row);
                }
            }

            RfcSessionManager.BeginContext(destination);
            try
            {
                rfcCallTransUsing.Invoke(destination);
            }
            catch { }
            RfcSessionManager.EndContext(destination);

            IRfcTable rfcReturn = rfcCallTransUsing["L_ERRORS"].GetTable();
            for (int j = 0; j < rfcReturn.RowCount; j++)
            {
                rfcReturn.CurrentIndex = j;
                switch (rfcReturn.GetString("MSGTYP"))
                {
                    case "A":
                        //MessageBox.Show(RfcReturn.GetString("DYNAME"), "RFC_CALL_TRANSACTION_USING", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        break;
                    case "E":
                        //MessageBox.Show(RfcReturn.GetString("DYNAME"), "RFC_CALL_TRANSACTION_USING", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        break;
                    case "W":
                        //MessageBox.Show(RfcReturn.GetString("MESSAGE"), "RFC_CALL_TRANSACTION_USING", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        break;
                    case "I":
                        //MessageBox.Show(RfcReturn.GetString("MESSAGE"), "RFC_CALL_TRANSACTION_USING", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        break;
                    default:
                        break;
                    //Ignore Success Messages
                }
            }
        }
        
        public static bool IsNumeric(string value)
        {
            bool ret = false;
            try
            {
                double num;
                bool isNum = double.TryParse(value.TrimEnd(), out num);
                if (isNum.Equals(true))
                {
                    ret = true;
                }
                else
                {
                    ret = false;
                }
            }
            catch { }
            return ret;
        }
        public static DataSet GetUserReportSelections(int reportID,int userID,int variantID)
        {
            string commandString = "EXEC dbo.prGetUserReportSelections " + reportID + "," + userID + "," + variantID;
            return SQLDataHandler.Functions.GetData("UserReportSelections", commandString);
        }
        public static IRfcTable GetBAPIMATERIALMRPLISTAll(string material,string plant)
        {
            RfcDestination destination = RfcConnection.GetDestination("0");

            IRfcFunction remoteQueryCall = destination.Repository.CreateFunction("BAPI_MATERIAL_MRP_LIST");
            remoteQueryCall.SetValue("MATERIAL", String.Format("{0:000000000000000000}", Convert.ToInt64(material)));
            remoteQueryCall.SetValue("PLANT", plant);
            remoteQueryCall.SetValue("GET_TOTAL_LINES", "X");
            remoteQueryCall.SetValue("GET_IND_LINES", " ");

            RfcSessionManager.BeginContext(destination);
            try
            {
                remoteQueryCall.Invoke(destination);
            }
            catch { Exception ex; }
            RfcSessionManager.EndContext(destination);

            IRfcTable rfcReturn = remoteQueryCall.GetTable("MRP_TOTAL_LINES");
            return rfcReturn;
        }
        public static DataTable GetBAPIMATERIALGETAll(string material, string plant)
        {
            RfcDestination destination = RfcConnection.GetDestination("0");

            IRfcFunction remoteQueryCall = destination.Repository.CreateFunction("BAPI_MATERIAL_GET_ALL");
            remoteQueryCall.SetValue("MATERIAL", String.Format("{0:000000000000000000}", Convert.ToInt64(material)));
            remoteQueryCall.SetValue("PLANT", plant);            

            RfcSessionManager.BeginContext(destination);
            try
            {
                remoteQueryCall.Invoke(destination);
            }
            catch { Exception ex; }
            RfcSessionManager.EndContext(destination);

            IRfcStructure maraReturn = remoteQueryCall.GetStructure("CLIENTDATA");
            IRfcStructure marcReturn = remoteQueryCall.GetStructure("PLANTDATA");
            IRfcStructure mardReturn = remoteQueryCall.GetStructure("STORAGELOCATIONDATA");

            IRfcTable materialReturn = remoteQueryCall.GetTable("MATERIALDESCRIPTION");
            string materialName = String.Empty;

            for (int i = 0; i < materialReturn.Count; i++)
            {
                materialReturn.CurrentIndex = i;
                if (materialReturn[i]["LANGU"].ToString().Split(new Char[] { '=', ';' }).GetValue(1).ToString().Equals("O"))
                {
                    materialName = materialReturn[i]["MATL_DESC"].ToString().Split(new Char[] { '=', ';' }).GetValue(1).ToString();
                }               
            }
            
            DataTable ret = new DataTable();

            for (int j = 0; j < 35; j++)
            {
                ret.Columns.Add("Column" + j.ToString(), typeof(String));
            }

            ret.Rows.Add(maraReturn["MATERIAL"].GetString(), marcReturn["PLANT"].GetString(), marcReturn["PUR_STATUS"].GetString(), marcReturn["PVALIDFROM"].GetString(),
            marcReturn["ABC_ID"].GetString(), marcReturn["CRIT_PART"].GetString(), marcReturn["PUR_GROUP"].GetString(), marcReturn["ISSUE_UNIT"].GetString(), marcReturn["MRPPROFILE"].GetString(),
            marcReturn["MRP_TYPE"].GetString(), marcReturn["MRP_CTRLER"].GetString(), "", marcReturn["PLND_DELRY"].GetString(), marcReturn["GR_PR_TIME"].GetString(),
            marcReturn["PERIOD_IND"].GetString(), marcReturn["LOTSIZEKEY"].GetString(), marcReturn["PROC_TYPE"].GetString(), marcReturn["SPPROCTYPE"].GetString(), marcReturn["REORDER_PT"].GetString(),
            marcReturn["SAFETY_STK"].GetString(), marcReturn["MINLOTSIZE"].GetString(), marcReturn["MAXLOTSIZE"].GetString(), marcReturn["FIXED_LOT"].GetString(), marcReturn["ROUND_VAL"].GetString(),
            marcReturn["SAFTY_T_ID"].GetString(), marcReturn["SAFETYTIME"].GetString(), mardReturn["CURR_PERIOD"].GetString(), mardReturn["FISC_YEAR"].GetString(), marcReturn["MIN_SAFETY_STK"].GetString(),
            maraReturn["BASE_UOM"].GetString(), maraReturn["MATL_GROUP"].GetString(), maraReturn["SHELF_LIFE"].GetString(), maraReturn["MINREMLIFE"].GetString(), materialName);

            return ret;
        }

        public static DataTable GetBAPIMATERIALSTOCKREQLIST(string material, string plant)
        {
            RfcDestination destination = RfcConnection.GetDestination("0");
           
            IRfcFunction remoteQueryCall = destination.Repository.CreateFunction("BAPI_MATERIAL_STOCK_REQ_LIST");
            remoteQueryCall.SetValue("MATERIAL", String.Format("{0:000000000000000000}", Convert.ToInt64(material)));
            remoteQueryCall.SetValue("PLANT", plant);

            RfcSessionManager.BeginContext(destination);
            try
            {
                remoteQueryCall.Invoke(destination);
            }
            catch { Exception ex; }
            RfcSessionManager.EndContext(destination);

            IRfcStructure stockReturn = remoteQueryCall.GetStructure("MRP_STOCK_DETAIL");    

            DataTable ret = new DataTable();
            for (int j = 0; j < 15; j++)
            {
                ret.Columns.Add("Column" + j.ToString(), typeof(String));
            }

            ret.Rows.Add(material, plant, "'" + stockReturn["UNRESTRICTED_STCK"].GetString(), "'" + stockReturn["STCK_IN_TFR"].GetString(), "'" + stockReturn["QUAL_INSPECTION"].GetString(), "'" + stockReturn["RESTR_USE"].GetString(), "'" + stockReturn["BLKD_STKC"].GetString(),
                         "'" + stockReturn["BLKD_RETURNS"].GetString(), String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty);


            return ret;
        }

        public static IRfcTable GetPurchaseCockPit(string material, string plant)
        {
            RfcDestination destination = RfcConnection.GetDestination("0");

            IRfcFunction remoteQueryCall = destination.Repository.CreateFunction("ZMM_PURCHCOCKPIT_DEX");
            remoteQueryCall.SetValue("P_MATNR", String.Format("{0:000000000000000000}", 502323));
            remoteQueryCall.SetValue("P_BYEAR", 2012);
            remoteQueryCall.SetValue("P_PLSCN", 212);
            remoteQueryCall.SetValue("GET_TOTAL_LINES", "X");
            remoteQueryCall.SetValue("GET_IND_LINES", " ");

            RfcSessionManager.BeginContext(destination);
            try
            {
                remoteQueryCall.Invoke(destination);
            }
            catch { Exception ex; }
            RfcSessionManager.EndContext(destination);

            IRfcTable rfcReturn = remoteQueryCall.GetTable("MRP_TOTAL_LINES");
            return rfcReturn;
        }
        public static IRfcTable GetMDSTOCKREQUIREMENTSLISTAPI(string material, string plant)
        {
            RfcDestination destination = RfcConnection.GetDestination("0");

            IRfcFunction remoteQueryCall = destination.Repository.CreateFunction("MD_STOCK_REQUIREMENTS_LIST_API");
            remoteQueryCall.SetValue("MATNR", String.Format("{0:000000000000000000}", Convert.ToInt64(material)));
            remoteQueryCall.SetValue("WERKS", plant);

            RfcSessionManager.BeginContext(destination);
            try
            {
                remoteQueryCall.Invoke(destination);
            }
            catch { Exception ex; }
            RfcSessionManager.EndContext(destination);

            IRfcTable rfcReturn = remoteQueryCall.GetTable("MDEZX");
            //IRfcTable rfcReturn = remoteQueryCall.GetTable("E_MT61D");

            return rfcReturn;
        }
        public static IRfcTable GetRFCREADTABLE(string tableName, string[] options, string[] fieldNames)
        {
            RfcDestination destination = RfcConnection.GetDestination("0");

            IRfcFunction fieldList = destination.Repository.CreateFunction("RFC_READ_TABLE");
            fieldList.SetValue("QUERY_TABLE", tableName);
            fieldList.SetValue("DELIMITER", ";");
            fieldList.SetValue("NO_DATA", String.Empty);

            IRfcTable rfcOptions = fieldList.GetTable("OPTIONS");
            for (int i = 0; i < options.Count(); i++)
            {
                rfcOptions.Append();
                rfcOptions.SetValue("TEXT", options[i]);
            }

            IRfcTable fields = fieldList.GetTable("FIELDS");
            for (int j = 0; j < fieldNames.Count(); j++)
            {
                fields.Append();
                fields.SetValue("FIELDNAME", fieldNames[j]);
            }

            RfcSessionManager.BeginContext(destination);
            try
            {
                fieldList.Invoke(destination);
            }
            catch { Exception ex; }
            RfcSessionManager.EndContext(destination);

            IRfcTable resultTable = fieldList.GetTable("DATA");

            return resultTable;
        }
        public static IRfcTable GetCLAFCLASSIFICATIONOFOBJECTSAll(string material)
        {
            RfcDestination destination = RfcConnection.GetDestination("0");

            IRfcFunction remoteQueryCall = destination.Repository.CreateFunction("CLAF_CLASSIFICATION_OF_OBJECTS");
            remoteQueryCall.SetValue("CLASSTYPE", "001");
            remoteQueryCall.SetValue("OBJECT", String.Format("{0:000000000000000000}", Convert.ToInt64(material)));

            RfcSessionManager.BeginContext(destination);
            try
            {
                remoteQueryCall.Invoke(destination);
            }
            catch { Exception ex; }
            RfcSessionManager.EndContext(destination);

            IRfcTable rfcReturn = remoteQueryCall.GetTable("T_OBJECTDATA");
            //IRfcTable rfcReturn = remoteQueryCall.GetTable("E_MT61D");

            return rfcReturn;
        }
        public static IRfcTable GetBAPIMATERIALAVAILABILITY(string material, string plant)
        {
            RfcDestination destination = RfcConnection.GetDestination("0");

            IRfcFunction remoteQueryCall = destination.Repository.CreateFunction("BAPI_MATERIAL_AVAILABILITY");
            remoteQueryCall.SetValue("MATERIAL", String.Format("{0:000000000000000000}", Convert.ToInt64(material)));
            remoteQueryCall.SetValue("PLANT", plant);
            remoteQueryCall.SetValue("UNIT", "CT");

            RfcSessionManager.BeginContext(destination);
            try
            {
                remoteQueryCall.Invoke(destination);
            }
            catch { Exception ex; }
            RfcSessionManager.EndContext(destination);

            IRfcTable rfcReturn = remoteQueryCall.GetTable("WMDVEX");
            return rfcReturn;
        }
        public static DataTable GetWorkCenterGroups(string plant)
        {
            RfcDestination destination = RfcConnection.GetDestination("0");

            IRfcFunction fieldList = destination.Repository.CreateFunction("RFC_READ_TABLE");
            fieldList.SetValue("QUERY_TABLE", "CRHH");
            fieldList.SetValue("DELIMITER", ";");
            fieldList.SetValue("NO_DATA", String.Empty);

            IRfcTable options = fieldList.GetTable("OPTIONS");
            options.Append();
            options.SetValue("TEXT", "WERKS ='" + plant + "'");

            IRfcTable fields = fieldList.GetTable("FIELDS");
            fields.Append();
            fields.SetValue("FIELDNAME", "NAME");

            RfcSessionManager.BeginContext(destination);
            try
            {
                fieldList.Invoke(destination);
            }
            catch { Exception ex; }
            RfcSessionManager.EndContext(destination);

            IRfcTable resultTable = fieldList.GetTable("DATA");
            DataTable ret = new DataTable();
            ret.TableName = "WorkCenterGroups";
            ret.Columns.Add("ID", typeof(String));
            ret.Columns.Add("Text", typeof(String));

            for (int j = 0; j < resultTable.RowCount; j++)
            {
                resultTable.CurrentIndex = j;
                ret.Rows.Add(resultTable[j][0].ToString().Split('=').GetValue(1).ToString(),resultTable[j][0].ToString().Split('=').GetValue(1).ToString());               
            }
            return ret;
        }

        public static DataTable GetProductGroups(string plant)
        {
            RfcDestination destination = RfcConnection.GetDestination("0");

            IRfcFunction fieldList = destination.Repository.CreateFunction("RFC_READ_TABLE");
            fieldList.SetValue("QUERY_TABLE", "M_MAT2W");
            fieldList.SetValue("DELIMITER", ";");
            fieldList.SetValue("NO_DATA", String.Empty);

            IRfcTable options = fieldList.GetTable("OPTIONS");
            options.Append();
            options.SetValue("TEXT", "WERKS ='" + plant + "' AND SPRAS ='O'");

            IRfcTable fields = fieldList.GetTable("FIELDS");
            fields.Append();
            fields.SetValue("FIELDNAME", "PRGRP");
            fields.Append();
            fields.SetValue("FIELDNAME", "MAKTG");

            RfcSessionManager.BeginContext(destination);
            try
            {
                fieldList.Invoke(destination);
            }
            catch { Exception ex; }
            RfcSessionManager.EndContext(destination);

            IRfcTable resultTable = fieldList.GetTable("DATA");
            DataTable ret = new DataTable();
            ret.TableName = "ProductGroups";
            ret.Columns.Add("ID", typeof(String));
            ret.Columns.Add("Text", typeof(String));

            for (int j = 0; j < resultTable.RowCount; j++)
            {
                resultTable.CurrentIndex = j;
                ret.Rows.Add(resultTable[j][0].ToString().Split(new Char[] { '=', ';' }).GetValue(1).ToString().Trim(), String.Empty);
            }
            return ret;
        }
        public static DataTable GetProductGroupMaterias(string productGroup,string plant)
        {
            RfcDestination destination = RfcConnection.GetDestination("0");

            IRfcFunction fieldList = destination.Repository.CreateFunction("RFC_READ_TABLE");
            fieldList.SetValue("QUERY_TABLE", "PGMI");
            fieldList.SetValue("DELIMITER", ";");
            fieldList.SetValue("NO_DATA", String.Empty);

            IRfcTable options = fieldList.GetTable("OPTIONS");
            options.Append();
            options.SetValue("TEXT", "WERKS ='" + plant + "' AND PRGRP ='" + productGroup + "'");

            IRfcTable fields = fieldList.GetTable("FIELDS");
            fields.Append();
            fields.SetValue("FIELDNAME", "NRMIT");

            RfcSessionManager.BeginContext(destination);
            try
            {
                fieldList.Invoke(destination);
            }
            catch { Exception ex; }
            RfcSessionManager.EndContext(destination);

            IRfcTable resultTable = fieldList.GetTable("DATA");

            DataTable ret = new DataTable();
            ret.TableName = "Materials";
            ret.Columns.Add("MaterialNumber", typeof(Int32));            

            for (int i = 0; i < resultTable.RowCount; i++)
            {
                resultTable.CurrentIndex = i;
                if (IsNumeric(resultTable[i][0].ToString().Split(new Char[] { '=', ';' }).GetValue(1).ToString()))
                {
                    ret.Rows.Add(Convert.ToInt32(resultTable[i][0].ToString().Split(new Char[] { '=', ';' }).GetValue(1).ToString()));
                }
                else
                {
                    IRfcFunction fieldList1 = destination.Repository.CreateFunction("RFC_READ_TABLE");
                    fieldList1.SetValue("QUERY_TABLE", "ENT5271");
                    fieldList1.SetValue("DELIMITER", ";");
                    fieldList1.SetValue("NO_DATA", String.Empty);

                    IRfcTable options1 = fieldList.GetTable("OPTIONS");
                    options1.Append();
                    options1.SetValue("TEXT", "PRGRP ='" + productGroup + "'");

                    IRfcTable fields1 = fieldList.GetTable("FIELDS");
                    fields1.Append();
                    fields1.SetValue("FIELDNAME", "NRMIT");

                    RfcSessionManager.BeginContext(destination);
                    try
                    {
                        fieldList1.Invoke(destination);
                    }
                    catch { Exception ex; }
                    RfcSessionManager.EndContext(destination);

                    IRfcTable resultTable1 = fieldList.GetTable("DATA");
                    for (int j = 0; j < resultTable1.RowCount; j++)
                    {
                        resultTable1.CurrentIndex = j;
                        if (IsNumeric(resultTable1[j][0].ToString().Split(new Char[] { '=', ';' }).GetValue(1).ToString()))
                        {
                            ret.Rows.Add(Convert.ToInt32(resultTable1[j][0].ToString().Split(new Char[] { '=', ';' }).GetValue(1).ToString()));
                        }                       
                    }
                }
            }
            return ret;
        }

    }
}
