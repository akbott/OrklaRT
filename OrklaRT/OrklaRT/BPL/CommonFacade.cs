using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;


namespace OrklaRTBPL
{
    public static class CommonFacade
    {
        public static DateTime GetStartOfLastWeek()
        {
            int DaysToSubtract = (int)DateTime.Now.DayOfWeek + 6;
            DateTime dt = DateTime.Now.Subtract(TimeSpan.FromDays(DaysToSubtract));
            return new DateTime(dt.Year, dt.Month, dt.Day, 0, 0, 0, 0);
        }

        public static DateTime GetEndOfLastWeek()
        {
            DateTime dt = GetStartOfLastWeek().AddDays(6);
            return new DateTime(dt.Year, dt.Month, dt.Day, 23, 59, 59, 999);
        }

        public static DateTime GetStartOfCurrentWeek()
        {
            int DaysToSubtract = (int)DateTime.Now.DayOfWeek;
            DateTime dt = DateTime.Now.Subtract(TimeSpan.FromDays(DaysToSubtract));
            return new DateTime(dt.Year, dt.Month, dt.Day, 0, 0, 0, 0);
        }

        public static DateTime GetEndOfCurrentWeek()
        {
            DateTime dt = GetStartOfCurrentWeek().AddDays(6);
            return new DateTime(dt.Year, dt.Month, dt.Day, 23, 59, 59, 999);
        }
        public static DateTime GetStartOfCurrentMonth()
        {
            return new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1, 0, 0, 0, 0);
        }
        public static DateTime GetEndOfCurrentMonth()
        {
            return new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month), 23, 59, 59, 999);
        }
        public static DateTime GetStartOfLastMonth()
        {
            if (DateTime.Now.Month.Equals(1))
                return new DateTime(DateTime.Now.Year - 1, 12, 1, 0, 0, 0, 0);
            else
                return new DateTime(DateTime.Now.Year, DateTime.Now.Month - 1, 1, 0, 0, 0, 0);
        }
        public static DateTime GetDateTime()
        {
            return new DateTime(1, 1, 1, 0, 0, 0, 0);
        }
        public static DateTime Get2MonthsPeriod()
        {
            if (DateTime.Now.Month.Equals(1))
                return new DateTime(DateTime.Now.Year, 11, DateTime.Now.Day + 1, 0, 0, 0, 0);
            else
                return new DateTime(DateTime.Now.Year, DateTime.Now.Month - 2, DateTime.Now.Day + 1, 0, 0, 0, 0);
        }
        public static DateTime Get3MonthsPeriod()
        {
            return new DateTime(DateTime.Now.Year, DateTime.Now.Month + 2, DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month + 2), 0, 0, 0, 0);
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
        public static int GetUserID()

        {
            SQLServerHandler.IDBBuilder bldr = SQLDataHandler.Functions.GetBuilder();
            string commandString = bldr.GetSelectStatement("vwCurrentUser", new string[] { "ID" }, null, null);
            return SQLDataHandler.Functions.GetIntData(commandString);
        }
        public static string GetUserName(int userId)
        {
            SQLServerHandler.IDBBuilder bldr = SQLDataHandler.Functions.GetBuilder();
            string commandString = bldr.GetSelectStatement("CurrentUsers", new string[] { "UserName" }, new string[] { "ID" }, new object[] { userId });
            return SQLDataHandler.Functions.GetStringData(commandString);
        }

        public static int GetUserGroup(int userId)
        {
            SQLServerHandler.IDBBuilder bldr = SQLDataHandler.Functions.GetBuilder();
            string commandString = bldr.GetSelectStatement("UserGroupSetup", new string[] { "UserGroupId" }, new string[] { "UserId" }, new object[] { userId });
            return SQLDataHandler.Functions.GetIntData(commandString);
        }

        public static DataTable GetCurrentUser()
        {
            SQLServerHandler.IDBBuilder bldr = SQLDataHandler.Functions.GetBuilder();
            string commandString = bldr.GetSelectStatement("vwCurrentUser", new string[] { "*" }, null, null);
            return SQLDataHandler.Functions.GetData("CurrentUser",commandString).Tables[0];
        }

        public static DataTable GetUser(string userName)
        {
            SQLServerHandler.IDBBuilder bldr = SQLDataHandler.Functions.GetBuilder();
            string commandString = bldr.GetSelectStatement("CurrentUsers", new string[] { "*" }, new string[] { "UserName" }, new object[] { userName });
            return SQLDataHandler.Functions.GetData("CurrentUsers", commandString).Tables[0];
        }

        public static DataSet GetReports(int reportGroup)
        {
            SQLServerHandler.IDBBuilder bldr = SQLDataHandler.Functions.GetBuilder();
            string commandString = "SELECT ReportID,BeginGroup,ReportName,Enabled FROM Reports WHERE ReportGroup=" + reportGroup +" AND Enabled = 1 ORDER BY ReportID ASC";
            return SQLDataHandler.Functions.GetData("Reports",commandString);
        }

        public static DataTable GetReportDefinition(int reportID)
        {
            SQLServerHandler.IDBBuilder bldr = SQLDataHandler.Functions.GetBuilder();
            string commandString = "SELECT ReportName,ReportDefinition FROM Reports WHERE ReportID=" + reportID + "";
            return SQLDataHandler.Functions.GetData("Reports", commandString).Tables[0];
        }

        public static int GetVariantID(int UserID, int reportID)
        {
            int ret = 0;
            if (!SQLDataHandler.Functions.GetIntData("SELECT COUNT(*) FROM CurrentUserReportVariants WHERE UserID = " + UserID + " AND ReportID = " + reportID).Equals(0))
            {
                SQLServerHandler.IDBBuilder bldr = SQLDataHandler.Functions.GetBuilder();
                string commandString = bldr.GetSelectStatement("CurrentUserReportVariants", new string[] { "VariantID" }, new string[] { "UserID", "ReportID" }, new object[] { UserID, reportID });
                ret = SQLDataHandler.Functions.GetIntData(commandString);
            }
            return ret;
        }

        public static DataSet GetReportRightClickMenu(int reportID)
        {
            string commandString = "SELECT RCM.MenuID,RCM.Caption,RCM.FunctionName,RCMS.BeginGroup,RCMS.Enable FROM RightClickMenu RCM LEFT JOIN RightClickMenuSetup RCMS ON RCM.MenuID = RCMS.MenuID WHERE RCMS.ReportID IN (0, " + reportID + ",999) ORDER BY RCMS.ReportID";
            return SQLDataHandler.Functions.GetData("RightClickMenu", commandString);
        }

        public static DataSet GetCurrentUserReportFields(int userId,int reportId)
        {
            string commandString = "SELECT * FROM CurrentUserReportFields WHERE UserId = " + userId + " AND ReportId = " + reportId;
            return SQLDataHandler.Functions.GetData("CurrentUserReportFields", commandString);
        }

        public static DataSet GetRightClickReportMenuFields(string funtionName)
        {
            string commandString = "SELECT ValueFieldName,FieldValue FROM BDCInput WHERE TransactionCode = '" + funtionName + "' AND ValueFieldName IS NOT NULL";
            return SQLDataHandler.Functions.GetData("BDCInput", commandString);
        }

        public static void UpdateCurrentUserReportFields(string columnName, string value, int userID,int reportID)
        {
            SQLServerHandler.IDBBuilder bldr = SQLDataHandler.Functions.GetBuilder();
            string commandString = bldr.GetUpdateStatement("CurrentUserReportFields", new string[] { columnName }, new object[] { value }, new string[] { "UserId", "ReportId" }, new object[] { userID, reportID });
            SQLDataHandler.Functions.ExecuteNonQuery(commandString);
        }
        //public static DataSet GetReportRightClickReportMenuFields(int reportID)
        //{
        //    string commandString = "SELECT DISTINCT RCMF.FieldName FROM RightClickMenu RCM INNER JOIN RighClickMenuFields RCMF ON RCMF.RightClickMenuID = RCM.ID WHERE RCM.ReportID = " + reportID;
        //    return SQLDataHandler.Functions.GetData("RightClickMenuFields", commandString);
        //}
        //public static DataSet GetRightClickReportMenuFields(int rightClickMenuID, string funtionName)
        //{
        //    string commandString = "SELECT BDC.ValueFieldName,BDC.FieldValue FROM RightClickMenu RCM INNER JOIN RighClickMenuFields RCMF ON RCMF.RightClickMenuID = RCM.MenuID " +
        //                           "INNER JOIN BDCInput BDC ON BDC.ValueFieldName = RCMF.FieldName WHERE RCMF.RightClickMenuID = " + rightClickMenuID + " AND BDC.TransactionCode = '" + funtionName + "'";
        //    return SQLDataHandler.Functions.GetData("BDCInput", commandString);
        //}
        public static int GetRightClickSubMenuCount(int rightClickMenuID)
        {
            string commandString = "SELECT COUNT(*) AS SubMenuCount FROM RightClickSubMenu WHERE RightClickMenuID = " + rightClickMenuID;
            return SQLDataHandler.Functions.GetIntData(commandString);
        }
        public static void UpdateOrklaRTVersion(int userID, string version)
        {
            string commandString = "SELECT COUNT(*) FROM UserOrklaRTVersion WHERE UserId = " + userID;
            if (SQLDataHandler.Functions.GetIntData(commandString) > 0)
            {
                SQLServerHandler.IDBBuilder bldr = SQLDataHandler.Functions.GetBuilder();
                string commandString1 = bldr.GetUpdateStatement("UserOrklaRTVersion", new string[] { "Version" }, new object[] { version }, new string[] { "UserId" }, new object[] { userID });
                SQLDataHandler.Functions.ExecuteNonQuery(commandString1);
            }
            else
            {
                string commandString2 = "INSERT INTO UserOrklaRTVersion(UserID,Version) VALUES(" + userID + ", '" + version + "')";
                SQLDataHandler.Functions.ExecuteNonQuery(commandString2);
            }
        }
        public static int GetReportSheetCount(int reportID, string sheetName)
        {
            string commandString = "SELECT COUNT(*) AS SheetOptionsCount FROM ReportSheetOptions WHERE ReportID = " + reportID + " AND ReportSheet LIKE '" + sheetName + "'";
            return SQLDataHandler.Functions.GetIntData(commandString);
        }
        public static DataSet GetExchangeRates(string year)
        {
            string commandString = "SELECT FromCurrency,ToCurrency,ValidFrom,ExchangeRate FROM ExchangeRates WHERE ValidFrom = CONVERT(date,'01.01." + year + "',103)";
            return SQLDataHandler.Functions.GetData("BDCInput", commandString);
        }
        public static DataSet GetLockedOrders(string plant)
        {
            string commandString = "SELECT OrderNumber FROM PPLockedOrders WHERE Plant = '" + plant + "'";
            return SQLDataHandler.Functions.GetData("PPLockedOrders", commandString);
        }
        public static DataSet GetReportComments(int reportID)
        {
            string commandString = "SELECT CommentID,Comment,Date1 FROM ReportComments WHERE ReportID = " + reportID;
            return SQLDataHandler.Functions.GetData("ReportComments", commandString);
        }
        public static DataSet GetReportDescriptions(int reportID)
        {
            string commandString = "EXEC dbo.PrGetReportDescriptions " + reportID;
            return SQLDataHandler.Functions.GetData("Descriptions", commandString);
        }
        public static void CreateCurrentUser(string userName, int userGroup)
        {
            string commandString = "EXEC dbo.prCreateCurrentUser '" + userName + "'," + userGroup;
            SQLDataHandler.Functions.ExecuteNonQuery(commandString);
        }
        public static void InsertErrorLog(string logMessage, string methodName, string formName,int userID,int reportID)
        {
            string commandString = "INSERT INTO ErrorLog(LogMessage,MethodName,FormName,UserID,ReportID,Date) VALUES('" + logMessage + "','" + methodName + "','" + formName + "', " + userID + ", " + reportID + ",CONVERT(datetime,'" + DateTime.Now.ToString() + "',103))";
            SQLDataHandler.Functions.ExecuteNonQuery(commandString);
        }
        public static void InsertReportLog(int userID, int reportID)
        {
            string commandString = "EXEC dbo.prInsertReportLog " + userID + "," + reportID;
            SQLDataHandler.Functions.ExecuteNonQuery(commandString);
        }
        public static int WorkDaysBetween(string calenderID, DateTime fromDate, DateTime toDate)
        {
            string commandString = "SELECT CAST(DayType AS Int) AS DayType FROM PlantCalender WHERE CalenderID = '" + calenderID + "' AND Date BETWEEN CONVERT(date,'" + fromDate.ToShortDateString() + "',103) AND CONVERT(date,'" + toDate.ToShortDateString() + "',103)";
            return SQLDataHandler.Functions.GetIntData(commandString);
            
        }
        //public static void CreateCurrentUser(string calenderID, DateTime date, int WorkDays)
        //{
        //    string commandString = "SELECT SUBSTRING((CASE MONTH(CONVERT(date,'" + date.ToShortDateString() + "',103)) WHEN 1 THEN January WHEN 2 THEN February WHEN 3 THEN March  " +
        //                           "WHEN 4 THEN April WHEN 5 THEN May WHEN 6 THEN June WHEN 7 THEN July WHEN 8 THEN August WHEN 9 THEN September  " +
        //                           "WHEN 10 THEN October WHEN 11 THEN November WHEN 12 THEN December ELSE '' END),0,1) AS test FROM FactoryCalender  " +
        //                           "WHERE CalenderID = 'NOAR2013'";
        //    SQLDataHandler.Functions.ExecuteNonQuery(commandString);
        //}
        public static ADODB.Recordset ConvertToRecordset(DataTable inTable)
        {
            ADODB.Recordset result = new ADODB.Recordset();
            result.CursorLocation = ADODB.CursorLocationEnum.adUseClient;

            ADODB.Fields resultFields = result.Fields;
            System.Data.DataColumnCollection inColumns = inTable.Columns;

            foreach (DataColumn inColumn in inColumns)
            {
                resultFields.Append(inColumn.ColumnName
                    , TranslateType(inColumn.DataType, inColumn.ColumnName)
                    , inColumn.MaxLength
                    , inColumn.AllowDBNull ? ADODB.FieldAttributeEnum.adFldIsNullable :
                                             ADODB.FieldAttributeEnum.adFldUnspecified
                    , null);
            }

            result.Open(System.Reflection.Missing.Value
                    , System.Reflection.Missing.Value
                    , ADODB.CursorTypeEnum.adOpenStatic
                    , ADODB.LockTypeEnum.adLockOptimistic, 0);

            try
            {
                foreach (DataRow dr in inTable.Rows)
                {
                    result.AddNew(System.Reflection.Missing.Value,
                                  System.Reflection.Missing.Value);
                    System.Diagnostics.Debug.Write(dr.Table.Rows.IndexOf(dr));
                    for (int columnIndex = 0; columnIndex < inColumns.Count; columnIndex++)
                    {
                        //if (dr.Table.Rows.IndexOf(dr).Equals(2488) || dr.Table.Rows.IndexOf(dr).Equals(3870) || dr.Table.Rows.IndexOf(dr).Equals(5206) || dr.Table.Rows.IndexOf(dr).Equals(5207))
                        //{

                        //}
                        //if (dr.Table.Columns[columnIndex].DataType.Name.Equals("String"))
                        //{
                        //    resultFields[columnIndex].Value = dr[columnIndex].ToString();
                        //}
                        //else
                        //{
                        resultFields[columnIndex].Value = dr[columnIndex];
                        //}
                    }
                }
            }
            catch (Exception ex)
            { }

            return result;
        }
        public static ADODB.DataTypeEnum TranslateType(Type columnType, string columnName)
        {
            switch (columnType.UnderlyingSystemType.ToString())
            {
                case "System.Boolean":
                    return ADODB.DataTypeEnum.adBoolean;

                case "System.Byte":
                    return ADODB.DataTypeEnum.adUnsignedTinyInt;

                case "System.Char":
                    return ADODB.DataTypeEnum.adChar;

                case "System.DateTime":
                    return ADODB.DataTypeEnum.adDate;

                case "System.Decimal":
                    return ADODB.DataTypeEnum.adCurrency;

                case "System.Double":
                    return ADODB.DataTypeEnum.adDouble;

                case "System.Int16":
                    return ADODB.DataTypeEnum.adSmallInt;

                case "System.Int32":
                    return ADODB.DataTypeEnum.adInteger;

                case "System.Int64":
                    return ADODB.DataTypeEnum.adBigInt;

                case "System.SByte":
                    return ADODB.DataTypeEnum.adTinyInt;

                case "System.Single":
                    return ADODB.DataTypeEnum.adSingle;

                case "System.UInt16":
                    return ADODB.DataTypeEnum.adUnsignedSmallInt;

                case "System.UInt32":
                    return ADODB.DataTypeEnum.adUnsignedInt;

                case "System.UInt64":
                    return ADODB.DataTypeEnum.adUnsignedBigInt;

                case "System.String":
                default:
                    return ADODB.DataTypeEnum.adLongVarChar;
            }
        }
    }
}
        

