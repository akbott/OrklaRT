using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace OrklaRTBPL
{
    public class SelectionFacade
    {
        public static void UpdateCurrentUserReportSelectionLowValue(int reportID, int userID, string screenID, string kind, string sign, string selectionOption, string lowValue)
        {
            var entities = new DAL.SAPExlEntities();
            if (!entities.CurrentUserReportSelections.Where(p => p.ReportID == reportID && p.UserID == userID && p.ScreenID == screenID).Count().Equals(0))
            {
                if (!lowValue.Equals(String.Empty))
                {
                    var row = entities.CurrentUserReportSelections.SingleOrDefault(p => p.ReportID == reportID && p.UserID == userID  && p.ScreenID == screenID);
                    row.ReportID = reportID;
                    row.UserID = userID;
                    row.ScreenID = screenID;
                    row.Kind = kind;
                    row.Sign = sign;
                    row.SelectionOption = selectionOption;
                    row.LowValue = lowValue;
                    entities.SaveChanges();
                }
                else
                {
                    var currentUserReportSelections = new DAL.CurrentUserReportSelections();
                    currentUserReportSelections = entities.CurrentUserReportSelections.SingleOrDefault(p => p.ReportID == reportID && p.UserID == userID && p.ScreenID == screenID);
                    entities.CurrentUserReportSelections.Remove(currentUserReportSelections);                   
                    entities.SaveChanges();
                }
            }
            else
            {
                var row = new DAL.CurrentUserReportSelections();
                row.ReportID = reportID;
                row.UserID = userID;
                row.VariantID = 0;
                row.ScreenID = screenID;
                row.Kind = kind;
                row.Sign = sign;
                row.SelectionOption = selectionOption;
                row.LowValue = lowValue;
                row.HighValue = String.Empty;
                entities.CurrentUserReportSelections.Add(row);
                entities.SaveChanges();
            }
        }

        public static void UpdateCurrentUserReportSelectionHighValue(int reportID, int userID, string screenID, string kind, string sign, string selectionOption, string highValue)
        {
            var entities = new DAL.SAPExlEntities();
            if (!entities.CurrentUserReportSelections.Where(p => p.ReportID == reportID && p.UserID == userID  && p.ScreenID == screenID).Count().Equals(0))
            {
                var row = entities.CurrentUserReportSelections.SingleOrDefault(p => p.ReportID == reportID && p.UserID == userID  && p.ScreenID == screenID);
                row.ReportID = reportID;
                row.UserID = userID;
                row.ScreenID = screenID;
                row.Kind = kind;
                row.Sign = sign;
                row.SelectionOption = selectionOption;
                row.HighValue = highValue;
                entities.SaveChanges();
            }
            else
            {
                var row = new DAL.CurrentUserReportSelections();
                row.ReportID = reportID;
                row.UserID = userID;
                row.VariantID = 0;
                row.ScreenID = screenID;
                row.Kind = kind;
                row.Sign = sign;
                row.SelectionOption = selectionOption;
                row.LowValue = String.Empty;
                row.HighValue = highValue;
                entities.CurrentUserReportSelections.Add(row);
                entities.SaveChanges();
            }
        }
        //public static void UpdateUserReportSelectionLowValue(int reportID, int userID, int variantID ,string screenID, string kind, string sign, string selectionOption, string lowValue)
        //{
        //    var entities = new DAL.SAPExlEntities();
        //    if (!entities.UserReportSelections.Where(p => p.ReportID == reportID && p.UserID == userID && p.VariantID == variantID && p.ScreenID == screenID).Count().Equals(0))
        //    {
        //        if (!lowValue.Equals(String.Empty))
        //        {
        //            var row = entities.UserReportSelections.SingleOrDefault(p => p.ReportID == reportID && p.UserID == userID && p.VariantID == variantID && p.ScreenID == screenID);
        //            row.SelectionOption = selectionOption;
        //            row.LowValue = lowValue;
        //            entities.SaveChanges();
        //        }
        //        else
        //        {
        //            var userReportSelections = new DAL.UserReportSelections();
        //            userReportSelections = entities.UserReportSelections.SingleOrDefault(p => p.ReportID == reportID && p.VariantID == variantID && p.UserID == userID && p.ScreenID == screenID);
        //            entities.UserReportSelections.Remove(userReportSelections);
        //            entities.SaveChanges();
        //        }
        //    }
        //    else
        //    {
        //        var row = new DAL.UserReportSelections();
        //        row.ReportID = reportID;
        //        row.UserID = userID;
        //        row.VariantID = variantID;
        //        row.ScreenID = screenID;
        //        row.Kind = kind;
        //        row.Sign = sign;
        //        row.SelectionOption = selectionOption;
        //        row.LowValue = lowValue;
        //        row.HighValue = String.Empty;
        //        entities.UserReportSelections.Add(row);
        //        entities.SaveChanges();
        //    }
        //}
        //public static void UpdateUserReportSelectionHighValue(int reportID, int userID, int variantID ,string screenID, string kind, string sign, string selectionOption, string highValue)
        //{
        //    var entities = new DAL.SAPExlEntities();
        //    if (!entities.UserReportSelections.Where(p => p.ReportID == reportID && p.UserID == userID && p.VariantID == variantID && p.ScreenID == screenID).Count().Equals(0))
        //    {
        //        var row = entities.UserReportSelections.SingleOrDefault(p => p.ReportID == reportID && p.UserID == userID && p.VariantID == variantID && p.ScreenID == screenID);
        //        row.HighValue = highValue;
        //        entities.SaveChanges();
        //    }
        //    else
        //    {
        //        var row = new DAL.UserReportSelections();
        //        row.ReportID = reportID;
        //        row.UserID = userID;
        //        row.VariantID = variantID;
        //        row.ScreenID = screenID;
        //        row.Kind = kind;
        //        row.Sign = sign;
        //        row.SelectionOption = selectionOption;
        //        row.LowValue = String.Empty;
        //        row.HighValue = highValue;
        //        entities.UserReportSelections.Add(row);
        //        entities.SaveChanges();
        //    }
        //}
        public static string GetLanguageKeyScreenID(int reportID)
        {
            string screenID = String.Empty;
            if (!SQLDataHandler.Functions.GetIntData("SELECT COUNT(*) FROM ReportSelections WHERE FieldName = 'Language Key' AND ControlType IS NULL AND ReportID = " + reportID).Equals(0))
            {
                string commandString = "SELECT ScreenID FROM ReportSelections WHERE FieldName = 'Language Key' AND ControlType IS NULL AND ReportID = " + reportID;
                screenID = SQLDataHandler.Functions.GetStringData(commandString);
            }
            return screenID;
        }
        public static string InsertLanguageCodes(int reportID, int userID, string language)
        {
            string screenID = String.Empty;
            if (!SQLDataHandler.Functions.GetIntData("SELECT COUNT(*) FROM ReportSelections WHERE FieldName = 'Language Key' AND ControlType IS NULL AND ReportID = " + reportID).Equals(0))
            {
                SQLDataHandler.Functions.ExecuteNonQuery("DELETE FROM CurrentUserReportSelections WHERE  ReportID = " + reportID + " AND UserID = " + userID + " AND ScreenID IN (SELECT ScreenID FROM ReportSelections WHERE ReportID = " + reportID + " AND FieldName = 'Language Key' AND UpdateLanguage <> 1)");
                string commandString = "INSERT INTO CurrentUserReportSelections(ReportID,UserID,VariantID,ScreenID,Kind,Sign,SelectionOption,LowValue,HighValue,MultipleSelection) (SELECT " + reportID + "," + userID + ", 0,ScreenID,'S',Sign,SelectionOption,'" + language + "',HighValue,MultipleSelection FROM ReportSelections WHERE ReportID = " + reportID + " AND FieldName = 'Language Key' AND UpdateLanguage <> 1)";
                SQLDataHandler.Functions.ExecuteNonQuery(commandString);
            }
            return screenID;
        }
        public static DataSet GetUserReportVariants(int userID, int reportID)
        {
            string commandString = "SELECT 0 AS ID, 'Standard' AS VariantName UNION ALL SELECT ID,VariantName FROM UserReportVariants WHERE UserID = " + userID + " AND ReportID = " + reportID + " ORDER BY ID";
            return SQLDataHandler.Functions.GetData("UserReportVariants", commandString);
        }
        public static DataSet GetUserReportVariants(int variantID)
        {
            string commandString = "SELECT ID,VariantName,VariantDescription FROM UserReportVariants WHERE ID = " + variantID;
            return SQLDataHandler.Functions.GetData("UserReportVariants", commandString);
        }
        public static string GetCurrentUserFirstMultipleValue(int reportID, int userID, string screenID)
        {
            string commandString = "SELECT TOP 1 LowValue FROM CurrentUserReportMultipleSelections WHERE UserID = " + userID + " AND ReportID = " + reportID + " AND VariantID = 0 AND ScreenID = '" + screenID + "' ORDER BY ID ASC";
            return SQLDataHandler.Functions.GetStringData(commandString);
        }
        public static int GetCurrentUserMultipleValueCount(int reportID, int userID, string screenID)
        {
            string commandString = "SELECT COUNT(*) FROM CurrentUserReportMultipleSelections WHERE UserID = " + userID + " AND ReportID = " + reportID + " AND VariantID = 0 AND ScreenID = '" + screenID + "'";
            return SQLDataHandler.Functions.GetIntData(commandString);
        }
        public static int GetCurrentUserVariant(int reportID, int userID)
        {
            int variant = 0;
            if (!SQLDataHandler.Functions.GetIntData("SELECT COUNT(*) FROM CurrentUserReportVariants WHERE UserID = " + userID + " AND ReportID = " + reportID).Equals(0))
            {
                string commandString = "SELECT VariantID FROM CurrentUserReportVariants WHERE UserID = " + userID + " AND ReportID = " + reportID;
                variant = SQLDataHandler.Functions.GetIntData(commandString);
            }
            return variant;
        }
        public static int GetLastVariantID(int userID, int reportID)
        {
            string commandString = "SELECT TOP 1 ID FROM UserReportVariants WHERE ReportID = " + reportID + " AND UserID = " + userID + " ORDER BY ID DESC";
            return SQLDataHandler.Functions.GetIntData(commandString);
        }
        public static void InsertUserReportVariant(int reportID, int userID, string variantName, string description)
        {
            SQLServerHandler.IDBBuilder bldr = SQLDataHandler.Functions.GetBuilder();
            string commandString = bldr.GetInsertStatement("UserReportVariants", new string[] { "ReportID", "UserID", "VariantName", "VariantDescription" }, new object[] { reportID, userID, variantName, description });
            SQLDataHandler.Functions.ExecuteNonQuery(commandString);
        }
        public static void DeleteUserReportVariant(int reportID, int userID, int variantID)
        {
            SQLServerHandler.IDBBuilder bldr = SQLDataHandler.Functions.GetBuilder();
            string commandString = bldr.GetDeleteStatement("UserReportVariants", new string[] { "ReportID", "UserID", "ID" }, new object[] { reportID, userID, variantID });
            SQLDataHandler.Functions.ExecuteNonQuery(commandString);
            string commandString1 = bldr.GetDeleteStatement("UserReportSelections", new string[] { "ReportID", "UserID", "VariantID" }, new object[] { reportID, userID, variantID });
            SQLDataHandler.Functions.ExecuteNonQuery(commandString1);
            string commandString2 = bldr.GetDeleteStatement("UserReportMultipleSelections", new string[] { "ReportID", "UserID", "VariantID"}, new object[] { reportID, userID, variantID });
            SQLDataHandler.Functions.ExecuteNonQuery(commandString2);
            
        }
        public static void DeleteCurrentUserMultipleSelections(int reportID, int userID, string screenID)
        {
            SQLServerHandler.IDBBuilder bldr = SQLDataHandler.Functions.GetBuilder();
            string commandString = bldr.GetDeleteStatement("CurrentUserReportMultipleSelections", new string[] { "ReportID", "UserID", "VariantID", "ScreenID" }, new object[] { reportID, userID, 0, screenID });
            SQLDataHandler.Functions.ExecuteNonQuery(commandString);
            string commandString1 = bldr.GetDeleteStatement("CurrentUserReportSelections", new string[] { "ReportID", "UserID", "ScreenID" }, new object[] { reportID, userID, screenID });
            SQLDataHandler.Functions.ExecuteNonQuery(commandString1);
        }
        public static void DeleteCurrentUserMultipleSelections(int id)
        {
            SQLServerHandler.IDBBuilder bldr = SQLDataHandler.Functions.GetBuilder();
            string commandString = bldr.GetDeleteStatement("CurrentUserReportMultipleSelections", new string[] { "ID" }, new object[] { id });
            SQLDataHandler.Functions.ExecuteNonQuery(commandString);
        }
        public static void DeleteEmptyCurrentUserReportSelections(int reportID, int userID)
        {
            SQLServerHandler.IDBBuilder bldr = SQLDataHandler.Functions.GetBuilder();
            string commandString = bldr.GetDeleteStatement("CurrentUserReportSelections", new string[] { "ReportID", "UserID", "VariantID", "LowValue", "HighValue" }, new object[] { reportID, userID, 0, String.Empty, String.Empty });
            SQLDataHandler.Functions.ExecuteNonQuery(commandString);
        }
        public static void DeleteCurrentUserReportSelections(int reportID,int userID)
        {
            SQLServerHandler.IDBBuilder bldr = SQLDataHandler.Functions.GetBuilder();
            string commandString = bldr.GetDeleteStatement("CurrentUserReportSelections", new string[] { "ReportID","UserID", "VariantID" }, new object[] { reportID,userID, 0});
            SQLDataHandler.Functions.ExecuteNonQuery(commandString);
            string commandString1 = bldr.GetDeleteStatement("CurrentUserReportMultipleSelections", new string[] { "ReportID", "UserID", "VariantID" }, new object[] { reportID, userID, 0 });
            SQLDataHandler.Functions.ExecuteNonQuery(commandString1);
            
        }


        public static void DeleteCurrentUserReportSelectionTempFields(int reportID, int userID)
        {
            SQLServerHandler.IDBBuilder bldr = SQLDataHandler.Functions.GetBuilder();
            string commandString = bldr.GetDeleteStatement("CurrentUserReportSelections", new string[] { "ReportID", "UserID", "VariantID", "ScreenID" }, new object[] { reportID, userID, 0, "SP$00000" });
            SQLDataHandler.Functions.ExecuteNonQuery(commandString);
            string commandString1 = bldr.GetDeleteStatement("CurrentUserReportSelections", new string[] { "ReportID", "UserID", "VariantID", "ScreenID" }, new object[] { reportID, userID, 0, "SP$11111" });
            SQLDataHandler.Functions.ExecuteNonQuery(commandString1);
        }


        public static void InsertUserReportSelections(int reportID,int userID,int variantID)
        {
            SQLServerHandler.IDBBuilder bldr = SQLDataHandler.Functions.GetBuilder();
            string commandString1 = bldr.GetDeleteStatement("UserReportSelections", new string[] { "ReportID", "UserID", "VariantID" }, new object[] { reportID, userID, variantID });
            SQLDataHandler.Functions.ExecuteNonQuery(commandString1);
            string commandString2 = bldr.GetDeleteStatement("UserReportMultipleSelections", new string[] { "ReportID", "UserID", "VariantID" }, new object[] { reportID, userID, variantID });
            SQLDataHandler.Functions.ExecuteNonQuery(commandString2);
            string commandString3 = "INSERT INTO UserReportSelections(ReportID,UserID,VariantID,ScreenID,Kind,Sign,SelectionOption,LowValue,HighValue,MultipleSelection) (SELECT " + reportID + "," + userID + ", " + variantID + ",ScreenID,Kind,Sign,SelectionOption,LowValue,HighValue,MultipleSelection FROM CurrentUserReportSelections WHERE ReportID = " + reportID + " AND UserID = " + userID + " AND ScreenID IN (SELECT ScreenID FROM ReportSelections WHERE ReportID = " + reportID + " AND Include <> 0))";
            SQLDataHandler.Functions.ExecuteNonQuery(commandString3);
            string commandString4 = "INSERT INTO UserReportMultipleSelections(ReportID,UserID,VariantID,ScreenID,LowValue) (SELECT ReportID,UserID," + variantID + ",ScreenID,LowValue FROM CurrentUserReportMultipleSelections WHERE ReportID = " + reportID + " AND UserID = " + userID + ")";
            SQLDataHandler.Functions.ExecuteNonQuery(commandString4);
        }
        public static void InsertCurrentUserReportSelections(int reportID, int userID, int variantID)
        {
            SQLServerHandler.IDBBuilder bldr = SQLDataHandler.Functions.GetBuilder();
            string commandString1 = bldr.GetDeleteStatement("CurrentUserReportSelections", new string[] { "ReportID", "UserID" }, new object[] { reportID, userID });
            SQLDataHandler.Functions.ExecuteNonQuery(commandString1);
            string commandString2 = bldr.GetDeleteStatement("CurrentUserReportMultipleSelections", new string[] { "ReportID", "UserID"}, new object[] { reportID, userID });
            SQLDataHandler.Functions.ExecuteNonQuery(commandString2);
            string commandString3 = "INSERT INTO CurrentUserReportSelections(ReportID,UserID,VariantID,ScreenID,Kind,Sign,SelectionOption,LowValue,HighValue,MultipleSelection) (SELECT " + reportID + "," + userID + ", 0,ScreenID,Kind,Sign,SelectionOption,LowValue,HighValue,MultipleSelection FROM UserReportSelections WHERE ReportID = " + reportID + " AND UserID = " + userID + " AND VariantID = " + variantID + ")";
            SQLDataHandler.Functions.ExecuteNonQuery(commandString3);
            string commandString4 = "INSERT INTO CurrentUserReportMultipleSelections(ReportID,UserID,VariantID,ScreenID,LowValue) (SELECT ReportID,UserID,0,ScreenID,LowValue FROM UserReportMultipleSelections WHERE ReportID = " + reportID + " AND UserID = " + userID + " AND VariantID = " + variantID + ")";
            SQLDataHandler.Functions.ExecuteNonQuery(commandString4);
        }

        public static void InsertReportSelectionToCurrentUserReportSelections(int reportID, int userID)
        {
            SQLServerHandler.IDBBuilder bldr = SQLDataHandler.Functions.GetBuilder();
            string commandString1 = bldr.GetDeleteStatement("CurrentUserReportSelections", new string[] { "ReportID", "UserID" }, new object[] { reportID, userID });
            SQLDataHandler.Functions.ExecuteNonQuery(commandString1);
            string commandString3 = "INSERT INTO CurrentUserReportSelections(ReportID,UserID,VariantID,ScreenID,Kind,Sign,SelectionOption,LowValue,HighValue,MultipleSelection) (SELECT " + reportID + "," + userID + ", 0,ScreenID,'S',Sign,SelectionOption,LowValue,HighValue,MultipleSelection FROM ReportSelections WHERE ReportID = " + reportID + ")";
            SQLDataHandler.Functions.ExecuteNonQuery(commandString3);
        }

        public static void InsertMixingPlanReportSelectionToCurrentUserReportSelections(int reportID, int userID)
        {

            InsertReportSelectionToCurrentUserReportSelections(reportID, userID);
            UpdateCurrentUserReportSelectionLowValue(reportID, userID, "0P_PLANT", "S", "I", "EQ", OrklaRTBPL.SelectionFacade.ProductionPlanSelectionPlant);
            UpdateCurrentUserReportSelectionLowValue(reportID, userID, "ZVR0062", "S", "I", "BT", DateTime.Now.AddDays(-30).Date.ToString("yyyyMMdd"));
            UpdateCurrentUserReportSelectionHighValue(reportID, userID, "ZVR0062", "S", "I", "BT", DateTime.Now.AddDays(14).ToString("yyyyMMdd"));
            UpdateCurrentUserReportSelectionLowValue(reportID, userID, "SP$00000", "S", "I", "EQ", DateTime.Now.Date.ToString("yyyyMMdd"));
        }
        public static bool GetCurrentUserMultipleSelectionValue(int reportID, int userID, string screenID)
        {
            bool ret = false;
            var entities = new DAL.SAPExlEntities();
            {
                if (!entities.CurrentUserReportSelections.Where(p => p.ReportID == reportID && p.UserID == userID  && p.ScreenID == screenID).Count().Equals(0))
                {
                    ret = entities.CurrentUserReportSelections.SingleOrDefault(p => p.ReportID == reportID && p.UserID == userID  && p.ScreenID == screenID).MultipleSelection;
                }
            }
            return ret;
        }
        public static void UpdateCurrentUserReportVariants(int reportID, int userID, int variantID)
        {
            var entities = new DAL.SAPExlEntities();
            if (!entities.CurrentUserReportVariants.Where(p => p.ReportID == reportID && p.UserID == userID).Count().Equals(0))
            {
                var row = entities.CurrentUserReportVariants.SingleOrDefault(p => p.ReportID == reportID && p.UserID == userID);
                row.ReportID = reportID;
                row.UserID = userID;
                row.VariantID = variantID;
                entities.SaveChanges();
            }
            else
            {
                var row = new DAL.CurrentUserReportVariants();
                row.ReportID = reportID;
                row.UserID = userID;
                row.VariantID = variantID;
                entities.CurrentUserReportVariants.Add(row);
                entities.SaveChanges();
            }
        }
        public static void InsertCurrentUserReportMultipleSelections(int reportID, int userID, string screenID, string lowValue)
        {
            var entities = new DAL.SAPExlEntities();
            {
                var row = new DAL.CurrentUserReportMultipleSelections();
                row.ReportID = reportID;
                row.UserID = userID;
                row.VariantID = 0;
                row.ScreenID = screenID;
                row.LowValue = lowValue;
                entities.CurrentUserReportMultipleSelections.Add(row);
                entities.SaveChanges();
            }
        }
        public static void UpdateCurrentUserReportSelectionMultiplSelected(int reportID, int userID, string screenID,bool multipleSelected)
        {
            var entities = new DAL.SAPExlEntities();
            var row = entities.CurrentUserReportSelections.SingleOrDefault(p => p.ReportID == reportID && p.UserID == userID && p.ScreenID == screenID);
            row.MultipleSelection = multipleSelected;
            entities.SaveChanges();
        }
        public static void UpdateCurrentUserReportMultipleSelectionValue(int reportID, int userID, string screenID, string lowValue,bool Combo)
        {
            var entities = new DAL.SAPExlEntities();
            {
                if (entities.CurrentUserReportMultipleSelections.Where(p => p.ReportID == reportID && p.UserID == userID && p.ScreenID == screenID && p.ComboBox == Combo).Count() > 0)
                {
                    var row = entities.CurrentUserReportMultipleSelections.SingleOrDefault(p => p.ReportID == reportID && p.UserID == userID && p.ScreenID == screenID && p.ComboBox == Combo);
                    row.LowValue = lowValue;
                    entities.SaveChanges();
                }
                else
                {
                    if (entities.CurrentUserReportMultipleSelections.Where(p => p.ReportID == reportID && p.UserID == userID && p.ScreenID == screenID && p.ComboBox == Combo && p.LowValue == lowValue).Count().Equals(0))
                    {
                        var row = new DAL.CurrentUserReportMultipleSelections();
                        row.ReportID = reportID;
                        row.UserID = userID;
                        row.VariantID = 0;
                        row.ScreenID = screenID;
                        row.LowValue = lowValue;
                        row.ComboBox = Combo;
                        entities.CurrentUserReportMultipleSelections.Add(row);
                        entities.SaveChanges();
                    }
                }
            }
        }
        public static void UpdateCurrentUserReportMultipleSelectionValue(int id, string lowValue)
        {
            var entities = new DAL.SAPExlEntities();
            var row = entities.CurrentUserReportMultipleSelections.SingleOrDefault(p => p.ID == id);
            row.LowValue  = lowValue;
            entities.SaveChanges();
        }
        public static void UpdateCurrentUserReportCount(int reportID, int userID)
        {
            var entities = new DAL.SAPExlEntities();
            if (!entities.ReportStatistics.Where(p => p.ReportID == reportID && p.UserID == userID).Count().Equals(0))
            {
                var row = entities.ReportStatistics.SingleOrDefault(p => p.ReportID == reportID && p.UserID == userID);
                row.ReportCount = row.ReportCount + 1;
                row.LastOpened = DateTime.Now;
                entities.SaveChanges();
            }
            else
            {
                var row = new DAL.ReportStatistics();
                row.ReportID = reportID;
                row.UserID = userID;
                row.ReportCount = 1;
                row.LastOpened = DateTime.Now;
                entities.ReportStatistics.Add(row);
                entities.SaveChanges();
            }
        }
        //public static void UpdateDateOptions(int reportID,int userID,int variantID,int rollingDays,int dayType,bool useStandard)
        //{
        //    SQLServerHandler.IDBBuilder bldr = SQLDataHandler.Functions.GetBuilder();
        //    string commandString = bldr.GetUpdateStatement("UserReportSelections", new string[] { "RollingDays", "DayType", "UseStandard" }, new object[] { rollingDays, dayType, useStandard == true ? 1 : 0 }, new string[] { "ReportID", "UserID", "VariantID", }, new object[] { reportID, userID, variantID });
        //    SQLDataHandler.Functions.ExecuteSqlCommand(commandString);
        //}
        public static int CheckMandatoryFields(int reportID,int userID)
        {
            //string commandString = "SELECT COUNT(*) FROM ReportSelections WHERE ReportID = " + reportID + " AND Mandatory = 1 AND ScreenID NOT IN " +
            //                       "(SELECT ScreenID FROM CurrentUserReportSelections WHERE ReportID = " + reportID + " AND UserID = " + userID + " AND VariantID = 0 AND MultipleSelection <> 1 " +
            //                       "UNION ALL SELECT ScreenID FROM CurrentUserReportMultipleSelections WHERE ReportID = " + reportID + " AND UserID = " + userID + " AND VariantID = 0)";
            string commandString = "SELECT COUNT(*) FROM ReportSelections WHERE ReportID = " + reportID + " AND Mandatory = 1 AND " +
                                    "ScreenID IN (SELECT ScreenID FROM CurrentUserReportSelections WHERE ReportID = " + reportID + " AND UserID = " + userID + " AND VariantID = 0 AND " +
                                    "MultipleSelection <> 1 AND LowValue IS NULL)";
            return SQLDataHandler.Functions.GetIntData(commandString);
        }

        public static bool CheckRequiredFields(int reportID, int userID)
        {
            bool ret;
            string commandString = "SELECT COUNT(*) FROM ReportSelections WHERE ReportID = " + reportID + " AND Mandatory = 2 AND " +
                                    "ScreenID IN (SELECT ScreenID FROM CurrentUserReportSelections WHERE ReportID = " + reportID + " AND UserID = " + userID + " AND VariantID = 0)";            
            string commandString1 = "SELECT COUNT(*) FROM ReportSelections WHERE ReportID = " + reportID + " AND Mandatory = 2 AND " +
                                    "ScreenID IN (SELECT ScreenID FROM CurrentUserReportSelections WHERE ReportID = " + reportID + " AND UserID = " + userID + " AND VariantID = 0 AND (LowValue IS NULL OR LowValue = ' '))";
            if (!SQLDataHandler.Functions.GetIntData(commandString).Equals(0))
            {
                if (SQLDataHandler.Functions.GetIntData(commandString1) < SQLDataHandler.Functions.GetIntData(commandString))
                    ret = true;
                else
                    ret = false;
            }
            else
            {
                ret = true;
            }
            return ret;
        }

        public static string ReportSelectionLanguage;

        public static string ProductionPlanSelectionPlant;
        public static string ProductionPlanSelectionDate;

        public static string CapacityLevellingSelectionPlant;
        public static string CapacityLevellingSelectionWorkCenter;
        public static string CapacityLevellingWorkGroupCenter;

        public static string DeliveryAgentSelectionPlant;

        public static string PurchaseCockpitSelectionMaterial;
        public static string PurchaseCockpitSelectionYear;
        public static string PurchaseCockpitSelectionScenario;

        public static string MD04SelectionPlant;

        public static string MixingPlanSelectionPlant;
        public static string MixingPlanProdPlanSelectionDate;

        public static string StockTransferSelectionWarehouse;

        public static string ScrappingOverviewSelectionLanguage;

        public static string StockValuesAndCoverageProdPlanSelectionPlant;

        public static string ShelfLifeSelectionFirmakode;
        public static string ShelfLifeSelectionPlant;
        public static string ShelfLifeSelectionStorageLocation;
        public static string ShelfLifeSelectionMaterialType;      
        public static string ShelfLifeSelectionLanguage;

        public static string StockSimulationSelectionPlant;
        public static string StockSimulationSelectionMaterial;
        public static string StockSimulationSelectionFromDate;

        public static string StockHistorySelectionPlant;
        public static string StockHistorySelectionMaterial;
        public static string StockHistorySelectionFromDate;

        public static string SalesOrderSelectionSalesOrg;

        public static string DailyProductionPlanPlant;
        public static DataSet GetMaterials(string plant, string materialType)
        {
            string commandString = "SELECT CAST(Material AS varchar) AS ID,MaterialDescription AS Text, (CAST(Material AS varchar) + '   ' + MaterialDescription) AS CombinedText FROM Materials WHERE Plant = '" + plant + "' AND MaterialType = '" + materialType + "'";
            return SQLDataHandler.Functions.GetData("Materials", commandString);
        }

        public static DataSet GetProductGroupMaterials(string plant, string productGroup)
        {
            string commandString = "SELECT Material FROM ProductGroupMaterials WHERE Plant = '" + plant + "' AND ProductGroup = '" + productGroup + "'";
            return SQLDataHandler.Functions.GetData("ProductGroupMaterials", commandString);
        }

        public static DataSet GetWorkCenters(string plant)
        {
            string commandString = "SELECT [ID],[Text],([ID]+'   ' +[Text]) AS CombinedText FROM vwWorkCenters WHERE FilterField = '" + plant + "'";
            return SQLDataHandler.Functions.GetData("WorkCenters", commandString);
        }

        public static DataSet GetWorkCenterGroups(string plant)
        {
            string commandString = "SELECT [ID],[Text] FROM vwWorkCenterGroups WHERE FilterField = '" + plant + "'";
            return SQLDataHandler.Functions.GetData("WorkCenterGroups", commandString);
        }

        public static DataSet GetFixedProductionPlants()
        {
            string commandString = "SELECT [ID],[Text] FROM vwPlants WHERE [ID] IN ('CZBY','NOAR','NOEL','NOLA','PLM2')";
            return SQLDataHandler.Functions.GetData("vwPlants", commandString);
        }

        public static DataSet GetDailyProductionPlants()
        {
            string commandString = "SELECT [ID],[Text] FROM vwPlants WHERE [ID] IN ('LA00')";
            return SQLDataHandler.Functions.GetData("vwPlants", commandString);
        }

        public static DataSet GetMRPControllers(string plant)
        {
            string commandString = "SELECT [ID],[Text] FROM vwMRPControllers WHERE FilterField = '" + plant + "'";
            return SQLDataHandler.Functions.GetData("MRPControllers", commandString);
        }

        public static DataSet GetStorageTypes(string warehouseNumber)
        {
            string commandString = "SELECT [ID],[Text] FROM vwStorageTypes WHERE FilterField = '" + warehouseNumber + "'";
            return SQLDataHandler.Functions.GetData("StorageTypes", commandString);
        }

        public static DataSet GetSubReportSelection(int report)
        {
            string commandString = "SELECT FieldName,ScreenID,SelectionOption,Sign FROM ReportSelections WHERE FieldName NOT LIKE '%Field%' AND ReportID =" + report;
            return SQLDataHandler.Functions.GetData("ReportSelections", commandString);
        }

        public static DataSet GetBrands()
        {
            string commandString = "SELECT * FROM vwBrands";
            return SQLDataHandler.Functions.GetData("Brands", commandString);
        }

        public static DataSet GetPlants()
        {
            string commandString = "SELECT * FROM vwPlants";
            return SQLDataHandler.Functions.GetData("Plants", commandString);
        }

        public static DataSet GetMaterialTypes()
        {
            string commandString = "SELECT * FROM vwMaterialTypes";
            return SQLDataHandler.Functions.GetData("MaterialTypes", commandString);
        }

        public static DataSet GetCompanyCodes()
        {
            string commandString = "SELECT * FROM vwCompanyCodes";
            return SQLDataHandler.Functions.GetData("CompanyCodes", commandString);
        }

        public static DataSet GetSalesOrganizations()
        {
            string commandString = "SELECT * FROM vwSalesOrganizations";
            return SQLDataHandler.Functions.GetData("SalesOrganizations", commandString);
        }

        public static DataSet GetMaterialGroups()
        {
            string commandString = "SELECT * FROM vwMaterialGroups";
            return SQLDataHandler.Functions.GetData("MaterialGroups", commandString);
        }

        public static DataSet GetProductionScheduler()
        {
            string commandString = "SELECT * FROM vwProductionScheduler";
            return SQLDataHandler.Functions.GetData("ProductionScheduler", commandString);
        }

        public static DataSet GetPurchasingGroups()
        {
            string commandString = "SELECT * FROM vwPurchasingGroups";
            return SQLDataHandler.Functions.GetData("PurchasingGroups", commandString);
        }

        public static DataSet GetWarehouseNumbers()
        {
            string commandString = "SELECT * FROM vwWarehouseNumbers";
            return SQLDataHandler.Functions.GetData("WarehouseNumbers", commandString);
        }

        public static DataSet GetStorageLocations()
        {
            string commandString = "SELECT * FROM vwStorageLocations";
            return SQLDataHandler.Functions.GetData("StorageLocations", commandString);
        }

        public static DataSet GetMRPControllers()
        {
            string commandString = "SELECT * FROM vwMRPControllers";
            return SQLDataHandler.Functions.GetData("MRPControllers", commandString);
        }

        public static DataSet GetProductResponsibles()
        {
            string commandString = "SELECT * FROM vwProductResponsibles";
            return SQLDataHandler.Functions.GetData("ProductResponsibles", commandString);
        }
        public static DataSet GetMaterialArts()
        {
            string commandString = "SELECT * FROM vwMaterialArts";
            return SQLDataHandler.Functions.GetData("MaterialArts", commandString);
        }
    }
}
