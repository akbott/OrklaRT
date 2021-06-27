using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace OrklaRTBPL
{
    public class PivotFacade
    {
        public static int GetPivotLayoutVariantID(int reportID, int userID, string variantName, string variantDescription)
        {
            var entities = new DAL.SAPExlEntities();
            return entities.PivotLayoutVariants.SingleOrDefault(p => p.ReportID == reportID && p.UserID == userID && p.VariantName == variantName && p.VariantDescription == variantDescription).ID;
        }
        public static void SavePivotComment(int reportID,string commentID, string comment, DateTime date,string user)
        {
            var entities = new DAL.SAPExlEntities();
            if (!entities.ReportComments.Where(p => p.CommentID == commentID && p.ModifiedBy == user).Count().Equals(0))
            {
                var row = entities.ReportComments.Where(p => p.CommentID == commentID && p.ModifiedBy == user).SingleOrDefault();
                row.Comment = comment;
                row.Date1 = date.Date;
                row.ModifiedBy = user;
                row.ModifiedDate = DateTime.Now.Date;
                entities.SaveChanges();

            }
            else
            {
                var row = new DAL.ReportComments();
                row.ReportID = reportID;
                row.CommentID = commentID;
                row.Comment = comment;
                row.Date1 = date.Date;
                row.ModifiedBy = user;
                row.ModifiedDate = DateTime.Now.Date;
                entities.ReportComments.Add(row);
                entities.SaveChanges();
            }
        }
        public static void SavePivotLayoutVariant(int reportID, int userID, string variantName,string variantDescription)
        {
            var entities = new DAL.SAPExlEntities();
            if (!entities.PivotLayoutVariants.Where(p => p.ReportID == reportID && p.UserID == userID && p.VariantName == variantName && p.VariantDescription == variantDescription).Count().Equals(0))
            {
                var row = entities.PivotLayoutVariants.SingleOrDefault(p => p.ReportID == reportID && p.UserID == userID );
                row.VariantName = variantName;
                row.VariantDescription = variantDescription;
                entities.SaveChanges();

            }
            else
            {
                var row = new DAL.PivotLayoutVariants();
                row.ReportID = reportID;
                row.UserID = userID;
                row.VariantName = variantName;
                row.VariantDescription = variantDescription;
                entities.PivotLayoutVariants.Add(row);
                entities.SaveChanges();
            }
        }
        public static void SavePivotLayout(int reportID, int userID, int variantID, string pivotLayout)
        {
            var entities = new DAL.SAPExlEntities();
            if (!entities.PivotLayouts.Where(p => p.ReportID == reportID && p.UserID == userID && p.VariantID == variantID).Count().Equals(0))
            {
                var row = entities.PivotLayouts.Where(p => p.ReportID == reportID && p.UserID == userID && p.VariantID == variantID).SingleOrDefault();
                row.PivotLayout = pivotLayout;
                entities.SaveChanges();
              
            }
            else
            {
                var row = new DAL.PivotLayouts();
                row.ReportID = reportID;
                row.UserID = userID;
                row.VariantID = variantID;
                row.PivotLayout = pivotLayout;
                entities.PivotLayouts.Add(row);
                entities.SaveChanges();
            }
        }
        public static void InsertPivotTableDef(int userID, int reportID, int variantID, string sheetName, string tableName, int seqNum, string pivotELement, string sourceName)
        {
            SQLServerHandler.IDBBuilder bldr = SQLDataHandler.Functions.GetBuilder();
            string commandString = bldr.GetInsertStatement("PivotTableDef", new string[] { "UserID", "ReportID", "VariantID", "SheetName", "TableName", "SeqNum", "PivotElement", "SourceName" }, new object[] { userID, reportID, variantID, sheetName, tableName, seqNum, pivotELement, sourceName });
            SQLDataHandler.Functions.ExecuteSqlCommand(commandString);
        }
        public static void UpdatePivotTableDef(string columnName,object value)
        {
            SQLServerHandler.IDBBuilder bldr = SQLDataHandler.Functions.GetBuilder();
            string commandString = bldr.GetUpdateStatement("PivotTableDef", new string[] { columnName }, new object[] { value }, new string[] { "ID" }, new object[] { GetLastPivotTableDef() });
            SQLDataHandler.Functions.ExecuteSqlCommand(commandString);
        }
        public static int GetLastPivotTableDef()
        {
            string commandString = "SELECT TOP 1 ID FROM PivotTableDef ORDER BY ID DESC";
            return SQLDataHandler.Functions.GetIntData(commandString);
        }
        public static DataSet GetPivotTableDef(int UserID, int reportID,int variantID)
        {
            SQLServerHandler.IDBBuilder bldr = SQLDataHandler.Functions.GetBuilder();
            string commandString = bldr.GetSelectStatement("PivotTableDef",new string[] { "*" }, new string[] { "UserID","ReportID","VariantID" }, new object[] { UserID,reportID,variantID });
            return SQLDataHandler.Functions.GetData("PivotTableDef", commandString);
        }

        public static DataTable GetCurrentUserReportPivotLayoutVariant(int UserID, int reportID)
        {
            SQLServerHandler.IDBBuilder bldr = SQLDataHandler.Functions.GetBuilder();
            string commandString = bldr.GetSelectStatement("CurrentUserPivotLayoutVariants", new string[] { "*" }, new string[] { "UserID", "ReportID" }, new object[] { UserID, reportID });
            return SQLDataHandler.Functions.GetData("CurrentUserPivotLayoutVariants", commandString).Tables[0];
        }

        public static void UpdateCurrentUserReportPivotLayoutVariant(int UserID, int reportID, int pivotLayoutVariantID)
        {
            SQLServerHandler.IDBBuilder bldr = SQLDataHandler.Functions.GetBuilder();
            string commandString = bldr.GetUpdateStatement("CurrentUserPivotLayoutVariants", new string[] { "PivotLayoutVariantID" }, new object[] { pivotLayoutVariantID }, new string[] { "UserID", "ReportID" }, new object[] { UserID, reportID });
            SQLDataHandler.Functions.ExecuteSqlCommand(commandString);
        }

        public static DataSet GetMaterialUsage()
        {
            SQLServerHandler.IDBBuilder bldr = SQLDataHandler.Functions.GetBuilder();
            string commandString = bldr.GetSelectStatement("vwMaterialUsage", new string[] { "*" }, null, null);
            return SQLDataHandler.Functions.GetData("MaterialUsage", commandString);
        }
        public static void InsertPivotLayout(int variantID, string layout)
        {
            SQLServerHandler.IDBBuilder bldr = SQLDataHandler.Functions.GetBuilder();
            string commandString = bldr.GetInsertStatement("PivotLayouts", new string[] { "VariantID", "PivotLayout" }, new object[] { variantID, layout });
            SQLDataHandler.Functions.ExecuteSqlCommand(commandString);
        }
        public static string GetPivotLayout(int variantID)
        {
            string commandString = "SELECT PivotLayout FROM PivotLayouts WHERE VariantID = " + variantID;
            return SQLDataHandler.Functions.GetStringData(commandString);
        }
        public static void DeletePivotLayout(int variantID)
        {
            string commandString = "DELETE FROM PivotLayoutVariants WHERE ID = " + variantID;
            SQLDataHandler.Functions.ExecuteSqlCommand(commandString);

            string commandString1 = "DELETE FROM PivotLayouts WHERE VariantID = " + variantID;
            SQLDataHandler.Functions.ExecuteSqlCommand(commandString1);
        }
    }
}
