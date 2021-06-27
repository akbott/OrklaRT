using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace OrklaRTBPL
{
    public class ReportSpecific
    {
        public static DataSet ClonePlanData()
        {
            string commandString = "SELECT * FROM [dbo].[ProductionPlanData] ";
            return SQLDataHandler.Functions.GetData("ProductionPlanData", commandString);
        }
        public static DataSet CheckPlanExists(string plant, string planDate)
        {
            string commandString = "SELECT [Order],[ConfirmedOrderFinishDate1],[ActualFinishDate],[ScheduledFinishDate],[ScheduledFinishTime],[ActualStartDate],[ScheduledStartDate],[ScheduledStartTime],[ActualFinishExecutionDate]," +
                                  "[ActualStartExecutionDate],[DeliveryCompletedIndicator],[OrderNumber],[EarliestScheduledStartDate],[Language],[MixMaterialDescription],[LogicalSystem],[PlantDescription],[DeletionFlag]," +
                                  "[OrderPriority],[ItemCategory],[WorkCenterDescription],[Counter],[ResourceObjectID],[Temp],[EarliestScheduledStartTime],[ConfirmedOrderFinishDate],[ConfirmedOrderFinishTime],[ActualStartTime]," +
                                  "[ActualStartExecutionTime],[ActualFinishTime],[CapacityID],[MaterialDescription],[IndicatorPhase],[LatestScheduledStartDate],[LatestScheduledStartTime],[ControlKey],[Temp1],[ShortDescription],[MixUnit],[LanguageCode]," +
                                  "[MixMaterial],[Plant],[WorkCenter],[MaterialNumber],[BaseQuantity],[RequiredMixingQuantity],[ConfirmedScrap],[ConfirmedYield],[ConfirmedYield1],[QuantityOfGoodsReceived],[ProcessingTime],[ScrapQuantity]," +
                                  "[TotalOrderQuantity],[TotalScrapQuantity],[SetupTime],[OperationQuantity],[ConfirmedQuantityYield],[TotalStock],[QuantityWithdrawn],[SetupTid],[ProcTid]," +
                                  "[CMaterialName],[CRemainingQuantity],[CRtTotal],[CSetup],[CRemainingTime],[CSequence],[CStartDateH]" +
                                  ",[CStartTimeH],[CActive],[CEndDate],[CEndTime],[CStart],[CStop],[CStartSec],[CWorkCenter],[CStopSec],[CAccCap],[CStartCap],[CAccCap1],[CStartValue],[CProdValue],[CNowValue],[CTotalValue]" +
                                  ",[CAktHelp],[CAccAkt],[CAccStarted],[CLocked],[CCapToday],[CLeftNow],[CCapStart],[CMixStatus],[CHelp2],[CHelp3],[CStartDate],[CStartTime] FROM [dbo].[ProductionPlanData] WHERE Plant = '" + plant + "' AND CONVERT(date,PlanTime,103) = CONVERT(date,'" + planDate + "',103)";
                                 //" [OrderNumber],[Description],[Plant],[ActualStartDate],[TotalOrderQuantity],[BaseUnitOfMeasure],[TotalScrapQuantity],[BaseUnitOfMeasure1],[MaterialNumber],[ActualStartTime]" +
                                 //",[IndicatorPhase],[MaterialDescription],[WorkCenter],[OrderPriority],[ControlKey],[UnitOfMeasure],[BaseQuantity],[UnitOfMeasure1],[OperationQuantity],[UnitOfMeasure2],[ScrapQuantity]" +
                                 //",[UnitOfMeasure3],[ConfirmedYield],[UnitOfMeasure4],[ConfirmedScrap],[UnitOfMeasure5],[ConfirmedYield1],[UnitOfMeasure6],[Counter],[ActualStartExecutionDate],[ActualStartExecutionTime]" +
                                 //",[ConfirmedOrderFinishDate],[ConfirmedOrderFinishTime],[SetupTime],[SetupUnit],[ProcessingTime],[UnitOfWork7],[EarliestScheduledStartDate],[EarliestScheduledStartTime],[LatestScheduledStartDate]" +
                                 //",[LatestScheduledStartTime],[ScheduledStartDate],[ScheduledFinishDate],[ScheduledFinishTime],[ScheduledStartTime],[ResourceObjectID],[QuantityOfGoodsReceived],[UnitOfMeasureInHouse]" +
                                 //",[ShortDescription],[ActualOrderFinishDate],[ConfirmedQuantityYield],[BaseUnitOfMeasure2],[ActualFinishTime],[ActualFinishDate],[DeliveryCompletedIndicator],[MaterialNumberMix],[RequiredMixingQuantity]" +
                                 //",[BaseUnitOfMeasure3],[QuantityWithdrawn],[BaseUnitOfMeasure4],[TotalStock],[BaseUnitOfMeasure5],                                  
            return SQLDataHandler.Functions.GetData("ProductionPlanData", commandString);
        }
        public static DataSet GetProductionPlanData(string plant, string planDate)
        {
            string commandString = "SELECT [Order],[ConfirmedOrderFinishDate1],[ActualFinishDate],[ScheduledFinishDate],[ScheduledFinishTime],[ActualStartDate],[ScheduledStartDate],[ScheduledStartTime],[ActualFinishExecutionDate]," +
                                  "[ActualStartExecutionDate],[DeliveryCompletedIndicator],[OrderNumber],[EarliestScheduledStartDate],[Language],[MixMaterialDescription],[LogicalSystem],[PlantDescription],[DeletionFlag]," +
                                  "[OrderPriority],[ItemCategory],[WorkCenterDescription],[Counter],[ResourceObjectID],[Temp],[EarliestScheduledStartTime],[ConfirmedOrderFinishDate],[ConfirmedOrderFinishTime],[ActualStartTime]," +
                                  "[ActualStartExecutionTime],[ActualFinishTime],[CapacityID],[MaterialDescription],[IndicatorPhase],[LatestScheduledStartDate],[LatestScheduledStartTime],[ControlKey],[Temp1],[ShortDescription],[MixUnit],[LanguageCode]," +
                                  "[MixMaterial],[Plant],[WorkCenter],[MaterialNumber],[BaseQuantity],[RequiredMixingQuantity],[ConfirmedScrap],[ConfirmedYield],[ConfirmedYield1],[QuantityOfGoodsReceived],[ProcessingTime],[ScrapQuantity]," +
                                  "[TotalOrderQuantity],[TotalScrapQuantity],[SetupTime],[OperationQuantity],[ConfirmedQuantityYield],[TotalStock],[QuantityWithdrawn],[SetupTid],[ProcTid] WHERE Plant = '" + plant + "' AND CONVERT(date,PlanTime,103) = CONVERT(date,'" + planDate + "',103)";            
            return SQLDataHandler.Functions.GetData("ProductionPlanData", commandString);
        }
        public static int GetPlannerId(string plant, string planDate)
        {
            string commandString = "SELECT DISTINCT UserId FROM [dbo].[ProductionPlanData] WHERE Plant = '" + plant + "' AND CONVERT(date,PlanTime,103) = CONVERT(date,'" + planDate + "',103)";
            return SQLDataHandler.Functions.GetIntData(commandString);
        }
        public static void UpdateProductionPlanPriorities(string plant,string workCenter,int colorIndex,string updatedOn, int userID)
        {
            //string commandString1 = "DELETE FROM PPPriorities WHERE Plant = '" + plant + "' AND CONVERT(date,UpdatedOn,103) = CONVERT(date,'" + updatedOn + "',103)";
            //SQLDataHandler.Functions.ExecuteNonQuery(commandString1);
            string commandString = "UPDATE PPPriorities SET ColorIndex=" + colorIndex + ",UpdatedOn=CONVERT(datetime,'" + updatedOn + "',103) FROM PPPriorities WHERE WorkCenter='" + workCenter + "'";
            SQLDataHandler.Functions.ExecuteNonQuery(commandString);
        }
        public static void DeleteProductionPlanData(string plant, string planTime)
        {
            string commandString = "DELETE FROM ProductionPlanData WHERE Plant = '" + plant + "' AND CONVERT(date,PlanTime,103) = CONVERT(date,'" + planTime + "',103)";
            SQLDataHandler.Functions.ExecuteNonQuery(commandString);
        }
        public static void DeleteLockedOrder(string plant, string orderNumber, int userID)
        {
            string commandString = "DELETE FROM PPLockedOrders WHERE Plant = '" + plant + "' AND OrderNumber = '" + orderNumber + "'";
            SQLDataHandler.Functions.ExecuteNonQuery(commandString);
        }
        public static void InsertLockedOrder(string plant, string orderNumber, int userID)
        {
            string commandString = "INSERT INTO PPLockedOrders(Plant,OrderNumber,LockedDate,UserId) VALUES('" + plant + "','" + orderNumber + "',GETDATE()," + userID + ")";
            SQLDataHandler.Functions.ExecuteNonQuery(commandString);
        }
        public static DataSet GetPriorities(string plant, string planDate)
        {
            string commandString = "SELECT WorkCenter,ColorIndex,UserId FROM [dbo].[PPPriorities] WHERE Plant = '" + plant + "'";
            return SQLDataHandler.Functions.GetData("PPPriorities", commandString);
        }
        public static DateTime GetPlanTime(string plant, string planDate)
        {
            string commandString = "SELECT DISTINCT PlanTime FROM ProductionPlanData WHERE Plant = '" + plant + "' AND CONVERT(date,PlanTime,103) = CONVERT(date,'" + planDate + "',103)";
            return (DateTime)SQLDataHandler.Functions.GetObjectData(commandString);
        }

        public static void InsertMPOrderStart(string orderNumber,string newDate, string plant, int userID)
        {            
            string commandString = "SELECT COUNT(*) FROM MPOrderStart WHERE OrderNumber ='" + orderNumber + "' AND Plant ='" + plant + "'";
            string commandString1 = String.Empty;
            if (SQLDataHandler.Functions.GetIntData(commandString).Equals(1))
            {
                commandString1 = "UPDATE MPOrderStart SET NewDate = CONVERT(datetime,'" + newDate + "',103),UserId = " + userID + "  FROM MPOrderStart WHERE OrderNumber = '" + orderNumber + "' AND Plant ='" + plant + "'";
            }
            else
            {
                commandString1 = "INSERT INTO MPOrderStart(OrderNumber,NewDate,Plant,UserId) VALUES('" + orderNumber + "'," + "CONVERT(datetime,'" + newDate + "',103)" + ",'" + plant + "'," + userID + ")";
            }
            SQLDataHandler.Functions.ExecuteNonQuery(commandString1);           
        }
        public static void DeleteMPOrderStart(string orderNumber, string newDate, string plant, int userID)
        {
            string commandString = "DELETE FROM MPOrderStart WHERE OrderNumber = '" + orderNumber + "' AND NewDate = CONVERT(datetime,'" + newDate + "',103) AND Plant = '" + plant + "' AND UserId = " + userID;
            SQLDataHandler.Functions.ExecuteNonQuery(commandString);
        }
        public static DataSet GetMPOrderStart(string plant)
        {
            string commandString = "SELECT OrderNumber,NewDate FROM MPOrderStart WHERE Plant = '" + plant + "'";
            return SQLDataHandler.Functions.GetData("MPOrderStart", commandString);
        }

        public static void InsertMPMixPlan(string orderNumber, int machine, string plant, int userID)
        {
            string commandString = "SELECT COUNT(*) FROM MPMixPlan WHERE OrderNumber ='" + orderNumber + "' AND Plant ='" + plant + "'";
            string commandString1 = String.Empty;
            if (SQLDataHandler.Functions.GetIntData(commandString).Equals(1))
            {
                commandString1 = "UPDATE MPMixPlan SET Machine = " + machine + " FROM MPMixPlan WHERE OrderNumber = '" + orderNumber + "' AND Plant = '" + plant + "'";
            }
            else
            {
                commandString1 = "INSERT INTO MPMixPlan(OrderNumber,Machine,Plant,UserId) VALUES('" + orderNumber + "'," + machine + ",'" + plant + "'," + userID + ")";
            }
            SQLDataHandler.Functions.ExecuteNonQuery(commandString1);            
        }
        public static DataSet GetMPMixPlan(string plant)
        {
            string commandString = "SELECT OrderNumber,Machine FROM MPMixPlan WHERE Plant = '" + plant + "'";
            return SQLDataHandler.Functions.GetData("MPMixPlan", commandString);
        }

        public static void InsertMPRSTest(string orderNumber, string rs, string plant, int userID)
        {
            string commandString = "INSERT INTO MPRSTest(OrderNumber,RS,Plant,UserId) VALUES('" + orderNumber + "','" + rs + "','" + plant + "'," + userID + ")";
            SQLDataHandler.Functions.ExecuteNonQuery(commandString);
        }
        public static DataSet GetMPRSTest(string plant)
        {
            string commandString = "SELECT OrderNumber,RS FROM MPRSTest WHERE Plant = '" + plant + "'";
            return SQLDataHandler.Functions.GetData("MPRSTest", commandString);
        }

        public static void InsertMPPriPlan(string orderNumber, int priority, string plant, int userID)
        {
            string commandString = "SELECT COUNT(*) FROM MPPriPlan WHERE OrderNumber ='" + orderNumber + "' AND Plant ='" + plant + "'";
            string commandString1 = String.Empty;
            if(SQLDataHandler.Functions.GetIntData(commandString).Equals(1))
            {
                commandString1 = "UPDATE MPPriPlan SET Priority = " + priority + " FROM MPPriPlan WHERE OrderNumber = '" + orderNumber + "'";
            }
            else
            {
                commandString1 = "INSERT INTO MPPriPlan(OrderNumber,Priority,Plant,UserId) VALUES('" + orderNumber + "'," + priority + ",'" + plant + "'," + userID + ")";
            }           
            SQLDataHandler.Functions.ExecuteNonQuery(commandString1);
        }
        public static DataSet GetMPPriPlan(string plant)
        {
            string commandString = "SELECT OrderNumber,Priority FROM MPPriPlan WHERE Plant = '" + plant + "'";
            return SQLDataHandler.Functions.GetData("MPPriPlan", commandString);
        }

        public static void InsertMPMixWC(string workCenter, string plant, int userID)
        {
            string commandString = "INSERT INTO MPMixWC(WorkCenter,Plant,UserId) VALUES('" + workCenter + "','" + plant + "'," + userID + ")";
            SQLDataHandler.Functions.ExecuteNonQuery(commandString);
        }

        public static void DeleteMPMixWC(string plant)
        {
            string commandString = "DELETE FROM MPMixWC WHERE Plant = '" + plant + "'";
            SQLDataHandler.Functions.ExecuteNonQuery(commandString);
         }
        public static DataSet GetMPMixWC(string plant)
        {
            string commandString = "SELECT WorkCenter FROM MPMixWC WHERE Plant = '" + plant + "'";
            return SQLDataHandler.Functions.GetData("MPMixWC", commandString);
        }

        public static DataSet GetProdPlanData(string plant, string planDate)
        {
            string commandString = "SELECT MaterialNumber,OrderNUmber,CStart,MaterialDescription,WorkCenter,OperationQuantity FROM ProductionPlanData WHERE Plant = '" + plant + "' AND CONVERT(date,PlanTime,103) = CONVERT(date,'" + planDate + "',103)";
            return SQLDataHandler.Functions.GetData("ProductionPlanData", commandString);
        }

        public static DataSet GetMPProdStatus(string plant)
        {
            string commandString = "SELECT OrderNumber,Material,Material1,StartValue FROM MPProdStatus WHERE Plant = '" + plant + "'";
            return SQLDataHandler.Functions.GetData("MPProdStatus", commandString);
        }

        public static DataSet GetSTTransportGroups(int warehouseNumber)
        {
            string commandString = "SELECT IndexValue,Department,FromWarehouse,ToWarehouse,TestBin,Time,GroupName FROM STTransportGroups WHERE WarehouseNumber = " + warehouseNumber;
            return SQLDataHandler.Functions.GetData("STTransportGroups", commandString);
        }

        public static DataSet GetSTBinTest(int warehouseNumber)
        {
            string commandString = "SELECT TestBin,Department FROM STBinTest WHERE WarehouseNumber = " + warehouseNumber;
            return SQLDataHandler.Functions.GetData("STBinTest", commandString);
        }

        public static DataSet GetSTExcludedTypes(int warehouseNumber)
        {
            string commandString = "SELECT ExcludedTypes FROM STExcludedTypes WHERE WarehouseNumber = " + warehouseNumber;
            return SQLDataHandler.Functions.GetData("STExcludedTypes", commandString);
        }

        public static void InsertSTTransportGroups(int warehouseNumber, string indexValue, string department, int from, int to, string testBin,int time,string group)
        {
            string commandString = "INSERT INTO STTransportGroups(WarehouseNumber,IndexValue,Department,FromWarehouse,ToWarehouse,TestBin,Time,GroupName) VALUES(" + warehouseNumber + ",'" + indexValue + "','" + department + "'," + from + "," + to + ",'" + testBin + "'," + time + ",'" + group + "')";
            SQLDataHandler.Functions.ExecuteNonQuery(commandString);
        }

        public static void InsertSTBinTest(int warehouseNumber, string testBin, string department)
        {
            string commandString = "INSERT INTO STBinTest(WarehouseNumber,TestBin,Department) VALUES(" + warehouseNumber + ",'" + testBin + "','" + department + "')";
            SQLDataHandler.Functions.ExecuteNonQuery(commandString);
        }

        public static void InsertSTExcludedTypes(int warehouseNumber, string excludedTypes)
        {
            string commandString = "INSERT INTO STExcludedTypes(WarehouseNumber,ExcludedTypes) VALUES(" + warehouseNumber + ",'" + excludedTypes + "')";
            SQLDataHandler.Functions.ExecuteNonQuery(commandString);
        }

        public static DataSet GetT157EData(string languageCode)
        {
            string commandString = "SELECT BWART,GRUND,GRTXT FROM T157E WHERE LanguageCode = '" + languageCode + "'";
            return SQLDataHandler.Functions.GetData("T157E", commandString);
        }

        public static DataSet GetStockValuesAndCoverageProdPlanData(string plant, string planDate)
        {
            string commandString = "SELECT MaterialNumber,CStart,CWorkCenter FROM ProductionPlanData WHERE Plant LIKE '%" + plant + "%' AND CONVERT(date,PlanTime,103) = CONVERT(date,'" + planDate + "',103)";
            return SQLDataHandler.Functions.GetData("ProductionPlanData", commandString);
        }
        public static DataSet GetDPKData()
        {
            string commandString = "SELECT MARC.WERKS+MARM.MATNR,MARM.MEINH,MARM.UMREZ,MARM.UMREN FROM MARM INNER JOIN MARC ON MARC.MATNR = MARM.MATNR WHERE MARM.MEINH IN ('DPK') AND MARC.WERKS = '" + OrklaRTBPL.SelectionFacade.DailyProductionPlanPlant + "'";
            return SQLDataHandler.Functions.GetData("MARMData", commandString);
        }
        public static DataSet GetFPKData()
        {
            string commandString = "SELECT MARC.WERKS+MARM.MATNR,MARM.MEINH,MARM.UMREZ,MARM.UMREN FROM MARM INNER JOIN MARC ON MARC.MATNR = MARM.MATNR WHERE MARM.MEINH IN ('FPK') AND MARC.WERKS = '" + OrklaRTBPL.SelectionFacade.DailyProductionPlanPlant + "'";
            return SQLDataHandler.Functions.GetData("MARMData", commandString);
        }
        public static DataSet GetTPKData()
        {
            string commandString = "SELECT MARC.WERKS+MARM.MATNR,MARM.MEINH,MARM.UMREZ,MARM.UMREN FROM MARM INNER JOIN MARC ON MARC.MATNR = MARM.MATNR WHERE MARM.MEINH IN ('TPK') AND MARC.WERKS = '" + OrklaRTBPL.SelectionFacade.ShelfLifeSelectionPlant + "'";
            return SQLDataHandler.Functions.GetData("MARMData", commandString);
        }
        public static DataSet GetAllergenData()
        {
            string commandString = "SELECT Material,Allergen FROM Allergen";
            return SQLDataHandler.Functions.GetData("Allergenata", commandString);
        }
        public static DataSet GetAllergenType(bool mixing)
        {
            string commandString = String.Empty;
            if (mixing.Equals(true))
            {
                commandString = "SELECT MARA.[MATNR],[ZZMUST],[ZZSESA],[ZZFISK],[ZZSKAL],[ZZRES2],[ZZNUTS],[ZZPEAN] FROM MARA INNER JOIN MARC ON MARC.MATNR = MARA.MATNR WHERE MARC.WERKS = '" + OrklaRTBPL.SelectionFacade.MixingPlanSelectionPlant + "'";
            }
            else
            {
                commandString = "SELECT MARA.[MATNR],[ZZMUST],[ZZSESA],[ZZFISK],[ZZSKAL],[ZZRES2],[ZZNUTS],[ZZPEAN] FROM MARA INNER JOIN MARC ON MARC.MATNR = MARA.MATNR WHERE MARC.WERKS = '" + OrklaRTBPL.SelectionFacade.ProductionPlanSelectionPlant + "'";
            }
            return SQLDataHandler.Functions.GetData("AllergenTypeData", commandString);
        }
        public static DataSet GetWorkCentersCapacityData()
        {
            string commandString = "SELECT * FROM vwWorkCentersCapacity";
            return SQLDataHandler.Functions.GetData("WorkCentersCapacity", commandString);
        }
    }
}
