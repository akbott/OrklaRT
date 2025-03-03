﻿//------------------------------------------------------------------------------
// <auto-generated>
//    This code was generated from a template.
//
//    Manual changes to this file may cause unexpected behavior in your application.
//    Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace DAL
{
    using System;
    using System.Data.Entity;
    using System.Data.Entity.Infrastructure;
    
    public partial class SAPExlEntities : DbContext
    {
        public SAPExlEntities()
            : base("name=SAPExlEntities")
        {
        }
    
        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            throw new UnintentionalCodeFirstException();
        }
    
        public DbSet<BDCInput> BDCInput { get; set; }
        public DbSet<ExchangeRates> ExchangeRates { get; set; }
        public DbSet<ReportGroups> ReportGroups { get; set; }
        public DbSet<ReportOptions> ReportOptions { get; set; }
        public DbSet<ReportSheetOptions> ReportSheetOptions { get; set; }
        public DbSet<RighClickMenuFields> RighClickMenuFields { get; set; }
        public DbSet<RightClickSubMenu> RightClickSubMenu { get; set; }
        public DbSet<Views> Views { get; set; }
        public DbSet<vwCompanyCodes> vwCompanyCodes { get; set; }
        public DbSet<vwCurrency> vwCurrency { get; set; }
        public DbSet<vwCustomerHierarchy> vwCustomerHierarchy { get; set; }
        public DbSet<vwDayType> vwDayType { get; set; }
        public DbSet<vwMaterialPrice> vwMaterialPrice { get; set; }
        public DbSet<vwMaterialTypes> vwMaterialTypes { get; set; }
        public DbSet<vwPlants> vwPlants { get; set; }
        public DbSet<vwProductHierarchy> vwProductHierarchy { get; set; }
        public DbSet<vwQuantityUnit> vwQuantityUnit { get; set; }
        public DbSet<vwRelativePeriods> vwRelativePeriods { get; set; }
        public DbSet<vwReportGroups> vwReportGroups { get; set; }
        public DbSet<vwSalesValue> vwSalesValue { get; set; }
        public DbSet<PivotLayouts> PivotLayouts { get; set; }
        public DbSet<PivotLayoutVariants> PivotLayoutVariants { get; set; }
        public DbSet<vwSalesOrganizations> vwSalesOrganizations { get; set; }
        public DbSet<ReportComments> ReportComments { get; set; }
        public DbSet<UserReportMultipleSelections> UserReportMultipleSelections { get; set; }
        public DbSet<UserReportVariants> UserReportVariants { get; set; }
        public DbSet<CurrentUserReportVariants> CurrentUserReportVariants { get; set; }
        public DbSet<ReportSelectionDefaultValues> ReportSelectionDefaultValues { get; set; }
        public DbSet<vwMaterialGroups> vwMaterialGroups { get; set; }
        public DbSet<vwProfitCenters> vwProfitCenters { get; set; }
        public DbSet<Reports> Reports { get; set; }
        public DbSet<CurrentUserReportMultipleSelections> CurrentUserReportMultipleSelections { get; set; }
        public DbSet<CurrentUserReportSelections> CurrentUserReportSelections { get; set; }
        public DbSet<UserReportSelections> UserReportSelections { get; set; }
        public DbSet<ReportStatistics> ReportStatistics { get; set; }
        public DbSet<vwProductionScheduler> vwProductionScheduler { get; set; }
        public DbSet<vwBrands> vwBrands { get; set; }
        public DbSet<vwLogicalSystems> vwLogicalSystems { get; set; }
        public DbSet<vwSAPSystems> vwSAPSystems { get; set; }
        public DbSet<UserSAPSystems> UserSAPSystems { get; set; }
        public DbSet<vwUserSAPSystems> vwUserSAPSystems { get; set; }
        public DbSet<vwWorkCenterGroups> vwWorkCenterGroups { get; set; }
        public DbSet<ReportsLinkedQuery> ReportsLinkedQuery { get; set; }
        public DbSet<vwBudgetVersions> vwBudgetVersions { get; set; }
        public DbSet<CurrentUsers> CurrentUsers { get; set; }
        public DbSet<vwCurrentUser> vwCurrentUser { get; set; }
        public DbSet<UserGroups> UserGroups { get; set; }
        public DbSet<vwValuationClass> vwValuationClass { get; set; }
        public DbSet<vwMRPControllers> vwMRPControllers { get; set; }
        public DbSet<vwPurchasingGroups> vwPurchasingGroups { get; set; }
        public DbSet<vwShowOptions> vwShowOptions { get; set; }
        public DbSet<vwShelfLifeTypes> vwShelfLifeTypes { get; set; }
        public DbSet<vwWarehouseNumbers> vwWarehouseNumbers { get; set; }
        public DbSet<vwMaterialsIncluded> vwMaterialsIncluded { get; set; }
        public DbSet<vwStorageLocations> vwStorageLocations { get; set; }
        public DbSet<RightClickMenu> RightClickMenu { get; set; }
        public DbSet<ReportSelections> ReportSelections { get; set; }
        public DbSet<RfcConnection> RfcConnection { get; set; }
        public DbSet<vwDistributionChannel> vwDistributionChannel { get; set; }
        public DbSet<CurrentUserReportFields> CurrentUserReportFields { get; set; }
        public DbSet<vwShowStocks> vwShowStocks { get; set; }
        public DbSet<vwShowMD04Data> vwShowMD04Data { get; set; }
        public DbSet<PPLockedOrders> PPLockedOrders { get; set; }
        public DbSet<ProductionPlanData> ProductionPlanData { get; set; }
    }
}
