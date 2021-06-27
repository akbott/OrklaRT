Imports System.IO

Module GlobalDeclarations
    'This module contents declarations of global constants and variables
    Public Application As Excel.Application = Globals.ThisAddIn.Application
    Public gsErrTest As String
    Public Const gSysTitle As String = "OrklaRT"
    Public gwbReport As Excel.Workbook
    Public sql As String
    Public gReportID As Integer
    Public ProductionPlanTable As System.Data.DataTable
    Public MD04Table As System.Data.DataTable       
    Public gUserId As Integer
    '= OrklaRTBPL.CommonFacade.GetUserID()
    Public systemIni As String = Path.GetDirectoryName(New Uri(System.Reflection.Assembly.GetExecutingAssembly().CodeBase).LocalPath.ToString()) + "\Files\SystemIni_07_M.xls"
End Module
