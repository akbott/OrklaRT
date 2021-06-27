Imports System.Runtime.InteropServices

<ComVisible(True)>
Public Interface ICOMCalls
    Sub LaunchSAPOrder(orderNumber As String)
    Sub LaunchMD04(materialNumber As String)
    Sub LoopStockInLevels()
    Sub LoopForecasts()
    Sub LoopAllLevels()
    Sub CreateForecast()    
    Sub CreatePlainForecast() 
    Sub CreateActualForecast()
    Sub ShowStockIn()
    Sub ShowForecast1()
    Sub ShowCurrentSim()
    Sub ShowActualFC()
    Sub ShowMRP()
    Sub ShowDailyUsage()
    Sub PurchasingCockpitGet_Requsitions()
    Sub PurchasingCockpitGet_OpenQty()
    Sub PurchasingCockpitGet_SafetyTime()
    Sub PurchasingCockpitGet_Requirements()
    Sub PurchasingCockpitGet_GRTime()
    Sub PurchasingCockpitGet_SafStock()
    Sub PurchasingCockpitSet_Formulas()
    Sub PurchasingCockpitSet_PlantRefresh()
End Interface
