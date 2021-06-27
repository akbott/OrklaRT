Imports System.Windows.Forms

Public Class Selection
    Public Sub New(reportID As Integer, userID As Integer, fromRightClick As Boolean)
        InitializeComponent()
        Dim selectionPane As New SelectionPane.Selection(reportID, userID, fromRightClick)
        selectionElementHost.Child = selectionPane
        selectionElementHost.Dock = DockStyle.Fill
    End Sub
End Class
