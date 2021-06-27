Imports Microsoft.Office.Tools.Ribbon
Imports System.Net.Sockets

Public Class SAPLogin
    Public Sub New(ByVal SapSystem As String, Optional loginFailed As Boolean = False)
        InitializeComponent()
        cboSAPSystems.Visible = True
        txtPassword.PasswordChar = "*"
        Using entities = New DAL.SAPExlEntities()
            cboSAPSystems.DataSource = entities.vwUserSAPSystems.ToList()
            cboSAPSystems.ValueMember = "SAPSystem"
            cboSAPSystems.DisplayMember = "SAPSystem"
            cboSAPSystems.Invalidate()
            If Not String.IsNullOrWhiteSpace(SapSystem) Then
                cboSAPSystems.SelectedValue = SapSystem
                Dim currentSapSystem = entities.vwUserSAPSystems.Where(Function(ss) ss.SAPSystem = SapSystem).SingleOrDefault()
                txtUserName.Text = currentSapSystem.SAPUserName
                txtPassword.Text = currentSapSystem.SAPPassword
            End If
        End Using
        If loginFailed Then lblMessage.Text = "Navn eller passord er ikke korrekt"
    End Sub

    Private Sub cboSAPSystems_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cboSAPSystems.SelectionChangeCommitted
        Using entities = New DAL.SAPExlEntities()
            If Not String.IsNullOrWhiteSpace(cboSAPSystems.SelectedValue.ToString()) Then
                Dim sapSystem = cboSAPSystems.SelectedValue.ToString()
                Dim currentSapSystem = entities.vwUserSAPSystems.Where(Function(ss) ss.SAPSystem = sapSystem).SingleOrDefault()
                txtUserName.Text = currentSapSystem.SAPUserName
                txtPassword.Text = currentSapSystem.SAPPassword
            End If
        End Using
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Try
            LoginSAP()
        Catch
        End Try
    End Sub

    Public Function CheckServerAvailablity() As Boolean
        Try
            Dim TcpClient As New TcpClient()
            TcpClient.Connect("10.195.9.174", 1433)
            TcpClient.Close()
            Return True
        Catch
            Return False
        End Try
    End Function

    Private Sub txtPassword_KeyDown(sender As Object, e As KeyEventArgs) Handles txtPassword.KeyDown
        If e.KeyCode = Keys.Enter Then
            LoginSAP()
        End If
    End Sub
    Private Sub LoginSAP()
        lblMessage.Text = String.Empty
        Using entities = New DAL.SAPExlEntities()
            Dim sapSystem = cboSAPSystems.SelectedValue.ToString()
            Dim userId = entities.vwCurrentUser.SingleOrDefault().ID
            Dim row = entities.UserSAPSystems.SingleOrDefault(Function(ss) ss.SAPSystem = sapSystem And ss.UserId = userId)
            row.SAPUserName = txtUserName.Text
            row.SAPPassword = txtPassword.Text
            entities.SaveChanges()
            Globals.Ribbons.OrklaRT.edtSAPSystem.Text = sapSystem
        End Using
        Me.Close()
    End Sub

End Class