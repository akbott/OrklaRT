Imports System.DirectoryServices
Imports System.DirectoryServices.ActiveDirectory


Partial Public Class UserSettings
    Public dirSearch As DirectorySearcher = Nothing
    Public userName As String = String.Empty

    Private Sub Settings_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            txtAddress.Text = GetSystemDomain()
            txtSearchUser.Text = Environment.UserName
            GetUserInformation(txtAddress.Text)
            If OrklaRTBPL.CommonFacade.GetCurrentUser().Rows.Count > 0 Then
                If OrklaRTBPL.CommonFacade.GetUserGroup(gUserId).Equals(2) Or OrklaRTBPL.CommonFacade.GetUserGroup(gUserId).Equals(3) Or OrklaRTBPL.CommonFacade.GetUserGroup(gUserId).Equals(4) Then
                    grpUserSettings.Enabled = True
                    txtSearchUser.Focus()
                    btnSearchUserName.[Select]()
                    txtSearchUser.Text = Environment.UserName
                    grpSapUserDetails.Enabled = True
                End If               
            End If
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.InnerException.InnerException.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "Settings", gUserId, gReportID)
        End Try
    End Sub

    Private Function GetSystemDomain() As String
        Try
            Return Environment.UserDomainName.ToUpper()
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "Settings", gUserId, gReportID)
            Return String.Empty
        End Try
    End Function

    Private Sub GetUserInformation(domain As String)
        Try

            Cursor.Current = Cursors.WaitCursor

            pnlBlock.BringToFront()
            pnlBlock.Visible = True

            Dim rs As SearchResult = Nothing
            If txtSearchUser.Text.Trim().IndexOf("@") > 0 Then
                rs = SearchUserByEmail(GetDirectorySearcher(domain), txtSearchUser.Text.Trim())
            Else
                rs = SearchUserByUserName(GetDirectorySearcher(domain), txtSearchUser.Text.Trim())
                If IsNothing(rs) Then rs = SearchUserByLastName(GetDirectorySearcher(domain), txtSearchUser.Text.Trim())
            End If

            If rs IsNot Nothing Then
                ShowUserInformation(rs)
                If Not txtSearchUser.Text.Equals(rs.GetDirectoryEntry().Properties("samaccountname").Value.ToString()) Then
                    txtSearchUser.Text = rs.GetDirectoryEntry().Properties("samaccountname").Value.ToString()
                End If

                Dim userTable = OrklaRTBPL.CommonFacade.GetUser(txtAddress.Text & "\" & txtSearchUser.Text)
                If userTable.Rows.Count > 0 Then
                    Select Case OrklaRTBPL.CommonFacade.GetUserGroup(CInt(userTable.Rows(0)("ID")))
                        Case 2
                            rbSysAdmin.Checked = True
                        Case 3
                            rbProdPlanner.Checked = True
                        Case 4
                            rbUserAdmin.Checked = True
                        Case Else
                            rbNormalUser.Checked = True
                    End Select
                    lblMessage.ForeColor = Drawing.Color.DarkGreen
                    lblMessage.Text = "OrklaRT bruker"
                    Using entities = New DAL.SAPExlEntities()
                        Dim currentUserSapSystem = entities.vwCurrentUser.SingleOrDefault().SAPSystem
                        Dim currentUserBWSystem = entities.vwCurrentUser.SingleOrDefault().BwHana
                        Dim userSapSystem = entities.vwUserSAPSystems.Where(Function(ss) ss.SAPSystem = currentUserSapSystem).SingleOrDefault()
                        txtSAPSystem.Text = currentUserSapSystem
                        txtSAPUsername.Text = userSapSystem.SAPUserName
                        txtSAPPassword.Text = userSapSystem.SAPPassword
                        Dim userBWSystem = entities.vwUserSAPSystems.Where(Function(ss) ss.SAPSystem = currentUserBWSystem).SingleOrDefault()
                        txtBWSystem.Text = currentUserBWSystem
                        txtBwUserName.Text = userBWSystem.SAPUserName
                        txtBWPassword.Text = userBWSystem.SAPPassword
                        btnChangeSystem.Text = String.Empty
                        If currentUserSapSystem.Equals("R3P") Then
                            btnChangeSystem.Text = "&Bytte til R3Q/BHQ"
                        Else
                            btnChangeSystem.Text = "&Bytte til R3P/BHP"
                        End If
                    End Using
                Else
                    lblMessage.ForeColor = Drawing.Color.DarkRed
                    lblMessage.Text = "Ikke OrklaRT bruker"
                    rbNormalUser.Checked = True
                End If
            Else
                MessageBox.Show("User not found!!!", "Search Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.InnerException.InnerException.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "Settings", gUserId, gReportID)
        End Try
    End Sub

    Private Sub ShowUserInformation(rs As SearchResult)
        Try
            Cursor.Current = Cursors.[Default]

            pnlBlock.Visible = False

            If rs.GetDirectoryEntry().Properties("samaccountname").Value IsNot Nothing Then
                lblUsernameDisplay.Text = "Username : " & rs.GetDirectoryEntry().Properties("samaccountname").Value.ToString()
                userName = rs.GetDirectoryEntry().Properties("samaccountname").Value.ToString()
            End If

            If rs.GetDirectoryEntry().Properties("givenName").Value IsNot Nothing Then
                lblFirstname.Text = "First Name : " & rs.GetDirectoryEntry().Properties("givenName").Value.ToString()
            Else
                lblFirstname.Text = "First Name : " & String.Empty
            End If

            If rs.GetDirectoryEntry().Properties("initials").Value IsNot Nothing Then
                lblMiddleName.Text = "Middle Name : " & rs.GetDirectoryEntry().Properties("initials").Value.ToString()
            Else
                lblMiddleName.Text = "Middle Name : " & String.Empty
            End If

            If rs.GetDirectoryEntry().Properties("sn").Value IsNot Nothing Then
                lblLastName.Text = "Last Name : " & rs.GetDirectoryEntry().Properties("sn").Value.ToString()
            Else
                lblLastName.Text = "Last Name : " & String.Empty
            End If

            If rs.GetDirectoryEntry().Properties("mail").Value IsNot Nothing Then
                lblEmailId.Text = "Email ID : " & rs.GetDirectoryEntry().Properties("mail").Value.ToString()
            Else
                lblEmailId.Text = "Email ID : " & String.Empty
            End If

            If rs.GetDirectoryEntry().Properties("title").Value IsNot Nothing Then
                lblTitle.Text = "Title : " & rs.GetDirectoryEntry().Properties("title").Value.ToString()
            Else
                lblTitle.Text = "Title : " & String.Empty
            End If

            If rs.GetDirectoryEntry().Properties("company").Value IsNot Nothing Then
                lblCompany.Text = "Company : " & rs.GetDirectoryEntry().Properties("company").Value.ToString()
            Else
                lblCompany.Text = "Company : " & String.Empty
            End If

            If rs.GetDirectoryEntry().Properties("l").Value IsNot Nothing Then
                lblCity.Text = "City : " & rs.GetDirectoryEntry().Properties("l").Value.ToString()
            Else
                lblCity.Text = "City : " & String.Empty
            End If

            If rs.GetDirectoryEntry().Properties("st").Value IsNot Nothing Then
                lblState.Text = "State : " & rs.GetDirectoryEntry().Properties("st").Value.ToString()
            Else
                lblState.Text = "State : " & String.Empty
            End If

            If rs.GetDirectoryEntry().Properties("co").Value IsNot Nothing Then
                lblCountry.Text = "Country : " & rs.GetDirectoryEntry().Properties("co").Value.ToString()
            Else
                lblCountry.Text = "Country : " & String.Empty
            End If

            If rs.GetDirectoryEntry().Properties("postalCode").Value IsNot Nothing Then
                lblPostal.Text = "Postal Code : " & rs.GetDirectoryEntry().Properties("postalCode").Value.ToString()
            Else
                lblPostal.Text = "Postal Code : " & String.Empty
            End If

            If rs.GetDirectoryEntry().Properties("telephoneNumber").Value IsNot Nothing Then
                lblTelephone.Text = "Telephone No. : " & rs.GetDirectoryEntry().Properties("telephoneNumber").Value.ToString()
            Else
                lblTelephone.Text = "Telephone No : " & String.Empty
            End If
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "Settings", gUserId, gReportID)
        End Try

    End Sub

    Private Function GetDirectorySearcher(domain As String) As DirectorySearcher
        If dirSearch Is Nothing Then
            Try
                dirSearch = New DirectorySearcher(New DirectoryEntry("LDAP://" & domain))
            Catch e As DirectoryServicesCOMException
                MessageBox.Show("Connection Creditial is Wrong!!!, please Check.", "Erro Info", MessageBoxButtons.OK, MessageBoxIcon.[Error])
                e.Message.ToString()
            End Try
            Return dirSearch
        Else
            Return dirSearch
        End If
    End Function

    Private Function SearchUserByUserName(ds As DirectorySearcher, username As String) As SearchResult
        ds.Filter = "(&((&(objectCategory=Person)(objectClass=User)))(samaccountname=" & username & "))"

        ds.SearchScope = SearchScope.Subtree
        ds.ServerTimeLimit = TimeSpan.FromSeconds(90)

        Dim userObject As SearchResult = ds.FindOne()

        If userObject IsNot Nothing Then
            Return userObject
        Else
            Return Nothing
        End If
    End Function

    Private Function SearchUserByEmail(ds As DirectorySearcher, email As String) As SearchResult
        ds.Filter = "(&((&(objectCategory=Person)(objectClass=User)))(mail=" & email & "))"

        ds.SearchScope = SearchScope.Subtree
        ds.ServerTimeLimit = TimeSpan.FromSeconds(90)

        Dim userObject As SearchResult = ds.FindOne()

        If userObject IsNot Nothing Then
            Return userObject
        Else
            Return Nothing
        End If
    End Function

    Private Function SearchUserByFirstName(ds As DirectorySearcher, firstName As String) As SearchResult
        ds.Filter = "(&((&(objectCategory=Person)(objectClass=User)))(givenName=" & firstName & "))"

        ds.SearchScope = SearchScope.Subtree
        ds.ServerTimeLimit = TimeSpan.FromSeconds(90)

        Dim userObject As SearchResult = ds.FindOne()

        If userObject IsNot Nothing Then
            Return userObject
        Else
            Return Nothing
        End If
    End Function

    Private Function SearchUserByLastName(ds As DirectorySearcher, lastName As String) As SearchResult
        ds.Filter = "(&((&(objectCategory=Person)(objectClass=User)))(sn=" & lastName & "))"

        ds.SearchScope = SearchScope.Subtree
        ds.ServerTimeLimit = TimeSpan.FromSeconds(90)

        Dim userObject As SearchResult = ds.FindOne()

        If userObject IsNot Nothing Then
            Return userObject
        Else
            Return Nothing
        End If
    End Function

    Private Sub btnSearchUserName_Click(sender As Object, e As EventArgs) Handles btnSearchUserName.Click
        Try
            If txtSearchUser.Text.Trim().Length <> 0 Then
                GetUserInformation(txtAddress.Text.Trim())
            Else
                MessageBox.Show("Please enter all required inputs.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "Settings", gUserId, gReportID)
        End Try
    End Sub

    Private Sub txtSearchUser_KeyDown(sender As Object, e As KeyEventArgs) Handles txtSearchUser.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then
                If txtSearchUser.Text.Trim().Length <> 0 Then
                    GetUserInformation(txtAddress.Text.Trim())
                Else
                    MessageBox.Show("Please enter all required inputs.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End If
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "Settings", gUserId, gReportID)
        End Try
    End Sub

    Private Sub btnAddToOrklaRT_Click(sender As Object, e As EventArgs) Handles btnAddToOrklaRT.Click
        Try
            Dim sqlServer As New Microsoft.SqlServer.Management.Smo.Server("RIHQITDEV1")
            Dim dataBase = sqlServer.Databases("OrklaRT")
            Dim userGroup As Integer
            If userName <> String.Empty And dataBase IsNot Nothing Then
                Dim sqlLogin As New Microsoft.SqlServer.Management.Smo.Login(sqlServer, Environment.UserDomainName.ToUpper() & "\" & userName)
                sqlLogin.DefaultDatabase = "master"
                sqlLogin.LoginType = Microsoft.SqlServer.Management.Smo.LoginType.WindowsUser
                sqlLogin.Create()

                Dim sqlUser As New Microsoft.SqlServer.Management.Smo.User(dataBase, Environment.UserDomainName.ToUpper() & "\" & userName)
                sqlUser.UserType = Microsoft.SqlServer.Management.Smo.LoginType.WindowsUser
                sqlUser.Login = sqlLogin.Name
                sqlUser.Create()
                sqlUser.AddToRole("db_owner")

                If rbSysAdmin.Checked Then
                    userGroup = 2
                ElseIf rbProdPlanner.Checked Then
                    userGroup = 3
                ElseIf rbUserAdmin.Checked Then
                    userGroup = 4
                ElseIf rbNormalUser.Checked Then
                    userGroup = 1
                End If


                OrklaRTBPL.CommonFacade.CreateCurrentUser(Environment.UserDomainName + "\" + userName, userGroup)
                userName = String.Empty
                lblMessage.ForeColor = Drawing.Color.DarkGreen
                lblMessage.Text = "Bruker tildelt"
            Else
                MessageBox.Show("Please search the user first.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            lblMessage.Text = ex.InnerException.InnerException.Message
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.InnerException.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "Settings", gUserId, gReportID)
        End Try
    End Sub


    Private Sub btnUserSystemUpdate_Click(sender As Object, e As EventArgs) Handles btnUserSystemUpdate.Click
        Try
            Using entities = New DAL.SAPExlEntities()
                Dim userId = entities.vwCurrentUser.SingleOrDefault().ID
                Dim sapSystemRow = entities.UserSAPSystems.SingleOrDefault(Function(ss) ss.SAPSystem = txtSAPSystem.Text And ss.UserId = userId)
                sapSystemRow.SAPUserName = txtSAPUsername.Text
                sapSystemRow.SAPPassword = txtSAPPassword.Text
                entities.SaveChanges()

                'Dim bwSystemRow = entities.UserSAPSystems.SingleOrDefault(Function(ss) ss.SAPSystem = txtBWSystem.Text And ss.UserId = userId)
                'bwSystemRow.SAPUserName = txtBwUserName.Text
                'bwSystemRow.SAPPassword = txtBWPassword.Text
                'entities.SaveChanges()
            End Using
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.InnerException.InnerException.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "Settings", gUserId, gReportID)
        End Try
    End Sub

    Private Sub btnChangeSystem_Click(sender As Object, e As EventArgs) Handles btnChangeSystem.Click
        Try
            Using entities = New DAL.SAPExlEntities()
                Dim userId = entities.vwCurrentUser.SingleOrDefault().ID
                Dim currentUser = entities.CurrentUsers.SingleOrDefault(Function(ss) ss.ID = userId)
                btnChangeSystem.Text = String.Empty
                If currentUser.SAPSystem.Equals("R3P") Then
                    txtSAPSystem.Text = "R3Q"
                    txtBWSystem.Text = "BHQ"
                    btnChangeSystem.Text = "&Bytte til R3P/BHP"
                Else
                    txtSAPSystem.Text = "R3P"
                    txtBWSystem.Text = "BHP"
                    btnChangeSystem.Text = "&Bytte til R3Q/BHQ"
                End If

                Dim sapSystemRow = entities.UserSAPSystems.SingleOrDefault(Function(ss) ss.SAPSystem = currentUser.SAPSystem And ss.UserId = userId)
                sapSystemRow.SAPSystem = txtSAPSystem.Text
                sapSystemRow.SAPUserName = txtSAPUsername.Text
                sapSystemRow.SAPPassword = txtSAPPassword.Text
                entities.SaveChanges()

                Dim bwSystemRow = entities.UserSAPSystems.SingleOrDefault(Function(ss) ss.SAPSystem = currentUser.BwHana And ss.UserId = userId)
                bwSystemRow.SAPSystem = txtBWSystem.Text
                bwSystemRow.SAPUserName = txtBwUserName.Text
                bwSystemRow.SAPPassword = txtBWPassword.Text
                entities.SaveChanges()

                currentUser.SAPSystem = txtSAPSystem.Text
                currentUser.BwHana = txtBWSystem.Text
                entities.SaveChanges()

                If BPL.RfcConnection.CheckR3PConnection() Then
                    Globals.Ribbons.OrklaRT.edtSAPSystem.Text = txtSAPSystem.Text
                    Me.Close()
                Else
                    Dim sapLoginForm As New SAPLogin(txtSAPSystem.Text, True)
                    Call sapLoginForm.ShowDialog()
                    Me.Close()
                End If
            End Using
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.InnerException.InnerException.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "Settings", gUserId, gReportID)
        End Try
    End Sub
End Class
