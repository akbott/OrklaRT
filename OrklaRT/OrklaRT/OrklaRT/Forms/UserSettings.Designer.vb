<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class UserSettings
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.grpUserSettings = New System.Windows.Forms.GroupBox()
        Me.grpUserLevel = New System.Windows.Forms.GroupBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.rbSysAdmin = New System.Windows.Forms.RadioButton()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.rbProdPlanner = New System.Windows.Forms.RadioButton()
        Me.lblUserAdmin = New System.Windows.Forms.Label()
        Me.lblNormalUser = New System.Windows.Forms.Label()
        Me.rbUserAdmin = New System.Windows.Forms.RadioButton()
        Me.rbNormalUser = New System.Windows.Forms.RadioButton()
        Me.lblMessage = New System.Windows.Forms.Label()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.grpUserInformation = New System.Windows.Forms.GroupBox()
        Me.pnlBlock = New System.Windows.Forms.Panel()
        Me.lblUsernameDisplay = New System.Windows.Forms.Label()
        Me.lblTelephone = New System.Windows.Forms.Label()
        Me.lblPostal = New System.Windows.Forms.Label()
        Me.lblCountry = New System.Windows.Forms.Label()
        Me.lblState = New System.Windows.Forms.Label()
        Me.lblCity = New System.Windows.Forms.Label()
        Me.lblCompany = New System.Windows.Forms.Label()
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.lblEmailId = New System.Windows.Forms.Label()
        Me.lblLastName = New System.Windows.Forms.Label()
        Me.lblMiddleName = New System.Windows.Forms.Label()
        Me.lblFirstname = New System.Windows.Forms.Label()
        Me.grpSearchUser = New System.Windows.Forms.GroupBox()
        Me.btnAddToOrklaRT = New System.Windows.Forms.Button()
        Me.txtAddress = New System.Windows.Forms.TextBox()
        Me.lblAddress = New System.Windows.Forms.Label()
        Me.txtSearchUser = New System.Windows.Forms.TextBox()
        Me.label2 = New System.Windows.Forms.Label()
        Me.btnSearchUserName = New System.Windows.Forms.Button()
        Me.grpSapUserDetails = New System.Windows.Forms.GroupBox()
        Me.btnChangeSystem = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.btnUserSystemUpdate = New System.Windows.Forms.Button()
        Me.txtSAPPassword = New System.Windows.Forms.TextBox()
        Me.txtBWPassword = New System.Windows.Forms.TextBox()
        Me.txtSAPUsername = New System.Windows.Forms.TextBox()
        Me.txtBwUserName = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtSAPSystem = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtBWSystem = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.grpUserSettings.SuspendLayout()
        Me.grpUserLevel.SuspendLayout()
        Me.grpUserInformation.SuspendLayout()
        Me.grpSearchUser.SuspendLayout()
        Me.grpSapUserDetails.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpUserSettings
        '
        Me.grpUserSettings.Controls.Add(Me.grpUserLevel)
        Me.grpUserSettings.Controls.Add(Me.lblMessage)
        Me.grpUserSettings.Controls.Add(Me.lblStatus)
        Me.grpUserSettings.Controls.Add(Me.grpUserInformation)
        Me.grpUserSettings.Controls.Add(Me.grpSearchUser)
        Me.grpUserSettings.Enabled = False
        Me.grpUserSettings.Location = New System.Drawing.Point(5, 13)
        Me.grpUserSettings.Name = "grpUserSettings"
        Me.grpUserSettings.Size = New System.Drawing.Size(926, 645)
        Me.grpUserSettings.TabIndex = 44
        Me.grpUserSettings.TabStop = False
        Me.grpUserSettings.Text = "Bruker oppsett"
        '
        'grpUserLevel
        '
        Me.grpUserLevel.Controls.Add(Me.Label7)
        Me.grpUserLevel.Controls.Add(Me.rbSysAdmin)
        Me.grpUserLevel.Controls.Add(Me.Label6)
        Me.grpUserLevel.Controls.Add(Me.rbProdPlanner)
        Me.grpUserLevel.Controls.Add(Me.lblUserAdmin)
        Me.grpUserLevel.Controls.Add(Me.lblNormalUser)
        Me.grpUserLevel.Controls.Add(Me.rbUserAdmin)
        Me.grpUserLevel.Controls.Add(Me.rbNormalUser)
        Me.grpUserLevel.Location = New System.Drawing.Point(16, 489)
        Me.grpUserLevel.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.grpUserLevel.Name = "grpUserLevel"
        Me.grpUserLevel.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.grpUserLevel.Size = New System.Drawing.Size(888, 107)
        Me.grpUserLevel.TabIndex = 47
        Me.grpUserLevel.TabStop = False
        Me.grpUserLevel.Text = "Bruker nivå"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(558, 64)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(257, 20)
        Me.Label7.TabIndex = 7
        Me.Label7.Text = "(OrklaRT administrator - full tilgang)"
        '
        'rbSysAdmin
        '
        Me.rbSysAdmin.AutoSize = True
        Me.rbSysAdmin.Location = New System.Drawing.Point(444, 64)
        Me.rbSysAdmin.Name = "rbSysAdmin"
        Me.rbSysAdmin.Size = New System.Drawing.Size(107, 24)
        Me.rbSysAdmin.TabIndex = 6
        Me.rbSysAdmin.Text = "Sys admin"
        Me.rbSysAdmin.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(594, 32)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(255, 20)
        Me.Label6.TabIndex = 5
        Me.Label6.Text = "(kan lage prod plan bilde i OrklaRT)"
        '
        'rbProdPlanner
        '
        Me.rbProdPlanner.AutoSize = True
        Me.rbProdPlanner.Location = New System.Drawing.Point(445, 32)
        Me.rbProdPlanner.Name = "rbProdPlanner"
        Me.rbProdPlanner.Size = New System.Drawing.Size(145, 24)
        Me.rbProdPlanner.TabIndex = 4
        Me.rbProdPlanner.Text = "Prod.planlegger"
        Me.rbProdPlanner.UseVisualStyleBackColor = True
        '
        'lblUserAdmin
        '
        Me.lblUserAdmin.AutoSize = True
        Me.lblUserAdmin.Location = New System.Drawing.Point(146, 64)
        Me.lblUserAdmin.Name = "lblUserAdmin"
        Me.lblUserAdmin.Size = New System.Drawing.Size(247, 20)
        Me.lblUserAdmin.TabIndex = 3
        Me.lblUserAdmin.Text = "(kan operette ny bruker i OrklaRT)"
        '
        'lblNormalUser
        '
        Me.lblNormalUser.AutoSize = True
        Me.lblNormalUser.Location = New System.Drawing.Point(145, 30)
        Me.lblNormalUser.Name = "lblNormalUser"
        Me.lblNormalUser.Size = New System.Drawing.Size(249, 20)
        Me.lblNormalUser.TabIndex = 2
        Me.lblNormalUser.Text = "(Vannlig bruker tilgang til OrklaRT)"
        '
        'rbUserAdmin
        '
        Me.rbUserAdmin.AutoSize = True
        Me.rbUserAdmin.Location = New System.Drawing.Point(16, 64)
        Me.rbUserAdmin.Name = "rbUserAdmin"
        Me.rbUserAdmin.Size = New System.Drawing.Size(128, 24)
        Me.rbUserAdmin.TabIndex = 1
        Me.rbUserAdmin.Text = "Bruker admin"
        Me.rbUserAdmin.UseVisualStyleBackColor = True
        '
        'rbNormalUser
        '
        Me.rbNormalUser.AutoSize = True
        Me.rbNormalUser.Checked = True
        Me.rbNormalUser.Location = New System.Drawing.Point(16, 28)
        Me.rbNormalUser.Name = "rbNormalUser"
        Me.rbNormalUser.Size = New System.Drawing.Size(127, 24)
        Me.rbNormalUser.TabIndex = 0
        Me.rbNormalUser.TabStop = True
        Me.rbNormalUser.Text = "Vanlig bruker"
        Me.rbNormalUser.UseVisualStyleBackColor = True
        '
        'lblMessage
        '
        Me.lblMessage.AutoSize = True
        Me.lblMessage.Location = New System.Drawing.Point(85, 613)
        Me.lblMessage.Name = "lblMessage"
        Me.lblMessage.Size = New System.Drawing.Size(0, 20)
        Me.lblMessage.TabIndex = 46
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Location = New System.Drawing.Point(13, 613)
        Me.lblStatus.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(64, 20)
        Me.lblStatus.TabIndex = 45
        Me.lblStatus.Text = "Status :"
        Me.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'grpUserInformation
        '
        Me.grpUserInformation.Controls.Add(Me.pnlBlock)
        Me.grpUserInformation.Controls.Add(Me.lblUsernameDisplay)
        Me.grpUserInformation.Controls.Add(Me.lblTelephone)
        Me.grpUserInformation.Controls.Add(Me.lblPostal)
        Me.grpUserInformation.Controls.Add(Me.lblCountry)
        Me.grpUserInformation.Controls.Add(Me.lblState)
        Me.grpUserInformation.Controls.Add(Me.lblCity)
        Me.grpUserInformation.Controls.Add(Me.lblCompany)
        Me.grpUserInformation.Controls.Add(Me.lblTitle)
        Me.grpUserInformation.Controls.Add(Me.lblEmailId)
        Me.grpUserInformation.Controls.Add(Me.lblLastName)
        Me.grpUserInformation.Controls.Add(Me.lblMiddleName)
        Me.grpUserInformation.Controls.Add(Me.lblFirstname)
        Me.grpUserInformation.Location = New System.Drawing.Point(16, 159)
        Me.grpUserInformation.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.grpUserInformation.Name = "grpUserInformation"
        Me.grpUserInformation.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.grpUserInformation.Size = New System.Drawing.Size(890, 320)
        Me.grpUserInformation.TabIndex = 43
        Me.grpUserInformation.TabStop = False
        Me.grpUserInformation.Text = "Bruker informasjon"
        '
        'pnlBlock
        '
        Me.pnlBlock.Location = New System.Drawing.Point(9, 25)
        Me.pnlBlock.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.pnlBlock.Name = "pnlBlock"
        Me.pnlBlock.Size = New System.Drawing.Size(870, 286)
        Me.pnlBlock.TabIndex = 34
        '
        'lblUsernameDisplay
        '
        Me.lblUsernameDisplay.AutoSize = True
        Me.lblUsernameDisplay.Location = New System.Drawing.Point(46, 34)
        Me.lblUsernameDisplay.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblUsernameDisplay.Name = "lblUsernameDisplay"
        Me.lblUsernameDisplay.Size = New System.Drawing.Size(91, 20)
        Me.lblUsernameDisplay.TabIndex = 32
        Me.lblUsernameDisplay.Text = "Username :"
        '
        'lblTelephone
        '
        Me.lblTelephone.AutoSize = True
        Me.lblTelephone.Location = New System.Drawing.Point(518, 158)
        Me.lblTelephone.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblTelephone.Name = "lblTelephone"
        Me.lblTelephone.Size = New System.Drawing.Size(120, 20)
        Me.lblTelephone.TabIndex = 31
        Me.lblTelephone.Text = "Telephone No. :"
        Me.lblTelephone.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblPostal
        '
        Me.lblPostal.AutoSize = True
        Me.lblPostal.Location = New System.Drawing.Point(538, 120)
        Me.lblPostal.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblPostal.Name = "lblPostal"
        Me.lblPostal.Size = New System.Drawing.Size(103, 20)
        Me.lblPostal.TabIndex = 30
        Me.lblPostal.Text = "Postal Code :"
        Me.lblPostal.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblCountry
        '
        Me.lblCountry.AutoSize = True
        Me.lblCountry.Location = New System.Drawing.Point(570, 85)
        Me.lblCountry.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblCountry.Name = "lblCountry"
        Me.lblCountry.Size = New System.Drawing.Size(72, 20)
        Me.lblCountry.TabIndex = 29
        Me.lblCountry.Text = "Country :"
        Me.lblCountry.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblState
        '
        Me.lblState.AutoSize = True
        Me.lblState.Location = New System.Drawing.Point(586, 48)
        Me.lblState.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblState.Name = "lblState"
        Me.lblState.Size = New System.Drawing.Size(56, 20)
        Me.lblState.TabIndex = 28
        Me.lblState.Text = "State :"
        Me.lblState.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblCity
        '
        Me.lblCity.AutoSize = True
        Me.lblCity.Location = New System.Drawing.Point(93, 286)
        Me.lblCity.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblCity.Name = "lblCity"
        Me.lblCity.Size = New System.Drawing.Size(43, 20)
        Me.lblCity.TabIndex = 27
        Me.lblCity.Text = "City :"
        Me.lblCity.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblCompany
        '
        Me.lblCompany.AutoSize = True
        Me.lblCompany.Location = New System.Drawing.Point(52, 251)
        Me.lblCompany.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblCompany.Name = "lblCompany"
        Me.lblCompany.Size = New System.Drawing.Size(84, 20)
        Me.lblCompany.TabIndex = 26
        Me.lblCompany.Text = "Company :"
        Me.lblCompany.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblTitle
        '
        Me.lblTitle.AutoSize = True
        Me.lblTitle.Location = New System.Drawing.Point(88, 214)
        Me.lblTitle.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(46, 20)
        Me.lblTitle.TabIndex = 25
        Me.lblTitle.Text = "Title :"
        Me.lblTitle.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblEmailId
        '
        Me.lblEmailId.AutoSize = True
        Me.lblEmailId.Location = New System.Drawing.Point(60, 177)
        Me.lblEmailId.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblEmailId.Name = "lblEmailId"
        Me.lblEmailId.Size = New System.Drawing.Size(77, 20)
        Me.lblEmailId.TabIndex = 24
        Me.lblEmailId.Text = "Email ID :"
        Me.lblEmailId.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblLastName
        '
        Me.lblLastName.AutoSize = True
        Me.lblLastName.Location = New System.Drawing.Point(42, 138)
        Me.lblLastName.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblLastName.Name = "lblLastName"
        Me.lblLastName.Size = New System.Drawing.Size(94, 20)
        Me.lblLastName.TabIndex = 23
        Me.lblLastName.Text = "Last Name :"
        Me.lblLastName.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblMiddleName
        '
        Me.lblMiddleName.AutoSize = True
        Me.lblMiddleName.Location = New System.Drawing.Point(26, 103)
        Me.lblMiddleName.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblMiddleName.Name = "lblMiddleName"
        Me.lblMiddleName.Size = New System.Drawing.Size(109, 20)
        Me.lblMiddleName.TabIndex = 22
        Me.lblMiddleName.Text = "Middle Name :"
        Me.lblMiddleName.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblFirstname
        '
        Me.lblFirstname.AutoSize = True
        Me.lblFirstname.Location = New System.Drawing.Point(44, 66)
        Me.lblFirstname.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblFirstname.Name = "lblFirstname"
        Me.lblFirstname.Size = New System.Drawing.Size(94, 20)
        Me.lblFirstname.TabIndex = 21
        Me.lblFirstname.Text = "First Name :"
        Me.lblFirstname.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'grpSearchUser
        '
        Me.grpSearchUser.Controls.Add(Me.btnAddToOrklaRT)
        Me.grpSearchUser.Controls.Add(Me.txtAddress)
        Me.grpSearchUser.Controls.Add(Me.lblAddress)
        Me.grpSearchUser.Controls.Add(Me.txtSearchUser)
        Me.grpSearchUser.Controls.Add(Me.label2)
        Me.grpSearchUser.Controls.Add(Me.btnSearchUserName)
        Me.grpSearchUser.Location = New System.Drawing.Point(17, 39)
        Me.grpSearchUser.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.grpSearchUser.Name = "grpSearchUser"
        Me.grpSearchUser.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.grpSearchUser.Size = New System.Drawing.Size(888, 111)
        Me.grpSearchUser.TabIndex = 44
        Me.grpSearchUser.TabStop = False
        Me.grpSearchUser.Text = "Søk bruker"
        '
        'btnAddToOrklaRT
        '
        Me.btnAddToOrklaRT.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.btnAddToOrklaRT.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnAddToOrklaRT.FlatAppearance.BorderColor = System.Drawing.SystemColors.ButtonShadow
        Me.btnAddToOrklaRT.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnAddToOrklaRT.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnAddToOrklaRT.Location = New System.Drawing.Point(712, 65)
        Me.btnAddToOrklaRT.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.btnAddToOrklaRT.Name = "btnAddToOrklaRT"
        Me.btnAddToOrklaRT.Size = New System.Drawing.Size(148, 35)
        Me.btnAddToOrklaRT.TabIndex = 25
        Me.btnAddToOrklaRT.Text = "&Legg til bruker"
        Me.btnAddToOrklaRT.UseVisualStyleBackColor = False
        '
        'txtAddress
        '
        Me.txtAddress.Location = New System.Drawing.Point(260, 30)
        Me.txtAddress.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtAddress.Name = "txtAddress"
        Me.txtAddress.Size = New System.Drawing.Size(368, 26)
        Me.txtAddress.TabIndex = 23
        '
        'lblAddress
        '
        Me.lblAddress.AutoSize = True
        Me.lblAddress.Location = New System.Drawing.Point(164, 33)
        Me.lblAddress.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblAddress.Name = "lblAddress"
        Me.lblAddress.Size = New System.Drawing.Size(78, 20)
        Me.lblAddress.TabIndex = 24
        Me.lblAddress.Text = "Domene :"
        Me.lblAddress.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtSearchUser
        '
        Me.txtSearchUser.Location = New System.Drawing.Point(261, 66)
        Me.txtSearchUser.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtSearchUser.Name = "txtSearchUser"
        Me.txtSearchUser.Size = New System.Drawing.Size(367, 26)
        Me.txtSearchUser.TabIndex = 3
        '
        'label2
        '
        Me.label2.AutoSize = True
        Me.label2.Location = New System.Drawing.Point(3, 69)
        Me.label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.label2.Name = "label2"
        Me.label2.Size = New System.Drawing.Size(248, 20)
        Me.label2.TabIndex = 22
        Me.label2.Text = "Bruk etternavn/brukernavn/epost :"
        Me.label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btnSearchUserName
        '
        Me.btnSearchUserName.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.btnSearchUserName.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnSearchUserName.FlatAppearance.BorderColor = System.Drawing.SystemColors.ButtonShadow
        Me.btnSearchUserName.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnSearchUserName.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnSearchUserName.Location = New System.Drawing.Point(732, 20)
        Me.btnSearchUserName.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.btnSearchUserName.Name = "btnSearchUserName"
        Me.btnSearchUserName.Size = New System.Drawing.Size(112, 35)
        Me.btnSearchUserName.TabIndex = 4
        Me.btnSearchUserName.Text = "&Søk"
        Me.btnSearchUserName.UseVisualStyleBackColor = False
        '
        'grpSapUserDetails
        '
        Me.grpSapUserDetails.Controls.Add(Me.btnChangeSystem)
        Me.grpSapUserDetails.Controls.Add(Me.Label3)
        Me.grpSapUserDetails.Controls.Add(Me.btnUserSystemUpdate)
        Me.grpSapUserDetails.Controls.Add(Me.txtSAPPassword)
        Me.grpSapUserDetails.Controls.Add(Me.txtBWPassword)
        Me.grpSapUserDetails.Controls.Add(Me.txtSAPUsername)
        Me.grpSapUserDetails.Controls.Add(Me.txtBwUserName)
        Me.grpSapUserDetails.Controls.Add(Me.Label1)
        Me.grpSapUserDetails.Controls.Add(Me.txtSAPSystem)
        Me.grpSapUserDetails.Controls.Add(Me.Label4)
        Me.grpSapUserDetails.Controls.Add(Me.txtBWSystem)
        Me.grpSapUserDetails.Controls.Add(Me.Label5)
        Me.grpSapUserDetails.Enabled = False
        Me.grpSapUserDetails.Location = New System.Drawing.Point(5, 666)
        Me.grpSapUserDetails.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.grpSapUserDetails.Name = "grpSapUserDetails"
        Me.grpSapUserDetails.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.grpSapUserDetails.Size = New System.Drawing.Size(926, 202)
        Me.grpSapUserDetails.TabIndex = 45
        Me.grpSapUserDetails.TabStop = False
        Me.grpSapUserDetails.Text = "SAP bruker info"
        '
        'btnChangeSystem
        '
        Me.btnChangeSystem.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.btnChangeSystem.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnChangeSystem.FlatAppearance.BorderColor = System.Drawing.SystemColors.ButtonShadow
        Me.btnChangeSystem.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnChangeSystem.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnChangeSystem.Location = New System.Drawing.Point(741, 64)
        Me.btnChangeSystem.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.btnChangeSystem.Name = "btnChangeSystem"
        Me.btnChangeSystem.Size = New System.Drawing.Size(136, 64)
        Me.btnChangeSystem.TabIndex = 47
        Me.btnChangeSystem.Text = "&Bytte til "
        Me.btnChangeSystem.UseVisualStyleBackColor = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label3.Location = New System.Drawing.Point(56, 31)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(175, 20)
        Me.Label3.TabIndex = 46
        Me.Label3.Text = "Kun for informasjon :"
        '
        'btnUserSystemUpdate
        '
        Me.btnUserSystemUpdate.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.btnUserSystemUpdate.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnUserSystemUpdate.FlatAppearance.BorderColor = System.Drawing.SystemColors.ButtonShadow
        Me.btnUserSystemUpdate.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnUserSystemUpdate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnUserSystemUpdate.Location = New System.Drawing.Point(749, 139)
        Me.btnUserSystemUpdate.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.btnUserSystemUpdate.Name = "btnUserSystemUpdate"
        Me.btnUserSystemUpdate.Size = New System.Drawing.Size(112, 35)
        Me.btnUserSystemUpdate.TabIndex = 30
        Me.btnUserSystemUpdate.Text = "&Oppdatere"
        Me.btnUserSystemUpdate.UseVisualStyleBackColor = False
        Me.btnUserSystemUpdate.Visible = False
        '
        'txtSAPPassword
        '
        Me.txtSAPPassword.Enabled = False
        Me.txtSAPPassword.Location = New System.Drawing.Point(526, 102)
        Me.txtSAPPassword.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtSAPPassword.Name = "txtSAPPassword"
        Me.txtSAPPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtSAPPassword.Size = New System.Drawing.Size(169, 26)
        Me.txtSAPPassword.TabIndex = 29
        '
        'txtBWPassword
        '
        Me.txtBWPassword.Location = New System.Drawing.Point(526, 148)
        Me.txtBWPassword.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtBWPassword.Name = "txtBWPassword"
        Me.txtBWPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtBWPassword.Size = New System.Drawing.Size(169, 26)
        Me.txtBWPassword.TabIndex = 28
        Me.txtBWPassword.Visible = False
        '
        'txtSAPUsername
        '
        Me.txtSAPUsername.Enabled = False
        Me.txtSAPUsername.Location = New System.Drawing.Point(282, 102)
        Me.txtSAPUsername.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtSAPUsername.Name = "txtSAPUsername"
        Me.txtSAPUsername.Size = New System.Drawing.Size(169, 26)
        Me.txtSAPUsername.TabIndex = 27
        '
        'txtBwUserName
        '
        Me.txtBwUserName.Location = New System.Drawing.Point(282, 148)
        Me.txtBwUserName.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtBwUserName.Name = "txtBwUserName"
        Me.txtBwUserName.Size = New System.Drawing.Size(169, 26)
        Me.txtBwUserName.TabIndex = 26
        Me.txtBwUserName.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(531, 64)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(75, 20)
        Me.Label1.TabIndex = 25
        Me.Label1.Text = "Passord :"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtSAPSystem
        '
        Me.txtSAPSystem.Enabled = False
        Me.txtSAPSystem.Location = New System.Drawing.Point(52, 102)
        Me.txtSAPSystem.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtSAPSystem.Name = "txtSAPSystem"
        Me.txtSAPSystem.Size = New System.Drawing.Size(169, 26)
        Me.txtSAPSystem.TabIndex = 23
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(56, 64)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(84, 20)
        Me.Label4.TabIndex = 24
        Me.Label4.Text = "Systemer :"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtBWSystem
        '
        Me.txtBWSystem.Location = New System.Drawing.Point(52, 148)
        Me.txtBWSystem.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtBWSystem.Name = "txtBWSystem"
        Me.txtBWSystem.Size = New System.Drawing.Size(169, 26)
        Me.txtBWSystem.TabIndex = 3
        Me.txtBWSystem.Visible = False
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(287, 64)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(90, 20)
        Me.Label5.TabIndex = 22
        Me.Label5.Text = "Brukernavn"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'UserSettings
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(933, 876)
        Me.Controls.Add(Me.grpSapUserDetails)
        Me.Controls.Add(Me.grpUserSettings)
        Me.Name = "UserSettings"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Bruker oppsett"
        Me.grpUserSettings.ResumeLayout(False)
        Me.grpUserSettings.PerformLayout()
        Me.grpUserLevel.ResumeLayout(False)
        Me.grpUserLevel.PerformLayout()
        Me.grpUserInformation.ResumeLayout(False)
        Me.grpUserInformation.PerformLayout()
        Me.grpSearchUser.ResumeLayout(False)
        Me.grpSearchUser.PerformLayout()
        Me.grpSapUserDetails.ResumeLayout(False)
        Me.grpSapUserDetails.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents grpUserSettings As System.Windows.Forms.GroupBox
    Private WithEvents grpUserInformation As System.Windows.Forms.GroupBox
    Private WithEvents pnlBlock As System.Windows.Forms.Panel
    Private WithEvents lblUsernameDisplay As System.Windows.Forms.Label
    Private WithEvents lblTelephone As System.Windows.Forms.Label
    Private WithEvents lblPostal As System.Windows.Forms.Label
    Private WithEvents lblCountry As System.Windows.Forms.Label
    Private WithEvents lblState As System.Windows.Forms.Label
    Private WithEvents lblCity As System.Windows.Forms.Label
    Private WithEvents lblCompany As System.Windows.Forms.Label
    Private WithEvents lblTitle As System.Windows.Forms.Label
    Private WithEvents lblEmailId As System.Windows.Forms.Label
    Private WithEvents lblLastName As System.Windows.Forms.Label
    Private WithEvents lblMiddleName As System.Windows.Forms.Label
    Private WithEvents lblFirstname As System.Windows.Forms.Label
    Private WithEvents grpSearchUser As System.Windows.Forms.GroupBox
    Private WithEvents btnAddToOrklaRT As System.Windows.Forms.Button
    Private WithEvents txtAddress As System.Windows.Forms.TextBox
    Private WithEvents lblAddress As System.Windows.Forms.Label
    Private WithEvents txtSearchUser As System.Windows.Forms.TextBox
    Private WithEvents label2 As System.Windows.Forms.Label
    Private WithEvents btnSearchUserName As System.Windows.Forms.Button
    Private WithEvents grpSapUserDetails As System.Windows.Forms.GroupBox
    Private WithEvents txtSAPPassword As System.Windows.Forms.TextBox
    Private WithEvents txtBWPassword As System.Windows.Forms.TextBox
    Private WithEvents txtSAPUsername As System.Windows.Forms.TextBox
    Private WithEvents txtBwUserName As System.Windows.Forms.TextBox
    Private WithEvents Label1 As System.Windows.Forms.Label
    Private WithEvents txtSAPSystem As System.Windows.Forms.TextBox
    Private WithEvents Label4 As System.Windows.Forms.Label
    Private WithEvents txtBWSystem As System.Windows.Forms.TextBox
    Private WithEvents Label5 As System.Windows.Forms.Label
    Private WithEvents lblStatus As System.Windows.Forms.Label
    Private WithEvents btnUserSystemUpdate As System.Windows.Forms.Button
    Friend WithEvents lblMessage As System.Windows.Forms.Label
    Private WithEvents Label3 As System.Windows.Forms.Label
    Private WithEvents grpUserLevel As System.Windows.Forms.GroupBox
    Friend WithEvents rbUserAdmin As System.Windows.Forms.RadioButton
    Friend WithEvents rbNormalUser As System.Windows.Forms.RadioButton
    Friend WithEvents lblUserAdmin As System.Windows.Forms.Label
    Friend WithEvents lblNormalUser As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents rbProdPlanner As System.Windows.Forms.RadioButton
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents rbSysAdmin As System.Windows.Forms.RadioButton
    Private WithEvents btnChangeSystem As System.Windows.Forms.Button
End Class
