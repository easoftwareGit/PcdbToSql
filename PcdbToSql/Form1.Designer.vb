Partial Public Class Form1
    Inherits DevExpress.XtraEditors.XtraForm

    ''' <summary>
    ''' Required designer variable.
    ''' </summary>
    Private components As System.ComponentModel.IContainer = Nothing

    ''' <summary>
    ''' Clean up any resources being used.
    ''' </summary>
    ''' <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso (components IsNot Nothing) Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

#Region "Windows Form Designer generated code"

    ''' <summary>
    ''' Required method for Designer support - do not modify
    ''' the contents of this method with the code editor.
    ''' </summary>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim EditorButtonImageOptions1 As DevExpress.XtraEditors.Controls.EditorButtonImageOptions = New DevExpress.XtraEditors.Controls.EditorButtonImageOptions()
        Dim SerializableAppearanceObject1 As DevExpress.Utils.SerializableAppearanceObject = New DevExpress.Utils.SerializableAppearanceObject()
        Dim SerializableAppearanceObject2 As DevExpress.Utils.SerializableAppearanceObject = New DevExpress.Utils.SerializableAppearanceObject()
        Dim SerializableAppearanceObject3 As DevExpress.Utils.SerializableAppearanceObject = New DevExpress.Utils.SerializableAppearanceObject()
        Dim SerializableAppearanceObject4 As DevExpress.Utils.SerializableAppearanceObject = New DevExpress.Utils.SerializableAppearanceObject()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.CurrentLabelControl = New DevExpress.XtraEditors.LabelControl()
        Me.XtraOpenFileDialog1 = New DevExpress.XtraEditors.XtraOpenFileDialog(Me.components)
        Me.GetTablesButton = New DevExpress.XtraEditors.SimpleButton()
        Me.AccessOleDbConnection = New System.Data.OleDb.OleDbConnection()
        Me.StratTimeLabelControl = New DevExpress.XtraEditors.LabelControl()
        Me.StartTimeEdit = New DevExpress.XtraEditors.TimeEdit()
        Me.TotalProgressLabelControl = New DevExpress.XtraEditors.LabelControl()
        Me.TotalProgressBarControl = New DevExpress.XtraEditors.ProgressBarControl()
        Me.ActionProgressBarControl = New DevExpress.XtraEditors.ProgressBarControl()
        Me.ActionProgressLabelControl = New DevExpress.XtraEditors.LabelControl()
        Me.ActionTextEdit = New DevExpress.XtraEditors.TextEdit()
        Me.ActionLabelControl = New DevExpress.XtraEditors.LabelControl()
        Me.MySqlServerTextEdit = New DevExpress.XtraEditors.TextEdit()
        Me.MyServerLabelControl = New DevExpress.XtraEditors.LabelControl()
        Me.MySqlPcdbNameTextEdit = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl2 = New DevExpress.XtraEditors.LabelControl()
        Me.UserIdTextEdit = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl3 = New DevExpress.XtraEditors.LabelControl()
        Me.PasswordTextEdit = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl4 = New DevExpress.XtraEditors.LabelControl()
        Me.MySqlConnectButton = New DevExpress.XtraEditors.SimpleButton()
        Me.TablesCheckedListBox = New DevExpress.XtraEditors.CheckedListBoxControl()
        Me.AccessButtonEdit = New DevExpress.XtraEditors.ButtonEdit()
        Me.CopyTablesButton = New DevExpress.XtraEditors.SimpleButton()
        Me.SuccessLabel = New DevExpress.XtraEditors.LabelControl()
        Me.CopyQueriesButton = New DevExpress.XtraEditors.SimpleButton()
        Me.QueriesCheckedListBox = New DevExpress.XtraEditors.CheckedListBoxControl()
        Me.Select3Point0Button = New DevExpress.XtraEditors.SimpleButton()
        Me.Non3Point0Button = New DevExpress.XtraEditors.SimpleButton()
        Me.Non3Point0ListBoxControl = New DevExpress.XtraEditors.ListBoxControl()
        Me.CopyAllButton = New DevExpress.XtraEditors.SimpleButton()
        Me.CreateDatabaseButton = New DevExpress.XtraEditors.SimpleButton()
        Me.CreateTestUserButton = New DevExpress.XtraEditors.SimpleButton()
        Me.ChangePasswordButton = New DevExpress.XtraEditors.SimpleButton()
        Me.WriteConfigButton = New DevExpress.XtraEditors.SimpleButton()
        Me.ReadConfigButton = New DevExpress.XtraEditors.SimpleButton()
        Me.GetTableSchemaButton = New DevExpress.XtraEditors.SimpleButton()
        Me.LabelControl1 = New DevExpress.XtraEditors.LabelControl()
        Me.PersistSecurityInfoCheckBox = New System.Windows.Forms.CheckBox()
        Me.NoneRadioButton = New System.Windows.Forms.RadioButton()
        Me.PreferredRadioButton = New System.Windows.Forms.RadioButton()
        Me.RequiredRadioButton = New System.Windows.Forms.RadioButton()
        Me.VerifyCARadioButton = New System.Windows.Forms.RadioButton()
        Me.VerifyFullRadioButton = New System.Windows.Forms.RadioButton()
        Me.LabelControl5 = New DevExpress.XtraEditors.LabelControl()
        Me.OtherButton = New DevExpress.XtraEditors.SimpleButton()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.RunningTimeLabelControl = New DevExpress.XtraEditors.LabelControl()
        Me.RunningTimeTextEdit = New DevExpress.XtraEditors.TextEdit()
        Me.SimpleButton1 = New DevExpress.XtraEditors.SimpleButton()
        Me.SimpleButton2 = New DevExpress.XtraEditors.SimpleButton()
        Me.DoCancelButton = New DevExpress.XtraEditors.SimpleButton()
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.CanceledLabel = New DevExpress.XtraEditors.LabelControl()
        Me.CopiedTablesLabel = New DevExpress.XtraEditors.LabelControl()
        Me.CopiedQueriesLabel = New DevExpress.XtraEditors.LabelControl()
        CType(Me.StartTimeEdit.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TotalProgressBarControl.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ActionProgressBarControl.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ActionTextEdit.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MySqlServerTextEdit.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MySqlPcdbNameTextEdit.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UserIdTextEdit.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PasswordTextEdit.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TablesCheckedListBox, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AccessButtonEdit.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.QueriesCheckedListBox, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Non3Point0ListBoxControl, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RunningTimeTextEdit.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'CurrentLabelControl
        '
        Me.CurrentLabelControl.Location = New System.Drawing.Point(12, 16)
        Me.CurrentLabelControl.Name = "CurrentLabelControl"
        Me.CurrentLabelControl.Size = New System.Drawing.Size(128, 13)
        Me.CurrentLabelControl.TabIndex = 0
        Me.CurrentLabelControl.Text = "Access PCDB Database file"
        '
        'XtraOpenFileDialog1
        '
        Me.XtraOpenFileDialog1.DefaultExt = "mdb"
        Me.XtraOpenFileDialog1.FileName = "Ewcg2000.mdb"
        Me.XtraOpenFileDialog1.Filter = "Access Database files (*.mdb)|*.mdb"
        Me.XtraOpenFileDialog1.Title = "Find Ewcg2000.mdb"
        '
        'GetTablesButton
        '
        Me.GetTablesButton.Location = New System.Drawing.Point(12, 270)
        Me.GetTablesButton.Name = "GetTablesButton"
        Me.GetTablesButton.Size = New System.Drawing.Size(118, 23)
        Me.GetTablesButton.TabIndex = 18
        Me.GetTablesButton.Text = "Get Tables && Queries"
        '
        'AccessOleDbConnection
        '
        Me.AccessOleDbConnection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Projects\PcdbData\Ewcg2000.mdb"
        '
        'StratTimeLabelControl
        '
        Me.StratTimeLabelControl.Location = New System.Drawing.Point(13, 246)
        Me.StratTimeLabelControl.Name = "StratTimeLabelControl"
        Me.StratTimeLabelControl.Size = New System.Drawing.Size(49, 13)
        Me.StratTimeLabelControl.TabIndex = 16
        Me.StratTimeLabelControl.Text = "Start Time"
        Me.StratTimeLabelControl.Visible = False
        '
        'StartTimeEdit
        '
        Me.StartTimeEdit.EditValue = New Date(2015, 3, 23, 0, 0, 0, 0)
        Me.StartTimeEdit.Location = New System.Drawing.Point(151, 244)
        Me.StartTimeEdit.Name = "StartTimeEdit"
        Me.StartTimeEdit.Properties.Appearance.BackColor = System.Drawing.Color.White
        Me.StartTimeEdit.Properties.Appearance.Options.UseBackColor = True
        Me.StartTimeEdit.Properties.DisplayFormat.FormatString = "MM/dd/yyyy hh:mm tt"
        Me.StartTimeEdit.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime
        Me.StartTimeEdit.Properties.ReadOnly = True
        Me.StartTimeEdit.Size = New System.Drawing.Size(169, 20)
        Me.StartTimeEdit.TabIndex = 17
        Me.StartTimeEdit.Visible = False
        '
        'TotalProgressLabelControl
        '
        Me.TotalProgressLabelControl.Location = New System.Drawing.Point(13, 220)
        Me.TotalProgressLabelControl.Name = "TotalProgressLabelControl"
        Me.TotalProgressLabelControl.Size = New System.Drawing.Size(69, 13)
        Me.TotalProgressLabelControl.TabIndex = 14
        Me.TotalProgressLabelControl.Text = "Total Progress"
        Me.TotalProgressLabelControl.Visible = False
        '
        'TotalProgressBarControl
        '
        Me.TotalProgressBarControl.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TotalProgressBarControl.EditValue = 99
        Me.TotalProgressBarControl.Location = New System.Drawing.Point(151, 220)
        Me.TotalProgressBarControl.Name = "TotalProgressBarControl"
        Me.TotalProgressBarControl.Properties.EndColor = System.Drawing.Color.DeepSkyBlue
        Me.TotalProgressBarControl.Properties.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.Style3D
        Me.TotalProgressBarControl.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.TotalProgressBarControl.Properties.ProgressViewStyle = DevExpress.XtraEditors.Controls.ProgressViewStyle.Solid
        Me.TotalProgressBarControl.Properties.ShowTitle = True
        Me.TotalProgressBarControl.Properties.StartColor = System.Drawing.Color.RoyalBlue
        Me.TotalProgressBarControl.Properties.Step = 1
        Me.TotalProgressBarControl.ShowProgressInTaskBar = True
        Me.TotalProgressBarControl.Size = New System.Drawing.Size(539, 18)
        Me.TotalProgressBarControl.TabIndex = 15
        Me.TotalProgressBarControl.Visible = False
        '
        'ActionProgressBarControl
        '
        Me.ActionProgressBarControl.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ActionProgressBarControl.EditValue = 99
        Me.ActionProgressBarControl.Location = New System.Drawing.Point(151, 196)
        Me.ActionProgressBarControl.Name = "ActionProgressBarControl"
        Me.ActionProgressBarControl.Properties.EndColor = System.Drawing.Color.DeepSkyBlue
        Me.ActionProgressBarControl.Properties.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.Style3D
        Me.ActionProgressBarControl.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.ActionProgressBarControl.Properties.ProgressViewStyle = DevExpress.XtraEditors.Controls.ProgressViewStyle.Solid
        Me.ActionProgressBarControl.Properties.ShowTitle = True
        Me.ActionProgressBarControl.Properties.StartColor = System.Drawing.Color.RoyalBlue
        Me.ActionProgressBarControl.Properties.Step = 1
        Me.ActionProgressBarControl.Size = New System.Drawing.Size(539, 18)
        Me.ActionProgressBarControl.TabIndex = 13
        Me.ActionProgressBarControl.Visible = False
        '
        'ActionProgressLabelControl
        '
        Me.ActionProgressLabelControl.Location = New System.Drawing.Point(13, 196)
        Me.ActionProgressLabelControl.Name = "ActionProgressLabelControl"
        Me.ActionProgressLabelControl.Size = New System.Drawing.Size(75, 13)
        Me.ActionProgressLabelControl.TabIndex = 12
        Me.ActionProgressLabelControl.Text = "Action Progress"
        Me.ActionProgressLabelControl.Visible = False
        '
        'ActionTextEdit
        '
        Me.ActionTextEdit.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ActionTextEdit.Location = New System.Drawing.Point(151, 170)
        Me.ActionTextEdit.Name = "ActionTextEdit"
        Me.ActionTextEdit.Properties.Appearance.BackColor = System.Drawing.Color.White
        Me.ActionTextEdit.Properties.Appearance.Options.UseBackColor = True
        Me.ActionTextEdit.Properties.ReadOnly = True
        Me.ActionTextEdit.Size = New System.Drawing.Size(539, 20)
        Me.ActionTextEdit.TabIndex = 11
        Me.ActionTextEdit.Visible = False
        '
        'ActionLabelControl
        '
        Me.ActionLabelControl.Location = New System.Drawing.Point(13, 173)
        Me.ActionLabelControl.Name = "ActionLabelControl"
        Me.ActionLabelControl.Size = New System.Drawing.Size(30, 13)
        Me.ActionLabelControl.TabIndex = 10
        Me.ActionLabelControl.Text = "Action"
        Me.ActionLabelControl.Visible = False
        '
        'MySqlServerTextEdit
        '
        Me.MySqlServerTextEdit.EditValue = "localhost"
        Me.MySqlServerTextEdit.Location = New System.Drawing.Point(150, 40)
        Me.MySqlServerTextEdit.Name = "MySqlServerTextEdit"
        Me.MySqlServerTextEdit.Size = New System.Drawing.Size(539, 20)
        Me.MySqlServerTextEdit.TabIndex = 3
        '
        'MyServerLabelControl
        '
        Me.MyServerLabelControl.Location = New System.Drawing.Point(12, 43)
        Me.MyServerLabelControl.Name = "MyServerLabelControl"
        Me.MyServerLabelControl.Size = New System.Drawing.Size(63, 13)
        Me.MyServerLabelControl.TabIndex = 2
        Me.MyServerLabelControl.Text = "MySql Server"
        '
        'MySqlPcdbNameTextEdit
        '
        Me.MySqlPcdbNameTextEdit.EditValue = "pcdb_test"
        Me.MySqlPcdbNameTextEdit.Location = New System.Drawing.Point(150, 66)
        Me.MySqlPcdbNameTextEdit.Name = "MySqlPcdbNameTextEdit"
        Me.MySqlPcdbNameTextEdit.Size = New System.Drawing.Size(135, 20)
        Me.MySqlPcdbNameTextEdit.TabIndex = 5
        '
        'LabelControl2
        '
        Me.LabelControl2.Location = New System.Drawing.Point(12, 69)
        Me.LabelControl2.Name = "LabelControl2"
        Me.LabelControl2.Size = New System.Drawing.Size(106, 13)
        Me.LabelControl2.TabIndex = 4
        Me.LabelControl2.Text = "MySql PCDB Database"
        '
        'UserIdTextEdit
        '
        Me.UserIdTextEdit.EditValue = "root"
        Me.UserIdTextEdit.Location = New System.Drawing.Point(377, 66)
        Me.UserIdTextEdit.Name = "UserIdTextEdit"
        Me.UserIdTextEdit.Size = New System.Drawing.Size(103, 20)
        Me.UserIdTextEdit.TabIndex = 7
        '
        'LabelControl3
        '
        Me.LabelControl3.Location = New System.Drawing.Point(304, 69)
        Me.LabelControl3.Name = "LabelControl3"
        Me.LabelControl3.Size = New System.Drawing.Size(67, 13)
        Me.LabelControl3.TabIndex = 6
        Me.LabelControl3.Text = "MySql User ID"
        '
        'PasswordTextEdit
        '
        Me.PasswordTextEdit.EditValue = "password"
        Me.PasswordTextEdit.Location = New System.Drawing.Point(588, 66)
        Me.PasswordTextEdit.Name = "PasswordTextEdit"
        Me.PasswordTextEdit.Size = New System.Drawing.Size(101, 20)
        Me.PasswordTextEdit.TabIndex = 9
        '
        'LabelControl4
        '
        Me.LabelControl4.Location = New System.Drawing.Point(505, 69)
        Me.LabelControl4.Name = "LabelControl4"
        Me.LabelControl4.Size = New System.Drawing.Size(77, 13)
        Me.LabelControl4.TabIndex = 8
        Me.LabelControl4.Text = "MySql Password"
        '
        'MySqlConnectButton
        '
        Me.MySqlConnectButton.Location = New System.Drawing.Point(13, 301)
        Me.MySqlConnectButton.Name = "MySqlConnectButton"
        Me.MySqlConnectButton.Size = New System.Drawing.Size(118, 23)
        Me.MySqlConnectButton.TabIndex = 19
        Me.MySqlConnectButton.Text = "Connect to MySQL"
        '
        'TablesCheckedListBox
        '
        Me.TablesCheckedListBox.HotTrackItems = True
        Me.TablesCheckedListBox.Location = New System.Drawing.Point(151, 301)
        Me.TablesCheckedListBox.Name = "TablesCheckedListBox"
        Me.TablesCheckedListBox.Size = New System.Drawing.Size(169, 322)
        Me.TablesCheckedListBox.TabIndex = 20
        '
        'AccessButtonEdit
        '
        Me.AccessButtonEdit.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.AccessButtonEdit.EditValue = ""
        Me.AccessButtonEdit.Location = New System.Drawing.Point(150, 12)
        Me.AccessButtonEdit.Name = "AccessButtonEdit"
        EditorButtonImageOptions1.Location = DevExpress.XtraEditors.ImageLocation.MiddleLeft
        Me.AccessButtonEdit.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Glyph, "Browse", -1, True, True, False, EditorButtonImageOptions1, New DevExpress.Utils.KeyShortcut(System.Windows.Forms.Keys.None), SerializableAppearanceObject1, SerializableAppearanceObject2, SerializableAppearanceObject3, SerializableAppearanceObject4, "", Nothing, Nothing, DevExpress.Utils.ToolTipAnchor.[Default])})
        Me.AccessButtonEdit.Size = New System.Drawing.Size(539, 22)
        Me.AccessButtonEdit.TabIndex = 1
        '
        'CopyTablesButton
        '
        Me.CopyTablesButton.Location = New System.Drawing.Point(151, 270)
        Me.CopyTablesButton.Name = "CopyTablesButton"
        Me.CopyTablesButton.Size = New System.Drawing.Size(169, 23)
        Me.CopyTablesButton.TabIndex = 21
        Me.CopyTablesButton.Text = "Copy Selected Tables to MySql"
        '
        'SuccessLabel
        '
        Me.SuccessLabel.Appearance.BackColor = System.Drawing.Color.Lime
        Me.SuccessLabel.Appearance.Font = New System.Drawing.Font("Tahoma", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SuccessLabel.Appearance.Options.UseBackColor = True
        Me.SuccessLabel.Appearance.Options.UseFont = True
        Me.SuccessLabel.Location = New System.Drawing.Point(528, 385)
        Me.SuccessLabel.Name = "SuccessLabel"
        Me.SuccessLabel.Size = New System.Drawing.Size(89, 23)
        Me.SuccessLabel.TabIndex = 22
        Me.SuccessLabel.Text = "Success!!"
        Me.SuccessLabel.Visible = False
        '
        'CopyQueriesButton
        '
        Me.CopyQueriesButton.Location = New System.Drawing.Point(342, 270)
        Me.CopyQueriesButton.Name = "CopyQueriesButton"
        Me.CopyQueriesButton.Size = New System.Drawing.Size(169, 25)
        Me.CopyQueriesButton.TabIndex = 24
        Me.CopyQueriesButton.Text = "Copy Selected Queries to MySql"
        '
        'QueriesCheckedListBox
        '
        Me.QueriesCheckedListBox.HotTrackItems = True
        Me.QueriesCheckedListBox.Location = New System.Drawing.Point(342, 301)
        Me.QueriesCheckedListBox.Name = "QueriesCheckedListBox"
        Me.QueriesCheckedListBox.Size = New System.Drawing.Size(169, 322)
        Me.QueriesCheckedListBox.TabIndex = 23
        '
        'Select3Point0Button
        '
        Me.Select3Point0Button.Location = New System.Drawing.Point(12, 330)
        Me.Select3Point0Button.Name = "Select3Point0Button"
        Me.Select3Point0Button.Size = New System.Drawing.Size(118, 23)
        Me.Select3Point0Button.TabIndex = 25
        Me.Select3Point0Button.Text = "Select 3.0 Items"
        '
        'Non3Point0Button
        '
        Me.Non3Point0Button.Location = New System.Drawing.Point(12, 359)
        Me.Non3Point0Button.Name = "Non3Point0Button"
        Me.Non3Point0Button.Size = New System.Drawing.Size(118, 23)
        Me.Non3Point0Button.TabIndex = 26
        Me.Non3Point0Button.Text = "Get Non 3.0 Tables"
        '
        'Non3Point0ListBoxControl
        '
        Me.Non3Point0ListBoxControl.Items.AddRange(New Object() {"AutoInc", "CloudDels", "CloudEdits", "CompanyRoles2", "CompanyRolesNL", "CompOls00000", "CompOls00001", "CompOls00002", "CompOls00006", "CompOls00008", "CompOls00009", "CompOls00010", "CompOls00011", "CompOls00012", "CompOls00013", "CompOls00014", "CompOls00016", "CompOls00017", "CompOls00018", "CompOls00019", "CompOls00020", "CompOls00021", "CompOls00022", "CompOls00023", "CompOls00024", "Contacts2", "ContactsNL", "ContOls00000", "ContOls00001", "ContOls00002", "ContOls00006", "ContOls00008", "ContOls00009", "ContOls00010", "ContOls00011", "ContOls00012", "ContOls00013", "ContOls00014", "ContOls00016", "ContOls00017", "ContOls00018", "ContOls00019", "ContOls00020", "ContOls00021", "ContOls00022", "ContOls00023", "ContOls00024", "FfHistory2", "FixedFeeNL", "FixedFeeProbs", "NonLinked"})
        Me.Non3Point0ListBoxControl.Location = New System.Drawing.Point(12, 603)
        Me.Non3Point0ListBoxControl.Name = "Non3Point0ListBoxControl"
        Me.Non3Point0ListBoxControl.Size = New System.Drawing.Size(117, 20)
        Me.Non3Point0ListBoxControl.TabIndex = 27
        Me.Non3Point0ListBoxControl.TabStop = False
        Me.Non3Point0ListBoxControl.Visible = False
        '
        'CopyAllButton
        '
        Me.CopyAllButton.Location = New System.Drawing.Point(517, 270)
        Me.CopyAllButton.Name = "CopyAllButton"
        Me.CopyAllButton.Size = New System.Drawing.Size(173, 25)
        Me.CopyAllButton.TabIndex = 28
        Me.CopyAllButton.Text = "Copy All Selected items to MySql"
        '
        'CreateDatabaseButton
        '
        Me.CreateDatabaseButton.Location = New System.Drawing.Point(13, 388)
        Me.CreateDatabaseButton.Name = "CreateDatabaseButton"
        Me.CreateDatabaseButton.Size = New System.Drawing.Size(118, 23)
        Me.CreateDatabaseButton.TabIndex = 29
        Me.CreateDatabaseButton.Text = "Create Database"
        '
        'CreateTestUserButton
        '
        Me.CreateTestUserButton.Location = New System.Drawing.Point(13, 417)
        Me.CreateTestUserButton.Name = "CreateTestUserButton"
        Me.CreateTestUserButton.Size = New System.Drawing.Size(118, 23)
        Me.CreateTestUserButton.TabIndex = 30
        Me.CreateTestUserButton.Text = "Create Test User"
        '
        'ChangePasswordButton
        '
        Me.ChangePasswordButton.Location = New System.Drawing.Point(13, 446)
        Me.ChangePasswordButton.Name = "ChangePasswordButton"
        Me.ChangePasswordButton.Size = New System.Drawing.Size(118, 23)
        Me.ChangePasswordButton.TabIndex = 31
        Me.ChangePasswordButton.Text = "Change User Password"
        '
        'WriteConfigButton
        '
        Me.WriteConfigButton.Location = New System.Drawing.Point(12, 475)
        Me.WriteConfigButton.Name = "WriteConfigButton"
        Me.WriteConfigButton.Size = New System.Drawing.Size(118, 23)
        Me.WriteConfigButton.TabIndex = 32
        Me.WriteConfigButton.Text = "Write Config Info"
        '
        'ReadConfigButton
        '
        Me.ReadConfigButton.Location = New System.Drawing.Point(12, 504)
        Me.ReadConfigButton.Name = "ReadConfigButton"
        Me.ReadConfigButton.Size = New System.Drawing.Size(118, 23)
        Me.ReadConfigButton.TabIndex = 33
        Me.ReadConfigButton.Text = "Read Config Info"
        '
        'GetTableSchemaButton
        '
        Me.GetTableSchemaButton.Location = New System.Drawing.Point(12, 533)
        Me.GetTableSchemaButton.Name = "GetTableSchemaButton"
        Me.GetTableSchemaButton.Size = New System.Drawing.Size(118, 23)
        Me.GetTableSchemaButton.TabIndex = 34
        Me.GetTableSchemaButton.Text = "Get Table Schema"
        '
        'LabelControl1
        '
        Me.LabelControl1.Location = New System.Drawing.Point(12, 144)
        Me.LabelControl1.Name = "LabelControl1"
        Me.LabelControl1.Size = New System.Drawing.Size(106, 13)
        Me.LabelControl1.TabIndex = 35
        Me.LabelControl1.Text = "MySql PCDB Database"
        '
        'PersistSecurityInfoCheckBox
        '
        Me.PersistSecurityInfoCheckBox.AutoSize = True
        Me.PersistSecurityInfoCheckBox.Checked = True
        Me.PersistSecurityInfoCheckBox.CheckState = System.Windows.Forms.CheckState.Checked
        Me.PersistSecurityInfoCheckBox.Location = New System.Drawing.Point(150, 93)
        Me.PersistSecurityInfoCheckBox.Name = "PersistSecurityInfoCheckBox"
        Me.PersistSecurityInfoCheckBox.Size = New System.Drawing.Size(144, 17)
        Me.PersistSecurityInfoCheckBox.TabIndex = 36
        Me.PersistSecurityInfoCheckBox.Text = "Use Persist Security Info"
        Me.PersistSecurityInfoCheckBox.UseVisualStyleBackColor = True
        '
        'NoneRadioButton
        '
        Me.NoneRadioButton.AutoSize = True
        Me.NoneRadioButton.Location = New System.Drawing.Point(150, 116)
        Me.NoneRadioButton.Name = "NoneRadioButton"
        Me.NoneRadioButton.Size = New System.Drawing.Size(50, 17)
        Me.NoneRadioButton.TabIndex = 37
        Me.NoneRadioButton.Text = "None"
        Me.NoneRadioButton.UseVisualStyleBackColor = True
        '
        'PreferredRadioButton
        '
        Me.PreferredRadioButton.AutoSize = True
        Me.PreferredRadioButton.Checked = True
        Me.PreferredRadioButton.Location = New System.Drawing.Point(242, 116)
        Me.PreferredRadioButton.Name = "PreferredRadioButton"
        Me.PreferredRadioButton.Size = New System.Drawing.Size(71, 17)
        Me.PreferredRadioButton.TabIndex = 38
        Me.PreferredRadioButton.TabStop = True
        Me.PreferredRadioButton.Text = "Preferred"
        Me.PreferredRadioButton.UseVisualStyleBackColor = True
        '
        'RequiredRadioButton
        '
        Me.RequiredRadioButton.AutoSize = True
        Me.RequiredRadioButton.Location = New System.Drawing.Point(341, 116)
        Me.RequiredRadioButton.Name = "RequiredRadioButton"
        Me.RequiredRadioButton.Size = New System.Drawing.Size(68, 17)
        Me.RequiredRadioButton.TabIndex = 39
        Me.RequiredRadioButton.Text = "Required"
        Me.RequiredRadioButton.UseVisualStyleBackColor = True
        '
        'VerifyCARadioButton
        '
        Me.VerifyCARadioButton.AutoSize = True
        Me.VerifyCARadioButton.Location = New System.Drawing.Point(440, 116)
        Me.VerifyCARadioButton.Name = "VerifyCARadioButton"
        Me.VerifyCARadioButton.Size = New System.Drawing.Size(70, 17)
        Me.VerifyCARadioButton.TabIndex = 40
        Me.VerifyCARadioButton.Text = "Verify CA"
        Me.VerifyCARadioButton.UseVisualStyleBackColor = True
        '
        'VerifyFullRadioButton
        '
        Me.VerifyFullRadioButton.AutoSize = True
        Me.VerifyFullRadioButton.Location = New System.Drawing.Point(542, 116)
        Me.VerifyFullRadioButton.Name = "VerifyFullRadioButton"
        Me.VerifyFullRadioButton.Size = New System.Drawing.Size(72, 17)
        Me.VerifyFullRadioButton.TabIndex = 41
        Me.VerifyFullRadioButton.Text = "Verify Full"
        Me.VerifyFullRadioButton.UseVisualStyleBackColor = True
        '
        'LabelControl5
        '
        Me.LabelControl5.Location = New System.Drawing.Point(12, 116)
        Me.LabelControl5.Name = "LabelControl5"
        Me.LabelControl5.Size = New System.Drawing.Size(46, 13)
        Me.LabelControl5.TabIndex = 42
        Me.LabelControl5.Text = "SSL Mode"
        '
        'OtherButton
        '
        Me.OtherButton.Location = New System.Drawing.Point(11, 562)
        Me.OtherButton.Name = "OtherButton"
        Me.OtherButton.Size = New System.Drawing.Size(118, 23)
        Me.OtherButton.TabIndex = 43
        Me.OtherButton.Text = "Other Test"
        '
        'Timer1
        '
        Me.Timer1.Interval = 1000
        '
        'RunningTimeLabelControl
        '
        Me.RunningTimeLabelControl.Location = New System.Drawing.Point(447, 247)
        Me.RunningTimeLabelControl.Name = "RunningTimeLabelControl"
        Me.RunningTimeLabelControl.Size = New System.Drawing.Size(64, 13)
        Me.RunningTimeLabelControl.TabIndex = 45
        Me.RunningTimeLabelControl.Text = "Running Time"
        Me.RunningTimeLabelControl.Visible = False
        '
        'RunningTimeTextEdit
        '
        Me.RunningTimeTextEdit.Location = New System.Drawing.Point(517, 243)
        Me.RunningTimeTextEdit.Name = "RunningTimeTextEdit"
        Me.RunningTimeTextEdit.Properties.Appearance.BackColor = System.Drawing.Color.White
        Me.RunningTimeTextEdit.Properties.Appearance.Options.UseBackColor = True
        Me.RunningTimeTextEdit.Size = New System.Drawing.Size(137, 20)
        Me.RunningTimeTextEdit.TabIndex = 46
        Me.RunningTimeTextEdit.Visible = False
        '
        'SimpleButton1
        '
        Me.SimpleButton1.Location = New System.Drawing.Point(517, 301)
        Me.SimpleButton1.Name = "SimpleButton1"
        Me.SimpleButton1.Size = New System.Drawing.Size(118, 23)
        Me.SimpleButton1.TabIndex = 47
        Me.SimpleButton1.Text = "Start Timer"
        '
        'SimpleButton2
        '
        Me.SimpleButton2.Location = New System.Drawing.Point(517, 330)
        Me.SimpleButton2.Name = "SimpleButton2"
        Me.SimpleButton2.Size = New System.Drawing.Size(118, 23)
        Me.SimpleButton2.TabIndex = 48
        Me.SimpleButton2.Text = "Stop Timer"
        '
        'DoCancelButton
        '
        Me.DoCancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.DoCancelButton.Location = New System.Drawing.Point(517, 598)
        Me.DoCancelButton.Name = "DoCancelButton"
        Me.DoCancelButton.Size = New System.Drawing.Size(173, 25)
        Me.DoCancelButton.TabIndex = 49
        Me.DoCancelButton.Text = "Cancel"
        '
        'BackgroundWorker1
        '
        Me.BackgroundWorker1.WorkerReportsProgress = True
        Me.BackgroundWorker1.WorkerSupportsCancellation = True
        '
        'CanceledLabel
        '
        Me.CanceledLabel.Appearance.BackColor = System.Drawing.Color.Yellow
        Me.CanceledLabel.Appearance.Font = New System.Drawing.Font("Tahoma", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CanceledLabel.Appearance.ForeColor = System.Drawing.Color.Black
        Me.CanceledLabel.Appearance.Options.UseBackColor = True
        Me.CanceledLabel.Appearance.Options.UseFont = True
        Me.CanceledLabel.Appearance.Options.UseForeColor = True
        Me.CanceledLabel.Location = New System.Drawing.Point(528, 417)
        Me.CanceledLabel.Name = "CanceledLabel"
        Me.CanceledLabel.Size = New System.Drawing.Size(100, 23)
        Me.CanceledLabel.TabIndex = 50
        Me.CanceledLabel.Text = "Canceled!!"
        Me.CanceledLabel.Visible = False
        '
        'CopiedTablesLabel
        '
        Me.CopiedTablesLabel.Appearance.BackColor = System.Drawing.Color.Aqua
        Me.CopiedTablesLabel.Appearance.Font = New System.Drawing.Font("Tahoma", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CopiedTablesLabel.Appearance.Options.UseBackColor = True
        Me.CopiedTablesLabel.Appearance.Options.UseFont = True
        Me.CopiedTablesLabel.Location = New System.Drawing.Point(528, 446)
        Me.CopiedTablesLabel.Name = "CopiedTablesLabel"
        Me.CopiedTablesLabel.Size = New System.Drawing.Size(148, 23)
        Me.CopiedTablesLabel.TabIndex = 51
        Me.CopiedTablesLabel.Text = "Copied Tables!!"
        Me.CopiedTablesLabel.Visible = False
        '
        'CopiedQueriesLabel
        '
        Me.CopiedQueriesLabel.Appearance.BackColor = System.Drawing.Color.Fuchsia
        Me.CopiedQueriesLabel.Appearance.Font = New System.Drawing.Font("Tahoma", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CopiedQueriesLabel.Appearance.Options.UseBackColor = True
        Me.CopiedQueriesLabel.Appearance.Options.UseFont = True
        Me.CopiedQueriesLabel.Location = New System.Drawing.Point(528, 475)
        Me.CopiedQueriesLabel.Name = "CopiedQueriesLabel"
        Me.CopiedQueriesLabel.Size = New System.Drawing.Size(159, 23)
        Me.CopiedQueriesLabel.TabIndex = 52
        Me.CopiedQueriesLabel.Text = "Copied Queries!!"
        Me.CopiedQueriesLabel.Visible = False
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.DoCancelButton
        Me.ClientSize = New System.Drawing.Size(701, 638)
        Me.Controls.Add(Me.CopiedQueriesLabel)
        Me.Controls.Add(Me.CopiedTablesLabel)
        Me.Controls.Add(Me.CanceledLabel)
        Me.Controls.Add(Me.DoCancelButton)
        Me.Controls.Add(Me.SimpleButton2)
        Me.Controls.Add(Me.SimpleButton1)
        Me.Controls.Add(Me.RunningTimeTextEdit)
        Me.Controls.Add(Me.RunningTimeLabelControl)
        Me.Controls.Add(Me.OtherButton)
        Me.Controls.Add(Me.LabelControl5)
        Me.Controls.Add(Me.VerifyFullRadioButton)
        Me.Controls.Add(Me.VerifyCARadioButton)
        Me.Controls.Add(Me.RequiredRadioButton)
        Me.Controls.Add(Me.PreferredRadioButton)
        Me.Controls.Add(Me.NoneRadioButton)
        Me.Controls.Add(Me.PersistSecurityInfoCheckBox)
        Me.Controls.Add(Me.LabelControl1)
        Me.Controls.Add(Me.GetTableSchemaButton)
        Me.Controls.Add(Me.ReadConfigButton)
        Me.Controls.Add(Me.WriteConfigButton)
        Me.Controls.Add(Me.ChangePasswordButton)
        Me.Controls.Add(Me.CreateTestUserButton)
        Me.Controls.Add(Me.CreateDatabaseButton)
        Me.Controls.Add(Me.CopyAllButton)
        Me.Controls.Add(Me.Non3Point0ListBoxControl)
        Me.Controls.Add(Me.Non3Point0Button)
        Me.Controls.Add(Me.Select3Point0Button)
        Me.Controls.Add(Me.CopyQueriesButton)
        Me.Controls.Add(Me.QueriesCheckedListBox)
        Me.Controls.Add(Me.SuccessLabel)
        Me.Controls.Add(Me.CopyTablesButton)
        Me.Controls.Add(Me.TablesCheckedListBox)
        Me.Controls.Add(Me.MySqlConnectButton)
        Me.Controls.Add(Me.PasswordTextEdit)
        Me.Controls.Add(Me.LabelControl4)
        Me.Controls.Add(Me.UserIdTextEdit)
        Me.Controls.Add(Me.LabelControl3)
        Me.Controls.Add(Me.MySqlPcdbNameTextEdit)
        Me.Controls.Add(Me.LabelControl2)
        Me.Controls.Add(Me.MySqlServerTextEdit)
        Me.Controls.Add(Me.MyServerLabelControl)
        Me.Controls.Add(Me.StratTimeLabelControl)
        Me.Controls.Add(Me.StartTimeEdit)
        Me.Controls.Add(Me.TotalProgressLabelControl)
        Me.Controls.Add(Me.TotalProgressBarControl)
        Me.Controls.Add(Me.ActionProgressBarControl)
        Me.Controls.Add(Me.ActionProgressLabelControl)
        Me.Controls.Add(Me.ActionTextEdit)
        Me.Controls.Add(Me.ActionLabelControl)
        Me.Controls.Add(Me.GetTablesButton)
        Me.Controls.Add(Me.AccessButtonEdit)
        Me.Controls.Add(Me.CurrentLabelControl)
        Me.IconOptions.Icon = CType(resources.GetObject("Form1.IconOptions.Icon"), System.Drawing.Icon)
        Me.Name = "Form1"
        Me.Text = "PCDB: Access to SQL"
        CType(Me.StartTimeEdit.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TotalProgressBarControl.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ActionProgressBarControl.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ActionTextEdit.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MySqlServerTextEdit.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MySqlPcdbNameTextEdit.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UserIdTextEdit.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PasswordTextEdit.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TablesCheckedListBox, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AccessButtonEdit.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.QueriesCheckedListBox, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Non3Point0ListBoxControl, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RunningTimeTextEdit.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents CurrentLabelControl As DevExpress.XtraEditors.LabelControl
    Friend WithEvents XtraOpenFileDialog1 As DevExpress.XtraEditors.XtraOpenFileDialog
    Friend WithEvents GetTablesButton As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents AccessOleDbConnection As OleDb.OleDbConnection
    Friend WithEvents StratTimeLabelControl As DevExpress.XtraEditors.LabelControl
    Friend WithEvents StartTimeEdit As DevExpress.XtraEditors.TimeEdit
    Friend WithEvents TotalProgressLabelControl As DevExpress.XtraEditors.LabelControl
    Friend WithEvents TotalProgressBarControl As DevExpress.XtraEditors.ProgressBarControl
    Friend WithEvents ActionProgressBarControl As DevExpress.XtraEditors.ProgressBarControl
    Friend WithEvents ActionProgressLabelControl As DevExpress.XtraEditors.LabelControl
    Friend WithEvents ActionTextEdit As DevExpress.XtraEditors.TextEdit
    Friend WithEvents ActionLabelControl As DevExpress.XtraEditors.LabelControl
    Friend WithEvents MySqlServerTextEdit As DevExpress.XtraEditors.TextEdit
    Friend WithEvents MyServerLabelControl As DevExpress.XtraEditors.LabelControl
    Friend WithEvents MySqlPcdbNameTextEdit As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl2 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents UserIdTextEdit As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl3 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents PasswordTextEdit As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl4 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents MySqlConnectButton As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents TablesCheckedListBox As DevExpress.XtraEditors.CheckedListBoxControl
    Friend WithEvents AccessButtonEdit As DevExpress.XtraEditors.ButtonEdit
    Friend WithEvents CopyTablesButton As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents SuccessLabel As DevExpress.XtraEditors.LabelControl
    Friend WithEvents CopyQueriesButton As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents QueriesCheckedListBox As DevExpress.XtraEditors.CheckedListBoxControl
    Friend WithEvents Select3Point0Button As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents Non3Point0Button As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents Non3Point0ListBoxControl As DevExpress.XtraEditors.ListBoxControl
    Friend WithEvents CopyAllButton As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents CreateDatabaseButton As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents CreateTestUserButton As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents ChangePasswordButton As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents WriteConfigButton As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents ReadConfigButton As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents GetTableSchemaButton As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents LabelControl1 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents PersistSecurityInfoCheckBox As CheckBox
    Friend WithEvents NoneRadioButton As RadioButton
    Friend WithEvents PreferredRadioButton As RadioButton
    Friend WithEvents RequiredRadioButton As RadioButton
    Friend WithEvents VerifyCARadioButton As RadioButton
    Friend WithEvents VerifyFullRadioButton As RadioButton
    Friend WithEvents LabelControl5 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents OtherButton As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents Timer1 As Timer
    Friend WithEvents RunningTimeLabelControl As DevExpress.XtraEditors.LabelControl
    Friend WithEvents RunningTimeTextEdit As DevExpress.XtraEditors.TextEdit
    Friend WithEvents SimpleButton1 As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents SimpleButton2 As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents DoCancelButton As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents CanceledLabel As DevExpress.XtraEditors.LabelControl
    Friend WithEvents CopiedTablesLabel As DevExpress.XtraEditors.LabelControl
    Friend WithEvents CopiedQueriesLabel As DevExpress.XtraEditors.LabelControl

#End Region

End Class
