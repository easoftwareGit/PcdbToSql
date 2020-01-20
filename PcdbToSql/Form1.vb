Imports EaTools.DataTools
Imports EaTools.IO
Imports EaAccess.AccessSql
Imports EaMySql
Imports MySql.Data.MySqlClient
'Imports DevExpress.Data.Async.Helpers
'Imports DevExpress.DirectX.Common

Partial Public Class Form1

#Region " Class Constants and Variables "

#Region " Constants "

    Private Const cnCharMaxLen As String = "CHARACTER_MAXIMUM_LENGTH"
    Private Const cnColumnName As String = "COLUMN_NAME"
    Private Const cnDataType As String = "DATA_TYPE"
    Private Const cnIndexName As String = "INDEX_NAME"
    Private Const cnIsNullable As String = "IS_NULLABLE"
    Private Const cnOrdinalPos As String = "ORDINAL_POSITION"
    Private Const cnPrimaryKey As String = "PRIMARY_KEY"
    Private Const cnProcedureDef As String = "PROCEDURE_DEFINITION"
    Private Const cnProcedureName As String = "PROCEDURE_NAME"
    Private Const cnQueryDef As String = "QUERY_DEFINITION"
    Private Const cnQueryOrder As String = "QUERY_ORDER"
    Private Const cnQueryName As String = "QUERY_NAME"
    Private Const cnTableName As String = "TABLE_NAME"
    Private Const cnViewDef As String = "VIEW_DEFINITION"
    Private Const cnViewOrProc As String = "VIEW_OR_PROC"

    Private Const qtView As String = "VIEW"
    Private Const qtProc As String = "PROC"

    'Private Const cnColHasDefault As String = "COLUMN_HASDEFAULT"
    'Private Const cnColFlags As String = "COLUMN_FLAGS"
    'Private Const cnCharOctLen As String = "CHARACTER_OCTET_LENGTH"
    'Private Const cnDateTimePrecision As String = "DATETIME_PRECISION"
    'Private Const cnNumPrecision As String = "NUMERIC_PRECISION"

    Private Const etCreateMySqlDataAdapter As String = "Error Creating MySqlDataAdapter"
    Private Const etCreateOleDbDataAdapter As String = "Error Creating OleDbDataAdapter"
    Private Const etCreatingTable As String = "Error Creating Table"
    Private Const etCreatingQuery As String = "Error Creating Query"
    Private Const etSortingQueries As String = "Error Sorting Queries"
    Private Const etRowCount As String = "Row Count Error"
    Private Const etVerifyTableData As String = "Error Verifying Table Data"
    Private Const etVerifyQuery As String = "Error Verifying Query"
    'Private Const SqlConnStrFormat As String = "server=localhost;user id=root;password=password;persistsecurityinfo=True;database=pcdb_test;sslmode=None"
    'Private Const SqlConnStrFormat As String = "server={0};user id={1};password={2};persistsecurityinfo=True;database={3};sslmode=None"

    Private Const TimeStampColName As String = "UpdatedWhen"
    'Private Const TimeStampValue As String = "CURRENT_TIMESTAMP"

    Private Const TestUserName As String = "JohnDoe"
    Private Const TestPassword As String = "qwerty123"
    Private Const TestNewPassword As String = "newpassword"

#End Region

#Region " Enums "

    Enum CopyTasks
        Tables = 0
        Queries = 1
        All = 2
    End Enum

#End Region

#Region " Private Vars "

    Private AccessFolder As String
    Private AccessDatabaseName As String
    Private AccessFullDatabaseName As String
    Private MySqlPcdbConn As MySqlConnection

    Private AccessColumns As DataTable
    Private AccessIndexes As DataTable
    Private AccessProcs As DataTable
    Private AccessQueries As DataTable
    Private AccessTables As DataTable
    Private AccessViews As DataTable
    Private MySqlQueries As DataTable

    Private startTime As DateTime
    Private stCount As Integer = 1

    Private totalSteps As Integer

#End Region

#End Region

#Region " New "

    ' connect to access pcdb
    ' connect to mySql pcdb
    ' for each table in access pcdb
    '   get table struct
    '   create table in mySql
    '     use autoInc for id's
    '     include timestamp column
    '     include user id column
    '   select all data from access data table
    '   for each access data table row
    '     get new row for mySql table
    '     populate mySql data table row
    '     add row to data table
    '   update mySql data table

    'CREATE TABLE `pcdb_test`.`companies` (
    '  `CompId` INT NOT NULL AUTO_INCREMENT,
    '  `Name` VARCHAR(75) NULL,
    '  `FormalName` VARCHAR(75) NULL,
    '  `Address1` VARCHAR(60) NULL,
    '  `Address2` VARCHAR(60) NULL,
    '  `City` VARCHAR(50) NULL,
    '  `StateRegion` VARCHAR(30) NULL,
    '  `PostalCode` VARCHAR(30) NULL,
    '  `Country` VARCHAR(50) NULL,
    '  `Phone` VARCHAR(30) NULL,
    '  `Fax` VARCHAR(30) NULL,
    '  `WpSoftware` VARCHAR(50) NULL,
    '  `CADSoftware` VARCHAR(50) NULL,
    '  `WebPage` VARCHAR(100) NULL,
    '  `EMail` VARCHAR(50) NULL,
    '  `Prospect` VARCHAR(10) NULL,
    '  `Customer` TINYINT NULL,
    '  `FocusLocal` TINYINT NULL,
    '  `FocusNational` TINYINT NULL,
    '  `Comments` TEXT NULL,
    '  `PDxo35` TINYINT NULL,
    '  `Active` TINYINT NULL,
    '  `DataLocInt` INT NULL,
    '  `UserId` INT NULL,
    '  `LastUpdate` TIMESTAMP NULL DEFAULT CURRENT_TIMESTAMP,
    'PRIMARY KEY (`CompId`),
    'UNIQUE INDEX `CompId_UNIQUE` (`CompId` ASC));


    Shared Sub New()
        DevExpress.UserSkins.BonusSkins.Register()
    End Sub
    Public Sub New()
        InitializeComponent()
    End Sub

#End Region

#Region " Form Events "

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        AccessButtonEdit.Text = "C:\Projects 2017\PcdbData\Ewcg2000.mdb"
        Dim fInfo As IO.FileInfo = My.Computer.FileSystem.GetFileInfo(AccessButtonEdit.Text)                ' get file info for file
        If Not My.Computer.FileSystem.DirectoryExists(fInfo.DirectoryName) Then                             ' if file does not exist
            MessageBox.Show(String.Format("Folder {1} not found for {0}.", CurrentLabelControl.Text, fInfo.DirectoryName),
                                    "Folder Not Found", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            AccessButtonEdit.Text = String.Empty                                                            ' clear values
            AccessFolder = String.Empty
            AccessDatabaseName = String.Empty
            AccessFullDatabaseName = String.Empty

            Return
        End If
        ' if got here, the access PCDB database entry is valid, set values
        AccessFolder = fInfo.DirectoryName                                                                  ' set the current database folder name
        AccessDatabaseName = fInfo.Name                                                                     ' set the current database file name 
        AccessFullDatabaseName = fInfo.FullName                                                             ' set the current database full file name

        Dim Err As Integer = MakeAccessConnection(AccessFolder, AccessDatabaseName, AccessOleDbConnection)  ' make connection to access database
        If Err <> Pcm.NoErrors Then                                                                         ' if got an error
            Return                                                                                          ' return error
        End If

        Err = GetAccessTablesAndQueries()
        If Err <> Pcm.NoErrors Then                                                                         ' if got an error
            Return                                                                                          ' return error
        End If

        PopulateTablesCheckedListBox()
        PopulateQueriesCheckedListBox()
        ShowActions(False)
    End Sub

#End Region

#Region " BackgroundWorker "

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork

        Dim worker As System.ComponentModel.BackgroundWorker = CType(sender, System.ComponentModel.BackgroundWorker) ' get BackgroundWorker that raised event
        Dim task As CopyTasks = CType(e.Argument, CopyTasks)
        StartStopTimer(True)
        Select Case task
            Case CopyTasks.Tables
                e.Result = CopyTablesToMySql(worker, e)
            Case CopyTasks.Queries
                e.Result = CopyQueriesToMySql(worker, e)
            Case CopyTasks.All
                e.Result = CopyAllToMySql(worker, e)
        End Select
        StartStopTimer(False)
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged

        Try
            Dim p As PcdbToMySqlProgress = CType(e.UserState, PcdbToMySqlProgress)
            p.ActionMessageUpdate()
            If p.ResetActionMax Then
                p.ActionPbParams.Maximum = p.ActionMax
                p.ResetActionMax = False
            End If
            p.ActionUpdate()
            p.TimeUpdate()
            If p.ResetTotalMax Then
                p.TotalPbParams.Maximum = p.TotalMax
                p.ResetTotalMax = False
            End If
            p.TotalUpdate()
            If p.CopiedTables Then
                p.ShowCopiedTables()
            End If
            If p.CopiedQueries Then
                p.ShowCopiedQueries()
            End If
        Catch ex As Exception
            Return
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted

        If e.Error IsNot Nothing Then
            MessageBox.Show(e.Error.Message)

        ElseIf e.Cancelled Then

            ' Next, handle the case where the user canceled the operation.
            ' Note that due to a race condition in the DoWork event handler,
            ' the Cancelled flag may Not have been set, even though CancelAsync was called.

            CanceledLabel.Visible = True

        Else                                                ' Finally, handle the case where the operation succeeded.
            SuccessLabel.Visible = True
        End If
        StartStopTimer(False)
    End Sub

    Private Sub CancelBackgroundWorker()

        Timer1.Stop()
        BackgroundWorker1.CancelAsync()
    End Sub

    Private Sub ShowProgress(worker As System.ComponentModel.BackgroundWorker,
                             progress As PcdbToMySqlProgress)

        ' shows the progress of the task being done by the background worker 
        '
        ' vars passed:
        '   worker - background worker
        '   progress - progress values 

        worker.ReportProgress(1, progress)                  ' send message to background worker progressed changed event
    End Sub

    Private Sub ShowProgress(worker As System.ComponentModel.BackgroundWorker,
                             progress As PcdbToMySqlProgress,
                             actionPosition As Integer)

        ' shows the progress of the task being done by the background worker 
        '
        ' vars passed:
        '   worker - background worker
        '   progress - progress values 
        '   actionPosition - new position for the current action progress bar

        progress.ActionPosition = actionPosition            ' set action progress bar position
        worker.ReportProgress(1, progress)                  ' send message to background worker progressed changed event
    End Sub

    Private Sub ShowProgress(worker As System.ComponentModel.BackgroundWorker,
                             progress As PcdbToMySqlProgress,
                             actionMessage As String,
                             Optional actionPosition As Integer = 0)

        ' shows the progress of the task being done by the background worker 
        '
        ' vars passed:
        '   worker - background worker
        '   progress - progress values 
        '   action message - new action message to be displayed
        '   actionPosition - new position for the current action progress bar (pass -1 or less to ignore)   

        progress.ActionMessage = actionMessage              ' set new action message
        If actionPosition > -1 Then                         ' if want to set a different progress bar position
            progress.ActionPosition = actionPosition        ' set action progress bar position
        End If
        worker.ReportProgress(1, progress)                  ' send message to background worker progressed changed event
    End Sub

#End Region

#Region " Timer "

    Private Sub StartStopTimer(DoStart As Boolean)

        If DoStart Then
            startTime = DateAndTime.Now()
            Timer1.Enabled = True
            Timer1.Start()
        Else
            Timer1.Stop()
        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        Dim elapsedTime As TimeSpan = DateTime.Now - startTime
        elapsedTime = EaTools.Tools.TimeSpanRound(elapsedTime, 0)
        RunningTimeTextEdit.Text = String.Format("{0:00}:{1:00}:{2:00}", elapsedTime.Hours, elapsedTime.Minutes, elapsedTime.Seconds)
    End Sub

#End Region

#Region " Show Actions "

    Private Sub ActionPosition(Position As Integer)

        ' sets the position of the action progress bar
        '
        ' vars passed:
        '   Position - position of the action progress bar

        ActionProgressBarControl.Position = Position
        ActionProgressBarControl.Update()
    End Sub

    Private Sub ActionProgressBarStep()

        ' move the action progress bar 1 step

        ActionProgressBarControl.PerformStep()
        ActionProgressBarControl.Update()
    End Sub

    Private Sub SetAction(ActionMessage As String,
                          Optional Maximum As Integer = 100)

        ' sets the action text and resets the action progress bar position to 0
        '
        ' vars passed:
        '   ActionMessage - action message to display

        ActionTextEdit.Text = ActionMessage
        ActionTextEdit.Update()
        ActionProgressBarControl.Properties.Maximum = Maximum
        ActionPosition(0)
    End Sub

    Private Sub ShowActions(DoShow As Boolean)

        ' shows the action

        StartTimeEdit.EditValue = Now()
        ActionLabelControl.Visible = DoShow
        ActionTextEdit.Visible = DoShow
        ActionProgressLabelControl.Visible = DoShow
        ActionProgressBarControl.Visible = DoShow
        TotalProgressLabelControl.Visible = DoShow
        TotalProgressBarControl.Visible = DoShow
        StratTimeLabelControl.Visible = DoShow
        StartTimeEdit.Visible = DoShow
        RunningTimeLabelControl.Visible = DoShow
        RunningTimeTextEdit.Visible = DoShow

        DoCancelButton.Enabled = Not DoShow

        If DoShow Then
            SuccessLabel.Visible = False
            CanceledLabel.Visible = False
            CopiedTablesLabel.Visible = False
            CopiedQueriesLabel.Visible = False
        End If

        Refresh()
    End Sub

    Private Sub TotalProgressBarStep()

        ' move the total progress bar 1 step

        TotalProgressBarControl.PerformStep()
        TotalProgressBarControl.Update()
    End Sub

#End Region

#Region " ConvertColNames "

    Private Function ConvertAccessColumnType(ByVal acColumnType As EaAccess.AccessColumn.ColumnTypes) As MySqlDataType

        Try
            Select Case acColumnType
                Case EaAccess.AccessColumn.ColumnTypes.ctBoolean
                    Return MySqlDataType.dtTinyInt
                Case EaAccess.AccessColumn.ColumnTypes.ctCurrency
                    Return MySqlDataType.dtDecimal
                Case EaAccess.AccessColumn.ColumnTypes.ctDate
                    Return MySqlDataType.dtDate
                Case EaAccess.AccessColumn.ColumnTypes.ctDouble
                    Return MySqlDataType.dtDouble
                Case EaAccess.AccessColumn.ColumnTypes.ctInteger
                    Return MySqlDataType.dtInteger
                Case EaAccess.AccessColumn.ColumnTypes.ctMemo
                    Return MySqlDataType.dtText
                Case EaAccess.AccessColumn.ColumnTypes.ctNone
                    Return MySqlDataType.dtNone
                Case EaAccess.AccessColumn.ColumnTypes.ctSmallInt
                    Return MySqlDataType.dtSmallInt
                Case EaAccess.AccessColumn.ColumnTypes.ctText
                    Return MySqlDataType.dtVarChar
                Case Else
                    Return MySqlDataType.dtNone
            End Select
        Catch ex As Exception
            Return MySqlDataType.dtNone
        End Try
    End Function

    Private Function ConvertCodeCol(TableName As String,
                                    ColumnName As String) As String

        ' gets the converted column name when the column name is "CODE"
        '   
        ' vars passed:
        '   TableName - name of table with column
        '   ColumnName - column name
        '   
        ' returns:
        '   ColumnName - if column name <> "CODE", then just return the original column name
        '   "...Code" - converted column name if column name = "CODE"        

        If ColumnName.ToUpper <> "CODE" Then
            Return ColumnName
        End If
        Select Case TableName
            Case Pcm.CountriesTableName
                Return "CodeName"
            Case Else
                Return ColumnName
        End Select
    End Function

    Private Function ConvertNameCol(TableName As String,
                                    ColumnName As String) As String

        ' gets the converted column name when the column name is "NAME"
        '   
        ' vars passed:
        '   TableName - name of table with column
        '   ColumnName - column name
        '   
        ' returns:
        '   ColumnName - if column name <> "NAME", then just return the original column name
        '   "...Name" - converted column name if column name = "NAME"        

        If ColumnName.ToUpper <> "NAME" Then
            Return ColumnName
        End If
        Select Case TableName
            Case Pcm.AffilTableName
                Return "AffilName"
            Case Pcm.CitiesTableName
                Return "CityName"
            Case Pcm.CodesTableName
                Return "CodeName"
            Case Pcm.CompaniesTableName
                Return "CompanyName"
            Case Pcm.ContactAffilTableName
                Return "CaName"
            Case Pcm.CountriesTableName
                Return "CountryName"
            Case Pcm.ElevCntrsTableName
                Return "EcName"
            Case Pcm.ProjectsTableName
                Return "ProjName"
            Case Pcm.TitlesTableName
                Return "Title"
            Case Else
                Return ColumnName
        End Select
    End Function

    Private Function ConvertTypesCol(TableName As String,
                                     ColumnName As String) As String

        ' gets the converted column name when the column name is "TYPES"
        '   
        ' vars passed:
        '   TableName - name of table with column
        '   ColumnName - column name
        '   
        ' returns:
        '   ColumnName - if column name <> "TYPES", then just return the original column name
        '   "...Type" - converted column name if column name = "TYPES"        

        If ColumnName.ToUpper <> "TYPES" Then
            Return ColumnName
        End If
        Select Case TableName
            Case Pcm.BldgTypesTableName
                Return "BldgType"
            Case Else
                Return ColumnName
        End Select
    End Function

    Private Function ConvertColNames(MySqlConn As MySql.Data.MySqlClient.MySqlConnection,
                                     TableName As String,
                                     MySqlColumns As List(Of MySqlColumn),
                                     IndexesView As DataView) As Integer

        ' converts all columns names with specific MySql Key Words to non key words
        '
        ' vars passed:
        '   MySqlConn - MySql database connection
        '   TableName - Name of table with column to change
        '   MySqlColumns - list of columns in table
        '   IndexView - data view of table indexes
        '
        ' returns:
        '   >= 0 - No Errors
        '   <0 - SqlCommand.ExecuteNonQuery error value
        '   Pcm.sqlNoTableName - no table name 
        '   Pcm.sqlNoKeyName - no key name
        '   Pcm.sqlNoColumn - no columns
        '   Pcm.sqlErr- other error in sql 
        '   Pcm.sqlInvalidAlter - invalid alter option (not atAddColumn or atDropColumn)
        '   Pcm.sqlAlterErr  - error in alter sql         
        '   Pcm.teOtherErr - something went wrong

        Try
            Dim Err As Integer = Pcm.NoErrors
            Dim colNameUpper As String
            For Each col As MySqlColumn In MySqlColumns                                                     ' for each column                
                colNameUpper = col.ColumnName.ToUpper                                                       ' get column name to UPPERCASE
                If colNameUpper = "NAME" OrElse colNameUpper = "TYPES" OrElse colNameUpper = "CODE" Then    ' if a specific name
                    ' convert column name (converts col name in indexes, and index name if needed
                    Err = ConvertColName(MySqlConn, TableName, col, IndexesView)
                    If Err < 0 Then                                                                         ' if got an error
                        Return Err                                                                          ' return the error
                    End If
                End If
            Next

            Return Err
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.CreateTableIndexes.  " & ex.Message)
            Return Pcm.teOtherErr
        End Try
    End Function

    Private Function ConvertColName(MySqlConn As MySql.Data.MySqlClient.MySqlConnection,
                                    TableName As String,
                                    MySqlCol As MySqlColumn,
                                    IndexesView As DataView) As Integer

        ' converts one column name with specific MySql Key Words to non key words
        '
        ' vars passed:
        '   MySqlConn - MySql database connection
        '   TableName - Name of table with column to change
        '   MySqlCol - MySql column with column to change
        '   IndexView - data view of table indexes
        '   UpperCaseColName - name to column in UPPER CASE
        '
        ' returns:
        '   >= 0 - No Errors
        '   <0 - SqlCommand.ExecuteNonQuery error value
        '   Pcm.sqlNoTableName - no table name 
        '   Pcm.sqlNoKeyName - no key name
        '   Pcm.sqlNoColumn - no columns
        '   Pcm.sqlErr- other error in sql 
        '   Pcm.sqlInvalidAlter - invalid alter option (not atAddColumn or atDropColumn)
        '   Pcm.sqlAlterErr  - error in alter sql         
        '   Pcm.teOtherErr - something went wrong

        Try
            Dim Err As Integer = Pcm.NoErrors
            Dim fromColName As String = MySqlCol.ColumnName
            Dim fromColNameUpper As String = MySqlCol.ColumnName.ToUpper
            Dim indexColName As String
            Dim IndexName As String
            ' convert the column name with the correct convertion function
            Select Case fromColNameUpper
                Case "NAME"
                    MySqlCol.ColumnName = ConvertNameCol(TableName, MySqlCol.ColumnName)
                Case "TYPES"
                    MySqlCol.ColumnName = ConvertTypesCol(TableName, MySqlCol.ColumnName)
                Case "CODE"
                    MySqlCol.ColumnName = ConvertCodeCol(TableName, MySqlCol.ColumnName)
            End Select

            Err = MySqlTools.RenameColumn(MySqlConn, TableName, fromColName, MySqlCol)          ' rename column in MySql database table
            For i As Integer = 0 To IndexesView.Count - 1                                       ' for each index view row
                indexColName = GcrcvAs(IndexesView(i).Row, cnColumnName, "")                    ' get the index column name
                If indexColName.ToUpper = fromColNameUpper Then                                 ' if convert name column in the index
                    IndexesView(i).Row(cnColumnName) = MySqlCol.ColumnName                      ' change the index column name to converted name
                End If
                IndexName = GcrcvAs(IndexesView(i).Row, cnIndexName, "")                        ' get index name
                If IndexName.ToUpper = fromColNameUpper Then                                    ' if the index name same as converted column name
                    IndexesView(i).Row(cnIndexName) = MySqlCol.ColumnName                       ' change the index name to converted column name
                End If
            Next
            Return Pcm.NoErrors
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.ConvertIndexNames.  " & ex.Message)
            Return Pcm.teOtherErr
        End Try
    End Function

#End Region

#Region " Copy View/Proc "

    Private Function ConvertQueryToMySql(SqlText As String) As String

        ' converts an Access Sql Query to MySql
        '
        ' vars passed:
        '   SqlText -Access sql query text to convert
        '
        ' returns:
        '   <> "" - converted sql text
        '   string.empty - something went wrong

        Const Space As Char = " "c
        Const Exclimation As Char = "!"c
        Const Period As Char = "."c
        Const AccessSqlIf As String = "IIf("
        Const MySqlIf As String = "IF("
        Const LeftSquareBracket As Char = "["c
        Const RightSquareBracket As Char = "]"c

        Try
            While SqlText.Contains(vbCrLf)                                                      ' if a line break in the sql text
                SqlText = SqlText.Replace(vbCrLf, Space)                                        ' replace line break with just a space
            End While
            While SqlText.Contains(Exclimation)                                                 ' if an exclamation in the sql text
                SqlText = SqlText.Replace(Exclimation, Period)                                  ' replace exclamation with a period
            End While
            While SqlText.Contains(AccessSqlIf)                                                 ' if a "IIf(" in the sql text
                SqlText = SqlText.Replace(AccessSqlIf, MySqlIf)                                 ' replace "IIf(" with "IF("
            End While
            While SqlText.Contains(LeftSquareBracket)                                           ' if got a left square bracket "["
                SqlText = SqlText.Remove(SqlText.IndexOf(LeftSquareBracket), 1)                 ' remove left square bracket 
            End While
            While SqlText.Contains(RightSquareBracket)                                          ' if got a right square bracket "]"
                SqlText = SqlText.Remove(SqlText.IndexOf(RightSquareBracket), 1)                ' remove right square bracket 
            End While

            ' make sure last chr is semi colon 
            Dim lastColon As Integer = SqlText.LastIndexOf(";")                                 ' find last semi colon 
            If lastColon < 0 Then                                                               ' if no ending semi colon
                SqlText &= ";"                                                                  ' add in semi colon
            Else                                                                                ' else found semi colon
                If SqlText.Length - 1 > lastColon Then                                          ' if semi colon is not last char
                    SqlText = SqlText.Substring(0, lastColon + 1)                               ' remove all chars after semi colon
                End If
            End If

            Return SqlText                                                                      ' return converted sql text
        Catch ex As Exception
            Return String.Empty
        End Try
    End Function

    Private Function CopyProcedure(worker As System.ComponentModel.BackgroundWorker,
                                   progress As PcdbToMySqlProgress,
                                   ProcName As String,
                                   SqlText As String) As Integer

        ' creates a query in the refreshed database
        '
        ' vars passed:
        '   worker - background worker
        '   progress - progress values class 
        '   ProcName - procedure to create
        '   SqlText - SQL command (SQL Query text)
        '
        ' returns:
        '   Pcm.NoErrors - no errors
        '   Pcm.deCreateQuery - could not create query
        '   Pcm.deVerifyQuery - query not verified
        '   Return Pcm.teOtherErr - something went wrong

        Try
            ShowProgress(worker, progress, String.Format("Procedure {0}: Creating...", ProcName))
            ' 1) Get the params and reformat query with params
            Dim Params As List(Of MySqlParam) = GetAccessProcParams(SqlText)                                ' get params and also updates SqlText
            ShowProgress(worker, progress, 33)

            ' 2) convert procedure
            Dim MySqlProcSqlText As String = ConvertQueryToMySql(SqlText)                                   ' convert to MySql query
            If MySqlProcSqlText = String.Empty Then                                                         ' if returned empty string
                Return Pcm.deVerifyQuery                                                                    ' return error
            End If
            ShowProgress(worker, progress, 67)

            ' 3) create the procedure
            Dim FullMySqlProcText As String = String.Empty
            Dim Err As Integer = MySqlTools.CreateProcedure(MySqlPcdbConn, ProcName, MySqlProcSqlText, Params, FullMySqlProcText) ' create the proc
            If Err < 0 Then                                                                                 ' if error creating proc
                Return Pcm.deCreateQuery                                                                    ' return error
            End If
            progress.TotalStep()
            ShowProgress(worker, progress, 100)

            '4 ) verify the procedure
            If Not VerifyProc(worker, progress, ProcName, SqlText, FullMySqlProcText, MySqlPcdbConn) Then       ' if did not verify the procedure
                Return Pcm.deVerifyQuery                                                                    ' return error
            End If

            Return Pcm.NoErrors     ' if got here, then return no errors
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.CopyProcedure.  " & ex.Message)
            Return Pcm.teOtherErr
        End Try
    End Function

    Private Function CopyView(worker As System.ComponentModel.BackgroundWorker,
                              progress As PcdbToMySqlProgress,
                              ViewName As String,
                              SqlText As String) As Integer

        ' creates a query in the refreshed database
        '
        ' vars passed:
        '   worker - background worker
        '   progress - progress values class 
        '   ViewName - view to create
        '   SqlText - SQL command (SQL Query text)
        '
        ' returns:
        '   Pcm.NoErrors - no errors
        '   Pcm.deCreateQuery - could not create query
        '   Pcm.deVerifyQuery - query not verified
        '   Return Pcm.teOtherErr - something went wrong

        Try
            ' 1) convert the view
            ShowProgress(worker, progress, String.Format("View {0}: Creating...", ViewName))
            Dim ConvertedSqlText As String = ConvertQueryToMySql(SqlText)                                   ' convert to MySql query
            If SqlText = String.Empty Then                                                                  ' if returned empty string
                Return Pcm.deVerifyQuery                                                                    ' return error
            End If
            ShowProgress(worker, progress, 50)

            '2) create the view
            Dim Err As Integer = MySqlTools.CreateView(MySqlPcdbConn, ViewName, ConvertedSqlText)          ' create the view
            If Err < 0 Then                                                                                 ' if error creating query
                Return Pcm.deCreateQuery                                                                    ' return error
            End If
            progress.TotalStep()                                                                            ' one more total step done
            ShowProgress(worker, progress, 100)                                                             ' show 100 % done

            ' 3) verify the view
            If Not VerifyView(worker, progress, ViewName, SqlText, ConvertedSqlText, MySqlPcdbConn) Then    ' if did not verify the view
                Return Pcm.deVerifyQuery                                                                    ' return error
            End If

            Return Pcm.NoErrors     ' if got here, then return no errors
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.CopyView.  " & ex.Message)
            Return Pcm.teOtherErr
        End Try
    End Function

    Private Function GetMySqlColumn(TableName As String,
                                    ColumnName As String) As MySqlColumn

        ' gets the corresponding mySqlColumn for an Access table/column combo
        '
        ' vars passed:
        '   TableName - name of table
        '   ColumnName - name of column
        '
        ' returns:
        '   MySqlColumn
        '   Nothing - something went wrong

        Const DefaultDecLength As Integer = 14
        Const DefaultDecPercision As Integer = 2

        Try
            Dim Err As Integer = GetTableStructure(AccessOleDbConnection, TableName, AccessColumns, AccessIndexes)  ' get access table structure
            If Err < 0 Then                                                                                         ' if got an error
                Return Nothing                                                                                      ' return nothing
            End If

            Dim SelectStr As String = PcdbModule.FieldSelectString(cnColumnName, ColumnName)                        ' to select just desired column's row
            Dim rows() As DataRow = AccessColumns.Select(SelectStr)                                                 ' select just desired column's row
            If rows Is Nothing OrElse rows.Length <> 1 Then                                                         ' if no or multi rows returned
                Return Nothing                                                                                      ' return nothing
            End If

            Dim eaColType As EaAccess.AccessColumn.ColumnTypes = GetAccessColumnType(rows(0))                       ' get eaColumn type
            Select Case eaColType
                Case EaAccess.AccessColumn.ColumnTypes.ctBoolean
                    Return New MySqlColumn(ColumnName, MySqlDataType.dtTinyInt)
                Case EaAccess.AccessColumn.ColumnTypes.ctCurrency
                    Return New MySqlColumn(ColumnName, DefaultDecLength, DefaultDecPercision)
                Case EaAccess.AccessColumn.ColumnTypes.ctDate
                    Return New MySqlColumn(ColumnName, MySqlDataType.dtDateTime)
                Case EaAccess.AccessColumn.ColumnTypes.ctDouble
                    Return New MySqlColumn(ColumnName, MySqlDataType.dtDouble)
                Case EaAccess.AccessColumn.ColumnTypes.ctInteger
                    Return New MySqlColumn(ColumnName, MySqlDataType.dtInteger)
                Case EaAccess.AccessColumn.ColumnTypes.ctMemo
                    Return New MySqlColumn(ColumnName, MySqlDataType.dtText)
                Case EaAccess.AccessColumn.ColumnTypes.ctNone
                    Return Nothing
                Case EaAccess.AccessColumn.ColumnTypes.ctSmallInt
                    Return New MySqlColumn(ColumnName, MySqlDataType.dtSmallInt)
                Case EaAccess.AccessColumn.ColumnTypes.ctText
                    Dim vcLength As Integer = GcrcvAs(rows(0), cnCharMaxLen, 0)
                    Return New MySqlColumn(ColumnName, MySqlDataType.dtVarChar, vcLength)
                Case Else
                    Return Nothing
            End Select
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.GetParamType.  " & ex.Message)
            Return Nothing
        End Try
    End Function

    Private Function GetAccessProcParams(SqlText As String) As List(Of MySqlParam)

        ' gets the my sql parameters from an Access Sql query.  Params are surrounded by [] also converts params to lower case
        '
        ' vars passed
        '   SqlText - SQL command (SQL Query text)
        '
        ' returns:
        '   list of parameters.  empty list is OK
        '   Nothing - something went wrong

        Const LeftSquareBracket As String = "["
        Const LeftBracket As String = "("
        Const RightSquareBracket As String = "]"
        Const RightBracket As String = ")"

        Try
            Dim Params As New List(Of MySqlParam)
            Dim Param As MySqlParam
            Dim lb As Integer = SqlText.IndexOf(LeftSquareBracket)                              ' find left square bracket
            Dim rb As Integer = -1
            Dim colEndIndex As Integer = -1
            Dim tableStartIndex As Integer = -1
            Dim dotIndex As Integer = -1
            Dim ColumnName As String
            Dim TableName As String
            Dim pCol As MySqlColumn
            Dim VarName As String
            While lb >= 0                                                                       ' while got a left square bracket
                rb = SqlText.IndexOf(RightSquareBracket, lb)                                    ' get next right square bracket index
                If rb < 0 OrElse rb < lb Then                                                   ' if no right square bracket or right before left
                    Return Nothing                                                              ' return nothing
                End If
                colEndIndex = lb                                                                ' start looking for end of column name
                While colEndIndex >= 0 AndAlso SqlText(colEndIndex) <> RightBracket             ' while not past beginning of text and don't have right bracket
                    colEndIndex -= 1                                                            ' go to prior char in sql text
                End While
                If colEndIndex < 0 Then                                                         ' if did not find right bracket
                    Return Nothing                                                              ' return nothing
                End If
                tableStartIndex = colEndIndex                                                   ' start table index search at right bracket
                dotIndex = -1                                                                   ' reset dot index to -1
                While tableStartIndex >= 0 AndAlso SqlText(tableStartIndex) <> LeftBracket      ' while not past beginning of text and don't have left bracket
                    If SqlText(tableStartIndex) = "." Then                                      ' if found dot (separator between table and column)
                        dotIndex = tableStartIndex                                              ' set dot position
                    End If
                    tableStartIndex -= 1                                                        ' go to prior char in sql text
                End While
                If tableStartIndex < 0 OrElse dotIndex < 0 Then                                 ' if table start or dot index
                    Return Nothing                                                              ' return nothing
                End If
                TableName = SqlText.Substring(tableStartIndex + 1, dotIndex - tableStartIndex - 1) ' get table name for param in query
                ColumnName = SqlText.Substring(dotIndex + 1, colEndIndex - dotIndex - 1)        ' get column name for param in query
                pCol = GetMySqlColumn(TableName, ColumnName)                                    ' get param column type info
                If pCol Is Nothing Then                                                         ' if no param column type info
                    Return Nothing                                                              ' return nothing
                End If
                VarName = SqlText.Substring(lb + 1, rb - lb - 1).ToLower                        ' get variable as lower case
                Param = New MySqlParam(VarName, pCol)                                           ' get new param
                Params.Add(Param)                                                               ' add param to list of params

                lb = SqlText.IndexOf(LeftSquareBracket, rb)                                     ' get next left square bracket index                
            End While

            Return Params
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.GetProcParams.  " & ex.Message)
            Return Nothing
        End Try
    End Function

#End Region

#Region " Create Column Lists "

    Private Function CreateColumnLists(TableName As String,
                                       ColumnsView As DataView,
                                       PrimariesView As DataView,
                                       ByRef acColumns As List(Of EaAccess.AccessColumn),
                                       ByRef acPrimaryKeyCols As List(Of EaAccess.AccessColumn),
                                       ByRef MySqlColumns As List(Of MySqlColumn),
                                       ByRef MySqlPrimaryKeyCols As List(Of MySqlColumn)) As Integer

        ' creates the lists of eaColumns and MySqlColumns for the tables
        ' also sorts key fields to beginning of lists
        '
        ' vars passed:
        '   TableName As String - name of table for columns
        '   ColumnsView - data view with column information, sorted by column position in table
        '   PrimariesView - data view with primary key information (table is IndexesTable created in GetTableStructure)
        '   ByRef eaColumns - list of ea columns for the table
        '   ByRef eaPrimaryKeyCols - list of ea primary key columns
        '   ByRef MySqlColumns - list of mySqlColumns for the table
        '   ByRef MySqlPrimaryKeyCols - list of mySql primary key columns

        Const trillionWidth As Integer = 14  ' enough for 999,999,999,999.99 or a trillion $ - 0.01
        Const centsWidth As Integer = 2

        Try
            Dim ColumnName As String
            Dim DefaultValue As String = String.Empty
            Dim acCol As EaAccess.AccessColumn
            Dim acColType As EaAccess.AccessColumn.ColumnTypes
            Dim GotKey As Boolean = False
            Dim IsKeyCol As Boolean = False
            Dim IsNullable As Boolean
            Dim MySqlCol As MySqlColumn
            Dim colDataType As MySqlDataType
            Dim pvRow As DataRow = Nothing
            Dim TextLength As Integer

            Dim FindKeyDataView As New DataView(PrimariesView.Table)                                    ' create new data view
            FindKeyDataView.Sort = cnColumnName & " ASC"                                                ' sort on column name column
            FindKeyDataView.RowFilter = PcdbModule.FieldSelectString(cnPrimaryKey, True)                ' [PRIMARY_KEY] = TRUE
            For i As Integer = 0 To ColumnsView.Count - 1                                               ' for each row in columns view
                DefaultValue = String.Empty                                                             ' reset default value
                ColumnName = GcrcvAs(ColumnsView(i).Row, cnColumnName, String.Empty)                    ' get the column name
                If ColumnName = String.Empty Then                                                       ' if no column name
                    MessageBox.Show(String.Format("No column name when creating table {0}.", TableName),
                                    etCreatingTable, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return Pcm.deNoColumnName                                                           ' return error
                End If
                acColType = GetAccessColumnType(ColumnsView(i).Row)                                     ' get the column type
                If acColType = EaAccess.AccessColumn.ColumnTypes.ctNone Then                            ' if did not get a column type
                    Return Pcm.deInvalidColumnType                                                      ' return error
                End If
                If acColType = EaAccess.AccessColumn.ColumnTypes.ctText Then
                    TextLength = GcrcvAs(ColumnsView(i).Row, cnCharMaxLen, -1)                          ' get length for text
                    If TextLength <= 0 Then                                                             ' if no text length
                        MessageBox.Show(String.Format("No text length when creating text column {0} for table {1}.", ColumnName, TableName),
                                        etCreatingTable, MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return Pcm.deNoTextLength                                                       ' return error
                    End If
                End If
                'DefaultValue = GcrcvAs(ColumnsView(i).Row, cnColDefault, "")                            ' get default value                
                IsNullable = GcrcvAs(ColumnsView(i).Row, cnIsNullable, True)                            ' get is nullable value
                colDataType = ConvertAccessColumnType(acColType)                                        ' convert the column type to mySql                

                If acColType = EaAccess.AccessColumn.ColumnTypes.ctBoolean Then                         ' if an access boolean column
                    'DefaultValue = GetDefaultBooleanValue(TableName, ColumnName).ToString               ' get the default value
                    IsNullable = False
                End If

                If ColumnName = TimeStampColName Then                                                   ' if the time stamp column
                    colDataType = MySqlDataType.dtTimeStamp                                             ' set the column type to time stamp
                    DefaultValue = MySqlTools.TimeStampValue                                            ' set the default value
                End If

                IsKeyCol = False                                                                        ' reset IsKeyCol to false
                If FindKeyDataView.Find(ColumnName) > -1 Then                                           ' if found ColumnName in data view
                    IsKeyCol = True                                                                     ' column is a key column
                    IsNullable = False                                                                  ' cannot be null if a key col
                    GotKey = True                                                                       ' got a key column
                End If

                If colDataType Is MySqlDataType.dtVarChar Then
                    MySqlCol = New MySqlColumn(ColumnName, colDataType, TextLength, IsKeyCol, DefaultValue, IsNullable)
                ElseIf colDataType Is MySqlDataType.dtDecimal Then
                    MySqlCol = New MySqlColumn(ColumnName, trillionWidth, centsWidth, IsKeyCol, DefaultValue, IsNullable)
                Else
                    MySqlCol = New MySqlColumn(ColumnName, colDataType, False, IsKeyCol, DefaultValue, IsNullable)
                End If
                If ColumnName = TimeStampColName Then                                                   ' if time stamp column
                    MySqlCol.OnUpdate = MySqlTools.TimeStampValue                                       ' add on update value
                End If

                If acColType = EaAccess.AccessColumn.ColumnTypes.ctText Then
                    acCol = New EaAccess.AccessColumn(ColumnName, acColType, TextLength, IsKeyCol, DefaultValue, IsNullable)
                Else
                    acCol = New EaAccess.AccessColumn(ColumnName, acColType, IsKeyCol, DefaultValue, IsNullable)
                End If

                MySqlColumns.Add(MySqlCol)
                acColumns.Add(acCol)
            Next

            For Each col As EaAccess.AccessColumn In acColumns                              ' for each ea column
                If col.KeyField Then                                                        ' if a key column
                    acCol = New EaAccess.AccessColumn(col.ColumnName, col.ColumnType)       ' get a new column
                    acPrimaryKeyCols.Add(acCol)                                             ' add column to list of key columns
                End If
            Next
            For Each col As MySqlColumn In MySqlColumns                                     ' for each mySql column
                If col.KeyField Then                                                        ' if a key field
                    MySqlCol = New MySqlColumn(col.ColumnName, col.DataType)                ' get a new column
                    MySqlPrimaryKeyCols.Add(MySqlCol)                                       ' add column to list of key columns
                End If
            Next
            MovePrimaryKeyColsToBeginning(acColumns, MySqlColumns, MySqlPrimaryKeyCols)     ' move primary key columns to beginning of list

            ' if the first column is an int, and only one key column
            If MySqlColumns(0).DataType Is MySqlDataType.dtInteger _
                        AndAlso MySqlColumns(0).KeyField = True _
                        AndAlso PrimariesView.Count = 1 Then
                MySqlColumns(0).IsNullable = False
            End If

            Return Pcm.NoErrors
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.CreateMySqlColumnList.  " & ex.Message)
            Return Nothing
        End Try
    End Function

    Private Function GetDefaultBooleanValue(TableName As String, ColumnName As String) As Boolean

        ' gets the default value for a boolean column
        '
        ' vars passed:
        '   TableName - name of table for column
        '   ColumnName - column name
        '
        ' returns:
        '   TRUE - default value for column is TRUE
        '   FALSE - 
        '     - default value for column is FALSE
        '     - could not file table name
        '     - something went wrong

        Try
            Select Case TableName
                Case Pcm.CompaniesTableName
                    If ColumnName = Pcm.CompActiveColName Then          ' if "Active" column
                        Return True                                     ' return TRUE
                    Else                                                ' else all other boolean columns
                        Return False                                    ' return FALSE
                    End If
                Case Pcm.ContactsTableName
                    If ColumnName = Pcm.ContActiveColName Then          ' if "Active" column
                        Return True                                     ' return TRUE
                    Else                                                ' else all other boolean columns
                        Return False                                    ' return FALSE
                    End If
                Case Pcm.DelivTableName
                    Return False
                Case EwEmpsTableName
                    If ColumnName = Pcm.EwEmpActiveColName Then         ' if "Active" column
                        Return True                                     ' return TRUE
                    Else                                                ' else all other boolean columns
                        Return False                                    ' return FALSE
                    End If
                Case Pcm.JobsTableName
                    Select Case ColumnName
                        Case Pcm.JobsActiveColName
                            Return True
                        Case Pcm.JobsBillableColName
                            Return True
                        Case Else
                            Return False
                    End Select
                Case Pcm.PlayersTableName
                    Return False                                        ' return false for all Players table boolean columns
                Case Pcm.ProjectsTableName
                    Return False                                        ' return false for all Projects table boolean columns
                Case Pcm.ProjHistPlayersTableName
                    If ColumnName = Pcm.ProjHPPrimColName Then          ' if "PrimaryContact" column
                        Return True                                     ' return TRUE
                    Else                                                ' else all other boolean columns
                        Return False                                    ' return FALSE
                    End If
                Case Pcm.ReimToBillTableName
                    If ColumnName = Pcm.Reim2BOkToBillColName Then      ' if "OkToBill" column
                        Return True                                     ' return TRUE
                    Else                                                ' else all other boolean columns
                        Return False                                    ' return FALSE
                    End If
                Case Pcm.TravelTableName
                    Return False                                        ' return false for all Travel table boolean columns
                Case Else                                               ' table not found
                    Return False                                        ' return FALSE
            End Select
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.SetDefaultBooleanValue.  " & ex.Message)
            Return False
        End Try
    End Function

#End Region

#Region " Create Table "

    Private Function CreateTables(TableName As String,
                                  acColumns As List(Of EaAccess.AccessColumn),
                                  MySqlColumns As List(Of MySqlColumn),
                                  ByRef AccessTable As DataTable) As Integer

        ' 1) create the table in the mySql database
        ' 2) make the Access table in memory        
        '   
        ' vars passed:
        '   TableName - name of table to create
        '   acColumns - list of access columns
        '   MySqlColumns - list of MySql Columns
        '   ByRef AccessTable - access table to create        
        '   
        ' returns:
        '   Pcm.NoErrors - no errors
        '   Pcm.deInvalidColumnType - invalid column type
        '   Pcm.deCreateTable - could not create table
        '   Pcm.deCreatePrimaryKey - could not create primary key
        '   Pcm.teOtherErr - something went wrong

        Try
            'SetAction(String.Format("Table {0}: Creating...", TableName))

            ' 1) create the table in the mySql database
            Dim Err As Integer = MySqlTools.CreateTable(MySqlPcdbConn, TableName, MySqlColumns, True)       ' create table             
            If Err < 0 Then                                                                                 ' if error creating table
                CancelBackgroundWorker()
                MessageBox.Show(String.Format("Could not create table {0} in database {1}.", TableName, MySqlPcdbConn.Database),
                                etCreatingTable, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return Pcm.deCreateTable                                                                    ' return error
            End If

            ' 2) make the Access table in memory
            AccessTable = MakeAccessDataTable(TableName, acColumns)                                         ' make access table in memory
            If AccessTable Is Nothing Then                                                                  ' if could not make data table
                Return Pcm.teNoTableErr                                                                     ' return error
            End If

            Return Pcm.NoErrors             ' if got here, return no errors
        Catch ex As Exception
            CancelBackgroundWorker()
            PcdbModule.ShowOtherErrorMessage("MainForm.CreateMySqlTable.  " & ex.Message)
            Return Pcm.teOtherErr
        End Try
    End Function

    Private Sub MovePrimaryKeyColsToBeginning(acColumns As List(Of EaAccess.AccessColumn),
                                              MySqlColumns As List(Of MySqlColumn),
                                              PrimaryKeyCols As List(Of MySqlColumn))

        ' moves primary key columns to start of list
        '
        ' vars passed:
        '   acColumns - list of access columns for table
        '   MySqlColumns - list of mySqlColumns for table
        '   PrimariesView - dataview with rows of primary key columns

        Try
            Dim foundCol As MySqlColumn = Nothing
            Dim KeyColName As String = String.Empty
            Dim foundIndex As Integer
            Dim tempeaCol As New EaAccess.AccessColumn(String.Empty, EaAccess.AccessColumn.ColumnTypes.ctNone)
            Dim tempMySqlCol As New MySqlColumn(String.Empty, MySqlDataType.dtNone)                     ' get temp column, no data type needed
            Dim EndPos As Integer = 1                                                                   ' end switching when switch 1 and 0
            For Each pCol As MySqlColumn In PrimaryKeyCols                                              ' for each primary key column
                foundCol = MySqlColumns.Find(Function(c As MySqlColumn) c.ColumnName = pCol.ColumnName) ' find the primary key column in list of columns
                If foundCol IsNot Nothing Then                                                          ' if found the primary key column
                    foundIndex = MySqlColumns.FindIndex(Function(f) foundCol.Equals(f))                 ' get the index in the list of columns
                    For j As Integer = foundIndex To EndPos Step -1                                     ' for each column to switch
                        tempMySqlCol = MySqlColumns(j)                                                  ' get column to move up in list
                        tempeaCol = acColumns(j)
                        MySqlColumns(j) = MySqlColumns(j - 1)                                           ' move prior column down in list
                        acColumns(j) = acColumns(j - 1)
                        MySqlColumns(j - 1) = tempMySqlCol                                              ' move current column up in list
                        acColumns(j - 1) = tempeaCol
                    Next
                    EndPos += 1                         ' now end at next key column (if ended at 1 & 0, next end at 2 & 1)
                End If
            Next
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.MovePrimaryKeyColsToBeginning.  " & ex.Message)
        End Try
    End Sub

#End Region

#Region " Database Connection "

    Private Function GetSslMode() As MySqlSslMode

        If PreferredRadioButton.Checked Then
            Return MySqlSslMode.Preferred
        ElseIf RequiredRadioButton.Checked Then
            Return MySqlSslMode.Required
        ElseIf VerifyCARadioButton.Checked Then
            Return MySqlSslMode.VerifyCA
        ElseIf VerifyFullRadioButton.Checked Then
            Return MySqlSslMode.VerifyFull
        Else
            Return MySqlSslMode.None
        End If
    End Function

    Private Function MakeAccessConnection(DatabaseFolder As String,
                                          DatabaseFileName As String,
                                          OleDbConn As OleDb.OleDbConnection) As Integer

        ' connects to the access database
        '
        ' vars passed:
        '   DatabaseFolder - folder where the data base is
        '   DatabaseFileName - database file name (no path)
        '   OleDbConn - connection to use
        '
        ' returns:
        '   Pcm.NoErrors - no errors
        '   Pcm.teNoDataPathErr - no data path
        '   Pcm.teDataPathNoFoundErr - datapath not found
        '   Pcm.teDatabaseNotFoundErr - database not found in datapath
        '   Pcm.teNoConnectionStringErr - no connection string
        '   Pcm.teConnectionErr - error setting connection string

        Try
            Dim ConnErr As Integer = SetAccessConnectionString(DatabaseFolder, DatabaseFileName, OleDbConn) ' set the data connection
            If ConnErr = Pcm.teNoDataPathErr _
                    OrElse ConnErr = teDataPathNoFoundErr _
                    OrElse ConnErr = Pcm.teDatabaseNotFoundErr Then                                         ' if database not found
                ShowAccessConnectionErr(DatabaseFolder, DatabaseFileName, ConnErr)                          ' show connection error 
            Else                                                                                            ' else no error or connection error                
                If ConnErr <> Pcm.NoErrors Then                                                             ' if error connection 
                    ShowAccessConnectionErr(DatabaseFolder, DatabaseFileName, ConnErr)                      ' show connection error
                End If
            End If
            Return ConnErr
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.MakeAccessConnection().  " & ex.Message)
            Return Pcm.teOtherErr
        End Try
    End Function

    Private Function MakeAccessAndMySqlConnections() As Integer

        ' makes both the access and MySql connections. updates the action message and progress bars
        '
        ' returns:
        '   NoErrors - no errors
        '   Pcm.teNoDataPathErr - no data path
        '   Pcm.teDataPathNoFoundErr - datapath not found
        '   Pcm.teDatabaseNotFoundErr - database not found in datapath
        '   Pcm.teNoConnectionStringErr - no connection string
        '   Pcm.teConnectionErr - error setting connection string
        '   teOtherErr - something went wrong

        Try
            SetAction("Connecting to to Databases")                                                             ' show action being done
            Dim Err As Integer = MakeAccessConnection(AccessFolder, AccessDatabaseName, AccessOleDbConnection)  ' make connection to access database
            If Err <> Pcm.NoErrors Then                                                                         ' if got an error
                Return Err                                                                                      ' return error
            End If
            ActionPosition(50)
            TotalProgressBarStep()

            Err = MakeMySqlConnection(MySqlServerTextEdit.Text, UserIdTextEdit.Text, PasswordTextEdit.Text, MySqlPcdbNameTextEdit.Text) ' make sql connection
            If Err <> Pcm.NoErrors Then                                                                         ' if got an error
                Return Err                                                                                      ' return error
            End If
            ActionPosition(100)
            TotalProgressBarStep()

            Return NoErrors
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.MakeAccessAndSqlConnections().  " & ex.Message)
            Return Pcm.teOtherErr
        End Try
    End Function

    Private Function MakeMySqlConnection(serverName As String,
                                         userId As String,
                                         password As String,
                                         dbName As String) As Integer

        ' makes the MySql connection
        '
        ' vars passed
        '   serverName - name or address of server
        '   userId - user name
        '   password - user's password
        '   dbName - database name
        '
        ' returns:
        '   NoErrors - no errors
        '   EaSql.Ea.sqlConnectionStringErr - connection string error
        '   Pcm.teNoConnectionStringErr - no connection string
        '   EaMySql.sqlNoConnection - error connection to mySql
        '   teOtherErr - something went wrong

        Dim MySqlConnStr As String = String.Empty
        Try
            If MySqlPcdbConn IsNot Nothing Then
                MySqlPcdbConn.Dispose()
                MySqlPcdbConn = Nothing
            End If
            MySqlConnStr = MySqlTools.ConnectionString(serverName, userId, password, dbName,
                                                       PersistSecurityInfoCheckBox.Checked, GetSslMode)        ' get connection string
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.MakeMySqlConnection(): Getting MySql Connection string" & ex.Message)
            Return EaSql.Ea.sqlConnectionStringErr
        End Try

        If MySqlConnStr = String.Empty Then
            Return Pcm.teNoConnectionStringErr
        End If

        Try
            Dim em As String = MySqlTools.ConnectionErrMsg
            Dim ec As Integer = MySqlTools.ConnectionErr

            MySqlPcdbConn = MySqlTools.MySqlConn(MySqlConnStr)
            If MySqlPcdbConn Is Nothing Then
                Dim errMsg As String = MySqlTools.CustomConnErrorMessage
                PcdbModule.ShowErrorMessage("Error connecting to MySql.  " & errMsg, "SQL Connect Error")
                Return MySqlTools.sqlNoConnection
            End If
            MySqlPcdbConn.Open()

            Return Pcm.NoErrors                                         ' Finally section will close connection, so OK to return before close
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.MakeMySqlConnection(): Opening" & ex.Message)
            Return Pcm.teConnectionErr
        Finally
            If MySqlPcdbConn IsNot Nothing Then
                MySqlPcdbConn.Close()
            End If
        End Try
    End Function

    Private Function SetAccessConnectionString(DataPathToUse As String,
                                               DatabaseFileName As String,
                                               OleDbConn As OleDb.OleDbConnection) As Integer

        ' sets the connection string for the access database
        '
        ' vars passed:
        '   DataPathToUse - data path for DataBase file
        '   DatabaseFileName - database file name (no path)
        '   ConnectionToUse - ole db connection to set data path for
        '
        ' returns:
        '   Pcm.NoErrors - no errors
        '   Pcm.teNoDataPathErr - no data path
        '   Pcm.teDataPathNoFoundErr - datapath not found
        '   Pcm.teDatabaseNotFoundErr - database not found in datapath
        '   Pcm.teNoConnectionStringErr - no connection string
        '   Pcm.teConnectionErr - error setting connection string

        If DataPathToUse = "" Then          ' if no data path
            Return Pcm.teNoDataPathErr      ' return no data path err
        End If

        ' get the current connection string
        Dim ConnStr As String = OleDbConn.ConnectionString
        Try
            EndWithSlash(DataPathToUse)                                                     ' make sure data path ends with slash
            DataPathToUse = ConvertSpecialDirectory(DataPathToUse)                          ' convert if a special dir

            If Not FolderExits(DataPathToUse) Then                                          ' if cannot find datapath
                Return Pcm.teDataPathNoFoundErr                                             ' return error
            End If
            If Not My.Computer.FileSystem.FileExists(DataPathToUse & DatabaseFileName) Then ' if no database
                Return Pcm.teDatabaseNotFoundErr                                            ' return no database error
            End If

            Dim DpIndex As Integer = ConnStr.IndexOf(OleDbConn.DataSource)                  ' get the data source index in the connection string
            ConnStr = ConnStr.Remove(DpIndex, OleDbConn.DataSource.Length)                  ' remove the data source
            ConnStr = ConnStr.Insert(DpIndex, DataPathToUse & DatabaseFileName)             ' insert the new data source (DataPathToUse + DataFileName)
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.SetAccessConnectionString(DataPathToUse,DatabaseFileName,ConnectionToUse) - Get Connection String.  " & ex.Message)
            ConnStr = String.Empty
        End Try

        If ConnStr = String.Empty Then                                                      ' if no connection string
            Return Pcm.teNoConnectionStringErr                                              ' return no connection string err
        End If

        Try
            OleDbConn.ConnectionString = ConnStr                                            ' set new connection string
            Return Pcm.NoErrors                                                             ' if got here, then no errors
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.SetAccessConnectionString(DataPathToUse,DatabaseFileName,ConnectionToUse) - Set New Connection String.  " & ex.Message)
            Return Pcm.teConnectionErr
        End Try
    End Function

    Private Sub ShowAccessConnectionErr(DataPathToUse As String,
                                        DatabaseFileName As String,
                                        ConnErr As Integer)

        ' shows the connection error
        '
        ' vars passed:
        '   DataPathToUse - path for database
        '   DatabaseFileName - database file name (no path)
        '   ConnErr - connection error

        If ConnErr = Pcm.NoErrors Then
            Return
        End If
        Try
            Select Case ConnErr
                Case Pcm.teNoDataPathErr
                    PcdbModule.ShowErrorMessage("There is no Folder for the database.",
                                                Pcm.etDatabaseNotFound, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Case Pcm.teDataPathNoFoundErr
                    PcdbModule.ShowErrorMessage(String.Format("Could not find the folder {0}.", DataPathToUse),
                                                Pcm.etDatabaseNotFound, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Case Pcm.teDatabaseNotFoundErr
                    PcdbModule.ShowErrorMessage(String.Format("Could not find the database {0} in the folder {1}.", DatabaseFileName, DataPathToUse),
                                                Pcm.etDatabaseNotFound, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Case Pcm.teNoConnectionStringErr
                    PcdbModule.ShowErrorMessage(String.Format("Could not create a connection string for the database {0} in the folder {1}.", DatabaseFileName, DataPathToUse),
                                                Pcm.etDatabaseNotFound, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Case Pcm.teConnectionErr
                    PcdbModule.ShowErrorMessage(String.Format("Could not set connection for the database {0} in the folder {1}.", DatabaseFileName, DataPathToUse),
                                                Pcm.etDatabaseNotFound, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Case Else
                    PcdbModule.ShowErrorMessage(String.Format("Other error trying to connect to the database {0} in the folder {1}.", DatabaseFileName, DataPathToUse),
                                                Pcm.etDatabaseNotFound, MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Select
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.ShowConnectionErr.  " & ex.Message)
        End Try
    End Sub

#End Region

#Region " Drop DataLocInt column "

    Private Function DropDataLocIntCol(MySqlConn As MySql.Data.MySqlClient.MySqlConnection,
                                       TableName As String) As Integer

        ' drops the DataLocInt column (it is not needed in MySql)
        '
        ' vars passed:
        '   MySqlConn - MySql database connection
        '   TableName - Name of table with column to change
        '
        ' returns:
        '   >= 0 - No Errors
        '   <0 - SqlCommand.ExecuteNonQuery error value
        '   Pcm.sqlNoTableName - no table name 
        '   Pcm.sqlErr- other error in sql 
        '   Pcm.sqlInvalidAlter - invalid alter option (not atAddColumn or atDropColumn)
        '   Pcm.sqlAlterErr  - error in alter sql         
        '   Pcm.teOtherErr - something went wrong

        Try
            Return MySqlTools.DropColumn(MySqlConn, TableName, Pcm.ccDataLocColName)
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.DropDataLocIntCol.  " & ex.Message)
            Return Pcm.teOtherErr
        End Try
    End Function

#End Region

#Region " Fill/Populate Table "

    Private Function FillTable(TableToFill As DataTable,
                               OleDbAdapt As OleDb.OleDbDataAdapter,
                               Optional Params As List(Of EaAccess.AccessParam) = Nothing,
                               Optional DoClear As Boolean = True) As Integer

        ' fills the table by using OleDbDataAdapter.Fill(aTable)
        ' 
        ' vars passed:
        '   TableToFill - data table that will be filled
        '   OleDbAdapt - OleDbDataAdapter for table
        '   Params - list of params for sql text
        '   DoClear - if true, clears the table before filling it
        '
        ' returns:
        '   >= 0 the number of data rows that will filled/loaded
        '   Pcm.deNoDataAdapt - no oleDbDataAdapter was found for table
        '   Pcm.teConstraintErr - there was a constraint error while filling table
        '   Pcm.teLoadErr - there was some other error
        '   Pcm.teNoTableErr - no table 
        '   Pcm.teNoColumnErr - column not found in table
        '   Pcm.tePrimaryKeyErr - other error

        Dim TryCount As Integer = 0
        Do
            Try
                If TableToFill Is Nothing Then                                  ' if no table
                    MessageBox.Show("Table to Fill is nothing.", Pcm.etLoadTable, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return teNoTableErr                                         ' return no table err
                End If
                If OleDbAdapt Is Nothing Then                                   ' if no ole db adapter
                    MessageBox.Show(String.Format("No OleDbDataAdapter for table {0}.", TableToFill.TableName),
                                    Pcm.etLoadTable, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return Pcm.deNoDataAdapt                                    ' return no data adapter err
                End If
                If Params IsNot Nothing Then
                    If OleDbAdapt.SelectCommand.Parameters.Count = 0 Then
                        For Each param In Params
                            OleDbAdapt.SelectCommand.Parameters.AddWithValue(param.Name, param.Value)
                        Next
                    Else
                        For Each param In Params
                            For i As Integer = 0 To OleDbAdapt.SelectCommand.Parameters.Count - 1
                                If OleDbAdapt.SelectCommand.Parameters(i).ParameterName.ToUpper = param.Name.ToUpper Then
                                    OleDbAdapt.SelectCommand.Parameters(i).Value = param.Value
                                    Exit For
                                End If
                            Next
                        Next
                    End If
                End If
                'Cursor = Cursors.WaitCursor                                     ' show wait cursor
                If TableToFill.Rows.Count > 0 AndAlso DoClear Then              ' if got rows, and want to clear
                    If PcdbModule.SafeClear(TableToFill) <> Pcm.NoErrors Then   ' if could not clear table
                        Return Pcm.teCouldNotClearErr                           ' return error
                    End If
                End If
                Return OleDbAdapt.Fill(TableToFill)                             ' fill table, return # of rows filled
            Catch OleDbEx As Data.OleDb.OleDbException
                If OleDbEx.ErrorCode = Pcm.oleCannotOpen AndAlso TryCount < Pcm.oleMaxTries Then
                    TryCount += 1
                    'My.Application.DoEvents()
                    'Debug.WriteLine("No DoEvents")
                    System.Threading.Thread.Sleep(Pcm.oleWaitMiliSeconds)
                Else
                    MessageBox.Show(String.Format("OleDb Error while loading data table {0}{1}Error Code:{4}{1}Database: {2}{1}{3}",
                                                  TableToFill.TableName, vbCrLf, OleDbAdapt.SelectCommand.Connection.DataSource, OleDbEx.Message, OleDbEx.ErrorCode),
                                    Pcm.etLoadTable, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return Pcm.teOtherErr
                End If
            Catch ex As Exception
                PcdbModule.ShowOtherErrorMessage("MainForm.FillTable.  " & ex.Message)
                Return Pcm.teLoadErr
            Finally
                'Cursor = Cursors.Default                                        ' always set cursor back to normal
            End Try
        Loop
    End Function

#End Region

#Region " Get Column Type "

    Private Function GetAccessColumnType(ColumnInfoRow As DataRow) As EaAccess.AccessColumn.ColumnTypes

        ' gets the eaColumnClass column type 
        '
        ' vars passed:
        '   ColumnInfoRow - data row from the columns table 
        '
        ' returns:
        '   eaColumnClass.ColumnTypes.ctInteger 
        '   eaColumnClass.ColumnTypes.ctDouble
        '   eaColumnClass.ColumnTypes.ctCurrency
        '   eaColumnClass.ColumnTypes.ctDate
        '   eaColumnClass.ColumnTypes.ctBoolean
        '   eaColumnClass.ColumnTypes.ctMemo                         
        '   eaColumnClass.ColumnTypes.ctText
        '   eaColumnClass.ColumnTypes.ctNone - 
        '     - Data type for column info row not compatible with eaColumnClass.ColumnTypes
        '     - something went wrong

        ' the column type conversions are:
        '	DATA_TYPE INT	DATA_TYPE NAME			
        '	0				Empty
        '	2				SmallInt
        '	3				Integer
        '	4				Single
        '	5				Double
        '	6				Currency
        '	7				Date
        '	8				BSTR
        '	9				IDispatch
        '	10				Error
        '	11				Boolean
        '	12				Variant
        '	13				IUnknown
        '	14				Decimal
        '	16				TinyInt
        '	17				UnsignedTinyInt
        '	18				UnsignedSmallInt
        '	19				UnsignedInt
        '	20				BigInt
        '	21				UnsignedBigInt
        '	64				File time
        '	72				Guid
        '	128				Binary
        '	129				Char
        '	130				WChar
        '	131				Numeric
        '	133				DBDate
        '	134				DBTime
        '	135				DBTimeStamp
        '	138				PropVariant
        '	139				VarNumeric
        '	200				VarChar
        '	201				LongVarChar
        '	202				VarWChar
        '	203				LongVarWChar
        '	204				VarBinary
        '	205				LongVarBinary

        ' since the func only cares about eaColumnClass.ColumnTypes, the only valid values are:
        '	DATA_TYPE INT	DATA_TYPE NAME			
        '	0				Empty
        '	3				Integer
        '	5				Double
        '	6				Currency
        '	7				Date
        '	11				Boolean
        '	130				WChar

        Try
            Dim DataTypeInt As Integer = GcrcvAs(ColumnInfoRow, cnDataType, -1)
            Select Case DataTypeInt
                Case 2
                    Return EaAccess.AccessColumn.ColumnTypes.ctSmallInt
                Case 3
                    Return EaAccess.AccessColumn.ColumnTypes.ctInteger
                Case 5
                    Return EaAccess.AccessColumn.ColumnTypes.ctDouble
                Case 6
                    Return EaAccess.AccessColumn.ColumnTypes.ctCurrency
                Case 7
                    Return EaAccess.AccessColumn.ColumnTypes.ctDate
                Case 11
                    Return EaAccess.AccessColumn.ColumnTypes.ctBoolean
                Case 130
                    Dim MaxLength As Integer = GcrcvAs(ColumnInfoRow, cnCharMaxLen, -1) ' get max length for text
                    If MaxLength = 0 Then                                               ' if no max length
                        Return EaAccess.AccessColumn.ColumnTypes.ctMemo                 ' then a memo column
                    Else                                                                ' else got a max length
                        Return EaAccess.AccessColumn.ColumnTypes.ctText                 ' a text column
                    End If
                Case Else
                    MessageBox.Show(String.Format("Error in GetEaColumnType.  DataIntType: {0} not used in eaColumnClass.ColumnTypes.", DataTypeInt),
                                    "Invalid Column Type", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return EaAccess.AccessColumn.ColumnTypes.ctNone
            End Select
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.GetEaColumnType.  " & ex.Message)
            Return EaAccess.AccessColumn.ColumnTypes.ctNone
        End Try
    End Function

    Private Function GetSystemColumnTypeStr(ColType As EaAccess.AccessColumn.ColumnTypes) As String

        ' gets the string ColTypeStr used to create a column in a call to New DataColumn(ColName, ColTypeStr)
        '
        ' vars passed:
        '   ColType - column type to create
        '
        ' returns:
        '   "System.columnType"
        '   "" -
        '     ColType not valid
        '     - something went wrong

        Try
            Select Case ColType
                Case EaAccess.AccessColumn.ColumnTypes.ctBoolean
                    Return "System.Boolean"
                Case EaAccess.AccessColumn.ColumnTypes.ctCurrency
                    Return "System.Decimal"
                Case EaAccess.AccessColumn.ColumnTypes.ctDate
                    Return "System.DateTime"
                Case EaAccess.AccessColumn.ColumnTypes.ctDouble
                    Return "System.Double"
                Case EaAccess.AccessColumn.ColumnTypes.ctInteger
                    Return "System.Int32"
                Case EaAccess.AccessColumn.ColumnTypes.ctMemo
                    Return "System.String"
                Case EaAccess.AccessColumn.ColumnTypes.ctText
                    Return "System.String"
                Case EaAccess.AccessColumn.ColumnTypes.ctSmallInt
                    Return "System.Int16"
                Case Else
                    MessageBox.Show(String.Format("Error in GetSystemColumnType.  ColType: {0} not used in eaColumnClass.ColumnTypes.", ColType),
                                    "Invalid Column Type", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return String.Empty
            End Select
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.GetSystemColumnType.  " & ex.Message)
            Return ""
        End Try
    End Function

#End Region

#Region " Get Tables, Queries Views and Procedures "

    Private Function GetAccessTablesAndQueries() As Integer

        ' fills CurTables with data about all tables in CurrentOleDbConnection
        ' fills CurViews with data about all views in CurrentOleDbConnection
        ' fills CurrProcs with data about all procedures in OleDbConn
        '
        ' returns:
        '   Pcm.NoErrors - no errors
        '   Pcm.teOtherErr - something went wrong

        Try
            ' fill tables            
            Dim Err As Integer = EaAccess.AccessSql.GetSchemeTablesAndQueries(AccessOleDbConnection, AccessTables, AccessViews, AccessProcs)
            If Err <> Pcm.NoErrors Then                                                             ' if got an error
                Return Err                                                                          ' return error
            End If

            Return Pcm.NoErrors
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.GetAccessTablesAndQueries.  " & ex.Message)
            Return Pcm.teOtherErr
        End Try
    End Function

#End Region

#Region " MakeSingleKeyFieldAutoInc "

    Private Function AlterKeyColumnToAutoInc(MySqlConn As MySql.Data.MySqlClient.MySqlConnection,
                                             TableName As String,
                                             ColumnName As String) As Integer


        ' alters a column to an auto inc primary key column
        '
        ' vars passed:
        '   MySqlConn - MySql database connection
        '   TableName - name of table to check
        '   MySqlColumns - list of MySqlColumns for table
        '   PrimariesView - data view of primary key fields
        '
        ' returns:
        '   Pcm.NoErrors - no errors
        '   <0 - SqlCommand.ExecuteNonQuery error value
        '   Pcm.sqlErr - other error in sql 
        '   Pcm.teOtherErr - something went wrong


        '   "ALTER TABLE <TableName> MODIFY COLUMN <KeyColumnName> INT AUTO_INCREMENT"
        '   "ALTER TABLE <TableName> MODIFY <KeyColumnName> INT AUTO_INCREMENT PRIMARY KEY"
        '   "ALTER TABLE <TableName> MODIFY COLUMN <KeyColumnName> INT NOT NULL AUTO_INCREMENT"        

        Try
            Dim SqlText As String = String.Format("ALTER TABLE `{0}` MODIFY COLUMN `{1}` INT NOT NULL AUTO_INCREMENT", TableName, ColumnName)
            Dim ErrMsg As String = String.Format("Error modifying column.{0}Column: {1}{0}Table: {2}{0}DataBase: {3}{0}Server: {4}",
                                                 vbCrLf, ColumnName, TableName, MySqlConn.Database, MySqlConn.DataSource)
            Return MySqlTools.ExecuteNonQuery(MySqlConn, SqlText)
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.AlterKeyColumnToAutoInc.  " & ex.Message)
            Return Pcm.teOtherErr
        End Try
    End Function

    Private Function FixWorkDoneTableId(MySqlConn As MySql.Data.MySqlClient.MySqlConnection,
                                        MySqlAdapt As MySql.Data.MySqlClient.MySqlDataAdapter,
                                        wdTable As DataTable) As Integer

        ' 

        Try
            ' SELECT MAX(WorkDoneId) as WorkDoneIdMax from affiliations
            Dim SqlText As String = String.Format("SELECT MAX(`{0}`) AS {0}Max FROM `{1}`", Pcm.WdIdColName, Pcm.WorkDoneTableName)
            Dim MaxIdObj As Object = MySqlTools.ExecuteScalar(MySqlConn, SqlText)
            Dim MaxWdId As Integer = CInt(MaxIdObj)                                             ' get the max work done id
            If MaxWdId < 0 Then                                                                 ' if got an error
                Return MaxWdId                                                                  ' return error
            End If
            Dim wdRow As DataRow = PcdbModule.FindRow(wdTable, 0)                               ' get the row to fix
            If wdRow Is Nothing Then                                                            ' if did not find row
                Return Pcm.teRowNotFound                                                        ' return error
            End If
            MaxWdId += 1                                                                        ' get the next max id
            wdRow(Pcm.WdIdColName) = MaxWdId                                                    ' fix the id in the row

            Dim RowCount As Integer = UpdateTable(wdTable, MySqlConn, MySqlAdapt, 1)            ' update the table
            If RowCount <> 1 Then                                                               ' if did not update one row
                Return Pcm.teRowCountErr                                                        ' return error
            End If

            Return MaxWdId                                                                      ' return new max id
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.FixWorkDoneTableIds.  " & ex.Message)
            Return Pcm.teOtherErr
        End Try
    End Function

    Private Function MakeSingleKeyFieldAutoInc(MySqlConn As MySql.Data.MySqlClient.MySqlConnection,
                                               TableName As String,
                                               MySqlColumns As List(Of MySqlColumn),
                                               PrimariesView As DataView) As Integer

        ' makes a single int key column an auto inc column
        '
        ' vars passed:
        '   MySqlConn - MySql database connection
        '   TableName - name of table to check
        '   MySqlColumns - list of MySqlColumns for table
        '   PrimariesView - data view of primary key fields
        '
        ' returns:
        '   Pcm.NoErrors - no errors
        '   <0 - SqlCommand.ExecuteNonQuery error value
        '   Pcm.sqlErr - other error in sql 
        '   Pcm.teOtherErr - something went wrong

        Try
            If TableName = Pcm.CalcInvsTableName _
                    OrElse TableName = Pcm.DelivToBillTableName _
                    OrElse TableName = Pcm.ExpDisbTableName _
                    OrElse TableName = Pcm.FeeToBillTableName _
                    OrElse TableName = Pcm.ReimToBillTableName _
                    OrElse TableName = Pcm.InvJobCompAttnTableName _
                    OrElse TableName = Pcm.PastDueHistTableName _
                    OrElse TableName = Pcm.PreBillTableName _
                    OrElse TableName = Pcm.PbFixedFeeTableName _
                    OrElse TableName = Pcm.TravPerMileTableName _
                    OrElse TableName = Pcm.TravToBillTableName Then
                Return Pcm.NoErrors                                                                 ' return no errors
            End If
            Dim Err As Integer = Pcm.NoErrors

            If PrimariesView.Count = 1 Then                                                         ' if one key field
                ' if first col is int and a key field
                If MySqlColumns(0).DataType Is MySqlDataType.dtInteger AndAlso MySqlColumns(0).KeyField = True Then
                    MySqlColumns(0).AutoInc = True                                                  ' change column to auto inc
                    MySqlColumns(0).IsNullable = False                                              ' column is not nullable
                    Err = AlterKeyColumnToAutoInc(MySqlConn, TableName, MySqlColumns(0).ColumnName) ' change col type in MySql database
                    If Err < 0 Then                                                                 ' if got an error
                        Return Err                                                                  ' return the error
                    End If
                Else                                                                                ' else first col is not int and key
                    PcdbModule.ShowErrorMessage(String.Format("First column ""{0}"" is not an INT or not a key column", MySqlColumns(0).ColumnName),
                                                "Change Structure Error")
                    Return Pcm.teNoPrimaryKey                                                       ' return error
                End If
            End If
            Return Pcm.NoErrors              ' if got here, return no errors
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.MakeSingleKeyFieldAutoInc.  " & ex.Message)
            Return Pcm.teOtherErr
        End Try
    End Function

#End Region

#Region " Sort Queries "

    Private Function GetDirectQueriesUsed(QueryInfo As QueryInfoClass) As Integer

        ' gets the other queries directly used by a query
        '
        ' vars passed:
        '   QueryInfo - query info class for query to find
        '
        ' returns:
        '   Pcm.NoErrors - no errors
        '   Pcm.deNoQueryName - no query name
        '   Pcm.deNoQueryText - no query text
        '   Pcm.teOtherErr - something went wrong

        Try
            Dim WorkingQueryName As String = QueryInfo.QueryName.ToUpper
            Dim QueryToFindName As String
            For i As Integer = 0 To AccessQueries.Rows.Count - 1                                    ' for each current query
                QueryToFindName = GcrcvAs(AccessQueries.Rows(i), cnQueryName, "")                   ' get query name to find
                If QueryToFindName = "" Then                                                        ' if no query name
                    MessageBox.Show("No query name to find when sorting queries.",
                                    etSortingQueries, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return Pcm.deNoQueryName                                                        ' return error
                End If
                If QueryToFindName.ToUpper <> WorkingQueryName Then                                 ' if not the working query
                    If QueryInfo.QueryDefText.ToUpper.Contains(QueryToFindName.ToUpper) Then        ' if query to find used in query def text
                        QueryInfo.QueriesUsed.Add(QueryToFindName)                                  ' add query name to collection of query names
                    End If
                End If
            Next

            Return Pcm.NoErrors         ' if got here, then return no errors
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.GetDirectQueriesUsed.  " & ex.Message)
            Return Pcm.teOtherErr
        End Try
    End Function

    Private Function GetReferencedQueriesUsed(AllQueryInfos() As QueryInfoClass) As Integer

        ' gets all referenced queries used
        '
        ' vars passed:
        '   AllQueryInfos - list of all query infos
        '
        ' returns:
        '   Pcm.NoErrors - no errors
        '   Pcm.teOtherErr - something went wrong

        Try
            ' need to loop thru the list n-1 times so all referenced queries can be found.
            ' for example, if query a uses query b, query b uses query c, and query c uses query d
            ' looping through this only once would have query c as used by query a, but not query d
            For i As Integer = 0 To AllQueryInfos.Count - 2                             ' loop through the list n-1 times
                For Each aQi As QueryInfoClass In AllQueryInfos                         ' for each query info in list of query infos
                    For Each bQi As QueryInfoClass In AllQueryInfos                     ' for each query info in list of query infos
                        If aQi.QueryName <> bQi.QueryName Then                          ' if not the same query name
                            If bQi.QueriesUsed.IndexOf(aQi.QueryName) >= 0 Then         ' if query a is used by query b
                                For a As Integer = 0 To aQi.QueriesUsed.Count - 1       ' for each of query a's used query
                                    If bQi.QueriesUsed.IndexOf(aQi.QueriesUsed(a)) = -1 Then ' if query a's used query not already used by b
                                        bQi.QueriesUsed.Add(aQi.QueriesUsed(a))         ' add query a's used query to query b
                                    End If
                                Next
                            End If
                        End If
                    Next
                Next
            Next
            Return Pcm.NoErrors         ' if got here, then return no errors
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.GetReferencedQueriesUsed.  " & ex.Message)
            Return Pcm.teOtherErr
        End Try
    End Function

    Private Function SortQueryAfter(QueryInfo As QueryInfoClass,
                                    SortedQueryInfos() As QueryInfoClass) As Integer

        ' find the sort order of the query in SortedQueryInfos QueryInfo must sort after
        '
        ' vars passed:
        '   QueryInfo - query with used query names to find in SortedQueryInfos
        '   SortedQueryInfos - list of sorted queries, sorted by creation order
        '
        ' returns:
        '   >= 0 - sort the query after this value
        '   -1 - do not adjust the sort order

        Try
            Dim SortAfter As Integer = -1
            For i As Integer = 0 To QueryInfo.QueriesUsed.Count - 1                     ' for each query used
                For s As Integer = 0 To SortedQueryInfos.Count - 1                      ' for each query in sorted list
                    If QueryInfo.QueriesUsed(i) = SortedQueryInfos(s).QueryName Then    ' if found used query name in sorted list
                        If (SortAfter < SortedQueryInfos(s).SortOrder) Then             ' if used query has higher sort order 
                            SortAfter = SortedQueryInfos(s).SortOrder                   ' QueryInfo must sort after found query name
                        End If
                        Exit For                                                        ' exit for s loop
                    End If
                Next
            Next

            Return SortAfter            ' return the sort after value
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.SortQueryAfter.  " & ex.Message)
            Return -1
        End Try
    End Function

    Private Function SortQueries() As Integer

        ' sorts the AccessQueries table into creation order.  If query a uses query b, then query b must be copied first
        '
        ' returns:
        '   Pcm.NoErrors - no errors
        '   Pcm.deNoQueryName - no query name
        '   Pcm.deNoQueryText - no query text
        '   Pcm.teOtherErr - something went wrong

        Try
            Dim QueryDefText As String
            Dim QueryName As String
            Dim Err As Integer
            Dim AllQueriesInfo(AccessQueries.Rows.Count - 1) As QueryInfoClass              ' create list of query info for all queries
            Dim a As Integer = 0
            For Each CurQueryRow As DataRow In AccessQueries.Rows
                QueryName = GcrcvAs(CurQueryRow, cnQueryName, String.Empty)                 ' get the query name 
                If QueryName = String.Empty Then                                            ' if no query name
                    MessageBox.Show("No query name when sorting queries.",
                                    etSortingQueries, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return Pcm.deSortQuery                                                  ' return error
                End If
                QueryDefText = GcrcvAs(CurQueryRow, cnQueryDef, String.Empty)               ' get the query definition text
                If QueryDefText = String.Empty Then                                         ' if no query definition text
                    MessageBox.Show("No query definition text when sorting queries.",
                                    etCreatingQuery, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return Pcm.deCreateQuery                                                ' return error
                End If
                AllQueriesInfo(a) = New QueryInfoClass(QueryName, QueryDefText)             ' create new query info
                AllQueriesInfo(a).SortOrder = a                                             ' set the initial sort order
                a += 1                                                                      ' increment AllQueriesInfo index 
            Next

            For Each QueryInfo As QueryInfoClass In AllQueriesInfo                          ' for each query info
                Err = GetDirectQueriesUsed(QueryInfo)                                       ' get directly used queries
                If Err <> Pcm.NoErrors Then                                                 ' if got an error
                    Return Err                                                              ' return the error
                End If
            Next

            Err = GetReferencedQueriesUsed(AllQueriesInfo)                                  ' get all referenced queries used
            If Err <> Pcm.NoErrors Then                                                     ' if got an error
                Return Err                                                                  ' return the error
            End If

            Dim SortAfter As Integer = 0
            For i As Integer = 0 To AllQueriesInfo.Count - 1                                ' for each query in AllQueriesInfos
                SortAfter = SortQueryAfter(AllQueriesInfo(i), AllQueriesInfo)               ' get the order value to sort after
                If SortAfter >= 0 Then                                                      ' if got a sort after                      
                    For s As Integer = 0 To AllQueriesInfo.Count - 1                        ' for all sorted query infos
                        If AllQueriesInfo(s).SortOrder > SortAfter Then                     ' if sort order more than sort after
                            AllQueriesInfo(s).SortOrder += 1                                ' increment sort order for query
                        End If
                    Next
                    AllQueriesInfo(i).SortOrder = SortAfter + 1                             ' set sort order for query
                End If
            Next

            Dim SortedQueryRow As DataRow
            Dim CurQueriesView As DataView = New DataView(AccessQueries)                    ' create dataview so can search on query name
            For i As Integer = 0 To AllQueriesInfo.Count - 1                                ' for each sorted query
                ' find the query
                SortedQueryRow = PcdbModule.FindRowUsingDataView(CurQueriesView, AllQueriesInfo(i).QueryName, cnQueryName)
                If SortedQueryRow Is Nothing Then                                           ' if did not find the query
                    Return Pcm.deSortQuery                                                  ' return error
                End If
                SortedQueryRow(cnQueryOrder) = AllQueriesInfo(i).SortOrder                  ' set the sort order in CurQuery table
            Next

            Return Pcm.NoErrors             ' if got here, then return no errors
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.SortQueries.  " & ex.Message)
            Return Pcm.teOtherErr
        End Try
    End Function

#End Region

#Region " Update Table "

    Private Function UpdateTable(Table As DataTable,
                                 MySqlConn As MySqlConnection,
                                 MySqlDataAdapt As MySqlDataAdapter,
                                 Optional DesiredRowCount As Integer = -1) As Integer

        ' updates a table to a MySql database
        '
        ' vars passed:
        '   Table - table to updated
        '   MySqlConn - MySql database connection
        '   MySqlDataAdapt - MySql data adapter to use for the update
        '   DesiredRowCount - # rows that should be updated, pass -1 to ignore
        '
        ' returns:
        '   >= 0 - # rows updated
        '   Pcm.teWrongRowCountErr - updated wrong # of rows
        '   Pcm.teOtherErr - something went wrong

        Try
            MySqlConn.Open()
            Dim RowCount As Integer = MySqlDataAdapt.Update(Table)
            If DesiredRowCount >= 0 AndAlso RowCount <> DesiredRowCount Then
                CancelBackgroundWorker()
                MessageBox.Show(String.Format("Row Counts do not match.  Table ""{0}"" has {1} rows, but MySql updated {2} rows.",
                                              Table.TableName, DesiredRowCount, RowCount),
                                etRowCount, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return Pcm.teWrongRowCountErr                                                           ' return error
            End If
            Return RowCount
        Catch ex As Exception
            CancelBackgroundWorker()
            PcdbModule.ShowOtherErrorMessage("MainForm.UpdateTable.  " & ex.Message)
            Return Pcm.teOtherErr
        Finally
            Try
                MySqlConn.Close()
            Catch ex As Exception
                CancelBackgroundWorker()
                PcdbModule.ShowOtherErrorMessage("MainForm.UpdateTable.  " & ex.Message)
            End Try
        End Try
    End Function

#End Region

#Region " Verify Query "

#Region " Verify Proc "

    Private Function GetParamValue(ProcName As String,
                                   ParamName As String,
                                   VerifyNum As Integer) As Object

        ' gets parameter values used for query verification 
        '
        ' vars passed:
        '   ProcName - name of procedure being verified
        '   ParamName - name of parameter getting value
        '   VerifyNum - verification #
        '
        ' returns:
        '   Object - parameter value
        '   nothing - something went wrong

        ParamName = ParamName.ToUpper
        Select Case ProcName
            Case "TimecardJobsToBill"
                Select Case ParamName
                    Case "FOM"
                        Select Case VerifyNum
                            Case 1
                                Return #01/01/2019#
                            Case 2
                                Return #01/01/2018#
                        End Select
                    Case "FONM"
                        Select Case VerifyNum
                            Case 1
                                Return #02/01/2019#
                            Case 2
                                Return #02/01/2018#
                        End Select
                    Case "BT"
                        Select Case VerifyNum
                            Case 1
                                Return "Fixed Fee"
                            Case 2
                                Return "Hourly"
                        End Select
                End Select
            Case "WorkDoneJobsToBill"
                Select Case ParamName
                    Case "FOM"
                        Select Case VerifyNum
                            Case 1
                                Return #07/01/2008#
                            Case 2
                                Return #07/01/2007#
                        End Select
                    Case "FONM"
                        Select Case VerifyNum
                            Case 1
                                Return #08/01/2008#
                            Case 2
                                Return #08/01/2007#
                        End Select
                    Case "BT"
                        Select Case VerifyNum
                            Case 1
                                Return "Fixed Fee"
                            Case 2
                                Return "Hourly"
                        End Select
                End Select
        End Select
        Return Nothing
    End Function

    Private Sub PopulateParams(ProcName As String,
                               aParams As List(Of EaAccess.AccessParam),
                               mParams As List(Of MySqlParam),
                               verifyNum As Integer)

        ' populates parameters for MySql procedure verification
        '
        ' vars passed;
        ' vars passed:
        '   ProcName - name of procedure being verified
        '   aParams - list of access parameters
        '   mParams - list of mySql parameters
        '   VerifyNum - verification #

        For Each ap In aParams
            ap.Value = GetParamValue(ProcName, ap.Name, verifyNum)
        Next
        For Each mp In mParams
            mp.Value = GetParamValue(ProcName, mp.Name, verifyNum)
        Next
    End Sub

    Private Function VerifyProc(worker As System.ComponentModel.BackgroundWorker,
                                progress As PcdbToMySqlProgress,
                                ProcName As String,
                                AccessSqlText As String,
                                MySqlSqlText As String,
                                MySqlConn As MySql.Data.MySqlClient.MySqlConnection) As Boolean

        ' verifies a procedure in the MySql database
        ' creates a 2nd temp procedure, hopefully identical to the procedure to verify.  
        ' if the sql command text of the two procedures are identical, then procedure is verified 
        '
        ' 1) make sure mySql proc exists
        ' 2) create access params list for proc
        ' 3) create mySql params list for proc
        ' 4) verify all access params are in mysql params        
        ' 5) create access oleDbDataAdapter
        ' 6) create MySqlDbDataAdapter
        ' for each verify try
        ' 7) populate access and mysql params
        ' 8) fill access table using access oleDbDataAdapter
        ' 9) fill MySql table data using MySql MySqlDbDataAdapter
        ' 10) verify table data match
        '
        ' vars passed:
        '   worker - background worker
        '   progress - progress values class 
        '   MySqlConn - connection to use
        '   ProcName - procedure to verify
        '   sqlText - procedure command
        '   Params - list of procedure parameters
        '
        ' returns:
        '   TRUE - procedure verified
        '   FALSE - 
        '     - procedure not verified
        '     - something went wrong

        Const AccessTempTableName As String = "TempA"
        Const MySqlTempTableName As String = "TempM"

        Try
            ShowProgress(worker, progress, String.Format("Procedure {0}: verifying...", ProcName))
            ' 1) make sure mySql proc exists
            If MySqlTools.ProcExits(MySqlConn, ProcName) <> MySqlTools.ExistsTypes.Yes Then     ' if cannot find procedure
                CancelBackgroundWorker()
                MessageBox.Show(String.Format("Procedure not found in MySql database.{0}Procedure Name: {1}{0}Database: {2}", vbCrLf, ProcName, MySqlConn.Database),
                                etVerifyQuery, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False                                                                    ' return not verified
            End If
            ShowProgress(worker, progress, String.Format("Setup to verify procedure: {0}", ProcName), 15)

            ' 2) create access params list for proc            
            Dim aParams As List(Of EaAccess.AccessParam) = EaAccess.AccessSql.GetProcParams(AccessOleDbConnection, AccessSqlText)
            If aParams Is Nothing Then
                CancelBackgroundWorker()
                MessageBox.Show(String.Format("Could not get list of parameters from Access procedure.{0}Procedure Name: {1}{0}Database: {2}", vbCrLf, ProcName, MySqlConn.Database),
                                etVerifyQuery, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False                                                                    ' return not verified
            End If
            ShowProgress(worker, progress, 30)

            ' 3) create mySql params list for proc
            Dim mParams As List(Of MySqlParam) = MySqlTools.GetProcParams(MySqlSqlText)
            If mParams Is Nothing Then
                CancelBackgroundWorker()
                MessageBox.Show(String.Format("Could not get list of parameters from MySql procedure.{0}Procedure Name: {1}{0}Database: {2}", vbCrLf, ProcName, MySqlConn.Database),
                                etVerifyQuery, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False                                                                    ' return not verified
            End If
            ShowProgress(worker, progress, 50)

            ' 4) verify all access params are in mysql params
            If aParams.Count <> mParams.Count Then                                              ' if # params don't match
                Return False                                                                    ' return not verified
            Else                                                                                ' else # params the same 
                ' verify list of params are the same
                Dim foundList As New List(Of String)                                            ' list of found param names
                Dim paramsToFind As New List(Of MySqlParam)
                For Each p In mParams
                    paramsToFind.Add(New MySqlParam(p))
                Next
                For i As Integer = 0 To aParams.Count - 1                                       ' for each access param
                    For j As Integer = 0 To paramsToFind.Count - 1                              ' for each param to find
                        If paramsToFind(j).Name = aParams(i).Name Then                          ' if param names match
                            foundList.Add(aParams(i).Name)                                      ' add param name to found list
                            paramsToFind.RemoveAt(j)                                            ' remove verify param from list
                            Exit For                                                            ' exit for j
                        End If
                    Next
                Next
                If foundList.Count <> aParams.Count Then                                        ' if didn't find all params
                    CancelBackgroundWorker()
                    MessageBox.Show(String.Format("Could not find all parameters in MySql procedure.{0}Procedure Name: {1}{0}Database: {2}", vbCrLf, ProcName, MySqlConn.Database),
                                etVerifyQuery, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return False                                                                ' return not verified
                End If
            End If
            ShowProgress(worker, progress, 80)

            ' 5) create access oleDbDataAdapter
            Dim oleDbDataAdapt As OleDb.OleDbDataAdapter = CreateOleDbDataAdapt(AccessOleDbConnection, AccessTempTableName, True, AccessSqlText)
            If oleDbDataAdapt Is Nothing Then
                CancelBackgroundWorker()
                MessageBox.Show(String.Format("Could not create OleDbDataAdapt.{0}Procedure Name: {1}{0}Database: {2}", vbCrLf, ProcName, AccessOleDbConnection.Database),
                                etVerifyQuery, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False                                                                                ' return not verified
            End If
            ShowProgress(worker, progress, 90)

            ' 6) create MySqlDbDataAdapter
            'Dim MySqlDataAdapt As MySqlDataAdapter = EaMySql.CreateMySqlDataAdapt(MySqlPcdbConn, MySqlTempTableName, True, MySqlSqlText)
            Dim MySqlDataAdapt As MySqlDataAdapter = MySqlTools.CreateMySqlDataAdapt(MySqlPcdbConn, ProcName)
            If MySqlDataAdapt Is Nothing Then
                CancelBackgroundWorker()
                MessageBox.Show(String.Format("Could not create MySqlDataAdapt.{0}Procedure Name: {1}{0}Database: {2}", vbCrLf, ProcName, MySqlPcdbConn.Database),
                                etVerifyQuery, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False                                                                                ' return not verified
            End If

            progress.TotalStep()
            ShowProgress(worker, progress, 100)

            Dim tempTableA As New DataTable(AccessTempTableName)
            Dim tempTableM As New DataTable(MySqlTempTableName)
            Dim rowCountA As Integer
            Dim rowCountM As Integer
            Dim err As Integer
            For i As Integer = 1 To 2
                ShowProgress(worker, progress, String.Format("Filling data for procedure: {0}", ProcName))

                ' 7) populate access and mysql params
                PopulateParams(ProcName, aParams, mParams, i)

                ' 8) fill access table using access oleDbDataAdapter
                rowCountA = FillTable(tempTableA, oleDbDataAdapt, aParams)                                  ' fill access view
                If rowCountA < 0 Then
                    CancelBackgroundWorker()
                    MessageBox.Show(String.Format("Error filling table with Access Procedure.{0}Procedure Name: {1}{0}Database: {2}", vbCrLf, ProcName, AccessOleDbConnection.Database),
                                    etVerifyQuery, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return False                                                                            ' return not verified
                End If
                ShowProgress(worker, progress, 10)

                ' 9) fill MySql table data using MySql MySqlDbDataAdapter
                rowCountM = MySqlTools.FillTable(tempTableM, MySqlDataAdapt, mParams)
                If rowCountM < 0 Then                                                                       ' if error filling table
                    CancelBackgroundWorker()
                    MessageBox.Show(String.Format("Error filling table with MySql Procedure.{0}Procedure Name: {1}{0}Database: {2}", vbCrLf, ProcName, MySqlPcdbConn.Database),
                                    etVerifyQuery, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return False
                End If
                ShowProgress(worker, progress, 55)

                ' 10) verify table data match
                If rowCountM <> rowCountA Then                                                              ' if did not fill correct # of rows
                    CancelBackgroundWorker()
                    MessageBox.Show(String.Format("Filled row count <> desired row count for Procedure: {1}{0}Filled row count: {2}{0}Desired row count: {3}",
                                                  vbCrLf, ProcName, rowCountM, rowCountA),
                                    etVerifyQuery, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return False
                End If
                progress.TotalStep()
                ShowProgress(worker, progress, 100)

                err = VerifyProcOrViewData(worker, progress, ProcName, tempTableA, tempTableM)
                If err < 0 Then
                    CancelBackgroundWorker()
                    Return False
                End If
                progress.TotalStep()
            Next

            Return True                                                                                     ' if got here, then verified
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.VerifyProc.  " & ex.Message)
            Return False
        Finally
            progress.TotalStep()
            ShowProgress(worker, progress, 100)
        End Try
    End Function

#End Region

#Region " Verify ProcOrView Data"

    Private Function VerifyProcOrViewData(worker As System.ComponentModel.BackgroundWorker,
                                          progress As PcdbToMySqlProgress,
                                          PvName As String,
                                          AccessTable As DataTable,
                                          MySqlTable As DataTable) As Integer

        ' verifies the data in the access table and MySql table match.  tables have no primary key
        ' NOTE: tabled are filled via the matching Access and MySql procedures/views.
        '
        ' vars passed:
        '   worker - background worker
        '   progress - progress values 
        '   TableName - name of procedure or view to verify
        '   AccessTable - temp Access table filled from access procedure/view 
        '   MySqlTable - temp MySql table filled from MySql procedure/view 
        '
        ' returns:
        '   Pcm.NoErrors - no errors, table data verified
        '   Pcm.teRowCountErr - selected row counts do not match
        '   Pcm.teOtherErr - something went wrong

        Try
            ShowProgress(worker, progress, "Sorting view results for data verification...", 0)

            Dim sortStr As String = String.Empty
            For Each col As DataColumn In AccessTable.Columns
                sortStr &= String.Format("{0} ASC,", col.ColumnName)
            Next

            sortStr = sortStr.Substring(0, sortStr.Length - 1)
            AccessTable.DefaultView.Sort = sortStr
            ShowProgress(worker, progress, 50)

            MySqlTable.DefaultView.Sort = sortStr
            ShowProgress(worker, progress, 100)

            For r As Integer = 0 To AccessTable.DefaultView.Count - 1
                If PcdbModule.ConfirmRowsMatch(AccessTable.DefaultView(r).Row, MySqlTable.DefaultView(r).Row) <> Pcm.NoErrors Then    ' if row column values do not match
                    Return Pcm.deVerifyTableData                                                                ' return error
                End If
                progress.ActionPosition = progress.ActionPosition + 1
                ShowProgress(worker, progress, String.Format("View {0}: Verifying view data.  Row {1} of {2}...",
                                                             PvName, r, AccessTable.Rows.Count), progress.ActionPosition)
                r += 1                                                                          ' increment row counter
            Next
            Return Pcm.NoErrors
        Catch ex As Exception
            CancelBackgroundWorker()
            PcdbModule.ShowOtherErrorMessage("MainForm.VerifyTableDataNoKey.  " & ex.Message)
            Return Pcm.teOtherErr
        End Try
    End Function


#End Region

#Region " Verify View "

    Private Function VerifyView(worker As System.ComponentModel.BackgroundWorker,
                                progress As PcdbToMySqlProgress,
                                ViewName As String,
                                AccessSqlText As String,
                                MySqlSqlText As String,
                                MySqlConn As MySql.Data.MySqlClient.MySqlConnection) As Boolean

        ' verifies a view in the MySql database
        ' creates a 2nd temp view, hopefully identical to the view to verify.  
        ' if the sql command text of the two views are identical, then view is verified 
        '
        ' 1) make sure mySql view exists
        ' 2) create access oleDbDataAdapter
        ' 3) fill access table using access oleDbDataAdapter
        ' 4) create MySqlDbDataAdapter
        ' 5) load MySql table data using MySql MySqlDbDataAdapter
        ' 6) verify table data match
        '
        ' vars passed:
        '   worker - background worker
        '   progress - progress values class 
        '   ViewName - view to verify
        '   AccessSqlText - view command (SQL Query text for Access)
        '   MySqlSqlText - view command (SQL Query text for MySql)
        '   MySqlConn - connection to use
        '
        ' returns:
        '   TRUE - view verified
        '   FALSE - 
        '     - view not verified
        '     - something went wrong

        Const AccessTempTableName As String = "TempA"
        Const MySqlTempTableName As String = "TempM"

        Try
            ShowProgress(worker, progress, String.Format("View {0}: verifying...", ViewName))
            progress.TotalStep()                                                                            ' one more total step done

            ' 1) make sure mySql view exists
            If MySqlTools.ViewExits(MySqlConn, ViewName) <> MySqlTools.ExistsTypes.Yes Then                 ' if cannot find view
                CancelBackgroundWorker()
                MessageBox.Show(String.Format("View not found in MySql database.{0}View Name: {1}{0}Database: {2}", vbCrLf, ViewName, MySqlConn.Database),
                                etVerifyQuery, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False                                                                                ' return not verified
            End If
            ShowProgress(worker, progress, 5)

            ' 2) create access oleDbDataAdapter
            Dim oleDbDataAdapt As OleDb.OleDbDataAdapter = CreateOleDbDataAdapt(AccessOleDbConnection, AccessTempTableName, True, AccessSqlText)
            If oleDbDataAdapt Is Nothing Then
                CancelBackgroundWorker()
                MessageBox.Show(String.Format("Could not create OleDbDataAdapt.{0}View Name: {1}{0}Database: {2}", vbCrLf, ViewName, AccessOleDbConnection.Database),
                                etVerifyQuery, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False                                                                                ' return not verified
            End If
            ShowProgress(worker, progress, 10)

            ' 3) fill access table using access oleDbDataAdapter
            Dim tempTableA As New DataTable(AccessTempTableName)
            Dim rowCountA As Integer = FillTable(tempTableA, oleDbDataAdapt)                                ' fill access view
            If rowCountA < 0 Then
                CancelBackgroundWorker()
                MessageBox.Show(String.Format("Error filling table with Access View.{0}View Name: {1}{0}Database: {2}", vbCrLf, ViewName, AccessOleDbConnection.Database),
                                etVerifyQuery, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End If
            ShowProgress(worker, progress, 50)

            ' 4) create MySqlDbDataAdapter
            Dim MySqlDataAdapt As MySqlDataAdapter = MySqlTools.CreateMySqlDataAdapt(MySqlPcdbConn, MySqlTempTableName, True, MySqlSqlText)
            If MySqlDataAdapt Is Nothing Then
                CancelBackgroundWorker()
                MessageBox.Show(String.Format("Could not create MySqlDataAdapt.{0}View Name: {1}{0}Database: {2}", vbCrLf, ViewName, MySqlPcdbConn.Database),
                                etVerifyQuery, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False                                                                                ' return not verified
            End If
            ShowProgress(worker, progress, 60)

            ' 5) load MySql table data using MySql MySqlDbDataAdapter
            Dim tempTableM As New DataTable(MySqlTempTableName)
            Dim rowCountM = MySqlTools.FillTable(tempTableM, MySqlDataAdapt)
            If rowCountM < 0 Then                                                                           ' if error filling table
                CancelBackgroundWorker()
                MessageBox.Show(String.Format("Error filling table with MySql View.{0}View Name: {1}{0}Database: {2}", vbCrLf, ViewName, MySqlPcdbConn.Database),
                                etVerifyQuery, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End If
            If rowCountM <> rowCountA Then                                                                  ' if did not fill correct # of rows
                CancelBackgroundWorker()
                MessageBox.Show(String.Format("Filled row count <> desired row count for View: {1}{0}Filled row count: {2}{0}Desired row count: {3}",
                                              vbCrLf, ViewName, rowCountM, rowCountA),
                                etVerifyQuery, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End If
            progress.TotalStep()
            ShowProgress(worker, progress, 100)

            ' 6) verify view result data match
            Dim err As Integer = VerifyProcOrViewData(worker, progress, ViewName, tempTableA, tempTableM)   ' verify result data
            If err <> Pcm.NoErrors Then                                                                     ' if error verifying table
                Return False                                                                                ' return error
            End If
            progress.TotalStep()                                                                            ' one more total step done

            Return True                                                                                     ' if got here, view verified
        Catch ex As Exception
            CancelBackgroundWorker()
            PcdbModule.ShowOtherErrorMessage("MainForm.VerifyView.  " & ex.Message)
            Return False
        Finally
            progress.TotalStep()
            ShowProgress(worker, progress, 100)
        End Try
    End Function

#End Region

#End Region

#Region " Verify Table "

    Private Function GetKeyValuesAsText(KeyValues() As Object,
                                        KeyColNames() As String) As String

        ' gets the key values text to display in an error message
        '
        ' vars passed:
        '   KeyValues() - key values
        '   KeyColNames() - key column names
        '
        ' returns:
        '   string with key columns names and their column values 
        '   "" - something went wrong

        Try
            Dim KeyText As String = String.Empty
            For i As Integer = 0 To KeyValues.Count - 1
                KeyText &= String.Format("   Column: {0}; Value: {1}", KeyColNames(i), KeyValues(i))
                If i > 0 AndAlso i < KeyValues.Count - 1 Then
                    KeyText &= vbCrLf
                End If
            Next
            Return KeyText
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.GetKeyValuesAsText.  " & ex.Message)
            Return ""
        End Try
    End Function

    Private Function VerifyTableData(worker As System.ComponentModel.BackgroundWorker,
                                     progress As PcdbToMySqlProgress,
                                     TableName As String,
                                     AccessTable As DataTable,
                                     MySqlTable As DataTable) As Integer

        ' verifies the data in the AccessTable table and MySqlTable table match
        '
        ' vars passed:
        '   worker - background worker
        '   progress - progress values 
        '   TableName - name of table to verify
        '   AccessTable - Access table
        '   MySqlTable - MySql table
        '
        ' returns:
        '   Pcm.NoErrors - no errors, table data verified
        '   Pcm.teRowCountErr - row counts do not match between tables
        '   Pcm.deVerifyTableData - data not verified
        '   Pcm.teOtherErr - something went wrong

        Try
            If AccessTable.Rows.Count <> MySqlTable.Rows.Count Then                                 ' if row counts do not match
                CancelBackgroundWorker()
                MessageBox.Show(String.Format("Row counts do not match.{0}Table: {1}{0}Access Table Row Count: {2}{0}MySql Table Row Count: {3}",
                                              vbCrLf, TableName, AccessTable.Rows.Count, MySqlTable.Rows.Count),
                                etVerifyTableData, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return Pcm.teRowCountErr                                                            ' return error
            End If
            If AccessTable.Rows.Count = 0 Then                                                      ' row counts match, and no rows to check
                Return Pcm.NoErrors                                                                 ' return no errors now
            End If

            ' no need to check if PrimaryKey.Count match.  that was done in verify table structure
            If AccessTable.PrimaryKey.Count = 0 AndAlso MySqlTable.PrimaryKey.Count = 0 Then        ' if tables not keyed
                Return VerifyTableDataNoKey(worker, progress, TableName, AccessTable, MySqlTable)   ' verify non keyed data
            Else                                                                                    ' else table keyed
                Return VerifyTableDataHasKey(worker, progress, TableName, AccessTable, MySqlTable)  ' verify keyed data
            End If
        Catch ex As Exception
            CancelBackgroundWorker()
            PcdbModule.ShowOtherErrorMessage("MainForm.VerifyTableData.  " & ex.Message)
            Return Pcm.teOtherErr
        End Try
    End Function

    Private Function VerifyTableDataHasKey(worker As System.ComponentModel.BackgroundWorker,
                                           progress As PcdbToMySqlProgress,
                                           TableName As String,
                                           AccessTable As DataTable,
                                           MySqlTable As DataTable) As Integer

        ' verifies the data in the current table and refreshed table match.  tables have primary key(s)
        '
        ' vars passed:
        '   worker - background worker
        '   progress - progress values 
        '   TableName - name of table to verify
        '   AccessTable - Access table
        '   MySqlTable - MySql table
        '
        ' returns:
        '   Pcm.NoErrors - no errors, table data verified
        '   Pcm.deVerifyTableData - data not verified
        '   Pcm.teOtherErr - something went wrong

        Try
            Dim MySqlRow As DataRow
            Dim KeyValues(AccessTable.PrimaryKey.Count - 1) As Object                           ' key values
            Dim KeyColNames(AccessTable.PrimaryKey.Count - 1) As String                         ' key column names
            For i As Integer = 0 To AccessTable.PrimaryKey.Count - 1                            ' for each key column
                KeyColNames(i) = AccessTable.PrimaryKey(i).ColumnName                           ' get the key column name
            Next
            Dim r As Integer = 1                                                                ' row counter
            For Each AccessRow As DataRow In AccessTable.Rows                                   ' for each row
                If AccessTable.PrimaryKey.Count = 1 Then                                        ' if only 1 key column
                    KeyValues(0) = AccessRow(KeyColNames(0))                                    ' get the key value
                    MySqlRow = PcdbModule.FindRow(MySqlTable, KeyValues(0))                     ' find the matching MySql row
                Else                                                                            ' else multi column key
                    For i As Integer = 0 To AccessTable.PrimaryKey.Count - 1                    ' for each key column
                        KeyValues(i) = AccessRow(KeyColNames(i))                                ' get the key value
                    Next
                    MySqlRow = PcdbModule.FindRow(MySqlTable, KeyValues, KeyColNames)           ' find the matching MySql row
                End If
                If MySqlRow Is Nothing Then                                                     ' if did not find matching refreshed row
                    CancelBackgroundWorker()
                    Dim KeyValuesText As String = GetKeyValuesAsText(KeyValues, KeyColNames)
                    MessageBox.Show(String.Format("Row not found in refreshed table.{0}Table Name: {1}{0}Key Values{0}{2}", vbCrLf, TableName, KeyValuesText),
                                    etVerifyTableData, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return Pcm.deVerifyTableData                                                ' return error
                End If
                If PcdbModule.ConfirmRowsMatch(AccessRow, MySqlRow) <> Pcm.NoErrors Then        ' if row column values do not match
                    Return Pcm.deVerifyTableData                                                ' return error
                End If
                progress.ActionPosition = progress.ActionPosition + 1
                ShowProgress(worker, progress, String.Format("Table {0}: Verifying table data.  Row {1} of {2}...", TableName, r, AccessTable.Rows.Count), progress.ActionPosition)
                r += 1                                                                          ' increment row counter
            Next

            Return Pcm.NoErrors     ' if got here, then all data in the table has been verified, return no errors
        Catch ex As Exception
            CancelBackgroundWorker()
            PcdbModule.ShowOtherErrorMessage("MainForm.VerifyTableDataHasKey.  " & ex.Message)
            Return Pcm.teOtherErr
        End Try
    End Function

    Private Function VerifyTableDataNoKey(worker As System.ComponentModel.BackgroundWorker,
                                          progress As PcdbToMySqlProgress,
                                          TableName As String,
                                          AccessTable As DataTable,
                                          MySqlTable As DataTable) As Integer

        ' verifies the data in the current table and refreshed table match.  tables have no primary key
        '
        ' vars passed:
        '   worker - background worker
        '   progress - progress values 
        '   TableName - name of table to verify
        '   AccessTable - Access table
        '   MySqlTable - MySql table
        '
        ' returns:
        '   Pcm.NoErrors - no errors, table data verified
        '   Pcm.teRowCountErr - selected row counts do not match
        '   Pcm.teOtherErr - something went wrong

        Try
            Dim SelectDistinctSql As String = String.Format("SELECT DISTINCT {0}.* FROM {0}", TableName)
            Dim DistinctOleDbdataAspt As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(SelectDistinctSql, AccessOleDbConnection)
            Dim DistinctTable As DataTable = New DataTable
            Dim RowCount As Integer = FillTable(DistinctTable, DistinctOleDbdataAspt)       ' load distinct rows for Access table
            If RowCount < 0 Then                                                            ' if error loading data
                Return RowCount                                                             ' return error
            End If
            ' update action progress bar maximum to DistinctTable.Rows.Count
            'SetAction(String.Format("Table {0}: Verifying table data...", TableName), DistinctTable.Rows.Count)
            Dim SelectStr As String                                                         ' row selection string
            Dim c As Integer = 0                                                            ' column counter
            Dim accessRows() As DataRow                                                     ' selected access rows
            Dim mySqlRows() As DataRow                                                      ' selected mysql rows
            Dim r As Integer = 1                                                            ' row counter
            For Each dRow As DataRow In DistinctTable.Rows                                  ' for each distinct row
                SelectStr = ""                                                              ' reset the selection string
                c = 1                                                                       ' reset column counter to 1
                For Each col As DataColumn In DistinctTable.Columns                         ' for each column
                    If c > 1 Then                                                           ' if not the first column
                        SelectStr &= " AND "                                                ' add " AND " before selection string
                    End If
                    SelectStr &= PcdbModule.FieldSelectString(col.ColumnName, dRow(col))    ' get selection string for column
                    c += 1                                                                  ' increment column counter
                Next

                accessRows = AccessTable.Select(SelectStr)                                  ' select matching rows from current table
                mySqlRows = MySqlTable.Select(SelectStr)                                    ' select matching rows from refreshed table
                ' there is no need to check if the values match.  by definition rows selected 
                ' will have all the column values matching.  just need to check # of rows selected
                If accessRows.Count <> mySqlRows.Count Then                                 ' if # of rows selected do not match
                    CancelBackgroundWorker()
                    MessageBox.Show(String.Format("Data row counts do not match for non keyed table.{0}Table: {1}{0}Current Table Selected Row Count: {2}{0}Refreshed Table Selected Row Count: {3}{0}Selection String: {4}",
                                                  vbCrLf, TableName, accessRows.Count, mySqlRows.Count, SelectStr),
                                    etVerifyTableData, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return Pcm.teRowCountErr
                End If
                progress.ActionStep()
                ShowProgress(worker, progress, String.Format("Table {0}: Verifying table data.  Row {1} of {2}...", TableName, r, AccessTable.Rows.Count))

                'ActionProgressBarStep()                                                     ' update action progress bar
                'ActionTextEdit.Text = String.Format("Table {0}: Verifying table data.  Distinct Row {1} of {2}...", TableName, r, DistinctTable.Rows.Count)
                'ActionTextEdit.Update()                                                     ' update action text
                r += 1                                                                      ' increment row counter
            Next

            Return Pcm.NoErrors
        Catch ex As Exception
            CancelBackgroundWorker()
            PcdbModule.ShowOtherErrorMessage("MainForm.VerifyTableDataNoKey.  " & ex.Message)
            Return Pcm.teOtherErr
        End Try
    End Function

#End Region

#Region " Pre Copy "

    Private Function PreCopy(TotalSteps As Integer) As Integer

        ' does the pre copy setup. 
        '   verify settings
        '   set up progress
        '
        ' vars passed:
        '   TotalSteps - total number of steps for total progress bar
        '
        ' returns:
        '   NoErrors - no errors
        '   Pcm.teConnectionErr - error in access sql connection settings
        '   EaSql.Ea.sqlErr - error in mysql connection settings
        '   OtherExErr - something went wrong  

        Try
            If Not ValidAccessDatabase() Then                                       ' if not valid access database settings
                Return Pcm.teConnectionErr                                          ' return error
            End If
            If Not ValidMySqlSettings() Then                                        ' if not valid MySql settings
                Return EaSql.Ea.sqlErr                                              ' return error
            End If

            ShowActions(True)                                                       ' show all the action controls
            StartStopTimer(True)

            TotalProgressBarControl.Properties.Maximum = TotalSteps                 ' set # of total steps for progress bar
            TotalProgressBarControl.Position = 0

            Dim err As Integer = MakeAccessAndMySqlConnections()                    ' make the database connections
            If err <> NoErrors Then
                Return err
            End If

            Return NoErrors
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.PreCopy.  " & ex.Message)
            Return OtherExErr
        End Try
    End Function

#End Region

#Region " Main Form Controls "

#Region " 3.0 buttons "

    Private Sub Non3Point0Button_Click(sender As Object, e As EventArgs) Handles Non3Point0Button.Click

        Try
            Dim Non3TableNames As String = String.Empty
            Dim tableName As String
            For i As Integer = 0 To TablesCheckedListBox.CheckedItemsCount - 1                                  ' for each table to copy
                tableName = CStr(TablesCheckedListBox.CheckedItems(i))                                          ' get table name
                Non3TableNames &= tableName & vbCrLf                                                            ' add table name and new line
            Next
            My.Computer.FileSystem.WriteAllText("C:\Projects 2017\PcdbToSql\PcdbToSql\Non3Tables.txt", Non3TableNames, False) ' write text file
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.Non3Point0Button_Click.  " & ex.Message)
        End Try
    End Sub

    Private Sub Select3Point0Button_Click(sender As Object, e As EventArgs) Handles Select3Point0Button.Click

        Try
            TablesCheckedListBox.CheckAll()                                                     ' check all table items
            Dim Non3TableName As String
            Dim cbIndex As Integer
            For i As Integer = 0 To Non3Point0ListBoxControl.ItemCount - 1                      ' for each item in non 3.0 table names
                Non3TableName = Non3Point0ListBoxControl.Items(i).ToString                      ' get non 3.0 table name
                cbIndex = TablesCheckedListBox.Items.IndexOf(Non3TableName)                     ' find non 3.0 table name is list of all tables
                If cbIndex >= 0 Then                                                            ' if found non 3.0 table name
                    TablesCheckedListBox.Items(cbIndex).CheckState = CheckState.Unchecked       ' uncheck non 3.0 table name in list of all tables
                End If
            Next

            QueriesCheckedListBox.CheckAll()                                                    ' check all queries (views and procedures)
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.Select3Point0Button_Click.  " & ex.Message)
        End Try
    End Sub

#End Region

#Region " AccessButtonEdit "

    Private Sub AccessButtonEdit_ButtonClick(sender As Object, e As DevExpress.XtraEditors.Controls.ButtonPressedEventArgs) Handles AccessButtonEdit.ButtonClick

        Try
            XtraOpenFileDialog1.InitialDirectory = "C:\Projects\PcdbData"
            If XtraOpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                AccessButtonEdit.Text = XtraOpenFileDialog1.FileName

                Dim fInfo As IO.FileInfo = My.Computer.FileSystem.GetFileInfo(AccessButtonEdit.Text)    ' get file info for file
                If Not My.Computer.FileSystem.DirectoryExists(fInfo.DirectoryName) Then                 ' if file does not exist
                    MessageBox.Show(String.Format("Folder {1} not found for {0}.", CurrentLabelControl.Text, fInfo.DirectoryName),
                                    "Folder Not Found", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    AccessButtonEdit.Text = String.Empty                                                ' clear values
                    AccessFolder = String.Empty
                    AccessDatabaseName = String.Empty
                    AccessFullDatabaseName = String.Empty
                    Return
                End If

                ' if got here, the access PCDB database entry is valid, set values
                AccessFolder = fInfo.DirectoryName                          ' set the current database folder name
                AccessDatabaseName = fInfo.Name                             ' set the current database file name 
                AccessFullDatabaseName = fInfo.FullName                     ' set the current database full file name
            End If
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.AccessButtonEdit_ButtonClick.  " & ex.Message)
            Return
        End Try
    End Sub

#End Region

#Region " Copy All "

    Private Function CopyAllToMySql(worker As System.ComponentModel.BackgroundWorker, e As System.ComponentModel.DoWorkEventArgs) As Integer

        Try
            Dim err As Integer = CopyTablesToMySql(worker, e)
            If err < 0 Then
                CancelBackgroundWorker()
                Return err
            End If
            'CopiedTablesLabel.Visible = True

            err = CopyQueriesToMySql(worker, e)
            If err < 0 Then
                CancelBackgroundWorker()
                Return err
            End If
            'CopiedQueriesLabel.Visible = True

            'SuccessLabel.Visible = True
            Return NoErrors
        Catch ex As Exception
            CancelBackgroundWorker()
            PcdbModule.ShowOtherErrorMessage("MainForm.CopyAllMySql.  " & ex.Message)
            Return Pcm.OtherExErr
        End Try
    End Function

    Private Sub CopyAllButton_Click(sender As Object, e As EventArgs) Handles CopyAllButton.Click

        Const StepsPerTable As Integer = 4
        Const SetupSteps As Integer = 6         ' 2 for database, 2 for tables and 2 for queries
        Const StepsPerQuery As Integer = 6
        Const FinalSteps As Integer = 0

        totalSteps = ((TablesCheckedListBox.CheckedItemsCount - 1) * StepsPerTable) +
                     ((QueriesCheckedListBox.CheckedItemsCount - 1) * StepsPerQuery) +
                     SetupSteps + FinalSteps
        If PreCopy(totalSteps) <> NoErrors Then
            Return
        End If

        BackgroundWorker1.RunWorkerAsync(CopyTasks.All)
    End Sub

#End Region

#Region " Copy Queries "

    Private Function CreateQueryTables() As Integer

        ' creates the AccessQueries table and populates it
        ' creates the MySqlQueries table, but DOES NOT populate it
        '
        ' note: this func assumes CurViews and CurProcs have been created and populated
        '
        ' returns:
        '   >= 0 - # of rows in CurQueries table
        '   Pcm.teOtherErr - something went wrong

        Const AccessQueriesTableName As String = "AccessQueries"
        Const MySqlQueriesTableName As String = "MySqlQueries"

        Try
            Dim QueryDefText As String
            Dim QueryName As String

            ' make CurQueries table
            AccessQueries = New DataTable(AccessQueriesTableName)                                       ' create data table
            Dim ColTypeStr As String = GetSystemColumnTypeStr(EaAccess.AccessColumn.ColumnTypes.ctText) ' get string for column creation
            Dim col As DataColumn = New DataColumn(cnQueryName, Type.GetType(ColTypeStr))               ' create QUERY_NAME column
            AccessQueries.Columns.Add(col)                                                              ' add QUERY_NAME column to table
            col = New DataColumn(cnQueryDef, Type.GetType(ColTypeStr))                                  ' create QUERY_DEFINITION column
            AccessQueries.Columns.Add(col)                                                              ' add QUERY_DEFINITION column to table
            ColTypeStr = GetSystemColumnTypeStr(EaAccess.AccessColumn.ColumnTypes.ctInteger)            ' get integer string for column creation
            col = New DataColumn(cnQueryOrder, Type.GetType(ColTypeStr))                                ' create QUERY_ORDER column
            AccessQueries.Columns.Add(col)                                                              ' add QUERY_ORDER column to table
            ColTypeStr = GetSystemColumnTypeStr(EaAccess.AccessColumn.ColumnTypes.ctText)               ' need a text column next
            col = New DataColumn(cnViewOrProc, Type.GetType(ColTypeStr))                                ' create VIEW_OR_PROC column
            AccessQueries.Columns.Add(col)                                                              ' add VIEW_OR_PROC column to table

            ' make MySqlQueries table (no QUERY_ORDER columns for sorting)
            MySqlQueries = New DataTable(MySqlQueriesTableName)                                         ' create data table
            ColTypeStr = GetSystemColumnTypeStr(EaAccess.AccessColumn.ColumnTypes.ctText)               ' get string for column creation
            col = New DataColumn(cnQueryName, Type.GetType(ColTypeStr))                                 ' create QUERY_NAME column
            MySqlQueries.Columns.Add(col)                                                               ' add QUERY_NAME column to table
            col = New DataColumn(cnQueryDef, Type.GetType(ColTypeStr))                                  ' create QUERY_DEFINITION column
            MySqlQueries.Columns.Add(col)                                                               ' add QUERY_DEFINITION column to table

            ' move current view data into queries table
            Dim QueryRow As DataRow
            Dim rows() As DataRow
            Dim RowSelectStr As String
            If AccessViews IsNot Nothing Then                                                           ' if got access views
                For i As Integer = 0 To QueriesCheckedListBox.CheckedItemsCount - 1                     ' for each selected query
                    QueryName = CStr(QueriesCheckedListBox.CheckedItems(i))                             ' get the query name
                    RowSelectStr = PcdbModule.FieldSelectString(cnTableName, QueryName)                 ' set selection string
                    rows = AccessViews.Select(RowSelectStr)                                             ' find row in access queries
                    If rows IsNot Nothing AndAlso rows.Length = 1 Then                                  ' if found query name                        
                        QueryDefText = GcrcvAs(rows(0), cnViewDef, String.Empty)                        ' get view definition
                        QueryRow = AccessQueries.NewRow()                                               ' create new CurQueries row
                        QueryRow(cnQueryName) = QueryName                                               ' populate CurQueries row
                        QueryRow(cnQueryDef) = QueryDefText
                        QueryRow(cnViewOrProc) = qtView
                        AccessQueries.Rows.Add(QueryRow)                                                ' add CurQueries row to table
                    End If
                Next
            End If

            ' move current proc data into queries table        
            If AccessProcs IsNot Nothing Then                                                           ' if got current procs
                For i As Integer = 0 To QueriesCheckedListBox.CheckedItemsCount - 1                     ' for each selected query
                    QueryName = CStr(QueriesCheckedListBox.CheckedItems(i))                             ' get the query name
                    RowSelectStr = PcdbModule.FieldSelectString(cnProcedureName, QueryName)             ' set selection string
                    rows = AccessProcs.Select(RowSelectStr)                                             ' find row in access queries
                    If rows IsNot Nothing AndAlso rows.Length = 1 Then                                  ' if found row                        
                        QueryDefText = GcrcvAs(rows(0), cnProcedureDef, String.Empty)                   ' get proc definition
                        QueryRow = AccessQueries.NewRow()                                               ' create new CurQueries row
                        QueryRow(cnQueryName) = QueryName                                               ' populate CurQueries row
                        QueryRow(cnQueryDef) = QueryDefText
                        QueryRow(cnViewOrProc) = qtProc
                        AccessQueries.Rows.Add(QueryRow)                                                ' add CurQueries row to table
                    End If
                Next
            End If
            Return AccessQueries.Rows.Count
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.CreateQueryTables.  " & ex.Message)
            Return Pcm.teOtherErr
        End Try
    End Function

    Private Function CopyQueriesToMySql(worker As System.ComponentModel.BackgroundWorker, e As System.ComponentModel.DoWorkEventArgs) As Integer


        ' vars passed:
        '   worker - background worker
        '   e - worker event arguments

        Dim progress As New PcdbToMySqlProgress(ActionTextEdit, ActionProgressBarControl, RunningTimeTextEdit, totalSteps,
                                                TotalProgressBarControl, CopiedTablesLabel, CopiedQueriesLabel)
        ShowProgress(worker, progress, "Copying Queries")

        Try
            Dim QueryDefText As String
            Dim QueryName As String
            Dim QueryType As String

            ' queries can contain references to other queries.  if a query has a reference to another query, the referenced query 
            ' must be created before the referencing query can be created.  the table of current queries must be sorted in to the
            ' correct query creation order

            Dim QueryCount As Integer = CreateQueryTables()                                         ' make the Query tables (populates AccessQueries)
            If QueryCount < 0 Then                                                                  ' if got an error
                Return EaTools.Ea.teMakeTableErr                                                    ' exit now
            End If
            progress.TotalStep()
            ShowProgress(worker, progress)

            Dim err As Integer = SortQueries()                                                      ' sort the cur queries to QUERY_ORDER column
            If err <> Pcm.NoErrors Then                                                             ' if got an error
                Return EaTools.Ea.teTableHasErrors                                                  ' exit now
            End If
            progress.TotalStep()
            ShowProgress(worker, progress)

            Dim CurQueriesView As DataView = New DataView(AccessQueries) With {.Sort = cnQueryOrder} ' create dataview sorted by order
            For i As Integer = 0 To QueriesCheckedListBox.CheckedItemsCount - 1                     ' for each table to copy
                QueryName = GcrcvAs(CurQueriesView(i).Row, cnQueryName, String.Empty)               ' get the query name 
                If QueryName = String.Empty Then                                                    ' if no query name
                    CancelBackgroundWorker()
                    MessageBox.Show("No query name when creating queries.",
                                    etCreatingQuery, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return EaSql.Ea.sqlNoQueryName                                                  ' exit now
                End If
                QueryDefText = GcrcvAs(CurQueriesView(i).Row, cnQueryDef, String.Empty)             ' get the query definition text
                If QueryDefText = String.Empty Then                                                 ' if no query definition text
                    CancelBackgroundWorker()
                    MessageBox.Show("No query definition text when creating queries.",
                                    etCreatingQuery, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return EaSql.Ea.sqlNoCommandText                                                ' exit now
                End If
                QueryType = GcrcvAs(CurQueriesView(i).Row, cnViewOrProc, String.Empty)              ' get the query type
                If QueryType = String.Empty Then                                                    ' if no query type
                    CancelBackgroundWorker()
                    MessageBox.Show("No query type (View or Procedure) when creating queries.",
                                    etCreatingQuery, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return EaSql.Ea.sqlNoQueryType                                                  ' exit now
                End If

                If QueryType = qtView Then                                                          ' if a view 
                    err = CopyView(worker, progress, QueryName, QueryDefText)                       ' copy the view
                Else                                                                                ' else a procedure
                    err = CopyProcedure(worker, progress, QueryName, QueryDefText)                  ' copy the procedure
                End If
                If err <> Pcm.NoErrors Then                                                         ' if error refreshing query                    
                    Return err                                                                      ' exit now
                End If
            Next
            progress.CopiedQueries = True
            ShowProgress(worker, progress, 100)

            Return NoErrors
        Catch ex As Exception
            Return EaTools.Ea.OtherExErr
        End Try
    End Function

    Private Sub CopyQueriesButton_Click(sender As Object, e As EventArgs) Handles CopyQueriesButton.Click

        Const TableCount As Integer = 0
        Const StepsPerTable As Integer = 0
        Const SetupSteps As Integer = 4
        Const QuerySetupSteps As Integer = 2
        Const StepsPerQuery As Integer = 6
        Const FinalSteps As Integer = 0

        totalSteps = (TableCount * StepsPerTable) + ((QueriesCheckedListBox.CheckedItemsCount - 1) * StepsPerQuery) + SetupSteps + QuerySetupSteps + FinalSteps  ' calc total # of steps
        If PreCopy(totalSteps) <> NoErrors Then
            Return
        End If

        BackgroundWorker1.RunWorkerAsync(CopyTasks.Queries)
    End Sub

#End Region

#Region " CopyTables "

    Private Function CopyTablesToMySql(worker As System.ComponentModel.BackgroundWorker, e As System.ComponentModel.DoWorkEventArgs) As Integer

        ' copies selected tables from PCDB to MySql
        '
        ' vars passed:
        '   worker - background worker
        '   e - worker event arguments
        '
        ' for each table:
        ' 0) get table structure from Access database
        ' 1) create dataViews needed when making the tables
        ' 2) create list of columns for table                
        ' 3) make the data tables (NO AUTO INC COLUMNS FOR MySql Table)
        ' 4) populate the Access table 
        ' 5) Copy data to to MySql table
        ' 6) Create MySql table adapter for MySql Table
        ' 7) update the MySql table in the MySql database
        ' 8) re-fill MySql table so data is from MySql database, not AccessTable.Copy
        ' 9) verify the data in the MySql table
        ' 10) Change key column in MySql table to Auto Inc 
        ' 11) Change Column names that are MySql key words
        ' 12) drop DataLocInt column
        ' 13) add table indexes

        Try
            Dim progress As New PcdbToMySqlProgress(ActionTextEdit, ActionProgressBarControl, RunningTimeTextEdit, totalSteps,
                                                    TotalProgressBarControl, CopiedTablesLabel, CopiedQueriesLabel)
            ShowProgress(worker, progress, "Copying Tables")

            Dim err As Integer
            Dim tableName As String
            For i As Integer = 0 To TablesCheckedListBox.CheckedItemsCount - 1                      ' for each table to copy

                ' 0) get table structure from Access database                
                tableName = CStr(TablesCheckedListBox.CheckedItems(i))                              ' get table name
                err = GetTableStructure(AccessOleDbConnection, tableName, AccessColumns, AccessIndexes)
                If err <> Pcm.NoErrors Then                                                          ' if error getting table structure
                    Return err
                End If
                ShowProgress(worker, progress, 20)

                ' 1) create dataViews needed when making the tables
                Dim SortStr As String = cnOrdinalPos & " ASC"                                       ' sort of ordinal pos column
                Dim KeyFilterStr As String = PcdbModule.FieldSelectString(cnPrimaryKey, True)       ' [PRIMARY_KEY] = TRUE
                Dim IndexFilterStr As String = PcdbModule.FieldSelectString(cnPrimaryKey, False)    ' [PRIMARY_KEY] = FALSE

                Dim ColumnsView As New DataView(AccessColumns) With {.Sort = SortStr}               ' create column data view
                Dim PrimariesView As New DataView(AccessIndexes) With
                    {.RowFilter = KeyFilterStr,
                     .Sort = SortStr}                                                                ' create primary key data view
                SortStr = String.Format("{0} {1}, {2} {1}", cnIndexName, "ASC", cnOrdinalPos)       ' sort by index name, ordinal pos
                Dim IndexesView As New DataView(AccessIndexes) With
                    {.RowFilter = IndexFilterStr,
                     .Sort = SortStr}                                                               ' create indexes data view
                ShowProgress(worker, progress, 30)

                ' 2) create list of columns for table
                Dim acColumns As New List(Of EaAccess.AccessColumn)
                Dim acPrimaryKeyCols As New List(Of EaAccess.AccessColumn)
                Dim MySqlColumns As New List(Of MySqlColumn)
                Dim MySqlPrimaryKeyCols As New List(Of MySqlColumn)
                err = CreateColumnLists(tableName, ColumnsView, PrimariesView, acColumns, acPrimaryKeyCols, MySqlColumns, MySqlPrimaryKeyCols)
                If err <> Pcm.NoErrors Then                                                         ' if error creating table
                    Return err                                                                      ' return error
                End If
                ShowProgress(worker, progress, 50)

                ' 3) make the data tables
                Dim AccessTable As DataTable = Nothing
                Dim MySqlTable As DataTable = Nothing
                err = CreateTables(tableName, acColumns, MySqlColumns, AccessTable)                 ' create the tables
                If err <> Pcm.NoErrors Then                                                         ' if error creating table
                    Return err                                                                      ' return error
                End If
                ShowProgress(worker, progress, 70)

                ' 4) populate the Access table 
                ' get OleDbDataAdapter
                Dim AccessOleDbDataAdapt As OleDb.OleDbDataAdapter = CreateOleDbDataAdapt(AccessOleDbConnection, tableName, True)
                If AccessOleDbDataAdapt Is Nothing Then                                             ' if error getting OleDbDataAdapter
                    Return Pcm.deNoDataAdapt                                                        ' return error
                End If
                Dim RowCount As Integer = FillTable(AccessTable, AccessOleDbDataAdapt)
                If RowCount < 0 Then                                                                ' if error populating table
                    Return RowCount                                                                 ' return error
                End If
                ShowProgress(worker, progress, 80)

                ' 5) Copy data to to MySql table
                MySqlTable = AccessTable.Copy()
                If MySqlTable Is Nothing Then
                    Return Pcm.deCouldNotCreate
                End If
                progress.TotalStep()                                                                ' one more total step done
                ShowProgress(worker, progress, 100)

                ' 6) Create MySql table adapter for MySql Table
                ShowProgress(worker, progress, String.Format("Saving MySql table: {0}", tableName))
                Dim MySqlDataAdapt As MySqlDataAdapter = MySqlTools.CreateMySqlDataAdapt(MySqlPcdbConn, tableName, False)
                If MySqlDataAdapt Is Nothing Then
                    Return Pcm.deNoDataAdapt
                End If
                ShowProgress(worker, progress, 10)

                ' 7) update the MySql table in the MySql database
                RowCount = UpdateTable(MySqlTable, MySqlPcdbConn, MySqlDataAdapt, AccessTable.Rows.Count)
                If RowCount < 0 Then                                                                ' if error updating table
                    Return RowCount                                                                 ' return error
                End If
                ShowProgress(worker, progress, 90)

                ' 8) re-fill MySql table so data is from MySql database, not AccessTable.Copy
                RowCount = MySqlTools.FillTable(MySqlTable, MySqlDataAdapt)
                If RowCount < 0 Then                                                                ' if error filling table
                    Return RowCount                                                                 ' return error
                End If
                If RowCount <> AccessTable.Rows.Count Then                                          ' if did not fill correct # of rows
                    CancelBackgroundWorker()
                    MessageBox.Show(String.Format("Filled row count <> desired row count for table: {1}{0}Filled row count: {2}{0}Desired row count: {3}",
                                                  vbCrLf, tableName, RowCount, AccessTable.Rows.Count),
                                   "Fill Table Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return Pcm.teWrongRowCountErr                                                   ' return error
                End If
                progress.TotalStep()
                ShowProgress(worker, progress, 100)

                ' 9) verify the data in the MySql table                
                progress.ActionMax = AccessTable.Rows.Count
                ShowProgress(worker, progress, String.Format("Table {0}: Verifying table data...", tableName))

                err = VerifyTableData(worker, progress, tableName, AccessTable, MySqlTable)
                If err <> Pcm.NoErrors Then                                                         ' if error verifying table
                    Return err                                                                      ' return error
                End If
                If tableName = Pcm.WorkDoneTableName Then                                           ' if the WorkDone table
                    err = FixWorkDoneTableId(MySqlPcdbConn, MySqlDataAdapt, MySqlTable)             ' fix the row with id problem
                    If err < 0 Then                                                                 ' if got an error
                        Return err                                                                  ' return error
                    End If
                End If
                progress.TotalStep()                                                                ' one more total step done
                ShowProgress(worker, progress)

                ' 10) if VerifyTableData = NoErrors, but NeedToAdjustValues = TRUE
                ' 

                ' 10) Change key column in MySql table to Auto Inc 
                progress.ActionMax = 100
                ShowProgress(worker, progress, String.Format("Final MySql changes for table: {0}", tableName))

                err = MakeSingleKeyFieldAutoInc(MySqlPcdbConn, tableName, MySqlColumns, PrimariesView)
                If err <> Pcm.NoErrors Then                                                         ' if error making single key field auto inc
                    Return err                                                                      ' return error
                End If
                ShowProgress(worker, progress, 33)

                ' 11) Change Column names that are MySql key words
                err = ConvertColNames(MySqlPcdbConn, tableName, MySqlColumns, IndexesView)
                If err <> Pcm.NoErrors Then                                                         ' if error converting column names
                    Return err                                                                      ' return error
                End If
                ShowProgress(worker, progress, 67)

                ' 12) drop DataLocInt column
                err = DropDataLocIntCol(MySqlPcdbConn, tableName)                                   ' drop the DataLocInt column
                If err <> Pcm.NoErrors Then                                                         ' if error dropping column
                    Return err                                                                      ' return error
                End If
                ShowProgress(worker, progress, 70)

                ' 13) add table indexes
                err = MySqlTools.CreateTableIndexes(MySqlPcdbConn, tableName, IndexesView)          ' create indexes after converting name column
                If err <> Pcm.NoErrors Then                                                         ' if error adding table indexes
                    Return err                                                                      ' return error
                End If
                progress.TotalStep()                                                                ' one more total step done
                ShowProgress(worker, progress, 100)
            Next
            progress.CopiedTables = True
            ShowProgress(worker, progress, 100)
            Return NoErrors
        Catch ex As Exception
            CancelBackgroundWorker()
            PcdbModule.ShowOtherErrorMessage("MainForm.CopyTables.  " & ex.Message)
            Return Pcm.OtherExErr
        End Try
    End Function

    Private Sub CopyTablesButton_Click(sender As Object, e As EventArgs) Handles CopyTablesButton.Click

        Const StepsPerTable As Integer = 4
        Const SetupSteps As Integer = 2
        Const QueryCount As Integer = 0
        Const StepsPerQuery As Integer = 0
        Const FinalSteps As Integer = 0

        totalSteps = ((TablesCheckedListBox.CheckedItemsCount - 1) * StepsPerTable) + (QueryCount * StepsPerQuery) + SetupSteps + FinalSteps
        If PreCopy(totalSteps) <> NoErrors Then
            Return
        End If

        BackgroundWorker1.RunWorkerAsync(CopyTasks.Tables)
    End Sub

#End Region

#Region " Create Database "

    Private Sub CreateDatabaseButton_Click(sender As Object, e As EventArgs) Handles CreateDatabaseButton.Click

        Try
            SuccessLabel.Visible = False
            If Not ValidMySqlSettings() Then
                Return
            End If
            Dim dbToCreate As String = MySqlPcdbNameTextEdit.Text
            Dim MySqlConnStr As String = MySqlTools.ConnectionString(MySqlServerTextEdit.Text, UserIdTextEdit.Text, PasswordTextEdit.Text)
            Dim MySqlConn As MySqlConnection = New MySqlConnection(MySqlConnStr)
            Dim err As Integer = MySqlTools.CreateDatabase(MySqlConn, dbToCreate)
            If err >= 0 Then
                SuccessLabel.Visible = True
            End If
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.CreateDatabaseButton_Click.  " & ex.Message)
        End Try
    End Sub

#End Region

#Region " Get Tables "

    Private Sub GetTablesButton_Click(sender As Object, e As EventArgs) Handles GetTablesButton.Click

        Try
            If Not ValidAccessDatabase() Then
                Return
            End If

            ShowActions(True)
            'SetAction("Getting Tables and Queries for Access Database")                             ' show action being done
            SetAction("Connection to Access Database")                                              ' show action being done

            If AccessFolder = String.Empty OrElse AccessDatabaseName = String.Empty Then
                PcdbModule.ShowErrorMessage("No Access PCDB database selected", "Cannot Get Tables")
            End If
            Dim Err As Integer = MakeAccessConnection(AccessFolder, AccessDatabaseName, AccessOleDbConnection)  ' make connection to access database
            If Err <> Pcm.NoErrors Then                                                                         ' if got an error
                Return                                                                                          ' return error
            End If
            ActionPosition(50)

            Err = GetAccessTablesAndQueries()
            ActionPosition(90)

            PopulateTablesCheckedListBox()
            PopulateQueriesCheckedListBox()
            ActionPosition(100)

        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.GetTablesButton_Click.  " & ex.Message)
        End Try
    End Sub

    Private Sub PopulateQueriesCheckedListBox()

        Try
            Dim queryName As String
            Dim i As Integer = 1
            For Each row As DataRow In AccessViews.Rows                             ' for each row in accessViews
                queryName = GcrcvAs(row, cnTableName, String.Empty)                 ' get the query name
                If queryName = String.Empty Then
                    PcdbModule.ShowErrorMessage(String.Format("Error getting query name in list of queries from row # {0}.", i), "No Query")
                End If
                QueriesCheckedListBox.Items.Add(queryName)
                i += 1
            Next
            i = 1
            For Each row As DataRow In AccessProcs.Rows                             ' for each row in accessProcs
                queryName = GcrcvAs(row, cnProcedureName, String.Empty)             ' get the query name
                If queryName = String.Empty Then
                    PcdbModule.ShowErrorMessage(String.Format("Error getting procedure name in list of procedures from row # {0}.", i), "No Procedure")
                End If
                QueriesCheckedListBox.Items.Add(queryName)
                i += 1
            Next
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.PopulateQueriesCheckedListBox.  " & ex.Message)
        End Try
    End Sub

    Private Sub PopulateTablesCheckedListBox()

        Try
            Dim tableName As String
            Dim i As Integer = 1
            For Each row As DataRow In AccessTables.Rows                            ' for each row in accessTables
                tableName = GcrcvAs(row, cnTableName, String.Empty)                 ' get the table name
                If tableName = String.Empty Then
                    PcdbModule.ShowErrorMessage(String.Format("Error getting table name in list of tables from row # {0}.", i), Pcm.etNoTable)
                End If
                TablesCheckedListBox.Items.Add(tableName)
                i += 1
            Next
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.PopulateTablesCheckedListBox.  " & ex.Message)
        End Try
    End Sub

    Private Function ValidAccessDatabase() As Boolean

        Try
            If AccessButtonEdit.Text = String.Empty OrElse AccessFolder = String.Empty _
                    OrElse AccessDatabaseName = String.Empty OrElse AccessFullDatabaseName = String.Empty Then
                AccessButtonEdit.Focus()
                PcdbModule.ShowErrorMessage("There is no Access database", Pcm.etDataEntryErr)
                Return False
            End If
            Return True
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.ValidAccessDatabase.  " & ex.Message)
            Return False
        End Try
    End Function

#End Region

#Region " MySql "

    Private Function ValidMySqlSettings() As Boolean

        Try
            If MySqlServerTextEdit.Text = String.Empty Then
                MySqlServerTextEdit.Focus()
                PcdbModule.ShowErrorMessage("There is no MySql Server", Pcm.etDataEntryErr)
                Return False
            End If
            If MySqlPcdbNameTextEdit.Text = String.Empty Then
                MySqlPcdbNameTextEdit.Focus()
                PcdbModule.ShowErrorMessage("There is no MySql PCDB Database", Pcm.etDataEntryErr)
                Return False
            End If
            If UserIdTextEdit.Text = String.Empty Then
                UserIdTextEdit.Focus()
                PcdbModule.ShowErrorMessage("There is no MySql User Id", Pcm.etDataEntryErr)
                Return False
            End If
            If PasswordTextEdit.Text = String.Empty Then
                PasswordTextEdit.Focus()
                PcdbModule.ShowErrorMessage("There is no MySql Password", Pcm.etDataEntryErr)
                Return False
            End If
            Return True
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.ValidMySqlSettings.  " & ex.Message)
            Return False
        End Try
    End Function

    Private Sub MySqlConnectButton_Click(sender As Object, e As EventArgs) Handles MySqlConnectButton.Click

        Try
            If Not ValidMySqlSettings() Then
                Return
            End If
            Dim err As Integer = MakeMySqlConnection(MySqlServerTextEdit.Text, UserIdTextEdit.Text, PasswordTextEdit.Text, MySqlPcdbNameTextEdit.Text)
            If err = Pcm.NoErrors Then
                PcdbModule.ShowErrorMessage(String.Format("Connected to MySql database {0}", MySqlPcdbNameTextEdit.Text), "Connected", , MessageBoxIcon.Information)
                SuccessLabel.Visible = True
            End If
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.MySqlConnectButton_Click.  " & ex.Message)
        End Try
    End Sub

    Private Sub GetTableSchemaButton_Click(sender As Object, e As EventArgs) Handles GetTableSchemaButton.Click

        Try
            SuccessLabel.Visible = False
            If Not ValidMySqlSettings() Then
                Return
            End If

            Dim MySqlConnStr As String = MySqlTools.ConnectionString(MySqlServerTextEdit.Text, UserIdTextEdit.Text,
                                                                  PasswordTextEdit.Text, MySqlPcdbNameTextEdit.Text)
            Dim MySqlConn As MySqlConnection = MySqlTools.MySqlConn(MySqlConnStr)

            Dim SchmTbl As DataTable = MySqlTools.TableInfo(MySqlConn, "Companies")
            If SchmTbl Is Nothing Then
                MessageBox.Show("No Schema Table for ""Companies"".")
                Return
            End If
            If SchmTbl.Rows.Count = 0 Then
                MessageBox.Show("No rows in Schema Table for ""Companies"".")
                Return
            End If
            MessageBox.Show(String.Format("Got {0} row(s) for in Schema Table for ""Companies"".", SchmTbl.Rows.Count))

            SuccessLabel.Visible = True
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.ReadConfigButton_Click.  " & ex.Message)
        End Try
    End Sub

#End Region

#Region " Users "

    Private Sub CreateTestUserButton_Click(sender As Object, e As EventArgs) Handles CreateTestUserButton.Click

        Try
            SuccessLabel.Visible = False
            If Not ValidMySqlSettings() Then
                Return
            End If

            Dim MySqlConnStr As String = MySqlTools.ConnectionString(MySqlServerTextEdit.Text, UserIdTextEdit.Text, PasswordTextEdit.Text)
            Dim MySqlConn As MySqlConnection = MySqlTools.MySqlConn(MySqlConnStr)

            Dim err As Integer = MySqlTools.DropUser(MySqlConn, TestUserName)
            If err < 0 Then
                Return
            End If
            err = MySqlTools.CreateUser(MySqlConn, TestUserName, TestPassword)
            If err < 0 Then
                Return
            End If
            err = MySqlTools.UserGrantAllAccess(MySqlConn, MySqlPcdbNameTextEdit.Text, TestUserName)
            If err < 0 Then
                Return
            End If

            SuccessLabel.Visible = True
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.CreateTestUserButton_Click.  " & ex.Message)
        End Try
    End Sub

    Private Sub ChangePasswordButton_Click(sender As Object, e As EventArgs) Handles ChangePasswordButton.Click

        Try
            SuccessLabel.Visible = False
            If Not ValidMySqlSettings() Then
                Return
            End If
            Dim MySqlConnStr As String = MySqlTools.ConnectionString(MySqlServerTextEdit.Text, UserIdTextEdit.Text, PasswordTextEdit.Text)
            Dim MySqlConn As MySqlConnection = MySqlTools.MySqlConn(MySqlConnStr)

            Dim err As Integer = MySqlTools.ChangePassword(MySqlConn, TestUserName, TestNewPassword)
            If err < 0 Then
                Return
            End If

            SuccessLabel.Visible = True
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.ChangePasswordButton_Click.  " & ex.Message)
        End Try
    End Sub

    Private Sub WriteConfigButton_Click(sender As Object, e As EventArgs) Handles WriteConfigButton.Click

        Try
            SuccessLabel.Visible = False

            If Not ValidMySqlSettings() Then
                Return
            End If
            Dim connSettings As New DcSettings()

            connSettings.Server = MySqlServerTextEdit.Text
            connSettings.UserName = UserIdTextEdit.Text
            connSettings.Password = PasswordTextEdit.Text
            connSettings.Database = MySqlPcdbNameTextEdit.Text
            connSettings.PerSecInfo = PersistSecurityInfoCheckBox.Checked
            connSettings.sslMode = MySqlSslMode.None

            Dim err As Integer = connSettings.WriteUserSettings()
            If err < 0 Then
                Return
            End If

            SuccessLabel.Visible = True
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.WriteConfigButton_Click.  " & ex.Message)
        End Try
    End Sub

    Private Sub ReadConfigButton_Click(sender As Object, e As EventArgs) Handles ReadConfigButton.Click

        Try
            SuccessLabel.Visible = False

            If Not ValidMySqlSettings() Then
                Return
            End If
            Dim connSettings As New DcSettings()

            Dim err As Integer = connSettings.ReadUserSettings()
            If err < 0 Then
                Return
            End If

            SuccessLabel.Visible = True
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("MainForm.ReadConfigButton_Click.  " & ex.Message)
        End Try
    End Sub

    Private Sub OtherButton_Click(sender As Object, e As EventArgs) Handles OtherButton.Click

        If Not ValidMySqlSettings() Then                                        ' if not valid MySql settings
            Return
        End If

        Dim MySqlConnStr As String = MySqlTools.ConnectionString(MySqlServerTextEdit.Text, UserIdTextEdit.Text, PasswordTextEdit.Text)
        Dim MySqlConn As MySqlConnection = MySqlTools.MySqlConn(MySqlConnStr)

        Dim exists As MySqlTools.ExistsTypes = MySqlTools.UserExists(MySqlConn, TestUserName)
        MessageBox.Show(String.Format("User: {0} exits: {1}", TestUserName, exists), "Before Drop User")

        Dim err As Integer = MySqlTools.DropUser(MySqlConn, TestUserName)
        If err < 0 Then
            Return
        End If

        exists = MySqlTools.UserExists(MySqlConn, TestUserName)
        MessageBox.Show(String.Format("User: {0} exits: {1}", TestUserName, exists), "After Drop User")

        SuccessLabel.Visible = True

        'Dim connStr As String = MySqlTools.ConnectionString(MySqlServerTextEdit.Text, UserIdTextEdit.Text, PasswordTextEdit.Text,
        '                                                    MySqlPcdbNameTextEdit.Text, PersistSecurityInfoCheckBox.Checked, GetSslMode)
        'Using conn As New MySqlConnection(connStr)
        '    Using cmd As New MySqlCommand("CabTest", conn)
        '        cmd.CommandType = CommandType.StoredProcedure
        '        'cmd.Parameters.AddWithValue("TravelId", 100)
        '        cmd.Parameters.AddWithValue("TRAVELID", 100)
        '        Using sda As New MySqlDataAdapter(cmd)
        '            Dim dt As New DataTable
        '            Dim rowCount As Integer = sda.Fill(dt)
        '            MessageBox.Show(String.Format("rowCount = {0}", rowCount))
        '            If rowCount > 0 Then
        '                MessageBox.Show(String.Format("Cab Fares: {0:C} for TravelId: {1}", dt.Rows(0)(0), 100), "Success!")
        '            End If
        '        End Using
        '    End Using
        'End Using

    End Sub

    Private Sub SimpleButton1_Click(sender As Object, e As EventArgs) Handles SimpleButton1.Click
        ShowActions(True)
        StartStopTimer(True)
    End Sub

    Private Sub SimpleButton2_Click(sender As Object, e As EventArgs) Handles SimpleButton2.Click
        'Timer1.Stop()
        StartStopTimer(False)
    End Sub

#End Region

#End Region

End Class