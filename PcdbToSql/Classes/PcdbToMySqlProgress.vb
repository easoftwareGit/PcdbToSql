Public Class PcdbToMySqlProgress

    Private _actionMax As Integer
    Private _totalMax As Integer

    Public ReadOnly Property ActionMessageTextEdit As DevExpress.XtraEditors.TextEdit
    Public ReadOnly Property ActionPbParams As DevExpress.XtraEditors.Repository.RepositoryItemProgressBar
    Public ReadOnly Property CopiedQueriesLabel As DevExpress.XtraEditors.LabelControl
    Public ReadOnly Property CopiedTablesLabel As DevExpress.XtraEditors.LabelControl
    Public ReadOnly Property TotalPbParams As DevExpress.XtraEditors.Repository.RepositoryItemProgressBar
    Public ReadOnly Property TotalTimeTextEdit As DevExpress.XtraEditors.TextEdit

    Public Property ActionMax As Integer
        Get
            Return _actionMax
        End Get
        Set(value As Integer)
            _actionMax = value
            ResetActionMax = True
        End Set
    End Property
    Public Property ActionMessage As String
    Public Property ActionPosition As Integer
    Public Property ActionProgressBar As DevExpress.XtraEditors.ProgressBarBaseControl
    Public Property ResetActionMax As Boolean = False
    Public Property ResetTotalMax As Boolean = False

    Public Property CopiedTables As Boolean = False
    Public Property CopiedQueries As Boolean = False

    Public Property TotalMax As Integer
        Get
            Return _totalMax
        End Get
        Set(value As Integer)
            _totalMax = value
            ResetTotalMax = True
        End Set
    End Property
    Public Property TotalPosition As Integer
    Public Property TotalProgressBar As DevExpress.XtraEditors.ProgressBarBaseControl
    Public Property TotalTimeText As String

    Public Sub New(actionTe As DevExpress.XtraEditors.TextEdit,
                   actionPb As DevExpress.XtraEditors.ProgressBarBaseControl,
                   totalTe As DevExpress.XtraEditors.TextEdit,
                   totalMax As Integer,
                   totalPb As DevExpress.XtraEditors.ProgressBarBaseControl,
                   copiedTablesLbl As DevExpress.XtraEditors.LabelControl,
                   copiedQueriesLbl As DevExpress.XtraEditors.LabelControl)

        ActionMessage = String.Empty
        TotalTimeText = String.Empty

        ActionMessageTextEdit = actionTe
        ActionProgressBar = actionPb
        TotalTimeTextEdit = totalTe
        TotalProgressBar = totalPb
        CopiedTablesLabel = copiedTablesLbl
        CopiedQueriesLabel = copiedQueriesLbl

        ActionPbParams = CType(actionPb.Properties, DevExpress.XtraEditors.Repository.RepositoryItemProgressBar)
        ActionPosition = CInt(actionPb.EditValue)
        _actionMax = ActionPbParams.Maximum

        TotalPbParams = CType(totalPb.Properties, DevExpress.XtraEditors.Repository.RepositoryItemProgressBar)
        TotalPosition = CInt(totalPb.EditValue)
        Me.TotalMax = totalMax
    End Sub

    Public Sub ActionMessageUpdate()

        ActionMessageTextEdit.Text = ActionMessage
        ActionMessageTextEdit.Update()
    End Sub

    Public Sub ActionStep()

        ActionPosition += 1
        ' ActionPosition = CInt(ActionProgressBar.EditValue) + 1
    End Sub

    Public Sub ActionUpdate()

        ActionProgressBar.EditValue = ActionPosition
        ActionProgressBar.Update()
    End Sub

    Public Sub ShowCopiedQueries()
        If CopiedQueries AndAlso CopiedQueriesLabel.Visible = False Then
            CopiedQueriesLabel.Visible = True
            CopiedQueriesLabel.Update()
        End If
    End Sub

    Public Sub ShowCopiedTables()
        If CopiedTables AndAlso CopiedTablesLabel.Visible = False Then
            CopiedTablesLabel.Visible = True
            CopiedTablesLabel.Update()
        End If
    End Sub

    Public Sub TotalStep()

        TotalPosition += 1
        'TotalPosition = CInt(TotalProgressBar.EditValue) + 1
    End Sub

    Public Sub TotalUpdate()

        TotalProgressBar.EditValue = TotalPosition
        TotalProgressBar.Update()
    End Sub

    Public Sub TimeUpdate()

        TotalTimeTextEdit.Text = TotalTimeText
        TotalTimeTextEdit.Update()
    End Sub

End Class
