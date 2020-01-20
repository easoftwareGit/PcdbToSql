Public Class QueryInfoClass

#Region " Class Constants and Variables "

#Region " Constants "

#End Region

#Region " Enums "

#End Region

#Region " Private Vars "

    Private _sortOrder As Integer
    Private _queriesUsed As List(Of String)
    Private _queryDefText As String
    Private _queryName As String

#End Region

#End Region

#Region " New "

    Public Sub New()
        Me._sortOrder = -1
    End Sub

    Public Sub New(QueryName As String, QueryDefText As String)
        Me.New()
        _queryName = QueryName
        _queryDefText = QueryDefText
        _queriesUsed = New List(Of String)
    End Sub

    Public Sub New(QueryName As String, QueryDefText As String, QueriesUsed As List(Of String))
        Me.New(QueryName, QueryDefText)
        For Each qStr As String In QueriesUsed
            _queriesUsed.Add(qStr)
        Next
    End Sub

    Public Sub New(QueryInfo As QueryInfoClass)
        Me.New(QueryInfo.QueryName, QueryInfo.QueryDefText)
        QueryInfo.CopyTo(Me)
    End Sub

#End Region

#Region " Properties and Methods "

#Region " Properties "

    Public Property SortOrder As Integer
        Get
            Return _sortOrder
        End Get
        Set(ByVal Value As Integer)
            _sortOrder = Value
        End Set
    End Property

    Public Property QueryDefText As String
        Get
            Return _queryDefText
        End Get
        Set(ByVal Value As String)
            _queryDefText = Value
        End Set
    End Property

    Public Property QueryName As String
        Get
            Return _queryName
        End Get
        Set(ByVal Value As String)
            _queryName = Value
        End Set
    End Property

    Public Property QueriesUsed As List(Of String)
        Get
            Return _queriesUsed
        End Get
        Set(ByVal Value As List(Of String))
            _queriesUsed = Value
        End Set
    End Property

#End Region

#Region " Methods "

    Public Sub CopyTo(QueryInfo As QueryInfoClass)

        QueryInfo.SortOrder = Me._sortOrder
        QueryInfo.QueryDefText = Me._queryDefText
        QueryInfo.QueryName = Me._queryName
        QueryInfo.QueriesUsed.Clear()
        For Each queryName As String In Me._queriesUsed
            QueryInfo.QueriesUsed.Add(queryName)
        Next
    End Sub

#End Region

#End Region

End Class
