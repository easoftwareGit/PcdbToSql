Public Class eaColumnClass

#Region " Class Constants and Variables "

#Region " Constants "

    'Private Const sqlCnBoolean As String = "Logical"
    Private Const sqlCnBoolean As String = "Bit"
    Private Const sqlCnCurrency As String = "Currency"
    Private Const sqlCnDate As String = "Date"
    Private Const sqlCnDouble As String = "Double"
    Private Const sqlCnInteger As String = "Integer"
    Private Const sqlCnMemo As String = "Memo"
    Private Const sqlCnSmallInt As String = "SmallInt"
    Private Const sqlCnText As String = "Text"
    'Private Const sqlCnYesNo As String = "YesNo"

#End Region

#Region " Enums "

    Enum ColumnTypes
        ctNone = 0
        ctBoolean = 1
        ctCurrency = 2
        ctDate = 3
        ctDouble = 4
        ctInteger = 5
        ctMemo = 6
        ctText = 7
        ctSmallInt = 8
    End Enum

    Enum AlterTableTypes
        atNone = 0
        atAddColumn = 1
        atAlterColumn = 2
        atDropColumn = 3
        atAddKey = 4
        atDropKey = 5
    End Enum

#End Region

#Region " Private Vars "

    Private ea_ColumnName As String
    Private ea_ColumnType As ColumnTypes
    Private ea_DefaultValue As String
    Private ea_KeyField As Boolean
    Private ea_IsNullable As Boolean
    Private ea_TextLength As Integer

#End Region

#End Region

#Region " New "

    Public Sub New(ByVal aColumnName As String,
                   ByVal aColumnType As ColumnTypes,
                   Optional ByVal aKeyField As Boolean = False,
                   Optional ByVal aDefaultValue As String = "",
                   Optional ByVal aIsNullable As Boolean = True)

        ' vars passed:
        '   aColumnName - name for the column
        '   aColumnType - column type 
        '   aKeyField - is column a key field
        '   aDefaultValue - columns's default value 

        ColumnName = aColumnName
        ColumnType = aColumnType
        KeyField = aKeyField
        DefaultValue = aDefaultValue
        IsNullable = aIsNullable
    End Sub

    Public Sub New(ByVal aColumnName As String,
                   ByVal aColumnType As ColumnTypes,
                   ByVal aTextLength As Integer,
                   Optional ByVal aKeyField As Boolean = False,
                   Optional ByVal aDefaultValue As String = "",
                   Optional ByVal aIsNullable As Boolean = True)

        ' vars passed:
        '   aColumnName - name for the column
        '   aColumnType - column type 
        '   aTextLength - length of text in a text column
        '   aKeyField - is column a key field
        '   aDefaultValue - columns's default value as a string

        Me.New(aColumnName, aColumnType, aKeyField, aDefaultValue, aIsNullable)
        TextLength = aTextLength
    End Sub

#End Region

#Region " Properties "

    Property ColumnName() As String
        Get
            Return Me.ea_ColumnName
        End Get
        Set(ByVal Value As String)
            Try
                ea_ColumnName = Value
            Catch ex As Exception
                ea_ColumnName = ""
            End Try
        End Set
    End Property

    Property ColumnType() As ColumnTypes
        Get
            Return Me.ea_ColumnType
        End Get
        Set(ByVal Value As ColumnTypes)
            Try
                If (Value = eaColumnClass.ColumnTypes.ctBoolean) OrElse (Value = eaColumnClass.ColumnTypes.ctCurrency) OrElse _
                        (Value = eaColumnClass.ColumnTypes.ctDate) OrElse (Value = eaColumnClass.ColumnTypes.ctDouble) OrElse _
                        (Value = eaColumnClass.ColumnTypes.ctInteger) OrElse (Value = eaColumnClass.ColumnTypes.ctMemo) OrElse _
                        (Value = eaColumnClass.ColumnTypes.ctText) OrElse (Value = ColumnTypes.ctSmallInt) Then ' if a valid type
                    ea_ColumnType = Value                                       ' then set column type
                Else                                                            ' else type is none or not valid
                    ea_ColumnType = eaColumnClass.ColumnTypes.ctNone
                End If
            Catch ex As Exception
                ea_ColumnType = eaColumnClass.ColumnTypes.ctNone
            End Try
        End Set
    End Property

    Property DefaultValue As String
        Get
            Return Me.ea_DefaultValue
        End Get
        Set(value As String)
            Try
                ea_DefaultValue = value
            Catch ex As Exception
                ea_DefaultValue = ""
            End Try
        End Set
    End Property

    Property KeyField() As Boolean
        Get
            Return Me.ea_KeyField
        End Get
        Set(ByVal Value As Boolean)
            Try
                If (Me.ea_ColumnType = eaColumnClass.ColumnTypes.ctNone) OrElse _
                        (Me.ea_ColumnType = eaColumnClass.ColumnTypes.ctMemo) Then  ' if col type is none or memo
                    Me.ea_KeyField = False                                          ' cannot be key field
                Else                                                                ' else col can be key field
                    Me.ea_KeyField = Value                                          ' set key field property
                End If
            Catch ex As Exception
                Me.ea_KeyField = False
            End Try
        End Set
    End Property

    Property IsNullable As Boolean
        Get
            Return ea_IsNullable
        End Get
        Set(value As Boolean)
            Try
                ea_IsNullable = value
            Catch ex As Exception
                ea_IsNullable = True
            End Try
        End Set
    End Property

    ReadOnly Property OleDbDataType As OleDb.OleDbType
        Get
            Select Case ea_ColumnType
                Case ColumnTypes.ctBoolean
                    Return OleDb.OleDbType.Boolean
                Case ColumnTypes.ctCurrency
                    Return OleDb.OleDbType.Currency
                Case ColumnTypes.ctDate
                    Return OleDb.OleDbType.Date
                Case ColumnTypes.ctDouble
                    Return OleDb.OleDbType.Double
                Case ColumnTypes.ctInteger
                    Return OleDb.OleDbType.Integer
                Case ColumnTypes.ctMemo
                    Return OleDb.OleDbType.LongVarWChar
                Case ColumnTypes.ctNone
                    Return OleDb.OleDbType.Empty
                Case ColumnTypes.ctSmallInt
                    Return OleDb.OleDbType.SmallInt
                Case ColumnTypes.ctText
                    Return OleDb.OleDbType.VarWChar
                Case Else
                    Return OleDb.OleDbType.Empty
            End Select
        End Get
    End Property

    ReadOnly Property SqlColumnType() As String
        Get
            Try
                Select Case Me.ea_ColumnType
                    Case eaColumnClass.ColumnTypes.ctNone
                        Return ""
                    Case eaColumnClass.ColumnTypes.ctBoolean
                        Return eaColumnClass.sqlCnBoolean
                        'Return eaColumnClass.sqlCnYesNo
                    Case eaColumnClass.ColumnTypes.ctCurrency
                        Return eaColumnClass.sqlCnCurrency
                    Case eaColumnClass.ColumnTypes.ctDate
                        Return eaColumnClass.sqlCnDate
                    Case eaColumnClass.ColumnTypes.ctDouble
                        Return eaColumnClass.sqlCnDouble
                    Case eaColumnClass.ColumnTypes.ctInteger
                        Return eaColumnClass.sqlCnInteger
                    Case eaColumnClass.ColumnTypes.ctMemo
                        Return eaColumnClass.sqlCnMemo
                    Case eaColumnClass.ColumnTypes.ctText
                        Return eaColumnClass.sqlCnText
                    Case ColumnTypes.ctSmallInt
                        Return eaColumnClass.sqlCnSmallInt
                    Case Else
                        Return ""
                End Select
            Catch ex As Exception
                Return ""
            End Try
        End Get
    End Property

    Property TextLength() As Integer
        Get
            Return Me.ea_TextLength
        End Get
        Set(ByVal Value As Integer)
            Try
                If Me.ea_ColumnType = eaColumnClass.ColumnTypes.ctText Then ' if col type is test
                    Me.ea_TextLength = Value                                ' then set text length
                Else                                                        ' else not a a text col
                    Me.ea_TextLength = -1                                   ' set text length to -1
                End If
            Catch ex As Exception
                Me.ea_TextLength = -1
            End Try
        End Set
    End Property

#End Region

#Region " Methods "

#End Region

End Class
