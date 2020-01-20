Namespace EaSql

    Public Class PcdbSql

#Region " KeyWords Array "

        Shared ReadOnly KeyWords As String() = {"ABSOLUTE",
                                                "ACTION",
                                                "ADA",
                                                "ADD",
                                                "ADMIN",
                                                "AFTER",
                                                "AGGREGATE",
                                                "ALIAS",
                                                "ALL",
                                                "ALLOCATE",
                                                "ALTER",
                                                "AND",
                                                "ANY",
                                                "ARE",
                                                "ARRAY",
                                                "AS",
                                                "ASC",
                                                "ASSERTION",
                                                "AT",
                                                "AUTHORIZATION",
                                                "AVG",
                                                "BACKUP",
                                                "BEFORE",
                                                "BEGIN",
                                                "BETWEEN",
                                                "BINARY",
                                                "BIT",
                                                "BIT_LENGTH",
                                                "BLOB",
                                                "BOOLEAN",
                                                "BOTH",
                                                "BREADTH",
                                                "BREAK",
                                                "BROWSE",
                                                "BULK",
                                                "BY",
                                                "CALL",
                                                "CASCADE",
                                                "CASCADED",
                                                "CASE",
                                                "CAST",
                                                "CATALOG",
                                                "CHAR",
                                                "CHAR_LENGTH",
                                                "CHARACTER",
                                                "CHARACTER_LENGTH",
                                                "CHECK",
                                                "CHECKPOINT",
                                                "CLASS",
                                                "CLOB",
                                                "CLOSE",
                                                "CLUSTERED",
                                                "COALESCE",
                                                "COLLATE",
                                                "COLLATION",
                                                "COLUMN",
                                                "COMMIT",
                                                "COMPLETION",
                                                "COMPUTE",
                                                "CONNECT",
                                                "CONNECTION",
                                                "CONSTRAINT",
                                                "CONSTRAINTS",
                                                "CONSTRUCTOR",
                                                "CONTAINS",
                                                "CONTAINSTABLE",
                                                "CONTINUE",
                                                "CONVERT",
                                                "CORRESPONDING",
                                                "COUNT",
                                                "CREATE",
                                                "CROSS",
                                                "CUBE",
                                                "CURRENT",
                                                "CURRENT_DATE",
                                                "CURRENT_PATH",
                                                "CURRENT_ROLE",
                                                "CURRENT_TIME",
                                                "CURRENT_TIMESTAMP",
                                                "CURRENT_USER",
                                                "CURSOR",
                                                "CYCLE",
                                                "DATA",
                                                "DATABASE",
                                                "DATE",
                                                "DAY",
                                                "DBCC",
                                                "DEALLOCATE",
                                                "DEC",
                                                "DECIMAL",
                                                "DECLARE",
                                                "DEFAULT",
                                                "DEFERRABLE",
                                                "DEFERRED",
                                                "DELETE",
                                                "DENY",
                                                "DEPTH",
                                                "DEREF",
                                                "DESC",
                                                "DESCRIBE",
                                                "DESCRIPTOR",
                                                "DESTROY",
                                                "DESTRUCTOR",
                                                "DETERMINISTIC",
                                                "DIAGNOSTICS",
                                                "DICTIONARY",
                                                "DISCONNECT",
                                                "DISK",
                                                "DISTINCT",
                                                "DISTRIBUTED",
                                                "DOMAIN",
                                                "DOUBLE",
                                                "DROP",
                                                "DUMMY",
                                                "DUMP",
                                                "DYNAMIC",
                                                "EACH",
                                                "ELSE",
                                                "END",
                                                "END-EXEC",
                                                "EQUALS",
                                                "ERRLVL",
                                                "ESCAPE",
                                                "EVERY",
                                                "EXCEPT",
                                                "EXCEPTION",
                                                "EXEC",
                                                "EXECUTE",
                                                "EXISTS",
                                                "EXIT",
                                                "EXTERNAL",
                                                "EXTRACT",
                                                "FALSE",
                                                "FETCH",
                                                "FILE",
                                                "FILLFACTOR",
                                                "FIRST",
                                                "FLOAT",
                                                "FOR",
                                                "FOREIGN",
                                                "FORTRAN",
                                                "FOUND",
                                                "FREE",
                                                "FREETEXT",
                                                "FREETEXTTABLE",
                                                "FROM",
                                                "FULL",
                                                "FUNCTION",
                                                "GENERAL",
                                                "GET",
                                                "GLOBAL",
                                                "GO",
                                                "GOTO",
                                                "GRANT",
                                                "GROUP",
                                                "GROUPING",
                                                "HAVING",
                                                "HOLDLOCK",
                                                "HOST",
                                                "HOUR",
                                                "IDENTITY",
                                                "IDENTITY_INSERT",
                                                "IDENTITYCOL",
                                                "IF",
                                                "IGNORE",
                                                "IMMEDIATE",
                                                "IN",
                                                "INCLUDE",
                                                "INDEX",
                                                "INDICATOR",
                                                "INITIALIZE",
                                                "INITIALLY",
                                                "INNER",
                                                "INOUT",
                                                "INPUT",
                                                "INSENSITIVE",
                                                "INSERT",
                                                "INT",
                                                "INTEGER",
                                                "INTERSECT",
                                                "INTERVAL",
                                                "INTO",
                                                "IS",
                                                "ISOLATION",
                                                "ITERATE",
                                                "JOIN",
                                                "KEY",
                                                "KILL",
                                                "LANGUAGE",
                                                "LARGE",
                                                "LAST",
                                                "LATERAL",
                                                "LEADING",
                                                "LEFT",
                                                "LESS",
                                                "LEVEL",
                                                "LIKE",
                                                "LIMIT",
                                                "LINENO",
                                                "LOAD",
                                                "LOCAL",
                                                "LOCALTIME",
                                                "LOCALTIMESTAMP",
                                                "LOCATOR",
                                                "LOWER",
                                                "MAP",
                                                "MATCH",
                                                "MAX",
                                                "MIN",
                                                "MINUTE",
                                                "MODIFIES",
                                                "MODIFY",
                                                "MODULE",
                                                "MONTH",
                                                "NAMES",
                                                "NATIONAL",
                                                "NATURAL",
                                                "NCHAR",
                                                "NCLOB",
                                                "NEW",
                                                "NEXT",
                                                "NO",
                                                "NOCHECK",
                                                "NONCLUSTERED",
                                                "NONE",
                                                "NOT",
                                                "NULL",
                                                "NULLIF",
                                                "NUMERIC",
                                                "OBJECT",
                                                "OCTET_LENGTH",
                                                "OF",
                                                "OFF",
                                                "OFFSETS",
                                                "OLD",
                                                "ON",
                                                "ONLY",
                                                "OPEN",
                                                "OPENDATASOURCE",
                                                "OPENQUERY",
                                                "OPENROWSET",
                                                "OPENXML",
                                                "OPERATION",
                                                "OPTION",
                                                "OR",
                                                "ORDER",
                                                "ORDINALITY",
                                                "OUT",
                                                "OUTER",
                                                "OUTPUT",
                                                "OVER",
                                                "OVERLAPS",
                                                "PAD",
                                                "PARAMETER",
                                                "PARAMETERS",
                                                "PARTIAL",
                                                "PASCAL",
                                                "PATH",
                                                "PERCENT",
                                                "PLAN",
                                                "POSITION",
                                                "POSTFIX",
                                                "PRECISION",
                                                "PREFIX",
                                                "PREORDER",
                                                "PREPARE",
                                                "PRESERVE",
                                                "PRIMARY",
                                                "PRINT",
                                                "PRIOR",
                                                "PRIVILEGES",
                                                "PROC",
                                                "PROCEDURE",
                                                "PUBLIC",
                                                "RAISERROR",
                                                "READ",
                                                "READS",
                                                "READTEXT",
                                                "REAL",
                                                "RECONFIGURE",
                                                "RECURSIVE",
                                                "REF",
                                                "REFERENCES",
                                                "REFERENCING",
                                                "RELATIVE",
                                                "REPLICATION",
                                                "RESTORE",
                                                "RESTRICT",
                                                "RESULT",
                                                "RETURN",
                                                "RETURNS",
                                                "REVOKE",
                                                "RIGHT",
                                                "ROLE",
                                                "ROLLBACK",
                                                "ROLLUP",
                                                "ROUTINE",
                                                "ROW",
                                                "ROWCOUNT",
                                                "ROWGUIDCOL",
                                                "ROWS",
                                                "RULE",
                                                "SAVE",
                                                "SAVEPOINT",
                                                "SCHEMA",
                                                "SCOPE",
                                                "SCROLL",
                                                "SEARCH",
                                                "SECOND",
                                                "SECTION",
                                                "SELECT",
                                                "SEQUENCE",
                                                "SESSION",
                                                "SESSION_USER",
                                                "SET",
                                                "SETS",
                                                "SETUSER",
                                                "SHUTDOWN",
                                                "SIZE",
                                                "SMALLINT",
                                                "SOME",
                                                "SPACE",
                                                "SPECIFIC",
                                                "SPECIFICTYPE",
                                                "SQL",
                                                "SQLCA",
                                                "SQLCODE",
                                                "SQLERROR",
                                                "SQLEXCEPTION",
                                                "SQLSTATE",
                                                "SQLWARNING",
                                                "START",
                                                "STATE",
                                                "STATEMENT",
                                                "STATIC",
                                                "STATISTICS",
                                                "STRUCTURE",
                                                "SUBSTRING",
                                                "SUM",
                                                "SYSTEM_USER",
                                                "TABLE",
                                                "TEMPORARY",
                                                "TERMINATE",
                                                "TEXTSIZE",
                                                "THAN",
                                                "THEN",
                                                "TIME",
                                                "TIMESTAMP",
                                                "TIMEZONE_HOUR",
                                                "TIMEZONE_MINUTE",
                                                "TO",
                                                "TOP",
                                                "TRAILING",
                                                "TRAN",
                                                "TRANSACTION",
                                                "TRANSLATE",
                                                "TRANSLATION",
                                                "TREAT",
                                                "TRIGGER",
                                                "TRIM",
                                                "TRUE",
                                                "TRUNCATE",
                                                "TSEQUAL",
                                                "UNDER",
                                                "UNION",
                                                "UNIQUE",
                                                "UNKNOWN",
                                                "UNNEST",
                                                "UPDATE",
                                                "UPDATETEXT",
                                                "UPPER",
                                                "USAGE",
                                                "USE",
                                                "USER",
                                                "USING",
                                                "VALUE",
                                                "VALUES",
                                                "VARCHAR",
                                                "VARIABLE",
                                                "VARYING",
                                                "VIEW",
                                                "WAITFOR",
                                                "WHEN",
                                                "WHENEVER",
                                                "WHERE",
                                                "WHILE",
                                                "WITH",
                                                "WITHOUT",
                                                "WORK",
                                                "WRITE",
                                                "WRITETEXT",
                                                "YEAR",
                                                "ZONE"}

#End Region

        '#Region " Got SQL Param "

        '        Public Shared Function GotSqlParam(ByVal aId As Integer) As Boolean

        '            ' checks to see if the value in aId is a valid param for SQL select command
        '            '
        '            ' vars passed:
        '            '   aId - a id number
        '            ' 
        '            ' returns:
        '            '   TRUE - if aId <> Nothing And aId <> Pcm.idNotSet
        '            '   FALSE - if aId = Nothing or aId = Pcm.idNotSet

        '            Try
        '                If (aId <> Nothing) AndAlso (aId <> Pcm.idNotSet) Then
        '                    Return True
        '                Else
        '                    Return False
        '                End If
        '            Catch ex As Exception
        '                Return False
        '            End Try
        '        End Function

        '        Public Shared Function GotSqlParam(ByVal aDate As Date) As Boolean

        '            ' checks to see if the value in aDate is a valid param for SQL select command
        '            '
        '            ' vars passed:
        '            '   aDate - a date
        '            ' 
        '            ' returns:
        '            '   TRUE - if aDate <> Nothing And aDate <> Pcm.NoDate
        '            '   FALSE - if aDate = Nothing or aDate = Pcm.NoDate

        '            Try
        '                If (aDate <> Nothing) AndAlso (aDate <> Pcm.NoDate) Then
        '                    Return True
        '                Else
        '                    Return False
        '                End If
        '            Catch ex As Exception
        '                Return False
        '            End Try
        '        End Function

        '        Public Shared Function GotSqlParam(ByVal aString As String) As Boolean

        '            ' checks to see if the value in aId is a valid param for SQL select command
        '            '
        '            ' vars passed:
        '            '   aString - a string
        '            ' 
        '            ' returns:
        '            '   TRUE - if aString <> Nothing And aString <> ""
        '            '   FALSE - if aString = Nothing or aString = ""

        '            Try
        '                If (aString <> Nothing) AndAlso (aString <> "") Then
        '                    Return True
        '                Else
        '                    Return False
        '                End If
        '            Catch ex As Exception
        '                Return False
        '            End Try
        '        End Function

        '#End Region

        '#Region " SQL Column value string "

        '        Private Shared Function SqlColumnValueString(ByVal TableName As String,
        '                                                     ByVal ColumnName As String,
        '                                                     ByVal aValue As Object,
        '                                                     Optional ByVal aOperator As String = "=",
        '                                                     Optional ByVal NeedBrackets As Boolean = True) As String

        '            ' creates the selection string for one field
        '            '
        '            ' vars passed:
        '            '   TableName - Name of Table
        '            '   ColumnName - Name of column
        '            '   aValue - value for field
        '            '   aOperator - operator to use in selection
        '            '     valid values are "=", "<", ">", "<=", ">=", "<>"
        '            '       "IN", "LIKE" - when aValue is a string
        '            '   NeedBrackets - flag to use brackets
        '            '     TRUE if want "(TableName.ColumnName = Value)"
        '            '     FALSE if want "TableName.ColumnName = Value"
        '            '     
        '            ' returns:
        '            '   note: if TableName = "", then "TableName." is omitted in all return values
        '            '   if aValue is Nothing:
        '            '       "(TableName.ColumnName Is Null)"
        '            '   if aValue is a boolean:
        '            '       "(TableName.ColumnName = TRUE)" or "(TableName.ColumnName = FALSE)" 
        '            '       (no single quotes for boolean value)
        '            '   if aValue is a date:
        '            '       "(TableName.ColumnName op #mm/dd/yyyy#)" 
        '            '       where op is the operator
        '            '   if aValue is a string
        '            '       "(TableName.ColumnName op "XXX")" or "(TableName.ColumnName Is Null)"
        '            '       where op is the operator
        '            '   if aValue is a number
        '            '       "(TableName.ColumnName op NNN)" or "(TableName.ColumnName Is Null)"
        '            '       where op is the operator

        '            Dim ValueText As String = ""

        '            Try
        '                If Not aValue Is Nothing Then
        '                    If TypeOf aValue Is Date Then
        '                        ValueText = aValue.ToShortDateString
        '                    Else
        '                        ValueText = aValue.ToString
        '                    End If
        '                End If
        '            Catch ex As Exception
        '                ValueText = ""
        '            End Try

        '            aOperator = aOperator.ToUpper()
        '            Select Case aOperator
        '            ' normal comparison operators
        '                Case "=", "<", ">", "<=", ">=", "<>"
        '                ' do nothing

        '                ' if an operator for just strings
        '                Case "IN", "LIKE"
        '                    Try
        '                        If Not (aValue Is Nothing) Then             ' if got a value
        '                            If Not (TypeOf aValue Is String) Then   ' and value is not a string
        '                                aOperator = "="                     ' then use default operator
        '                            End If
        '                        Else                                        ' else no value
        '                            aOperator = "="                         ' then use default operator
        '                        End If
        '                    Catch ex As Exception
        '                        aOperator = "="
        '                    End Try

        '                Case Else
        '                    aOperator = "="
        '            End Select

        '            Dim SqlColValStr As String = ""                                     ' initial Sql Column Value string
        '            If NeedBrackets Then                                                ' if need brackets
        '                SqlColValStr = "("                                              ' then add starting bracket
        '            End If
        '            If TableName = "" Then                                              ' if no table name
        '                SqlColValStr &= ColumnName                                      ' then exclude TableName.
        '            Else                                                                ' else got table name
        '                SqlColValStr &= String.Format("{0}.{1}", TableName, ColumnName) ' include table name
        '            End If

        '            If TypeOf aValue Is Boolean Then
        '                SqlColValStr &= " = " & ValueText
        '            Else
        '                If ValueText = "" Then
        '                    If aOperator = "<>" Then
        '                        SqlColValStr &= " Is Not Null"
        '                    Else
        '                        SqlColValStr &= " Is Null"
        '                    End If
        '                Else
        '                    If TypeOf aValue Is Date Then                                           ' if value is a date
        '                        SqlColValStr &= String.Format(" {0} #{1}#", aOperator, ValueText)   ' format value text as a date
        '                    ElseIf TypeOf aValue Is Boolean Then                                    ' if value is a boolean
        '                        If aOperator = "<>" Then                                            ' if not equal
        '                            SqlColValStr &= " <> " & ValueText                              ' use not equal value text
        '                        Else                                                                ' else not using not equal
        '                            SqlColValStr &= " = " & ValueText                               ' alway use equal (no >, < for booleans)
        '                        End If
        '                    ElseIf TypeOf aValue Is String Then
        '                        SqlColValStr &= String.Format(" {0}""{1}""", aOperator, ValueText.Replace("'", "''"))
        '                    Else
        '                        SqlColValStr &= String.Format(" {0} {1}", aOperator, ValueText)
        '                    End If
        '                End If
        '            End If
        '            If NeedBrackets Then
        '                SqlColValStr &= ")"
        '            End If
        '            Return SqlColValStr
        '        End Function

        '#End Region

        '#Region " SQL WHERE "

        '        Public Shared Sub SqlRemoveWhere(ByRef SqlSelectText As String)

        '            ' removes the ending " WHERE" and text after from a Sql Select command
        '            '
        '            ' vars passed:
        '            '   SqlSelectText - sql select command text to remove "where" from.  
        '            '     Note: Passed ByRef

        '            Try
        '                Dim WherePos As Integer = SqlSelectText.ToUpper().LastIndexOf(" WHERE")
        '                If WherePos > 0 Then
        '                    SqlSelectText = SqlSelectText.Remove(WherePos, SqlSelectText.Length - WherePos)
        '                End If
        '            Catch ex As Exception
        '                Return
        '            End Try
        '        End Sub

        '        Public Shared Function SqlWhereColumnValueString(ByVal TableName As String,
        '                                                         ByVal ColumnName As String,
        '                                                         ByVal aValue As Object,
        '                                                         Optional ByVal aOperator As String = "=",
        '                                                         Optional ByVal NeedBrackets As Boolean = False) As String

        '            ' creates the selection string for one field
        '            '
        '            ' vars passed:
        '            '   TableName - Name of Table
        '            '   ColumnName - Name of column
        '            '   aValue - value for field
        '            '   aOperator - operator to use in selection
        '            '     valid values are "=", "<", ">", "<=", ">=", "<>"
        '            '       "IN", "LIKE" - when aValue is a string
        '            '   NeedBrackets - flag to use brackets
        '            '     TRUE if want "(TableName.ColumnName = Value)"
        '            '     FALSE if want "TableName.ColumnName = Value"
        '            '     
        '            ' returns:
        '            '   note: if TableName = "", then "TableName." is omitted in all return values
        '            '   if aValue is Nothing:
        '            '       "(TableName.ColumnName Is Null)"
        '            '   if aValue is a boolean:
        '            '       "(TableName.ColumnName = TRUE)" or "(TableName.ColumnName = FALSE)" 
        '            '       (no single quotes for boolean value)
        '            '   if aValue is a date:
        '            '       "(TableName.ColumnName op #mm/dd/yyyy#)" 
        '            '       where op is the operator
        '            '   if aValue is a string
        '            '       "(TableName.ColumnName op "XXX")" or "(TableName.ColumnName Is Null)"
        '            '       where op is the operator
        '            '   if aValue is a number
        '            '       "(TableName.ColumnName op NNN)" or "(TableName.ColumnName Is Null)"
        '            '       where op is the operator

        '            Try
        '                Return SqlColumnValueString(TableName, ColumnName, aValue, aOperator, NeedBrackets)
        '            Catch ex As Exception
        '                Return ""
        '            End Try
        '        End Function

        '        Public Shared Function SqlWhereColumnValueString(ByVal aParamCount As Integer,
        '                                                         ByVal TableName As String,
        '                                                         ByVal ColumnName As String,
        '                                                         ByVal aValue As Object,
        '                                                         Optional ByVal aOperator As String = "=",
        '                                                         Optional ByVal aWhereOpp As String = "AND",
        '                                                         Optional ByVal NeedBrackets As Boolean = False) As String

        '            ' creates the selection string for one field
        '            '
        '            ' vars passed:
        '            '   aParamCount - how many WHERE parameters already
        '            '   TableName - Name of Table
        '            '   ColumnName - Name of column
        '            '   aValue - value for field
        '            '   aOperator - operator to use in selection
        '            '     valid values are "=", "<", ">", "<=", ">=", "<>"
        '            '       "IN", "LIKE" - when aValue is a string
        '            '   aWhereOpp - logical operator (AND/OR)
        '            '   NeedBrackets - flag to use brackets
        '            '     TRUE if want "(TableName.ColumnName = Value)"
        '            '     FALSE if want "TableName.ColumnName = Value"
        '            '     
        '            ' returns:
        '            '   note: if TableName = "", then "TableName." is omitted in all return values
        '            '   if aValue is Nothing:
        '            '       "(TableName.ColumnName Is Null)"
        '            '   if aValue is a boolean:
        '            '       "(TableName.ColumnName = TRUE)" or "(TableName.ColumnName = FALSE)" 
        '            '       (no single quotes for boolean value)
        '            '   if aValue is a date:
        '            '       "(TableName.ColumnName op #mm/dd/yyyy#)" 
        '            '       where op is the operator
        '            '   if aValue is a string
        '            '       "(TableName.ColumnName op "XXX")" or "(TableName.ColumnName Is Null)"
        '            '       where op is the operator
        '            '   if aValue is a number
        '            '       "(TableName.ColumnName op NNN)" or "(TableName.ColumnName Is Null)"
        '            '       where op is the operator
        '            '   if aParamCount > 0, add "AND " or "OR" to start of return value

        '            Dim WhereSql As String = ""                 ' initial Where string
        '            Try
        '                If aParamCount > 0 Then                 ' if not first param
        '                    aWhereOpp = aWhereOpp.ToUpper
        '                    If (aWhereOpp <> "AND") AndAlso (aWhereOpp <> "OR") Then
        '                        aWhereOpp = "AND"
        '                    End If
        '                    WhereSql = String.Format(" {0} ", aWhereOpp)    ' add where opp (space before and after)
        '                End If
        '                Return WhereSql & SqlColumnValueString(TableName, ColumnName, aValue, aOperator, NeedBrackets)
        '            Catch ex As Exception
        '                Return ""
        '            End Try
        '        End Function

        '#End Region

        '#Region " SQL INNER JOIN "

        '        Public Shared Function InnerJoin(ByVal JoinToTableName As String,
        '                                         ByVal JoinToColumnName As String,
        '                                         ByVal JoinFromTableName As String,
        '                                         ByVal JoinFromColumnName As String,
        '                                         Optional ByVal InsertINNERJOIN As Boolean = False,
        '                                         Optional ByVal InsertAND As Boolean = False) As String

        '            ' gets the "INNER JOIN" SQL command text.
        '            '
        '            ' vars passed:
        '            '   JoinToTableName - table name to join to
        '            '   JoinToColumnName - column name in JoinToTable to join to JoinFromColumnName
        '            '   JoinFromTableName - table name to join from (main table in query)
        '            '   JoinFromColumnName - column name in JoinFromTable to join to JoinToColumnName
        '            '   InsertINNERJOIN - if TRUE, add text " INNER JOIN JoinToTableName ON " to start of InnerJoin text
        '            '   InsertAND - if TRUE and InsertINNERJOIN is FALSE, add " AND " to start of InnerJoin text
        '            '
        '            ' returns:
        '            '   if InsertINNERJOIN = TRUE 
        '            '     "FROM (JoinFromTableName INNER JOIN 
        '            '       JoinToTableName ON JoinToTableName.JoinToColumnName = JoinFromTAbleName.JoinFromColumnName)"
        '            '   if InsertAND = TRUE and InsertINNERJOIN = FALSE
        '            '     "AND JoinToTableName.JoinToColumnName = JoinFromTAbleName.JoinFromColumnName)"
        '            '   else
        '            '     "(JoinToTableName.JoinToColumnName = JoinFromTAbleName.JoinFromColumnName)"
        '            '   "" - something went wrong

        '            Try
        '                Dim IjStr As String = ""
        '                If InsertINNERJOIN Then                                 ' if want to insert inner join text
        '                    IjStr = String.Format("FROM ({0} INNER JOIN {1} ON ", JoinFromTableName, JoinToTableName) ' add inner join SQL text
        '                ElseIf InsertAND Then                                   ' else if do not want inner join but want AND
        '                    IjStr = "AND "                                      ' add AND SQL text
        '                End If
        '                ' add in join SQL text
        '                IjStr &= String.Format("{0}.{1} = {2}.{3})", JoinFromTableName, JoinFromColumnName, JoinToTableName, JoinToColumnName)
        '                Return IjStr
        '            Catch ex As Exception
        '                Return ""
        '            End Try
        '        End Function

        '        Public Shared Function AddOnInnerJoinAnd(ByVal InnerJoinSql As String, ByVal JoinToAdd As String) As String

        '            ' adds an additional join element to the INNER JOIN SQL 
        '            '
        '            ' vars passed:
        '            '   InnerJoinSql - current INNER JOIN SQL statement, ends with ")"
        '            '   JoinToAdd - join element to add, starts with " AND " and ends with ")"
        '            '
        '            ' returns:
        '            '   InnerJoinSql with out the ending ")", and the JoinToAdd added to it
        '            '   "" - something went wrong

        '            Try
        '                If InnerJoinSql.IndexOf(")") = InnerJoinSql.Length - 1 Then
        '                    InnerJoinSql = InnerJoinSql.Remove(InnerJoinSql.Length - 1, 1)
        '                End If

        '                Return String.Format("{0} {1}", InnerJoinSql, JoinToAdd)
        '            Catch ex As Exception
        '                Return ""
        '            End Try
        '        End Function

        '#End Region

        '#Region " SQL SET "

        '        Public Shared Function SqlSetColumnValueString(ByVal TableName As String,
        '                                                       ByVal ColumnName As String,
        '                                                       ByVal aValue As Object,
        '                                                       Optional ByVal aOperator As String = "=",
        '                                                       Optional ByVal NeedBrackets As Boolean = False) As String

        '            ' creates the selection string for one field
        '            '
        '            ' vars passed:
        '            '   TableName - Name of Table
        '            '   ColumnName - Name of column
        '            '   aValue - value for field
        '            '   aOperator - operator to use in selection
        '            '     valid values are "=", "<", ">", "<=", ">=", "<>"
        '            '       "IN", "LIKE" - when aValue is a string
        '            '   NeedBrackets - flag to use brackets
        '            '     TRUE if want "(TableName.ColumnName = Value)"
        '            '     FALSE if want "TableName.ColumnName = Value"
        '            '     
        '            ' returns:
        '            '   note: if TableName = "", then "TableName." is omitted in all return values
        '            '   if aValue is Nothing:
        '            '       "TableName.ColumnName Is Null"
        '            '   if aValue is a boolean:
        '            '       "TableName.ColumnName = TRUE" or "TableName.ColumnName = FALSE" 
        '            '       (no single quotes for boolean value)
        '            '   if aValue is a date:
        '            '       "TableName.ColumnName op #mm/dd/yyyy#" 
        '            '       where op is the operator
        '            '   if aValue is a string
        '            '       "TableName.ColumnName op "XXX"" or "TableName.ColumnName Is Null"
        '            '       where op is the operator
        '            '   if aValue is a number
        '            '       "TableName.ColumnName op NNN" or "TableName.ColumnName Is Null"
        '            '       where op is the operator

        '            Try
        '                Return SqlColumnValueString(TableName, ColumnName, aValue, aOperator, NeedBrackets)
        '            Catch ex As Exception
        '                Return ""
        '            End Try
        '        End Function

        '#End Region

#Region " Execute Non Query "

        Public Shared Function ExecuteNonQuery(ByVal SqlText As String,
                                               ByVal OleDbConn As OleDb.OleDbConnection,
                                               ByVal ErrMsg As String,
                                               ByVal ErrTitle As String,
                                               Optional ByVal ParamNames() As String = Nothing,
                                               Optional ByVal ParamValues() As Object = Nothing) As Integer

            ' executes a non query SQL command
            '
            ' vars passed:
            '   SqlText - SqlText to use
            '   OleDbConn - ole db connection to use
            '   ErrMsg - error message to display
            '   ErrTitle - error title for error in executing non query
            '   ParamNames() - names of parameters in query
            '   ParamValues() - values of parameters in query
            '
            ' returns:
            '   >= 0 - no errors
            '   <0 - SqlCommand.ExecuteNonQuery error value
            '   Pcm.sqlErr - other error in sql 

            Dim SqlCommand As OleDb.OleDbCommand = Nothing                          ' sql command
            Dim EnqReturn As Integer = 0                                            ' exec non query return value
            Dim Done As Boolean = False
            Dim TryCount As Integer = 0
            Do
                Try
                    SqlCommand = New OleDb.OleDbCommand(SqlText, OleDbConn)         ' get the connection
                    SqlCommand.Connection.Open()                                    ' open the connection

                    ' if got ParamNames and ParamValues and # of ParamNames = # of ParamValues
                    If ParamNames IsNot Nothing AndAlso ParamValues IsNot Nothing AndAlso ParamNames.Count = ParamValues.Count Then
                        If SqlCommand.Parameters.Count = 0 Then                     ' sql command does not have params saved
                            For p As Integer = 0 To ParamNames.Count - 1            ' for each param name
                                SqlCommand.Parameters.AddWithValue(ParamNames(p), ParamValues(p)) ' add name with value
                            Next
                        Else                                                        ' else if SQL command has params saved 
                            Dim i As Integer                                        ' parameter index 
                            For p As Integer = 0 To ParamNames.Count - 1            ' for each param name
                                i = SqlCommand.Parameters.IndexOf(ParamNames(p))    ' find param name is list of params
                                If i >= 0 Then                                      ' if found param name in list of params
                                    SqlCommand.Parameters(i).Value = ParamValues(p) ' set the param value
                                End If
                            Next
                        End If
                    End If

                    EnqReturn = SqlCommand.ExecuteNonQuery()                        ' execute the non query
                    Done = True                                                     ' done with loop
                Catch OleEx As OleDb.OleDbException
                    If (OleEx.ErrorCode = Pcm.oleCannotOpen OrElse OleEx.ErrorCode = Pcm.oleCannotModify) AndAlso TryCount < Pcm.oleMaxTries Then
                        TryCount += 1                                               ' incremnet try count
                        'My.Application.DoEvents()                                   ' make sure everything is done

                        Debug.WriteLine("No DoEvents")

                        System.Threading.Thread.Sleep(Pcm.oleWaitMiliSeconds)       ' wait for a time
                    Else                                                            ' else got an unhandled error
                        MessageBox.Show(String.Format("{0}{1}Error Code:{2}{1}{3}{1}{4}",
                                                  ErrMsg, vbCrLf, CStr(OleEx.ErrorCode), OleDbConn.DataSource, OleEx.Message),
                                    ErrTitle, MessageBoxButtons.OK, MessageBoxIcon.Error) ' show error
                        EnqReturn = OleEx.ErrorCode                                 ' set return error
                        Done = True                                                 ' ok to exit 
                    End If
                Catch ex As Exception
                    MessageBox.Show(String.Format("{0}{1}{2}{1}{3}", ErrMsg, vbCrLf, OleDbConn.DataSource, ex.Message),
                                ErrTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return Pcm.sqlErr
                Finally
                    Try
                        If SqlCommand IsNot Nothing Then                                    ' if got a command
                            If SqlCommand.Connection.State <> ConnectionState.Closed Then   ' if connection not closed
                                SqlCommand.Connection.Close()                               ' close the connection
                            End If
                            SqlCommand.Dispose()                                            ' free sql command
                            SqlCommand = Nothing                                            ' set sql command to nothing 
                        End If
                    Catch ex As Exception
                        MessageBox.Show(String.Format("{0}  {1}", ErrMsg, ex.Message),
                                    Pcm.etSqlCanNotCloseErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                        EnqReturn = Pcm.sqlErr
                        Done = True
                    End Try
                End Try
            Loop Until Done
            Return EnqReturn
        End Function

        Public Shared Function ExecuteNonQuery(ByVal SqlText As String,
                                               ByVal OleDbConn As OleDb.OleDbConnection,
                                               ByVal ErrMsg As String,
                                               ByVal ErrTitle As String,
                                               ByVal ErrToHandle As Integer,
                                               ByVal HandledReturn As Integer) As Integer

            ' executes a non query for the 
            '
            ' vars passed:
            '   SqlText - SqlText to use
            '   OleDbConn - ole db connection to use
            '   ErrMsg - error message to display
            '   ErrTitle - error title for error in executing non query
            '   ErrToHandle - error value to handle
            '   HandledReturn - return value for error handled
            '
            ' returns:
            '   >= 0 - no errors
            '   HandledReturn - return value if error handled
            '   <0 - SqlCommand.ExecuteNonQuery error value
            '   Pcm.sqlErr - other error in sql 

            Dim SqlCommand As OleDb.OleDbCommand = Nothing                  ' sql command
            Dim EnqReturn As Integer = 0                                    ' exec non query return value
            Dim Done As Boolean = False
            Dim TryCount As Integer = 0
            Do
                Try
                    SqlCommand = New OleDb.OleDbCommand(SqlText, OleDbConn)     ' get the connection
                    SqlCommand.Connection.Open()                                ' open the connection
                    EnqReturn = SqlCommand.ExecuteNonQuery()                    ' execute the non query
                    Done = True                                                 ' done with loop
                Catch OleEx As OleDb.OleDbException
                    If OleEx.ErrorCode = ErrToHandle Then                       ' if got the error to handle
                        EnqReturn = HandledReturn                               ' set to return the handled return value
                        Done = True                                             ' done with loop
                    ElseIf (OleEx.ErrorCode = Pcm.oleCannotOpen OrElse OleEx.ErrorCode = Pcm.oleCannotModify) AndAlso TryCount < Pcm.oleMaxTries Then
                        TryCount += 1                                           ' increment try count
                        'My.Application.DoEvents()                               ' make sure everything is done

                        Debug.WriteLine("No DoEvents")

                        System.Threading.Thread.Sleep(Pcm.oleWaitMiliSeconds)   ' wait some time
                    Else                                                        ' else got an unhandled error
                        MessageBox.Show(String.Format("{0}{1}Error Code:{2}{1}{3}{1}{4}",
                                                  ErrMsg, vbCrLf, CStr(OleEx.ErrorCode), OleDbConn.DataSource, OleEx.Message),
                                    ErrTitle, MessageBoxButtons.OK, MessageBoxIcon.Error) ' show error
                        EnqReturn = OleEx.ErrorCode                             ' return error
                        Done = True                                             ' ok to exit
                    End If
                Catch ex As Exception
                    MessageBox.Show(String.Format("{0}{1}{2}{1}{3}", ErrMsg, vbCrLf, OleDbConn.DataSource, ex.Message),
                                ErrTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return Pcm.sqlErr
                Finally
                    Try
                        If SqlCommand IsNot Nothing Then                                    ' if got a command
                            If SqlCommand.Connection.State <> ConnectionState.Closed Then   ' if connection not closed
                                SqlCommand.Connection.Close()                               ' close the connection
                            End If
                            SqlCommand.Dispose()                                            ' free sql command
                            SqlCommand = Nothing                                            ' set sql command to nothing 
                        End If
                    Catch ex As Exception
                        MessageBox.Show(String.Format("{0}  {1}", ErrMsg, ex.Message),
                                    Pcm.etSqlCanNotCloseErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                        EnqReturn = Pcm.sqlErr
                        Done = True                                                         ' Ok to exit now
                    End Try
                End Try
            Loop Until Done
            Return EnqReturn
        End Function

#End Region

#Region " Execute Scalar "

        Public Shared Function ExecuteScalar(ByVal SqlText As String,
                                             ByVal OleDbConn As OleDb.OleDbConnection,
                                             ByVal ErrMsg As String,
                                             ByVal ErrTitle As String) As Integer

            ' executes a scalar SQL command
            '
            ' vars passed:
            '   SqlText - SqlText to use
            '   OleDbConn - ole db connection to use
            '   ErrMsg - error message to display
            '   ErrTitle - error title for error in executing non query
            '
            ' returns:
            '   >= 0 - no errors
            '   <0 - SqlCommand.ExecuteNonQuery error value
            '   Pcm.sqlErr - other error in sql 

            Dim SqlCommand As OleDb.OleDbCommand = Nothing                      ' sql command
            Dim ScalarReturn As Integer = 0                                     ' scalar command return value
            Dim Done As Boolean = False
            Dim TryCount As Integer = 0
            Do
                Try
                    SqlCommand = New OleDb.OleDbCommand(SqlText, OleDbConn)     ' get the connection
                    SqlCommand.Connection.Open()                                ' open the connection
                    ScalarReturn = Convert.ToInt32(SqlCommand.ExecuteScalar())  ' execute the scalar command
                    Done = True                                                 ' done with loop
                Catch OleEx As OleDb.OleDbException
                    If (OleEx.ErrorCode = Pcm.oleCannotOpen OrElse OleEx.ErrorCode = Pcm.oleCannotModify) AndAlso TryCount < Pcm.oleMaxTries Then
                        TryCount += 1                                           ' increment try count
                        'My.Application.DoEvents()                               ' make sure everything is done

                        Debug.WriteLine("No DoEvents")

                        System.Threading.Thread.Sleep(Pcm.oleWaitMiliSeconds)   ' wait for a time
                    Else                                                        ' else got an unhandled error
                        MessageBox.Show(String.Format("{0}{1}Error Code:{2}{1}{3}{1}{4}",
                                                  ErrMsg, vbCrLf, CStr(OleEx.ErrorCode), OleDbConn.DataSource, OleEx.Message),
                                    ErrTitle, MessageBoxButtons.OK, MessageBoxIcon.Error) ' show error
                        ScalarReturn = OleEx.ErrorCode                          ' set return error
                        Done = True                                             ' ok to exit 
                    End If
                Catch ex As Exception
                    MessageBox.Show(String.Format("{0}{1}{2}{1}{3}", ErrMsg, vbCrLf, OleDbConn.DataSource, ex.Message),
                                ErrTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return Pcm.sqlErr
                Finally
                    Try
                        If SqlCommand IsNot Nothing Then                                    ' if got a command
                            If SqlCommand.Connection.State <> ConnectionState.Closed Then   ' if connection not closed
                                SqlCommand.Connection.Close()                               ' close the connection
                            End If
                            SqlCommand.Dispose()                                            ' free sql command
                            SqlCommand = Nothing                                            ' set sql command to nothing 
                        End If
                    Catch ex As Exception
                        MessageBox.Show(String.Format("{0}  {1}", ErrMsg, ex.Message),
                                    Pcm.etSqlCanNotCloseErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                        ScalarReturn = Pcm.sqlErr
                        Done = True
                    End Try
                End Try
            Loop Until Done
            Return ScalarReturn
        End Function

#End Region

#Region " Update via SQL "

        Public Shared Function UpdateViaSql(ByVal UpdateSqlText As String,
                                            ByVal aOleDbConnection As OleDb.OleDbConnection,
                                            ByVal ErrMsg As String,
                                            ByVal ErrTitle As String) As Integer

            ' Updates a table via a SQL command.  Makes sure sql command starts with "UPDATE", and then calls
            '   ExecuteNonQuery to do the query.  
            '
            ' vars passed:
            '   UpdateSqlText - SqlText to use
            '   aOleDbConnection - ole db connection to use
            '   ErrMsg - error message to display
            '   ErrTitle - error title for error in executing non query
            '
            ' returns:
            '   >= 0 - no errors
            '   <0 - SqlCommand.ExecuteNonQuery error value
            '   Pcm.sqlErr- other error in sql 

            Try
                If Not UpdateSqlText.ToUpper.StartsWith("UPDATE".ToUpper) Then
                    MessageBox.Show(String.Format("The Update SQL command does not start with ""UPDATE"".  SQL Command Text:{0}{1}", vbCrLf, UpdateSqlText),
                                Pcm.etUpdateErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
                ' do the insert via SQL
                Dim UpdateReturn As Integer = ExecuteNonQuery(UpdateSqlText, aOleDbConnection, ErrMsg, ErrTitle)
                If UpdateReturn < 0 Then        ' if got an err
                    Return UpdateReturn         ' return the error
                End If
                Return UpdateReturn
            Catch ex As Exception
                MessageBox.Show(String.Format("{0}  {1}", ErrMsg, ex.Message),
                            ErrTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return Pcm.sqlErr
            End Try
        End Function

#End Region

#Region " Insert/Add Row via SQL "

        Public Shared Function InsertViaSql(ByVal InsertSqlText As String,
                                            ByVal aOleDbConnection As OleDb.OleDbConnection,
                                            ByVal ErrMsg As String,
                                            ByVal ErrTitle As String,
                                            Optional ByVal ParamNames() As String = Nothing,
                                            Optional ByVal ParamValues() As Object = Nothing) As Integer

            ' Inserts data into a table via a SQL command.  Makes sure sql command starts with "INSERT INTO ", 
            '   and then calls ExecuteNonQuery to do the query.  
            '
            ' vars passed:
            '   InsertSqlText - SqlText to use
            '   aOleDbConnection - ole db connection to use
            '   ErrMsg - error message to display
            '   ErrTitle - error title for error in executing non query
            '   ParamNames() - names of parameters in query
            '   ParamValues() - values of parameters in query
            '
            ' note: ParamValues(i) has the value for the parameter named ParamNames(i)
            '
            ' returns:
            '   >= 0 - # of rows inserted
            '   Pcm.sqlErr - other error in sql 
            '   Pcm.teOtherErr - other error

            Try
                If Not InsertSqlText.ToUpper.StartsWith("INSERT INTO".ToUpper) Then
                    MessageBox.Show(String.Format("The Insert SQL command does not start with ""INSERT INTO"".  SQL Command Text:{0}{1}", vbCrLf, InsertSqlText),
                                Pcm.etInsertErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
                ' do the insert via SQL
                Dim InsertReturn As Integer = ExecuteNonQuery(InsertSqlText, aOleDbConnection, ErrMsg, ErrTitle, ParamNames, ParamValues)
                If InsertReturn < 0 Then            ' if got an error
                    Return InsertReturn             ' return the error
                End If

                Return InsertReturn
            Catch ex As Exception
                MessageBox.Show(String.Format("{0}  {1}", ErrMsg, ex.Message),
                            Pcm.etInsertErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return Pcm.sqlErr
            End Try
        End Function

        Public Shared Function AddRowViaSql(ByVal aRow As DataRow,
                                            ByVal aTableName As String,
                                            ByVal OleDbConn As OleDb.OleDbConnection,
                                            ByVal InsertErrMsg As String) As Integer

            ' inserts a new row into a history table via sql.  This way the table does not 
            ' need to be loaded for a new row to be inserted.  
            '
            ' vars passed:
            '   aRow - history data row to be inserted
            '   aTableName - name of table to insert into
            '   OleDbConn - connection to use
            '   InsertErrMsg - insert error message
            '
            ' returns:
            '   Pcm.NoErrors - row was inserted
            '   Pcm.teParametersErr - one column in Row was Null
            '   <0 - SqlCommand.ExecuteNonQuery error value
            '   Pcm.sqlErr - other error in sql 

            Try
                ' set starting values  
                Dim InsertInto As String = String.Format("INSERT INTO {0} (", aTableName)
                Dim Values As String = "VALUES ("

                Dim c As Integer
                For c = 0 To aRow.Table.Columns.Count - 1
                    AddColumnParamsToInsert(InsertInto, Values, aRow, aRow.Table.Columns(c).ColumnName)
                Next

                InsertInto &= ") "      ' add in trailing ")", and add a space after
                Values &= ")"           ' add in trailing ")"

                Dim AddSqlErr As Integer = PcdbSql.ExecuteNonQuery(InsertInto & Values, OleDbConn,
                        InsertErrMsg, "SQL Insert Error")   ' do the insert via SQL
                If AddSqlErr < 0 Then           ' if got an error
                    Return AddSqlErr            ' return the error
                End If

                Return Pcm.NoErrors             ' if got here, return no errors
            Catch ex As Exception
                MessageBox.Show(String.Format("{0}  {1}", InsertErrMsg, ex.Message),
                            Pcm.etInsertErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return Pcm.teOtherErr
            End Try
        End Function

        Private Shared Sub AddColumnParamsToInsert(ByRef InsertStr As String,
                                                   ByRef ValueStr As String,
                                                   ByVal aRow As DataRow,
                                                   ByVal ColumnName As String)

            ' updated the Insert string and Values string with the column name and value
            ' When combined, the Insert and Value strings make an INSERT SQL statement
            ' 
            ' NOTE: InsertStr and ValueStr are passed ByRef
            '
            ' vars passed:
            '   InsertStr - ByRef - insert part of the INSERT SQL statement
            '   ValueStr - ByRef - values part of the INSERT SQL statement
            '   aRow - datarow with a value in column ColumnName
            '   aColumnName - name of a column to include in the INSERT SQL statement

            Try
                If aRow.IsNull(ColumnName) Then    ' if no value in column
                    Exit Sub                        ' exit now, do nothing
                End If

                If Not InsertStr.EndsWith("(") Then                             ' if not the first entry
                    InsertStr &= ", "                                           ' add a comma separator
                    ValueStr &= ", "
                End If
                InsertStr &= ColumnName                                        ' add in value column name
                If TypeOf aRow(ColumnName) Is Date Then                        ' if a date value
                    ValueStr &= String.Format("#{0}#", CStr(aRow(ColumnName))) ' add in value, enclosed with #
                ElseIf TypeOf aRow(ColumnName) Is String Then                  ' else if a string
                    ' add in value, enclosed with '.  also, if text has a ', then convert to ''
                    ValueStr &= String.Format("'{0}'", aRow(ColumnName).ToString.Replace("'", "''"))
                Else                                                            ' else a regular value
                    ValueStr &= aRow(ColumnName).ToString                      ' add in value 
                End If
            Catch ex As Exception
                MessageBox.Show("Unknown error while generating SQL INSERT statement.  " & ex.Message,
                            Pcm.etInternalErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

#End Region

#Region " Delete Row via SQL "

        Public Shared Function DeleteRowViaSQL(ByVal aRow As DataRow,
                                               ByVal TableName As String,
                                               ByVal OleDbConn As OleDb.OleDbConnection,
                                               ByVal DeleteErrMsg As String) As Integer

            ' deletes a row from a table using SQL.  Note, the table must be keyed
            '
            ' vars passed:
            '   aRow - row with key values matching row tp be deleted
            '   aTableName - name of table to delete row from
            '   OleDbConn - connection to use
            '   DeleteErrMsg - error message if something goes wrong
            '
            ' returns:
            '   >= 0 - # of rows deleted
            '   <0 - SqlCommand.ExecuteNonQuery error value
            '   Pcm.sqlErr - other error in sql 
            '   Pcm.OtherExErr - something went wrong

            Try
                If aRow.Table.PrimaryKey.Length = 0 Then                ' if no key field
                    Return 0                                            ' then return no rows deleted
                End If

                ' set the first line of the DELETE command 
                Dim DeleteSQL As String = "DELETE "
                Dim k As Integer = 0
                For Each KeyCol As DataColumn In aRow.Table.PrimaryKey                  ' for each key column 
                    If k > 0 Then                                                       ' if not the first key column
                        DeleteSQL &= ", "                                               ' add a comma and a space
                    End If
                    DeleteSQL &= String.Format("{0}.{1}", TableName, KeyCol.ColumnName) ' add the Table.KeyColumnName
                    k += 1                                                              ' increment key column counter
                Next
                DeleteSQL &= String.Format(" FROM {0} WHERE ", TableName)               ' add the 2nd line, FROM Table, and start 3rd line
                k = 0                                                                   ' reset key column counter (used as param count)
                For Each KeyCol As DataColumn In aRow.Table.PrimaryKey                  ' for each key column
                    ' add the WHERE criteria or each key column
                    DeleteSQL &= EaSql.Sql.SqlWhereColumnValueString(k, TableName, KeyCol.ColumnName, aRow(KeyCol), , , True)
                    k += 1                                                              ' increment key column counter
                Next

                ' delete the row
                Return PcdbSql.ExecuteNonQuery(DeleteSQL, OleDbConn, DeleteErrMsg, "SQL Delete Error")

                Return Pcm.NoErrors             ' if got here, return no errors
            Catch ex As Exception
                MessageBox.Show(String.Format("{0}  {1}", DeleteErrMsg, ex.Message),
                            Pcm.etDeleteRow, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return Pcm.OtherExErr
            End Try
        End Function

        Public Shared Function DeleteRowViaSQL(ByVal KeyValues() As Object,
                                               ByVal Table As DataTable,
                                               ByVal OleDbConn As OleDb.OleDbConnection,
                                               ByVal DeleteErrMsg As String) As Integer

            ' deletes a row from a table using SQL.  Note, the table must be keyed
            '
            ' vars passed:
            '   KeyValues - key values matching row to be deleted.  values MUST be in teh same order as listed in Table.PrimaryKey
            '   Table - table to delete row from
            '   OleDbConn - connection to use
            '   DeleteErrMsg - error message if something goes wrong
            '
            ' returns:
            '   >= 0 - # of rows deleted
            '   <0 - SqlCommand.ExecuteNonQuery error value
            '   Pcm.sqlErr - other error in sql 
            '   Pcm.OtherExErr - something went wrong

            Try
                If KeyValues.Length = 0 Then                            ' if no key field(s)
                    Return 0                                            ' then return no rows deleted
                End If

                ' set the first line of the DELETE command 
                Dim DeleteSQL As String = "DELETE "
                Dim k As Integer = 0
                For Each KeyCol As DataColumn In Table.PrimaryKey                       ' for each key column 
                    If k > 0 Then                                                       ' if not the first key column
                        DeleteSQL &= ", "                                               ' add a comma and a space
                    End If
                    DeleteSQL &= String.Format("{0}.{1}", Table.TableName, KeyCol.ColumnName) ' add the Table.KeyColumnName
                    k += 1                                                              ' increment key column counter
                Next
                DeleteSQL &= String.Format(" FROM {0} WHERE ", Table.TableName)         ' add the 2nd line, FROM Table, and start 3rd line
                k = 0                                                                   ' reset key column counter (used as param count)
                For Each KeyCol As DataColumn In Table.PrimaryKey                       ' for each key column
                    ' add the WHERE criteria or each key column
                    DeleteSQL &= EaSql.Sql.SqlWhereColumnValueString(k + 1, Table.TableName, KeyCol.ColumnName, KeyValues(k), , , True)
                Next

                ' delete the row
                Return PcdbSql.ExecuteNonQuery(DeleteSQL, OleDbConn, DeleteErrMsg, "SQL Delete Error")
            Catch ex As Exception
                MessageBox.Show(String.Format("{0}  {1}", DeleteErrMsg, ex.Message),
                            Pcm.etDeleteRow, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return Pcm.teOtherErr
            End Try
        End Function

#End Region

#Region " Table Actions "

#Region " Alter Table "

        Public Shared Function AlterTable(ByVal OleDbConn As OleDb.OleDbConnection,
                                          ByVal TableName As String,
                                          ByVal SqlText As String) As Integer

            ' alters a data table
            '
            ' vars passed:
            '   OleDbConn - connection to use
            '   TableName - table to alter
            '   SqlText - sql text to use for the altering
            '
            ' returns:
            '   >= 0 - no errors
            '   <0 - SqlCommand.ExecuteNonQuery error value
            '   Pcm.sqlNoTableName - no table name 
            '   Pcm.sqlErr- other error in sql 
            '   Pcm.sqlAlterErr  - error in delete sql 

            If TableName = "" Then
                Return Pcm.sqlNoTableName
            End If

            Dim AtReturn As Integer = 0
            Dim ErrMsg As String = AlterTableErrMessage(TableName)
            Try
                ' alter the table via execute non query
                AtReturn = ExecuteNonQuery(SqlText, OleDbConn, ErrMsg, Pcm.etSqlAlterErr)
            Catch ex As Exception
                MessageBox.Show(String.Format("{0}  {1}", ErrMsg, ex.Message),
                            Pcm.etSqlAlterErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return Pcm.sqlAlterErr
            End Try
            Return AtReturn
        End Function

        Public Shared Function AlterTableAddKey(ByVal OleDbConn As OleDb.OleDbConnection,
                                                ByVal TableName As String,
                                                ByVal eaColumns() As eaColumnClass,
                                                ByVal KeyName As String) As Integer

            ' adds a multi column key to a table
            ' 
            ' vars passed:
            '   OleDbConn - connection to use
            '   TableName - table name to add primary key to
            '   eaColumns() - array of column names and types
            '   KeyName - name of primary key
            '
            ' returns: 
            '   >= 0 - no errors
            '   <0 - SqlCommand.ExecuteNonQuery error value
            '   Pcm.sqlNoTableName - no table name 
            '   Pcm.sqlNoKeyName - no key name
            '   Pcm.sqlNoColumn - no columns
            '   Pcm.sqlErr- other error in sql 
            '   Pcm.sqlInvalidAlter - invalid alter option (not atAddColumn or atDropColumn)
            '   Pcm.sqlAlterErr  - error in alter sql 

            If TableName = "" Then
                Return Pcm.sqlNoTableName
            End If
            If KeyName = "" Then
                Return Pcm.sqlNoKeyName
            End If

            Dim SqlText As String
            Dim ErrMsg As String = AlterTableErrMessage(TableName)
            Dim aCol As eaColumnClass
            Try
                ' make sure got columns
                If (eaColumns Is Nothing) OrElse (eaColumns.Length = 0) Then
                    Return Pcm.sqlNoColumn  ' return error if no columns
                End If
                ' ALTER TABLE <TableName> ADD CONSTRAINT <KeyName> PRIMARY KEY (<ColumnName1, ColumnName2,..>)
                SqlText = String.Format("ALTER TABLE {0} ADD CONSTRAINT {1} PRIMARY KEY (", TableName, KeyName)
                Dim aColCount As Integer = 0
                For Each aCol In eaColumns
                    SqlText &= aCol.ColumnName
                    If aColCount < eaColumns.Length - 1 Then
                        SqlText &= ","
                    End If
                    aColCount += 1
                Next
                SqlText &= ")"
            Catch ex As Exception
                MessageBox.Show(String.Format("{0}  {1}", ErrMsg, ex.Message),
                            Pcm.etSqlAlterErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return Pcm.sqlAlterErr
            End Try
            Return AlterTable(OleDbConn, TableName, SqlText)
        End Function

        Public Shared Function AlterTableDropKey(ByVal OleDbConn As OleDb.OleDbConnection,
                                                 ByVal TableName As String,
                                                 ByVal KeyName As String) As Integer

            ' adds a multi column key to a table
            ' 
            ' vars passed:
            '   OleDbConn - connection to use
            '   TableName - table to store foreign key rows to delete        
            '   KeyName - name of primary key
            '
            ' returns: 
            '   >= 0 - no errors
            '   <0 - SqlCommand.ExecuteNonQuery error value
            '   Pcm.sqlNoTableName - no table name 
            '   Pcm.sqlNoKeyName - no key name
            '   Pcm.sqlNoColumn - no columns
            '   Pcm.sqlErr- other error in sql 
            '   Pcm.sqlInvalidAlter - invalid alter option (not atAddColumn or atDropColumn)
            '   Pcm.sqlAlterErr  - error in alter sql 

            If TableName = "" Then
                Return Pcm.sqlNoTableName
            End If
            If KeyName = "" Then
                Return Pcm.sqlNoKeyName
            End If

            ' aColumn param not used when dropping a key
            Return AlterTableColumn(OleDbConn, TableName, Nothing, eaColumnClass.AlterTableTypes.atDropKey, KeyName)
        End Function

        Private Shared Function AlterTableErrMessage(TableName As String) As String

            ' gets the initial error message when altering a table
            '
            ' vars passed:
            '   Table Name - name of table being altered
            '
            ' returns:
            '   initial error message

            Return String.Format("Error altering table ""{0}"".", TableName)
        End Function

#End Region

#Region " Create Index "

        Public Shared Function CreateIndex(ByVal OleDbConn As OleDb.OleDbConnection,
                                           ByVal TableName As String,
                                           ByVal IndexName As String,
                                           ByVal eaColumns() As eaColumnClass,
                                           ByVal SortASC() As Boolean,
                                           ByVal Unique As Boolean) As Integer

            ' adds a index to a table
            ' 
            ' vars passed:
            '   OleDbConn - connection to use
            '   TableName - table name for index
            '   IndexName - name of index
            '   eaColumns() - array of column names and types to be used in index
            '   SortASC() - array of sort orders, length must match eaColumns
            '   Unique - TRUE if index requires unique rows
            '
            ' returns: 
            '   >= 0 - no errors
            '   <0 - SqlCommand.ExecuteNonQuery error value
            '   Pcm.sqlNoTableName - no table name 
            '   Pcm.sqlNoIndexName - no index name
            '   Pcm.sqlNoColumn - no columns
            '   Pcm.sqlErr - other error in sql 
            '   Pcm.sqlCreateIndexErr - error in create index sql

            If TableName = "" Then
                Return Pcm.sqlNoTableName
            End If
            If IndexName = "" Then
                Return Pcm.sqlNoIndexName
            End If

            Dim SqlText As String = "CREATE "
            Dim ErrMsg As String = String.Format("Error creating index ""{0}"" for table ""{1}"".", IndexName, TableName)
            Dim aCol As eaColumnClass
            Try
                If (eaColumns Is Nothing) OrElse (eaColumns.Length = 0) Then    ' make sure got columns
                    Return Pcm.sqlNoColumn                                      ' return error if no columns
                End If
                If (SortASC Is Nothing) OrElse (SortASC.Length = 0) Then        ' make sure got sorts
                    Return Pcm.sqlNoColumn                                      ' return error if no columns
                End If
                If eaColumns.Length <> SortASC.Length Then                      ' make sure array sizes match 
                    Return Pcm.sqlNoColumn                                      ' return error if array sizes dont match 
                End If
                ' CREATE [UNIQUE] INDEX <IndexName> ON <TableName> (<ColumnName1> [ASC|DESC], <ColumnName2> [ASC|DESC], ...)
                If Unique Then                                                  ' if unique rows in index
                    SqlText &= "UNIQUE "                                        ' add UNIQUE to sql text
                End If
                SqlText &= String.Format("INDEX {0} ON {1} (", IndexName, TableName) ' add index's name and table name
                Dim aColCount As Integer = 0
                For Each aCol In eaColumns                                      ' for each column in index
                    SqlText &= aCol.ColumnName                                  ' add column name
                    If SortASC(aColCount) Then                                  ' if ascending sort
                        SqlText &= " ASC"                                       ' add sort direction
                    Else                                                        ' else descending sort
                        SqlText &= " DESC"                                      ' add sort direction
                    End If
                    If aColCount < eaColumns.Length - 1 Then                    ' if not the last column
                        SqlText &= ","                                          ' add a space
                    End If
                    aColCount += 1                                              ' increment column counter
                Next
                SqlText &= ")"                                                  ' add ending bracket for column names
            Catch ex As Exception
                MessageBox.Show(String.Format("{0}  {1}", ErrMsg, ex.Message),
                            Pcm.etSqlCreateIndexErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return Pcm.sqlCreateIndexErr
            End Try
            Return PcdbSql.ExecuteNonQuery(SqlText, OleDbConn, ErrMsg, Pcm.etSqlCreateIndexErr)
        End Function

#End Region

#Region " Create Table "

        Public Shared Function CreateTable(ByVal OleDbConn As OleDb.OleDbConnection,
                                           ByVal TableName As String,
                                           ByVal SqlText As String) As Integer

            ' create a new data table
            '
            ' vars passed:
            '   OleDbConn - connection to use
            '   TableName - table to store foreign key rows to delete
            '   SqlText - Sql Text to use
            '
            ' returns:
            '   >= 0 - no errors
            '   <0 - SqlCommand.ExecuteNonQuery error value
            '   Pcm.sqlErr- other error in sql 
            '   Pcm.sqlDropErr - other error in dropping table
            '   Pcm.sqlCreateErr - other error in creating table

            If TableName = "" Then
                Return Pcm.sqlNoTableName
            End If

            Dim CtReturn As Integer = 0
            Dim ErrMsg As String = String.Format("Error creating table ""{0}"".", TableName)
            Try
                CtReturn = DropTable(OleDbConn, TableName)  ' try to drop table (if it exists)
                If CtReturn < 0 Then                        ' if error in drop table
                    Return CtReturn                         ' return drop error
                End If
                ' create the table via execute non query
                CtReturn = ExecuteNonQuery(SqlText, OleDbConn, ErrMsg, Pcm.etSqlCreateErr)
            Catch ex As Exception
                MessageBox.Show(String.Format("{0}  {1}", ErrMsg, ex.Message),
                            Pcm.etSqlCreateErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return Pcm.sqlCreateErr
            End Try
            Return CtReturn
        End Function

        Public Shared Function CreateTable(ByVal OleDbConn As OleDb.OleDbConnection,
                                           ByVal TableName As String,
                                           ByVal eaColumns() As eaColumnClass,
                                           Optional ByVal KeyName As String = "") As Integer

            ' create a new data table
            '
            ' vars passed:
            '   OleDbConn - connection to use
            '   TableName - table to store foreign key rows to delete
            '   Columns() - array of column names and types
            '   KeyName - name of primary key
            '
            ' returns:
            '   >= 0 - no errors
            '   <0 - SqlCommand.ExecuteNonQuery error value
            '   Pcm.sqlNoTableName - no table name 
            '   Pcm.sqlNoColumn - no columns
            '   Pcm.sqlErr - other error in sql 
            '   Pcm.sqlDropErr - other error in dropping table
            '   Pcm.sqlCreateErr - other error in creating table
            '
            '   CREATE TABLE <TableName> (
            '       <Field1Name> <Field1Type>, <Field2Name> <Field2Type>, ..
            '       [CONSTRAINT <KeyName> 
            '           [PRIMARY KEY (<Field1Name>, <Field2Name>,..)] 
            '       ]
            '   )
            '   if FieldType is Text, then use Text(<Size>) for FieldType

            If TableName = "" Then
                Return Pcm.sqlNoTableName
            End If

            Dim SqlText As String
            Dim ErrMsg As String = String.Format("Error creating table ""{0}"".", TableName)
            Try
                ' make sure got columns
                If (eaColumns Is Nothing) OrElse (eaColumns.Length = 0) Then
                    Return Pcm.sqlNoColumn  ' return error if no columns
                End If

                ' create table + table name + start bracket for fields
                SqlText = String.Format("CREATE TABLE {0} (", TableName)
                Dim ColText As String
                Dim aColCount As Integer = 0
                Dim aCol As eaColumnClass
                Dim KeyCount As Integer = 0

                ' count # of key columns
                For Each aCol In eaColumns              ' for each column
                    If aCol.KeyField Then               ' if a key column   
                        KeyCount += 1                   ' increment key col count
                    End If
                Next

                If KeyCount > 0 AndAlso KeyName = "" Then
                    Return Pcm.sqlNoKeyName
                End If
                ' get field names string for SQL
                For Each aCol In eaColumns              ' for each column
                    ' no space for first aCol
                    If aColCount = 0 Then               ' if first column
                        ColText = ""                    ' no need for a space before column name
                    Else                                ' else not first column
                        ColText = " "                   ' need space before column name
                    End If

                    ' get column name and type "Name Type"
                    If IsKeyWord(aCol.ColumnName) Then                                              ' if column name is a Key word
                        ColText &= String.Format("[{0}] {1}", aCol.ColumnName, aCol.SqlColumnType)  ' place column name in square brackets []
                    Else                                                                            ' else column name is not SQL key word
                        ColText &= String.Format("{0} {1}", aCol.ColumnName, aCol.SqlColumnType)    ' no square brackets needed
                    End If

                    If aCol.ColumnType = eaColumnClass.ColumnTypes.ctText Then                  ' if a text column
                        ColText &= String.Format("({0})", CStr(aCol.TextLength))                ' add in text size "(XX)"

                        'ElseIf aCol.ColumnType = eaColumnClass.ColumnTypes.ctBoolean Then
                        '    ColText &= " YESNO"

                    End If
                    If aCol.DefaultValue <> "" Then                                             ' if have a default value
                        ColText &= " DEFAULT " & aCol.DefaultValue                              ' add default and default value
                    End If
                    If Not aCol.IsNullable Then                                                 ' if column is not nullable (required)
                        ColText &= " NOT NULL"                                                  ' add NOT NULL 
                    End If
                    ' get ending comma (if needed) 
                    If aColCount < eaColumns.Length - 1 Then                                    ' if not last column
                        ColText &= ","                                                          ' add comma after field type
                    End If
                    SqlText &= ColText                                                          ' add column text to sql text
                    aColCount += 1                                                              ' increment column count
                Next
                If KeyCount > 0 Then                                                            ' if got a key field(s)
                    ' add ", CONSTRAINT " and KeyName and " PRIMARY KEY " and start bracket for primary key
                    SqlText &= String.Format(", CONSTRAINT {0} PRIMARY KEY (", KeyName)
                    ColText = ""
                    aColCount = 0                                                               ' reset column count
                    For Each aCol In eaColumns                                                  ' for each column
                        If aCol.KeyField Then                                                   ' if a key field 
                            ColText &= aCol.ColumnName                                          ' add column name
                            If aColCount < KeyCount - 1 Then                                    ' if not last aCol
                                ColText &= ", "                                                 ' add a comma after column info
                            End If
                            aColCount += 1                                                      ' increment column count
                        End If
                        If aColCount >= KeyCount Then                                           ' if got all key columns
                            Exit For                                                            ' exit the for loop
                        End If
                    Next
                    SqlText &= ColText & ")"            ' add column names and ending brackey for primary key
                End If
                SqlText &= ")"                          ' add ending bracket for fields
            Catch ex As Exception
                MessageBox.Show(String.Format("{0}  {1}", ErrMsg, ex.Message),
                            Pcm.etSqlCreateErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return Pcm.sqlErr
            End Try
            Return CreateTable(OleDbConn, TableName, SqlText)
        End Function

        Public Shared Function IsKeyWord(text As String) As Boolean

            ' checks if text is a key word (see KeyWords array for full list of words)
            '
            ' vars passed:
            '   text - text to check if a key word
            '   
            ' returns:
            '   TRUE - text is a key word
            '   FALSE - text is not a key word

            Try
                Return KeyWords.Contains(text.ToUpper)
            Catch ex As Exception
                Return False
            End Try
        End Function

#End Region

#Region " Delete Foreign Key Rows/Empty Table "

        Public Shared Function DeleteForeignKeyRows(ByVal fkToDelTableName As String,
                                                    ByVal fkLinkFieldName As String,
                                                    ByVal fkLinkId As Object,
                                                    ByVal OleDbConn As OleDb.OleDbConnection) As Integer

            ' deletes foreign key rows in different, non linked table
            ' note: this is a manual deletion of foreign key rows.
            '
            ' vars passed:
            '   fkToDelTableName - table to store foreign key rows to delete
            '   fkLinkFieldName - name of field that links foreign key table
            '   fkLinkId - value of linking field in in the foreign key table 
            '   OleDbConn - connection to use
            '
            ' returns:
            '   >= 0 - no errors
            '   Pcm.sqlDeleteErr  - error in delete sql 
            '   Pcm.sqlErr - other error in SQL

            Dim fkDelCmd As OleDb.OleDbCommand = Nothing
            Dim fkLinkIdStr As String = ""
            Dim fkReturn As Integer = 0
            Try
                ' get fk link id as a string
                Try
                    fkLinkIdStr = fkLinkId.ToString
                Catch ex1 As Exception
                    MessageBox.Show(String.Format("Error deleting row from ""{0}"".  Could not convert fkLinkId to string.  {1}{2}{3}", fkToDelTableName, ex1.Message, vbCrLf, ex1),
                    Pcm.etDeleteRow, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return Pcm.sqlDeleteErr
                End Try

                ' create the SQL delete command text
                Dim SqlDelText As String = String.Format("DELETE FROM {0} WHERE {1}", fkToDelTableName, EaSql.Sql.SqlWhereColumnValueString("", fkLinkFieldName, fkLinkId))
                '" WHERE (" & fkLinkFieldName & " = " & fkLinkIdStr & ")"
                ' create the SQL command
                fkDelCmd = New OleDb.OleDbCommand(SqlDelText, OleDbConn)
                fkDelCmd.Connection.Open()              ' open the connection
                fkReturn = fkDelCmd.ExecuteNonQuery()   ' delete the foreign key rows
                If fkReturn < 0 Then                    ' if ExecuteNonQuery
                    fkReturn = Pcm.sqlDeleteErr         ' set delete sql error
                End If
            Catch ex As Exception
                MessageBox.Show(String.Format("Error deleting row from ""{0}"" with value ""{1}"" in column ""{2}"".  {3}", fkToDelTableName, fkLinkIdStr, fkLinkFieldName, ex.Message),
                            Pcm.etDeleteRow, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return Pcm.sqlErr
            Finally
                Try
                    If fkDelCmd.Connection.State <> ConnectionState.Closed Then
                        fkDelCmd.Connection.Close()             ' close the connection
                    End If
                Catch ex As Exception
                    MessageBox.Show(String.Format("Error deleting row from '{0}'.  Could not save deletion.  {1}", fkToDelTableName, ex.Message),
                                Pcm.etDeleteRow, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    fkReturn = Pcm.sqlErr
                End Try
            End Try
            Return fkReturn
        End Function

        Public Shared Function EmptyTable(ByVal TableName As String,
                                          ByVal OleDbConn As OleDb.OleDbConnection) As Integer

            ' clears data from a table
            ' 
            ' vars passed:
            '   TableName - name of table with data to clear
            '   OleDbConn - connection to use
            ' 
            ' returns:
            '   Pcm.NoErrors 
            '   Pcm.teEmptyTableErr - error clearing table

            Try
                Using EmptyTblDelCmd As New OleDb.OleDbCommand("DELETE * FROM " & TableName, OleDbConn) ' get delete all rows all columns command
                    Try
                        EmptyTblDelCmd.Connection.Open()                                    ' open connection
                        EmptyTblDelCmd.ExecuteNonQuery()                                    ' empty table
                    Catch ex As Exception
                        MessageBox.Show(String.Format("Error emptying table ""{0}"".  {1}", TableName, ex.Message),
                                    "Empty Table Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return Pcm.teEmptyTableErr
                    Finally
                        If EmptyTblDelCmd.Connection.State = ConnectionState.Open Then      ' if connection still open 
                            EmptyTblDelCmd.Connection.Close()                               ' close it
                        End If
                    End Try
                End Using
                Return Pcm.NoErrors         ' if got here, then no errors
            Catch ex As Exception
                MessageBox.Show(String.Format("Other error emptying table ""{0}"".  {1}", TableName, ex.Message),
                            "Empty Table Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return Pcm.teEmptyTableErr
            End Try
        End Function

#End Region

#Region " Drop Index "

        Public Shared Function DropIndex(ByVal OleDbConn As OleDb.OleDbConnection,
                                         ByVal TableName As String,
                                         ByVal IndexName As String) As Integer

            ' drops (deletes) a index on a table
            '
            ' vars passed:
            '   OleDbConn - connection to use
            '   TableName - table with index
            '   IndexName - name of index to drop (delete)
            '
            ' returns:
            '   >= 0 - no errors 
            '   <0 - SqlCommand.ExecuteNonQuery error value
            '   Pcm.sqlNoTableName - no table name
            '   Pcm.sqlNoIndexName - no index name
            '   Pcm.sqlErr - other error in sql 
            '   Pcm.sqlDropErr - other error in dropping index

            Const TableDoesNotExistErrCode As Integer = -2147217865

            If TableName = "" Then
                Return Pcm.sqlNoTableName
            End If
            If IndexName = "" Then
                Return Pcm.sqlNoIndexName
            End If

            Dim DtReturn As Integer = 0
            Dim ErrMsg As String = String.Format("Error dropping index ""{0}"" for table ""{1}"".", IndexName, TableName)
            Try
                ' DROP INDEX <IndexName> on <TableName>
                Dim SqlText As String = String.Format("DROP INDEX {0} ON {1}", IndexName, TableName)  ' drop index sql text

                ' drop the index via execute non query
                DtReturn = ExecuteNonQuery(SqlText, OleDbConn, ErrMsg,
                Pcm.etSqlDropIndexErr, TableDoesNotExistErrCode, Pcm.NoErrors)
            Catch ex As Exception
                MessageBox.Show(String.Format("{0}  {1}", ErrMsg, ex.Message),
                            Pcm.etSqlDropIndexErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return Pcm.sqlDropErr
            End Try
            Return DtReturn
        End Function

#End Region

#Region " Drop Table "

        Public Shared Function DropTable(ByVal OleDbConn As OleDb.OleDbConnection,
                                         ByVal TableName As String) As Integer

            ' drops (deletes) a data table
            '
            ' vars passed:
            '   TableName - table to delete
            '   OleDbConn - connection to use
            '
            ' returns:
            '   >= 0 - no errors or Table did not exist
            '   <0 - SqlCommand.ExecuteNonQuery error value
            '   Pcm.sqlNoTableName - no table name
            '   Pcm.sqlErr - other error in sql 
            '   Pcm.sqlDropErr - other error in dropping table

            Const TableDoesNotExistErrCode As Integer = -2147217865

            If TableName = "" Then
                Return Pcm.sqlNoTableName
            End If

            Dim ErrMsg As String = String.Format("Error dropping table ""{0}"".", TableName)
            Try
                If Not TableExists(TableName, OleDbConn) Then               ' if table does not exist, cannot delete
                    Return Pcm.NoErrors                                     ' return no errors
                End If

                ' DROP TABLE <TableName> 
                Dim SqlText As String = "DROP TABLE " & TableName           ' set sql text to drop table

                ' drop the table via execute non query
                Return ExecuteNonQuery(SqlText, OleDbConn, ErrMsg,
                                   Pcm.etSqlDropErr, TableDoesNotExistErrCode, Pcm.NoErrors)
            Catch ex As Exception
                MessageBox.Show(String.Format("{0}  {1}", ErrMsg, ex.Message),
                            Pcm.etSqlDropErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return Pcm.sqlDropErr
            End Try
        End Function

#End Region

#Region " Rename Table "

        Public Shared Function RenameTable(ByVal OleDbConn As OleDb.OleDbConnection,
                                           ByVal FromTableName As String,
                                           ByVal ToTableName As String) As Integer

            ' renames a table
            '   
            ' note: any PrimaryKey or indexes will be lost
            ' 
            ' vars passed:
            '   OleDbConn - connection to use
            '   FromTableName - old table name
            '   ToTableName - new table name
            '
            ' returns:
            '   >= 0 - no errors
            '   <0 - SqlCommand.ExecuteNonQuery error value
            '   Pcm.sqlNoTableName - from or to table name not set
            '   Pcm.sqlErr- other error in sql 
            '   Pcm.sqlMakeTableErr - error in make table sql 
            '   Pcm.sqlDropErr - other error in dropping table
            '   Pcm.sqlRenameTableErr - other error in rename
            '
            ' 1) delete ToTable if needed 
            ' 2) copy all of FromTable to ToTable
            ' 3) delete FromTable

            If FromTableName = "" OrElse ToTableName = "" Then
                Return Pcm.sqlNoTableName
            End If

            Try
                ' 1) delete ToTable if needed
                Dim RenameErr As Integer = Pcm.NoErrors
                If TableExists(ToTableName, OleDbConn) Then
                    RenameErr = DropTable(OleDbConn, ToTableName)
                End If

                ' 2) copy all of FromTable to ToTable
                Dim SqlText As String
                Dim ErrMsg As String = String.Format("Error copying table ""{0}"" to ""{1}"".", FromTableName, ToTableName)
                Try
                    ' SELECT * INTO <ToTableName> FROM <FromTableName>
                    SqlText = String.Format("SELECT * INTO {0} FROM {1}", ToTableName, FromTableName)
                Catch ex As Exception
                    MessageBox.Show(String.Format("{0}  {1}", ErrMsg, ex.Message),
                                Pcm.etSqlRenameTableErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return Pcm.sqlMakeTableErr
                End Try
                RenameErr = ExecuteNonQuery(SqlText, OleDbConn, ErrMsg, Pcm.etSqlAlterErr)
                If RenameErr < 0 Then
                    Return RenameErr
                End If

                ' 3) delete FromTable        
                If TableExists(FromTableName, OleDbConn) Then
                    RenameErr = DropTable(OleDbConn, FromTableName)
                End If

                Return Pcm.NoErrors
            Catch ex As Exception
                MessageBox.Show(String.Format("Other error renaming table ""{0}"" to ""{1}"".  {2}", FromTableName, ToTableName, ex.Message),
                            Pcm.etSqlRenameTableErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return sqlRenameTableErr
            End Try
        End Function

#End Region

#Region " Row Count "

        Public Shared Function RowCount(ByVal OleDbConn As OleDb.OleDbConnection,
                                        ByVal TableName As String) As Integer

            ' gets the # or rows in a table
            '
            ' vars passed:
            '   OleDbConn - connection to use
            '   TableName - name of the table
            '
            ' returns 
            '   >=0 - number of rows in table
            '   Pcm.teNoTableErr - table not found
            '   Pcm.sqlErr - other error in sql 
            '   Pcm.teOtherErr - something went wrong
            '   <0 - SqlCommand.ExecuteNonQuery error value

            Try
                If Not TableExists(TableName, OleDbConn) Then
                    Return Pcm.teNoTableErr
                End If
                Dim SqlText As String = String.Format("SELECT COUNT(*) FROM {0}", TableName)
                Dim ErrMsg As String = String.Format("Error getting row count from table: {0}", TableName)
                Return ExecuteScalar(SqlText, OleDbConn, ErrMsg, "Error Getting Row Count")
            Catch ex As Exception
                Return Pcm.teOtherErr
            End Try
        End Function

#End Region

#Region " Table Exists "

        Public Shared Function TableExists(ByVal TableName As String,
                                           ByVal OleDbConn As OleDb.OleDbConnection) As Boolean

            ' checks to see if a table exits in a connection
            '
            ' vars passed:
            '   TableName - table name to find
            '   OleDbConn - ole db connection to use
            '
            ' returns:
            '   TRUE - table is found
            '   FALSE - 
            '     Table was not found
            '     Error in find

            Try
                If OleDbConn Is Nothing Then        ' if no connection
                    Return False                    ' return false
                End If
                Dim TryCount As Integer = 0
                Do
                    Try
                        OleDbConn.Open()                ' open the connection
                        ' get the table from the connection
                        Dim aTable As DataTable = OleDbConn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Columns,
                            New Object() {Nothing, Nothing, TableName})
                        If aTable.Rows.Count <> 0 Then  ' if got rows
                            Return True                 ' then got table
                        Else                            ' else no rows
                            Return False                ' so no table
                        End If
                    Catch OleDbEx As OleDb.OleDbException
                        If (OleDbEx.ErrorCode = Pcm.oleCannotOpen) AndAlso (TryCount < Pcm.oleMaxTries) Then
                            TryCount += 1
                            'My.Application.DoEvents()

                            Debug.WriteLine("No DoEvents")

                            System.Threading.Thread.Sleep(Pcm.oleWaitMiliSeconds)
                        Else
                            MessageBox.Show(String.Format("Error trying to find if table ""{0}"" exists.{1}Database: {2}{1}Error Code: {3}{1}Error Message: {4}",
                                                      TableName, vbCrLf, OleDbConn.DataSource, OleDbEx.ErrorCode, OleDbEx.Message),
                                        Pcm.etInternalErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End If
                    Catch ex As Exception
                        MessageBox.Show(String.Format("Error trying to find if table ""{0}"" exists.{1}Database: {2}{1}{3}",
                                                  TableName, vbCrLf, OleDbConn.DataSource, ex.Message),
                                    Pcm.etInternalErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return False
                    Finally
                        OleDbConn.Close()        ' close the connection
                    End Try
                Loop
            Catch ex As Exception
                MessageBox.Show(String.Format("Unknown error trying to find if table ""{0}"" exists.  {1}", TableName, ex.Message),
                            Pcm.etInternalErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End Try
        End Function

#End Region

#End Region

#Region " Column Actions "

#Region " Add Column "

        Public Shared Function AddAutoNumberColumn(ByVal OleDbConn As OleDb.OleDbConnection,
                                                   ByVal TableName As String,
                                                   ByVal aColumn As eaColumnClass,
                                                   Optional ByVal AutoNumStart As Integer = 1,
                                                   Optional ByVal AutoNumIncrement As Integer = 1) As Integer

            ' adds a AutoNumber column to a table.  AutoNumber column is not key column
            '
            ' vars passed:
            '   OldDbConn - connection to use
            '   TableName - table to alter
            '   aColumn - column to alter
            '   AutoNumStart - starting auto number value
            '   AutoNumIncrement - auto number increment value
            '
            ' returns:
            '   >= 0 - no errors
            '   <0 - SqlCommand.ExecuteNonQuery error value
            '   Pcm.sqlNoTableName - no table name 
            '   Pcm.sqlNoKeyName - no key name
            '   Pcm.sqlNoColumn - no columns
            '   Pcm.sqlErr- other error in sql 
            '   Pcm.sqlInvalidAlter - invalid alter option (not atAddColumn or atDropColumn)
            '   Pcm.sqlAlterErr  - error in alter sql 

            If TableName = "" Then
                Return Pcm.sqlNoTableName
            End If

            Dim SqlText As String
            Dim ErrMsg As String = AlterTableErrMessage(TableName)
            Try
                ' ALTER TABLE <TableName> ADD COLUMN <ColumnName> COUNTER(<Start>,<Increment>)
                SqlText = String.Format("ALTER TABLE {0} ADD COLUMN {1} COUNTER({2},{3})", TableName, aColumn.ColumnName, AutoNumStart, AutoNumIncrement)
            Catch ex As Exception
                MessageBox.Show(String.Format("{0}  {1}", ErrMsg, ex.Message),
                            Pcm.etSqlAlterErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return Pcm.sqlAlterErr
            End Try
            Return AlterTable(OleDbConn, TableName, SqlText)
        End Function

        Public Shared Function AddColumn(ByVal OleDbConn As OleDb.OleDbConnection,
                                         ByVal TableName As String,
                                         ByVal ColToAdd As eaColumnClass) As Integer

            ' adds a column to a table
            '
            ' vars passed:
            '   OleDbConn - connection to use
            '   aTableName - name of table to add column to
            '   ColToAdd - eaColumn type to add
            '
            ' returns:
            '   >= 0 - no errors
            '   Pcm.sqlNoTableName - no table name err
            '   Pcm.sqlNoKeyName - no key name err
            '   Pcm.sqlNoColumn - no column name err
            '   Pcm.sqlErr - other sql error
            '   Pcm.sqlInvalidAlter - invalid ALTER option err
            '   Pcm.sqlAlterErr - sql ALTER err
            '   Pcm.OtherExErr - other error
            '   <0 - Invalid SQL command err

            Const etSqlAddColErr As String = "Add Column Error"
            Try
                ' make new id column
                Dim AcErr As Integer = AlterTableColumn(OleDbConn, TableName, ColToAdd, eaColumnClass.AlterTableTypes.atAddColumn)
                Select Case AcErr
                    Case Is >= 0
                    ' do nothing 
                    Case Pcm.sqlNoTableName
                        MessageBox.Show(String.Format("No table name in AddColumn, aTableName = ""{0}"".", TableName),
                                    etSqlAddColErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Case Pcm.sqlNoKeyName
                        MessageBox.Show(String.Format("No key name in AddColumn, aColumnName = ""{0}"".", ColToAdd.ColumnName),
                                    etSqlAddColErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Case Pcm.sqlNoColumn
                        MessageBox.Show(String.Format("No column name in AddColumn, aColumnName = ""{0}"".", ColToAdd.ColumnName),
                                    etSqlAddColErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Case Pcm.sqlErr
                        MessageBox.Show(String.Format("Other SQL error in AddColumn, aTableName = ""{0}"", aColumnName = ""{1}"".", TableName, ColToAdd.ColumnName),
                                    etSqlAddColErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Case Pcm.sqlInvalidAlter
                        MessageBox.Show(String.Format("Invalid ALTER option in AddColumn, aTableName = ""{0}"", aColumnName = ""{1}"".", TableName, ColToAdd.ColumnName),
                                    etSqlAddColErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Case Pcm.sqlAlterErr
                        MessageBox.Show(String.Format("ALTER error in AddColumn, aTableName = ""{0}"", aColumnName = ""{1}"".", TableName, ColToAdd.ColumnName),
                                    etSqlAddColErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Case Else
                        MessageBox.Show(String.Format("Invalid SQL in AddColumn, aTableName = ""{0}"", aColumnName = ""{1}"".", TableName, ColToAdd.ColumnName),
                                    etSqlAddColErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Select
                Return AcErr
            Catch ex As Exception
                MessageBox.Show("Other error in AddColumn.  " & ex.Message,
                            etSqlAddColErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return Pcm.OtherExErr
            End Try
        End Function

#End Region

#Region " Alter Table Column "

        Public Shared Function AlterTableColumn(ByVal OleDbConn As OleDb.OleDbConnection,
                                                ByVal TableName As String,
                                                ByVal aColumn As eaColumnClass,
                                                ByVal AlterType As eaColumnClass.AlterTableTypes,
                                                Optional ByVal KeyName As String = "") As Integer

            ' alters a data table, 
            '   adding or dropping a column
            '   adding key (one field), not multi field
            '   dropping key
            '
            ' vars passed:
            '   OldDbConn - connection to use
            '   TableName - table to alter
            '   aColumn - column to alter
            '   AlterType - 
            '     atAddColumn (add column)
            '     atDropColumn (drop column)
            '     atAddKey (add key) 
            '     atDropKey (drop key)
            '   KeyName - name of primary key (required if AlterType is atAddKey or atDropKey)
            '
            ' returns:
            '   >= 0 - no errors
            '   <0 - SqlCommand.ExecuteNonQuery error value
            '   Pcm.sqlNoTableName - no table name 
            '   Pcm.sqlNoKeyName - no key name
            '   Pcm.sqlNoColumn - no columns
            '   Pcm.sqlErr- other error in sql 
            '   Pcm.sqlInvalidAlter - invalid alter option (not atAddColumn or atDropColumn)
            '   Pcm.sqlAlterErr  - error in alter sql 

            If TableName = "" Then
                Return Pcm.sqlNoTableName
            End If
            ' if got a alter key, must have a key name
            If (AlterType = eaColumnClass.AlterTableTypes.atAddKey) OrElse (AlterType = eaColumnClass.AlterTableTypes.atDropKey) Then
                If KeyName = "" Then
                    Return Pcm.sqlNoKeyName
                End If
                If KeyName.IndexOfAny(CType(" #", Char())) >= 0 Then    ' if key name has space or special chars
                    KeyName = String.Format("[{0}]", KeyName)           ' put square brackets around KeyName
                End If
            End If

            Dim SqlText As String
            Dim ErrMsg As String = AlterTableErrMessage(TableName)
            Try
                ' if no column to alter and not dropping key, must have aColumn
                If (aColumn Is Nothing) AndAlso (AlterType <> eaColumnClass.AlterTableTypes.atDropKey) Then
                    Return Pcm.sqlNoColumn          ' then return error
                End If

                Select Case AlterType
                    Case eaColumnClass.AlterTableTypes.atAddColumn
                        ' ALTER TABLE <TableName> ADD COLUMN <ColumnName> <ColumnType>
                        ' ALTER TABLE <TableName> ADD COLUMN <ColumnName> Text(<Size>)                    
                        SqlText = String.Format("ALTER TABLE {0} ADD COLUMN {1} {2}", TableName, aColumn.ColumnName, aColumn.SqlColumnType)
                        If aColumn.ColumnType = eaColumnClass.ColumnTypes.ctText Then   ' if a text column
                            SqlText &= String.Format("({0})", CStr(aColumn.TextLength)) ' need to add "(size)"
                        End If
                    Case eaColumnClass.AlterTableTypes.atDropColumn
                        ' ALTER TABLE <TableName> DROP COLUMN <ColumnName>
                        SqlText = String.Format("ALTER TABLE {0} DROP COLUMN {1}", TableName, aColumn.ColumnName)
                    Case eaColumnClass.AlterTableTypes.atAddKey
                        ' ALTER TABLE <TableName> ADD CONSTRAINT <KeyName> PRIMARY KEY (<ColumnName>)
                        SqlText = String.Format("ALTER TABLE {0} ADD CONSTRAINT {1} PRIMARY KEY ({2})", TableName, KeyName, aColumn.ColumnName)
                    Case eaColumnClass.AlterTableTypes.atDropKey
                        ' ALTER TABLE <TableName> DROP CONSTRAINT <KeyName> 
                        SqlText = String.Format("ALTER TABLE {0} DROP CONSTRAINT {1}", TableName, KeyName)
                    Case Else
                        Return Pcm.sqlInvalidAlter
                End Select
            Catch ex As Exception
                MessageBox.Show(String.Format("{0}  {1}", ErrMsg, ex.Message),
                            Pcm.etSqlAlterErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return Pcm.sqlAlterErr
            End Try
            Return AlterTable(OleDbConn, TableName, SqlText)
        End Function

#End Region

#Region " Alter Column Type "

        Public Shared Function AlterColumnType(ByVal OleDbConn As OleDb.OleDbConnection,
                                               ByVal TableName As String,
                                               ByVal FromColumn As eaColumnClass,
                                               ByVal ToColumn As eaColumnClass) As Integer

            ' alters a column, changing the data type
            ' 
            ' note: this function DOES NOT change the column name, just the data type
            '
            ' vars passed:
            '   OldDbConn - connection to use
            '   TableName - table to alter
            '   FromColumn - column to alter
            '   ToColumn - column to change to
            '
            ' returns:
            '   >= 0 - no errors
            '   <0 - SqlCommand.ExecuteNonQuery error value
            '   Pcm.sqlErr- other error in sql 
            '   Pcm.sqlInvalidAlter - invalid alter option (not atAddColumn or atDropColumn)
            '   Pcm.sqlAlterErr  - error in alter sql 

            If TableName = "" Then
                Return Pcm.sqlNoTableName
            End If

            Dim SqlText As String
            Dim ErrMsg As String = AlterTableErrMessage(TableName)
            Try
                If (FromColumn Is Nothing) OrElse (ToColumn Is Nothing) Then ' if no from or to column 
                    Return Pcm.sqlNoColumn                                  ' then return error
                End If
                ' ALTER TABLE <TableName> ALTER COLUMN <ColumnName> <ColumnType>
                ' ALTER TABLE <TableName> ALTER COLUMN <ColumnName> Text(<Size>)
                SqlText = String.Format("ALTER TABLE {0} ALTER COLUMN {1} {2}", TableName, FromColumn.ColumnName, ToColumn.SqlColumnType)
                If ToColumn.ColumnType = eaColumnClass.ColumnTypes.ctText Then
                    SqlText += String.Format("({0})", CStr(ToColumn.TextLength))
                End If
            Catch ex As Exception
                MessageBox.Show(String.Format("{0}  {1}", ErrMsg, ex.Message),
                            Pcm.etSqlAlterErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return Pcm.sqlAlterErr
            End Try
            Return AlterTable(OleDbConn, TableName, SqlText)
        End Function

#End Region

#Region " Column Exists "

        Public Shared Function ColumnExists(ByVal aTable As DataTable, ByVal ColumnName As String) As Boolean

            ' sees if a column exists in a table
            '
            ' vars passed:
            '   aTable - data table to check
            '   ColumnName - name of column to look for
            '
            ' returns:
            '   TRUE - table contains the column 
            '   FALSE - 
            '     aTable is nothing
            '     aTable does not contain the column
            '     something went wrong

            Try
                If aTable Is Nothing Then
                    Return False
                End If
                Return aTable.Columns.Contains(ColumnName)
            Catch ex As Exception
                Return False
            End Try
        End Function

        Public Shared Function ColumnExists(ByVal TableName As String,
                                            ByVal ColumnName As String,
                                            ByVal OleDbConn As OleDb.OleDbConnection) As Boolean

            ' sees if a column exists in a table.  Use this function if the table definition is not in the dataset for
            ' the table.  this version tries to perform a SELECT SQL query for the table and column.  The query will 
            ' cause an OleDb.OleDbException exception if the table or column does not exist.  This exception is handled 
            ' here, and FALSE is returned.  If there is no exception thrown, then both the table and column exist, and 
            ' TRUE is returned.  
            '
            ' vars passed:
            '   TableName - name of data table to check
            '   ColumnName - name of column to look for
            '   aOleDbConnection - ole db connection to use
            '   ErrMsg - error message to display
            '   ErrTitle - error title for error in executing non query
            '
            ' returns:
            '   TRUE - TableName contains the column 
            '   FALSE - 
            '     table for TableName not found 
            '     table for TableName does not contain the column
            '     something went wrong

            Dim SqlCommand As OleDb.OleDbCommand = Nothing  ' sql command
            Dim aReader As OleDb.OleDbDataReader = Nothing
            Dim colErr As Integer
            Dim Done As Boolean = False
            Dim TryCount As Integer = 0
            Do
                Try
                    ' only select 1 row
                    Dim SqlText As String = String.Format("SELECT TOP 1 ({0}.{1}) FROM {0}", TableName, ColumnName)
                    SqlCommand = New OleDb.OleDbCommand(SqlText, OleDbConn)             ' get the connection
                    SqlCommand.Connection.Open()                                        ' open the connection
                    aReader = SqlCommand.ExecuteReader                                  ' get the reader
                    colErr = Pcm.NoErrors                                               ' got the column
                    Done = True                                                         ' done with loop
                Catch OleDbEx As OleDb.OleDbException
                    If OleDbEx.ErrorCode = Pcm.oleCannotOpen AndAlso TryCount < Pcm.oleMaxTries Then ' if cant open and less than max tries
                        TryCount += 1                                                   ' increment try count
                        'My.Application.DoEvents()                                       ' process all events

                        Debug.WriteLine("No DoEvents")

                        System.Threading.Thread.Sleep(Pcm.oleWaitMiliSeconds)           ' wait some time
                    ElseIf (OleDbEx.ErrorCode = Pcm.oleNoColumn) Then                   ' if no column error
                        colErr = OleDbEx.ErrorCode                                      ' set error code
                        Done = True                                                     ' done with do loop
                    Else                                                                ' else other error or over max tries
                        MessageBox.Show(String.Format("OleDbException in PcdbSql.ColumnExists.  Table: {0}, Column: {1}, Error Code: {2}.  {3}",
                                                  TableName, ColumnName, OleDbEx.ErrorCode, OleDbEx.Message),
                                    Pcm.etSqlColumnExistsErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                        colErr = OleDbEx.ErrorCode                                      ' set error code
                        Done = True                                                     ' done with do loop
                    End If
                Catch ex As Exception
                    MessageBox.Show("Error in ColumnExists.  " & ex.Message,
                                Pcm.etSqlColumnExistsErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    colErr = Pcm.sqlErr
                Finally
                    Try
                        If SqlCommand IsNot Nothing Then                                    ' if got a command
                            If SqlCommand.Connection.State <> ConnectionState.Closed Then   ' if connection not closed 
                                SqlCommand.Connection.Close()                               ' close the connection
                            End If
                            SqlCommand.Dispose()                                            ' dispose of the command
                            SqlCommand = Nothing
                        End If
                    Catch ex As Exception
                        MessageBox.Show("Error in ColumnExists closing connection.  " & ex.Message,
                                    Pcm.etSqlCanNotCloseErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                        colErr = Pcm.sqlErr
                    End Try
                End Try
            Loop Until Done
            Return (colErr >= 0)    ' returns true if colErr >= 0; else returns false
        End Function

#End Region

#Region " Rename Column "

        Public Shared Function RenameColumn(ByVal OleDbConn As OleDb.OleDbConnection,
                                            ByVal TableName As String,
                                            ByVal FromColumn As eaColumnClass,
                                            ByVal ToColumn As eaColumnClass) As Integer

            ' renames a column by 
            '   1) adding a new column (ToColumn)
            '   2) copying all data from FromColumn to ToColumn
            '   3) deleting FromColumn
            '
            ' NOTE: do not rename a column in a Primary Key or an Index
            '
            ' vars passed:
            '   OleDbConn - connection to use
            '   TableName - name of table with column to rename
            '   FromColumn - column to be renamed
            '   ToColumn - new column 
            '
            ' returns:
            '   >= 0 - no errors
            '   <0 - SqlCommand.ExecuteNonQuery error value
            '   Pcm.sqlNoTableName - no table name 
            '   Pcm.sqlNoKeyName - no key name
            '   Pcm.sqlNoColumn - no columns
            '   Pcm.sqlErr- other error in sql 
            '   Pcm.sqlInvalidAlter - invalid alter option (not atAddColumn or atDropColumn)
            '   Pcm.sqlAlterErr  - error in alter sql 
            '   sqlRenameTableErr - something went wrong

            If TableName = "" Then
                Return Pcm.sqlNoTableName
            End If
            Try
                ' 1) adding a new column (ToColumn)
                Dim RenameErr As Integer = AlterTableColumn(OleDbConn, TableName, ToColumn, eaColumnClass.AlterTableTypes.atAddColumn)
                If RenameErr < 0 Then
                    Return RenameErr
                End If

                ' 2) copying all data from FromColumn to ToColumn
                ' UPDATE <TableName> SET <ToColumn> = <FromColumn>
                Dim ErrMsg As String = String.Format("Error renaming column ""{0}"" to ""{1}"" in table ""{2}"".", FromColumn.ColumnName, ToColumn.ColumnName, TableName)
                Dim SqlText As String = String.Format("UPDATE {0} SET {1} = {2}", TableName, ToColumn.ColumnName, FromColumn.ColumnName)
                RenameErr = ExecuteNonQuery(SqlText, OleDbConn, ErrMsg, Pcm.etSqlRenameColumnErr)
                If RenameErr < 0 Then
                    Return RenameErr
                End If

                ' 3) deleting FromColumn
                RenameErr = AlterTableColumn(OleDbConn, TableName, FromColumn, eaColumnClass.AlterTableTypes.atDropColumn)
                If RenameErr < 0 Then
                    Return RenameErr
                End If

                Return Pcm.NoErrors
            Catch ex As Exception
                MessageBox.Show(String.Format("Other error renaming column ""{0}"" to ""{1}"" in table ""{2}"".  {3}", FromColumn.ColumnName, ToColumn.ColumnName, TableName, ex.Message),
                            Pcm.etSqlRenameTableErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return sqlRenameTableErr
            End Try
        End Function

#End Region

#End Region

#Region " Query Actions "

#Region " Notes "

        ' in an Access database, Procedures are Views are saved Queries.
        '   Views do not have parameters
        '   Procedures have parameter(s)

#End Region

#Region " CreateQuery "

        Public Shared Function CreateQuery(ByVal OleDbConn As OleDb.OleDbConnection,
                                           ByVal QueryName As String,
                                           ByVal QueryCommand As String) As Integer

            ' creates a query
            '
            ' vars passed:
            '   OleDbConn - connection to use
            '   QueryName - query to create
            '   QueryCommand - query command (SQL Query text)
            '
            ' returns:
            '   >= 0 - no errors
            '   <0 - SqlCommand.ExecuteNonQuery error value
            '   Pcm.sqlErr- other error in sql 
            '   Pcm.sqlDropErr - other error in dropping query
            '   Pcm.sqlCreateErr - other error in creating query

            Try
                If ProcedureExists(OleDbConn, QueryName) Then                       ' if query exists
                    Dim DropErr As Integer = DropQuery(OleDbConn, QueryName)        ' delete the query
                    If DropErr < 0 Then                                             ' if error dropping query
                        Return DropErr                                              ' return the error
                    End If
                End If
                Dim QuerySql As String = String.Format("CREATE PROC {0} AS {1}", QueryName, QueryCommand) ' set the SQL to create the query
                Return PcdbSql.ExecuteNonQuery(QuerySql, OleDbConn, "Error in creating a query", Pcm.etSqlCreateQueryErr) ' create the query
            Catch ex As Exception
                Return Pcm.OtherExErr
            End Try
        End Function

#End Region

#Region " Drop Query "

        Public Shared Function DropQuery(ByVal OleDbConn As OleDb.OleDbConnection,
                                         ByVal QueryName As String) As Integer

            ' drops (deletes) a query (procedure or view)
            '
            ' vars passed:
            '   OleDbConn - connection to use
            '   QueryName - procedure to delete
            '
            ' returns:
            '   >= 0 - no errors or query did not exist
            '   <0 - SqlCommand.ExecuteNonQuery error value
            '   Pcm.sqlNoProcName - no query name
            '   Pcm.sqlErr - other error in sql 
            '   Pcm.sqlDropErr - other error in dropping query

            Const TableDoesNotExistErrCode As Integer = -2147217865

            If QueryName = "" Then
                Return Pcm.sqlNoQueryName
            End If

            Dim ErrMsg As String = String.Format("Error dropping query ""{0}"".", QueryName)
            Try
                If Not QueryExists(OleDbConn, QueryName) Then                       ' if query does not exist, cannot delete
                    Return Pcm.NoErrors                                             ' return no errors
                End If

                ' DROP PROC <ProcName> 
                Dim SqlText As String = "DROP PROC " & QueryName                    ' set sql text to drop procedure

                ' drop the procedure via execute non query
                Return ExecuteNonQuery(SqlText, OleDbConn, ErrMsg,
                                   Pcm.etSqlDropProcErr, TableDoesNotExistErrCode, Pcm.NoErrors)
            Catch ex As Exception
                MessageBox.Show(String.Format("{0}  {1}", ErrMsg, ex.Message),
                            Pcm.etSqlDropProcErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return Pcm.sqlDropErr
            End Try
        End Function

#End Region

#Region " Procedure Exists "

        Public Shared Function ProcedureExists(ByVal OleDbConn As OleDb.OleDbConnection,
                                               ByVal ProcName As String) As Boolean

            ' checks to see if a procedure exits in a connection
            '
            ' vars passed:
            '   OleDbConn - ole db connection to use
            '   ProcName - procedure name to find
            '
            ' returns:
            '   TRUE - procedure is found
            '   FALSE - 
            '     procedure was not found
            '     Error in find

            Try
                If OleDbConn Is Nothing Then    ' if no connection
                    Return False                ' return false
                End If
                Dim Done As Boolean = False
                Dim TryCount As Integer = 0
                Do
                    Try
                        Dim ProcTbl As DataTable = Nothing
                        ' using ... end using makes sure connection is closed and disposed
                        Using connection As New System.Data.OleDb.OleDbConnection(OleDbConn.ConnectionString)
                            connection.Open()                                                                   ' open the connection
                            ProcTbl = connection.GetOleDbSchemaTable(OleDb.OleDbSchemaGuid.Procedures,
                                                               New Object() {Nothing, Nothing, ProcName})   ' get info for one procedure
                        End Using
                        Done = True
                        If ProcTbl IsNot Nothing AndAlso ProcTbl.Rows.Count <> 0 Then   ' if got rows
                            Return True                                                 ' then got procedure
                        Else                                                            ' else no rows
                            Return False                                                ' so no procedure
                        End If
                    Catch OleDbEx As OleDb.OleDbException
                        If (OleDbEx.ErrorCode = Pcm.oleCannotOpen) AndAlso (TryCount < Pcm.oleMaxTries) Then    ' if cant open & not over max tries
                            TryCount += 1                                                                       ' increment try count
                            'My.Application.DoEvents()                                                           ' process all windows messages

                            Debug.WriteLine("No DoEvents")

                            System.Threading.Thread.Sleep(Pcm.oleWaitMiliSeconds)                               ' sleep for a time
                        Else
                            MessageBox.Show(String.Format("Error trying to find if procedure ""{0}"" exists.{1}Database: {2}{1}Error Code: {3}{1}Error Message: {4}",
                                                      ProcName, vbCrLf, OleDbConn.DataSource, OleDbEx.ErrorCode, OleDbEx.Message),
                                        Pcm.etInternalErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End If
                    Catch ex As Exception
                        MessageBox.Show(String.Format("Error trying to find if procedure ""{0}"" exists.  {1}", ProcName, ex.Message),
                                    Pcm.etInternalErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return False
                    Finally
                        OleDbConn.Close()        ' close the connection
                    End Try
                Loop Until Done
            Catch ex As Exception
                MessageBox.Show(String.Format("Unknown error trying to find if procedure ""{0}"" exists.  {1}", ProcName, ex.Message),
                            Pcm.etInternalErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End Try
        End Function

#End Region

#Region " QueryExists "

        Public Shared Function QueryExists(ByVal OleDbConn As OleDb.OleDbConnection,
                                           ByVal QueryName As String) As Boolean

            ' checks to see if a query exits in a connection
            '
            ' vars passed:
            '   OleDbConn - ole db connection to use
            '   QueryName - query name to find
            '
            ' returns:
            '   TRUE - query is found
            '   FALSE - 
            '     query was not found
            '     Error in find

            Try
                If OleDbConn Is Nothing Then        ' if no connection
                    Return False                    ' return false
                End If
                Dim Done As Boolean = False
                Dim TryCount As Integer = 0
                Do
                    Try
                        Dim ProcTbl As DataTable = Nothing
                        Dim ViewTbl As DataTable = Nothing
                        ' using ... end using makes sure connection is closed and disposed
                        Using connection As New System.Data.OleDb.OleDbConnection(OleDbConn.ConnectionString)
                            connection.Open()                                                                   ' open the connection
                            ProcTbl = connection.GetOleDbSchemaTable(OleDb.OleDbSchemaGuid.Procedures,
                                                               New Object() {Nothing, Nothing, QueryName})  ' get info for one procedure
                            If ProcTbl IsNot Nothing AndAlso ProcTbl.Rows.Count <> 0 Then                       ' if got procTbl and rows
                                Done = True
                                Return True                                                                     ' then got procedure (query)
                            Else                                                                                ' else do not have procedure
                                ViewTbl = connection.GetOleDbSchemaTable(OleDb.OleDbSchemaGuid.Views,
                                                               New Object() {Nothing, Nothing, QueryName})  ' get info for one view
                                Done = True
                                If ViewTbl IsNot Nothing AndAlso ViewTbl.Rows.Count <> 0 Then                   ' if got viewTbl and rows
                                    Return True                                                                 ' then got view (query)
                                Else                                                                            ' else no tables and rows
                                    Return False                                                                ' so no query
                                End If
                            End If
                        End Using
                    Catch OleDbEx As OleDb.OleDbException
                        If (OleDbEx.ErrorCode = Pcm.oleCannotOpen) AndAlso (TryCount < Pcm.oleMaxTries) Then    ' if cant open & not over max tries
                            TryCount += 1                                                                       ' increment try count
                            'My.Application.DoEvents()                                                           ' process all windows messages

                            Debug.WriteLine("No DoEvents")

                            System.Threading.Thread.Sleep(Pcm.oleWaitMiliSeconds)                               ' sleep for a time
                        Else
                            MessageBox.Show(String.Format("Error trying to find if query ""{0}"" exists.{1}Database: {2}{1}Error Code: {3}{1}Error Message: {4}",
                                                      QueryName, vbCrLf, OleDbConn.DataSource, OleDbEx.ErrorCode, OleDbEx.Message),
                                        Pcm.etInternalErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End If
                    Catch ex As Exception
                        MessageBox.Show(String.Format("Error trying to find if query ""{0}"" exists.  {1}", QueryName, ex.Message),
                                    Pcm.etInternalErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Done = True
                        Return False
                    Finally
                        OleDbConn.Close()        ' close the connection
                    End Try
                Loop Until Done
            Catch ex As Exception
                MessageBox.Show(String.Format("Unknown error trying to find if query ""{0}"" exists.  {1}", QueryName, ex.Message),
                            Pcm.etInternalErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End Try
        End Function

#End Region

#Region " View Exists "

        Public Shared Function ViewExists(ByVal OleDbConn As OleDb.OleDbConnection,
                                          ByVal ViewName As String) As Boolean

            ' checks to see if a procedure exits in a connection
            '
            ' vars passed:
            '   OleDbConn - ole db connection to use
            '   ViewName - view name to find
            '
            ' returns:
            '   TRUE - view is found
            '   FALSE - 
            '     view was not found
            '     Error in find

            Try
                If OleDbConn Is Nothing Then        ' if no connection
                    Return False                    ' return false
                End If
                Dim Done As Boolean = False
                Dim TryCount As Integer = 0
                Do
                    Try
                        Dim ViewTbl As DataTable = Nothing
                        ' using ... end using makes sure connection is closed and disposed
                        Using connection As New System.Data.OleDb.OleDbConnection(OleDbConn.ConnectionString)
                            connection.Open()                                                                   ' open the connection
                            ViewTbl = connection.GetOleDbSchemaTable(OleDb.OleDbSchemaGuid.Views,
                                                               New Object() {Nothing, Nothing, ViewName})   ' get info for one procedure                    
                            Return True
                        End Using
                        Done = True
                        If ViewTbl IsNot Nothing AndAlso ViewTbl.Rows.Count <> 0 Then   ' if got table and rows
                            Return True                                                 ' then got view
                        Else                                                            ' else no table or rows
                            Return False                                                ' so no view
                        End If
                    Catch OleDbEx As OleDb.OleDbException
                        If (OleDbEx.ErrorCode = Pcm.oleCannotOpen) AndAlso (TryCount < Pcm.oleMaxTries) Then
                            TryCount += 1
                            'My.Application.DoEvents()

                            Debug.WriteLine("No DoEvents")

                            System.Threading.Thread.Sleep(Pcm.oleWaitMiliSeconds)
                        Else
                            MessageBox.Show(String.Format("Error trying to find if view ""{0}"" exists.{1}Database: {2}{1}Error Code: {3}{1}Error Message: {4}",
                                                      ViewName, vbCrLf, OleDbConn.DataSource, OleDbEx.ErrorCode, OleDbEx.Message),
                                        Pcm.etInternalErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End If
                    Catch ex As Exception
                        MessageBox.Show(String.Format("Error trying to find if view ""{0}"" exists.  {1}", ViewName, ex.Message),
                                    Pcm.etInternalErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Done = True
                        Return False
                    Finally
                        OleDbConn.Close()        ' close the connection
                    End Try
                Loop Until Done
            Catch ex As Exception
                MessageBox.Show(String.Format("Unknown error trying to find if view ""{0}"" exists.  {1}", ViewName, ex.Message),
                            Pcm.etInternalErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End Try
        End Function

#End Region

#End Region

#Region " OleDbAdaptToString  "

        'Public Function DaoDbToString(DaoDb As Microsoft.Office.Interop.Access.Dao.Database) As String

        '    Try
        '        Dim aString As New System.Text.StringBuilder
        '        aString.Append(String.Format("Dao.Database{0}{1}{2}", vbTab, DaoDb, vbCrLf))
        '        If DaoDb Is Nothing Then
        '            aString.Append(String.Format("{0}DaoDb = [Nothing]", vbTab))
        '            Return aString.ToString
        '        End If
        '        aString.Append(String.Format("{0}CollatingOrder{0}{1}{2}", vbTab, DaoDb.CollatingOrder, vbCrLf))
        '        'aString.Append(String.Format("{0}Connect{0}{1}{2}", vbTab, DaoDb.Connect, vbCrLf))
        '        'aString.Append(String.Format("{0}Connection{0}{1}{2}", vbTab, DaoDb.Connection, vbCrLf))
        '        'aString.Append(String.Format("{0}{0}Connect{0}{1}{2}", vbTab, DaoDb.Connection.Connect, vbCrLf))
        '        'aString.Append(String.Format("{0}{0}Database{0}{1}{2}", vbTab, DaoDb.Connection.Database, vbCrLf))
        '        'aString.Append(String.Format("{0}{0}Name{0}{1}{2}", vbTab, DaoDb.Connection.Name, vbCrLf))
        '        'aString.Append(String.Format("{0}{0}QueryDefs.Count{0}{1}{2}", vbTab, DaoDb.Connection.QueryDefs.Count, vbCrLf))
        '        'aString.Append(String.Format("{0}{0}QueryTimeout{0}{1}{2}", vbTab, DaoDb.Connection.QueryTimeout, vbCrLf))
        '        'aString.Append(String.Format("{0}{0}RecordsAffected{0}{1}{2}", vbTab, DaoDb.Connection.RecordsAffected, vbCrLf))
        '        'aString.Append(String.Format("{0}{0}Recordsets.Count{0}{1}{2}", vbTab, DaoDb.Connection.Recordsets.Count, vbCrLf))
        '        'aString.Append(String.Format("{0}{0}StillExecuting{0}{1}{2}", vbTab, DaoDb.Connection.StillExecuting, vbCrLf))
        '        'aString.Append(String.Format("{0}{0}Transactions{0}{1}{2}", vbTab, DaoDb.Connection.Transactions, vbCrLf))
        '        'aString.Append(String.Format("{0}{0}Updatable{0}{1}{2}", vbTab, DaoDb.Connection.Updatable, vbCrLf))
        '        'aString.Append(String.Format("{0}Containers.Count{0}{1}{2}", vbTab, DaoDb.Containers.Count, vbCrLf))
        '        'aString.Append(String.Format("{0}DesignMasterID{0}{1}{2}", vbTab, DaoDb.DesignMasterID, vbCrLf))
        '        aString.Append(String.Format("{0}Name{0}{1}{2}", vbTab, DaoDb.Name, vbCrLf))
        '        'aString.Append(String.Format("{0}Properties.Count{0}{1}{2}", vbTab, DaoDb.Properties.Count, vbCrLf))
        '        'aString.Append(String.Format("{0}QueryDefs.Count{0}{1}{2}", vbTab, DaoDb.QueryDefs.Count, vbCrLf))
        '        'aString.Append(String.Format("{0}QueryTimeout{0}{1}{2}", vbTab, DaoDb.QueryTimeout, vbCrLf))
        '        'aString.Append(String.Format("{0}RecordsAffected{0}{1}{2}", vbTab, DaoDb.RecordsAffected, vbCrLf))
        '        'aString.Append(String.Format("{0}Recordsets.Count{0}{1}{2}", vbTab, DaoDb.Recordsets.Count, vbCrLf))
        '        'aString.Append(String.Format("{0}Relations.Count{0}{1}{2}", vbTab, DaoDb.Relations.Count, vbCrLf))
        '        'aString.Append(String.Format("{0}ReplicaID{0}{1}{2}", vbTab, DaoDb.ReplicaID, vbCrLf))
        '        aString.Append(String.Format("{0}TableDefs.Count{0}{1}{2}", vbTab, DaoDb.TableDefs.Count, vbCrLf))
        '        aString.Append(String.Format("{0}Transactions{0}{1}{2}", vbTab, DaoDb.Transactions, vbCrLf))
        '        aString.Append(String.Format("{0}Updatable{0}{1}", vbTab, DaoDb.Updatable))
        '        'aString.Append(String.Format("{0}Version{0}{1}", vbTab, DaoDb.Version))
        '        Return aString.ToString
        '    Catch ex As Exception
        '        Return ""
        '    End Try
        'End Function

        'Public Function OleDbCommandToString(CommandName As String, OleDbCommand As OleDb.OleDbCommand) As String

        '    Try
        '        Dim aString As New System.Text.StringBuilder
        '        aString.Append(String.Format("CommandName{0}{1}{2}", vbTab, CommandName, vbCrLf))
        '        aString.Append(String.Format("{0}CommandText{0}{1}{2}", vbTab, OleDbCommand.CommandText, vbCrLf))
        '        'aString.Append(String.Format("{0}CommandTimeout{0}{1}{2}", vbTab, OleDbCommand.CommandTimeout, vbCrLf))
        '        'aString.Append(String.Format("{0}CommandType{0}{1}{2}", vbTab, OleDbCommand.CommandType, vbCrLf))
        '        aString.Append(OleDbConnToString(OleDbCommand.Connection, True))
        '        'aString.Append(String.Format("{0}Container{0}{1}{2}", vbTab, OleDbCommand.Container, vbCrLf))
        '        'aString.Append(String.Format("{0}Parameters{0}{1}{2}", vbTab, OleDbCommand.Parameters, vbCrLf))
        '        'aString.Append(String.Format("{0}{0}Count{0}{1}{2}", vbTab, OleDbCommand.Parameters.Count, vbCrLf))
        '        'aString.Append(String.Format("{0}Site{0}{1}{2}", vbTab, OleDbCommand.Site, vbCrLf))
        '        'aString.Append(String.Format("{0}Transaction{0}{1}{2}", vbTab, OleDbCommand.Transaction, vbCrLf))
        '        'aString.Append(String.Format("{0}UpdatedRowSource{0}{1}{2}", vbTab, OleDbCommand.UpdatedRowSource, vbCrLf))
        '        Return aString.ToString
        '    Catch ex As Exception
        '        Return ""
        '    End Try
        'End Function

        'Public Function OleDbConnToString(OleDbConn As OleDb.OleDbConnection, Optional UseTwoTabs As Boolean = False) As String

        '    Try
        '        Dim aString As New System.Text.StringBuilder
        '        aString.Append(String.Format("{0}Connection{0}{1}{2}", vbTab, OleDbConn, vbCrLf))
        '        If UseTwoTabs Then
        '            'aString.Append(String.Format("{0}{0}ConnectionString{0}{1}{2}", vbTab, OleDbConn.ConnectionString, vbCrLf))
        '            'aString.Append(String.Format("{0}{0}ConnectionTimeout{0}{1}{2}", vbTab, OleDbConn.ConnectionTimeout, vbCrLf))
        '            'aString.Append(String.Format("{0}{0}Container{0}{1}{2}", vbTab, OleDbConn.Container, vbCrLf))
        '            'aString.Append(String.Format("{0}{0}Database{0}{1}{2}", vbTab, OleDbConn.Database, vbCrLf))
        '            aString.Append(String.Format("{0}{0}DataSource{0}{1}{2}", vbTab, OleDbConn.DataSource, vbCrLf))
        '            'aString.Append(String.Format("{0}{0}Provider{0}{1}{2}", vbTab, OleDbConn.Provider, vbCrLf))
        '            'If OleDbConn.State = ConnectionState.Closed Then
        '            '    aString.Append(String.Format("{0}{0}ServerVersion{0}{1}{2}", vbTab, "Invalid operation.  The connection is closed", vbCrLf))
        '            'Else
        '            '    aString.Append(String.Format("{0}{0}ServerVersion{0}{1}{2}", vbTab, OleDbConn.ServerVersion, vbCrLf))
        '            'End If
        '            'aString.Append(String.Format("{0}{0}Site{0}{1}{2}", vbTab, OleDbConn.Site, vbCrLf))
        '            aString.Append(String.Format("{0}{0}State{0}{1}", vbTab, OleDbConn.State))
        '        Else
        '            'aString.Append(String.Format("{0}ConnectionString{0}{1}{2}", vbTab, OleDbConn.ConnectionString, vbCrLf))
        '            'aString.Append(String.Format("{0}ConnectionTimeout{0}{1}{2}", vbTab, OleDbConn.ConnectionTimeout, vbCrLf))
        '            'aString.Append(String.Format("{0}Container{0}{1}{2}", vbTab, OleDbConn.Container, vbCrLf))
        '            'aString.Append(String.Format("{0}Database{0}{1}{2}", vbTab, OleDbConn.Database, vbCrLf))
        '            aString.Append(String.Format("{0}DataSource{0}{1}{2}", vbTab, OleDbConn.DataSource, vbCrLf))
        '            'aString.Append(String.Format("{0}Provider{0}{1}{2}", vbTab, OleDbConn.Provider, vbCrLf))
        '            'If OleDbConn.State = ConnectionState.Closed Then
        '            '    aString.Append(String.Format("{0}ServerVersion{0}{1}{2}", vbTab, "Invalid operation.  The connection is closed", vbCrLf))
        '            'Else
        '            '    aString.Append(String.Format("{0}ServerVersion{0}{1}{2}", vbTab, OleDbConn.ServerVersion, vbCrLf))
        '            'End If
        '            'aString.Append(String.Format("{0}Site{0}{1}{2}", vbTab, OleDbConn.Site, vbCrLf))
        '            aString.Append(String.Format("{0}State{0}{1}", vbTab, OleDbConn.State))
        '        End If
        '        Return aString.ToString
        '    Catch ex As Exception
        '        Return ""
        '    End Try
        'End Function

#End Region

#Region " Samples "

        'Private Sub MakeTableMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MakeTableMenuItem.Click

        '    Const TestTableName As String = "Test"
        '    Const Field1Name As String = "Id"
        '    Const Field1aName As String = "CheckNo"
        '    Const Field2Name As String = "Name"
        '    Const Field3Name As String = "Amount"
        '    Const Field4Name As String = "OnDate"
        '    Const Field5Name As String = "APercent"
        '    Const Field6Name As String = "AYesOrNo"
        '    Const Field7Name As String = "Comments"

        'Dim CtSql As String = "CREATE TABLE " & TestTableName & " (" & _
        '                Field1Name & " Integer " & "PRIMARY KEY, " & _
        '                Field2Name & " Text(25), " & _
        '                Field3Name & " Currency, " & _
        '                Field4Name & " Date, " & _
        '                Field5Name & " Double, " & _
        '                Field6Name & " Logical, " & _
        '                Field7Name & " Memo" & _
        '                ")"

        'Dim CtSql As String = "CREATE TABLE " & TestTableName & " (" & _
        '                Field1Name & " Integer," & _
        '                Field2Name & " Text(25), " & _
        '                Field3Name & " Currency, " & _
        '                Field4Name & " Date, " & _
        '                Field5Name & " Double, " & _
        '                Field6Name & " Logical, " & _
        '                Field7Name & " Memo, " & _
        '                "CONSTRAINT " & "Pk" & Field1Name & _
        '                    " PRIMARY KEY (" & Field1Name & ")" & _
        '                ")"

        'Dim CtSql As String = "CREATE TABLE " & TestTableName & " (" & _
        '                Field1Name & " Integer, " & _
        '                Field1aName & " Integer, " & _
        '                Field2Name & " Text(25), " & _
        '                Field3Name & " Currency, " & _
        '                Field4Name & " Date, " & _
        '                Field5Name & " Double, " & _
        '                Field6Name & " Logical, " & _
        '                Field7Name & " Memo, " & _
        '                "CONSTRAINT " & Field1Name & Field1aName & _
        '                    " PRIMARY KEY (" & Field1Name & ", " & Field1aName & ")" & _
        '                ")"


        'Dim CtErr As Integer = PcdbSql.DropTable(TestTableName)
        'Dim CtErr As Integer = PcdbSql.CreateTable(TestTableName, CtSql)
        'MessageBox.Show("CtErr = " & CStr(CtErr))

        'Dim myColumns(6) As eaColumnClass
        'Dim MyCol As eaColumnClass
        'myColumns(0) = New eaColumnClass(Field1Name, eaColumnClass.ctInteger, True)
        'myColumns(1) = New eaColumnClass(Field2Name, eaColumnClass.ctText, 25, False)
        'myColumns(2) = New eaColumnClass(Field3Name, eaColumnClass.ctCurrency, False)
        'myColumns(3) = New eaColumnClass(Field4Name, eaColumnClass.ctDate, False)
        'myColumns(4) = New eaColumnClass(Field5Name, eaColumnClass.ctDouble, False)
        'myColumns(5) = New eaColumnClass(Field6Name, eaColumnClass.ctBoolean, False)
        'myColumns(6) = New eaColumnClass(Field7Name, eaColumnClass.ctMemo, False)
        'Dim CtErr As Integer = CreateTable(TestTableName, myColumns, "Pk" & Field1Name)

        'Dim myColumns(7) As eaColumnClass
        'Dim MyCol As eaColumnClass
        'myColumns(0) = New eaColumnClass(Field1Name, eaColumnClass.ctInteger, True)
        'myColumns(1) = New eaColumnClass(Field1aName, eaColumnClass.ctInteger, True)
        'myColumns(2) = New eaColumnClass(Field2Name, eaColumnClass.ctText, 25, False)
        'myColumns(3) = New eaColumnClass(Field3Name, eaColumnClass.ctCurrency, False)
        'myColumns(4) = New eaColumnClass(Field4Name, eaColumnClass.ctDate, False)
        'myColumns(5) = New eaColumnClass(Field5Name, eaColumnClass.ctDouble, False)
        'myColumns(6) = New eaColumnClass(Field6Name, eaColumnClass.ctBoolean, False)
        'myColumns(7) = New eaColumnClass(Field7Name, eaColumnClass.ctMemo, False)
        'Dim CtErr As Integer = CreateTable(TestTableName, myColumns, "Pk" & Field1Name & Field1aName)

        'Dim ColToAdd As eaColumnClass = New eaColumnClass(Field1aName, eaColumnClass.ctInteger, False)
        'Dim ColFrom As eaColumnClass = New eaColumnClass(Field1aName, eaColumnClass.ctInteger, False)
        'Dim ColTo As eaColumnClass = New eaColumnClass(Field1aName, eaColumnClass.ctText, 25, False)
        'Dim CtErr As Integer = PcdbSql.AlterTableColumn(TestTableName, ColToAdd, eaColumnClass.atAddColumn)
        'Dim CtErr As Integer = PcdbSql.AlterTableColumn(TestTableName, ColToAdd, eaColumnClass.atDropColumn)
        'Dim CtErr As Integer = PcdbSql.AlterColumnType(TestTableName, ColFrom, ColTo)
        'Dim CtErr As Integer = PcdbSql.AlterTableColumn(TestTableName, ColTo, eaColumnClass.atDropColumn)

        'Dim KeyCol1 As eaColumnClass = New eaColumnClass(Field1Name, eaColumnClass.ctInteger, True)
        'Dim ColTo As eaColumnClass = New eaColumnClass(Field1aName, eaColumnClass.ctText, 25, False)

        'Dim CtErr As Integer = PcdbSql.AlterTableColumn(TestTableName, KeyCol1, eaColumnClass.atDropKey, "Pk" & Field1Name)
        'Dim CtErr As Integer = PcdbSql.AlterTableColumn(TestTableName, KeyCol1, eaColumnClass.atAddKey, "Pk" & Field1Name)

        'Dim myColumns(1) As eaColumnClass
        'Dim MyCol As eaColumnClass
        'myColumns(0) = New eaColumnClass(Field1Name, eaColumnClass.ctInteger, True)
        'myColumns(1) = New eaColumnClass(Field1aName, eaColumnClass.ctInteger, True)
        'Dim CtErr As Integer = PcdbSql.AlterTableAddKey(TestTableName, myColumns, "Pk" & Field1Name & Field1aName)
        'Dim CtErr As Integer = PcdbSql.AlterTableColumn(TestTableName, Nothing, eaColumnClass.atDropKey, "Pk" & Field1Name & Field1aName)

        'MessageBox.Show("CtErr = " & CStr(CtErr))

        'End Sub

#End Region

    End Class

End Namespace