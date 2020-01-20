Namespace EaSql

    Public Class Sql

#Region " Got SQL Param "

        Public Shared Function GotSqlParam(ByVal id As Integer) As Boolean

            ' checks to see if the value in id is a valid param for SQL select command
            '
            ' vars passed:
            '   id - a id number
            ' 
            ' returns:
            '   TRUE - if id <> Nothing And id <> Pcm.idNotSet
            '   FALSE - if id = Nothing or id = Pcm.idNotSet

            Try
                If (id <> Nothing) AndAlso (id <> Pcm.idNotSet) Then
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                Return False
            End Try
        End Function

        Public Shared Function GotSqlParam(ByVal aDate As Date) As Boolean

            ' checks to see if the value in aDate is a valid param for SQL select command
            '
            ' vars passed:
            '   aDate - a date
            ' 
            ' returns:
            '   TRUE - if aDate <> Nothing And aDate <> Pcm.NoDate
            '   FALSE - if aDate = Nothing or aDate = Pcm.NoDate

            Try
                If (aDate <> Nothing) AndAlso (aDate <> Pcm.NoDate) Then
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                Return False
            End Try
        End Function

        Public Shared Function GotSqlParam(ByVal aString As String) As Boolean

            ' checks to see if the value in aId is a valid param for SQL select command
            '
            ' vars passed:
            '   aString - a string
            ' 
            ' returns:
            '   TRUE - if aString <> Nothing And aString <> ""
            '   FALSE - if aString = Nothing or aString = ""

            Try
                If (aString <> Nothing) AndAlso (aString <> "") Then
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                Return False
            End Try
        End Function

#End Region

#Region " SQL Column value string "

        Private Shared Function SqlColumnValueString(ByVal TableName As String,
                                                     ByVal ColumnName As String,
                                                     ByVal Value As Object,
                                                     Optional ByVal aOperator As String = "=",
                                                     Optional ByVal NeedBrackets As Boolean = True) As String

            ' creates the selection string for one field
            '
            ' vars passed:
            '   TableName - Name of Table
            '   ColumnName - Name of column
            '   Value - value for field
            '   Operator - operator to use in selection
            '     valid values are "=", "<", ">", "<=", ">=", "<>"
            '       "IN", "LIKE" - when aValue is a string
            '   NeedBrackets - flag to use brackets
            '     TRUE if want "(TableName.ColumnName = Value)"
            '     FALSE if want "TableName.ColumnName = Value"
            '     
            ' returns:
            '   note: if TableName = "", then "TableName." is omitted in all return values
            '   if aValue is Nothing:
            '       "(TableName.ColumnName Is Null)"
            '   if aValue is a boolean:
            '       "(TableName.ColumnName = TRUE)" or "(TableName.ColumnName = FALSE)" 
            '       (no single quotes for boolean value)
            '   if aValue is a date:
            '       "(TableName.ColumnName op #mm/dd/yyyy#)" 
            '       where op is the operator
            '   if aValue is a string
            '       "(TableName.ColumnName op "XXX")" or "(TableName.ColumnName Is Null)"
            '       where op is the operator
            '   if aValue is a number
            '       "(TableName.ColumnName op NNN)" or "(TableName.ColumnName Is Null)"
            '       where op is the operator

            Dim ValueText As String = String.Empty

            Try
                If Not Value Is Nothing Then
                    If TypeOf Value Is Date Then
                        Dim ValueDate As Date = CDate(Value)
                        ValueText = ValueDate.ToShortDateString
                    Else
                        ValueText = Value.ToString
                    End If
                End If
            Catch ex As Exception
                ValueText = ""
            End Try

            aOperator = aOperator.ToUpper()
            Select Case aOperator
            ' normal comparison operators
                Case "=", "<", ">", "<=", ">=", "<>"
                ' do nothing

                ' if an operator for just strings
                Case "IN", "LIKE"
                    Try
                        If Not (Value Is Nothing) Then              ' if got a value
                            If Not (TypeOf Value Is String) Then    ' and value is not a string
                                aOperator = "="                     ' then use default operator
                            End If
                        Else                                        ' else no value
                            aOperator = "="                         ' then use default operator
                        End If
                    Catch ex As Exception
                        aOperator = "="
                    End Try

                Case Else
                    aOperator = "="
            End Select

            Dim SqlColValStr As String = ""                                     ' initial Sql Column Value string
            If NeedBrackets Then                                                ' if need brackets
                SqlColValStr = "("                                              ' then add starting bracket
            End If
            If TableName = "" Then                                              ' if no table name
                SqlColValStr &= ColumnName                                      ' then exclude TableName.
            Else                                                                ' else got table name
                SqlColValStr &= String.Format("{0}.{1}", TableName, ColumnName) ' include table name
            End If

            If TypeOf Value Is Boolean Then
                SqlColValStr &= " = " & ValueText
            Else
                If ValueText = "" Then
                    If aOperator = "<>" Then
                        SqlColValStr &= " Is Not Null"
                    Else
                        SqlColValStr &= " Is Null"
                    End If
                Else
                    If TypeOf Value Is Date Then                                            ' if value is a date
                        SqlColValStr &= String.Format(" {0} #{1}#", aOperator, ValueText)   ' format value text as a date
                    ElseIf TypeOf Value Is Boolean Then                                     ' if value is a boolean
                        If aOperator = "<>" Then                                            ' if not equal
                            SqlColValStr &= " <> " & ValueText                              ' use not equal value text
                        Else                                                                ' else not using not equal
                            SqlColValStr &= " = " & ValueText                               ' alway use equal (no >, < for booleans)
                        End If
                    ElseIf TypeOf Value Is String Then
                        SqlColValStr &= String.Format(" {0}""{1}""", aOperator, ValueText.Replace("'", "''"))
                    Else
                        SqlColValStr &= String.Format(" {0} {1}", aOperator, ValueText)
                    End If
                End If
            End If
            If NeedBrackets Then
                SqlColValStr &= ")"
            End If
            Return SqlColValStr
        End Function

#End Region

#Region " SQL WHERE "

        Public Shared Sub SqlRemoveWhere(ByRef SqlSelectText As String)

            ' removes the ending " WHERE" and text after from a Sql Select command
            '
            ' vars passed:
            '   SqlSelectText - sql select command text to remove "where" from.  
            '     Note: Passed ByRef

            Try
                Dim WherePos As Integer = SqlSelectText.ToUpper().LastIndexOf(" WHERE")
                If WherePos > 0 Then
                    SqlSelectText = SqlSelectText.Remove(WherePos, SqlSelectText.Length - WherePos)
                End If
            Catch ex As Exception
                Return
            End Try
        End Sub

        Public Shared Function SqlWhereColumnValueString(ByVal TableName As String,
                                                         ByVal ColumnName As String,
                                                         ByVal aValue As Object,
                                                         Optional ByVal aOperator As String = "=",
                                                         Optional ByVal NeedBrackets As Boolean = False) As String

            ' creates the selection string for one field
            '
            ' vars passed:
            '   TableName - Name of Table
            '   ColumnName - Name of column
            '   aValue - value for field
            '   aOperator - operator to use in selection
            '     valid values are "=", "<", ">", "<=", ">=", "<>"
            '       "IN", "LIKE" - when aValue is a string
            '   NeedBrackets - flag to use brackets
            '     TRUE if want "(TableName.ColumnName = Value)"
            '     FALSE if want "TableName.ColumnName = Value"
            '     
            ' returns:
            '   note: if TableName = "", then "TableName." is omitted in all return values
            '   if aValue is Nothing:
            '       "(TableName.ColumnName Is Null)"
            '   if aValue is a boolean:
            '       "(TableName.ColumnName = TRUE)" or "(TableName.ColumnName = FALSE)" 
            '       (no single quotes for boolean value)
            '   if aValue is a date:
            '       "(TableName.ColumnName op #mm/dd/yyyy#)" 
            '       where op is the operator
            '   if aValue is a string
            '       "(TableName.ColumnName op "XXX")" or "(TableName.ColumnName Is Null)"
            '       where op is the operator
            '   if aValue is a number
            '       "(TableName.ColumnName op NNN)" or "(TableName.ColumnName Is Null)"
            '       where op is the operator

            Try
                Return SqlColumnValueString(TableName, ColumnName, aValue, aOperator, NeedBrackets)
            Catch ex As Exception
                Return ""
            End Try
        End Function

        Public Shared Function SqlWhereColumnValueString(ByVal aParamCount As Integer,
                                                         ByVal TableName As String,
                                                         ByVal ColumnName As String,
                                                         ByVal aValue As Object,
                                                         Optional ByVal aOperator As String = "=",
                                                         Optional ByVal aWhereOpp As String = "AND",
                                                         Optional ByVal NeedBrackets As Boolean = False) As String

            ' creates the selection string for one field
            '
            ' vars passed:
            '   aParamCount - how many WHERE parameters already
            '   TableName - Name of Table
            '   ColumnName - Name of column
            '   aValue - value for field
            '   aOperator - operator to use in selection
            '     valid values are "=", "<", ">", "<=", ">=", "<>"
            '       "IN", "LIKE" - when aValue is a string
            '   aWhereOpp - logical operator (AND/OR)
            '   NeedBrackets - flag to use brackets
            '     TRUE if want "(TableName.ColumnName = Value)"
            '     FALSE if want "TableName.ColumnName = Value"
            '     
            ' returns:
            '   note: if TableName = "", then "TableName." is omitted in all return values
            '   if aValue is Nothing:
            '       "(TableName.ColumnName Is Null)"
            '   if aValue is a boolean:
            '       "(TableName.ColumnName = TRUE)" or "(TableName.ColumnName = FALSE)" 
            '       (no single quotes for boolean value)
            '   if aValue is a date:
            '       "(TableName.ColumnName op #mm/dd/yyyy#)" 
            '       where op is the operator
            '   if aValue is a string
            '       "(TableName.ColumnName op "XXX")" or "(TableName.ColumnName Is Null)"
            '       where op is the operator
            '   if aValue is a number
            '       "(TableName.ColumnName op NNN)" or "(TableName.ColumnName Is Null)"
            '       where op is the operator
            '   if aParamCount > 0, add "AND " or "OR" to start of return value

            Dim WhereSql As String = ""                 ' initial Where string
            Try
                If aParamCount > 0 Then                 ' if not first param
                    aWhereOpp = aWhereOpp.ToUpper
                    If (aWhereOpp <> "AND") AndAlso (aWhereOpp <> "OR") Then
                        aWhereOpp = "AND"
                    End If
                    WhereSql = String.Format(" {0} ", aWhereOpp)    ' add where opp (space before and after)
                End If
                Return WhereSql & SqlColumnValueString(TableName, ColumnName, aValue, aOperator, NeedBrackets)
            Catch ex As Exception
                Return ""
            End Try
        End Function

#End Region

#Region " SQL INNER JOIN "

        Public Shared Function InnerJoin(ByVal JoinToTableName As String,
                                         ByVal JoinToColumnName As String,
                                         ByVal JoinFromTableName As String,
                                         ByVal JoinFromColumnName As String,
                                         Optional ByVal InsertINNERJOIN As Boolean = False,
                                         Optional ByVal InsertAND As Boolean = False) As String

            ' gets the "INNER JOIN" SQL command text.
            '
            ' vars passed:
            '   JoinToTableName - table name to join to
            '   JoinToColumnName - column name in JoinToTable to join to JoinFromColumnName
            '   JoinFromTableName - table name to join from (main table in query)
            '   JoinFromColumnName - column name in JoinFromTable to join to JoinToColumnName
            '   InsertINNERJOIN - if TRUE, add text " INNER JOIN JoinToTableName ON " to start of InnerJoin text
            '   InsertAND - if TRUE and InsertINNERJOIN is FALSE, add " AND " to start of InnerJoin text
            '
            ' returns:
            '   if InsertINNERJOIN = TRUE 
            '     "FROM (JoinFromTableName INNER JOIN 
            '       JoinToTableName ON JoinToTableName.JoinToColumnName = JoinFromTAbleName.JoinFromColumnName)"
            '   if InsertAND = TRUE and InsertINNERJOIN = FALSE
            '     "AND JoinToTableName.JoinToColumnName = JoinFromTAbleName.JoinFromColumnName)"
            '   else
            '     "(JoinToTableName.JoinToColumnName = JoinFromTAbleName.JoinFromColumnName)"
            '   "" - something went wrong

            Try
                Dim IjStr As String = ""
                If InsertINNERJOIN Then                                 ' if want to insert inner join text
                    IjStr = String.Format("FROM ({0} INNER JOIN {1} ON ", JoinFromTableName, JoinToTableName) ' add inner join SQL text
                ElseIf InsertAND Then                                   ' else if do not want inner join but want AND
                    IjStr = "AND "                                      ' add AND SQL text
                End If
                ' add in join SQL text
                IjStr &= String.Format("{0}.{1} = {2}.{3})", JoinFromTableName, JoinFromColumnName, JoinToTableName, JoinToColumnName)
                Return IjStr
            Catch ex As Exception
                Return ""
            End Try
        End Function

        Public Shared Function AddOnInnerJoinAnd(ByVal InnerJoinSql As String, ByVal JoinToAdd As String) As String

            ' adds an additional join element to the INNER JOIN SQL 
            '
            ' vars passed:
            '   InnerJoinSql - current INNER JOIN SQL statement, ends with ")"
            '   JoinToAdd - join element to add, starts with " AND " and ends with ")"
            '
            ' returns:
            '   InnerJoinSql with out the ending ")", and the JoinToAdd added to it
            '   "" - something went wrong

            Try
                If InnerJoinSql.IndexOf(")") = InnerJoinSql.Length - 1 Then
                    InnerJoinSql = InnerJoinSql.Remove(InnerJoinSql.Length - 1, 1)
                End If

                Return String.Format("{0} {1}", InnerJoinSql, JoinToAdd)
            Catch ex As Exception
                Return ""
            End Try
        End Function

#End Region

#Region " SQL SET "

        Public Shared Function SqlSetColumnValueString(ByVal TableName As String,
                                                       ByVal ColumnName As String,
                                                       ByVal aValue As Object,
                                                       Optional ByVal aOperator As String = "=",
                                                       Optional ByVal NeedBrackets As Boolean = False) As String

            ' creates the selection string for one field
            '
            ' vars passed:
            '   TableName - Name of Table
            '   ColumnName - Name of column
            '   aValue - value for field
            '   aOperator - operator to use in selection
            '     valid values are "=", "<", ">", "<=", ">=", "<>"
            '       "IN", "LIKE" - when aValue is a string
            '   NeedBrackets - flag to use brackets
            '     TRUE if want "(TableName.ColumnName = Value)"
            '     FALSE if want "TableName.ColumnName = Value"
            '     
            ' returns:
            '   note: if TableName = "", then "TableName." is omitted in all return values
            '   if aValue is Nothing:
            '       "TableName.ColumnName Is Null"
            '   if aValue is a boolean:
            '       "TableName.ColumnName = TRUE" or "TableName.ColumnName = FALSE" 
            '       (no single quotes for boolean value)
            '   if aValue is a date:
            '       "TableName.ColumnName op #mm/dd/yyyy#" 
            '       where op is the operator
            '   if aValue is a string
            '       "TableName.ColumnName op "XXX"" or "TableName.ColumnName Is Null"
            '       where op is the operator
            '   if aValue is a number
            '       "TableName.ColumnName op NNN" or "TableName.ColumnName Is Null"
            '       where op is the operator

            Try
                Return SqlColumnValueString(TableName, ColumnName, aValue, aOperator, NeedBrackets)
            Catch ex As Exception
                Return ""
            End Try
        End Function

#End Region

    End Class

End Namespace
