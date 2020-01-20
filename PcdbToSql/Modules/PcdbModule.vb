Imports DevExpress.XtraEditors
Imports DevExpress.XtraGrid
Imports EaTools.IO

Friend Module PcdbModule

#Region " Constants "

    Private Const NotFoundMsg1 As String = "Value """
    Private Const NotFoundMsg2 As String = """ not found in list.  Please select a value from the list."

#End Region

#Region " Get Key Field Column Name "

    Public Function GetKeyFieldColumnName(ByVal aTable As DataTable, _
                                          Optional ByVal ShowErr As Boolean = True) As String

        ' gets the key field column name for a table
        '
        ' vars passed:
        '   aTable - table to get key field name for
        '   ShowErr - set to FALSE if do not want to show error messages
        '   
        ' returns:
        '   <> "" - name of key field
        '   "" - 
        '       table not found
        '       no key field
        '       multi-column key
        '       something went wrong

        Try
            If aTable Is Nothing Then
                If ShowErr Then
                    MessageBox.Show("Could not get key field name from table.  No table.", _
                                    Pcm.etKeyColumnErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
                Return String.Empty
            End If
            If aTable.PrimaryKey.Count = 1 Then
                Return aTable.PrimaryKey(0).ColumnName
            ElseIf aTable.PrimaryKey.Count > 1 Then
                If ShowErr Then
                    MessageBox.Show(String.Format("Could not get key value for table ""{0}"".  Table has multi-part key", aTable.TableName), _
                                    Pcm.etKeyColumnErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
                Return String.Empty
            Else
                If ShowErr Then
                    MessageBox.Show(String.Format("Could not get key value for table ""{0}"".  Table has no key field.", aTable.TableName), _
                                    Pcm.etKeyColumnErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
                Return String.Empty
            End If
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("PcdbModule.GetKeyFieldColumnName(aTable,ShowErr).  " & ex.Message)
            Return ""
        End Try
    End Function

#End Region

#Region " Control Info "

    Public Function MaskBoxConvert(ByVal aControl As Control) As Control

        ' makes sure control is not a DevExpress TextBoxMaskBox
        '
        ' vars passed:
        '   aControl - the control to convert
        '
        ' returns:
        '   if aControl is a DevExpress TextBoxMaskBox - aControl.Parent
        '   else - aControl

        Try
            If TypeOf aControl Is TextBoxMaskBox Then
                Return aControl.Parent
            End If
            Return aControl
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("PcdbModule.MaskBoxConvert.  " & ex.Message)
            Return aControl
        End Try
    End Function

    Public Function ControlName(ByVal aControl As Control) As String

        ' returns the name of a control
        ' notes from question asked to DevExpress:
        '   Based on the following design, some of our edit controls consist of two 
        '   controls. The internal control represents an editable area and is placed
        '   inside the external control The External control is a TextEdit class 
        '   instance which is used to implement the editor's appearance. The internal 
        '   Control() 's type is TextBoxMaskBox. It also gets focused. So, the 
        '   ActiveControl method returns an inner control which is actually 
        '   TextBoxMaskBox. To access to Name and other properties of the external
        '   control you can use its Parent property
        '
        ' vars passed:
        '   aControl - the control with the name to get
        '
        ' returns:
        '   the name of the control
        ' 

        Try
            Return MaskBoxConvert(aControl).Name
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("PcdbModule.ControlName.  " & ex.Message)
            Return Nothing
        End Try
    End Function

#End Region

#Region " GetObjectValue "

    Public Function GetObjectValue(ByVal aValue As Object, _
                                   ByVal aNullValue As Object) As Object

        ' gets the value from an object.  Converts a null to the desired null value
        '
        ' vars passed:
        '   aValue - usually e.Value from the CellValueChanged event
        '   aNullValue - the value to return if aCellValue is null
        '
        ' returns:
        '   if aValue = Null
        '     aNullValue
        '   else
        '     aValue  

        Try
            If aValue Is DBNull.Value Then
                Return aNullValue
            Else
                Return aValue
            End If
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("PcdbModule.GetObjectValue.  " & ex.Message)
            Return aNullValue
        End Try
    End Function

#End Region

#Region " Row Subs and Funcs "

#Region " Clear Row Column Values "

    Public Sub ClearRowColumnValues(row As DataRow)

        ' makes all column values in the row dbNull
        '
        ' vars passed:
        '   row - data row to have all values set to dbNull

        Try
            For i As Integer = 0 To row.Table.Columns.Count - 1     ' for each column in row
                row(i) = DBNull.Value                               ' clear the column value
            Next
        Catch ex As Exception
            Return
        End Try
    End Sub

#End Region

#Region " Confirm Rows Match "

    Public Function ConfirmRowsMatch(Row1 As DataRow, _
                                     Row2 As DataRow) As Integer

        ' confirms values in rows match 
        '
        ' vars passed:
        '   Row1 - first row to compare
        '   Row2 - row to compare row1 to
        '
        ' returns:
        '   Pcm.NoErrors - no errors, all values match
        '   Pcm.teUpdateErr - values do not match
        '   Pcm.OtherExErr - something went wrong

        Const RoundToDecimals As Integer = 12
        Try
            ' confirm all values in updated rows match values in tables
            For i As Integer = 0 To Row2.ItemArray.Length - 1                           ' for each value in row2
                If IsDBNull(Row2.ItemArray(i)) Then                                     ' if value is dbNull
                    If Not IsDBNull(Row1.ItemArray(i)) Then                             ' if row1 value is not DbNull
                        ShowConfirmationError(Row2, Row1, i)                            ' show error
                        Return Pcm.teUpdateErr                                          ' return error
                    End If
                Else                                                                    ' else value in row2 is NOT DbNull
                    If IsDBNull(Row1.ItemArray(i)) Then                                 ' if rpw1 value is DbNll
                        ShowConfirmationError(Row2, Row1, i)                            ' show err
                        Return Pcm.teUpdateErr                                          ' return error
                    Else                                                                ' else got values in row2 and row1
                        If Not (Row2.ItemArray(i).Equals(Row1.ItemArray(i))) Then       ' if values do not match
                            If TypeOf (Row2.ItemArray(i)) Is Double Then                ' if values are double
                                ' get values rounded to 12 decimals (loaded value will be rounded to 12 decimals)
                                Dim Row2Dbl As Double = Math.Round(CDbl(Row2.ItemArray(i)), RoundToDecimals, MidpointRounding.AwayFromZero)
                                Dim Row1Dbl As Double = Math.Round(CDbl(Row1.ItemArray(i)), RoundToDecimals, MidpointRounding.AwayFromZero)
                                If Row2Dbl <> Row1Dbl Then                              ' if still do not match
                                    ShowConfirmationError(Row2, Row2Dbl, Row1Dbl, i)    ' show error
                                    Return Pcm.teUpdateErr                              ' return error
                                End If
                            ElseIf TypeOf (Row2.ItemArray(i)) Is Date Then              ' if a date
                                ' decimal seconds can be rounded off, so do not check anything past integer seconds
                                Dim Row2Date As Date = CDate(Row2.ItemArray(i))         ' get date row 2
                                Dim Row1Date As Date = CDate(Row1.ItemArray(i))         ' get date row 1
                                Dim diff As TimeSpan = Row2Date - Row1Date
                                If Math.Abs(diff.TotalSeconds) >= 1 Then
                                    ShowConfirmationError(Row2, Row1, i)                ' show error
                                    Return Pcm.teUpdateErr                              ' return error
                                End If
                            ElseIf TypeOf (Row2.ItemArray(i)) Is Decimal Then           ' if decimal #'s
                                Dim Row2Dec As Decimal = CDec(Row2.ItemArray(i))        ' get decimal row 2
                                Dim Row1Dec As Decimal = CDec(Row1.ItemArray(i))        ' get decimal row 1
                                If Math.Abs(Row2Dec - Row1Dec) > 0.005D Then            ' if more than half a cent off
                                    ShowConfirmationError(Row2, Row1, i)                ' show err
                                    Return Pcm.teUpdateErr                              ' return error
                                End If
                            Else                                                        ' else not a double, date, or decimal
                                ShowConfirmationError(Row2, Row1, i)                    ' show err
                                Return Pcm.teUpdateErr                                  ' return error
                            End If
                        End If
                    End If
                End If
            Next
            Return Pcm.NoErrors
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("PcdbModule.ConfirmRowsMatch.  " & ex.Message)
            Return Pcm.OtherExErr
        End Try
    End Function

    Private Function GetConfirmationColumnValue(aRow As DataRow, _
                                                ItemIndex As Integer) As String

        ' gets the column value as a string from the ItemArray
        '
        ' vars passed:
        '   aRow - data row with column value to get
        '   ItemIndex - index into ItemArray for column value to get
        '
        ' returns:
        '   the value as a string with default format for the value type

        Try
            Dim ColumnValue As String
            If IsDBNull(aRow.ItemArray(ItemIndex)) Then                             ' if value is DbNull
                ColumnValue = "[Null]"                                              ' use "Null" text
            Else                                                                    ' else got a value
                If TypeOf aRow.ItemArray(ItemIndex) Is Decimal Then                 ' if a decimal value
                    ColumnValue = String.Format("{0:C}", aRow.ItemArray(ItemIndex)) ' use currency format
                ElseIf TypeOf aRow.ItemArray(ItemIndex) Is Date Then                ' else if a date
                    ColumnValue = String.Format("{0:G}", aRow.ItemArray(ItemIndex)) ' use general long data format
                ElseIf TypeOf aRow.ItemArray(ItemIndex) Is Double Then              ' else if a double
                    ColumnValue = String.Format("{0:###############.###############}", aRow.ItemArray(ItemIndex))   ' use 15 digits
                Else
                    ColumnValue = String.Format("{0}", aRow.ItemArray(ItemIndex))   ' use value with default formatting
                End If
            End If
            Return ColumnValue                                                      ' return column value as string
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Private Sub ShowConfirmationError(Row1 As DataRow, _
                                      Row2 As DataRow, _
                                      ErrorIndex As Integer)

        ' shows the error message when an row's column value does not match another row column value in the 
        '
        ' vars passed:
        '   Row1 - first row to compare
        '   Row2 - row to compare row1 to
        '   ErrorIndex - index into Row2.ItemArray where values do not match

        Try
            Dim KeyValuesStr As String = String.Empty
            Dim KeyColumnsStr As String = String.Empty
            Dim i As Integer = 0
            For Each KeyCol As DataColumn In Row1.Table.PrimaryKey
                If i > 0 Then
                    KeyValuesStr &= ", "
                    KeyColumnsStr &= ", "
                End If
                If IsDBNull(Row1(KeyCol)) Then
                    KeyValuesStr &= "[Null]"
                Else
                    KeyValuesStr &= Row1(KeyCol).ToString
                End If
                KeyColumnsStr &= KeyCol.ColumnName
                i += 1
            Next

            Dim Row1ValStr As String = GetConfirmationColumnValue(Row1, ErrorIndex)
            Dim Row2ValStr As String = GetConfirmationColumnValue(Row2, ErrorIndex)

            MessageBox.Show(String.Format("Row values do not match in table.{0}Table: {1}{0}Key Column(s): {2}{0}Key Value(s): {3}{0}Column Name with non matching values:{4}{0}Non matching value 1: {5}{0}Non matching value 2: {6}", _
                                          vbCrLf, Row2.Table.TableName, KeyValuesStr, KeyColumnsStr, Row2.Table.Columns(ErrorIndex).ColumnName, Row1ValStr, Row2ValStr), _
                            Pcm.etConfirmationErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("PcdbModule.ShowOtherErrorMessage.  " & ex.Message)
        End Try
    End Sub

    Private Sub ShowConfirmationError(Row2 As DataRow, _
                                      Row2Value As Double, _
                                      Row1Value As Double, _
                                      ErrorIndex As Integer)

        ' shows the error message when an row's column value does not match another row column value in the 
        '
        ' vars passed:
        '   Row2 - row to compare to
        '   Row2Value - non matching row 2 value
        '   Row1Value - non matching row 1 value
        '   ErrorIndex - index into Row2.ItemArray where values do not match

        Try
            Dim KeyValuesStr As String = String.Empty
            Dim KeyColumnsStr As String = String.Empty
            Dim i As Integer = 0
            For Each KeyCol As DataColumn In Row2.Table.PrimaryKey
                If i > 0 Then
                    KeyValuesStr &= ", "
                    KeyColumnsStr &= ", "
                End If
                If IsDBNull(Row2(KeyCol)) Then
                    KeyValuesStr &= "[Null]"
                Else
                    KeyValuesStr &= Row2(KeyCol).ToString
                End If
                KeyColumnsStr &= KeyCol.ColumnName
                i += 1
            Next

            Dim Row1ValStr As String = String.Format("{0:###############.###############}", Row1Value)
            Dim Row2ValStr As String = String.Format("{0:###############.###############}", Row2Value)

            MessageBox.Show(String.Format("Row values do not match in table.{0}Table: {1}{0}Key Column(s): {2}{0}Key Value(s): {3}{0}Column Name with non matching values:{4}{0}Non matching value 1: {5}{0}Non matching value 2: {6}", _
                                          vbCrLf, Row2.Table.TableName, KeyValuesStr, KeyColumnsStr, Row2.Table.Columns(ErrorIndex).ColumnName, Row1ValStr, Row2ValStr), _
                            Pcm.etConfirmationErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("PcdbModule.ShowOtherErrorMessage.  " & ex.Message)
        End Try
    End Sub

#End Region

#Region " Delete Row Sorted Grid / Remove All BM Rows "

    Public Function DeleteRowSortedGrid(gc As GridControl) As Boolean

        ' this is a workaround to a bug in DevExp XtraGrids.  Use this func to
        ' delete the focused row from a sorted grid
        '
        ' vars passed:
        '   aGridControl - GridControl with row to delete
        '   
        ' returns:
        '   TRUE - row was deleted
        '   FALSE - row was not deleted

        Try
            Dim gv As Views.Grid.GridView = CType(gc.MainView, Views.Grid.GridView)                     ' get grid view
            Dim rh As Integer = gv.FocusedRowHandle                                                     ' get current row handle
            If (rh = GridControl.InvalidRowHandle) OrElse (rh = GridControl.AutoFilterRowHandle) Then   ' if not a data row
                Return False                                                                            ' return false, cannot delete
            End If
            gv.DeleteRow(rh)                                                                            ' delete the row
            Return True                                                                                 ' return TRUE, row was deleted
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("PcdbModule.DeleteRowSortedGrid.  " & ex.Message)
            Return False
        End Try
    End Function

    Public Function RemoveAllBmRows(aBindingManager As BindingManagerBase) As Integer

        ' removes all rows for a BindingManager
        '
        ' vars passed:
        '   aBindingManager - the binding manager with rows to remove
        '
        ' Returns:
        '   NoErrors - no errors in removing
        '   teDeleteErr - there was an error wil removing a row

        Try
            Dim i As Integer
            ' to ensure will not get "Index x is not non-negative and below total rows count" error message
            ' set position to -1 before removing rows from the binding manager
            aBindingManager.Position = -1
            For i = 0 To aBindingManager.Count - 1  ' for each row in the binding manager
                aBindingManager.RemoveAt(0)         ' remove the row
            Next
        Catch ex As Exception                       ' if had an error
            PcdbModule.ShowOtherErrorMessage("PcdbModule.RemoveAllBmRows.  " & ex.Message)
            Return Pcm.teDeleteErr                  ' return delete error
        End Try
        Return Pcm.NoErrors                         ' if no errors, return NoError
    End Function

#End Region

#Region " Field Select String "

    Public Function FieldSelectString(ByVal FieldName As String, _
                                      ByVal aValue As Object, _
                                      Optional ByVal aOperator As String = "=", _
                                      Optional ByVal UseSingleQuotes As Boolean = True) As String

        ' creates the selection string for one field
        '
        ' vars passed:
        '   FieldName - Name of field for selection string
        '   aValue - value for field
        '   aOperator - operator to use in selection
        '     valid values are "=", "<", ">", "<=", ">=", "<>"
        '       "IN", "LIKE" - when aValue is a string
        '     
        ' returns:
        '   if aValue is Nothing:
        '       "[FieldName] Is Null"
        '   if aValue is a boolean:
        '       "[FieldName] = TRUE" or "[FieldName] = FALSE" 
        '       (no single quotes for boolean value)
        '   if aValue is a date:
        '     if the time is at midnight (no time)
        '       "[FieldName] op #MM/dd/yyyy#" 
        '     if the tine is not 0
        '       "[FieldName] op #MM/dd/yyyy HH:mm:SS tt#" 
        '       where op is the operator
        '   if aValue is not a boolean
        '       "[FieldName] op 'XXX'" or "[FieldName] Is Null"
        '       where op is the operator
        '   "" - something went wrong

        Dim ValueText As String = ""

        Try
            If aValue IsNot Nothing Then
                If TypeOf aValue Is Date Then
                    Dim vDate As Date = CDate(aValue)
                    Dim NoTime As System.TimeSpan = Pcm.NoDate.TimeOfDay
                    If vDate.TimeOfDay = NoTime Then
                        ValueText = vDate.ToShortDateString
                    Else
                        ValueText = String.Format("{0} {1}", vDate.ToShortDateString, vDate.ToLongTimeString)
                    End If
                Else
                    ValueText = aValue.ToString
                End If
            End If
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("PcdbModule.FieldSelectString - ValueText.  " & ex.Message)
        End Try

        Try
            aOperator = aOperator.ToUpper()
            Select Case aOperator
                ' normal comparison operators
                Case "=", "<", ">", "<=", ">=", "<>"
                    ' do nothing

                    ' if an opetaor for just strings
                Case "IN", "LIKE"
                    Try
                        If Not (aValue Is Nothing) Then             ' if got a value
                            If Not (TypeOf aValue Is String) Then   ' and value is not a string
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

            Dim FieldNameText As String = String.Format("[{0}]", FieldName)

            ' if no value and (want equal or not equal) and aValue is not a string
            If ValueText = "" AndAlso (aOperator = "=" OrElse aOperator = "<>") AndAlso _
                    Not (TypeOf aValue Is String) Then
                FieldNameText &= " Is "                     ' add in " is "
                If aOperator = "<>" Then                    ' if want <> nothing
                    FieldNameText &= "Not "                 ' add in " not "
                End If
                Return FieldNameText & "Null"               ' returns "FieldName Is Null" or "FieldName Is Not Null" 

                ' else got a value 
            Else
                If TypeOf aValue Is Date Then                                       ' if a date value, format ValueText as date text
                    Return String.Format("{0} {1} #{2}#", FieldNameText, aOperator, ValueText)
                ElseIf TypeOf aValue Is Boolean Then                                ' if a boolean value
                    If aOperator = "<>" Then                                        ' if not equal
                        Return String.Format("{0} <> {1}", FieldNameText, ValueText) ' then return not equal value text
                    Else                                                            ' else other than not equal
                        Return String.Format("{0} = {1}", FieldNameText, ValueText) ' else use equal (no greater than or less than)
                    End If
                Else
                    If UseSingleQuotes Then
                        Return String.Format("{0} {1} '{2}'", FieldNameText, aOperator, ValueText.Replace("'", "''"))
                    Else
                        Return String.Format("{0} {1} {2}", FieldNameText, aOperator, ValueText.Replace("'", "''"))
                    End If
                End If
            End If
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("PcdbModule.FieldSelectString.  " & ex.Message)
            Return ""
        End Try
    End Function

#End Region

#Region " MatchingValuesCount "

    Public Function MatchingValuesCount(ByVal aTable As DataTable, _
                                        ByVal ColumnName As String, _
                                        ByVal aValue As Object) As Integer

        ' gets the number of rows in a table with a specific value in a column
        '
        ' vars passed:
        '   aTable - table to check
        '   ColumnName - name of column in aTable 
        '   aValue - value to check for

        Try
            Dim SelectStr As String = FieldSelectString(ColumnName, aValue)
            Dim DupRows As DataRow() = aTable.Select(SelectStr)
            Return DupRows.Length
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("PcdbModule.MatchingValuesCount(aTable,ColumnName,aValue).  " & ex.Message)
            Return 0
        End Try
    End Function

#End Region

#Region " Rows Match "

    Public Function RowsMatch(Row1 As DataRow, _
                              Row2 As DataRow, _
                              ColumnNames() As String) As Integer

        ' verifies if two rows match.  Only columns with names in ColumnNames are checked
        '
        ' vars passed:
        '   Row1 - one row to check
        '   Row2 - row to match with Row1
        '   ColumnNames() - array of column names to check.  
        '
        ' returns:
        '   ColumnNames.Length - all values match
        '   0 to ColumnNames.Count - 1 - index in ColumnNames of mismatch
        '   -1 - something went wrong

        Try
            For i As Integer = 0 To ColumnNames.Count - 1                       ' for each column to check
                If Row1.IsNull(ColumnNames(i)) Then                             ' if row 1 column value is null
                    If Not Row2.IsNull(ColumnNames(i)) Then                     ' if row 2 column value is not null
                        Return i                                                ' return column name index
                    End If
                Else                                                            ' else row 1 column value is not null
                    If Row2.IsNull(ColumnNames(i)) Then                         ' if row 2 column value is null
                        Return i                                                ' return column name index
                    Else                                                        ' else row 2 is not null
                        If Row1(ColumnNames(i)) IsNot Row2(ColumnNames(i)) Then ' if row 1 column value <> row 2 column value
                            Return i                                            ' return column name index
                        End If
                    End If
                End If
            Next
            Return ColumnNames.Length          ' if got here, then return -1, rows match
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("PcdbModule.RowMatch.  " & ex.Message)
            Return -1
        End Try
    End Function

#End Region

#End Region

#Region " Find... "

#Region " Find Row "

    Public Function FindRow(ByVal aTable As DataTable, ByVal aKeyValue As Object) As DataRow

        ' tries to find a row in a table using a key value
        '
        ' vars passed:
        '   aTable - table with data to find
        '   aKeyValue - key field value to find
        '
        ' returns:
        '   data row of row with matching key value
        '   Nothing - 
        '       table not found
        '       no data in table
        '       row key value not found
        '       some error happened

        Try
            If aTable Is Nothing Then               ' if no data table
                Return Nothing                      ' return nothing
            End If
            If aTable.Rows.Count = 0 Then           ' if no data
                Return Nothing                      ' return nothing
            End If
            Return aTable.Rows.Find(aKeyValue)      ' return find value
        Catch ex As Exception
            ' no error message here
            Return Nothing
        End Try
    End Function

    Public Function FindRow(ByVal aTable As DataTable, ByVal KeyValues() As Object) As DataRow

        ' tries to find a row in a table using a multi-part key
        '
        ' vars passed:
        '   aTable - table with data to find
        '   KeyValues - array of values for the multi part key.  
        '       NOTE - the order is not checked.  make sure to pass correct values in correct array locations
        '
        ' returns:
        '   data row of row with matching key value
        '   Nothing - 
        '       table not found
        '       no data in table
        '       row key value not found
        '       some error happened

        Try
            If aTable Is Nothing Then                               ' if no data table
                Return Nothing                                      ' return nothing
            End If
            If aTable.Rows.Count = 0 Then                           ' if no data
                Return Nothing                                      ' return nothing
            End If
            If aTable.PrimaryKey.Length <> KeyValues.Length Then    ' if primary key length <> key values length
                Return Nothing                                      ' return nothing
            End If
            Return aTable.Rows.Find(KeyValues)      ' return find value
        Catch ex As Exception
            ' no error message here
            Return Nothing
        End Try
    End Function

    Public Function FindRow(ByVal aTable As DataTable, _
                            ByVal KeyValues() As Object, _
                            ByVal KeyColNames() As String) As DataRow

        ' tries to find a row in a table using a multi-part key
        ' orders the KeyValues to match the order of the primary key in aTableName
        '
        ' vars passed:
        '   aTableName - name of table with data to find
        '   KeyValues - array of values for the multi part key.  
        '   KeyColNames - array of key column names, in the order of the keyValues array is set
        '
        ' returns:
        '   data row of row with matching key value
        '   Nothing - 
        '       table not found
        '       no data in table
        '       row key value not found
        '       some error happened

        Try
            If aTable Is Nothing Then                               ' if no data table
                Return Nothing                                      ' return nothing
            End If
            If KeyValues.Length <> KeyColNames.Length Then          ' if key values length <> key col names length
                Return Nothing                                      ' return nothing 
            End If
            Dim OrderedKeys(KeyValues.Length - 1) As Object         ' get empty key values array
            Dim i As Integer
            Dim c As Integer
            For i = 0 To KeyColNames.Length - 1                     ' for each key column name
                For c = 0 To aTable.PrimaryKey.Length - 1           ' for each primary key column
                    If KeyColNames(i).ToUpper = aTable.PrimaryKey(c).ColumnName.ToUpper Then    ' if col names match
                        OrderedKeys(c) = KeyValues(i)               ' set key value in ordered array
                        Exit For
                    End If
                Next
                If c >= KeyColNames.Length Then                     ' if key column not found
                    Return Nothing                                  ' return nothing
                End If
            Next

            Return PcdbModule.FindRow(aTable, OrderedKeys)          ' return the FindRow value
        Catch ex As Exception
            ' no error message here
            Return Nothing
        End Try
    End Function

#End Region

#Region " FindRowUsingDataView "

    Public Function FindRowUsingDataView(ByVal aDataView As DataView, _
                                         ByVal aValue As Object, _
                                         ByVal aColumnName As String) As DataRow

        ' finds the matching row using a dataview and row filter
        '
        ' vars passed:
        '   aDataView - data view to use
        '   aValue - value to find
        '   aColumnName - name of column to for aValue to be in
        '
        ' returns: 
        '   aDataRow - 1 and only 1 filtered row was found
        '   Nothing - 
        '     No matching rows found
        '     More than 1 matching row was found
        '     something went wrong

        Dim InitSort As String = aDataView.Sort                             ' save current sort 
        Try
            aDataView.Sort = aColumnName                                    ' sort on column to find value in   
            Dim SomeRowViews() As DataRowView = aDataView.FindRows(aValue)  ' find matching row(s)
            If SomeRowViews.Length = 1 Then                                 ' if only 1 row found
                Return SomeRowViews(0).Row                                  ' return the row
            ElseIf SomeRowViews.Length = 0 Then                             ' if no matches
                Return Nothing                                              ' return nothing
            Else                                                            ' else multiple matches
                MessageBox.Show("Multiple matches found in ""FindRowUsingDataView"".", _
                                Pcm.etInternalErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return Nothing                                              ' return nothing 
            End If
        Catch ex As Exception
            ' no error message here
            Return Nothing
        Finally
            aDataView.Sort = InitSort                                       ' always reset sort
        End Try
    End Function

    Public Function FindRowUsingDataView(ByVal aDataView As DataView, _
                                         ByVal Values() As Object, _
                                         ByVal ColumnNames() As String) As DataRow

        ' finds the matching row using a dataview and row sort
        '
        ' vars passed:
        '   aDataView - data view to use
        '   Values() - values to find
        '   ColumnNames() - name of columns to for Value() to be in
        '
        ' note: Values(i) will be in column ColumnsName(i)
        '
        ' returns: 
        '   DataRowView() - matching rows
        '   Nothing - 
        '     No matching rows found
        '     More than 1 matching row was found
        '     something went wrong

        Dim InitSort As String = aDataView.Sort                             ' save current sort
        Try
            Dim SortStr As String = ""
            For i As Integer = 0 To ColumnNames.Count - 1                   ' for each column name
                If i > 0 Then                                               ' if not the first column name
                    SortStr &= ", "                                         ' add a ", " before the column name
                End If
                SortStr &= ColumnNames(i)                                   ' add the column name to the sort string
            Next
            aDataView.Sort = SortStr                                        ' sort the filter on columns to find values in
            Dim SomeRowViews() As DataRowView = aDataView.FindRows(Values)  ' find the matching row(s)
            If SomeRowViews.Length = 1 Then                                 ' if found 1 row
                Return SomeRowViews(0).Row                                  ' return the row
            ElseIf SomeRowViews.Length = 0 Then                             ' if found 0 rows
                Return Nothing                                              ' return nothing 
            Else                                                            ' else found multiple matching rows
                MessageBox.Show("Multiple matches found in ""FindRowUsingDataView"".", _
                                Pcm.etInternalErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return Nothing                                              ' return nothing 
            End If
        Catch ex As Exception
            ' no error message here
            Return Nothing
        Finally
            aDataView.Sort = InitSort                                       ' always reset sort
        End Try
    End Function

    Public Function FindRowUsingDataView(ByVal aTable As DataTable, _
                                         ByVal aValue As Object, _
                                         ByVal aColumnName As String) As DataRow


        ' finds the matching row using a dataview and row filter
        '
        ' vars passed:
        '   aTable - table to search
        '   aValue - value to find
        '   aColumnName - name of column to for a value to be in
        '
        ' returns: 
        '   aDataRow - 1 and only 1 filtered row was found
        '   Nothing - 
        '     No matching rows found
        '     Could not get a dataview
        '     More than 1 matching row was found
        '     something went wrong

        Try
            Dim aDataView As DataView = New DataView(aTable)
            If aDataView Is Nothing Then
                MessageBox.Show(String.Format("Could not get a DataView for table ""{0}"".", aTable.TableName), _
                                Pcm.etNoDataView, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return Nothing
            End If
            Return FindRowUsingDataView(aDataView, aValue, aColumnName)
        Catch ex As Exception
            ' no error message here
            Return Nothing
        End Try
    End Function

#End Region

#Region " FindRows "

    Public Function FindRows(ByVal aTable As DataTable, _
                             ByVal Values() As Object, _
                             ByVal ColNames() As String) As DataRow()

        ' searches a table and finds all rows with matching values in columns
        '
        ' vars passed:
        '   aTable - table to search
        '   ColNames - array of columns names to search
        '   Values - array of values to search for
        '
        ' note: ColNames and Values must be the same length.  Look for value Values(i) in column ColNames(i)
        '
        ' returns:
        '   array of matching rows
        '   Nothing -
        '     ColNames and Values not same length
        '     something went wrong

        Try
            If ColNames.Count <> Values.Count Then                  ' if list of column names and values not same length
                Return Nothing                                      ' return nothing
            End If
            Dim Filter As String = ""
            For i As Integer = 0 To ColNames.Count - 1              ' for each column
                If i > 0 Then                                       ' if not first column
                    Filter &= " and "                               ' add " and " to filter
                End If
                Filter &= FieldSelectString(ColNames(i), Values(i)) ' add filter for column and value
            Next
            Return aTable.Select(Filter)                            ' return matching rows
        Catch ex As Exception
            ' no error message here
            Return Nothing
        End Try
    End Function

#End Region

#End Region

#Region " MergeRow "

    Public Sub CopyRow(FromRow As DataRow, ToRow As DataRow)

        Try
            For i As Integer = 0 To FromRow.Table.Columns.Count - 1
                ToRow(i) = FromRow(i)
            Next
        Catch ex As Exception
            Return
        End Try
    End Sub

    Public Function MergeRow(RowToBeMerged As DataRow, _
                             Optional MergeNullColumns As Boolean = False, _
                             Optional AcceptChanges As Boolean = True) As Integer

        ' merges that values in RowToBeMerged into the matching row in RowToBeMerged.Table
        ' requirements to use:
        '   1) RowToBeMerged.Table must be keyed.
        '   2) Row matching RowToBeMerged key values must already be in RowToBeMerged.Table
        '
        ' vars passed:
        '   RowToBeMerged - row with data to be merged
        '   MergeNullColumns - 
        '     TRUE: columns with null values in RowsToBeMerged will have the null value merged into the table
        '     FALSE: columns with null values in RowsToBeMerged will be skipped
        '   AcceptChanges - 
        '     TRUE - RowToBeMerged.Table.AcceptChanges will be called after the merge, so there are no pending changes to the table
        '     FALSE - RowToBeMerged.Table.AcceptChanges is not called after the merge.  Use this if there are multiple rows to be 
        '             merged, and manually call RowToBeMerged.Table.AcceptChanges after the last row is merged.
        '
        ' returns:
        '   0 - could not find matching row in RowToBeMerged.Table, need to add RowToBeMerged to table
        '   1 - 1 row merged
        '   Pcm.teNoPrimaryKey - RowToBeMerged.Table has no primary key
        '   Pcm.teOtherErr- something went wrong

        Try
            Dim MTable As DataTable = RowToBeMerged.Table                                               ' get the table to merge into
            If MTable.PrimaryKey.Length = 0 Then                                                        ' if no primary key
                Return Pcm.teNoPrimaryKey                                                               ' return no primary key error
            End If
            Dim Keys(MTable.PrimaryKey.Length - 1) As Object                                            ' array of key values for Row To Be Merged
            Dim k As Integer = 0                                                                        ' start at key array index 0
            For Each KeyCol As DataColumn In MTable.PrimaryKey                                          ' for each key column 
                Keys(k) = RowToBeMerged(KeyCol)                                                         ' populate key values array 
                k += 1                                                                                  ' increment key values index
            Next
            Dim MergeIntoRow As DataRow = MTable.Rows.Find(Keys)                                        ' find row to be merged into
            If MergeIntoRow Is Nothing Then                                                             ' if could not find row to be merges into
                Return 0                                                                                ' return no matching row, no rows merged
            End If
            For Each col As DataColumn In MTable.Columns                                                ' for each column in RowToBeMerged
                If RowToBeMerged.IsNull(col) Then                                                       ' if to be merged value is null
                    If MergeNullColumns Then                                                            ' if ok to merge null values
                        MergeIntoRow(col) = RowToBeMerged(col)                                          ' then set merge into row value as null
                    End If
                Else                                                                                    ' else to be merged value is not null
                    If MergeIntoRow.IsNull(col) OrElse MergeIntoRow(col) IsNot RowToBeMerged(col) Then  ' if no MergeInto value or values no match
                        MergeIntoRow(col) = RowToBeMerged(col)                                          ' set merge into row value 
                    End If
                End If
            Next
            If AcceptChanges Then                                                                       ' if want to accept changes now
                MTable.AcceptChanges()                                                                  ' then accept changes now
            End If
            Return 1                ' if got here, then return 1 row merged
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("PcdbModule.MergeRow.  " & ex.Message)
            Return Pcm.teOtherErr   ' return other error
        End Try
    End Function

#End Region

#Region " Delete Rows "

    Public Function DeleteSelectedRows(ByVal aTable As DataTable, _
                                       ByVal ColumnName As String, _
                                       ByVal IdValue As Integer) As Integer

        ' selects rows from a table, and then deletes them
        '
        ' vars passed:
        '   aTable - table to select rows from
        '   ColumnName - name of column with value to match (all rows with matching value in this column will be deleted)
        '   IdValue - value to match in ColumnName
        '
        ' returns:
        '   >= 0 - # of rows deleted
        '   Pcm.teDeleteErr - something went wrong

        Try
            Dim SelectStr As String = PcdbModule.FieldSelectString(ColumnName, IdValue, , False)    ' get selection string
            Dim RowsToDelete() As DataRow = aTable.Select(SelectStr)                                ' get rows for id value in table
            For Each aRow As DataRow In RowsToDelete                                                ' for each row to delete

                ' calling aRow().Delete does not delete the row here.  It just changes the .RowState to Deleted.
                ' when AcceptChanges is called, the row is removed from the table. 
                ' OK to use For..Each here because using it in an Array of DataRow, not a DataRowCollection

                aRow.Delete()                                                                       ' delete it
            Next
            Return RowsToDelete.Length                                                              ' return # of rows to be deleted
        Catch ex As Exception
            ShowOtherErrorMessage("PcdbModule.DeleteRows(aTable, ColumnName, IdValue).  " & ex.Message)
            Return Pcm.teDeleteErr
        End Try
    End Function

#End Region

#Region " ColumnSummary "

    Public Function ColumnSummary(ByVal aGridView As Views.Grid.GridView, _
                                  ByVal aFieldName As String) As Object

        ' tries to get the summary total for a column in the fixed fee grid
        ' 
        ' vars passed:
        '   aGridView - DevExpress grid view to use
        '   aFieldName - field name of column for summary

        Try
            If aGridView.RowCount > 0 Then
                Dim SumCol As DevExpress.XtraGrid.Columns.GridColumn = aGridView.Columns(aFieldName)
                Return SumCol.SummaryItem.SummaryValue
            Else
                Return Nothing
            End If
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("PcdbModule.ColumnSummary.  " & ex.Message)
            Return Nothing
        End Try

    End Function

#End Region

#Region " SafeClear "

    Public Function SafeClear(ByVal aTable As DataTable) As Integer

        ' clears table, but checks if any child table has data.  If child table has data, then
        ' recursively clears child table data
        '
        ' note: child tables must have ForeignKeyConstraints (ChildKeyConstraint IsNot Nothing).  
        '   If child table has just a Relation (ChildKeyConstraint Is Nothing) then the child table is not
        '   cleared.
        '
        ' vars passed:
        '   aTable - table to clear
        '
        ' returns:
        '   Pcm.NoErrors - no errors
        '   Pcm.teCouldNotClearErr - could not clear
        '   Pcm.teOtherErr - other error

        Try
            If aTable Is Nothing Then
                Return Pcm.teNoTableErr
            End If
            Dim aErr As Integer = Pcm.NoErrors
            If aTable.ChildRelations.Count > 0 Then                             ' if got child tables     
                Dim cr As Integer
                For cr = 0 To aTable.ChildRelations.Count - 1                   ' for each child relation
                    If aTable.ChildRelations(cr).ChildKeyConstraint IsNot Nothing Then ' if a foreign key constraint
                        aErr = SafeClear(aTable.ChildRelations(cr).ChildTable()) ' clear all child table recursively
                        If aErr <> Pcm.NoErrors Then                            ' if got an error
                            Return aErr                                         ' return the error
                        End If
                    End If
                Next
            End If
            aTable.Clear()                                      ' clear data from aTable
            Return Pcm.NoErrors
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("PcdbModule.SafeClear(aTable).  " & ex.Message)
            Return Pcm.teOtherErr
        End Try
    End Function

#End Region

#Region " Force To Post "

    Public Sub ForceLuEditsToPost(ByVal ActiveDevExpForm As XtraForm, _
                                  ByVal TempFocusCtrl As Control)

        ' forces a lookup edit to post it changes to the underlying data, by focusing another control,
        ' then refocusing the lookup edit
        '
        ' vars passed:
        '   ActiveDevExpForm - DevExp form that is(active)
        '   TempFocusCtrl - control to temp focus

        Try
            Dim ActCtrl As Control = PcdbModule.MaskBoxConvert(ActiveDevExpForm.ActiveControl)  ' get active control
            If TypeOf ActCtrl Is LookUpEdit Then     ' if active control is lookup edit
                TempFocusCtrl.Focus()
                ActCtrl.Focus()
            End If
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("PcdbModule.ForceLuEditsToPost.  " & ex.Message)
        End Try
    End Sub

    Public Sub ForceRowToPost(ByVal aGridControl As GridControl)

        ' forces a row in a grid control to post

        Dim aBaseView As DevExpress.XtraGrid.Views.Base.BaseView
        Try                                                 ' get the focused grid control
            aBaseView = aGridControl.MainView               ' get the main view of the grid
            If Not aBaseView Is Nothing Then                ' if got a main view
                aBaseView.CloseEditor()                     ' close the editor
                aBaseView.UpdateCurrentRow()                ' force row to post
            End If
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("PcdbModule.ForceRowToPost.  " & ex.Message)
        End Try
    End Sub

#End Region

#Region " GfrcvAs - GetFocusedRowCellValue as ... "

    '' there are six GfrcvAs funcs, with the NoValue Param and returning
    ''   Boolean
    ''   DateTime
    ''   Decimal
    ''   Double
    ''   Integer
    ''   String
    '' no error message in Catch part of try..catch

    'Public Function GfrcvAs(ByVal aGridView As Views.Grid.GridView, _
    '                        ByVal aFieldName As String, _
    '                        ByVal NoValue As Boolean) As Boolean

    '    ' uses the gridview's GetFocusedRowCellValue method to return the value of the 
    '    ' cell in a field for the focused row.  If there is no value in that cell, or 
    '    ' there is an error, then the NoValue is returned.
    '    '
    '    ' vars passed:
    '    '   aGridView - the gridview to use
    '    '   aFieldName - the name of the field with cell to get value from
    '    '   NoValue - the value to return if the cell is DBNull or there is an error 

    '    Try
    '        If TypeOf aGridView.GetFocusedRowCellValue(aFieldName) Is DBNull Then
    '            Return NoValue
    '        Else
    '            Return CBool(aGridView.GetFocusedRowCellValue(aFieldName))
    '        End If
    '    Catch ex As Exception
    '        Return NoValue
    '    End Try
    'End Function

    'Public Function GfrcvAs(ByVal aGridView As Views.Grid.GridView, _
    '                        ByVal aFieldName As String, _
    '                        ByVal NoValue As DateTime) As DateTime

    '    Try
    '        If TypeOf aGridView.GetFocusedRowCellValue(aFieldName) Is DBNull Then
    '            Return NoValue
    '        Else
    '            Return CType(aGridView.GetFocusedRowCellValue(aFieldName), DateTime)
    '        End If
    '    Catch ex As Exception
    '        Return NoValue
    '    End Try
    'End Function

    'Public Function GfrcvAs(ByVal aGridView As Views.Grid.GridView, _
    '                        ByVal aFieldName As String, _
    '                        ByVal NoValue As Decimal, _
    '                        Optional ByVal RoundDigits As Integer = 2) As Decimal

    '    Try
    '        If TypeOf aGridView.GetFocusedRowCellValue(aFieldName) Is DBNull Then
    '            Return NoValue
    '        Else
    '            Dim aDecimalValue As Decimal = CDec(aGridView.GetFocusedRowCellValue(aFieldName))   ' get decimal value
    '            If RoundDigits >= 0 AndAlso RoundDigits <= 28 Then                                  ' if round is within range
    '                aDecimalValue = Decimal.Round(aDecimalValue, RoundDigits)                       ' round the value as desired
    '            End If
    '            Return aDecimalValue
    '        End If
    '    Catch ex As Exception
    '        Return NoValue
    '    End Try
    'End Function

    'Public Function GfrcvAs(ByVal aGridView As Views.Grid.GridView, _
    '                        ByVal aFieldName As String, _
    '                        ByVal NoValue As Double) As Double

    '    Try
    '        If TypeOf aGridView.GetFocusedRowCellValue(aFieldName) Is DBNull Then
    '            Return NoValue
    '        Else
    '            Return CDbl(aGridView.GetFocusedRowCellValue(aFieldName))
    '        End If
    '    Catch ex As Exception
    '        Return NoValue
    '    End Try
    'End Function

    'Public Function GfrcvAs(ByVal aGridView As Views.Grid.GridView, _
    '                        ByVal aFieldName As String, _
    '                        ByVal NoValue As Integer) As Integer

    '    Try
    '        If TypeOf aGridView.GetFocusedRowCellValue(aFieldName) Is DBNull Then
    '            Return NoValue
    '        Else
    '            Return CInt(aGridView.GetFocusedRowCellValue(aFieldName))
    '        End If
    '    Catch ex As Exception
    '        Return NoValue
    '    End Try
    'End Function

    'Public Function GfrcvAs(ByVal aGridView As Views.Grid.GridView, _
    '                        ByVal aFieldName As String, _
    '                        ByVal NoValue As String) As String

    '    Try
    '        If TypeOf aGridView.GetFocusedRowCellValue(aFieldName) Is DBNull Then
    '            Return NoValue
    '        Else
    '            Return CStr(aGridView.GetFocusedRowCellValue(aFieldName))
    '        End If
    '    Catch ex As Exception
    '        Return NoValue
    '    End Try
    'End Function

#End Region

#Region " GrcvAs - GetRowCellValue as ... "

    '' there are six GrcvAs funcs, with the NoValue Param and returning
    ''   Boolean
    ''   DateTime
    ''   Decimal
    ''   Double
    ''   Integer
    ''   String
    '' no error message in Catch part of try..catch

    'Public Function GrcvAs(ByVal aGridView As Views.Grid.GridView, _
    '                       ByVal aRowHandle As Integer, _
    '                       ByVal aFieldName As String, _
    '                       ByVal NoValue As Boolean) As Boolean

    '    ' uses the gridview's GetRowCellValue method to return the value of the 
    '    ' cell in a field for the desired row.  If there is no value in that cell, or 
    '    ' there is an error, then the NoValue is returned.
    '    '
    '    ' vars passed:
    '    '   aGridView - the gridview to use
    '    '   aRowHandle - the row handle for the desired row in the grid
    '    '   aFieldName - the name of the field with cell to get value from
    '    '   NoValue - the value to return if the cell is DBNull or there is an error 

    '    Try
    '        If TypeOf aGridView.GetRowCellValue(aRowHandle, aFieldName) Is DBNull Then
    '            Return NoValue
    '        Else
    '            Return CBool(aGridView.GetRowCellValue(aRowHandle, aFieldName))
    '        End If
    '    Catch ex As Exception
    '        Return NoValue
    '    End Try
    'End Function

    'Public Function GrcvAs(ByVal aGridView As Views.Grid.GridView, _
    '                       ByVal aRowHandle As Integer, _
    '                       ByVal aFieldName As String, _
    '                       ByVal NoValue As DateTime) As DateTime

    '    Try
    '        If TypeOf aGridView.GetRowCellValue(aRowHandle, aFieldName) Is DBNull Then
    '            Return NoValue
    '        Else
    '            Return CType(aGridView.GetRowCellValue(aRowHandle, aFieldName), DateTime)
    '        End If
    '    Catch ex As Exception
    '        Return NoValue
    '    End Try
    'End Function

    'Public Function GrcvAs(ByVal aGridView As Views.Grid.GridView, _
    '                       ByVal aRowHandle As Integer, _
    '                       ByVal aFieldName As String, _
    '                       ByVal NoValue As Decimal, _
    '                       Optional ByVal RoundDigits As Integer = 2) As Decimal

    '    Try
    '        If TypeOf aGridView.GetRowCellValue(aRowHandle, aFieldName) Is DBNull Then
    '            Return NoValue
    '        Else
    '            Dim aDecimalValue As Decimal = CDec(aGridView.GetRowCellValue(aRowHandle, aFieldName))  ' get decimal value
    '            If RoundDigits >= 0 AndAlso RoundDigits <= 28 Then                                      ' if round is within range
    '                aDecimalValue = Decimal.Round(aDecimalValue, RoundDigits)                           ' round the value as desired
    '            End If
    '            Return aDecimalValue
    '        End If
    '    Catch ex As Exception
    '        Return NoValue
    '    End Try
    'End Function

    'Public Function GrcvAs(ByVal aGridView As Views.Grid.GridView, _
    '                       ByVal aRowHandle As Integer, _
    '                       ByVal aFieldName As String, _
    '                       ByVal NoValue As Double) As Double

    '    Try
    '        If TypeOf aGridView.GetRowCellValue(aRowHandle, aFieldName) Is DBNull Then
    '            Return NoValue
    '        Else
    '            Return CDbl(aGridView.GetRowCellValue(aRowHandle, aFieldName))
    '        End If
    '    Catch ex As Exception
    '        Return NoValue
    '    End Try
    'End Function

    'Public Function GrcvAs(ByVal aGridView As Views.Grid.GridView, _
    '                       ByVal aRowHandle As Integer, _
    '                       ByVal aFieldName As String, _
    '                       ByVal NoValue As Integer) As Integer

    '    Try
    '        If TypeOf aGridView.GetRowCellValue(aRowHandle, aFieldName) Is DBNull Then
    '            Return NoValue
    '        Else
    '            Return CInt(aGridView.GetRowCellValue(aRowHandle, aFieldName))
    '        End If
    '    Catch ex As Exception
    '        Return NoValue
    '    End Try
    'End Function

    'Public Function GrcvAs(ByVal aGridView As Views.Grid.GridView, _
    '                       ByVal aRowHandle As Integer, _
    '                       ByVal aFieldName As String, _
    '                       ByVal NoValue As String) As String

    '    Try
    '        If TypeOf aGridView.GetRowCellValue(aRowHandle, aFieldName) Is DBNull Then
    '            Return NoValue
    '        Else
    '            Return CStr(aGridView.GetRowCellValue(aRowHandle, aFieldName))
    '        End If
    '    Catch ex As Exception
    '        Return NoValue
    '    End Try
    'End Function

#End Region

#Region " GcrcvAs - GetCurrentRowCellValue as ... "

    'Public Function GetCurrentDataRow(ByVal aBm As BindingManagerBase) As DataRow

    '    ' gets the current data row for a binding manager base
    '    ' 
    '    ' vars passed:
    '    '   aBm - the binding manager base, bound to a datasource/datatable
    '    '
    '    ' returns:
    '    '   a data row from the data table aBm is bound to
    '    '   nothing - an error occurred when trying to get the data row

    '    Try
    '        If aBm Is Nothing OrElse aBm.Count = 0 Then             ' if no binding manager or no data in binding manager
    '            Return Nothing                                      ' return nothing
    '        Else                                                    ' else got data 
    '            Return CType(aBm.Current, DataRowView).Row          ' return current data row
    '        End If
    '    Catch ex As Exception
    '        ' no error message here
    '        Return Nothing
    '    End Try
    'End Function

    '' there are six GcrcvAs funcs using a BindingManagerBase to get the row data 
    '' with the NoValue Param and returning
    ''   Boolean
    ''   DateTime
    ''   Decimal
    ''   Double
    ''   Integer
    ''   String
    '' no error message in Catch part of try..catch

    'Public Function GcrcvAs(ByVal aBmb As BindingManagerBase, _
    '                        ByVal aColumnName As String, _
    '                        ByVal NoValue As Boolean) As Boolean

    '    ' uses the BindingManagerBase.Current method to return the DataRowView 
    '    ' of the current row of the table the BindingManagerBase is bound to.  Then, 
    '    ' get row from DataRowView, and then get the value in the desired column
    '    ' for that row.  If there is no value in that column, or there is an error, 
    '    ' then the NoValue is returned.
    '    '
    '    ' vars passed:
    '    '   aBm - the binding manager base to use
    '    '   aColumnName - the name of the column with value to get 
    '    '   NoValue - the value to return if 
    '    '       there are no rows
    '    '       the column is DBNull 
    '    '       there is an error 
    '    '
    '    ' returns:
    '    '   value in the desired column of current row
    '    '   NoValue - 
    '    '       there are no rows
    '    '       the column is DBNull 
    '    '       there is an error 

    '    Try
    '        If aBmb.Count = 0 Then              ' if no rows
    '            Return NoValue                  ' return NoValue
    '        End If
    '        Dim aRow As DataRow = GetCurrentDataRow(aBmb)  ' get datarow for current item
    '        If aRow.IsNull(aColumnName) Then    ' if no value in column
    '            Return NoValue                  ' return NoValue
    '        Else                                ' else got a value
    '            Return CBool(aRow(aColumnName)) ' return the value in the column
    '        End If
    '    Catch ex As Exception                   ' if have an error
    '        Return NoValue                      ' return no value
    '    End Try
    'End Function

    'Public Function GcrcvAs(ByVal aBmb As BindingManagerBase, _
    '                        ByVal aColumnName As String, _
    '                        ByVal NoValue As DateTime) As DateTime

    '    Try
    '        If aBmb.Count = 0 Then
    '            Return NoValue
    '        End If
    '        Dim aRow As DataRow = GetCurrentDataRow(aBmb)
    '        If aRow.IsNull(aColumnName) Then
    '            Return NoValue
    '        Else
    '            Return CType(aRow(aColumnName), DateTime)
    '        End If
    '    Catch ex As Exception
    '        Return NoValue
    '    End Try
    'End Function

    'Public Function GcrcvAs(ByVal aBmb As BindingManagerBase, _
    '                        ByVal aColumnName As String, _
    '                        ByVal NoValue As Decimal, _
    '                        Optional ByVal RoundDigits As Integer = 2) As Decimal

    '    Try
    '        If aBmb.Count = 0 Then
    '            Return NoValue
    '        End If
    '        Dim aRow As DataRow = GetCurrentDataRow(aBmb)
    '        If aRow.IsNull(aColumnName) Then
    '            Return NoValue
    '        Else
    '            Dim aDecimalValue As Decimal = CDec(aRow(aColumnName))          ' get decimal value
    '            If RoundDigits >= 0 AndAlso RoundDigits <= 28 Then              ' if round is within range
    '                aDecimalValue = Decimal.Round(aDecimalValue, RoundDigits)   ' round the value as desired
    '            End If
    '            Return aDecimalValue
    '        End If
    '    Catch ex As Exception
    '        Return NoValue
    '    End Try
    'End Function

    'Public Function GcrcvAs(ByVal aBmb As BindingManagerBase, _
    '                        ByVal aColumnName As String, _
    '                        ByVal NoValue As Double) As Double

    '    Try
    '        If aBmb.Count = 0 Then
    '            Return NoValue
    '        End If
    '        Dim aRow As DataRow = GetCurrentDataRow(aBmb)
    '        If aRow.IsNull(aColumnName) Then
    '            Return NoValue
    '        Else
    '            Return CDbl(aRow(aColumnName))
    '        End If
    '    Catch ex As Exception
    '        Return NoValue
    '    End Try
    'End Function

    'Public Function GcrcvAs(ByVal aBmb As BindingManagerBase, _
    '                        ByVal aColumnName As String, _
    '                        ByVal NoValue As Integer) As Integer

    '    Try
    '        If aBmb.Count = 0 Then
    '            Return NoValue
    '        End If
    '        Dim aRow As DataRow = GetCurrentDataRow(aBmb)
    '        If aRow.IsNull(aColumnName) Then
    '            Return NoValue
    '        Else
    '            Return CInt(aRow(aColumnName))
    '        End If
    '    Catch ex As Exception
    '        Return NoValue
    '    End Try
    'End Function

    'Public Function GcrcvAs(ByVal aBmb As BindingManagerBase, _
    '                        ByVal aColumnName As String, _
    '                        ByVal NoValue As String) As String

    '    Try
    '        If aBmb.Count = 0 Then
    '            Return NoValue
    '        End If
    '        Dim aRow As DataRow = GetCurrentDataRow(aBmb)
    '        If aRow.IsNull(aColumnName) Then
    '            Return NoValue
    '        Else
    '            Return CStr(aRow(aColumnName))
    '        End If
    '    Catch ex As Exception
    '        Return NoValue
    '    End Try
    'End Function

    '' there are six GcrcvAs funcs using a row data 
    '' with the NoValue Param and returning
    ''   Boolean
    ''   DateTime
    ''   Decimal
    ''   Double
    ''   Integer
    ''   String

    'Public Function GcrcvAs(ByVal aRow As DataRow, _
    '                        ByVal aColumnName As String, _
    '                        ByVal NoValue As Boolean) As Boolean

    '    ' get the value in the desired column for a row.  If there is no value 
    '    ' in that column, or there is an error, then the NoValue is returned.
    '    '
    '    ' vars passed:
    '    '   aRow - the data row to use
    '    '   aColumnName - the name of the column with value to get 
    '    '   NoValue - the value to return if 
    '    '       there are no rows
    '    '       the column is DBNull 
    '    '       there is an error 
    '    '
    '    ' returns:
    '    '   value in the desired column of current row
    '    '   NoValue - 
    '    '       there are no rows
    '    '       the column is DBNull 
    '    '       there is an error 

    '    Try
    '        If aRow.IsNull(aColumnName) Then    ' if no value in column
    '            Return NoValue                  ' return NoValue
    '        Else                                ' else got a value
    '            Return CBool(aRow(aColumnName)) ' return the value in the column
    '        End If
    '    Catch ex As Exception                   ' if have an error
    '        Return NoValue                      ' return no value
    '    End Try
    'End Function

    'Public Function GcrcvAs(ByVal aRow As DataRow, _
    '                        ByVal aColumnName As String, _
    '                        ByVal NoValue As DateTime) As DateTime

    '    Try
    '        If aRow.IsNull(aColumnName) Then
    '            Return NoValue
    '        Else
    '            Return CType(aRow(aColumnName), DateTime)
    '        End If
    '    Catch ex As Exception
    '        Return NoValue
    '    End Try
    'End Function

    'Public Function GcrcvAs(ByVal aRow As DataRow, _
    '                        ByVal aColumnName As String, _
    '                        ByVal NoValue As Decimal, _
    '                        Optional ByVal RoundDigits As Integer = 2) As Decimal

    '    ' RoundDigits - rounds the decimal value to this many digits

    '    Try
    '        If aRow.IsNull(aColumnName) Then
    '            Return NoValue
    '        Else
    '            Dim aDecimalValue As Decimal = CDec(aRow(aColumnName))
    '            If RoundDigits >= 0 AndAlso RoundDigits <= 28 Then              ' if round is within range
    '                aDecimalValue = Decimal.Round(aDecimalValue, RoundDigits)   ' round the value as desired
    '            End If
    '            Return aDecimalValue
    '        End If
    '    Catch ex As Exception
    '        Return NoValue
    '    End Try
    'End Function

    'Public Function GcrcvAs(ByVal aRow As DataRow, _
    '                        ByVal aColumnName As String, _
    '                        ByVal NoValue As Double) As Double

    '    Try
    '        If aRow.IsNull(aColumnName) Then
    '            Return NoValue
    '        Else
    '            Return CDbl(aRow(aColumnName))
    '        End If
    '    Catch ex As Exception
    '        Return NoValue
    '    End Try
    'End Function

    'Public Function GcrcvAs(ByVal aRow As DataRow, _
    '                        ByVal aColumnName As String, _
    '                        ByVal NoValue As Integer) As Integer

    '    Try
    '        If aRow.IsNull(aColumnName) Then
    '            Return NoValue
    '        Else
    '            Return CInt(aRow(aColumnName))
    '        End If
    '    Catch ex As Exception
    '        Return NoValue
    '    End Try
    'End Function

    'Public Function GcrcvAs(ByVal aRow As DataRow, _
    '                        ByVal aColumnName As String, _
    '                        ByVal NoValue As String) As String

    '    Try
    '        If aRow.IsNull(aColumnName) Then
    '            Return NoValue
    '        Else
    '            Return CStr(aRow(aColumnName))
    '        End If
    '    Catch ex As Exception
    '        Return NoValue
    '    End Try
    'End Function

#End Region

#Region " GetValueAs.. "

    '' there are six GetValueAs funcs using to convert an object to a data value
    ''   Boolean
    ''   DateTime
    ''   Decimal
    ''   Double
    ''   Integer
    ''   String

    'Public Function GetValueAs(Value As Object, _
    '                           NoValue As Boolean) As Boolean

    '    ' get the value.  If there is no value, or there is an error, then the NoValue is returned.
    '    '
    '    ' vars passed:
    '    '   Value - the object to convert
    '    '   NoValue - the value to return if 
    '    '       there are no rows
    '    '       the column is DBNull 
    '    '       there is an error 
    '    '
    '    ' returns:
    '    '   value in the desired column of current row
    '    '   NoValue - 
    '    '       there are no rows
    '    '       the column is DBNull 
    '    '       there is an error 

    '    Try
    '        If Value Is Nothing OrElse Value Is DBNull.Value Then
    '            Return NoValue
    '        Else
    '            Return CType(Value, Boolean)
    '        End If
    '    Catch ex As Exception
    '        Return NoValue
    '    End Try
    'End Function

    'Public Function GetValueAs(Value As Object, _
    '                           NoValue As DateTime) As DateTime

    '    Try
    '        If Value Is Nothing OrElse Value Is DBNull.Value Then
    '            Return NoValue
    '        Else
    '            Return CType(Value, DateTime)
    '        End If
    '    Catch ex As Exception
    '        Return NoValue
    '    End Try
    'End Function

    'Public Function GetValueAs(Value As Object, _
    '                           NoValue As Decimal, _
    '                           Optional ByVal RoundDigits As Integer = 2) As Decimal

    '    Try
    '        If Value Is Nothing OrElse Value Is DBNull.Value Then
    '            Return NoValue
    '        Else
    '            Dim aDecimalValue As Decimal = CType(Value, Decimal)            ' get decimal value
    '            If RoundDigits >= 0 AndAlso RoundDigits <= 28 Then              ' if round is within range
    '                aDecimalValue = Decimal.Round(aDecimalValue, RoundDigits)   ' round the value as desired
    '            End If
    '            Return aDecimalValue
    '        End If
    '    Catch ex As Exception
    '        Return NoValue
    '    End Try
    'End Function

    'Public Function GetValueAs(Value As Object, _
    '                           NoValue As Double) As Double

    '    Try
    '        If Value Is Nothing OrElse Value Is DBNull.Value Then
    '            Return NoValue
    '        Else
    '            Return CType(Value, Double)
    '        End If
    '    Catch ex As Exception
    '        Return NoValue
    '    End Try
    'End Function

    'Public Function GetValueAs(Value As Object, _
    '                           NoValue As Integer) As Integer

    '    Try
    '        If Value Is Nothing OrElse Value Is DBNull.Value Then
    '            Return NoValue
    '        Else
    '            Return CType(Value, Integer)
    '        End If
    '    Catch ex As Exception
    '        Return NoValue
    '    End Try
    'End Function

    'Public Function GetValueAs(Value As Object, _
    '                           NoValue As String) As String

    '    Try
    '        If Value Is Nothing OrElse Value Is DBNull.Value Then
    '            Return NoValue
    '        Else
    '            Return CType(Value, String)
    '        End If
    '    Catch ex As Exception
    '        Return NoValue
    '    End Try
    'End Function

#End Region

#Region " Get____Value "

    ' no error message in Catch part of try..catch for Get___Value functions

    Public Function GetIntegerValue(ByVal aBaseEdit As BaseEdit) As Integer

        ' gets the value as a integer
        '
        ' vars passed:
        '   aBaseEdit - dev express edit control with a value
        ' 
        ' returns:
        '   the value as a integer
        '   0 - if got a error

        Try
            If aBaseEdit.Text = "" Then                     ' if no text
                Return 0                                    ' return 0
            End If
            Return CType(aBaseEdit.EditValue, Integer)
        Catch ex As Exception
            Return 0
        End Try
    End Function

    Public Function GetDecimalValue(ByVal aBaseEdit As BaseEdit, _
                                    Optional ByVal RoundDigits As Integer = Pcm.RoundToCents) As Decimal

        ' gets the value as a decimal
        '
        ' vars passed:
        '   aBaseEdit - dev express edit control with a value
        '   RoundDigits - round value to digits
        ' 
        ' returns:
        '   the value as a decimal
        '   0 - if got a error

        Try
            If aBaseEdit.Text = "" Then                                         ' if no text
                Return Pcm.ZeroMoney                                            ' return $0
            End If
            Dim aDecimalValue As Decimal = CType(aBaseEdit.EditValue, Decimal)  ' get decimal value
            If RoundDigits >= 0 AndAlso RoundDigits <= 28 Then                  ' if round is within range
                aDecimalValue = Decimal.Round(aDecimalValue, RoundDigits)       ' round the value as desired
            End If
            Return aDecimalValue

        Catch ex As Exception
            Return Pcm.ZeroMoney
        End Try
    End Function

    Public Function GetDoubleValue(ByVal aBaseEdit As BaseEdit) As Double

        ' gets the value as a double
        '
        ' vars passed:
        '   aBaseEdit - dev express edit control with a value
        ' 
        ' returns:
        '   the value as a double
        '   0 - if got a error

        Try
            If aBaseEdit.Text = "" Then                     ' if no text
                Return 0.0                                  ' return 0.0
            End If
            Return CType(aBaseEdit.EditValue, Double)
        Catch ex As Exception
            Return 0.0
        End Try
    End Function

#End Region

#Region " RowCountToBeUpdated "

    'Private Sub RowCountToBeUpdatedDebugging(rcDataView As DataView)

    '    Debug.WriteLine(String.Format("DataView for Table: {0}, RowStateFilter: {1}", rcDataView.Table.TableName, rcDataView.RowStateFilter.ToString))
    '    Debug.WriteLine(String.Format("DataView Count: {0}", rcDataView.Count))
    '    For i As Integer = 0 To rcDataView.Count - 1
    '        Debug.WriteLine(String.Format("DataView(0): {0}, KeyField: {1}, Key: {2}", i, rcDataView.Table.PrimaryKey(0).ColumnName, rcDataView(i)(rcDataView.Table.PrimaryKey(0).ColumnName)))
    '    Next
    'End Sub

    Public Function RowCountToBeUpdated(Table As DataTable) As Integer

        ' calculates how many rows will be updated. count # of added, deleted and modified rows
        '
        ' vars passed:
        '   Table - data table to check
        '
        ' returns
        '   >= 0 - # of rows that will be updated
        '   Pcm.teNoTableErr - no data table
        '   Pcm.OtherExErr - something went wrong

        Try
            If Table Is Nothing Then                                        ' if no data table
                Return Pcm.teNoTableErr                                     ' return error
            End If
            Dim rcDataView As New DataView(Table) _
                With {.RowStateFilter = DataViewRowState.Added}             ' create data view with rowState filter as added
            Dim RowCount As Integer = rcDataView.Count                      ' get # of rows to be added

            'RowCountToBeUpdatedDebugging(rcDataView)

            rcDataView.RowStateFilter = DataViewRowState.Deleted            ' filter data view by deleted rows
            RowCount += rcDataView.Count                                    ' add in # of rows to be deleted

            'RowCountToBeUpdatedDebugging(rcDataView)

            rcDataView.RowStateFilter = DataViewRowState.ModifiedOriginal   ' filter data view by modified rows
            RowCount += rcDataView.Count                                    ' add in # of rows to be modified

            'RowCountToBeUpdatedDebugging(rcDataView)

            Return RowCount                                                 ' return # of rows to be updated
        Catch ex As Exception
            Return Pcm.OtherExErr
        End Try
    End Function

#End Region

#Region " StartExplorer "

    Public Sub StartExplorer(ByVal StartPath As String)

        ' starts windows explorer for the desired path
        '
        ' vars passed:
        '   StartPath - path for explorer to start exploring

        Const ExplorerName As String = "explorer.exe"
        Const ERROR_FILE_NOT_FOUND As Integer = 2

        'Dim aProcess As Process = New Process
        Try
            If StartPath = Nothing OrElse StartPath = "" Then
                Process.Start(ExplorerName)
            Else
                Process.Start(ExplorerName, "/e, " & StartPath)
            End If
        Catch ex As System.ComponentModel.Win32Exception
            If ex.NativeErrorCode = ERROR_FILE_NOT_FOUND Then
                MessageBox.Show(String.Format("Could not find {0}.", ExplorerName.ToUpper), _
                                "Program not found", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("PcdbModule.StartExplorer.  " & ex.Message)
        End Try
    End Sub

#End Region

#Region " Active Button Edit "

    Public Sub SetActiveButtonEditText(ByVal aButtonEdit As ButtonEdit, _
                                       ByVal aDataRow As DataRow, _
                                       ByVal aColName As String)

        ' sets ButtonEdit.Text to:
        '   "" (Column Value = DBNull)
        '   "Active" (Column Value = TRUE) 
        '   "Suspended" (Column Value = FALSE)
        '
        ' vars passed:
        '   aButtonEdit - button edit
        '   aDataRow - data row for column value
        '   aColName - name of column with active boolean value

        Const Abt_Active As String = "Active"
        Const Abt_Suspended As String = "Suspended"
        Try
            If aDataRow Is Nothing Then
                aButtonEdit.Text = ""
            Else
                ' do not use GcrcvAs here, because want to check if value is DBNull
                If TypeOf aDataRow(aColName) Is DBNull Then ' do not .IsNull 
                    aButtonEdit.Text = ""
                Else
                    If CBool(aDataRow(aColName)) Then
                        aButtonEdit.Text = Abt_Active
                    Else
                        aButtonEdit.Text = Abt_Suspended
                    End If
                End If
            End If
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("PcdbModule.SetActiveButtonEditText.  " & ex.Message)
            aButtonEdit.Text = ""
        End Try
    End Sub

#End Region

#Region " Show Error Message "

    Public Function ShowErrorMessage(ByVal Text As String,
                                     ByVal Caption As String,
                                     Optional ByVal Buttons As MessageBoxButtons = MessageBoxButtons.OK,
                                     Optional ByVal Icon As MessageBoxIcon = MessageBoxIcon.Error) As DialogResult

        ' shows an error message.  
        '
        ' vars passed:
        '   Text - text for message box
        '   Caption - caption for message box
        '   Buttons - message box buttons
        '   Icon - message box icon
        '
        ' returns:
        '   buttons pressed in the message box

        ' no try..catch here
        Return MessageBox.Show(Text, Caption, Buttons, Icon)
    End Function

    Public Function ShowOtherErrorMessage(ByVal Text As String) As DialogResult

        ' show the Other error message.  
        '   Caption is "Other Error"
        '   Button is [OK]
        '   Icon is Error
        '
        ' vars passed:
        '   Text - text for message box
        '
        ' returns:
        '   buttons pressed in the message box

        ' no try..catch here
        Return MessageBox.Show("Other error in " & Text, Pcm.etOtherErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Function

#End Region

#Region " SetSearchForEditToolTips "

    Public Sub SetSearchForEditToolTips(ByVal aGridControl As GridControl,
                                        ByVal ConciseGridView As DevExpress.XtraGrid.Views.Grid.GridView,
                                        ByVal CurrentView As DevExpress.XtraGrid.Views.Grid.GridView,
                                        ByVal ViewInFormToolTip As String,
                                        ByVal SearchFor As ButtonEdit)

        ' sets the tool tips for the search for button edit's buttons
        '
        ' vars passed:
        '   aGridControl - grid control with concise and detail view
        '   ConciseGridView - concise grid view
        '   CurrentView - aGridControl's current main view
        '   ViewInFormToolTip - tip to display al data for one row
        '   SearchFor - search for button edit.  needs to have two buttons:
        '     Index: Pcm.sfConciseIndex; button for Concise/view in form
        '     Index: Pcm.sfDetailIndex; button for Detail/view in form

        Try
            If aGridControl.Visible Then                        ' if showing grid
                If CurrentView.Name = ConciseGridView.Name Then ' if showing concise grid
                    SearchFor.Properties.Buttons(Pcm.sfConciseIndex).ToolTip = ViewInFormToolTip
                    SearchFor.Properties.Buttons(Pcm.sfDetailIndex).ToolTip = Pcm.sfDetailToolTip
                Else                                            ' else showing detail grid
                    SearchFor.Properties.Buttons(Pcm.sfConciseIndex).ToolTip = Pcm.sfConciseToolTip
                    SearchFor.Properties.Buttons(Pcm.sfDetailIndex).ToolTip = ViewInFormToolTip
                End If
            Else                                                ' else showing data in form
                SearchFor.Properties.Buttons(Pcm.sfConciseIndex).ToolTip = Pcm.sfConciseToolTip
                SearchFor.Properties.Buttons(Pcm.sfDetailIndex).ToolTip = Pcm.sfDetailToolTip
            End If
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("PcdbModule.SetSearchForEditToolTips.  " & ex.Message)
        End Try
    End Sub

#End Region

#Region " DataSet "

    Public Function AddTableToDataSet(ByVal aDataSet As DataSet, _
                                      ByVal TableToAddName As String, _
                                      ByVal BaseTableName As String) As Integer

        ' adds a table to a dataset
        '   
        ' vars passed:
        '   aDataSet - dataset to add table to
        '   TableToAddName - name of table to add
        '   KeyColumnName - name of key column
        '
        ' returns:
        '   Pcm.NoErrors - no errors, table added or table already in dataset
        '   Pcm.
        '   Pcm.teDataSetErr - error adding table to dataset

        Try
            If Not aDataSet.Tables.Contains(TableToAddName) Then            ' if table not already in dataset
                aDataSet.Tables.Add(TableToAddName)                         ' then add table to dataset
            End If

            Dim TableToAdd As DataTable = aDataSet.Tables(TableToAddName)   ' get the datatable after added 
            If TableToAdd Is Nothing Then                                   ' if did not get a table
                Return Pcm.teNoTableErr                                     ' return error
            End If

            If TableToAdd.Columns.Count = 0 Then                            ' if no columns in Table to add
                Dim BaseTable As DataTable = aDataSet.Tables(BaseTableName) ' get base table
                If BaseTable Is Nothing Then                                ' if did not get a table
                    Return Pcm.teNoTableErr                                 ' return error
                End If
                ' add columns
                Dim colToAdd As DataColumn
                For Each col As DataColumn In BaseTable.Columns             ' for each column in base table
                    colToAdd = New DataColumn(col.ColumnName, col.DataType) ' create new column 
                    TableToAdd.Columns.Add(colToAdd)                        ' add new column to TableToAdd
                Next

                Dim KeyLength As Integer = BaseTable.PrimaryKey.Length      ' get Base table primary key length
                If KeyLength > 0 Then                                       ' if base table has a primary key
                    Dim KeyColumns(KeyLength - 1) As DataColumn             ' create key columns array
                    Dim KeyColIndex As Integer                              ' key column's index in .Columns
                    For c As Integer = 0 To KeyLength - 1                   ' for each base table key column
                        KeyColIndex = TableToAdd.Columns.IndexOf(BaseTable.PrimaryKey(c).ColumnName) ' get col index
                        KeyColumns(c) = TableToAdd.Columns(c)               ' add to key columns array
                    Next
                    TableToAdd.PrimaryKey = KeyColumns                      ' set key columns for TableToAdd
                End If
            End If

            Return Pcm.NoErrors

        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("PcdbModule.AddTableToDataSet.  " & ex.Message)
            Return Pcm.teDataSetErr
        End Try
    End Function

    Public Function RemoveTableFromDataSet(ByVal aDataSet As DataSet, _
                                           ByVal aTableName As String) As Integer

        ' removes a table from a dataset
        '   
        ' vars passed:
        '   aDataSet - dataset to remove table from
        '   aTableName - name of table to remove
        '
        ' returns:
        '   Pcm.NoErrors - no errors, table removed or table not in dataset
        '   Pcm.teNoTableErr - no data table
        '   Pcm.teDataSetErr - error removing table to dataset

        Try
            If aDataSet.Tables.Contains(aTableName) Then                ' if table in dataset
                Dim aTable As DataTable = aDataSet.Tables(aTableName)   ' get the datatable
                If aTable Is Nothing Then                               ' if did not get a data table
                    Return Pcm.teNoTableErr                             ' return error
                End If
                If aDataSet.Tables.CanRemove(aTable) Then               ' if can remove table from dataset
                    aDataSet.Tables.Remove(aTable)                      ' then remove table
                End If
            End If
            Return Pcm.NoErrors

        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("PcdbModule.RemoveTableFromDataSet.  " & ex.Message)
            Return Pcm.teDataSetErr
        End Try
    End Function

#End Region

#Region " SelectAll Text when Click "

    Public Sub GridView_ShownEditor(ByVal sender As Object, ByVal e As EventArgs)

        ' generic ShownEditor (do not delete parameter ByVal e As EventArgs)

        Try
            ' get sender as a gridview
            Dim aGridView As DevExpress.XtraGrid.Views.Grid.GridView = TryCast(sender, DevExpress.XtraGrid.Views.Grid.GridView)
            If aGridView Is Nothing Then
                Return
            End If
            Dim aTextEdit As TextEdit = TryCast(aGridView.ActiveEditor, TextEdit)
            If aTextEdit Is Nothing Then
                Return
            End If
            AddHandler aTextEdit.Click, AddressOf TextEditSelectAllWhenClick
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("PcdbModule.GridView_ShownEditor.  " & ex.Message)
        End Try
    End Sub

    Public Sub TextEditSelectAllWhenClick(ByVal sender As Object, ByVal e As EventArgs)

        ' selects all text if the edit value is 0
        '
        ' this works for all DevExpress controls derived from the TextExit class.  Some are:
        '   ButtonEdit, CalcEdit, ComboBox, ComboBoxEdit, LookupEdit, SpinEdit

        Try
            Dim aTextEdit As TextEdit = TryCast(sender, TextEdit)
            If aTextEdit Is Nothing Then
                Return
            End If
            If (aTextEdit.EditValue Is DBNull.Value) _
                    OrElse (aTextEdit.EditValue Is Nothing) _
                    OrElse ((Not (TypeOf (aTextEdit.EditValue) Is String)) AndAlso (CInt(aTextEdit.EditValue) = 0)) Then
                aTextEdit.SelectAll()
            End If
        Catch ex As Exception
            PcdbModule.ShowOtherErrorMessage("PcdbModule.TextEditSelectAllWhenClick.  " & ex.Message)
        End Try
    End Sub

#End Region

#Region " DataMemberName "

    Public Function DataMemberName(TableName As String, ColumnName As String) As String

        ' gets the dataMember name for a column in a table
        '
        ' vars passed:
        '   TableName - name of table
        '   ColumnName - name of column
        '
        ' returns:
        '   TableName.ColumnName - value dataMember 
        '   "" - something went wrong

        Try
            Return String.Format("{0}.{1}", TableName, ColumnName)
        Catch ex As Exception
            Return ""
        End Try
    End Function
#End Region

#Region " CurrentUserAppDataRoot "

    Public Function CurrentUserAppDataRoot() As String

        ' gets the current user application data folder name, but without the trailing version #
        ' for example, My.Computer.FileSystem.SpecialDirectories.CurrentUserApplicationData returns:
        '     C:\Users\UserName\AppData\Roaming\CompanyName\ApplicationName\Version 
        '   or 
        '     C:\Users\Eric\AppData\Roaming\EWCG\PCDB\2.0.6.8
        '
        ' this function returns: 
        '     C:\Users\UserName\AppData\Roaming\CompanyName\ApplicationName
        '   or 
        '     C:\Users\Eric\AppData\Roaming\EWCG\PCDB
        '
        ' returns:
        '   My.Computer.FileSystem.SpecialDirectories.CurrentUserApplicationData without the trailing version info
        '   "" - something went wrong
        Try
            Dim FolderName As String = My.Computer.FileSystem.SpecialDirectories.CurrentUserApplicationData
            If FolderName.EndsWith(My.Application.Info.Version.ToString) Then               ' if folder name ends with version info
                Dim VersionLength As Integer = My.Application.Info.Version.ToString.Length  ' get the version info length
                Dim RemoveAt As Integer = FolderName.Length - VersionLength - 1             ' calc starting remove index -1 for slash
                FolderName = FolderName.Remove(RemoveAt, VersionLength + 1)                 ' remove version info + 1 for slash
            End If
            If Not FolderExits(FolderName) Then
                Return String.Empty
            End If
            Return FolderName
        Catch ex As Exception
            Return String.Empty
        End Try
    End Function

#End Region

#Region " DebugWriteRow "

    Public Sub DebugWriteRow(FromFunc As String, DebugRow As DataRow)

        Debug.WriteLine(String.Format("{0}Calling Sub/Func Name{1}{2}", vbCrLf, vbTab, FromFunc))
        Debug.WriteLine(String.Format("Table Name{0}{1}{0}Column Count{0}{2}", vbTab, DebugRow.Table.TableName, DebugRow.Table.Columns.Count))
        For i As Integer = 0 To DebugRow.Table.Columns.Count - 1
            Debug.WriteLine(String.Format("Column Name{0}{1}{0}Value{0}{2}", vbTab, DebugRow.Table.Columns(i).ColumnName, DebugRow(i)))
        Next
    End Sub

#End Region

End Module
