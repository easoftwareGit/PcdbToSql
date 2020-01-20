Module PcdbIo

#Region " File Dialogs "

    'Public Function GetFilter(ByVal aDesc As String, ByVal aExt As String) As String

    '    Try
    '        Return String.Format("{0}(*.{1})|*.{1}", aDesc, aExt)
    '    Catch ex As Exception
    '        Return ""
    '    End Try

    'End Function

#End Region

#Region " Folders "

    'Public Function ConvertSpecialDirectory(ByVal FolderName As String) As String

    '    ' converts a special folder name into the actual folder name
    '    '
    '    ' vars passed:
    '    '   FolderName - name of folder to convert
    '    '   
    '    ' returns:
    '    '   the actual file system folder name for the folder

    '    Try
    '        Dim SlashPos As Integer = FolderName.IndexOf("\")
    '        Dim SpecialName As String
    '        Dim AfterSpecial As String = String.Empty
    '        If SlashPos < 0 Then
    '            SpecialName = FolderName
    '        Else
    '            SpecialName = FolderName.Substring(0, SlashPos)
    '            AfterSpecial = FolderName.Substring(SlashPos, FolderName.Length - SlashPos)
    '        End If
    '        Select Case SpecialName
    '            Case Pcm.sdMyDocs
    '                Return My.Computer.FileSystem.SpecialDirectories.MyDocuments & AfterSpecial
    '            Case Pcm.sdDeskTop
    '                Return My.Computer.FileSystem.SpecialDirectories.Desktop & AfterSpecial
    '            Case Else
    '                Return FolderName
    '        End Select
    '    Catch ex As Exception
    '        Return ""
    '    End Try
    'End Function

    'Public Sub EndWithSlash(ByRef FolderName As String)

    '    ' makes sure a folder names ends with a slash
    '    '
    '    ' vars passed: 
    '    '   Folder Name: name of folder to check (passed By Ref)

    '    Try
    '        If Not FolderName.EndsWith("\") Then    ' if does not end with a "\"
    '            FolderName &= "\"                   ' add the "\"
    '        End If
    '    Catch ex As Exception
    '        Return
    '    End Try
    'End Sub

    'Public Function FolderExits(ByVal FolderName As String, _
    '                            Optional ByRef FolderLabelText As String = "", _
    '                            Optional ByRef ShowErr As Boolean = False) As Boolean

    '    ' checks to see if a folder exists on the computer
    '    '
    '    ' vars passed:
    '    '   FolderName - name of folder to find
    '    '   FolderLabelText - prompt for folder name
    '    '   ShowErr - TRUE of err message is to be displayed (if needed)
    '    '   
    '    ' returns:
    '    '   TRUE - folder is found on this computer
    '    '   FALSE - folder is NOT found on this computer

    '    Try
    '        FolderName = ConvertSpecialDirectory(FolderName)            ' convert folder name if needed
    '        If My.Computer.FileSystem.DirectoryExists(FolderName) Then  ' if found folder
    '            Return True                                             ' return TRUE
    '        Else                                                        ' else folder not found
    '            If ShowErr Then                                         ' if want to show error
    '                If FolderLabelText <> "" Then                       ' if got folder label text
    '                    FolderLabelText &= ": "                         ' add in ": " as separator
    '                End If

    '                MessageBox.Show(String.Format("{0}""{1}"" not found on this computer.", FolderLabelText, UnconvertSpecialDirectory(FolderName)), _
    '                                Pcm.etFolderNotFound, MessageBoxButtons.OK, MessageBoxIcon.Error)
    '            End If
    '            Return False                                            ' return FALSE (not found)
    '        End If
    '    Catch ex As Exception
    '        If ShowErr Then
    '            MessageBox.Show(String.Format("Other error while checking if folder ""{0}"" exists.  {1}", UnconvertSpecialDirectory(FolderName), ex.Message), _
    '                            Pcm.etOtherErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        End If
    '        Return False
    '    End Try
    'End Function

    'Public Function MakeFolder(ByVal FolderName As String, _
    '                           Optional ByRef FolderLabelText As String = "", _
    '                           Optional ByRef ShowErr As Boolean = False) As Integer

    '    ' tries to make a folder (directory) on the current computer
    '    '
    '    ' vars passed:
    '    '   FolderName - name of folder to find
    '    '   FolderLabelText - prompt for folder name
    '    '   ShowErr - TRUE of err message is to be displayed (if needed)
    '    '   
    '    ' returns:
    '    '   Pcm.NoErrors - no errors
    '    '   Pcm.ioArgumentErr - bad folder name
    '    '   Pcm.ioArgumentNullErr - blank folder name
    '    '   Pcm.ioErr - folder is read only
    '    '   Pcm.ioPathTooLongErr - folder name too long
    '    '   Pcm.ioNotSupportedErr - folder name is a colon
    '    '   Pcm.ioUnAuthErr As Integer - user does not have permission to make folder
    '    '   Pcm.ioSecurityErr As Integer - user does not have permission in partial trust
    '    '   Pcm.ioOtherErr As Integer - other error

    '    Dim MakeFolderErr As Integer = Pcm.NoErrors
    '    Try
    '        My.Computer.FileSystem.CreateDirectory(FolderName)
    '    Catch ArgEx As ArgumentException
    '        If TypeOf (ArgEx) Is ArgumentNullException Then
    '            MakeFolderErr = Pcm.ioArgumentNullErr
    '        Else
    '            MakeFolderErr = Pcm.ioArgumentErr
    '        End If
    '    Catch IoEx As IO.IOException
    '        If TypeOf (IoEx) Is IO.PathTooLongException Then
    '            MakeFolderErr = Pcm.ioPathTooLongErr
    '        Else
    '            MakeFolderErr = Pcm.ioErr
    '        End If
    '    Catch NoSupEx As NotSupportedException
    '        MakeFolderErr = Pcm.ioNotSupportedErr
    '    Catch SecEx As Security.SecurityException
    '        MakeFolderErr = Pcm.ioSecurityErr
    '    Catch UnAuthEx As UnauthorizedAccessException
    '        MakeFolderErr = Pcm.ioUnAuthErr
    '    Catch ex As Exception
    '        MakeFolderErr = Pcm.ioOtherErr
    '    Finally
    '        If MakeFolderErr <> Pcm.NoErrors AndAlso ShowErr Then
    '            Dim CouldNotMake As String = "Could not make folder"
    '            If FolderLabelText = String.Empty Then
    '                CouldNotMake &= ".  "
    '            Else
    '                CouldNotMake &= String.Format(" for {0}.  ", FolderLabelText)
    '            End If
    '            Select Case MakeFolderErr
    '                Case Pcm.ioErr
    '                    MessageBox.Show(String.Format("{0}The parent folder of ""{1}"" is read only.", CouldNotMake, FolderName), _
    '                                    Pcm.etIoErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
    '                Case Pcm.ioArgumentErr
    '                    MessageBox.Show(String.Format("{0}The folder name ""{1}"" contains illegal characters or is only white space.", CouldNotMake, FolderName), _
    '                                    Pcm.etIoErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
    '                Case Pcm.ioArgumentNullErr
    '                    MessageBox.Show(String.Format("{0}The folder name ""{1}"" is nothing.", CouldNotMake, FolderName), _
    '                                    Pcm.etIoErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
    '                Case Pcm.ioPathTooLongErr
    '                    MessageBox.Show(String.Format("{0}The folder name ""{1}"" is too long.", CouldNotMake, FolderName), _
    '                                    Pcm.etIoErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
    '                Case Pcm.ioNotSupportedErr
    '                    MessageBox.Show(String.Format("{0}The folder name ""{1}"" is a colon "":"".", CouldNotMake, FolderName), _
    '                                    Pcm.etIoErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
    '                Case Pcm.ioUnAuthErr
    '                    MessageBox.Show(String.Format("{0}The user does not have permission to create the folder ""{1}"".", CouldNotMake, FolderName), _
    '                                    Pcm.etIoErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
    '                Case Pcm.ioSecurityErr
    '                    MessageBox.Show(String.Format("{0}The user lacks permissions in a partial-trust situation to create the folder ""{1}"".", CouldNotMake, FolderName), _
    '                                    Pcm.etIoErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
    '                Case Else
    '                    MessageBox.Show(String.Format("{0}Other I/O error trying to create the folder ""{1}"".", CouldNotMake, FolderName), _
    '                                    Pcm.etIoErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
    '            End Select
    '        End If
    '    End Try
    '    Return MakeFolderErr
    'End Function

    'Public Function UnconvertSpecialDirectory(ByVal FolderName As String) As String

    '    ' unconverts a special actual folder name into the special folder name
    '    '
    '    ' vars passed:
    '    '   FolderName - name of folder to unconvert
    '    '   
    '    ' returns:
    '    '   the actual file system folder name for the folder
    '    '   the special folder name of the folder

    '    Try
    '        If FolderName.StartsWith(My.Computer.FileSystem.SpecialDirectories.Desktop) Then
    '            FolderName = Pcm.sdDeskTop & FolderName.Remove(0, My.Computer.FileSystem.SpecialDirectories.Desktop.Length)
    '        ElseIf FolderName.StartsWith(My.Computer.FileSystem.SpecialDirectories.MyDocuments) Then
    '            FolderName = Pcm.sdMyDocs & FolderName.Remove(0, My.Computer.FileSystem.SpecialDirectories.MyDocuments.Length)
    '        End If
    '    Catch ex As Exception
    '        Return String.Empty
    '    End Try
    '    Return FolderName
    'End Function

#End Region

#Region " Copy File "

    'Public Function CopyFile(ByVal FromFileName As String, ByVal ToFileName As String) As Integer

    '    ' copies a file
    '    '
    '    ' vars passed:
    '    '   FromFileName - file name to copy from
    '    '   ToFileName - file name to copy to
    '    '
    '    ' returns:
    '    '   Pcm.NoErrors - no errors
    '    '   Pcm.ioArgumentErr - bad path name
    '    '   Pcm.ioArgumentNullErr - blank file name
    '    '   Pcm.ioErr - source file points to existing folder, 
    '    '       file exists and overwrite is FALSE
    '    '       user does not have sufficient permissions to access the file  
    '    '   Pcm.ioPathTooLongErr - file name too long
    '    '   Pcm.ioNotSupportedErr - folder name is a colon, or is in invalid format
    '    '   Pcm.ioUnAuthErr As Integer - user does not have permission to make folder
    '    '   Pcm.ioSecurityErr As Integer - user does not have permission in partial trust
    '    '   Pcm.ioFileNotFoundErr - source file not found/not valid
    '    '   Pcm.ioOpCanceledErr - ShowUI is set to True, onUserCancel is set to ThrowException, user has cancelled 
    '    '       ShowUI is set to True, onUserCancel is set to ThrowException, and an unspecified I/O error occurs 
    '    '   Pcm.ioOtherErr As Integer - other error

    '    Dim CopyFileErr As Integer = Pcm.NoErrors
    '    Try
    '        My.Computer.FileSystem.CopyFile(FromFileName, ToFileName)
    '    Catch ArgEx As ArgumentException
    '        If TypeOf (ArgEx) Is ArgumentNullException Then
    '            CopyFileErr = Pcm.ioArgumentNullErr
    '        Else
    '            CopyFileErr = Pcm.ioArgumentErr
    '        End If
    '    Catch IoEx As IO.IOException
    '        If TypeOf (IoEx) Is IO.PathTooLongException Then
    '            CopyFileErr = Pcm.ioPathTooLongErr
    '        ElseIf TypeOf (IoEx) Is IO.FileNotFoundException Then
    '            CopyFileErr = Pcm.ioFileNotFoundErr
    '        Else
    '            CopyFileErr = Pcm.ioErr
    '        End If
    '    Catch NoSupEx As NotSupportedException
    '        CopyFileErr = Pcm.ioNotSupportedErr
    '    Catch OpCancelErr As OperationCanceledException
    '        CopyFileErr = Pcm.ioOpCanceledErr
    '    Catch SecEx As Security.SecurityException
    '        CopyFileErr = Pcm.ioSecurityErr
    '    Catch UnAuthEx As UnauthorizedAccessException
    '        CopyFileErr = Pcm.ioUnAuthErr
    '    Catch ex As Exception
    '        CopyFileErr = Pcm.ioOtherErr
    '    Finally
    '        If CopyFileErr <> Pcm.NoErrors Then
    '            ShowCopyErr(CopyFileErr, FromFileName, ToFileName)
    '        End If
    '    End Try
    '    Return CopyFileErr
    'End Function

    'Public Function CopyFile(ByVal FromFileName As String,
    '                         ByVal ToFileName As String,
    '                         ByVal OverWrite As Boolean) As Integer

    '    ' copies a file
    '    '
    '    ' vars passed:
    '    '   FromFileName - file name to copy from
    '    '   ToFileName - file name to copy to
    '    '   OverWrite - TRUE: do overwrite file; FALSE: do not overwrite file
    '    '
    '    ' returns:
    '    '   Pcm.NoErrors - no errors
    '    '   Pcm.ioArgumentErr - bad path name
    '    '   Pcm.ioArgumentNullErr - blank file name
    '    '   Pcm.ioErr - source file points to existing folder, 
    '    '       file exists and overwrite is FALSE
    '    '       user does not have sufficient permissions to access the file  
    '    '   Pcm.ioPathTooLongErr - file name too long
    '    '   Pcm.ioNotSupportedErr - folder name is a colon, or is in invalid format
    '    '   Pcm.ioUnAuthErr As Integer - user does not have permission to make folder
    '    '   Pcm.ioSecurityErr As Integer - user does not have permission in partial trust
    '    '   Pcm.ioFileNotFoundErr - source file not found/not valid
    '    '   Pcm.ioOpCanceledErr - ShowUI is set to True, onUserCancel is set to ThrowException, user has cancelled 
    '    '       ShowUI is set to True, onUserCancel is set to ThrowException, and an unspecified I/O error occurs 
    '    '   Pcm.ioOtherErr As Integer - other error

    '    Dim CopyFileErr As Integer = Pcm.NoErrors
    '    Try
    '        If Not OverWrite Then                                       ' if not overwriting by default
    '            If My.Computer.FileSystem.FileExists(ToFileName) Then   ' if to file exists
    '                ' have user confirm overwrite
    '                If MessageBox.Show(String.Format("File ""{0}"" already exists.   Do you want to overwrite it?", ToFileName),
    '                                   Pcm.etConfirmOverwrite, MessageBoxButtons.YesNo, MessageBoxIcon.Question,
    '                                   MessageBoxDefaultButton.Button2) <> DialogResult.Yes Then
    '                    Return Pcm.ioOpCanceledErr                      ' if user did not confirm, return cancel
    '                End If
    '            End If
    '        End If

    '        ' copy the file (always overwrite here, because if file exists and user does not want to overwrite, 
    '        ' then have already returned Pcm.ioOpCanceledErr in first part of this func

    '        My.Computer.FileSystem.CopyFile(FromFileName, ToFileName, True)
    '    Catch ArgEx As ArgumentException
    '        If TypeOf (ArgEx) Is ArgumentNullException Then
    '            CopyFileErr = Pcm.ioArgumentNullErr
    '        Else
    '            CopyFileErr = Pcm.ioArgumentErr
    '        End If
    '    Catch IoEx As IO.IOException
    '        If TypeOf (IoEx) Is IO.PathTooLongException Then
    '            CopyFileErr = Pcm.ioPathTooLongErr
    '        ElseIf TypeOf (IoEx) Is IO.FileNotFoundException Then
    '            CopyFileErr = Pcm.ioFileNotFoundErr
    '        Else
    '            CopyFileErr = Pcm.ioErr
    '        End If
    '    Catch NoSupEx As NotSupportedException
    '        CopyFileErr = Pcm.ioNotSupportedErr
    '    Catch OpCancelErr As OperationCanceledException
    '        CopyFileErr = Pcm.ioOpCanceledErr
    '    Catch SecEx As Security.SecurityException
    '        CopyFileErr = Pcm.ioSecurityErr
    '    Catch UnAuthEx As UnauthorizedAccessException
    '        CopyFileErr = Pcm.ioUnAuthErr
    '    Catch ex As Exception
    '        CopyFileErr = Pcm.ioOtherErr
    '    Finally
    '        If CopyFileErr <> Pcm.NoErrors Then
    '            ShowCopyErr(CopyFileErr, FromFileName, ToFileName)
    '        End If
    '    End Try
    '    Return CopyFileErr
    'End Function

    'Public Function CopyFile(ByVal FromFileName As String,
    '                         ByVal ToFileName As String,
    '                         ByVal ShowUI As Microsoft.VisualBasic.FileIO.UIOption,
    '                         Optional ByVal OnUserCancel As Microsoft.VisualBasic.FileIO.UICancelOption = 0) As Integer

    '    ' copies a file
    '    '
    '    ' vars passed:
    '    '   FromFileName - file name to copy from
    '    '   ToFileName - file name to copy to
    '    '   ShowUI - 
    '    '       FileIO.UIOption.AllDialogs - show progress and error dialog boxes
    '    '       FileIO.UIOption.OnlyErrorDialogs - show error dialog boxes, hide progress dialog boxes
    '    '   OnUserCancel - 
    '    '       0 - param not used
    '    '       FileIO.UICancelOption.DoNothing - do nothing when user cancels
    '    '       FileIO.UICancelOption.ThrowException - throw exception when user cancels
    '    '
    '    ' returns:
    '    '   Pcm.NoErrors - no errors
    '    '   Pcm.ioArgumentErr - bad path name
    '    '   Pcm.ioArgumentNullErr - blank file name
    '    '   Pcm.ioErr - source file points to existing folder, 
    '    '       file exists and overwrite is FALSE
    '    '       user does not have sufficient permissions to access the file  
    '    '   Pcm.ioPathTooLongErr - file name too long
    '    '   Pcm.ioNotSupportedErr - folder name is a colon, or is in invalid format
    '    '   Pcm.ioUnAuthErr As Integer - user does not have permission to make folder
    '    '   Pcm.ioSecurityErr As Integer - user does not have permission in partial trust
    '    '   Pcm.ioFileNotFoundErr - source file not found/not valid
    '    '   Pcm.ioOpCanceledErr - ShowUI is set to True, onUserCancel is set to ThrowException, user has cancelled 
    '    '       ShowUI is set to True, onUserCancel is set to ThrowException, and an unspecified I/O error occurs 
    '    '   Pcm.ioOtherErr As Integer - other error

    '    Dim CopyFileErr As Integer = Pcm.NoErrors
    '    Try
    '        ' make suer got valid ShowUi param
    '        If ShowUI <> FileIO.UIOption.AllDialogs OrElse ShowUI <> FileIO.UIOption.OnlyErrorDialogs Then
    '            CopyFileErr = Pcm.ioInvalidParamErr
    '            Exit Try
    '        End If
    '        ' make sure got valid OnUserCancel param
    '        If OnUserCancel <> 0 Then
    '            If OnUserCancel <> FileIO.UICancelOption.DoNothing OrElse OnUserCancel <> FileIO.UICancelOption.ThrowException Then
    '                CopyFileErr = Pcm.ioInvalidParamErr
    '                Exit Try
    '            End If
    '        End If

    '        If OnUserCancel = 0 Then                                                ' if no OnUserCancel
    '            My.Computer.FileSystem.CopyFile(FromFileName, ToFileName, ShowUI)   ' then copy file w/o OnUserCancel
    '        Else                                                                    ' else got OnUserCancel
    '            My.Computer.FileSystem.CopyFile(FromFileName, ToFileName, ShowUI, OnUserCancel) ' copy file 
    '        End If
    '    Catch ArgEx As ArgumentException
    '        If TypeOf (ArgEx) Is ArgumentNullException Then
    '            CopyFileErr = Pcm.ioArgumentNullErr
    '        Else
    '            CopyFileErr = Pcm.ioArgumentErr
    '        End If
    '    Catch IoEx As IO.IOException
    '        If TypeOf (IoEx) Is IO.PathTooLongException Then
    '            CopyFileErr = Pcm.ioPathTooLongErr
    '        ElseIf TypeOf (IoEx) Is IO.FileNotFoundException Then
    '            CopyFileErr = Pcm.ioFileNotFoundErr
    '        Else
    '            CopyFileErr = Pcm.ioErr
    '        End If
    '    Catch NoSupEx As NotSupportedException
    '        CopyFileErr = Pcm.ioNotSupportedErr
    '    Catch OpCancelErr As OperationCanceledException
    '        CopyFileErr = Pcm.ioOpCanceledErr
    '    Catch SecEx As Security.SecurityException
    '        CopyFileErr = Pcm.ioSecurityErr
    '    Catch UnAuthEx As UnauthorizedAccessException
    '        CopyFileErr = Pcm.ioUnAuthErr
    '    Catch ex As Exception
    '        CopyFileErr = Pcm.ioOtherErr
    '    Finally
    '        If CopyFileErr <> Pcm.NoErrors Then
    '            ShowCopyErr(CopyFileErr, FromFileName, ToFileName)
    '        End If
    '    End Try
    '    Return CopyFileErr
    'End Function

    'Private Sub ShowCopyErr(ByVal CopyFileErr As Integer, ByVal FromFileName As String, ByVal ToFileName As String)

    '    ' shows the copy file error message
    '    '
    '    ' vars passed:
    '    '   CopyFileErr - copy file error value
    '    '   FromFileName - file name to copy from
    '    '   ToFileName - file name to copy to

    '    If CopyFileErr = Pcm.NoErrors Then
    '        Exit Sub
    '    End If
    '    Try
    '        Dim InitMsg As String = String.Format("Cannot copy from file ""{0}"" to file ""{1}"".  ", FromFileName, ToFileName)
    '        Select Case CopyFileErr
    '            Case Pcm.ioArgumentErr
    '                MessageBox.Show(InitMsg & "Invalid path in one of the file names.", _
    '                                Pcm.etIoErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
    '            Case Pcm.ioArgumentNullErr
    '                MessageBox.Show(InitMsg & "One of the file names is blank.", _
    '                                Pcm.etIoErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
    '            Case Pcm.ioFileNotFoundErr
    '                MessageBox.Show(String.Format("{0}The file ""{1}"" is invalid or does not exist.", InitMsg, FromFileName),
    '                                Pcm.etIoErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
    '            Case Pcm.ioErr
    '                MessageBox.Show(String.Format("{0}Could not create the file ""{1}"".", InitMsg, ToFileName), _
    '                                Pcm.etIoErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
    '            Case Pcm.ioNotSupportedErr
    '                MessageBox.Show(InitMsg & "A file or folder name in the path contains a colon (:) or is in an invalid format.", _
    '                                Pcm.etIoErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
    '            Case Pcm.ioOpCanceledErr
    '                MessageBox.Show(InitMsg & "User canceled or unspecified I/O error.", _
    '                                Pcm.etIoErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
    '            Case Pcm.ioPathTooLongErr
    '                MessageBox.Show(InitMsg & "One of the folder/file names is too long.", _
    '                                Pcm.etIoErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
    '            Case Pcm.ioSecurityErr
    '                MessageBox.Show(String.Format("{0}User does not have permissions to view path for file ""{1}"".", InitMsg, ToFileName), _
    '                                Pcm.etIoErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
    '            Case Pcm.ioUnAuthErr
    '                MessageBox.Show(InitMsg & "User does not have required permission.", _
    '                                Pcm.etIoErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
    '            Case Pcm.ioInvalidParamErr
    '                MessageBox.Show(InitMsg & "Invalid parameters in call to ""CopyFile"".", _
    '                                Pcm.etIoErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
    '            Case Else
    '                MessageBox.Show(InitMsg & "Other I/O error.", _
    '                                Pcm.etIoErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        End Select
    '    Catch ex As Exception
    '        MessageBox.Show("Other I/O error.  " & ex.Message, _
    '                        Pcm.etIoErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    End Try
    'End Sub

#End Region

#Region " Rename File "

    'Public Function RenameFile(ByVal FromFileName As String, ByVal ToFileName As String) As Integer

    '    ' renames a file
    '    '
    '    ' vars passed:
    '    '   FromFileName - file name to rename from
    '    '   ToFileName - file name to rename to
    '    '
    '    ' returns:
    '    '   Pcm.NoErrors - no errors
    '    '   Pcm.ioArgumentErr - bad path name
    '    '   Pcm.ioArgumentNullErr - blank file name
    '    '   Pcm.ioErr - source file points to existing folder, 
    '    '       file exists and overwrite is FALSE
    '    '       user does not have sufficient permissions to access the file  
    '    '   Pcm.ioPathTooLongErr - file name too long
    '    '   Pcm.ioNotSupportedErr - folder name is a colon, or is in invalid format
    '    '   Pcm.ioUnAuthErr As Integer - user does not have permission to make folder
    '    '   Pcm.ioSecurityErr As Integer - user does not have permission in partial trust
    '    '   Pcm.ioFileNotFoundErr - source file not found/not valid
    '    '   Pcm.ioOpCanceledErr - ShowUI is set to True, onUserCancel is set to ThrowException, user has cancelled 
    '    '       ShowUI is set to True, onUserCancel is set to ThrowException, and an unspecified I/O error occurs 
    '    '   Pcm.ioOtherErr As Integer - other error

    '    Dim RenameFileErr As Integer = Pcm.NoErrors
    '    Try
    '        My.Computer.FileSystem.RenameFile(FromFileName, ToFileName)
    '    Catch ArgEx As ArgumentException
    '        If TypeOf (ArgEx) Is ArgumentNullException Then
    '            RenameFileErr = Pcm.ioArgumentNullErr
    '        Else
    '            RenameFileErr = Pcm.ioArgumentErr
    '        End If
    '    Catch IoEx As IO.IOException
    '        If TypeOf (IoEx) Is IO.PathTooLongException Then
    '            RenameFileErr = Pcm.ioPathTooLongErr
    '        ElseIf TypeOf (IoEx) Is IO.FileNotFoundException Then
    '            RenameFileErr = Pcm.ioFileNotFoundErr
    '        Else
    '            RenameFileErr = Pcm.ioErr
    '        End If
    '    Catch NoSupEx As NotSupportedException
    '        RenameFileErr = Pcm.ioNotSupportedErr
    '    Catch OpCancelErr As OperationCanceledException
    '        RenameFileErr = Pcm.ioOpCanceledErr
    '    Catch SecEx As Security.SecurityException
    '        RenameFileErr = Pcm.ioSecurityErr
    '    Catch UnAuthEx As UnauthorizedAccessException
    '        RenameFileErr = Pcm.ioUnAuthErr
    '    Catch ex As Exception
    '        RenameFileErr = Pcm.ioOtherErr
    '    Finally
    '        If RenameFileErr <> Pcm.NoErrors Then
    '            ShowRenameErr(RenameFileErr, FromFileName, ToFileName)
    '        End If
    '    End Try
    '    Return RenameFileErr
    'End Function

    'Public Function RenameFile(ByVal FromFileName As String,
    '                           ByVal ToFileName As String,
    '                           ByVal OverWrite As Boolean) As Integer

    '    ' renames a file
    '    '
    '    ' vars passed:
    '    '   FromFileName - file name to copy from
    '    '   ToFileName - file name to copy to
    '    '   OverWrite - TRUE: do overwrite file; FALSE: do not overwrite file
    '    '
    '    ' returns:
    '    '   Pcm.NoErrors - no errors
    '    '   Pcm.ioArgumentErr - bad path name
    '    '   Pcm.ioArgumentNullErr - blank file name
    '    '   Pcm.ioErr - source file points to existing folder, 
    '    '       file exists and overwrite is FALSE
    '    '       user does not have sufficient permissions to access the file  
    '    '   Pcm.ioPathTooLongErr - file name too long
    '    '   Pcm.ioNotSupportedErr - folder name is a colon, or is in invalid format
    '    '   Pcm.ioUnAuthErr As Integer - user does not have permission to make folder
    '    '   Pcm.ioSecurityErr As Integer - user does not have permission in partial trust
    '    '   Pcm.ioFileNotFoundErr - source file not found/not valid
    '    '   Pcm.ioOpCanceledErr - ShowUI is set to True, onUserCancel is set to ThrowException, user has cancelled 
    '    '       ShowUI is set to True, onUserCancel is set to ThrowException, and an unspecified I/O error occurs 
    '    '   Pcm.ioOtherErr As Integer - other error

    '    Dim RenameFileErr As Integer = Pcm.NoErrors
    '    Try
    '        If Not OverWrite Then                                       ' if not overwriting by default
    '            If My.Computer.FileSystem.FileExists(ToFileName) Then   ' if to file exists
    '                ' have user confirm overwrite
    '                If MessageBox.Show(String.Format("File ""{0}"" already exists.   Do you want to overwrite it?", ToFileName),
    '                                   Pcm.etConfirmOverwrite, MessageBoxButtons.YesNo, MessageBoxIcon.Question,
    '                                   MessageBoxDefaultButton.Button2) <> DialogResult.Yes Then
    '                    Return Pcm.ioOpCanceledErr                      ' if user did not confirm, return cancel
    '                End If
    '            End If
    '        End If

    '        RenameFileErr = RenameFile(FromFileName, ToFileName)
    '    Catch ex As Exception
    '        RenameFileErr = Pcm.ioOtherErr
    '        ShowRenameErr(RenameFileErr, FromFileName, ToFileName)
    '    End Try
    '    Return RenameFileErr
    'End Function

    'Private Sub ShowRenameErr(ByVal RenameFileErr As Integer, ByVal FromFileName As String, ByVal ToFileName As String)

    '    ' shows the copy file error message
    '    '
    '    ' vars passed:
    '    '   RenameFileErr - rename file error value
    '    '   FromFileName - file name to rename from
    '    '   ToFileName - file name to rename to

    '    If RenameFileErr = Pcm.NoErrors Then
    '        Exit Sub
    '    End If
    '    Try
    '        Dim InitMsg As String = String.Format("Cannot rename from file ""{0}"" to file ""{1}"".  ", FromFileName, ToFileName)
    '        Select Case RenameFileErr
    '            Case Pcm.ioArgumentErr
    '                MessageBox.Show(InitMsg & "Invalid path in one of the file names.", _
    '                                Pcm.etIoErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
    '            Case Pcm.ioArgumentNullErr
    '                MessageBox.Show(InitMsg & "One of the file names is blank.", _
    '                                Pcm.etIoErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
    '            Case Pcm.ioFileNotFoundErr
    '                MessageBox.Show(String.Format("{0}The file name ""{1}"" is invalid or does not exist.", InitMsg, FromFileName),
    '                                Pcm.etIoErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
    '            Case Pcm.ioErr
    '                MessageBox.Show(String.Format("{0}Could not create the file ""{1}"".", InitMsg, ToFileName), _
    '                                Pcm.etIoErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
    '            Case Pcm.ioNotSupportedErr
    '                MessageBox.Show(InitMsg & "A file or folder name in the path contains a colon (:) or is in an invalid format.", _
    '                                Pcm.etIoErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
    '            Case Pcm.ioOpCanceledErr
    '                MessageBox.Show(InitMsg & "User canceled or unspecified I/O error.", _
    '                                Pcm.etIoErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
    '            Case Pcm.ioPathTooLongErr
    '                MessageBox.Show(InitMsg & "One of the folder/file names is too long.", _
    '                                Pcm.etIoErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
    '            Case Pcm.ioSecurityErr
    '                MessageBox.Show(String.Format("{0}User does not have permissions to view path for file ""{1}"".", InitMsg, ToFileName), _
    '                                Pcm.etIoErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
    '            Case Pcm.ioUnAuthErr
    '                MessageBox.Show(InitMsg & "User does not have required permission.", _
    '                                Pcm.etIoErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
    '            Case Pcm.ioInvalidParamErr
    '                MessageBox.Show(InitMsg & "Invalid parameters in call to ""CopyFile"".", _
    '                                Pcm.etIoErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
    '            Case Else
    '                MessageBox.Show(InitMsg & "Other I/O error.", _
    '                                Pcm.etIoErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        End Select
    '    Catch ex As Exception
    '        MessageBox.Show("Other I/O error.  " & ex.Message, _
    '                        Pcm.etIoErr, MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    End Try
    'End Sub

#End Region

End Module
