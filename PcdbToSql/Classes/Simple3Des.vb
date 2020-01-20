Imports System.Security.Cryptography

Public NotInheritable Class Simple3Des

#Region " Notes "

    ' https://docs.microsoft.com/en-us/dotnet/visual-basic/programming-guide/language-features/strings/walkthrough-encrypting-and-decrypting-strings

    'Sub TestEncoding()
    '    Dim plainText As String = InputBox("Enter the plain text:")
    '    Dim password As String = InputBox("Enter the password:")

    '    Dim wrapper As New Simple3Des(password)
    '    Dim cipherText As String = wrapper.EncryptData(plainText)

    '    MsgBox("The cipher text is: " & cipherText)
    '    My.Computer.FileSystem.WriteAllText(
    '    My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\cipherText.txt", cipherText, False)
    'End Sub

    'Sub TestDecoding()
    '    Dim cipherText As String = My.Computer.FileSystem.ReadAllText(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\cipherText.txt")
    '    Dim password As String = InputBox("Enter the password:")
    '    Dim wrapper As New Simple3Des(password)

    '    ' DecryptData throws if the wrong password is used.
    '    Try
    '        Dim plainText As String = wrapper.DecryptData(cipherText)
    '        MsgBox("The plain text is: " & plainText)
    '    Catch ex As System.Security.Cryptography.CryptographicException
    '        MsgBox("The data could not be decrypted with the password.")
    '    End Try
    'End Sub

#End Region

#Region " Class Constants, ENums and Variables "

#Region " Constants "

#End Region

#Region " Enums "

#End Region

#Region " Private Vars "

    Private TripleDes As New TripleDESCryptoServiceProvider

#End Region

#End Region

#Region " New "

    Sub New(ByVal key As String)

        ' vars passed
        '   key - password key to use for encryption

        ' Initialize the crypto provider.
        TripleDes.Key = TruncateHash(key, TripleDes.KeySize \ 8)
        TripleDes.IV = TruncateHash("", TripleDes.BlockSize \ 8)
    End Sub

#End Region

#Region " Properties & Methods "

#Region " Properties "

#End Region

#Region " Methods "

    Public Function DecryptData(ByVal encryptedText As String) As String

        ' decrypts encrypted text into plain text
        '
        ' vars passed:
        '   encryptedText - text to decrypt
        '
        ' returns:
        '   decrypted plain text 

        If encryptedText = String.Empty Then
            Return String.Empty
        End If

        ' Convert the encrypted text string to a byte array.
        Dim encryptedBytes() As Byte = Convert.FromBase64String(encryptedText)

        ' Create the stream.
        Dim ms As New System.IO.MemoryStream
        ' Create the decoder to write to the stream.
        Dim decStream As New CryptoStream(ms, TripleDes.CreateDecryptor(), System.Security.Cryptography.CryptoStreamMode.Write)

        ' Use the crypto stream to write the byte array to the stream.
        decStream.Write(encryptedBytes, 0, encryptedBytes.Length)
        decStream.FlushFinalBlock()

        ' Convert the plaintext stream to a string.
        Return System.Text.Encoding.Unicode.GetString(ms.ToArray)
    End Function

    Public Function EncryptData(ByVal plaintext As String) As String

        ' encrypts plain text 
        '
        ' vars passed:
        '   plaintext - text to encrypt
        '
        ' returns:
        '   encrypted text 

        ' Convert the plaintext string to a byte array.
        Dim plaintextBytes() As Byte = System.Text.Encoding.Unicode.GetBytes(plaintext)

        ' Create the stream.
        Dim ms As New System.IO.MemoryStream
        ' Create the encoder to write to the stream.
        Dim encStream As New CryptoStream(ms, TripleDes.CreateEncryptor(), System.Security.Cryptography.CryptoStreamMode.Write)

        ' Use the crypto stream to write the byte array to the stream.
        encStream.Write(plaintextBytes, 0, plaintextBytes.Length)
        encStream.FlushFinalBlock()

        ' Convert the encrypted stream to a printable string.
        Return Convert.ToBase64String(ms.ToArray)
    End Function

#End Region

#End Region

#Region " TruncateHash "

    Private Function TruncateHash(ByVal key As String,
                                  ByVal length As Integer) As Byte()

        ' creates a byte array of a specified length from the hash of the specified key.
        '
        ' vars passed:
        '   key - password key
        '   length - length of the hash array of bytes

        Dim sha1 As New SHA1CryptoServiceProvider

        ' Hash the key.
        Dim keyBytes() As Byte = System.Text.Encoding.Unicode.GetBytes(key)
        Dim hash() As Byte = sha1.ComputeHash(keyBytes)

        ' Truncate or pad the hash.
        ReDim Preserve hash(length - 1)
        Return hash
    End Function

#End Region

End Class
