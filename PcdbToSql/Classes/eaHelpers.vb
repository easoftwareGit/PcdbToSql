Public Class eaHelpers

    '<Extension()>
    Public Shared Function GetPrivatePropertyValue(Of T)(ByVal obj As Object, ByVal propName As String) As T
        If obj Is Nothing Then Throw New ArgumentNullException("obj")
        Dim pi As System.Reflection.PropertyInfo =
                obj.GetType().GetProperty(propName, System.Reflection.BindingFlags.Public Or
                                                    System.Reflection.BindingFlags.NonPublic Or
                                                    System.Reflection.BindingFlags.Instance)
        If pi Is Nothing Then
            Throw New ArgumentOutOfRangeException("propName", String.Format("Property {0} was not found in Type {1}", propName, obj.GetType().FullName))
        End If
        Return CType(pi.GetValue(obj, Nothing), T)
    End Function


    '<Extension()>
    Public Shared Function GetPrivateFieldValue(Of T)(ByVal obj As Object, ByVal propName As String) As T
        If obj Is Nothing Then Throw New ArgumentNullException("obj")
        Dim objType As Type = obj.GetType()
        Dim fi As System.Reflection.FieldInfo = Nothing

        While fi Is Nothing AndAlso objType IsNot Nothing
            fi = objType.GetField(propName, System.Reflection.BindingFlags.Public Or
                                            System.Reflection.BindingFlags.NonPublic Or
                                            System.Reflection.BindingFlags.Instance)
            objType = objType.BaseType
        End While

        If fi Is Nothing Then
            Throw New ArgumentOutOfRangeException("propName", String.Format("Field {0} was not found in Type {1}", propName, obj.GetType().FullName))
        End If
        Return CType(fi.GetValue(obj), T)
    End Function

    '<Extension()>
    Public Shared Sub SetPrivatePropertyValue(Of T)(ByVal obj As Object, ByVal propName As String, ByVal val As T)
        Dim objType As Type = obj.GetType()
        If objType.GetProperty(propName, System.Reflection.BindingFlags.Public Or
                                         System.Reflection.BindingFlags.NonPublic Or
                                         System.Reflection.BindingFlags.Instance) Is Nothing Then
            Throw New ArgumentOutOfRangeException("propName", String.Format("Property {0} was not found in Type {1}", propName, obj.[GetType]().FullName))
        End If
        objType.InvokeMember(propName, System.Reflection.BindingFlags.Public Or
                                       System.Reflection.BindingFlags.NonPublic Or
                                       System.Reflection.BindingFlags.SetProperty Or
                                       System.Reflection.BindingFlags.Instance,
                             Nothing, obj, New Object() {val})
    End Sub

    '<Extension()>
    Public Shared Sub SetPrivateFieldValue(Of T)(ByVal obj As Object, ByVal propName As String, ByVal val As T)
        If obj Is Nothing Then Throw New ArgumentNullException("obj")
        Dim objType As Type = obj.GetType()
        Dim fi As System.Reflection.FieldInfo = Nothing

        While fi Is Nothing AndAlso objType IsNot Nothing
            fi = objType.GetField(propName, System.Reflection.BindingFlags.Public Or
                                            System.Reflection.BindingFlags.NonPublic Or
                                            System.Reflection.BindingFlags.Instance)
            objType = objType.BaseType
        End While

        If fi Is Nothing Then
            Throw New ArgumentOutOfRangeException("propName", String.Format("Field {0} was not found in Type {1}", propName, obj.GetType().FullName))
        End If
        fi.SetValue(obj, val)
    End Sub

End Class
