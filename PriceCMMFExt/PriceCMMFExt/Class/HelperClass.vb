Imports System
Imports System.Security.Principal
Imports System.DirectoryServices
Imports System.IO


Public Class HelperClass
    Implements IDisposable

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls


    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).

            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

    Dim _userInfo As UserInfo
    Dim _UserId As String

    Public ReadOnly Property UserInfo As UserInfo
        Get
            Return _userInfo
        End Get
    End Property


    Public Property UserId As String
        Get
            Return _UserId
        End Get
        Set(ByVal value As String)
            _UserId = value
        End Set
    End Property

    Public Sub New()
        GetCurrentUser()
    End Sub


    Private Sub GetCurrentUser()

        Dim genericPrincipal As GenericPrincipal = GetGenericPrincipal()
        Dim principalIdentity As GenericIdentity = CType(genericPrincipal.Identity, GenericIdentity)
        ' Display identity name and authentication type.

        If Not principalIdentity.IsAuthenticated Then
            Throw New System.Exception("You are not allowed to use this application")
        Else
            _UserId = principalIdentity.Name
            _userInfo = New UserInfo

            getDataAD(_UserId)
        End If
    End Sub

    Private Function GetGenericPrincipal() As GenericPrincipal
        ' Use values from the current WindowsIdentity to construct
        ' a set of GenericPrincipal roles.
        Dim roles(10) As String
        Dim windowsIdentity As WindowsIdentity = windowsIdentity.GetCurrent()

        If (windowsIdentity.IsAuthenticated) Then
            ' Add custom NetworkUser role.
            roles(0) = "NetworkUser"
        End If

        If (windowsIdentity.IsGuest) Then
            ' Add custom GuestUser role.
            roles(1) = "GuestUser"
        End If


        If (windowsIdentity.IsSystem) Then
            ' Add custom SystemUser role.
            roles(2) = "SystemUser"
        End If

        ' Construct a GenericIdentity object based on the current Windows
        ' identity name and authentication type.
        Dim authenticationType As String = windowsIdentity.AuthenticationType
        Dim userName As String = windowsIdentity.Name
        Dim genericIdentity = New GenericIdentity(userName, authenticationType)

        ' Construct a GenericPrincipal object based on the generic identity
        ' and custom roles for the user.
        Dim genericPrincipal As New GenericPrincipal(genericIdentity, roles)

        Return genericPrincipal
    End Function

    Public Sub getDataAD(ByRef userid As String)

        Dim entry As DirectoryEntry = New DirectoryEntry


        Dim myuser() As String = userid.Split("\")
        _userInfo.Domain = myuser(0)
        _userInfo.userid = myuser(1)
        Select Case _userInfo.Domain.ToString.ToLower
            Case "as"
                entry.Path = "LDAP://as/DC=as;DC=seb;DC=com"
            Case "eu"
                entry.Path = "LDAP://eu/DC=eu;DC=seb;DC=com"
            Case "supor"
                entry.Path = "LDAP://supor/DC=supor;DC=seb;DC=com"
            Case "sa"
                entry.Path = "LDAP://sa/DC=sa;DC=seb;DC=com"
            Case "na"
                entry.Path = "LDAP://na/DC=na;DC=seb;DC=com"
        End Select

        Try
            Dim mysearcher As DirectorySearcher = New DirectorySearcher(entry)
            mysearcher.Filter = "sAMAccountName=" & _userInfo.userid
            Dim result As SearchResult = mysearcher.FindOne
            If Not (result Is Nothing) Then
                Dim myPerson As DirectoryEntry = New DirectoryEntry
                myPerson.Path = result.Path

                Try
                    _userInfo.email = myPerson.Properties("mail").Value.ToString
                Catch ex As Exception
                End Try

                Try
                    _userInfo.DisplayName = UCase(myPerson.Properties("givenname").Value.ToString) & " " & UCase(myPerson.Properties("sn").Value.ToString)
                Catch
                End Try

                Try
                    _userInfo.Department = UCase(myPerson.Properties("department").Value.ToString) & " " & UCase(myPerson.Properties("sn").Value.ToString)
                Catch
                End Try
                'Dim counter As Integer = 0
                'For Each elemantName In myPerson.Properties.PropertyNames

                '    Dim valuecollection As PropertyValueCollection = myPerson.Properties(elemantName)
                '    For i = 0 To valuecollection.Count - 1
                '        'Debug.WriteLine(elemantName + "[" + i.ToString() + "]=" + valuecollection(i).ToString())
                '        Debug.WriteLine("{0} {1}[{2}] = {3}", counter, elemantName, i, valuecollection(i).ToString)
                '    Next
                '    counter += 1
                'Next

            End If
        Catch ex As Exception

        End Try
        myuser = Nothing
    End Sub


    Public Sub fadeout(ByVal o As System.Windows.Forms.Form)
        Dim accelerator As Double = 0
        For i = -1 To 0 Step 0.01
            o.Opacity = System.Math.Abs(i)
            o.Refresh()
            accelerator += 0.01
            i += accelerator
        Next
    End Sub
    
End Class

Public Class UserInfo
    Public Property email As String
    Public Property userid As String
    Public Property Domain As String
    Public Property DisplayName As String
    Public Property Department As String
    Public Property isAdmin As Boolean
End Class
