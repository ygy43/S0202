Imports System.Security.Principal
Imports KatabanBusinessLogic.KatabanWcfService

Namespace MyHelpers
    Public Class MyPrincipal
        Implements IPrincipal

        Public ReadOnly Property Identity As IIdentity Implements IPrincipal.Identity
        Public Property User As UserInfo

        Public Sub New(identity As IIdentity)
            Me.Identity = identity
        End Sub


        Public Function IsInRole(role As String) As Boolean Implements IPrincipal.IsInRole
            Return True
        End Function
    End Class
End NameSpace