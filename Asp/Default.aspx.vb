Imports System.Data
Imports System.DirectoryServices

Partial Class _Default
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
        Try
            Dim entry As New DirectoryEntry("LDAP://10.0.0.20/cn=daniel,cn=Users,DC=empresa,DC=com", "daniel@empresa.com", "Admin123")
            Dim searcher As New DirectorySearcher(entry)
            Dim res As SearchResult = searcher.FindOne()

            If res Is Nothing Then
                lblOutput.Text = "No existe usuario"
            Else
                lblOutput.Text = "Conexión exitosa<br>"

                ' Mostrar propiedades
                For Each s As String In entry.Properties("homePhone")
                    lblOutput.Text &= "Teléfono de casa: " & s & "<br>"
                Next

                For Each s As String In entry.Properties("mobile")
                    lblOutput.Text &= "Número de móvil: " & s & "<br>"
                Next

                For Each s As String In entry.Properties("name")
                    lblOutput.Text &= "Nombre de usuario: " & s & "<br>"
                Next
            End If

        Catch ex As Exception
            lblOutput.Text = "Error: " & ex.Message
        End Try
    End Sub
End Class
