<%@ Page aspcompat=true Language="VB" %>
<%@ Import NameSpace="System.Data" %>
<%@ Import NameSpace="System.Data.oleDB" %>
<%@ Import NameSpace="System.DirectoryServices" %>
<%@ Import NameSpace="System.Net" %>

<script runat="server" language="vb">
Sub Page_Load()

    Dim Entry As New DirectoryEntry("LDAP://10.0.0.1/cn=alumno,cn=Users,DC=empresa,DC=com", "empresa\alumno", "@lumn@s")

    Dim Searcher As New DirectorySearcher(Entry)
    Dim res As SearchResult = Searcher.FindOne()
    Dim s As String

    For Each s In Entry.Properties("homePhone")
        response.write("Telefono de casa: ")
        response.write(s)
        response.write("<br>")
    Next

    For Each s In Entry.Properties("mobile")
        response.write("Numero de movil: ")
        response.write(s)
        response.write("<br>")
    Next

    For Each s In Entry.Properties("name")
        response.write("Nombre de usuario: ")
        response.write(s)
        response.write("<br>")
    Next

    If res Is Nothing Then
        response.write("No existe usuario")
    Else
        response.write("Conexion exitosa")
    End If

End Sub
</script>
<html>
<head>
    <title>Ejemplo de ASP</title>
</head>
<body>
<asp:DataGrid id="dgrdAlumnos" runat="server"><asp:DataGrid>
</body>
</html>