<%@ Page aspcompat=true Language="VB" %>
<%@ Import NameSpace="System.Data" %>
<%@ Import NameSpace="System.Data.oleDB" %>
<%@ Import NameSpace="System.DirectoryServices" %>
<%@ Import NameSpace="System.Net" %>
<link rel="stylesheet" href="Styles/grid.css">
<link rel="stylesheet" href="Styles/styleGeneral.css">
<link rel="stylesheet" href="Styles/styleMenu.css">
<link rel="stylesheet" href="Styles/styleForms.css">
<link rel="stylesheet" href="Styles/styleLoader.css">
<link rel="stylesheet" href="Styles/styleTypes.css">
<script runat="server" language="vb">
Sub Page_Load()

    Dim Entry As New DirectoryEntry("LDAP://10.0.0.20/cn=daniel,cn=Users,DC=empresa,DC=com", "daniel@empresa.com", "Admin123")

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
<body class="VerticalLayout center w90 br-20 gap20 padding-30 fs-40 poppins-black bgbl">
    <form id="form1" runat="server">
        <asp:GridView ID="grdAlumnos" runat="server">
        </asp:GridView>
    </form>
</body>
</html>
