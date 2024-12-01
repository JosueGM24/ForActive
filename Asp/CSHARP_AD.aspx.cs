<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="Autentificar_LDAP.Default" %>

<!DOCTYPE html>
<html>
<head runat="server">
    <title>Autentificación LDAP</title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <label for="txtUsuario">Usuario:</label>
            <asp:TextBox ID="txtUsuario" runat="server"></asp:TextBox>
            <br />
            <label for="txtPassword">Contraseña:</label>
            <asp:TextBox ID="txtPassword" TextMode="Password" runat="server"></asp:TextBox>
            <br />
            <label for="txtDominio">Dominio:</label>
            <asp:TextBox ID="txtDominio" runat="server"></asp:TextBox>
            <br />
            <asp:Button ID="btnAutenticar" runat="server" Text="Autenticar" OnClick="btnAutenticar_Click" />
            <br />
            <asp:Label ID="lblResultado" runat="server"></asp:Label>
        </div>
    </form>
</body>
</html>