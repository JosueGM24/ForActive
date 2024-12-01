<%@ Page Language="VB" AutoEventWireup="true" %>
<%@ Import Namespace="System.DirectoryServices" %>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <link rel="stylesheet" href="Styles/grid.css">
    <link rel="stylesheet" href="Styles/styleGeneral.css">
    <link rel="stylesheet" href="Styles/styleMenu.css">
    <link rel="stylesheet" href="Styles/styleForms.css">
    <link rel="stylesheet" href="Styles/styleLoader.css">
    <link rel="stylesheet" href="Styles/styleTypes.css">
    <title>Gestionar Usuarios en Active Directory</title>
</head>
<body class="VerticalLayout center ww100 poppins-black fs-50">
    <form id="form1" runat="server" class="ww90 bgbl padding-50 br-50 VerticalLayout center">
        
        <!-- Campo para buscar usuario -->
        <div class="w100 VerticalLayout gap20 center">
            <label class="poppins-black fs-50" for="txtSearchUsername">Buscar Usuario:</label>
            <asp:TextBox ID="txtSearchUsername" class="inputNormal w100 fs-50" runat="server" />
            <asp:Button class="btnPrimary fs-50 poppins-black" type="button" runat="server" ID="btnSearch" Text="Buscar" OnClick="btnSearch_Click" />
        </div>
        <br />

        <!-- Formulario de datos del usuario -->
        <div class="VerticalLayout w100 center gap10">

            <div class="VerticalLayout">

                <label class="poppins-black fs-50" for="txtUsername">Nombre de Usuario:</label>
                <asp:TextBox ID="txtUsername" class="inputNormal w100 fs-50" runat="server" /><br />
            </div>
                
            <div class="VerticalLayout">

                <label class="poppins-black fs-50" for="txtFirstName">Primer Nombre:</label>
                <asp:TextBox ID="txtFirstName" class="inputNormal w100 fs-50" runat="server" /><br />
            </div>
                
            <div class="VerticalLayout">

                <label class="poppins-black fs-50" for="txtLastName">Apellido:</label>
                <asp:TextBox ID="txtLastName" class="inputNormal w100 fs-50" runat="server" /><br />
            </div>
                
            <div class="VerticalLayout">

                <label class="poppins-black fs-50" for="txtEmail">Correo Electrónico:</label>
                <asp:TextBox ID="txtEmail" class="inputNormal w100 fs-50" runat="server" /><br />
            </div>
                
            <div class="VerticalLayout">

                <label class="poppins-black fs-50" for="txtMobilePhone">Teléfono Móvil:</label>
                <asp:TextBox ID="txtMobilePhone" class="inputNormal w100 fs-50" runat="server" /><br />
            </div>
                
            <div class="VerticalLayout">

                <label class="poppins-black fs-50" for="txtHomePhone">Teléfono de Casa:</label>
                <asp:TextBox ID="txtHomePhone" class="inputNormal w100 fs-50" runat="server" /><br />
            </div>
                
            <div class="VerticalLayout">

                <label class="poppins-black fs-50" for="txtPassword">Contraseña:</label>
                <asp:TextBox ID="txtPassword" class="inputNormal w100 fs-50" runat="server" TextMode="Password" /><br />
            </div>
                
            <div class="VerticalLayout gap10">

                <asp:Button class="btnPrimary fs-50 poppins-black" type="submit" runat="server" ID="btnSubmit" Text="Guardar Usuario" OnClick="btnSubmit_Click" />
                <asp:Button class="btnPrimary fs-50 poppins-black" type="button" runat="server" ID="btnDelete" Text="Eliminar Usuario" OnClick="btnDelete_Click" />
            </div>
            </div>
    </form>


    <script runat="server">
        Protected Sub Page_Load(sender As Object, e As EventArgs)
            ' Lógica para cargar datos si es necesario en la carga de la página
        End Sub

        Protected Sub btnSearch_Click(sender As Object, e As EventArgs)
            Dim username As String = txtSearchUsername.Text.Trim()

            Try
                Dim entry As New DirectoryEntry("LDAP://10.0.0.20/DC=empresa,DC=com", "administrador@empresa.com", "Admin123")

                Dim searcher As New DirectorySearcher(entry) With {
                    .Filter = "(&(objectClass=user)(sAMAccountName=" & username & "))"
                }

                Dim result As SearchResult = searcher.FindOne()

                If result IsNot Nothing Then
                    Dim userEntry As DirectoryEntry = result.GetDirectoryEntry()
	            txtUsername.Text = userEntry.Properties("sAMAccountName").Value.ToString()
                    txtFirstName.Text = userEntry.Properties("givenName").Value.ToString()
                    txtLastName.Text = userEntry.Properties("sn").Value.ToString()
                    txtEmail.Text = userEntry.Properties("mail").Value.ToString()
                    txtMobilePhone.Text = If(userEntry.Properties("mobile").Value IsNot Nothing, userEntry.Properties("mobile").Value.ToString(), String.Empty)
                    txtHomePhone.Text = If(userEntry.Properties("homePhone").Value IsNot Nothing, userEntry.Properties("homePhone").Value.ToString(), String.Empty)

                    btnSubmit.Visible = True
                    btnDelete.Visible = True
                Else
                    ClearForm()
                    btnSubmit.Visible = True
                    btnDelete.Visible = False
                    Response.Write("Usuario no encontrado. Puede crear uno nuevo.")
                End If
            Catch ex As Exception
                Response.Write("Error: " & ex.Message)
            End Try
        End Sub

        Protected Sub btnSubmit_Click(sender As Object, e As EventArgs)
            Dim username As String = txtUsername.Text.Trim()
            Dim firstName As String = txtFirstName.Text.Trim()
            Dim lastName As String = txtLastName.Text.Trim()
            Dim email As String = txtEmail.Text.Trim()
            Dim password As String = txtPassword.Text.Trim()
            Dim mobilePhone As String = txtMobilePhone.Text.Trim()
            Dim homePhone As String = txtHomePhone.Text.Trim()

            Try
                Dim entry As New DirectoryEntry("LDAP://10.0.0.20/DC=empresa,DC=com", "administrador@empresa.com", "Admin123")
                Dim searcher As New DirectorySearcher(entry) With {
                    .Filter = "(&(objectClass=user)(sAMAccountName=" & username & "))"
                }

                Dim result As SearchResult = searcher.FindOne()

                If result IsNot Nothing Then
                    Dim userEntry As DirectoryEntry = result.GetDirectoryEntry()
                    userEntry.Properties("givenName").Value = firstName
                    userEntry.Properties("sn").Value = lastName
                    userEntry.Properties("mail").Value = email
                    userEntry.Properties("mobile").Value = mobilePhone
                    userEntry.Properties("homePhone").Value = homePhone
                    userEntry.CommitChanges()

                    Response.Write("Usuario actualizado correctamente.")
                Else
                    Dim entryUsers As New DirectoryEntry("LDAP://10.0.0.20/CN=Users,DC=empresa,DC=com", "administrador@empresa.com", "Admin123")
                    Dim newUser As DirectoryEntry = entryUsers.Children.Add("CN=" & username, "user")

                    newUser.Properties("sAMAccountName").Value = username
                    newUser.Properties("userPrincipalName").Value = username & "@empresa.com"
                    newUser.Properties("givenName").Value = If(Not String.IsNullOrEmpty(firstName), firstName, String.Empty)
                    newUser.Properties("sn").Value = If(Not String.IsNullOrEmpty(lastName), lastName, String.Empty)

                    If Not String.IsNullOrEmpty(email) Then
                        newUser.Properties("mail").Value = email
                    End If
                    If Not String.IsNullOrEmpty(mobilePhone) Then
                        newUser.Properties("mobile").Value = mobilePhone
                    End If
                    If Not String.IsNullOrEmpty(homePhone) Then
                        newUser.Properties("homePhone").Value = homePhone
                    End If

                    newUser.CommitChanges()
                    newUser.Invoke("SetPassword", password)
                    newUser.Properties("userAccountControl").Value = 512
                    newUser.CommitChanges()

                    Response.Write("Usuario creado correctamente.")
                End If
            Catch ex As Exception
                Response.Write("Error: " & ex.Message)
            End Try
        End Sub

        Protected Sub btnDelete_Click(sender As Object, e As EventArgs)
            Dim username As String = txtUsername.Text.Trim()

            Try
                Dim entry As New DirectoryEntry("LDAP://10.0.0.20/DC=empresa,DC=com", "administrador@empresa.com", "Admin123")
                Dim searcher As New DirectorySearcher(entry) With {
                    .Filter = "(&(objectClass=user)(sAMAccountName=" & username & "))"
                }

                Dim result As SearchResult = searcher.FindOne()

                If result IsNot Nothing Then
                    Dim userEntry As DirectoryEntry = result.GetDirectoryEntry()
                    userEntry.DeleteTree()
                    Response.Write("Usuario eliminado correctamente.")
                Else
                    Response.Write("El usuario no existe.")
                End If
            Catch ex As Exception
                Response.Write("Error: " & ex.Message)
            End Try
        End Sub

        Private Sub ClearForm()
            txtUsername.Text = String.Empty
            txtFirstName.Text = String.Empty
            txtLastName.Text = String.Empty
            txtEmail.Text = String.Empty
            txtPassword.Text = String.Empty
            txtMobilePhone.Text = String.Empty
            txtHomePhone.Text = String.Empty
        End Sub
    </script>
</body>
</html>
