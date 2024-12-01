<%@ Page Language="C#" AutoEventWireup="true" %>
<%@ Import Namespace="System.DirectoryServices" %>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Gestionar Usuarios en Active Directory</title>
    <link rel="stylesheet" href="Styles/grid.css">
    <link rel="stylesheet" href="Styles/styleGeneral.css">
    <link rel="stylesheet" href="Styles/styleMenu.css">
    <link rel="stylesheet" href="Styles/styleForms.css">
    <link rel="stylesheet" href="Styles/styleLoader.css">
    <link rel="stylesheet" href="Styles/styleTypes.css">
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
        protected void Page_Load(object sender, EventArgs e)
        {
            // Lógica para cargar datos si es necesario en la carga de la página
        }

        protected void btnSearch_Click(object sender, EventArgs e)
        {
            string username = txtSearchUsername.Text.Trim();

            try
            {
                // Conectar al servidor LDAP
                DirectoryEntry entry = new DirectoryEntry("LDAP://10.0.0.20/DC=empresa,DC=com", "administrador@empresa.com", "Admin123");
                
                // Buscar usuario por su sAMAccountName
                DirectorySearcher searcher = new DirectorySearcher(entry)
                {
                    Filter = "(&(objectClass=user)(sAMAccountName=" + username + "))"
                };

                SearchResult result = searcher.FindOne();

                if (result != null)
                {
                    // Si el usuario existe, cargar los datos en los campos
                    DirectoryEntry userEntry = result.GetDirectoryEntry();

                    txtUsername.Text = userEntry.Properties["sAMAccountName"].Value.ToString();
                    txtFirstName.Text = userEntry.Properties["givenName"].Value.ToString();
                    txtLastName.Text = userEntry.Properties["sn"].Value.ToString();
                    txtEmail.Text = userEntry.Properties["mail"].Value.ToString();
                    txtMobilePhone.Text = userEntry.Properties["mobile"].Value != null ? userEntry.Properties["mobile"].Value.ToString() : string.Empty;
                    txtHomePhone.Text = userEntry.Properties["homePhone"].Value != null ? userEntry.Properties["homePhone"].Value.ToString() : string.Empty;

                    // Se habilitan los botones para actualizar y eliminar
                    btnSubmit.Visible = true;
                    btnDelete.Visible = true;
                }
                else
                {
                    // Si no existe, limpiar los campos y permitir creación de nuevo usuario
                    ClearForm();
                    btnSubmit.Visible = true;  // Permitir creación de nuevo usuario
                    btnDelete.Visible = false; // No es posible eliminar si no existe
                    Response.Write("Usuario no encontrado. Puede crear uno nuevo.");
                }
            }
            catch (Exception ex)
            {
                // Manejo de errores
                Response.Write("Error: " + ex.Message);
            }
        }

        protected void btnSubmit_Click(object sender, EventArgs e)
        {
            string username = txtUsername.Text.Trim();
            string firstName = txtFirstName.Text.Trim();
            string lastName = txtLastName.Text.Trim();
            string email = txtEmail.Text.Trim();
            string password = txtPassword.Text.Trim();
            string mobilePhone = txtMobilePhone.Text.Trim();
            string homePhone = txtHomePhone.Text.Trim();

            try
            {
                DirectoryEntry entry = new DirectoryEntry("LDAP://10.0.0.20/DC=empresa,DC=com", "administrador@empresa.com", "Admin123");

                // Buscar el usuario
                DirectorySearcher searcher = new DirectorySearcher(entry)
                {
                    Filter = "(&(objectClass=user)(sAMAccountName=" + username + "))"
                };

                SearchResult result = searcher.FindOne();

                if (result != null)
                {
                    // Usuario encontrado, actualizar datos
                    DirectoryEntry userEntry = result.GetDirectoryEntry();
                    userEntry.Properties["givenName"].Value = firstName;
                    userEntry.Properties["sn"].Value = lastName;
                    userEntry.Properties["mail"].Value = email;
                    userEntry.Properties["mobile"].Value = mobilePhone;
                    userEntry.Properties["homePhone"].Value = homePhone;
                    userEntry.CommitChanges();
                    
                    Response.Write("Usuario actualizado correctamente.");
                }
                else
    	        {
    	            try
                    {
                        // Crear el nuevo usuario
                        DirectoryEntry entryUsers = new DirectoryEntry("LDAP://10.0.0.20/CN=Users,DC=empresa,DC=com", "administrador@empresa.com", "Admin123");

                        DirectoryEntry newUser = entryUsers.Children.Add("CN=" + username, "user");
                    
                        // Asignar propiedades obligatorias
                        newUser.Properties["sAMAccountName"].Value = username;
                        newUser.Properties["userPrincipalName"].Value = username + "@empresa.com";
                        newUser.Properties["givenName"].Value = !string.IsNullOrEmpty(firstName) ? firstName : string.Empty;
                        newUser.Properties["sn"].Value = !string.IsNullOrEmpty(lastName) ? lastName : string.Empty;
                    
                        // Propiedades opcionales
                        if (!string.IsNullOrEmpty(email))
                            newUser.Properties["mail"].Value = email;
                        if (!string.IsNullOrEmpty(mobilePhone))
                            newUser.Properties["mobile"].Value = mobilePhone;
                        if (!string.IsNullOrEmpty(homePhone))
                            newUser.Properties["homePhone"].Value = homePhone;
                    
                        // Guardar el usuario antes de establecer la contraseña
                        newUser.CommitChanges();
                    
                        // Establecer contraseña después de haber creado el usuario
                        newUser.Invoke("SetPassword", new object[] { password });
                    
                        // Habilitar cuenta
                        newUser.Properties["userAccountControl"].Value = 512; // 512 = Habilitada
                        newUser.CommitChanges();
                    
                        Response.Write("Usuario creado correctamente.");
                    }
                    catch (UnauthorizedAccessException uaEx)
                    {
                        Response.Write("Error de permisos: " + uaEx.Message);
                    }
                    catch (Exception ex)
                    {
                        Response.Write("Error: " + ex.Message);
                    }



                }
            }
            catch (Exception ex)
            {
                Response.Write("Error: " + ex.Message);
            }

            
        }
        
        

        protected void btnDelete_Click(object sender, EventArgs e)
        {
            string username = txtUsername.Text.Trim();

            try
            {
                DirectoryEntry entry = new DirectoryEntry("LDAP://10.0.0.20/DC=empresa,DC=com", "administrador@empresa.com", "Admin123");

                // Buscar el usuario
                DirectorySearcher searcher = new DirectorySearcher(entry)
                {
                    Filter = "(&(objectClass=user)(sAMAccountName=" + username + "))"
                };

                SearchResult result = searcher.FindOne();

                if (result != null)
                {
                    // Eliminar el usuario
                    DirectoryEntry userEntry = result.GetDirectoryEntry();
                    userEntry.DeleteTree();
                    Response.Write("Usuario eliminado correctamente.");
                }
                else
                {
                    Response.Write("El usuario no existe.");
                }
            }
            catch (Exception ex)
            {
                Response.Write("Error: " + ex.Message);
            }
        }

        private void ClearForm()
        {
            // Limpiar los campos del formulario
            txtUsername.Text = string.Empty;
            txtFirstName.Text = string.Empty;
            txtLastName.Text = string.Empty;
            txtEmail.Text = string.Empty;
            txtPassword.Text = string.Empty;
            txtMobilePhone.Text = string.Empty;
            txtHomePhone.Text = string.Empty;
        }
    </script>
</body>
</html>