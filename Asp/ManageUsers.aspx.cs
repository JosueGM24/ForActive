using System;
using System.DirectoryServices;
using System.Web.UI;

namespace ADUserManagement
{
    public partial class ManageUsers : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            // Lógica para cargar datos si es necesario en la carga de la página
        }

        protected void btnSearch_Click(object sender, EventArgs e)
        {
            string username = txtSearchUsername.Value.Trim();

            try
            {
                // Conectar al servidor LDAP
                DirectoryEntry entry = new DirectoryEntry("LDAP://10.0.0.20/DC=empresa,DC=com", "admin@empresa.com", "adminpassword");
                
                // Buscar usuario por su sAMAccountName
                DirectorySearcher searcher = new DirectorySearcher(entry)
                {
                    Filter = $"(&(objectClass=user)(sAMAccountName={username}))"
                };

                SearchResult result = searcher.FindOne();

                if (result != null)
                {
                    // Si el usuario existe, cargar los datos en los campos
                    DirectoryEntry userEntry = result.GetDirectoryEntry();

                    txtUsername.Value = userEntry.Properties["sAMAccountName"].Value.ToString();
                    txtFirstName.Value = userEntry.Properties["givenName"].Value.ToString();
                    txtLastName.Value = userEntry.Properties["sn"].Value.ToString();
                    txtEmail.Value = userEntry.Properties["mail"].Value.ToString();
                    
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
                Response.Write($"Error: {ex.Message}");
            }
        }

        protected void btnSubmit_Click(object sender, EventArgs e)
        {
            string username = txtUsername.Value.Trim();
            string firstName = txtFirstName.Value.Trim();
            string lastName = txtLastName.Value.Trim();
            string email = txtEmail.Value.Trim();
            string password = txtPassword.Value.Trim();

            try
            {
                DirectoryEntry entry = new DirectoryEntry("LDAP://10.0.0.20/DC=empresa,DC=com", "administrador@empresa.com", "Admin123");

                // Buscar el usuario
                DirectorySearcher searcher = new DirectorySearcher(entry)
                {
                    Filter = $"(&(objectClass=user)(sAMAccountName={username}))"
                };

                SearchResult result = searcher.FindOne();

                if (result != null)
                {
                    // Usuario encontrado, actualizar datos
                    DirectoryEntry userEntry = result.GetDirectoryEntry();
                    userEntry.Properties["givenName"].Value = firstName;
                    userEntry.Properties["sn"].Value = lastName;
                    userEntry.Properties["mail"].Value = email;
                    userEntry.CommitChanges();
                    
                    Response.Write("Usuario actualizado correctamente.");
                }
                else
                {
                    // Usuario no encontrado, crear uno nuevo
                    DirectoryEntry newUser = entry.Children.Add($"CN={username}", "user");
                    newUser.Properties["sAMAccountName"].Value = username;
                    newUser.Properties["givenName"].Value = firstName;
                    newUser.Properties["sn"].Value = lastName;
                    newUser.Properties["mail"].Value = email;
                    newUser.Properties["userPassword"].Value = password;
                    newUser.CommitChanges();

                    Response.Write("Usuario creado correctamente.");
                }
            }
            catch (Exception ex)
            {
                Response.Write($"Error: {ex.Message}");
            }
        }

        protected void btnDelete_Click(object sender, EventArgs e)
        {
            string username = txtUsername.Value.Trim();

            try
            {
                DirectoryEntry entry = new DirectoryEntry("LDAP://10.0.0.20/DC=empresa,DC=com", "admin@empresa.com", "adminpassword");

                // Buscar el usuario
                DirectorySearcher searcher = new DirectorySearcher(entry)
                {
                    Filter = $"(&(objectClass=user)(sAMAccountName={username}))"
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
                Response.Write($"Error: {ex.Message}");
            }
        }

        private void ClearForm()
        {
            // Limpiar los campos del formulario
            txtUsername.Value = string.Empty;
            txtFirstName.Value = string.Empty;
            txtLastName.Value = string.Empty;
            txtEmail.Value = string.Empty;
            txtPassword.Value = string.Empty;
        }
    }
}
