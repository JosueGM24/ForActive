using System;
using System.DirectoryServices;

public partial class _Default : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        try
        {
            // Configuraci�n de conexi�n LDAP
            DirectoryEntry entry = new DirectoryEntry(
                "LDAP://10.0.0.20/cn=daniel,cn=Users,DC=empresa,DC=com", 
                "daniel@empresa.com", 
                "Admin123"
            );

            DirectorySearcher searcher = new DirectorySearcher(entry);
            SearchResult result = searcher.FindOne();

            if (result == null)
            {
                lblOutput.Text = "No existe usuario.";
            }
            else
            {
                lblOutput.Text = "Conexi�n exitosa<br>";

                // Obtener propiedades y mostrarlas
                foreach (string homePhone in entry.Properties["homePhone"])
                {
                    lblOutput.Text += "Tel�fono de casa: " + homePhone + "<br>";
                }

                foreach (string mobile in entry.Properties["mobile"])
                {
                    lblOutput.Text += "N�mero de m�vil: " + mobile + "<br>";
                }

                foreach (string name in entry.Properties["name"])
                {
                    lblOutput.Text += "Nombre de usuario: " + name + "<br>";
                }
            }
        }
        catch (Exception ex)
        {
            lblOutput.Text = "Error: " + ex.Message;
        }
    }
}
