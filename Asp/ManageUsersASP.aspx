<%@ Language="VBScript" %>
<!DOCTYPE html>
<html>
<head>
    <title>Gestión de Usuarios en Active Directory</title>
</head>
<body>
    <h2>Gestión de Usuarios de Active Directory</h2>
    <hr />

    <!-- Campo para buscar usuario -->
    <form method="post" action="ad_manager.asp">
        <div>
            <label for="txtSearchUsername">Buscar Usuario:</label>
            <input type="text" id="txtSearchUsername" name="txtSearchUsername" />
            <input type="submit" name="action" value="Buscar" />
        </div>
        <br />

        <!-- Formulario de datos del usuario -->
        <div>
            <label for="txtUsername">Nombre de Usuario:</label>
            <input type="text" id="txtUsername" name="txtUsername" /><br />

            <label for="txtFirstName">Primer Nombre:</label>
            <input type="text" id="txtFirstName" name="txtFirstName" /><br />

            <label for="txtLastName">Apellido:</label>
            <input type="text" id="txtLastName" name="txtLastName" /><br />

            <label for="txtEmail">Correo Electrónico:</label>
            <input type="email" id="txtEmail" name="txtEmail" /><br />

            <label for="txtMobilePhone">Teléfono Móvil:</label>
            <input type="text" id="txtMobilePhone" name="txtMobilePhone" /><br />

            <label for="txtHomePhone">Teléfono de Casa:</label>
            <input type="text" id="txtHomePhone" name="txtHomePhone" /><br />

            <label for="txtPassword">Contraseña:</label>
            <input type="password" id="txtPassword" name="txtPassword" /><br />

            <input type="submit" name="action" value="Guardar" />
            <input type="submit" name="action" value="Eliminar" />
        </div>
    </form>

<%
    Const ADS_PROPERTY_UPDATE = 2
    Dim action, username, firstName, lastName, email, mobilePhone, homePhone, password
    action = Request.Form("action")
    username = Trim(Request.Form("txtUsername"))
    firstName = Trim(Request.Form("txtFirstName"))
    lastName = Trim(Request.Form("txtLastName"))
    email = Trim(Request.Form("txtEmail"))
    mobilePhone = Trim(Request.Form("txtMobilePhone"))
    homePhone = Trim(Request.Form("txtHomePhone"))
    password = Trim(Request.Form("txtPassword"))

    Dim ldapPath, adminUser, adminPassword
    ldapPath = "LDAP://10.0.0.20/DC=empresa,DC=com"
    adminUser = "administrador@empresa.com"
    adminPassword = "Admin123"

    If action = "Buscar" Then
        Dim searchUsername
        searchUsername = Trim(Request.Form("txtSearchUsername"))
        Call SearchUser(ldapPath, adminUser, adminPassword, searchUsername)
    ElseIf action = "Guardar" Then
        Call SaveUser(ldapPath, adminUser, adminPassword, username, firstName, lastName, email, mobilePhone, homePhone, password)
    ElseIf action = "Eliminar" Then
        Call DeleteUser(ldapPath, adminUser, adminPassword, username)
    End If

    Sub SearchUser(ldapPath, adminUser, adminPassword, searchUsername)
        On Error Resume Next
        Dim conn, command, result
        Set conn = GetObject(ldapPath)
        conn.Username = adminUser
        conn.Password = adminPassword

        Set command = CreateObject("ADODB.Command")
        Set command.ActiveConnection = conn
        command.CommandText = "<LDAP://" & ldapPath & ">;(&(objectClass=user)(sAMAccountName=" & searchUsername & "));givenName,sn,mail,mobile,homePhone;subtree"

        Set result = command.Execute
        If Not result.EOF Then
            Response.Write("<p>Usuario encontrado:</p>")
            Response.Write("Nombre: " & result.Fields("givenName").Value & "<br>")
            Response.Write("Apellido: " & result.Fields("sn").Value & "<br>")
            Response.Write("Correo: " & result.Fields("mail").Value & "<br>")
            Response.Write("Teléfono Móvil: " & result.Fields("mobile").Value & "<br>")
        Else
            Response.Write("<p>Usuario no encontrado.</p>")
        End If
        On Error GoTo 0
    End Sub

    Sub SaveUser(ldapPath, adminUser, adminPassword, username, firstName, lastName, email, mobilePhone, homePhone, password)
        On Error Resume Next
        Dim conn, user
        Set conn = GetObject(ldapPath)
        conn.Username = adminUser
        conn.Password = adminPassword

        Set user = conn.Create("user", "CN=" & username)
        user.Put "sAMAccountName", username
        user.Put "givenName", firstName
        user.Put "sn", lastName
        user.Put "mail", email
        If mobilePhone <> "" Then user.Put "mobile", mobilePhone
        If homePhone <> "" Then user.Put "homePhone", homePhone
        user.SetInfo()

        ' Establecer contraseña
        user.SetPassword(password)
        user.Put "userAccountControl", 512
        user.SetInfo()
        Response.Write("<p>Usuario guardado correctamente.</p>")
        On Error GoTo 0
    End Sub

    Sub DeleteUser(ldapPath, adminUser, adminPassword, username)
        On Error Resume Next
        Dim conn, user
        Set conn = GetObject(ldapPath)
        conn.Username = adminUser
        conn.Password = adminPassword

        Set user = conn.GetObject("user", "CN=" & username)
        user.Delete
        Response.Write("<p>Usuario eliminado correctamente.</p>")
        On Error GoTo 0
    End Sub
%>
</body>
</html>
