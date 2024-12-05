Imports System.Data.SqlClient

Public Class LoginForm
    Inherits Form

    ' Controles para el formulario de Login
    Private txtUsuario As TextBox
    Private txtContraseña As TextBox
    Private btnIniciar As Button
    Private lblMensaje As Label
    Private btnRegistrar As Button
    Private btnCancelar As Button

    Public Property Rol As String = "" ' Almacenará el rol: "Administrador" o "Usuario"

    ' Definir la cadena de conexión a la base de datos
    Private connectionString As String = "Data Source=DESKTOP-B2C5UKG\SQLEXPRESS;Initial Catalog=SISTICKET;Integrated Security=True"

    Public Sub New()
        ' Configuración del formulario de Login
        Me.Text = "Inicio de Sesión - SISTICKET"
        Me.Size = New Size(400, 300)

        ' Crear y configurar los controles
        lblMensaje = New Label() With {
            .Location = New Point(20, 20),
            .Text = "Por favor, ingrese su usuario y contraseña."
        }

        txtUsuario = New TextBox() With {.Location = New Point(20, 50), .Width = 200}
        txtContraseña = New TextBox() With {.Location = New Point(20, 100), .Width = 200, .PasswordChar = "*"c}
        btnIniciar = New Button() With {.Text = "Iniciar Sesión", .Location = New Point(20, 150), .Size = New Size(200, 40)}
        btnRegistrar = New Button() With {.Text = "Registrar Usuario", .Location = New Point(20, 200), .Size = New Size(200, 40)}
        btnCancelar = New Button() With {.Text = "Cancelar", .Location = New Point(20, 250), .Size = New Size(200, 40)}

        ' Agregar los "placeholders" a los TextBox
        SetPlaceholderText(txtUsuario, "Usuario")
        SetPlaceholderText(txtContraseña, "Contraseña")

        ' Agregar controles al formulario
        Me.Controls.Add(lblMensaje)
        Me.Controls.Add(txtUsuario)
        Me.Controls.Add(txtContraseña)
        Me.Controls.Add(btnIniciar)
        Me.Controls.Add(btnRegistrar)
        Me.Controls.Add(btnCancelar)

        ' Asociar eventos
        AddHandler btnIniciar.Click, AddressOf btnIniciar_Click
        AddHandler btnRegistrar.Click, AddressOf btnRegistrar_Click
        AddHandler btnCancelar.Click, AddressOf btnCancelar_Click
    End Sub

    ' Método para agregar texto de placeholder
    Private Sub SetPlaceholderText(ByVal textBox As TextBox, ByVal placeholder As String)
        textBox.Text = placeholder
        textBox.ForeColor = Color.Gray

        ' Evento cuando el usuario hace clic en el campo
        AddHandler textBox.Enter, Sub(sender, e)
                                      If textBox.Text = placeholder Then
                                          textBox.Text = ""
                                          textBox.ForeColor = Color.Black
                                      End If
                                  End Sub

        ' Evento cuando el usuario sale del campo
        AddHandler textBox.Leave, Sub(sender, e)
                                      If String.IsNullOrEmpty(textBox.Text) Then
                                          textBox.Text = placeholder
                                          textBox.ForeColor = Color.Gray
                                      End If
                                  End Sub
    End Sub

    ' Evento al hacer clic en el botón Iniciar Sesión
    Private Sub btnIniciar_Click(sender As Object, e As EventArgs)
        ' Verificar si el usuario existe y si las credenciales son correctas
        Dim query As String = "SELECT rol FROM Usuarios WHERE nombreUsuario = @usuario AND contrasena = @contrasena"

        ' Manejo de errores
        Try
            ' Crear la conexión a la base de datos
            Using conn As New SqlConnection(connectionString)
                ' Crear el comando SQL
                Using cmd As New SqlCommand(query, conn)
                    ' Agregar los parámetros para evitar inyecciones SQL
                    cmd.Parameters.AddWithValue("@usuario", txtUsuario.Text)
                    cmd.Parameters.AddWithValue("@contrasena", txtContraseña.Text)

                    ' Abrir la conexión a la base de datos
                    conn.Open()

                    ' Ejecutar la consulta y obtener el rol
                    Dim rol As Object = cmd.ExecuteScalar()

                    ' Verificar si el rol fue encontrado
                    If rol IsNot Nothing Then
                        ' Si el rol existe, asignarlo
                        Me.Rol = rol.ToString()

                        ' Cerrar el formulario de login y mostrar el formulario principal
                        Me.DialogResult = DialogResult.OK
                        Me.Close()
                    Else
                        ' Si las credenciales son incorrectas, mostrar mensaje
                        MessageBox.Show("Usuario o contraseña incorrectos.")
                    End If
                End Using
            End Using
        Catch ex As Exception
            ' Si hay un error de conexión o de ejecución, mostrar el mensaje de error
            MessageBox.Show("Error al conectar con la base de datos: " & ex.Message)
        End Try
    End Sub

    ' Evento al hacer clic en el botón Registrar Usuario
    Private Sub btnRegistrar_Click(sender As Object, e As EventArgs)
        ' Confirmar si realmente desea registrar el nuevo usuario
        Dim result As DialogResult = MessageBox.Show("¿Estás seguro de que deseas registrar un nuevo usuario?", "Confirmar Registro", MessageBoxButtons.YesNo)

        If result = DialogResult.Yes Then
            RegistrarUsuario()
        End If
    End Sub

    ' Evento al hacer clic en el botón Cancelar
    Private Sub btnCancelar_Click(sender As Object, e As EventArgs)
        ' Cerrar el formulario de login sin hacer nada
        Me.Close()
    End Sub

    ' Método para registrar un nuevo usuario
    Private Sub RegistrarUsuario()
        ' Mostrar un formulario de registro de usuario o crear un nuevo usuario directamente
        Dim query As String = "INSERT INTO Usuarios (nombreUsuario, contrasena, rol) VALUES (@usuario, @contrasena, @rol)"

        ' Manejo de errores
        Try
            ' Crear la conexión a la base de datos
            Using conn As New SqlConnection(connectionString)
                ' Crear el comando SQL
                Using cmd As New SqlCommand(query, conn)
                    ' Agregar los parámetros para evitar inyecciones SQL
                    cmd.Parameters.AddWithValue("@usuario", txtUsuario.Text)
                    cmd.Parameters.AddWithValue("@contrasena", txtContraseña.Text) ' Asegúrate de cifrar la contraseña
                    cmd.Parameters.AddWithValue("@rol", "Usuario") ' Asignamos rol "Usuario" por defecto

                    ' Abrir la conexión a la base de datos
                    conn.Open()

                    ' Ejecutar el comando para insertar el nuevo usuario
                    cmd.ExecuteNonQuery()

                    MessageBox.Show("Usuario registrado exitosamente. Inicie sesión con su nuevo usuario.")
                End Using
            End Using
        Catch ex As Exception
            ' Si hay un error en el registro, mostrar el mensaje de error
            MessageBox.Show("Error al registrar el usuario: " & ex.Message)
        End Try
    End Sub
End Class
