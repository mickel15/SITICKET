Imports System.IO

Public Class form1
    Inherits Form

    Private loginForm As LoginForm ' Referencia al formulario de login
    Private administradorMenu As ToolStripMenuItem
    Private itinerarioMenu As ToolStripMenuItem
    Private reporteMenu As ToolStripMenuItem
    Private ayudaMenu As ToolStripMenuItem

    Private Sub GestiónDeSoportes_Click(sender As Object, e As EventArgs)
        Dim adminForm As New AdminForm()
        adminForm.Show()
    End Sub

    Public Sub New()
        ' Mostrar el formulario de login primero
        loginForm = New LoginForm()
        If loginForm.ShowDialog() = DialogResult.OK Then
            ' Si el inicio de sesión fue exitoso, configurar la interfaz según el rol
            Me.Text = "Sistema SITICKET"
            Me.Size = New Size(600, 500)
            ConfigurarMenu(loginForm.Rol)
        Else
            ' Si el usuario no inició sesión, cerrar la aplicación
            Application.Exit()
        End If
    End Sub

    ' Configurar el menú de acuerdo al rol del usuario
    Private Sub ConfigurarMenu(rol As String)
        Dim menuStrip As New MenuStrip()
        Dim administraciónMenu As New ToolStripMenuItem("Administración")
        Dim itinerarioMenu As New ToolStripMenuItem("Itinerario")
        Dim reporteMenu As New ToolStripMenuItem("Reporte")
        Dim ayudaMenu As New ToolStripMenuItem("Ayuda")

        ' Configurar elementos del menú según el rol
        If rol = "Administrador" Then
            ' Administrador: Acceso completo
            administraciónMenu.DropDownItems.Add(New ToolStripMenuItem("Gestión de Soportes", Nothing, AddressOf GestiónDeSoportes_Click))
        ElseIf rol = "Usuario" Then
            ' Usuario: Solo acceso a ciertos módulos
            administraciónMenu.Enabled = False
            reporteMenu.Enabled = False ' Si no quieres que el usuario acceda al módulo de reporte
        End If

        ' Resto de menús
        itinerarioMenu.DropDownItems.Add(New ToolStripMenuItem("Itinerario", Nothing, AddressOf ItinerarioToolStripMenuItem_Click))
        ayudaMenu.DropDownItems.Add(New ToolStripMenuItem("Ayuda", Nothing, AddressOf AyudaToolStripMenuItem_Click))

        menuStrip.Items.AddRange(New ToolStripItem() {administraciónMenu, itinerarioMenu, reporteMenu, ayudaMenu})
        Me.Controls.Add(menuStrip)
        MainMenuStrip = menuStrip
    End Sub

    ' Evento para abrir el formulario de gestión de soportes


    ' Evento para Itinerario
    Private Sub ItinerarioToolStripMenuItem_Click(sender As Object, e As EventArgs)
        ' Solo abrir el formulario de Itinerario sin pasar por Administrador ni Usuario
        Dim itinerarioForm As New ItinerarioForm()
        itinerarioForm.Show()
    End Sub

    ' Evento para Reporte
    Private Sub ReporteToolStripMenuItem_Click(sender As Object, e As EventArgs)
        ' Solo abrir el formulario de Reporte sin pasar por Administrador ni Usuario
        Dim reporteForm As New ReporteForm()
        reporteForm.Show()
    End Sub

    ' Evento para Ayuda
    Private Sub AyudaToolStripMenuItem_Click(sender As Object, e As EventArgs)
        MessageBox.Show("Versión 1.0 - Sistema SITICKET", "Ayuda")
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    End Sub
End Class



Public Class AdminForm
    Inherits Form

    ' Controles
    Private dgvSoportes As DataGridView
    Private WithEvents btnAgregar As Button
    Private WithEvents btnEditar As Button
    Private WithEvents btnEliminar As Button
    Private WithEvents btnCerrar As Button
    Private WithEvents btnExportar As Button ' Botón para exportar a CSV

    Public Sub New()
        ' Configuración del formulario
        Me.Text = "Gestión de Soportes"
        Me.Size = New Size(700, 500)

        ' Configurar DataGridView
        dgvSoportes = New DataGridView() With {
            .Location = New Point(20, 20),
            .Size = New Size(640, 300),
            .AllowUserToAddRows = False,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
        }

        ' Agregar columnas al DataGridView
        dgvSoportes.Columns.Add("Id", "ID")
        dgvSoportes.Columns.Add("TipoSoporte", "Tipo de Soporte")
        dgvSoportes.Columns.Add("Detalle", "Detalle")
        dgvSoportes.Columns.Add("CostoInicial", "Costo Inicial")
        dgvSoportes.Columns.Add("CostoFinal", "Costo Final")
        dgvSoportes.Columns.Add("Estado", "Estado")

        ' Botón Agregar
        btnAgregar = New Button() With {
            .Text = "Agregar",
            .Location = New Point(20, 350),
            .Size = New Size(150, 40)
        }

        ' Botón Editar
        btnEditar = New Button() With {
            .Text = "Editar",
            .Location = New Point(200, 350),
            .Size = New Size(150, 40)
        }

        ' Botón Eliminar
        btnEliminar = New Button() With {
            .Text = "Eliminar",
            .Location = New Point(380, 350),
            .Size = New Size(150, 40)
        }

        ' Botón Cerrar
        btnCerrar = New Button() With {
            .Text = "Cerrar",
            .Location = New Point(560, 350),
            .Size = New Size(150, 40)
        }

        ' Botón Exportar a CSV
        btnExportar = New Button() With {
            .Text = "Exportar a CSV",
            .Location = New Point(20, 400),
            .Size = New Size(150, 40)
        }

        ' Agregar controles al formulario
        Me.Controls.AddRange(New Control() {dgvSoportes, btnAgregar, btnEditar, btnEliminar, btnCerrar, btnExportar})

        ' Asociar eventos
        AddHandler btnAgregar.Click, AddressOf btnAgregar_Click
        AddHandler btnEditar.Click, AddressOf btnEditar_Click
        AddHandler btnEliminar.Click, AddressOf btnEliminar_Click
        AddHandler btnCerrar.Click, AddressOf btnCerrar_Click
        AddHandler btnExportar.Click, AddressOf btnExportar_Click ' Asociar evento de exportación
    End Sub

    ' Evento para agregar un nuevo soporte
    Private Sub btnAgregar_Click(sender As Object, e As EventArgs)
        Dim formSoporte As New SoporteForm("Agregar")
        If formSoporte.ShowDialog() = DialogResult.OK Then
            ' Añadir el nuevo soporte a la tabla
            Dim row As DataGridViewRow = dgvSoportes.Rows(dgvSoportes.Rows.Add())
            row.Cells("Id").Value = dgvSoportes.Rows.Count
            row.Cells("TipoSoporte").Value = formSoporte.TipoSoporte
            row.Cells("Detalle").Value = formSoporte.Detalle
            row.Cells("CostoInicial").Value = formSoporte.CostoInicial
            row.Cells("CostoFinal").Value = formSoporte.CostoFinal
            row.Cells("Estado").Value = "En Progreso"
        End If
    End Sub

    ' Evento para editar un soporte existente
    Private Sub btnEditar_Click(sender As Object, e As EventArgs)
        If dgvSoportes.SelectedRows.Count = 0 Then
            MessageBox.Show("Seleccione un soporte para editar.", "Advertencia")
            Return
        End If

        Dim row As DataGridViewRow = dgvSoportes.SelectedRows(0)
        Dim formSoporte As New SoporteForm("Editar") With {
            .TipoSoporte = row.Cells("TipoSoporte").Value?.ToString(),
            .Detalle = row.Cells("Detalle").Value?.ToString(),
            .CostoInicial = Convert.ToDecimal(row.Cells("CostoInicial").Value),
            .CostoFinal = Convert.ToDecimal(row.Cells("CostoFinal").Value)
        }

        If formSoporte.ShowDialog() = DialogResult.OK Then
            ' Actualizar los datos en el DataGridView
            row.Cells("TipoSoporte").Value = formSoporte.TipoSoporte
            row.Cells("Detalle").Value = formSoporte.Detalle
            row.Cells("CostoInicial").Value = formSoporte.CostoInicial
            row.Cells("CostoFinal").Value = formSoporte.CostoFinal
        End If
    End Sub

    ' Evento para eliminar un soporte
    Private Sub btnEliminar_Click(sender As Object, e As EventArgs)
        If dgvSoportes.SelectedRows.Count = 0 Then
            MessageBox.Show("Seleccione un soporte para eliminar.", "Advertencia")
            Return
        End If

        If MessageBox.Show("¿Está seguro de que desea eliminar este soporte?", "Confirmación", MessageBoxButtons.YesNo) = DialogResult.Yes Then
            dgvSoportes.Rows.Remove(dgvSoportes.SelectedRows(0))
        End If
    End Sub

    ' Evento para cerrar el formulario
    Private Sub btnCerrar_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

    ' Evento para exportar los datos a un archivo CSV
    Private Sub btnExportar_Click(sender As Object, e As EventArgs)
        ' Crear el archivo CSV
        Dim saveFileDialog As New SaveFileDialog() With {
            .Filter = "CSV Files (*.csv)|*.csv",
            .FileName = "soportes.csv"
        }

        If saveFileDialog.ShowDialog() = DialogResult.OK Then
            Try
                Using writer As New StreamWriter(saveFileDialog.FileName, False, System.Text.Encoding.UTF8)
                    ' Escribir encabezados organizados
                    Dim headers = dgvSoportes.Columns.Cast(Of DataGridViewColumn)().
                                  Select(Function(c) c.HeaderText.Replace(",", ";"))
                    writer.WriteLine(String.Join(";", headers))

                    ' Escribir los datos de cada fila
                    For Each row As DataGridViewRow In dgvSoportes.Rows
                        If Not row.IsNewRow Then
                            Dim cells = row.Cells.Cast(Of DataGridViewCell)().
                                        Select(Function(c) If(c.Value?.ToString().Replace(",", ";"), ""))
                            writer.WriteLine(String.Join(";", cells))
                        End If
                    Next
                End Using

                MessageBox.Show("Datos exportados correctamente en formato organizado.", "Éxito")
            Catch ex As Exception
                MessageBox.Show($"Error al exportar los datos: {ex.Message}", "Error")
            End Try
        End If
    End Sub

End Class



Public Class ReporteForm
    Inherits Form ' Asegúrate de que ReporteForm herede de la clase Form

    ' Declaración de controles
    WithEvents btnGenerar As Button
    WithEvents btnCerrar As Button

    ' Evento cuando se carga el formulario
    Public Sub New()
        ' Configuración del formulario
        Me.Text = "Módulo de Reporte"
        Me.Size = New Size(400, 300) ' Ajusta el tamaño del formulario según sea necesario

        ' Crear el botón Generar Reporte dinámicamente
        btnGenerar = New Button()
        btnGenerar.Text = "Generar Reporte"
        btnGenerar.Size = New Size(150, 40)
        btnGenerar.Location = New Point(100, 100)
        Me.Controls.Add(btnGenerar)

        ' Crear el botón Cerrar
        btnCerrar = New Button()
        btnCerrar.Text = "Cerrar"
        btnCerrar.Size = New Size(150, 40)
        btnCerrar.Location = New Point(100, 150)
        Me.Controls.Add(btnCerrar)

        ' Vincular los eventos Click a los botones
        AddHandler btnGenerar.Click, AddressOf btnGenerar_Click
        AddHandler btnCerrar.Click, AddressOf btnCerrar_Click
    End Sub

    ' Evento cuando se hace clic en el botón Generar Reporte
    Private Sub btnGenerar_Click(sender As Object, e As EventArgs)
        ' Código para generar el reporte
        MessageBox.Show("Reporte generado correctamente.", "Información")
    End Sub

    ' Evento cuando se hace clic en el botón Cerrar
    Private Sub btnCerrar_Click(sender As Object, e As EventArgs)
        Me.Close() ' Cierra el formulario ReporteForm
    End Sub
End Class

Public Class ItinerarioForm
    Inherits Form

    ' Controles
    Private dgvSoportes As DataGridView
    Private calendario As MonthCalendar
    Private WithEvents btnCerrar As Button

    Public Sub New()
        ' Configuración del formulario
        Me.Text = "Gestión de Itinerarios"
        Me.Size = New Size(800, 600)

        ' Crear DataGridView para los soportes en stock
        dgvSoportes = New DataGridView() With {
            .Location = New Point(20, 20),
            .Size = New Size(500, 300),
            .AllowUserToAddRows = False,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
        }

        ' Agregar columnas al DataGridView
        dgvSoportes.Columns.Add("Id", "ID")
        dgvSoportes.Columns.Add("Nombre", "Nombre del Soporte")
        dgvSoportes.Columns.Add("Cantidad", "Cantidad")
        dgvSoportes.Columns.Add("Estado", "Estado")

        ' Llenar datos simulados para soportes en stock
        dgvSoportes.Rows.Add(1, "Soporte Técnico", 10, "Disponible")
        dgvSoportes.Rows.Add(2, "Mantenimiento Preventivo", 5, "Disponible")
        dgvSoportes.Rows.Add(3, "Revisión", 2, "En Uso")

        ' Crear MonthCalendar
        calendario = New MonthCalendar() With {
            .Location = New Point(550, 20),
            .MaxSelectionCount = 1
        }

        ' Botón Cerrar
        btnCerrar = New Button() With {
            .Text = "Cerrar",
            .Location = New Point(320, 350),
            .Size = New Size(150, 40)
        }

        ' Agregar controles al formulario
        Me.Controls.AddRange(New Control() {dgvSoportes, calendario, btnCerrar})

        ' Asociar evento al botón Cerrar
        AddHandler btnCerrar.Click, AddressOf btnCerrar_Click
    End Sub

    ' Evento para cerrar el formulario
    Private Sub btnCerrar_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub
End Class


Public Class SoporteForm
    Inherits Form

    ' Propiedades públicas para recoger datos
    Public Property TipoSoporte As String
    Public Property Detalle As String
    Public Property CostoInicial As Decimal
    Public Property CostoFinal As Decimal

    ' Controles
    Private txtTipoSoporte As TextBox
    Private txtDetalle As TextBox
    Private nudCostoInicial As NumericUpDown
    Private nudCostoFinal As NumericUpDown
    Private btnGuardar As Button
    Private btnCancelar As Button

    Public Sub New(operation As String)
        ' Configuración del formulario
        Me.Text = $"{operation} Soporte"
        Me.Size = New Size(400, 400)

        ' Configurar controles
        txtTipoSoporte = New TextBox() With {.Location = New Point(20, 50), .Width = 200}
        txtDetalle = New TextBox() With {.Location = New Point(20, 100), .Width = 200, .Height = 80, .Multiline = True}
        nudCostoInicial = New NumericUpDown() With {.Location = New Point(20, 200), .Maximum = 10000, .DecimalPlaces = 2}
        nudCostoFinal = New NumericUpDown() With {.Location = New Point(20, 250), .Maximum = 10000, .DecimalPlaces = 2}

        btnGuardar = New Button() With {.Text = "Guardar", .Location = New Point(50, 300)}
        btnCancelar = New Button() With {.Text = "Cancelar", .Location = New Point(150, 300)}

        ' Etiquetas
        Me.Controls.Add(New Label() With {.Text = "Tipo de Soporte", .Location = New Point(20, 30)})
        Me.Controls.Add(txtTipoSoporte)
        Me.Controls.Add(New Label() With {.Text = "Detalle", .Location = New Point(20, 80)})
        Me.Controls.Add(txtDetalle)
        Me.Controls.Add(New Label() With {.Text = "Costo Inicial", .Location = New Point(20, 180)})
        Me.Controls.Add(nudCostoInicial)
        Me.Controls.Add(New Label() With {.Text = "Costo Final", .Location = New Point(20, 230)})
        Me.Controls.Add(nudCostoFinal)
        Me.Controls.AddRange(New Control() {btnGuardar, btnCancelar})

        ' Eventos
        AddHandler btnGuardar.Click, AddressOf btnGuardar_Click
        AddHandler btnCancelar.Click, AddressOf btnCancelar_Click
    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs)
        TipoSoporte = txtTipoSoporte.Text
        Detalle = txtDetalle.Text
        CostoInicial = nudCostoInicial.Value
        CostoFinal = nudCostoFinal.Value

        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub btnCancelar_Click(sender As Object, e As EventArgs)
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
End Class


Public Class EquiposForm
    Inherits Form

    Private WithEvents btnAlta As Button
    Private WithEvents btnBaja As Button
    Private WithEvents btnCancelar As Button

    Public Sub New()
        ' Configuración del formulario
        Me.Text = "Gestión de Equipos"
        Me.Size = New Size(400, 300)

        ' Crear controles
        btnAlta = New Button() With {
            .Text = "Alta de Equipos",
            .Location = New Point(100, 50),
            .Size = New Size(200, 40)
        }

        btnBaja = New Button() With {
            .Text = "Baja de Equipos",
            .Location = New Point(100, 120),
            .Size = New Size(200, 40)
        }

        btnCancelar = New Button() With {
            .Text = "Cerrar",
            .Location = New Point(100, 190),
            .Size = New Size(200, 40)
        }

        ' Agregar controles
        Me.Controls.AddRange(New Control() {btnAlta, btnBaja, btnCancelar})

        ' Asociar eventos
        AddHandler btnAlta.Click, AddressOf btnAlta_Click
        AddHandler btnBaja.Click, AddressOf btnBaja_Click
        AddHandler btnCancelar.Click, AddressOf btnCancelar_Click
    End Sub

    Private Sub btnAlta_Click(sender As Object, e As EventArgs)
        MessageBox.Show("Equipo dado de alta correctamente.", "Éxito")
    End Sub

    Private Sub btnBaja_Click(sender As Object, e As EventArgs)
        MessageBox.Show("Equipo dado de baja correctamente.", "Éxito")
    End Sub

    Private Sub btnCancelar_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub



End Class