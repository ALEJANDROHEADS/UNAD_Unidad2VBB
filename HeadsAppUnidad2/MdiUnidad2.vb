
Public Class MdiUnidad2


    '<<<═══════════════════════════════════════════════════════════════════════════════════════
    '<<< (CesarAlejandroCabezas) 20/marzo/2022//  Fase 3 - Diseño de Modulos  
    '<<< Declaración de colores constantes para Reutilización.
    Dim ColorFondoMove = Color.FromArgb(234, 233, 239)
    Dim ColorTextMove = Color.FromArgb(0, 0, 64)
    Dim ColorTextInAct = Color.FromArgb(140, 0, 0)
    Dim ColorFondoOrg = Color.FromArgb(0, 0, 64)
    Dim ColorTextOrg = Color.White
    '>>>═══════════════════════════════════════════════════════════════════════════════════════

#Region "MenuLateral"
    '<<<═══════════════════════════════════════════════════════════════════════════════════════
    '<<< (CesarAlejandroCabezas) 20/marzo/2022//  Fase 3 - Diseño de Modulos  
    ' Visualización y personalización Menu Lateral izquierdo 
    Private Sub BtnMenuTema1_Click(sender As Object, e As EventArgs) Handles BtnMenuTema1.Click
        VisualizarSubMenu(PnlSubMenuTema1)
    End Sub
    Private Sub BtnMenuTema2_Click(sender As Object, e As EventArgs) Handles BtnMenuTema2.Click
        VisualizarSubMenu(PnlSubMenuTema2)
    End Sub
    Private Sub BtnMenuTema3_Click(sender As Object, e As EventArgs) Handles BtnMenuTema3.Click
        VisualizarSubMenu(PnlSubMenuTema3)
    End Sub
    Private Sub PersonalizaDiseño()
        PnlSubMenuTema1.Visible = False
        PnlSubMenuTema2.Visible = False
        PnlSubMenuTema3.Visible = False
    End Sub
    Private Sub OcultarSubMenu()
        If PnlSubMenuTema1.Visible Then PnlSubMenuTema1.Visible = False
        If PnlSubMenuTema2.Visible Then PnlSubMenuTema2.Visible = False
        If PnlSubMenuTema3.Visible Then PnlSubMenuTema3.Visible = False
    End Sub
    Private Sub VisualizarSubMenu(ByVal SubMenu As Panel)
        If SubMenu.Visible = False Then
            OcultarSubMenu()
            SubMenu.Visible = True
        Else
            SubMenu.Visible = False
        End If
    End Sub
    '<<<═══════════════════════════════════════════════════════════════════════════════════════
    '<<< (CesarAlejandroCabezas) 20/marzo/2022//  Fase 3 - Diseño de Modulos  
    '<<< Comportamiento colores y fondo al parar el cursor sobre ellos
    Private Sub BtnEjericio5_MouseMove(sender As Object, e As MouseEventArgs) Handles BtnEjericio5.MouseMove
        BtnEjericio5.BackColor = ColorFondoMove
        BtnEjericio5.ForeColor = ColorTextMove
    End Sub
    Private Sub BtnEjericio5_MouseLeave(sender As Object, e As EventArgs) Handles BtnEjericio5.MouseLeave
        BtnEjericio5.BackColor = ColorFondoOrg
        BtnEjericio5.ForeColor = ColorTextOrg
    End Sub
    '>>>═══════════════════════════════════════════════════════════════════════════════════════
    Private Sub BtnEjericio7_MouseMove(sender As Object, e As MouseEventArgs) Handles BtnEjericio7.MouseMove
        BtnEjericio7.BackColor = ColorFondoMove
        BtnEjericio7.ForeColor = ColorTextMove
    End Sub
    Private Sub BtnEjericio7_MouseLeave(sender As Object, e As EventArgs) Handles BtnEjericio7.MouseLeave
        BtnEjericio7.BackColor = ColorFondoOrg
        BtnEjericio7.ForeColor = ColorTextOrg
    End Sub
    '>>>═══════════════════════════════════════════════════════════════════════════════════════
    Private Sub BtnEjericio13_MouseMove(sender As Object, e As MouseEventArgs) Handles BtnEjericio13.MouseMove
        BtnEjericio13.BackColor = ColorFondoMove
        BtnEjericio13.ForeColor = ColorTextInAct
    End Sub
    Private Sub BtnEjericio13_MouseLeave(sender As Object, e As EventArgs) Handles BtnEjericio13.MouseLeave
        BtnEjericio13.BackColor = ColorFondoOrg
        BtnEjericio13.ForeColor = ColorTextOrg
    End Sub
    '>>>═══════════════════════════════════════════════════════════════════════════════════════
    Private Sub BtnEjericio15_MouseMove(sender As Object, e As MouseEventArgs) Handles BtnEjericio15.MouseMove
        BtnEjericio15.BackColor = ColorFondoMove
        BtnEjericio15.ForeColor = ColorTextInAct
    End Sub
    Private Sub BtnEjericio15_MouseLeave(sender As Object, e As EventArgs) Handles BtnEjericio15.MouseLeave
        BtnEjericio15.BackColor = ColorFondoOrg
        BtnEjericio15.ForeColor = ColorTextOrg
    End Sub
    '>>>═══════════════════════════════════════════════════════════════════════════════════════
    Private Sub BtnEjericio19_MouseMove(sender As Object, e As MouseEventArgs) Handles BtnEjericio19.MouseMove
        BtnEjericio19.BackColor = ColorFondoMove
        BtnEjericio19.ForeColor = ColorTextInAct
    End Sub
    Private Sub BtnEjericio19_MouseLeave(sender As Object, e As EventArgs) Handles BtnEjericio19.MouseLeave
        BtnEjericio19.BackColor = ColorFondoOrg
        BtnEjericio19.ForeColor = ColorTextOrg
    End Sub
    '>>>═══════════════════════════════════════════════════════════════════════════════════════
    Private Sub BtnEjericio20_MouseMove(sender As Object, e As MouseEventArgs) Handles BtnEjericio20.MouseMove
        BtnEjericio20.BackColor = ColorFondoMove
        BtnEjericio20.ForeColor = ColorTextInAct
    End Sub
    Private Sub BtnEjericio20_MouseLeave(sender As Object, e As EventArgs) Handles BtnEjericio20.MouseLeave
        BtnEjericio20.BackColor = ColorFondoOrg
        BtnEjericio20.ForeColor = ColorTextOrg
    End Sub
    '>>>═══════════════════════════════════════════════════════════════════════════════════════
#End Region

    Private Sub MdiUnidad2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Me.WindowState = FormWindowState.Maximized

        PersonalizaDiseño()

    End Sub
    Private Sub MensajeItemInActivo(ByVal NumEj As String)
        MsgBox("El Ejercicio ( " & NumEj & " ) Se Encuentra Temporalmente In-Activo", MsgBoxStyle.Information, "Fase - 3 Visual Basic ")
        OcultarSubMenu()
    End Sub
    Private Sub BtnEjericio13_Click(sender As Object, e As EventArgs) Handles BtnEjericio13.Click
        MensajeItemInActivo("13")
        OcultarSubMenu()
    End Sub
    Private Sub BtnEjericio15_Click(sender As Object, e As EventArgs) Handles BtnEjericio15.Click
        MensajeItemInActivo("15")
        OcultarSubMenu()
    End Sub
    Private Sub BtnEjericio19_Click(sender As Object, e As EventArgs) Handles BtnEjericio19.Click
        MensajeItemInActivo("19")
        OcultarSubMenu()
    End Sub
    Private Sub BtnEjericio20_Click(sender As Object, e As EventArgs) Handles BtnEjericio20.Click
        MensajeItemInActivo("20")
        Me.StatusBarInf.BackColor = ColorFondoMove
        OcultarSubMenu()
    End Sub

    Private FormActivo As Form = Nothing

    Private Sub AbrirFormulario(ByVal FormChild As Form)

        If (IsNothing(FormActivo) = False) Then
            If (FormActivo.Text.TrimEnd.Length > 0) Then
                If (FormActivo.Name.Trim() <> FormChild.Name.Trim()) Then
                    If (MsgBox("Se desea Cerrar el Formulario Actual ( " & FormActivo.Text & ")", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Fase - 3  Visual Basic ") = MsgBoxResult.Yes) Then
                        FormActivo.Close()
                    Else
                        Exit Sub
                    End If
                Else
                    Exit Sub
                End If
            Else
                FormActivo.Close()
            End If
        End If

        FormActivo = FormChild
        FormChild.FormBorderStyle = FormBorderStyle.None
        FormChild.Dock = DockStyle.Fill
        FormChild.TopLevel = False
        PanelForm.Dock = DockStyle.Fill
        PanelForm.Controls.Add(FormChild)
        PanelForm.Tag = FormChild
        FormChild.BringToFront()
        FormChild.Show()
    End Sub

    Private Sub BtnEjericio5_Click(sender As Object, e As EventArgs) Handles BtnEjericio5.Click
        AbrirFormulario(FrmEjercicio5)
    End Sub

    Private Sub BtnEjericio7_Click(sender As Object, e As EventArgs) Handles BtnEjericio7.Click
        AbrirFormulario(FrmEjercicio7)
    End Sub
End Class