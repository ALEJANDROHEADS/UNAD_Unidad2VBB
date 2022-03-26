Imports System.Math

Public Class FrmEjercicio5

    Const TituloApp = "Herramienta Cálculo de Áreas .."

    Private Sub FrmEjercicio5_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.CBoxTrianguloMedida.SelectedIndex = 0
        Me.CBoxRectangulo.SelectedIndex = 0
        Me.CBoxRombo.SelectedIndex = 0
        Me.CBoxTrapecio.SelectedIndex = 0
        Me.CBoxRomboide.SelectedIndex = 0

    End Sub
    Private Sub BtnCerrar_Click(sender As Object, e As EventArgs) Handles BtnCerrar.Click
        If (MsgBox("¿Confirma que se desea cerrar la ventana?", MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, "Herramienta Cálculo de Áreas ..") = MsgBoxResult.Yes) Then
            Me.Close()
        End If
    End Sub
    Private Function SeleccionMedida(ByVal Codigo As Integer) As String
        Dim sMedida As String = ""
        Select Case Codigo
            Case 1
                sMedida = "Cm"
            Case 2
                sMedida = "Mtr"
            Case 3
                sMedida = "Km"
        End Select
        Return sMedida
    End Function

#Region "VALIDACION INGRESO NUMÉRICO"
    '<<<═══════════════════════════════════════════════════════════════════════════════════════
    '<<< (CesarAlejandroCabezas) 20/marzo/2022//  Fase 3 - Diseño de Modulos  
    '>>> Se Valida que el Dato Ingresado se Numerico 
    Private Sub TxtTrianguloAltura_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtTrianguloAltura.KeyPress
        If Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TxtTrianguloBase_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtTrianguloBase.KeyPress
        If Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TxtTrapecioAltura_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtTrapecioAltura.KeyPress
        If Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TxtTrapecioBaseSup_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtTrapecioBaseSup.KeyPress
        If Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TxtTrapecioInf_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtTrapecioInf.KeyPress
        If Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TxtRectanguloAltura_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtRectanguloAltura.KeyPress
        If Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TxtRectanguloBase_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtRectanguloBase.KeyPress
        If Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TxtRomboD1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtRomboD1.KeyPress
        If Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TxtRomboD2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtRomboD2.KeyPress
        If Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TxtRomboideAltura_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtRomboideAltura.KeyPress
        If Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TxtRomboideBase_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtRomboideBase.KeyPress
        If Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub
    '>>>═══════════════════════════════════════════════════════════════════════════════════════
#End Region

#Region "TRIANGULO... "
    Private Sub BtnTrianguloLimpiar_Click(sender As Object, e As EventArgs) Handles BtnTrianguloLimpiar.Click
        If (MsgBox("¿Deseas borrar los datos ingresados ?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, TituloApp) = MsgBoxResult.Yes) Then
            Me.CBoxTrianguloMedida.SelectedIndex = 0
            Me.TxtTrianguloAltura.Text = ""
            Me.TxtTrianguloBase.Text = ""
            Me.TxtTrianguloResul.Text = ""
            Me.CBoxTrianguloMedida.Select()
        End If
    End Sub
    Private Sub BtnTrianguloCalcular_Click(sender As Object, e As EventArgs) Handles BtnTrianguloCalcular.Click

        Dim sAltura As String = Me.TxtTrianguloAltura.Text.Trim()
        Dim sBase As String = Me.TxtTrianguloBase.Text.Trim()
        Dim Medida As Int16 = Me.CBoxTrianguloMedida.SelectedIndex
        Dim dArea As Double = 0
        Dim sMedida As String = ""


        If (sAltura.Length = 0 Or sBase.Length = 0 Or Medida = 0) Then
            MsgBox("Todos los datos son obligatorios; Verificar los datos " & vbCrLf & "e intentalo nuevamente!.", MsgBoxStyle.Exclamation, "Herramienta Cálculo de Áreas ..")
            Exit Sub
        Else
            sMedida = SeleccionMedida(Medida)

            dArea = (Convert.ToDouble(sBase) * Convert.ToDouble(sAltura)) / 2
            Me.TxtTrianguloResul.Text = "Área Calculada :" & vbCrLf & " * " & Convert.ToString(dArea) & " " & sMedida & Char.ConvertFromUtf32(178)
        End If

    End Sub

#End Region

#Region "TRAPECIO..."
    Private Sub BtnTrapecioCalcular_Click(sender As Object, e As EventArgs) Handles BtnTrapecioCalcular.Click
        Dim sAltura As String = Me.TxtTrapecioAltura.Text.Trim()
        Dim sBaseSup As String = Me.TxtTrapecioBaseSup.Text.Trim()
        Dim sBaseinf As String = Me.TxtTrapecioInf.Text.Trim()

        Dim Medida As Int16 = Me.CBoxTrapecio.SelectedIndex
        Dim dArea As Double = 0
        Dim sMedida As String = ""


        If (sAltura.Length = 0 Or sBaseSup.Length = 0 Or sBaseinf.Length = 0 Or Medida = 0) Then
            MsgBox("Todos los datos son obligatorios; Verificar los datos " & vbCrLf & "e intentalo nuevamente!.", MsgBoxStyle.Exclamation, TituloApp)
            Exit Sub
        Else
            sMedida = SeleccionMedida(Medida)

            dArea = ((Convert.ToDouble(sBaseSup) + Convert.ToDouble(sBaseinf)) / 2) * Convert.ToDouble(sAltura)
            Me.TxtTrapecioResul.Text = "Área Calculada :" & vbCrLf & " * " & Convert.ToString(dArea) & " " & sMedida & Char.ConvertFromUtf32(178)
        End If

    End Sub
    Private Sub BtnTrapecioLimpiar_Click(sender As Object, e As EventArgs) Handles BtnTrapecioLimpiar.Click
        If (MsgBox("¿Deseas borrar los datos ingresados ?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, TituloApp) = MsgBoxResult.Yes) Then
            Me.CBoxTrapecio.SelectedIndex = 0
            Me.TxtTrapecioAltura.Text = ""
            Me.TxtTrapecioInf.Text = ""
            Me.TxtTrapecioBaseSup.Text = ""
            Me.TxtTrapecioResul.Text = ""
            Me.CBoxTrapecio.Select()
        End If
    End Sub
#End Region

#Region "RECTANGULO..."

    Private Sub BtnRectanguloCalcular_Click(sender As Object, e As EventArgs) Handles BtnRectanguloCalcular.Click
        Dim sAltura As String = Me.TxtRectanguloAltura.Text.Trim()
        Dim sBase As String = Me.TxtRectanguloBase.Text.Trim()
        Dim Medida As Int16 = Me.CBoxRectangulo.SelectedIndex
        Dim dArea As Double = 0
        Dim sMedida As String = ""


        If (sAltura.Length = 0 Or sBase.Length = 0 Or Medida = 0) Then
            MsgBox("Todos los datos son obligatorios; Verificar los datos " & vbCrLf & "e intentalo nuevamente!.", MsgBoxStyle.Exclamation, TituloApp)
            Exit Sub
        Else
            sMedida = SeleccionMedida(Medida)

            dArea = (Convert.ToDouble(sBase) * Convert.ToDouble(sAltura))
            Me.TxtRectanguloResult.Text = "Área Calculada :" & vbCrLf & " * " & Convert.ToString(dArea) & " " & sMedida & Char.ConvertFromUtf32(178)
        End If

    End Sub

    Private Sub BtnRectanguloLimpiar_Click(sender As Object, e As EventArgs) Handles BtnRectanguloLimpiar.Click
        If (MsgBox("¿Deseas borrar los datos ingresados ?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, TituloApp) = MsgBoxResult.Yes) Then
            Me.CBoxRectangulo.SelectedIndex = 0
            Me.TxtRectanguloAltura.Text = ""
            Me.TxtRectanguloBase.Text = ""
            Me.TxtRectanguloResult.Text = ""
            Me.CBoxRectangulo.Select()
        End If
    End Sub
#End Region

#Region "ROMBO"
    Private Sub BtnRomboCalcular_Click(sender As Object, e As EventArgs) Handles BtnRomboCalcular.Click
        Dim sDiagonal1 As String = Me.TxtRomboD1.Text.Trim()
        Dim sDiagonal2 As String = Me.TxtRomboD2.Text.Trim()
        Dim Medida As Int16 = Me.CBoxRombo.SelectedIndex
        Dim dArea As Double = 0
        Dim sMedida As String = ""


        If (sDiagonal1.Length = 0 Or sDiagonal2.Length = 0 Or Medida = 0) Then
            MsgBox("Todos los datos son obligatorios; Verificar los datos " & vbCrLf & "e intentalo nuevamente!.", MsgBoxStyle.Exclamation, TituloApp)
            Exit Sub
        Else
            sMedida = SeleccionMedida(Medida)

            dArea = (Convert.ToDouble(sDiagonal1) * Convert.ToDouble(sDiagonal2)) / 2
            Me.TxtRomboResult.Text = "Área Calculada :" & vbCrLf & " * " & Convert.ToString(dArea) & " " & sMedida & Char.ConvertFromUtf32(178)
        End If
    End Sub
    Private Sub BtnRomboLimpiar_Click(sender As Object, e As EventArgs) Handles BtnRomboLimpiar.Click
        If (MsgBox("¿Deseas borrar los datos ingresados ?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, TituloApp) = MsgBoxResult.Yes) Then
            Me.CBoxRombo.SelectedIndex = 0
            Me.TxtRomboD1.Text = ""
            Me.TxtRomboD2.Text = ""
            Me.TxtRomboResult.Text = ""
            Me.CBoxRombo.Select()
        End If
    End Sub
#End Region

#Region "ROMBOIDE"
    Private Sub BtnRomboideCalcular_Click(sender As Object, e As EventArgs) Handles BtnRomboideCalcular.Click
        Dim sAltura As String = Me.TxtRomboideAltura.Text.Trim()
        Dim sBase As String = Me.TxtRomboideBase.Text.Trim()
        Dim Medida As Int16 = Me.CBoxRomboide.SelectedIndex
        Dim dArea As Double = 0
        Dim sMedida As String = ""


        If (sAltura.Length = 0 Or sBase.Length = 0 Or Medida = 0) Then
            MsgBox("Todos los datos son obligatorios; Verificar los datos " & vbCrLf & "e intentalo nuevamente!.", MsgBoxStyle.Exclamation, TituloApp)
            Exit Sub
        Else
            sMedida = SeleccionMedida(Medida)

            dArea = (Convert.ToDouble(sBase) * Convert.ToDouble(sAltura))
            Me.TxtRomboideResul.Text = "Área Calculada :" & vbCrLf & " * " & Convert.ToString(dArea) & " " & sMedida & Char.ConvertFromUtf32(178)
        End If
    End Sub

    Private Sub BtnRomboideLimpiar_Click(sender As Object, e As EventArgs) Handles BtnRomboideLimpiar.Click
        If (MsgBox("¿Deseas borrar los datos ingresados ?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, TituloApp) = MsgBoxResult.Yes) Then
            Me.TxtRomboideAltura.Text = ""
            Me.TxtRomboideBase.Text = ""
            Me.TxtRomboideResul.Text = ""

            Me.CBoxRomboide.SelectedIndex = 0
            Me.CBoxRomboide.Select()
        End If
    End Sub
#End Region

End Class