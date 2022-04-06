Public Class FrmEjercicio7

    Const B = 0
    Const KB = 1
    Const MB = 2
    Const GB = 3
    Const Bi = -1

    Const sTitleMensaje = "Conversión de Unidades de Almacenamiento "
    Private Sub BtnLimpiar_Click(sender As Object, e As EventArgs) Handles BtnLimpiar.Click
        Limpiar()
    End Sub

    Private Sub Limpiar()
        Me.TxtResultado.Text = ""
        Me.TxtUnidades.Text = ""
        CBoxA.SelectedIndex = 0
        CBoxDe.SelectedIndex = 0
    End Sub

    Private Sub BtnConvertir_Click(sender As Object, e As EventArgs) Handles BtnConvertir.Click
        Dim StrResult As String = ""
        Dim StrMensaje As String = ""
        Dim iUnidades As Int32 = 0

        iUnidades = Convert.ToInt32(IIf(TxtUnidades.Text.Trim() = "", "0", TxtUnidades.Text.Trim()))
        If iUnidades > 0 And CBoxA.Text.Length > 0 And CBoxDe.Text.Length > 0 Then

            StrResult = Conversion(iUnidades)
            If StrResult.Trim().Length > 0 Then

                StrMensaje = "" & TxtUnidades.Text.Trim & " - " & Me.CBoxDe.Text & vbCrLf
                StrMensaje += "     Equivale a : " & vbCrLf & " " & StrResult

                Me.TxtResultado.Text = StrMensaje
            Else
                MsgBox("Se han presentado inconsistencias al momento de la conversión", vbExclamation, sTitleMensaje)
            End If
        Else
            MsgBox("Seleccionar Unidades e Ingresar el valor para la Conversión", vbInformation, sTitleMensaje)
        End If


    End Sub

    Private Function Conversion(ByVal Unid As Int32) As String
        Dim sResult As String = ""
        Dim ArrDes As String()
        Try
            Dim UniOrigen As String = CBoxDe.Text.Trim()
            Dim UniDestino As String = CBoxA.Text.Trim()
            Dim ValorBytes As Double = 0
            Dim ValorDestino As Double = 0
            Dim iDes As Int32 = 0
            Dim iOrg As Int32 = 0
            Dim DestinoSup As Boolean = False
            Dim ConverIgual As Boolean = False

            'Se identifican los elementos de Almacenamientos seleccionado en el Origen "DE"
            Select Case UniOrigen
                Case "BYTES"
                    iOrg = B
                Case "KILOBYTES - KB"
                    iOrg = KB
                Case "MEGABYTES - MG"
                    iOrg = MB
                Case "GIGABYTES - GB"
                    iOrg = GB
                Case "Bits"
                    iOrg = Bi
            End Select

            'Se identifican los elementos de Almacenamientos seleccionado para Destino  "A"
            Select Case UniDestino
                Case "BYTES"
                    iDes = B
                Case "KILOBYTES - KB"
                    iDes = KB
                Case "MEGABYTES - MG"
                    iDes = MB
                Case "GIGABYTES - GB"
                    iDes = GB
                Case "Bits"
                    iDes = Bi
            End Select

            'Se identifica si la Conversión es ascendente o descendente
            If (iOrg = iDes) Then
                ConverIgual = True
            ElseIf iOrg < iDes Then
                DestinoSup = True
            ElseIf iOrg > iDes Then
                DestinoSup = False
            End If

            If Not ConverIgual Then
                If DestinoSup Then
                    ValorBytes = 0
                    If iOrg = B Then
                        ValorBytes = Unid
                    ElseIf iOrg = Bi Then
                        ValorBytes = Unid / 8
                    Else
                        ValorBytes = Unid * (1024 ^ iOrg)
                    End If
                Else
                    ValorBytes = Unid * (1024 ^ iOrg)
                End If

                If iDes <> Bi Then
                    ValorDestino = ValorBytes / (1024 ^ iDes)
                Else
                    ValorDestino = ValorBytes * 8
                End If
            Else
                ValorDestino = Unid
            End If

            If ValorDestino < 1 And ValorDestino > 0 Then
                sResult = Format(ValorDestino, "##,##0.0000000")
            Else
                sResult = Format(ValorDestino, "##,##0")
            End If

            If iDes <> Bi And iDes <> B Then
                ArrDes = UniDestino.Split("-")
                sResult += " " & ArrDes(1).Trim()
            Else
                sResult += "  " & UniDestino.ToUpper()
            End If

        Catch ex As Exception
            sResult = ""
        End Try

        Return sResult
    End Function
    Private Sub TxtUnidades_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtUnidades.KeyPress
        If Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub CBoxDe_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CBoxDe.KeyPress
        e.Handled = True
    End Sub

    Private Sub CBoxDe_KeyDown(sender As Object, e As KeyEventArgs) Handles CBoxDe.KeyDown
        e.Handled = True
    End Sub

    Private Sub FrmEjercicio7_Load(sender As Object, e As EventArgs) Handles Me.Load
        CBoxA.SelectedIndex = 0
        CBoxDe.SelectedIndex = 0
    End Sub
End Class