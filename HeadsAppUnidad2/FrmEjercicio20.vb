Public Class FrmEjercicio20

    Const TituloForm = "Sumatoria de Matrices VBB."
    Dim Matriz22 As Boolean
    Private Sub FrmEjercicio20_Load(sender As Object, e As EventArgs) Handles Me.Load
        '>>>═══════════════════════════════════════════════════════════════════════════════════════
        '>>> Se inician los campos activando inicialmente la Matriz 2 x 2 
        Matriz22 = True
        GBoxMatriz22.Visible = True
        GBoxMatriz33.Visible = False
        '>>>═══════════════════════════════════════════════════════════════════════════════════════
    End Sub

    Private Sub BtnProcesar_Click(sender As Object, e As EventArgs) Handles BtnProcesar.Click
        '>>>═══════════════════════════════════════════════════════════════════════════════════════
        '>>> Seberá existir la validacion para iniciar el Proceso "Sumatorioa de Matrices."
        If (ValidadorMatrices()) Then
            If sumatoriaMatrices() Then
                MsgBox("¡Se ha realizado la sumatoria satisfacoriemante! ", vbInformation, TituloForm)
            Else
                MsgBox("Ups Se han presentado Inconsistencias en el cálculo! ", vbCritical, TituloForm)
            End If
        Else
            MsgBox("¡Es Necesario Ingresar los valores en las dos Matrices! ", vbExclamation, TituloForm)
        End If
        '>>>═══════════════════════════════════════════════════════════════════════════════════════
    End Sub

    Private Sub BtnLimpiar_Click(sender As Object, e As EventArgs) Handles BtnLimpiar.Click
        '>>>═══════════════════════════════════════════════════════════════════════════════════════
        '>>> Se Realiza Confirmacion antes del borrado de la información 
        If (MsgBox("¿Se Desea hacer el borrado y limpieza de los campos ?", vbQuestion + vbYesNo, TituloForm) = vbYes) Then
            Limpiar()
        End If
        '>>>═══════════════════════════════════════════════════════════════════════════════════════
    End Sub

    Private Function sumatoriaMatrices() As Boolean
        Dim bResp As Boolean = False

        Try
            '>>>═══════════════════════════════════════════════════════════════════════════════════════
            '>>> Dependiendo de la selección se llama el proceso de la Sumatoria
            If Matriz22 Then
                bResp = Procesar_M22()
                If bResp Then
                    Me.GBoxResp_M22.Visible = True
                End If
            Else
                bResp = Procesar_M33()
                If bResp Then
                    Me.GBoxResp_M33.Visible = True
                End If
            End If
            '>>>═══════════════════════════════════════════════════════════════════════════════════════
        Catch ex As Exception
            bResp = False
        End Try
        Return bResp

    End Function

    Private Function Procesar_M22() As Boolean
        Dim bProcesado As Boolean = False
        Dim listResp As List(Of String) = New List(Of String)    ' Se crea Lista para Alamcenar Descripción de cada proceso 
        Dim SumaPunto As Int32 = 0              ' Variable Suma de los puntos de la Matriz 
        Dim ElementoA As String = ""           ' Variables DEfinicion de Nombres de los elemtos de la Matriz 
        Dim ElementoB As String = ""
        Dim ElemeResp As String = ""
        Dim sPto As String = ""

        Try
            '>> Se agrega Progres bar para simular el tiempo del proceso 
            ProgressBar.Value = 0
            ProgressBar.Visible = True
            '>>> Se Diseñan 2 Ciclos For Para que recorren Las Filas y las Columnas d ela Matriz
            listResp.Add("* Sumatorias de Matrices 2 x 2 ")
            For iCol As Integer = 0 To 1

                For iRow As Integer = 0 To 1
                    '>>>═══════════════════════════════════════════════════════════════════════════════════════
                    '>>> Se realiza REcuperación del Control por medio del Nombre
                    '>>> El cual contiene el punto exacto de las Matrices.
                    ElementoA = "TxtM22_A" & iCol.ToString() & iRow.ToString()
                    ElementoB = "TxtM22_B" & iCol.ToString() & iRow.ToString()
                    ElemeResp = "TxtResp22_" & iCol.ToString() & iRow.ToString()
                    Dim TextA As TextBox = Me.Controls.Find(ElementoA, True)(0)
                    Dim TextB As TextBox = Me.Controls.Find(ElementoB, True)(0)
                    Dim Resp As TextBox = Me.Controls.Find(ElemeResp, True)(0)

                    sPto = iCol & "," & iRow
                    '>>> Se Realiza la Sumatoria de lso 2 Puntos Matriz A y B 
                    SumaPunto = Convert.ToInt32(TextA.Text) + Convert.ToInt32(TextB.Text)
                    '>>> Se Agrega la Descripción del proceso para el resumen del Resultado 
                    listResp.Add("    >   [ A(" & sPto & ") =" & TextA.Text & " ] + [ B(" & sPto & ") =" & TextB.Text & " ] = " & SumaPunto.ToString)
                    '>>> Se le Asigna el Valor a la Matris de Resultados 
                    Resp.Text = SumaPunto.ToString()
                    '>>>═══════════════════════════════════════════════════════════════════════════════════════
                    Threading.Thread.Sleep(700)
                    ProgressBar.Value += 20
                Next
            Next

            '>>>═══════════════════════════════════════════════════════════════════════════════════════
            '>>> Se crea cursor para Agregar la Descripción en el Label de Resultados 
            Me.LblResultado.Text = listResp(0)
            For iFila As Integer = 1 To listResp.Count - 1
                LblResultado.Text += vbCrLf & listResp(iFila)
            Next
            '>>>═══════════════════════════════════════════════════════════════════════════════════════
            ProgressBar.Value = 100
            bProcesado = True
        Catch ex As Exception
            bProcesado = False
        End Try

        Return bProcesado
    End Function

    Private Function Procesar_M33() As Boolean
        Dim bProcesado As Boolean = False
        Dim listResp As List(Of String) = New List(Of String)    ' Se crea Lista para Alamcenar Descripción de cada proceso 
        Dim SumaPunto As Int32 = 0              ' Variable Suma de los puntos de la Matriz 
        Dim ElementoA As String = ""           ' Variables DEfinicion de Nombres de los elemtos de la Matriz 
        Dim ElementoB As String = ""
        Dim ElemeResp As String = ""
        Dim sPto As String = ""

        Try
            '>> Se agrega Progres bar para simular el tiempo del proceso 
            ProgressBar.Value = 0
            ProgressBar.Visible = True
            '>>> Se Diseñan 2 Ciclos For Para que recorren Las Filas y las Columnas d ela Matriz
            listResp.Add("* Sumatorias de Matrice 3 x 3 ")
            For iCol As Integer = 0 To 2

                For iRow As Integer = 0 To 2
                    '>>>═══════════════════════════════════════════════════════════════════════════════════════
                    '>>> Se realiza REcuperación del Control por medio del Nombre
                    '>>> El cual contiene el punto exacto de las Matrices.
                    ElementoA = "TxtM33_A" & iCol.ToString() & iRow.ToString()
                    ElementoB = "TxtM33_B" & iCol.ToString() & iRow.ToString()
                    ElemeResp = "TxtResp33_" & iCol.ToString() & iRow.ToString()
                    Dim TextA As TextBox = Me.Controls.Find(ElementoA, True)(0)
                    Dim TextB As TextBox = Me.Controls.Find(ElementoB, True)(0)
                    Dim Resp As TextBox = Me.Controls.Find(ElemeResp, True)(0)

                    sPto = iCol & "," & iRow
                    '>>> Se Realiza la Sumatoria de lso 2 Puntos Matriz A y B 
                    SumaPunto = Convert.ToInt32(TextA.Text) + Convert.ToInt32(TextB.Text)
                    '>>> Se Agrega la Descripción del proceso para el resumen del Resultado 
                    listResp.Add("       °   [ A(" & sPto & ") =" & TextA.Text & " ] + [ B(" & sPto & ") =" & TextB.Text & " ] = " & SumaPunto.ToString)
                    '>>> Se le Asigna el Valor a la Matris de Resultados 
                    Resp.Text = SumaPunto.ToString()
                    '>>>═══════════════════════════════════════════════════════════════════════════════════════
                    Threading.Thread.Sleep(500)
                    ProgressBar.Value += 10
                Next
            Next

            '>>>═══════════════════════════════════════════════════════════════════════════════════════
            '>>> Se crea cursor para Agregar la Descripción en el Label de Resultados 
            Me.LblResultado.Text = listResp(0)
            For iFila As Integer = 1 To listResp.Count - 1
                LblResultado.Text += vbCrLf & listResp(iFila)
            Next
            '>>>═══════════════════════════════════════════════════════════════════════════════════════
            ProgressBar.Value = 100
            bProcesado = True
        Catch ex As Exception
            bProcesado = False
        End Try

        Return bProcesado
    End Function

    Private Sub Limpiar()

        '>>>════════════════════════════════════════════════════════════════════════════════════════════════════════
        '<<< (CesarAlejandroCabezas) 06/abril/2022//  Fase 3 - Diseño de Modulos  
        '>>> Se realiza un cursor por cada uno de los grupos donde se encuentran las matrices
        '>>> y se limpian cada una de las cajas de texto en las matrices 
        '>>>════════════════════════════════════════════════════════════════════════════════════════════════════════
        For Each ctrl As Control In GBoxMatriz22.Controls
            If ((TypeOf ctrl Is TextBox)) Then
                ctrl.Text = ""
            End If
        Next
        '>>>════════════════════════════════════════════════════════════════════════════════════════════════════════
        For Each ctrl As Control In GBoxMatriz33.Controls
            If ((TypeOf ctrl Is TextBox)) Then
                ctrl.Text = ""
            End If
        Next
        '>>>════════════════════════════════════════════════════════════════════════════════════════════════════════
        For Each ctrl As Control In GBoxResp_M22.Controls
            If ((TypeOf ctrl Is TextBox)) Then
                ctrl.Text = ""
            End If
        Next
        '>>>════════════════════════════════════════════════════════════════════════════════════════════════════════
        For Each ctrl As Control In GBoxResp_M33.Controls
            If ((TypeOf ctrl Is TextBox)) Then
                ctrl.Text = ""
            End If
        Next
        '>>>════════════════════════════════════════════════════════════════════════════════════════════════════════
        GBoxResp_M33.Visible = False
        GBoxResp_M22.Visible = False
        LblResultado.Text = ""
        ProgressBar.Value = 0
        ProgressBar.Visible = False
    End Sub

    '>>>═══════════════════════════════════════════════════════════════════════════════════════
    '>>> Se verifican cada uno de los campos de las matrices, con el proposito que se posea
    '>>> Información en cada una de las cajas de texto 
    Private Function ValidadorMatrices() As Boolean
        Dim bResp As Boolean = False
        Try

            bResp = True
        Catch ex As Exception
            bResp = False
        End Try

        Return bResp
    End Function

    '>>>═══════════════════════════════════════════════════════════════════════════════════════
    '>>> Se valida los cambios o las Selecciones de la Matriz que se desea procesar
    Private Sub RBtn3x3_CheckedChanged(sender As Object, e As EventArgs) Handles RBtn3x3.CheckedChanged
        Dim RB As RadioButton = CType(sender, RadioButton)

        If (RB.Checked) Then
            Matriz22 = False
            GBoxMatriz33.Visible = True
            GBoxMatriz22.Visible = False
        End If
    End Sub
    Private Sub RBtn2x2_CheckedChanged(sender As Object, e As EventArgs) Handles RBtn2x2.CheckedChanged
        Dim RB As RadioButton = CType(sender, RadioButton)
        If (RB.Checked) Then
            Matriz22 = True
            GBoxMatriz22.Visible = True
            GBoxMatriz33.Visible = False
        End If
    End Sub

    '>>>═══════════════════════════════════════════════════════════════════════════════════════

End Class