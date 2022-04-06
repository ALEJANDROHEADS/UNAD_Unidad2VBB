<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmEjercicio7
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.TxtResultado = New System.Windows.Forms.Label()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.BtnLimpiar = New System.Windows.Forms.Button()
        Me.TxtUnidades = New System.Windows.Forms.TextBox()
        Me.CBoxA = New System.Windows.Forms.ComboBox()
        Me.CBoxDe = New System.Windows.Forms.ComboBox()
        Me.BtnConvertir = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox3.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(12, 48)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(634, 462)
        Me.TabControl1.TabIndex = 8
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.GroupBox2)
        Me.TabPage1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.TabPage1.Location = New System.Drawing.Point(4, 24)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(626, 434)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Convertir"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.GroupBox1)
        Me.GroupBox2.Controls.Add(Me.PictureBox2)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.BtnLimpiar)
        Me.GroupBox2.Controls.Add(Me.TxtUnidades)
        Me.GroupBox2.Controls.Add(Me.CBoxA)
        Me.GroupBox2.Controls.Add(Me.CBoxDe)
        Me.GroupBox2.Controls.Add(Me.BtnConvertir)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.GroupBox3)
        Me.GroupBox2.Font = New System.Drawing.Font("Segoe UI", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.GroupBox2.Location = New System.Drawing.Point(20, 18)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(585, 398)
        Me.GroupBox2.TabIndex = 2
        Me.GroupBox2.TabStop = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.TxtResultado)
        Me.GroupBox1.Location = New System.Drawing.Point(32, 231)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(519, 97)
        Me.GroupBox1.TabIndex = 22
        Me.GroupBox1.TabStop = False
        '
        'TxtResultado
        '
        Me.TxtResultado.AutoSize = True
        Me.TxtResultado.ForeColor = System.Drawing.Color.ForestGreen
        Me.TxtResultado.Location = New System.Drawing.Point(6, 23)
        Me.TxtResultado.Name = "TxtResultado"
        Me.TxtResultado.Size = New System.Drawing.Size(12, 20)
        Me.TxtResultado.TabIndex = 0
        Me.TxtResultado.Text = ":"
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = Global.HeadsAppUnidad2.My.Resources.Resources.Download_256x256
        Me.PictureBox2.Location = New System.Drawing.Point(483, 14)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(96, 94)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 21
        Me.PictureBox2.TabStop = False
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Segoe UI", 11.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point)
        Me.Label6.Location = New System.Drawing.Point(216, 110)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(36, 20)
        Me.Label6.TabIndex = 20
        Me.Label6.Text = "De :"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Segoe UI", 11.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point)
        Me.Label5.Location = New System.Drawing.Point(404, 110)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(28, 20)
        Me.Label5.TabIndex = 19
        Me.Label5.Text = "A :"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Segoe UI", 11.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point)
        Me.Label4.Location = New System.Drawing.Point(36, 110)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(53, 20)
        Me.Label4.TabIndex = 18
        Me.Label4.Text = "Valor :"
        '
        'BtnLimpiar
        '
        Me.BtnLimpiar.Location = New System.Drawing.Point(32, 193)
        Me.BtnLimpiar.Name = "BtnLimpiar"
        Me.BtnLimpiar.Size = New System.Drawing.Size(226, 32)
        Me.BtnLimpiar.TabIndex = 17
        Me.BtnLimpiar.Text = "Limpiar "
        Me.BtnLimpiar.UseVisualStyleBackColor = True
        '
        'TxtUnidades
        '
        Me.TxtUnidades.Location = New System.Drawing.Point(36, 134)
        Me.TxtUnidades.Name = "TxtUnidades"
        Me.TxtUnidades.Size = New System.Drawing.Size(147, 27)
        Me.TxtUnidades.TabIndex = 15
        '
        'CBoxA
        '
        Me.CBoxA.FormattingEnabled = True
        Me.CBoxA.Items.AddRange(New Object() {"BYTES", "KILOBYTES - KB", "MEGABYTES - MG", "GIGABYTES - GB", "Bits"})
        Me.CBoxA.Location = New System.Drawing.Point(404, 133)
        Me.CBoxA.Name = "CBoxA"
        Me.CBoxA.Size = New System.Drawing.Size(147, 28)
        Me.CBoxA.TabIndex = 14
        '
        'CBoxDe
        '
        Me.CBoxDe.FormattingEnabled = True
        Me.CBoxDe.Items.AddRange(New Object() {"BYTES", "KILOBYTES - KB", "MEGABYTES - MG", "GIGABYTES - GB", "Bits"})
        Me.CBoxDe.Location = New System.Drawing.Point(216, 133)
        Me.CBoxDe.Name = "CBoxDe"
        Me.CBoxDe.Size = New System.Drawing.Size(148, 28)
        Me.CBoxDe.TabIndex = 13
        '
        'BtnConvertir
        '
        Me.BtnConvertir.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.BtnConvertir.Location = New System.Drawing.Point(340, 193)
        Me.BtnConvertir.Name = "BtnConvertir"
        Me.BtnConvertir.Size = New System.Drawing.Size(211, 32)
        Me.BtnConvertir.TabIndex = 12
        Me.BtnConvertir.Text = "Convertir"
        Me.BtnConvertir.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Segoe UI", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point)
        Me.Label3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Label3.Location = New System.Drawing.Point(6, 23)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(261, 25)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "CONVERSOR DE UNIDADES "
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Label2)
        Me.GroupBox3.Font = New System.Drawing.Font("Segoe UI", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.GroupBox3.Location = New System.Drawing.Point(32, 324)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(519, 58)
        Me.GroupBox3.TabIndex = 4
        Me.GroupBox3.TabStop = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.ForeColor = System.Drawing.Color.CornflowerBlue
        Me.Label2.Location = New System.Drawing.Point(19, 19)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(449, 26)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Herramienta para realizar e identificar Conversión entre unidades de Almacenamien" &
    "to." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Puedes seleccionar la unidad a convertir y la unidad a la cual la deseas co" &
    "nvertir. " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.PictureBox1)
        Me.TabPage2.Location = New System.Drawing.Point(4, 24)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(626, 434)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Unidades "
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.HeadsAppUnidad2.My.Resources.Resources.Almacenamiento
        Me.PictureBox1.Location = New System.Drawing.Point(15, 16)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(594, 402)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 0
        Me.PictureBox1.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point)
        Me.Label1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(155, 25)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "EJERCICIO No. 7"
        '
        'FrmEjercicio7
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(658, 522)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TabControl1)
        Me.Name = "FrmEjercicio7"
        Me.Text = "Conversión Tamaño Almacenamiento "
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents TabPage1 As TabPage
    Friend WithEvents TabPage2 As TabPage
    Friend WithEvents Label1 As Label
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents Label3 As Label
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents BtnLimpiar As Button
    Friend WithEvents TxtUnidades As TextBox
    Friend WithEvents CBoxDe As ComboBox
    Friend WithEvents BtnConvertir As Button
    Friend WithEvents Label6 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents PictureBox2 As PictureBox
    Protected WithEvents CBoxA As ComboBox
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents TxtResultado As Label
End Class
