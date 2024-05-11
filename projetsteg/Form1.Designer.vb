<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Button1 = New Button()
        Button2 = New Button()
        Button3 = New Button()
        Button4 = New Button()
        Button5 = New Button()
        Button6 = New Button()
        Button7 = New Button()
        Button8 = New Button()
        data = New DataGridView()
        CType(data, ComponentModel.ISupportInitialize).BeginInit()
        SuspendLayout()
        ' 
        ' Button1
        ' 
        Button1.BackColor = SystemColors.GradientActiveCaption
        Button1.Location = New Point(327, 214)
        Button1.Name = "Button1"
        Button1.Size = New Size(227, 23)
        Button1.TabIndex = 0
        Button1.Text = "Connexions AutoCAD"
        Button1.UseVisualStyleBackColor = False
        ' 
        ' Button2
        ' 
        Button2.BackColor = SystemColors.GradientActiveCaption
        Button2.Location = New Point(327, 301)
        Button2.Name = "Button2"
        Button2.Size = New Size(227, 23)
        Button2.TabIndex = 1
        Button2.Text = "PDF Schedules"
        Button2.UseVisualStyleBackColor = False
        ' 
        ' Button3
        ' 
        Button3.BackColor = SystemColors.GradientActiveCaption
        Button3.Location = New Point(327, 272)
        Button3.Name = "Button3"
        Button3.Size = New Size(227, 23)
        Button3.TabIndex = 2
        Button3.Text = "Google Earth"
        Button3.UseVisualStyleBackColor = False
        ' 
        ' Button4
        ' 
        Button4.BackColor = SystemColors.GradientActiveCaption
        Button4.Location = New Point(327, 243)
        Button4.Name = "Button4"
        Button4.Size = New Size(227, 23)
        Button4.TabIndex = 3
        Button4.Text = "TrackMaker"
        Button4.UseVisualStyleBackColor = False
        ' 
        ' Button5
        ' 
        Button5.BackColor = SystemColors.GradientActiveCaption
        Button5.Location = New Point(327, 185)
        Button5.Name = "Button5"
        Button5.Size = New Size(227, 23)
        Button5.TabIndex = 4
        Button5.Text = "Traitement Connexions"
        Button5.UseVisualStyleBackColor = False
        ' 
        ' Button6
        ' 
        Button6.BackColor = SystemColors.GradientActiveCaption
        Button6.Location = New Point(327, 153)
        Button6.Name = "Button6"
        Button6.Size = New Size(227, 23)
        Button6.TabIndex = 5
        Button6.Text = "Planimétrie BT sur AutoCAD"
        Button6.UseVisualStyleBackColor = False
        ' 
        ' Button7
        ' 
        Button7.BackColor = SystemColors.GradientActiveCaption
        Button7.Location = New Point(327, 124)
        Button7.Name = "Button7"
        Button7.Size = New Size(227, 23)
        Button7.TabIndex = 6
        Button7.Text = "Feeders"
        Button7.UseVisualStyleBackColor = False
        ' 
        ' Button8
        ' 
        Button8.BackColor = SystemColors.GradientActiveCaption
        Button8.Location = New Point(327, 95)
        Button8.Name = "Button8"
        Button8.Size = New Size(227, 23)
        Button8.TabIndex = 7
        Button8.Text = "Traitement Réseau  Base Tension "
        Button8.UseVisualStyleBackColor = False
        ' 
        ' data
        ' 
        data.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        data.Location = New Point(1, 330)
        data.Name = "data"
        data.Size = New Size(907, 213)
        data.TabIndex = 8
        ' 
        ' Form1
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(920, 538)
        Controls.Add(data)
        Controls.Add(Button8)
        Controls.Add(Button7)
        Controls.Add(Button6)
        Controls.Add(Button5)
        Controls.Add(Button4)
        Controls.Add(Button3)
        Controls.Add(Button2)
        Controls.Add(Button1)
        Name = "Form1"
        Text = "Form1"
        CType(data, ComponentModel.ISupportInitialize).EndInit()
        ResumeLayout(False)
    End Sub

    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents Button5 As Button
    Friend WithEvents Button6 As Button
    Friend WithEvents Button7 As Button
    Friend WithEvents Button8 As Button
    Friend WithEvents data As DataGridView

End Class
