<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.TestButton = New System.Windows.Forms.Button()
        Me.ArchiveButton = New System.Windows.Forms.Button()
        Me.LiveButton = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'TestButton
        '
        Me.TestButton.BackColor = System.Drawing.Color.DimGray
        Me.TestButton.FlatAppearance.BorderColor = System.Drawing.Color.White
        Me.TestButton.FlatAppearance.BorderSize = 0
        Me.TestButton.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.TestButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TestButton.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.TestButton.Location = New System.Drawing.Point(12, 81)
        Me.TestButton.Name = "TestButton"
        Me.TestButton.Size = New System.Drawing.Size(145, 63)
        Me.TestButton.TabIndex = 5
        Me.TestButton.Text = "TEST File Monitor"
        Me.TestButton.UseVisualStyleBackColor = False
        '
        'ArchiveButton
        '
        Me.ArchiveButton.BackColor = System.Drawing.Color.ForestGreen
        Me.ArchiveButton.FlatAppearance.BorderColor = System.Drawing.Color.White
        Me.ArchiveButton.FlatAppearance.BorderSize = 0
        Me.ArchiveButton.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.ArchiveButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ArchiveButton.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.ArchiveButton.Location = New System.Drawing.Point(169, 12)
        Me.ArchiveButton.Name = "ArchiveButton"
        Me.ArchiveButton.Size = New System.Drawing.Size(145, 63)
        Me.ArchiveButton.TabIndex = 4
        Me.ArchiveButton.Text = "Archive Files"
        Me.ArchiveButton.UseVisualStyleBackColor = False
        '
        'LiveButton
        '
        Me.LiveButton.BackColor = System.Drawing.Color.MediumBlue
        Me.LiveButton.FlatAppearance.BorderColor = System.Drawing.Color.White
        Me.LiveButton.FlatAppearance.BorderSize = 0
        Me.LiveButton.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.LiveButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LiveButton.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.LiveButton.Location = New System.Drawing.Point(12, 12)
        Me.LiveButton.Name = "LiveButton"
        Me.LiveButton.Size = New System.Drawing.Size(145, 63)
        Me.LiveButton.TabIndex = 3
        Me.LiveButton.Text = "XML File Monitor"
        Me.LiveButton.UseVisualStyleBackColor = False
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Indigo
        Me.ClientSize = New System.Drawing.Size(329, 158)
        Me.Controls.Add(Me.TestButton)
        Me.Controls.Add(Me.ArchiveButton)
        Me.Controls.Add(Me.LiveButton)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "Form1"
        Me.Text = "XML File Monitor"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TestButton As Button
    Friend WithEvents ArchiveButton As Button
    Friend WithEvents LiveButton As Button
End Class
