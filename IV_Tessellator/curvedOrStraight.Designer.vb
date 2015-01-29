<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class curvedOrStraight
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(curvedOrStraight))
        Me.cbCurved = New System.Windows.Forms.Button()
        Me.cbStraight = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'cbCurved
        '
        Me.cbCurved.Location = New System.Drawing.Point(12, 12)
        Me.cbCurved.Name = "cbCurved"
        Me.cbCurved.Size = New System.Drawing.Size(75, 23)
        Me.cbCurved.TabIndex = 0
        Me.cbCurved.Text = "Curved!"
        Me.cbCurved.UseVisualStyleBackColor = True
        '
        'cbStraight
        '
        Me.cbStraight.Location = New System.Drawing.Point(173, 12)
        Me.cbStraight.Name = "cbStraight"
        Me.cbStraight.Size = New System.Drawing.Size(75, 23)
        Me.cbStraight.TabIndex = 1
        Me.cbStraight.Text = "Straight!"
        Me.cbStraight.UseVisualStyleBackColor = True
        '
        'curvedOrStraight
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 54)
        Me.ControlBox = False
        Me.Controls.Add(Me.cbStraight)
        Me.Controls.Add(Me.cbCurved)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "curvedOrStraight"
        Me.Text = "Curved Or Straight Segments?"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents cbCurved As System.Windows.Forms.Button
    Friend WithEvents cbStraight As System.Windows.Forms.Button
End Class
