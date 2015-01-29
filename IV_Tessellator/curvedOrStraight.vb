Imports System.Windows.Forms
Imports System.Diagnostics.CodeAnalysis
Public Class curvedOrStraight

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles cbCurved.Click
        MessageBox.Show("Curved!")
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles cbStraight.Click
        MessageBox.Show("Straight!")
    End Sub
End Class