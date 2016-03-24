Imports System.Windows.Forms

Public Class PinyinDialog

    Public pytype As Integer

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If RadioButton1.Checked Then
            pytype = 1
        End If
        If RadioButton2.Checked Then
            pytype = 2
        End If
        If RadioButton3.Checked Then
            pytype = 3
        End If
        If RadioButton4.Checked Then
            pytype = 4
        End If
        If RadioButton5.Checked Then
            pytype = 5
        End If
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

End Class
