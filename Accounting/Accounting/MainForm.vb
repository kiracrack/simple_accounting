Public Class MainForm
    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, ByVal keyData As Keys) As Boolean
        If keyData = (Keys.Shift) + Keys.F12 Then
            frmConnectionSetup.Show()
        
        End If
        Return ProcessCmdKey
    End Function
    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If System.IO.File.Exists(file_conn) = False Then
            frmConnectionSetup.ShowDialog()
            Me.Close()
        End If
        If ConnectVerify() Then
            'LoadStoreProfile()
            'LoadGeneralSetting()
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        frmGLItem.Show()
    End Sub
 
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        frmPostTransaction.Show()
    End Sub
 
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        frmTrnSettingsItems.Show()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs)
        frmAccountTitleLedger.Show()
    End Sub

   
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        frmGLGroup.Show()
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        frmPostingPredefineGL.Show()
    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        frmTrialBalance.Show()
    End Sub
 
End Class
