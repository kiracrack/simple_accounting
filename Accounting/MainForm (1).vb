Public Class MainForm
    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, ByVal keyData As Keys) As Boolean
        If keyData = (Keys.Shift) + Keys.F12 Then
            frmConnectionSetup.Show()
        ElseIf keyData = (Keys.Shift) + Keys.F11 Then
            frmGLSettings.Show()
        ElseIf keyData = (Keys.Shift) + Keys.F10 Then
            frmStakeholder.Show()
        End If
        Return ProcessCmdKey
    End Function
    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If System.IO.File.Exists(file_conn) = False Then
            frmConnectionSetup.ShowDialog()
            Me.Close()
        End If
        If ConnectVerify() Then
            LoadStoreProfile()
            LoadGeneralSetting()
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        frmItems.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        frmStoreAccount.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        frmJournalMode.Show()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        frmPassbookLedger.Show()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        frmCreditAccount.Show()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        frmCreditLedger.Show()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles cmdFinancial.Click
        frmGeneralSummary.Show()
    End Sub

    Private Sub Button7_Click_1(sender As Object, e As EventArgs) Handles Button7.Click
        frmSavingsLedger.Show()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If countrecord("tblglsettings") = 0 Then
            MessageBox.Show("GL Financial GL Setting not available!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        frmFinancialCapital.Show()
    End Sub
End Class
