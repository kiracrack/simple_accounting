Imports MySql.Data.MySqlClient ' this is to import MySQL.NET
Imports System
Imports System.IO
Imports System.Net

Public Class frmConnectionSetup

    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, ByVal keyData As Keys) As Boolean
        If keyData = (Keys.Escape) Then
            Me.Close()
        End If
        Return ProcessCmdKey
    End Function
    Private Sub frmRequestType_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If System.IO.File.Exists(file_conn) = True Then
            Dim sr As StreamReader = File.OpenText(file_conn)
            Dim br As String = sr.ReadLine() : sr.Close()
            Dim strSetup = DecryptTripleDES(br) : Dim cnt As Integer = 0
            For Each word In strSetup.Split(New Char() {","c})
                If cnt = 0 Then
                    txtServerhost.Text = word
                ElseIf cnt = 1 Then
                    txtPort.Text = word
                ElseIf cnt = 2 Then
                    txtusername.Text = word
                ElseIf cnt = 3 Then
                    txtpassword.Text = word
                ElseIf cnt = 4 Then
                    txtDatabase.Text = word
                End If
                cnt = cnt + 1
            Next
        End If
        clientserver = ""
        clientuser = ""
        clientpass = ""
        clientdatabase = ""
    End Sub
  
    Private Sub cmdset_Click(sender As Object, e As EventArgs) Handles cmdset.Click
        If txtServerhost.Text = "" Then
            MessageBox.Show("Please enter Server host!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            txtServerhost.Focus()
            Exit Sub
        ElseIf txtusername.Text = "" Then
            MessageBox.Show("Please enter username!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            txtusername.Focus()
            Exit Sub
        ElseIf txtpassword.Text = "" Then
            MessageBox.Show("Please enter password!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            txtpassword.Focus()
            Exit Sub
        End If
        clientserver = LCase(txtServerhost.Text)
        clientport = txtPort.Text
        clientuser = txtusername.Text
        clientpass = txtpassword.Text
        clientdatabase = txtDatabase.Text
        OpenClientServer()
        If connclient.State <> 0 Then
            Try
                If System.IO.File.Exists(file_conn) = True Then
                    System.IO.File.Delete(file_conn)
                End If
                Dim detailsFile As StreamWriter = Nothing
                detailsFile = New StreamWriter(file_conn, True)
                detailsFile.WriteLine(EncryptTripleDES(txtServerhost.Text & "," & txtPort.Text & "," & txtusername.Text & "," & txtpassword.Text & "," & txtDatabase.Text))
                detailsFile.Close()

                MessageBox.Show("System successfully activated", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.Close()
            Catch errMYSQL As MySqlException
                MessageBox.Show("Message:" & errMYSQL.Message, vbCrLf _
                                & "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            Catch errMS As Exception
                MessageBox.Show("Message:" & errMS.Message & vbCrLf, _
                                  "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End Try
        End If
    End Sub
End Class