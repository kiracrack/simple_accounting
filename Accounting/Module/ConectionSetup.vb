Imports MySql.Data.MySqlClient ' this is to import MySQL.NET
Imports System.Data
Imports System.Management
Imports Microsoft.VisualBasic
Imports System.Net.Mail
Imports System.Text
Imports System.IO

Module ConectionSetup
    Public li As String = Environment.NewLine
    Public Em As DataGridView
    Public fversion As Date = System.IO.File.GetLastWriteTime(My.Application.Info.DirectoryPath).ToShortDateString()
    Public file_conn As String = Application.StartupPath.ToString & "\" & My.Application.Info.AssemblyName & ".conn"
    Public formclose As Boolean
    Public conn As New MySqlConnection
    Public msda As MySqlDataAdapter
    Public dst As New DataSet
    Public com As New MySqlCommand
    Public rst As MySqlDataReader
    Public branchcode As String
   
    Public connclient As New MySqlConnection 'for MySQLDatabase Connection
    Public msdaclient As MySqlDataAdapter 'is use to update the dataset and datasource
    Public dstclient As New DataSet 'miniature of your table - cache table to client
    Public comclient As New MySqlCommand
    Public rstclient As MySqlDataReader

    Public sqlserver As String
    Public sqlport As String
    Public sqluser As String
    Public sqlpass As String
    Public sqldatabase As String

    Public clientserver As String
    Public clientport As String
    Public clientuser As String
    Public clientpass As String
    Public clientdatabase As String

    Public GLobalStoreID As String
    Public GLobalStoreName As String
    Public GLobalStorebegAsset As Double
    Public GLobalGLfinancial As String
    Public GlobalEnableSavings As Boolean
    Public GlobalSavings As Double

    Public Function ConnectVerify()
        Try
            Dim strSetup As String = ""
            Dim sr As StreamReader = File.OpenText(file_conn)
            Dim br As String = sr.ReadLine() : sr.Close()
            strSetup = DecryptTripleDES(br) : Dim cnt As Integer = 0
            For Each word In strSetup.Split(New Char() {","c})
                If cnt = 0 Then
                    sqlserver = word
                ElseIf cnt = 1 Then
                    sqlport = word
                ElseIf cnt = 2 Then
                    sqluser = word
                ElseIf cnt = 3 Then
                    sqlpass = word
                ElseIf cnt = 4 Then
                    sqldatabase = word
                End If
                cnt = cnt + 1
            Next

            If sqlserver <> "localhost" Then
                Dim ipAddress As System.Net.IPAddress = Nothing
                If System.Net.IPAddress.TryParse(sqlserver, ipAddress) = False Then
                    sqlserver = GetIPAddress(sqlserver)
                End If
            End If

            conn = New MySql.Data.MySqlClient.MySqlConnection
            conn.ConnectionString = "server=" & sqlserver & "; Port=" & sqlport & "; user id=" & sqluser & "; password=" & sqlpass & "; database=" & sqldatabase & "; Connection Timeout=6000000 ; Allow Zero Datetime=True"
            conn.Open()
            com.Connection = conn
            com.CommandTimeout = 6000000
        Catch ex As Exception
            MessageBox.Show("Can't connect database server on '" & sqlserver & "'", _
                          "Connection Server Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
        Return True
    End Function
    Public Function OpenClientServer() As Boolean
        Try
            If sqlserver <> "localhost" Then
                Dim ipAddress As System.Net.IPAddress = Nothing
                If System.Net.IPAddress.TryParse(clientserver, ipAddress) = False Then
                    clientserver = GetIPAddress(clientserver)
                End If
            End If
            connclient = New MySql.Data.MySqlClient.MySqlConnection
            connclient.ConnectionString = "server=" & clientserver & "; Port=" & clientport & "; user id=" & clientuser & "; password=" & clientpass & "; database=" & clientdatabase & "; Connection Timeout=10; allow user variables=true"
            connclient.Open()
            comclient.Connection = connclient
            comclient.CommandTimeout = 0
            OpenClientServer = True

        Catch errMYSQL As MySqlException
            MessageBox.Show("Message:" & errMYSQL.Message & vbCrLf, _
                             "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            OpenClientServer = False
            Return False
        Catch errMS As Exception
            MessageBox.Show("Message:" & errMS.Message & vbCrLf, _
                              "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            OpenClientServer = False
            Return False
        End Try
    End Function
    Public Function GetIPAddress(ByVal computername) As String
        Dim strIPAddress As String
        strIPAddress = System.Net.Dns.GetHostByName(computername).AddressList(0).ToString()
        computername = strIPAddress
        Return computername
    End Function
    Public Sub LoadGeneralSetting()
        Dim cnt As Integer = 0
        com.CommandText = "SELECT * from tblglsettings" : rst = com.ExecuteReader()
        While rst.Read
            GLobalGLfinancial = rst("glfinancial").ToString
        End While
        rst.Close()
    End Sub
    Public Sub LoadStoreProfile()
        com.CommandText = "select * from tblstore" : rst = com.ExecuteReader()
        While rst.Read()
            GLobalStoreName = rst("accountname").ToString
            GLobalStoreID = rst("accountno").ToString
            GLobalStorebegAsset = Val(rst("beginningasset").ToString)
            If rst("savings") = 0 Then
                GlobalEnableSavings = False
            Else
                GlobalEnableSavings = True
                GlobalSavings = rst("savings")
            End If
        End While
        rst.Close()
    End Sub
    Public Class ComboBoxItem
        Private displayValue As String
        Private m_hiddenValue As String

        Public Sub New(ByVal d As String, ByVal h As String)
            displayValue = d
            m_hiddenValue = h
        End Sub

        Public ReadOnly Property HiddenValue() As String
            Get
                Return m_hiddenValue
            End Get
        End Property

        Public Overrides Function ToString() As String
            Return displayValue
        End Function
    End Class
    Public Function codeGenerator(ByVal table As String, ByVal field As String, ByVal length As Integer, ByVal initialcode As String, ByVal startfrom As String)
        Dim strng = ""

        If CInt(countrecord(table)) = 0 Then
            strng = initialcode & startfrom
        Else
            com.CommandText = "select " & field & " from " & table & " order by right(" & field & "," & length & ") desc limit 1" : rst = com.ExecuteReader()
            Dim removechar As Char() = initialcode.ToCharArray()
            Dim sb As New System.Text.StringBuilder
            While rst.Read
                Dim str As String = rst(field).ToString
                For Each ch As Char In str
                    If Array.IndexOf(removechar, ch) = -1 Then
                        sb.Append(ch)
                    End If
                Next
            End While
            rst.Close()
            strng = initialcode & Val(sb.ToString) + 1
        End If
        Return strng.ToString
    End Function

    Public Sub UpdateFinancialSummary(ByVal datetrn As Date)
        Dim query As String = "" : Dim sortorder As Integer = 0 : Dim asset As Double = 0 : Dim receivable As Double = 0 : Dim sales As Double = 0 : Dim expenses As Double = 0 : Dim netincome As Double = 0 : Dim cash As Double = 0 : Dim capital As Double = 0 : Dim earn As Double = 0

        Try
            Dim dateSummary As String = Format(datetrn.ToString("yyyy-MM-dd"))
            Dim ttlreceivable As Double = Val(qrysingledata("ttl", "sum(credit)- sum(debit)  as 'ttl'", "tbljournal where journalmode=1 and date_format(datetrn,'%Y-%m-%d') <= '" & dateSummary & "' "))
            com.CommandText = "delete from tblfinancial where financialdate ='" & dateSummary & "'" : com.ExecuteNonQuery()
            com.CommandText = "select 'BALANCE CONSUMABLE ASSET'  as description , (ifnull((select sum(beginningasset) from tblstore),0)+ifnull(sum(credit),0))- ifnull((select sum(debit) from tbljournal where journalmode=0 and date_format(datetrn,'%Y-%m-%d') <= '" & dateSummary & "' and referenceno is null),0) as 'asset'  from tbljournal inner join tblitem on tbljournal.itemcode = tblitem.itemcode where journalmode=0 and (category=0 or category=2) and date_format(datetrn,'%Y-%m-%d') <= '" & dateSummary & "'" : rst = com.ExecuteReader
            While rst.Read
                sortorder = sortorder + 1
                asset = Val(rst("asset").ToString) - Val(ttlreceivable)
                query = "insert into tblfinancial set financialdate='" & dateSummary & "',description='" & rchar(rst("description").ToString) & "', amount=" & asset & ",sortorder=" & sortorder & ";" & Environment.NewLine
            End While
            rst.Close()

            com.CommandText = "select 'EXPENSES' as description , sum(credit) as ttlcredit  from tbljournal where journalmode=0 and date_format(datetrn,'%Y-%m-%d') <= '" & dateSummary & "' " : rst = com.ExecuteReader
            While rst.Read
                sortorder = sortorder + 1
                query = query + "insert into tblfinancial set financialdate='" & dateSummary & "',description='" & rchar(rst("description").ToString) & "', amount='" & Val(rst("ttlcredit")) & "',sortorder=" & sortorder & ";" & Environment.NewLine
            End While
            rst.Close()

            com.CommandText = "select concat('  - ','Consumable Expenses')  as description ,  sum(credit) as 'val'  from tbljournal inner join tblitem on tbljournal.itemcode = tblitem.itemcode where journalmode=0 and category = 0 and date_format(datetrn,'%Y-%m-%d') <= '" & dateSummary & "' " : rst = com.ExecuteReader
            While rst.Read
                If Val(rst("val").ToString) > 0 Then
                    sortorder = sortorder + 1
                    query = query + "insert into tblfinancial set financialdate='" & dateSummary & "',description='" & rchar(rst("description").ToString) & "', amount=" & Val(rst("val").ToString) & ",sortorder=" & sortorder & ";" & Environment.NewLine
                End If
            End While
            rst.Close()

            com.CommandText = "select concat('  - ','Liabilities')  as description ,  sum(credit) as 'val'  from tbljournal inner join tblitem on tbljournal.itemcode = tblitem.itemcode where journalmode=0 and category = 3 and date_format(datetrn,'%Y-%m-%d') <= '" & dateSummary & "' " : rst = com.ExecuteReader
            While rst.Read
                If Val(rst("val").ToString) > 0 Then
                    sortorder = sortorder + 1
                    query = query + "insert into tblfinancial set financialdate='" & dateSummary & "',description='" & rchar(rst("description").ToString) & "', amount=" & Val(rst("val").ToString) & ",sortorder=" & sortorder & ";" & Environment.NewLine
                End If
            End While
            rst.Close()

            com.CommandText = "select concat('  - ','Billing''s and Payments')  as description ,  sum(credit) as 'val'  from tbljournal inner join tblitem on tbljournal.itemcode = tblitem.itemcode where journalmode=0 and category = 1 and date_format(datetrn,'%Y-%m-%d') <= '" & dateSummary & "' " : rst = com.ExecuteReader
            While rst.Read
                If Val(rst("val").ToString) > 0 Then
                    sortorder = sortorder + 1
                    query = query + "insert into tblfinancial set financialdate='" & dateSummary & "',description='" & rchar(rst("description").ToString) & "', amount=" & Val(rst("val").ToString) & ",sortorder=" & sortorder & ";" & Environment.NewLine
                End If
            End While
            rst.Close()

            com.CommandText = "select concat('  - ','Other Assets')  as description , sum(credit) as 'val'  from tbljournal inner join tblitem on tbljournal.itemcode = tblitem.itemcode where journalmode=0 and category = 2 and date_format(datetrn,'%Y-%m-%d') <= '" & dateSummary & "' " : rst = com.ExecuteReader
            While rst.Read
                If Val(rst("val").ToString) > 0 Then
                    sortorder = sortorder + 1
                    query = query + "insert into tblfinancial set financialdate='" & dateSummary & "',description='" & rchar(rst("description").ToString) & "', amount='" & Val(rst("val").ToString) & "',sortorder=" & sortorder & ";" & Environment.NewLine
                End If
            End While
            rst.Close()

            com.CommandText = "select 'ACCOUNT RECEIVABLE' as description, sum(credit)- sum(debit)  as 'balance'   from tbljournal where journalmode=1 and date_format(datetrn,'%Y-%m-%d') <= '" & dateSummary & "' " : rst = com.ExecuteReader
            While rst.Read
                sortorder = sortorder + 1
                receivable = Val(rst("balance").ToString)
                query = query + "insert into tblfinancial set financialdate='" & dateSummary & "',description='" & rchar(rst("description").ToString) & "', amount='" & receivable & "',sortorder=" & sortorder & ";" & Environment.NewLine
            End While
            rst.Close()

            com.CommandText = "select concat('  - ',(select accountname from tblcreditaccount where acctnumber=tbljournal.accountno)) as description, sum(credit)- sum(debit)  as 'balance'   from tbljournal where journalmode=1 and date_format(datetrn,'%Y-%m-%d') <= '" & dateSummary & "'  group by accountno" : rst = com.ExecuteReader
            While rst.Read
                If Val(rst("balance").ToString) > 0 Then
                    sortorder = sortorder + 1
                    query = query + "insert into tblfinancial set financialdate='" & dateSummary & "',description='" & StrConv(rchar(rst("description").ToString), vbProperCase) & "', amount='" & Val(rst("balance").ToString) & "',sortorder=" & sortorder & ";" & Environment.NewLine
                End If
            End While
            rst.Close()

            com.CommandText = "select  'TOTAL SALES'  as description , sum(debit) as ttldebit  from tbljournal where journalmode=0 and date_format(datetrn,'%Y-%m-%d') <= '" & dateSummary & "' " : rst = com.ExecuteReader
            While rst.Read
                sortorder = sortorder + 1
                sales = sales + Val(rst("ttldebit").ToString)
                query = query + "insert into tblfinancial set financialdate='" & dateSummary & "',description='" & rchar(rst("description").ToString) & "', amount='" & Val(rst("ttldebit").ToString) & "',sortorder=" & sortorder & ";" & Environment.NewLine
            End While
            rst.Close()

            com.CommandText = "select 'TOTAL EXPENSES' as description , sum(credit) as ttlcredit  from tbljournal where journalmode=0 and date_format(datetrn,'%Y-%m-%d') <= '" & dateSummary & "' " : rst = com.ExecuteReader
            While rst.Read
                sortorder = sortorder + 1
                expenses = expenses + Val(rst("ttlcredit").ToString)
                query = query + "insert into tblfinancial set financialdate='" & dateSummary & "',description='" & rchar(rst("description").ToString) & "', amount='" & Val(rst("ttlcredit")) & "',sortorder=" & sortorder & ";" & Environment.NewLine
            End While
            rst.Close()

            sortorder = sortorder + 1
            netincome = sales - expenses
            query = query + "insert into tblfinancial set financialdate='" & dateSummary & "',description='NET INCOME', amount=(" & netincome & "),sortorder=" & sortorder & ";" & Environment.NewLine

            sortorder = sortorder + 1
            query = query + "insert into tblfinancial set financialdate='" & dateSummary & "',description='TOTAL ASSETS (Consumables+Receivable+Cash)', amount=(" & Val(netincome) + Val(asset) + Val(receivable) & "),sortorder=" & sortorder & ";" & Environment.NewLine

            sortorder = sortorder + 1
            query = query + "insert into tblfinancial set financialdate='" & dateSummary & "',description='TOTAL SAVINGS', amount=ifnull((select sum(totalsavings) from tblsavings),0),sortorder=" & sortorder & ";" & Environment.NewLine

            com.CommandText = query : com.ExecuteNonQuery()

        Catch errMYSQL As MySqlException
            MessageBox.Show("Message:" & errMYSQL.Message, _
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Catch errMS As Exception
            MessageBox.Show("Message:" & errMS.Message, _
                             "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Module
