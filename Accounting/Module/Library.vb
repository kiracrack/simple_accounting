Imports System.IO
Imports MySql.Data.MySqlClient
Imports System.Net.Mail
Imports System.Text
Imports System.Net
Imports System.Collections.Generic

Module library
    Public removechar As Char() = "\".ToCharArray()
    Public sb As New System.Text.StringBuilder
    Public imgBytes As Byte() = Nothing
    Public stream As MemoryStream = Nothing
    Public img As Image = Nothing
    Public sqlcmd As New MySqlCommand
    Public sql As String
    Public arrImage() As Byte = Nothing
    Public proFileImg As Boolean
    Public TargetFile As String
    Public ico As Icon
    '----------------email variables ----------------
    Public SendTo(1) As String
    Public FileAttach(10) As String
    Public strSubject As String
    Public strMessage As String
    Public ForceCloseSystem As Boolean = False
 

    Public Function rchar(ByVal str As String)
        str = str.Replace("'", "''")
        str = str.Replace("\", "\\")
        Return str
    End Function
    Public Sub loadIcons()
        TargetFile = Application.StartupPath + "\ico.ico"
        If File.Exists(TargetFile) = True Then
            ico = New Icon(TargetFile)
        End If
    End Sub
    Public Function Rowcount(ByVal tbl As String)
        Dim cnt As Integer = 0
        com.CommandText = "SELECT count(*) as cnt from " & tbl : rst = com.ExecuteReader()
        While rst.Read
            cnt = rst("cnt")
        End While
        rst.Close()
        Return cnt
    End Function

    Public Function qrysingledata(ByVal field As String, ByVal fqry As String, ByVal tblandqry As String)
        Dim def As String = ""
        com.CommandText = "select " & fqry & " from " & tblandqry : rst = com.ExecuteReader
        While rst.Read
            def = rst(field).ToString
        End While
        rst.Close()
        Return def
    End Function

    Public Function qryDate(ByVal field As String, ByVal fqry As String)
        Dim def As String = ""
        com.CommandText = "select " & fqry : rst = com.ExecuteReader
        While rst.Read
            def = rst(field).ToString
        End While
        rst.Close()
        Return def
    End Function

    Public Function countqry(ByVal tbl As String, ByVal cond As String)
        Dim cnt As Integer = 0
        com.CommandText = "select count(*) as cnt from " & tbl & " where " & cond
        rst = com.ExecuteReader
        While rst.Read
            cnt = Val(rst("cnt").ToString)
        End While
        rst.Close()
        Return cnt
    End Function

    Public Function CenterDashColumns(ByVal grdView As DataGridView) As DataGridView
        For Each clm As DataGridViewColumn In grdView.Columns
            If clm.Visible = True Then
                Dim visibility As Boolean = False
                For Each row As DataGridViewRow In grdView.Rows
                    If row.Cells(clm.Index).Value.ToString() = "-" Then
                        grdView.Columns(clm.Index).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                        grdView.Columns(clm.Index).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                        Exit For
                    End If
                Next
            End If
        Next
        Return grdView
    End Function

    Public Function UpdateImage(ByVal qry As String, ByVal fld As String, ByVal tbl As String, ByVal picbox As System.Windows.Forms.PictureBox)
        arrImage = Nothing
        Try
            If Not picbox.Image Is Nothing Then
                Dim mstream As New System.IO.MemoryStream()
                picbox.Image.Save(mstream, System.Drawing.Imaging.ImageFormat.Jpeg)
                arrImage = mstream.GetBuffer()
                mstream.Close()
            End If

            sql = "Update " & tbl & " set " & fld & " = @file where " & qry

            With sqlcmd
                .CommandText = sql
                .Connection = conn
                .Parameters.AddWithValue("@file", arrImage)
                .ExecuteNonQuery()
            End With
            sqlcmd.Parameters.Clear()

        Catch errMYSQL As MySqlException
            MessageBox.Show("Message:" & errMYSQL.Message & vbCrLf, _
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch errMS As Exception
            MessageBox.Show("Message:" & errMS.Message & vbCrLf, _
                              "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Return 0
    End Function

   
    Public Function ConvertImage(ByVal fld As String) As Image
        Try
            If rst(fld).ToString <> "" Then
                imgBytes = CType(rst(fld), Byte())
                stream = New MemoryStream(imgBytes, 0, imgBytes.Length)
                img = Image.FromStream(stream)
                ConvertImage = img
            Else
                ConvertImage = Nothing
            End If
        Catch errMYSQL As MySqlException
            MessageBox.Show("Message:" & errMYSQL.Message & vbCrLf, _
                             "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch errMS As Exception
            MessageBox.Show("Message:" & errMS.Message & vbCrLf, _
                              "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Return ConvertImage
    End Function

    Public Function ConvertDate(ByVal d As Date)
        Return d.ToString("yyyy-MM-dd")
    End Function
    Public Function ConvertTime(ByVal d As Date)
        Return d.ToString("HH:mm:ss")
    End Function
    Public Function ConvertDateTime(ByVal d As Date)
        Return d.ToString("yyyy-MM-dd HH:mm:ss")
    End Function
    Public Function countrecord(ByVal tbl As String)
        Dim cnt As Integer = 0
        com.CommandText = "select count(*) as cnt from " & tbl & " "
        rst = com.ExecuteReader
        While rst.Read
            cnt = rst("cnt")
        End While
        rst.Close()
        Return cnt
    End Function

    Public Function GridColumnAlignment(ByVal grdView As DataGridView, ByVal column As Array, ByVal alignment As DataGridViewContentAlignment) As DataGridView
        '   Dim array() As String = {a}
        For Each valueArr As String In column
            For i = 0 To grdView.ColumnCount - 1
                If valueArr = grdView.Columns(i).Name Then
                    grdView.Columns(i).DefaultCellStyle.Alignment = alignment
                    grdView.Columns(i).HeaderCell.Style.Alignment = alignment
                End If
            Next
        Next
        Return grdView
    End Function

    Public Function GridDisableColumn(ByVal grdView As DataGridView, ByVal column As Array) As DataGridView
        '   Dim array() As String = {a}
        For Each valueArr As String In column
            For i = 0 To grdView.ColumnCount - 1
                If valueArr = grdView.Columns(i).Name Then
                    grdView.Columns(i).ReadOnly = True
                    ' MyDataGridView_room.Rows(MyDataGridView_room.CurrentRow.Index).Cells("Rate Type").ReadOnly = False
                End If
            Next
        Next
        Return grdView
    End Function
    Public Function GridClearCell(ByVal grdView As DataGridView, ByVal column As Array, ByVal row As Integer, ByVal valuenumeric As Boolean)
        For Each valueArr As String In column
            For i = 0 To grdView.ColumnCount - 1
                If valueArr = grdView.Columns(i).Name Then
                    If valuenumeric = True Then
                        grdView.Item(i, row).Value = 0
                    Else
                        grdView.Item(i, row).Value = ""
                    End If
                End If
            Next
        Next
        Return grdView
    End Function
    Public Function GridCurrencyColumnDecimalCount(ByVal grdView As DataGridView, ByVal column As Array, ByVal decimalplaces As Integer) As DataGridView
        For Each valueArr As String In column
            For i = 0 To grdView.ColumnCount - 1
                If valueArr = grdView.Columns(i).Name Then
                    ' grdView.Columns(i).ValueType = System.Type.GetType("System.Decimal")
                    grdView.Columns(i).ValueType = GetType(Decimal)
                    grdView.Columns(i).DefaultCellStyle.Format = "n" & decimalplaces
                    grdView.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    grdView.Columns(i).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight

                End If
            Next
        Next
        Return grdView
    End Function
    Public Function GridCurrencyColumn(ByVal grdView As DataGridView, ByVal column As Array) As DataGridView
        For Each valueArr As String In column
            For i = 0 To grdView.ColumnCount - 1
                If valueArr = grdView.Columns(i).Name Then
                    ' grdView.Columns(i).ValueType = System.Type.GetType("System.Decimal")
                    grdView.Columns(i).ValueType = GetType(Decimal)
                    grdView.Columns(i).DefaultCellStyle.Format = "n2"
                    grdView.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    grdView.Columns(i).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight

                End If
            Next
        Next
        Return grdView
    End Function
    Public Function GridColumnWidth(ByVal grdView As DataGridView, ByVal column As Array, ByVal width As Double) As DataGridView
        For Each valueArr As String In column
            For i = 0 To grdView.ColumnCount - 1
                If valueArr = grdView.Columns(i).Name Then
                    grdView.Columns(i).Width = width
                End If
            Next
        Next
        Return grdView
    End Function
    Public Function GridColumAutonWidth(ByVal grdView As DataGridView, ByVal column As Array) As DataGridView
        For Each valueArr As String In column
            For i = 0 To grdView.ColumnCount - 1
                If valueArr = grdView.Columns(i).Name Then
                    grdView.Columns(i).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                End If
            Next
        Next
        Return grdView
    End Function
    Public Function GridNumberColumn(ByVal grdView As DataGridView, ByVal column As Array) As DataGridView
        For Each valueArr As String In column
            For i = 0 To grdView.ColumnCount - 1
                If valueArr = grdView.Columns(i).Name Then
                    ' grdView.Columns(i).ValueType = System.Type.GetType("System.Decimal")
                    grdView.Columns(i).ValueType = GetType(Decimal)
                    grdView.Columns(i).DefaultCellStyle.Format = "n0"
                    grdView.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    grdView.Columns(i).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

                End If
            Next
        Next
        Return grdView
    End Function
    Public Function GridHideColumn(ByVal grdView As DataGridView, ByVal column As Array) As DataGridView
        For Each valueArr As String In column
            For i = 0 To grdView.ColumnCount - 1
                If valueArr = grdView.Columns(i).Name Then
                    grdView.Columns(i).Visible = False
                End If
            Next
        Next
        Return grdView
    End Function
    Public Sub GridColumnChoosed(ByVal grdView As DataGridView, ByVal file_dir As String)
        If System.IO.File.Exists(Application.StartupPath & "\Config\" & file_dir) = True Then
            Dim sr As StreamReader = File.OpenText(Application.StartupPath & "\Config\" & file_dir)
            Try
                Dim columnChoosed As String = sr.ReadLine()
                For Each col In grdView.Columns
                    If Not DecryptTripleDES(columnChoosed).Contains(col.Name) Then
                        col.Visible = False
                    End If
                Next
                sr.Close()
            Catch errMS As Exception
                ' System.IO.File.Delete(Application.StartupPath & "\" & file_dir)
            End Try
        End If
    End Sub
    Public Function LoadToComboBox(ByVal cb As Windows.Forms.ComboBox, ByVal path As String)
        Dim chpname As String = ""
        Dim strSetup As String = ""
        Dim sr As StreamReader = File.OpenText(path)
        Do While sr.Peek() >= 0
            Dim description As String = "" : Dim id As String = "" : Dim cnt As Integer = 0
            For Each word In sr.ReadLine().Split(New Char() {"|"c})
                If cnt = 0 Then
                    id = word
                ElseIf cnt = 1 Then
                    description = word
                End If
                cnt = cnt + 1
            Next
            If id <> "" Then
                cb.Items.Add(New ComboBoxItem(description, id))
            End If
        Loop
        sr.Close()
        Return 0
    End Function
    Public Function ShowGridTotalSummary(ByVal captionLocation As String, ByVal totalColumn As String, ByVal grdView As DataGridView)
        grdView.AllowUserToAddRows = True
        grdView.Columns(totalColumn).Width = 200
        If grdView.RowCount - 1 > 0 Then
            Dim totalsum As Double = 0
            For i = 0 To grdView.RowCount - 1
                totalsum = totalsum + grdView.Rows(i).Cells(totalColumn).Value()
            Next
            If captionLocation.Length > 0 Then
                grdView.Rows(grdView.RowCount - 1).Cells(captionLocation).Value = "Total"
            End If
            grdView.Rows(grdView.RowCount - 1).Cells(totalColumn).Value = totalsum
            grdView.Rows(grdView.RowCount - 1).DefaultCellStyle.BackColor = Color.Red
            grdView.Rows(grdView.RowCount - 1).DefaultCellStyle.ForeColor = Color.White
        End If
    End Function
    Public Function LoadToComboBoxTxt(ByVal cb As Windows.Forms.ComboBox, ByVal path As String)
        Dim chpname As String = ""
        Dim strSetup As String = ""
        Dim sr As StreamReader = File.OpenText(path)
        cb.Items.Clear()
        Do While sr.Peek() >= 0
            For Each word In sr.ReadLine().Split(New Char() {vbCrLf})
                cb.Items.Add(word)
            Next
        Loop
        sr.Close()
        Return 0
    End Function
    Public Function LoadToComboBoxDBWithID(ByVal cb As Windows.Forms.ComboBox, ByVal path As String)
        Dim chpname As String = ""
        Dim strSetup As String = ""
        Dim sr As StreamReader = File.OpenText(path)
        cb.Items.Clear()
        Do While sr.Peek() >= 0
            For Each word In sr.ReadLine().Split(New Char() {vbCrLf})
                If word <> "" Then
                    cb.Items.Add(New ComboBoxItem(word.Split("|".ToCharArray)(0), word.Split("|".ToCharArray)(1)))
                End If
            Next
        Loop
        sr.Close()
        Return 0
    End Function
    Public Function LoadToComboBoxDB(ByVal query As String, ByVal fields As String, ByVal id As String, ByVal cb As Windows.Forms.ComboBox)
        cb.Items.Clear()
        com.CommandText = query : rst = com.ExecuteReader
        While rst.Read
            If rst(fields).ToString <> "" Then
                cb.Items.Add(New ComboBoxItem(rst(fields).ToString, rst(id.ToString)))
            End If
        End While
        rst.Close()
        Return 0
    End Function
    Public Function CC(ByVal m As String)
        If m <> "" Then
            CC = Val(m.Replace(",", ""))
        End If
        Return CC
    End Function
 
   
    Public Function LoadToGridComboBox(ByVal query As String, ByVal fields As String, ByVal cb As Windows.Forms.DataGridViewComboBoxColumn)
        cb.Items.Clear()
        com.CommandText = query : rst = com.ExecuteReader
        While rst.Read
            If rst(fields).ToString <> "" Then
                cb.Items.Add(rst(fields).ToString)
            End If
        End While
        rst.Close()
        Return 0
    End Function
    Public Function LoadToGridComboBoxCell(ByVal columnname As String, ByVal rowIndex As Integer, ByVal query As String, ByVal fields As String, ByVal allowBlankRow As Boolean, ByVal gridview As DataGridView)
        Dim dgvcc As New DataGridViewComboBoxCell
        dgvcc.Items.Clear()
        If allowBlankRow = True Then
            dgvcc.Items.Add("")
        End If
        com.CommandText = query : rst = com.ExecuteReader
        While rst.Read
            If rst(fields).ToString <> "" Then
                dgvcc.Items.Add(rst(fields).ToString)
            End If
        End While
        rst.Close()
        gridview.Item(columnname, rowIndex) = dgvcc
        Return 0
    End Function
    Public Function getStockhouseid()
        Dim strng = ""

        If CInt(countrecord("tblstockhouse")) = 0 Then
            strng = "H100001"
        Else
            com.CommandText = "select stockhouseid from tblstockhouse order by right(stockhouseid,6) desc limit 1" : rst = com.ExecuteReader()
            Dim removechar As Char() = "H".ToCharArray()
            Dim sb As New System.Text.StringBuilder
            While rst.Read
                Dim str As String = rst("stockhouseid").ToString
                For Each ch As Char In str
                    If Array.IndexOf(removechar, ch) = -1 Then
                        sb.Append(ch)
                    End If
                Next
            End While
            rst.Close()
            strng = "H" & Val(sb.ToString) + 1
        End If
        Return strng.ToString
    End Function
    Public Function getClientid()
        Dim compid = ""

        If CInt(countrecord("tblclientaccounts")) = 0 Then
            compid = "CIF1000001"
        Else
            com.CommandText = "select cifid from tblclientaccounts order by right(cifid,7) desc limit 1" : rst = com.ExecuteReader()
            Dim removechar As Char() = "ABCDEFGHIJKLMNOPQRSTUVWXYZ- ".ToCharArray()
            Dim sb As New System.Text.StringBuilder
            While rst.Read
                Dim str As String = rst("cifid").ToString
                For Each ch As Char In str
                    If Array.IndexOf(removechar, ch) = -1 Then
                        sb.Append(ch)
                    End If
                Next
            End While
            rst.Close()
            compid = "CIF" & Val(sb.ToString) + 1
        End If
        Return compid.ToString
    End Function

    Public Function getproid()
        Dim strng As Integer = 0 : Dim newprocode As String = ""
        If CInt(countrecord("tblglobalproductssequence")) = 0 Then
            If CInt(countrecord("tblglobalproducts")) = 0 Then
                strng = 1000001
            Else
                com.CommandText = "select productid from tblglobalproducts order by productid desc limit 1" : rst = com.ExecuteReader()
                While rst.Read
                    strng = Val(rst("productid").ToString) + 1
                End While
                rst.Close()
            End If
        Else
            com.CommandText = "select productid from tblglobalproductssequence" : rst = com.ExecuteReader()
            While rst.Read
                strng = Val(rst("productid").ToString) + 1
            End While
            rst.Close()
        End If
        com.CommandText = "delete from tblglobalproductssequence" : com.ExecuteNonQuery()
        com.CommandText = "insert into tblglobalproductssequence set productid='" & strng & "'" : com.ExecuteNonQuery()
        newprocode = strng.ToString
        Return newprocode
    End Function

    Public Function InputNumberOnly(ByVal textbox As System.Windows.Forms.TextBox, e As KeyPressEventArgs)
        Dim keyChar = e.KeyChar
        If Char.IsControl(keyChar) Then
        ElseIf Char.IsDigit(keyChar) OrElse keyChar = "."c Then
            Dim text = textbox.Text
            Dim selectionStart = textbox.SelectionStart
            Dim selectionLength = textbox.SelectionLength
            text = text.Substring(0, selectionStart) & keyChar & text.Substring(selectionStart + selectionLength)

            If Integer.TryParse(text, New Integer) AndAlso text.Length > 16 Then
                'Reject an integer that is longer than 16 digits.
                e.Handled = True
            End If
        Else
            'Reject all other characters.
            e.Handled = True
        End If
    End Function
    Public Function RemoveEmptyColumns(ByVal grdView As DataGridView) As DataGridView
        For Each clm As DataGridViewColumn In grdView.Columns
            Dim visibility As Boolean = False
            For Each row As DataGridViewRow In grdView.Rows
                If row.Cells(clm.Index).Value.ToString() <> String.Empty Or Val(row.Cells(clm.Index).Value.ToString()) <> 0 Then
                    visibility = True
                    Exit For
                End If
            Next
            ' MsgBox(clm.Name)
            If clm.Name <> "id" And clm.Name <> "productid" Then
                grdView.Columns(clm.Name).Visible = visibility
            End If
        Next
        Return grdView
    End Function


    Public Function getvendorid()
        Dim compid = ""
        If CInt(countrecord("tblglobalvendor")) = 0 Then
            compid = "SPR-1001"
        Else
            com.CommandText = "select vendorid from tblglobalvendor order by right(vendorid,4) desc limit 1" : rst = com.ExecuteReader()
            Dim removechar As Char() = "ABCDEFGHIJKLMNOPQRSTUVWXYZ- ".ToCharArray()
            Dim sb As New System.Text.StringBuilder
            While rst.Read
                Dim str As String = rst("vendorid").ToString
                For Each ch As Char In str
                    If Array.IndexOf(removechar, ch) = -1 Then
                        sb.Append(ch)
                    End If
                Next
            End While
            rst.Close()
            compid = "SPR-" & Val(sb.ToString) + 1
        End If
        Return compid.ToString
    End Function
    Public Function RecordApprovingHistory(ByVal approvalDescription As String, ByVal mainreference As String, ByVal refno As String, ByVal title As String, ByVal remarks As String)
        com.CommandText = "INSERT INTO `tblapprovalhistory` set approvaltype='-', " _
                                                                 + " appdescription='" & approvalDescription & "', " _
                                                                 + " mainreference='" & mainreference & "', " _
                                                                 + " referenceno='" & refno & "', " _
                                                                 + " status='" & rchar(title) & "', " _
                                                                 + " remarks='" & rchar(remarks) & "', " _
                                                                 + " apptitle='" & globalposition & "', " _
                                                                 + " applevel='-', " _
                                                                 + " confirmid='" & globaluserid & "', " _
                                                                 + " confirmby='" & globalfullname & "', " _
                                                                 + " position='" & globalposition & "', " _
                                                                 + " dateconfirm=current_timestamp, " _
                                                                 + " finalapprover=0 " : com.ExecuteNonQuery()
        Return 0
    End Function

    Public Function GetFFECode(ByVal officeid As String, ByVal ffecode As String)
        Dim strng = ""
        If CInt(countrecord("tblinventoryffe")) = 0 Then
            strng = ffecode & "100001"
        Else
            com.CommandText = "select ffecode from tblinventoryffe where officeid='" & officeid & "' order by right(ffecode,6) desc limit 1" : rst = com.ExecuteReader()
            Dim removechar As Char() = "ABCDEFGHIJKLMNOPQRSTUVWXYZ-".ToCharArray()
            Dim sb As New System.Text.StringBuilder
            While rst.Read
                Dim str As String = rst("ffecode").ToString
                For Each ch As Char In str
                    If Array.IndexOf(removechar, ch) = -1 Then
                        sb.Append(ch)
                    End If
                Next
            End While
            rst.Close()
            strng = ffecode & Val(sb.ToString) + 1
        End If
        Return strng.ToString
    End Function
    Public Function LessQuantityConsumable(ByVal productid As String, ByVal quantity As Double, ByVal officeid As String)
        Dim strquery As String = "" : Dim remquantity As Double = 0
        com.CommandText = "select * from tblinventory where productid='" & productid & "' order by dateinventory,trnid asc" : rst = com.ExecuteReader
        While rst.Read
            If remquantity = 0 Then
                If quantity > Val(rst("quantity").ToString) Then
                    remquantity = quantity - Val(rst("quantity").ToString)
                    strquery = "update tblinventory set quantity=quantity-" & Val(rst("quantity").ToString) & " where trnid='" & rst("trnid").ToString & "' and officeid='" & officeid & "';" & Chr(13)
                Else
                    strquery = "update tblinventory set quantity=quantity-" & quantity & " where trnid='" & rst("trnid").ToString & "' and officeid='" & officeid & "';" & Chr(13)
                End If
            Else
                If remquantity > Val(rst("quantity").ToString) Then
                    remquantity = remquantity - Val(rst("quantity").ToString)
                    strquery += "update tblinventory set quantity=quantity-" & Val(rst("quantity").ToString) & " where trnid='" & rst("trnid").ToString & "' and officeid='" & officeid & "';" & Chr(13)
                Else
                    strquery += "update tblinventory set quantity=quantity-" & remquantity & " where trnid='" & rst("trnid").ToString & "' and officeid='" & officeid & "';" & Chr(13)
                End If
            End If
        End While
        rst.Close()
        MsgBox(strquery)
        Return 0
    End Function
  
    Public Function RemoveSpecialCharacter(ByVal str As String)
        Dim removechar As Char() = "!@#$%^&*()_+-={}|[]\:;'<>?/".ToCharArray()
        Dim sb As New System.Text.StringBuilder
        For Each ch As Char In str
            If Array.IndexOf(removechar, ch) = -1 Then
                sb.Append(ch)
            End If
        Next
        Return sb.ToString
    End Function
    Public Function RemoveFilenameCharacter(ByVal str As String)
        Dim removechar As Char() = "!@#$%^&*+={}|[]\:;'<>?/".ToCharArray()
        Dim sb As New System.Text.StringBuilder
        For Each ch As Char In str
            If Array.IndexOf(removechar, ch) = -1 Then
                sb.Append(ch)
            End If
        Next
        Return sb.ToString
    End Function
    Public Function PopulateGridViewControls(ByVal ColumnName As String, ByVal ColumnWidth As Double, ByVal ColumnType As String, ByVal gridview As DataGridView, ByVal visible As Boolean, ByVal readonlycolumn As Boolean)
        If ColumnType = "COMBO" Then
            Dim dgcmbcol As DataGridViewComboBoxColumn
            dgcmbcol = New DataGridViewComboBoxColumn
            dgcmbcol.HeaderText = ColumnName
            dgcmbcol.Width = ColumnWidth
            dgcmbcol.Name = ColumnName
            dgcmbcol.ReadOnly = False
            dgcmbcol.AutoComplete = False
            dgcmbcol.FlatStyle = FlatStyle.System
            gridview.Columns.Add(dgcmbcol)

        ElseIf ColumnType = "CHECKBOX" Then
            Dim colCheckbox As New DataGridViewCheckBoxColumn()
            colCheckbox.AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
            colCheckbox.ThreeState = False
            colCheckbox.TrueValue = 1
            colCheckbox.FalseValue = 0
            colCheckbox.IndeterminateValue = System.DBNull.Value
            colCheckbox.HeaderText = ColumnName
            colCheckbox.Width = ColumnWidth
            colCheckbox.Name = ColumnName
            colCheckbox.ReadOnly = False
            gridview.Columns.Add(colCheckbox)
        Else
            Dim dgcmbcol As DataGridViewColumn
            dgcmbcol = New DataGridViewColumn
            dgcmbcol.HeaderText = ColumnName
            dgcmbcol.Width = ColumnWidth
            dgcmbcol.Name = ColumnName
            dgcmbcol.CellTemplate = New DataGridViewTextBoxCell
            gridview.Columns.Add(dgcmbcol)
        End If
        gridview.Columns(ColumnName).Visible = visible
        If readonlycolumn = True Then
            gridview.Columns(ColumnName).ReadOnly = True
            gridview.Columns(ColumnName).DefaultCellStyle.BackColor = Color.LemonChiffon
            gridview.Columns(ColumnName).DefaultCellStyle.SelectionBackColor = Color.LemonChiffon
            gridview.Columns(ColumnName).DefaultCellStyle.SelectionForeColor = Color.Black
        Else
            gridview.Columns(ColumnName).ReadOnly = False
            gridview.Columns(ColumnName).DefaultCellStyle.BackColor = Color.White
            gridview.Columns(ColumnName).DefaultCellStyle.SelectionForeColor = Color.Black
        End If

        Return 0
    End Function


    Public Function PopulateGridViewColumns(ByVal ColumnName As String, ByVal ColumnWidth As Double, ByVal ComboBoxColumn As Boolean, ByVal gridview As DataGridView, ByVal visible As Boolean, ByVal readonlycolumn As Boolean)
        If ComboBoxColumn = True Then
            Dim dgcmbcol As DataGridViewComboBoxColumn
            dgcmbcol = New DataGridViewComboBoxColumn
            dgcmbcol.HeaderText = ColumnName
            dgcmbcol.Width = ColumnWidth
            dgcmbcol.Name = ColumnName
            dgcmbcol.ReadOnly = False
            dgcmbcol.AutoComplete = False
            dgcmbcol.FlatStyle = FlatStyle.Flat
            gridview.Columns.Add(dgcmbcol)
        Else
            Dim dgcmbcol As DataGridViewColumn
            dgcmbcol = New DataGridViewColumn
            dgcmbcol.HeaderText = ColumnName
            dgcmbcol.Width = ColumnWidth
            dgcmbcol.Name = ColumnName
            dgcmbcol.CellTemplate = New DataGridViewTextBoxCell
            gridview.Columns.Add(dgcmbcol)
        End If
        gridview.Columns(ColumnName).Visible = visible
        If readonlycolumn = True Then
            gridview.Columns(ColumnName).ReadOnly = True
            gridview.Columns(ColumnName).DefaultCellStyle.BackColor = Color.LemonChiffon
            gridview.Columns(ColumnName).DefaultCellStyle.SelectionBackColor = Color.LemonChiffon

        Else
            gridview.Columns(ColumnName).ReadOnly = False
            gridview.Columns(ColumnName).DefaultCellStyle.BackColor = Color.White
        End If

        Return 0
    End Function

    Public Function LoadComboSuggestion(ByVal ColumnName As String, ByVal gridview As DataGridView, ByVal suggestionQuery As String, ByVal suggestionValue As String)
        Dim dgvcc As DataGridViewComboBoxColumn
        dgvcc = gridview.Columns(ColumnName)
        If suggestionQuery <> "" Then
            LoadToGridComboBox(suggestionQuery, suggestionValue, dgvcc)
        End If
        Return 0
    End Function

    Public Sub ExportGridToExcel(ByVal filename As String, ByVal dst As DataSet)
        Dim f As FolderBrowserDialog = New FolderBrowserDialog
        Try
            If f.ShowDialog() = DialogResult.OK Then
                dst.WriteXml(f.SelectedPath & "\" & LCase(filename) & ".xls")
                MessageBox.Show(LCase(filename) & ".xls successfully exported!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Warning", MessageBoxButtons.OK)
        End Try
    End Sub

    Public Sub ColumnGridSetup(ByVal setupName As String, ByVal gridview As DataGridView, ByVal form As Form)
        Dim colname As String = ""
        For i = 0 To gridview.ColumnCount - 1
            colname += gridview.Columns(i).Name & ","
        Next
        
    End Sub

    Public Function getIncrementID(ByVal tableName As String) As String
        getIncrementID = ""
        com.CommandText = "SELECT `AUTO_INCREMENT` as ID FROM  INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = '" & sqldatabase & "' AND TABLE_NAME   = '" & tableName & "';" : rst = com.ExecuteReader
        While rst.Read
            getIncrementID = rst("ID").ToString
        End While
        rst.Close()
        com.CommandText = "ALTER TABLE `" & tableName & "` AUTO_INCREMENT = " & Val(getIncrementID) + 1 & ";" : com.ExecuteNonQuery()
        Return getIncrementID
    End Function
    Public Function getcodeid(ByVal code As String, ByVal table As String, ByVal initialcode As String)
        Dim strng = ""
        Try
            If CInt(countrecord(table)) = 0 Then
                strng = initialcode
            Else
                com.CommandText = "select (" & code & " + 1) as code from " & table & " order by " & code & " desc limit 1" : rst = com.ExecuteReader()
                While rst.Read
                    strng = rst("code").ToString
                End While
                rst.Close()
            End If
        Catch errMYSQL As MySqlException
            MessageBox.Show("Message:" & errMYSQL.Message & vbCrLf, _
                             "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Catch errMS As Exception
            MessageBox.Show("Message:" & errMS.Message & vbCrLf, _
                              "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
        Return strng.ToString
    End Function

    Public Function RGBColorConverter(ByVal RGBString As String) As Color
        Dim RGB As String() = RGBString.Split(",")
        If RGBString.Length > 0 Then
            RGBColorConverter = System.Drawing.Color.FromArgb(CType(CType(Val(RGB(0)), Byte), Integer), CType(CType(Val(RGB(1)), Byte), Integer), CType(CType(Val(RGB(2)), Byte), Integer))
        Else
            RGBColorConverter = Color.Black
        End If
    End Function

    Public Sub PrintDatagridview(ByVal ReportTitle As String, ByVal TableTitle As String, ByVal ReportDescription As String, ByVal gv As DataGridView, ByVal form As Form)
        If gv.RowCount = 0 Then
            MessageBox.Show("No data report to print!", _
                       "Error Print", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        
        Dim Template As String = Application.StartupPath.ToString & "\Templates\StandardReports.html"
        Dim SaveLocation As String = Application.StartupPath.ToString & "\Transaction\REPORTS\" & RemoveSpecialCharacter(form.Text) & ".html"
        If System.IO.File.Exists(SaveLocation) = True Then
            System.IO.File.Delete(SaveLocation)
        End If
        My.Computer.FileSystem.CopyFile(Template, SaveLocation)
        If System.IO.File.Exists(Application.StartupPath.ToString & "\Logo\logo.png") = True Then
            My.Computer.FileSystem.WriteAllText(SaveLocation, My.Computer.FileSystem.ReadAllText(SaveLocation).Replace("[logo]", "<img style='float:left;  position: absolute;' src='" & Application.StartupPath.ToString.Replace("\", "/") & "/Logo/logo.png'>"), False)
        Else
            My.Computer.FileSystem.WriteAllText(SaveLocation, My.Computer.FileSystem.ReadAllText(SaveLocation).Replace("[logo]", ""), False)
        End If
        My.Computer.FileSystem.WriteAllText(SaveLocation, My.Computer.FileSystem.ReadAllText(SaveLocation).Replace("[report header]", PrintSetupHeader()), False)
        My.Computer.FileSystem.WriteAllText(SaveLocation, My.Computer.FileSystem.ReadAllText(SaveLocation).Replace("[title]", UCase(ReportTitle)), False)

        Dim ReportDetails As String = "<p align='left'> " & If(ReportDescription = "", "&nbsp;", ReportDescription) & " <span style='float:right'>Date Report: " & CDate(Now).ToString("MMMM dd, yyyy") & "</span></p>"
        My.Computer.FileSystem.WriteAllText(SaveLocation, My.Computer.FileSystem.ReadAllText(SaveLocation).Replace("[report details]", ReportDetails), False)
        My.Computer.FileSystem.WriteAllText(SaveLocation, My.Computer.FileSystem.ReadAllText(SaveLocation).Replace("[report table]", SetupTheGridviewPrinter(TableTitle, gv)), False)
        My.Computer.FileSystem.WriteAllText(SaveLocation, My.Computer.FileSystem.ReadAllText(SaveLocation).Replace("[prepared by]", UCase(globalfullname)), False)
        My.Computer.FileSystem.WriteAllText(SaveLocation, My.Computer.FileSystem.ReadAllText(SaveLocation).Replace("[prepared position]", UCase(globalposition)), False)

       

        ' Me.WindowState = FormWindowState.Minimized
        PrintViaInternetExplorer(SaveLocation.Replace("\", "/"), form)
    End Sub
    Public Function SetupTheGridviewPrinter(ByVal TableTitle As String, ByVal gv As DataGridView) As String
        On Error Resume Next
        Dim TableHeaderStart As String = "" : Dim TableHeaderEnd As String = "" : Dim ColumnName As String = "" : Dim rows As String = "" : Dim rowRowTableData As String = "" : Dim RowFooter As String = ""
        TableHeaderStart = "<table border='1' style='min-width:650px; margin:3px 0' cellpadding='4' cellspacing='0' style='border-collapse:collapse;'> " _
                                       + " <tr> " _
                                           + " <td colspan='" & gv.ColumnCount & "' align='center'><b>" & UCase(TableTitle) & "</b></td> " _
                                       + " </tr> " & Chr(13) _
                                       + " <tr> "
        For i = 0 To gv.ColumnCount - 1
            If gv.Columns(i).Visible = True Then
                ColumnName += "<th>" & gv.Columns(i).Name & "</th>"
            End If
        Next
        TableHeaderEnd = " </tr> "
        For i = 0 To gv.RowCount - 1
            rowRowTableData = "" : Dim rowColor As String = ""
            For x = 0 To gv.ColumnCount - 1
                If gv.Columns(x).Visible = True Then
                    Dim colalignment As String = "" : Dim strvalue As String = "" : Dim columnSize As String = ""
                    If gv.Columns(gv.Columns(x).Name).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter Then
                        colalignment = "align='center'"

                        If gv.Item(gv.Columns(x).Name, i).Value() Is Nothing Then
                            strvalue = "&nbsp;"
                        Else
                            strvalue = gv.Item(gv.Columns(x).Name, i).Value().ToString
                        End If

                    ElseIf gv.Columns(gv.Columns(x).Name).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight Then
                        colalignment = "align='right'"
                        If gv.Item(gv.Columns(x).Name, i).Value().ToString = "" Then
                            strvalue = "&nbsp;"
                        Else
                            strvalue = FormatNumber(gv.Item(gv.Columns(x).Name, i).Value().ToString, 2)
                        End If
                    Else
                        If gv.Item(gv.Columns(x).Name, i).Value() Is Nothing Then
                            strvalue = "&nbsp;"
                        Else
                            strvalue = gv.Item(gv.Columns(x).Name, i).Value().ToString
                        End If
                    End If
                    If gv.Columns(x).Width = 300 Then
                        columnSize = " width='" & gv.Columns(x).Width.ToString & "' "
                    End If

                    rowRowTableData += "<td " & colalignment & columnSize & ">" & strvalue.Replace("True", "YES").Replace("False", "-").Replace(Chr(13), "<br/>") & "</td> "
                End If
            Next
            If gv.Rows(i).DefaultCellStyle.ForeColor = Color.Red Then
                rowColor = "ff0000"
            ElseIf gv.Rows(i).DefaultCellStyle.ForeColor = Color.Blue Then
                rowColor = "001a7a"
            Else
                rowColor = "000000"
            End If
            rows += "<tr style='color:#" & rowColor & "'>" + rowRowTableData + "</tr>"
        Next
        SetupTheGridviewPrinter = TableHeaderStart + ColumnName + TableHeaderEnd + rows + "</table>"
        Return SetupTheGridviewPrinter
    End Function

    Public Function PrintSetupHeader()
        PrintSetupHeader += "<p align='center' ><strong>WIN E-SOFT TECHNOLOGIES</strong></br>" _
                            + "Bldg. #044, Gomez St. Corner Lacaya St. Biasong Dipolog City, ZN<br/> " _
                            + "065 908 3550 / 065 212 1727<br/> "
        PrintSetupHeader += "<p align='center'><b>Accounting Dept.</b></br>"

        Return PrintSetupHeader
    End Function
    Public Sub PrintViaInternetExplorer(ByVal location As String, ByVal form As Windows.Forms.Form)
        Try
            System.Diagnostics.Process.Start(location)
            'form.WindowState = FormWindowState.Minimized
        Catch ex As Exception
            MessageBox.Show("File not found!", _
                          "Error Reprint Transaction", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Module
