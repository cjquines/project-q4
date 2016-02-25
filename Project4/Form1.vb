Public Class Form1
    Dim connect As New OleDb.OleDbConnection
    Dim sql As String
    Dim ds As New DataSet
    Dim da As OleDb.OleDbDataAdapter

    Dim rows As Integer
    Dim index As Integer

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        connect.ConnectionString = "Provider=Microsoft.jet.OLEDB.4.0; data source = |datadirectory|/Quines.mdb"
        connect.Open()
        sql = "select * from tblStudents"
        da = New OleDb.OleDbDataAdapter(sql, connect)
        da.Fill(ds, "tblStudents")
        connect.Close()

        rows = ds.Tables("tblStudents").Rows.Count
        index = 0
        navigate_records(index)
    End Sub

    Private Sub navigate_records(index As Integer)
        txtId.Text = ds.Tables("tblStudents").Rows(index).Item(0)
        txtName.Text = ds.Tables("tblStudents").Rows(index).Item(1)
        txtSection.Text = ds.Tables("tblStudents").Rows(index).Item(2)
        Select Case ds.Tables("tblStudents").Rows(index).Item(3)
            Case 1
                rdoSubject1.Checked = True
                rdoSubject2.Checked = False
                rdoSubject3.Checked = False
            Case 2
                rdoSubject1.Checked = False
                rdoSubject2.Checked = True
                rdoSubject3.Checked = False
            Case 3
                rdoSubject1.Checked = False
                rdoSubject2.Checked = False
                rdoSubject3.Checked = True
        End Select
        txtWwRaw.Clear()
        txtWwTotal.Clear()
        txtPtRaw.Clear()
        txtPtTotal.Clear()
        txtQaRaw.Clear()
        txtQaTotal.Clear()
        txtComputed.Text = ds.Tables("tblStudents").Rows(index).Item(4)
        txtDown.Text = ds.Tables("tblStudents").Rows(index).Item(5)
        txtTotal.Text = ds.Tables("tblStudents").Rows(index).Item(6)
        Select Case ds.Tables("tblStudents").Rows(index).Item(7)
            Case 1
                rdoNumber1.Checked = True
                rdoNumber2.Checked = False
                rdoNumber4.Checked = False
                rdoNumber10.Checked = False
                rdoNumberOther.Checked = False
                txtOther.Clear()
                txtOther.Enabled = False
            Case 2
                rdoNumber1.Checked = False
                rdoNumber2.Checked = True
                rdoNumber4.Checked = False
                rdoNumber10.Checked = False
                rdoNumberOther.Checked = False
                txtOther.Clear()
                txtOther.Enabled = False
            Case 4
                rdoNumber1.Checked = False
                rdoNumber2.Checked = False
                rdoNumber4.Checked = True
                rdoNumber10.Checked = False
                rdoNumberOther.Checked = False
                txtOther.Clear()
                txtOther.Enabled = False
            Case 10
                rdoNumber1.Checked = False
                rdoNumber2.Checked = False
                rdoNumber4.Checked = False
                rdoNumber10.Checked = True
                rdoNumberOther.Checked = False
                txtOther.Clear()
                txtOther.Enabled = False
            Case Else
                rdoNumber1.Checked = False
                rdoNumber2.Checked = False
                rdoNumber4.Checked = False
                rdoNumber10.Checked = False
                rdoNumberOther.Checked = True
                txtOther.Text = ds.Tables("tblStudents").Rows(index).Item(7)
                txtOther.Enabled = True
        End Select
        txtContrib.Text = ds.Tables("tblStudents").Rows(index).Item(8)
    End Sub

    Private Sub btnCalculate_Click(sender As Object, e As EventArgs) Handles btnCalculate.Click
        If txtId.Text = "" Or txtName.Text = "" Or txtSection.Text = "" Then
            MsgBox("Please fill in student ID, name, and section.")
        ElseIf Not (rdoSubject1.Checked Or rdoSubject2.Checked Or rdoSubject3.Checked) Then
            MsgBox("Please select subject.")
        ElseIf txtWwRaw.Text <> "" Or txtWwTotal.Text <> "" Or txtPtRaw.Text <> "" Or
           txtPtTotal.Text <> "" Or txtQaRaw.Text <> "" Or txtQaTotal.Text <> "" Then
            If txtWwRaw.Text = "" Or txtWwTotal.Text = "" Or txtPtRaw.Text = "" Or
               txtPtTotal.Text = "" Or txtQaRaw.Text = "" Or txtQaTotal.Text = "" Then
                MsgBox("Please fill in all raw scores and number of items.")
            ElseIf Not (IsNumeric(txtWwRaw.Text) And IsNumeric(txtWwTotal.Text) And IsNumeric(txtPtRaw.Text) And
                        IsNumeric(txtPtTotal.Text) And IsNumeric(txtQaRaw.Text) And IsNumeric(txtQaTotal.Text)) Then
                MsgBox("Raw scores and number of items must be integers.")
            ElseIf Val(txtWwRaw.Text) < 0 Or Val(txtPtRaw.Text) < 0 Or Val(txtQaRaw.Text) < 0 Or
                   Val(txtWwTotal.Text) <= 0 Or Val(txtPtTotal.Text) <= 0 Or Val(txtQaTotal.Text) <= 0 Then
                MsgBox("Raw score and number of items must be positive.")
            ElseIf Val(txtWwRaw.Text) > Val(txtWwTotal.Text) Or Val(txtPtRaw.Text) > Val(txtPtTotal.Text) Or
                   Val(txtQaRaw.Text) > Val(txtQaTotal.Text) Then
                MsgBox("Raw score must be less than or equal to number of items.")
            End If
        ElseIf Not IsNumeric(txtComputed.Text) Then
            MsgBox("Computed grade must be numeric.")
        ElseIf txtDown.Text = "" Or txtTotal.Text = "" Or
               Not (rdoNumber1.Checked Or rdoNumber2.Checked Or rdoNumber4.Checked Or
                    rdoNumber10.Checked Or rdoNumberOther.Checked) Or
               (rdoNumberOther.Checked And txtOther.Text = "") Then
            MsgBox("Please fill in downpayment, total, and select number of payments.")
        ElseIf Not ((IsNumeric(txtDown.Text) And IsNumeric(txtTotal.Text))) Then
            MsgBox("Downpayment and total must be integers.")
        ElseIf Val(txtDown.Text) < 1 Or Val(txtTotal.Text) < 0 Then
            MsgBox("Downpayment must be greater than 1, and total must be positive.")
        ElseIf Val(txtDown.Text) > Val(txtTotal.Text) Then
            MsgBox("Downpayment must be less than or equal to total payment.")
        ElseIf rdoNumberOther.Checked And (Not IsNumeric(txtOther.Text) Or (txtOther.Text) <= 0) Then
            MsgBox("Number of payments must be positive integer.")
        End If

        If rdoSubject1.Checked Then
            txtComputed.Text = transmutation(0.3 * (Val(txtWwRaw.Text) / Val(txtWwTotal.Text)) +
                                             0.5 * (Val(txtPtRaw.Text) / Val(txtPtTotal.Text)) +
                                             0.2 * (Val(txtQaRaw.Text) / Val(txtQaTotal.Text)))
        ElseIf rdoSubject2.Checked Then
            txtComputed.Text = transmutation(0.4 * (Val(txtWwRaw.Text) / Val(txtWwTotal.Text)) +
                                             0.4 * (Val(txtPtRaw.Text) / Val(txtPtTotal.Text)) +
                                             0.2 * (Val(txtQaRaw.Text) / Val(txtQaTotal.Text)))
        ElseIf rdoSubject3.Checked Then
            txtComputed.Text = transmutation(0.2 * (Val(txtWwRaw.Text) / Val(txtWwTotal.Text)) +
                                             0.6 * (Val(txtPtRaw.Text) / Val(txtPtTotal.Text)) +
                                             0.2 * (Val(txtQaRaw.Text) / Val(txtQaTotal.Text)))
        End If

        If rdoNumber1.Checked Then
            txtContrib.Text = Math.Ceiling((Val(txtTotal.Text) - Val(txtDown.Text)) / 1)
        ElseIf rdoNumber2.Checked Then
            txtContrib.Text = Math.Ceiling((Val(txtTotal.Text) - Val(txtDown.Text)) / 2)
        ElseIf rdoNumber4.Checked Then
            txtContrib.Text = Math.Ceiling((Val(txtTotal.Text) - Val(txtDown.Text)) / 4)
        ElseIf rdoNumber10.Checked Then
            txtContrib.Text = Math.Ceiling((Val(txtTotal.Text) - Val(txtDown.Text)) / 10)
        ElseIf rdoNumberOther.Checked Then
            txtContrib.Text = Math.Ceiling((Val(txtTotal.Text) - Val(txtDown.Text)) / Val(txtOther.Text))
        End If
    End Sub

    Private Function transmutation(grade As Double)
        If grade > 0.6 Then
            Return Math.Floor(62.5 * grade + 37.5)
        Else
            Return Math.Floor(25 * grade + 60)
        End If
    End Function

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        txtId.Clear()
        txtName.Clear()
        txtSection.Clear()
        rdoSubject1.Checked = False
        rdoSubject2.Checked = False
        rdoSubject3.Checked = False
        txtWwRaw.Clear()
        txtWwTotal.Clear()
        txtPtRaw.Clear()
        txtPtTotal.Clear()
        txtQaRaw.Clear()
        txtQaTotal.Clear()
        txtComputed.Clear()
        txtDown.Clear()
        txtTotal.Clear()
        rdoNumber1.Checked = False
        rdoNumber2.Checked = False
        rdoNumber4.Checked = False
        rdoNumber10.Checked = False
        rdoNumberOther.Checked = False
        txtOther.Enabled = False
        txtOther.Clear()
        txtContrib.Clear()
    End Sub

    Private Sub rdoNumber_CheckedChanged(sender As Object, e As EventArgs) Handles rdoNumber1.CheckedChanged, rdoNumber2.CheckedChanged, rdoNumber4.CheckedChanged, rdoNumber10.CheckedChanged, rdoNumberOther.CheckedChanged
        If rdoNumberOther.Checked Then
            txtOther.Enabled = True
        Else
            txtOther.Enabled = False
        End If
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        Dim subject, payments As Integer

        If rdoSubject1.Checked Then
            subject = 1
        ElseIf rdoSubject2.Checked Then
            subject = 2
        ElseIf rdoSubject3.Checked Then
            subject = 3
        Else
            subject = 0
        End If

        If rdoNumber1.Checked Then
            payments = 1
        ElseIf rdoNumber2.Checked Then
            payments = 2
        ElseIf rdoNumber4.Checked Then
            payments = 4
        ElseIf rdoNumber10.Checked Then
            payments = 10
        ElseIf rdoNumberOther.Checked Then
            payments = Val(txtOther.Text)
        Else
            payments = 0
        End If

        connect.Open()
        Dim cmdadd As OleDb.OleDbCommand = New OleDb.OleDbCommand("SELECT * FROM tblStudents", connect)
        sql = "INSERT INTO tblStudents VALUES ('" & txtId.Text & "','" & txtName.Text & "','" & txtSection.Text & "'," &
            subject & "," & Val(txtComputed.Text) & "," & Val(txtDown.Text) & "," & Val(txtTotal.Text) & "," &
            payments & "," & Val(txtContrib.Text) & ")"
        cmdadd = New OleDb.OleDbCommand(sql, connect)
        cmdadd.ExecuteNonQuery()
        connect.Close()
        MsgBox("Record added.")
    End Sub

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        Dim subject, payments As Integer

        If rdoSubject1.Checked Then
            subject = 1
        ElseIf rdoSubject2.Checked Then
            subject = 2
        ElseIf rdoSubject3.Checked Then
            subject = 3
        Else
            subject = 0
        End If

        If rdoNumber1.Checked Then
            payments = 1
        ElseIf rdoNumber2.Checked Then
            payments = 2
        ElseIf rdoNumber4.Checked Then
            payments = 4
        ElseIf rdoNumber10.Checked Then
            payments = 10
        ElseIf rdoNumberOther.Checked Then
            payments = Val(txtOther.Text)
        Else
            payments = 0
        End If

        connect.Open()
        Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand("SELECT * FROM tblStudents", connect)
        sql = "UPDATE tblStudents SET student_id='" & txtId.Text & "', student_name='" & txtName.Text &
            "', student_section='" & txtSection.Text & "', student_subject=" & subject &
            ", grade_computed=" & Val(txtComputed.Text) & ", payment_down=" & Val(txtDown.Text) &
            ", payment_total=" & Val(txtTotal.Text) & ", payment_number=" & payments &
            ", payment_contrib=" & Val(txtContrib.Text) & " WHERE student_id='" & txtId.Text & "'"
        cmd = New OleDb.OleDbCommand(sql, connect)
        cmd.ExecuteNonQuery()
        connect.Close()
        MsgBox("Record updated.")
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        connect.Open()
        Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand("SELECT * FROM tblStudents WHERE student_id='" & txtId.Text & "'", connect)
        Dim sdr As OleDb.OleDbDataReader = cmd.ExecuteReader
        If (MsgBox("Are you sure you want to delete this record?", vbOKCancel) = vbOK) Then
            sql = "DELETE * from tblStudents WHERE student_id='" & txtId.Text & "'"
            cmd = New OleDb.OleDbCommand(sql, connect)
            cmd.ExecuteNonQuery()
            MsgBox("Record deleted.")
        Else
            MsgBox("Operation cancelled.")
        End If
        connect.Close()
    End Sub

    Private Sub btnFirst_Click(sender As Object, e As EventArgs) Handles btnFirst.Click
        index = 0
        navigate_records(index)
    End Sub

    Private Sub btnPrevious_Click(sender As Object, e As EventArgs) Handles btnPrevious.Click
        If index > 0 Then
            index -= 1
            navigate_records(index)
        Else
            MsgBox("First record.")
        End If
    End Sub

    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        If index < rows - 1 Then
            index += 1
            navigate_records(index)
        Else
            MsgBox("Last record.")
        End If
    End Sub

    Private Sub btnLast_Click(sender As Object, e As EventArgs) Handles btnLast.Click
        index = rows - 1
        navigate_records(index)
    End Sub

    Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
        Form2.Show()
        Me.Close()
    End Sub
End Class