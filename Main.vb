Imports System.Data
Imports System.Data.SqlClient

Public Class frmMain

    Private gobjConn As New SqlConnection
    Private gstrDBName As String = ""
    Private gblnHasSort As Boolean = False
    Private gblnHasActive As Boolean = False
    Private gblnHasApproved As Boolean = False
    Private gtabSPSort As TabPage = Nothing
    Private gtabSPDeactivate As TabPage = Nothing
    Private gtabSPApprove As TabPage = Nothing

#Region " Event Handlers "

    Private Sub frmMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not My.Settings.FirstRun Then
            My.Settings.Upgrade()
            My.Settings.FirstRun = True
        End If

        gblnHasSort = False
        gblnHasActive = False
        gblnHasApproved = False
        gtabSPSort = tabsOutput.TabPages("tabSPSort")
        tabsOutput.TabPages.RemoveByKey("tabSPSort")
        gtabSPDeactivate = tabsOutput.TabPages("tabSPDeactivate")
        tabsOutput.TabPages.RemoveByKey("tabSPDeactivate")
        gtabSPApprove = tabsOutput.TabPages("tabSPApprove")
        tabsOutput.TabPages.RemoveByKey("tabSPApprove")

        lblVersion.Text = "Version " & Application.ProductVersion

        txtConnectionString.Text = My.Settings.ConnString
        txtName.Text = My.Settings.UserName
        txtConnectionName.Text = My.Settings.ConnName
    End Sub

    Private Sub frmMain_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If gobjConn.State <> ConnectionState.Closed Then
            gobjConn.Close()
        End If

        lblConnStatus.Text = "Disconnected"

        txtConnectionString.Enabled = True
        btnConnect.Visible = True
        btnDisconnect.Visible = False
        ddlTable.Enabled = False
        btnGenerate.Enabled = False

        gobjConn.Dispose()
        gobjConn = Nothing

        My.Settings.UserName = txtName.Text.Trim
        My.Settings.ConnName = txtConnectionName.Text.Trim
        My.Settings.ConnString = txtConnectionString.Text.Trim
    End Sub

    Private Sub btnConnect_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnConnect.Click
        If txtConnectionString.Text.Trim <> "" Then
            Try
                gobjConn.ConnectionString = txtConnectionString.Text.Trim
                gobjConn.Open()

                lblConnStatus.Text = "Connected"

                txtConnectionString.Enabled = False
                btnConnect.Visible = False
                btnDisconnect.Visible = True
                ddlTable.Enabled = True

                LoadTables()

                My.Settings.ConnString = txtConnectionString.Text.Trim
            Catch ex As Exception
                MsgBox("Connection Failed" & vbCrLf & vbCrLf & ex.Message, MsgBoxStyle.OkOnly, "CodeGen: Connection Failed")
            End Try
        End If
    End Sub

    Private Sub btnDisconnect_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDisconnect.Click
        If gobjConn.State <> ConnectionState.Closed Then
            gobjConn.Close()
        End If
        lblConnStatus.Text = "Disconnected"

        txtConnectionString.Enabled = True
        btnConnect.Visible = True
        btnDisconnect.Visible = False
        ddlTable.Enabled = False
        txtObjSingle.Enabled = False
        txtObjPlural.Enabled = False
        btnGenerate.Enabled = False
        tabsOutput.Enabled = True

        txtObjPlural.Text = ""
        txtObjSingle.Text = ""
    End Sub

    Private Sub txtConnectionString_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtConnectionString.TextChanged
        If txtConnectionString.Text.Trim <> "" Then
            btnConnect.Enabled = True
        Else
            btnConnect.Enabled = False
        End If
    End Sub

    Private Sub ddlTable_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlTable.SelectedIndexChanged
        If ddlTable.SelectedItem.ToString <> "" Then
            txtObjSingle.Enabled = True
            txtObjPlural.Enabled = True
            btnGenerate.Enabled = True

            txtObjPlural.Text = ddlTable.SelectedItem.ToString.Replace("LU_", "").Replace("_", " ").Replace("-", " ")
            txtObjSingle.Text = Singularize(txtObjPlural.Text)
        Else
            txtObjSingle.Enabled = False
            txtObjPlural.Enabled = False
            btnGenerate.Enabled = False

            txtObjPlural.Text = ""
            txtObjSingle.Text = ""
        End If
    End Sub

    Private Sub btnGenerate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnGenerate.Click
        Dim cmdGet As New SqlCommand()
        Dim objAdapter As New SqlDataAdapter(cmdGet)
        Dim tblColumns As New DataTable
        Dim strTable As String = ddlTable.SelectedItem.ToString

        My.Settings.UserName = txtName.Text.Trim
        My.Settings.ConnName = txtConnectionName.Text.Trim

        cmdGet.Connection = gobjConn
        cmdGet.CommandType = CommandType.Text
        cmdGet.CommandText = "SELECT COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH " & _
            "FROM INFORMATION_SCHEMA.COLUMNS " & _
            "WHERE TABLE_NAME = '" & strTable & "'"

        objAdapter.Fill(tblColumns)

        gblnHasSort = False
        gblnHasActive = False
        gblnHasApproved = False
        If gtabSPSort Is Nothing Then
            gtabSPSort = tabsOutput.TabPages("tabSPSort")
            tabsOutput.TabPages.RemoveByKey("tabSPSort")
        End If
        If gtabSPDeactivate Is Nothing Then
            gtabSPDeactivate = tabsOutput.TabPages("tabSPDeactivate")
            tabsOutput.TabPages.RemoveByKey("tabSPDeactivate")
        End If
        If gtabSPApprove Is Nothing Then
            gtabSPApprove = tabsOutput.TabPages("tabSPApprove")
            tabsOutput.TabPages.RemoveByKey("tabSPApprove")
        End If

        GenerateSave(strTable, tblColumns)
        GenerateGetAll(strTable)
        GenerateGetByID(strTable)
        GenerateDelete(strTable)
        If gblnHasActive Then
            GenerateDeactivate(strTable)
            tabsOutput.TabPages.Insert(3, gtabSPDeactivate)
            gtabSPDeactivate = Nothing
        End If
        If gblnHasApproved Then
            GenerateApprove(strTable)
            tabsOutput.TabPages.Insert(3, tabSPApprove)
            gtabSPSort = Nothing
        End If
        If gblnHasSort Then
            GenerateUpdateSort(strTable)
            tabsOutput.TabPages.Insert(3, gtabSPSort)
            gtabSPSort = Nothing
        End If
        GenerateVBModule(tblColumns)
        GenerateASPForm(tblColumns)
        GenerateVBForm(tblColumns)

        tabsOutput.Enabled = True

        SetMessage("Successfully generated all code")
    End Sub

    Private Sub objTimer_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles objTimer.Tick
        lblMessage.Text = ""
        lblMessage.Visible = False
        objTimer.Enabled = False
    End Sub

#Region " Copy Code Click Events "

    Private Sub btnCopySPGetAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCopySPGetAll.Click
        Clipboard.SetText(txtSPGetAll.Text.Trim)
        SetMessage("spGetAll has been copied to the clipboard")
    End Sub

    Private Sub btnCopySPGetByID_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCopySPGetByID.Click
        Clipboard.SetText(txtSPGetByID.Text.Trim)
        SetMessage("spGetByID has been copied to the clipboard")
    End Sub

    Private Sub btnCopySPSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCopySPSave.Click
        Clipboard.SetText(txtSPSave.Text.Trim)
        SetMessage("spSave has been copied to the clipboard")
    End Sub

    Private Sub btnCopySPSort_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCopySPSort.Click
        Clipboard.SetText(txtSPSort.Text.Trim)
        SetMessage("spUpdateSort has been copied to the clipboard")
    End Sub

    Private Sub btnCopySPApprove_Click(sender As System.Object, e As System.EventArgs) Handles btnCopySPApprove.Click
        Clipboard.SetText(txtSPApprove.Text.Trim)
        SetMessage("spApprove has been copied to the clipboard")
    End Sub

    Private Sub btnCopySPDeactivate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCopySPDeactivate.Click
        Clipboard.SetText(txtSPDeactivate.Text.Trim)
        SetMessage("spDeactivate has been copied to the clipboard")
    End Sub

    Private Sub btnCopySPDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCopySPDelete.Click
        Clipboard.SetText(txtSPDelete.Text.Trim)
        SetMessage("spDelete has been copied to the clipboard")
    End Sub

    Private Sub btnCopyVBModule_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCopyVBModule.Click
        Clipboard.SetText(txtVBModule.Text.Trim)
        SetMessage("Module.vb has been copied to the clipboard")
    End Sub

    Private Sub btnCopyASPForm_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCopyASPForm.Click
        Clipboard.SetText(txtASPForm.Text.Trim)
        SetMessage("Form.aspx has been copied to the clipboard")
    End Sub

    Private Sub btnCopyVBForm_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCopyVBForm.Click
        Clipboard.SetText(txtVBForm.Text.Trim)
        SetMessage("Form.aspx.vb has been copied to the clipboard")
    End Sub

#End Region

    '#Region " Select All Key Press Events "

    '    Private Sub txtSPGetAll_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
    '        If e.KeyChar = Convert.ToChar(1) Then
    '            DirectCast(sender, TextBox).SelectAll()
    '            e.Handled = True
    '        End If
    '    End Sub

    '    Private Sub txtSPGetByID_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
    '        If e.KeyChar = Convert.ToChar(1) Then
    '            DirectCast(sender, TextBox).SelectAll()
    '            e.Handled = True
    '        End If
    '    End Sub

    '    Private Sub txtSPSave_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSPSave.KeyPress
    '        If e.KeyChar = Convert.ToChar(1) Then
    '            DirectCast(sender, TextBox).SelectAll()
    '            e.Handled = True
    '        End If
    '    End Sub

    '    Private Sub txtSPDelete_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSPDelete.KeyPress
    '        If e.KeyChar = Convert.ToChar(1) Then
    '            DirectCast(sender, TextBox).SelectAll()
    '            e.Handled = True
    '        End If
    '    End Sub

    '    Private Sub txtVBModule_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtVBModule.KeyPress
    '        If e.KeyChar = Convert.ToChar(1) Then
    '            DirectCast(sender, TextBox).SelectAll()
    '            e.Handled = True
    '        End If
    '    End Sub

    '#End Region

#End Region

#Region " Generate Code Methods "

    Private Sub GenerateGetAll(ByVal strTable As String)
        Dim strCode As New System.Text.StringBuilder
        Dim strObjSingle As String = txtObjSingle.Text.Trim
        Dim strObjPlural As String = txtObjPlural.Text.Trim
        Dim strObjSingleNoSpace As String = strObjSingle.Replace(" ", "")
        Dim strObjPluralNoSpace As String = strObjPlural.Replace(" ", "")

        With strCode
            .AppendLine("USE [" & gstrDBName & "]")
            .AppendLine("GO")
            .AppendLine("/****** Object: StoredProcedure [dbo].[sp" & strObjPluralNoSpace & "_GetAll] Script Date: " & Now.ToString("MM/dd/yyyy HH:mm:ss") & " ******/")
            .AppendLine("SET ANSI_NULLS ON")
            .AppendLine("GO")
            .AppendLine("SET QUOTED_IDENTIFIER ON")
            .AppendLine("GO")
            .AppendLine("-- =========================================================================")
            .AppendLine("-- Author:      " & txtName.Text.Trim)
            .AppendLine("-- Create date: " & Now.ToString("yyyy.MM.dd"))
            .AppendLine("-- Description: Returns all " & strObjPlural.ToLower & " in the system")
            .AppendLine("-- =========================================================================")
            .AppendLine("CREATE PROCEDURE [dbo].[sp" & strObjPluralNoSpace & "_GetAll]")
            If gblnHasActive Then .Append(vbTab & "@ActiveOnly AS bit")
            If gblnHasActive And gblnHasApproved Then .AppendLine(",")
            If gblnHasApproved Then .AppendLine(vbTab & "@ApprovedOnly AS bit") Else .AppendLine()
            .AppendLine("AS")
            .AppendLine("BEGIN")
            .AppendLine(vbTab & "SET NOCOUNT ON")
            .AppendLine()
            .AppendLine(vbTab & "SELECT *")
            .AppendLine(vbTab & "FROM " & strTable)
            If gblnHasActive Or gblnHasApproved Then
                .Append(vbTab & "WHERE ")
                If gblnHasActive Then .AppendLine("(Active = 1 OR @ActiveOnly = 0)")
                If gblnHasActive And gblnHasApproved Then .Append(vbTab & vbTab & "AND ")
                If gblnHasApproved Then .AppendLine("(Approved = 1 OR @ApprovedOnly = 0)")
            End If
            If gblnHasSort Then
                .AppendLine(vbTab & "ORDER BY Sort")
            End If
            .AppendLine()
            .Append("END")
        End With

        txtSPGetAll.Text = strCode.ToString
    End Sub

    Private Sub GenerateGetByID(ByVal strTable As String)
        Dim strCode As New System.Text.StringBuilder
        Dim strObjSingle As String = txtObjSingle.Text.Trim
        Dim strObjPlural As String = txtObjPlural.Text.Trim
        Dim strObjSingleNoSpace As String = strObjSingle.Replace(" ", "")
        Dim strObjPluralNoSpace As String = strObjPlural.Replace(" ", "")

        With strCode
            .AppendLine("USE [" & gstrDBName & "]")
            .AppendLine("GO")
            .AppendLine("/****** Object: StoredProcedure [dbo].[sp" & strObjPluralNoSpace & "_GetByID] Script Date: " & Now.ToString("MM/dd/yyyy HH:mm:ss") & " ******/")
            .AppendLine("SET ANSI_NULLS ON")
            .AppendLine("GO")
            .AppendLine("SET QUOTED_IDENTIFIER ON")
            .AppendLine("GO")
            .AppendLine("-- =========================================================================")
            .AppendLine("-- Author:      " & txtName.Text.Trim)
            .AppendLine("-- Create date: " & Now.ToString("yyyy.MM.dd"))
            .AppendLine("-- Description: Returns data for the requested " & strObjSingle.ToLower & " ID")
            .AppendLine("-- =========================================================================")
            .AppendLine("CREATE PROCEDURE [dbo].[sp" & strObjPluralNoSpace & "_GetByID]")
            .AppendLine(vbTab & "@" & strObjSingleNoSpace & "ID AS int")
            .AppendLine("AS")
            .AppendLine("BEGIN")
            .AppendLine(vbTab & "SET NOCOUNT ON")
            .AppendLine()
            .AppendLine(vbTab & "SELECT *")
            .AppendLine(vbTab & "FROM " & strTable)
            .AppendLine(vbTab & "WHERE ID = @" & strObjSingleNoSpace & "ID")
            .AppendLine()
            .Append("END")
        End With

        txtSPGetByID.Text = strCode.ToString
    End Sub

    Private Sub GenerateSave(ByVal strTable As String, ByRef tblColumns As DataTable)
        Dim strCode As New System.Text.StringBuilder
        Dim strObjSingle As String = txtObjSingle.Text.Trim
        Dim strObjPlural As String = txtObjPlural.Text.Trim
        Dim strObjSingleNoSpace As String = strObjSingle.Replace(" ", "")
        Dim strObjPluralNoSpace As String = strObjPlural.Replace(" ", "")
        Dim intCount As Integer = 0
        Dim intSkipped As Integer = 0
        Dim objRow As DataRow

        With strCode
            .AppendLine("USE [" & gstrDBName & "]")
            .AppendLine("GO")
            .AppendLine("/****** Object: StoredProcedure [dbo].[sp" & strObjPluralNoSpace & "_Save] Script Date: " & Now.ToString("MM/dd/yyyy HH:mm:ss") & " ******/")
            .AppendLine("SET ANSI_NULLS ON")
            .AppendLine("GO")
            .AppendLine("SET QUOTED_IDENTIFIER ON")
            .AppendLine("GO")
            .AppendLine("-- =========================================================================")
            .AppendLine("-- Author:      " & txtName.Text.Trim)
            .AppendLine("-- Create date: " & Now.ToString("yyyy.MM.dd"))
            .AppendLine("-- Description: Updates or inserts a new " & strObjSingle.ToLower & " in the system")
            .AppendLine("-- =========================================================================")
            .AppendLine("CREATE PROCEDURE [dbo].[sp" & strObjPluralNoSpace & "_Save]")

            For Each objRow In tblColumns.Rows
                If objRow("COLUMN_NAME") <> "Sort" Then
                    If intCount > 0 Then
                        .Append("," & vbCrLf)
                    End If

                    .Append(vbTab & "@" & objRow("COLUMN_NAME") & " AS " & objRow("DATA_TYPE"))
                    If objRow("CHARACTER_MAXIMUM_LENGTH") & "" <> "" Then
                        .Append("(" & objRow("CHARACTER_MAXIMUM_LENGTH").ToString.Replace("-1", "MAX") & ")")
                    End If

                    intCount += 1
                Else
                    intSkipped += 1

                    gblnHasSort = True
                End If

                If intCount + intSkipped = tblColumns.Rows.Count Then
                    .Append(vbCrLf)
                End If

                If objRow("COLUMN_NAME") = "Active" Then gblnHasActive = True

                If objRow("COLUMN_NAME") = "Approved" Then gblnHasApproved = True
            Next objRow

            .AppendLine("AS")
            .AppendLine("BEGIN")
            .AppendLine(vbTab & "SET NOCOUNT ON")
            .AppendLine()
            .AppendLine(vbTab & "IF EXISTS(SELECT * FROM " & strTable & " WHERE ID = @ID) BEGIN")
            .AppendLine(vbTab & vbTab & "--The " & strObjSingle.ToLower & " already exists, so just update the existing record")
            .AppendLine(vbTab & vbTab & "UPDATE " & strTable)

            intCount = 0
            intSkipped = 0
            For Each objRow In tblColumns.Rows
                If objRow("COLUMN_NAME") <> "ID" And objRow("COLUMN_NAME") <> "Sort" Then
                    If intCount = 0 Then
                        .Append(vbTab & vbTab & "SET ")
                    Else
                        .Append("," & vbCrLf & vbTab & vbTab & vbTab)
                    End If

                    .Append(objRow("COLUMN_NAME") & " = @" & objRow("COLUMN_NAME"))

                    intCount += 1
                Else
                    intSkipped += 1
                End If
                If intCount + intSkipped = tblColumns.Rows.Count Then
                    .Append(vbCrLf)
                End If
            Next objRow

            .AppendLine(vbTab & vbTab & "WHERE ID = @ID")
            .AppendLine()
            .AppendLine(vbTab & vbTab & "SELECT @ID AS ID")
            .AppendLine(vbTab & "END")
            .AppendLine(vbTab & "ELSE BEGIN")
            .AppendLine(vbTab & vbTab & "--The " & strObjSingle.ToLower & " does not exist, so create a new record")
            .Append(vbTab & vbTab & "INSERT INTO " & strTable & " (")

            intCount = 0
            For Each objRow In tblColumns.Rows
                If objRow("COLUMN_NAME") <> "ID" And objRow("COLUMN_NAME") <> "Sort" Then
                    If intCount > 0 Then
                        .Append(", ")
                    End If

                    .Append(objRow("COLUMN_NAME"))

                    intCount += 1
                End If
            Next objRow

            .Append(")" & vbCrLf)
            .Append(vbTab & vbTab & "VALUES (")

            intCount = 0
            For Each objRow In tblColumns.Rows
                If objRow("COLUMN_NAME") <> "ID" And objRow("COLUMN_NAME") <> "Sort" Then
                    If intCount > 0 Then
                        .Append(", ")
                    End If

                    .Append("@" & objRow("COLUMN_NAME"))

                    intCount += 1
                End If
            Next objRow

            .Append(")" & vbCrLf)
            .AppendLine()
            .AppendLine(vbTab & vbTab & "SELECT @@IDENTITY AS ID")
            .AppendLine(vbTab & "END")
            .AppendLine()
            .Append("END")
        End With

        txtSPSave.Text = strCode.ToString
    End Sub

    Private Sub GenerateUpdateSort(ByVal strTable As String)
        Dim strCode As New System.Text.StringBuilder
        Dim strObjSingle As String = txtObjSingle.Text.Trim
        Dim strObjPlural As String = txtObjPlural.Text.Trim
        Dim strObjSingleNoSpace As String = strObjSingle.Replace(" ", "")
        Dim strObjPluralNoSpace As String = strObjPlural.Replace(" ", "")

        With strCode
            .AppendLine("USE [" & gstrDBName & "]")
            .AppendLine("GO")
            .AppendLine("/****** Object: StoredProcedure [dbo].[sp" & strObjPluralNoSpace & "_GetByID] Script Date: " & Now.ToString("MM/dd/yyyy HH:mm:ss") & " ******/")
            .AppendLine("SET ANSI_NULLS ON")
            .AppendLine("GO")
            .AppendLine("SET QUOTED_IDENTIFIER ON")
            .AppendLine("GO")
            .AppendLine("-- =========================================================================")
            .AppendLine("-- Author:      " & txtName.Text.Trim)
            .AppendLine("-- Create date: " & Now.ToString("yyyy.MM.dd"))
            .AppendLine("-- Description: Updates the sort value for the given " & strObjSingle.ToLower & " ID")
            .AppendLine("-- =========================================================================")
            .AppendLine("CREATE PROCEDURE [dbo].[sp" & strObjPluralNoSpace & "_UpdateSort]")
            .AppendLine(vbTab & "@" & strObjSingleNoSpace & "ID AS int,")
            .AppendLine(vbTab & "@Sort AS int")
            .AppendLine("AS")
            .AppendLine("BEGIN")
            .AppendLine(vbTab & "SET NOCOUNT ON")
            .AppendLine()
            .AppendLine(vbTab & "UPDATE " & strTable)
            .AppendLine(vbTab & "SET Sort = @Sort")
            .AppendLine(vbTab & "WHERE ID = @" & strObjSingleNoSpace & "ID")
            .AppendLine()
            .Append("END")
        End With

        txtSPSort.Text = strCode.ToString
    End Sub

    Private Sub GenerateApprove(ByVal strTable As String)
        Dim strCode As New System.Text.StringBuilder
        Dim strObjSingle As String = txtObjSingle.Text.Trim
        Dim strObjPlural As String = txtObjPlural.Text.Trim
        Dim strObjSingleNoSpace As String = strObjSingle.Replace(" ", "")
        Dim strObjPluralNoSpace As String = strObjPlural.Replace(" ", "")

        With strCode
            .AppendLine("USE [" & gstrDBName & "]")
            .AppendLine("GO")
            .AppendLine("/****** Object: StoredProcedure [dbo].[sp" & strObjPluralNoSpace & "_Approve] Script Date: " & Now.ToString("MM/dd/yyyy HH:mm:ss") & " ******/")
            .AppendLine("SET ANSI_NULLS ON")
            .AppendLine("GO")
            .AppendLine("SET QUOTED_IDENTIFIER ON")
            .AppendLine("GO")
            .AppendLine("-- =========================================================================")
            .AppendLine("-- Author:      " & txtName.Text.Trim)
            .AppendLine("-- Create date: " & Now.ToString("yyyy.MM.dd"))
            .AppendLine("-- Description: Approves the " & strObjSingle.ToLower & " associated with the given ID")
            .AppendLine("-- =========================================================================")
            .AppendLine("CREATE PROCEDURE [dbo].[sp" & strObjPluralNoSpace & "_Approve]")
            .AppendLine(vbTab & "@" & strObjSingleNoSpace & "ID AS int")
            .AppendLine("AS")
            .AppendLine("BEGIN")
            .AppendLine(vbTab & "SET NOCOUNT ON")
            .AppendLine()
            .AppendLine(vbTab & "UPDATE " & strTable)
            .AppendLine(vbTab & "SET Approved = 1")
            .AppendLine(vbTab & "WHERE ID = @" & strObjSingleNoSpace & "ID")
            .AppendLine()
            .Append("END")
        End With

        txtSPApprove.Text = strCode.ToString
    End Sub

    Private Sub GenerateDeactivate(ByVal strTable As String)
        Dim strCode As New System.Text.StringBuilder
        Dim strObjSingle As String = txtObjSingle.Text.Trim
        Dim strObjPlural As String = txtObjPlural.Text.Trim
        Dim strObjSingleNoSpace As String = strObjSingle.Replace(" ", "")
        Dim strObjPluralNoSpace As String = strObjPlural.Replace(" ", "")

        With strCode
            .AppendLine("USE [" & gstrDBName & "]")
            .AppendLine("GO")
            .AppendLine("/****** Object: StoredProcedure [dbo].[sp" & strObjPluralNoSpace & "_Deactivate] Script Date: " & Now.ToString("MM/dd/yyyy HH:mm:ss") & " ******/")
            .AppendLine("SET ANSI_NULLS ON")
            .AppendLine("GO")
            .AppendLine("SET QUOTED_IDENTIFIER ON")
            .AppendLine("GO")
            .AppendLine("-- =========================================================================")
            .AppendLine("-- Author:      " & txtName.Text.Trim)
            .AppendLine("-- Create date: " & Now.ToString("yyyy.MM.dd"))
            .AppendLine("-- Description: Deactivates the " & strObjSingle.ToLower & " associated with the given ID")
            .AppendLine("-- =========================================================================")
            .AppendLine("CREATE PROCEDURE [dbo].[sp" & strObjPluralNoSpace & "_Deactivate]")
            .AppendLine(vbTab & "@" & strObjSingleNoSpace & "ID AS int")
            .AppendLine("AS")
            .AppendLine("BEGIN")
            .AppendLine(vbTab & "SET NOCOUNT ON")
            .AppendLine()
            .AppendLine(vbTab & "UPDATE " & strTable)
            .AppendLine(vbTab & "SET Active = 0")
            .AppendLine(vbTab & "WHERE ID = @" & strObjSingleNoSpace & "ID")
            .AppendLine()
            .Append("END")
        End With

        txtSPDeactivate.Text = strCode.ToString
    End Sub

    Private Sub GenerateDelete(ByVal strTable As String)
        Dim strCode As New System.Text.StringBuilder
        Dim strObjSingle As String = txtObjSingle.Text.Trim
        Dim strObjPlural As String = txtObjPlural.Text.Trim
        Dim strObjSingleNoSpace As String = strObjSingle.Replace(" ", "")
        Dim strObjPluralNoSpace As String = strObjPlural.Replace(" ", "")

        With strCode
            .AppendLine("USE [" & gstrDBName & "]")
            .AppendLine("GO")
            .AppendLine("/****** Object: StoredProcedure [dbo].[sp" & strObjPluralNoSpace & "_Delete] Script Date: " & Now.ToString("MM/dd/yyyy HH:mm:ss") & " ******/")
            .AppendLine("SET ANSI_NULLS ON")
            .AppendLine("GO")
            .AppendLine("SET QUOTED_IDENTIFIER ON")
            .AppendLine("GO")
            .AppendLine("-- =========================================================================")
            .AppendLine("-- Author:      " & txtName.Text.Trim)
            .AppendLine("-- Create date: " & Now.ToString("yyyy.MM.dd"))
            .AppendLine("-- Description: Deletes all data associated with the given " & strObjSingle.ToLower & " ID")
            .AppendLine("-- =========================================================================")
            .AppendLine("CREATE PROCEDURE [dbo].[sp" & strObjPluralNoSpace & "_Delete]")
            .AppendLine(vbTab & "@" & strObjSingleNoSpace & "ID AS int")
            .AppendLine("AS")
            .AppendLine("BEGIN")
            .AppendLine(vbTab & "SET NOCOUNT ON")
            .AppendLine()
            .AppendLine(vbTab & "DELETE")
            .AppendLine(vbTab & "FROM " & strTable)
            .AppendLine(vbTab & "WHERE ID = @" & strObjSingleNoSpace & "ID")
            .AppendLine()
            .Append("END")
        End With

        txtSPDelete.Text = strCode.ToString
    End Sub

    Private Sub GenerateVBModule(ByRef tblColumns As DataTable)
        Dim strCode As New System.Text.StringBuilder
        Dim strObjSingle As String = txtObjSingle.Text.Trim
        Dim strObjPlural As String = txtObjPlural.Text.Trim
        Dim strObjSingleNoSpace As String = strObjSingle.Replace(" ", "")
        Dim strObjPluralNoSpace As String = strObjPlural.Replace(" ", "")
        Dim objrow As DataRow
        Dim intCount As Integer = 0
        Dim objType As VBType

        With strCode
            .AppendLine("Imports Microsoft.VisualBasic")
            .AppendLine("Imports System.Data")
            .AppendLine("Imports System.Data.SqlClient")
            .AppendLine()
            .AppendLine("Public Module m" & strObjPluralNoSpace)
            .AppendLine()
            .AppendLine("    ''' <summary>")
            .AppendLine("    ''' Returns a datatable of all " & strObjPlural.ToLower & " in the database")
            .AppendLine("    ''' </summary>")
            If gblnHasActive Then .AppendLine("    ''' <param name=""blnActiveOnly"">Boolean value indicating if only active records should be included</param>")
            If gblnHasApproved Then .AppendLine("    ''' <param name=""blnApprovedOnly"">Boolean value indicating if only approved records should be included</param>")
            .AppendLine("    ''' <returns>A datatable of all " & strObjPlural.ToLower & " in the database</returns>")
            If gblnHasActive Or gblnHasApproved Then
                .Append("    Public Function GetAll" & strObjPluralNoSpace & "(")
                If gblnHasActive Then
                    .Append("Optional ByVal blnActiveOnly AS boolean = True")
                End If
                If gblnHasActive And gblnHasApproved Then .Append(", ")
                If gblnHasApproved Then
                    .Append("Optional ByVal blnApprovedOnly AS boolean = True")
                End If
                .AppendLine(") AS DataTable")
            Else
                .AppendLine("    Public Function GetAll" & strObjPluralNoSpace & "() AS DataTable")
            End If
            .AppendLine("        ' *****************************************************************************")
            .AppendLine("        ' Author:       " & txtName.Text.Trim)
            .AppendLine("        ' Created Date: " & Now.ToString("yyyy.MM.dd"))
            .AppendLine("        ' Description:  Returns a datatable of all " & strObjPlural.ToLower & " in the database")
            .AppendLine("        ' *****************************************************************************")
            .AppendLine("        Dim objConn As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings(""" & txtConnectionName.Text.Trim & """).ToString)")
            .AppendLine("        Dim cmdGet As New SqlCommand(""sp" & strObjPluralNoSpace & "_GetAll"", objConn)")
            If gblnHasActive Then .AppendLine("        Dim parActiveOnly As New SqlParameter(""@ActiveOnly"", blnActiveOnly)")
            If gblnHasApproved Then .AppendLine("        Dim parApprovedOnly As New SqlParameter(""@ApprovedOnly"", blnApprovedOnly)")
            .AppendLine("        Dim objAdapter As New SqlDataAdapter(cmdGet)")
            .AppendLine("        Dim objTable As New DataTable")
            .AppendLine()
            .AppendLine("        'set up the command object")
            .AppendLine("        cmdGet.CommandType = CommandType.StoredProcedure")
            If gblnHasActive Then .AppendLine("        cmdGet.Parameters.Add(parActiveOnly)")
            If gblnHasApproved Then .AppendLine("        cmdGet.Parameters.Add(parApprovedOnly)")
            .AppendLine()
            .AppendLine("        'execute the stored procedure")
            .AppendLine("        objConn.Open()")
            .AppendLine("        objAdapter.Fill(objTable)")
            .AppendLine()
            .AppendLine("        'return the DataTable")
            .AppendLine("        GetAll" & strObjPluralNoSpace & " = objTable")
            .AppendLine()
            .AppendLine("        'close the connection")
            .AppendLine("        objConn.Close()")
            .AppendLine("        objConn = Nothing")
            .AppendLine("    End Function")
            .AppendLine()

            If gblnHasActive Then
                .AppendLine("    ''' <summary>")
                .AppendLine("    ''' Deactivates the " & strObjSingle.ToLower & " associated with the given ID")
                .AppendLine("    ''' </summary>")
                .AppendLine("    ''' <param name=""int" & strObjSingleNoSpace & "ID"">ID of the " & strObjSingle.ToLower & " to deactivate</param>")
                .AppendLine("    Public Sub Deactivate" & strObjSingleNoSpace & "(ByVal int" & strObjSingleNoSpace & "ID AS integer)")
                .AppendLine("        ' *****************************************************************************")
                .AppendLine("        ' Author:       " & txtName.Text.Trim)
                .AppendLine("        ' Created Date: " & Now.ToString("yyyy.MM.dd"))
                .AppendLine("        ' Description:  Deactivates the " & strObjSingle.ToLower & " associated with the given ID")
                .AppendLine("        ' *****************************************************************************")
                .AppendLine("        If int" & strObjSingleNoSpace & "ID <= 0 Then Throw New Exception(""Cannot deactivate. A valid " & strObjSingle.ToLower & " ID was not specified."")")
                .AppendLine()
                .AppendLine("        Dim objConn As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings(""" & txtConnectionName.Text.Trim & """).ToString)")
                .AppendLine("        Dim cmdDel As New SqlCommand(""sp" & strObjPluralNoSpace & "_Deactivate"", objConn)")
                .AppendLine("        Dim par" & strObjSingleNoSpace & "ID As New SqlParameter(""@" & strObjSingleNoSpace & "ID"", int" & strObjSingleNoSpace & "ID)")
                .AppendLine()
                .AppendLine("        'set up the command object")
                .AppendLine("        cmdDel.CommandType = CommandType.StoredProcedure")
                .AppendLine("        cmdDel.Parameters.Add(par" & strObjSingleNoSpace & "ID)")
                .AppendLine()
                .AppendLine("        'execute the stored procedure")
                .AppendLine("        objConn.Open()")
                .AppendLine("        cmdDel.ExecuteNonQuery()")
                .AppendLine()
                .AppendLine("        'close the connection")
                .AppendLine("        objConn.Close()")
                .AppendLine("        objConn = Nothing")
                .AppendLine("    End Sub")
                .AppendLine()
            End If

            If gblnHasApproved Then
                .AppendLine("    ''' <summary>")
                .AppendLine("    ''' Approves the " & strObjSingle.ToLower & " associated with the given ID")
                .AppendLine("    ''' </summary>")
                .AppendLine("    ''' <param name=""int" & strObjSingleNoSpace & "ID"">ID of the " & strObjSingle.ToLower & " to deactivate</param>")
                .AppendLine("    Public Sub Approve" & strObjSingleNoSpace & "(ByVal int" & strObjSingleNoSpace & "ID AS integer)")
                .AppendLine("        ' *****************************************************************************")
                .AppendLine("        ' Author:       " & txtName.Text.Trim)
                .AppendLine("        ' Created Date: " & Now.ToString("yyyy.MM.dd"))
                .AppendLine("        ' Description:  Approves the " & strObjSingle.ToLower & " associated with the given ID")
                .AppendLine("        ' *****************************************************************************")
                .AppendLine("        If int" & strObjSingleNoSpace & "ID <= 0 Then Throw New Exception(""Cannot approve. A valid " & strObjSingle.ToLower & " ID was not specified."")")
                .AppendLine()
                .AppendLine("        Dim objConn As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings(""" & txtConnectionName.Text.Trim & """).ToString)")
                .AppendLine("        Dim cmdSave As New SqlCommand(""sp" & strObjPluralNoSpace & "_Approve"", objConn)")
                .AppendLine("        Dim par" & strObjSingleNoSpace & "ID As New SqlParameter(""@" & strObjSingleNoSpace & "ID"", int" & strObjSingleNoSpace & "ID)")
                .AppendLine()
                .AppendLine("        'set up the command object")
                .AppendLine("        cmdSave.CommandType = CommandType.StoredProcedure")
                .AppendLine("        cmdSave.Parameters.Add(par" & strObjSingleNoSpace & "ID)")
                .AppendLine()
                .AppendLine("        'execute the stored procedure")
                .AppendLine("        objConn.Open()")
                .AppendLine("        cmdSave.ExecuteNonQuery()")
                .AppendLine()
                .AppendLine("        'close the connection")
                .AppendLine("        objConn.Close()")
                .AppendLine("        objConn = Nothing")
                .AppendLine("    End Sub")
                .AppendLine()
            End If

            .AppendLine("    ''' <summary>")
            .AppendLine("    ''' Deletes the " & strObjSingle.ToLower & " associated with the given ID from the DB")
            .AppendLine("    ''' </summary>")
            .AppendLine("    ''' <param name=""int" & strObjSingleNoSpace & "ID"">ID of the " & strObjSingle.ToLower & " to delete from the database</param>")
            .AppendLine("    Public Sub Delete" & strObjSingleNoSpace & "(ByVal int" & strObjSingleNoSpace & "ID AS integer)")
            .AppendLine("        ' *****************************************************************************")
            .AppendLine("        ' Author:       " & txtName.Text.Trim)
            .AppendLine("        ' Created Date: " & Now.ToString("yyyy.MM.dd"))
            .AppendLine("        ' Description:  Deletes the " & strObjSingle.ToLower & " associated with the given ID from the DB")
            .AppendLine("        ' *****************************************************************************")
            .AppendLine("        If int" & strObjSingleNoSpace & "ID <= 0 Then Throw New Exception(""Cannot perform delete. A valid " & strObjSingle.ToLower & " ID was not specified."")")
            .AppendLine()
            .AppendLine("        Dim objConn As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings(""" & txtConnectionName.Text.Trim & """).ToString)")
            .AppendLine("        Dim cmdDel As New SqlCommand(""sp" & strObjPluralNoSpace & "_Delete"", objConn)")
            .AppendLine("        Dim par" & strObjSingleNoSpace & "ID As New SqlParameter(""@" & strObjSingleNoSpace & "ID"", int" & strObjSingleNoSpace & "ID)")
            .AppendLine()
            .AppendLine("        'set up the command object")
            .AppendLine("        cmdDel.CommandType = CommandType.StoredProcedure")
            .AppendLine("        cmdDel.Parameters.Add(par" & strObjSingleNoSpace & "ID)")
            .AppendLine()
            .AppendLine("        'execute the stored procedure")
            .AppendLine("        objConn.Open()")
            .AppendLine("        cmdDel.ExecuteNonQuery()")
            .AppendLine()
            .AppendLine("        'close the connection")
            .AppendLine("        objConn.Close()")
            .AppendLine("        objConn = Nothing")
            .AppendLine("    End Sub")
            .AppendLine()

            If gblnHasSort Then
                .AppendLine("    ''' <summary>")
                .AppendLine("    ''' Updates the sort value for the given " & strObjSingle.ToLower & " ID")
                .AppendLine("    ''' </summary>")
                .AppendLine("    ''' <param name=""int" & strObjSingleNoSpace & "ID"">ID of the " & strObjSingle.ToLower & " to update the sorting for</param>")
                .AppendLine("    ''' <param name=""intSort"">New sort value of the specified " & strObjSingle.ToLower & "</param>")
                .AppendLine("    Public Sub Update" & strObjSingleNoSpace & "Sort(ByVal int" & strObjSingleNoSpace & "ID AS integer, ByVal intSort As Integer)")
                .AppendLine("        ' *****************************************************************************")
                .AppendLine("        ' Author:       " & txtName.Text.Trim)
                .AppendLine("        ' Created Date: " & Now.ToString("yyyy.MM.dd"))
                .AppendLine("        ' Description:  Updates the sort value for the given " & strObjSingle.ToLower & " ID")
                .AppendLine("        ' *****************************************************************************")
                .AppendLine("        If int" & strObjSingleNoSpace & "ID <= 0 Then Throw New Exception(""Cannot update sort value. A valid " & strObjSingle.ToLower & " ID was not specified."")")
                .AppendLine()
                .AppendLine("        Dim objConn As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings(""" & txtConnectionName.Text.Trim & """).ToString)")
                .AppendLine("        Dim cmdUpdate As New SqlCommand(""sp" & strObjPluralNoSpace & "_UpdateSort"", objConn)")
                .AppendLine("        Dim par" & strObjSingleNoSpace & "ID As New SqlParameter(""@" & strObjSingleNoSpace & "ID"", int" & strObjSingleNoSpace & "ID)")
                .AppendLine("        Dim parSort As New SqlParameter(""@Sort"", intSort)")
                .AppendLine()
                .AppendLine("        'set up the command object")
                .AppendLine("        cmdUpdate.CommandType = CommandType.StoredProcedure")
                .AppendLine("        cmdUpdate.Parameters.Add(par" & strObjSingleNoSpace & "ID)")
                .AppendLine("        cmdUpdate.Parameters.Add(parSort)")
                .AppendLine()
                .AppendLine("        'execute the stored procedure")
                .AppendLine("        objConn.Open()")
                .AppendLine("        cmdUpdate.ExecuteNonQuery()")
                .AppendLine()
                .AppendLine("        'close the connection")
                .AppendLine("        objConn.Close()")
                .AppendLine("        objConn = Nothing")
                .AppendLine("    End Sub")
                .AppendLine()
            End If

            .AppendLine("    Public Class " & strObjSingleNoSpace)
            .AppendLine()
            .AppendLine("        Private gobjWebConfig As System.Configuration.Configuration")
            .AppendLine("        Private gobjConn As SqlConnection")
            .AppendLine()
            .AppendLine("#Region "" Properties """)

            For Each objrow In tblColumns.Rows
                objType = ConvertType(objrow("DATA_TYPE"))
                .AppendLine("        Private g" & objType.Prefix & objrow("COLUMN_NAME") & " As " & objType.Type & " = " & objType.Default)
            Next objrow

            .AppendLine()

            For Each objrow In tblColumns.Rows
                objType = ConvertType(objrow("DATA_TYPE"))
                .AppendLine("        Public Property " & objrow("COLUMN_NAME") & "() As " & objType.Type)
                .AppendLine("            Get")
                .AppendLine("                Return g" & objType.Prefix & objrow("COLUMN_NAME"))
                .AppendLine("            End Get")
                .AppendLine("            Set (ByVal value AS " & objType.Type & ")")
                .AppendLine("                g" & objType.Prefix & objrow("COLUMN_NAME") & " = value")
                .AppendLine("            End Set")
                .AppendLine("        End Property")
            Next objrow

            .AppendLine()
            .AppendLine("#End Region")
            .AppendLine()
            .AppendLine("#Region "" Constructors/Destructors """)
            .AppendLine()
            .AppendLine("        ''' <summary>")
            .AppendLine("        ''' Creates a new sql connection and clears out any existing data")
            .AppendLine("        ''' </summary>")
            .AppendLine("        Public Sub New()")
            .AppendLine("            ' *****************************************************************************")
            .AppendLine("            ' Author:       " & txtName.Text.Trim)
            .AppendLine("            ' Created Date: " & Now.ToString("yyyy.MM.dd"))
            .AppendLine("            ' Description:  Creates a new sql connection and clears out any existing data")
            .AppendLine("            ' *****************************************************************************")
            .AppendLine("            Dim objConnString As ConnectionStringSettings = System.Web.Configuration.WebConfigurationManager.ConnectionStrings(""" & txtConnectionName.Text.Trim & """)")
            .AppendLine("            gobjConn = New SqlConnection(objConnString.ConnectionString)")
            .AppendLine("            Clear()")
            .AppendLine("        End Sub")
            .AppendLine()
            .AppendLine("        ''' <summary>")
            .AppendLine("        ''' Creates a new sql connection and clears out any existing data. Pre-populates the " & strObjSingle.ToLower & " object with data for the given " & strObjSingle.ToLower & " ID")
            .AppendLine("        ''' </summary>")
            .AppendLine("        ''' <param name=""int" & strObjSingleNoSpace & "ID"">ID of the " & strObjSingle.ToLower & " to pre-populate this object with</param>")
            .AppendLine("        Public Sub New(ByVal int" & strObjSingleNoSpace & "ID As Integer)")
            .AppendLine("            ' *****************************************************************************")
            .AppendLine("            ' Author:       " & txtName.Text.Trim)
            .AppendLine("            ' Created Date: " & Now.ToString("yyyy.MM.dd"))
            .AppendLine("            ' Description:  Creates a new sql connection and clears out any existing data")
            .AppendLine("            '               Pre-populates the " & strObjSingle.ToLower & " object with data for the given " & strObjSingle.ToLower & " ID")
            .AppendLine("            ' *****************************************************************************")
            .AppendLine("            Me.New()")
            .AppendLine("            GetByID(int" & strObjSingleNoSpace & "ID)")
            .AppendLine("        End Sub")
            .AppendLine()
            .AppendLine("        ''' <summary>")
            .AppendLine("        ''' Overkills the closing and disposal of the connection")
            .AppendLine("        ''' </summary>")
            .AppendLine("        Protected Overrides Sub Finalize()")
            .AppendLine("            ' *****************************************************************************")
            .AppendLine("            ' Author:       " & txtName.Text.Trim)
            .AppendLine("            ' Created Date: " & Now.ToString("yyyy.MM.dd"))
            .AppendLine("            ' Description:  Overkills the closing and disposal of the connection")
            .AppendLine("            ' *****************************************************************************")
            .AppendLine("            gobjConn.Dispose()")
            .AppendLine("            gobjConn = Nothing")
            .AppendLine("        End Sub")
            .AppendLine()
            .AppendLine("#End Region")
            .AppendLine()
            .AppendLine("#Region "" Public Functions """)
            .AppendLine()
            .AppendLine("        ''' <summary>")
            .AppendLine("        ''' Inserts a new " & strObjSingle.ToLower & " in the database")
            .AppendLine("        ''' </summary>")
            .AppendLine("        ''' <returns>The ID of the new " & strObjSingle.ToLower & " that was created in the database</returns>")
            .AppendLine("        Public Function Insert() As Integer")
            .AppendLine("            ' *****************************************************************************")
            .AppendLine("            ' Author:       " & txtName.Text.Trim)
            .AppendLine("            ' Created Date: " & Now.ToString("yyyy.MM.dd"))
            .AppendLine("            ' Description:  Inserts a new " & strObjSingle.ToLower & " in the database")
            .AppendLine("            ' *****************************************************************************")
            .AppendLine("            Insert = Update(0)")
            .AppendLine("        End Sub")
            .AppendLine()
            .AppendLine("        ''' <summary>")
            .AppendLine("        ''' Updates the database with the data contained in this " & strObjSingle.ToLower & " object")
            .AppendLine("        ''' </summary>")
            .AppendLine("        ''' <returns>The ID of the " & strObjSingle.ToLower & " that was updated in the database</returns>")
            .AppendLine("        Public Function Update() As Integer")
            .AppendLine("            ' *****************************************************************************")
            .AppendLine("            ' Author:       " & txtName.Text.Trim)
            .AppendLine("            ' Created Date: " & Now.ToString("yyyy.MM.dd"))
            .AppendLine("            ' Description:  Updates the database with the data contained in this " & strObjSingle.ToLower & " object")
            .AppendLine("            ' *****************************************************************************")
            .AppendLine("            Return Update(ID)")
            .AppendLine("        End Sub")
            .AppendLine()
            .AppendLine("        ''' <summary>")
            .AppendLine("        ''' Updates the database with the data contained in this " & strObjSingle.ToLower & " object")
            .AppendLine("        ''' </summary>")
            .AppendLine("        ''' <param name=""int" & strObjSingleNoSpace & "ID"">ID of the " & strObjSingle.ToLower & " to update in the database</param>")
            .AppendLine("        ''' <returns>The ID of the " & strObjSingle.ToLower & " that was updated in the database</returns>")
            .AppendLine("        Public Function Update(ByVal int" & strObjSingleNoSpace & "ID As Integer) As Integer")
            .AppendLine("            ' *****************************************************************************")
            .AppendLine("            ' Author:       " & txtName.Text.Trim)
            .AppendLine("            ' Created Date: " & Now.ToString("yyyy.MM.dd"))
            .AppendLine("            ' Description:  Updates the database with the data contained in this " & strObjSingle.ToLower & " object")
            .AppendLine("            ' *****************************************************************************")
            .AppendLine("            Update = 0")
            .AppendLine()
            .AppendLine("            If int" & strObjSingleNoSpace & "ID < 0 Then Throw New Exception(""Cannot perform update. A valid ID was not specified."")")
            .AppendLine()
            .AppendLine("            Dim cmdSave As New SqlCommand(""sp" & strObjPluralNoSpace & "_Save"", gobjConn)")
            .AppendLine()
            .AppendLine("            'Create the parameters")

            For Each objrow In tblColumns.Rows
                If objrow("COLUMN_NAME") <> "Sort" Then
                    objType = ConvertType(objrow("DATA_TYPE"))
                    .Append("            Dim par" & objrow("COLUMN_NAME") & " As New SqlParameter(""@" & objrow("COLUMN_NAME") & """, SqlDbType." & objrow("DATA_TYPE"))
                    If objrow("CHARACTER_MAXIMUM_LENGTH") & "" = "" Then
                        .Append(")" & vbCrLf)
                    Else
                        If objrow("CHARACTER_MAXIMUM_LENGTH") = -1 Then
                            .Append(")" & vbCrLf)
                        Else
                            .Append(", " & objrow("CHARACTER_MAXIMUM_LENGTH") & ")" & vbCrLf)
                        End If
                    End If
                End If
            Next objrow

            .AppendLine()
            .AppendLine("            'Add the parameters")
            .AppendLine("            With cmdSave.Parameters")

            For Each objrow In tblColumns.Rows
                If objrow("COLUMN_NAME") <> "Sort" Then
                    .AppendLine("                .Add(par" & objrow("COLUMN_NAME") & ")")
                End If
            Next objrow

            .AppendLine("            End With")
            .AppendLine()
            .AppendLine("            'Add values to the parameters")

            For Each objrow In tblColumns.Rows
                If objrow("COLUMN_NAME") <> "Sort" Then
                    If objrow("COLUMN_NAME") = "ID" Then
                        .AppendLine("            parID.value = int" & strObjSingleNoSpace & "ID")
                    Else
                        objType = ConvertType(objrow("DATA_TYPE"))
                        .AppendLine("            par" & objrow("COLUMN_NAME") & ".Value = " & objrow("COLUMN_NAME"))
                    End If
                End If
            Next objrow

            .AppendLine()
            .AppendLine("            'run the save stored procedure")
            .AppendLine("            cmdSave.CommandType = CommandType.StoredProcedure")
            .AppendLine("            gobjConn.Open()")
            .AppendLine("            Update = cmdSave.ExecuteScalar()")
            .AppendLine()
            .AppendLine("            'close the connection")
            .AppendLine("            gobjConn.Close()")
            .AppendLine("        End Sub")
            .AppendLine()
            .AppendLine("        ''' <summary>")
            .AppendLine("        ''' Deletes this " & strObjSingle.ToLower & " from the database")
            .AppendLine("        ''' </summary>")
            .AppendLine("        Public Sub Delete()")
            .AppendLine("            ' *****************************************************************************")
            .AppendLine("            ' Author:       " & txtName.Text.Trim)
            .AppendLine("            ' Created Date: " & Now.ToString("yyyy.MM.dd"))
            .AppendLine("            ' Description:  Deletes this " & strObjSingle.ToLower & " from the DB")
            .AppendLine("            ' *****************************************************************************")
            .AppendLine("            Delete" & strObjSingleNoSpace & "(ID)")
            .AppendLine("            Clear()")
            .AppendLine("        End Sub")
            .AppendLine()

            If gblnHasActive Then
                .AppendLine("        ''' <summary>")
                .AppendLine("        ''' Deactivates this " & strObjSingle.ToLower)
                .AppendLine("        ''' </summary>")
                .AppendLine("        Public Sub Deactivate()")
                .AppendLine("            ' *****************************************************************************")
                .AppendLine("            ' Author:       " & txtName.Text.Trim)
                .AppendLine("            ' Created Date: " & Now.ToString("yyyy.MM.dd"))
                .AppendLine("            ' Description:  Deactivates this " & strObjSingle.ToLower)
                .AppendLine("            ' *****************************************************************************")
                .AppendLine("            Deactivate" & strObjSingleNoSpace & "(ID)")
                .AppendLine("            Clear()")
                .AppendLine("        End Sub")
                .AppendLine()
            End If

            If gblnHasApproved Then
                .AppendLine("        ''' <summary>")
                .AppendLine("        ''' Approves this " & strObjSingle.ToLower)
                .AppendLine("        ''' </summary>")
                .AppendLine("        Public Sub Approve()")
                .AppendLine("            ' *****************************************************************************")
                .AppendLine("            ' Author:       " & txtName.Text.Trim)
                .AppendLine("            ' Created Date: " & Now.ToString("yyyy.MM.dd"))
                .AppendLine("            ' Description:  Approves this " & strObjSingle.ToLower)
                .AppendLine("            ' *****************************************************************************")
                .AppendLine("            Approve" & strObjSingleNoSpace & "(ID)")
                .AppendLine("            Clear()")
                .AppendLine("        End Sub")
                .AppendLine()
            End If

            .AppendLine("        ''' <summary>")
            .AppendLine("        ''' Loads this " & strObjSingle.ToLower & " with the data associated with the given ID")
            .AppendLine("        ''' </summary>")
            .AppendLine("        ''' <param name=""int" & strObjSingleNoSpace & "ID"">ID of the " & strObjSingle.ToLower & " to get from the database</param>")
            .AppendLine("        Public Sub GetByID(ByVal int" & strObjSingleNoSpace & "ID As Integer)")
            .AppendLine("            ' *****************************************************************************")
            .AppendLine("            ' Author:       " & txtName.Text.Trim)
            .AppendLine("            ' Created Date: " & Now.ToString("yyyy.MM.dd"))
            .AppendLine("            ' Description:  Loads this " & strObjSingle.ToLower & " with the data associated with the given ID")
            .AppendLine("            ' *****************************************************************************")
            .AppendLine("            If int" & strObjSingleNoSpace & "ID <= 0 Then Throw New Exception(""Cannot load " & strObjSingle.ToLower & " data. A valid ID was not specified."")")
            .AppendLine()
            .AppendLine("            Dim cmdGet As New SqlCommand(""sp" & strObjPluralNoSpace & "_GetByID"", gobjConn)")
            .AppendLine("            Dim par" & strObjSingleNoSpace & "ID As New SqlParameter(""@" & strObjSingleNoSpace & "ID"", int" & strObjSingleNoSpace & "ID)")
            .AppendLine("            Dim objReader As SqlDataReader")
            .AppendLine()
            .AppendLine("            'set up the command object")
            .AppendLine("            cmdGet.CommandType = CommandType.StoredProcedure")
            .AppendLine("            cmdGet.Parameters.Add(par" & strObjSingleNoSpace & "ID)")
            .AppendLine()
            .AppendLine("            'execute the stored procedure")
            .AppendLine("            gobjConn.Open()")
            .AppendLine("            objReader = cmdGet.ExecuteReader")
            .AppendLine()
            .AppendLine("            'get the values from the reader")
            .AppendLine("            If objReader.Read Then")

            For Each objrow In tblColumns.Rows
                objType = ConvertType(objrow("DATA_TYPE"))
                .AppendLine("                " & objrow("COLUMN_NAME") & " = objReader(""" & objrow("COLUMN_NAME") & """) & """"")
            Next objrow

            .AppendLine("            End If")
            .AppendLine()
            .AppendLine("            'close the connection")
            .AppendLine("            objReader.Close()")
            .AppendLine("            objReader = Nothing")
            .AppendLine("            gobjConn.Close()")
            .AppendLine("        End Sub")
            .AppendLine()
            .AppendLine("        ''' <summary>")
            .AppendLine("        ''' Clears all of the properties for this object")
            .AppendLine("        ''' </summary>")
            .AppendLine("        Public Sub Clear()")
            .AppendLine("            ' *****************************************************************************")
            .AppendLine("            ' Author:       " & txtName.Text.Trim)
            .AppendLine("            ' Created Date: " & Now.ToString("yyyy.MM.dd"))
            .AppendLine("            ' Description:  Clears all of the properties for this object")
            .AppendLine("            ' *****************************************************************************")

            For Each objrow In tblColumns.Rows
                objType = ConvertType(objrow("DATA_TYPE"))
                .AppendLine("            " & objrow("COLUMN_NAME") & " = " & objType.Default)
            Next objrow

            .AppendLine("        End Sub")
            .AppendLine()
            .AppendLine("#End Region")
            .AppendLine()
            .AppendLine("    End Class")
            .AppendLine()
            .Append("End Module")
        End With

        txtVBModule.Text = strCode.ToString
    End Sub

    Private Sub GenerateASPForm(ByVal tblColumns As DataTable)
        Dim strCode As New System.Text.StringBuilder
        Dim strObjSingle As String = txtObjSingle.Text.Trim
        Dim strObjPlural As String = txtObjPlural.Text.Trim
        Dim strObjSingleNoSpace As String = strObjSingle.Replace(" ", "")
        Dim strObjPluralNoSpace As String = strObjPlural.Replace(" ", "")
        Dim objtype As VBType
        Dim objRow As DataRow

        With strCode
            .AppendLine("<fieldset>")
            .AppendLine("    <asp:HiddenField runat=""server"" id=""hdn" & strObjSingleNoSpace & "ID"" value=""0"" />")
            For Each objRow In tblColumns.Rows
                If objRow("COLUMN_NAME") <> "ID" And objRow("COLUMN_NAME") <> "Sort" Then
                    objtype = ConvertType(objRow("DATA_TYPE"))
                    .AppendLine("    <div class=""Field"">")
                    Select Case objtype.Type
                        Case "String"
                            .AppendLine("        <asp:Label runat=""server"" ID=""lbl" & objRow("COLUMN_NAME") & """ AssociatedControlID=""txt" & objRow("COLUMN_NAME") & """>" & objRow("COLUMN_NAME") & ":</asp:Label>")
                            .AppendLine("        <asp:TextBox runat=""server"" ID=""txt" & objRow("COLUMN_NAME") & """ CssClass=""L3"" />")
                        Case "Integer", "Double"
                            .AppendLine("        <asp:Label runat=""server"" ID=""lbl" & objRow("COLUMN_NAME") & """ AssociatedControlID=""txt" & objRow("COLUMN_NAME") & """>" & objRow("COLUMN_NAME") & ":</asp:Label>")
                            .AppendLine("        <asp:TextBox runat=""server"" ID=""txt" & objRow("COLUMN_NAME") & """ CssClass=""L6 Numeric"" />")
                        Case "Boolean"
                            .AppendLine("        <asp:CheckBox runat=""server"" ID=""chk" & objRow("COLUMN_NAME") & """ Text=""" & objRow("COLUMN_NAME") & """ TextAlign=""Right"" CssClass=""CheckBox"" />")
                        Case "Date"
                            .AppendLine("        <asp:Label runat=""server"" ID=""lbl" & objRow("COLUMN_NAME") & """ AssociatedControlID=""txt" & objRow("COLUMN_NAME") & """>" & objRow("COLUMN_NAME") & ":</asp:Label>")
                            .AppendLine("        <asp:TextBox runat=""server"" ID=""txt" & objRow("COLUMN_NAME") & """ CssClass=""DateCal"" />")
                        Case Else
                    End Select
                    .AppendLine("    </div>")
                End If
            Next objRow

            .AppendLine("    <div class=""Buttons"">")
            .AppendLine("        <asp:Button runat=""server"" id=""btnSave"" text=""Save"" CssClass=""Save"" />")
            .AppendLine("        <asp:Button runat=""server"" id=""btnCancel"" text=""Cancel"" CssClass=""Cancel"" />")
            .AppendLine("    </div>")
            .Append("</fieldset>")
        End With

        txtASPForm.Text = strCode.ToString
    End Sub

    Private Sub GenerateVBForm(ByVal tblColumns As DataTable)
        Dim strCode As New System.Text.StringBuilder
        Dim strObjSingle As String = txtObjSingle.Text.Trim
        Dim strObjPlural As String = txtObjPlural.Text.Trim
        Dim strObjSingleNoSpace As String = strObjSingle.Replace(" ", "")
        Dim strObjPluralNoSpace As String = strObjPlural.Replace(" ", "")
        Dim objtype As VBType
        Dim objRow As DataRow

        With strCode
            .AppendLine("#Region "" Event Handlers """)
            .AppendLine()
            .AppendLine("    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load")
            .AppendLine("        ' *****************************************************************************")
            .AppendLine("        ' Author:       " & txtName.Text.Trim)
            .AppendLine("        ' Created Date: " & Now.ToString("yyyy.MM.dd"))
            .AppendLine("        ' Description:  Handles page load; Sets up the form for adding or editing")
            .AppendLine("        ' *****************************************************************************")
            .AppendLine("        Try")
            .AppendLine("            If Not IsPostBack Then")
            .AppendLine("                'TODO: Load any DDLs on the page")
            .AppendLine()
            .AppendLine("                If Request(""ID"") <> """" Then")
            .AppendLine("                    'Set up form for edit")
            .AppendLine("                    hdn" & strObjSingleNoSpace & "ID.Value = Request(""ID"")")
            .AppendLine("                    btnSave.CommandName = ""Edit""")
            .AppendLine("                    Load" & strObjSingleNoSpace & "(hdn" & strObjSingleNoSpace & "ID.Value)")
            .AppendLine("                    btnSave.Text = ""Update " & strObjSingle & """")
            .AppendLine("                Else")
            .AppendLine("                    'Set up form for add")
            .AppendLine("                    btnSave.CommandName = ""Add""")
            .AppendLine("                    btnSave.Text = ""Add " & strObjSingle & """")
            .AppendLine("                End If")
            .AppendLine("            End If")
            .AppendLine("        Catch ex As Exception")
            .AppendLine("            Master.HandleError(Ex, AppRelativeVirtualPath, System.Reflection.MethodBase.GetCurrentMethod.Name, User.Identity.Name, , Not IsPostBack)")
            .AppendLine("        End Try")
            .AppendLine("    End Sub")
            .AppendLine()
            .AppendLine("    Protected Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click")
            .AppendLine("        ' *****************************************************************************")
            .AppendLine("        ' Author:       " & txtName.Text.Trim)
            .AppendLine("        ' Created Date: " & Now.ToString("yyyy.MM.dd"))
            .AppendLine("        ' Description:  Handles cancel click; Redirects to " & strObjPlural.ToLower & " index page")
            .AppendLine("        ' *****************************************************************************")
            .AppendLine("        Response.Redirect(""./?Msg=Cancel"")")
            .AppendLine("    End Sub")
            .AppendLine()
            .AppendLine("    Protected Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click")
            .AppendLine("        ' *****************************************************************************")
            .AppendLine("        ' Author:       " & txtName.Text.Trim)
            .AppendLine("        ' Created Date: " & Now.ToString("yyyy.MM.dd"))
            .AppendLine("        ' Description:  Handles save click; Validates and saves the " & strObjSingle.ToLower)
            .AppendLine("        ' *****************************************************************************")
            .AppendLine("        Dim strRedirect As String = """"")
            .AppendLine()
            .AppendLine("        Try")
            .AppendLine("            If ValidateForm() Then")
            .AppendLine("                Dim obj" & strObjSingleNoSpace & " As New " & strObjSingleNoSpace & "")
            .AppendLine()
            .AppendLine("                'fill the " & strObjSingle.ToLower & " object")
            .AppendLine("                With obj" & strObjSingleNoSpace)
            .AppendLine("                    .ID = hdn" & strObjSingleNoSpace & "ID.Value")

            For Each objRow In tblColumns.Rows
                If objRow("COLUMN_NAME") <> "ID" And objRow("COLUMN_NAME") <> "Sort" Then
                    objtype = ConvertType(objRow("DATA_TYPE"))
                    Select Case objtype.Type
                        Case "String"
                            .AppendLine("                    ." & objRow("COLUMN_NAME") & " = txt" & objRow("COLUMN_NAME") & ".Text.Trim")
                        Case "Integer"
                            .AppendLine("                    ." & objRow("COLUMN_NAME") & " = NumbersOnly(txt" & objRow("COLUMN_NAME") & ".Text.Trim, , , True)")
                        Case "Double"
                            .AppendLine("                    ." & objRow("COLUMN_NAME") & " = NumbersOnly(txt" & objRow("COLUMN_NAME") & ".Text.Trim, , True, True)")
                        Case "Boolean"
                            .AppendLine("                    ." & objRow("COLUMN_NAME") & " = chk" & objRow("COLUMN_NAME") & ".Checked")
                        Case "Date"
                            .AppendLine("                    If txt" & objRow("COLUMN_NAME") & ".Text.Trim = """" Or Not IsDate(txt" & objRow("COLUMN_NAME") & ".Text.Trim) Then")
                            .AppendLine("                        ." & objRow("COLUMN_NAME") & " = Nothing")
                            .AppendLine("                    Else")
                            .AppendLine("                        ." & objRow("COLUMN_NAME") & " = txt" & objRow("COLUMN_NAME") & ".Text.Trim")
                            .AppendLine("                    End If")
                        Case Else
                    End Select
                End If
            Next objRow

            .AppendLine("                End With")
            .AppendLine()
            .AppendLine("                If btnSave.CommandName = ""Edit"" Then")
            .AppendLine("                    'update the existing " & strObjSingle.ToLower)
            .AppendLine("                    obj" & strObjSingleNoSpace & ".Update()")
            .AppendLine()
            .AppendLine("                    'set the redirect URL")
            .AppendLine("                    strRedirect = ""./?Msg=Updated""")
            .AppendLine("                Else")
            .AppendLine("                    'insert a new " & strObjSingle.ToLower)
            .AppendLine("                    obj" & strObjSingleNoSpace & ".Insert()")
            .AppendLine()
            .AppendLine("                    'set the redirect URL")
            .AppendLine("                    strRedirect = ""./?Msg=Saved""")
            .AppendLine("                End If")
            .AppendLine()
            .AppendLine("                'clean up")
            .AppendLine("                obj" & strObjSingleNoSpace & " = Nothing")
            .AppendLine("            End If")
            .AppendLine("        Catch ex As Exception")
            .AppendLine("            Master.HandleError(Ex, AppRelativeVirtualPath, System.Reflection.MethodBase.GetCurrentMethod.Name, User.Identity.Name, , Not IsPostBack)")
            .AppendLine("        End Try")
            .AppendLine()
            .AppendLine("        'follow preset redirect")
            .AppendLine("        If strRedirect <> """" Then")
            .AppendLine("            Response.Redirect(strRedirect)")
            .AppendLine("        End If")
            .AppendLine("    End Sub")
            .AppendLine()
            .AppendLine("#End Region")
            .AppendLine()
            .AppendLine("#Region "" Miscellaneous Methods """)
            .AppendLine()
            .AppendLine("    Private Sub Load" & strObjSingleNoSpace & "(ByVal int" & strObjSingleNoSpace & "ID As Integer)")
            .AppendLine("        ' *****************************************************************************")
            .AppendLine("        ' Author:       " & txtName.Text.Trim)
            .AppendLine("        ' Created Date: " & Now.ToString("yyyy.MM.dd"))
            .AppendLine("        ' Description:  Loads the form with the " & strObjSingle.ToLower & " to be edited")
            .AppendLine("        ' *****************************************************************************")
            .AppendLine("        Dim obj" & strObjSingleNoSpace & " As New " & strObjSingleNoSpace & "(int" & strObjSingleNoSpace & "ID)")
            .AppendLine()
            .AppendLine("        With obj" & strObjSingleNoSpace)

            For Each objRow In tblColumns.Rows
                objtype = ConvertType(objRow("DATA_TYPE"))
                If objRow("COLUMN_NAME") <> "ID" And objRow("COLUMN_NAME") <> "Sort" Then
                    Select Case objtype.Type
                        Case "String"
                            .AppendLine("            txt" & objRow("COLUMN_NAME") & ".Text = ." & objRow("COLUMN_NAME"))
                        Case "Integer", "Double"
                            .AppendLine("            txt" & objRow("COLUMN_NAME") & ".Text = FormatNumber(." & objRow("COLUMN_NAME") & ")")
                        Case "Boolean"
                            .AppendLine("            chk" & objRow("COLUMN_NAME") & ".Checked = ." & objRow("COLUMN_NAME"))
                        Case "Date"
                            .AppendLine("            txt" & objRow("COLUMN_NAME") & ".Text = ." & objRow("COLUMN_NAME") & ".ToString(""M/d/yyyy"")")
                        Case Else
                    End Select
                End If
            Next objRow

            .AppendLine("        End With")
            .AppendLine()
            .AppendLine("        'clean up")
            .AppendLine("        obj" & strObjSingleNoSpace & "  = Nothing")
            .AppendLine("    End Sub")
            .AppendLine()
            .AppendLine("    Private Function ValidateForm() As Boolean")
            .AppendLine("        ' *****************************************************************************")
            .AppendLine("        ' Author:       " & txtName.Text.Trim)
            .AppendLine("        ' Created Date: " & Now.ToString("yyyy.MM.dd"))
            .AppendLine("        ' Description:  Validates the user entered form data")
            .AppendLine("        ' *****************************************************************************")
            .AppendLine("        ValidateForm = True")
            .AppendLine("        ResetHighlights()")
            .AppendLine()
            .AppendLine("        Dim strMessage As String = """"")
            .AppendLine()
            .AppendLine("        'start validation results message")
            .AppendLine("        strMessage = ""<p>The following fields are either empty or contain invalid data. Please correct them and try again.</p><ul>""")
            .AppendLine()

            For Each objRow In tblColumns.Rows
                objtype = ConvertType(objRow("DATA_TYPE"))
                If objRow("COLUMN_NAME") <> "ID" And objRow("COLUMN_NAME") <> "Sort" Then
                    Select Case objtype.Type
                        Case "String", "Integer", "Double"
                            .AppendLine("        If txt" & objRow("COLUMN_NAME") & ".Text.Trim = """" Then")
                            .AppendLine("            ValidateForm = False")
                            .AppendLine("            strMessage &= ""<li>" & objRow("COLUMN_NAME") & " is required</li>""")
                            .AppendLine("            txt" & objRow("COLUMN_NAME") & ".CssClass &= "" Invalid""")
                            .AppendLine("        End If")
                        Case "Date"
                            .AppendLine("        If txt" & objRow("COLUMN_NAME") & ".Text.Trim = """" Or Not IsDate(txt" & objRow("COLUMN_NAME") & ".Text.Trim) Then")
                            .AppendLine("            ValidateForm = False")
                            .AppendLine("            strMessage &= ""<li>" & objRow("COLUMN_NAME") & " is required</li>""")
                            .AppendLine("            txt" & objRow("COLUMN_NAME") & ".CssClass &= "" Invalid""")
                            .AppendLine("        End If")
                        Case Else
                    End Select
                End If
            Next objRow


            .AppendLine("        'close the validation message list that was started above")
            .AppendLine("        strMessage &= ""</ul>""")
            .AppendLine()
            .AppendLine("        'display the validation message if a problem was found")
            .AppendLine("        If Not ValidateForm Then")
            .AppendLine("            Master.DisplayNote(MsgType.Warn, strMessage)")
            .AppendLine("        End If")
            .AppendLine("    End Function")
            .AppendLine()
            .AppendLine("    Private Sub ResetHighlights()")
            .AppendLine("        ' *****************************************************************************")
            .AppendLine("        ' Author:       " & txtName.Text.Trim)
            .AppendLine("        ' Created Date: " & Now.ToString("yyyy.MM.dd"))
            .AppendLine("        ' Description:  Reset Invalid class to remove field highlighting")
            .AppendLine("        ' *****************************************************************************")

            For Each objRow In tblColumns.Rows
                objtype = ConvertType(objRow("DATA_TYPE"))
                If objRow("COLUMN_NAME") <> "ID" And objRow("COLUMN_NAME") <> "Sort" And objtype.Type <> "Boolean" Then
                    .AppendLine("        txt" & objRow("COLUMN_NAME") & ".CssClass = txt" & objRow("COLUMN_NAME") & ".CssClass.Replace("" Invalid"", """")")
                End If
            Next objRow

            .AppendLine("    End Sub")
            .AppendLine()
            .AppendLine("#End Region")
        End With

        txtVBForm.Text = strCode.ToString
    End Sub

#End Region

#Region " Load/Save Methods "

    Private Sub LoadTables()
        Dim cmdGet As New SqlCommand()
        Dim objReader As SqlDataReader

        cmdGet.Connection = gobjConn
        cmdGet.CommandType = CommandType.Text
        cmdGet.CommandText = "SELECT TABLE_NAME, TABLE_CATALOG " & _
            "FROM INFORMATION_SCHEMA.TABLES " & _
            "WHERE TABLE_TYPE <> 'VIEW' " & _
            "ORDER BY TABLE_NAME"

        objReader = cmdGet.ExecuteReader()

        ddlTable.Items.Clear()

        While objReader.Read()
            gstrDBName = objReader("TABLE_CATALOG") & ""
            ddlTable.Items.Add(objReader("TABLE_NAME") & "")
        End While

        ddlTable.SelectedIndex = 0

        objReader.Close()
        objReader = Nothing
    End Sub

#End Region

#Region " Support Methods "

    Private Structure VBType
        Dim Prefix As String
        Dim Type As String
        Dim [Default] As String
    End Structure
    Private Function ConvertType(ByVal strType As String) As VBType
        Dim objType As VBType

        With objType
            Select Case strType
                Case "int"
                    .Prefix = "int"
                    .Type = "Integer"
                    .Default = "0"
                Case "varchar"
                    .Prefix = "str"
                    .Type = "String"
                    .Default = """"""
                Case "char"
                    .Prefix = "str"
                    .Type = "String"
                    .Default = """"""
                Case "bit"
                    .Prefix = "bln"
                    .Type = "Boolean"
                    .Default = "False"
                Case "money"
                    .Prefix = "dbl"
                    .Type = "Double"
                    .Default = "0"
                Case "float"
                    .Prefix = "dbl"
                    .Type = "Double"
                    .Default = "0"
                Case "datetime"
                    .Prefix = "dat"
                    .Type = "Date"
                    .Default = "Nothing"
                Case Else
                    .Prefix = "obj"
                    .Type = "Object"
                    .Default = "Nothing"
            End Select
        End With

        Return objType
    End Function

    Private Function Singularize(ByVal strPlural As String) As String
        If strPlural.ToLower.EndsWith("ies") AndAlso "aeiou".IndexOf(strPlural.ToLower.Substring(strPlural.Length - 4, 1)) < 0 Then
            Singularize = strPlural.Substring(0, strPlural.Length - 3) & "y"
        ElseIf strPlural.ToLower.EndsWith("es") AndAlso "sszz".IndexOf(strPlural.ToLower.Substring(strPlural.Length - 4, 2)) >= 0 Then
            Singularize = strPlural.Substring(0, strPlural.Length - 3)
        ElseIf strPlural.ToLower.EndsWith("es") AndAlso "szx".IndexOf(strPlural.ToLower.Substring(strPlural.Length - 3, 1)) >= 0 Then
            Singularize = strPlural.Substring(0, strPlural.Length - 2)
        ElseIf strPlural.ToLower.EndsWith("es") AndAlso "shch".IndexOf(strPlural.ToLower.Substring(strPlural.Length - 4, 2)) >= 0 Then
            Singularize = strPlural.Substring(0, strPlural.Length - 2)
        ElseIf strPlural.ToLower.Substring(strPlural.Length - 1).ToLower = "s" Then
            Singularize = strPlural.Substring(0, strPlural.Length - 1)
        ElseIf strPlural.ToLower.EndsWith("men") Then
            Singularize = strPlural.Substring(0, strPlural.Length - 2) & "an"
        ElseIf strPlural.ToLower = "people" Then
            Singularize = "Person"
        Else
            Singularize = strPlural
        End If
    End Function

    Private Sub SetMessage(ByVal strMessage As String)
        lblMessage.Text = strMessage
        lblMessage.Visible = True
        objTimer.Enabled = True
    End Sub

#End Region

End Class