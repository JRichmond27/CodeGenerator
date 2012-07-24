<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.txtConnectionString = New System.Windows.Forms.TextBox()
        Me.sbStatus = New System.Windows.Forms.StatusStrip()
        Me.lblConnStatus = New System.Windows.Forms.ToolStripStatusLabel()
        Me.lblMessage = New System.Windows.Forms.ToolStripStatusLabel()
        Me.lblConnectionString = New System.Windows.Forms.Label()
        Me.btnConnect = New System.Windows.Forms.Button()
        Me.btnDisconnect = New System.Windows.Forms.Button()
        Me.btnGenerate = New System.Windows.Forms.Button()
        Me.lblTable = New System.Windows.Forms.Label()
        Me.ddlTable = New System.Windows.Forms.ComboBox()
        Me.txtName = New System.Windows.Forms.TextBox()
        Me.lblName = New System.Windows.Forms.Label()
        Me.tabVBModule = New System.Windows.Forms.TabPage()
        Me.btnCopyVBModule = New System.Windows.Forms.Button()
        Me.txtVBModule = New System.Windows.Forms.TextBox()
        Me.tabSPDelete = New System.Windows.Forms.TabPage()
        Me.btnCopySPDelete = New System.Windows.Forms.Button()
        Me.txtSPDelete = New System.Windows.Forms.TextBox()
        Me.tabsOutput = New System.Windows.Forms.TabControl()
        Me.tabSPGetAll = New System.Windows.Forms.TabPage()
        Me.btnCopySPGetAll = New System.Windows.Forms.Button()
        Me.txtSPGetAll = New System.Windows.Forms.TextBox()
        Me.tabSPGetByID = New System.Windows.Forms.TabPage()
        Me.btnCopySPGetByID = New System.Windows.Forms.Button()
        Me.txtSPGetByID = New System.Windows.Forms.TextBox()
        Me.tabSPSave = New System.Windows.Forms.TabPage()
        Me.btnCopySPSave = New System.Windows.Forms.Button()
        Me.txtSPSave = New System.Windows.Forms.TextBox()
        Me.tabSPSort = New System.Windows.Forms.TabPage()
        Me.btnCopySPSort = New System.Windows.Forms.Button()
        Me.txtSPSort = New System.Windows.Forms.TextBox()
        Me.tabSPApprove = New System.Windows.Forms.TabPage()
        Me.btnCopySPApprove = New System.Windows.Forms.Button()
        Me.txtSPApprove = New System.Windows.Forms.TextBox()
        Me.tabSPDeactivate = New System.Windows.Forms.TabPage()
        Me.btnCopySPDeactivate = New System.Windows.Forms.Button()
        Me.txtSPDeactivate = New System.Windows.Forms.TextBox()
        Me.tabASPForm = New System.Windows.Forms.TabPage()
        Me.btnCopyASPForm = New System.Windows.Forms.Button()
        Me.txtASPForm = New System.Windows.Forms.TextBox()
        Me.tabVBForm = New System.Windows.Forms.TabPage()
        Me.btnCopyVBForm = New System.Windows.Forms.Button()
        Me.txtVBForm = New System.Windows.Forms.TextBox()
        Me.lblConnectionName = New System.Windows.Forms.Label()
        Me.txtConnectionName = New System.Windows.Forms.TextBox()
        Me.lblObjSingle = New System.Windows.Forms.Label()
        Me.txtObjSingle = New System.Windows.Forms.TextBox()
        Me.lblObjPlural = New System.Windows.Forms.Label()
        Me.txtObjPlural = New System.Windows.Forms.TextBox()
        Me.objTimer = New System.Windows.Forms.Timer(Me.components)
        Me.lblVersion = New System.Windows.Forms.Label()
        Me.sbStatus.SuspendLayout()
        Me.tabVBModule.SuspendLayout()
        Me.tabSPDelete.SuspendLayout()
        Me.tabsOutput.SuspendLayout()
        Me.tabSPGetAll.SuspendLayout()
        Me.tabSPGetByID.SuspendLayout()
        Me.tabSPSave.SuspendLayout()
        Me.tabSPSort.SuspendLayout()
        Me.tabSPApprove.SuspendLayout()
        Me.tabSPDeactivate.SuspendLayout()
        Me.tabASPForm.SuspendLayout()
        Me.tabVBForm.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtConnectionString
        '
        Me.txtConnectionString.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtConnectionString.Location = New System.Drawing.Point(8, 72)
        Me.txtConnectionString.Multiline = True
        Me.txtConnectionString.Name = "txtConnectionString"
        Me.txtConnectionString.Size = New System.Drawing.Size(504, 72)
        Me.txtConnectionString.TabIndex = 1
        '
        'sbStatus
        '
        Me.sbStatus.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.lblConnStatus, Me.lblMessage})
        Me.sbStatus.Location = New System.Drawing.Point(0, 538)
        Me.sbStatus.Name = "sbStatus"
        Me.sbStatus.Size = New System.Drawing.Size(784, 24)
        Me.sbStatus.TabIndex = 2
        Me.sbStatus.Text = "Status"
        '
        'lblConnStatus
        '
        Me.lblConnStatus.BorderSides = System.Windows.Forms.ToolStripStatusLabelBorderSides.Right
        Me.lblConnStatus.Name = "lblConnStatus"
        Me.lblConnStatus.Size = New System.Drawing.Size(83, 19)
        Me.lblConnStatus.Text = "Disconnected"
        '
        'lblMessage
        '
        Me.lblMessage.Name = "lblMessage"
        Me.lblMessage.Size = New System.Drawing.Size(53, 19)
        Me.lblMessage.Text = "Message"
        Me.lblMessage.Visible = False
        '
        'lblConnectionString
        '
        Me.lblConnectionString.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblConnectionString.AutoSize = True
        Me.lblConnectionString.Location = New System.Drawing.Point(8, 56)
        Me.lblConnectionString.Name = "lblConnectionString"
        Me.lblConnectionString.Size = New System.Drawing.Size(94, 13)
        Me.lblConnectionString.TabIndex = 3
        Me.lblConnectionString.Text = "Connection String:"
        '
        'btnConnect
        '
        Me.btnConnect.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnConnect.Enabled = False
        Me.btnConnect.Location = New System.Drawing.Point(8, 152)
        Me.btnConnect.Name = "btnConnect"
        Me.btnConnect.Size = New System.Drawing.Size(504, 32)
        Me.btnConnect.TabIndex = 4
        Me.btnConnect.Text = "Connect"
        Me.btnConnect.UseVisualStyleBackColor = True
        '
        'btnDisconnect
        '
        Me.btnDisconnect.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnDisconnect.Location = New System.Drawing.Point(8, 152)
        Me.btnDisconnect.Name = "btnDisconnect"
        Me.btnDisconnect.Size = New System.Drawing.Size(504, 32)
        Me.btnDisconnect.TabIndex = 4
        Me.btnDisconnect.Text = "Disconnect"
        Me.btnDisconnect.UseVisualStyleBackColor = True
        Me.btnDisconnect.Visible = False
        '
        'btnGenerate
        '
        Me.btnGenerate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnGenerate.Enabled = False
        Me.btnGenerate.Location = New System.Drawing.Point(528, 152)
        Me.btnGenerate.Name = "btnGenerate"
        Me.btnGenerate.Size = New System.Drawing.Size(248, 32)
        Me.btnGenerate.TabIndex = 4
        Me.btnGenerate.Text = "Generate"
        Me.btnGenerate.UseVisualStyleBackColor = True
        '
        'lblTable
        '
        Me.lblTable.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblTable.AutoSize = True
        Me.lblTable.Location = New System.Drawing.Point(528, 8)
        Me.lblTable.Name = "lblTable"
        Me.lblTable.Size = New System.Drawing.Size(86, 13)
        Me.lblTable.TabIndex = 3
        Me.lblTable.Text = "Database Table:"
        '
        'ddlTable
        '
        Me.ddlTable.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ddlTable.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ddlTable.Enabled = False
        Me.ddlTable.FormattingEnabled = True
        Me.ddlTable.Location = New System.Drawing.Point(528, 24)
        Me.ddlTable.Name = "ddlTable"
        Me.ddlTable.Size = New System.Drawing.Size(248, 21)
        Me.ddlTable.Sorted = True
        Me.ddlTable.TabIndex = 5
        '
        'txtName
        '
        Me.txtName.Location = New System.Drawing.Point(264, 24)
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(248, 20)
        Me.txtName.TabIndex = 6
        '
        'lblName
        '
        Me.lblName.AutoSize = True
        Me.lblName.Location = New System.Drawing.Point(264, 8)
        Me.lblName.Name = "lblName"
        Me.lblName.Size = New System.Drawing.Size(90, 13)
        Me.lblName.TabIndex = 7
        Me.lblName.Text = "Developer Name:"
        '
        'tabVBModule
        '
        Me.tabVBModule.Controls.Add(Me.btnCopyVBModule)
        Me.tabVBModule.Controls.Add(Me.txtVBModule)
        Me.tabVBModule.Location = New System.Drawing.Point(4, 22)
        Me.tabVBModule.Name = "tabVBModule"
        Me.tabVBModule.Padding = New System.Windows.Forms.Padding(3)
        Me.tabVBModule.Size = New System.Drawing.Size(760, 302)
        Me.tabVBModule.TabIndex = 5
        Me.tabVBModule.Text = "Module.vb"
        Me.tabVBModule.UseVisualStyleBackColor = True
        '
        'btnCopyVBModule
        '
        Me.btnCopyVBModule.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCopyVBModule.Location = New System.Drawing.Point(616, 272)
        Me.btnCopyVBModule.Name = "btnCopyVBModule"
        Me.btnCopyVBModule.Size = New System.Drawing.Size(120, 23)
        Me.btnCopyVBModule.TabIndex = 17
        Me.btnCopyVBModule.Text = "Copy Module.vb"
        Me.btnCopyVBModule.UseVisualStyleBackColor = True
        '
        'txtVBModule
        '
        Me.txtVBModule.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtVBModule.Font = New System.Drawing.Font("Consolas", 8.25!)
        Me.txtVBModule.Location = New System.Drawing.Point(0, 0)
        Me.txtVBModule.Multiline = True
        Me.txtVBModule.Name = "txtVBModule"
        Me.txtVBModule.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtVBModule.Size = New System.Drawing.Size(760, 304)
        Me.txtVBModule.TabIndex = 2
        Me.txtVBModule.WordWrap = False
        '
        'tabSPDelete
        '
        Me.tabSPDelete.Controls.Add(Me.btnCopySPDelete)
        Me.tabSPDelete.Controls.Add(Me.txtSPDelete)
        Me.tabSPDelete.Location = New System.Drawing.Point(4, 22)
        Me.tabSPDelete.Name = "tabSPDelete"
        Me.tabSPDelete.Padding = New System.Windows.Forms.Padding(3)
        Me.tabSPDelete.Size = New System.Drawing.Size(760, 302)
        Me.tabSPDelete.TabIndex = 3
        Me.tabSPDelete.Text = "spDelete"
        Me.tabSPDelete.UseVisualStyleBackColor = True
        '
        'btnCopySPDelete
        '
        Me.btnCopySPDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCopySPDelete.Location = New System.Drawing.Point(616, 272)
        Me.btnCopySPDelete.Name = "btnCopySPDelete"
        Me.btnCopySPDelete.Size = New System.Drawing.Size(120, 23)
        Me.btnCopySPDelete.TabIndex = 18
        Me.btnCopySPDelete.Text = "Copy spDelete"
        Me.btnCopySPDelete.UseVisualStyleBackColor = True
        '
        'txtSPDelete
        '
        Me.txtSPDelete.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSPDelete.Font = New System.Drawing.Font("Consolas", 8.25!)
        Me.txtSPDelete.Location = New System.Drawing.Point(0, 0)
        Me.txtSPDelete.Multiline = True
        Me.txtSPDelete.Name = "txtSPDelete"
        Me.txtSPDelete.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtSPDelete.Size = New System.Drawing.Size(760, 304)
        Me.txtSPDelete.TabIndex = 2
        Me.txtSPDelete.WordWrap = False
        '
        'tabsOutput
        '
        Me.tabsOutput.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tabsOutput.Controls.Add(Me.tabSPGetAll)
        Me.tabsOutput.Controls.Add(Me.tabSPGetByID)
        Me.tabsOutput.Controls.Add(Me.tabSPSave)
        Me.tabsOutput.Controls.Add(Me.tabSPSort)
        Me.tabsOutput.Controls.Add(Me.tabSPApprove)
        Me.tabsOutput.Controls.Add(Me.tabSPDeactivate)
        Me.tabsOutput.Controls.Add(Me.tabSPDelete)
        Me.tabsOutput.Controls.Add(Me.tabVBModule)
        Me.tabsOutput.Controls.Add(Me.tabASPForm)
        Me.tabsOutput.Controls.Add(Me.tabVBForm)
        Me.tabsOutput.Enabled = False
        Me.tabsOutput.Location = New System.Drawing.Point(8, 200)
        Me.tabsOutput.Name = "tabsOutput"
        Me.tabsOutput.SelectedIndex = 0
        Me.tabsOutput.Size = New System.Drawing.Size(768, 328)
        Me.tabsOutput.TabIndex = 0
        '
        'tabSPGetAll
        '
        Me.tabSPGetAll.Controls.Add(Me.btnCopySPGetAll)
        Me.tabSPGetAll.Controls.Add(Me.txtSPGetAll)
        Me.tabSPGetAll.Location = New System.Drawing.Point(4, 22)
        Me.tabSPGetAll.Name = "tabSPGetAll"
        Me.tabSPGetAll.Padding = New System.Windows.Forms.Padding(3)
        Me.tabSPGetAll.Size = New System.Drawing.Size(760, 302)
        Me.tabSPGetAll.TabIndex = 11
        Me.tabSPGetAll.Text = "spGetAll"
        Me.tabSPGetAll.UseVisualStyleBackColor = True
        '
        'btnCopySPGetAll
        '
        Me.btnCopySPGetAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCopySPGetAll.Location = New System.Drawing.Point(616, 272)
        Me.btnCopySPGetAll.Name = "btnCopySPGetAll"
        Me.btnCopySPGetAll.Size = New System.Drawing.Size(120, 23)
        Me.btnCopySPGetAll.TabIndex = 21
        Me.btnCopySPGetAll.Text = "Copy spGetAll"
        Me.btnCopySPGetAll.UseVisualStyleBackColor = True
        '
        'txtSPGetAll
        '
        Me.txtSPGetAll.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSPGetAll.Font = New System.Drawing.Font("Consolas", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSPGetAll.Location = New System.Drawing.Point(0, 0)
        Me.txtSPGetAll.Multiline = True
        Me.txtSPGetAll.Name = "txtSPGetAll"
        Me.txtSPGetAll.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtSPGetAll.Size = New System.Drawing.Size(760, 304)
        Me.txtSPGetAll.TabIndex = 20
        Me.txtSPGetAll.WordWrap = False
        '
        'tabSPGetByID
        '
        Me.tabSPGetByID.Controls.Add(Me.btnCopySPGetByID)
        Me.tabSPGetByID.Controls.Add(Me.txtSPGetByID)
        Me.tabSPGetByID.Location = New System.Drawing.Point(4, 22)
        Me.tabSPGetByID.Name = "tabSPGetByID"
        Me.tabSPGetByID.Padding = New System.Windows.Forms.Padding(3)
        Me.tabSPGetByID.Size = New System.Drawing.Size(760, 302)
        Me.tabSPGetByID.TabIndex = 12
        Me.tabSPGetByID.Text = "spGetByID"
        Me.tabSPGetByID.UseVisualStyleBackColor = True
        '
        'btnCopySPGetByID
        '
        Me.btnCopySPGetByID.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCopySPGetByID.Location = New System.Drawing.Point(616, 272)
        Me.btnCopySPGetByID.Name = "btnCopySPGetByID"
        Me.btnCopySPGetByID.Size = New System.Drawing.Size(120, 23)
        Me.btnCopySPGetByID.TabIndex = 21
        Me.btnCopySPGetByID.Text = "Copy spGetByID"
        Me.btnCopySPGetByID.UseVisualStyleBackColor = True
        '
        'txtSPGetByID
        '
        Me.txtSPGetByID.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSPGetByID.Font = New System.Drawing.Font("Consolas", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSPGetByID.Location = New System.Drawing.Point(0, 0)
        Me.txtSPGetByID.Multiline = True
        Me.txtSPGetByID.Name = "txtSPGetByID"
        Me.txtSPGetByID.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtSPGetByID.Size = New System.Drawing.Size(760, 304)
        Me.txtSPGetByID.TabIndex = 20
        Me.txtSPGetByID.WordWrap = False
        '
        'tabSPSave
        '
        Me.tabSPSave.Controls.Add(Me.btnCopySPSave)
        Me.tabSPSave.Controls.Add(Me.txtSPSave)
        Me.tabSPSave.Location = New System.Drawing.Point(4, 22)
        Me.tabSPSave.Name = "tabSPSave"
        Me.tabSPSave.Padding = New System.Windows.Forms.Padding(3)
        Me.tabSPSave.Size = New System.Drawing.Size(760, 302)
        Me.tabSPSave.TabIndex = 2
        Me.tabSPSave.Text = "spSave"
        Me.tabSPSave.UseVisualStyleBackColor = True
        '
        'btnCopySPSave
        '
        Me.btnCopySPSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCopySPSave.Location = New System.Drawing.Point(616, 272)
        Me.btnCopySPSave.Name = "btnCopySPSave"
        Me.btnCopySPSave.Size = New System.Drawing.Size(120, 23)
        Me.btnCopySPSave.TabIndex = 18
        Me.btnCopySPSave.Text = "Copy spSave"
        Me.btnCopySPSave.UseVisualStyleBackColor = True
        '
        'txtSPSave
        '
        Me.txtSPSave.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSPSave.Font = New System.Drawing.Font("Consolas", 8.25!)
        Me.txtSPSave.Location = New System.Drawing.Point(0, 0)
        Me.txtSPSave.Multiline = True
        Me.txtSPSave.Name = "txtSPSave"
        Me.txtSPSave.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtSPSave.Size = New System.Drawing.Size(760, 304)
        Me.txtSPSave.TabIndex = 2
        Me.txtSPSave.WordWrap = False
        '
        'tabSPSort
        '
        Me.tabSPSort.Controls.Add(Me.btnCopySPSort)
        Me.tabSPSort.Controls.Add(Me.txtSPSort)
        Me.tabSPSort.Location = New System.Drawing.Point(4, 22)
        Me.tabSPSort.Name = "tabSPSort"
        Me.tabSPSort.Padding = New System.Windows.Forms.Padding(3)
        Me.tabSPSort.Size = New System.Drawing.Size(760, 302)
        Me.tabSPSort.TabIndex = 10
        Me.tabSPSort.Text = "spUpdateSort"
        Me.tabSPSort.UseVisualStyleBackColor = True
        '
        'btnCopySPSort
        '
        Me.btnCopySPSort.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCopySPSort.Location = New System.Drawing.Point(616, 272)
        Me.btnCopySPSort.Name = "btnCopySPSort"
        Me.btnCopySPSort.Size = New System.Drawing.Size(120, 23)
        Me.btnCopySPSort.TabIndex = 19
        Me.btnCopySPSort.Text = "Copy spUpdateSort"
        Me.btnCopySPSort.UseVisualStyleBackColor = True
        '
        'txtSPSort
        '
        Me.txtSPSort.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSPSort.Font = New System.Drawing.Font("Consolas", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSPSort.Location = New System.Drawing.Point(0, 0)
        Me.txtSPSort.Multiline = True
        Me.txtSPSort.Name = "txtSPSort"
        Me.txtSPSort.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtSPSort.Size = New System.Drawing.Size(760, 304)
        Me.txtSPSort.TabIndex = 3
        Me.txtSPSort.WordWrap = False
        '
        'tabSPApprove
        '
        Me.tabSPApprove.Controls.Add(Me.btnCopySPApprove)
        Me.tabSPApprove.Controls.Add(Me.txtSPApprove)
        Me.tabSPApprove.Location = New System.Drawing.Point(4, 22)
        Me.tabSPApprove.Name = "tabSPApprove"
        Me.tabSPApprove.Padding = New System.Windows.Forms.Padding(3)
        Me.tabSPApprove.Size = New System.Drawing.Size(760, 302)
        Me.tabSPApprove.TabIndex = 14
        Me.tabSPApprove.Text = "spApprove"
        Me.tabSPApprove.UseVisualStyleBackColor = True
        '
        'btnCopySPApprove
        '
        Me.btnCopySPApprove.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCopySPApprove.Location = New System.Drawing.Point(616, 272)
        Me.btnCopySPApprove.Name = "btnCopySPApprove"
        Me.btnCopySPApprove.Size = New System.Drawing.Size(120, 23)
        Me.btnCopySPApprove.TabIndex = 23
        Me.btnCopySPApprove.Text = "Copy spApprove"
        Me.btnCopySPApprove.UseVisualStyleBackColor = True
        '
        'txtSPApprove
        '
        Me.txtSPApprove.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSPApprove.Font = New System.Drawing.Font("Consolas", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSPApprove.Location = New System.Drawing.Point(0, 0)
        Me.txtSPApprove.Multiline = True
        Me.txtSPApprove.Name = "txtSPApprove"
        Me.txtSPApprove.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtSPApprove.Size = New System.Drawing.Size(760, 304)
        Me.txtSPApprove.TabIndex = 22
        Me.txtSPApprove.WordWrap = False
        '
        'tabSPDeactivate
        '
        Me.tabSPDeactivate.Controls.Add(Me.btnCopySPDeactivate)
        Me.tabSPDeactivate.Controls.Add(Me.txtSPDeactivate)
        Me.tabSPDeactivate.Location = New System.Drawing.Point(4, 22)
        Me.tabSPDeactivate.Name = "tabSPDeactivate"
        Me.tabSPDeactivate.Padding = New System.Windows.Forms.Padding(3)
        Me.tabSPDeactivate.Size = New System.Drawing.Size(760, 302)
        Me.tabSPDeactivate.TabIndex = 13
        Me.tabSPDeactivate.Text = "spDeactivate"
        Me.tabSPDeactivate.UseVisualStyleBackColor = True
        '
        'btnCopySPDeactivate
        '
        Me.btnCopySPDeactivate.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCopySPDeactivate.Location = New System.Drawing.Point(616, 272)
        Me.btnCopySPDeactivate.Name = "btnCopySPDeactivate"
        Me.btnCopySPDeactivate.Size = New System.Drawing.Size(120, 23)
        Me.btnCopySPDeactivate.TabIndex = 21
        Me.btnCopySPDeactivate.Text = "Copy spDeactivate"
        Me.btnCopySPDeactivate.UseVisualStyleBackColor = True
        '
        'txtSPDeactivate
        '
        Me.txtSPDeactivate.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSPDeactivate.Font = New System.Drawing.Font("Consolas", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSPDeactivate.Location = New System.Drawing.Point(0, 0)
        Me.txtSPDeactivate.Multiline = True
        Me.txtSPDeactivate.Name = "txtSPDeactivate"
        Me.txtSPDeactivate.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtSPDeactivate.Size = New System.Drawing.Size(760, 304)
        Me.txtSPDeactivate.TabIndex = 20
        Me.txtSPDeactivate.WordWrap = False
        '
        'tabASPForm
        '
        Me.tabASPForm.Controls.Add(Me.btnCopyASPForm)
        Me.tabASPForm.Controls.Add(Me.txtASPForm)
        Me.tabASPForm.Location = New System.Drawing.Point(4, 22)
        Me.tabASPForm.Name = "tabASPForm"
        Me.tabASPForm.Padding = New System.Windows.Forms.Padding(3)
        Me.tabASPForm.Size = New System.Drawing.Size(760, 302)
        Me.tabASPForm.TabIndex = 7
        Me.tabASPForm.Text = "Form.aspx"
        Me.tabASPForm.UseVisualStyleBackColor = True
        '
        'btnCopyASPForm
        '
        Me.btnCopyASPForm.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCopyASPForm.Location = New System.Drawing.Point(616, 272)
        Me.btnCopyASPForm.Name = "btnCopyASPForm"
        Me.btnCopyASPForm.Size = New System.Drawing.Size(120, 23)
        Me.btnCopyASPForm.TabIndex = 18
        Me.btnCopyASPForm.Text = "Copy Form.aspx"
        Me.btnCopyASPForm.UseVisualStyleBackColor = True
        '
        'txtASPForm
        '
        Me.txtASPForm.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtASPForm.Font = New System.Drawing.Font("Consolas", 8.25!)
        Me.txtASPForm.Location = New System.Drawing.Point(0, 0)
        Me.txtASPForm.Multiline = True
        Me.txtASPForm.Name = "txtASPForm"
        Me.txtASPForm.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtASPForm.Size = New System.Drawing.Size(760, 304)
        Me.txtASPForm.TabIndex = 3
        Me.txtASPForm.WordWrap = False
        '
        'tabVBForm
        '
        Me.tabVBForm.Controls.Add(Me.btnCopyVBForm)
        Me.tabVBForm.Controls.Add(Me.txtVBForm)
        Me.tabVBForm.Location = New System.Drawing.Point(4, 22)
        Me.tabVBForm.Name = "tabVBForm"
        Me.tabVBForm.Padding = New System.Windows.Forms.Padding(3)
        Me.tabVBForm.Size = New System.Drawing.Size(760, 302)
        Me.tabVBForm.TabIndex = 8
        Me.tabVBForm.Text = "Form.aspx.vb"
        Me.tabVBForm.UseVisualStyleBackColor = True
        '
        'btnCopyVBForm
        '
        Me.btnCopyVBForm.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCopyVBForm.Location = New System.Drawing.Point(616, 272)
        Me.btnCopyVBForm.Name = "btnCopyVBForm"
        Me.btnCopyVBForm.Size = New System.Drawing.Size(120, 23)
        Me.btnCopyVBForm.TabIndex = 18
        Me.btnCopyVBForm.Text = "Copy Form.aspx.vb"
        Me.btnCopyVBForm.UseVisualStyleBackColor = True
        '
        'txtVBForm
        '
        Me.txtVBForm.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtVBForm.Font = New System.Drawing.Font("Consolas", 8.25!)
        Me.txtVBForm.Location = New System.Drawing.Point(0, 0)
        Me.txtVBForm.Multiline = True
        Me.txtVBForm.Name = "txtVBForm"
        Me.txtVBForm.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtVBForm.Size = New System.Drawing.Size(760, 304)
        Me.txtVBForm.TabIndex = 4
        Me.txtVBForm.WordWrap = False
        '
        'lblConnectionName
        '
        Me.lblConnectionName.AutoSize = True
        Me.lblConnectionName.Location = New System.Drawing.Point(8, 8)
        Me.lblConnectionName.Name = "lblConnectionName"
        Me.lblConnectionName.Size = New System.Drawing.Size(95, 13)
        Me.lblConnectionName.TabIndex = 9
        Me.lblConnectionName.Text = "Connection Name:"
        '
        'txtConnectionName
        '
        Me.txtConnectionName.Location = New System.Drawing.Point(8, 24)
        Me.txtConnectionName.Name = "txtConnectionName"
        Me.txtConnectionName.Size = New System.Drawing.Size(248, 20)
        Me.txtConnectionName.TabIndex = 8
        '
        'lblObjSingle
        '
        Me.lblObjSingle.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblObjSingle.AutoSize = True
        Me.lblObjSingle.Location = New System.Drawing.Point(528, 56)
        Me.lblObjSingle.Name = "lblObjSingle"
        Me.lblObjSingle.Size = New System.Drawing.Size(119, 13)
        Me.lblObjSingle.TabIndex = 13
        Me.lblObjSingle.Text = "Object Name (Singular):"
        '
        'txtObjSingle
        '
        Me.txtObjSingle.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtObjSingle.Enabled = False
        Me.txtObjSingle.Location = New System.Drawing.Point(528, 72)
        Me.txtObjSingle.Name = "txtObjSingle"
        Me.txtObjSingle.Size = New System.Drawing.Size(248, 20)
        Me.txtObjSingle.TabIndex = 12
        '
        'lblObjPlural
        '
        Me.lblObjPlural.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblObjPlural.AutoSize = True
        Me.lblObjPlural.Location = New System.Drawing.Point(528, 104)
        Me.lblObjPlural.Name = "lblObjPlural"
        Me.lblObjPlural.Size = New System.Drawing.Size(107, 13)
        Me.lblObjPlural.TabIndex = 11
        Me.lblObjPlural.Text = "Object Name (Plural):"
        '
        'txtObjPlural
        '
        Me.txtObjPlural.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtObjPlural.Enabled = False
        Me.txtObjPlural.Location = New System.Drawing.Point(528, 120)
        Me.txtObjPlural.Name = "txtObjPlural"
        Me.txtObjPlural.Size = New System.Drawing.Size(248, 20)
        Me.txtObjPlural.TabIndex = 10
        '
        'objTimer
        '
        Me.objTimer.Interval = 3000
        '
        'lblVersion
        '
        Me.lblVersion.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblVersion.BackColor = System.Drawing.SystemColors.MenuBar
        Me.lblVersion.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVersion.ForeColor = System.Drawing.SystemColors.ControlDarkDark
        Me.lblVersion.Location = New System.Drawing.Point(592, 544)
        Me.lblVersion.Name = "lblVersion"
        Me.lblVersion.Size = New System.Drawing.Size(168, 16)
        Me.lblVersion.TabIndex = 14
        Me.lblVersion.Text = "Version"
        Me.lblVersion.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(784, 562)
        Me.Controls.Add(Me.lblVersion)
        Me.Controls.Add(Me.ddlTable)
        Me.Controls.Add(Me.btnGenerate)
        Me.Controls.Add(Me.lblTable)
        Me.Controls.Add(Me.btnConnect)
        Me.Controls.Add(Me.lblConnectionString)
        Me.Controls.Add(Me.sbStatus)
        Me.Controls.Add(Me.tabsOutput)
        Me.Controls.Add(Me.btnDisconnect)
        Me.Controls.Add(Me.txtConnectionString)
        Me.Controls.Add(Me.lblObjPlural)
        Me.Controls.Add(Me.txtObjPlural)
        Me.Controls.Add(Me.lblName)
        Me.Controls.Add(Me.txtName)
        Me.Controls.Add(Me.lblConnectionName)
        Me.Controls.Add(Me.txtConnectionName)
        Me.Controls.Add(Me.lblObjSingle)
        Me.Controls.Add(Me.txtObjSingle)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MinimumSize = New System.Drawing.Size(800, 600)
        Me.Name = "frmMain"
        Me.Text = "Code Generator"
        Me.sbStatus.ResumeLayout(False)
        Me.sbStatus.PerformLayout()
        Me.tabVBModule.ResumeLayout(False)
        Me.tabVBModule.PerformLayout()
        Me.tabSPDelete.ResumeLayout(False)
        Me.tabSPDelete.PerformLayout()
        Me.tabsOutput.ResumeLayout(False)
        Me.tabSPGetAll.ResumeLayout(False)
        Me.tabSPGetAll.PerformLayout()
        Me.tabSPGetByID.ResumeLayout(False)
        Me.tabSPGetByID.PerformLayout()
        Me.tabSPSave.ResumeLayout(False)
        Me.tabSPSave.PerformLayout()
        Me.tabSPSort.ResumeLayout(False)
        Me.tabSPSort.PerformLayout()
        Me.tabSPApprove.ResumeLayout(False)
        Me.tabSPApprove.PerformLayout()
        Me.tabSPDeactivate.ResumeLayout(False)
        Me.tabSPDeactivate.PerformLayout()
        Me.tabASPForm.ResumeLayout(False)
        Me.tabASPForm.PerformLayout()
        Me.tabVBForm.ResumeLayout(False)
        Me.tabVBForm.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtConnectionString As System.Windows.Forms.TextBox
    Friend WithEvents sbStatus As System.Windows.Forms.StatusStrip
    Friend WithEvents lblConnectionString As System.Windows.Forms.Label
    Friend WithEvents lblConnStatus As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnConnect As System.Windows.Forms.Button
    Friend WithEvents btnDisconnect As System.Windows.Forms.Button
    Friend WithEvents btnGenerate As System.Windows.Forms.Button
    Friend WithEvents lblTable As System.Windows.Forms.Label
    Friend WithEvents ddlTable As System.Windows.Forms.ComboBox
    Friend WithEvents txtName As System.Windows.Forms.TextBox
    Friend WithEvents lblName As System.Windows.Forms.Label
    Friend WithEvents tabVBModule As System.Windows.Forms.TabPage
    Friend WithEvents txtVBModule As System.Windows.Forms.TextBox
    Friend WithEvents tabSPDelete As System.Windows.Forms.TabPage
    Friend WithEvents txtSPDelete As System.Windows.Forms.TextBox
    Friend WithEvents tabsOutput As System.Windows.Forms.TabControl
    Friend WithEvents lblConnectionName As System.Windows.Forms.Label
    Friend WithEvents txtConnectionName As System.Windows.Forms.TextBox
    Friend WithEvents lblObjSingle As System.Windows.Forms.Label
    Friend WithEvents txtObjSingle As System.Windows.Forms.TextBox
    Friend WithEvents lblObjPlural As System.Windows.Forms.Label
    Friend WithEvents txtObjPlural As System.Windows.Forms.TextBox
    Friend WithEvents btnCopyVBModule As System.Windows.Forms.Button
    Friend WithEvents objTimer As System.Windows.Forms.Timer
    Friend WithEvents lblMessage As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnCopySPDelete As System.Windows.Forms.Button
    Friend WithEvents tabASPForm As System.Windows.Forms.TabPage
    Friend WithEvents btnCopyASPForm As System.Windows.Forms.Button
    Friend WithEvents txtASPForm As System.Windows.Forms.TextBox
    Friend WithEvents tabVBForm As System.Windows.Forms.TabPage
    Friend WithEvents btnCopyVBForm As System.Windows.Forms.Button
    Friend WithEvents txtVBForm As System.Windows.Forms.TextBox
    Friend WithEvents lblVersion As System.Windows.Forms.Label
    Friend WithEvents tabSPSort As System.Windows.Forms.TabPage
    Friend WithEvents btnCopySPSort As System.Windows.Forms.Button
    Friend WithEvents txtSPSort As System.Windows.Forms.TextBox
    Friend WithEvents tabSPGetAll As System.Windows.Forms.TabPage
    Friend WithEvents tabSPGetByID As System.Windows.Forms.TabPage
    Friend WithEvents tabSPSave As System.Windows.Forms.TabPage
    Friend WithEvents btnCopySPSave As System.Windows.Forms.Button
    Friend WithEvents txtSPSave As System.Windows.Forms.TextBox
    Friend WithEvents tabSPDeactivate As System.Windows.Forms.TabPage
    Friend WithEvents btnCopySPGetAll As System.Windows.Forms.Button
    Friend WithEvents txtSPGetAll As System.Windows.Forms.TextBox
    Friend WithEvents btnCopySPGetByID As System.Windows.Forms.Button
    Friend WithEvents txtSPGetByID As System.Windows.Forms.TextBox
    Friend WithEvents btnCopySPDeactivate As System.Windows.Forms.Button
    Friend WithEvents txtSPDeactivate As System.Windows.Forms.TextBox
    Friend WithEvents tabSPApprove As System.Windows.Forms.TabPage
    Friend WithEvents btnCopySPApprove As System.Windows.Forms.Button
    Friend WithEvents txtSPApprove As System.Windows.Forms.TextBox

End Class
