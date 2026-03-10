<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class EventLogTester
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        components = New ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EventLogTester))
        TableLayoutPanel1 = New TableLayoutPanel()
        lblLog = New Label()
        cboLogName = New ComboBox()
        lblSource = New Label()
        cboSourceName = New ComboBox()
        lblMessage = New Label()
        txtMessage = New TextBox()
        ckbSeverity = New CheckBox()
        cboSeverity = New ComboBox()
        ckbEventID = New CheckBox()
        numEventID = New NumericUpDown()
        ckbCategory = New CheckBox()
        numCategory = New NumericUpDown()
        ckbRawData = New CheckBox()
        txtRawData = New TextBox()
        btnSend = New Button()
        cboEventType = New ComboBox()
        ckbEventType = New CheckBox()
        lblStatus = New Label()
        tmr = New Timer(components)
        TableLayoutPanel1.SuspendLayout()
        CType(numEventID, ComponentModel.ISupportInitialize).BeginInit()
        CType(numCategory, ComponentModel.ISupportInitialize).BeginInit()
        SuspendLayout()
        ' 
        ' TableLayoutPanel1
        ' 
        TableLayoutPanel1.ColumnCount = 6
        TableLayoutPanel1.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0F))
        TableLayoutPanel1.ColumnStyles.Add(New ColumnStyle())
        TableLayoutPanel1.ColumnStyles.Add(New ColumnStyle())
        TableLayoutPanel1.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 240.0F))
        TableLayoutPanel1.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 240.0F))
        TableLayoutPanel1.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0F))
        TableLayoutPanel1.Controls.Add(lblLog, 1, 1)
        TableLayoutPanel1.Controls.Add(cboLogName, 2, 1)
        TableLayoutPanel1.Controls.Add(lblSource, 1, 2)
        TableLayoutPanel1.Controls.Add(cboSourceName, 2, 2)
        TableLayoutPanel1.Controls.Add(lblMessage, 1, 7)
        TableLayoutPanel1.Controls.Add(txtMessage, 2, 7)
        TableLayoutPanel1.Controls.Add(ckbSeverity, 1, 4)
        TableLayoutPanel1.Controls.Add(cboSeverity, 2, 4)
        TableLayoutPanel1.Controls.Add(ckbEventID, 1, 5)
        TableLayoutPanel1.Controls.Add(numEventID, 2, 5)
        TableLayoutPanel1.Controls.Add(ckbCategory, 1, 6)
        TableLayoutPanel1.Controls.Add(numCategory, 2, 6)
        TableLayoutPanel1.Controls.Add(ckbRawData, 1, 8)
        TableLayoutPanel1.Controls.Add(txtRawData, 2, 8)
        TableLayoutPanel1.Controls.Add(btnSend, 4, 10)
        TableLayoutPanel1.Controls.Add(cboEventType, 2, 3)
        TableLayoutPanel1.Controls.Add(ckbEventType, 1, 3)
        TableLayoutPanel1.Controls.Add(lblStatus, 1, 10)
        TableLayoutPanel1.Dock = DockStyle.Fill
        TableLayoutPanel1.Location = New Point(0, 0)
        TableLayoutPanel1.Name = "TableLayoutPanel1"
        TableLayoutPanel1.RowCount = 12
        TableLayoutPanel1.RowStyles.Add(New RowStyle(SizeType.Percent, 50.0F))
        TableLayoutPanel1.RowStyles.Add(New RowStyle())
        TableLayoutPanel1.RowStyles.Add(New RowStyle())
        TableLayoutPanel1.RowStyles.Add(New RowStyle())
        TableLayoutPanel1.RowStyles.Add(New RowStyle())
        TableLayoutPanel1.RowStyles.Add(New RowStyle())
        TableLayoutPanel1.RowStyles.Add(New RowStyle())
        TableLayoutPanel1.RowStyles.Add(New RowStyle(SizeType.Absolute, 75.0F))
        TableLayoutPanel1.RowStyles.Add(New RowStyle(SizeType.Absolute, 75.0F))
        TableLayoutPanel1.RowStyles.Add(New RowStyle(SizeType.Percent, 50.0F))
        TableLayoutPanel1.RowStyles.Add(New RowStyle(SizeType.Absolute, 40.0F))
        TableLayoutPanel1.RowStyles.Add(New RowStyle(SizeType.Absolute, 10.0F))
        TableLayoutPanel1.Size = New Size(800, 450)
        TableLayoutPanel1.TabIndex = 0
        ' 
        ' lblLog
        ' 
        lblLog.AutoSize = True
        lblLog.Dock = DockStyle.Right
        lblLog.Location = New Point(71, 38)
        lblLog.Name = "lblLog"
        lblLog.Size = New Size(62, 29)
        lblLog.TabIndex = 0
        lblLog.Text = "Log Name"
        lblLog.TextAlign = ContentAlignment.MiddleRight
        ' 
        ' cboLogName
        ' 
        TableLayoutPanel1.SetColumnSpan(cboLogName, 2)
        cboLogName.Dock = DockStyle.Fill
        cboLogName.FormattingEnabled = True
        cboLogName.Location = New Point(139, 41)
        cboLogName.Name = "cboLogName"
        cboLogName.Size = New Size(370, 23)
        cboLogName.TabIndex = 1
        ' 
        ' lblSource
        ' 
        lblSource.AutoSize = True
        lblSource.Dock = DockStyle.Right
        lblSource.Location = New Point(55, 67)
        lblSource.Name = "lblSource"
        lblSource.Size = New Size(78, 29)
        lblSource.TabIndex = 2
        lblSource.Text = "Source Name"
        lblSource.TextAlign = ContentAlignment.MiddleRight
        ' 
        ' cboSourceName
        ' 
        TableLayoutPanel1.SetColumnSpan(cboSourceName, 2)
        cboSourceName.Dock = DockStyle.Fill
        cboSourceName.FormattingEnabled = True
        cboSourceName.Location = New Point(139, 70)
        cboSourceName.Name = "cboSourceName"
        cboSourceName.Size = New Size(370, 23)
        cboSourceName.TabIndex = 3
        ' 
        ' lblMessage
        ' 
        lblMessage.AutoSize = True
        lblMessage.Dock = DockStyle.Top
        lblMessage.Location = New Point(50, 212)
        lblMessage.Name = "lblMessage"
        lblMessage.Size = New Size(83, 15)
        lblMessage.TabIndex = 10
        lblMessage.Text = "Message"
        lblMessage.TextAlign = ContentAlignment.TopRight
        ' 
        ' txtMessage
        ' 
        TableLayoutPanel1.SetColumnSpan(txtMessage, 3)
        txtMessage.Dock = DockStyle.Fill
        txtMessage.Location = New Point(139, 215)
        txtMessage.Multiline = True
        txtMessage.Name = "txtMessage"
        txtMessage.Size = New Size(610, 69)
        txtMessage.TabIndex = 11
        ' 
        ' ckbSeverity
        ' 
        ckbSeverity.AutoSize = True
        ckbSeverity.CheckAlign = ContentAlignment.MiddleRight
        ckbSeverity.Dock = DockStyle.Right
        ckbSeverity.Location = New Point(66, 128)
        ckbSeverity.Name = "ckbSeverity"
        ckbSeverity.Size = New Size(67, 23)
        ckbSeverity.TabIndex = 4
        ckbSeverity.Text = "Severity"
        ckbSeverity.TextAlign = ContentAlignment.MiddleRight
        ' 
        ' cboSeverity
        ' 
        cboSeverity.Dock = DockStyle.Fill
        cboSeverity.DropDownStyle = ComboBoxStyle.DropDownList
        cboSeverity.Location = New Point(139, 128)
        cboSeverity.Name = "cboSeverity"
        cboSeverity.Size = New Size(130, 23)
        cboSeverity.TabIndex = 5
        ' 
        ' ckbEventID
        ' 
        ckbEventID.AutoSize = True
        ckbEventID.CheckAlign = ContentAlignment.MiddleRight
        ckbEventID.Dock = DockStyle.Right
        ckbEventID.Location = New Point(64, 157)
        ckbEventID.Name = "ckbEventID"
        ckbEventID.Size = New Size(69, 23)
        ckbEventID.TabIndex = 6
        ckbEventID.Text = "Event ID"
        ckbEventID.TextAlign = ContentAlignment.MiddleRight
        ' 
        ' numEventID
        ' 
        numEventID.Dock = DockStyle.Fill
        numEventID.Location = New Point(139, 157)
        numEventID.Name = "numEventID"
        numEventID.Size = New Size(130, 23)
        numEventID.TabIndex = 7
        ' 
        ' ckbCategory
        ' 
        ckbCategory.AutoSize = True
        ckbCategory.CheckAlign = ContentAlignment.MiddleRight
        ckbCategory.Dock = DockStyle.Right
        ckbCategory.Location = New Point(59, 186)
        ckbCategory.Name = "ckbCategory"
        ckbCategory.Size = New Size(74, 23)
        ckbCategory.TabIndex = 8
        ckbCategory.Text = "Category"
        ckbCategory.TextAlign = ContentAlignment.MiddleRight
        ' 
        ' numCategory
        ' 
        numCategory.Dock = DockStyle.Fill
        numCategory.Location = New Point(139, 186)
        numCategory.Name = "numCategory"
        numCategory.Size = New Size(130, 23)
        numCategory.TabIndex = 9
        ' 
        ' ckbRawData
        ' 
        Me.ckbRawData.AutoSize = True
        Me.ckbRawData.CheckAlign = ContentAlignment.MiddleRight
        Me.ckbRawData.Dock = DockStyle.Top
        Me.ckbRawData.Location = New Point(50, 290)
        Me.ckbRawData.Name = "ckbRawData"
        Me.ckbRawData.Size = New Size(83, 19)
        Me.ckbRawData.TabIndex = 12
        Me.ckbRawData.Text = "Raw Data"
        Me.ckbRawData.TextAlign = ContentAlignment.TopRight
        ' 
        ' txtRawData
        ' 
        TableLayoutPanel1.SetColumnSpan(txtRawData, 3)
        txtRawData.Dock = DockStyle.Fill
        txtRawData.Location = New Point(139, 290)
        txtRawData.Multiline = True
        txtRawData.Name = "txtRawData"
        txtRawData.Size = New Size(610, 69)
        txtRawData.TabIndex = 13
        ' 
        ' btnSend
        ' 
        btnSend.Dock = DockStyle.Right
        btnSend.Location = New Point(639, 403)
        btnSend.Name = "btnSend"
        btnSend.Size = New Size(110, 34)
        btnSend.TabIndex = 14
        btnSend.Text = "Write Event Log"
        btnSend.UseVisualStyleBackColor = True
        ' 
        ' cboEventType
        ' 
        TableLayoutPanel1.SetColumnSpan(cboEventType, 2)
        cboEventType.Dock = DockStyle.Fill
        cboEventType.FormattingEnabled = True
        cboEventType.Location = New Point(139, 99)
        cboEventType.Name = "cboEventType"
        cboEventType.Size = New Size(370, 23)
        cboEventType.TabIndex = 16
        ' 
        ' ckbEventType
        ' 
        ckbEventType.AutoSize = True
        ckbEventType.CheckAlign = ContentAlignment.MiddleRight
        ckbEventType.Location = New Point(50, 99)
        ckbEventType.Name = "ckbEventType"
        ckbEventType.Size = New Size(83, 19)
        ckbEventType.TabIndex = 17
        ckbEventType.Text = "Event Type"
        ckbEventType.TextAlign = ContentAlignment.MiddleRight
        ckbEventType.UseVisualStyleBackColor = True
        ' 
        ' lblStatus
        ' 
        lblStatus.AutoSize = True
        TableLayoutPanel1.SetColumnSpan(lblStatus, 3)
        lblStatus.ForeColor = Color.Green
        lblStatus.Location = New Point(50, 400)
        lblStatus.Name = "lblStatus"
        lblStatus.Size = New Size(393, 15)
        lblStatus.TabIndex = 18
        lblStatus.Text = "This is the label to hold the status information when a log event happens."
        ' 
        ' tmr
        ' 
        tmr.Interval = 3000
        ' 
        ' EventLogTester
        ' 
        AutoScaleDimensions = New SizeF(7.0F, 15.0F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(800, 450)
        Controls.Add(TableLayoutPanel1)
        Icon = CType(resources.GetObject("$this.Icon"), Icon)
        Name = "EventLogTester"
        Text = "EventLogTester"
        TableLayoutPanel1.ResumeLayout(False)
        TableLayoutPanel1.PerformLayout()
        CType(numEventID, ComponentModel.ISupportInitialize).EndInit()
        CType(numCategory, ComponentModel.ISupportInitialize).EndInit()
        ResumeLayout(False)
    End Sub

    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents lblLog As Label
    Friend WithEvents cboLogName As ComboBox
    Friend WithEvents lblSource As Label
    Friend WithEvents cboSourceName As ComboBox
    Friend WithEvents lblMessage As Label
    Friend WithEvents txtMessage As TextBox
    Friend WithEvents ckbSeverity As CheckBox
    Friend WithEvents cboSeverity As ComboBox
    Friend WithEvents ckbEventID As CheckBox
    Friend WithEvents numEventID As NumericUpDown
    Friend WithEvents ckbCategory As CheckBox
    Friend WithEvents numCategory As NumericUpDown
    Friend WithEvents ckbRawData As CheckBox
    Friend WithEvents txtRawData As TextBox
    Friend WithEvents btnSend As Button
    Friend WithEvents ckbEventType As CheckBox
    Friend WithEvents cboEventType As ComboBox
    Friend WithEvents lblStatus As Label
    Friend WithEvents tmr As Timer

End Class
