<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EventLogTester
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EventLogTester))
        TableLayoutPanel1 = New TableLayoutPanel()
        lblLog = New Label()
        ComboBox1 = New ComboBox()
        lblSource = New Label()
        ComboBox2 = New ComboBox()
        lblMessage = New Label()
        TextBox1 = New TextBox()
        TableLayoutPanel1.SuspendLayout()
        SuspendLayout()
        ' 
        ' TableLayoutPanel1
        ' 
        TableLayoutPanel1.ColumnCount = 5
        TableLayoutPanel1.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50F))
        TableLayoutPanel1.ColumnStyles.Add(New ColumnStyle())
        TableLayoutPanel1.ColumnStyles.Add(New ColumnStyle())
        TableLayoutPanel1.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 20F))
        TableLayoutPanel1.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50F))
        TableLayoutPanel1.Controls.Add(lblLog, 1, 1)
        TableLayoutPanel1.Controls.Add(ComboBox1, 2, 1)
        TableLayoutPanel1.Controls.Add(lblSource, 1, 2)
        TableLayoutPanel1.Controls.Add(ComboBox2, 2, 2)
        TableLayoutPanel1.Controls.Add(lblMessage, 1, 3)
        TableLayoutPanel1.Controls.Add(TextBox1, 2, 3)
        TableLayoutPanel1.Dock = DockStyle.Fill
        TableLayoutPanel1.Location = New Point(0, 0)
        TableLayoutPanel1.Name = "TableLayoutPanel1"
        TableLayoutPanel1.RowCount = 7
        TableLayoutPanel1.RowStyles.Add(New RowStyle(SizeType.Percent, 50F))
        TableLayoutPanel1.RowStyles.Add(New RowStyle())
        TableLayoutPanel1.RowStyles.Add(New RowStyle())
        TableLayoutPanel1.RowStyles.Add(New RowStyle())
        TableLayoutPanel1.RowStyles.Add(New RowStyle(SizeType.Absolute, 20F))
        TableLayoutPanel1.RowStyles.Add(New RowStyle(SizeType.Absolute, 20F))
        TableLayoutPanel1.RowStyles.Add(New RowStyle(SizeType.Percent, 50F))
        TableLayoutPanel1.Size = New Size(800, 450)
        TableLayoutPanel1.TabIndex = 0
        ' 
        ' lblLog
        ' 
        lblLog.AutoSize = True
        lblLog.Dock = DockStyle.Right
        lblLog.Location = New Point(303, 135)
        lblLog.Name = "lblLog"
        lblLog.Size = New Size(62, 29)
        lblLog.TabIndex = 0
        lblLog.Text = "Log Name"
        lblLog.TextAlign = ContentAlignment.MiddleRight
        ' 
        ' ComboBox1
        ' 
        ComboBox1.Dock = DockStyle.Left
        ComboBox1.FormattingEnabled = True
        ComboBox1.Location = New Point(371, 138)
        ComboBox1.Name = "ComboBox1"
        ComboBox1.Size = New Size(121, 23)
        ComboBox1.TabIndex = 1
        ' 
        ' lblSource
        ' 
        lblSource.AutoSize = True
        lblSource.Dock = DockStyle.Right
        lblSource.Location = New Point(287, 164)
        lblSource.Name = "lblSource"
        lblSource.Size = New Size(78, 29)
        lblSource.TabIndex = 2
        lblSource.Text = "Source Name"
        lblSource.TextAlign = ContentAlignment.MiddleRight
        ' 
        ' ComboBox2
        ' 
        ComboBox2.Dock = DockStyle.Left
        ComboBox2.FormattingEnabled = True
        ComboBox2.Location = New Point(371, 167)
        ComboBox2.Name = "ComboBox2"
        ComboBox2.Size = New Size(121, 23)
        ComboBox2.TabIndex = 3
        ' 
        ' lblMessage
        ' 
        lblMessage.AutoSize = True
        lblMessage.Dock = DockStyle.Right
        lblMessage.Location = New Point(312, 193)
        lblMessage.Name = "lblMessage"
        lblMessage.Size = New Size(53, 81)
        lblMessage.TabIndex = 4
        lblMessage.Text = "Message"
        lblMessage.TextAlign = ContentAlignment.TopRight
        ' 
        ' TextBox1
        ' 
        TableLayoutPanel1.SetColumnSpan(TextBox1, 2)
        TextBox1.Dock = DockStyle.Fill
        TextBox1.Location = New Point(371, 196)
        TextBox1.Multiline = True
        TextBox1.Name = "TextBox1"
        TextBox1.Size = New Size(141, 75)
        TextBox1.TabIndex = 5
        ' 
        ' EventLogTester
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(800, 450)
        Controls.Add(TableLayoutPanel1)
        Icon = CType(resources.GetObject("$this.Icon"), Icon)
        Name = "EventLogTester"
        Text = "EventLogTester"
        TableLayoutPanel1.ResumeLayout(False)
        TableLayoutPanel1.PerformLayout()
        ResumeLayout(False)
    End Sub

    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents lblLog As Label
    Friend WithEvents ComboBox1 As ComboBox
    Friend WithEvents lblSource As Label
    Friend WithEvents ComboBox2 As ComboBox
    Friend WithEvents lblMessage As Label
    Friend WithEvents TextBox1 As TextBox
End Class
