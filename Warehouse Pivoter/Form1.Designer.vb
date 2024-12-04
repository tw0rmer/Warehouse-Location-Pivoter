<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.ListView1 = New System.Windows.Forms.ListView()
        Me.TabPage4 = New System.Windows.Forms.TabPage()
        Me.ListView2 = New System.Windows.Forms.ListView()
        Me.TabPage3 = New System.Windows.Forms.TabPage()
        Me.rtbDebuggerOutput = New System.Windows.Forms.RichTextBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.btnExportLocation = New System.Windows.Forms.Button()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.txtLocationFilter = New System.Windows.Forms.TextBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.cbMovableUnitLabel = New System.Windows.Forms.ComboBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.cbSKU = New System.Windows.Forms.ComboBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.cbLocation = New System.Windows.Forms.ComboBox()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OpenxlsxToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.TableLayoutPanel1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage4.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 1
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 632.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.TabControl1, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Panel1, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.ProgressBar1, 0, 2)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 3
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 88.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 26.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(640, 620)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage4)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(3, 91)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(634, 500)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.ListView1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(626, 474)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "DATA"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'ListView1
        '
        Me.ListView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ListView1.FullRowSelect = True
        Me.ListView1.HideSelection = False
        Me.ListView1.Location = New System.Drawing.Point(3, 3)
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(620, 468)
        Me.ListView1.TabIndex = 0
        Me.ListView1.UseCompatibleStateImageBehavior = False
        Me.ListView1.View = System.Windows.Forms.View.Details
        '
        'TabPage4
        '
        Me.TabPage4.Controls.Add(Me.ListView2)
        Me.TabPage4.Location = New System.Drawing.Point(4, 22)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage4.Size = New System.Drawing.Size(626, 474)
        Me.TabPage4.TabIndex = 3
        Me.TabPage4.Text = "TEMP"
        Me.TabPage4.UseVisualStyleBackColor = True
        '
        'ListView2
        '
        Me.ListView2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ListView2.FullRowSelect = True
        Me.ListView2.HideSelection = False
        Me.ListView2.Location = New System.Drawing.Point(3, 3)
        Me.ListView2.Name = "ListView2"
        Me.ListView2.Size = New System.Drawing.Size(620, 468)
        Me.ListView2.TabIndex = 0
        Me.ListView2.UseCompatibleStateImageBehavior = False
        Me.ListView2.View = System.Windows.Forms.View.Details
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.rtbDebuggerOutput)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Size = New System.Drawing.Size(626, 474)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Debugger"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'rtbDebuggerOutput
        '
        Me.rtbDebuggerOutput.Dock = System.Windows.Forms.DockStyle.Fill
        Me.rtbDebuggerOutput.Location = New System.Drawing.Point(0, 0)
        Me.rtbDebuggerOutput.Name = "rtbDebuggerOutput"
        Me.rtbDebuggerOutput.Size = New System.Drawing.Size(626, 474)
        Me.rtbDebuggerOutput.TabIndex = 2
        Me.rtbDebuggerOutput.Text = ""
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.btnExportLocation)
        Me.Panel1.Controls.Add(Me.GroupBox4)
        Me.Panel1.Controls.Add(Me.GroupBox3)
        Me.Panel1.Controls.Add(Me.GroupBox2)
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Controls.Add(Me.MenuStrip1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(3, 3)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(634, 82)
        Me.Panel1.TabIndex = 1
        '
        'btnExportLocation
        '
        Me.btnExportLocation.Location = New System.Drawing.Point(460, 35)
        Me.btnExportLocation.Name = "btnExportLocation"
        Me.btnExportLocation.Size = New System.Drawing.Size(75, 36)
        Me.btnExportLocation.TabIndex = 6
        Me.btnExportLocation.Text = "Export html"
        Me.btnExportLocation.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.txtLocationFilter)
        Me.GroupBox4.Location = New System.Drawing.Point(317, 31)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(137, 40)
        Me.GroupBox4.TabIndex = 5
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Location Prefix Filter"
        '
        'txtLocationFilter
        '
        Me.txtLocationFilter.Location = New System.Drawing.Point(6, 15)
        Me.txtLocationFilter.Name = "txtLocationFilter"
        Me.txtLocationFilter.Size = New System.Drawing.Size(125, 20)
        Me.txtLocationFilter.TabIndex = 4
        Me.txtLocationFilter.Text = "A1"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.cbMovableUnitLabel)
        Me.GroupBox3.Location = New System.Drawing.Point(197, 31)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(114, 40)
        Me.GroupBox3.TabIndex = 3
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Movable Unit Label"
        '
        'cbMovableUnitLabel
        '
        Me.cbMovableUnitLabel.FormattingEnabled = True
        Me.cbMovableUnitLabel.Location = New System.Drawing.Point(6, 15)
        Me.cbMovableUnitLabel.Name = "cbMovableUnitLabel"
        Me.cbMovableUnitLabel.Size = New System.Drawing.Size(99, 21)
        Me.cbMovableUnitLabel.TabIndex = 0
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.cbSKU)
        Me.GroupBox2.Location = New System.Drawing.Point(4, 31)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(90, 40)
        Me.GroupBox2.TabIndex = 2
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "SKU"
        '
        'cbSKU
        '
        Me.cbSKU.FormattingEnabled = True
        Me.cbSKU.Location = New System.Drawing.Point(5, 15)
        Me.cbSKU.Name = "cbSKU"
        Me.cbSKU.Size = New System.Drawing.Size(74, 21)
        Me.cbSKU.TabIndex = 0
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.cbLocation)
        Me.GroupBox1.Location = New System.Drawing.Point(100, 31)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(91, 40)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Location"
        '
        'cbLocation
        '
        Me.cbLocation.FormattingEnabled = True
        Me.cbLocation.Location = New System.Drawing.Point(6, 15)
        Me.cbLocation.Name = "cbLocation"
        Me.cbLocation.Size = New System.Drawing.Size(74, 21)
        Me.cbLocation.TabIndex = 0
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(634, 24)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.OpenxlsxToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.FileToolStripMenuItem.Text = "File"
        '
        'OpenxlsxToolStripMenuItem
        '
        Me.OpenxlsxToolStripMenuItem.Name = "OpenxlsxToolStripMenuItem"
        Me.OpenxlsxToolStripMenuItem.Size = New System.Drawing.Size(129, 22)
        Me.OpenxlsxToolStripMenuItem.Text = "Open .xlsx"
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ProgressBar1.Location = New System.Drawing.Point(3, 597)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(634, 20)
        Me.ProgressBar1.TabIndex = 2
        '
        'Timer1
        '
        Me.Timer1.Interval = 2000
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(640, 620)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MinimumSize = New System.Drawing.Size(656, 659)
        Me.Name = "Form1"
        Me.Text = "Warehouse Pivoter (specific to bulletproof logistics)"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage4.ResumeLayout(False)
        Me.TabPage3.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents TabPage1 As TabPage
    Friend WithEvents ListView1 As ListView
    Friend WithEvents Panel1 As Panel
    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents FileToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents OpenxlsxToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents cbLocation As ComboBox
    Friend WithEvents TabPage3 As TabPage
    Friend WithEvents rtbDebuggerOutput As RichTextBox
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents cbMovableUnitLabel As ComboBox
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents cbSKU As ComboBox
    Friend WithEvents TabPage4 As TabPage
    Friend WithEvents ListView2 As ListView
    Friend WithEvents btnExportLocation As Button
    Friend WithEvents GroupBox4 As GroupBox
    Friend WithEvents txtLocationFilter As TextBox
    Friend WithEvents ProgressBar1 As ProgressBar
    Friend WithEvents Timer1 As Timer
End Class
