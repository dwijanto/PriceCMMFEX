<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormHistoryDetail
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormHistoryDetail))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.ToolStripComboBox1 = New System.Windows.Forms.ToolStripComboBox()
        Me.ToolStripTextBox1 = New System.Windows.Forms.ToolStripTextBox()
        Me.ToolStripButton1 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton2 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton3 = New System.Windows.Forms.ToolStripButton()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripStatusLabel2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripProgressBar1 = New System.Windows.Forms.ToolStripProgressBar()
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column24 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column21 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column6 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column11 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column7 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column16 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column31 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column8 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column9 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column20 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column18 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column10 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column22 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column23 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column32 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column12 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column14 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column17 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column15 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column19 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column13 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column25 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column26 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column27 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column28 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column29 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column30 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ToolStrip1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StatusStrip1.SuspendLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ToolStrip1
        '
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripComboBox1, Me.ToolStripTextBox1, Me.ToolStripButton1, Me.ToolStripButton2, Me.ToolStripButton3})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(1015, 25)
        Me.ToolStrip1.TabIndex = 1
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'ToolStripComboBox1
        '
        Me.ToolStripComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ToolStripComboBox1.Items.AddRange(New Object() {"CMMF (leave blank for All records)", "Supplier Code", "Valid Date yyyy-mm-dd", "Submit Date yyyy-mm-dd"})
        Me.ToolStripComboBox1.Name = "ToolStripComboBox1"
        Me.ToolStripComboBox1.Size = New System.Drawing.Size(250, 25)
        '
        'ToolStripTextBox1
        '
        Me.ToolStripTextBox1.Name = "ToolStripTextBox1"
        Me.ToolStripTextBox1.Size = New System.Drawing.Size(100, 25)
        '
        'ToolStripButton1
        '
        Me.ToolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButton1.Image = CType(resources.GetObject("ToolStripButton1.Image"), System.Drawing.Image)
        Me.ToolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton1.Name = "ToolStripButton1"
        Me.ToolStripButton1.Size = New System.Drawing.Size(46, 22)
        Me.ToolStripButton1.Text = "Search"
        '
        'ToolStripButton2
        '
        Me.ToolStripButton2.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButton2.Image = CType(resources.GetObject("ToolStripButton2.Image"), System.Drawing.Image)
        Me.ToolStripButton2.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton2.Name = "ToolStripButton2"
        Me.ToolStripButton2.Size = New System.Drawing.Size(92, 22)
        Me.ToolStripButton2.Text = "Extract To Excel"
        '
        'ToolStripButton3
        '
        Me.ToolStripButton3.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButton3.Image = CType(resources.GetObject("ToolStripButton3.Image"), System.Drawing.Image)
        Me.ToolStripButton3.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton3.Name = "ToolStripButton3"
        Me.ToolStripButton3.Size = New System.Drawing.Size(44, 22)
        Me.ToolStripButton3.Text = "Delete"
        Me.ToolStripButton3.Visible = False
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.AllowUserToOrderColumns = True
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.BackgroundColor = System.Drawing.SystemColors.Window
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column1, Me.Column2, Me.Column24, Me.Column3, Me.Column4, Me.Column21, Me.Column5, Me.Column6, Me.Column11, Me.Column7, Me.Column16, Me.Column31, Me.Column8, Me.Column9, Me.Column20, Me.Column18, Me.Column10, Me.Column22, Me.Column23, Me.Column32, Me.Column12, Me.Column14, Me.Column17, Me.Column15, Me.Column19, Me.Column13, Me.Column25, Me.Column26, Me.Column27, Me.Column28, Me.Column29, Me.Column30})
        Me.DataGridView1.Location = New System.Drawing.Point(0, 28)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView1.Size = New System.Drawing.Size(1015, 429)
        Me.DataGridView1.TabIndex = 2
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripStatusLabel2, Me.ToolStripProgressBar1})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 460)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(1015, 22)
        Me.StatusStrip1.TabIndex = 64
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(0, 17)
        '
        'ToolStripStatusLabel2
        '
        Me.ToolStripStatusLabel2.BackColor = System.Drawing.SystemColors.Control
        Me.ToolStripStatusLabel2.Name = "ToolStripStatusLabel2"
        Me.ToolStripStatusLabel2.Size = New System.Drawing.Size(898, 17)
        Me.ToolStripStatusLabel2.Spring = True
        '
        'ToolStripProgressBar1
        '
        Me.ToolStripProgressBar1.Name = "ToolStripProgressBar1"
        Me.ToolStripProgressBar1.Size = New System.Drawing.Size(100, 16)
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.ContainerControl = Me
        '
        'Column1
        '
        Me.Column1.DataPropertyName = "cmmf"
        Me.Column1.HeaderText = "CMMF"
        Me.Column1.Name = "Column1"
        Me.Column1.ReadOnly = True
        '
        'Column2
        '
        Me.Column2.DataPropertyName = "materialdesc"
        Me.Column2.HeaderText = "Material Description"
        Me.Column2.Name = "Column2"
        Me.Column2.ReadOnly = True
        Me.Column2.Width = 200
        '
        'Column24
        '
        Me.Column24.DataPropertyName = "commercialref"
        Me.Column24.HeaderText = "Commercial Code"
        Me.Column24.Name = "Column24"
        Me.Column24.ReadOnly = True
        '
        'Column3
        '
        Me.Column3.DataPropertyName = "vendorcode"
        Me.Column3.HeaderText = "Vendor Code"
        Me.Column3.Name = "Column3"
        Me.Column3.ReadOnly = True
        '
        'Column4
        '
        Me.Column4.DataPropertyName = "vendorname"
        Me.Column4.HeaderText = "Vendor Name"
        Me.Column4.Name = "Column4"
        Me.Column4.ReadOnly = True
        Me.Column4.Width = 150
        '
        'Column21
        '
        Me.Column21.DataPropertyName = "shortname"
        Me.Column21.HeaderText = "Short Name"
        Me.Column21.Name = "Column21"
        Me.Column21.ReadOnly = True
        '
        'Column5
        '
        Me.Column5.DataPropertyName = "purchorg"
        Me.Column5.HeaderText = "PurchOrg"
        Me.Column5.Name = "Column5"
        Me.Column5.ReadOnly = True
        '
        'Column6
        '
        Me.Column6.DataPropertyName = "plant"
        Me.Column6.HeaderText = "Plant"
        Me.Column6.Name = "Column6"
        Me.Column6.ReadOnly = True
        '
        'Column11
        '
        Me.Column11.DataPropertyName = "submitdate"
        DataGridViewCellStyle1.Format = "dd-MMM-yyyy"
        Me.Column11.DefaultCellStyle = DataGridViewCellStyle1
        Me.Column11.HeaderText = "Submit Date"
        Me.Column11.Name = "Column11"
        Me.Column11.ReadOnly = True
        '
        'Column7
        '
        Me.Column7.DataPropertyName = "validon"
        DataGridViewCellStyle2.Format = "dd-MMM-yyyy"
        Me.Column7.DefaultCellStyle = DataGridViewCellStyle2
        Me.Column7.HeaderText = "New Valid Date"
        Me.Column7.Name = "Column7"
        Me.Column7.ReadOnly = True
        '
        'Column16
        '
        Me.Column16.DataPropertyName = "negotiateddate"
        DataGridViewCellStyle3.Format = "dd-MMM-yyyy"
        Me.Column16.DefaultCellStyle = DataGridViewCellStyle3
        Me.Column16.HeaderText = "Negotiated Effective Date"
        Me.Column16.Name = "Column16"
        Me.Column16.ReadOnly = True
        '
        'Column31
        '
        Me.Column31.DataPropertyName = "crcy"
        Me.Column31.HeaderText = "Currency"
        Me.Column31.Name = "Column31"
        Me.Column31.ReadOnly = True
        Me.Column31.Width = 70
        '
        'Column8
        '
        Me.Column8.DataPropertyName = "price"
        Me.Column8.HeaderText = "Price"
        Me.Column8.Name = "Column8"
        Me.Column8.ReadOnly = True
        '
        'Column9
        '
        Me.Column9.DataPropertyName = "pricingunit"
        Me.Column9.HeaderText = "Pricing Unit"
        Me.Column9.Name = "Column9"
        Me.Column9.ReadOnly = True
        '
        'Column20
        '
        Me.Column20.DataPropertyName = "unitprice"
        DataGridViewCellStyle4.Format = "0.000,#####"
        Me.Column20.DefaultCellStyle = DataGridViewCellStyle4
        Me.Column20.HeaderText = "Unit Price"
        Me.Column20.Name = "Column20"
        Me.Column20.ReadOnly = True
        '
        'Column18
        '
        Me.Column18.DataPropertyName = "pricetype"
        Me.Column18.HeaderText = "Price Type"
        Me.Column18.Name = "Column18"
        Me.Column18.ReadOnly = True
        '
        'Column10
        '
        Me.Column10.DataPropertyName = "comment"
        Me.Column10.HeaderText = "Comment"
        Me.Column10.Name = "Column10"
        Me.Column10.ReadOnly = True
        '
        'Column22
        '
        Me.Column22.DataPropertyName = "description"
        Me.Column22.HeaderText = "Description"
        Me.Column22.Name = "Column22"
        Me.Column22.ReadOnly = True
        '
        'Column23
        '
        Me.Column23.DataPropertyName = "reasonname"
        Me.Column23.HeaderText = "Reason"
        Me.Column23.Name = "Column23"
        Me.Column23.ReadOnly = True
        '
        'Column32
        '
        Me.Column32.DataPropertyName = "specialproject"
        Me.Column32.HeaderText = "Special Project"
        Me.Column32.Name = "Column32"
        Me.Column32.ReadOnly = True
        '
        'Column12
        '
        Me.Column12.DataPropertyName = "creatorname"
        Me.Column12.HeaderText = "Creator"
        Me.Column12.Name = "Column12"
        Me.Column12.ReadOnly = True
        '
        'Column14
        '
        Me.Column14.DataPropertyName = "validator"
        Me.Column14.HeaderText = "Validator"
        Me.Column14.Name = "Column14"
        Me.Column14.ReadOnly = True
        '
        'Column17
        '
        Me.Column17.DataPropertyName = "cc"
        Me.Column17.HeaderText = "CC"
        Me.Column17.Name = "Column17"
        Me.Column17.ReadOnly = True
        '
        'Column15
        '
        Me.Column15.DataPropertyName = "statusname"
        Me.Column15.HeaderText = "Status"
        Me.Column15.Name = "Column15"
        Me.Column15.ReadOnly = True
        '
        'Column19
        '
        Me.Column19.DataPropertyName = "changeby"
        Me.Column19.HeaderText = "Changeby"
        Me.Column19.Name = "Column19"
        Me.Column19.ReadOnly = True
        '
        'Column13
        '
        Me.Column13.DataPropertyName = "pricechangehdid"
        Me.Column13.HeaderText = "PriceChangeHDId"
        Me.Column13.Name = "Column13"
        Me.Column13.ReadOnly = True
        '
        'Column25
        '
        Me.Column25.DataPropertyName = "amount"
        Me.Column25.HeaderText = "SAP Price"
        Me.Column25.Name = "Column25"
        Me.Column25.ReadOnly = True
        '
        'Column26
        '
        Me.Column26.DataPropertyName = "perunit"
        Me.Column26.HeaderText = "SAP Perunit"
        Me.Column26.Name = "Column26"
        Me.Column26.ReadOnly = True
        '
        'Column27
        '
        Me.Column27.DataPropertyName = "validfrom"
        DataGridViewCellStyle5.Format = "dd-MM-yyyy"
        Me.Column27.DefaultCellStyle = DataGridViewCellStyle5
        Me.Column27.HeaderText = "SAP Valid From"
        Me.Column27.Name = "Column27"
        Me.Column27.ReadOnly = True
        '
        'Column28
        '
        Me.Column28.DataPropertyName = "deltasap"
        DataGridViewCellStyle6.Format = "0.00%"
        Me.Column28.DefaultCellStyle = DataGridViewCellStyle6
        Me.Column28.HeaderText = "Delta SAP"
        Me.Column28.Name = "Column28"
        Me.Column28.ReadOnly = True
        '
        'Column29
        '
        Me.Column29.DataPropertyName = "deltastd"
        DataGridViewCellStyle7.Format = "0.00%"
        Me.Column29.DefaultCellStyle = DataGridViewCellStyle7
        Me.Column29.HeaderText = "Delta STD"
        Me.Column29.Name = "Column29"
        Me.Column29.ReadOnly = True
        '
        'Column30
        '
        Me.Column30.DataPropertyName = "alert"
        Me.Column30.HeaderText = "Alert"
        Me.Column30.Name = "Column30"
        Me.Column30.ReadOnly = True
        '
        'FormHistoryDetail
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1015, 482)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Name = "FormHistoryDetail"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FormHistoryDetail"
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents ToolStripComboBox1 As System.Windows.Forms.ToolStripComboBox
    Friend WithEvents ToolStripTextBox1 As System.Windows.Forms.ToolStripTextBox
    Friend WithEvents ToolStripButton1 As System.Windows.Forms.ToolStripButton
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Public WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Public WithEvents ToolStripStatusLabel2 As System.Windows.Forms.ToolStripStatusLabel
    Public WithEvents ToolStripProgressBar1 As System.Windows.Forms.ToolStripProgressBar
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents ToolStripButton2 As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButton3 As System.Windows.Forms.ToolStripButton
    Friend WithEvents Column1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column24 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column21 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column6 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column11 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column7 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column16 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column31 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column8 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column9 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column20 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column18 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column10 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column22 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column23 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column32 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column12 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column14 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column17 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column15 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column19 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column13 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column25 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column26 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column27 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column28 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column29 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column30 As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
