<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormMenu
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormMenu))
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.ImportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.BSEGSQ01ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ZZA0035ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PO40SQ01PO40AndPO41LocalfileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PO39ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripSeparator()
        Me.ImportSavingsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ImportZZA037ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AveragePriceIndexToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AveragePriceIndexSavingsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PriceChangeTaskToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SupplierDocumentsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.UserGuideToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ImportToolStripMenuItem, Me.ReportToolStripMenuItem, Me.PriceChangeTaskToolStripMenuItem, Me.SupplierDocumentsToolStripMenuItem, Me.HelpToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(547, 24)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ImportToolStripMenuItem
        '
        Me.ImportToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BSEGSQ01ToolStripMenuItem, Me.ZZA0035ToolStripMenuItem, Me.PO40SQ01PO40AndPO41LocalfileToolStripMenuItem, Me.PO39ToolStripMenuItem, Me.ToolStripMenuItem1, Me.ImportSavingsToolStripMenuItem, Me.ImportZZA037ToolStripMenuItem})
        Me.ImportToolStripMenuItem.Name = "ImportToolStripMenuItem"
        Me.ImportToolStripMenuItem.Size = New System.Drawing.Size(55, 20)
        Me.ImportToolStripMenuItem.Text = "Import"
        '
        'BSEGSQ01ToolStripMenuItem
        '
        Me.BSEGSQ01ToolStripMenuItem.Name = "BSEGSQ01ToolStripMenuItem"
        Me.BSEGSQ01ToolStripMenuItem.Size = New System.Drawing.Size(597, 22)
        Me.BSEGSQ01ToolStripMenuItem.Tag = "FormImportBSEG"
        Me.BSEGSQ01ToolStripMenuItem.Text = "BSEG (SQ01,F037-FGPURCH,0029,OUTPUT Format->File Store,Encoding 1133,With Column " & _
            "Header)"
        '
        'ZZA0035ToolStripMenuItem
        '
        Me.ZZA0035ToolStripMenuItem.Name = "ZZA0035ToolStripMenuItem"
        Me.ZZA0035ToolStripMenuItem.Size = New System.Drawing.Size(597, 22)
        Me.ZZA0035ToolStripMenuItem.Tag = "FormImportZZ0035"
        Me.ZZA0035ToolStripMenuItem.Text = "ZZ0035 ( Local File, Spreadsheet)"
        '
        'PO40SQ01PO40AndPO41LocalfileToolStripMenuItem
        '
        Me.PO40SQ01PO40AndPO41LocalfileToolStripMenuItem.Name = "PO40SQ01PO40AndPO41LocalfileToolStripMenuItem"
        Me.PO40SQ01PO40AndPO41LocalfileToolStripMenuItem.Size = New System.Drawing.Size(597, 22)
        Me.PO40SQ01PO40AndPO41LocalfileToolStripMenuItem.Tag = "FormImportPO40Plus"
        Me.PO40SQ01PO40AndPO41LocalfileToolStripMenuItem.Text = "PO40 (SQ01,PO40  And PO41, Localfile)"
        Me.PO40SQ01PO40AndPO41LocalfileToolStripMenuItem.Visible = False
        '
        'PO39ToolStripMenuItem
        '
        Me.PO39ToolStripMenuItem.Name = "PO39ToolStripMenuItem"
        Me.PO39ToolStripMenuItem.Size = New System.Drawing.Size(597, 22)
        Me.PO39ToolStripMenuItem.Tag = "FormPO39"
        Me.PO39ToolStripMenuItem.Text = "PO39 (SQ01,PO39 Plant 3750, Localfile)"
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(594, 6)
        '
        'ImportSavingsToolStripMenuItem
        '
        Me.ImportSavingsToolStripMenuItem.Name = "ImportSavingsToolStripMenuItem"
        Me.ImportSavingsToolStripMenuItem.Size = New System.Drawing.Size(597, 22)
        Me.ImportSavingsToolStripMenuItem.Tag = "FormImportSaving"
        Me.ImportSavingsToolStripMenuItem.Text = "Import Savings (Text File -> Tab Delimited)"
        '
        'ImportZZA037ToolStripMenuItem
        '
        Me.ImportZZA037ToolStripMenuItem.Name = "ImportZZA037ToolStripMenuItem"
        Me.ImportZZA037ToolStripMenuItem.Size = New System.Drawing.Size(597, 22)
        Me.ImportZZA037ToolStripMenuItem.Tag = "FormImportZFA037"
        Me.ImportZZA037ToolStripMenuItem.Text = "Import ZFA037 (Price List)"
        '
        'ReportToolStripMenuItem
        '
        Me.ReportToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AveragePriceIndexToolStripMenuItem, Me.AveragePriceIndexSavingsToolStripMenuItem})
        Me.ReportToolStripMenuItem.Name = "ReportToolStripMenuItem"
        Me.ReportToolStripMenuItem.Size = New System.Drawing.Size(54, 20)
        Me.ReportToolStripMenuItem.Text = "Report"
        '
        'AveragePriceIndexToolStripMenuItem
        '
        Me.AveragePriceIndexToolStripMenuItem.Name = "AveragePriceIndexToolStripMenuItem"
        Me.AveragePriceIndexToolStripMenuItem.Size = New System.Drawing.Size(233, 22)
        Me.AveragePriceIndexToolStripMenuItem.Tag = "ReportAVPI"
        Me.AveragePriceIndexToolStripMenuItem.Text = "Average Price Index"
        '
        'AveragePriceIndexSavingsToolStripMenuItem
        '
        Me.AveragePriceIndexSavingsToolStripMenuItem.Name = "AveragePriceIndexSavingsToolStripMenuItem"
        Me.AveragePriceIndexSavingsToolStripMenuItem.Size = New System.Drawing.Size(233, 22)
        Me.AveragePriceIndexSavingsToolStripMenuItem.Tag = "ReportAVPISaving"
        Me.AveragePriceIndexSavingsToolStripMenuItem.Text = "Average Price Index && Savings"
        '
        'PriceChangeTaskToolStripMenuItem
        '
        Me.PriceChangeTaskToolStripMenuItem.Name = "PriceChangeTaskToolStripMenuItem"
        Me.PriceChangeTaskToolStripMenuItem.Size = New System.Drawing.Size(116, 20)
        Me.PriceChangeTaskToolStripMenuItem.Text = "Price Change Task"
        '
        'SupplierDocumentsToolStripMenuItem
        '
        Me.SupplierDocumentsToolStripMenuItem.Name = "SupplierDocumentsToolStripMenuItem"
        Me.SupplierDocumentsToolStripMenuItem.Size = New System.Drawing.Size(126, 20)
        Me.SupplierDocumentsToolStripMenuItem.Tag = "FormDocumentHeader"
        Me.SupplierDocumentsToolStripMenuItem.Text = "Supplier Documents"
        Me.SupplierDocumentsToolStripMenuItem.Visible = False
        '
        'HelpToolStripMenuItem
        '
        Me.HelpToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.UserGuideToolStripMenuItem})
        Me.HelpToolStripMenuItem.Name = "HelpToolStripMenuItem"
        Me.HelpToolStripMenuItem.Size = New System.Drawing.Size(44, 20)
        Me.HelpToolStripMenuItem.Text = "Help"
        '
        'UserGuideToolStripMenuItem
        '
        Me.UserGuideToolStripMenuItem.Name = "UserGuideToolStripMenuItem"
        Me.UserGuideToolStripMenuItem.Size = New System.Drawing.Size(131, 22)
        Me.UserGuideToolStripMenuItem.Text = "User Guide"
        '
        'FormMenu
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(547, 111)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "FormMenu"
        Me.Text = "Price CMMF Ext"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents ImportToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BSEGSQ01ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ReportToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AveragePriceIndexToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ZZA0035ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PO40SQ01PO40AndPO41LocalfileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ImportSavingsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AveragePriceIndexSavingsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PriceChangeTaskToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ImportZZA037ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SupplierDocumentsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HelpToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UserGuideToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PO39ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem

End Class
