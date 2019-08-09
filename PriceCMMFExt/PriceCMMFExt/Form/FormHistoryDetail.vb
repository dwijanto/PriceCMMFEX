Imports System.Threading
Imports System.Text
Imports PriceCMMFExt.PublicClass
Public Class FormHistoryDetail
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf doSearch)
    Private DS As DataSet
    Private BS As BindingSource
    Private SB As StringBuilder
    Private myUser As String
    Dim mycriteria As String = String.Empty
    Dim recordcount As Integer = 0
    Dim sqlstr As String = String.Empty
    Protected CM As CurrencyManager
    Dim NeedRefresh As Boolean = False
    Dim MyPeriod As UCPeriodRange
    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        If Not myThread.IsAlive Then
            If Me.Validate() Then
                myThread = New Thread(AddressOf doSearch)
                myThread.Start()
            End If
        Else
            MessageBox.Show("Please wait until the current process finished.")
        End If
    End Sub

    Private Overloads Function Validate() As Boolean
        Dim myret As Boolean = True
        mycriteria = ""
        MyBase.Validate()
        Select Case ToolStripComboBox1.SelectedIndex
            Case 0
                If ToolStripTextBox1.Text = "" Then
                    mycriteria = ""
                Else
                    mycriteria = "c.cmmf =" & ToolStripTextBox1.Text.Replace("'", "''")
                End If

            Case 1
                If IsNumeric(ToolStripTextBox1.Text) Then

                    mycriteria = "v.vendorcode =" & ToolStripTextBox1.Text
                Else
                    MessageBox.Show("Please check the value.")
                    myret = False
                End If

            Case 3
                If IsDate(ToolStripTextBox1.Text) Then
                    mycriteria = "submitdate ='" & DbAdapter1.dateformatYYYYMMdd(ToolStripTextBox1.Text) & "'"
                ElseIf MyPeriod.CheckBox1.Checked Then
                    mycriteria = String.Format("submitdate >= ''{0}-01-01'' and submitdate <= ''{1}-12-31'' ", MyPeriod.Year1, MyPeriod.Year2)
                Else
                    MessageBox.Show("Please check the value.")
                    myret = False
                End If

            Case 2
                If IsDate(ToolStripTextBox1.Text) Then
                    mycriteria = "validon = '" & DbAdapter1.dateformatYYYYMMdd(ToolStripTextBox1.Text) & "'"
                ElseIf MyPeriod.CheckBox1.Checked Then
                    mycriteria = String.Format("validon >= ''{0}-01-01'' and  validon <= ''{1}-12-31'' ", MyPeriod.Year1, MyPeriod.Year2)

                Else
                    MessageBox.Show("Please check the value.")
                    myret = False
                End If

        End Select
        'mycriteria = "vendorcode = 10023684"       

        Return myret
    End Function

    Sub doSearch()
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Loading Data.")

        DS = New DataSet
        Dim mymessage As String = String.Empty
        'SB.Clear()


        myuser = HelperClass1.UserId.ToLower
        'myuser = "as\tchen"
        'myuser = "as\elai"
        'myuser = "as\rlo"
        'myUser = "as\afok"
        'myUser = "as\weho"
        myUser = "as\dummy" 'show all records

        'sqlstr = "select * from  sp_gethistorydetail001('" & myUser & "','" & mycriteria & "') as tb(cmmf bigint,commercialref character varying,vendorcode bigint,shortname character varying,purchorg integer,plant integer,submitdate date, validon date,price numeric,pricingunit integer,unitprice numeric,comment text,description text,reasonname text,creator character varying,materialdesc character varying,vendorname character varying,pricechangehdid bigint,creatorname character varying,validator character varying,cc character varying,pricetype character varying,negotiateddate date,changeby text,  statusname text,pricechangedtlid bigint)"
        'sqlstr = "with my as (select * from  sp_gethistorydetail001('" & myUser & "','" & mycriteria & "') as tb(cmmf bigint,commercialref character varying,vendorcode bigint,shortname character varying,purchorg integer,plant integer,submitdate date, validon date,price numeric,pricingunit integer,unitprice numeric,comment text,description text,reasonname text,creator character varying,materialdesc character varying,vendorname character varying,pricechangehdid bigint,creatorname character varying,validator character varying,cc character varying,pricetype character varying,negotiateddate date,changeby text,  statusname text,pricechangedtlid bigint))," &
        '         " std as(select cmmf, max(validfrom) as validfrom from standardcostad group by cmmf) " &
        '         " select my.* ,(getpriceinfo(my.statusname,my.cmmf,my.vendorcode,my.validon,my.price,my.pricingunit,ad.planprice1,ad.per)).*" &
        '         " from my" &
        '         " left join std on std.cmmf = my.cmmf " &
        '         " left join standardcostad ad on ad.cmmf = std.cmmf and ad.validfrom = std.validfrom "
        'sqlstr = "with my as (select * from  sp_gethistorydetail001('" & myUser & "','" & mycriteria & "') as tb(cmmf bigint,commercialref character varying,vendorcode bigint,shortname character varying,purchorg integer,plant integer,submitdate date, validon date,price numeric,pricingunit integer,unitprice numeric,comment text,description text,reasonname text,creator character varying,materialdesc character varying,vendorname character varying,pricechangehdid bigint,creatorname character varying,validator character varying,cc character varying,pricetype character varying,negotiateddate date,changeby text,  statusname text,pricechangedtlid bigint))," &
        ' " std as(select cmmf, max(validfrom) as validfrom from standardcostad group by cmmf) " &
        ' " select my.* ,(getpriceplantinfo(my.statusname,my.cmmf,my.vendorcode,my.plant,my.validon,my.price,my.pricingunit,ad.planprice1,ad.per)).*,doc.getvendorcurr(my.vendorcode,my.validon) as crcy" &
        ' " from my" &
        ' " left join std on std.cmmf = my.cmmf " &
        ' " left join standardcostad ad on ad.cmmf = std.cmmf and ad.validfrom = std.validfrom "
        sqlstr = "with my as (select * from  sp_gethistorydetail002('" & myUser & "','" & mycriteria & "') as tb(cmmf bigint,commercialref character varying,vendorcode bigint,shortname character varying,purchorg integer,plant integer,submitdate date, validon date,price numeric,pricingunit integer,unitprice numeric,comment text,description text,reasonname text,specialproject text,creator character varying,materialdesc character varying,vendorname character varying,pricechangehdid bigint,creatorname character varying,validator character varying,cc character varying,pricetype character varying,negotiateddate date,changeby text,  statusname text,pricechangedtlid bigint))," &
        " std as(select cmmf, max(validfrom) as validfrom from standardcostad group by cmmf) " &
        " select my.cmmf,my.commercialref,my.vendorcode,my.shortname,my.purchorg,my.plant,my.submitdate,my.validon,doc.getvendorcurr(my.vendorcode,my.validon) as crcy,my.price,my.pricingunit,my.unitprice ,my.comment,my.description,my.reasonname,my.specialproject,my.creator,my.materialdesc,my.vendorname ,my.pricechangehdid,my.creatorname ,my.validator,my.cc,my.pricetype ,my.negotiateddate,my.changeby,my.statusname,my.pricechangedtlid,(getpriceplantinfo(my.statusname,my.cmmf,my.vendorcode,my.plant,my.validon,my.price,my.pricingunit,ad.planprice1,ad.per)).*" &
        " from my" &
        " left join std on std.cmmf = my.cmmf " &
        " left join standardcostad ad on ad.cmmf = std.cmmf and ad.validfrom = std.validfrom "
        If DbAdapter1.TbgetDataSet(sqlstr, DS, mymessage) Then
            ProgressReport(4, "InitData")
        Else
            ProgressReport(1, "Loading Data. Error::" & mymessage)
            ProgressReport(5, "Continuous")
            Exit Sub
        End If
        ProgressReport(1, "Loading Data.Done! Record Found " & recordcount)
        ProgressReport(5, "Continuous")
    End Sub
    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Try
                Select Case id
                    Case 1
                        ToolStripStatusLabel1.Text = message
                    Case 2
                        ToolStripStatusLabel1.Text = message
                    Case 4
                        Try
                            BS = New BindingSource
                            BS.DataSource = DS.Tables(0)
                            DataGridView1.AutoGenerateColumns = False
                            DataGridView1.DataSource = BS
                            recordcount = DS.Tables(0).Rows.Count
                            CM = CType(BindingContext(BS), CurrencyManager)
                        Catch ex As Exception
                            message = ex.Message
                        End Try

                    Case 5
                        ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                    Case 6
                        ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                End Select
            Catch ex As Exception

            End Try
        End If

    End Sub

    Private Sub FormHistoryDetail_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'MessageBox.Show(Me.DialogResult)
        If NeedRefresh Then
            Me.DialogResult = DialogResult.OK
        End If


    End Sub


    Private Sub FormHistoryDetail_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ToolStripComboBox1.SelectedIndex = 0
    End Sub

    

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        If myUser = "" Then
            Exit Sub
        End If
        Dim myQueryWorksheetList As New List(Of QueryWorksheet)
        Me.ToolStripStatusLabel1.Text = ""
        Me.ToolStripStatusLabel2.Text = ""

        Dim mymessage As String = String.Empty

        Dim priceChangeId As String = String.Empty

        'Dim sqlstr As String
        'Dim sqlstr1 As String = String.Empty
        'sqlstr = "select * from  sp_gethistorydetail('" & myUser & "','" & mycriteria & "') as tb(cmmf bigint,vendorcode bigint,purchorg integer,plant integer,submitdate date, validon date,price numeric,pricingunit integer,comment text,creator character varying,materialdesc character varying,vendorname character varying,pricechangehdid bigint,creatorname character varying,validator character varying,cc character varying,pricetype character varying,negotiateddate date,changeby text,  statusname text)"
        'sqlstr = "select * from  sp_gethistorydetail('" & myUser & "','" & mycriteria & "') as tb(cmmf bigint,vendorcode bigint,purchorg integer,plant integer,submitdate date, validon date,price numeric,pricingunit integer,comment text,creator character varying,materialdesc character varying,vendorname character varying,pricechangehdid bigint,creatorname character varying,validator character varying,statusname text)"

        Dim DirectoryBrowser As FolderBrowserDialog = New FolderBrowserDialog
        DirectoryBrowser.Description = "Which directory do you want to use?"
        If (DirectoryBrowser.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = DirectoryBrowser.SelectedPath 'Application.StartupPath & "\PrintOut"
            Dim reportname = "PriceChange" & "-" & priceChangeId '& GetCompanyName()
            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

            Dim myqueryworksheet = New QueryWorksheet With {.DataSheet = 1,
                                                            .SheetName = "PriceChange",
                                                            .Sqlstr = sqlstr
                                                            }
            myQueryWorksheetList.Add(myqueryworksheet)


            'Dim myreport As New ExportToExcelFile(Me, sqlstr, filename, reportname, mycallback, PivotCallback)
            Dim myreport As New ExportToExcelFile(Me, myQueryWorksheetList, filename, reportname, mycallback, PivotCallback)

            myreport.Run(Me, e)

        End If
    End Sub

    Private Sub FormattingReport()
        'Throw New NotImplementedException
    End Sub

    Private Sub PivotTable()
        'Throw New NotImplementedException
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        If MessageBox.Show("Delete selected record?", "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Try
                If DataGridView1.SelectedRows.Count = 0 Then
                    BS.RemoveAt(CM.Position)
                Else
                    For Each a As DataGridViewRow In DataGridView1.SelectedRows
                        BS.RemoveAt(a.Index)
                    Next
                End If
                UpdateRecord()
            Catch ex As Exception
            End Try
        End If
    End Sub

    Public Sub UpdateRecord()

        Me.Validate()
        BS.EndEdit()
        Dim ds2 = DS.GetChanges()
        If Not IsNothing(ds2) Then
            Dim mymessage As String = String.Empty
            Dim ra As Integer
            Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
            If DbAdapter1.PriceChangeDTLTx(Me, mye) Then
                'delete the modfied row for Merged
                Dim modifiedRows = From row In ds2.Tables(0)
                   Where row.RowState = DataRowState.Added
                For Each row In modifiedRows.ToArray
                    row.Delete()
                Next
            Else
                MessageBox.Show(mye.message)
                Exit Sub
            End If
            DS.Merge(ds2)
            DS.AcceptChanges()
            Me.NeedRefresh = True
            MessageBox.Show("Saved.")
        End If


    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        myPeriod = New UCPeriodRange
        Dim PeriodHost = New ToolStripControlHost(myPeriod)
        ToolStrip1.Items.Add(PeriodHost)

    End Sub

    Private Sub StatusStrip1_ItemClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles StatusStrip1.ItemClicked

    End Sub

    Private Sub ToolStripComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox1.SelectedIndexChanged
        If Not IsNothing(ToolStripComboBox1.SelectedItem) Then
            Select Case ToolStripComboBox1.SelectedIndex
                Case 0, 1
                    MyPeriod.CheckBox1.Enabled = False
                    MyPeriod.DateTimePicker1.Enabled = False
                    MyPeriod.DateTimePicker2.Enabled = False

                Case Else
                    MyPeriod.CheckBox1.Enabled = True
                    MyPeriod.DateTimePicker1.Enabled = True
                    MyPeriod.DateTimePicker2.Enabled = True

            End Select
            'MessageBox.Show("selected")
        End If
    End Sub
End Class