Imports PriceCMMFExt.SharedClass
Imports System.Threading
Imports System.Text
Imports PriceCMMFExt.PublicClass
Public Class FormPriceChange2
    Private Enum TXType
        Task = 1
        History = 2
    End Enum

    Private Enum TaskStatus
        StatusDraft = 1
        StatusNew = 2
        StatusRejected = 3
        StatusReSubmit = 4
        StatusValidated = 5
        StatusCancelled = 6
        StatusCompleted = 7
    End Enum
    Dim myTXType As TXType
    Dim bsHeader As BindingSource
    Dim bsdetail As New BindingSource
    Dim bsdetailtmp As New BindingSource
    Dim comboBS As New BindingSource
    Dim comboRS As New BindingSource
    Dim DS As DataSet
    Dim DS2 As DataSet
    Dim DSDetail As DataSet
    Dim myrowhd As DataRowView
    Dim myrowdtl As DataRowView
    Dim fieldtohelp As String
    Dim validator1id As String
    Dim validator2id As String
    Dim validator3id As String
    Dim myfilename As String
    Dim ImportDetail As Boolean = False

    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByRef message As String)
    Dim myThread As New System.Threading.Thread(AddressOf doimport)
    Dim creator As String
    Dim IsNewRecord As Boolean
    Dim recordcount As Integer

    Dim sb As StringBuilder

    Public Property myUser As String
    Dim MyDS As DataSet

    Private Sub AddRecordToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddRecordToolStripMenuItem.Click
        myrowdtl = bsdetail.AddNew()
        Dim dr = myrowdtl.Row
        dr.Item("pricechangehdid") = myrowhd.Item("pricechangehdid")
        Dim myform = New FormInputPriceChange(bsdetail)

        'DS.Tables(4).Rows.Add(dr)
        'bs.AddNew()

        'If Not myform.ShowDialog = DialogResult.OK Then
        ' bs.CancelEdit()
        ' Else
        ' bs.EndEdit()
        ' End If

    End Sub
    Private Sub AddRecordWithPopUpToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddRecordWithPopUpToolStripMenuItem.Click
        Dim myform = New FormInputPriceChange(bsdetail)
        myrowdtl = bsdetail.AddNew()
        'bs.AddNew()

        If Not myform.ShowDialog = DialogResult.OK Then
            bsdetail.CancelEdit()
        Else
            bsdetail.EndEdit()
        End If
    End Sub
    Private Sub DataGridView1_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DataGridView1.CellBeginEdit
        fieldtohelp = DataGridView1.Columns(e.ColumnIndex).HeaderText
    End Sub

    Private Sub DataGridView1_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        If e.ColumnIndex = 8 Then
            Dim deltasap = 0.0
            If Not (IsDBNull(DataGridView1.Rows(e.RowIndex).Cells("price").Value) And (IsDBNull(DataGridView1.Rows(e.RowIndex).Cells("pricingunit").Value))) Then
                deltasap = ((DataGridView1.Rows(e.RowIndex).Cells("price").Value / DataGridView1.Rows(e.RowIndex).Cells("pricingunit").Value) - (DataGridView1.Rows(e.RowIndex).Cells("amount").Value / DataGridView1.Rows(e.RowIndex).Cells("perunit").Value)) / (DataGridView1.Rows(e.RowIndex).Cells("amount").Value / DataGridView1.Rows(e.RowIndex).Cells("perunit").Value)
            End If

            Dim deltastd = 0.0
            If Not (IsDBNull(DataGridView1.Rows(e.RowIndex).Cells("planprice1").Value) And (IsDBNull(DataGridView1.Rows(e.RowIndex).Cells("per").Value))) Then
                deltastd = ((DataGridView1.Rows(e.RowIndex).Cells("price").Value / DataGridView1.Rows(e.RowIndex).Cells("pricingunit").Value) - (DataGridView1.Rows(e.RowIndex).Cells("planprice1").Value / DataGridView1.Rows(e.RowIndex).Cells("per").Value)) / (DataGridView1.Rows(e.RowIndex).Cells("planprice1").Value / DataGridView1.Rows(e.RowIndex).Cells("per").Value)
            End If

            DataGridView1.Rows(e.RowIndex).Cells("deltasap").Value = deltasap
            DataGridView1.Rows(e.RowIndex).Cells("deltastd").Value = deltastd
            DataGridView1.Rows(e.RowIndex).Cells("alert").Value = IIf(deltasap > 0.05 Or deltastd > 0.05 Or deltasap < -0.05 Or deltasap < -0.05, ">5% or <-5% ", "") & IIf(DataGridView1.Rows(e.RowIndex).Cells("validon").Value = DataGridView1.Rows(e.RowIndex).Cells("validfrom").Value, "[Invalid date]", "")
        End If
        'MessageBox.Show("EndEdit " & e.ColumnIndex)
    End Sub

    Private Sub DataGridView1_EditingControlShowing(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles DataGridView1.EditingControlShowing
        Dim tb As DataGridViewTextBoxEditingControl = DirectCast(e.Control, DataGridViewTextBoxEditingControl)
        RemoveHandler (tb.KeyDown), AddressOf datagridviewTextBox_Keypdown
        AddHandler (tb.KeyDown), AddressOf datagridviewTextBox_Keypdown
    End Sub

    Private Sub datagridviewTextBox_Keypdown(ByVal sender As Object, ByVal e As KeyEventArgs)
        If e.KeyValue = 112 Then 'F1 
            MessageBox.Show("Help " & fieldtohelp)

        End If
    End Sub

    Private Sub DeleteRecordToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteRecordToolStripMenuItem.Click
        Dim myform = New FormInputPriceChange(bsdetail)
        If Not myform.ShowDialog = DialogResult.OK Then
            bsdetail.CancelEdit()
        Else
            bsdetail.EndEdit()
        End If
    End Sub

    Private Sub FormPriceChange_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Me.DialogResult = DialogResult.OK
        'bsHeader.CancelEdit()
        'bsdetail.CancelEdit()
    End Sub


    Private Sub FormPriceChange_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        loaddata()

    End Sub

    Private Sub DeleteRecordToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteRecordToolStripMenuItem1.Click
        If Not IsNothing(bsdetail.Current) Then
            If MessageBox.Show("Delete selected Record(s)?", "Question", System.Windows.Forms.MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.OK Then
                For Each dsrow In DataGridView1.SelectedRows
                    bsdetail.RemoveAt(CType(dsrow, DataGridViewRow).Index)
                Next
                'bsdetail.EndEdit()
                'ProgressReport(1, "Done!Record Found " & bsdetail.Count)
            End If

            'update records
            Dim mydrv As DataRowView = bsHeader.Current
            If Not mydrv.Row.Item("pricechangehdid") < 1 Then

                '' Dim ds2 = Me.MyDS.Tables(0).GetChanges

                Dim ds2 As New DataTable
                If myTXType = TXType.Task Then
                    ds2 = Me.DS.Tables(4).GetChanges
                ElseIf myTXType = TXType.History Then
                    ds2 = Me.MyDS.Tables(0).GetChanges
                End If


                'Dim ds2 = Me.DS.Tables(4).GetChanges
                If Not IsNothing(ds2) Then
                    Dim myds2 As New DataSet
                    myds2.Tables.Add(ds2)
                    Dim mymessage As String = String.Empty
                    Dim ra As Integer
                    Dim mye As New ContentBaseEventArgs(myds2, True, mymessage, ra, True)
                    If DbAdapter1.PriceChangeDTLTx(Me, mye) Then
                        'delete the modfied row for Merged
                        Dim modifiedRows = From row In myds2.Tables(0)
                           Where row.RowState = DataRowState.Added
                        For Each row In modifiedRows.ToArray
                            row.Delete()
                        Next
                    Else
                        MessageBox.Show(mye.message)
                        Exit Sub
                    End If
                    DS.Merge(myds2)
                    DS.AcceptChanges()

                    MessageBox.Show("Saved.")
                End If
            End If

                ProgressReport(1, "Done!Record Found " & bsdetail.Count)
            Else
                MessageBox.Show("No record to delete.")
            End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub


    Public Sub New(ByVal bsheader As BindingSource, ByVal DS As DataSet, ByVal DS2 As DataSet, Optional ByVal isnewrecord As Boolean = True)

        ' This call is required by the designer. New Task
        InitializeComponent()
        myTXType = TXType.Task
        Me.bsHeader = bsheader
        myrowhd = bsheader.Current
        'creator = myrowhd.Row.Item("creator").ToString.Replace("\", "")
        creator = myrowhd.Row.Item("creator").ToString
        Me.DS = DS
        Me.DS2 = DS2

        myUser = HelperClass1.UserId.ToLower
        sb = New StringBuilder
        'sb.Append(String.Format("with my as (select distinct pricechangehdid from sp_getmytasks('{0}'::text,false) as tb(pricechangehdid bigint,creator character varying,creatorname character varying,validator1 character varying,validator1name character varying,validator2 character varying,validator2name character varying,pricetype character varying,description text,submitdate date,negotiateddate date,attachment text,status integer,statusname text,actiondate date,actionby character varying,reasonid integer))" &
        '         " , pl as (select cmmf, max(validfrom) as validfrom,vendorcode from pricelist  group by cmmf,vendorcode)" &
        '         " , std as (select cmmf, max(validfrom) as validfrom from standardcostad group by cmmf)" &
        '         " select dt.*,c.commercialref,r.rangedesc,r.range,v.vendorname::character varying,v.shortname::character varying,materialdesc::character varying, pr.amount::numeric,pr.perunit::numeric,ad.planprice1,ad.per," &
        '         " (getdelta(dt.price,dt.pricingunit,ad.planprice1,ad.per)) as deltastd, " &
        '         " (getdelta(dt.price,dt.pricingunit,pr.amount::numeric,pr.perunit::numeric) ) as deltasap," &
        '          " getalert(dt.price,dt.pricingunit,pr.amount::numeric,pr.perunit::numeric,ad.planprice1,ad.per) as alert,pr.validfrom " &
        '         "   from pricechangedtl dt  " &
        '         " left join cmmf c on c.cmmf = dt.cmmf  inner join my on my.pricechangehdid = dt.pricechangehdid left join vendor v on v.vendorcode = dt.vendorcode left join range r on r.rangeid = c.rangeid " &
        '         " left join std on std.cmmf = c.cmmf" &
        '         " left join standardcostad ad on ad.cmmf = std.cmmf and ad.validfrom = std.validfrom" &
        '         " left join pl on pl.cmmf = c.cmmf and pl.vendorcode = v.vendorcode" &
        '         " left join pricelistscalelatest pls on pls.cmmf = c.cmmf and pls.vendorcode = v.vendorcode and pls.plant = dt.plant" &
        '         " left join pricelistscale pr on pr.cmmf = pls.cmmf and pr.validfrom = pls.validfrom and pr.vendorcode = v.vendorcode and pr.plant = dt.plant where dt.pricechangehdid = {1};", myUser, myrowhd.Item("pricechangehdid")))

        'sb.Append(String.Format("with " &
        '                        " pls as ( SELECT foo.cmmf, foo.vendorcode, foo.plant, max(foo.validfrom) AS validfrom  " &
        '                        "   FROM ( select pd.cmmf,pd.vendorcode,pp.plant,pl.validfrom from pricechangedtl pd " &
        '                        " left join pricelist pl on pl.cmmf = pd.cmmf" &
        '                        " LEFT JOIN priceplantscale pp ON pp.pricelistid = pl.pricelistid " &
        '                        " WHERE pl.scaleqty = pp.scale::double precision and pd.pricechangehdid = {0}) foo " &
        '                        " GROUP BY foo.cmmf, foo.vendorcode, foo.plant ), " &
        '                        " pr as (SELECT pl.pricelistid, pl.cmmf, pl.scaleqty, pp.amount, pp.perunit, pl.validfrom, pl.validto, pl.vendorcode, pl.currency, pp.plant, pp.id  from pricechangedtl pd" &
        '                        " left join  pricelist pl on pl.cmmf = pd.cmmf  LEFT JOIN priceplantscale pp ON pp.pricelistid = pl.pricelistid " &
        '                        " WHERE pl.scaleqty = pp.scale::double precision and pd.pricechangehdid = {0})," &
        '        "  std as (select cmmf, max(validfrom) as validfrom from standardcostad group by cmmf)" &
        '        " select dt.*,c.commercialref,r.rangedesc,r.range,v.vendorname::character varying,v.shortname::character varying,materialdesc::character varying, pr.amount::numeric,pr.perunit::numeric,ad.planprice1,ad.per," &
        '        " (getdelta(dt.price,dt.pricingunit,ad.planprice1,ad.per)) as deltastd, " &
        '        " (getdelta(dt.price,dt.pricingunit,pr.amount::numeric,pr.perunit::numeric) ) as deltasap," &
        '         " getalert(dt.price,dt.pricingunit,pr.amount::numeric,pr.perunit::numeric,ad.planprice1,ad.per) as alert,pr.validfrom " &
        '        "   from pricechangedtl dt  " &
        '        " left join cmmf c on c.cmmf = dt.cmmf  left join vendor v on v.vendorcode = dt.vendorcode left join range r on r.rangeid = c.rangeid " &
        '        " left join std on std.cmmf = c.cmmf" &
        '        " left join standardcostad ad on ad.cmmf = std.cmmf and ad.validfrom = std.validfrom" &
        '        " left join  pls on pls.cmmf = c.cmmf and pls.vendorcode = v.vendorcode and pls.plant = dt.plant" &
        '        " left join pr on pr.cmmf = pls.cmmf and pr.validfrom = pls.validfrom and pr.vendorcode = v.vendorcode and pr.plant = dt.plant where dt.pricechangehdid = {0};", myrowhd.Item("pricechangehdid")))
        sb.Append("select 1;")
        'sb.Append("with dup as (select commercialref from cmmf where length(commercialref) > 1" &
        '          " group by commercialref" &
        '          " having count(commercialref) = 1)" &
        '          " select c.commercialref,c.cmmf from cmmf c inner join dup on dup.commercialref = c.commercialref;")
        sb.Append("select 2;")



        'ClearBindingObject()

        ''BSdetail filter
        'Me.IsNewRecord = isnewrecord
        ''Me.comboBS.DataSource = DS.Tables(6)
        'Me.comboBS.DataSource = DS.Tables("PriceType")
        'Me.comboRS.DataSource = DS.Tables("Reason")
        'bsdetail.DataSource = DS.Tables(4)
        'bsdetailtmp.DataSource = DS.Tables(4) 'temp needed for create new record using import method

        'bsdetail.Filter = "pricechangehdid=" & myrowhd.Row.Item("pricechangehdid").ToString
        'BindingObject()
        '' Add any initialization after the InitializeComponent() call.
        ''DataGridView1.DataSource = bsdetail
    End Sub

    Public Sub New(ByVal DS As DataSet, ByVal bshistory As BindingSource)
        'History
        myTXType = TXType.History
        InitializeComponent()
        Me.DS = DS
        Me.bsHeader = bshistory
        myrowhd = bsHeader.Current
        creator = myrowhd.Row.Item("creator").ToString
        sb = New StringBuilder
        myUser = HelperClass1.UserId.ToLower
        'sb.Append(String.Format("with my as (select distinct pricechangehdid,statusname from sp_getmytasks('{0}'::text,true) as tb(pricechangehdid bigint,creator character varying,creatorname character varying,validator1 character varying,validator1name character varying,validator2 character varying,validator2name character varying,pricetype character varying,description text,submitdate date,negotiateddate date,attachment text,status integer,statusname text,actiondate date,actionby character varying,reasonid integer))" &
        '         " , std as (select cmmf, max(validfrom) as validfrom from standardcostad group by cmmf)" &
        '         " select dt.*,ad.planprice1,c.commercialref,r.rangedesc,r.range,v.vendorname::character varying,v.shortname::character varying,materialdesc::character varying, " &
        '         " (getpriceinfo(my.statusname,dt.cmmf,dt.vendorcode,dt.validon,dt.price,dt.pricingunit,ad.planprice1,ad.per)).* " &
        '         "   from pricechangedtl dt  " &
        '         " left join cmmf c on c.cmmf = dt.cmmf  inner join my on my.pricechangehdid = dt.pricechangehdid left join vendor v on v.vendorcode = dt.vendorcode left join range r on r.rangeid = c.rangeid " &
        '         " left join std on std.cmmf = c.cmmf" &
        '         " left join standardcostad ad on ad.cmmf = std.cmmf and ad.validfrom = std.validfrom where dt.pricechangehdid = {1};", myUser, myrowhd.Item("pricechangehdid")))
        'sb.Append(String.Format("with my as (select distinct pricechangehdid,statusname from sp_getmytasks('{0}'::text,true) as tb(pricechangehdid bigint,creator character varying,creatorname character varying,validator1 character varying,validator1name character varying,validator2 character varying,validator2name character varying,pricetype character varying,description text,submitdate date,negotiateddate date,attachment text,status integer,statusname text,actiondate date,actionby character varying,reasonid integer))" &
        '        " , std as (select cmmf, max(validfrom) as validfrom from standardcostad group by cmmf)" &
        '        " select dt.*,ad.planprice1,c.commercialref,r.rangedesc,r.range,v.vendorname::character varying,v.shortname::character varying,materialdesc::character varying, " &
        '        " (getpriceplantinfo(my.statusname,dt.cmmf,dt.vendorcode,dt.plant,dt.validon,dt.price,dt.pricingunit,ad.planprice1,ad.per)).* ,doc.getvendorcurr(dt.vendorcode,dt.validon) as crcy " &
        '        "   from pricechangedtl dt  " &
        '        " left join cmmf c on c.cmmf = dt.cmmf  inner join my on my.pricechangehdid = dt.pricechangehdid left join vendor v on v.vendorcode = dt.vendorcode left join range r on r.rangeid = c.rangeid " &
        '        " left join std on std.cmmf = c.cmmf" &
        '        " left join standardcostad ad on ad.cmmf = std.cmmf and ad.validfrom = std.validfrom where dt.pricechangehdid = {1};", myUser, myrowhd.Item("pricechangehdid")))
        sb.Append(String.Format("with my as (select distinct pricechangehdid,statusname from sp_getmytasks4('{0}'::text,true) as tb(pricechangehdid bigint,creator character varying,creatorname character varying,validator1 character varying,validator1name character varying,validator2 character varying,validator2name character varying,validator3 character varying,validator3name character varying,pricetype character varying,description text,submitdate date,negotiateddate date,attachment text,status integer,statusname text,actiondate date,actionby character varying,reasonid integer))" &
                " , std as (select cmmf, max(validfrom) as validfrom from standardcostad group by cmmf)" &
                " select dt.*,ad.planprice1,c.commercialref,r.rangedesc,r.range,v.vendorname::character varying,v.shortname::character varying,materialdesc::character varying, " &
                " (getpriceplantinfo(my.statusname,dt.cmmf,dt.vendorcode,dt.plant,dt.validon,dt.price,dt.pricingunit,ad.planprice1,ad.per)).* ,doc.getvendorcurr(dt.vendorcode,dt.validon) as crcy " &
                "   from pricechangedtl dt  " &
                " left join cmmf c on c.cmmf = dt.cmmf  inner join my on my.pricechangehdid = dt.pricechangehdid left join vendor v on v.vendorcode = dt.vendorcode left join range r on r.rangeid = c.rangeid " &
                " left join std on std.cmmf = c.cmmf" &
                " left join standardcostad ad on ad.cmmf = std.cmmf and ad.validfrom = std.validfrom where dt.pricechangehdid = {1};", myUser, myrowhd.Item("pricechangehdid")))
        sb.Append("select 1;")
    End Sub

    Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

    Private Overloads Function validate() As Boolean
        MyBase.Validate()
        Dim myret As Boolean = True
        If Not validator1id <> "" Then
            TextBox3.Focus()
            ErrorProvider1.SetError(TextBox3, "Validator not found.")
            myret = False
        Else
            ErrorProvider1.SetError(TextBox3, "")
        End If

        'If Not validator2id <> "" Then
        '    TextBox4.Focus()
        '    ErrorProvider1.SetError(TextBox4, "Validator not found.")
        '    myret = False
        'Else
        '    ErrorProvider1.SetError(TextBox4, "")
        'End If

        'If TextBox3.Text.Contains("PD -") Then
        '    If ComboBox1.Text = "FOB" Then
        '        ErrorProvider1.SetError(TextBox3, "Director only validate STD Price Change.")
        '        myret = False
        '    Else
        '        ErrorProvider1.SetError(TextBox3, "")
        '    End If
        'End If
        'If TextBox4.Text.Contains("PD -") Then
        '    If ComboBox1.Text = "FOB" Then
        '        ErrorProvider1.SetError(TextBox4, "Director only validate  STD Price Change.")
        '        myret = False
        '    Else
        '        ErrorProvider1.SetError(TextBox4, "")
        '    End If
        'End If

        'check datagridview. at least has one record
        If IsNothing(ComboBox1.SelectedValue) Then
            ErrorProvider1.SetError(ComboBox1, "Select from the list.")
            myret = False
        Else
            ErrorProvider1.SetError(ComboBox1, "")
        End If

        If IsNothing(bsdetail.Current) Then
            MessageBox.Show("No details to submit.")
            myret = False
        End If

        'If IsNothing(ComboBox2.SelectedValue) Then
        '    ErrorProvider1.SetError(ComboBox2, "Select from the list.")
        '    myret = False
        'Else
        '    ErrorProvider1.SetError(ComboBox2, "")
        'End If

        'Validate Detail For Fob
        'If ComboBox1.Text = "FOB" Then
        '    For Each drv As DataRowView In bsdetail.List

        '        Dim mysb As New StringBuilder
        '        If Not (drv.Item("pricingunit") = 100 Or drv.Item("pricingunit") = 1000) Then
        '            myret = False
        '            mysb.Append("the unit not equal to ""100""or""1000"".")
        '        End If

        '        If drv.Item("validon") < drv.Item("validfrom") Then
        '            myret = False
        '            If mysb.Length > 0 Then
        '                mysb.Append(",")
        '            End If
        '            mysb.Append("The new valid date should be greater than the SAP Price Valid From")
        '        End If
        '        If IsDBNull(drv.Item("amount")) Then
        '            myret = False
        '            If mysb.Length > 0 Then
        '                mysb.Append(",")
        '            End If
        '            mysb.Append("Missing current price. Please check whether the item is created or not")
        '        End If
        '        drv.Row.RowError = mysb.ToString
        '    Next
        'End If

        For Each drv As DataRowView In bsdetail.List

            Dim mysb As New StringBuilder
            If Not (drv.Item("pricingunit") = 100 Or drv.Item("pricingunit") = 1000) Then
                myret = False
                mysb.Append("the unit not equal to ""100""or""1000"".")
            End If
            If ComboBox1.Text = "FOB" Then

                If IsDBNull(drv.Item("amount")) Then
                    myret = False
                    If mysb.Length > 0 Then
                        mysb.Append(",")
                    End If
                    mysb.Append("Missing current price. Please check whether the item is created or not")
                End If
                If Not IsDBNull(drv.Item("validfrom")) Then
                    If drv.Item("validon") < drv.Item("validfrom") Then
                        myret = False
                        If mysb.Length > 0 Then
                            mysb.Append(",")
                        End If
                        mysb.Append("The new valid date should be greater than the SAP Price Valid From")
                    End If
                End If
            End If

            drv.Row.RowError = mysb.ToString
        Next

        Return myret
    End Function
    Private Sub ToolStripButton2_Clickori(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If Me.validate Then
            bsHeader.EndEdit()
            bsdetail.EndEdit()
            If MessageBox.Show("Submit the records?", "Submit Record", MessageBoxButtons.OKCancel) = DialogResult.OK Then
                Dim mymessage As String = String.Empty
                If Not IsDBNull(myrowhd.Row.Item("status")) Then
                    If myrowhd.Row.Item("status") = TaskStatus.StatusRejected Then '= 3
                        myrowhd.Row.Item("status") = TaskStatus.StatusReSubmit  '4
                    End If
                    'modified New

                End If

                Dim myitem As String = String.Empty
                If Not DateTimePicker2.Checked Then
                    myrowhd.Row.Item("negotiateddate") = DBNull.Value
                End If
                For Each item In ListBox1.Items
                    myitem = myitem & IIf(myitem = "", "", ",") & item.ToString
                Next
                'No need importdetail checking
                If ImportDetail Then

                    myrowhd.Row.Item("attachment") = myitem

                    'Dim ds2 As New DataSet 'Tables(0) -> header
                    'ds2 = DS.GetChanges

                    'Dim ds3 As New DataSet 'Tables(0) ->details
                    'bsdetail.EndEdit()

                    'Dim mynewds As New DataSet

                    'ds2.Tables.Add(DSDetail.Tables(0))

                    'mynewds.Tables.Add(ds2.Tables(0))
                    'mynewds.Tables.Add(ds2.Tables(1))
                    'mynewds.Tables.Add(ds2.Tables(2))
                    'mynewds.Tables.Add(ds2.Tables(3))
                    'mynewds.Tables.Add(DSDetail.Tables(0))

                    'If Not IsNothing(mynewds) Then
                    '    mynewds.Tables(0).Rows(0).Item("attachment") = myitem
                    '    'ds2.Tables(0).Rows(0).Item("status") = 4
                    '    'Dim mymessage As String = String.Empty
                    '    Dim ra As Integer
                    '    'reset sequence number
                    '    Dim mye As New ContentBaseEventArgs(mynewds, True, mymessage, ra, True)
                    '    If Not DbAdapter1.PriceChangeTx(Me, mye) Then
                    '        ProgressReport(1, mye.message)
                    '    End If
                    '    'copy Attachment

                    '    Me.DialogResult = DialogResult.OK
                    '    Me.Close()
                    'End If
                    'Me.Close()

                    '   
                    '    *********************
                    If Not DbAdapter1.copyToPriceChange(creator, myrowhd.Row, IsNewRecord, mymessage) Then
                        ProgressReport(1, mymessage)
                    Else
                        Me.DialogResult = DialogResult.OK
                        'copy data from pricechangedtltmp to pricechangehd
                        Me.Close()
                    End If
                    '    ********************
                Else
                    Dim ds2 As New DataSet
                    ds2 = DS.GetChanges

                    If Not IsNothing(ds2) Then
                        ds2.Tables(0).Rows(0).Item("attachment") = myitem
                        'ds2.Tables(0).Rows(0).Item("status") = 4
                        'Dim mymessage As String = String.Empty
                        Dim ra As Integer
                        'reset sequence number
                        Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)

                        If Not DbAdapter1.PriceChangeTx(Me, mye) Then
                            ProgressReport(1, mye.message)
                        End If
                        'copy Attachment

                        Me.DialogResult = DialogResult.OK
                        Me.Close()
                    End If
                    Me.Close()
                End If
            End If

        End If



    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        If Me.validate Then
            Dim vendordict As New Dictionary(Of Long, String)
            Dim vendorlist As New StringBuilder
            If bsdetail.Count > 0 Then
                For Each dr As DataRowView In bsdetail.List
                    If Not vendordict.ContainsKey(dr.Item("vendorcode")) Then
                        vendordict.Add(dr.Item("vendorcode"), "")
                    End If
                Next
                For Each vd As KeyValuePair(Of Long, String) In vendordict
                    If vendorlist.Length <> 0 Then
                        vendorlist.Append(",")
                    End If
                    vendorlist.Append(vd.Key)
                Next
                Dim drv As DataRowView = bsHeader.Current
                If Not DbAdapter1.CanFindUserVendor(drv.Row.Item("validator1"), vendorlist) Then
                    If MessageBox.Show("Validator did not match with the Short Name, Continue?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1) = DialogResult.No Then
                        Exit Sub
                    End If
                End If

            End If


            bsHeader.EndEdit()
            bsdetail.EndEdit()


            If MessageBox.Show("Submit the records?", "Submit Record", MessageBoxButtons.OKCancel) = DialogResult.OK Then
                Dim mymessage As String = String.Empty
                If Not IsDBNull(myrowhd.Row.Item("status")) Then
                    If myrowhd.Row.Item("status") = TaskStatus.StatusRejected Then '3 
                        myrowhd.Row.Item("status") = TaskStatus.StatusReSubmit '4
                    End If
                    'modified New

                End If

                Dim myitem As String = String.Empty
                If Not DateTimePicker2.Checked Then
                    myrowhd.Row.Item("negotiateddate") = DBNull.Value
                End If
                For Each item In ListBox1.Items
                    myitem = myitem & IIf(myitem = "", "", ",") & item.ToString
                Next
                'No need importdetail checking
                If ImportDetail Then
                    'getchanges from bsdetail.datasource (dsdetail.table(0))
                    'create record in ds.tables(4)
                    'myrowhd.Row.Item("attachment") = myitem
                    Dim ds3 = DSDetail.GetChanges
                    ' If Not IsNothing(ds3) Then
                    For Each dr As DataRow In DSDetail.Tables(0).Rows
                        'DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "pricechangehdid").SourceVersion = DataRowVersion.Original
                        'DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                        'DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "cmmf").SourceVersion = DataRowVersion.Current
                        'DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "purchorg").SourceVersion = DataRowVersion.Current
                        'DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "plant").SourceVersion = DataRowVersion.Current
                        'DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "validon").SourceVersion = DataRowVersion.Current
                        'DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "price").SourceVersion = DataRowVersion.Current
                        'DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "pricingunit").SourceVersion = DataRowVersion.Current
                        'DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "comment").SourceVersion = DataRowVersion.Current
                        If dr.RowState <> DataRowState.Deleted Then


                            Dim mydrtmp As DataRowView = bsdetailtmp.AddNew
                            mydrtmp.Row.Item("pricechangehdid") = myrowhd.Row.Item("pricechangehdid") 'dr.item("pricechangehdid")
                            mydrtmp.Row.Item("vendorcode") = dr.Item("vendorcode")
                            mydrtmp.Row.Item("cmmf") = dr.Item("cmmf")
                            mydrtmp.Row.Item("purchorg") = dr.Item("purchorg")
                            mydrtmp.Row.Item("plant") = dr.Item("plant")
                            mydrtmp.Row.Item("validon") = dr.Item("validon")
                            mydrtmp.Row.Item("price") = dr.Item("price")
                            mydrtmp.Row.Item("pricingunit") = dr.Item("pricingunit")
                            mydrtmp.Row.Item("comment") = dr.Item("comment")
                            DS.Tables(4).Rows.Add(mydrtmp.Row)
                            'MyDS.Tables(0).Rows.Add(mydrtmp.Row)
                            'DS.Tables(10).Rows.Add(mydrtmp.Row)
                        End If
                    Next
                    ' End If


                    'Dim ds2 As New DataSet 'Tables(0) -> header
                    'ds2 = DS.GetChanges

                    'Dim ds3 As New DataSet 'Tables(0) ->details
                    'bsdetail.EndEdit()

                    'Dim mynewds As New DataSet

                    'ds2.Tables.Add(DSDetail.Tables(0))

                    'mynewds.Tables.Add(ds2.Tables(0))
                    'mynewds.Tables.Add(ds2.Tables(1))
                    'mynewds.Tables.Add(ds2.Tables(2))
                    'mynewds.Tables.Add(ds2.Tables(3))
                    'mynewds.Tables.Add(DSDetail.Tables(0))

                    'If Not IsNothing(mynewds) Then
                    '    mynewds.Tables(0).Rows(0).Item("attachment") = myitem
                    '    'ds2.Tables(0).Rows(0).Item("status") = 4
                    '    'Dim mymessage As String = String.Empty
                    '    Dim ra As Integer
                    '    'reset sequence number
                    '    Dim mye As New ContentBaseEventArgs(mynewds, True, mymessage, ra, True)
                    '    If Not DbAdapter1.PriceChangeTx(Me, mye) Then
                    '        ProgressReport(1, mye.message)
                    '    End If
                    '    'copy Attachment

                    '    Me.DialogResult = DialogResult.OK
                    '    Me.Close()
                    'End If
                    'Me.Close()

                    '   
                    '    *********************
                    'If Not DbAdapter1.copyToPriceChange(creator, myrowhd.Row, IsNewRecord, mymessage) Then
                    '    ProgressReport(1, mymessage)
                    'Else
                    '    Me.DialogResult = DialogResult.OK
                    '    'copy data from pricechangedtltmp to pricechangehd
                    '    Me.Close()
                    'End If
                    '    ********************
                    'Else
                End If
                Dim ds2 As New DataSet
                'ds2 = DS.GetChanges
                ds2 = DS.GetChanges

                If Not IsNothing(ds2) Then
                    ds2.Tables(0).Rows(0).Item("attachment") = myitem
                    'ds2.Tables(0).Rows(0).Item("status") = 4
                    'Dim mymessage As String = String.Empty
                    Dim ra As Integer
                    'reset sequence number
                    Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)

                    If Not DbAdapter1.PriceChangeTx(Me, mye) Then
                        ProgressReport(1, mye.message)
                        Exit Sub
                    End If
                    'copy Attachment

                    Me.DialogResult = DialogResult.OK
                    Me.Close()
                    'End If
                    Me.Close()
                End If
            End If
        Else
            ProgressReport(1, "Error found. Please check.")
        End If



    End Sub



    Private Sub BindingObject()

        With ComboBox1
            .DisplayMember = "paramname"
            .ValueMember = "paramname"
            .DataSource = comboBS
        End With


        TextBox1.DataBindings.Add(New Binding("Text", bsHeader, "creatorname", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox2.DataBindings.Add(New Binding("Text", bsHeader, "description", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox3.DataBindings.Add(New Binding("Text", bsHeader, "validator1name", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox4.DataBindings.Add(New Binding("Text", bsHeader, "validator2name", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox5.DataBindings.Add(New Binding("Text", bsHeader, "validator3name", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox6.DataBindings.Add(New Binding("Text", bsHeader, "reasonname", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox7.DataBindings.Add(New Binding("Text", bsHeader, "specialproject", True, DataSourceUpdateMode.OnPropertyChanged))

        DateTimePicker1.DataBindings.Add(New Binding("value", bsHeader, "submitdate"))
        DateTimePicker2.DataBindings.Add(New Binding("Text", bsHeader, "negotiateddate"))

        'With ComboBox1

        '    .DisplayMember = "paramname"
        '    .ValueMember = "paramname"
        '    .DataSource = comboBS
        '    .DataBindings.Add(New Binding("SelectedValue", bsHeader, "pricetype"))

        'End With

        ComboBox1.DataBindings.Add(New Binding("SelectedValue", bsHeader, "pricetype", True, DataSourceUpdateMode.OnPropertyChanged))

        'With ComboBox2

        '    .DisplayMember = "reasonname"
        '    .ValueMember = "id"
        '    .DataSource = comboRS
        '    .DataBindings.Add(New Binding("SelectedValue", bsHeader, "reasonid"))

        'End With

        'If Not IsDBNull(myrowhd.Row("pricetype")) Then
        '    ComboBox1.Text = myrowhd.Row("pricetype")
        'End If

        If Not IsDBNull(myrowhd.Row.Item("validator1")) Then
            validator1id = myrowhd.Row.Item("validator1")
        End If
        If Not IsDBNull(myrowhd.Row.Item("validator2")) Then
            validator2id = myrowhd.Row.Item("validator2")
        End If
        If Not IsDBNull(myrowhd.Row.Item("validator3")) Then
            validator3id = myrowhd.Row.Item("validator3")
        End If
        Dim mycontent = myrowhd.Row.Item("attachment").ToString.Split(",")

        For Each Data As String In mycontent
            If Data <> "" Then
                ListBox1.Items.Add(Data)
            End If
        Next

        DataGridView1.AutoGenerateColumns = False
        DataGridView1.DataSource = bsdetail
        recordcount = DataGridView1.RowCount
        DataGridView1.Invalidate()

    End Sub

    Private Sub ClearBindingObject()
        TextBox1.DataBindings.Clear()
        TextBox2.DataBindings.Clear()
        TextBox3.DataBindings.Clear()
        TextBox4.DataBindings.Clear()
        TextBox5.DataBindings.Clear()
        TextBox6.DataBindings.Clear()
        TextBox7.DataBindings.Clear()
        DateTimePicker1.DataBindings.Clear()
        DateTimePicker2.DataBindings.Clear()
        ComboBox1.DataBindings.Clear()
        ListBox1.DataBindings.Clear()

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click, Button3.Click, Button1.Click
        Dim myobj = CType(sender, Button)
        Dim bs As New BindingSource
        bs.DataSource = DS.Tables(2)
        Dim myform As New FormGetValidator(bs)
        bs.Filter = ""
        If myobj.Name = "Button2" Then
            bs.Filter = "teamtitleshortname <> 'PO'"
        End If
        If myform.ShowDialog = DialogResult.OK Then
            Dim drv As DataRowView = bs.Current

            Select Case myobj.Name
                Case "Button2"
                    TextBox3.Text = drv.Row.Item("name")
                    validator1id = drv.Row.Item("userid")
                    myrowhd.Row.Item("validator1") = validator1id.ToLower
                    myrowhd.Row.Item("validator1name") = drv.Row.Item("name")
                Case "Button3"
                    TextBox4.Text = drv.Row.Item("name")
                    validator2id = drv.Row.Item("userid")
                    myrowhd.Row.Item("validator2") = validator2id.ToLower
                    myrowhd.Row.Item("validator2name") = drv.Row.Item("name")
                Case "Button1"
                    TextBox5.Text = drv.Row.Item("name")
                    validator3id = drv.Row.Item("userid")
                    myrowhd.Row.Item("validator3") = validator3id.ToLower
                    myrowhd.Row.Item("validator3name") = drv.Row.Item("name")
            End Select
        End If
    End Sub


    Private Sub AddAttachmentToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddAttachmentToolStripMenuItem.Click
        Dim FileDialog1 As New OpenFileDialog
        Dim destination As String = DS.Tables("AttachmentFolder").Rows(0).Item(0).ToString
        Dim targetFilename As String = ""
        If FileDialog1.ShowDialog = DialogResult.OK Then
            targetFilename = Replace(FileDialog1.SafeFileName, ",", "_")
            ListBox1.Items.Add(targetFilename)
            'copy to default folder
            Try
                IO.File.Copy(FileDialog1.FileName, destination & "\" & targetFilename, True)
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If
    End Sub

    Private Sub DeleteAttachmentToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteAttachmentToolStripMenuItem.Click
        If MessageBox.Show("Delete selected item?", "Delete Item") = DialogResult.OK Then
            If ListBox1.SelectedIndex > -1 Then
                ListBox1.Items.Remove(ListBox1.Items(ListBox1.SelectedIndex))
            End If

        End If
    End Sub

    Private Sub ToolStripButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton7.Click
        If Not myThread.IsAlive Then
            ImportDetail = True
            Dim FileDialog1 As New OpenFileDialog
            If FileDialog1.ShowDialog = DialogResult.OK Then
                Dim errmsg As String = String.Empty
                myfilename = FileDialog1.FileName

                bsdetail = New BindingSource
                bsdetail.DataSource = MyDS.Tables(0) 'DS.Tables(4)
                For i = 0 To bsdetail.Count - 1
                    bsdetail.RemoveCurrent()
                Next
                bsdetail.EndEdit()

                myThread = New Thread(AddressOf doimport)
                myThread.Start()

                'end thread

            End If
        Else
            MessageBox.Show("Please wait until the current process finished.")
        End If

    End Sub

    Public Sub ProgressReport(ByVal id As Integer, ByRef message As String)
        If Me.Label4.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    'TextBox2.Text = message
                    Me.ToolStripStatusLabel1.Text = message
                Case 2
                    'TextBox2.Text = message
                    Me.ToolStripStatusLabel1.Text = message
                Case 4
                    DataGridView1.Invalidate()
                Case 5

                    Me.ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                Case 6

                    Me.ToolStripProgressBar1.Style = ProgressBarStyle.Marquee

                Case 7
                    bsdetail = New BindingSource
                    bsdetail.DataSource = DSDetail.Tables(0)
                    DataGridView1.DataSource = bsdetail
                    DataGridView1.Invalidate()
                    recordcount = DSDetail.Tables(0).Rows.Count

                Case 8
                   

                    ClearBindingObject()

                    'BSdetail filter
                    Me.IsNewRecord = IsNewRecord
                    Me.comboBS.DataSource = DS.Tables("PriceType")

                    'Dim ReasonBS As New BindingSource
                    'ReasonBS.DataSource = DS.Tables("Reason")
                    'If myTXType = TXType.Task Then
                    '    ReasonBS.Filter = "isactive = true"
                    'End If
                    'Me.comboRS.DataSource = ReasonBS 'DS.Tables("Reason")




                    'bsdetail.Filter = "pricechangehdid=" & myrowhd.Row.Item("pricechangehdid").ToString
                    BindingObject()


                    With DataGridView1
                        .DataSource = bsdetail
                    End With
                    'Me.ComboBox3.DataSource = DS.Tables("PriceType")
                    'Me.ComboBox3.SelectedIndex = 0
                    If Me.ComboBox1.SelectedIndex = -1 Then
                        Me.ComboBox1.SelectedIndex = 0
                    End If
            End Select

        End If

    End Sub
    Private Sub doimport()

        ProgressReport(4, "Refresh Datagridview")
        ProgressReport(6, "Marguee")
        Dim sb As New StringBuilder
        Dim DuplicateCheck As String = String.Empty
        Dim errmsg As String = String.Empty
        If Not ImportMyTask(myfilename, bsdetail, Me, errmsg) Then
            ProgressReport(1, errmsg)
            Exit Sub
        End If

        ProgressReport(2, "Get Initial Data.")
        'Get ExistingData
        Dim mysqlstr = "select vendorcode , cmmf , validon , plant from pricechangedtl;"
        'Dim mymessage As String = String.Empty
        Dim DSLocal As New DataSet
        If DbAdapter1.TbgetDataSet(mysqlstr, DSLocal, errmsg) Then
            Try
                DSLocal.Tables(0).TableName = "PriceChangeDtl"
                Dim idx0(3) As DataColumn
                idx0(0) = DSLocal.Tables(0).Columns(0)
                idx0(1) = DSLocal.Tables(0).Columns(1)
                idx0(2) = DSLocal.Tables(0).Columns(2)
                idx0(3) = DSLocal.Tables(0).Columns(3)
                DSLocal.Tables(0).PrimaryKey = idx0

            Catch ex As Exception
                ProgressReport(1, "Loading Data. Error::" & ex.Message)
                ProgressReport(5, "Continuous")
                Exit Sub
            End Try
        Else
            ProgressReport(1, "Get Initial Data. Error::" & errmsg)
            ProgressReport(5, "Continuous")
            Exit Sub
        End If
        ProgressReport(1, "Get Initial Data.Done!")
        ProgressReport(5, "Continuous")


        Using objTFParser = New FileIO.TextFieldParser(Replace(myfilename, ".xlsx", ".txt"))
            Dim myrecord() As String
            Dim mylist As New List(Of String())
            With objTFParser
                .TextFieldType = FileIO.FieldType.Delimited
                .SetDelimiters(Chr(9))
                .HasFieldsEnclosedInQuotes = True
                Dim count As Long = 0

                Do Until .EndOfData
                    myrecord = .ReadFields
                    If count > 0 Then
                        mylist.Add(myrecord)
                    End If
                    count += 1
                Loop
                ProgressReport(2, "Build Record...")
                ProgressReport(6, "Marguee")

                Try

                    For i = 0 To mylist.Count - 1
                        'Dim drv As DataRowView = bsdetail.AddNew
                        'Dim dr As DataRow = drv.Row
                        myrecord = mylist(i)

                        'dr.Item(1) = myrowhd.Row.Item("pricechangehdid")
                        'dr.Item(2) = myrecord(0) 'vendorcode long
                        Dim mycmmf As Object
                        Dim result As Object
                        If myrecord(1) = "" Then
                            mycmmf = DBNull.Value
                            Dim mykey(0) As Object
                            mykey(0) = myrecord(2)
                            'result = DS.Tables(5).Rows.Find(mykey)
                            'result = MyDS.Tables(1).Rows.Find(mykey)
                            result = DS2.Tables(0).Rows.Find(mykey)
                            If Not IsNothing(result) Then
                                mycmmf = result.Item("cmmf")
                            Else
                                Err.Raise(1, Description:="Error:: No CMMF found for " & myrecord(2))

                            End If
                        Else
                            mycmmf = myrecord(1)
                        End If
                        Dim mykey1(3) As Object
                        mykey1(0) = myrecord(0)
                        mykey1(1) = mycmmf 'myrecord(1)
                        mykey1(2) = CDate(myrecord(5))
                        mykey1(3) = myrecord(4)

                        result = DSLocal.Tables(0).Rows.Find(mykey1)
                        Dim duplicate As Boolean = False
                        If Not IsNothing(result) Then
                            duplicate = True
                            DuplicateCheck = ". Duplicate data found!!"
                        End If

                        'dr.Item(3) = mycmmf
                        ''dr.Item(3) = If(myrecord(1) = "", DBNull.Value, myrecord(1)) 'cmmf long
                        'dr.Item(4) = myrecord(3) 'purchorg  integer
                        'dr.Item(5) = myrecord(4) 'plant integer
                        'dr.Item(6) = myrecord(5) 'validon date
                        'dr.Item(7) = myrecord(6) 'price numeric
                        'dr.Item(8) = myrecord(7) 'priceingunit integer
                        'dr.Item(9) = myrecord(8) 'comment text
                        'dr.Item(10) = myrecord(2) 'commercialcode text
                        'DS.Tables(4).Rows.Add(dr)


                        sb.Append(creator.Replace("\", "\\") & vbTab &
                                  myrowhd.Row.Item("pricechangehdid") & vbTab &
                                  DbAdapter1.validlongNull(myrecord(0)) & vbTab &
                                  DbAdapter1.validlongNull(mycmmf) & vbTab &
                                  DbAdapter1.validint(myrecord(3)) & vbTab &
                                  DbAdapter1.validint(myrecord(4)) & vbTab &
                                  DbAdapter1.dateformatYYYYMMdd(myrecord(5)) & vbTab &
                                  DbAdapter1.validdec(myrecord(6)) & vbTab &
                                  DbAdapter1.validdec(myrecord(7)) & vbTab &
                                  DbAdapter1.validcharNull(myrecord(8)) & vbTab &
                                  DbAdapter1.validcharNull(myrecord(2)) & vbTab &
                                  IIf(duplicate, duplicate, "Null") & vbCrLf)
                        '
                    Next

                    'copy to db
                    If sb.Length > 0 Then
                        Dim ra As Long = 0
                        Dim sqlstr As String = ""
                        sqlstr = "delete from pricechangedtltemp where creator =  '" & creator & "';" &
                                 "copy pricechangedtltemp(creator,pricechangehdid,vendorcode,cmmf,purchorg,plant,validon,price,pricingunit,comment,commercialcode,duplicate) from stdin with null as 'Null';"

                        ProgressReport(1, "Start Add New Records")
                        'mystr.Append(sqlstr)

                        Dim errmessage As String = String.Empty
                        Dim myret As Boolean = False
                        'If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, errmessage) Then
                        '    MessageBox.Show(errmessage)
                        'Else
                        '    ProgressReport(1, "Update Done.")
                        'End If
                        Try
                            Dim mymessage As String = String.Empty
                            DSDetail = New DataSet
                            errmessage = DbAdapter1.copy(sqlstr, sb.ToString, myret)
                            If myret Then
                                'sqlstr = "with pl as (select cmmf, max(validfrom) as validfrom from pricelist " &
                                '         " group by cmmf)," &
                                '         "  std as (select cmmf, max(validfrom) as validfrom from standardcostad group by cmmf) " &
                                '         " select p.cmmf,c.commercialref,p.vendorcode,p.purchorg,p.plant,p.validon,p.price,p.pricingunit,p.comment,c.commercialref,v.vendorname::character varying,v.shortname::character varying,materialdesc::character varying,r.range,r.rangedesc,ad.planprice1, ad.per,pr.amount::numeric, pr.perunit::numeric," &
                                '         " (getdelta(p.price,p.pricingunit,ad.planprice1,ad.per)) as deltastd," &
                                '         " (getdelta(p.price,p.pricingunit,pr.amount::numeric,pr.perunit::numeric) ) as deltasap," &
                                '         " getalert(p.price,p.pricingunit,pr.amount::numeric,pr.perunit::numeric,ad.planprice1,ad.per) as alert " &
                                '         " from pricechangedtltemp p" &
                                '         " left join vendor v on v.vendorcode = p.vendorcode" &
                                '         " left join cmmf c on c.cmmf = p.cmmf" &
                                '         " left join range r on r.rangeid = c.rangeid" &
                                '         " left join std on std.cmmf = p.cmmf" &
                                '         " left join standardcostad ad on ad.cmmf = std.cmmf and ad.validfrom = std.validfrom" &
                                '         " left join pl on pl.cmmf = p.cmmf" &
                                '         " left join pricelist pr on pr.cmmf = pl.cmmf and pr.validfrom = pl.validfrom and pr.vendorcode = v.vendorcode" &
                                '         " where creator = '" & creator & "'"

                                'sqlstr = "with std as (select cmmf, max(validfrom) as validfrom from standardcostad group by cmmf) " &
                                '         " select p.cmmf,c.commercialref,p.vendorcode,p.purchorg,p.plant,p.validon,p.price,p.pricingunit,p.comment,c.commercialref,v.vendorname::character varying,v.shortname::character varying,materialdesc::character varying,r.range,r.rangedesc,ad.planprice1, ad.per,pr.amount::numeric, pr.perunit::numeric," &
                                '         " (getdelta(p.price,p.pricingunit,ad.planprice1,ad.per)) as deltastd," &
                                '         " (getdelta(p.price,p.pricingunit,pr.amount::numeric,pr.perunit::numeric) ) as deltasap," &
                                '         " getalert(p.price,p.pricingunit,pr.amount::numeric,pr.perunit::numeric,ad.planprice1,ad.per,p.validon,pr.validfrom) as alert,pr.validfrom ,duplicate" &
                                '         " from pricechangedtltemp p" &
                                '         " left join vendor v on v.vendorcode = p.vendorcode" &
                                '         " left join cmmf c on c.cmmf = p.cmmf" &
                                '         " left join range r on r.rangeid = c.rangeid" &
                                '         " left join std on std.cmmf = p.cmmf" &
                                '         " left join standardcostad ad on ad.cmmf = std.cmmf and ad.validfrom = std.validfrom" &
                                '         " left join pricelist pr on pr.pricelistid = (select pricelistid from pricelist " &
                                '         " where cmmf = p.cmmf and vendorcode = p.vendorcode order by validfrom desc limit 1)" &
                                '         " where creator = '" & creator & "'"
                                sqlstr = "with std as (select cmmf, max(validfrom) as validfrom from standardcostad group by cmmf) " &
                                         " select p.cmmf,c.commercialref,p.vendorcode,p.purchorg,p.plant,p.validon,p.price,p.pricingunit,p.comment,c.commercialref,v.vendorname::character varying,v.shortname::character varying,materialdesc::character varying,r.range,r.rangedesc,ad.planprice1, ad.per,pr.amount::numeric, pr.perunit::numeric," &
                                         " (getdelta(p.price,p.pricingunit,ad.planprice1,ad.per)) as deltastd," &
                                         " (getdelta(p.price,p.pricingunit,pr.amount::numeric,pr.perunit::numeric) ) as deltasap," &
                                         " getalert(p.price,p.pricingunit,pr.amount::numeric,pr.perunit::numeric,ad.planprice1,ad.per,p.validon,pr.validfrom) as alert,pr.validfrom ,duplicate" &
                                         " from pricechangedtltemp p" &
                                         " left join vendor v on v.vendorcode = p.vendorcode" &
                                         " left join cmmf c on c.cmmf = p.cmmf" &
                                         " left join range r on r.rangeid = c.rangeid" &
                                         " left join std on std.cmmf = p.cmmf" &
                                         " left join standardcostad ad on ad.cmmf = std.cmmf and ad.validfrom = std.validfrom" &
                                         " left join pricelistscale pr on pr.id = (select id from pricelistscale ps" &
                                         " where ps.cmmf = p.cmmf and ps.vendorcode = p.vendorcode and ps.plant = p.plant order by validfrom desc limit 1)" &
                                         " where creator = '" & creator & "'"
                                If DbAdapter1.TbgetDataSet(sqlstr, DSDetail, mymessage) Then
                                    ProgressReport(7, "assign to db")
                                Else
                                    ProgressReport(1, errmessage)
                                End If
                            End If


                        Catch ex As Exception
                            ProgressReport(1, ex.Message)
                        End Try
                    Else

                    End If
                    'retrive data to datagridview

                Catch ex As Exception
                    ProgressReport(2, ex.Message)
                    ProgressReport(5, "Continuous")
                    Exit Sub
                End Try
            End With
        End Using
        ProgressReport(4, "Refresh Datagridview")
        ProgressReport(5, "Continuous")
        ProgressReport(2, "Done. Record found " & recordcount & DuplicateCheck)
    End Sub
    Private Sub doimport2()
        ProgressReport(4, "Refresh Datagridview")
        ProgressReport(5, "Marguee")
        Dim sb As New StringBuilder
        Dim errmsg As String = String.Empty
        If Not ImportMyTask(myfilename, bsdetail, Me, errmsg) Then
            ProgressReport(1, errmsg)
            Exit Sub
        End If

        Using objTFParser = New FileIO.TextFieldParser(Replace(myfilename, ".xlsx", ".txt"))
            Dim myrecord() As String
            Dim mylist As New List(Of String())
            With objTFParser
                .TextFieldType = FileIO.FieldType.Delimited
                .SetDelimiters(Chr(9))
                .HasFieldsEnclosedInQuotes = True
                Dim count As Long = 0

                Do Until .EndOfData
                    myrecord = .ReadFields
                    If count > 0 Then
                        mylist.Add(myrecord)
                    End If
                    count += 1
                Loop
                ProgressReport(2, "Build Record...")
                ProgressReport(5, "Marguee")

                Try

                    For i = 0 To mylist.Count - 1
                        'Dim drv As DataRowView = bsdetail.AddNew
                        'Dim dr As DataRow = drv.Row
                        myrecord = mylist(i)

                        'dr.Item(1) = myrowhd.Row.Item("pricechangehdid")
                        'dr.Item(2) = myrecord(0) 'vendorcode long
                        Dim mycmmf As Object
                        If myrecord(1) = "" Then
                            mycmmf = DBNull.Value
                            Dim mykey(0) As Object
                            mykey(0) = myrecord(2)
                            'Dim result = DS.Tables(5).Rows.Find(mykey)
                            Dim result = MyDS.Tables(1).Rows.Find(mykey)

                            If Not IsNothing(result) Then
                                mycmmf = result.Item("cmmf")
                            Else
                                Err.Raise(1, Description:="Error:: No CMMF found for " & myrecord(2))

                            End If
                        Else
                            mycmmf = myrecord(1)
                        End If
                        'dr.Item(3) = mycmmf
                        ''dr.Item(3) = If(myrecord(1) = "", DBNull.Value, myrecord(1)) 'cmmf long
                        'dr.Item(4) = myrecord(3) 'purchorg  integer
                        'dr.Item(5) = myrecord(4) 'plant integer
                        'dr.Item(6) = myrecord(5) 'validon date
                        'dr.Item(7) = myrecord(6) 'price numeric
                        'dr.Item(8) = myrecord(7) 'priceingunit integer
                        'dr.Item(9) = myrecord(8) 'comment text
                        'dr.Item(10) = myrecord(2) 'commercialcode text
                        'DS.Tables(4).Rows.Add(dr)


                        sb.Append(creator.Replace("\", "\\") & vbTab &
                                  myrowhd.Row.Item("pricechangehdid") & vbTab &
                                  DbAdapter1.validlongNull(myrecord(0)) & vbTab &
                                  DbAdapter1.validlongNull(mycmmf) & vbTab &
                                  DbAdapter1.validint(myrecord(3)) & vbTab &
                                  DbAdapter1.validint(myrecord(4)) & vbTab &
                                  DbAdapter1.dateformatYYYYMMdd(myrecord(5)) & vbTab &
                                  DbAdapter1.validdec(myrecord(6)) & vbTab &
                                  DbAdapter1.validdec(myrecord(7)) & vbTab &
                                  DbAdapter1.validcharNull(myrecord(8)) & vbTab &
                                  DbAdapter1.validcharNull(myrecord(2)) & vbCrLf)
                    Next

                    'copy to db
                    If sb.Length > 0 Then
                        Dim ra As Long = 0
                        Dim sqlstr As String = ""
                        sqlstr = "delete from pricechangedtltemp where creator =  '" & creator & "';" &
                                 "copy pricechangedtltemp(creator,pricechangehdid,vendorcode,cmmf,purchorg,plant,validon,price,pricingunit,comment,commercialcode) from stdin with null as 'Null';"

                        ProgressReport(1, "Start Add New Records")
                        'mystr.Append(sqlstr)

                        Dim errmessage As String = String.Empty
                        Dim myret As Boolean = False
                        'If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, errmessage) Then
                        '    MessageBox.Show(errmessage)
                        'Else
                        '    ProgressReport(1, "Update Done.")
                        'End If
                        Try
                            Dim mymessage As String = String.Empty
                            DSDetail = New DataSet
                            errmessage = DbAdapter1.copy(sqlstr, sb.ToString, myret)
                            If myret Then
                                'sqlstr = "with pl as (select cmmf, max(validfrom) as validfrom from pricelist " &
                                '         " group by cmmf)," &
                                '         "  std as (select cmmf, max(validfrom) as validfrom from standardcostad group by cmmf) " &
                                '         " select p.cmmf,c.commercialref,p.vendorcode,p.purchorg,p.plant,p.validon,p.price,p.pricingunit,p.comment,c.commercialref,v.vendorname::character varying,v.shortname::character varying,materialdesc::character varying,r.range,r.rangedesc,ad.planprice1, ad.per,pr.amount::numeric, pr.perunit::numeric," &
                                '         " (getdelta(p.price,p.pricingunit,ad.planprice1,ad.per)) as deltastd," &
                                '         " (getdelta(p.price,p.pricingunit,pr.amount::numeric,pr.perunit::numeric) ) as deltasap," &
                                '         " getalert(p.price,p.pricingunit,pr.amount::numeric,pr.perunit::numeric,ad.planprice1,ad.per) as alert " &
                                '         " from pricechangedtltemp p" &
                                '         " left join vendor v on v.vendorcode = p.vendorcode" &
                                '         " left join cmmf c on c.cmmf = p.cmmf" &
                                '         " left join range r on r.rangeid = c.rangeid" &
                                '         " left join std on std.cmmf = p.cmmf" &
                                '         " left join standardcostad ad on ad.cmmf = std.cmmf and ad.validfrom = std.validfrom" &
                                '         " left join pl on pl.cmmf = p.cmmf" &
                                '         " left join pricelist pr on pr.cmmf = pl.cmmf and pr.validfrom = pl.validfrom and pr.vendorcode = v.vendorcode" &
                                '         " where creator = '" & creator & "'"

                                sqlstr = "with std as (select cmmf, max(validfrom) as validfrom from standardcostad group by cmmf) " &
                                         " select p.cmmf,c.commercialref,p.vendorcode,p.purchorg,p.plant,p.validon,p.price,p.pricingunit,p.comment,c.commercialref,v.vendorname::character varying,v.shortname::character varying,materialdesc::character varying,r.range,r.rangedesc,ad.planprice1, ad.per,pr.amount::numeric, pr.perunit::numeric," &
                                         " (getdelta(p.price,p.pricingunit,ad.planprice1,ad.per)) as deltastd," &
                                         " (getdelta(p.price,p.pricingunit,pr.amount::numeric,pr.perunit::numeric) ) as deltasap," &
                                         " getalert(p.price,p.pricingunit,pr.amount::numeric,pr.perunit::numeric,ad.planprice1,ad.per) as alert,pr.validfrom " &
                                         " from pricechangedtltemp p" &
                                         " left join vendor v on v.vendorcode = p.vendorcode" &
                                         " left join cmmf c on c.cmmf = p.cmmf" &
                                         " left join range r on r.rangeid = c.rangeid" &
                                         " left join std on std.cmmf = p.cmmf" &
                                         " left join standardcostad ad on ad.cmmf = std.cmmf and ad.validfrom = std.validfrom" &
                                         " left join pricelist pr on pr.pricelistid = (select pricelistid from pricelist " &
                                         " where cmmf = p.cmmf and vendorcode = p.vendorcode order by validfrom desc limit 1)" &
                                         " where creator = '" & creator & "'"
                                If DbAdapter1.TbgetDataSet(sqlstr, DSDetail, mymessage) Then
                                    ProgressReport(7, "assign to db")
                                Else
                                    ProgressReport(1, errmessage)
                                End If
                            End If


                        Catch ex As Exception
                            ProgressReport(1, ex.Message)
                        End Try
                    Else

                    End If
                    'retrive data to datagridview

                Catch ex As Exception
                    ProgressReport(2, ex.Message)
                    ProgressReport(6, "Continuous")
                    Exit Sub
                End Try
            End With
        End Using
        ProgressReport(4, "Refresh Datagridview")
        ProgressReport(6, "Continuous")
        ProgressReport(2, "Done. Record found " & recordcount)
    End Sub

    Private Sub doimportOld()
        ProgressReport(4, "Refresh Datagridview")
        Dim errmsg As String = String.Empty
        If Not ImportMyTask(myfilename, bsdetail, Me, errmsg) Then
            ProgressReport(1, errmsg)
            Exit Sub
        End If
        Using objTFParser = New FileIO.TextFieldParser(Replace(myfilename, ".xlsx", ".txt"))
            Dim myrecord() As String
            Dim mylist As New List(Of String())
            With objTFParser
                .TextFieldType = FileIO.FieldType.Delimited
                .SetDelimiters(Chr(9))
                .HasFieldsEnclosedInQuotes = True
                Dim count As Long = 0

                Do Until .EndOfData
                    myrecord = .ReadFields
                    If count > 0 Then
                        mylist.Add(myrecord)
                    End If
                    count += 1
                Loop
                ProgressReport(2, "Build Record...")
                ProgressReport(6, "Marquee")
                Try
                    For i = 0 To mylist.Count - 1
                        Dim drv As DataRowView = bsdetail.AddNew
                        Dim dr As DataRow = drv.Row
                        myrecord = mylist(i)

                        dr.Item(1) = CType(bsHeader.Current, DataRowView).Row.Item("pricechangehdid")
                        dr.Item(2) = myrecord(0) 'vendorcode
                        dr.Item(3) = If(myrecord(1) = "", DBNull.Value, myrecord(1)) 'cmmf
                        dr.Item(4) = myrecord(3) ' 
                        dr.Item(5) = myrecord(4)
                        dr.Item(6) = myrecord(5)
                        dr.Item(7) = myrecord(6)
                        dr.Item(8) = myrecord(7)
                        dr.Item(9) = myrecord(8)
                        dr.Item(10) = myrecord(2)
                        'DS.Tables(4).Rows.Add(dr)
                        MyDS.Tables(0).Rows.Add(dr)
                    Next
                Catch ex As Exception
                    ProgressReport(2, ex.Message)
                    Exit Sub
                End Try
            End With
        End Using
        ProgressReport(4, "Refresh Datagridview")
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click

        Dim dr As DataRow = CType(bsHeader.Current, DataRowView).Row
        Dim myQueryWorksheetList As New List(Of QueryWorksheet)
        Me.ToolStripStatusLabel1.Text = ""
        Me.ToolStripStatusLabel2.Text = ""

        Dim mymessage As String = String.Empty

        Dim priceChangeId As String = String.Empty

        Dim sqlstr As String
        Dim sqlstr1 As String = String.Empty
        Dim myusername As String
        'If ComboBox1.Text <> "" Then
        priceChangeId = dr.Item("pricechangehdid")
        myusername = dr.Item("creatorname")


        ' sqlstr = "select tb.* from sp_gethardcopy(" & myuser & "::character varying," & DateFormatyyyyMMdd(dateupload) & "::date) as tb(invoicehardcopydtid bigint,supplierinvoice character varying,amount numeric,vendorname character varying,readydate date,accountingdoc bigint,delivery bigint,item integer,pohd bigint,poitem integer,deliverydate date,sebinvoice bigint,billoflading character varying)"
        'sqlstr = "with pl as (select cmmf, max(validfrom) as validfrom from pricelist  group by cmmf) ," &
        '         " std as (select cmmf, max(validfrom) as validfrom from standardcostad group by cmmf) " &
        '         " select dt.vendorcode,dt.cmmf as material,c.commercialref,dt.purchorg,dt.plant,dt.validon,dt.price as rate,dt.pricingunit,dt.comment,r.rangedesc,r.range,v.vendorname::character varying,v.shortname::character varying,materialdesc::character varying, pr.amount::numeric,pr.perunit::numeric, (getdelta(dt.price,dt.pricingunit,ad.planprice1,ad.per)) as deltastd,  (getdelta(dt.price,dt.pricingunit,pr.amount::numeric,pr.perunit::numeric) ) as deltasap, getalert(dt.price,dt.pricingunit,pr.amount::numeric,pr.perunit::numeric,ad.planprice1,ad.per) as alert    " &
        '         " from pricechangedtl dt " &
        '         " left join cmmf c on c.cmmf = dt.cmmf  " &
        '         " left join vendor v on v.vendorcode = dt.vendorcode left join range r on r.rangeid = c.rangeid" &
        '         " left join std on std.cmmf = c.cmmf left join standardcostad ad on ad.cmmf = std.cmmf and ad.validfrom = std.validfrom left join pl on pl.cmmf = c.cmmf" &
        '         " left join pricelist pr on pr.cmmf = pl.cmmf and pr.validfrom = pl.validfrom" &
        '         " where(dt.pricechangehdid = " & priceChangeId & ")"
        sqlstr = "with std as (select cmmf, max(validfrom) as validfrom from standardcostad group by cmmf) " &
                 " select dt.vendorcode,dt.cmmf as material,c.commercialref,dt.purchorg,dt.plant,dt.validon,dt.price as rate,dt.pricingunit,dt.comment,r.rangedesc,r.range,v.vendorname::character varying,v.shortname::character varying,materialdesc::character varying, pr.amount::numeric,pr.perunit::numeric, (getdelta(dt.price,dt.pricingunit,ad.planprice1,ad.per)) as deltastd,  (getdelta(dt.price,dt.pricingunit,pr.amount::numeric,pr.perunit::numeric) ) as deltasap, getalert(dt.price,dt.pricingunit,pr.amount::numeric,pr.perunit::numeric,ad.planprice1,ad.per) as alert,dt.sap" &
                 " from pricechangedtl dt " &
                 " left join cmmf c on c.cmmf = dt.cmmf  " &
                 " left join vendor v on v.vendorcode = dt.vendorcode left join range r on r.rangeid = c.rangeid" &
                 " left join std on std.cmmf = c.cmmf left join standardcostad ad on ad.cmmf = std.cmmf and ad.validfrom = std.validfrom" &
                 " left join pricelist pr on pr.pricelistid =  (select pricelistid from pricelist " &
                 " where cmmf = dt.cmmf and vendorcode = dt.vendorcode order by validfrom desc limit 1)" &
                 " where(dt.pricechangehdid = " & priceChangeId & ")"
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
    Private Sub FormattingReport(ByRef sender As Object, ByRef e As EventArgs)
        'Throw New NotImplementedException
    End Sub

    Private Sub PivotTable(ByRef sender As Object, ByRef e As EventArgs)
        'Throw New NotImplementedException
    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        If Me.validate Then
            If MessageBox.Show("Reject the records?", "Reject Record", MessageBoxButtons.OKCancel) = DialogResult.OK Then
                Dim drv As DataRowView = bsHeader.Current
                drv.Item("status") = TaskStatus.StatusRejected '3
                drv.Item("actiondate") = Date.Today
                drv.Item("actionby") = HelperClass1.UserId.ToLower
                drv.EndEdit()
                'DS.Tables(0).Rows(0).Item("status") = 3
                'DS.Tables(0).Rows(0).Item("actiondate") = Date.Today
                'DS.Tables(0).Rows(0).Item("actionby") = HelperClass1.UserId.ToLower
                bsHeader.EndEdit()
                bsdetail.EndEdit()
                Dim ds2 As New DataSet
                ds2 = DS.GetChanges
                If Not IsNothing(ds2) Then
                    Dim mymessage As String = String.Empty
                    Dim ra As Integer
                    'reset sequence number
                    Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
                    If Not DbAdapter1.PriceChangeTx(Me, mye) Then
                        ProgressReport(1, mye.message)
                    End If
                    Me.DialogResult = DialogResult.OK
                    Me.Close()
                End If
            End If
        End If
    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click

        If Me.validate Then
            ProgressReport(1, "")
            bsHeader.EndEdit()
            bsdetail.EndEdit()
            If MessageBox.Show("Validate the records?", "Validate Record", MessageBoxButtons.OKCancel) = DialogResult.OK Then
                Dim drv As DataRowView = bsHeader.Current
                drv.Item("status") = TaskStatus.StatusValidated '5
                drv.Item("actiondate") = Date.Today
                drv.Item("actionby") = HelperClass1.UserId.ToLower
                drv.EndEdit()
                'DS.Tables(0).Rows(0).Item("status") = 5
                'DS.Tables(0).Rows(0).Item("actiondate") = Date.Today
                'DS.Tables(0).Rows(0).Item("actionby") = HelperClass1.UserId.ToLower
                'Me.validate()
                Dim ds2 As New DataSet
                ds2 = DS.GetChanges
                If Not IsNothing(ds2) Then
                    Dim mymessage As String = String.Empty
                    Dim ra As Integer
                    'reset sequence number
                    Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)

                    If Not DbAdapter1.PriceChangeTx(Me, mye) Then
                        ProgressReport(1, mye.message)
                        Exit Sub
                    End If

                    'Update Price Comment
                    'If Not DbAdapter1.PriceCommentTx(myrowhd.Row.Item("pricechangehdid"), mymessage) Then
                    '    ProgressReport(1, mye.message)
                    '    Exit Sub
                    'End If
                    Me.DialogResult = DialogResult.OK
                    Me.Close()
                End If
            End If
        Else
            ProgressReport(1, "Found error(s). Please check.")
            Exit Sub
        End If
    End Sub

    Private Sub ListBox1_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ListBox1.MouseDoubleClick
        Select Case System.IO.Path.GetExtension(ListBox1.SelectedItem)
            Case ".msg"
                Call SavingFile()
            Case Else
                Dim destination As String = DS.Tables(8).Rows(0).Item("cvalue").ToString
                Dim p As New System.Diagnostics.Process
                p.StartInfo.FileName = destination & "\" & ListBox1.SelectedItem
                Try
                    p.Start()
                Catch ex As Exception

                End Try
        End Select


    End Sub

    Sub SavingFile()
        Dim selectfolder As New FolderBrowserDialog
        selectfolder.Description = "Please select destination folder to copy the file " & ListBox1.SelectedItem
        If selectfolder.ShowDialog = DialogResult.OK Then
            Dim destination As String = DS.Tables(8).Rows(0).Item("cvalue").ToString
            Dim sourcefilename = destination & "\" & ListBox1.SelectedItem
            Dim destfilename = selectfolder.SelectedPath & "\" & ListBox1.SelectedItem
            Try
                IO.File.Copy(sourcefilename, destfilename, True)
                MessageBox.Show("File copied to " & destfilename)
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If

    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged, TextBox4.TextChanged
        Dim myobj As TextBox = DirectCast(sender, TextBox)
        ErrorProvider1.SetError(myobj, "")
    End Sub

    Private Sub loaddata()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Private Sub DoWork()

        Dim mymessage As String = String.Empty
        MyDS = New DataSet
        ProgressReport(1, "Loading Data... Please wait.")
        ProgressReport(6, "Marque")
        If DbAdapter1.TbgetDataSet(sb.ToString, MyDS, mymessage) Then
            Try
                'bsdetail.DataSource = MyDS.Tables(0)
                ''bsdetailtmp.DataSource = MyDS.Tables(0)
                'MyDS.Tables(0).TableName = "Detail Add"
                If myTXType = TXType.Task Then
                    bsdetail.DataSource = DS.Tables(4)
                    bsdetailtmp.DataSource = DS.Tables(4) 'temp needed for create new record using import method
                Else
                    bsdetail.DataSource = MyDS.Tables(0)
                    bsdetailtmp.DataSource = MyDS.Tables(0)
                End If
                
                bsdetail.Filter = "pricechangehdid=" & myrowhd.Row.Item("pricechangehdid").ToString

                ''If New
                'DS.Tables.Add(MyDS.Tables(0).Copy)
                'bsdetailtmp.DataSource = DS.Tables(10)
                ''add relationship
                'Dim rel As DataRelation
                'Dim hcol As DataColumn
                'Dim dcol As DataColumn
                ''create relation ds.table(0) and ds.table(4)
                'hcol = DS.Tables(0).Columns("pricechangehdid") 'docemailhdid in table header
                'dcol = DS.Tables(10).Columns("pricechangehdid") 'docemailhdid in table dtl
                'rel = New DataRelation("hdrel", hcol, dcol)
                'DS.Relations.Add(rel)


                'Dim idx1(0) As DataColumn
                'idx1(0) = MyDS.Tables(1).Columns(0)
                'MyDS.Tables(1).PrimaryKey = idx1
            Catch ex As Exception
                ProgressReport(1, "Loading Data. Error::" & ex.Message)
                ProgressReport(5, "Continuous")
                Exit Sub
            End Try
            ProgressReport(8, "InitData")
        Else
            ProgressReport(1, "Loading Data. Error::" & mymessage)
            ProgressReport(5, "Continuous")
            Exit Sub
        End If
        ProgressReport(1, "Loading Data.Done!")
        ProgressReport(5, "Continuous")




        ProgressReport(1, "Done!Record Found " & recordcount)
    End Sub


    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click, Button5.Click
        Dim myobj = CType(sender, Button)
        Dim bs As New BindingSource
        Dim myform As New FormHelper(bs)
        bs.Filter = ""
        Select Case myobj.Name
            Case "Button4" 'Reason
                bs.DataSource = DS.Tables(9)
                myform.DataGridView1.Columns(0).DataPropertyName = "reasonname"
            Case "Button5" 'SpecialProject
                bs.DataSource = DS.Tables(12)
                myform.DataGridView1.Columns(0).DataPropertyName = "specialproject"
        End Select

        If myform.ShowDialog = DialogResult.OK Then
            Dim drv As DataRowView = bs.Current
            Select Case myobj.Name
                Case "Button4"
                    TextBox6.Text = drv.Row.Item("reasonname")
                    myrowhd.Row.Item("reasonid") = drv.Row.Item("id")
                    'myrowhd.Row.Item("reasonname") = drv.Row.Item("reasonname")
                Case "Button5"
                    TextBox7.Text = "" + drv.Row.Item("specialproject")
                    myrowhd.Row.Item("specialprojectid") = drv.Row.Item("id")
                    'myrowhd.Row.Item("specialproject") = drv.Row.Item("specialproject")
            End Select
        End If
    End Sub

End Class