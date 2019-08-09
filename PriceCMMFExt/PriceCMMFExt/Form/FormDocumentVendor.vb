Imports PriceCMMFExt.SharedClass
Imports System.Threading
Imports System.Text
Imports PriceCMMFExt.PublicClass
Public Class FormDocumentVendor
    Dim WithEvents bsheader As BindingSource
    Dim WithEvents bsDetail As BindingSource

    Dim bsShortname As BindingSource
    Dim bsVendorname As BindingSource
    Dim bsShortnameVendor As BindingSource
    Dim WithEvents bsDocType As BindingSource
    Dim bsDocLevel As BindingSource
    Dim bsPaymentTerm As BindingSource

    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)

    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Public Property DS As DataSet

    Dim sb As New StringBuilder
    Dim myuser As String
    Dim headerid As Long
    Dim validatorid As String

    Dim cc1id As String
    Dim cc2id As String
    Dim cc3id As String
    Dim cc4id As String

    Public Sub New(ByVal headerid As Long)
        InitializeComponent()
        'Update
        Me.headerid = headerid
        loaddata(headerid)
    End Sub
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        'Create New
        headerid = 0
        loaddata(headerid)

    End Sub

    Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click
        If MessageBox.Show("Cancel current task?", "Cancel", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) = DialogResult.Yes Then
            bsheader.CancelEdit()
            Me.DialogResult = DialogResult.Cancel
        End If
        
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click



        If Me.Validate() Then
            bsheader.EndEdit()
            bsDetail.EndEdit()
            Try
                'get modified rows, send all rows to stored procedure. let the stored procedure create a new record.
                Dim ds2 As DataSet
                ds2 = DS.GetChanges
                If Not IsNothing(ds2) Then
                    Dim mymessage As String = String.Empty
                    Dim ra As Integer
                    Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
                    If DbAdapter1.DocumentVendorTx(Me, mye) Then                        
                        'delete original Dataset (DS) for those table having added record -> Merged with modified Dataset (DS2)
                        'For update record, no need to delete the original dataset (DS) because the id is the same. 
                        'Why need to delete the added one, because when we create new record, the id started with 0,-1,-2 and so on.
                        'when we update to database, we put the real id from database.
                        'so we have different value id for DS and DS2. if we do merged without deleting the original one, we will have 2 records.
                        For i = 0 To 1
                            Dim modifiedRows = From row In DS.Tables(i)
                                Where row.RowState = DataRowState.Modified
                            For Each row In modifiedRows.ToArray
                                row.Delete()
                            Next
                        Next
                    Else
                        MessageBox.Show(mye.message)
                        Exit Sub
                    End If
                    DS.Merge(ds2)
                    DS.AcceptChanges()
                    MessageBox.Show("Saved.")
                End If

                'copy file
                For Each drv As DataRowView In bsDetail.List

                Next
                'create record
                Me.DialogResult = DialogResult.OK
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
            


        Else
            'bsheader.CancelEdit()
            'Me.DialogResult = DialogResult.Cancel
        End If

    End Sub
    Public Overloads Function validate() As Boolean
        Dim myret As Boolean = True
        MyBase.Validate()


        For Each drv As DataRowView In bsDetail.List
            'drv.Row.RowError = "Has Error loh"
            If Not validaterecord(drv) Then
                myret = False
            End If            
        Next

        DataGridView1.Invalidate()
        Return myret
    End Function

    Private Function validaterecord(ByVal drv As DataRowView) As Boolean
        Dim myerror As New StringBuilder
        Dim myret As Boolean = True
        If IsDBNull(drv.Row.Item("doctypename")) Then
            myerror.Append("Document Type cannot be blank.")
            myret = False
        End If

        If IsDBNull(drv.Row.Item("vendorcode")) And IsDBNull(drv.Row.Item("shortname")) Then
            myerror.Append("Vendor Name and Shortname cannot be blank.")        
            myret = False
        End If
        drv.Row.RowError = myerror.ToString
        Return myret

    End Function
    Public Sub loaddata(ByVal id As Long)
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Sub DoWork()
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Loading Data.")
        'ProgressReport(4, "InitData")
        '2 Dataset 
        '1 contains All tx except Completed
        'the other only contains Completed

        DS = New DataSet
        Dim mymessage As String = String.Empty
        sb.Clear()
        'Admin checking first
        'sb.Append("select * from pricechangehd ph where (ph.creator = '" & HelperClass1.UserId & "')")

        'myuser = HelperClass1.UserId.ToLower
        'myuser = "as\dlie"
        'myuser = "as\elai"
        'myuser = "as\rlo"
        myuser = "AS\afok".ToLower
        'myuser = "as\weho"
        'myuser = "AS\shxu".ToLower
        'myuser = "as\jdai"
        'myuser = "AS\SCHAN".ToLower
        sb.Append("select distinct h.* ,o.officersebname::text as username,o2.officersebname::text as validatorname,o3.officersebname::text as cc1name ,o4.officersebname::text as cc2name,o5.officersebname::text as cc3name,o6.officersebname::text as cc4name" &
                  " from doc.header h" &
                  " left join officerseb o on o.userid = h.userid" &
                  " left join officerseb o2 on o2.userid = h.validator" &
                  " left join officerseb o3 on o3.userid = h.cc1" &
                  " left join officerseb o4 on o4.userid = h.cc2" &
                  " left join officerseb o5 on o5.userid = h.cc3" &
                  " left join officerseb o6 on o6.userid = h.cc4  where h.id = " & headerid & ";") 'add join with ofsebid
        sb.Append("select vd.*,v.vendorname::text,v.shortname::text,d.*,vr.version,gt.paymentcode,sc.leadtime,sc.sasl,q.nqsu,p.projectname,sa.auditby,sa.audittype,sa.auditgrade,sef.score,sif.myyear,sif.turnovery,sif.turnovery1,sif.turnovery2,sif.turnovery3,sif.turnovery4,sif.ratioy,sif.ratioy1,sif.ratioy2,sif.ratioy3,sif.ratioy4,'' as filename,dt.doctypename,dl.levelname from doc.vendordoc vd" &
                  " left join doc.header h on h.id = vd.headerid" &
                  " left join vendor v on v.vendorcode = vd.vendorcode" &
                  " left join doc.document d on d.id = vd.documentid" &
                  " left join doc.version vr on vr.documentid = d.id" &
                  " left join doc.generalcontract gt on gt.documentid = d.id" &
                  " left join doc.supplychain sc on sc.documentid = d.id" &
                  " left join doc.qualityappendix q on q.documentid = d.id" &
                  " left join doc.project p on p.documentid = d.id" &
                  " left join doc.socialaudit sa on sa.documentid = d.id" &
                  " left join doc.sef sef on sef.documentid = d.id" &
                  " left join doc.sif sif on sif.documentid = d.id" &
                  " left join doc.doctype dt on dt.id = d.doctypeid" &
                  " left join doc.doclevel dl on dl.id = d.doclevelid" &
                  " where h.id = " & headerid & ";")
        sb.Append("select null::text as shortname union all (select distinct shortname::text from vendor  where not shortname isnull order by shortname);")
        sb.Append("select shortname::text,vendorcode from vendor where not shortname isnull order by shortname;")
        sb.Append("select null as vendorcode,''::text as description,''::text as vendorname union all (select vendorcode, vendorcode::text || ' - ' || vendorname::text as description,vendorname::text from vendor order by vendorname);")
        sb.Append("select null as id,''::text as doctypename union all (select id,doctypename from doc.doctype order by doctypename);")
        sb.Append("select null as id,''::text as levelname union all (select id,levelname from doc.doclevel order by id);")
        sb.Append("select paymenttermid,payt  from paymentterm  order by payt;")
        'sb.Append("select * from officerseb o  where lower(o.userid) = '" & myuser & "' limit 1;")
        sb.Append("select ''::text as name,'' as userid,null as teamtitleid,'' as officersebname union all (select distinct teamtitleshortname || ' - ' || officersebname as name,lower(userid) as userid,tt.teamtitleid,officersebname from officerseb o left join teamtitle tt on tt.teamtitleid = o.teamtitleid where teamtitleshortname in ('PD','SPM','PM') and o.isactive and userid <> 'as\lili2' order by tt.teamtitleid,officersebname);")
        If DbAdapter1.TbgetDataSet(sb.ToString, DS, mymessage) Then
            Try

                DS.Tables(0).TableName = "Header"
                DS.Tables(1).TableName = "Detail"
                DS.Tables(2).TableName = "ShortName"
                DS.Tables(3).TableName = "ShortNameVendorCode"
                DS.Tables(4).TableName = "VendorName"
                DS.Tables(5).TableName = "DocType"
                DS.Tables(6).TableName = "DocLevel"
                DS.Tables(7).TableName = "PaymentTerm"
                DS.Tables(8).TableName = "User"
            Catch ex As Exception
                ProgressReport(1, "Loading Data. Error::" & ex.Message)
                ProgressReport(5, "Continuous")
                Exit Sub
            End Try
            ProgressReport(4, "InitData")
        Else
            ProgressReport(1, "Loading Data. Error::" & mymessage)
            ProgressReport(5, "Continuous")
            Exit Sub
        End If
        ProgressReport(1, "Loading Data.Done!")
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
                            bsheader = New BindingSource
                            bsDetail = New BindingSource

                            bsShortname = New BindingSource
                            bsShortnameVendor = New BindingSource
                            bsVendorname = New BindingSource
                            bsDocType = New BindingSource
                            bsDocLevel = New BindingSource
                            bsPaymentTerm = New BindingSource

                            Dim pk(0) As DataColumn
                            pk(0) = DS.Tables(0).Columns("id")
                            DS.Tables(0).PrimaryKey = pk
                            DS.Tables(0).Columns(0).AutoIncrement = True
                            DS.Tables(0).Columns(0).AutoIncrementSeed = 0
                            DS.Tables(0).Columns(0).AutoIncrementStep = -1
                            DS.Tables(0).TableName = "Header"


                            Dim rel As DataRelation
                            Dim hcol As DataColumn
                            Dim dcol As DataColumn
                            'create relation ds.table(0) and ds.table(1)
                            hcol = DS.Tables(0).Columns("id") 'id in table header
                            dcol = DS.Tables(1).Columns("headerid") 'headerid in table vendordoc
                            rel = New DataRelation("hdrel", hcol, dcol)
                            DS.Relations.Add(rel)

                            bsheader.DataSource = DS.Tables(0)
                            bsDetail.DataSource = DS.Tables(1)
                            bsShortname.DataSource = DS.Tables(2)
                            bsShortnameVendor.DataSource = DS.Tables(3)
                            bsVendorname.DataSource = DS.Tables(4)
                            bsDocType.DataSource = DS.Tables(5)
                            bsDocLevel.DataSource = DS.Tables(6)
                            bsPaymentTerm.DataSource = DS.Tables(7)

                            TextBox1.DataBindings.Clear()
                            TextBox2.DataBindings.Clear()
                            TextBox3.DataBindings.Clear()
                            TextBox4.DataBindings.Clear()
                            TextBox5.DataBindings.Clear()
                            TextBox6.DataBindings.Clear()
                            TextBox7.DataBindings.Clear()

                            If headerid = 0 Then ' New Record
                                'bsheader.AddNew()
                                Dim drv As DataRowView = bsheader.AddNew()                                
                                drv.Row.Item("creationdate") = Date.Today
                                drv.Row.Item("userid") = HelperClass1.UserId.ToLower
                                drv.Row.Item("username") = HelperClass1.UserInfo.DisplayName

                            End If



                            

                            ComboBox1.DataBindings.Clear()

                            ComboBox1.DisplayMember = "shortname"
                            ComboBox1.ValueMember = "shortname"
                            'ComboBox1.SelectedIndex = -1
                            ComboBox1.DataSource = bsShortname
                            'ComboBox1.DataBindings.Add("SelectedValue", bsDetail, "shortname")

                            ComboBox2.DataBindings.Clear()

                            ComboBox2.DisplayMember = "description"
                            ComboBox2.ValueMember = "vendorcode"
                            ComboBox2.DataSource = bsVendorname
                            ComboBox2.DataBindings.Add("Selectedvalue", bsDetail, "vendorcode", True, DataSourceUpdateMode.OnPropertyChanged)


                            ComboBox3.DataBindings.Clear()
                            ComboBox3.DataSource = bsDocType
                            ComboBox3.DisplayMember = "doctypename"
                            ComboBox3.ValueMember = "id"
                            ComboBox3.DataBindings.Add("SelectedValue", bsDetail, "doctypeid", True, DataSourceUpdateMode.OnPropertyChanged)

                            ComboBox4.DataBindings.Clear()
                            ComboBox4.DataSource = bsDocLevel
                            ComboBox4.DisplayMember = "levelname"
                            ComboBox4.ValueMember = "id"
                            ComboBox4.DataBindings.Add("SelectedValue", bsDetail, "doclevelid", True, DataSourceUpdateMode.OnPropertyChanged)


                            ComboBox5.DataBindings.Clear()
                            ComboBox5.DataSource = bsPaymentTerm
                            ComboBox5.DisplayMember = "payt"
                            ComboBox5.ValueMember = "paymenttermid"
                            ComboBox5.DataBindings.Add("SelectedValue", bsDetail, "paymentcode", True, DataSourceUpdateMode.OnPropertyChanged)

                            DateTimePicker1.DataBindings.Clear()

                            TextBox8.DataBindings.Clear()
                            TextBox9.DataBindings.Clear()
                            TextBox10.DataBindings.Clear()
                            TextBox11.DataBindings.Clear()
                            TextBox12.DataBindings.Clear()
                            TextBox13.DataBindings.Clear()
                            TextBox14.DataBindings.Clear()
                            TextBox15.DataBindings.Clear()
                            TextBox16.DataBindings.Clear()
                            TextBox17.DataBindings.Clear()
                            TextBox18.DataBindings.Clear()
                            TextBox19.DataBindings.Clear()
                            TextBox20.DataBindings.Clear()
                            TextBox21.DataBindings.Clear()
                            TextBox22.DataBindings.Clear()
                            TextBox23.DataBindings.Clear()
                            TextBox24.DataBindings.Clear()


                            TextBox1.DataBindings.Add(New Binding("Text", bsheader, "username", True, DataSourceUpdateMode.OnPropertyChanged))
                            TextBox2.DataBindings.Add(New Binding("Text", bsheader, "validatorname", True, DataSourceUpdateMode.OnPropertyChanged))
                            TextBox3.DataBindings.Add(New Binding("Text", bsheader, "cc1name", True, DataSourceUpdateMode.OnPropertyChanged))
                            TextBox4.DataBindings.Add(New Binding("Text", bsheader, "cc2name", True, DataSourceUpdateMode.OnPropertyChanged))
                            TextBox5.DataBindings.Add(New Binding("Text", bsheader, "cc3name", True, DataSourceUpdateMode.OnPropertyChanged))
                            TextBox6.DataBindings.Add(New Binding("Text", bsheader, "cc4name", True, DataSourceUpdateMode.OnPropertyChanged))
                            TextBox7.DataBindings.Add(New Binding("Text", bsheader, "otheremail", True, DataSourceUpdateMode.OnPropertyChanged))
                            DateTimePicker1.DataBindings.Add(New Binding("Text", bsheader, "creationdate"))

                            TextBox8.DataBindings.Add(New Binding("Text", bsDetail, "remarks", True, DataSourceUpdateMode.OnPropertyChanged))
                            TextBox9.DataBindings.Add(New Binding("Text", bsDetail, "version", True, DataSourceUpdateMode.OnPropertyChanged))
                            TextBox10.DataBindings.Add(New Binding("Text", bsDetail, "filename", True, DataSourceUpdateMode.OnPropertyChanged)) 'if no value means update
                            TextBox11.DataBindings.Add(New Binding("Text", bsDetail, "leadtime", True, DataSourceUpdateMode.OnPropertyChanged))
                            TextBox12.DataBindings.Add(New Binding("Text", bsDetail, "sasl", True, DataSourceUpdateMode.OnPropertyChanged, "", "##0 %"))
                            TextBox13.DataBindings.Add(New Binding("Text", bsDetail, "nqsu", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0.00"))
                            TextBox14.DataBindings.Add(New Binding("Text", bsDetail, "projectname", True, DataSourceUpdateMode.OnPropertyChanged))
                            TextBox15.DataBindings.Add(New Binding("Text", bsDetail, "auditby", True, DataSourceUpdateMode.OnPropertyChanged))
                            TextBox16.DataBindings.Add(New Binding("Text", bsDetail, "audittype", True, DataSourceUpdateMode.OnPropertyChanged))
                            TextBox17.DataBindings.Add(New Binding("Text", bsDetail, "auditgrade", True, DataSourceUpdateMode.OnPropertyChanged))
                            TextBox18.DataBindings.Add(New Binding("Text", bsDetail, "score", True, DataSourceUpdateMode.OnPropertyChanged, "", "##0 %"))

                            Dim b19 As Binding = New Binding("Text", bsDetail, "turnovery", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0.00")
                            Dim b20 As Binding = New Binding("Text", bsDetail, "turnovery1", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0.00")
                            Dim b21 As Binding = New Binding("Text", bsDetail, "turnovery2", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0.00")
                            Dim b22 As Binding = New Binding("Text", bsDetail, "turnovery3", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0.00")
                            Dim b23 As Binding = New Binding("Text", bsDetail, "turnovery4", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0.00")
                            Dim b24 As Binding = New Binding("Text", bsDetail, "ratioy", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0.00")
                            Dim b25 As Binding = New Binding("Text", bsDetail, "ratioy1", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0.00")
                            Dim b26 As Binding = New Binding("Text", bsDetail, "ratioy2", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0.00")
                            Dim b27 As Binding = New Binding("Text", bsDetail, "ratioy3", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0.00")
                            Dim b28 As Binding = New Binding("Text", bsDetail, "ratioy4", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0.00")
                            TextBox19.DataBindings.Add(b19)
                            TextBox20.DataBindings.Add(b20)
                            TextBox21.DataBindings.Add(b21)
                            TextBox22.DataBindings.Add(b22)
                            TextBox23.DataBindings.Add(b23)
                            TextBox24.DataBindings.Add(b24)
                            TextBox25.DataBindings.Add(b25)
                            TextBox26.DataBindings.Add(b26)
                            TextBox27.DataBindings.Add(b27)
                            TextBox28.DataBindings.Add(b28)
                            'AddHandler b19.Parse, AddressOf onTextBoxBindingParse
                            'AddHandler b20.Parse, AddressOf onTextBoxBindingParse
                            'AddHandler b21.Parse, AddressOf onTextBoxBindingParse
                            'AddHandler b22.Parse, AddressOf onTextBoxBindingParse
                            'AddHandler b23.Parse, AddressOf onTextBoxBindingParse
                            'AddHandler b24.Parse, AddressOf onTextBoxBindingParse
                            'AddHandler b25.Parse, AddressOf onTextBoxBindingParse
                            'AddHandler b26.Parse, AddressOf onTextBoxBindingParse
                            'AddHandler b27.Parse, AddressOf onTextBoxBindingParse
                            'AddHandler b28.Parse, AddressOf onTextBoxBindingParse
                            'TextBox19.DataBindings.Add(New Binding("Text", bsDetail, "turnovery", True))
                            'TextBox20.DataBindings.Add(New Binding("Text", bsDetail, "turnovery1", True))
                            'TextBox21.DataBindings.Add(New Binding("Text", bsDetail, "turnovery2", True))
                            'TextBox22.DataBindings.Add(New Binding("Text", bsDetail, "turnovery3", True))
                            'TextBox23.DataBindings.Add(New Binding("Text", bsDetail, "turnovery4", True))
                            'TextBox24.DataBindings.Add(New Binding("Text", bsDetail, "ratioy", True))
                            'TextBox25.DataBindings.Add(New Binding("Text", bsDetail, "ratioy1", True))
                            'TextBox26.DataBindings.Add(New Binding("Text", bsDetail, "ratioy2", True))
                            'TextBox27.DataBindings.Add(New Binding("Text", bsDetail, "ratioy3", True))
                            'TextBox28.DataBindings.Add(New Binding("Text", bsDetail, "ratioy4", True))
                            DataGridView1.AutoGenerateColumns = False
                            DataGridView1.DataSource = bsDetail
                            DataGridView1.RowTemplate.Height = 22
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


    Private Sub bsheader_ListChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ListChangedEventArgs) Handles bsheader.ListChanged

        TextBox1.Enabled = Not IsNothing(bsheader.Current)
        TextBox2.Enabled = Not IsNothing(bsheader.Current)
        TextBox3.Enabled = Not IsNothing(bsheader.Current)
        TextBox4.Enabled = Not IsNothing(bsheader.Current)
        TextBox5.Enabled = Not IsNothing(bsheader.Current)
        TextBox6.Enabled = Not IsNothing(bsheader.Current)
        TextBox7.Enabled = Not IsNothing(bsheader.Current)
    End Sub

    Private Sub bsDetail_ListChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ListChangedEventArgs) Handles bsDetail.ListChanged
        TextBox8.Enabled = Not IsNothing(bsDetail.Current)
        TextBox9.Enabled = Not IsNothing(bsDetail.Current)
        TextBox10.Enabled = Not IsNothing(bsDetail.Current)
        TextBox11.Enabled = Not IsNothing(bsDetail.Current)
        TextBox12.Enabled = Not IsNothing(bsDetail.Current)
        TextBox13.Enabled = Not IsNothing(bsDetail.Current)
        TextBox14.Enabled = Not IsNothing(bsDetail.Current)
        TextBox15.Enabled = Not IsNothing(bsDetail.Current)
        TextBox16.Enabled = Not IsNothing(bsDetail.Current)
        TextBox17.Enabled = Not IsNothing(bsDetail.Current)
        TextBox18.Enabled = Not IsNothing(bsDetail.Current)
        TextBox19.Enabled = Not IsNothing(bsDetail.Current)
        TextBox20.Enabled = Not IsNothing(bsDetail.Current)
        TextBox21.Enabled = Not IsNothing(bsDetail.Current)
        TextBox22.Enabled = Not IsNothing(bsDetail.Current)
        TextBox23.Enabled = Not IsNothing(bsDetail.Current)
        TextBox24.Enabled = Not IsNothing(bsDetail.Current)
        RadioButton1.Enabled = Not IsNothing(bsDetail.Current)
        RadioButton2.Enabled = Not IsNothing(bsDetail.Current)
        
        If Not IsNothing(bsDetail.Current) Then
            ComboBox1.Enabled = RadioButton1.Checked
            ComboBox2.Enabled = RadioButton2.Checked
        Else
            ComboBox1.Enabled = Not IsNothing(bsDetail.Current)
            ComboBox2.Enabled = Not IsNothing(bsDetail.Current)
            
        End If        
        ComboBox3.Enabled = Not IsNothing(bsDetail.Current)
        ComboBox4.Enabled = Not IsNothing(bsDetail.Current)
        ComboBox5.Enabled = Not IsNothing(bsDetail.Current)
        DateTimePicker2.Enabled = Not IsNothing(bsDetail.Current)
        DateTimePicker3.Enabled = Not IsNothing(bsDetail.Current)
    End Sub


    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Dim mydrv As DataRowView = bsDetail.AddNew()
        ComboBox4.SelectedIndex = 1
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged, RadioButton2.CheckedChanged
        ComboBox1.Enabled = RadioButton1.Checked
        ComboBox2.Enabled = RadioButton2.Checked
        Button7.Enabled = RadioButton1.Checked
        Button8.Enabled = RadioButton2.Checked

        'If RadioButton2.Checked Then
        '    ComboBox1.SelectedIndex = -1
        'Else
        '    ComboBox2.SelectedIndex = -1
        'End If
    End Sub


    Private Sub ComboBox3_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        'General Contract
        enabledTextBox()

        
    End Sub

    Private Sub enabledTextBox()
        GroupBox1.Enabled = False
        GroupBox2.Enabled = False
        GroupBox3.Enabled = False
        GroupBox5.Enabled = False
        GroupBox6.Enabled = False
        GroupBox7.Enabled = False
        If Not IsNothing(ComboBox3.SelectedValue) Then
            Dim drv As DataRowView = ComboBox3.SelectedItem
            Select Case drv.Row.Item("doctypename") 'ComboBox3.SelectedText
                Case "General Contract"
                    GroupBox1.Enabled = True

                    'General Contract
                    'ComboBox5.SelectedIndex = -1


                    'supply chain Appendix
                    TextBox11.Text = ""
                    TextBox12.Text = ""

                    'Quality Appendix
                    TextBox13.Text = ""

                    'Social Audit
                    TextBox15.Text = ""
                    TextBox16.Text = ""
                    TextBox17.Text = ""

                    'SEF
                    TextBox18.Text = ""

                    'SIF
                    TextBox19.Text = ""
                    TextBox20.Text = ""
                    TextBox21.Text = ""
                    TextBox22.Text = ""
                    TextBox23.Text = ""
                    TextBox24.Text = ""
                    TextBox25.Text = ""
                    TextBox26.Text = ""
                    TextBox27.Text = ""
                    TextBox28.Text = ""

                Case "Supply Chain Appendix"
                    GroupBox2.Enabled = True
                    'General Contract
                    ComboBox5.SelectedIndex = -1


                    'supply chain Appendix
                    'TextBox11.Text = ""
                    'TextBox12.Text = ""

                    'Quality Appendix
                    TextBox13.Text = ""

                    'Social Audit
                    TextBox15.Text = ""
                    TextBox16.Text = ""
                    TextBox17.Text = ""

                    'SEF
                    TextBox18.Text = ""

                    'SIF
                    TextBox19.Text = ""
                    TextBox20.Text = ""
                    TextBox21.Text = ""
                    TextBox22.Text = ""
                    TextBox23.Text = ""
                    TextBox24.Text = ""
                    TextBox25.Text = ""
                    TextBox26.Text = ""
                    TextBox27.Text = ""
                    TextBox28.Text = ""
                Case "Quality Appendix"
                    GroupBox3.Enabled = True
                    'General Contract
                    ComboBox5.SelectedIndex = -1


                    'supply chain Appendix
                    TextBox11.Text = ""
                    TextBox12.Text = ""

                    'Quality Appendix
                    'TextBox13.Text = ""

                    'Social Audit
                    TextBox15.Text = ""
                    TextBox16.Text = ""
                    TextBox17.Text = ""

                    'SEF
                    TextBox18.Text = ""

                    'SIF
                    TextBox19.Text = ""
                    TextBox20.Text = ""
                    TextBox21.Text = ""
                    TextBox22.Text = ""
                    TextBox23.Text = ""
                    TextBox24.Text = ""
                    TextBox25.Text = ""
                    TextBox26.Text = ""
                    TextBox27.Text = ""
                    TextBox28.Text = ""
                Case "Social Audit"
                    GroupBox5.Enabled = True
                    'General Contract
                    ComboBox5.SelectedIndex = -1


                    'supply chain Appendix
                    TextBox11.Text = ""
                    TextBox12.Text = ""

                    'Quality Appendix
                    TextBox13.Text = ""

                    'Social Audit
                    'TextBox15.Text = ""
                    'TextBox16.Text = ""
                    'TextBox17.Text = ""

                    'SEF
                    TextBox18.Text = ""

                    'SIF
                    TextBox19.Text = ""
                    TextBox20.Text = ""
                    TextBox21.Text = ""
                    TextBox22.Text = ""
                    TextBox23.Text = ""
                    TextBox24.Text = ""
                    TextBox25.Text = ""
                    TextBox26.Text = ""
                    TextBox27.Text = ""
                    TextBox28.Text = ""
                Case "SEF"
                    GroupBox6.Enabled = True
                    'General Contract
                    ComboBox5.SelectedIndex = -1


                    'supply chain Appendix
                    TextBox11.Text = ""
                    TextBox12.Text = ""

                    'Quality Appendix
                    TextBox13.Text = ""

                    'Social Audit
                    TextBox15.Text = ""
                    TextBox16.Text = ""
                    TextBox17.Text = ""

                    'SEF
                    'TextBox18.Text = ""

                    'SIF
                    TextBox19.Text = ""
                    TextBox20.Text = ""
                    TextBox21.Text = ""
                    TextBox22.Text = ""
                    TextBox23.Text = ""
                    TextBox24.Text = ""
                    TextBox25.Text = ""
                    TextBox26.Text = ""
                    TextBox27.Text = ""
                    TextBox28.Text = ""
                Case "SIF"
                    GroupBox7.Enabled = True
                    'General Contract
                    ComboBox5.SelectedIndex = -1


                    'supply chain Appendix
                    TextBox11.Text = ""
                    TextBox12.Text = ""

                    'Quality Appendix
                    TextBox13.Text = ""

                    'Social Audit
                    TextBox15.Text = ""
                    TextBox16.Text = ""
                    TextBox17.Text = ""

                    'SEF
                    TextBox18.Text = ""

                    'SIF
                    'TextBox19.Text = ""
                    'TextBox20.Text = ""
                    'TextBox21.Text = ""
                    'TextBox22.Text = ""
                    'TextBox23.Text = ""
                    'TextBox24.Text = ""
                    'TextBox25.Text = ""
                    'TextBox26.Text = ""
                    'TextBox27.Text = ""
                    'TextBox28.Text = ""
                Case Else
                    ComboBox5.SelectedIndex = -1


                    'supply chain Appendix
                    TextBox11.Text = ""
                    TextBox12.Text = ""

                    'Quality Appendix
                    TextBox13.Text = ""

                    'Social Audit
                    TextBox15.Text = ""
                    TextBox16.Text = ""
                    TextBox17.Text = ""

                    'SEF
                    TextBox18.Text = ""

                    'SIF
                    TextBox19.Text = ""
                    TextBox20.Text = ""
                    TextBox21.Text = ""
                    TextBox22.Text = ""
                    TextBox23.Text = ""
                    TextBox24.Text = ""
                    TextBox25.Text = ""
                    TextBox26.Text = ""
                    TextBox27.Text = ""
                    TextBox28.Text = ""
            End Select
        End If
    End Sub
    Private Sub ComboBox3_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectionChangeCommitted


        If Not IsNothing(ComboBox3.SelectedValue) Then
            Dim drv As DataRowView = ComboBox3.SelectedItem
            'If drv.Row.Item("doctypename") <> "" Then
            If Not IsNothing(bsDetail.Current) Then
                Dim mydrv As DataRowView = bsDetail.Current
                mydrv.Row.BeginEdit()
                If IsDBNull(mydrv.Row.Item("doctypename")) Then
                    mydrv.Row.Item("doctypename") = drv.Row.Item("doctypename")
                Else
                    If Not mydrv.Row.Item("doctypename") = drv.Row.Item("doctypename") Then
                        mydrv.Row.Item("doctypename") = drv.Row.Item("doctypename")
                    End If
                End If

                DataGridView1.Invalidate()
            End If

            'End If
            'If ComboBox3.Text <> "" Then
            '    If Not IsNothing(bsDetail.Current) Then
            '        Dim mydrv As DataRowView = bsDetail.Current
            '        If IsDBNull(mydrv.Row.Item("doctypename")) Then
            '            mydrv.Row.Item("doctypename") = ComboBox3.Text
            '        Else
            '            If Not mydrv.Row.Item("doctypename") = ComboBox3.Text Then
            '                mydrv.Row.Item("doctypename") = ComboBox3.Text
            '            End If
            '        End If

            '        DataGridView1.Invalidate()
            '    End If
            'End If


            'Select Case drv.Row.Item("doctypename")


            'Select Case drv.Row.Item("doctypename") 'ComboBox3.SelectedText
            '    Case "General Contract"
            '        GroupBox1.Enabled = True

            '        'General Contract
            '        'ComboBox5.SelectedIndex = -1


            '        'supply chain Appendix
            '        TextBox11.Text = ""
            '        TextBox12.Text = ""

            '        'Quality Appendix
            '        TextBox13.Text = ""

            '        'Social Audit
            '        TextBox15.Text = ""
            '        TextBox16.Text = ""
            '        TextBox17.Text = ""

            '        'SEF
            '        TextBox18.Text = ""

            '        'SIF
            '        TextBox19.Text = ""
            '        TextBox20.Text = ""
            '        TextBox21.Text = ""
            '        TextBox22.Text = ""
            '        TextBox23.Text = ""
            '        TextBox24.Text = ""
            '        TextBox25.Text = ""
            '        TextBox26.Text = ""
            '        TextBox27.Text = ""
            '        TextBox28.Text = ""

            '    Case "Supply Chain Appendix"
            '        GroupBox2.Enabled = True
            '        'General Contract
            '        ComboBox5.SelectedIndex = -1


            '        'supply chain Appendix
            '        'TextBox11.Text = ""
            '        'TextBox12.Text = ""

            '        'Quality Appendix
            '        TextBox13.Text = ""

            '        'Social Audit
            '        TextBox15.Text = ""
            '        TextBox16.Text = ""
            '        TextBox17.Text = ""

            '        'SEF
            '        TextBox18.Text = ""

            '        'SIF
            '        TextBox19.Text = ""
            '        TextBox20.Text = ""
            '        TextBox21.Text = ""
            '        TextBox22.Text = ""
            '        TextBox23.Text = ""
            '        TextBox24.Text = ""
            '        TextBox25.Text = ""
            '        TextBox26.Text = ""
            '        TextBox27.Text = ""
            '        TextBox28.Text = ""
            '    Case "Quality Appendix"
            '        GroupBox3.Enabled = True
            '        'General Contract
            '        ComboBox5.SelectedIndex = -1


            '        'supply chain Appendix
            '        TextBox11.Text = ""
            '        TextBox12.Text = ""

            '        'Quality Appendix
            '        'TextBox13.Text = ""

            '        'Social Audit
            '        TextBox15.Text = ""
            '        TextBox16.Text = ""
            '        TextBox17.Text = ""

            '        'SEF
            '        TextBox18.Text = ""

            '        'SIF
            '        TextBox19.Text = ""
            '        TextBox20.Text = ""
            '        TextBox21.Text = ""
            '        TextBox22.Text = ""
            '        TextBox23.Text = ""
            '        TextBox24.Text = ""
            '        TextBox25.Text = ""
            '        TextBox26.Text = ""
            '        TextBox27.Text = ""
            '        TextBox28.Text = ""
            '    Case "Social Audit"
            '        GroupBox5.Enabled = True
            '        'General Contract
            '        ComboBox5.SelectedIndex = -1


            '        'supply chain Appendix
            '        TextBox11.Text = ""
            '        TextBox12.Text = ""

            '        'Quality Appendix
            '        TextBox13.Text = ""

            '        'Social Audit
            '        'TextBox15.Text = ""
            '        'TextBox16.Text = ""
            '        'TextBox17.Text = ""

            '        'SEF
            '        TextBox18.Text = ""

            '        'SIF
            '        TextBox19.Text = ""
            '        TextBox20.Text = ""
            '        TextBox21.Text = ""
            '        TextBox22.Text = ""
            '        TextBox23.Text = ""
            '        TextBox24.Text = ""
            '        TextBox25.Text = ""
            '        TextBox26.Text = ""
            '        TextBox27.Text = ""
            '        TextBox28.Text = ""
            '    Case "SEF"
            '        GroupBox6.Enabled = True
            '        'General Contract
            '        ComboBox5.SelectedIndex = -1


            '        'supply chain Appendix
            '        TextBox11.Text = ""
            '        TextBox12.Text = ""

            '        'Quality Appendix
            '        TextBox13.Text = ""

            '        'Social Audit
            '        TextBox15.Text = ""
            '        TextBox16.Text = ""
            '        TextBox17.Text = ""

            '        'SEF
            '        'TextBox18.Text = ""

            '        'SIF
            '        TextBox19.Text = ""
            '        TextBox20.Text = ""
            '        TextBox21.Text = ""
            '        TextBox22.Text = ""
            '        TextBox23.Text = ""
            '        TextBox24.Text = ""
            '        TextBox25.Text = ""
            '        TextBox26.Text = ""
            '        TextBox27.Text = ""
            '        TextBox28.Text = ""
            '    Case "SIF"
            '        GroupBox7.Enabled = True
            '        'General Contract
            '        ComboBox5.SelectedIndex = -1


            '        'supply chain Appendix
            '        TextBox11.Text = ""
            '        TextBox12.Text = ""

            '        'Quality Appendix
            '        TextBox13.Text = ""

            '        'Social Audit
            '        TextBox15.Text = ""
            '        TextBox16.Text = ""
            '        TextBox17.Text = ""

            '        'SEF
            '        TextBox18.Text = ""

            '        'SIF
            '        'TextBox19.Text = ""
            '        'TextBox20.Text = ""
            '        'TextBox21.Text = ""
            '        'TextBox22.Text = ""
            '        'TextBox23.Text = ""
            '        'TextBox24.Text = ""
            '        'TextBox25.Text = ""
            '        'TextBox26.Text = ""
            '        'TextBox27.Text = ""
            '        'TextBox28.Text = ""
            '    Case Else
            '        ComboBox5.SelectedIndex = -1


            '        'supply chain Appendix
            '        TextBox11.Text = ""
            '        TextBox12.Text = ""

            '        'Quality Appendix
            '        TextBox13.Text = ""

            '        'Social Audit
            '        TextBox15.Text = ""
            '        TextBox16.Text = ""
            '        TextBox17.Text = ""

            '        'SEF
            '        TextBox18.Text = ""

            '        'SIF
            '        TextBox19.Text = ""
            '        TextBox20.Text = ""
            '        TextBox21.Text = ""
            '        TextBox22.Text = ""
            '        TextBox23.Text = ""
            '        TextBox24.Text = ""
            '        TextBox25.Text = ""
            '        TextBox26.Text = ""
            '        TextBox27.Text = ""
            '        TextBox28.Text = ""
            'End Select
        End If
    End Sub
    Private Sub ComboBox4_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        GroupBox4.Enabled = False
        If Not IsNothing(ComboBox4.SelectedValue) Then
            Dim drv As DataRowView = ComboBox4.SelectedItem
            Dim mydrv As DataRowView = bsDetail.Current

            Select Case drv.Row.Item("levelname")
                Case "Project"
                    GroupBox4.Enabled = True
                Case Else
                    'mydrv.Row.BeginEdit()
                    TextBox14.Text = ""
            End Select
        End If
    End Sub
    Private Sub ComboBox4_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectionChangeCommitted

        If Not IsNothing(ComboBox4.SelectedValue) Then
            Dim drv As DataRowView = ComboBox4.SelectedItem
            'If drv.Row.Item("levelname") <> "" Then
            If Not IsNothing(bsDetail.Current) Then
                Dim mydrv As DataRowView = bsDetail.Current
                mydrv.Row.BeginEdit()
                If IsDBNull(mydrv.Row.Item("levelname")) Then
                    mydrv.Row.Item("levelname") = drv.Row.Item("levelname")
                Else
                    If Not mydrv.Row.Item("levelname") = drv.Row.Item("levelname") Then
                        mydrv.Row.Item("levelname") = drv.Row.Item("levelname")
                    End If
                End If
                If mydrv.Row.Item("levelname") <> "Project" Then
                    'mydrv.Row.Item("projectname") = ""
                    TextBox14.Text = ""
                End If
                DataGridView1.Invalidate()
            End If
            'End If
            'GroupBox4.Enabled = False
            Select Case drv.Row.Item("levelname")
                Case "Project"
                    GroupBox4.Enabled = True
                Case Else
                    TextBox14.Text = ""
            End Select


        End If
    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
       
            'If Not IsNothing(bsDetail.Current) Then
            '    Dim drvd As DataRowView = bsDetail.Current
            '    If Not drvd.Row.RowState = DataRowState.Unchanged Then
            '        ComboBox1.SelectedIndex = -1
            '        If Not IsNothing(bsDetail.Current) Then
            '            Dim mydrv As DataRowView = bsDetail.Current
            '            If IsDBNull(mydrv.Row.Item("vendorname")) Then
            '                mydrv.Row.Item("vendorname") = drv.Row.Item("vendorname")
            '                mydrv.Row.Item("vendorcode") = drv.Row.Item("vendorcode")

            '            Else
            '                If Not mydrv.Row.Item("vendorname") = drv.Row.Item("vendorname") Then
            '                    mydrv.Row.Item("vendorname") = drv.Row.Item("vendorname")
            '                    mydrv.Row.Item("vendorcode") = drv.Row.Item("vendorcode")
            '                    mydrv.Row.Item("shortname") = ""
            '                End If
            '            End If
            '            DataGridView1.Invalidate()
            '        End If
            '    End If
            'End If

    End Sub
    Private Sub ComboBox2_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectionChangeCommitted
        If Not IsNothing(ComboBox2.SelectedItem) Then
            Dim drv As DataRowView = ComboBox2.SelectedItem
            'If drv.Row.Item("description") <> "" Then
            ComboBox1.SelectedIndex = -1
            If Not IsNothing(bsDetail.Current) Then
                Dim mydrv As DataRowView = bsDetail.Current
                mydrv.Row.BeginEdit()
                If IsDBNull(mydrv.Row.Item("vendorname")) Then
                    mydrv.Row.Item("vendorname") = drv.Row.Item("vendorname")
                    'mydrv.Row.Item("vendorcode") = drv.Row.Item("vendorcode")

                Else
                    If Not mydrv.Row.Item("vendorname") = drv.Row.Item("vendorname") Then
                        mydrv.Row.Item("vendorname") = drv.Row.Item("vendorname")
                        'mydrv.Row.Item("vendorcode") = drv.Row.Item("vendorcode")
                        'mydrv.Row.Item("shortname") = ""
                        mydrv.Row.Item("shortname") = DBNull.Value
                    End If
                End If
                DataGridView1.Invalidate()
            End If
            'End If
        End If
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub
    Private Sub ComboBox1_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectionChangeCommitted
        If Not IsNothing(ComboBox1.SelectedItem) Then
            Dim drv As DataRowView = ComboBox1.SelectedItem
            'If drv.Row.Item("shortname") <> "" Then
            ComboBox2.SelectedIndex = -1
            Dim mydrv As DataRowView = bsDetail.Current
            mydrv.Row.BeginEdit()
            'mydrv.Row.Item("vendorcode") = DBNull.Value
            mydrv.Row.Item("vendorname") = DBNull.Value
            mydrv.Row.Item("shortname") = drv.Row.Item("shortname")
            'End If
            DataGridView1.Invalidate()
        End If
    End Sub


    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        If Not IsNothing(bsDetail.Current) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                'Dim mydrv As DataRowView = bsDetail.Current
                'mydrv.Row.CancelEdit()
                bsDetail.RemoveCurrent()
            End If
        End If
        
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim openfiledialog1 As New OpenFileDialog
        If openfiledialog1.ShowDialog = DialogResult.OK Then
            Dim mydrv As DataRowView = bsDetail.Current
            mydrv.Row.Item("docname") = IO.Path.GetFileName(openfiledialog1.FileName)
            mydrv.Row.Item("docext") = IO.Path.GetExtension(openfiledialog1.FileName)
            TextBox10.Text = openfiledialog1.FileName
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click, Button2.Click, Button3.Click, Button4.Click, Button5.Click
        Dim myobj = CType(sender, Button)
        Dim bs As New BindingSource
        bs.DataSource = DS.Tables(8)
        Dim myform As New FormGetValidator(bs)
        If myform.ShowDialog = DialogResult.OK Then
            Dim drv As DataRowView = bs.Current
            Dim myrowhd As DataRowView = bsheader.Current

            Select Case myobj.Name
                Case "Button1"
                    TextBox2.Text = drv.Row.Item("name")
                    validatorid = drv.Row.Item("userid")
                    myrowhd.Row.Item("validator") = validatorid.ToLower
                    myrowhd.Row.Item("validatorname") = drv.Row.Item("name")
                Case "Button2"
                    TextBox3.Text = drv.Row.Item("name")
                    cc1id = drv.Row.Item("userid")
                    myrowhd.Row.Item("cc1") = cc1id.ToLower
                    myrowhd.Row.Item("cc1name") = drv.Row.Item("name")
                Case "Button3"
                    TextBox4.Text = drv.Row.Item("name")
                    cc2id = drv.Row.Item("userid")
                    myrowhd.Row.Item("cc2") = IIf(cc2id.ToLower = "", DBNull.Value, cc2id.ToLower)
                    myrowhd.Row.Item("cc2name") = drv.Row.Item("name")
                Case "Button4"
                    TextBox5.Text = drv.Row.Item("name")
                    cc3id = drv.Row.Item("userid")
                    myrowhd.Row.Item("cc3") = IIf(cc3id.ToLower = "", DBNull.Value, cc3id.ToLower)
                    myrowhd.Row.Item("cc3name") = drv.Row.Item("name")
                Case "Button5"
                    TextBox6.Text = drv.Row.Item("name")
                    cc4id = drv.Row.Item("userid")
                    myrowhd.Row.Item("cc4") = IIf(cc4id.ToLower = "", DBNull.Value, cc4id.ToLower)
                    myrowhd.Row.Item("cc4name") = drv.Row.Item("name")
            End Select
        End If
    End Sub



    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click, Button8.Click, Button9.Click, Button10.Click, Button11.Click
        Dim myobj As Button = CType(sender, Button)
        If Not IsNothing(bsDetail.Current) Then
            Dim drv As DataRowView = bsDetail.Current
            drv.Row.BeginEdit()
            Select Case myobj.Name
                Case "Button7"
                    ComboBox1.SelectedIndex = -1
                    drv.Row.Item("shortname") = DBNull.Value
                Case "Button8"
                    ComboBox2.SelectedIndex = -1
                    'drv.Row.Item("vendorcode") = DBNull.Value
                    drv.Row.Item("vendorname") = DBNull.Value
                Case "Button9"
                    ComboBox3.SelectedIndex = -1
                    drv.Row.Item("doctypename") = DBNull.Value
                    'drv.Row.Item("doctypeid") = DBNull.Value
                    ComboBox5.SelectedIndex = -1
                    'drv.Row.Item("paymentcode") = DBNull.Value
                    TextBox11.Text = ""
                    TextBox12.Text = ""
                    TextBox13.Text = ""

                    TextBox15.Text = ""
                    TextBox16.Text = ""
                    TextBox17.Text = ""
                    TextBox18.Text = ""
                    TextBox19.Text = ""
                    TextBox20.Text = ""
                    TextBox21.Text = ""
                    TextBox22.Text = ""
                    TextBox23.Text = ""
                    TextBox24.Text = ""
                    TextBox25.Text = ""
                    TextBox26.Text = ""
                    TextBox27.Text = ""
                    TextBox28.Text = ""



                Case "Button10"
                    ComboBox4.SelectedIndex = -1
                    drv.Row.Item("levelname") = DBNull.Value
                    TextBox14.Text = ""


                Case "Button11"
                    ComboBox5.SelectedIndex = -1
                    'drv.Row.Item("paymentcode") = DBNull.Value

            End Select
            DataGridView1.Invalidate()
        End If

    End Sub

    Private Sub TextBox14_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox14.TextChanged
        DataGridView1.Invalidate()
    End Sub

    Private Sub onTextBoxBindingParse(ByVal sender As Object, ByVal e As ConvertEventArgs)
        If (IsDBNull(e.Value)) Then
            e.Value = String.Empty
        ElseIf (e.Value = "") Then
            e.Value = DBNull.Value
        End If
    End Sub




End Class