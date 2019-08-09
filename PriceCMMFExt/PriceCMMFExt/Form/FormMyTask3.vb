Imports System.Threading
Imports System.Text
Imports PriceCMMFExt.PublicClass
Imports PriceCMMFExt.SharedClass
Public Class FormMyTask3
    Private Enum TaskStatus
        StatusDraft = 1
        StatusNew = 2
        StatusRejected = 3
        StatusReSubmit = 4
        StatusValidated = 5
        StatusCancelled = 6
        StatusCompleted = 7
    End Enum
    Dim limit As String = " limit 1"
    Dim startdate As Date
    Dim enddate As Date
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)

    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim myThread2 As New System.Threading.Thread(AddressOf DoLoadCMMF)
    Dim bsheader As New BindingSource
    Dim bshistory As New BindingSource
    Dim DS As DataSet
    Dim DS2 As DataSet
    Dim sb As New StringBuilder
    Dim creator As String
    Dim myuser As String = String.Empty
    Dim bsVendorName As BindingSource
    Dim bsShortName As BindingSource
    Dim VendornameDict As New Dictionary(Of Long, String)
    Dim ShortnameHT As New Hashtable
    Dim SupplierFilter As String = String.Empty

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        If Not myThread.IsAlive Then
            If creator = "" Then
                MessageBox.Show("You are not allowed to create new Task.")
                Exit Sub
            End If


            Dim myrow As DataRowView = bsheader.AddNew
            myrow.Row.Item("creatorname") = creator.Trim
            myrow.Row.Item("creator") = myuser
            myrow.Row.Item("submitdate") = Today.Date
            'myrow.Row.Item("negotiateddate") = Today.Date
            myrow.Row.Item("pricetype") = "FOB"
            myrow.Row.Item("status") = 2
            DS.Tables(0).Rows.Add(myrow.Row)

            Dim myform = New FormPriceChange2(bsheader, DS, DS2)
            myform.ToolStripButton4.Visible = False
            myform.ToolStripButton5.Visible = False
            myform.ToolStripButton1.Visible = False
            'myform.ComboBox1.SelectedIndex = 0
            If Not myform.ShowDialog = DialogResult.OK Then
                'MessageBox.Show("Add New One")
                'bsheader.CancelEdit()
                bsheader.RemoveCurrent()
            Else
                bsheader.EndEdit()

            End If
            loaddata()
        Else
            MessageBox.Show("Still loading... Please wait.")
        End If


    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim obj = DirectCast(sender, ComboBox)
        If obj.Text = "All" Then
            limit = ""
        Else
            limit = " limit " & obj.Text
        End If
        loaddata()
    End Sub

    Private Sub loaddata()
        startdate = DateTimePicker1.Value.Date
        enddate = DateTimePicker2.Value.Date
        myuser = HelperClass1.UserId.ToLower
        'MessageBox.Show(myuser)
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
        'Dim myFilter As String = String.Empty
        DS = New DataSet
        Dim mymessage As String = String.Empty
        sb.Clear()

        myuser = HelperClass1.UserId.ToLower
        'myuser = "as\cchiu"
        'sb.Append("with dtl  as (select pricechangehdid,count(0) as count from pricechangedtl dtl group by pricechangehdid order by pricechangehdid) select distinct tb.*,dtl.count,pcr.reasonname from sp_getmytasksshortname('" & myuser & "'::text,false) as tb(pricechangehdid bigint,creator character varying,creatorname character varying,validator1 character varying,validator1name character varying,validator2 character varying,validator2name character varying,pricetype character varying,description text,submitdate date,negotiateddate date,attachment text,status integer,statusname text,actiondate date,actionby character varying,reasonid integer,shortname text) left join dtl on dtl.pricechangehdid = tb.pricechangehdid  left join pricechangereason pcr on pcr.id = tb.reasonid order by tb.pricechangehdid desc;")
        'sb.Append("with dtl  as (select pricechangehdid,count(0) as count from pricechangedtl dtl group by pricechangehdid order by pricechangehdid) select distinct tb.*,dtl.count,pcr.reasonname from sp_getmytasks5shortname('" & myuser & "'::text,false) as tb(pricechangehdid bigint,creator character varying,creatorname character varying,validator1 character varying,validator1name character varying,validator2 character varying,validator2name character varying,validator3 character varying,validator3name character varying,pricetype character varying,description text,submitdate date,negotiateddate date,attachment text,status integer,statusname text,actiondate date,actionby character varying,reasonid integer,shortname text) left join dtl on dtl.pricechangehdid = tb.pricechangehdid  left join pricechangereason pcr on pcr.id = tb.reasonid order by tb.pricechangehdid desc;")
        sb.Append("with dtl  as (select pricechangehdid,count(0) as count from pricechangedtl dtl group by pricechangehdid order by pricechangehdid) select distinct tb.*,dtl.count,pcr.reasonname,ps.specialproject from sp_getmytasks5shortname01('" & myuser & "'::text,false) as tb(pricechangehdid bigint,creator character varying,creatorname character varying,validator1 character varying,validator1name character varying,validator2 character varying,validator2name character varying,validator3 character varying,validator3name character varying,pricetype character varying,description text,submitdate date,negotiateddate date,attachment text,status integer,statusname text,actiondate date,actionby character varying,reasonid integer,shortname text,specialprojectid integer) left join dtl on dtl.pricechangehdid = tb.pricechangehdid  left join pricechangereason pcr on pcr.id = tb.reasonid left join pricechangespecialproject ps on ps.id = tb.specialprojectid order by tb.pricechangehdid desc;")
        'sb.Append(String.Format("with dtl  as (select pricechangehdid,count(0) as count from pricechangedtl dtl group by pricechangehdid order by pricechangehdid) select distinct tb.*,dtl.count,pcr.reasonname,officersebname || ' on ' || actiondate as statuschangedby,exportfiledate::character varying  || ' ' || exportfileid::character varying as myexportfile from sp_getmytasks3shortname('{0}'::text,true,'{1:yyyy-MM-dd}'::date,'{2:yyyy-MM-dd}'::date,'{3}'::text) as tb(pricechangehdid bigint,creator character varying,creatorname character varying,validator1 character varying,validator1name character varying,validator2 character varying,validator2name character varying,pricetype character varying,description text,submitdate date,negotiateddate date,attachment text,status integer,statusname text,actiondate date,actionby character varying,reasonid integer,exportfiledate date,exportfileid bigint,shortname text,vendorcode text) left join dtl on dtl.pricechangehdid = tb.pricechangehdid left join officerseb o on o.userid = actionby left join pricechangereason pcr on pcr.id = tb.reasonid order by tb.pricechangehdid desc;", myuser, startdate, enddate, limit))
        'sb.Append(String.Format("with dtl  as (select pricechangehdid,count(0) as count from pricechangedtl dtl group by pricechangehdid order by pricechangehdid) select distinct tb.*,dtl.count,pcr.reasonname,officersebname || ' on ' || actiondate as statuschangedby,exportfiledate::character varying  || ' ' || exportfileid::character varying as myexportfile from sp_getmytasks4shortname('{0}'::text,true,'{1:yyyy-MM-dd}'::date,'{2:yyyy-MM-dd}'::date,'{3}'::text,'{4}'::text) as tb(pricechangehdid bigint,creator character varying,creatorname character varying,validator1 character varying,validator1name character varying,validator2 character varying,validator2name character varying,pricetype character varying,description text,submitdate date,negotiateddate date,attachment text,status integer,statusname text,actiondate date,actionby character varying,reasonid integer,exportfiledate date,exportfileid bigint,shortname text,vendorcode text) left join dtl on dtl.pricechangehdid = tb.pricechangehdid left join officerseb o on o.userid = actionby left join pricechangereason pcr on pcr.id = tb.reasonid order by tb.pricechangehdid desc;", myuser, startdate, enddate, limit, SupplierFilter))
        'sb.Append(String.Format("with dtl  as (select pricechangehdid,count(0) as count from pricechangedtl dtl group by pricechangehdid order by pricechangehdid) select distinct tb.*,dtl.count,pcr.reasonname,officersebname || ' on ' || actiondate as statuschangedby,exportfiledate::character varying  || ' ' || exportfileid::character varying as myexportfile from                    sp_getmytasks5shortname('{0}'::text,true,'{1:yyyy-MM-dd}'::date,'{2:yyyy-MM-dd}'::date,'{3}'::text,'{4}'::text) as tb(pricechangehdid bigint,creator character varying,creatorname character varying,validator1 character varying,validator1name character varying,validator2 character varying,validator2name character varying,validator3 character varying,validator3name character varying,pricetype character varying,description text,submitdate date,negotiateddate date,attachment text,status integer,statusname text,actiondate date,actionby character varying,reasonid integer,exportfiledate date,exportfileid bigint,shortname text,vendorcode text) left join dtl on dtl.pricechangehdid = tb.pricechangehdid left join officerseb o on o.userid = actionby left join pricechangereason pcr on pcr.id = tb.reasonid order by tb.pricechangehdid desc;", myuser, startdate, enddate, limit, SupplierFilter))
        sb.Append(String.Format("with dtl  as (select pricechangehdid,count(0) as count from pricechangedtl dtl group by pricechangehdid order by pricechangehdid) select distinct tb.*,dtl.count,pcr.reasonname,ps.specialproject,officersebname || ' on ' || actiondate as statuschangedby,exportfiledate::character varying  || ' ' || exportfileid::character varying as myexportfile from sp_getmytasks5shortname01('{0}'::text,true,'{1:yyyy-MM-dd}'::date,'{2:yyyy-MM-dd}'::date,'{3}'::text,'{4}'::text) as tb(pricechangehdid bigint,creator character varying,creatorname character varying,validator1 character varying,validator1name character varying,validator2 character varying,validator2name character varying,validator3 character varying,validator3name character varying,pricetype character varying,description text,submitdate date,negotiateddate date,attachment text,status integer,statusname text,actiondate date,actionby character varying,reasonid integer,exportfiledate date,exportfileid bigint,shortname text,vendorcode text,specialprojectid integer) left join dtl on dtl.pricechangehdid = tb.pricechangehdid left join officerseb o on o.userid = actionby left join pricechangereason pcr on pcr.id = tb.reasonid left join pricechangespecialproject ps on ps.id = tb.specialprojectid order by tb.pricechangehdid desc;", myuser, startdate, enddate, limit, SupplierFilter))
        'sb.Append("select distinct teamtitleshortname || ' - ' || officersebname as name,lower(o.userid) as userid,tt.teamtitleid,officersebname,tt.teamtitleshortname from doc.user u left join officerseb o on o.userid = u.userid left join teamtitle tt on tt.teamtitleid = o.teamtitleid where teamtitleshortname in ('PD','SPM','PM','PO') and o.isactive and o.userid <> 'as\lili2' order by tt.teamtitleid,officersebname;")
        sb.Append("select distinct teamtitleshortname || ' - ' || mu.username as name,lower(mu.userid) as userid," &
                  " tt.teamtitleid,mu.username as officersebname,tt.teamtitleshortname " &
                  " from doc.user u " &
                  " left join officerseb o on o.userid = u.userid " &
                  " left join masteruser mu on mu.id = o.muid" &
                  " left join teamtitle tt on tt.teamtitleid = o.teamtitleid " &
                  " where teamtitleshortname in ('PD','SPM','PM','PO','WMF','PCL') " &
                  " and u.isactive " &
                  " order by tt.teamtitleid,officersebname;")
        'sb.Append("select * from officerseb o left join teamtitle t on t.teamtitleid = o.teamtitleid where lower(userid) = '" & myuser & "' order by t.teamtitleid limit 1;")
        sb.Append("select * from officerseb o left join teamtitle t on t.teamtitleid = o.teamtitleid left join masteruser mu on mu.id = o.muid where lower(mu.userid) = '" & myuser & "' order by t.teamtitleid limit 1;")

        'sb.Append("with my as (select distinct pricechangehdid from sp_getmytasks('" & myuser & "'::text,false) as tb(pricechangehdid bigint,creator character varying,creatorname character varying,validator1 character varying,validator1name character varying,validator2 character varying,validator2name character varying,pricetype character varying,description text,submitdate date,negotiateddate date,attachment text,status integer,statusname text,actiondate date,actionby character varying,reasonid integer))," &
        '        " pls as ( SELECT foo.cmmf, foo.vendorcode, foo.plant, max(foo.validfrom) AS validfrom " &
        '        " FROM ( select pd.cmmf,pl.vendorcode,pp.plant,pl.validfrom from my left join pricechangedtl pd  on pd.pricechangehdid = my.pricechangehdid left join pricelist pl on pl.cmmf = pd.cmmf LEFT JOIN priceplantscale pp ON pp.pricelistid = pl.pricelistid WHERE pl.scaleqty = pp.scale::double precision  ) foo GROUP BY foo.cmmf, foo.vendorcode, foo.plant ) ," &
        '        " pr as (SELECT pl.pricelistid, pl.cmmf, pl.scaleqty, pp.amount, pp.perunit, pl.validfrom, pl.validto, pl.vendorcode, pl.currency, pp.plant, pp.id  from my left join pricechangedtl pd on pd.pricechangehdid = my.pricechangehdid	left join  pricelist pl on pl.cmmf = pd.cmmf  LEFT JOIN priceplantscale pp ON pp.pricelistid = pl.pricelistid WHERE pl.scaleqty = pp.scale::double precision )," &
        '        " std as (select cmmf, max(validfrom) as validfrom from standardcostad group by cmmf)" &
        '        " select distinct dt.*,c.commercialref,r.rangedesc,r.range,v.vendorname::character varying,v.shortname::character varying,materialdesc::character varying, pr.amount::numeric,pr.perunit::numeric,ad.planprice1,ad.per," &
        '        " (getdelta(dt.price,dt.pricingunit,ad.planprice1,ad.per)) as deltastd, " &
        '        " (getdelta(dt.price,dt.pricingunit,pr.amount::numeric,pr.perunit::numeric) ) as deltasap," &
        '         " getalert(dt.price,dt.pricingunit,pr.amount::numeric,pr.perunit::numeric,ad.planprice1,ad.per) as alert,pr.validfrom,doc.getvendorcurr(dt.vendorcode,dt.validon) as crcy  " &
        '        "   from pricechangedtl dt  " &
        '        " left join cmmf c on c.cmmf = dt.cmmf  inner join my on my.pricechangehdid = dt.pricechangehdid left join vendor v on v.vendorcode = dt.vendorcode left join range r on r.rangeid = c.rangeid " &
        '        " left join std on std.cmmf = c.cmmf" &
        '        " left join standardcostad ad on ad.cmmf = std.cmmf and ad.validfrom = std.validfrom" &
        '        " left join  pls on pls.cmmf = c.cmmf and pls.vendorcode = v.vendorcode and pls.plant = dt.plant" &
        '        " left join  pr on pr.cmmf = pls.cmmf and pr.validfrom = pls.validfrom and pr.vendorcode = v.vendorcode and pr.plant = dt.plant;")
        'sb.Append("with my as (select distinct pricechangehdid from sp_getmytasks4('" & myuser & "'::text,false) as tb(pricechangehdid bigint,creator character varying,creatorname character varying,validator1 character varying,validator1name character varying,validator2 character varying,validator2name character varying,validator3 character varying,validator3name character varying,pricetype character varying,description text,submitdate date,negotiateddate date,attachment text,status integer,statusname text,actiondate date,actionby character varying,reasonid integer))," &
        '       " pls as ( SELECT foo.cmmf, foo.vendorcode, foo.plant, max(foo.validfrom) AS validfrom " &
        '       " FROM ( select pd.cmmf,pl.vendorcode,pp.plant,pl.validfrom from my left join pricechangedtl pd  on pd.pricechangehdid = my.pricechangehdid left join pricelist pl on pl.cmmf = pd.cmmf LEFT JOIN priceplantscale pp ON pp.pricelistid = pl.pricelistid WHERE pl.scaleqty = pp.scale::double precision  ) foo GROUP BY foo.cmmf, foo.vendorcode, foo.plant ) ," &
        '       " pr as (SELECT pl.pricelistid, pl.cmmf, pl.scaleqty, pp.amount, pp.perunit, pl.validfrom, pl.validto, pl.vendorcode, pl.currency, pp.plant, pp.id  from my left join pricechangedtl pd on pd.pricechangehdid = my.pricechangehdid	left join  pricelist pl on pl.cmmf = pd.cmmf  LEFT JOIN priceplantscale pp ON pp.pricelistid = pl.pricelistid WHERE pl.scaleqty = pp.scale::double precision )," &
        '       " std as (select cmmf, max(validfrom) as validfrom from standardcostad group by cmmf)" &
        '       " select distinct dt.*,c.commercialref,r.rangedesc,r.range,v.vendorname::character varying,v.shortname::character varying,materialdesc::character varying, pr.amount::numeric,pr.perunit::numeric,ad.planprice1,ad.per," &
        '       " (getdelta(dt.price,dt.pricingunit,ad.planprice1,ad.per)) as deltastd, " &
        '       " (getdelta(dt.price,dt.pricingunit,pr.amount::numeric,pr.perunit::numeric) ) as deltasap," &
        '        " getalert(dt.price,dt.pricingunit,pr.amount::numeric,pr.perunit::numeric,ad.planprice1,ad.per) as alert,pr.validfrom,doc.getvendorcurr(dt.vendorcode,dt.validon) as crcy  " &
        '       "   from pricechangedtl dt  " &
        '       " left join cmmf c on c.cmmf = dt.cmmf  inner join my on my.pricechangehdid = dt.pricechangehdid left join vendor v on v.vendorcode = dt.vendorcode left join range r on r.rangeid = c.rangeid " &
        '       " left join std on std.cmmf = c.cmmf" &
        '       " left join standardcostad ad on ad.cmmf = std.cmmf and ad.validfrom = std.validfrom" &
        '       " left join  pls on pls.cmmf = c.cmmf and pls.vendorcode = v.vendorcode and pls.plant = dt.plant" &
        '       " left join  pr on pr.cmmf = pls.cmmf and pr.validfrom = pls.validfrom and pr.vendorcode = v.vendorcode and pr.plant = dt.plant;")
        sb.Append("with my as (select distinct pricechangehdid from sp_getmytasks5('" & myuser & "'::text,false) as tb(pricechangehdid bigint,creator character varying,creatorname character varying,validator1 character varying,validator1name character varying,validator2 character varying,validator2name character varying,validator3 character varying,validator3name character varying,pricetype character varying,description text,submitdate date,negotiateddate date,attachment text,status integer,statusname text,actiondate date,actionby character varying,reasonid integer))," &
               " pls as ( SELECT foo.cmmf, foo.vendorcode, foo.plant, max(foo.validfrom) AS validfrom " &
               " FROM ( select pd.cmmf,pl.vendorcode,pp.plant,pl.validfrom from my left join pricechangedtl pd  on pd.pricechangehdid = my.pricechangehdid left join pricelist pl on pl.cmmf = pd.cmmf LEFT JOIN priceplantscale pp ON pp.pricelistid = pl.pricelistid WHERE pl.scaleqty = pp.scale::double precision  ) foo GROUP BY foo.cmmf, foo.vendorcode, foo.plant ) ," &
               " pr as (SELECT pl.pricelistid, pl.cmmf, pl.scaleqty, pp.amount, pp.perunit, pl.validfrom, pl.validto, pl.vendorcode, pl.currency, pp.plant, pp.id  from my left join pricechangedtl pd on pd.pricechangehdid = my.pricechangehdid	left join  pricelist pl on pl.cmmf = pd.cmmf  LEFT JOIN priceplantscale pp ON pp.pricelistid = pl.pricelistid WHERE pl.scaleqty = pp.scale::double precision )," &
               " std as (select cmmf, max(validfrom) as validfrom from standardcostad group by cmmf)" &
               " select distinct dt.*,c.commercialref,r.rangedesc,r.range,v.vendorname::character varying,v.shortname::character varying,materialdesc::character varying, pr.amount::numeric,pr.perunit::numeric,ad.planprice1,ad.per," &
               " (getdelta(dt.price,dt.pricingunit,ad.planprice1,ad.per)) as deltastd, " &
               " (getdelta(dt.price,dt.pricingunit,pr.amount::numeric,pr.perunit::numeric) ) as deltasap," &
                " getalert(dt.price,dt.pricingunit,pr.amount::numeric,pr.perunit::numeric,ad.planprice1,ad.per) as alert,pr.validfrom,doc.getvendorcurr(dt.vendorcode,dt.validon) as crcy  " &
               "   from pricechangedtl dt  " &
               " left join cmmf c on c.cmmf = dt.cmmf  inner join my on my.pricechangehdid = dt.pricechangehdid left join vendor v on v.vendorcode = dt.vendorcode left join range r on r.rangeid = c.rangeid " &
               " left join std on std.cmmf = c.cmmf" &
               " left join standardcostad ad on ad.cmmf = std.cmmf and ad.validfrom = std.validfrom" &
               " left join  pls on pls.cmmf = c.cmmf and pls.vendorcode = v.vendorcode and pls.plant = dt.plant" &
               " left join  pr on pr.cmmf = pls.cmmf and pr.validfrom = pls.validfrom and pr.vendorcode = v.vendorcode and pr.plant = dt.plant;")

        sb.Append("select 5;")
        sb.Append("select dt.paramname from paramdt dt left join paramhd ph on ph.paramhdid = dt.paramhdid where ph.paramname = 'PriceType' order by paramdtid;")
        sb.Append("select 2;")
        sb.Append("select hd.cvalue from paramhd hd where paramname='PriceCmmfAttachmentFolder';")
        sb.Append("select id,reasonname,isactive from pricechangereason where isactive order by lineno;")
        sb.Append("select NULL::bigint as vendorcode,''::text as description, ''::text as vendorname union all (with vc as (select distinct vendorcode from pricechangedtl)" &
                  " select vc.vendorcode, vc.vendorcode || ' - ' || v.vendorname as description,v.vendorname from vc" &
                  " left join vendor v on v.vendorcode = vc.vendorcode order by vendorname);")
        sb.Append("select ''::text as shortname" &
                  " union all" &
                  " (with vc as (select distinct vendorcode from pricechangedtl) " &
                  " select distinct v.shortname from vc" &
                  " left join vendor v on v.vendorcode = vc.vendorcode order by v.shortname);")
        sb.Append("select Null::bigint as id,Null::text as specialproject,null::boolean as isactive union all(select id,specialproject,isactive from pricechangespecialproject where isactive order by lineno);")
        If DbAdapter1.TbgetDataSet(sb.ToString, DS, mymessage) Then
            Try
                If DS.Tables(3).Rows.Count > 0 Then
                    'creator = DS.Tables(3).Rows(0).Item("officersebname").ToString.Trim
                    creator = DS.Tables(3).Rows(0).Item("username").ToString.Trim
                End If
                DS.Tables(0).TableName = "Vendor"
                Dim idx5(0) As DataColumn
                idx5(0) = DS.Tables(5).Columns(0)
                DS.Tables(5).PrimaryKey = idx5
                DS.Tables(6).TableName = "PriceType"
                DS.Tables(8).TableName = "AttachmentFolder"
                DS.Tables(9).TableName = "Reason"

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
    Sub DoWorkOri()
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

        myuser = HelperClass1.UserId.ToLower
        'myuser = "as\tchen"
        'myuser = "as\elai"
        'myuser = "as\rlo"
        'myuser = "AS\afok".ToLower
        'myuser = "as\tckwok"
        'myuser = "AS\shxu".ToLower
        'myuser = "as\jdai"
        'myuser = "as\ovalance"
        'myuser = "AS\SCHAN".ToLower
        'myuser = "as\dlam".ToLower
        'myuser = "as\vhui".ToLower
        'myuser = "as\cchiu"
        'myuser = "as\eyang"
        'sb.Append("with dtl  as (select pricechangehdid,count(0) as count from pricechangedtl dtl group by pricechangehdid order by pricechangehdid) select distinct tb.*,dtl.count,pcr.reasonname from sp_getmytasks('" & myuser & "'::text,false) as tb(pricechangehdid bigint,creator character varying,creatorname character varying,validator1 character varying,validator1name character varying,validator2 character varying,validator2name character varying,pricetype character varying,description text,submitdate date,negotiateddate date,attachment text,status integer,statusname text,actiondate date,actionby character varying,reasonid integer,shortname text) left join dtl on dtl.pricechangehdid = tb.pricechangehdid  left join pricechangereason pcr on pcr.id = tb.reasonid order by tb.pricechangehdid desc;")
        sb.Append("with dtl  as (select pricechangehdid,count(0) as count from pricechangedtl dtl group by pricechangehdid order by pricechangehdid) select distinct tb.*,dtl.count,pcr.reasonname from sp_getmytasksshortname('" & myuser & "'::text,false) as tb(pricechangehdid bigint,creator character varying,creatorname character varying,validator1 character varying,validator1name character varying,validator2 character varying,validator2name character varying,pricetype character varying,description text,submitdate date,negotiateddate date,attachment text,status integer,statusname text,actiondate date,actionby character varying,reasonid integer,shortname text) left join dtl on dtl.pricechangehdid = tb.pricechangehdid  left join pricechangereason pcr on pcr.id = tb.reasonid order by tb.pricechangehdid desc;")
        sb.Append("with dtl  as (select pricechangehdid,count(0) as count from pricechangedtl dtl group by pricechangehdid order by pricechangehdid) select distinct tb.*,dtl.count,pcr.reasonname,officersebname || ' on ' || actiondate as statuschangedby,exportfiledate::character varying  || ' ' || exportfileid::character varying as myexportfile from sp_getmytasks2shortname('" & myuser & "'::text,true) as tb(pricechangehdid bigint,creator character varying,creatorname character varying,validator1 character varying,validator1name character varying,validator2 character varying,validator2name character varying,pricetype character varying,description text,submitdate date,negotiateddate date,attachment text,status integer,statusname text,actiondate date,actionby character varying,reasonid integer,exportfiledate date,exportfileid bigint,shortname text) left join dtl on dtl.pricechangehdid = tb.pricechangehdid left join officerseb o on o.userid = actionby left join pricechangereason pcr on pcr.id = tb.reasonid order by tb.pricechangehdid desc " & limit & ";")
        'sb.Append("select distinct teamtitleshortname || ' - ' || officersebname as name,lower(userid) as userid,tt.teamtitleid,officersebname,tt.teamtitleshortname from officerseb o left join teamtitle tt on tt.teamtitleid = o.teamtitleid where teamtitleshortname in ('PD','SPM','PM','PO') and isactive and userid <> 'as\lili2' order by tt.teamtitleid,officersebname;")
        sb.Append("select distinct teamtitleshortname || ' - ' || officersebname as name,lower(o.userid) as userid,tt.teamtitleid,officersebname,tt.teamtitleshortname from doc.user u left join officerseb o on o.userid = u.userid left join teamtitle tt on tt.teamtitleid = o.teamtitleid where teamtitleshortname in ('PD','SPM','PM','PO') and o.isactive and o.userid <> 'as\lili2' order by tt.teamtitleid,officersebname;")
        sb.Append("select * from officerseb o left join teamtitle t on t.teamtitleid = o.teamtitleid where lower(userid) = '" & myuser & "' order by t.teamtitleid limit 1;")
        'sb.Append("with my as (select distinct pricechangehdid from sp_getmytasks('" & myuser & "'::text,false) as tb(pricechangehdid bigint,creator character varying,creatorname character varying,validator1 character varying,validator1name character varying,validator2 character varying,validator2name character varying,pricetype character varying,description text,submitdate date,negotiateddate date,attachment text,status integer,statusname text,actiondate date,actionby character varying,reasonid integer))" &
        '          " , pl as (select cmmf, max(validfrom) as validfrom,vendorcode from pricelist  group by cmmf,vendorcode)" &
        '          " , std as (select cmmf, max(validfrom) as validfrom from standardcostad group by cmmf)" &
        '          " select dt.*,c.commercialref,r.rangedesc,r.range,v.vendorname::character varying,v.shortname::character varying,materialdesc::character varying, pr.amount::numeric,pr.perunit::numeric,ad.planprice1,ad.per," &
        '          " (getdelta(dt.price,dt.pricingunit,ad.planprice1,ad.per)) as deltastd, " &
        '          " (getdelta(dt.price,dt.pricingunit,pr.amount::numeric,pr.perunit::numeric) ) as deltasap," &
        '           " getalert(dt.price,dt.pricingunit,pr.amount::numeric,pr.perunit::numeric,ad.planprice1,ad.per) as alert,pr.validfrom " &
        '          "   from pricechangedtl dt  " &
        '          " left join cmmf c on c.cmmf = dt.cmmf  inner join my on my.pricechangehdid = dt.pricechangehdid left join vendor v on v.vendorcode = dt.vendorcode left join range r on r.rangeid = c.rangeid " &
        '          " left join std on std.cmmf = c.cmmf" &
        '          " left join standardcostad ad on ad.cmmf = std.cmmf and ad.validfrom = std.validfrom" &
        '          " left join pl on pl.cmmf = c.cmmf and pl.vendorcode = v.vendorcode" &
        '          " left join pricelistscale pr on pr.cmmf = pl.cmmf and pr.validfrom = pl.validfrom and pr.vendorcode = v.vendorcode and pr.plant = dt.plant;")
        '1st       " left join pricelist pr on pr.cmmf = pl.cmmf and pr.validfrom = pl.validfrom and pr.vendorcode = v.vendorcode;")
        'Replace pricelist pr with below code to get Plant
        sb.Append("with my as (select distinct pricechangehdid from sp_getmytasks('" & myuser & "'::text,false) as tb(pricechangehdid bigint,creator character varying,creatorname character varying,validator1 character varying,validator1name character varying,validator2 character varying,validator2name character varying,pricetype character varying,description text,submitdate date,negotiateddate date,attachment text,status integer,statusname text,actiondate date,actionby character varying,reasonid integer))" &
                 " , pl as (select cmmf, max(validfrom) as validfrom,vendorcode from pricelist  group by cmmf,vendorcode)" &
                 " , std as (select cmmf, max(validfrom) as validfrom from standardcostad group by cmmf)" &
                 " select dt.*,c.commercialref,r.rangedesc,r.range,v.vendorname::character varying,v.shortname::character varying,materialdesc::character varying, pr.amount::numeric,pr.perunit::numeric,ad.planprice1,ad.per," &
                 " (getdelta(dt.price,dt.pricingunit,ad.planprice1,ad.per)) as deltastd, " &
                 " (getdelta(dt.price,dt.pricingunit,pr.amount::numeric,pr.perunit::numeric) ) as deltasap," &
                  " getalert(dt.price,dt.pricingunit,pr.amount::numeric,pr.perunit::numeric,ad.planprice1,ad.per) as alert,pr.validfrom " &
                 "   from pricechangedtl dt  " &
                 " left join cmmf c on c.cmmf = dt.cmmf  inner join my on my.pricechangehdid = dt.pricechangehdid left join vendor v on v.vendorcode = dt.vendorcode left join range r on r.rangeid = c.rangeid " &
                 " left join std on std.cmmf = c.cmmf" &
                 " left join standardcostad ad on ad.cmmf = std.cmmf and ad.validfrom = std.validfrom" &
                 " left join pl on pl.cmmf = c.cmmf and pl.vendorcode = v.vendorcode" &
                 " left join pricelistscalelatest pls on pls.cmmf = c.cmmf and pls.vendorcode = v.vendorcode and pls.plant = dt.plant" &
                 " left join pricelistscale pr on pr.cmmf = pls.cmmf and pr.validfrom = pls.validfrom and pr.vendorcode = v.vendorcode and pr.plant = dt.plant;")


        sb.Append("with dup as (select commercialref from cmmf where length(commercialref) > 1" &
                  " group by commercialref" &
                  " having count(commercialref) = 1)" &
                  " select c.commercialref,c.cmmf from cmmf c inner join dup on dup.commercialref = c.commercialref;")
        sb.Append("select dt.paramname from paramdt dt left join paramhd ph on ph.paramhdid = dt.paramhdid where ph.paramname = 'PriceType' order by paramdtid;")
        'sb.Append("with my as (select distinct pricechangehdid from sp_getmytasks('" & myuser & "'::text,true) as tb(pricechangehdid bigint,creator character varying,creatorname character varying,validator1 character varying,validator1name character varying,validator2 character varying,validator2name character varying,pricetype character varying,description text,submitdate date,negotiateddate date,attachment text,status integer,statusname text,actiondate date,actionby character varying,reasonid integer))" &
        '          " , pl as (select cmmf, max(validfrom) as validfrom,vendorcode from pricelist  group by cmmf,vendorcode)" &
        '          " , std as (select cmmf, max(validfrom) as validfrom from standardcostad group by cmmf)" &
        '          " select dt.*,ad.planprice1,c.commercialref,r.rangedesc,r.range,v.vendorname::character varying,v.shortname::character varying,materialdesc::character varying, pr.amount::numeric,pr.perunit::numeric," &
        '          " (getdelta(dt.price,dt.pricingunit,ad.planprice1,ad.per)) as deltastd, " &
        '          " (getdelta(dt.price,dt.pricingunit,pr.amount::numeric,pr.perunit::numeric) ) as deltasap," &
        '           " getalert(dt.price,dt.pricingunit,pr.amount::numeric,pr.perunit::numeric,ad.planprice1,ad.per) as alert,pr.validfrom " &
        '          "   from pricechangedtl dt  " &
        '          " left join cmmf c on c.cmmf = dt.cmmf  inner join my on my.pricechangehdid = dt.pricechangehdid left join vendor v on v.vendorcode = dt.vendorcode left join range r on r.rangeid = c.rangeid " &
        '          " left join std on std.cmmf = c.cmmf" &
        '          " left join standardcostad ad on ad.cmmf = std.cmmf and ad.validfrom = std.validfrom" &
        '          " left join pl on pl.cmmf = c.cmmf and pl.vendorcode = v.vendorcode" &
        '          " left join pricelist pr on pr.cmmf = pl.cmmf and pr.validfrom = pl.validfrom and pr.vendorcode = v.vendorcode;")
        sb.Append("with my as (select distinct pricechangehdid,statusname from sp_getmytasks('" & myuser & "'::text,true) as tb(pricechangehdid bigint,creator character varying,creatorname character varying,validator1 character varying,validator1name character varying,validator2 character varying,validator2name character varying,pricetype character varying,description text,submitdate date,negotiateddate date,attachment text,status integer,statusname text,actiondate date,actionby character varying,reasonid integer))" &
                 " , std as (select cmmf, max(validfrom) as validfrom from standardcostad group by cmmf)" &
                 " select dt.*,ad.planprice1,c.commercialref,r.rangedesc,r.range,v.vendorname::character varying,v.shortname::character varying,materialdesc::character varying, " &
                 " (getpriceinfo(my.statusname,dt.cmmf,dt.vendorcode,dt.validon,dt.price,dt.pricingunit,ad.planprice1,ad.per)).* " &
                 "   from pricechangedtl dt  " &
                 " left join cmmf c on c.cmmf = dt.cmmf  inner join my on my.pricechangehdid = dt.pricechangehdid left join vendor v on v.vendorcode = dt.vendorcode left join range r on r.rangeid = c.rangeid " &
                 " left join std on std.cmmf = c.cmmf" &
                 " left join standardcostad ad on ad.cmmf = std.cmmf and ad.validfrom = std.validfrom;")
        sb.Append("select hd.cvalue from paramhd hd where paramname='PriceCmmfAttachmentFolder';")
        sb.Append("select id,reasonname from pricechangereason order by reasonname;")

        If DbAdapter1.TbgetDataSet(sb.ToString, DS, mymessage) Then
            Try
                If DS.Tables(3).Rows.Count > 0 Then
                    creator = DS.Tables(3).Rows(0).Item("officersebname").ToString.Trim
                End If
                DS.Tables(0).TableName = "Vendor"
                Dim idx5(0) As DataColumn
                idx5(0) = DS.Tables(5).Columns(0)
                DS.Tables(5).PrimaryKey = idx5
                DS.Tables(6).TableName = "PriceType"
                DS.Tables(8).TableName = "AttachmentFolder"
                DS.Tables(9).TableName = "Reason"

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
                            bshistory = New BindingSource
                            bsVendorName = New BindingSource
                            bsShortName = New BindingSource
                            Dim pk(0) As DataColumn
                            pk(0) = DS.Tables(0).Columns(0)
                            DS.Tables(0).PrimaryKey = pk
                            DS.Tables(0).Columns(0).AutoIncrement = True
                            DS.Tables(0).Columns(0).AutoIncrementSeed = 0
                            DS.Tables(0).Columns(0).AutoIncrementStep = -1
                            DS.Tables(0).TableName = "PriceChangeHD"

                            Dim pk4(0) As DataColumn
                            pk4(0) = DS.Tables(4).Columns(0)
                            DS.Tables(4).PrimaryKey = pk4
                            DS.Tables(4).Columns(0).AutoIncrement = True
                            DS.Tables(4).Columns(0).AutoIncrementSeed = 0
                            DS.Tables(4).Columns(0).AutoIncrementStep = -1
                            DS.Tables(4).TableName = "PriceChangeDtl"

                            Dim rel As DataRelation
                            Dim hcol As DataColumn
                            Dim dcol As DataColumn
                            'create relation ds.table(0) and ds.table(4)
                            hcol = DS.Tables(0).Columns("pricechangehdid") 'docemailhdid in table header
                            dcol = DS.Tables(4).Columns("pricechangehdid") 'docemailhdid in table dtl
                            rel = New DataRelation("hdrel", hcol, dcol)
                            DS.Relations.Add(rel)

                            DS.Tables(6).TableName = "PriceType"


                            bsheader.DataSource = DS.Tables(0)

                            bshistory.DataSource = DS.Tables(1)
                            DataGridView1.AutoGenerateColumns = False
                            DataGridView1.DataSource = bsheader
                            DataGridView2.AutoGenerateColumns = False
                            DataGridView2.DataSource = bshistory

                            bsVendorName.DataSource = DS.Tables(10)
                            bsShortName.DataSource = DS.Tables(11)
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

    Private Sub FormMyTask_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'loaddata()
        ComboBox1.SelectedIndex = 1
        ToolStripDropDownButton1.Visible = HelperClass1.UserInfo.isAdmin
    End Sub

    Private Sub DataGridView1_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DataGridView1.CellBeginEdit
        'MessageBox.Show(DataGridView1.Columns(e.ColumnIndex).HeaderText)
    End Sub

    Private Sub DataGridView1_EditingControlShowing(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles DataGridView1.EditingControlShowing
        Dim tb As DataGridViewTextBoxEditingControl = DirectCast(e.Control, DataGridViewTextBoxEditingControl)
        RemoveHandler (tb.KeyDown), AddressOf datagridviewTextBox_Keypdown
        AddHandler (tb.KeyDown), AddressOf datagridviewTextBox_Keypdown
    End Sub

    Private Sub datagridviewTextBox_Keypdown(ByVal sender As Object, ByVal e As KeyEventArgs)
        If e.KeyValue = 112 Then 'F1 
            MessageBox.Show("Help")

        End If
    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.Text = Me.Text & "-" & HelperClass1.UserId
        loadCMMF()
        'DateTimePicker1.Value = Date.Today.AddDays(-14)
    End Sub

    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        If Not myThread.IsAlive Then
            Dim myrow As DataRowView = bsheader.Current
            Dim myform = New FormPriceChange2(bsheader, DS, DS2, False)
            Select Case myrow.Row.Item("status")
                Case TaskStatus.StatusNew, TaskStatus.StatusReSubmit '2, 4
                    myform.ToolStripButton2.Visible = False
                    myform.ToolStripButton7.Visible = False
                    If myrow.Item("creator") = myuser Then
                        myform.ToolStripButton4.Visible = False
                        myform.ToolStripButton5.Visible = False
                    End If

                Case TaskStatus.StatusRejected '3
                    myform.ToolStripButton4.Visible = False
                    myform.ToolStripButton5.Visible = False
                Case TaskStatus.StatusValidated '5
                    myform.ToolStripButton2.Visible = False
                    myform.ToolStripButton4.Visible = False
                    myform.ToolStripButton5.Visible = False
                    myform.ToolStripButton7.Visible = False
            End Select

            If Not myform.ShowDialog = DialogResult.OK Then
                Try
                    bsheader.CancelEdit()
                Catch ex As Exception

                End Try

            Else
                bsheader.EndEdit()
            End If
            loaddata()
        End If

    End Sub

    Private Sub DataGridView2_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        If e.ColumnIndex <> -1 Then


            If DataGridView2.Columns(e.ColumnIndex).HeaderText = "" Then
                Dim drv As DataRowView = bshistory.Current
                If drv.Row.Item("status") = 6 Then 'drv.Row.Item("status") = 5 Or
                    MessageBox.Show("This record cannot be cancelled!")
                    Exit Sub
                End If
                If MessageBox.Show("Do you want to cancel this record?", "Cancel Record", MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.OK Then
                    'Dim drv As DataRowView = bshistory.Current
                    drv.Row.Item("status") = 6

                    drv.Row.Item("statusname") = "Cancelled"
                    drv.Row.Item("actiondate") = Date.Today
                    drv.Row.Item("actionby") = HelperClass1.UserId.ToLower
                    bshistory.EndEdit()

                    Dim ds2 As New DataSet
                    ds2 = DS.GetChanges

                    If Not IsNothing(ds2) Then
                        'ds2.Tables(0).Rows(0).Item("status") = 4
                        Dim mymessage As String = String.Empty
                        Dim ra As Integer
                        'reset sequence number
                        Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
                        If Not DbAdapter1.PriceChangeTx(Me, mye) Then
                            ProgressReport(1, mye.message)
                        End If
                        Me.DialogResult = DialogResult.OK
                        'Me.Close()
                    End If
                    'Me.Close()
                    loaddata()
                End If

            End If
        End If
    End Sub






    Private Sub DataGridView2_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellDoubleClick
        If Not myThread.IsAlive Then
            'Dim bhhistory As New BindingSource
            'bshistory.DataSource = DS.Tables(1)

            Dim myrow As DataRowView = bshistory.Current
            Dim myform = New FormPriceChange2(DS, bshistory)
            'myform.myUser = HelperClass1.UserId.ToLower
            'Select Case myrow.Row.Item("status")
            'Case 2, 4
            '    myform.ToolStripButton4.Visible = False
            '    myform.ToolStripButton5.Visible = False

            'Case 3
            '    myform.ToolStripButton2.Visible = False
            '    myform.ToolStripButton7.Visible = False
            'Case 5
            myform.ToolStripButton2.Visible = False
            myform.ToolStripButton4.Visible = False
            myform.ToolStripButton6.Visible = False
            myform.ToolStripButton5.Visible = False
            myform.ToolStripButton7.Visible = False
            myform.TextBox2.ReadOnly = True
            myform.TextBox2.BackColor = SystemColors.Window
            'myform.TextBox2.ForeColor = SystemColors.WindowText
            myform.Button2.Enabled = False
            myform.Button3.Enabled = False
            myform.DateTimePicker1.Enabled = False
            myform.DateTimePicker1.BackColor = SystemColors.Window
            myform.DateTimePicker2.Enabled = False
            myform.DateTimePicker2.BackColor = SystemColors.Window
            myform.ListBox1.Enabled = True
            myform.ListBox1.ContextMenuStrip = Nothing
            myform.ComboBox1.Enabled = False
            myform.ComboBox1.BackColor = SystemColors.Window
            'myform.ComboBox2.Enabled = False
            'myform.ComboBox2.BackColor = SystemColors.Window

            myform.DataGridView1.Columns("price").ReadOnly = True
            myform.DataGridView1.Columns("price").DefaultCellStyle.BackColor = Color.White
            myform.DataGridView1.Columns("column10").ReadOnly = True
            myform.DataGridView1.Columns("column10").DefaultCellStyle.BackColor = Color.White

            'End Select
            myform.ShowDialog()
            'bshistory.CancelEdit()
            'loaddata()
        End If

    End Sub

    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError

    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        If Not myThread.IsAlive Then
            Me.Validate()
            ComboBox1.Focus()
            loaddata()
        End If
    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim myform As New FormHistoryDetail
        If myform.ShowDialog() = DialogResult.OK Then
            loaddata()
        End If
    End Sub



    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        Dim fd As New FolderBrowserDialog
        fd.Description = "Select the folder."
        If fd.ShowDialog = DialogResult.OK Then
            Try
                FileIO.FileSystem.CopyFile(Application.StartupPath & "\templates\PriceCmmfTemplate.xlsx", fd.SelectedPath & "\PriceCmmfTemplate.xlsx", True)
                If MessageBox.Show("File Copied to " & fd.SelectedPath & "\PriceCmmfTemplate.xlsx. Open the file?", "Template", MessageBoxButtons.OKCancel) = DialogResult.OK Then
                    Process.Start(fd.SelectedPath & "\PriceCmmfTemplate.xlsx")
                End If
            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

        End If
    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click
        Dim myform = New FormPriceChangeReasonMaster
        myform.ShowDialog()

    End Sub

    Private Sub loadCMMF()
        If Not myThread2.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread2 = New Thread(AddressOf DoLoadCMMF)
            myThread2.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Private Sub DoLoadCMMF()
        ProgressReport(6, "Marquee")
        DS2 = New DataSet
        Dim mymessage As String = String.Empty
        sb.Clear()
        sb.Append("with dup as (select commercialref from cmmf where length(commercialref) > 1" &
                 " group by commercialref" &
                 " having count(commercialref) = 1)" &
                 " select c.commercialref,c.cmmf from cmmf c inner join dup on dup.commercialref = c.commercialref;")

        If DbAdapter1.TbgetDataSet(sb.ToString, DS2, mymessage) Then
            Try
                Dim idx5(0) As DataColumn
                idx5(0) = DS2.Tables(0).Columns(0)
                DS2.Tables(0).PrimaryKey = idx5

            Catch ex As Exception
                ProgressReport(1, "Loading Data. Error::" & ex.Message)
                ProgressReport(5, "Continuous")
                Exit Sub
            End Try
        Else
            ProgressReport(1, "Loading Data. Error::" & mymessage)
            ProgressReport(5, "Continuous")
            Exit Sub
        End If
        'ProgressReport(1, "Loading Data.Done!")
        'ProgressReport(5, "Continuous")
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click, Button3.Click
        Dim obj As Button = DirectCast(sender, Button)
        Dim sbFilter As New StringBuilder
        Dim sbFilterValue As New StringBuilder
        Select Case obj.Name
            Case "Button2"
                'sbFilter.Append(TextBox1.Text)
                Dim myform = New FormHelper(bsShortName)
                myform.DataGridView1.Columns(0).DataPropertyName = "shortname"
                If myform.ShowDialog = Windows.Forms.DialogResult.OK Then
                    For Each sel As DataGridViewRow In myform.DataGridView1.SelectedRows
                        If Not ShortnameHT.ContainsKey(sel.Cells(0).FormattedValue) Then
                            ShortnameHT.Add(sel.Cells(0).FormattedValue, sel.Cells(0).FormattedValue)
                        End If
                    Next

                    For Each myobj As Object In ShortnameHT
                        If sbFilter.Length > 0 Then
                            sbFilter.Append(",")
                            sbFilterValue.Append(",")
                        End If
                        sbFilter.Append(myobj.Value)
                        sbFilterValue.Append(myobj.key)
                    Next
                    TextBox1.Text = sbFilter.ToString
                    TextBox2.Text = ""
                    VendornameDict.Clear()

                    SupplierFilter = ExplodeCriteria("shortname", sbFilterValue.ToString)
                End If
            Case "Button3"
                'sbFilter.Append(TextBox2.Text)
                Dim myform = New FormHelper(bsVendorName)
                myform.DataGridView1.Columns(0).DataPropertyName = "description"
                myform.DataGridView1.Columns(1).DataPropertyName = "vendorcode"
                myform.DataGridView1.Columns(2).DataPropertyName = "vendorname"
                If myform.ShowDialog = Windows.Forms.DialogResult.OK Then
                    For Each sel As DataGridViewRow In myform.DataGridView1.SelectedRows
                        If Not VendornameDict.ContainsKey(sel.Cells(1).FormattedValue) Then
                            VendornameDict.Add(sel.Cells(1).FormattedValue, sel.Cells(2).FormattedValue)
                        End If
                    Next

                    For Each myobj As KeyValuePair(Of Long, String) In VendornameDict
                        If sbFilter.Length > 0 Then
                            sbFilter.Append(",")
                            sbFilterValue.Append(",")
                        End If
                        sbFilter.Append(myobj.Value)
                        sbFilterValue.Append(myobj.Key)
                    Next
                    TextBox2.Text = sbFilter.ToString
                    TextBox1.Text = ""
                    ShortnameHT.Clear()
                    SupplierFilter = ExplodeCriteria("vendorcode", sbFilterValue.ToString)
                End If
        End Select


    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        VendornameDict.Clear()
        ShortnameHT.Clear()
        SupplierFilter = ""
    End Sub

    Private Sub UpdateTehPriceChangeReasonToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UpdateTehPriceChangeReasonToolStripMenuItem.Click
        Dim myform = New FormUpdatePriceChangeReason
        myform.ShowDialog()
    End Sub


    Private Sub SpecialProjectMasterToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SpecialProjectMasterToolStripMenuItem.Click
        Dim myform = New FormSpecialProjectMaster
        myform.ShowDialog()
    End Sub

End Class