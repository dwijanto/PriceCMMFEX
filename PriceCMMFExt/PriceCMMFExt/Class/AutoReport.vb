Imports PriceCMMFExt.PublicClass
Imports System.Text
Imports System.Net.Mail
Imports System.Net.Mime

Public Class AutoReport
    Inherits Email
    Enum ReportType
        AutoPriceChangeSummary = 0
    End Enum

    Dim MyReportType As ReportType
    Public errorMessage As String = String.Empty
    Dim BS As BindingSource
    Dim DTBS As BindingSource
    Dim minDay As Integer = -1
    Public Sub New(ByRef _reportType As ReportType)
        Me.MyReportType = _reportType
        If IsNothing(DbAdapter1) Then
            DbAdapter1 = New DbAdapter
        End If
    End Sub

    Public Function Execute() As Boolean
        Dim myret = False
        Try
            If MyReportType = ReportType.AutoPriceChangeSummary Then
                runAutoPriceChangeSummary()
            End If
            myret = True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        Return myret
    End Function

    Private Sub runAutoPriceChangeSummary()
        'Get Data
        Dim myContent As String

        Dim validdate As Date = Today
        'Dim MinDay As Integer = -1
        If validdate.DayOfWeek = DayOfWeek.Monday Then
            minDay = -3
        End If
        Dim ds As New DataSet
        Dim sb As New StringBuilder
        sb.Append(String.Format("with dtl as (select pricechangehdid,count(0) as count from pricechangedtl group by pricechangehdid)" &
                     " select o.officersebname::character varying as validator ,getshortnamepricechangehd(ph.pricechangehdid) as supplier,r.reasonname,ph.description,ph.submitdate,d.count as countofcmmf,getstatusname(status) as status" &
                     " from pricechangehd ph " &
                     " left join officerseb o on o.userid = ph.validator1" &
                     " left join pricechangereason r on r.id = ph.reasonid" &
                     " left join dtl d on d.pricechangehdid = ph.pricechangehdid" &
                     " where submitdate = '{0:yyyy-MM-dd}'  and d.count > 0" &
                     " order by o.officersebname,supplier,reasonname,description;", Today.Date.AddDays(minDay)))

        sb.Append("select pd.* from paramdt pd" &
                  " left join paramhd ph on ph.paramhdid = pd.paramhdid" &
                  " where ph.paramname = 'AutoPriceChangeSummary' ")

        Dim sqlstr = sb.ToString

        If DbAdapter1.TbgetDataSet(sqlstr, ds, errorMessage) Then
            ds.Tables(0).TableName = "HD"

            Dim pk(0) As DataColumn
            pk(0) = ds.Tables(0).Columns("pricechangehdid")
            ds.Tables(0).PrimaryKey = pk

            BS = New BindingSource
            DTBS = New BindingSource
            BS.DataSource = ds.Tables(0)
            DTBS.DataSource = ds.Tables(1)

        End If

        If ds.Tables(0).Rows.Count > 0 Then
            'generate body
            Dim drv As DataRowView = DTBS.Current
            Dim dbdate As Date = drv.Item("ts")
            If dbdate.Date <> Today.Date Then
                myContent = getbodyMessage(ds)


                Dim htmlView As AlternateView = AlternateView.CreateAlternateViewFromString(String.Format("{0} <br>Or click the Price CMMF Ext icon on your desktop: <br><p> <img src=cid:myLogo> <br></p><p>Price CMMF Ex System Administrator</p></body></html>", myContent), Nothing, MediaTypeNames.Text.Html)

                Dim logo As New LinkedResource(Application.StartupPath & "\PriceCMMFEx.png")
                logo.ContentId = "myLogo"
                htmlView.LinkedResources.Add(logo)


                'send email
                Dim myrecepient = drv.Row.Item("cvalue").ToString.Split(",")
                'Me.sendto = drv.Row.Item("cvalue") '"dwijanto@yahoo.com"
                Me.sendto = myrecepient(0) '"dwijanto@yahoo.com"
                Me.cc = myrecepient(1)
                Me.isBodyHtml = True
                Me.sender = "no-reply@groupeseb.com"
                Me.subject = String.Format("Price CMMF Ex: Tasks summary. (Date : {0:dd-MMM-yyyy}) ", Today.Date.AddDays(minDay)) '"***Do not reply to this e-mail.***"
                Me.body = myContent

                Me.htmlView = htmlView

                If Not Me.send(errorMessage) Then
                    Logger.log(errorMessage)
                End If

                sqlstr = "update paramdt set ts = now() where" &
                         " paramdtid in (select paramdtid from paramdt pt " &
                         " left join paramhd ph on ph.paramhdid = pt.paramhdid" &
                         " where ph.paramname = 'AutoPriceChangeSummary')"

                If Not DbAdapter1.ExecuteScalar(sqlstr, message:=errorMessage) Then
                    Logger.log(errorMessage)
                End If
            End If



        End If

    End Sub

    Private Function getbodyMessage(ByVal data As Object) As String

        Dim hdbs As New BindingSource
        Dim dtbs As New BindingSource
        hdbs.DataSource = DirectCast(data, DataSet).Tables(1)
        dtbs.DataSource = DirectCast(data, DataSet).Tables(0)

        Dim sb As New StringBuilder


        sb.Append("<!DOCTYPE html><html><head><meta name=""description"" content=""[PriceCMMFEx]"" /><meta http-equiv=""Content-Type"" content=""text/html; charset=us-ascii""></head><style>  td,th {padding-left:5px;         padding-right:10px;         text-align:left;  }  th {background-color:red;    color:white}  .defaultfont{    font-size:11.0pt;	font-family:""Calibri"",""sans-serif"";    }</style><body class=""defaultfont"">")
        sb.Append(String.Format("<p>Dear {0},</p> <p>Please be informed that we have the following Price Change tasks:<br>Date: {1:dd-MMM-yyyy}<br><br>", DirectCast(hdbs.Current, DataRowView).Item("paramname"), Today.Date.AddDays(minDay)))
        'sb.Append("    List of Tasks:</p>  <table border=1 cellspacing=0>    <tr>            <th>Status</th>      <th>Reason</th>            <th>Description</th>      <th>Price Type</th>      <th>Submit Date</th>      <th>Creator</th>      <th>Validator</th>          </tr>")
        sb.Append("List of Tasks:</p>  <table border=1 cellspacing=0 class=""defaultfont"">    <tr>            <th>Validator</th>      <th>Supplier</th>      <th>Reason</th>            <th>Description</th>      <th>Submit Date</th>      <th>Count of CMMF</th>   <th>Status</th>        </tr>")
        For Each n As DataRowView In dtbs.List
            sb.Append(String.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4:yyyy-MMM-dd}</td><td>{5}</td><td>{6}</td></tr>", n.Item("validator"), n.Item("supplier"), n.Item("reasonname"), n.Item("description"), n.Item("submitdate"), n.Item("countofcmmf"), n.Item("status")))

        Next
        'sb.Append("</table>  <br>  <p>Thank you.<br><br>You can access the system by below link:<br>   <a href=""http://hon08nt"">MyTask</a></p><br><br><p>Price CMMF Ex System Administrator</p></body></html>")
        'sb.Append("</table>  <br>  <p>Thank you.<br><br>You can access the system by below link:<br>    <a href=""http://hon08nt"">PriceCMMFEX</a></p><p>Price CMMF Ex System Administrator</p></body></html>")
        sb.Append("</table>  <br>  <p>Thank you.<br><br>You can access the system in Citrix by below link:<br>    <a href=""https://sw07e601/RDWeb"">PriceCMMFEX</a></p>")
        Return sb.ToString
    End Function

End Class
