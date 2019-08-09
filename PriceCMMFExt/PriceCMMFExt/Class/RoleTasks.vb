Imports PriceCMMFExt.PublicClass
Imports System.Text
Imports System.Net.Mail
Imports System.Net.Mime

Public Class RoleTasks
    Inherits Email

    Enum Role
        Creator = 0
        Validator = 1
        CC = 2
        WMF = 3
    End Enum

    Private ds As DataSet
    Public errormessage As String = String.Empty
    Private HDBS As BindingSource
    Private DTBS As BindingSource
    Private myRole As Role
    Public Sub New(ByVal _role As Role)
        myRole = _role
        If IsNothing(DbAdapter1) Then
            DbAdapter1 = New DbAdapter
        End If

    End Sub

    Public Function Execute() As Boolean
        'Initialize Dataset
        Dim myret = False
        Dim mygroup As Object = vbNull
        Dim roleTask As Object = vbNull
        If InitDataset() Then
            If myRole = Role.Validator Then
                roleTask = New ValidatorBM(HDBS)
            ElseIf myRole = Role.Creator Then
                roleTask = New CreatorBM(HDBS)
            ElseIf myRole = Role.CC Then
                roleTask = New CCBM(HDBS)
            ElseIf myRole = Role.WMF Then
                roleTask = New WMFBM(HDBS)
            End If
            myret = True
        End If
        'Dim sendtoname As String = String.Empty

        For Each n In roleTask.GetQuery
            'Using select is correct. Because Sendto is inside the roleTask.GetQuery
            Me.sendto = Nothing
            Select Case myRole
                Case Role.Validator
                    If Not IsDBNull(n.data(0).item("validatoremail")) Then
                        Me.sendto = n.data(0).item("validatoremail")
                        'sendtoname = n.data(0).item("validator1name")
                    End If
                Case Role.Creator
                    If Not IsDBNull(n.data(0).item("creatoremail")) Then
                        Me.sendto = n.data(0).item("creatoremail")
                        'sendtoname = n.data(0).item("creatorname")
                    End If
                Case Role.CC
                    If Not IsDBNull(n.data(0).item("ccemail")) Then
                        Me.sendto = n.data(0).item("ccemail")
                        'sendtoname = n.data(0).item("validator2name")
                    End If
                Case Role.WMF
                    If Not IsDBNull(n.data(0).item("wmfemail")) Then
                        Me.sendto = n.data(0).item("wmfemail")
                        'sendtoname = n.data(0).item("validator2name")
                    End If
            End Select
            If Not IsNothing(Me.sendto) Then
                Dim myContent = roleTask.getbodymessage(n.data)


                Dim htmlView As AlternateView = AlternateView.CreateAlternateViewFromString(String.Format("{0} <br>Or click the Price CMMF Ext icon on your desktop: <br><p> <img src=cid:myLogo> <br></p><p>Price CMMF Ex System Administrator</p></body></html>", myContent), Nothing, MediaTypeNames.Text.Html)

                Dim logo As New LinkedResource(Application.StartupPath & "\PriceCMMFEx.png")
                logo.ContentId = "myLogo"
                htmlView.LinkedResources.Add(logo)

                Me.htmlView = htmlView


                Me.sendto = Me.sendto '"dwijanto@yahoo.com"
                Me.isBodyHtml = True
                Me.sender = "no-reply@groupeseb.com"

                Me.subject = String.Format("Price CMMF Ex: Tasks status. (Date : {0:dd-MMM-yyyy}) ", Today.Date) '"***Do not reply to this e-mail.***" "Price CMMF Ex: Tasks status. " '"***Do not reply to this e-mail.***"
                Me.body = myContent 'roleTask.getbodymessage(n.data)
                Me.htmlView = htmlView
                If Not Me.send(errormessage) Then
                    Logger.log(errormessage)
                End If
            End If
            'HDBS.EndEdit()
        Next
        Dim ds2 = ds.GetChanges
        'Update Changes
        If Not IsNothing(ds2) Then
            Dim mymessage As String = String.Empty
            Dim ra As Integer
            Dim mye As New ContentBaseEventArgs(ds2, True, errormessage, ra, True)
            If Not DbAdapter1.PriceChangeSendEmail(Me, mye) Then
                Logger.log(mye.message)
            Else
                ds.AcceptChanges()
            End If
        End If

        Return myret
    End Function

    Private Function InitDataset() As Boolean
        Dim myret = False
        'Dim sqlstr = "select * from pricechangehd;" &
        '             "select * from pricechangedtl;"
        'Dim sqlstr = "select pricechangehdid,creator,c.officersebname::character varying as creatorname,validator1,v1.officersebname::character varying as validator1name,validator2,v2.officersebname::character varying as validator2name,pricetype,description,submitdate,negotiateddate,attachment,status,getstatusname(status)::text as statusname,actiondate,actionby,sendstdvalidatedtocreator,sendtocc,sendcompletedtocreator" &
        '             " ,c.email as creatoremail,v1.email as validatoremail , v2.email as ccemail,reasonname" &
        '             " from  pricechangehd phd " &
        '             " left join officerseb c on lower(c.userid) = lower(creator)" &
        '             " left join officerseb v1 on lower(v1.userid) = lower(validator1)" &
        '             " left join officerseb v2 on lower(v2.userid) = lower(validator2)" &
        '             " left join pricechangereason pcr on pcr.id = phd.reasonid " &
        '             " where (sendstdvalidatedtocreator isnull or sendcompletedtocreator isnull or sendtocc isnull ) and status in (2,3,4,5,7) order by submitdate,statusname"
        Dim sqlstr = "select pricechangehdid,creator,c.officersebname::character varying as creatorname,validator1,v1.officersebname::character varying as validator1name,validator2,v2.officersebname::character varying as validator2name,validator3,v3.officersebname::character varying as validator3name,pricetype,description,submitdate,negotiateddate,attachment,status,getstatusname(status)::text as statusname,actiondate,actionby,sendstdvalidatedtocreator,sendtocc,sendcompletedtocreator,sendtowmf" &
                     " ,c.email as creatoremail,v1.email as validatoremail , v2.email as ccemail,v3.email as wmfemail,reasonname" &
                     " from  pricechangehd phd " &
                     " left join officerseb c on lower(c.userid) = lower(creator)" &
                     " left join officerseb v1 on lower(v1.userid) = lower(validator1)" &
                     " left join officerseb v2 on lower(v2.userid) = lower(validator2)" &
                     " left join officerseb v3 on lower(v3.userid) = lower(validator3)" &
                     " left join pricechangereason pcr on pcr.id = phd.reasonid " &
                     " where (sendstdvalidatedtocreator isnull or sendcompletedtocreator isnull or sendtocc isnull  or sendtowmf isnull) and status in (2,3,4,5,7) order by submitdate,statusname"


        ds = New DataSet
        If DbAdapter1.TbgetDataSet(sqlstr, ds, errormessage) Then
            ds.Tables(0).TableName = "HD"

            Dim pk(0) As DataColumn
            pk(0) = ds.Tables(0).Columns("pricechangehdid")
            ds.Tables(0).PrimaryKey = pk

            'Dim rel As DataRelation
            'Dim hcol As DataColumn
            'Dim dcol As DataColumn
            ''create relation ds.table(0) and ds.table(4)
            'hcol = ds.Tables(0).Columns("pricechangehdid")
            'dcol = ds.Tables(1).Columns("pricechangehdid")
            'rel = New DataRelation("hdrel", hcol, dcol)
            'ds.Relations.Add(rel)

            'HDBS = New BindingSource
            'DTBS = New BindingSource
            'HDBS.DataSource = ds.Tables(0)
            'DTBS.DataSource = HDBS
            'DTBS.DataMember = "hdrel"

            HDBS = New BindingSource
            HDBS.DataSource = ds.Tables(0)
            myret = True
        End If
        Return myret
    End Function

    Private Function getbodymessage(ByVal Data As Object, ByVal sendtoname As String) As String
        Return ""
    End Function

End Class





Public Class ValidatorBM
    Implements iRolePriceTask
    Private HDBS As BindingSource

    Public Sub New(ByVal HDBS As BindingSource)
        Me.HDBS = HDBS
        HDBS.Filter = "status = 2 or status = 4"
    End Sub

    Public Function GetQuery() As Object Implements iRolePriceTask.GetQuery
        Return From n In HDBS.List
                          Group n By key = n.item("validator1") Into Group
                          Select key, data = Group
    End Function

    Public Function GetBodyMessage(ByVal data As Object) As String Implements iRolePriceTask.GetBodyMessage
        'Dim sb As New StringBuilder
        'For Each n As DataRowView In data
        '    sb.Append(n.Item("submitdate"))

        'Next
        'Return sb.ToString
        Dim sb As New StringBuilder
        sb.Append("<!DOCTYPE html><html><head><meta name=""description"" content=""[PriceCMMFEx]"" /><meta http-equiv=""Content-Type"" content=""text/html; charset=us-ascii""></head><style>  td,th {padding-left:5px;         padding-right:10px;         text-align:left;  }  th {background-color:red;    color:white}  .defaultfont{    font-size:11.0pt;	font-family:""Calibri"",""sans-serif"";    }</style><body class=""defaultfont"">")
        sb.Append(String.Format("<p>Dear {0},</p><br><p>Please be informed that you have Price Change tasks that need to follow up.<br>Date: {1:dd-MMM-yyyy}<br><br><br>", data(0).row.item("validator1name").ToString, Today.Date))
        sb.Append("    List of Tasks:</p>  <table border=1 cellspacing=0>    <tr>            <th>Status</th>      <th>Reason</th>            <th>Description</th>      <th>Price Type</th>      <th>Submit Date</th>      <th>Creator</th>      <th>Validator</th>          </tr>")
        For Each n As DataRowView In data
            sb.Append(String.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4:yyyy-MMM-dd}</td><td>{5}</td><td>{6}</td></tr>", n.Item("statusname"), n.Item("reasonname"), n.Item("description"), n.Item("pricetype"), n.Item("submitdate"), n.Item("creatorname"), n.Item("validator1name")))

        Next
        'sb.Append("</table>  <br>  <p>Thank you.<br><br>You can access the system by below link:<br>   <a href=""http://hon08nt"">MyTask</a></p><br><br><p>Price CMMF Ex System Administrator</p></body></html>")
        sb.Append("</table>  <br>  <p>Thank you.<br><br>You can access the system in Citrix by below link:<br>   <a href=""https://sw07e601/RDWeb"">MyTask</a></p><br>")
        Return sb.ToString
    End Function


End Class

Public Class CreatorBM
    Implements iRolePriceTask
    Private HDBS As BindingSource
    Public Property getEmail

    Public Sub New(ByVal HDBS As BindingSource)
        Me.HDBS = HDBS
        'HDBS.Filter = "status = 3 or (status = 5 and pricetype = 'STD' and sendstdvalidatedtocreator is null ) or (status = 7 and sendcompletedtocreator is null)"
        'HDBS.Filter = "status = 3 or (status = 5 and pricetype = 'STD' and (not(sendstdvalidatedtocreator) or sendstdvalidatedtocreator is null )) or (status = 7 and (not(sendcompletedtocreator) or sendcompletedtocreator is null))"
        HDBS.Filter = "status = 3 or (status = 5  and (not(sendstdvalidatedtocreator) or sendstdvalidatedtocreator is null )) or (status = 7 and (not(sendcompletedtocreator) or sendcompletedtocreator is null))"

    End Sub

    Public Function GetQuery() As Object Implements iRolePriceTask.GetQuery
        Return From n In HDBS.List
                         Group n By key = n.item("creator") Into Group
                         Select key, data = Group
    End Function

    Public Function GetBodyMessage(ByVal data As Object) As String Implements iRolePriceTask.GetBodyMessage
        'Dim sb As New StringBuilder
        'For Each n As DataRowView In data
        '    sb.Append(n.Item("submitdate"))
        '    If n.Item("status") = 5 Then
        '        n.Item("sendtocreator") = True
        '    End If

        'Next
        'Return sb.ToString
        Dim sb As New StringBuilder
        sb.Append("<!DOCTYPE html><html><head><meta name=""description"" content=""[PriceCMMFEx]"" /><meta http-equiv=""Content-Type"" content=""text/html; charset=us-ascii""></head><style>  td,th {padding-left:5px;         padding-right:10px;         text-align:left;  }  th {background-color:red;    color:white}  .defaultfont{    font-size:11.0pt;	font-family:""Calibri"",""sans-serif"";    }</style><body class=""defaultfont"">")
        'sb.Append(String.Format("<p>Dear {0},</p><br><p>Please be informed that you have Price Change tasks that need to follow up.<br><br>", data(0).item("creatorname").ToString))
        sb.Append(String.Format("<p>Dear {0},</p><br><p>Please be informed that you have Price Change tasks that need to follow up.<br>Date: {1:dd-MMM-yyyy}<br><br><br>", data(0).row.item("creatorname").ToString, Today.Date))
        sb.Append("    List of Tasks:</p>  <table border=1 cellspacing=0>    <tr>            <th>Status</th>      <th>Reason</th>            <th>Description</th>      <th>Price Type</th>      <th>Submit Date</th>      <th>Creator</th>      <th>Validator</th>          </tr>")
        For Each n As DataRowView In data
            sb.Append(String.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4:yyyy-MMM-dd}</td><td>{5}</td><td>{6}</td></tr>", n.Item("statusname"), n.Item("reasonname"), n.Item("description"), n.Item("pricetype"), n.Item("submitdate"), n.Item("creatorname"), n.Item("validator1name")))
            If n.Item("status") = 5 Then
                n.Item("sendstdvalidatedtocreator") = True
            End If
            If n.Item("status") = 7 Then
                n.Item("sendcompletedtocreator") = True
            End If
        Next
        HDBS.EndEdit()
        'sb.Append("</table>  <br>  <p>Thank you.<br><br>You can access the system by below link:<br>   <a href=""http://hon08nt"">MyTask</a></p><br><br><p>Price CMMF Ex System Administrator</p></body></html>")
        sb.Append("</table>  <br>  <p>Thank you.<br><br>You can access the system in Citrix by below link:<br>   <a href=""https://sw07e601/RDWeb"">MyTask</a></p><br>")
        Return sb.ToString
    End Function

End Class

Public Class CCBM
    Implements iRolePriceTask
    Private HDBS As BindingSource
    Public Sub New(ByVal HDBS As BindingSource)
        Me.HDBS = HDBS
        HDBS.Filter = "status = 7 and (sendtocc is null or not(sendtocc))"

    End Sub

    Public Function GetQuery() As Object Implements iRolePriceTask.GetQuery
        Return From n In HDBS.List
                         Group n By key = n.item("validator2") Into Group
                         Select key, data = Group
    End Function

    Public Function GetBodyMessage(ByVal data As Object) As String Implements iRolePriceTask.GetBodyMessage
        'Dim sb As New StringBuilder
        'For Each n As DataRowView In data
        '    sb.Append(n.Item("submitdate"))

        'Next
        'Return sb.ToString
        Dim sb As New StringBuilder
        sb.Append("<!DOCTYPE html><html><head><meta name=""description"" content=""[PriceCMMFEx]"" /><meta http-equiv=""Content-Type"" content=""text/html; charset=us-ascii""></head><style>  td,th {padding-left:5px;         padding-right:10px;         text-align:left;  }  th {background-color:red;    color:white}  .defaultfont{    font-size:11.0pt;	font-family:""Calibri"",""sans-serif"";    }</style><body class=""defaultfont"">")
        'sb.Append(String.Format("<p>Dear {0},</p><br><p>Please be informed that you have Price Change tasks with Completed status.<br><br>", data(0).item("validator2name")))
        sb.Append(String.Format("<p>Dear {0},</p><br><p>Please be informed that you have Price Change tasks with Completed status.<br>Date: {1:dd-MMM-yyyy}<br><br><br>", data(0).row.item("validator2name").ToString, Today.Date))
        sb.Append("    List of Tasks:</p>  <table border=1 cellspacing=0>    <tr>            <th>Status</th>      <th>Reason</th>            <th>Description</th>      <th>Price Type</th>      <th>Submit Date</th>      <th>Creator</th>      <th>Validator</th>          </tr>")
        For Each n As DataRowView In data
            sb.Append(String.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4:yyyy-MMM-dd}</td><td>{5}</td><td>{6}</td></tr>", n.Item("statusname"), n.Item("reasonname"), n.Item("description"), n.Item("pricetype"), n.Item("submitdate"), n.Item("creatorname"), n.Item("validator1name")))
            If n.Item("status") = 7 Then
                n.Item("sendtocc") = True
            End If
        Next
        HDBS.EndEdit()
        'sb.Append("</table>  <br>  <p>Thank you.<br><br>You can access the system by below link:<br>   <a href=""http://hon08nt"">MyTask</a></p><br><br><p>Price CMMF Ex System Administrator</p></body></html>")
        sb.Append("</table>  <br>  <p>Thank you.<br><br>You can access the system in Citrix by below link:<br>   <a href=""https://sw07e601/RDWeb"">MyTask</a></p><br>")
        Return sb.ToString
    End Function
End Class

Public Class WMFBM
    Implements iRolePriceTask
    Private HDBS As BindingSource
    Public Sub New(ByVal HDBS As BindingSource)
        Me.HDBS = HDBS
        HDBS.Filter = "status = 7 and (sendtowmf is null or not(sendtowmf))"

    End Sub

    Public Function GetQuery() As Object Implements iRolePriceTask.GetQuery
        Return From n In HDBS.List
                         Group n By key = n.item("validator3") Into Group
                         Select key, data = Group
    End Function

    Public Function GetBodyMessage(ByVal data As Object) As String Implements iRolePriceTask.GetBodyMessage
        'Dim sb As New StringBuilder
        'For Each n As DataRowView In data
        '    sb.Append(n.Item("submitdate"))

        'Next
        'Return sb.ToString
        Dim sb As New StringBuilder
        sb.Append("<!DOCTYPE html><html><head><meta name=""description"" content=""[PriceCMMFEx]"" /><meta http-equiv=""Content-Type"" content=""text/html; charset=us-ascii""></head><style>  td,th {padding-left:5px;         padding-right:10px;         text-align:left;  }  th {background-color:red;    color:white}  .defaultfont{    font-size:11.0pt;	font-family:""Calibri"",""sans-serif"";    }</style><body class=""defaultfont"">")
        'sb.Append(String.Format("<p>Dear {0},</p><br><p>Please be informed that you have Price Change tasks with Completed status.<br><br>", data(0).item("validator2name")))
        sb.Append(String.Format("<p>Dear {0},</p><br><p>Please be informed that you have Price Change tasks with Completed status.<br>Date: {1:dd-MMM-yyyy}<br><br><br>", data(0).row.item("validator3name").ToString, Today.Date))
        sb.Append("    List of Tasks:</p>  <table border=1 cellspacing=0>    <tr>            <th>Status</th>      <th>Reason</th>            <th>Description</th>      <th>Price Type</th>      <th>Submit Date</th>      <th>Creator</th>      <th>Validator</th>          </tr>")
        For Each n As DataRowView In data
            sb.Append(String.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4:yyyy-MMM-dd}</td><td>{5}</td><td>{6}</td></tr>", n.Item("statusname"), n.Item("reasonname"), n.Item("description"), n.Item("pricetype"), n.Item("submitdate"), n.Item("creatorname"), n.Item("validator1name")))
            If n.Item("status") = 7 Then
                n.Item("sendtowmf") = True
            End If
        Next
        HDBS.EndEdit()
        'sb.Append("</table>  <br>  <p>Thank you.<br><br>You can access the system by below link:<br>   <a href=""http://hon08nt"">MyTask</a></p><br><br><p>Price CMMF Ex System Administrator</p></body></html>")
        sb.Append("</table>  <br>  <p>Thank you.<br><br>You can access the system in Citrix by below link:<br>   <a href=""https://sw07e601/RDWeb"">MyTask</a></p><br>")
        Return sb.ToString
    End Function
End Class

Interface iRolePriceTask
    Function GetQuery() As Object
    Function GetBodyMessage(ByVal n As Object) As String
End Interface