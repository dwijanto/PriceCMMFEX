Imports PriceCMMFExt.PublicClass
Public Class ValidateSAPPrice
    Dim DS As DataSet
    Dim bs As BindingSource
    Property mymessage As String
    Public Sub New()
        If IsNothing(DbAdapter1) Then
            DbAdapter1 = New DbAdapter
        End If
    End Sub


    Public Function Validate() As Boolean
        Dim myret = False

        'DataSet 0 => PriceFromSAP -> get Unique based on CMMF,VendorCode,ValidOn
        'DataSet 1 => PriceChange Status = 4 

        Dim sqlstr = "with p as (select max(pricelistid) as id from pricelist group by cmmf,vendorcode,validfrom)" &
                     "select pl.cmmf,pl.vendorcode,pl.validfrom,(pl.amount::numeric / pl.perunit::numeric)as amount from p" &
                     " left join pricelist pl on pl.pricelistid = p.id;" &
                     "select *,price / pricingunit as amount from pricechangedtl where sap isnull;"
        Try
            DS = New DataSet
            If DbAdapter1.TbgetDataSet(sqlstr, DS, mymessage) Then
                DS.Tables(0).TableName = "SAPPrice"
                Dim pk0(3) As DataColumn
                pk0(0) = DS.Tables(0).Columns("cmmf")
                pk0(1) = DS.Tables(0).Columns("vendorcode")
                pk0(2) = DS.Tables(0).Columns("validfrom")
                pk0(3) = DS.Tables(0).Columns("amount")
                DS.Tables(0).PrimaryKey = pk0

                bs = New BindingSource
                bs.DataSource = DS.Tables(1)
            End If

            For Each dr As DataRowView In bs.List
                Dim myobj(3) As Object
                myobj(0) = dr.Row.Item("cmmf")
                myobj(1) = dr.Row.Item("vendorcode")
                myobj(2) = dr.Row.Item("validon")
                myobj(3) = dr.Row.Item("amount")

                Dim myresult = DS.Tables(0).Rows.Find(myobj)
                If Not IsNothing(myresult) Then
                    dr.Row.Item("sap") = True
                End If
                'Debug.Print("hello")
            Next

            Dim ds2 = DS.GetChanges
            If Not IsNothing(ds2) Then
                Dim mymessage As String = String.Empty
                Dim ra As Integer
                Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
                If Not DbAdapter1.PriceChangedtlsap(Me, mye) Then
                    Logger.log(mye.message)
                Else
                    DS.AcceptChanges()
                End If



            End If
            setCompleted()

            myret = True
        Catch ex As Exception
            mymessage = ex.Message
        End Try
        






        Return myret
    End Function

    Private Function setCompleted() As Boolean
        Dim myds As New DataSet
        Dim myret = False

        'if Select Count CMMF = Count SAP then set Complete in header where status is 5 (validated)
        Dim sqlstr = "select count(cmmf),count(sap),pd.pricechangehdid,ph.status from pricechangedtl pd" &
                     " left join pricechangehd ph on ph.pricechangehdid = pd.pricechangehdid" &
                     " group by pd.pricechangehdid,ph.status" &
                     " having count(cmmf) = count(sap) and status = 5"

        Try
            If DbAdapter1.TbgetDataSet(sqlstr, myds, mymessage) Then
                If myds.Tables(0).Rows.Count > 0 Then
                    For Each dr As DataRow In myds.Tables(0).Rows
                        dr.Item("status") = 7
                    Next

                    Dim ds2 = myds.GetChanges
                    If Not IsNothing(ds2) Then
                        Dim mymessage As String = String.Empty
                        Dim ra As Integer
                        Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
                        If Not DbAdapter1.PriceChangehdcompleted(Me, mye) Then
                            Logger.log(mye.message)
                        Else
                            myds.AcceptChanges()
                        End If



                    End If

                End If
            End If
        Catch ex As Exception

        End Try

        Return myret
    End Function
End Class
