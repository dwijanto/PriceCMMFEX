Imports Npgsql
Imports NpgsqlTypes
Imports System.IO

Public Class DbAdapter
    Implements IDisposable

    Dim _ConnectionStringDict As Dictionary(Of String, String)
    Dim _connectionstring As String
    Private CopyIn1 As NpgsqlCopyIn
    Dim _userid As String
    Dim _password As String
    Dim mytransaction As NpgsqlTransaction

    Public DataReaderCallback As FormatReportDelegate

    Public ReadOnly Property userid As String
        Get
            Return _userid
        End Get
    End Property
    Public ReadOnly Property password As String
        Get
            Return _password
        End Get
    End Property

    Public Property Connectionstring As String
        Get
            Return _connectionstring

        End Get
        Set(ByVal value As String)
            _connectionstring = value
        End Set
    End Property

    Public Sub New()
        InitConnectionStringDict()
        _connectionstring = getConnectionString()
    End Sub

    Public ReadOnly Property ConnectionStringDict As Dictionary(Of String, String)
        Get
            Return _ConnectionStringDict
        End Get
    End Property

    Private Sub InitConnectionStringDict()
        _ConnectionStringDict = New Dictionary(Of String, String)
        Dim connectionstring = getConnectionString()
        Dim connectionstrings() As String = connectionstring.Split(";")
        For i = 0 To (connectionstrings.Length - 1)
            Dim mystrs() As String = connectionstrings(i).Split("=")
            _ConnectionStringDict.Add(mystrs(0), mystrs(1))
        Next i

    End Sub

    Private Function getConnectionString() As String
        _userid = "admin"
        _password = "admin"
        Dim builder As New NpgsqlConnectionStringBuilder()
        builder.ConnectionString = My.Settings.Connectionstring1
        builder.Add("User Id", _userid)
        builder.Add("password", _password)
        'builder.Add("CommandTimeout", "300")
        'builder.Add("TimeOut", "300")
        Return builder.ConnectionString
    End Function

#Region "GetDataSet"
    Public Overloads Function TbgetDataSet(ByVal sqlstr As String, ByRef DataSet As DataSet, Optional ByRef message As String = "") As Boolean
        Dim DataAdapter As New NpgsqlDataAdapter

        Dim myret As Boolean = False
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                DataAdapter.SelectCommand = New NpgsqlCommand(sqlstr, conn)
                'DataAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey

                DataAdapter.Fill(DataSet)
            End Using
            myret = True
        Catch ex As NpgsqlException
            Dim obj = TryCast(ex.Errors(0), NpgsqlError)
            Dim myerror As String = String.Empty
            If Not IsNothing(obj) Then
                myerror = obj.InternalQuery
            End If
            message = ex.Message & " " & myerror
        End Try
        Return myret
    End Function
#End Region

    Function TBScorecardDataAdapter(ByRef DataSet As DataSet, Optional ByRef message As String = "", Optional ByRef RecordAffected As Integer = 0, Optional ByVal continueupdateonerror As Boolean = True) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        'Dim cmd As NpgsqlCommand
        Dim myret As Boolean = False

        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()
                    sqlstr = "sp_getscorecard() as tb(scorecardid bigint,supplierid bigint,mydate date,deptid integer,category integer,myvalue numeric)"
                    DataAdapter.ContinueUpdateOnError = continueupdateonerror
                    DataAdapter.SelectCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.SelectCommand.Connection = conn
                    DataAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
                    'DataAdapter.MissingMappingAction = MissingSchemaAction.AddWithKey

                    'DataAdapter.SelectCommand.Parameters.Add("col1", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = dbtools1.Region
                    DataAdapter.Fill(DataSet)

                    'Delete
                    sqlstr = "sp_deletescorecard"
                    DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "scorecardid").SourceVersion = DataRowVersion.Original
                    DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                    'Update
                    sqlstr = "sp_updatescorecard"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "scorecardid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "supplierid").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "mydate").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "category").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "myvalue").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    'insert
                    sqlstr = "sp_insertscorecard"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "supplierid").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "mydate").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "category").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "myvalue").SourceVersion = DataRowVersion.Current
                    'DataAdapter.InsertCommand.Parameters.Add("paramhdid", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = paramhdid
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure


                    RecordAffected = DataAdapter.Update(DataSet.Tables(0))

                Catch ex As Exception
                    message = ex.Message
                    Return False
                End Try
            End Using
            myret = True
        Catch ex As NpgsqlException
            message = ex.Message
        End Try
        Return myret
    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
                If Not IsNothing(_ConnectionStringDict) Then
                    _ConnectionStringDict.Clear()
                    _ConnectionStringDict = Nothing
                End If
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

    

    Public Function copy(ByVal sqlstr As String, ByVal InputString As String, Optional ByRef result As Boolean = False) As String
        result = False
        Dim myReturn As String = ""
        'Convert string to MemoryStream
        Dim MemoryStream1 As New IO.MemoryStream(System.Text.Encoding.ASCII.GetBytes(InputString))
        'Dim MemoryStream1 As New IO.MemoryStream(System.Text.Encoding.Default.GetBytes(InputString))
        Dim buf(9) As Byte
        Dim CopyInStream As Stream = Nothing
        Dim i As Long
        Using conn = New NpgsqlConnection(getConnectionString())
            conn.Open()
            Using command = New NpgsqlCommand(sqlstr, conn)
                CopyIn1 = New NpgsqlCopyIn(command, conn)
                Try
                    CopyIn1.Start()
                    CopyInStream = CopyIn1.CopyStream
                    i = MemoryStream1.Read(buf, 0, buf.Length)
                    While i > 0
                        CopyInStream.Write(buf, 0, i)
                        i = MemoryStream1.Read(buf, 0, buf.Length)
                        Application.DoEvents()
                    End While
                    CopyInStream.Close()
                    result = True
                Catch ex As NpgsqlException
                    Try
                        CopyIn1.Cancel("Undo Copy")
                        myReturn = ex.Message
                    Catch ex2 As NpgsqlException
                        If ex2.Message.Contains("Undo Copy") Then
                            myReturn = ex2.Message
                        End If
                    End Try
                End Try

            End Using
        End Using

        Return myReturn
    End Function

   
    'Public Function validint(ByVal sinqqty As String) As Object

    '    If sinqqty = "" Then
    '        Return DBNull.Value
    '    Else
    '        Return CInt(sinqqty.Replace(",", "").Replace("""", ""))
    '    End If
    'End Function
    Public Function validint(ByVal sinqqty As String) As Object

        If sinqqty = "" Then
            Return DBNull.Value
        Else
            Return CInt(sinqqty.Replace(",", "").Replace("""", ""))
        End If
    End Function
    Public Function validbool(ByVal mybool As String) As Object
        If mybool = "Y" Then
            Return "True"
        Else
            Return "False"
        End If
    End Function
    Public Function validdec(ByVal sunitprice As String) As Object
        If sunitprice = "" Then
            Return DBNull.Value
        Else
            Return CDec(sunitprice.Replace(",", "").Replace("""", ""))
        End If
    End Function

    Public Function validlong(ByVal myvalue As String) As Object
        If myvalue = "" Then
            Return DBNull.Value
        Else
            Return CLng(myvalue)
        End If
    End Function
    Public Function validlongNull(ByVal myvalue As String) As Object
        If myvalue = "" Then
            Return "Null"
        Else
            Return CLng(myvalue)
        End If
    End Function
    Public Function validchar(ByVal updateby As String) As Object
        If updateby = "" Then
            'Return DBNull.Value
            Return ""
        Else
            Return Trim(updateby.Replace("'", "''").Replace("""", "").Replace("\", "\\"))
        End If
    End Function
    Public Function validcharNull(ByVal updateby As String) As Object
        If updateby = "" Then
            'Return DBNull.Value
            Return "Null"
        Else
            Return Trim(updateby.Replace("'", "''").Replace("""", ""))
        End If
    End Function
    Public Function CDateddMMyyyy(ByVal updatedate As String) As Object
        Dim mydata() As String
        If updatedate.Contains(".") Then
            mydata = updatedate.Split(".")
        Else
            mydata = updatedate.Split("/")
        End If

        If mydata.Length > 1 Then
            Return CDate(mydata(2) & "-" & mydata(1) & "-" & mydata(0))
        End If
        Return DBNull.Value
    End Function
    Public Function ddMMyyyytoyyyyMMdd(ByVal updatedate As String) As Object
        Dim mydata() As String
        If updatedate.Contains(".") Then
            mydata = updatedate.Split(".")
        Else
            mydata = updatedate.Split("/")
        End If

        If mydata.Length > 1 Then
            Return "'" & mydata(2) & "-" & mydata(1) & "-" & mydata(0) & "'"
        End If
        Return DBNull.Value
    End Function

   
    Public Function ExNonQuery(ByVal sqlstr As String) As Long
        Dim myRet As Long
        Using conn = New NpgsqlConnection(getConnectionString)
            conn.Open()
            Using command = New NpgsqlCommand(sqlstr, conn)
                myRet = command.ExecuteNonQuery
            End Using
        End Using
        Return myRet
    End Function

    Public Function ExecuteNonQuery(ByVal sqlstr As String, Optional ByRef recordAffected As Int64 = 0, Optional ByRef message As String = "") As Boolean
        Dim myRet As Boolean = False
        Using conn = New NpgsqlConnection(getConnectionString)
            conn.Open()
            Using command = New NpgsqlCommand(sqlstr, conn)
                Try
                    recordAffected = command.ExecuteNonQuery
                    'recordAffected = command.ExecuteNonQuery
                    myRet = True
                Catch ex As NpgsqlException
                    message = ex.Message
                End Try
            End Using
        End Using
        Return myRet
    End Function
    Public Function ExecuteNonQueryAsync(ByVal sqlstr As String, Optional ByRef recordAffected As Int64 = 0, Optional ByRef message As String = "") As Boolean
        Dim myRet As Boolean = False
        Using conn = New NpgsqlConnection(getConnectionString)
            conn.Open()
            Using command = New NpgsqlCommand(sqlstr, conn)
                Try
                    recordAffected = command.ExecuteNonQuery

                    myRet = True
                Catch ex As NpgsqlException
                    message = ex.Message
                End Try
            End Using
        End Using
        Return myRet
    End Function
    Public Function ExecuteScalar(ByVal sqlstr As String, Optional ByRef recordAffected As Int64 = 0, Optional ByRef message As String = "") As Boolean
        Dim myRet As Boolean = False
        Using conn = New NpgsqlConnection(getConnectionString)
            conn.Open()
            Using command = New NpgsqlCommand(sqlstr, conn)
                Try
                    recordAffected = command.ExecuteScalar
                    myRet = True
                Catch ex As NpgsqlException
                    message = ex.Message
                End Try
            End Using
        End Using
        Return myRet
    End Function
    Public Function ExecuteScalar(ByVal sqlstr As String, ByRef recordAffected As Object, Optional ByRef message As String = "") As Boolean
        Dim myRet As Boolean = False
        Using conn = New NpgsqlConnection(getConnectionString)
            conn.Open()
            Using command = New NpgsqlCommand(sqlstr, conn)
                Try
                    recordAffected = command.ExecuteScalar
                    myRet = True
                Catch ex As NpgsqlException
                    message = ex.Message
                End Try
            End Using
        End Using
        Return myRet
    End Function

    Sub ExecuteStoreProcedure(ByVal storeprocedurename As String)
        Using conn As New NpgsqlConnection(Connectionstring)
            Try
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand(storeprocedurename, conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.ExecuteScalar()
            Catch ex As Exception

            End Try

        End Using
    End Sub

    Public Function getproglock(ByVal programname As String, ByVal userid As String, ByVal status As Integer) As Boolean
        Dim result As Object
        Using conn As New NpgsqlConnection(Connectionstring)
            conn.Open()
            Dim cmd As NpgsqlCommand = New NpgsqlCommand("proglock", conn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = programname
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = userid
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = status
            result = cmd.ExecuteScalar
        End Using
       
        Return result
    End Function

    Function dateformatdot(ByVal myrecord As String) As Object
        Dim myreturn = "Null"
        If myrecord = "" Then
            Return myreturn
        End If
        Dim mysplit = Split(myrecord, ".")
        myreturn = "'" & mysplit(2) & "-" & mysplit(1) & "-" & mysplit(0) & "'"
        Return myreturn
    End Function

    Function dateformatdotdate(ByVal myrecord As String) As Date
        Dim myreturn = "Null"
        If myrecord = "" Then
            Return myreturn
        End If
        Dim mysplit = Split(myrecord, ".")
        myreturn = CDate(mysplit(2) & "-" & mysplit(1) & "-" & mysplit(0))
        Return myreturn
    End Function

    Function dateformatYYYYMMdd(ByVal myrecord As Object) As Object
        Dim myreturn = "Null"

        myreturn = "'" & CDate(myrecord).Year & "-" & CDate(myrecord).Month & "-" & CDate(myrecord).Day & "'"
        Return myreturn
    End Function

    Function ImportTx(ByVal formImportSaving As FormImportSaving, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        'Dim cmd As NpgsqlCommand
        Dim myret As Boolean = False

        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()
                    myTransaction = conn.BeginTransaction
                    'Update
                    sqlstr = "sp_updatesavinglookup"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "savinglookupid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "savinglookupname").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "parentid").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    'insert
                    sqlstr = "sp_insertsavinglookup"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "savinglookupname").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "parentid").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                    DataAdapter.UpdateCommand.Transaction = mytransaction
                    DataAdapter.InsertCommand.Transaction = mytransaction

                    mye.ra = DataAdapter.Update(mye.dataset.Tables(1))

                    sqlstr = "sp_updatesaving"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "savingid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "mytotal").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "enddate").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    'insert
                    sqlstr = "sp_insertsaving"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "actionid").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "cmmf").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "mytotal").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "startdate").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "enddate").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                    DataAdapter.UpdateCommand.Transaction = mytransaction
                    DataAdapter.InsertCommand.Transaction = mytransaction

                    mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                    mytransaction.Commit()
                    myret = True
                Catch ex As Exception
                    mytransaction.Rollback()
                    mye.message = ex.Message
                    Return False
                End Try
            End Using

        Catch ex As NpgsqlException
            mye.message = ex.Message
        End Try
        Return myret

    End Function
    Function copyToPriceChange(ByVal creator As String, ByRef dr As DataRow, ByVal isnewrecord As Boolean, ByRef message As String) As Boolean
        Dim myret As Boolean
        Dim result As Object
        Dim myparam As NpgsqlParameter
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_copypricechange", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = creator
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0).Value = isnewrecord
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = dr.Item("validator1")
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = dr.Item("validator2")
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = dr.Item("pricetype")
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = dr.Item("description")
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = dr.Item("submitdate")
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = dr.Item("negotiateddate")
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = dr.Item("attachment")
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = dr.Item("reasonid")

                myparam = cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0)
                myparam.Value = dr.Item("pricechangehdid")
                myparam.Direction = ParameterDirection.InputOutput


                result = cmd.ExecuteScalar
                If IsDBNull(result) Then
                    myret = False
                End If
                myret = True
            End Using

        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            message = ex.Message & ". " & errordetail
        End Try
        Return myret
    End Function

    Function PriceChangeTx(ByVal formPriceChange As FormPriceChange, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        'Dim cmd As NpgsqlCommand
        Dim myret As Boolean = False

        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()
                    mytransaction = conn.BeginTransaction
                    'Update
                    sqlstr = "sp_updatepricechangehd"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "pricechangehdid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "creator").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "validator1").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "validator2").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "pricetype").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "description").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "submitdate").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "negotiateddate").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "attachment").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "status").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "actiondate").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "actionby").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "reasonid").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    sqlstr = "sp_insertpricechangehd"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "creator").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "validator1").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "validator2").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "pricetype").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "description").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "submitdate").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "negotiateddate").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "attachment").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "status").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "actiondate").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "actionby").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "reasonid").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "pricechangehdid").Direction = ParameterDirection.InputOutput
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure


                    DataAdapter.InsertCommand.Transaction = mytransaction
                    DataAdapter.UpdateCommand.Transaction = mytransaction

                    mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                    sqlstr = "sp_updatepricechangedtl"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "pricechangedtlid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "pricechangehdid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "cmmf").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "purchorg").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "plant").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "validon").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "price").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "pricingunit").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "comment").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    'insert
                    sqlstr = "sp_insertpricechangedtl"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "pricechangehdid").SourceVersion = DataRowVersion.Original
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "cmmf").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "purchorg").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "plant").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "validon").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "price").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "pricingunit").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "comment").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure


                    DataAdapter.UpdateCommand.Transaction = mytransaction
                    DataAdapter.InsertCommand.Transaction = mytransaction


                    sqlstr = "sp_deletepricechangedtl"
                    DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "pricechangedtlid").SourceVersion = DataRowVersion.Original
                    DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure


                    mye.ra = DataAdapter.Update(mye.dataset.Tables(4))

                    mytransaction.Commit()
                    myret = True
                    'Catch ex As Exception
                    '    Dim myerr = CType(ex, NpgsqlException)
                    '    mytransaction.Rollback()
                    '    mye.message = ex.Message & " " & myerr.Detail
                    '    Return False
                    'End Try
                Catch ex As NpgsqlException
                    Dim errordetail As String = String.Empty
                    errordetail = "" & ex.Detail
                    mye.message = ex.Message & ". " & errordetail
                    Return False
                End Try
            End Using

        Catch ex As NpgsqlException
            mye.message = ex.Message
        End Try
        Return myret
    End Function

   

    Function PriceCommentTx(ByVal pricechangehdid As Long, ByRef message As String) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim result As Object
        Dim myret As Boolean = False

        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_addpricecomment", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0).Value = pricechangehdid

                result = cmd.ExecuteScalar
                If IsDBNull(result) Then
                    myret = False
                End If
                myret = True
            End Using

        Catch ex As NpgsqlException
            Message = ex.Message
        End Try
        Return myret
    End Function
    Public Sub onRowInsertUpdate(ByVal sender As Object, ByVal e As NpgsqlRowUpdatedEventArgs)
        If e.StatementType = StatementType.Insert Or e.StatementType = StatementType.Update Then
            If e.Status <> UpdateStatus.ErrorsOccurred Then
                e.Status = UpdateStatus.SkipCurrentRow
            End If
        End If
    End Sub
    Function PriceChangeTx(ByVal formPriceChange As FormMyTask, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        'Dim cmd As NpgsqlCommand
        Dim myret As Boolean = False
        'AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowInsertUpdate)
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()
                    mytransaction = conn.BeginTransaction
                    'Update
                    sqlstr = "sp_updatepricechangehd"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "pricechangehdid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "creator").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "validator1").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "validator2").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "pricetype").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "description").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "submitdate").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "negotiateddate").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "attachment").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "status").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "actiondate").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "actionby").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "reasonid").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure


                    DataAdapter.UpdateCommand.Transaction = mytransaction


                    mye.ra = DataAdapter.Update(mye.dataset.Tables(1))

                    'sqlstr = "sp_updatepricechangedtl"
                    'DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "pricechangedtlid").SourceVersion = DataRowVersion.Original
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "pricechangehdid").SourceVersion = DataRowVersion.Original
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "cmmf").SourceVersion = DataRowVersion.Current
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "purchorg").SourceVersion = DataRowVersion.Current
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "plant").SourceVersion = DataRowVersion.Current
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "validon").SourceVersion = DataRowVersion.Current
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "price").SourceVersion = DataRowVersion.Current
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "pricingunit").SourceVersion = DataRowVersion.Current
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "comment").SourceVersion = DataRowVersion.Current
                    'DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure



                    'DataAdapter.UpdateCommand.Transaction = mytransaction


                    'mye.ra = DataAdapter.Update(mye.dataset.Tables(4))

                    mytransaction.Commit()
                    myret = True
                Catch ex As Exception
                    mytransaction.Rollback()
                    mye.message = ex.Message
                    Return False
                End Try
            End Using

        Catch ex As NpgsqlException
            mye.message = ex.Message
        End Try
        Return myret
    End Function

    'Function PriceChangeTx(ByVal formPriceChange As FormMyTask2, ByVal mye As ContentBaseEventArgs) As Boolean
    Function PriceChangeTx(ByVal formPriceChange As Object, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        'Dim cmd As NpgsqlCommand
        Dim myret As Boolean = False

        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()
                    mytransaction = conn.BeginTransaction
                    'Update
                    sqlstr = "sp_updatepricechangehd"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "pricechangehdid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "creator").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "validator1").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "validator2").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "pricetype").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "description").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "submitdate").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "negotiateddate").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "attachment").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "status").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "actiondate").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "actionby").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "reasonid").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure


                    DataAdapter.UpdateCommand.Transaction = mytransaction


                    mye.ra = DataAdapter.Update(mye.dataset.Tables(1))

                    'sqlstr = "sp_updatepricechangedtl"
                    'DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "pricechangedtlid").SourceVersion = DataRowVersion.Original
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "pricechangehdid").SourceVersion = DataRowVersion.Original
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "cmmf").SourceVersion = DataRowVersion.Current
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "purchorg").SourceVersion = DataRowVersion.Current
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "plant").SourceVersion = DataRowVersion.Current
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "validon").SourceVersion = DataRowVersion.Current
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "price").SourceVersion = DataRowVersion.Current
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "pricingunit").SourceVersion = DataRowVersion.Current
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "comment").SourceVersion = DataRowVersion.Current
                    'DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure



                    'DataAdapter.UpdateCommand.Transaction = mytransaction


                    'mye.ra = DataAdapter.Update(mye.dataset.Tables(4))

                    mytransaction.Commit()
                    myret = True
                Catch ex As Exception
                    mytransaction.Rollback()
                    mye.message = ex.Message
                    Return False
                End Try
            End Using

        Catch ex As NpgsqlException
            mye.message = ex.Message
        End Try
        Return myret
    End Function

    Function PriceChangeTx(ByVal formPriceChange As FormPriceChange2, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        'Dim cmd As NpgsqlCommand
        Dim myret As Boolean = False

        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()
                    mytransaction = conn.BeginTransaction
                    'Update
                    sqlstr = "sp_updatepricechangehd"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "pricechangehdid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "creator").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "validator1").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "validator2").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "validator3").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "pricetype").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "description").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "submitdate").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "negotiateddate").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "attachment").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "status").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "actiondate").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "actionby").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "reasonid").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "specialprojectid").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    sqlstr = "sp_insertpricechangehd"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "creator").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "validator1").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "validator2").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "validator3").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "pricetype").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "description").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "submitdate").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "negotiateddate").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "attachment").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "status").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "actiondate").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "actionby").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "reasonid").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "specialprojectid").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "pricechangehdid").Direction = ParameterDirection.InputOutput
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure


                    DataAdapter.InsertCommand.Transaction = mytransaction
                    DataAdapter.UpdateCommand.Transaction = mytransaction

                    mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                    sqlstr = "sp_updatepricechangedtl"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "pricechangedtlid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "pricechangehdid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "cmmf").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "purchorg").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "plant").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "validon").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "price").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "pricingunit").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "comment").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    'insert
                    sqlstr = "sp_insertpricechangedtl"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "pricechangehdid").SourceVersion = DataRowVersion.Original
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "cmmf").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "purchorg").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "plant").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "validon").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "price").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "pricingunit").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "comment").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure


                    DataAdapter.UpdateCommand.Transaction = mytransaction
                    DataAdapter.InsertCommand.Transaction = mytransaction


                    sqlstr = "sp_deletepricechangedtl"
                    DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "pricechangedtlid").SourceVersion = DataRowVersion.Original
                    DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure


                    mye.ra = DataAdapter.Update(mye.dataset.Tables(4))

                    mytransaction.Commit()
                    myret = True
                    'Catch ex As Exception
                    '    Dim myerr = CType(ex, NpgsqlException)
                    '    mytransaction.Rollback()
                    '    mye.message = ex.Message & " " & myerr.Detail
                    '    Return False
                    'End Try
                Catch ex As NpgsqlException
                    Dim errordetail As String = String.Empty
                    errordetail = "" & ex.Detail
                    mye.message = ex.Message & ". " & errordetail
                    Return False
                End Try
            End Using

        Catch ex As NpgsqlException
            mye.message = ex.Message
        End Try
        Return myret
    End Function
    Function PriceChangeReasonTx(ByVal formMasterProduct As FormPriceChangeReasonMaster, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowInsertUpdate)
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()

                    'Family
                    'Insert
                    sqlstr = "sp_insertpricechangereason"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "reasonname").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "isactive").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "lineno").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "id").Direction = ParameterDirection.InputOutput
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure


                    'Update
                    sqlstr = "sp_updatepricechangereason"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "id").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "reasonname").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "isactive").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "lineno").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure


                    sqlstr = "sp_deletepricechangereason"
                    DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "id").SourceVersion = DataRowVersion.Original
                    DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                    mytransaction = conn.BeginTransaction
                    DataAdapter.DeleteCommand.Transaction = mytransaction
                    DataAdapter.UpdateCommand.Transaction = mytransaction
                    DataAdapter.InsertCommand.Transaction = mytransaction

                    mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                    mytransaction.Commit()
                Catch ex As Exception
                    mytransaction.Rollback()
                    mye.message = ex.Message
                    Return False
                End Try
            End Using
            myret = True
        Catch ex As NpgsqlException
            mye.message = ex.Message
        End Try
        Return myret
    End Function

    Function CurrencyTx(ByVal MyForm As FormCurrency, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowInsertUpdate)
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()

                    'Family
                    'Insert
                    sqlstr = "doc.sp_insertcurency"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "crcy").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "currency").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "budgetcurrency").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "myyear").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "id").Direction = ParameterDirection.InputOutput
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure


                    'Update
                    sqlstr = "doc.sp_updatecurrency"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "id").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "crcy").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "currency").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "budgetcurrency").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "myyear").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure


                    sqlstr = "doc.sp_deletecurrency"
                    DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "id").SourceVersion = DataRowVersion.Original
                    DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                    mytransaction = conn.BeginTransaction
                    DataAdapter.DeleteCommand.Transaction = mytransaction
                    DataAdapter.UpdateCommand.Transaction = mytransaction
                    DataAdapter.InsertCommand.Transaction = mytransaction

                    mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                    mytransaction.Commit()
                Catch ex As Exception
                    mytransaction.Rollback()
                    mye.message = ex.Message
                    Return False
                End Try
            End Using
            myret = True
        Catch ex As NpgsqlException
            mye.message = ex.Message
        End Try
        Return myret
    End Function
    Function PriceChangeSpecialProject(ByVal formMasterProduct As FormSpecialProjectMaster, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowInsertUpdate)
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()
                    'Insert
                    sqlstr = "sp_insertpricechangespecialproject"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "specialproject").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "isactive").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "lineno").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "id").Direction = ParameterDirection.InputOutput
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure


                    'Update
                    sqlstr = "sp_updatepricechangespecialproject"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "id").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "specialproject").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "isactive").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "lineno").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure


                    sqlstr = "sp_deletepricechangespecialproject"
                    DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "id").SourceVersion = DataRowVersion.Original
                    DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                    mytransaction = conn.BeginTransaction
                    DataAdapter.DeleteCommand.Transaction = mytransaction
                    DataAdapter.UpdateCommand.Transaction = mytransaction
                    DataAdapter.InsertCommand.Transaction = mytransaction

                    mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                    mytransaction.Commit()
                Catch ex As Exception
                    mytransaction.Rollback()
                    mye.message = ex.Message
                    Return False
                End Try
            End Using
            myret = True
        Catch ex As NpgsqlException
            mye.message = ex.Message
        End Try
        Return myret
    End Function

    Function PriceChangeSendEmail(ByVal roleTasks As RoleTasks, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False

        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()

                    'Update
                    sqlstr = "sp_updatepricechangehdemailsend"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "pricechangehdid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "sendstdvalidatedtocreator").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "sendcompletedtocreator").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "sendtocc").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "sendtowmf").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    mytransaction = conn.BeginTransaction
                    DataAdapter.UpdateCommand.Transaction = mytransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                    mytransaction.Commit()
                Catch ex As Exception
                    mytransaction.Rollback()
                    mye.message = ex.Message
                    Return False
                End Try
            End Using
            myret = True
        Catch ex As NpgsqlException
            mye.message = ex.Message
        End Try
        Return myret
    End Function

    Function PriceChangeExportFile(ByVal exportSAPClass As ExportSAPClass, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False

        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()

                    'Update
                    sqlstr = "sp_updatepricechangehdexportfile"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "pricechangehdid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "exportfileid").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "exportfiledate").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    mytransaction = conn.BeginTransaction
                    DataAdapter.UpdateCommand.Transaction = mytransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                    mytransaction.Commit()
                Catch ex As Exception
                    mytransaction.Rollback()
                    mye.message = ex.Message
                    Return False
                End Try
            End Using
            myret = True
        Catch ex As NpgsqlException
            mye.message = ex.Message
        End Try
        Return myret
    End Function

    Function PriceChangedtlsap(ByVal validateSAPPrice As ValidateSAPPrice, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False

        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()

                    'Update
                    sqlstr = "sp_updatepricechangedtlsap"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "pricechangedtlid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "sap").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    mytransaction = conn.BeginTransaction
                    DataAdapter.UpdateCommand.Transaction = mytransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(1))

                    mytransaction.Commit()
                Catch ex As Exception
                    mytransaction.Rollback()
                    mye.message = ex.Message
                    Return False
                End Try
            End Using
            myret = True
        Catch ex As NpgsqlException
            mye.message = ex.Message
        End Try
        Return myret
    End Function

    Function PriceChangehdcompleted(ByVal validateSAPPrice As ValidateSAPPrice, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False

        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()

                    'Update
                    sqlstr = "sp_updatepricechangehdcompleted"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "pricechangehdid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "status").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    mytransaction = conn.BeginTransaction
                    DataAdapter.UpdateCommand.Transaction = mytransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                    mytransaction.Commit()
                Catch ex As Exception
                    mytransaction.Rollback()
                    mye.message = ex.Message
                    Return False
                End Try
            End Using
            myret = True
        Catch ex As NpgsqlException
            mye.message = ex.Message
        End Try
        Return myret
    End Function

    Function PriceChangeDTLTx(ByVal formHistoryDetail As Object, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False

        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()


                    sqlstr = "sp_deletepricechangedtl"
                    DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "pricechangedtlid").SourceVersion = DataRowVersion.Original
                    DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                    mytransaction = conn.BeginTransaction
                    DataAdapter.DeleteCommand.Transaction = mytransaction

                    mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                    mytransaction.Commit()
                Catch ex As Exception
                    mytransaction.Rollback()
                    mye.message = ex.Message
                    Return False
                End Try
            End Using
            myret = True
        Catch ex As NpgsqlException
            mye.message = ex.Message
        End Try
        Return myret
    End Function

    Function DocumentVendorTx(ByVal formDocumentVendor As FormDocumentVendor, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False

        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                mytransaction = conn.BeginTransaction
                'Update
                sqlstr = "doc.sp_updatedocumenthd"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "userid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "validator").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "cc1").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "cc2").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "cc3").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "cc4").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "otheremail").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "creationdate").SourceVersion = DataRowVersion.Current

                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_insertdocumenthd"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "userid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "validator").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "cc1").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "cc2").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "cc3").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "cc4").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "otheremail").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "creationdate").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure


                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction

                mye.ra = DataAdapter.Update(mye.dataset.Tables(0))
                'vd.*,v.vendorname::text,v.shortname::text,
                'd.*,vr.version,gt.paymentcode,sc.leadtime,sc.sasl,q.nqsu,p.projectname,
                'sa.auditby,sa.audittype,sa.auditgrade,
                'sef.score,sif.myyear,sif.turnovery,sif.turnovery1,sif.turnovery2,sif.turnovery3,sif.turnovery4,sif.ratioy,sif.ratioy1,sif.ratioy2,sif.ratioy3,sif.ratioy4,
                '' as filename,dt.doctypename,dl.levelname
                sqlstr = "doc.sp_updatedocumentdtl"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                'vendordoc
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "headerid").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "documentid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "status").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "shortname").SourceVersion = DataRowVersion.Current
                'document
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "docdate").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "docname").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "docext").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "uploaddate").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "userid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "remarks").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "doctypeid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "doclevelid").SourceVersion = DataRowVersion.Current
                'version
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "version").SourceVersion = DataRowVersion.Current
                'generalcontract
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "paymentcode").SourceVersion = DataRowVersion.Current
                'supplychain
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "leadtime").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "sasl").SourceVersion = DataRowVersion.Current
                'quality
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "nqsu").SourceVersion = DataRowVersion.Current
                'project
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "projectname").SourceVersion = DataRowVersion.Current
                'socialaudit
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "auditby").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "audittype").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "auditgrade").SourceVersion = DataRowVersion.Current
                'sef
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "score").SourceVersion = DataRowVersion.Current
                'sif
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "myyear").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "turnovery").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "turnovery1").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "turnovery2").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "turnovery3").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "turnovery4").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "ratioy").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "ratioy1").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "ratioy2").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "ratioy3").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "ratioy4").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                'insert
                sqlstr = "doc.sp_insertdocumentdetails"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                'vendordoc
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "headerid").SourceVersion = DataRowVersion.Original
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "documentid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "status").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "shortname").SourceVersion = DataRowVersion.Current
                'document
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "docdate").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "docname").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "docext").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "uploaddate").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "userid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "remarks").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "doctypeid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "doclevelid").SourceVersion = DataRowVersion.Current
                'version
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "version").SourceVersion = DataRowVersion.Current
                'generalcontract
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "paymentcode").SourceVersion = DataRowVersion.Current
                'supplychain
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "leadtime").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "sasl").SourceVersion = DataRowVersion.Current
                'quality
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "nqsu").SourceVersion = DataRowVersion.Current
                'project
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "nqsu").SourceVersion = DataRowVersion.Current
                'socialaudit
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "auditby").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "audittype").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "auditgrade").SourceVersion = DataRowVersion.Current
                'sef
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "score").SourceVersion = DataRowVersion.Current
                'sif
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "myyear").SourceVersion = DataRowVersion.Original
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "turnovery").SourceVersion = DataRowVersion.Original
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "turnovery1").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "turnovery2").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "turnovery3").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "turnovery4").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "ratio").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "ratio1").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "ratio2").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "ratio3").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "ratio4").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure


                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.InsertCommand.Transaction = mytransaction


                sqlstr = "doc.sp_deletedocumentdetails"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                'vendordoc
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure


                mye.ra = DataAdapter.Update(mye.dataset.Tables(1))

                mytransaction.Commit()
                myret = True

            End Using

        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            mye.message = ex.Message & ". " & errordetail
            Return False
        End Try
        Return myret
    End Function

    Function loglogin(ByVal applicationname As String, ByVal userid As String, ByVal username As String, ByVal computername As String, ByVal time_stamp As Date) As Boolean
        Dim result As Object
        Using conn As New NpgsqlConnection(Connectionstring)
            conn.Open()
            Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_insertlogonhistory", conn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = applicationname
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = userid
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = username
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = computername
            result = cmd.ExecuteNonQuery
        End Using
        Return result
    End Function

    Function CanFindUserVendor(ByVal p1 As String, ByVal vendorlist As System.Text.StringBuilder) As Boolean
        Dim result As Object
        Using conn As New NpgsqlConnection(Connectionstring)
            conn.Open()
            Dim cmd As NpgsqlCommand = New NpgsqlCommand("doc.sp_canfinduservendor", conn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = p1
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = vendorlist            
            result = cmd.ExecuteScalar
        End Using

        Return result <> 0
    End Function

    Public Function ExecuteReader(ByVal sqlstr As String, ByRef message As String) As Boolean
        Dim myRet As Boolean = False
        Using conn = New NpgsqlConnection(getConnectionString)
            conn.Open()
            Using command = New NpgsqlCommand(sqlstr, conn)
                Try
                    Dim dr As NpgsqlDataReader = command.ExecuteReader
                    DataReaderCallback.Invoke(dr, New EventArgs)
                    myRet = True
                Catch ex As NpgsqlException
                    message = ex.Message
                End Try
            End Using
        End Using
        Return myRet
    End Function

    Public Function isAdmin(ByVal userid As String) As Boolean
        Dim myret As Boolean = False

        Dim sqlstr = String.Format("select isadmin from doc.user where lower(userid) = '{0}';", userid.ToLower)
        Dim result As Object = Nothing
        If ExecuteScalar(sqlstr, result) Then
            myret = IIf(IsNothing(result), False, result)
        End If
        Return myret
    End Function



End Class
