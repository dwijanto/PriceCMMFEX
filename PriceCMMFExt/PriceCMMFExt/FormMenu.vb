Imports System.Reflection
Imports PriceCMMFExt.PublicClass

Public Class FormMenu
    Private Sub FormMenu_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            HelperClass1 = New HelperClass
            DbAdapter1 = New DbAdapter
            Me.Text = GetMenuDesc()
            Me.Location = New Point(300, 10)
            Try
                loglogin(DbAdapter1.userid)
                HelperClass1.UserInfo.isAdmin = DbAdapter1.isAdmin(HelperClass1.UserId)
            Catch ex As Exception
            End Try
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Me.Close()
        End Try

    End Sub
    Public Function GetMenuDesc() As String
        'Label1.Text = "Welcome, " & HelperClass1.UserInfo.DisplayName
        Return "App.Version: " & My.Application.Info.Version.ToString & " :: Server: " & DbAdapter1.ConnectionStringDict.Item("HOST") & ", Database: " & DbAdapter1.ConnectionStringDict.Item("DATABASE") & ", Userid: " & HelperClass1.UserId

    End Function
    Private Sub ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim ctrl As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
        Dim assembly1 As Assembly = Assembly.GetAssembly(GetType(FormMenu))
        Dim frm As Form = CType(assembly1.CreateInstance(assembly1.GetName.Name.ToString & "." & ctrl.Tag.ToString, True), Form)
        Dim inMemory As Boolean = False
        For i = 0 To My.Application.OpenForms.Count - 1
            If My.Application.OpenForms.Item(i).Name = frm.Name Then
                ExecuteForm(My.Application.OpenForms.Item(i))
                inMemory = True
            End If
        Next
        If Not inMemory Then
            ExecuteForm(frm)
        End If
    End Sub

    Private Sub ExecuteForm(ByVal obj As Windows.Forms.Form)
        With obj
            .WindowState = FormWindowState.Normal
            .StartPosition = FormStartPosition.CenterScreen
            .Show()
            .Focus()
        End With
    End Sub

    Private Sub FormMenu_Load1(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.FormMenu_Load(Me, New EventArgs)
        AddHandler BSEGSQ01ToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler AveragePriceIndexToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler ZZA0035ToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler PO40SQ01PO40AndPO41LocalfileToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler ImportSavingsToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler AveragePriceIndexSavingsToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler ImportZZA037ToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler SupplierDocumentsToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler PO39ToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
    End Sub

    Private Sub FormMenu_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Not e.CloseReason = CloseReason.ApplicationExitCall Then
            If MessageBox.Show("Are you sure?", "Exit", MessageBoxButtons.OKCancel) = DialogResult.OK Then
                Me.CloseOpenForm()
                HelperClass1.fadeout(Me)
                DbAdapter1.Dispose()
                HelperClass1.Dispose()
            Else
                e.Cancel = True
            End If
        End If
    End Sub
    Private Sub CloseOpenForm()
        For i = 1 To (My.Application.OpenForms.Count - 1)
            My.Application.OpenForms.Item(1).Close()
        Next
    End Sub

    Private Sub PO40SQ01PO40AndPO41LocalfileToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PO40SQ01PO40AndPO41LocalfileToolStripMenuItem.Click

    End Sub

    Private Sub PriceChangeTaskToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PriceChangeTaskToolStripMenuItem.Click
        'Dim myform As New FormMyTask2
        Dim myform As New FormMyTask3
        myform.Show()
    End Sub

    Private Sub ImportZZA037ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImportZZA037ToolStripMenuItem.Click

    End Sub

    Private Sub SupplierDocumentsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SupplierDocumentsToolStripMenuItem.Click

    End Sub


    Private Sub UserGuideToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UserGuideToolStripMenuItem.Click
        Dim p As New System.Diagnostics.Process
        p.StartInfo.FileName = "\\172.22.10.77\SharedFolder\PriceCMMF\New\template\PriceCMMFUserGuide.pdf"
        p.Start()
    End Sub

    Private Sub loglogin(ByVal userid As String)
        Dim applicationname As String = "PriceCMMFEX"
        Dim username As String = Environment.UserDomainName & "\" & Environment.UserName
        Dim computername As String = My.Computer.Name
        Dim time_stamp As DateTime = Now
        DbAdapter1.loglogin(applicationname, userid, username, computername, time_stamp)
    End Sub

    Private Sub ZZA0035ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ZZA0035ToolStripMenuItem.Click

    End Sub

    Private Sub PO39SQ01PO39Plant3750LocalfileToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PO39ToolStripMenuItem.Click

    End Sub


    Private Sub ReportToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReportToolStripMenuItem.Click

    End Sub
End Class
