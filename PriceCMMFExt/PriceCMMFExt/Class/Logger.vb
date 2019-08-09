Imports System.IO
Imports System.Text
Imports System.Reflection
Public Class Logger
    Private Shared Sub WriteLog(ByVal info As Object, ByVal message As Object)
        Dim filename = Application.StartupPath & "\log\application.log"
        Dim dir As DirectoryInfo = New DirectoryInfo(Application.StartupPath & "\log")
        If Not File.Exists(filename) Then
            dir = Directory.CreateDirectory(dir.FullName)
            Using fs As FileStream = File.Create(filename)
                Using sw As StreamWriter = New StreamWriter(fs)
                    Dim sb As New StringBuilder
                    sb.Append(String.Format("{0:yyyy-MM-dd hh:mm:ss.fff tt}", DateTime.Now) & vbTab)
                    sb.Append("****** Start Logging ******" & vbCrLf)
                    sw.Write(sb.ToString)
                End Using
            End Using
        End If
        Using fs As FileStream = New FileStream(filename, FileMode.Append, FileAccess.Write)
            Using sw As StreamWriter = New StreamWriter(fs)
                Dim sb As New StringBuilder
                sb.Append(String.Format("{0:yyyy-MM-dd hh:mm:ss.fff tt}", DateTime.Now) & vbTab)
                sb.Append(info & vbTab)
                sb.Append(message & vbCrLf)
                sw.Write(sb.ToString)
            End Using
        End Using
    End Sub

    Public Shared Sub log(ByVal message As String)
        Dim stackframe As New Diagnostics.StackFrame(1)
        WriteLog(stackframe.GetMethod.Name.ToString, message)
    End Sub

    Public Shared Sub log(ByVal info As String, ByVal message As String)
        WriteLog(info, message)
    End Sub
End Class
