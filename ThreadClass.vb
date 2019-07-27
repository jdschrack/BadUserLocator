Imports System.Data.SqlClient
Imports SqlConnector
Public Class ThreadClass
    Public Shared ConStr As String
    Public Shared Sub GetReportByMinute()
        ConStr = SqlConnector.CSettings.fnGetConnectionString
        If ConStr = Nothing Then
            Exit Sub
        End If
        Dim startTime As String = Format(DateAdd(DateInterval.Day, -7, Date.Now), "yyyy-MM-dd hh:mm:ss.000")
        'MsgBox(startTime)
        Dim endTime As String = Format(Date.Now, "yyyy-MM-dd hh:mm:ss.000")
        Dim sqlCmd As String = "declare @start as datetime " & _
            "declare @end as datetime " & _
            "set @start = '" & startTime & "' " & _
            "set @end = '" & endTime & "' " & _
            "Declare @table table( " & _
            "[date] DateTime, " & _
            "[FaxCount] int, " & _
            "[PageCount] int) " & _
            "insert into @table SELECT   Convert(varchar(4), datepart(year, dateadd(hour, -5, creationtime))) + '-' + Convert(varchar(4), datepart(month, dateadd(hour, -5, creationtime))) + '- ' + Convert(varchar(4), datepart(day, dateadd(hour, -5, creationtime))) + ' ' + CONVERT(varchar(2), DATEPART(hh, dateadd(hour, -5, Documents.CreationTime))) + ':' + CONVERT(varchar(2), DATEPART(n, Documents.CreationTime)) AS Time,  " & _
            "COUNT(Documents.handle) AS FaxCount, SUM(DocFiles.NumPages + LEN(Documents.FCSFile) / 8) AS PageCount " & _
            "FROM         Documents INNER JOIN " & _
            "DocFiles ON Documents.DocFileDBA = DocFiles.handle " & _
            "inner join users on documents.ownerid = users.handle " & _
            "WHERE     (dateadd(hour, -5, Documents.CreationTime) BETWEEN @start AND @end) " & _
            "GROUP BY DATEPART(year, dateadd(hour, -5, Documents.CreationTime)), DATEPART(month, dateadd(hour, -5, Documents.CreationTime)), DATEPART(day, dateadd(hour, -5, Documents.CreationTime)), DATEPART(hh, dateadd(hour, -5, Documents.CreationTime)), DATEPART(n, Documents.CreationTime) " & _
            "select * from @table " & _
            "where faxcount > 175 or pagecount > 800 " & _
            "order by date"
        Dim sqlCon As SqlConnection
        Dim ds As DataSet
        Dim sqlDs As SqlDataAdapter

        sqlCon = New SqlConnection(SqlConnector.CSettings.fnGetConnectionString)
        sqlCon.Open()
        ds = New DataSet
        sqlDs = New SqlDataAdapter(sqlCmd, sqlCon)
        sqlDs.Fill(ds)
        RaiseEvent MinuteDone(ds)
    End Sub
    Public Shared Event MinuteDone(ByVal table As DataSet)
End Class
