Imports System.Data.SqlClient
Imports SqlConnector
Public Class Form1
    Public Shared ConString As String
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ConString = SqlConnector.CSettings.fnGetConnectionString
        If ConString = Nothing Then
            Dim SetWin As New SqlConnector.frmSettings
            With SetWin
                .OpenerAction = 0
                If .ShowDialog = Windows.Forms.DialogResult.OK Then
                    Application.Restart()
                End If
            End With

        End If

        AddHandler ThreadClass.MinuteDone, AddressOf MinuteDone
        Dim tMinute As New Threading.Thread(AddressOf ThreadClass.GetReportByMinute)
        tMinute.Start()
    End Sub
    Private Sub EditSettingsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditSettingsToolStripMenuItem.Click
        Dim SetWin As New SqlConnector.frmSettings
        With SetWin
            .OpenerAction = 1
            If .ShowDialog = Windows.Forms.DialogResult.OK Then
                Application.Restart()
            Else
                Exit Sub
            End If
        End With
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Dim selDate As Date = DataGridView1.SelectedRows(0).Cells(0).Value
        Dim startDate As Date = selDate
        Dim endDate As Date = DateAdd(DateInterval.Minute, 1, startDate)
        Dim sqlCmd = "select users.userid as [UserID], count(documents.handle) as [FaxCount], SUM(DocFiles.NumPages + LEN(Documents.FCSFile) / 8) AS [PagesCount] from documents inner join users on documents.ownerid = users.handle INNER JOIN DocFiles ON Documents.DocFileDBA = DocFiles.handle " & _
            "WHERE     (dateadd(hour, -5, Documents.CreationTime) BETWEEN '" & startDate & "' AND '" & endDate & "') " & _
            "group by users.userid " & _
            "order by [PagesCount] Desc, [FaxCount] Desc"

        Dim sqlCon As SqlConnection
        Dim ds2 As DataSet
        Dim sqlDs As SqlDataAdapter
        sqlCon = New SqlConnection(SqlConnector.CSettings.fnGetConnectionString)
        sqlCon.Open()

        ds2 = New DataSet
        sqlDs = New SqlDataAdapter(sqlCmd, sqlCon)
        sqlDs.Fill(ds2)
        DataGridView2.DataSource = ds2.Tables(0)
    End Sub
    Public Sub MinuteDone(ByVal table As DataSet)
        DataGridView1.DataSource = table.Tables(0).DefaultView
        DataGridView1.Refresh()
    End Sub
End Class
