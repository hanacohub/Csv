Imports System.IO
Imports Hanaco.Csv.Parser

Public Class FormCsv

    ''' <summary>
    ''' Form_Load
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub FormCsv_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TextBoxFileName.Text = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) & "\sample.csv"
        DataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders
    End Sub

    ''' <summary>
    ''' [ファイル] - [終了]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ToolStripMenuItemClose_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItemClose.Click
        Me.Close()
    End Sub

    ''' <summary>
    ''' Read
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ButtonRead_Click(sender As Object, e As EventArgs) Handles ButtonRead.Click

        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()

        Dim headerRead As Boolean = False
        Dim config As New CsvConfiguration

        config.AllowVariableFieldCount = True

        Using parser = New CsvParser(TextBoxFileName.Text, config)

            Do Until parser.EndOfData

                Dim rec As String() = parser.Read()

                If Not headerRead Then
                    '1行目をヘッダーとして設定します。
                    Dim header As String = String.Empty
                    For Each itm In rec
                        DataGridView1.Columns.Add(itm, itm)
                        DataGridView1.Columns(itm).DefaultCellStyle.WrapMode = DataGridViewTriState.True
                    Next
                    headerRead = True
                Else
                    DataGridView1.Rows.Add(rec)
                End If

            Loop

            Me.AutoSizeColumnWidth(DataGridView1)

        End Using

    End Sub

    ''' <summary>
    ''' 列幅を調整します。
    ''' </summary>
    ''' <param name="grid">DataGridView コントロール。</param>
    Public Sub AutoSizeColumnWidth(grid As DataGridView)

        For i = 0 To grid.ColumnCount - 1

            Using g As Graphics = Graphics.FromHwnd(grid.Handle)

                Dim sf As New StringFormat(StringFormat.GenericTypographic)
                Dim maxWidth As Single = 0

                For j As Integer = 0 To grid.Rows.Count - 1
                    If grid.Rows(j).Cells(i).Value IsNot Nothing Then
                        Dim text As String = grid.Rows(j).Cells(i).Value.ToString
                        maxWidth = Math.Max(g.MeasureString(text, grid.Font, grid.Width, sf).Width, maxWidth)
                    End If
                Next

                Dim headerStyle As DataGridViewCellStyle = grid.ColumnHeadersDefaultCellStyle
                maxWidth = Math.Max(g.MeasureString(grid.Columns(i).HeaderText, headerStyle.Font, grid.Width, sf).Width, maxWidth)

                grid.Columns(i).Width = CInt(maxWidth) + 8

            End Using

        Next

    End Sub

End Class
