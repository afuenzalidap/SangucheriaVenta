Public Class Printer
    Private Shared Lines As New Queue(Of String)
    Private Shared _myfontTitle As Font
    Private Shared _myfontBody As Font
    Private Shared prn As Printing.PrintDocument

    Shared Sub New()
        _myfontTitle = New Font("Courier New", 12, FontStyle.Bold, GraphicsUnit.Point)
        _myfontBody = New Font("Courier New", 10, FontStyle.Bold, GraphicsUnit.Point)
        prn = New Printing.PrintDocument
        AddHandler prn.PrintPage, AddressOf PrintPageHandler
    End Sub

    Public Shared Sub Print(ByVal text As String)
        Dim linesarray() = text.Split(New String() {Environment.NewLine}, StringSplitOptions.None)

        For Each line As String In linesarray
            Lines.Enqueue(line)
        Next
        prn.Print()
    End Sub

    Private Shared Sub PrintPageHandler(ByVal sender As Object, ByVal e As Printing.PrintPageEventArgs)

        Dim sf As New StringFormat()
        Dim vpos As Single = e.PageSettings.HardMarginY
        Dim l As Integer = 0

        Do While Lines.Count > 0
            Dim line As String = Lines.Dequeue
            Dim _myFont As Font

            If l = 0 Then
                _myFont = _myfontTitle

                'imagen del logo en ticket
                'e.Graphics.DrawImage(Image.FromFile(Application.StartupPath & "/logo_ticket.jpg"), New Rectangle(1, 1, 90, 90))
            Else
                _myFont = _myfontBody
            End If

            Dim sz As SizeF = e.Graphics.MeasureString(line, _myfont, e.PageSettings.Bounds.Size, sf)

            Dim rct As New RectangleF(e.PageSettings.HardMarginX, vpos, e.PageBounds.Width - e.PageSettings.HardMarginX * 2, e.PageBounds.Height - e.PageSettings.HardMarginY * 2)

            e.Graphics.DrawString(line, _myfont, Brushes.Black, rct)
            vpos += sz.Height

            If vpos > e.PageSettings.Bounds.Height - e.PageSettings.HardMarginY * 2 Then
                e.HasMorePages = True
                Exit Sub
            End If

            l += 1
        Loop
    End Sub
End Class