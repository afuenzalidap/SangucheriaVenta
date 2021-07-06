Imports MySql.Data.MySqlClient
Imports System.Text
Imports Excel = Microsoft.Office.Interop.Excel
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports System.IO

Module Module1
    Public Cn As New MySqlConnection
    Public Da As New MySqlDataAdapter

    Public Dr As MySqlDataReader
    Dim Cmd As MySqlCommand
    Public DS As New DataSet

    Public cnStr As String

    Public gEditId As Integer
    Public gSwEdit As Boolean
    Public estaCambiando As Boolean

    Sub CnOpen()
        Try
            Cn.ConnectionString = cnStr
            'Cn.ConnectionString = "server=192.168.1.5;  user id=ti; password=chilfresh;port=3306;database=BD2019-2020;"
            Cn.Open()
        Catch
        End Try
    End Sub

    Public Async Function cmbLoadAsync(ByVal sSql As String) As Task(Of DataTable)
        Dim Dt As New DataTable
        Dim Da As New MySqlDataAdapter
        Dim Cmd As New MySqlCommand

        CnOpen()
        Cmd.CommandType = CommandType.Text
        Cmd.CommandText = sSql
        Cmd.Connection = Cn
        Da.SelectCommand = Cmd

        Await Da.FillAsync(Dt)
        'Da.Fill(Dt)
        Cn.Close()

        Return Dt
    End Function

    Public Async Function GetDataAsync(ByVal command As String) As Task(Of DataTable)
        Dim dt = New DataTable()

        Using da = New MySqlDataAdapter(command, cnStr)
            Await Task.Run(Function() da.Fill(dt))
        End Using

        Return dt
    End Function

    Async Sub GrillaLoad(ByVal Grid As DataGridView, ByVal nSql As String, ajusta As Boolean)
        CnOpen()
        Dim Cmd As New MySqlCommand("", Cn)
        Dim Tabla As New DataTable

        Cmd.CommandType = CommandType.Text
        Cmd.CommandText = nSql
        Cmd.CommandTimeout = 300

        Dim getDrTask As Common.DbDataReader = Await Cmd.ExecuteReaderAsync
        Tabla.Clear()
        Tabla.Load(getDrTask, LoadOption.OverwriteChanges)
        Grid.DataSource = Tabla
        Grid.Refresh()
        Cn.Close()
    End Sub

    Sub cmbLoad(ByVal oCombo As Object, ByVal sSql As String, ByVal nTabla As String, ByVal nValueMember As String, ByVal nDisplayMember As String)
        Dim Dt As New DataTable
        Dim Dac As New MySqlDataAdapter
        Dim Cmd As New MySqlCommand

        CnOpen()
        Cmd.CommandType = CommandType.Text
        Cmd.CommandText = sSql
        Cmd.Connection = Cn
        Da.SelectCommand = Cmd
        Da.Fill(Dt)
        oCombo.DataSource = Dt
        oCombo.ValueMember = nValueMember
        oCombo.DisplayMember = nDisplayMember
        Cn.Close()
    End Sub
    Sub sProcess(ByVal sSql As String)
        Dim Cmd As New MySqlCommand
        '    Try
        CnOpen()
        Cmd = New MySqlCommand(sSql, Cn)
        Dr = Cmd.ExecuteReader
        Cn.Close()
        '    Catch ex As Exception
        '    MsgBox(ex.Message)
        '   End Try
    End Sub
    Sub LoadImageDGV(DGV As DataGridView, FileName As String, Title As String, name As String)
        Dim DGVImagenCol As New DataGridViewImageColumn()
        Dim inImgPrint As System.Drawing.Image = System.Drawing.Image.FromFile(FileName)
        DGVImagenCol.Image = inImgPrint
        DGV.Columns.Add(DGVImagenCol)
        DGVImagenCol.HeaderText = Title
        DGVImagenCol.Name = name
    End Sub
    Sub LoadCheckDGV(DGV As DataGridView, Title As String, name As String)
        Dim DGVCheckCol As New DataGridViewCheckBoxColumn()
        DGV.Columns.Add(DGVCheckCol)
        DGVCheckCol.HeaderText = Title
        DGVCheckCol.Name = name
        DGVCheckCol.ReadOnly = True
    End Sub
    Function Searching(nSql As String)
        CnOpen()
        Dim Cmd As New MySqlCommand("", Cn)
        Dim Tabla As New DataTable
        Cmd.CommandType = CommandType.Text
        Cmd.CommandText = nSql
        Dr = Cmd.ExecuteReader
        Dr.Read()
        Return (IIf(Dr.HasRows, True, False))
    End Function
    Function fDataRead(nSql As String) As Boolean
        CnOpen()
        Dim Cmd As New MySqlCommand("", Cn)
        Dim Tabla As New DataTable
        Cmd.CommandType = CommandType.Text
        Cmd.CommandText = nSql
        Cmd.CommandTimeout = 0

        Dr = Cmd.ExecuteReader
        Dim retorno As Boolean = IIf(Dr.HasRows, True, False)
        '  Cn.Close()
        Return (retorno)
    End Function



    Sub Bordes(ByVal h As Excel.Worksheet, ByVal Range As String, ByVal Arriba As Boolean, ByVal abajo As Boolean, ByVal izquierda As Boolean, ByVal derecha As Boolean, ByVal bw As Excel.XlBorderWeight)
        Dim formatRange As Excel.Range
        formatRange = h.Range("A1")

        'formatRange.Range(Range).Borders(Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Excel.XlLineStyle.xlLineStyleNone
        'formatRange.Range(Range).Borders(Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Excel.XlLineStyle.xlLineStyleNone
        If izquierda Then
            With formatRange.Range(Range).Borders(Excel.XlBordersIndex.xlEdgeLeft)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic
                .TintAndShade = 0
                .Weight = bw
            End With
        End If
        If Arriba Then
            With formatRange.Range(Range).Borders(Excel.XlBordersIndex.xlEdgeTop)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic
                .TintAndShade = 0
                .Weight = bw
            End With
        End If
        If derecha Then
            With formatRange.Range(Range).Borders(Excel.XlBordersIndex.xlEdgeRight)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic
                .TintAndShade = 0
                .Weight = bw
            End With
        End If
        If abajo Then
            With formatRange.Range(Range).Borders(Excel.XlBordersIndex.xlEdgeBottom)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic
                .TintAndShade = 0
                .Weight = bw
            End With
        End If
        'formatRange.Range(Range).Borders(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlLineStyleNone
        'formatRange.Range(Range).Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlLineStyleNone
    End Sub

    Sub AjustarCelda(ByVal h As Excel.Worksheet, ByVal rango As String, ByVal ho As Excel.Constants, ByVal ve As Excel.Constants, ByVal ajuste As Boolean)
        Dim formatRange As Excel.Range
        formatRange = h.Range("A1")

        With formatRange.Range(rango)
            .HorizontalAlignment = ho
            .VerticalAlignment = ve
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = Excel.Constants.xlContext
            .MergeCells = False

            If ajuste Then
                .Merge()
            End If
        End With
    End Sub

    Sub AjustarTitulo(ByVal h As Excel.Worksheet)
        Dim formatRange As Excel.Range
        formatRange = h.Range("A1")

        With formatRange.Range("B2:U3")
            .HorizontalAlignment = Excel.Constants.xlCenter
            .VerticalAlignment = Excel.Constants.xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = Excel.Constants.xlContext
            .MergeCells = True
            .Merge()
            .Font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleSingle
        End With

        formatRange.Range("C10").NumberFormat = "[$-F800]dddd, mmmm dd, yyyy"

        With formatRange.Range("C5:H5")
            .HorizontalAlignment = Excel.Constants.xlLeft
            .VerticalAlignment = Excel.Constants.xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = Excel.Constants.xlContext
            .MergeCells = False
            .Merge()
        End With
        With formatRange.Range("C6:H6")
            .HorizontalAlignment = Excel.Constants.xlLeft
            .VerticalAlignment = Excel.Constants.xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = Excel.Constants.xlContext
            .MergeCells = False
            .Merge()
        End With
        With formatRange.Range("C7:H7")
            .HorizontalAlignment = Excel.Constants.xlLeft
            .VerticalAlignment = Excel.Constants.xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = Excel.Constants.xlContext
            .MergeCells = False
            .Merge()
        End With
        With formatRange.Range("C8:H8")
            .HorizontalAlignment = Excel.Constants.xlLeft
            .VerticalAlignment = Excel.Constants.xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = Excel.Constants.xlContext
            .MergeCells = False
            .Merge()
        End With
        With formatRange.Range("C9:H9")
            .HorizontalAlignment = Excel.Constants.xlLeft
            .VerticalAlignment = Excel.Constants.xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = Excel.Constants.xlContext
            .MergeCells = False
            .Merge()
        End With
        With formatRange.Range("C10:H10")
            .HorizontalAlignment = Excel.Constants.xlLeft
            .VerticalAlignment = Excel.Constants.xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = Excel.Constants.xlContext
            .MergeCells = False
            .Merge()
        End With

        With formatRange.Range("B15:BA15")
            .HorizontalAlignment = Excel.Constants.xlCenter
            .VerticalAlignment = Excel.Constants.xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = Excel.Constants.xlContext
            .MergeCells = False
        End With

        With formatRange.Range("B5:H10").Borders(Excel.XlBordersIndex.xlEdgeBottom)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .ColorIndex = Excel.Constants.xlAutomatic
            .TintAndShade = 0
            .Weight = Excel.XlBorderWeight.xlThin
        End With

        With formatRange.Range("B5:H10").Borders(Excel.XlBordersIndex.xlInsideHorizontal)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .ColorIndex = Excel.Constants.xlAutomatic
            .TintAndShade = 0
            .Weight = Excel.XlBorderWeight.xlThin
        End With
    End Sub

    Sub interiorCelda(ByVal h As Excel.Worksheet, ByVal rango As String)
        Dim formatRange As Excel.Range
        formatRange = h.Range("A1")

        With formatRange.Range(rango).Interior
            .Pattern = Excel.Constants.xlSolid
            .PatternColorIndex = Excel.Constants.xlAutomatic
            .ThemeColor = Excel.XlThemeColor.xlThemeColorAccent6
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
    End Sub

    Sub Totales(ByVal h As Excel.Worksheet, ByVal rango As String, ByVal negrita As Boolean)
        Dim formatRange As Excel.Range
        formatRange = h.Range("A1")

        With formatRange.Range(rango).Interior
            .Pattern = Excel.Constants.xlSolid
            .PatternColorIndex = Excel.Constants.xlAutomatic
            .ThemeColor = Excel.XlThemeColor.xlThemeColorAccent1
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
        With formatRange.Range(rango)
            .HorizontalAlignment = Excel.Constants.xlCenter
            .VerticalAlignment = Excel.Constants.xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = Excel.Constants.xlContext
            .MergeCells = False
            .Font.Bold = negrita
        End With
    End Sub

    Sub AjustarHoja(ByVal exApp As Excel.Application, ByVal h As Excel.Worksheet)
        exApp.PrintCommunication = False
        With h.PageSetup
            .PrintTitleRows = ""
            .PrintTitleColumns = ""
        End With
        exApp.PrintCommunication = True
        h.PageSetup.PrintArea = ""
        exApp.PrintCommunication = False
        With h.PageSetup
            .LeftMargin = (0.25)
            .RightMargin = (0.25)
            .TopMargin = (0.75)
            .BottomMargin = (0.75)
            .HeaderMargin = (0.3)
            .FooterMargin = (0.3)
            .PrintHeadings = False
            .PrintGridlines = False
            .PrintComments = Excel.XlPrintLocation.xlPrintNoComments
            .PrintQuality = 600
            .CenterHorizontally = False
            .CenterVertically = False
            .Orientation = Excel.XlPageOrientation.xlLandscape
            .Draft = False
            .PaperSize = Excel.XlPaperSize.xlPaperLetter
            .Order = Excel.XlOrder.xlDownThenOver
            .BlackAndWhite = False
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 1
            .PrintErrors = Excel.XlPrintErrors.xlPrintErrorsDisplayed
            .OddAndEvenPagesHeaderFooter = False
            .DifferentFirstPageHeaderFooter = False
            .ScaleWithDocHeaderFooter = True
            .AlignMarginsHeaderFooter = True
        End With
        'exApp.PrintCommunication = True
    End Sub
End Module
