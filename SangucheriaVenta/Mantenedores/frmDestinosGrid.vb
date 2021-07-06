Imports MySql.Data.MySqlClient
Imports System.Threading
Public Class frmDestinosGrid
    Private bindingSource1 As New BindingSource()
    Private dataAdapter As New MySqlDataAdapter
    Private sSql As String = "select * from destinos;"

    Private Sub DGV_SortStringChanged(sender As Object, e As EventArgs) Handles DGV.SortStringChanged
        bindingSource1.Sort = DGV.SortString
    End Sub
    Private Sub DGV_FilterStringChanged(sender As Object, e As EventArgs) Handles DGV.FilterStringChanged
        bindingSource1.Filter = DGV.FilterString
    End Sub

    Private Async Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.CheckForIllegalCrossThreadCalls = False
        Me.DGV.DataSource = Me.bindingSource1

        Dim data = Await GetDataAsync(sSql)
        Me.DGV.DataSource = data

        'GetData("select * from centroscostos;")



        'GrillaLoad(DGV, "select CategoriaId as Id, Categoria from tblcategorias " & IIf(txtBuscar.Text.Trim.Length > 0, " Where " & campo & " Like '%" & txtBuscar.Text & "%'", ""), True)
        LoadImageDGV(DGV, Application.StartupPath & "\iconos\pencil.ico", "Editar", "imgEdit")
        LoadImageDGV(DGV, Application.StartupPath & "\iconos\delete.ico", "Eliminar", "imgDelete")

        'DGV.Columns(0).Width = 50  ' Idfactura
        DGV.Columns(1).Width = 500 ' Fecha
        'DGV.Columns(2).Width = 40
        'DGV.Columns(3).Width = 40

        'DGV.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
    End Sub


    Private Async Sub ButtonClick(sender As Object, e As EventArgs) Handles btnSalir.Click, btnNuevo.Click
        Select Case DirectCast(sender, Button).Text.Trim
            Case "Nuevo" : gEditId = 0 : frmDestinos.ShowDialog()
                Dim data = Await GetDataAsync(sSql)
                Me.DGV.DataSource = data
            Case "Excel"
            Case "Printer"
            Case "Salir" : Me.Close()
        End Select
    End Sub
    Private Async Sub DGV_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGV.CellContentClick

        gEditId = DGV.Item("DestinoId", Me.DGV.CurrentRow.Index).Value
        Select Case DGV.Columns(e.ColumnIndex).HeaderText
            Case "Editar"
                frmDestinos.ShowDialog()

                Dim data = Await GetDataAsync(sSql)

                Me.DGV.DataSource = data
            Case "Eliminar" ' Elimina
                If MessageBox.Show("Desea eliminar el item seleccionado?", "Eliminar", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2, 0) = DialogResult.Yes Then
                    Try
                        sProcess("DELETE FROM destinos WHERE DestinoId='" & gEditId & "'")

                        DGV.Rows.RemoveAt(e.RowIndex)
                    Catch ex As MySqlException
                        MessageBox.Show(ex.Message)

                    End Try
                End If
        End Select
        gEditId = 0

    End Sub
    Private Sub DGV_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DGV.CellFormatting
        Try
            Select Case e.ColumnIndex
                Case 0 : DGV.Rows(e.RowIndex).Cells().Item(e.ColumnIndex).ToolTipText = "Editar"
                Case 1 : DGV.Rows(e.RowIndex).Cells().Item(e.ColumnIndex).ToolTipText = "Eliminar"
            End Select
        Catch ex As Exception
        End Try
    End Sub
    Sub GenerarExcel()
        Dim App As New Microsoft.Office.Interop.Excel.Application
        Dim Libro As Microsoft.Office.Interop.Excel.Workbook
        Dim Hoja As Microsoft.Office.Interop.Excel.Worksheet
        Try
            Libro = App.Workbooks.Add
            Hoja = Libro.Worksheets.Add()

            Dim i As Integer = 0

            Pb.Visible = True
            Pb.Minimum = 0
            Pb.Maximum = DGV.Rows.Count

            Hoja.Range("A1").Value = "DestinoId"
            Hoja.Range("B1").Value = "Destino"
            Hoja.Range("C1").Value = "Precio"

            While i < DGV.Rows.Count
                Hoja.Range("A" & i + 2).Value = DGV.Rows(i).Cells("DestinoId").Value
                Hoja.Range("B" & i + 2).Value = DGV.Rows(i).Cells("Destino").Value
                Hoja.Range("C" & i + 2).Value = DGV.Rows(i).Cells("Precio").Value

                i += 1

                Pb.Value = i
            End While

            Hoja.Rows.Item(1).Font.Bold = 1
            Hoja.Rows.Item(1).HorizontalAlignment = 3
            Hoja.Columns.AutoFit()
            App.Application.Visible = True

            Hoja = Nothing
            Libro = Nothing
            App = Nothing
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error exporting to excel")
        End Try
        Pb.Visible = False
    End Sub

    Private Async Sub BtnExcel_Click(sender As Object, e As EventArgs) Handles btnExcel.Click
        Dim tExcel2 = New Task(
        Sub()
            GenerarExcel()
        End Sub)

        tExcel2.Start()
        Await tExcel2
    End Sub
End Class