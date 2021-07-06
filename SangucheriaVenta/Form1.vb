Imports System.IO

Public Class Form1
    Private fCargado As Boolean
    Private stringToPrint As String

    Sub Calcular()
        Dim i As Integer = 0
        Dim sTotal As Long = 0
        Dim Total As Long = 0
        Dim tIng As Long = 0

        While i < DGVPEDIDO.Rows.Count
            sTotal = sTotal + CInt(DGVPEDIDO.Rows(i).Cells("txtSubTotal2").Value)

            i += 1
        End While
        txtSubTotal.Text = sTotal

        i = 0
        While i < DGVING.Rows.Count
            tIng = tIng + CInt(DGVING.Rows(i).Cells("txtPrecioIng3").Value)
            i += 1
        End While
        txtTotalExtra.Text = CDbl(txtExtra.Text) * CDbl(txtPrecioExtra.Text)

        Total = sTotal + CDbl(txtPrecioTraslado.Text) + tIng + CDbl(txtTotalExtra.Text)
        txtPrecioIngredientes.Text = tIng
        txtTotal.Text = Total
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        fCargado = False
        cnStr = "server=localhost;  user id=root; password=;port=3306;database=sdb;Connection Timeout=28800;"

        Dim i As Integer = 0
        Dim fil As Integer = 0

        'DGVGRUPOS.Columns(0).DefaultCellStyle.BackColor = Color.FromArgb(18, 2, 38)
        'DGVGRUPOS.Columns(0).DefaultCellStyle.ForeColor = Color.White

        If fDataRead("select Grupo from Grupos;") Then
            While Dr.Read
                If i = 0 Then
                    DGVGRUPOS.Rows.Add()
                    DGVGRUPOS.Rows(fil).Height = 59
                End If

                'DGVGRUPOS.Rows(fil).DefaultCellStyle.BackColor = Color.FromArgb(18, 2, 38)
                'DGVGRUPOS.Rows(fil).DefaultCellStyle.ForeColor = Color.White

                DGVGRUPOS.Rows(fil).Cells(i).Value = Dr("Grupo")

                i += 1

                If i = 4 Then
                    i = 0
                    fil += 1
                End If
            End While
        End If
        Dr.Close()

        cmbLoad(cmbDestino, "select DestinoId, Destino from Destinos order by Destino;", "", "DestinoId", "Destino")

        If Searching("select * from Destinos where DestinoId = " & cmbDestino.SelectedValue & ";") Then
            txtPrecioTraslado.Text = Dr("Precio")
        End If
        Dr.Close()

        Calcular()
        fCargado = True
    End Sub

    Private Sub DGVGRUPOS_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGVGRUPOS.CellContentClick
        If DGVGRUPOS.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = "" Then
            Exit Sub
        End If

        Dim Grupo As String = DGVGRUPOS.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
        Dim GrupoId As String = ""

        If Searching("select * from Grupos where Grupo = '" & Grupo & "';") Then
            GrupoId = Dr("GrupoId")
        End If
        Dr.Close()

        If fDataRead("select * from Productos where GrupoId = " & GrupoId & " order by Producto asc;") Then
            Dim i As Integer = 0
            DGVPRODUCTOS.Rows.Clear()

            While Dr.Read
                DGVPRODUCTOS.Rows.Add(Dr("ProductoId"), Dr("Producto"))
                'DGVPRODUCTOS.Rows(i).Height = 50

                'DGVPRODUCTOS.Rows(i).DefaultCellStyle.ForeColor = Color.Green

                i += 1
            End While

            DGVPRODUCTOS.Refresh()
        End If
        Dr.Close()
    End Sub

    Private Sub DGVPRODUCTOS_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGVPRODUCTOS.CellContentClick
        If pIng.Visible = True And DGVING.Rows.Count = 0 Then
            MessageBox.Show("Debe seleccionar al menos un ingrediente")

            Exit Sub
        End If

        Dim j As Integer = 0
        While j < DGVING.Rows.Count
            If DGVPRODUCTOS.Rows(e.RowIndex).Cells("txtProductoId").Value <> DGVING.Rows(j).Cells("txtProductoId3").Value Then
                DGVING.Rows(j).Visible = False
            Else
                DGVING.Rows(j).Visible = True
            End If
            j += 1
        End While

        Dim Receta As Integer = 0

        pIng.Visible = False
        Dim existe As Boolean = False

        If Searching("select * from Productos where ProductoId = " & DGVPRODUCTOS.Rows(e.RowIndex).Cells("txtProductoId").Value & ";") Then

            Dim i As Integer = 0
            While i < DGVPEDIDO.Rows.Count
                If DGVPEDIDO.Rows(i).Cells("txtProductoId2").Value = Dr("ProductoId") Then
                    DGVPEDIDO.Rows(i).Cells("txtCantidad2").Value += 1

                    DGVPEDIDO.Rows(i).Cells("txtSubTotal2").Value = DGVPEDIDO.Rows(i).Cells("txtCantidad2").Value * DGVPEDIDO.Rows(i).Cells("txtPrecio2").Value
                    existe = True
                End If
                i += 1
            End While

            Dim fil As Integer = 0
            If Not existe Then
                DGVPEDIDO.Rows.Add(1, Dr("ProductoId"), Dr("Producto"), Dr("Precio"))

                fil = DGVPEDIDO.Rows.Count - 1

                'If DGVPEDIDO.Rows(1).Cells("").Value Then
                DGVPEDIDO.Rows(fil).Cells("txtSubTotal2").Value = DGVPEDIDO.Rows(fil).Cells("txtCantidad2").Value * DGVPEDIDO.Rows(fil).Cells("txtPrecio2").Value
                'End If
            End If
            Receta = Dr("RecetaId")
        End If
        Dr.Close()

        If Not existe Then
            DGVRECETA.Rows.Clear()

            If fDataRead("select * from RecetaDtl where RecetaId = " & Receta & " order By Producto asc;") Then
                DGVPRODUCTOS.Enabled = False

                pIng.Visible = True
                Dim i As Integer = 0

                While Dr.Read
                    DGVRECETA.Rows.Add(Dr("Producto"), Dr("Precio"))

                    'DGVRECETA.Rows(i).Height = 50

                    i += 1
                End While
            End If
            Dr.Close()
        End If

        Calcular()
    End Sub

    Private Sub DGVPEDIDO_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGVPEDIDO.CellContentClick
        Dim recetaid As String = "0"
        Dim grupoid As String = ""

        If DGVPEDIDO.Columns(e.ColumnIndex).HeaderText = "Ingred" Then
            If Searching("select * from Productos where ProductoId = '" & DGVPEDIDO.Rows(e.RowIndex).Cells("txtProductoId2").Value & "';") Then
                grupoid = Dr("GrupoId")
            End If
            Dr.Close()

            If fDataRead("select * from Productos where GrupoId = '" & grupoid & "' order by Producto asc;") Then
                DGVPRODUCTOS.Rows.Clear()

                While Dr.Read
                    DGVPRODUCTOS.Rows.Add(Dr("ProductoId"), Dr("Producto"))
                End While
            End If
            Dr.Close()

            Dim k As Integer = 0
            While k < DGVPRODUCTOS.Rows.Count
                If DGVPRODUCTOS.Rows(k).Cells("txtProductoId").Value = DGVPEDIDO.Rows(e.RowIndex).Cells("txtProductoId2").Value Then

                    DGVPRODUCTOS.CurrentCell = DGVPRODUCTOS.Rows(k).Cells(1)
                End If
                k += 1
            End While

            If Searching("select * from Productos where ProductoId = " & DGVPEDIDO.Rows(e.RowIndex).Cells("txtProductoId2").Value & ";") Then
                recetaid = Dr("RecetaId")
            End If
            Dr.Close()

            If fDataRead("select * from RecetaDtl where RecetaId = " & recetaid & " order by Producto asc;") Then
                DGVPRODUCTOS.Enabled = False
                Dim l As Integer = 0

                While Dr.Read
                    DGVRECETA.Rows.Add(Dr("Producto"), Dr("Precio"))

                    'DGVRECETA.Rows(l).Height = 50

                    l += 1
                End While
            End If
            Dr.Close()

            Dim j As Integer = 0
            Dim cont As Integer = 0
            While j < DGVING.Rows.Count
                If DGVPEDIDO.Rows(e.RowIndex).Cells("txtProductoId2").Value <> DGVING.Rows(j).Cells("txtProductoId3").Value Then
                    DGVING.Rows(j).Visible = False
                Else
                    DGVING.Rows(j).Visible = True

                    cont += 1
                End If
                j += 1
            End While

            If cont > 0 Then
                pIng.Visible = True
            Else
                MessageBox.Show("El producto indicado no posee ingredientes para seleccionar")
            End If
        End If

        If DGVPEDIDO.Columns(e.ColumnIndex).HeaderText = "Eliminar" Then
            Dim i As Integer = 0
            While i < DGVING.Rows.Count
                If DGVPEDIDO.Rows(e.RowIndex).Cells("txtProductoId2").Value = DGVING.Rows(i).Cells("txtProductoId3").Value Then
                    DGVING.Rows.RemoveAt(i)

                    i -= 1
                End If
                i += 1
            End While

            DGVPEDIDO.Rows.RemoveAt(e.RowIndex)

            Calcular()
        End If
    End Sub

    Private Sub btnCancelar_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub cmbDestino_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbDestino.SelectedIndexChanged
        If Not fCargado Then
            Exit Sub
        End If

        If Searching("select * from Destinos where DestinoId = " & cmbDestino.SelectedValue & ";") Then
            txtPrecioTraslado.Text = Dr("Precio")
        End If
        Dr.Close()
    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If DGVPEDIDO.Rows.Count = 0 Then
            MessageBox.Show("Venta sin productos asignados")
            Exit Sub
        End If

        'If (CDbl(txtTotalEfectivo.Text) + CDbl(txtTotalTarjeta.Text)) <> CDbl(txtTotal.Text) Then
        'MessageBox.Show("No coincide los totales ingresados en el pago con el total de la venta")
        'Exit Sub
        'End If

        sProcess("UPDATE Pedidos set Finalizado = 'SI' WHERE PedidoId = " & txtId.Text & ";")

        MessageBox.Show("Venta finalizada")

        nuevo()
    End Sub

    Private Sub DGVRECETA_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGVRECETA.CellContentClick

        DGVING.Rows.Add(
            DGVPRODUCTOS.Rows(DGVPRODUCTOS.CurrentCell.RowIndex).Cells("txtProductoId").Value,
            DGVRECETA.Rows(DGVRECETA.CurrentCell.RowIndex).Cells("txtIngrediente").Value,
            DGVRECETA.Rows(DGVRECETA.CurrentCell.RowIndex).Cells("txtPrecioIng2").Value
        )

    End Sub

    Private Sub DGVING_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGVING.CellContentClick
        If DGVING.Columns(e.ColumnIndex).HeaderText = "Eliminar" Then
            DGVING.Rows.RemoveAt(e.RowIndex)
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim filasVisibles As Integer = 0
        Dim i As Integer = 0

        While i < DGVING.Rows.Count
            If DGVING.Rows(i).Visible = True Then
                filasVisibles += 1
            End If
            i += 1
        End While

        If filasVisibles = 0 Then
            MessageBox.Show("Debe seleccionar al menos un ingrediente")
            Exit Sub
        End If

        Calcular()

        DGVPRODUCTOS.Enabled = True
        DGVRECETA.Rows.Clear()
        pIng.Visible = False
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If txtContacto.Text.Trim = "" Then
            MessageBox.Show("El contacto es obligatorio")
            txtContacto.Focus()
            Exit Sub
        End If
        If DGVPEDIDO.Rows.Count < 1 Then
            MessageBox.Show("El pedido debe tener al menos un producto")
            Exit Sub
        End If

        Dim Fecha, Hora, Mesa, Observacion, UsuarioId, GarzonId, Finalizado, DestinoId, PrecioDestino, PrecioIngredientes, TiempoEspera, Direccion, Contacto, TotalEfectivo, TotalTarjeta, Extra, PrecioExtra As String
        Dim Efectivo, Transferencia, Transbank As Integer

        'Dim Metodopago As String

        Fecha = dtpFecha.Value.ToString("yyyy/MM/dd")
        Hora = dtpHora.Value.ToString("HH:mm")
        Mesa = "DOMICILIO"
        Observacion = txtObservacion.Text
        UsuarioId = "ADMIN"
        GarzonId = "0"
        Finalizado = "NO"
        DestinoId = cmbDestino.SelectedValue
        PrecioDestino = txtPrecioTraslado.Text
        PrecioIngredientes = txtPrecioIngredientes.Text

        'If txtTotalEfectivo.Text > 0 And txtTotalTarjeta.Text > 0 Then
        'MetodoPago = "MIXTO"
        'Else
        'If txtTotalEfectivo.Text > 0 Then
        'MetodoPago = "EFECTIVO"
        'Else
        'MetodoPago = "TARJETA"
        'End If
        'End If

        TiempoEspera = txtTiempoEspera.Text
        Direccion = txtDireccion.Text
        Contacto = txtContacto.Text
        TotalEfectivo = 0 'txtTotalEfectivo.Text
        TotalTarjeta = 0 'txtTotalTarjeta.Text
        Extra = txtExtra.Text
        PrecioExtra = txtPrecioExtra.Text

        Efectivo = IIf(chkEfectivo.Checked, 1, 0)
        Transferencia = IIf(chkTransferencia.Checked, 1, 0)
        Transbank = IIf(chkTransbank.Checked, 1, 0)

        If txtId.Text = "" Then
            sProcess("insert into Pedidos (     Fecha,          Hora,          Mesa,          Observacion,          UsuarioId,          GarzonId,          Finalizado,          DestinoId,          PrecioDestino,          PrecioIngredientes,          TiempoEspera,          Direccion,          Contacto,          TotalEfectivo,          TotalTarjeta,          Extra,          PrecioExtra,         Efectivo,        Transferencia,        Transbank) values 
                                          ('" & Fecha & "','" & Hora & "','" & Mesa & "','" & Observacion & "','" & UsuarioId & "','" & GarzonId & "','" & Finalizado & "','" & DestinoId & "','" & PrecioDestino & "','" & PrecioIngredientes & "','" & TiempoEspera & "','" & Direccion & "','" & Contacto & "','" & TotalEfectivo & "','" & TotalTarjeta & "','" & Extra & "','" & PrecioExtra & "'," & Efectivo & "," & Transferencia & "," & Transbank & ");")

            If Searching("select max(PedidoId) as m from Pedidos;") Then
                txtId.Text = Dr("m")
            End If
            Dr.Close()
        Else
            sProcess("update Pedidos set Observacion='" & Observacion & "',DestinoId=" & DestinoId & ",PrecioDestino=" & PrecioDestino & ",PrecioIngredientes=" & PrecioIngredientes & ",TiempoEspera=" & TiempoEspera & ",Direccion='" & Direccion & "',Contacto='" & Contacto & "',TotalEfectivo=" & TotalEfectivo & ",TotalTarjeta=" & TotalTarjeta & ",Extra=" & Extra & ",PrecioExtra=" & PrecioExtra & ",Efectivo=" & Efectivo & ",Transferencia=" & Transferencia & ",Transbank=" & Transbank & " where PedidoId = " & txtId.Text & ";")
        End If

        sProcess("delete from PedidosDtl where PedidoId = " & txtId.Text & ";")

        Dim i As Integer = 0
        While i < DGVPEDIDO.Rows.Count
            sProcess("insert into PedidosDtl (ProductoId,Precio,Cantidad,Observacion,PedidoId) values ('" & DGVPEDIDO.Rows(i).Cells("txtProductoId2").Value & "'," & DGVPEDIDO.Rows(i).Cells("txtPrecio2").Value & "," & DGVPEDIDO.Rows(i).Cells("txtCantidad2").Value & ",'" & DGVPEDIDO.Rows(i).Cells("txtObservacion2").Value & "'," & txtId.Text & ");")
            i += 1
        End While

        sProcess("delete from PedidosReceta where PedidoId = " & txtId.Text & ";")
        i = 0
        While i < DGVING.Rows.Count
            sProcess("insert into PedidosReceta (Producto,Precio,ProductoId,PedidoId) values ('" & DGVING.Rows(i).Cells("txtIngrediente3").Value & "'," & DGVING.Rows(i).Cells("txtPrecioIng3").Value & "," & DGVING.Rows(i).Cells("txtProductoId3").Value & "," & txtId.Text & ");")
            i += 1
        End While

        MessageBox.Show("Venta ingresada")

        Dim arrImpr As New ArrayList

        If fDataRead("SELECT 
                    g.`Impresora` 
                FROM pedidosdtl pd 
                LEFT JOIN productos p ON p.`ProductoId` = pd.`ProductoId` 
                LEFT JOIN grupos g ON g.`GrupoId` = p.`GrupoId` 
                WHERE pd.PedidoId = " & txtId.Text & " GROUP BY g.`Impresora`;") Then

            While Dr.Read
                arrImpr.Add(Dr("Impresora"))
            End While
        End If
        Dr.Close()

        Dim index As Integer = 0

        While index < arrImpr.Count
            If arrImpr(index) = "Cocina" Then
                imprimirExcelCocina(txtId.Text)
            Else
                imprimirExcelBarra(txtId.Text)
            End If

            index += 1
        End While

        imprimirExcelCliente(txtId.Text)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If fDataRead("select PedidoId,Contacto from Pedidos where Finalizado = 'NO' ORDER BY PedidoId DESC;") Then
            DGVPENDIENTES.Rows.Clear()

            While Dr.Read
                DGVPENDIENTES.Rows.Add(Dr("PedidoId"), Dr("Contacto"))
            End While
        End If
        Dr.Close()

        pPendientes.Visible = True
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If DGVPEDIDO.Rows.Count > 0 Then
            If MessageBox.Show(text:="Existen datos agregados. Desea continuar?", caption:=Me.Text, buttons:=MessageBoxButtons.YesNo) <> DialogResult.Yes Then
                Exit Sub
            End If
        End If

        Dim ts As TimeSpan

        If Searching("select * from Pedidos where PedidoId = " & DGVPENDIENTES.Rows(DGVPENDIENTES.CurrentCell.RowIndex).Cells("txtPedidoId").Value & ";") Then

            fCargado = False

            txtId.Text = Dr("PedidoId")
            dtpFecha.Value = Dr("Fecha")

            ts = Dr("Hora")

            dtpHora.Value = dtpFecha.Value & " " & ts.ToString

            txtObservacion.Text = Dr("Observacion")
            cmbDestino.SelectedValue = Dr("DestinoId")
            txtPrecioTraslado.Text = Dr("PrecioDestino")
            txtPrecioIngredientes.Text = Dr("PrecioIngredientes")
            txtTiempoEspera.Text = Dr("TiempoEspera")
            txtDireccion.Text = Dr("Direccion")
            txtContacto.Text = Dr("Contacto")
            'txtTotalEfectivo.Text = Dr("TotalEfectivo")
            'txtTotalTarjeta.Text = Dr("TotalTarjeta")
            txtExtra.Text = Dr("Extra")
            txtPrecioExtra.Text = Dr("PrecioExtra")

            chkEfectivo.Checked = IIf(Dr("Efectivo") = 1, True, False)
            chkTransferencia.Checked = IIf(Dr("Transferencia") = 1, True, False)
            chkTransbank.Checked = IIf(Dr("Transbank") = 1, True, False)

            'MessageBox.Show(Dr("Transferencia"))
            fCargado = True
        End If
        Dr.Close()


        DGVPEDIDO.Rows.Clear()

        If fDataRead("SELECT
	                    pd.`Cantidad`,
	                    pd.`ProductoId`,
	                    p.`Producto`,
	                    pd.`Precio`,
	                    (pd.`Precio` * pd.Cantidad) AS SubTotal,
	                    pd.`Observacion`
                    FROM PedidosDtl pd
                    INNER JOIN Productos p ON p.`ProductoId` = pd.`ProductoId` where pd.PedidoId = " & txtId.Text & ";") Then
            While Dr.Read
                DGVPEDIDO.Rows.Add(Dr("Cantidad"), Dr("ProductoId"), Dr("Producto"), Dr("Precio"), Dr("SubTotal"), Dr("Observacion"))
            End While
        End If
        Dr.Close()


        DGVING.Rows.Clear()

        If fDataRead("SELECT
                        pr.`ProductoId`,
                        pr.`Producto`,
	                    pr.`Precio`
                    FROM PedidosReceta pr where pr.PedidoId = " & txtId.Text & ";") Then
            While Dr.Read
                DGVING.Rows.Add(Dr("ProductoId"), Dr("Producto"), Dr("Precio"))
            End While
        End If
        Dr.Close()

        Calcular()
        pPendientes.Visible = False
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        pPendientes.Visible = False
    End Sub

    Sub imprimirExcelBarra(ByVal id As String)
        Dim f As Date
        Dim t As TimeSpan

        Dim sSql As String = ""


        Dim App As New Microsoft.Office.Interop.Excel.Application
        Dim Libro As Microsoft.Office.Interop.Excel.Workbook
        Dim Hoja As Microsoft.Office.Interop.Excel.Worksheet

        Libro = App.Workbooks.Open(Application.StartupPath & "\voucher\Voucher Barra.xltx")
        Hoja = Libro.Sheets("Barra")

        sSql = "SELECT
	                p.`PedidoId`,
	                p.Fecha,
	                p.`Hora`,
	                d.`Destino`,
	                p.`TiempoEspera`,
	                p.`Direccion`,
	                p.`Contacto`,
	                p.`MetodoPago`,
	                p.`Extra`,
	                p.`TiempoEspera`, 
	                p.`Observacion`
                FROM Pedidos p 
                LEFT JOIN Destinos d ON d.`DestinoId` = p.`DestinoId` 
                WHERE p.`PedidoId` = " & id & ";"

        If Searching(sSql) Then

            id = Format(Dr("PedidoId"), "0000")

            Hoja.Range("A2").Value = "PEDIDO " & id

            Hoja.Range("B4").Value = Dr("Fecha")
            Hoja.Range("B5").Value = Dr("Contacto")

            Hoja.Range("B6").Value = Convert.ToString(Dr("Hora"))


            Dim min As TimeSpan
            min = TimeSpan.FromMinutes(Dr("TiempoEspera"))

            t = Dr("Hora")

            t = t.Add(min)

            Hoja.Range("B7").Value = Convert.ToString(t)



            Hoja.Range("B19").Value = Dr("Extra")
            Hoja.Range("B20").Value = Dr("Observacion")
            Hoja.Range("B24").Value = Dr("Direccion")
        End If
        Dr.Close()



        sSql = "SELECT
	                pd.`Cantidad`,
	                pro.`Producto`,
	                IFNULL(pr.`Producto`,'') AS Variedad,
	                pd.`PedidoId` 
                FROM pedidosdtl pd 
                LEFT JOIN pedidos p ON p.`PedidoId` = pd.`PedidoId` 
                LEFT JOIN pedidosreceta pr ON pr.`PedidoId` = pd.`PedidoId` AND pr.`ProductoId` = pd.`ProductoId` 
                LEFT JOIN productos pro ON pro.`ProductoId` = pd.`ProductoId` 
                LEFT JOIN grupos g ON g.`GrupoId` = pro.`GrupoId`
                WHERE p.PedidoId = " & id & " AND g.Impresora = 'Barra' ORDER BY pro.Producto;"

        Dim i As Integer = 9
        Dim filInsert As Integer = 18
        Dim altoFila As Double = 0.0


        Dim auxProd As String = ""

        If fDataRead(sSql) Then

            While Dr.Read

                Hoja.Range(filInsert & ":" & filInsert).Insert(Shift:=Microsoft.Office.Interop.Excel.XlDirection.xlDown, CopyOrigin:=Microsoft.Office.Interop.Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove)

                If auxProd <> Dr("Producto") Then

                    i += 2

                    If Dr("Producto").ToString.Length > 15 Or Dr("Variedad").ToString.Length > 15 Then
                        altoFila = 54.6
                    Else
                        altoFila = 27.6
                    End If

                    'Hoja.Range(i + 1 & ":" & i + 1).Insert(Shift:=Microsoft.Office.Interop.Excel.XlDirection.xlDown, CopyOrigin:=Microsoft.Office.Interop.Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow)

                    'Aplicar formato a celdas
                    Hoja.Range(i & ":" & i).RowHeight = altoFila
                    Hoja.Range("A" & i & ":" & "D" & i).Font.Size = 22
                    Hoja.Range("A" & i & ":" & "D" & i).Font.Name = "Century Gothic"
                    Hoja.Range("A" & i & ":" & "D" & i).Font.Bold = True

                    'Asignar los valores
                    Hoja.Range("A" & i).Value = Dr("Cantidad")
                    Hoja.Range("B" & i).Value = Dr("Producto")
                    Hoja.Range("B" & i + 1).Value = Dr("Variedad")

                    'Aplicar Formato a celdas
                    Hoja.Range(i + 1 & ":" & i + 1).RowHeight = 27.6
                    Hoja.Range("A" & i + 1 & ":" & "D" & i + 1).Font.Size = 22
                    Hoja.Range("A" & i + 1 & ":" & "D" & i + 1).Font.Name = "Century Gothic"
                    Hoja.Range("A" & i + 1 & ":" & "D" & i + 1).Font.Bold = False

                    With Hoja.Range("A" & i & ":D" & i).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop)
                        .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                        .ColorIndex = Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic
                        .TintAndShade = 0
                        .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium
                    End With

                    With Hoja.Range("A" & i)
                        .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                        .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlTop
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = Microsoft.Office.Interop.Excel.Constants.xlContext
                        .MergeCells = True
                        .Merge()
                    End With


                    With Hoja.Range("B" & i & ":D" & i)
                        .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                        .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlTop
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = Microsoft.Office.Interop.Excel.Constants.xlContext
                        .MergeCells = True
                        .Merge()
                    End With

                    With Hoja.Range("B" & i + 1 & ":D" & i + 1)
                        .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                        .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlTop
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = Microsoft.Office.Interop.Excel.Constants.xlContext
                        .MergeCells = True
                        .Merge()
                    End With

                Else

                    'Hoja.Range(i + 1 & ":" & i + 1).Insert(Shift:=Microsoft.Office.Interop.Excel.XlDirection.xlDown, CopyOrigin:=Microsoft.Office.Interop.Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow)

                    If Dr("Producto").ToString.Length > 15 Or Dr("Variedad").ToString.Length > 15 Then
                        altoFila = 54.6
                    Else
                        altoFila = 27.6
                    End If

                    Hoja.Range(i + 1 & ":" & i + 1).RowHeight = altoFila
                    Hoja.Range("A" & i + 1 & ":" & "D" & i + 1).Font.Size = 22
                    Hoja.Range("A" & i + 1 & ":" & "D" & i + 1).Font.Name = "Century Gothic"
                    Hoja.Range("A" & i + 1 & ":" & "D" & i + 1).Font.Bold = False

                    Hoja.Range("B" & i + 1).Value = Dr("Variedad")

                    With Hoja.Range("B" & i + 1 & ":D" & i + 1)
                        .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                        .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlTop
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = Microsoft.Office.Interop.Excel.Constants.xlContext
                        .MergeCells = True
                        .Merge()
                    End With
                End If

                'MessageBox.Show(Hoja.Range("A" & filInsert + 2).Value)

                auxProd = Dr("Producto")

                filInsert += 1
                i += 1
            End While

        End If

        Dr.Close()

        Hoja.PrintOutEx(Copies:=1, Collate:=True, IgnorePrintAreas:=False)

        Libro.Close(SaveChanges:=False)
        App.Quit()


        Hoja = Nothing
        Libro = Nothing
        App = Nothing
    End Sub

    Sub imprimirExcelCocina(ByVal id As String)
        Dim f As Date
        Dim t As TimeSpan

        Dim sSql As String = ""


        Dim App As New Microsoft.Office.Interop.Excel.Application
        Dim Libro As Microsoft.Office.Interop.Excel.Workbook
        Dim Hoja As Microsoft.Office.Interop.Excel.Worksheet

        Libro = App.Workbooks.Open(Application.StartupPath & "\voucher\Voucher Cocina.xltx")
        Hoja = Libro.Sheets("Cocina")

        sSql = "SELECT
	                p.`PedidoId`,
	                p.Fecha,
	                p.`Hora`,
	                d.`Destino`,
	                p.`TiempoEspera`,
	                p.`Direccion`,
	                p.`Contacto`,
	                p.`MetodoPago`,
	                p.`Extra`,
	                p.`TiempoEspera`, 
	                p.`Observacion`
                FROM Pedidos p 
                LEFT JOIN Destinos d ON d.`DestinoId` = p.`DestinoId` 
                WHERE p.`PedidoId` = " & id & ";"

        If Searching(sSql) Then

            id = Format(Dr("PedidoId"), "0000")

            Hoja.Range("A2").Value = "PEDIDO " & id

            Hoja.Range("B4").Value = Dr("Fecha")
            Hoja.Range("B5").Value = Dr("Contacto")

            Hoja.Range("B6").Value = Convert.ToString(Dr("Hora"))


            Dim min As TimeSpan
            min = TimeSpan.FromMinutes(Dr("TiempoEspera"))

            t = Dr("Hora")

            t = t.Add(min)

            Hoja.Range("B7").Value = Convert.ToString(t)



            Hoja.Range("B19").Value = Dr("Extra")
            Hoja.Range("B20").Value = Dr("Observacion")
            Hoja.Range("B24").Value = Dr("Direccion")
        End If
        Dr.Close()



        sSql = "SELECT
	                pd.`Cantidad`,
	                pro.`Producto`,
	                IFNULL(pr.`Producto`,'') AS Variedad,
	                pd.`PedidoId` 
                FROM pedidosdtl pd 
                LEFT JOIN pedidos p ON p.`PedidoId` = pd.`PedidoId` 
                LEFT JOIN pedidosreceta pr ON pr.`PedidoId` = pd.`PedidoId` AND pr.`ProductoId` = pd.`ProductoId` 
                LEFT JOIN productos pro ON pro.`ProductoId` = pd.`ProductoId` 
                LEFT JOIN grupos g ON g.`GrupoId` = pro.`GrupoId`
                WHERE p.PedidoId = " & id & " AND g.Impresora = 'Cocina' ORDER BY pro.Producto;"

        Dim i As Integer = 9
        Dim filInsert As Integer = 18
        Dim altoFila As Double = 0.0


        Dim auxProd As String = ""

        If fDataRead(sSql) Then

            While Dr.Read

                Hoja.Range(filInsert & ":" & filInsert).Insert(Shift:=Microsoft.Office.Interop.Excel.XlDirection.xlDown, CopyOrigin:=Microsoft.Office.Interop.Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove)

                If auxProd <> Dr("Producto") Then

                    i += 2

                    If Dr("Producto").ToString.Length > 15 Or Dr("Variedad").ToString.Length > 15 Then
                        altoFila = 54.6
                    Else
                        altoFila = 27.6
                    End If

                    'Hoja.Range(i + 1 & ":" & i + 1).Insert(Shift:=Microsoft.Office.Interop.Excel.XlDirection.xlDown, CopyOrigin:=Microsoft.Office.Interop.Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow)

                    'Aplicar formato a celdas
                    Hoja.Range(i & ":" & i).RowHeight = altoFila
                    Hoja.Range("A" & i & ":" & "D" & i).Font.Size = 22
                    Hoja.Range("A" & i & ":" & "D" & i).Font.Name = "Century Gothic"
                    Hoja.Range("A" & i & ":" & "D" & i).Font.Bold = True

                    'Asignar los valores
                    Hoja.Range("A" & i).Value = Dr("Cantidad")
                    Hoja.Range("B" & i).Value = Dr("Producto")
                    Hoja.Range("B" & i + 1).Value = Dr("Variedad")

                    'Aplicar Formato a celdas
                    Hoja.Range(i + 1 & ":" & i + 1).RowHeight = 27.6
                    Hoja.Range("A" & i + 1 & ":" & "D" & i + 1).Font.Size = 22
                    Hoja.Range("A" & i + 1 & ":" & "D" & i + 1).Font.Name = "Century Gothic"
                    Hoja.Range("A" & i + 1 & ":" & "D" & i + 1).Font.Bold = False

                    With Hoja.Range("A" & i & ":D" & i).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop)
                        .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                        .ColorIndex = Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic
                        .TintAndShade = 0
                        .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium
                    End With

                    With Hoja.Range("A" & i)
                        .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                        .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlTop
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = Microsoft.Office.Interop.Excel.Constants.xlContext
                        .MergeCells = True
                        .Merge()
                    End With


                    With Hoja.Range("B" & i & ":D" & i)
                        .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                        .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlTop
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = Microsoft.Office.Interop.Excel.Constants.xlContext
                        .MergeCells = True
                        .Merge()
                    End With

                    With Hoja.Range("B" & i + 1 & ":D" & i + 1)
                        .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                        .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlTop
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = Microsoft.Office.Interop.Excel.Constants.xlContext
                        .MergeCells = True
                        .Merge()
                    End With

                Else

                    'Hoja.Range(i + 1 & ":" & i + 1).Insert(Shift:=Microsoft.Office.Interop.Excel.XlDirection.xlDown, CopyOrigin:=Microsoft.Office.Interop.Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow)

                    If Dr("Producto").ToString.Length > 15 Or Dr("Variedad").ToString.Length > 15 Then
                        altoFila = 54.6
                    Else
                        altoFila = 27.6
                    End If

                    Hoja.Range(i + 1 & ":" & i + 1).RowHeight = altoFila
                    Hoja.Range("A" & i + 1 & ":" & "D" & i + 1).Font.Size = 22
                    Hoja.Range("A" & i + 1 & ":" & "D" & i + 1).Font.Name = "Century Gothic"
                    Hoja.Range("A" & i + 1 & ":" & "D" & i + 1).Font.Bold = False

                    Hoja.Range("B" & i + 1).Value = Dr("Variedad")

                    With Hoja.Range("B" & i + 1 & ":D" & i + 1)
                        .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                        .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlTop
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = Microsoft.Office.Interop.Excel.Constants.xlContext
                        .MergeCells = True
                        .Merge()
                    End With
                End If

                'MessageBox.Show(Hoja.Range("A" & filInsert + 2).Value)

                auxProd = Dr("Producto")

                filInsert += 1
                i += 1
            End While

        End If

        Dr.Close()

        Hoja.PrintOutEx(Copies:=1, Collate:=True, IgnorePrintAreas:=False)



        Libro.Close(SaveChanges:=False)
        App.Quit()


        Hoja = Nothing
        Libro = Nothing
        App = Nothing
    End Sub

    Sub imprimirExcelCliente(ByVal id As String)
        Dim f As Date
        Dim t As TimeSpan

        Dim sSql As String = ""


        Dim App As New Microsoft.Office.Interop.Excel.Application
        Dim Libro As Microsoft.Office.Interop.Excel.Workbook
        Dim Hoja As Microsoft.Office.Interop.Excel.Worksheet

        Libro = App.Workbooks.Open(Application.StartupPath & "\voucher\Voucher Cliente.xltx")
        Hoja = Libro.Sheets("Cliente")

        sSql = "SELECT
	                p.`PedidoId`,
	                p.Fecha,
	                p.`Hora`,
	                p.`Extra` 
                FROM Pedidos p 
                LEFT JOIN Destinos d ON d.`DestinoId` = p.`DestinoId` 
                WHERE p.`PedidoId` = " & id & ";"

        If Searching(sSql) Then
            Hoja.Range("B6").Value = Dr("Fecha")
            Hoja.Range("B7").Value = Convert.ToString(Dr("Hora"))

            id = Format(Dr("PedidoId"), "0000")
            Hoja.Range("B8").Value = "PEDIDO " & id
        End If
        Dr.Close()

        sSql = "SELECT 
	                pd.`Cantidad`, 
	                pro.`Producto`, 
	                (pd.Cantidad * pd.Precio) as Total 
                FROM pedidosdtl pd 
                LEFT JOIN pedidos p ON p.`PedidoId` = pd.`PedidoId` 
                LEFT JOIN productos pro ON pro.`ProductoId` = pd.`ProductoId` 
                WHERE p.PedidoId = " & id & " ORDER BY pro.Producto;"

        Dim i As Integer = 11
        Dim altoFila As Double = 0.0

        If fDataRead(sSql) Then

            While Dr.Read

                Hoja.Range(i + 1 & ":" & i + 1).Insert(Shift:=Microsoft.Office.Interop.Excel.XlDirection.xlDown, CopyOrigin:=Microsoft.Office.Interop.Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove)

                If Dr("Producto").ToString.Length > 15 Then
                    altoFila = 27.4
                Else
                    altoFila = 17.4
                End If

                'Asignar los valores
                Hoja.Range("A" & i).Value = Dr("Cantidad")
                Hoja.Range("B" & i).Value = Dr("Producto")
                Hoja.Range("D" & i).Value = Dr("Total")

                'Hoja.Range("D" & i).NumberFormat = "$#,##0"

                'Aplicar Formato a celdas
                Hoja.Range(i & ":" & i).RowHeight = 17.4
                Hoja.Range("A" & i & ":" & "D" & i).Font.Size = 16
                Hoja.Range("A" & i & ":" & "D" & i).Font.Name = "Century Gothic"
                Hoja.Range("A" & i & ":" & "D" & i).Font.Bold = True

                With Hoja.Range("A" & i)
                    .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .WrapText = True
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = Microsoft.Office.Interop.Excel.Constants.xlContext
                    .MergeCells = True
                    .Merge()
                End With


                With Hoja.Range("B" & i & ":C" & i)
                    .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                    .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlTop
                    .WrapText = True
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = Microsoft.Office.Interop.Excel.Constants.xlContext
                    .MergeCells = True
                    .Merge()
                End With

                With Hoja.Range("D" & i)
                    .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlRight
                    .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlTop
                    .WrapText = True
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = Microsoft.Office.Interop.Excel.Constants.xlContext
                    .MergeCells = True
                    .Merge()
                End With

                i += 1
            End While
        End If

        Dr.Close()

        Hoja.PrintOutEx(Copies:=1, Collate:=True, IgnorePrintAreas:=False)
        Libro.Close(SaveChanges:=False)
        App.Quit()


        Hoja = Nothing
        Libro = Nothing
        App = Nothing
    End Sub

    Sub imprimir()
        Dim id As String
        Dim f As Date
        Dim t As TimeSpan

        Dim strPrint As String = ""
        Dim sSql As String = ""

        sSql = "SELECT
	                p.`PedidoId`,
	                p.Fecha,
	                p.`Hora`,
	                d.`Destino`,
	                p.`TiempoEspera`,
	                p.`Direccion`,
	                p.`Contacto`,
	                p.`MetodoPago`,
	                p.`Observacion`
                FROM Pedidos p
                LEFT JOIN Destinos d ON d.`DestinoId` = p.`DestinoId`
                WHERE p.`PedidoId` = " & txtId.Text & ";"

        If fDataRead(sSql) Then
            While Dr.Read

                t = Dr("Hora")

                id = Format(Dr("PedidoId"), "0000")

                strPrint = strPrint & "LA TERRAZA RESTOBAR" & vbCrLf
                strPrint = strPrint & "----------------------------------------" & vbCrLf
                strPrint = strPrint & "----------------------------------------" & vbCrLf
                strPrint = strPrint & $"PEDIDO: {id}" & vbCrLf
                strPrint = strPrint & "----------------------------------------" & vbCrLf
                strPrint = strPrint & "FECHA: " & Convert.ToString(FormatDateTime(Dr("Fecha"), DateFormat.ShortDate)) & "      HORA: " & Convert.ToString(FormatDateTime(dtpFecha.Value & " " & t.ToString, DateFormat.ShortTime)) & vbCrLf
                strPrint = strPrint & "----------------------------------------" & vbCrLf
                strPrint = strPrint & "TRASLADO: " & Dr("Destino") & vbCrLf
                strPrint = strPrint & "TIEMPO DE ESPERA: " & Dr("TiempoEspera") & vbCrLf
                strPrint = strPrint & "DIRECCION: " & Dr("Direccion") & vbCrLf
                strPrint = strPrint & "CONTACTO: " & Dr("Contacto") & vbCrLf
                strPrint = strPrint & "FORMA DE PAGO: " & Dr("MetodoPago") & vbCrLf
                strPrint = strPrint & "----------------------------------------" & vbCrLf
                strPrint = strPrint & "----------------------------------------" & vbCrLf
                strPrint = strPrint & "PRODUCTO                    SUBTOTAL" & vbCrLf
                strPrint = strPrint & "----------------------------------------" & vbCrLf
                strPrint = strPrint & "1 EMPANADAS POLLO MANDARIN 4UNI     $1.000" & vbCrLf
                strPrint = strPrint & "(POLLO, PALTA, TOMATE)" & vbCrLf
                strPrint = strPrint & "1 EMPANADAS POLLO MANDARIN 4UNI     $2.000" & vbCrLf
                strPrint = strPrint & "----------------------------------------" & vbCrLf
                strPrint = strPrint & "TOTAL                        $3.000" & vbCrLf & vbCrLf

                strPrint = strPrint & "----------------------------------------" & vbCrLf
                strPrint = strPrint & "----------------------------------------" & vbCrLf
                strPrint = strPrint & "OBSERVACION" & vbCrLf
                strPrint = strPrint & Dr("Observacion") & vbCrLf
                strPrint = strPrint & "----------------------------------------" & vbCrLf

                Dim ruta As String = System.Windows.Forms.Application.StartupPath & "\voucher\fichero.txt"
                Dim escritor As StreamWriter
                escritor = File.AppendText(ruta)
                escritor.Write(strPrint)
                escritor.Flush()
                escritor.Close()

                stringToPrint = strPrint

                PrintDocument1.Print()
                'Printer.Print(strPrint)
            End While
        End If
        Dr.Close()
    End Sub

    Private Sub imprimirTxt()

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If DGVING.Rows.Count = 0 Then
            Dim i As Integer = 0

            While i < DGVING.Rows.Count
                If DGVING.Rows(i).Cells("txtProductoId3").Value = DGVPEDIDO.Rows(DGVPEDIDO.CurrentCell.RowIndex).Cells("txtProductoId2").Value Then
                    DGVING.Rows.RemoveAt(i)
                    i -= 1
                End If
                i += 1
            End While

            DGVPEDIDO.Rows.RemoveAt(DGVPEDIDO.CurrentCell.RowIndex)

            DGVPRODUCTOS.Enabled = True
            pIng.Visible = False

            Calcular()
        Else
            DGVPRODUCTOS.Enabled = True
            pIng.Visible = False

            DGVRECETA.Rows.Clear()
        End If
    End Sub

    Private Sub txtExtra_LostFocus(sender As Object, e As EventArgs) Handles txtExtra.LostFocus
        txtTotalExtra.Text = CDbl(txtExtra.Text) * CDbl(txtPrecioExtra.Text)

        Calcular()
    End Sub

    Private Sub txtPrecioExtra_LostFocus(sender As Object, e As EventArgs) Handles txtPrecioExtra.LostFocus
        txtTotalExtra.Text = CDbl(txtExtra.Text) * CDbl(txtPrecioExtra.Text)

        Calcular()
    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Dim charactersOnPage As Integer = 0
        Dim linesPerPage As Integer = 0
        e.Graphics.MeasureString(stringToPrint, Me.Font, e.MarginBounds.Size, StringFormat.GenericTypographic, charactersOnPage, linesPerPage)
        e.Graphics.DrawString(stringToPrint, Me.Font, Brushes.Black, e.MarginBounds, StringFormat.GenericTypographic)
        stringToPrint = stringToPrint.Substring(charactersOnPage)
        e.HasMorePages = (stringToPrint.Length > 0)
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        nuevo()
    End Sub

    Sub nuevo()
        DGVPRODUCTOS.Rows.Clear()
        DGVRECETA.Rows.Clear()
        DGVPENDIENTES.Rows.Clear()
        DGVING.Rows.Clear()
        DGVPEDIDO.Rows.Clear()

        txtId.Text = ""

        dtpFecha.Value = Date.Now
        dtpHora.Value = Date.Now

        cmbDestino.SelectedIndex = 0
        txtTiempoEspera.Text = "40"

        txtDireccion.Text = ""
        txtContacto.Text = ""

        'txtTotalEfectivo.Text = 0
        'txtTotalTarjeta.Text = 0

        txtExtra.Text = 0
        txtPrecioExtra.Text = 0
        txtTotalExtra.Text = 0

        txtObservacion.Text = ""

        txtSubTotal.Text = 0
        txtPrecioIngredientes.Text = 0

        chkEfectivo.Checked = False
        chkTransferencia.Checked = False
        chkTransbank.Checked = False

        Calcular()
    End Sub

    Private Sub DGVGRUPOS_CellPainting(sender As Object, e As DataGridViewCellPaintingEventArgs) Handles DGVGRUPOS.CellPainting
        Exit Sub

        Dim img As Bitmap

        If DGVGRUPOS.Rows(e.RowIndex).Cells(e.ColumnIndex).Value <> "" Then
            If Searching("SELECT * FROM Grupos WHERE Grupo = '" & DGVGRUPOS.Rows(e.RowIndex).Cells(e.ColumnIndex).Value & "';") Then
                If Dr("Icono") <> "" Then
                    img = Image.FromFile(Application.StartupPath & "/iconos/" & Dr("Icono"))

                    e.Paint(e.CellBounds, DataGridViewPaintParts.All)
                    e.Graphics.DrawImage(img, CInt((e.CellBounds.Width / 2) - (img.Width / 2)) + e.CellBounds.X, CInt((e.CellBounds.Height / 2) - (img.Height / 2)) + e.CellBounds.Y)
                    e.Handled = True
                End If
            End If
            Dr.Close()
        End If
    End Sub

    Private colorActual, colorFuente As Color

    Private Sub DGVGRUPOS_CellMouseEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DGVGRUPOS.CellMouseEnter
        colorActual = DGVGRUPOS.Rows(e.RowIndex).Cells(e.ColumnIndex).Style.BackColor
        colorFuente = DGVGRUPOS.Rows(e.RowIndex).Cells(e.ColumnIndex).Style.ForeColor

        DGVGRUPOS.Rows(e.RowIndex).Cells(e.ColumnIndex).Style.BackColor = Color.FromArgb(201, 0, 44)
        DGVGRUPOS.Rows(e.RowIndex).Cells(e.ColumnIndex).Style.ForeColor = Color.White
    End Sub

    Private Sub DGVGRUPOS_CellMouseLeave(sender As Object, e As DataGridViewCellEventArgs) Handles DGVGRUPOS.CellMouseLeave
        DGVGRUPOS.Rows(e.RowIndex).Cells(e.ColumnIndex).Style.BackColor = colorActual
        DGVGRUPOS.Rows(e.RowIndex).Cells(e.ColumnIndex).Style.ForeColor = colorFuente
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click

        If txtContacto.Text = "" Then
            MessageBox.Show("Debe indicar un nombre de contacto")
            txtContacto.Focus()
            Exit Sub
        End If

        Dim Fecha, Hora, Mesa, Observacion, UsuarioId, GarzonId, Finalizado, DestinoId, PrecioDestino, PrecioIngredientes, TiempoEspera, Direccion, Contacto, TotalEfectivo, TotalTarjeta, Extra, PrecioExtra As String
        Fecha = dtpFecha.Value.ToString("yyyy/MM/dd")
        Hora = dtpHora.Value.ToString("HH:mm")
        Mesa = "DOMICILIO"
        Observacion = txtObservacion.Text
        UsuarioId = "ADMIN"
        GarzonId = "0"
        Finalizado = "NO"
        DestinoId = cmbDestino.SelectedValue
        PrecioDestino = txtPrecioTraslado.Text
        PrecioIngredientes = txtPrecioIngredientes.Text

        Dim Efectivo, Transferencia, Transbank As Integer

        Efectivo = IIf(chkEfectivo.Checked, 1, 0)
        Transferencia = IIf(chkTransferencia.Checked, 1, 0)
        Transbank = IIf(chkTransbank.Checked, 1, 0)

        'If txtTotalEfectivo.Text > 0 And txtTotalTarjeta.Text > 0 Then
        'MetodoPago = "MIXTO"
        'Else
        'If txtTotalEfectivo.Text > 0 Then
        'MetodoPago = "EFECTIVO"
        'Else
        'MetodoPago = "TARJETA"
        'End If
        'End If

        TiempoEspera = txtTiempoEspera.Text
        Direccion = txtDireccion.Text
        Contacto = txtContacto.Text
        TotalEfectivo = 0 'txtTotalEfectivo.Text
        TotalTarjeta = 0 'txtTotalTarjeta.Text
        Extra = txtExtra.Text
        PrecioExtra = txtPrecioExtra.Text

        If txtId.Text = "" Then
            sProcess("insert into Pedidos (     Fecha,          Hora,          Mesa,          Observacion,          UsuarioId,          GarzonId,          Finalizado,          DestinoId,          PrecioDestino,          PrecioIngredientes,          TiempoEspera,          Direccion,          Contacto,          TotalEfectivo,          TotalTarjeta,          Extra,          PrecioExtra,         Efectivo,        Transferencia,        Transbank) values 
                                          ('" & Fecha & "','" & Hora & "','" & Mesa & "','" & Observacion & "','" & UsuarioId & "','" & GarzonId & "','" & Finalizado & "','" & DestinoId & "','" & PrecioDestino & "','" & PrecioIngredientes & "','" & TiempoEspera & "','" & Direccion & "','" & Contacto & "','" & TotalEfectivo & "','" & TotalTarjeta & "','" & Extra & "','" & PrecioExtra & "'," & Efectivo & "," & Transferencia & "," & Transbank & ");")

            If Searching("select max(PedidoId) as m from Pedidos;") Then
                txtId.Text = Dr("m")
            End If
            Dr.Close()
        Else
            sProcess("update Pedidos set Observacion='" & Observacion & "',DestinoId=" & DestinoId & ",PrecioDestino=" & PrecioDestino & ",PrecioIngredientes=" & PrecioIngredientes & ",TiempoEspera=" & TiempoEspera & ",Direccion='" & Direccion & "',Contacto='" & Contacto & "',TotalEfectivo=" & TotalEfectivo & ",TotalTarjeta=" & TotalTarjeta & ",Extra=" & Extra & ",PrecioExtra=" & PrecioExtra & ",Efectivo=" & Efectivo & ",Transferencia=" & Transferencia & ",Transbank=" & Transbank & " where PedidoId = " & txtId.Text & ";")
        End If

        sProcess("delete from PedidosDtl where PedidoId = " & txtId.Text & ";")

        Dim i As Integer = 0
        While i < DGVPEDIDO.Rows.Count
            sProcess("insert into PedidosDtl (ProductoId,Precio,Cantidad,Observacion,PedidoId) values ('" & DGVPEDIDO.Rows(i).Cells("txtProductoId2").Value & "'," & DGVPEDIDO.Rows(i).Cells("txtPrecio2").Value & "," & DGVPEDIDO.Rows(i).Cells("txtCantidad2").Value & ",'" & DGVPEDIDO.Rows(i).Cells("txtObservacion2").Value & "'," & txtId.Text & ");")
            i += 1
        End While

        sProcess("delete from PedidosReceta where PedidoId = " & txtId.Text & ";")
        i = 0
        While i < DGVING.Rows.Count
            sProcess("insert into PedidosReceta (Producto,Precio,ProductoId,PedidoId) values ('" & DGVING.Rows(i).Cells("txtIngrediente3").Value & "'," & DGVING.Rows(i).Cells("txtPrecioIng3").Value & "," & DGVING.Rows(i).Cells("txtProductoId3").Value & "," & txtId.Text & ");")
            i += 1
        End While

        MessageBox.Show("Venta pendiente con el número " & txtId.Text)

        nuevo()
    End Sub

    Private Sub DGVPRODUCTOS_CellMouseLeave(sender As Object, e As DataGridViewCellEventArgs) Handles DGVPRODUCTOS.CellMouseLeave
        If e.RowIndex = -1 Then
            Exit Sub
        End If

        DGVPRODUCTOS.Rows(e.RowIndex).Cells(e.ColumnIndex).Style.BackColor = colorActual
        DGVPRODUCTOS.Rows(e.RowIndex).Cells(e.ColumnIndex).Style.ForeColor = colorFuente
    End Sub

    Private Sub DGVPRODUCTOS_CellMouseEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DGVPRODUCTOS.CellMouseEnter
        If e.RowIndex = -1 Then
            Exit Sub
        End If

        colorActual = DGVPRODUCTOS.Rows(e.RowIndex).Cells(e.ColumnIndex).Style.BackColor
        colorFuente = DGVPRODUCTOS.Rows(e.RowIndex).Cells(e.ColumnIndex).Style.ForeColor

        DGVPRODUCTOS.Rows(e.RowIndex).Cells(e.ColumnIndex).Style.BackColor = Color.FromArgb(201, 0, 44)
        DGVPRODUCTOS.Rows(e.RowIndex).Cells(e.ColumnIndex).Style.ForeColor = Color.White
    End Sub

    Private Sub DGVRECETA_CellMouseLeave(sender As Object, e As DataGridViewCellEventArgs) Handles DGVRECETA.CellMouseLeave
        If e.RowIndex = -1 Then
            Exit Sub
        End If

        DGVRECETA.Rows(e.RowIndex).Cells(e.ColumnIndex).Style.BackColor = colorActual
        DGVRECETA.Rows(e.RowIndex).Cells(e.ColumnIndex).Style.ForeColor = colorFuente
    End Sub

    Private Sub DGVRECETA_CellMouseEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DGVRECETA.CellMouseEnter
        If e.RowIndex = -1 Then
            Exit Sub
        End If

        colorActual = DGVRECETA.Rows(e.RowIndex).Cells(e.ColumnIndex).Style.BackColor
        colorFuente = DGVRECETA.Rows(e.RowIndex).Cells(e.ColumnIndex).Style.ForeColor

        DGVRECETA.Rows(e.RowIndex).Cells(e.ColumnIndex).Style.BackColor = Color.FromArgb(201, 0, 44)
        DGVRECETA.Rows(e.RowIndex).Cells(e.ColumnIndex).Style.ForeColor = Color.White
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        estaCambiando = False

        frmDestinosGrid.Show()
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        estaCambiando = False

        frmGruposGrid.Show()
    End Sub

    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        estaCambiando = False

        frmProductosGrid.Show()
    End Sub

    Private Sub LinkLabel3_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked
        estaCambiando = False

        frmRecetasGrid.Show()
    End Sub

    Sub ProcesoImprimir()
        Dim strPrint As String
        strPrint = "LA TERRAZA RESTOBAR" & vbCrLf
        strPrint = strPrint & "------------------------------" & vbCrLf
        strPrint = strPrint & "Nro     : TN1254389" & vbCrLf
        strPrint = strPrint & "Cashier: Soni" & vbCrLf
        strPrint = strPrint & " " & vbCrLf
        strPrint = strPrint & "Nama   Qty. Costs SubTotal" & vbCrLf
        strPrint = strPrint & "------------------------------" & vbCrLf
        strPrint = strPrint & "Sauce    2   5000    10000" & vbCrLf
        strPrint = strPrint & "Coffe    3   1000     3000" & vbCrLf
        strPrint = strPrint & "Sugar    1   8000     3000" & vbCrLf
        strPrint = strPrint & "------------------------------" & vbCrLf
        strPrint = strPrint & "Total                13000" & vbCrLf
        Printer.Print(strPrint)
    End Sub
End Class
