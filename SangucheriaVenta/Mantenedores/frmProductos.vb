Public Class frmProductos
    Private Async Sub Profiles_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim dt = Await cmbLoadAsync("select * from grupos;")

        cmbGrupo.DataSource = dt
        cmbGrupo.ValueMember = "GrupoId"
        cmbGrupo.DisplayMember = "Grupo"

        dt = Await cmbLoadAsync("select * from receta;")

        cmbReceta.DataSource = dt
        cmbReceta.ValueMember = "RecetaId"
        cmbReceta.DisplayMember = "Receta"

        If gEditId > 0 Then
            gSwEdit = True

            If Searching("select * from productos where ProductoId = " & gEditId & ";") Then
                txtId.Text = Dr("ProductoId")
                txtProducto.Text = Dr("Producto")
                txtDetalle.Text = Dr("Detalle")
                txtPrecio.Text = Dr("Precio")
                cmbGrupo.SelectedValue = Dr("GrupoId")
                cmbReceta.SelectedValue = Dr("RecetaId")
            End If
            Dr.Close()
        Else
            gSwEdit = False
            txtId.Text = ""
            txtProducto.Text = ""
            txtDetalle.Text = ""
            txtPrecio.Text = ""
            cmbGrupo.Text = ""
            cmbReceta.Text = ""
        End If

        txtProducto.Focus()
    End Sub
    Private Sub Botones(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGuardar.Click, btnCancelar.Click
        Select Case DirectCast(sender, Button).Text
            Case "Guardar" : Guardar()
            Case "Cancelar" : Me.Close()
        End Select
    End Sub
    Sub Guardar()
        Dim id, prod, det, prec, afe, gru, rec As String
        id = txtId.Text
        prod = txtProducto.Text
        det = txtDetalle.Text
        prec = txtPrecio.Text
        afe = "No"
        gru = IIf(CStr(cmbGrupo.SelectedValue) = "", "0", CStr(cmbGrupo.SelectedValue))
        rec = IIf(CStr(cmbReceta.SelectedValue) = "", "0", CStr(cmbReceta.SelectedValue))

        If prod.Length > 0 And gru > 0 And prec.Trim <> "" Then

            estaCambiando = True

            If txtId.Text.Trim <> "" Then
                sProcess("UPDATE productos set Producto='" & prod & "',Detalle='" & det & "',Precio=" & prec & ",Afecto='" & afe & "',GrupoId=" & gru & ",RecetaId=" & rec & " WHERE ProductoId='" & id.Trim & "';")
            Else
                sProcess("INSERT INTO productos (Producto,Detalle,Precio,Afecto,GrupoId,RecetaId) Values ('" & prod & "','" & det & "'," & prec & ",'" & afe & "'," & gru & "," & rec & ");")
            End If
            Me.Close()
        Else
            MessageBox.Show("Información incompleta(*)", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
        gSwEdit = False
    End Sub
End Class