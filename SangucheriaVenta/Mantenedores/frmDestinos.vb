Public Class frmDestinos
    Private Sub Profiles_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If gEditId > 0 Then
            gSwEdit = True

            If Searching("select * from destinos where DestinoId = " & gEditId & ";") Then
                txtId.Text = Dr("DestinoId")
                txtDestino.Text = Dr("Destino")
                txtPrecio.Text = Dr("Precio")

            End If
            Dr.Close()
        Else
            gSwEdit = False
            txtId.Text = ""
            txtDestino.Text = ""
            txtPrecio.Text = ""

        End If

        txtDestino.Focus()
    End Sub
    Private Sub Botones(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGuardar.Click, btnCancelar.Click
        Select Case DirectCast(sender, Button).Text
            Case "Guardar" : Guardar()
            Case "Cancelar" : Me.Close()
        End Select
    End Sub
    Sub Guardar()
        If txtDestino.Text.Length > 0 Then

            estaCambiando = True

            If txtId.Text.Trim <> "" Then
                sProcess("UPDATE destinos set Destino='" & txtDestino.Text & "',Precio=" & txtPrecio.Text.Replace(",", ".") & " WHERE DestinoId='" & txtId.Text.Trim & "';")
            Else
                sProcess("INSERT INTO destinos (Destino,Precio) Values('" & txtDestino.Text & "'," & txtPrecio.Text.Replace(",", ".") & ");")
            End If
            Me.Close()
        Else
            MessageBox.Show("Información incompleta(*)", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
        gSwEdit = False
    End Sub
End Class