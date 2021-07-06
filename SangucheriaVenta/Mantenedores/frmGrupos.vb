Public Class frmGrupos
    Private Sub Profiles_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If gEditId > 0 Then
            gSwEdit = True

            If Searching("select * from grupos where GrupoId = " & gEditId & ";") Then
                txtId.Text = Dr("GrupoId")
                txtGrupo.Text = Dr("Grupo")

            End If
            Dr.Close()
        Else
            gSwEdit = False
            txtId.Text = ""
            txtGrupo.Text = ""

        End If

        txtGrupo.Focus()
    End Sub
    Private Sub Botones(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGuardar.Click, btnCancelar.Click
        Select Case DirectCast(sender, Button).Text
            Case "Guardar" : Guardar()
            Case "Cancelar" : Me.Close()
        End Select
    End Sub
    Sub Guardar()
        If txtGrupo.Text.Length > 0 Then

            estaCambiando = True

            If txtId.Text.Trim <> "" Then
                sProcess("UPDATE grupos set Grupo='" & txtGrupo.Text & "',Icono='' WHERE GrupoId='" & txtId.Text.Trim & "';")
            Else
                sProcess("INSERT INTO grupos (Grupo,Icono) Values('" & txtGrupo.Text & "','');")
            End If
            Me.Close()
        Else
            MessageBox.Show("Información incompleta(*)", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
        gSwEdit = False
    End Sub
End Class