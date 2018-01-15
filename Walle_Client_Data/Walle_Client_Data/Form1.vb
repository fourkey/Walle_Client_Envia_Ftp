Public Class Frm_Principal

    'CRIPTOGRAFIA
    Public Shared UserCript As String
    Public Shared PassCript As String
    Public Shared CaminhoFtp As String
    Public Shared ClientLocation As String = ""
    Public Shared ClientFourkey As String = ""
    Public Shared ClientCod As String = ""
    Public ChaveFechar As Boolean = False
    Public ChaveDesligar As Boolean = False


    Dim Funcao As New Funcoes
    Public ListaDeArquivosNomes As New ArrayList
    Dim Pub As New Util

    Private Sub tb_password_KeyDown(sender As Object, e As KeyEventArgs) Handles tb_password.KeyDown

        If e.KeyCode = Keys.Enter Then

            If tb_password.Text = "fourkey" Then

                For Each processo As Process In Process.GetProcesses()

                    If processo.ProcessName = "Walle_Client" Then

                        processo.Kill()

                    End If

                Next processo

                ChaveFechar = True

                Application.Exit()

            Else

                tb_password.Clear()

            End If

        End If

    End Sub

    Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        If Me.WindowState = FormWindowState.Minimized Then

            Me.Hide()

        End If
    End Sub

    Private Sub Frm_Principal_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.ShowInTaskbar = False
        Me.WindowState = FormWindowState.Minimized

        Pub.Escreve_Log("Iniciando aplicativo Walle_Client_Data")

        UserCript = Pub.Decifra(My.Settings.CriptUser)
        PassCript = Pub.Decifra(My.Settings.CriptPass)
        CaminhoFtp = Pub.Decifra(My.Settings.PathFtp)

        ClientFourkey = Funcao.GetUserClient()
        ClientLocation = Funcao.GetLocationPath()
        ClientCod = Funcao.GetUserCod

        Funcao.LerPasta()

        Funcao.ExcluirArquivosProcessadosAntigos(ClientLocation)

        'Funcao.SubirArquivo()

        t_loop.Enabled = True
        ChaveDesligar = False
        If Pub.VerificaConexaoFtp() = True Then
            Funcao.SubirArquivo()
        Else
            Pub.Escreve_Log("WARNING - (Frm_Principal.Load) - Sem conexão com o FTP para subir os arquivos necessários.")
        End If

    End Sub

    Private Sub t_loop_Tick(sender As Object, e As EventArgs) Handles t_loop.Tick

        Try

            If Pub.VerificaConexaoFtp() = True Then
                Funcao.SubirArquivo()
            Else
                Pub.Escreve_Log("WARNING - (Frm_Principal.Load) - Sem conexão com o FTP para subir os arquivos necessários.")
            End If

        Catch ex As Exception

            Pub.Escreve_Log("CATCH - (Frm_Principal.t_loop_Tick) - Erro: " & ex.Message)

        End Try

    End Sub

    Private Sub Frm_Principal_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

        If ChaveFechar = False Then

            e.Cancel = True

            Me.ShowInTaskbar = False
            Me.WindowState = FormWindowState.Minimized

        End If

    End Sub

End Class
