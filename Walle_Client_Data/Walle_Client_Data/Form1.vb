Public Class Frm_Principal

    'CRIPTOGRAFIA
    Public Shared UserCript As String = My.Settings.CriptUser
    Public Shared PassCript As String = My.Settings.CriptPass
    Public Shared ClientLocation As String = ""
    Public Shared ClientFourkey As String = ""
    Public Shared ClientCod As String = ""
    Public ChaveFechar As Boolean = False
    Public ChaveDesligar As Boolean = False

    Dim Funcao As New Funcoes
    Public ListaDeArquivosNomes As New ArrayList

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

        ClientFourkey = Funcao.GetUserClient()
        ClientLocation = Funcao.GetLocationPath()
        ClientCod = Funcao.GetUserCod

        Funcao.LerPasta()

        Funcao.ExcluirArquivosProcessadosAntigos(ClientLocation)

        Me.ShowInTaskbar = False
        Me.WindowState = FormWindowState.Minimized

        'Funcao.SubirArquivo()

        t_loop.Enabled = True
        ChaveDesligar = False
        Funcao.SubirArquivo()

    End Sub

    Private Sub t_loop_Tick(sender As Object, e As EventArgs) Handles t_loop.Tick

        Try

            Funcao.SubirArquivo()

        Catch ex As Exception



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
