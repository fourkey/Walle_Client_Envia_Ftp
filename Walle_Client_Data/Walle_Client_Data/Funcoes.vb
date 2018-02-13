Imports System.Net
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text

Public Class Funcoes

    Dim textoCifrado As Byte()
    Dim sal() As Byte = {&H0, &H1, &H2, &H3, &H4, &H5, &H6, &H5, &H4, &H3, &H2, &H1, &H0}
    Dim Pub As New Util

    Public Sub SubirArquivo()

        Dim args(3) As String
        Dim Caminho As String
        Dim ArrayDeArquivosPendentes As New ArrayList
        Dim fluxoTexto As IO.StreamWriter
        Dim linhaTexto As String = ""
        Dim Local As String
        Dim Passou As Boolean = False
        Dim DataAgora As Date
        Dim PrimeiroCaminho As String = ""

        Pub.Escreve_Log("Verificando arquivos pendentes...")

        ArrayDeArquivosPendentes = LerPasta()

        For i As Integer = 0 To ArrayDeArquivosPendentes.Count - 1

            Passou = True

            Caminho = Frm_Principal.CaminhoFtp & "/Walle/Client/" & Frm_Principal.ClientFourkey & "/File/" & Frm_Principal.ListaDeArquivosNomes.Item(i)

            Try

                UploadFile(ArrayDeArquivosPendentes.Item(i), Caminho, Frm_Principal.UserCript, Frm_Principal.PassCript)

                Try

                    File.Copy(Frm_Principal.ClientLocation & "\Pendentes\" & Frm_Principal.ListaDeArquivosNomes.Item(i),
                         Frm_Principal.ClientLocation & "\Processados\" & Frm_Principal.ListaDeArquivosNomes.Item(i), True)

                    'File.Delete(Frm_Principal.ClientLocation & "\Pendentes\" & Frm_Principal.ListaDeArquivosNomes.Item(i))
                    My.Computer.FileSystem.DeleteFile(Frm_Principal.ClientLocation & "\Pendentes\" & Frm_Principal.ListaDeArquivosNomes.Item(i),
                                                      FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)

                Catch ex As Exception

                    Pub.Escreve_Log("Erro de execução:" & ex.Message)

                End Try

            Catch ex As Exception

                Pub.Escreve_Log("Erro de execução:" & ex.Message)

                Exit For

            End Try


        Next

        If Passou = True Then

            DataAgora = Now

            Local = System.Reflection.Assembly.GetExecutingAssembly().Location

            Local = Local.Replace("Walle_Client_Data.exe", "")
            PrimeiroCaminho = Local & "\" & Frm_Principal.ClientCod & "_" & DataAgora.ToString("yyyy-MM-dd-HH-mm-ss") & ".txt"

            fluxoTexto = New IO.StreamWriter(PrimeiroCaminho)

            fluxoTexto.Close()

            Caminho = Frm_Principal.CaminhoFtp & "/Walle/Client/" & Frm_Principal.ClientFourkey & "/PCSEND/" _
                & Frm_Principal.ClientCod & "_" & DataAgora.ToString("yyyy-MM-dd-HH-mm-ss") & ".txt"

            Try

                Pub.Escreve_Log("Fazendo upload do arquivo " & PrimeiroCaminho)

                UploadFile(PrimeiroCaminho, Caminho, Frm_Principal.UserCript, Frm_Principal.PassCript)

                My.Computer.FileSystem.DeleteFile(PrimeiroCaminho, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)

            Catch ex As Exception

                Pub.Escreve_Log("Erro de execução:" & ex.Message)

            End Try

        End If

        Pub.Escreve_Log("Total de arquivos enviados: " & Frm_Principal.ListaDeArquivosNomes.Count().ToString)

        Frm_Principal.ListaDeArquivosNomes.Clear()



    End Sub

    Public Sub UploadFile(ByVal _FileName As String, ByVal _UploadPath As String, ByVal _FTPUser As String, ByVal _FTPPass As String)
        Dim _FileInfo As New System.IO.FileInfo(_FileName)

        ' Create FtpWebRequest object from the Uri provided
        Dim _FtpWebRequest As System.Net.FtpWebRequest = CType(System.Net.FtpWebRequest.Create(New Uri(_UploadPath)), System.Net.FtpWebRequest)

        ' Provide the WebPermission Credintials
        _FtpWebRequest.Credentials = New System.Net.NetworkCredential(_FTPUser, _FTPPass)

        ' By default KeepAlive is true, where the control connection is not closed
        ' after a command is executed.
        _FtpWebRequest.KeepAlive = False

        ' set timeout for 20 seconds
        _FtpWebRequest.Timeout = 20000

        ' Specify the command to be executed.
        _FtpWebRequest.Method = System.Net.WebRequestMethods.Ftp.UploadFile

        ' Specify the data transfer type.
        _FtpWebRequest.UseBinary = True

        ' Notify the server about the size of the uploaded file
        _FtpWebRequest.ContentLength = _FileInfo.Length

        ' The buffer size is set to 2kb
        Dim buffLength As Integer = 2048
        Dim buff(buffLength - 1) As Byte

        ' Opens a file stream (System.IO.FileStream) to read the file to be uploaded
        Dim _FileStream As System.IO.FileStream = _FileInfo.OpenRead()

        ' Stream to which the file to be upload is written
        Dim _Stream As System.IO.Stream = _FtpWebRequest.GetRequestStream()

        ' Read from the file stream 2kb at a time
        Dim contentLen As Integer = _FileStream.Read(buff, 0, buffLength)

        ' Till Stream content ends
        Do While contentLen <> 0
            ' Write Content from the file stream to the FTP Upload Stream
            _Stream.Write(buff, 0, contentLen)
            contentLen = _FileStream.Read(buff, 0, buffLength)
        Loop

        ' Close the file stream and the Request Stream
        _Stream.Close()
        _Stream.Dispose()
        _FileStream.Close()
        _FileStream.Dispose()

    End Sub

    Public Function GetLocationPath() As String

        Dim LocationString As String = My.Settings.Location & "LOCATION.txt"
        Dim fluxoTexto As IO.StreamReader
        Dim linhaTexto As String = ""

        LocationString = LocationString.Replace("Walle_Client_Data.exe", "")

        Frm_Principal.PathEXE = My.Settings.Location.Replace("Walle_Client_Data.exe", "")

        If IO.File.Exists(LocationString) Then

            fluxoTexto = New IO.StreamReader(LocationString)
            linhaTexto = fluxoTexto.ReadLine
            fluxoTexto.Close()

        End If

        Return linhaTexto

    End Function

    Public Function LerPasta() As ArrayList

        Dim ListaDeArquivos As New ArrayList
        Dim NomeParte As String = ""
        Dim Cont As Integer = 0
        Dim Pasta As String

        Pasta = Frm_Principal.ClientLocation

        If Not Directory.Exists(Pasta & "\Pendentes\") Then

            Try

                Directory.CreateDirectory(Pasta & "\Pendentes\")

            Catch ex As Exception


                Pub.Escreve_Log("WARNING - (Funcoes.LerPasta) Falha na criação da pasta \Pendentes em " & Pasta & ". Verificar liberção de acesso para leitura e gravação.")
                Application.Exit()

            End Try



        End If

        If Not Directory.Exists(Pasta & "\Processados\") Then

            Try

                Directory.CreateDirectory(Pasta & "\Processados\")

            Catch ex As Exception

                Pub.Escreve_Log("CATCH - (Funcoes.LerPasta) Falha na criação da pasta \Processados em " & Pasta & ". Verificar liberção de acesso para leitura e gravação.")
                Application.Exit()

            End Try

        End If

        For Each nome In Directory.GetFiles(Pasta)

            ListaDeArquivos.Add(nome)
            Cont = nome.Length - 1

            While nome.Chars(Cont) <> "\"

                NomeParte = nome.Chars(Cont).ToString & NomeParte

                Cont = Cont - 1

            End While

            Frm_Principal.ListaDeArquivosNomes.Add(NomeParte)

            NomeParte = ""

        Next

        For i As Integer = 0 To ListaDeArquivos.Count - 1

            File.Copy(ListaDeArquivos.Item(i), Pasta & "\Pendentes\" & Frm_Principal.ListaDeArquivosNomes.Item(i))
            'File.Delete(ListaDeArquivos.Item(i))
            My.Computer.FileSystem.DeleteFile(ListaDeArquivos.Item(i), FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)

        Next

        ListaDeArquivos.Clear()
        Frm_Principal.ListaDeArquivosNomes.Clear()
        Cont = 0
        NomeParte = ""

        For Each nome In Directory.GetFiles(Pasta & "\Pendentes")

            ListaDeArquivos.Add(nome)
            Cont = nome.Length - 1

            While nome.Chars(Cont) <> "\"

                NomeParte = nome.Chars(Cont).ToString & NomeParte

                Cont = Cont - 1

            End While

            Frm_Principal.ListaDeArquivosNomes.Add(NomeParte)

            NomeParte = ""

        Next

        Return ListaDeArquivos

    End Function

    Public Function GetUserClient() As String

        Dim chave As New Rfc2898DeriveBytes(Pub.Decifra(My.Settings.KeyGetUser), sal)
        Dim algoritmo = New RijndaelManaged()
        Dim fluxoTexto As IO.StreamReader
        Dim linhaTexto As String
        Dim Cliente As String = ""
        Dim USERCLIENT As String = Frm_Principal.PathEXE & "USERCLIENT.txt"

        USERCLIENT = USERCLIENT.Replace("Walle_Client_Data.exe", "")

        If IO.File.Exists(USERCLIENT) Then

            fluxoTexto = New IO.StreamReader(USERCLIENT)
            linhaTexto = fluxoTexto.ReadLine

            fluxoTexto.Close()

            textoCifrado = Convert.FromBase64String(linhaTexto)

            algoritmo.Key = chave.GetBytes(16)
            algoritmo.IV = chave.GetBytes(16)

            Using StreamFonte = New MemoryStream(textoCifrado)

                Using StreamDestino As New MemoryStream()

                    Using crypto As New CryptoStream(StreamFonte, algoritmo.CreateDecryptor(), CryptoStreamMode.Read)

                        moveBytes(crypto, StreamDestino)

                        Dim bytesDescriptografados() As Byte = StreamDestino.ToArray()
                        Dim mensagemDescriptografada = New UnicodeEncoding().GetString(bytesDescriptografados)

                        Cliente = mensagemDescriptografada

                    End Using

                End Using

            End Using

        Else


            Pub.Escreve_Log("WARNING - (Funcoes.GetUserClient) - Walle: Atenção alguns arquivos foram movidos ou modificados indevidamente, por favor entre em contato com o fornecedor.")
            Application.Exit()

        End If

        Return Cliente

    End Function

    Public Function GetUserCod() As String

        Dim algoritmo = New RijndaelManaged()
        Dim fluxoTexto As IO.StreamReader
        Dim linhaTexto As String = ""
        Dim Cliente As String = ""
        Dim USERCLIENT As String = Frm_Principal.PathEXE & "BIN.txt"

        USERCLIENT = USERCLIENT.Replace("Walle_Client_Data.exe", "")

        If IO.File.Exists(USERCLIENT) Then

            fluxoTexto = New IO.StreamReader(USERCLIENT)
            linhaTexto = fluxoTexto.ReadLine

            fluxoTexto.Close()

        End If

        Return linhaTexto

    End Function

    Private Sub moveBytes(ByVal fonte As Stream, ByVal destino As Stream)
        Dim bytes(2048) As Byte
        Dim contador = fonte.Read(bytes, 0, bytes.Length - 1)
        While (0 <> contador)
            destino.Write(bytes, 0, contador)
            contador = fonte.Read(bytes, 0, bytes.Length - 1)
        End While
    End Sub

    Public Sub ExcluirArquivosProcessadosAntigos(ByVal PastaLocal As String)

        Dim Hoje As Date = Format(Now, "yyyy/MM/dd")
        Dim Caminho As String = PastaLocal & "\Processados"
        Dim NomePart As String = ""
        Dim NomePartFinal As String = ""
        Dim Cont As Integer = 1

        Hoje = Hoje.AddDays(-5)

        For Each nome In Directory.GetFiles(Caminho)

            NomePart = Mid(nome, nome.Length - 22, nome.Length)
            NomePart = NomePart.Replace(".csv", "")
            Cont = 1
            NomePartFinal = ""

            For i As Integer = 0 To NomePart.Length - 1

                If NomePart.Chars(i) = "-" Then

                    If Cont = 1 Or Cont = 2 Then

                        NomePartFinal = NomePartFinal & "/"

                    ElseIf Cont = 3 Then

                        Exit For

                    End If

                    Cont = Cont + 1

                Else

                    NomePartFinal = NomePartFinal & NomePart.Chars(i)

                End If

            Next

            If CDate(NomePartFinal) <= Hoje Then

                My.Computer.FileSystem.DeleteFile(nome, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
                'File.Delete(nome)

            End If

        Next

    End Sub

End Class
