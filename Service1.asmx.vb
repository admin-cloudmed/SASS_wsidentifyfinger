Imports System.Web.Services
Imports Registro
Imports System.Data.SqlClient
Imports System.Globalization
Imports System.Threading

<System.Web.Services.WebService(Namespace := "http://tempuri.org/wsIdentifyFinger/Service1")> _
Public Class Service1
    Inherits System.Web.Services.WebService
    Dim NumeroDoDedo As Integer

#Region " Web Services Designer Generated Code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Web Services Designer.
        InitializeComponent()

        'Add your own initialization code after the InitializeComponent() call

    End Sub

    'Required by the Web Services Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Web Services Designer
    'It can be modified using the Web Services Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub

    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        'CODEGEN: This procedure is required by the Web Services Designer
        'Do not modify it using the code editor.
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

#End Region

    Public Function RemoveDigitoVerificadorParaInclusaoDaDigital(ByVal codigo As String) As String
        If codigo.EndsWith("0") = False Then
            codigo = codigo.Remove(codigo.Length - 1, 1)
            codigo = codigo & "0"
        End If
        Return codigo

    End Function

    <WebMethod()> _
    Public Function CadastraFinger(ByVal codUsuario As String, ByVal template1 As Byte(), ByVal template2 As Byte(), ByVal template3 As Byte(), ByVal codDedo As Integer, ByVal cartaoFormatado As String, ByVal codPrestador As String) As String
        Dim connection As New SqlConnection(Misc.GetDadoWebConfig("conexaoBiometria"))
        Dim stringErroDebug As String

        codUsuario = RemoveDigitoVerificadorParaInclusaoDaDigital(codUsuario)

        Try
            NumeroDoDedo = codDedo

            Dim templateBin1 As New TTemplate : templateBin1.tpt = template1
            Dim templateBin2 As New TTemplate : templateBin2.tpt = template2
            Dim templateBin3 As New TTemplate : templateBin3.tpt = template3

            Dim UsuarioJaExiste As Boolean = False
            Dim DataSetIdentify As New DSIdentify

            cartaoFormatado = Misc.AjustaCartao(codUsuario, Registro.Misc.GetDadoWebConfig("AjusteCartao"), False)

            UsuarioJaExiste = VerificaSeUsuarioExiste(cartaoFormatado)

            If (UsuarioJaExiste) Then
                'Consulta dados do usuário no banco de dados (xml)
                LeDadosDoUsuarioNoBancoDeDados(cartaoFormatado, DataSetIdentify)
            Else
                'Consulta dados do usuário no autorizador da Dual
                LeDadosDoUsuarioNoAutorizador(cartaoFormatado, DataSetIdentify)
            End If

            If (UsuarioJaExiste) Then
                If (DataSetIdentify.BiometriaDP(0).IsTemplate1Null) = True Then 'Existe o registro, mas ainda não existe a biometria cadastrada
                    If (GravaTeplatesNoBanco(cartaoFormatado, templateBin1, templateBin2, templateBin3, codPrestador)) Then
                        insertlog(cartaoFormatado, DateTime.Now, "Impressão Digital Cadastrada com Sucesso.", True, codDedo, codPrestador)
                        Return "/ok"

                    Else
                        insertlog(cartaoFormatado, DateTime.Now, "Erro ao Cadastrar Impressão Digital.", False, codDedo, codPrestador)
                        Return "/erro Erro ao cadastrar Impressão Digital."
                    End If

                Else
                    insertlog(cartaoFormatado, DateTime.Now, "Impressão Digital já cadastrada.", False, codDedo, codPrestador)
                    Return "/erro Impressão Digital já cadastrada."
                End If

            Else
                If (GravaNovosTemplatesDoUsuarioNoBanco(cartaoFormatado, templateBin1, templateBin2, templateBin3, DataSetIdentify, codPrestador)) Then
                    insertlog(cartaoFormatado, DateTime.Now, "Impressão Digital Cadastrada com Sucesso.", True, codDedo, codPrestador)
                    Return "/ok"

                Else
                    insertlog(cartaoFormatado, DateTime.Now, "Erro ao Cadastrar Impressão Digital.", False, codDedo, codPrestador)
                    Return "/erro Erro ao cadastrar Impressão Digital."
                End If
            End If

        Catch ex As Exception
            stringErroDebug = "Erro na rotina de cadastro da digital: " & ex.Message
            insertlog(cartaoFormatado, DateTime.Now, "Erro: " & ex.Message, False, codDedo, codPrestador)
            Return "/erro Comunique erro " & ex.Message
        End Try

    End Function

    <WebMethod()> _
    Public Function validaFinger(ByVal codUsuario As String, ByVal template As Byte(), ByVal codDedo As Integer, ByVal cartaoFormatado As String, ByVal codPrestador As String) As String
        Try
            Dim DS As New DSIdentify
            Dim connection As New SqlConnection(Misc.GetDadoWebConfig("conexaoBiometria"))
            Dim adapter As SqlDataAdapter
            Dim stringCommand As String = ""

            If Misc.GetDadoWebConfig("layout") <> "stacasasauderiopreto" Then
                stringCommand = "SELECT * FROM BiometriaDP WHERE (Codigo = @Codigo) and (Dedo = @Dedo)"

            Else
                stringCommand = "SELECT * FROM BiometriaDP WHERE (Codigo like @Codigo) and (Dedo = @Dedo)"
            End If

            adapter = New SqlDataAdapter(stringCommand, connection)

            cartaoFormatado = Misc.AjustaCartao(codUsuario, Registro.Misc.GetDadoWebConfig("AjusteCartao"), False)
            cartaoFormatado = RemoveDigitoVerificadorParaInclusaoDaDigital(cartaoFormatado)

            Dim templateBin_recebido As New TTemplate
            templateBin_recebido.tpt = template

            'Lendo template binario gravado no banco de dados
            Dim templateGravado1 As Byte()
            Dim templateGravado2 As Byte()
            Dim templateGravado3 As Byte()

            If Misc.GetDadoWebConfig("layout") <> "stacasasauderiopreto" Then
                adapter.SelectCommand.Parameters.AddWithValue("@Codigo", cartaoFormatado)

            Else
                adapter.SelectCommand.Parameters.AddWithValue("@Codigo", cartaoFormatado & "%")
            End If

            adapter.SelectCommand.Parameters.AddWithValue("@Dedo", codDedo)
            adapter.Fill(DS, "BiometriaDP")

            If (DS.BiometriaDP.Rows.Count = 0) Then
                'Biometria não existe, grava no log
                insertlog(cartaoFormatado, DateTime.Now, "Código do usuário não cadastrado.", False, codDedo, codPrestador)
                Return "/erro Código do usuário não cadastrado."
                Exit Function
            End If

            If DS.BiometriaDP(0).IsTemplate1Null = True Or DS.BiometriaDP(0).Template1 Is Nothing = True Then
                'Biometria não existe, grava no log
                insertlog(cartaoFormatado, DateTime.Now, "Impressão Digital não cadastrada.", False, codDedo, codPrestador)
                Return "/erro Impressão Digital não cadastrada."

            Else
                Dim GrFinger As GrFingerXLib.GrFingerXCtrl
                Dim SensibilidadeDeVerificacao As Integer = CInt(Misc.GetDadoWebConfig("SensibilidadeDeVerificacao"))

                'O valor máximo permitido é 200, e o valor mínimo permitido é 10
                If (SensibilidadeDeVerificacao < 10) Or _
                   (SensibilidadeDeVerificacao > 200) Then
                    SensibilidadeDeVerificacao = 25 'Valor default = 25
                End If

                Dim Resultado_de_Parametrizacao As Integer

                templateGravado1 = DS.BiometriaDP(0).Template1
                templateGravado2 = DS.BiometriaDP(0).Template2
                templateGravado3 = DS.BiometriaDP(0).Template3

                Dim resultado, score1, score2, score3 As Integer

                GrFinger = New GrFingerXLib.GrFingerXCtrl
                GrFinger.Initialize()
                Resultado_de_Parametrizacao = GrFinger.SetVerifyParameters(SensibilidadeDeVerificacao, 180, GrFingerXLib.GRConstants.GR_DEFAULT_CONTEXT)

                resultado = GrFinger.Verify(templateBin_recebido.tpt, templateGravado1, score1, GrFingerXLib.GRConstants.GR_DEFAULT_CONTEXT)

                'Verificação com sucesso! Grava no log e retorna /ok
                If (resultado = GrFingerXLib.GRConstants.GR_MATCH) Then
                    insertlog(cartaoFormatado, DateTime.Now, "Biometria Verificada com Sucesso", True, codDedo, codPrestador)
                    Return "/ok "
                End If

                'Impressão digital não reconhecida, verificar novamente com 2o. template
                If (resultado = GrFingerXLib.GRConstants.GR_NOT_MATCH) Then
                    'não passou na primeira, tentando 2a.
                    resultado = GrFinger.Verify(templateBin_recebido.tpt, templateGravado2, score2, GrFingerXLib.GRConstants.GR_DEFAULT_CONTEXT)

                    'Verificação com sucesso! Grava no log e retorna /ok
                    If (resultado = GrFingerXLib.GRConstants.GR_MATCH) Then
                        insertlog(cartaoFormatado, DateTime.Now, "Biometria Verificada com Sucesso", True, codDedo, codPrestador)
                        Return "/ok "
                    End If

                    'Impressão digital não reconhecida, verificar novamente com 3o. template
                    If (resultado = GrFingerXLib.GRConstants.GR_NOT_MATCH) Then
                        'não passou na segunda, tentando 3a. e última
                        resultado = GrFinger.Verify(templateBin_recebido.tpt, templateGravado3, score3, GrFingerXLib.GRConstants.GR_DEFAULT_CONTEXT)

                        'Verificação com sucesso! Grava no log e retorna /ok
                        If (resultado = GrFingerXLib.GRConstants.GR_MATCH) Then
                            insertlog(cartaoFormatado, DateTime.Now, "Biometria Verificada com Sucesso", True, codDedo, codPrestador)
                            Return "/ok "
                        End If

                        'Impressão digital não reconhecida na 3a. tentativa, grava no log e retorna /erro
                        If (resultado = GrFingerXLib.GRConstants.GR_NOT_MATCH) Then
                            insertlog(cartaoFormatado, DateTime.Now, "Biometria Incompativel.(score1=" & score1.ToString & "),(score2=" & score2.ToString & "),(score3=" & score3.ToString & ")", False, codDedo, codPrestador)
                            Return "/erro Digital Incompativel."
                        End If
                    End If
                End If

                If (resultado <> GrFingerXLib.GRConstants.GR_NOT_MATCH) And (resultado <> GrFingerXLib.GRConstants.GR_MATCH) Then
                    Dim a As GrFingerXLib.GRConstants
                    a = resultado
                    insertlog(cartaoFormatado, DateTime.Now, "Erro Desconhecido: " & a, False, codDedo, codPrestador)
                    Return "/erro Digital nao Capturada com Sucesso. Detalhes do erro: " & a.ToString
                End If
            End If
            Return ""

        Catch ex As Exception
            insertlog(cartaoFormatado, DateTime.Now, "Erro ao Verificar Biometria: " & ex.Message, False, codDedo, codPrestador)
            Return "/erro " & ex.Message
        End Try

        insertlog(cartaoFormatado, DateTime.Now, "Falha de Verificacao de Biometria.", False, codDedo, codPrestador)
        Return "/erro Falha de Verificacao de Biometria."
    End Function

    <WebMethod()> _
    Public Function insertlog(ByVal codUsuario As String, ByVal dthr As Date, ByVal desc As String, ByVal verificOK As Boolean, ByVal dedo As Integer, ByVal codPrestador As String) As Boolean
        Try
            Dim ComandoSQL As String
            Dim cartaoFormatado As String = Misc.AjustaCartao(codUsuario, Registro.Misc.GetDadoWebConfig("AjusteCartao"), False)
            cartaoFormatado = RemoveDigitoVerificadorParaInclusaoDaDigital(cartaoFormatado)

            ComandoSQL = "INSERT INTO BiometriaDPLogs(Codigo, DataHora, Descricao, VerificacaoOK, CodDedo, CodigoPrestador) VALUES (@Codigo, @DataHora, @Descricao, @VerificacaoOK, @CodDedo, @CodigoPrestador)"
            Dim myConnection As SqlConnection = New SqlConnection(Misc.GetDadoWebConfig("conexaoBiometria"))
            Dim myCommand As SqlCommand = New SqlCommand(ComandoSQL, myConnection)

            myCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Codigo", System.Data.SqlDbType.VarChar, 50, "Codigo"))
            myCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@DataHora", System.Data.SqlDbType.DateTime, 8, "DataHora"))
            myCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Descricao", System.Data.SqlDbType.VarChar, 200, "Descricao"))
            myCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@VerificacaoOK", System.Data.SqlDbType.Int, 4, "VerificacaoOK"))
            myCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CodDedo", System.Data.SqlDbType.Int, 4, "CodDedo"))
            myCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CodigoPrestador", System.Data.SqlDbType.VarChar, 100, "CodigoPrestador"))

            myCommand.Parameters("@Codigo").Value = cartaoFormatado
            myCommand.Parameters("@DataHora").Value = dthr
            myCommand.Parameters("@Descricao").Value = desc
            myCommand.Parameters("@VerificacaoOK").Value = verificOK
            myCommand.Parameters("@CodDedo").Value = dedo
            myCommand.Parameters("@CodigoPrestador").Value = codPrestador

            myCommand.Connection.Open()
            myCommand.ExecuteNonQuery()
            myCommand.Connection.Close()

            Return True

        Catch ex As Exception
            Return "/Parametros Invalidos" & ex.Message
        End Try

    End Function

#Region " Funções usadas pela WebMethod CadastraFinger "

    Function LeDadosDoUsuarioNoAutorizador(ByVal cartaoFormatado As String, ByRef DataSetIdentify As DSIdentify) As Boolean
        Thread.CurrentThread.CurrentCulture = New CultureInfo("pt-BR")
        cartaoFormatado = RemoveDigitoVerificadorParaInclusaoDaDigital(cartaoFormatado)

        Dim WEBService As New WebReference.Autorizador : WEBService.Url = Misc.GetDadoWebConfig("UrlWebService")
        Dim SexoDoUsuario As String = ""
        Dim DataDeNascimento As String = ""
        Dim CPF_DoUsuario As String = ""
        Dim NomeDaMaeDoUsuario As String = ""
        Dim NomeDoUsuario As String = ""
        Dim RG_DoUsuario As String = ""
        Dim Dados_do_Usuario As String = ""
        Dim stringErroDebug As String
        Dim dtNascimento As Date
        Dim IdadeDoUsuario As Integer
        Dim dsDados As New dsDadosUsuario

        '***** Dados do Usuário HB: *****
        'PES,\"LEANDRO MATEUS ARAUJO MEDINA\",\"10/05/1979\",\"M\",
        '\"SOLANGE ARAUJO MEDINA\",\"28225242890\",\"333068142\"\n"

        '***** Dados do Usuário Santa Casa: *****
        '"<dsDadosUsuario xmlns=\"http://tempuri.org/dsDadosUsuario.xsd\">\n  <DadosUsuario>\n    <Nome>EMERSON DA         'SILVA MOREIRA</Nome>\n    <DataNasc>17/8/1973 00:00:00</DataNasc>\n    <Sexo>Masculino</Sexo>\n            '<CPF>VILMA RODRIGUES DO CARMO</CPF>\n    <RG>2016560-4</RG>\n  </DadosUsuario>\n</dsDadosUsuario>"

        'Para Santa Casa o cartão deve ir com 17 dígitos
        If Misc.GetDadoWebConfig("layout") = "uninet" Then
            Dados_do_Usuario = WEBService.Autoriza(9, cartaoFormatado.Remove(cartaoFormatado.Length - 1, 1).PadLeft(16, "0"), 0, 0, 0, 0, 0, "", "", "", "", Misc.GetDadoWebConfig("Prototipo"))

        ElseIf Misc.GetDadoWebConfig("layout") = "stacasasauderiopreto" Then
            Dados_do_Usuario = WEBService.Autoriza(9, cartaoFormatado, 0, 0, 0, 0, 0, "", "", "", "", Misc.GetDadoWebConfig("Prototipo"))
        End If

        stringErroDebug = "Dados retornados do autorizador: " & Dados_do_Usuario
        DataSetIdentify.Clear()

        If (Dados_do_Usuario.Trim <> "") Then
            Try
                Dim culture As New CultureInfo("pt-BR")
                If Misc.GetDadoWebConfig("layout") = "uninet" Then
                    NomeDoUsuario = Dados_do_Usuario.Split(",")(1).Replace("""", "")
                    DataDeNascimento = Dados_do_Usuario.Split(",")(2).Replace("""", "")
                    SexoDoUsuario = Dados_do_Usuario.Split(",")(3).Replace("""", "")
                    NomeDaMaeDoUsuario = Dados_do_Usuario.Split(",")(4).Replace("""", "")
                    CPF_DoUsuario = Dados_do_Usuario.Split(",")(5).Replace("""", "")
                    RG_DoUsuario = Dados_do_Usuario.Split(",")(6).Replace("""", "")
                    RG_DoUsuario = RG_DoUsuario.Replace(Chr(13) & Chr(10), "")
                    RG_DoUsuario = RG_DoUsuario.Replace(Chr(13), "")
                    RG_DoUsuario = RG_DoUsuario.Replace(Chr(10), "")

                ElseIf Misc.GetDadoWebConfig("layout") = "stacasasauderiopreto" Then
                    Misc.LoadXML(dsDados, Dados_do_Usuario)
                    NomeDoUsuario = dsDados.DadosUsuario(0).Nome
                    dtNascimento = Convert.ToDateTime(dsDados.DadosUsuario(0).DataNasc, culture)
                    SexoDoUsuario = dsDados.DadosUsuario(0).Sexo
                    NomeDaMaeDoUsuario = dsDados.DadosUsuario(0).NomeMae
                    CPF_DoUsuario = dsDados.DadosUsuario(0).CPF
                    RG_DoUsuario = dsDados.DadosUsuario(0).RG
                    RG_DoUsuario = RG_DoUsuario.Replace(Chr(13) & Chr(10), "")
                    RG_DoUsuario = RG_DoUsuario.Replace(Chr(13), "")
                    RG_DoUsuario = RG_DoUsuario.Replace(Chr(10), "")
                End If

                'Adicionando no dataset a linha databin passando os parametros que chegaram do layout
                If Misc.GetDadoWebConfig("layout") <> "stacasasauderiopreto" Then
                    Dim DataNascimentoTeste As DateTime
                    DataNascimentoTeste = Convert.ToDateTime(DataDeNascimento, culture)

                    IdadeDoUsuario = getIdade(DataNascimentoTeste)
                    DataSetIdentify.DataBin.AddDataBinRow(NomeDoUsuario, SexoDoUsuario, DataNascimentoTeste, "", "", NomeDaMaeDoUsuario, CPF_DoUsuario, RG_DoUsuario, IdadeDoUsuario)

                Else
                    IdadeDoUsuario = getIdade(dtNascimento)
                    DataSetIdentify.DataBin.AddDataBinRow(NomeDoUsuario, SexoDoUsuario, dtNascimento, "", "", NomeDaMaeDoUsuario, CPF_DoUsuario, RG_DoUsuario, IdadeDoUsuario)
                End If

                Return True

            Catch ex As Exception
                stringErroDebug = "Erro ao ler dados do usuário no autorizador: " & ex.Message
                insertlog(cartaoFormatado, DateTime.Now, "ERRO: Resposta do autorizador inválida, Erro=" & ex.Message, False, NumeroDoDedo, "")
                Return False
            End Try

        Else
            insertlog(cartaoFormatado, DateTime.Now, "ERRO: Resposta do autorizador está vazia", False, NumeroDoDedo, "")
            Return False
        End If

    End Function

    Function CarregaDatasetDoUsuario(ByVal codigoUsuario As String, ByRef ds As DSIdentify) As Boolean
        Try
            Dim connection As New SqlConnection(Misc.GetDadoWebConfig("conexaoBiometria"))
            Dim adapter As New SqlDataAdapter

            adapter.SelectCommand.Connection = connection
            codigoUsuario = RemoveDigitoVerificadorParaInclusaoDaDigital(codigoUsuario)
            ds.Clear()

            adapter.SelectCommand.CommandText = "SELECT * FROM BiometriaDP WHERE (Codigo like @Codigo) and (Dedo = @Dedo)"
            adapter.SelectCommand.Parameters.Clear()
            adapter.SelectCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Codigo", System.Data.SqlDbType.VarChar, 50, "Codigo"))
            adapter.SelectCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Dedo", System.Data.SqlDbType.Int, 4, "Dedo"))

            If Misc.GetDadoWebConfig("layout") <> "stacasasauderiopreto" Then
                adapter.SelectCommand.Parameters("@Codigo").Value = codigoUsuario.Substring(0, 15) & "%"

            Else
                adapter.SelectCommand.Parameters("@Codigo").Value = codigoUsuario & "%"
            End If

            adapter.SelectCommand.Parameters("@Dedo").Value = NumeroDoDedo
            adapter.Fill(ds)
            Return True

        Catch
            Return False
        End Try

    End Function

    Function VerificaSeUsuarioExiste(ByVal codigoUsuario As String) As Boolean
        Dim DS As New DSIdentify
        codigoUsuario = RemoveDigitoVerificadorParaInclusaoDaDigital(codigoUsuario)

        If (CarregaDatasetDoUsuario(codigoUsuario, DS)) Then
            If (DS.BiometriaDP.Rows.Count) > 0 Then Return True Else Return False

        Else
            Return False
        End If

    End Function

    Function LeDadosDoUsuarioNoBancoDeDados(ByVal codigoUsuario As String, ByRef DataSetIdentify As DSIdentify) As Boolean
        codigoUsuario = RemoveDigitoVerificadorParaInclusaoDaDigital(codigoUsuario)
        Return (CarregaDatasetDoUsuario(codigoUsuario, DataSetIdentify))
    End Function

    Function GravaTeplatesNoBanco(ByVal CodigoUsuario As String, ByVal TemplateBin1 As TTemplate, ByVal TemplateBin2 As TTemplate, ByVal TemplateBin3 As TTemplate, ByVal codPrestador As String) As Boolean
        Dim stringErroDebug As String
        CodigoUsuario = RemoveDigitoVerificadorParaInclusaoDaDigital(CodigoUsuario)

        Try
            Dim ComandoSQL As String
            ComandoSQL = "UPDATE BiometriaDP set Template1 = @Template1, Template2 = @Template2, Template3 = @Template3 WHERE (Codigo = @Codigo) and (Dedo = @Dedo)"
            Dim myConnection As SqlConnection = New SqlConnection(Misc.GetDadoWebConfig("conexaoBiometria"))
            Dim myCommand As SqlCommand = New SqlCommand(ComandoSQL, myConnection)

            myCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Template1", System.Data.SqlDbType.VarBinary, 2147483647, "Template1"))
            myCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Template2", System.Data.SqlDbType.VarBinary, 2147483647, "Template2"))
            myCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Template3", System.Data.SqlDbType.VarBinary, 2147483647, "Template3"))
            myCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Codigo", System.Data.SqlDbType.VarChar, 50, "Codigo"))
            myCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Dedo", System.Data.SqlDbType.Int, 4, "Dedo"))

            myCommand.Parameters("@Template1").Value = TemplateBin1.tpt
            myCommand.Parameters("@Template2").Value = TemplateBin2.tpt
            myCommand.Parameters("@Template3").Value = TemplateBin3.tpt
            myCommand.Parameters("@Codigo").Value = CodigoUsuario
            myCommand.Parameters("@Dedo").Value = NumeroDoDedo

            Dim QtdeItens As Integer
            myCommand.Connection.Open()
            QtdeItens = myCommand.ExecuteNonQuery()
            myCommand.Connection.Close()
            stringErroDebug = "Atualizar templates: " & QtdeItens

            If QtdeItens = 0 Then Return False
            Return True

        Catch ex As Exception
            stringErroDebug = "Erro em atualizar templates: " & ex.Message
            insertlog(CodigoUsuario, DateTime.Now, "ERRO ao gravar templates no Bando de Dados: " & ex.Message, False, NumeroDoDedo, codPrestador)
            Return False
        End Try

    End Function

    Public Function GravaNovosTemplatesDoUsuarioNoBanco(ByVal CodigoUsuario As String, ByVal TemplateBin1 As TTemplate, ByVal TemplateBin2 As TTemplate, ByVal TemplateBin3 As TTemplate, ByVal dataSetIdentify As DSIdentify, ByVal codPrestador As String) As Boolean
        Dim connection As New SqlConnection(Misc.GetDadoWebConfig("conexaoBiometria"))
        Dim command As New SqlCommand
        Dim stringErroDebug As String

        command.Connection = connection
        CodigoUsuario = RemoveDigitoVerificadorParaInclusaoDaDigital(CodigoUsuario)

        Try
            command.CommandText = "INSERT INTO BiometriaDP(Codigo, Dedo, DataBin, Template1, Template2, Template3) VALUES (@Codigo, @Dedo, @DataBin, @Template1, @Template2, @Template3)"

            command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Codigo", System.Data.SqlDbType.VarChar, 50, "Codigo"))
            command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Dedo", System.Data.SqlDbType.Int, 4, "Dedo"))
            command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@DataBin", System.Data.SqlDbType.NVarChar, 1073741823, "DataBin"))
            command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Template1", System.Data.SqlDbType.VarBinary, 2147483647, "Template1"))
            command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Template2", System.Data.SqlDbType.VarBinary, 2147483647, "Template2"))
            command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Template3", System.Data.SqlDbType.VarBinary, 2147483647, "Template3"))

            command.Parameters("@Codigo").Value = CodigoUsuario
            command.Parameters("@Dedo").Value = NumeroDoDedo
            command.Parameters("@DataBin").Value = dataSetIdentify.GetXml
            command.Parameters("@Template1").Value = TemplateBin1.tpt
            command.Parameters("@Template2").Value = TemplateBin2.tpt
            command.Parameters("@Template3").Value = TemplateBin3.tpt

            Dim QtdeItens As Integer
            command.Connection.Open()
            QtdeItens = command.ExecuteNonQuery()
            command.Connection.Close()

            If QtdeItens > 0 Then
                insertlog(CodigoUsuario, DateTime.Now, "Biometria Cadastrada com Sucesso", True, NumeroDoDedo, codPrestador)
                Return True

            Else
                Return False
            End If

        Catch ex As Exception
            stringErroDebug = "ERRO ao gravar novo usuário no Banco de Dados: " & ex.Message
            insertlog(CodigoUsuario, DateTime.Now, "ERRO ao gravar novo usuário no Bando de Dados: " & ex.Message, False, NumeroDoDedo, codPrestador)
            Return False
        End Try

    End Function

    Function getIdade(ByVal DataNascimento As Date) As Integer
        Dim AnoAtual, MesAtual, DiaAtual, Anos, Meses, Dias As Integer

        AnoAtual = Date.Now.Year
        MesAtual = Date.Now.Month
        DiaAtual = Date.Now.Day

        Anos = 0
        Meses = 0
        Dias = 0

        If AnoAtual >= DataNascimento.Year Then
            Anos = AnoAtual - DataNascimento.Year
            If MesAtual < DataNascimento.Month & Anos > 0 Then
                Anos = Anos - 1
                MesAtual = MesAtual + 12
            End If
        End If

        If MesAtual > DataNascimento.Month Then Meses = MesAtual - DataNascimento.Month

        If DiaAtual < DataNascimento.Day & Meses > 0 Then
            Meses = Meses - 1
            DiaAtual = DiaAtual + 30
        End If

        If DiaAtual > DataNascimento.Day Then Dias = DiaAtual - DataNascimento.Day

        If Anos >= 1 Then
            Return Anos
        Else
            Return 0
        End If

    End Function

#End Region

End Class