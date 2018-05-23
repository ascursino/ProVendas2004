Attribute VB_Name = "ModuloConecta"
Public Const SW_SHOW As Long = 5
Declare Function ShellExecute Lib "shell32.dll" Alias _
        "ShellExecuteA" (ByVal hwnd As Long, ByVal _
        lpOperation As String, ByVal lpFile As String, _
        ByVal lpParameters As String, ByVal lpDirectory _
        As String, ByVal nShowCmd As Long) As Long

Global vgCon As New ADODB.Connection 'variável de conexão
Global StrSql As String 'variável primária da string SQL
Global StrSql2 As String 'variável secundária da string SQL

'==== variável de criptografia ========
'Global VGCrypto As New clsCrypt
'Global VGCrypto As New Encryption
'======================================

Global VGIntCodCli As Long
Global VGIntCodRec As Long
Global VGIntCodMed As Long
Global VGIntCodForn As Long
Global VGIntCodEst As Long
Global VGIntCodProd As Long
Global VGIntCodParc As Long
Global VGIntCodCred As Long
Global VGIntCodCredsta As Long
Global VGIntCodCredstaVenda As Long
Global VGIntTotalCred As Long
Global VGIntCodCx As Long
Global VGIntCodPagar As Long
Global VGIntCodReceber As Long
Global VGIntCodVend As Long
Global VGIntCodOrc As Long
Global VGIntCodVenda As Long
Global VGIntCodVendaRel As Long
Global VGIntCodCredTemp As Long
Global VGIntLinha As Long
Global VGIntPropCodCred As Long
Global VGIntIdentCod As Long

Global VGStrAcesso As String
Global VGStrLocker As String
Global VGStrBox As String
Global VGStrNomeCredsta As String
Global VGStrTVenda As String
Global VGStrTCred As String
Global VGStrTDeb As String
Global VGStrTMov As String
Global VGStrStatusPagto As String
Global VGStrStatusReceb As String
Global VGStrDescrProd As String
Global VGStrTipoProd As String
Global VGStrCredLista As String
Global VGStrClienteRel As String
Global VGStrProposta As String
Global VGStrAssinatura As String
Global VGStrAssinaturaCarne As String
Global VGStrAssinaturaOrc As String
Global VGStrAssinaturaProp As String
Global VGStrAssinaturaProposta As String
Global VGStrTransfDados As String
Global VGStrParcelaCheque As String
Global VGStrReceb As String
Global VGStrIdentNome As String

Global VGStrBanco As String
Global VGStrChequeDigito As String
Global VGStrData As String
Global VGStrValor As String

Global VGStrEstoqueIncExtra As String
Global VGStrNomeCli As String
Global VGStrForm As String
Global VGStrIncluirProd As String
Global VGStrPersonalizar As String

'variáveis de acesso
Global VGStrMnuVendaCons As String
Global VGStrMnuVendaInc As String
Global VGStrMnuVendaExc As String
Global VGStrMnuVendaImp As String
Global VGStrMnuCliCons As String
Global VGStrMnuCliInc As String
Global VGStrMnuCliAlt As String
Global VGStrMnuCliExc As String
Global VGStrMnuCliImp As String
Global VGStrMnuFornCons As String
Global VGStrMnuFornInc As String
Global VGStrMnuFornAlt As String
Global VGStrMnuFornExc As String
Global VGStrMnuFornImp As String
Global VGStrMnuProdCons As String
Global VGStrMnuProdInc As String
Global VGStrMnuProdAlt As String
Global VGStrMnuProdExc As String
Global VGStrMnuProdImp As String
Global VGStrMnuEstCons As String
Global VGStrMnuEstInc As String
Global VGStrMnuEstAlt As String
Global VGStrMnuEstExc As String
Global VGStrMnuEstImp As String
Global VGStrMnuEstAlerta
Global VGStrMnuCredCons
Global VGStrMnuCredExc
Global VGStrMnuCredImp
Global VGStrMnuParcAlt
Global VGStrMnuParcExc
Global VGStrMnuCredstaCons
Global VGStrMnuCredstaInc
Global VGStrMnuCredstaAlt
Global VGStrMnuCredstaExc
Global VGStrMnuCredstaImp
Global VGStrMnuCxCons As String
Global VGStrMnuCxInc As String
Global VGStrMnuCxAlt As String
Global VGStrMnuCxExc As String
Global VGStrMnuCxImp As String
Global VGStrMnuPagCons As String
Global VGStrMnuPagInc As String
Global VGStrMnuPagAlt As String
Global VGStrMnuPagExc As String
Global VGStrMnuPagImp As String
Global VGStrMnuPagBx As String
Global VGStrMnuRecCons As String
Global VGStrMnuRecInc As String
Global VGStrMnuRecAlt As String
Global VGStrMnuRecExc As String
Global VGStrMnuRecImp As String
Global VGStrMnuRecBx As String
Global VGStrMnuOrcCons As String
Global VGStrMnuOrcInc As String
Global VGStrMnuOrcAlt As String
Global VGStrMnuOrcExc As String
Global VGStrMnuOrcImp As String
Global VGStrMnuVendCons As String
Global VGStrMnuVendInc As String
Global VGStrMnuVendAlt As String
Global VGStrMnuVendExc As String
Global VGStrMnuVendImp As String
Global VGStrMnuExNiver As String
Global VGStrMnuExCob As String
Global VGStrMnuExMala As String
Global VGStrMnuPropImp As String
Global VGStrMnuCarneImp As String

Global VGStrResolucao As String

Public Function Decipher(ByVal from_text As String) As String

Const MIN_ASC = 32  ' Space.
Const MAX_ASC = 126 ' ~.
Const NUM_ASC = MAX_ASC - MIN_ASC + 1

Dim offset As Long
Dim str_len As Integer
Dim i As Integer
Dim ch As Integer

    ' Initialize the random number generator.
    offset = 123
    Rnd -1
    Randomize offset

    ' Encipher the string.
    str_len = Len(from_text)
    For i = 1 To str_len
        ch = Asc(Mid$(from_text, i, 1))
        If ch >= MIN_ASC And ch <= MAX_ASC Then
            ch = ch - MIN_ASC
            offset = Int((NUM_ASC + 1) * Rnd)
            ch = ((ch - offset) Mod NUM_ASC)
            If ch < 0 Then ch = ch + NUM_ASC
            ch = ch + MIN_ASC
            Decipher = Decipher & Chr$(ch)
        End If
    Next i

End Function

Sub Conecta()       'Conecta com o banco prootica.mdb
    Dim vlFso, vlArquivo
    Dim vgStr As String
    
    '--- Abre arquivo criptografado com a string de conexão
    ''Set vlFso = CreateObject("Scripting.FileSystemObject")
    ''Set vlArquivo = vlFso.OpenTextFile(App.Path & "\prootica2004.dll", ForReading, False)
    
    '--- Descriptografa string de conexão
    ''vgStr = Decipher(vlArquivo.ReadLine)
    
    '--- Configura o tempo de pesquisa
    ''vgCon.ConnectionTimeout = 130
    
    '--- MontaString de conexão
    ''vgStr = "DBQ=" & App.Path & "\prootica2004.mdb;" & _
    ''    "Driver={Microsoft Access Driver (*.mdb)};"
    
    vgStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\provendas2004.mdb;Persist Security Info=False"
    
    '--- Abre conexão com SQL
    vgCon.Open vgStr
End Sub

Sub Desconecta()
    vgCon.Close
End Sub
