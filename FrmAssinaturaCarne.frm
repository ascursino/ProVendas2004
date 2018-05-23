VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmAssinaturaCarne 
   Caption         =   "Personaliza��o de carn�"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6480
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmAssinaturaCarne.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3585
   ScaleWidth      =   6480
   Begin VB.Frame FraBotaoCli 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   6255
      Begin VB.CommandButton CmdFechar 
         Caption         =   "&Fechar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4680
         TabIndex        =   7
         ToolTipText     =   "Fechar"
         Top             =   240
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   120
         OleObjectBlob   =   "FrmAssinaturaCarne.frx":0CCA
         Top             =   120
      End
      Begin VB.CommandButton CmdBranco 
         Caption         =   "&Em branco"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   6
         ToolTipText     =   "Personalizar carn� em branco"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton CmdOK 
         Caption         =   "&Personalizar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         ToolTipText     =   "Personalizar carn�"
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   6255
      Begin VB.CommandButton CmdListaArq 
         Caption         =   "&Listar arquivos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   8
         ToolTipText     =   "Listar arquivos de sua m�quina"
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox TxtLogo 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         ToolTipText     =   "Logomarca da empresa"
         Top             =   600
         Width           =   4935
      End
      Begin VB.TextBox TxtWeb 
         Height          =   285
         Left            =   1200
         MaxLength       =   60
         TabIndex        =   4
         ToolTipText     =   "Endere�o do site ou email da empresa"
         Top             =   2040
         Width           =   4935
      End
      Begin VB.TextBox TxtTel 
         Height          =   285
         Left            =   1200
         MaxLength       =   40
         TabIndex        =   3
         ToolTipText     =   "Telefone da empresa"
         Top             =   1680
         Width           =   4935
      End
      Begin VB.TextBox TxtBairro 
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   2
         ToolTipText     =   "Bairro da empresa"
         Top             =   1320
         Width           =   4935
      End
      Begin VB.TextBox TxtEndereco 
         Height          =   285
         Left            =   1200
         MaxLength       =   60
         TabIndex        =   1
         ToolTipText     =   "Endere�o da empresa"
         Top             =   960
         Width           =   4935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmAssinaturaCarne.frx":0EFE
         TabIndex        =   11
         Top             =   600
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmAssinaturaCarne.frx":0F6A
         TabIndex        =   12
         Top             =   960
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmAssinaturaCarne.frx":0FD4
         TabIndex        =   13
         Top             =   1320
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmAssinaturaCarne.frx":103A
         TabIndex        =   14
         Top             =   1680
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmAssinaturaCarne.frx":10A4
         TabIndex        =   15
         Top             =   2040
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmAssinaturaCarne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrBox As String
Public VPStrResponse As String

Private Sub CmdBranco_Click()
    Unload Me
    VGStrAssinaturaCarne = "branco"
    MDIPrincipal.Enabled = True
    Call ImprimirCarne
End Sub

Private Sub CmdFechar_Click()
    Unload Me
    MDIPrincipal.Enabled = True
End Sub

Private Sub CmdListaArq_Click()
    FrmListaArquivos.Show
End Sub

Private Sub CmdOK_Click()
    If TxtLogo.Text = "" And TxtEndereco.Text = "" And TxtBairro.Text = "" And TxtTel.Text = "" And TxtWeb.Text = "" Then
        VPStrBox = MsgBox("Preencha os campos para personalizar o carn�." & Chr(13) & "Caso n�o deseje personalizar escolha o bot�o 'Em Branco'", vbInformation, "Pr� Vendas 2004 - Informa��o")
    Else
        Conecta
        
        Dim RecAss As New ADODB.Recordset
        
        StrSql = "Select * From tb_AssinaturaCarne"
        RecAss.Open StrSql, vgCon, 1, 3
        
        If RecAss.EOF Then
            RecAss.AddNew
            RecAss("Logo") = TxtLogo.Text
            RecAss("Endereco") = TxtEndereco.Text
            RecAss("Bairro") = TxtBairro.Text
            RecAss("Telefone") = TxtTel.Text
            RecAss("Web") = TxtWeb.Text
            RecAss.Update
        Else
            RecAss("Logo") = TxtLogo.Text
            RecAss("Endereco") = TxtEndereco.Text
            RecAss("Bairro") = TxtBairro.Text
            RecAss("Telefone") = TxtTel.Text
            RecAss("Web") = TxtWeb.Text
            RecAss.Update
        End If
        
        Desconecta
        
        MDIPrincipal.Enabled = True
        
        Unload Me
        
        VGStrAssinaturaCarne = "personalizada"
        
        Call ImprimirCarne
    End If
End Sub

Private Sub Form_Resize()
  FrmAssinaturaCarne.Left = (MDIPrincipal.Width / 2) - (FrmAssinaturaCarne.Width / 2)
  FrmAssinaturaCarne.Top = (MDIPrincipal.Height / 3) - (FrmAssinaturaCarne.Height / 3)
End Sub

Private Sub Form_Load()
    Height = 4095
    Width = 6600
    
    MDIPrincipal.Enabled = False
    
    Skin1.LoadSkin (App.Path & "\ProVendas2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    Conecta
    
    Dim RecAss As New ADODB.Recordset
    
    StrSql = "Select * From tb_AssinaturaCarne"
    RecAss.Open StrSql, vgCon, 1, 3
    
    If Not RecAss.EOF Then
        If IsNull(RecAss!Logo) = False Then
            TxtLogo.Text = RecAss!Logo
        End If
        
        If IsNull(RecAss!endereco) = False Then
            TxtEndereco.Text = RecAss!endereco
        End If
        
        If IsNull(RecAss!bairro) = False Then
            TxtBairro.Text = RecAss!bairro
        End If
        
        If IsNull(RecAss!telefone) = False Then
            TxtTel.Text = RecAss!telefone
        End If
        
        If IsNull(RecAss!web) = False Then
            TxtWeb.Text = RecAss!web
        End If
    End If
    
    Desconecta
End Sub

Private Sub TxtBairro_GotFocus()
    TxtBairro.SelStart = 0
    TxtBairro.SelLength = Len(TxtBairro.Text)
End Sub

Private Sub TxtEndereco_GotFocus()
    TxtEndereco.SelStart = 0
    TxtEndereco.SelLength = Len(TxtEndereco.Text)
End Sub

Private Sub TxtLogo_GotFocus()
    TxtLogo.SelStart = 0
    TxtLogo.SelLength = Len(TxtLogo.Text)
End Sub

Private Sub TxtTel_GotFocus()
    TxtTel.SelStart = 0
    TxtTel.SelLength = Len(TxtTel.Text)
End Sub

Private Sub TxtTel_KeyPress(KeyAscii As Integer)
    '=== S� aceita n�meros, par�nteses, espa�o e - ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 40 And KeyAscii <> 41 And KeyAscii <> 32 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub

Sub ImprimirCarne()
    Dim RecPesq As New ADODB.Recordset
    Dim codparc As String
    Dim datacred As String
    Dim valortotal As String
    Dim parcela As String
    Dim vencimento As String
    Dim valor As String
    
    Conecta
    
    StrSql = "Select CR.DtCred,CR.Parcela,CR.ValorTotal,P.CodParc,P.NumParc,P.Vencimento,P.Valor,C.Nome " & _
             "From tb_Crediario as CR,tb_Crediario_Parcela as P,tb_Cliente as C,tb_Venda as V " & _
             "Where P.CodCred=CR.CodCred and C.CodCli=CR.CodCli and CR.CodCred=V.CodCred and V.CodVenda=" & VGIntCodVenda & " order by P.NumParc"
    RecPesq.Open StrSql, vgCon, 1, 3
    
    Do While Not RecPesq.EOF
        codparc = FormataNum(RecPesq.Fields.Item(3).Value)
        datacred = FormataData(RecPesq.Fields.Item(0).Value)
        valortotal = FormataMoeda(VerificaNulo(RecPesq.Fields.Item(2).Value))
        parcela = FormataNum(RecPesq.Fields.Item(4).Value) & "/" & FormataNum(RecPesq.Fields.Item(1).Value)
        vencimento = FormataData(RecPesq.Fields.Item(5).Value)
        valor = FormataMoeda(VerificaNulo(RecPesq.Fields.Item(6).Value))
        VGStrClienteRel = VerificaNulo(RecPesq.Fields.Item(7).Value)
        
        vgCon.Execute "INSERT INTO tb_Auxiliar " & _
        "(campo01,campo02,campo03,campo04,campo05,campo06) " & _
        "VALUES ('" & codparc & "','" & datacred & "','" & valortotal & "','" & parcela & "','" & vencimento & "','" & valor & " ')"
         
        RecPesq.MoveNext
    Loop
    
    Desconecta
    
    rptCarne.Show
End Sub

Private Sub TxtWeb_GotFocus()
    TxtWeb.SelStart = 0
    TxtWeb.SelLength = Len(TxtWeb.Text)
End Sub
