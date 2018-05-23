VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmCaixa_AReceber_Alt 
   Caption         =   "Alteração de Contas A Receber"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3945
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
   Icon            =   "FrmCaixa_AReceber_Alt.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   3945
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   3735
      Begin VB.TextBox TxtBanco 
         Height          =   285
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   4
         ToolTipText     =   "Número do banco do cheque"
         Top             =   3840
         Width           =   495
      End
      Begin VB.TextBox TxtCheque 
         Height          =   285
         Left            =   1200
         MaxLength       =   17
         TabIndex        =   5
         ToolTipText     =   "Número do cheque"
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox TxtDigito 
         Height          =   285
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   6
         ToolTipText     =   "Dígito do número do cheque"
         Top             =   4200
         Width           =   375
      End
      Begin VB.ComboBox CboTipo 
         Height          =   315
         ItemData        =   "FrmCaixa_AReceber_Alt.frx":0CCA
         Left            =   1320
         List            =   "FrmCaixa_AReceber_Alt.frx":0CDA
         TabIndex        =   0
         ToolTipText     =   "Tipo de conta a receber"
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox TxtDescr 
         Height          =   1365
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   3
         ToolTipText     =   "Descrição da conta a receber"
         Top             =   1920
         Width           =   3495
      End
      Begin VB.TextBox TxtDtVenc 
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   1
         Text            =   "__/__/____"
         ToolTipText     =   "Data do vencimento da conta a receber"
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox TxtValor 
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   2
         ToolTipText     =   "Valor da conta a receber"
         Top             =   1200
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCaixa_AReceber_Alt.frx":0D25
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCaixa_AReceber_Alt.frx":0D93
         TabIndex        =   12
         Top             =   720
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCaixa_AReceber_Alt.frx":0E01
         TabIndex        =   13
         Top             =   1200
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCaixa_AReceber_Alt.frx":0E65
         TabIndex        =   14
         Top             =   1680
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel42 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCaixa_AReceber_Alt.frx":0ED1
         TabIndex        =   15
         Top             =   3840
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel43 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCaixa_AReceber_Alt.frx":0F35
         TabIndex        =   16
         Top             =   4200
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCaixa_AReceber_Alt.frx":0FA1
         TabIndex        =   17
         Top             =   3480
         Width           =   3495
      End
   End
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
      TabIndex        =   9
      Top             =   4920
      Width           =   3735
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   360
         OleObjectBlob   =   "FrmCaixa_AReceber_Alt.frx":103B
         Top             =   240
      End
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
         Left            =   2400
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CmdOk 
         Caption         =   "&Ok"
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
         Left            =   1080
         TabIndex        =   7
         ToolTipText     =   "Efetuar a alteração"
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmCaixa_AReceber_Alt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrBox As String
Public VPStrResponse As String

Private Sub CmdFechar_Click()
    Unload Me
    
    FrmPrincipal.Enabled = True
End Sub

Private Sub CmdOK_Click()
    
    If CboTipo.Text = "" Or TxtDtVenc.Text = "" Or TxtValor.Text = "" Or TxtDescr.Text = "" Then
        VPStrBox = MsgBox("Preencha todos os campos.", vbInformation, "Pró Vendas 2004 - Informação")
    Else
        Conecta
        
        Dim RecReceb As New ADODB.Recordset
        
        StrSql = "Select * From tb_ContaReceber where CodCReceb=" & VGIntCodReceber
        RecReceb.Open StrSql, vgCon, 1, 3
        
        RecReceb("Tipo") = CboTipo.Text
        RecReceb("Vencimento") = FormataDataUS(TxtDtVenc.Text)
        RecReceb("Valor") = Mid(TxtValor.Text, 4)
        RecReceb("Descricao") = Trim(TxtDescr.Text)
        If TxtBanco.Text = "" Then
            RecReceb("NumBanco") = "0"
        Else
            RecReceb("NumBanco") = TxtBanco.Text
        End If
        RecReceb("NumCheque") = TxtCheque.Text & "-" & TxtDigito.Text
        RecReceb.Update
        
        Desconecta
        
        FrmPrincipal.CmdPesqAReceber.Value = True
        
        VPStrBox = MsgBox("Alteração efetuada.", vbInformation, "Pró Vendas 2004 - Informação")
            
        Unload Me
        
        FrmPrincipal.Enabled = True
    End If
    
End Sub

Private Sub Form_Resize()
  FrmCaixa_AReceber_Alt.Left = (MDIPrincipal.Width / 2) - (FrmCaixa_AReceber_Alt.Width / 2)
  FrmCaixa_AReceber_Alt.Top = (MDIPrincipal.Height / 3) - (FrmCaixa_AReceber_Alt.Height / 3)
End Sub

Private Sub Form_Load()
    Height = 6270
    Width = 4065
    
    Skin1.LoadSkin (App.Path & "\ProVendas2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    FrmPrincipal.Enabled = False
    
    Call MontaCbos
    
    Conecta
    
    Dim RecReceb As New ADODB.Recordset
    
    StrSql = "Select * From tb_ContaReceber where CodCReceb=" & VGIntCodReceber
    RecReceb.Open StrSql, vgCon, 1, 3
    
    CboTipo.Text = VerificaNulo(RecReceb!tipo)
    TxtDtVenc.Text = FormataData(RecReceb!vencimento)
    TxtValor.Text = FormataMoeda(VerificaNulo(RecReceb!valor))
    TxtDescr.Text = VerificaNulo(RecReceb!Descricao)
    If RecReceb!Numbanco = "0" Or IsNull(RecReceb!Numbanco) = True Then
        TxtBanco.Text = ""
    Else
        TxtBanco.Text = RecReceb!Numbanco
    End If
    
    If RecReceb!Numcheque <> "" Then
        If InStr(RecReceb!Numcheque, "-") <> 0 Then
            TxtCheque.Text = Mid(RecReceb!Numcheque, 1, InStr(RecReceb!Numcheque, "-") - 1)
            TxtDigito.Text = Mid(RecReceb!Numcheque, InStr(RecReceb!Numcheque, "-") + 1)
        Else
            TxtCheque.Text = RecReceb!Numcheque
            TxtDigito.Text = ""
        End If
    Else
        TxtCheque.Text = ""
        TxtDigito.Text = ""
    End If
    Desconecta
    
End Sub

Sub MontaCbos()
    Conecta
    
    Dim RecTipo As New ADODB.Recordset
    
    StrSql = "Select distinct Tipo From tb_ContaReceber"
    RecTipo.Open StrSql, vgCon, 1, 3
    
    CboTipo.Clear
    CboTipo.AddItem ("")
    Do While Not RecTipo.EOF
        CboTipo.AddItem (RecTipo!tipo)
        RecTipo.MoveNext
    Loop
    
    Desconecta
End Sub

Private Sub TxtBanco_GotFocus()
    TxtBanco.SelStart = 0
    TxtBanco.SelLength = Len(TxtBanco.Text)
End Sub

Private Sub TxtCheque_GotFocus()
    TxtCheque.SelStart = 0
    TxtCheque.SelLength = Len(TxtCheque.Text)
End Sub

Private Sub TxtDescr_GotFocus()
    TxtDescr.SelStart = 0
    TxtDescr.SelLength = Len(TxtDescr.Text)
End Sub

Private Sub TxtDigito_GotFocus()
    TxtDigito.SelStart = 0
    TxtDigito.SelLength = Len(TxtDigito.Text)
End Sub

Private Sub TxtDtVenc_GotFocus()
    If TxtDtVenc.Text = "__/__/____" Then
        TxtDtVenc.Text = ""
    End If
    TxtDtVenc.SelStart = 0
    TxtDtVenc.SelLength = Len(TxtDtVenc.Text)
End Sub

Private Sub TxtDtVenc_KeyPress(KeyAscii As Integer)
    '=== Só aceita números e barra ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtVenc_LostFocus()
    Dim VLStrData As String
    
    If TxtDtVenc.Text <> "" Then
        VLStrData = VerificaData(TxtDtVenc.Text)
        
        If VGStrDataErro = "sim" Then
            TxtDtVenc.SetFocus
        Else
            TxtDtVenc.Text = VLStrData
        End If
        
        VGStrDataErro = ""
    Else
        TxtDtVenc.Text = "__/__/____"
    End If
End Sub

Private Sub TxtValor_GotFocus()
    TxtValor.SelStart = 0
    TxtValor.SelLength = Len(TxtValor.Text)
End Sub

Private Sub TxtValor_LostFocus()
    TxtValor.Text = FormataMoeda(TxtValor.Text)
End Sub
