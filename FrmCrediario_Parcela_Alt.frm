VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmCrediario_Parcela_Alt 
   Caption         =   "Alteração de Parcela"
   ClientHeight    =   2625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3135
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
   Icon            =   "FrmCrediario_Parcela_Alt.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   3135
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
      TabIndex        =   5
      Top             =   1800
      Width           =   2895
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   0
         OleObjectBlob   =   "FrmCrediario_Parcela_Alt.frx":0CCA
         Top             =   0
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
         Left            =   1560
         TabIndex        =   3
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
         Left            =   240
         TabIndex        =   2
         ToolTipText     =   "Efetuar alteração"
         Top             =   240
         Width           =   1095
      End
   End
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
      Height          =   1695
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   2895
      Begin VB.TextBox TxtValor 
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   0
         ToolTipText     =   "Valor da parcela"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox TxtVenc 
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   1
         Text            =   "__/__/____"
         ToolTipText     =   "Data de vencimento da parcela"
         Top             =   1200
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCrediario_Parcela_Alt.frx":0EFE
         TabIndex        =   6
         Top             =   720
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCrediario_Parcela_Alt.frx":0F62
         TabIndex        =   7
         ToolTipText     =   "Vencimento das parcelas"
         Top             =   1200
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCrediario_Parcela_Alt.frx":0FD0
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblParc 
         Height          =   255
         Left            =   960
         OleObjectBlob   =   "FrmCrediario_Parcela_Alt.frx":1038
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "FrmCrediario_Parcela_Alt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrBox As String
Public VPStrResponse As String

Private Sub CmdFechar_Click()
    FrmResumo_Crediario.Enabled = True
    Unload Me
End Sub

Private Sub CmdOK_Click()
    If TxtValor.Text = "" Or TxtVenc.Text = "" Then
        VPStrBox = MsgBox("Não pode conter campos em branco.", vbInformation, "Pró Vendas 2004 - Informação")
    Else
        Conecta
        
        Dim RecParc As New ADODB.Recordset
        
        StrSql = "SELECT * FROM tb_Crediario_Parcela where CodParc=" & VGIntCodParc
        RecParc.Open StrSql, vgCon, 1, 3
        
        RecParc("Vencimento") = FormataDataUS(TxtVenc.Text)
        RecParc("Valor") = CCur(TxtValor.Text)
        RecParc.Update
        
        Desconecta
        
        VPStrBox = MsgBox("Alteração efetuada.", vbInformation, "Pró Vendas 2004 - Informação")

        Call FrmResumo_Crediario.MontaResumo
        
        Unload Me
        
        FrmResumo_Crediario.Enabled = True
    End If
End Sub

Private Sub Form_Resize()
  FrmCrediario_Parcela_Alt.Left = (MDIPrincipal.Width / 2) - (FrmCrediario_Parcela_Alt.Width / 2)
  FrmCrediario_Parcela_Alt.Top = (MDIPrincipal.Height / 3) - (FrmCrediario_Parcela_Alt.Height / 3)
End Sub

Private Sub Form_Load()
    Height = 3135
    Width = 3255
    
    Skin1.LoadSkin (App.Path & "\ProVendas2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    Conecta
    
    Dim RecParc As New ADODB.Recordset
    
    StrSql = "SELECT Valor,Vencimento FROM tb_Crediario_Parcela where CodParc=" & VGIntCodParc
    RecParc.Open StrSql, vgCon, 1, 3
    
    TxtValor.Text = FormataMoeda(VerificaNulo(RecParc!valor))
    TxtVenc.Text = FormataData(RecParc!vencimento)
    
    FrmResumo_Crediario.GridParcela.Row = FrmResumo_Crediario.GridParcela.ActiveRow
    FrmResumo_Crediario.GridParcela.Col = 1
    LblParc.Caption = FrmResumo_Crediario.GridParcela.Text
    
    Desconecta
    
    FrmResumo_Crediario.Enabled = False
End Sub

Private Sub TxtValor_GotFocus()
    TxtValor.SelStart = 0
    TxtValor.SelLength = Len(TxtValor.Text)
End Sub

Private Sub TxtVenc_GotFocus()
    TxtVenc.SelStart = 0
    TxtVenc.SelLength = Len(TxtVenc.Text)
End Sub

Private Sub TxtVenc_LostFocus()
    Dim VLStrData As String
    
    If TxtVenc.Text <> "" Then
        VLStrData = VerificaData(TxtVenc.Text)
        
        If VGStrDataErro = "sim" Then
            TxtVenc.SetFocus
        Else
            TxtVenc.Text = VLStrData
        End If
        
        VGStrDataErro = ""
    End If
End Sub
