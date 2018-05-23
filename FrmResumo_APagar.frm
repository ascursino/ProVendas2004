VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmResumo_APagar 
   Caption         =   "Resumo de Contas a Pagar"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
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
   Icon            =   "FrmResumo_APagar.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   6975
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
      Height          =   2895
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   6735
      Begin VB.TextBox TxtDescr 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   120
         TabIndex        =   10
         Text            =   "?????"
         Top             =   480
         Width           =   6495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_APagar.frx":0CCA
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblVenc 
         Height          =   255
         Left            =   1560
         OleObjectBlob   =   "FrmResumo_APagar.frx":0D36
         TabIndex        =   4
         Top             =   1800
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_APagar.frx":0DA2
         TabIndex        =   5
         Top             =   1800
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblTipo 
         Height          =   255
         Left            =   1560
         OleObjectBlob   =   "FrmResumo_APagar.frx":0E10
         TabIndex        =   6
         Top             =   1440
         Width           =   5055
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_APagar.frx":0E7C
         TabIndex        =   7
         Top             =   1440
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblValor 
         Height          =   255
         Left            =   1560
         OleObjectBlob   =   "FrmResumo_APagar.frx":0EF0
         TabIndex        =   8
         Top             =   2160
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_APagar.frx":0F5C
         TabIndex        =   9
         Top             =   2160
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblStatus 
         Height          =   255
         Left            =   1560
         OleObjectBlob   =   "FrmResumo_APagar.frx":0FC0
         TabIndex        =   11
         Top             =   2520
         Width           =   5055
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_APagar.frx":102C
         TabIndex        =   12
         Top             =   2520
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblPagto 
         Height          =   255
         Left            =   4920
         OleObjectBlob   =   "FrmResumo_APagar.frx":1092
         TabIndex        =   13
         Top             =   1800
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   3480
         OleObjectBlob   =   "FrmResumo_APagar.frx":10FE
         TabIndex        =   14
         Top             =   1800
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblValorPagto 
         Height          =   255
         Left            =   4920
         OleObjectBlob   =   "FrmResumo_APagar.frx":116A
         TabIndex        =   15
         Top             =   2160
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   3480
         OleObjectBlob   =   "FrmResumo_APagar.frx":11D6
         TabIndex        =   16
         Top             =   2160
         Width           =   1095
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
      TabIndex        =   1
      Top             =   3000
      Width           =   6735
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
         Left            =   5400
         TabIndex        =   0
         ToolTipText     =   "Fechar"
         Top             =   240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   360
         OleObjectBlob   =   "FrmResumo_APagar.frx":1244
         Top             =   120
      End
   End
End
Attribute VB_Name = "FrmResumo_APagar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdFechar_Click()
    Unload Me
    
    MDIPrincipal.Enabled = True
    MDIPrincipal.WindowState = 2
    MDIPrincipal.SetFocus
End Sub

Private Sub Form_Resize()
  FrmResumo_APagar.Left = (MDIPrincipal.Width / 2) - (FrmResumo_APagar.Width / 2)
  FrmResumo_APagar.Top = (MDIPrincipal.Height / 3) - (FrmResumo_APagar.Height / 3)
End Sub

Private Sub Form_Load()
    Skin1.LoadSkin (App.Path & "\ProVendas2004.skn")
    Skin1.ApplySkin (Me.hwnd)
        
    Height = 4335
    Width = 7095
    
    Call MontaResumo
    
    MDIPrincipal.Enabled = False
End Sub

Sub MontaResumo()
        
    FrmPrincipal.GridAPagar.Row = FrmPrincipal.GridAPagar.ActiveRow
    
    'Descrição
    FrmPrincipal.GridAPagar.Col = 1
    TxtDescr.Text = FrmPrincipal.GridAPagar.Text
    
    'Tipo
    FrmPrincipal.GridAPagar.Col = 2
    LblTipo.Caption = FrmPrincipal.GridAPagar.Text
    
    'Vencimento
    FrmPrincipal.GridAPagar.Col = 3
    LblVenc.Caption = FrmPrincipal.GridAPagar.Text
    
    'Valor
    FrmPrincipal.GridAPagar.Col = 4
    LblValor.Caption = FrmPrincipal.GridAPagar.Text

    'Status
    FrmPrincipal.GridAPagar.Col = 5
    If FrmPrincipal.GridAPagar.Text = "Em aberto" Then
        LblStatus.Caption = FrmPrincipal.GridAPagar.Text
        LblValorPagto.Caption = ""
        LblPagto.Caption = ""
    Else
        Conecta
        
        Dim RecPag As New ADODB.Recordset
        
        FrmPrincipal.GridAPagar.Col = 6
        
        StrSql = "SELECT DtPagto,ValorPago,TipoPagto,NumBanco,NumCheque " & _
                 "FROM tb_ContaPagar_Pagto WHERE CodCPag=" & Val(FrmPrincipal.GridAPagar.Text)
        RecPag.Open StrSql, vgCon, 1, 3
        
        LblValorPagto.Caption = FormataMoeda(RecPag!ValorPago)
        LblPagto.Caption = FormataData(RecPag!DtPagto)
        
        If RecPag!TipoPagto = "Cheque" Then
            LblStatus.Caption = "Pago - Cheque (banco:" & RecPag!Numbanco & "  cheque:" & RecPag!Numcheque & ")"
        Else
            LblStatus.Caption = "Pago - Dinheiro"
        End If
        
        Desconecta
    End If
End Sub
