VERSION 5.00
Begin VB.Form formInicial 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Automatizador do SABI"
   ClientHeight    =   4560
   ClientLeft      =   150
   ClientTop       =   330
   ClientWidth     =   8055
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "formInicial.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   8055
   Begin VB.PictureBox pctCopiaPartedaTelaCPF 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5040
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   127
      TabIndex        =   18
      Top             =   3720
      Width           =   1935
   End
   Begin VB.PictureBox pctImpressora 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   3720
      Picture         =   "formInicial.frx":08CA
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   17
      Top             =   840
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pctFundo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   120
      ScaleHeight     =   3855
      ScaleWidth      =   8175
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      Begin VB.Frame fraOrdem 
         Caption         =   "Requerimentos ordenados por"
         ForeColor       =   &H00404040&
         Height          =   855
         Left            =   0
         TabIndex        =   26
         Top             =   1800
         Width           =   7815
         Begin VB.CommandButton cmdFechar 
            Cancel          =   -1  'True
            Caption         =   "Fechar"
            Height          =   370
            Left            =   6240
            TabIndex        =   30
            Top             =   360
            Width           =   1200
         End
         Begin VB.CommandButton cmdIniciar 
            Caption         =   "&Processar"
            Default         =   -1  'True
            Height          =   370
            Left            =   4800
            TabIndex        =   29
            Top             =   360
            Width           =   1200
         End
         Begin VB.OptionButton optOrdem 
            Caption         =   "&Nome do Periciando"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   28
            ToolTipText     =   "Apresenta os requerimentos ordenados por nome do periciando"
            Top             =   420
            Value           =   -1  'True
            Width           =   2295
         End
         Begin VB.OptionButton optOrdem 
            Caption         =   "&Hora da Per�cia"
            Height          =   255
            Index           =   1
            Left            =   2640
            TabIndex        =   27
            ToolTipText     =   "Apresenta os requerimentos ordenados pelo horario da pericia"
            Top             =   420
            Width           =   2175
         End
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   2400
         Top             =   720
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   3000
         Top             =   720
      End
      Begin VB.Frame fraImprime 
         Caption         =   "Imprimir 2� Via da Marca��o de Exame"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   1005
         Left            =   0
         TabIndex        =   15
         Top             =   2760
         Visible         =   0   'False
         Width           =   7815
         Begin VB.CommandButton cmdFechar2 
            Caption         =   "Fechar"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   6240
            TabIndex        =   31
            ToolTipText     =   " Fechar aplicativo "
            Top             =   360
            Width           =   1200
         End
         Begin VB.CheckBox chkPP 
            Caption         =   "PP"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2880
            TabIndex        =   24
            ToolTipText     =   " marcar para imprimir os exames de PP "
            Top             =   680
            Width           =   735
         End
         Begin VB.CheckBox chkIniciais 
            Caption         =   "Exame Inicial"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   360
            TabIndex        =   23
            ToolTipText     =   " marcar para imprimir os exames iniciais "
            Top             =   680
            Width           =   1695
         End
         Begin VB.TextBox txtUltimo 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   330
            Left            =   2880
            TabIndex        =   20
            Text            =   "1"
            ToolTipText     =   " fixar o final da sequ�ncia de impress�o "
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txttPrimeiro 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   330
            Left            =   1800
            TabIndex        =   19
            Text            =   "1"
            ToolTipText     =   " fixar o �nicio da sequ�ncia de impress�o "
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton cmdContinua 
            Caption         =   "Confirma"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   4800
            TabIndex        =   16
            ToolTipText     =   " Confirmar a sequ�ncia e os tipos de exames e inciar a opera��o de impress�o "
            Top             =   375
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.Image parabaixo 
            Height          =   240
            Left            =   4200
            Picture         =   "formInicial.frx":0BAC
            Stretch         =   -1  'True
            ToolTipText     =   " mover a lista de requerimentos para baixo "
            Top             =   600
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image paracima 
            Height          =   240
            Left            =   4200
            Picture         =   "formInicial.frx":0FEE
            Stretch         =   -1  'True
            ToolTipText     =   " mover a lista de requerimentos para cima "
            Top             =   240
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "a "
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2280
            TabIndex        =   22
            Top             =   300
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Sequ�ncia: De"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   360
            TabIndex        =   21
            Top             =   300
            Width           =   1335
         End
      End
      Begin VB.PictureBox pctProgressoFundo 
         Appearance      =   0  'Flat
         FillColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   7440
         ScaleHeight     =   225
         ScaleWidth      =   465
         TabIndex        =   13
         Top             =   0
         Visible         =   0   'False
         Width           =   495
         Begin VB.PictureBox pctProgresso 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            FillColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   600
            Left            =   -30
            ScaleHeight     =   570
            ScaleWidth      =   165
            TabIndex        =   14
            Top             =   -30
            Width           =   200
         End
      End
      Begin VB.PictureBox pctFundoCopias 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   3480
         ScaleHeight     =   735
         ScaleWidth      =   2535
         TabIndex        =   9
         Top             =   600
         Visible         =   0   'False
         Width           =   2535
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1455
            Left            =   480
            ScaleHeight     =   97
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   129
            TabIndex        =   10
            Top             =   0
            Width           =   1935
         End
      End
      Begin VB.PictureBox pctApresentaPartedaTelaCopiada 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   3960
         ScaleHeight     =   41
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   121
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.ListBox lstMostrarRequerimentos 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   225
         ItemData        =   "formInicial.frx":1430
         Left            =   3480
         List            =   "formInicial.frx":1432
         TabIndex        =   7
         Top             =   1440
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.PictureBox pctCopiaPartedaTela 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1560
         ScaleHeight     =   41
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   121
         TabIndex        =   5
         Top             =   1200
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.ListBox lstClassificar 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   0
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblLocaleData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "local e data"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   5880
         TabIndex        =   25
         Top             =   1440
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Image imageIcone 
         Height          =   480
         Left            =   80
         Picture         =   "formInicial.frx":1434
         Top             =   80
         Width           =   480
      End
      Begin VB.Label lblRelogio 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 00 "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   225
         Left            =   7560
         TabIndex        =   6
         Top             =   840
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label lblRequerimentodoSABI 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Automatizador do SABI"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   4
         Top             =   120
         Width           =   2925
      End
      Begin VB.Label lblversao 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Compilado em 03-04-2019"
         ForeColor       =   &H00808080&
         Height          =   225
         Left            =   0
         TabIndex        =   3
         Top             =   1440
         Width           =   2100
      End
      Begin VB.Label lblImprimir 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         Caption         =   " Imprimir  "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   240
         Left            =   7080
         TabIndex        =   2
         Top             =   1200
         Visible         =   0   'False
         Width           =   840
      End
   End
   Begin VB.PictureBox pctDigitos 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   3120
      Picture         =   "formInicial.frx":1CFE
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   11
      Top             =   4200
      Width           =   1035
   End
   Begin VB.PictureBox pctEsteRequerimento 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   5280
      Picture         =   "formInicial.frx":2BE0
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   178
      TabIndex        =   12
      Top             =   3240
      Visible         =   0   'False
      Width           =   2670
   End
End
Attribute VB_Name = "formInicial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MenuName As New Collection
Dim MenuHandle As New Collection
Dim lHwnd As Long
Dim imprimereq As Boolean
Dim modoImprime As String
Dim LocalY As Long
Dim LocalCopiar As Boolean
Dim requerimentomostrado As Long
Dim mtempo1 As Long
Dim mtempo2 As Long
Dim deslocalista As Long
    
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, wParam As Any, ByVal lParam As String) As Long

Public Sub configuranomedoarquivo()
  Dim hwndDialog As Long  ' handle to the dialog box
  Dim hwndButton As Long  ' handle to the  button
  Dim retval As Long      ' return value
  Dim SaveAsDialog As Long
  Dim cDUIViewWndCIassName As Long
  Dim cDirectUIHWND As Long
  Dim cFloatNotifySink As Long
  Dim comboBox32win As Long
  Dim ComboBoxwin As Long
  Dim txtlen As Long
  Dim txt As String
  Dim conta As Long
  Dim EditBox As Long
  Dim memocritica As String
  Dim GlobalBot�oSalvar As Long
  Dim titletmp As String

  SaveAsDialog = FindWindow("#32770", "Choose Export File")
  cDUIViewWndCIassName = FindWindowEx(SaveAsDialog, 0, "DUIViewWndClassName", vbNullString)
  cDirectUIHWND = FindWindowEx(cDUIViewWndCIassName, 0, "DirectUIHWND", vbNullString)
  cFloatNotifySink = FindWindowEx(cDirectUIHWND, 0, "FloatNotifySink", vbNullString)
  ComboBoxwin = FindWindowEx(cFloatNotifySink, 0, "ComboBox", vbNullString)
  EditBox = FindWindowEx(ComboBoxwin, 0, "Edit", vbNullString)
  Debug.Print EditBox
  retval = SendMessage(ComboBoxwin, WM_SETTEXT, vbNullString, GlobalAreadeTrabalho & "\Agendamentos.txt")
  txtlen = SendMessage(EditBox, WM_GETTEXTLENGTH, vbNullString, vbNullString)
  txtlen = txtlen + 1
  txt = Space$(txtlen)
  Call SendMessage(EditBox, WM_GETTEXT, ByVal 260, txt)
  If InStr(1, txt, GlobalAreadeTrabalho & "\Agendamentos.txt") = 0 Then
    memocritica = "N�o foi poss�vel inserir na tela Salvar o destino: " & GlobalAreadeTrabalho & "\Agendamentos.txt"
    Exit Sub
  Else
    'procura o botao Salvar
    GlobalBot�oSalvar = 0
    GlobalBot�oSalvar = FindWindowEx(SaveAsDialog, 0, "Button", "Sa&lvar")
    conta = 0
    While GlobalBot�oSalvar = 0
      Sleep 20
      conta = conta + 1
      If conta > 50 Then
        memocritica = "N�o foi poss�vel encontrar o bot�o Salvar em 1 segundo"
        Exit Sub
      End If
      GlobalBot�oSalvar = FindWindowEx(SaveAsDialog, 0, "Button", "Sa&lvar")
    Wend
       
    'comanda o salvamento
    PostMessage GlobalBot�oSalvar, BM_CLICK, 0, 0
    
    'espera tela salvar ser fechada
    conta = 0
    titletmp = Space(256)
    GetWindowText SaveAsDialog, titletmp, Len(titletmp)
    While InStr(1, UCase(titletmp), "SALVAR COMO") <> 0
      'reafirma comando de salvamento
      PostMessage GlobalBot�oSalvar, BM_CLICK, 0, 0
      Sleep 50
      conta = conta + 1
      If conta > 50 Then
        memocritica = "A tela Salvar Como n�o foi fechada em 1 segundo"
        Exit Sub
      End If
      titletmp = Space(256)
      GetWindowText SaveAsDialog, titletmp, Len(titletmp)
      DoEvents
    Wend
    If Err.Number > 0 Then
      memocritica = "Erro: " & Err.Description
    Else
      memocritica = ""
     End If
  End If
End Sub

Sub mostratela()
  Dim RtnValue
  Dim win As Long
  Dim desloca As Long
  Dim esquerda, altura, largura, dimensao As Long
  If Val(GlobalRequerimentos(GlobalIDRequerimento).N�mero) = 0 Or (requerimentomostrado = Val(GlobalRequerimentos(GlobalIDRequerimento).N�mero)) Then Exit Sub
  If Val(GlobalRequerimentos(GlobalIDRequerimento).sequencia) > 40 Then 'numero de linhas
    desloca = ((Val(GlobalRequerimentos(GlobalIDRequerimento).sequencia) - 40) * 240) / 15
  Else
    desloca = 0
  End If
  esquerda = 760
  altura = 0
  largura = Picture1.Width
  dimensao = Picture1.Height
  win = GlobalhMDIClient
  Picture1.Refresh
  pctEsteRequerimento.Refresh
  requerimentomostrado = Val(GlobalRequerimentos(GlobalIDRequerimento).N�mero)
  RtnValue = BitBlt(GetDC(win), CLng(esquerda), CLng(-desloca), CLng(largura), CLng(dimensao), Picture1.hDC, CLng(0), CLng(0), SRCCOPY)
End Sub
    
Public Function CapturaNumeroDetalhes() As String
  On Error Resume Next
  Dim TopRequerimento As Long
  Dim LeftRequerimento As Long
  Dim algarismo As Long
  Dim soma As Long
  Dim indice As Long
  Dim digito As Long
  Dim letra As String
  Dim Requerimento As String
  Dim Deslocamento As Long
  Dim Digitos As Long

  Requerimento = ""
  TopRequerimento = 2
  Deslocamento = 1
  Digitos = 11
  LeftRequerimento = Deslocamento
  soma = 0
  algarismo = 0
  For digito = 0 To Digitos - 1
    soma = 0
    If pctCopiaPartedaTelaCPF.Point(algarismo + LeftRequerimento, TopRequerimento + 2) = 0 Then
      soma = 1
    End If
    If pctCopiaPartedaTelaCPF.Point(algarismo + LeftRequerimento, TopRequerimento + 6) = 0 Then
      soma = soma + 2
    End If
    If pctCopiaPartedaTelaCPF.Point(algarismo + LeftRequerimento + 2, TopRequerimento) = 0 Then
      soma = soma + 4
    End If
    If pctCopiaPartedaTelaCPF.Point(algarismo + LeftRequerimento + 2, TopRequerimento + 4) = 0 Then
      soma = soma + 8
    End If
    If pctCopiaPartedaTelaCPF.Point(algarismo + LeftRequerimento + 2, TopRequerimento + 8) = 0 Then
      soma = soma + 16
    End If
    If pctCopiaPartedaTelaCPF.Point(algarismo + LeftRequerimento + 4, TopRequerimento + 2) = 0 Then
      soma = soma + 32
    End If
    If pctCopiaPartedaTelaCPF.Point(algarismo + LeftRequerimento + 4, TopRequerimento + 6) = 0 Then
      soma = soma + 64
    End If
    Select Case soma
      Case 28
        letra = "1"
      Case 52
        letra = "2"
      Case 124
        letra = "3"
      Case 66
        letra = "4"
      Case 85
        letra = "5"
      Case 95
        letra = "6"
      Case 12
        letra = "7"
      Case 127
        letra = "8"
      Case 125
        letra = "9"
      Case 119
        letra = "0"
      Case 90
        letra = "INICIAL"
      Case 44
        letra = "PP"
      Case 108
        letra = "PR"
      Case 75
        letra = "NORMAL"
      Case 23
        letra = "DEFERIDO"
      Case 3
        letra = "INDEFERIDO"
      Case 15
        letra = "PENDENTE"
      Case Else
        letra = ""
    End Select
    Requerimento = Requerimento & letra
    algarismo = algarismo + 6
  Next digito
  CapturaNumeroDetalhes = Requerimento
End Function
    
Sub atualizaprogresso()
  On Error Resume Next
  pctProgressoFundo.Width = 2000
  pctProgresso.Width = pctProgressoFundo.Width * GlobalIDRequerimento / GlobalQuantidadedeRequerimentos
  pctProgressoFundo.Top = lblRelogio.Top
  pctProgressoFundo.Left = lblRelogio.Left + lblRelogio.Width + 40
  pctProgressoFundo.Height = lblRelogio.Height
  pctProgressoFundo.Visible = True
  pctProgresso.Cls
  pctProgresso.PSet (0, 0)
  pctProgresso.Top = -20
  pctProgresso.Left = -20
  pctProgresso.Print Space(10) & Format(GlobalIDRequerimento, "00") & " de " & Format(GlobalQuantidadedeRequerimentos, "00")
  pctProgressoFundo.Cls
  pctProgressoFundo.PSet (0, 0)
  pctProgressoFundo.Print Space(10) & Format(GlobalIDRequerimento, "00") & " de " & Format(GlobalQuantidadedeRequerimentos, "00")
  lblversao.Top = -1000
  lblversao.Left = Me.Width - lblversao.Width - 360
End Sub

Sub Escreve(pontox As Long, pontoy As Long, NUMERO As String)
  Dim conta, linha, coluna, digito As Long
  For conta = 1 To Len(NUMERO)
    digito = Mid(NUMERO, conta, 1)
    If digito = 0 Then
      digito = 9
    Else
      digito = digito - 1
    End If
    Picture1.Visible = False
    For linha = 4 To 12
      For coluna = 0 To 6 '160
        If pctDigitos.Point(coluna + 6 * digito + 3, linha) = 2631720 Then
          Picture1.PSet (6 * (conta - 1) + pontox + coluna, pontoy + linha), RGB(0, 0, 0)
        End If
      Next coluna
    Next linha
    Picture1.Visible = True
  Next conta
End Sub

Sub Requerimenton�oEncontrado(Requerimento As String, sequencia As String)
  Dim conta, linha, coluna, digito, pontoy As Long
  
  pontoy = GlobalLinhaPicture
  Picture1.Visible = False
  pctFundoCopias.Top = 3000
  pctFundoCopias.Visible = True
  Picture1.Height = GlobalLinhaPicture * 15 + 300
  If Picture1.Height + 200 > pctFundoCopias.Height Then
    Picture1.Top = 3000
  End If
  Picture1.Width = pctCopiaPartedaTela.Width

  'colore linha
  For linha = 2 To 15
    For coluna = 1 To 726
      Picture1.PSet (coluna, pontoy + linha), RGB(255, 220, 220)
    Next coluna
  Next linha
    
  'escreve o numero do requerimento
  Escreve 2, pontoy, Requerimento
    
  'escreve Este requerimento n�o foi encontrado
  For linha = 4 To 15
    For coluna = 0 To 178
      If pctEsteRequerimento.Point(coluna, linha) = 0 Then
        Picture1.PSet (80 + coluna, pontoy + linha), RGB(40, 40, 40)
      End If
    Next coluna
  Next linha

  'tra�a linha preta
  For coluna = 1 To 726
    Picture1.PSet (coluna, pontoy + 16), RGB(40, 40, 40)
    Picture1.PSet (coluna, pontoy + 17), RGB(40, 40, 40)
  Next coluna
  For linha = 2 To 15
    Picture1.PSet (65, pontoy + linha), RGB(40, 40, 40)
    Picture1.PSet (283, pontoy + linha), RGB(40, 40, 40)
    Picture1.PSet (357, pontoy + linha), RGB(40, 40, 40)
    Picture1.PSet (383, pontoy + linha), RGB(40, 40, 40)
    Picture1.PSet (450, pontoy + linha), RGB(40, 40, 40)
    Picture1.PSet (524, pontoy + linha), RGB(40, 40, 40)
    Picture1.PSet (700, pontoy + linha), RGB(40, 40, 40)
    Picture1.PSet (722, pontoy + linha), RGB(40, 40, 40)
    Picture1.PSet (723, pontoy + linha), RGB(40, 40, 40)
    Picture1.PSet (724, pontoy + linha), RGB(40, 40, 40)
    Picture1.PSet (725, pontoy + linha), RGB(40, 40, 40)
    Picture1.PSet (726, pontoy + linha), RGB(40, 40, 40)
  Next linha
  Escreve 703, GlobalLinhaPicture, sequencia
  GlobalLinhaPicture = GlobalLinhaPicture + 16
  SavePicture Picture1.Image, GlobalPastadeTrabalho & "\" & GlobalDatadosRequerimentos & GlobalAgenciaEscolhida & "Todos.bmp"
  Picture1.Visible = True
End Sub

Sub efeitos(Imprime As Boolean, Ordem As String)
  Dim linha, coluna As Long
  Dim sequencia As String
  Dim pontoy As Long

  sequencia = Ordem
  Picture1.Visible = False
  pctFundoCopias.Top = 3000
  pctFundoCopias.Visible = True
  Picture1.Height = GlobalLinhaPicture * 15 + 300
  If Picture1.Height + 200 > pctFundoCopias.Height Then
    Picture1.Top = 3000
  End If
  Picture1.Width = pctCopiaPartedaTela.Width
  If Imprime Then
    For linha = 1 To pctCopiaPartedaTela.Height / 15
      For coluna = 1 To pctCopiaPartedaTela.Width / 15
        If pctCopiaPartedaTela.Point(coluna, linha) = 0 Then
          Picture1.PSet (coluna, GlobalLinhaPicture + linha + 1), RGB(220, 220, 220)
        Else
          Picture1.PSet (coluna, GlobalLinhaPicture + linha + 1), RGB(40, 40, 40)
        End If
      Next coluna
    Next linha
  Else
    For linha = 1 To pctCopiaPartedaTela.Height / 15 - 2
      For coluna = 1 To pctCopiaPartedaTela.Width / 15
        If pctCopiaPartedaTela.Point(coluna, linha) = 0 Then
          Picture1.PSet (coluna, GlobalLinhaPicture + linha + 1), RGB(255, 255, 225)
        Else
          Picture1.PSet (coluna, GlobalLinhaPicture + linha + 1), RGB(120, 120, 120)
        End If
      Next coluna
    Next linha
    For coluna = 1 To pctCopiaPartedaTela.Width / 15
      Picture1.PSet (coluna, GlobalLinhaPicture + linha + 1), RGB(40, 40, 40)
      Picture1.PSet (coluna, GlobalLinhaPicture + linha + 2), RGB(40, 40, 40)
    Next coluna
  End If
  pontoy = GlobalLinhaPicture
  Escreve 703, pontoy, sequencia
  GlobalLinhaPicture = GlobalLinhaPicture + 16
  SavePicture Picture1.Image, GlobalPastadeTrabalho & "\" & GlobalDatadosRequerimentos & GlobalAgenciaEscolhida & "Todos.bmp"
  Picture1.Visible = True
End Sub

Private Sub vermelho()
  On Error Resume Next
  Dim linha, coluna As Long
  
  Picture1.Visible = False
  For linha = Picture1.Height / 15 - 16 - 2 To Picture1.Height / 15 - 2
    For coluna = 1 To formInicial.Picture1.Width / 15
      If Picture1.Point(coluna, linha) = 14474460 Then
        Picture1.PSet (coluna, linha), RGB(255, 220, 220)
      End If
    Next coluna
  Next linha
  Picture1.Visible = True
End Sub
  
Private Function RequerimentosAgendaAnterior(nomearquivo As String) As Long
  Dim FileNumber  As Long
  Dim mTexto, mLinha As String
  Dim contador As Long
  Dim ultimoNIT As Long
  Dim pos1, pos2, pos3, pos4, pos5, pos6, pos7, pos8 As Long
    
  contador = 0
  FileNumber = FreeFile
  Open GlobalPastadeTrabalho & "\" & nomearquivo For Input As #FileNumber
  Do While Not EOF(FileNumber)
    Line Input #FileNumber, mLinha
    If Len(mLinha) = 0 Then Exit Do
  contador = contador + 1
  GlobalQuantidadedeRequerimentos = contador
  pos1 = InStr(1, mLinha, Chr(9))
    If pos1 > 0 Then
    GlobalRequerimentos(contador).sequencia = Mid(mLinha, 1, pos1 - 1)
    pos2 = InStr(pos1 + 1, mLinha, Chr(9))
    If pos2 > 0 Then
      GlobalRequerimentos(contador).N�mero = Mid(mLinha, pos1 + 1, pos2 - pos1 - 1)
    End If
    pos3 = InStr(pos2 + 1, mLinha, Chr(9))
    If pos3 > 0 Then
      GlobalRequerimentos(contador).Tipo = Mid(mLinha, pos2 + 1, pos3 - pos2 - 1)
      pos4 = InStr(pos3 + 1, mLinha, Chr(9))
      If pos4 > 0 Then
        GlobalRequerimentos(contador).Status = Mid(mLinha, pos3 + 1, pos4 - pos3 - 1)
        pos5 = InStr(pos4 + 1, mLinha, Chr(9))
        If pos5 > 0 Then
          GlobalRequerimentos(contador).NIT = Mid(mLinha, pos4 + 1, pos5 - pos4 - 1)
          If GlobalRequerimentos(contador).NIT <> "" Then ultimoNIT = contador
            pos6 = InStr(pos5 + 1, mLinha, Chr(9))
            If pos6 > 0 Then
              GlobalRequerimentos(contador).Impresso = Mid(mLinha, pos5 + 1, pos6 - pos5 - 1)
              pos7 = InStr(pos6 + 1, mLinha, Chr(9))
              If pos7 > 0 Then
                GlobalRequerimentos(contador).Segurado = Mid(mLinha, pos6 + 1, pos7 - pos6 - 1)
                GlobalRequerimentos(contador).Cr�tica = Mid(mLinha, pos7 + 1)
              End If
            End If
          End If
        End If
      End If
    End If
    mTexto = mTexto & mLinha & Chr(13) & Chr(10)
  Loop
  Close #FileNumber
  RequerimentosAgendaAnterior = ultimoNIT
End Function
   
Function AcertaLarguraRelatorio(referencia As String, palavra As String)
  Dim largura As Long
  
  AcertaLarguraRelatorio = palavra & " "
  largura = Len(referencia)
  If Len(palavra) < largura Then AcertaLarguraRelatorio = palavra & Space(largura - Len(palavra)) & " "
  If Len(palavra) > largura Then AcertaLarguraRelatorio = Mid(palavra, 1, largura) & " "
End Function

Sub ColocaTelaControleOperacionanoModoNormal()
  On Error Resume Next
  
  If GlobalIDControleOperacional Then
    If AppToForeground(, GlobalIDControleOperacional, SW_NORMAL) Then
        ' a tela foi maximizada
    Else
      MsgBox "Falha ao maximizar a tela Controle Operacional", vbCritical, "Maximizar Controle Operacional"
    End If
  Else
    MsgBox "N�o foi encontrada tela Controle Operaciona", vbCritical, "Maximizar Controle Operacional"
  End If
End Sub

Sub verificaeapaga(tituladatela As String)
  Dim IDTelaExterna As Long

  IDTelaExterna = 0
  IDTelaExterna = ObtemIDdaTelaPrincipalporTitulo(tituladatela)
  If IDTelaExterna <> 0 Then SendMessage IDTelaExterna, WM_CLOSE, 0, 0
End Sub

Sub ConverteData(datalonga As String)
  On Error Resume Next
  Dim memo As String
  Dim dia As Long
  Dim mes As Long
  Dim ano As Long
  
  GlobalDatadosRequerimentos = "00000000"
  memo = UCase(datalonga)
  If InStr(1, memo, ",") Then memo = Trim(Mid(memo, InStr(1, memo, ",") + 1))
  dia = Val(memo)
  ano = Val(Right$(memo, 4))
  If InStr(1, memo, "DE") Then memo = Trim(Mid(memo, InStr(1, memo, "DE") + 3))
  If InStr(1, memo, "DE") Then memo = Trim(Mid(memo, 1, InStr(1, memo, "DE") - 1))
  Select Case memo
    Case "JANEIRO"
      mes = 1
    Case "FEVEREIRO"
      mes = 2
    Case "MAR�O"
      mes = 3
    Case "ABRIL"
      mes = 4
    Case "MAIO"
      mes = 5
    Case "JUNHO"
      mes = 6
    Case "JULHO"
      mes = 7
    Case "AGOSTO"
      mes = 8
    Case "SETEMBRO"
      mes = 9
    Case "OUTUBRO"
      mes = 10
    Case "NOVEMBRO"
      mes = 11
    Case "DEZEMBRO"
      mes = 12
  End Select
  GlobalDatadosRequerimentos = Format(ano, "0000") & Format(mes, "00") & Format(dia, "00")
End Sub

Function DialogGetHwnd(Optional ByVal sDialogCaption As String = vbNullString, Optional sClassName As String = vbNullString) As Long
  On Error Resume Next
  
  DialogGetHwnd = FindWindowA(sClassName, sDialogCaption)
  On Error GoTo 0
End Function

Function AppToForeground(Optional sFormCaption As String, Optional lHwnd As Long, Optional lWindowState As Long = SW_NORMAL) As Boolean
  On Error Resume Next
  Dim tWinPlace As WINDOWPLACEMENT

  If lHwnd = 0 Then
    lHwnd = DialogGetHwnd(sFormCaption)
  End If
  If lHwnd Then
    tWinPlace.Length = Len(tWinPlace)
    
    'Get the windows current placement
    Call GetWindowPlacement(lHwnd, tWinPlace)
    
    'Set the windows placement
    tWinPlace.showCmd = lWindowState
    
    'Change window state
    Call SetWindowPlacement(lHwnd, tWinPlace)
    
    'Bring to foreground
    AppToForeground = SetForegroundWindow(lHwnd)
  End If
End Function

Sub ColocaTelaControleOperacionanoModoMaximizado()
  On Error Resume Next
  If GlobalIDControleOperacional Then
    If AppToForeground(, GlobalIDControleOperacional, SW_MAXIMIZE) Then
      ' a tela foi maximizada
    Else
      MsgBox "Falha ao maximizar a tela Controle Operacional", vbCritical, "Maximiza Controle Operacional"
    End If
  Else
    MsgBox "N�o foi encontrada tela Controle Operacional"
  End If
End Sub

Sub AtualizaListadeRequerimentos(ATUAL As Long)
  On Error Resume Next
  Dim conta As Long
  Dim sLinha As String
  Dim sTodos As String
  Dim M As Long

  Me.Left = 600
  sTodos = ""
  If GlobalAgenciaEscolhida = "" Then
    lblRequerimentodoSABI.Caption = "Agendamentos  do dia " & Mid(GlobalDatadosRequerimentos, 7, 2) & "/" & Mid(GlobalDatadosRequerimentos, 5, 2) & "/" & Mid(GlobalDatadosRequerimentos, 1, 4)
  Else
    lblRequerimentodoSABI.Caption = Mid(GlobalAgenciaEscolhida, 1, 40) & ", " & Mid(GlobalDatadosRequerimentos, 7, 2) & "/" & Mid(GlobalDatadosRequerimentos, 5, 2) & "/" & Mid(GlobalDatadosRequerimentos, 1, 4)
  End If
  lstMostrarRequerimentos.Visible = False
  lstMostrarRequerimentos.Clear
  lstMostrarRequerimentos.AddItem AcertaLarguraRelatorio("000", "Seq.") & AcertaLarguraRelatorio("123456789", "Requerim.") & AcertaLarguraRelatorio("INICIAL", "Tipo") & AcertaLarguraRelatorio("INDEFERIDO", "Status") & AcertaLarguraRelatorio("12345678901", "NIT") & AcertaLarguraRelatorio("IMPRESSO", "Impresso") & AcertaLarguraRelatorio("JOSE GERALDO DA COSTA", "Segurado")
  For conta = 1 To GlobalQuantidadedeRequerimentos
    GlobalRequerimentos(conta).sequencia = Format(conta, "000")
    sLinha = AcertaLarguraRelatorio("000", GlobalRequerimentos(conta).sequencia)
    sLinha = sLinha & AcertaLarguraRelatorio("123456789", GlobalRequerimentos(conta).N�mero)
    sLinha = sLinha & AcertaLarguraRelatorio("INICIAL", GlobalRequerimentos(conta).Tipo)
    sLinha = sLinha & AcertaLarguraRelatorio("INDEFERIDO", GlobalRequerimentos(conta).Status)
    sLinha = sLinha & AcertaLarguraRelatorio("12345678901", GlobalRequerimentos(conta).NIT)
    If conta <= ATUAL And GlobalRequerimentos(conta).Impresso <> "SIM" Then GlobalRequerimentos(conta).Impresso = "N�O"
    sLinha = sLinha & AcertaLarguraRelatorio("IMPRESSO", GlobalRequerimentos(conta).Impresso)
    sLinha = sLinha & GlobalRequerimentos(conta).Segurado
    sLinha = sLinha & "     " & GlobalRequerimentos(conta).Cr�tica
    lstMostrarRequerimentos.AddItem sLinha
    sLinha = GlobalRequerimentos(conta).sequencia & Chr(9)
    sLinha = sLinha & GlobalRequerimentos(conta).N�mero & Chr(9)
    sLinha = sLinha & GlobalRequerimentos(conta).Tipo & Chr(9)
    sLinha = sLinha & GlobalRequerimentos(conta).Status & Chr(9)
    sLinha = sLinha & GlobalRequerimentos(conta).NIT & Chr(9)
    sLinha = sLinha & GlobalRequerimentos(conta).Impresso & Chr(9)
    sLinha = sLinha & GlobalRequerimentos(conta).Segurado & Chr(9)
    sLinha = sLinha & GlobalRequerimentos(conta).Cr�tica
    sTodos = sTodos & sLinha & Chr(13) & Chr(10)
  Next conta
  lstMostrarRequerimentos.ListIndex = ATUAL
  lstMostrarRequerimentos.Visible = True
  DoEvents
  Open GlobalPastadeTrabalho & "\" & GlobalDatadosRequerimentos & ".txt" For Output As #1
  Print #1, sTodos
  Close #1
End Sub
        
Sub ApresentaRelatorioFinal()
  On Error Resume Next
  Dim conta As Long
  Dim lhWndNotepad As Long
  Dim hPrimeira As Long
  Dim Size As RECT
  Dim memo As String
  Dim IDTela As Long
  Dim res As String
  Dim Nome As String
  Dim segundos As Long
  Dim minutos As Long
  
  lstMostrarRequerimentos.Enabled = True
  segundos = Int((GetTickCount - GlobalInicio) / 1000)
  minutos = Int(segundos / 60)
  segundos = segundos - minutos * 60
  lblRelogio.Caption = " " & minutos & ":" & Format(segundos, "00") & " "
  memo = "O Controle Operacional foi fechado por medida de seguran�a" & Chr(13) & Chr(10) & Chr(13) & Chr(10)
  memo = memo & lblLocaleData.Caption & Chr(13) & Chr(10)
  memo = memo & AcertaLarguraRelatorio("123", "Seq") & Chr(9) & AcertaLarguraRelatorio("123456789", "Requerimento") & Chr(9) & AcertaLarguraRelatorio("12345678901", "CPF") & Chr(9) & AcertaLarguraRelatorio("INICIAL", "Tipo") & Chr(9) & AcertaLarguraRelatorio("INDEFERIDO", "Status") & Chr(9) & AcertaLarguraRelatorio("12345678901", "NIT") & Chr(9) & AcertaLarguraRelatorio("N�O", "IMPRESSO") & Chr(9) & AcertaLarguraRelatorio("JOSE GERALDO DA COSTA", "Segurado") & Chr(9) & "Cr�tica"
  For conta = 1 To GlobalQuantidadedeRequerimentos
    memo = memo & Chr(13) & Chr(10)
    memo = memo & AcertaLarguraRelatorio("123", Format(conta, "000"))
    memo = memo & Chr(9) & AcertaLarguraRelatorio("123456789", GlobalRequerimentos(conta).N�mero)
    memo = memo & Chr(9) & AcertaLarguraRelatorio("12345678901", GlobalRequerimentos(conta).CPF)
    memo = memo & Chr(9) & AcertaLarguraRelatorio("INICIAL", GlobalRequerimentos(conta).Tipo)
    memo = memo & Chr(9) & AcertaLarguraRelatorio("INDEFERIDO", GlobalRequerimentos(conta).Status)
    memo = memo & Chr(9) & AcertaLarguraRelatorio("12345678901", GlobalRequerimentos(conta).NIT)
    memo = memo & Chr(9) & AcertaLarguraRelatorio("N�O", GlobalRequerimentos(conta).Impresso)
    memo = memo & Chr(9) & AcertaLarguraRelatorio("JOSE GERALDO DA COSTA", GlobalRequerimentos(conta).Segurado)
    memo = memo & Chr(9) & Trim(GlobalRequerimentos(conta).Cr�tica)
  Next conta
  lhWndNotepad = 0
  Call Shell("notepad", vbNormalFocus)    'you'll need notepad.exe on your PC for this to work
  DoEvents
  Do While lhWndNotepad = 0
    lhWndNotepad = FindWindow(vbNullString, "Sem t�tulo - Bloco de Notas")
  Loop
    hPrimeira = FindWindowEx(lhWndNotepad, 0, "Edit", vbNullString)
    SendMessage2 lhWndNotepad, WM_SETTEXT, 0, "Relat�rio DIB DIP E Gcont" & Chr$(0)
    res = SetWindowPos(lhWndNotepad, 0, 100, 50, 700, 600, 0)
    res = SetWindowPos(lhWndNotepad, -1, 0, 0, 0, 0, 3)
    SetForegroundWindow lhWndNotepad
    SendMessage2 hPrimeira, WM_SETTEXT, 0, memo & Chr$(0)
    GlobalRelatorioPronto = True
    res = SetWindowPos(lhWndNotepad, -1, 0, 0, 0, 0, 3)
  End
End Sub

Private Sub ClickOpen(hMsgBox As Long)
  On Error Resume Next
  Dim hButtonOpen As Long
  Dim hComboBox As Long
  
  hButtonOpen = FindWindowEx(hMsgBox, 0, "Button", "OK")
  SendMessage hButtonOpen, BM_CLICK, 0, 0
End Sub

Private Function Crystal(hMsgBox As Long) As Boolean
  On Error Resume Next
  Dim h1AfxWnd42 As Long
  Dim h2AfxWnd42 As Long
  Dim hAfxFrameOrView42 As Long
  
  h1AfxWnd42 = FindWindowEx(hMsgBox, 0, "AfxWnd42", "")
  h2AfxWnd42 = FindWindowEx(h1AfxWnd42, 0, "AfxWnd42", "")
  hAfxFrameOrView42 = FindWindowEx(h2AfxWnd42, 0, "AfxFrameOrView42", "")
  Crystal = hAfxFrameOrView42 <> 0
End Function

Function achaTelaInternaAtiva(NomedaTela As String) As Long
  On Error Resume Next
  Dim lngHWnd As Long
  Dim lngHWnd2 As Long
  Dim titletmp As String
  Dim nret As Long
  Dim Size As RECT
  Dim lhWndP As Long
    
  achaTelaInternaAtiva = 0
  lhWndP = FindWindow(vbNullString, vbNullString) 'PARENT WINDOW
  Do While lhWndP <> 0
    lngHWnd = FindWindowEx(lhWndP, 0, vbNullString, vbNullString)
    Do Until lngHWnd = 0
      If lngHWnd > 0 Then
        lngHWnd2 = FindWindowEx(lngHWnd, 0, vbNullString, vbNullString)
        Do Until lngHWnd2 = 0
          If lngHWnd2 > 0 Then
            titletmp = Space(256)
            nret = GetWindowText(lngHWnd2, titletmp, Len(titletmp))
            If InStr(1, titletmp, NomedaTela) > 0 Then
              achaTelaInternaAtiva = lngHWnd2
              Exit Function
            End If
          End If
          lngHWnd2 = FindWindowEx(lngHWnd2, lngHWnd, vbNullString, vbNullString)
        Loop
      End If
      lngHWnd = FindWindowEx(lhWndP, lngHWnd, vbNullString, vbNullString)
    Loop
    lhWndP = GetWindow(lhWndP, GW_HWNDNEXT)
  Loop
End Function

Function esperaCRYSTALREPORTeExporta() As String
  On Error Resume Next
  Dim COsize As RECT
  Dim Size As RECT
  Dim titletmp As String
  Dim nret As Long
  Dim TelaSize As RECT
  Dim arquivo As String
  Dim memo As String
  Dim hDC As Long
  Dim lcount As Long
  Dim LocalIDBot�oSalvar As Long
  Dim hNomedoArquivo As Long
  Dim hDUIView As Long
  Dim hDirectUI  As Long
  Dim hFloatNotify As Long
  Dim hComboBox As Long
  Dim hBot�oSalvar As Long
  Dim hDestinoExport As Long
  Dim hBot�oOKExport As Long
  Dim hFormatoExport As Long
  Dim conta As Long
  Dim hBlocodeNotas As Long
  Dim childhandle As Long
  Dim texto As String
  Dim SaveAsDialog As Long
  
  hDC = GetWindowDC(0)
  GlobalIDTelaSalvarComo = 0
  GlobalIDTelaRequerimentosCrystalReport = ObtemIDdoRelat�rioCrystalReport
  While GlobalIDTelaRequerimentosCrystalReport = 0
    GlobalIDTelaRequerimentosCrystalReport = ObtemIDdoRelat�rioCrystalReport
    Espera 300
    DoEvents
    conta = conta + 1
    If conta > 200 Then
      MsgBox "A tela 'Imprimir Agendamento' n�o apareceu.", vbCritical, "Tela Imprimir Agendamento"
      esperaCRYSTALREPORTeExporta = "A tela 'Imprimir Agendamento' n�o apareceu."
      Exit Function
    End If
  Wend
    
  'fecha a tela Imprimir Agendamentos
  SendMessage GlobalIDTelaImprimirAgendamento, WM_CLOSE, 0, 0
  res = SetWindowPos(GlobalIDTelaRequerimentosCrystalReport, 0, 0, 0, 800, 460, 0)
  SetForegroundWindow (GlobalIDTelaRequerimentosCrystalReport)
  Espera 1000  'com 300 falhou com o Bruno - clicou antes da hora
  
  'implementar rotina que repete o clique periodicamente ate� vir a nova tela
  SetForegroundWindow (GlobalIDTelaRequerimentosCrystalReport)
  DoEvents
  Me.Top = Screen.Height - 760 - 3000
  Espera 500
  MouseClique 262, 44
  DoEvents
  Espera 500
  '    nao funcionou
                
  'espera tela Export
  titletmp = Space(256)
  nret = GetWindowText(GlobalIDTelaAtiva, titletmp, Len(titletmp))
  GlobalT�tulodaTelaAtiva = titletmp
  While Mid(GlobalT�tulodaTelaAtiva, 1, 6) <> "Export"
    GlobalIDTelaAtiva = GetForegroundWindow
    titletmp = Space(256)
    nret = GetWindowText(GlobalIDTelaAtiva, titletmp, Len(titletmp))
    GlobalT�tulodaTelaAtiva = titletmp
    DoEvents
    Espera 300
    If InStr(1, titletmp, "SABI - Controle Operacional") > 0 Then
      ClickOpen (GlobalIDTelaAtiva)
    Else
      If Len(Trim(titletmp)) = 1 Then MouseClique 262, 44
    End If
  Wend
  
  'espera a tela export
  GlobalIDTelaExport = ObtemIDdaTelaPrincipalporTitulo("Export")
  While GlobalIDTelaExport = 0
    Espera 300
    DoEvents
    GlobalIDTelaExport = ObtemIDdaTelaPrincipalporTitulo("Export")
  Wend
  If GlobalIDTelaExport > 0 Then
    hBlocodeNotas = 0
    lcount = 0
    hDestinoExport = 0
    Do While hDestinoExport = 0 Or lcount > 10
      hBot�oOKExport = FindWindowEx(GlobalIDTelaExport, 0, "Button", "OK")
            
      'encontra DirectUIHWND
      hFormatoExport = FindWindowEx(GlobalIDTelaExport, 0, "ComboBox", "")
      hDestinoExport = FindWindowEx(GlobalIDTelaExport, hFormatoExport, "ComboBox", "")
      lcount = lcount + 1
      Espera 300
      DoEvents
    Loop
       
    'o destino deve ser escolhido antes do formato para n�o gerar erro de e-mail n�o configurado
    'destino
    SendMessageByLong hDestinoExport, CB_SETCURSEL, 1, 0&
    DoEvents
    Sleep 100
    conta = 0
    If ObtemTextodoControle(hDestinoExport) <> "Disk file" Then
      For conta = 0 To 100
        SendMessageByLong hDestinoExport, CB_SETCURSEL, conta, 0&
        DoEvents
        Sleep 100
        If ObtemTextodoControle(hDestinoExport) = "Disk file" Then Exit For
      Next conta
      DoEvents
    End If
    If conta > 99 Then
      MsgBox "N�o foi encontrada a op��o de exportar para 'Disk file'", vbCritical, "Export"
      End
    End If
        
    'formato
    SendMessageByLong hFormatoExport, CB_SETCURSEL, 22, 0&
    DoEvents
    Sleep 100
    conta = 0
    If ObtemTextodoControle(hFormatoExport) <> "Tab-separated text" Then
      For conta = 0 To 100
      SendMessageByLong hFormatoExport, CB_SETCURSEL, conta, 0&
      DoEvents
      Sleep 100
      If ObtemTextodoControle(hFormatoExport) = "Tab-separated text" Then Exit For
        Next conta
      End If
      If conta > 99 Then
        MsgBox "N�o foi encontrada a op��o de exportar para o formato 'Tab-separated text'", vbCritical, "Export"
      End
    End If
    Espera 100
    SendMessage hBot�oOKExport, BM_CLICK, 0, 0
    DoEvents
    Espera 300
    DoEvents
    Espera 1000
    SendMessage GlobalIDTelaExport, WM_CLOSE, 0, 0
    DoEvents
  Else
    MsgBox "tela export n�o abriu", vbCritical, "Tela Export"
    End
  End If
    
  'espera tela "Choose Export File"
  SaveAsDialog = FindWindow("#32770", "Choose Export File")
  conta = 0
  While SaveAsDialog = 0
    Espera 300
    conta = conta + 1
    If conta > 200 Then
      MsgBox "A tela 'Choose Export File' n�o apareceu.", vbCritical, "Tela Choose Export File"
      esperaCRYSTALREPORTeExporta = "A tela 'Choose Export File' n�o apareceu."
      Exit Function
    End If
  Wend
  configuranomedoarquivo
    
  'fecha relatorio Crystal Report
  SendMessage GlobalIDTelaRequerimentosCrystalReport, WM_CLOSE, 0, 0
  DoEvents
  While ObtemIDdoRelat�rioCrystalReport <> 0
    SendMessage GlobalIDTelaRequerimentosCrystalReport, WM_CLOSE, 0, 0
    DoEvents
    Sleep 300
  Wend
  decodeRequerimentos texto
End Function


Sub MontaListadeRequerimentos(memotexto As String)
  Dim pos As Long
  Dim linha As String
  Dim conta As Long
  Dim indice As Long
  Dim nomedosegurado As String
  Dim GlobalRequerimentosProv(1000) As Requerimento
  Dim LINHA2 As String
  Dim pos3 As Long
  Dim conta2 As Long
  
  indice = 0
  conta = 0
  GlobalQuantidadedeRequerimentos = GlobalAgendamentosQuandidade
  lstClassificar.Clear
  If optOrdem(1).Value = True Then
    For conta = 1 To GlobalAgendamentosQuandidade
      lstClassificar.AddItem Mid(GlobalAgendamentosConsulta(conta).Horario, 1, 5) & " - " & GlobalAgendamentosConsulta(conta).Segurado
      lstClassificar.ItemData(lstClassificar.NewIndex) = GlobalAgendamentosConsulta(conta).Requerimento
    Next conta
  Else
    For conta = 1 To GlobalAgendamentosQuandidade
      lstClassificar.AddItem GlobalAgendamentosConsulta(conta).Segurado
      lstClassificar.ItemData(lstClassificar.NewIndex) = GlobalAgendamentosConsulta(conta).Requerimento
    Next conta
  End If
  lstMostrarRequerimentos.Clear
  lstMostrarRequerimentos.AddItem "Seq." & Chr(9) & "Requerim." & Chr(9) & Chr(9) & Chr(9) & "Segurado"
  For conta = 0 To lstClassificar.ListCount - 1
    For conta2 = 1 To GlobalAgendamentosQuandidade
      If GlobalAgendamentosConsulta(conta2).Requerimento = (lstClassificar.ItemData(conta)) Then
        Exit For
      End If
    Next conta2
    indice = indice + 1
    GlobalRequerimentos(conta + 1).N�mero = lstClassificar.ItemData(conta)
    GlobalRequerimentos(conta + 1).Segurado = lstClassificar.List(conta)
    lstMostrarRequerimentos.AddItem Format(indice, "000") & Chr(9) & lstClassificar.ItemData(conta) & Chr(9) & lstClassificar.List(conta)
    
    'valores iniciais
    GlobalRequerimentos(conta).Tipo = ""
    GlobalRequerimentos(conta).Status = ""
    GlobalRequerimentos(conta).NIT = ""
    GlobalRequerimentos(conta).Cr�tica = ""
  Next conta
  lstMostrarRequerimentos.Visible = True
  If GlobalQuantidadedeRequerimentos > 0 Then
    LocalCopiar = True
    Me.Top = Screen.Height - 760 - 3000
    Me.Width = 12540 + 1600
    Me.Height = Screen.Height - 600
    SetForegroundWindow (Me.hWnd)
    fraOrdem.Visible = False
    fraImprime.Visible = True
    mostralista
    If GlobalQuantidadedeRequerimentos > 40 Then
      paracima.Visible = True
      parabaixo.Visible = True
    End If
    txtUltimo.Text = GlobalQuantidadedeRequerimentos
  Else
    res = SetWindowPos(Me.hWnd, -2, 0, 0, 0, 0, 3)
    MsgBox "N�o foi encontrado nenhum agendamento de per�cia para esta data" & Chr(13) & Chr(10) & GlobalAgendamentosConsultaCabecalho, vbCritical, "Agendamentos do SABI"
    End
  End If
 
End Sub
  
Private Function ObtemIDdaTelaPrincipalporTitulo(ByVal sCaption As String) As Long
  On Error Resume Next
  Dim lhWndP As Long
  Dim sStr As String

  ObtemIDdaTelaPrincipalporTitulo = False
  lhWndP = FindWindow(vbNullString, vbNullString) 'PARENT WINDOW
  Do While lhWndP <> 0
    sStr = String(GetWindowTextLength(lhWndP) + 1, Chr$(0))
    GetWindowText lhWndP, sStr, Len(sStr)
    sStr = Left$(sStr, Len(sStr) - 1)
    If InStr(1, sStr, sCaption) > 0 Then
      ObtemIDdaTelaPrincipalporTitulo = lhWndP
      Exit Function
    End If
    lhWndP = GetWindow(lhWndP, GW_HWNDNEXT)
  Loop
  ObtemIDdaTelaPrincipalporTitulo = 0
End Function

Sub MouseClique(posx As Long, posy As Long)
  On Error Resume Next
  Dim Size As RECT
  Dim IDTelaAtiva As Long
  Dim pt As POINTAPI
  
  GetCursorPos pt
  IDTelaAtiva = GetForegroundWindow
  res = GetWindowRect(IDTelaAtiva, Size)
  SetCursorPos Size.Left + posx, Size.Top + posy
  mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
  SetCursorPos pt.x, pt.y
End Sub
    
Sub Get_User_Name()
  On Error Resume Next 'voltar
  Dim lpBuff As String * 25
  Dim ret As Long

  'Get the user name minus any trailing spaces found in the name.
  ret = GetUserName(lpBuff, 25)
  GlobalUserName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
  GlobalAreadeTrabalho = getSpecialFolder(CSIDL_DESKTOP)
  GlobalPastadeTrabalho = getSpecialFolder(CSIDL_LOCAL_APPDATA) & "\" & NomeAplicacao
End Sub
    
Private Sub chkIniciais_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  On Error Resume Next
  If chkIniciais.Value = 1 And chkPP.Value = 0 Then SaveSetting "AGENDAMENTODOSABI", "IMPRIMIR", "EXAMES", "INICIAIS"
  If chkIniciais.Value = 1 And chkPP.Value = 1 Then SaveSetting "AGENDAMENTODOSABI", "IMPRIMIR", "EXAMES", "TODOS"
  If chkIniciais.Value = 0 And chkPP.Value = 0 Then SaveSetting "AGENDAMENTODOSABI", "IMPRIMIR", "EXAMES", "NENHUM"
  If chkIniciais.Value = 0 And chkPP.Value = 1 Then SaveSetting "AGENDAMENTODOSABI", "IMPRIMIR", "EXAMES", "PP"
End Sub

Private Sub chkPP_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  On Error Resume Next
  If chkIniciais.Value = 1 And chkPP.Value = 0 Then SaveSetting "AGENDAMENTODOSABI", "IMPRIMIR", "EXAMES", "INICIAIS"
  If chkIniciais.Value = 1 And chkPP.Value = 1 Then SaveSetting "AGENDAMENTODOSABI", "IMPRIMIR", "EXAMES", "TODOS"
  If chkIniciais.Value = 0 And chkPP.Value = 0 Then SaveSetting "AGENDAMENTODOSABI", "IMPRIMIR", "EXAMES", "NENHUM"
  If chkIniciais.Value = 0 And chkPP.Value = 1 Then SaveSetting "AGENDAMENTODOSABI", "IMPRIMIR", "EXAMES", "PP"
End Sub

Private Sub cmdContinua_Click()
  On Error Resume Next
  Dim res As String

  GlobalMedidadeSeguranca = True
  Timer2.Enabled = False
  paracima.Visible = False
  parabaixo.Visible = False
  cmdIniciar.Enabled = False
  cmdContinua.Enabled = False
  txttPrimeiro.Enabled = False
  txtUltimo.Enabled = False
  chkIniciais.Enabled = False
  chkPP.Enabled = False
  lblversao.Visible = False
  GlobalIDControleOperacional = ObtemIDdaTelaPrincipalporTitulo("SABI - M�dulo de Controle Operacional")
  If GlobalIDControleOperacional <> 0 Then
  Else
    res = SetWindowPos(Me.hWnd, -2, 0, 0, 0, 0, 3)
    Me.Left = 600
    MsgBox "Abra o m�dulo Controle Operacional do SABI.", vbCritical, "Agendamentos do SABI"
    res = SetWindowPos(Me.hWnd, -1, 0, 0, 0, 0, 3)
    Exit Sub
  End If

  Dim titletmp As String
  Dim nret As Long
  Dim localMenu As Long
  Dim Rela��odeMenu As String
  Dim newMenu As Long
  Dim hToolbar20WndClass As Long
  Dim hmsvb_lib_Tollbar As Long
  Dim hPrimeira As Long
  Dim hSegunda As Long
  Dim ClassedoControle As String
  Dim classlength As Long
  Dim IDtelasInternasdoSABI As Long
  Dim IDApagaTela As Long
  Dim contavezes As Long
  Dim IDPrimeiroToolBar As Long
    
  'limpa todas telas internas do SABI
  'resta agora as telas externas
  verificaeapaga "Imprimir Agendamento"
  verificaeapaga "Imprimir Escala"
  verificaeapaga "Marca��o da Avalia��o Social"
  verificaeapaga "Segunda Via de Carta de Exig�ncia"
  verificaeapaga "Pesquisa de Requerente"
  GlobalhMDIClient = 0
  GlobalhMDIClient = FindWindowEx(GlobalIDControleOperacional, 0, "MDIClient", "")
  If GlobalhMDIClient = 0 Then
    MsgBox "N�o foi encontrada o indentificador da tela de fundo do Controle Operacional.", vbCritical, "Agendamentos do SABI"
    End
  End If
  IDtelasInternasdoSABI = 0
  IDtelasInternasdoSABI = FindWindowEx(GlobalhMDIClient, 0, vbNullString, vbNullString)
  While IDtelasInternasdoSABI <> 0
    If IDtelasInternasdoSABI <> 0 Then IDApagaTela = IDtelasInternasdoSABI
    DoEvents
    Espera 300
    IDtelasInternasdoSABI = FindWindowEx(GlobalhMDIClient, IDtelasInternasdoSABI, vbNullString, vbNullString)
    If IDApagaTela <> 0 Then SendMessage IDApagaTela, WM_CLOSE, 0, 0
  Wend
  
  'todas as telas internas foram limpas
  DoEvents
  Espera 300
  
  'acerta a tela
  ColocaTelaControleOperacionanoModoNormal
  ColocaTelaControleOperacionanoModoMaximizado

 'Abre tela Consulta Requerimento/Benef�cio
  ClickMenu GlobalIDControleOperacional, 2, 0
  Espera 3000
  IDtelasInternasdoSABI = 0
  contavezes = 0
  While IDtelasInternasdoSABI = 0
    IDtelasInternasdoSABI = FindWindowEx(GlobalhMDIClient, 0, "ThunderRT6FormDC", "Consulta Requerimento/Benef�cio")
    Espera 300
    DoEvents
    contavezes = contavezes + 1
    If contavezes > 400 Then
      MsgBox "O SABI est� muito lento. Tente outra hora.", vbCritical, "Consulta Requerimento/Benef�cio"
      End
    End If
  Wend
  DoEvents
  GlobalIDTelaConsultaRequerimentoBenef�cio = 0
  While GlobalIDTelaConsultaRequerimentoBenef�cio = 0
    Espera 300
    DoEvents
    GlobalIDTelaConsultaRequerimentoBenef�cio = achaTelaInternaAtiva("Consulta Requerimento/Benef�cio")
  Wend
  IDPrimeiroToolBar = FindWindowEx(IDtelasInternasdoSABI, 0, "Toolbar20WndClass", "")
  GlobalToolbarConsultaRequerimentoOCX = FindWindowEx(IDtelasInternasdoSABI, IDPrimeiroToolBar, "Toolbar20WndClass", "")
  Espera 300
  GlobalToolbarConsultaRequerimento = FindWindowEx(GlobalToolbarConsultaRequerimentoOCX, 0, "msvb_lib_toolbar", vbNullString)
  res = SetWindowPos(Me.hWnd, -1, 0, 0, 0, 0, 3)
  If chkIniciais.Value = 1 And chkPP.Value = 1 Then Imprimeosrequerimentos ("TODOS")
  If chkIniciais.Value = 1 And chkPP.Value = 0 Then Imprimeosrequerimentos ("INICIAL")
  If chkIniciais.Value = 0 And chkPP.Value = 1 Then Imprimeosrequerimentos ("PP")
  If chkIniciais.Value = 0 And chkPP.Value = 0 Then Imprimeosrequerimentos ("NENHUM")
  SendMessageByLong lstMostrarRequerimentos.hWnd, LB_SETHORIZONTALEXTENT, 1200, 0
  SendMessageByLong lstMostrarRequerimentos.hWnd, WM_VSCROLL, SB_BOTTOM, 0
  SetForegroundWindow (GlobalIDControleOperacional)
  Sleep 300
  ClickMenu GlobalIDControleOperacional, 0, 6
  End
End Sub

Private Sub cmdFechar_Click()
  End
End Sub

Private Sub cmdFechar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Me.MousePointer = 0
End Sub

Private Sub Imprimeosrequerimentos(abrangencia As String)
  Dim RequerimentoAtual As Requerimento
  Dim requerimentosn�omarcados As String
  Dim conta As Long
  Dim res As String
  Dim memo
  Dim contador As Long
  Dim titletmp  As String
  Dim nret As Long
  Dim idteladibdip As Long
  Dim hThunderRT6CommandButtonDIBDIPGCONT As Long
  Dim hThunderRT6CommandButtonFECHAR As Long
  Dim Size As RECT
  Dim IDTelaAtiva As Long
  Dim pt As POINTAPI
  Dim numCPF As String
  Dim idtelanafrente As Long

  lstMostrarRequerimentos.Enabled = False
  DoEvents
  
  'for�a resize
  GlobalRelatorioPronto = False
  pctCopiaPartedaTela.Visible = False
  LocalCopiar = True
  Me.Height = 2000
  
  'le lstMostrarRequerimentos
  requerimentosn�omarcados = ""
  For conta = 1 To lstMostrarRequerimentos.ListCount - 1
    If lstMostrarRequerimentos.Selected(conta) = False Then
      requerimentosn�omarcados = requerimentosn�omarcados & lstMostrarRequerimentos.List(conta)
    End If
  Next conta
  GlobalHoradeInicio = Time
  ColocaTelaControleOperacionanoModoMaximizado
  GlobalInicio = GetTickCount
  For GlobalIDRequerimento = 1 To GlobalQuantidadedeRequerimentos
    atualizaprogresso
    Me.Visible = True
    DoEvents
    mostratela
  
    'so atua nos requerimentos sem marca de impress�o
    If GlobalRequerimentos(GlobalIDRequerimento).Impresso <> "SIM" And GlobalRequerimentos(GlobalIDRequerimento).Impresso <> "N�O" Then
      RequerimentoAtual = ConsultaRequerimento(GlobalRequerimentos(GlobalIDRequerimento).N�mero, GlobalImpress�oAutom�tica)
      If RequerimentoAtual.Cr�tica = "" Then
        GlobalRequerimentos(GlobalIDRequerimento).NIT = Mid(RequerimentoAtual.NIT, 1, 11)
        GlobalRequerimentos(GlobalIDRequerimento).Tipo = Mid(RequerimentoAtual.Tipo, 1, 7)
        GlobalRequerimentos(GlobalIDRequerimento).Status = Mid(RequerimentoAtual.Status, 1, 10)
        AtualizaListadeRequerimentos (GlobalIDRequerimento)
        If IsNumeric(GlobalRequerimentos(GlobalIDRequerimento).NIT) And GlobalRequerimentos(GlobalIDRequerimento).Status = "NORMAL" Then
          'rotina de clicar em DIB/DIP e Gcont
          Me.Left = 600
          DoEvents
          conta = 0
          While GetForegroundWindow <> GlobalIDControleOperacional
            SetForegroundWindow (GlobalIDControleOperacional)
            DoEvents
            Espera 100
            conta = conta + 1
            If conta > 100 Then Exit Sub
          Wend
          GetCursorPos pt
          SetCursorPos 800, 320
          MouseClique 800, 320
          SetCursorPos pt.x, pt.y
    
          'espera tela "Detalhes Requerimento/Benef�cio"
          titletmp = Space(256)
          nret = GetWindowText(GlobalIDTelaAtiva, titletmp, Len(titletmp))
          GlobalT�tulodaTelaAtiva = titletmp
          While InStr(1, GlobalT�tulodaTelaAtiva, "Detalhes Requerimento/Benef�cio") = 0
            GlobalIDTelaAtiva = GetForegroundWindow
            idteladibdip = GetForegroundWindow
            titletmp = Space(256)
            nret = GetWindowText(GlobalIDTelaAtiva, titletmp, Len(titletmp))
            GlobalT�tulodaTelaAtiva = titletmp
            DoEvents
            Espera 300
            mostratela
          Wend
          Dim DimensoesdatelaImprimir As RECT
          res = GetWindowRect(idteladibdip, DimensoesdatelaImprimir)
          Espera 300
          mostratela
    
          'clica na Aba Documentos para indentificar o CPF
          GetCursorPos pt
          IDTelaAtiva = GetForegroundWindow
          res = GetWindowRect(IDTelaAtiva, Size)
          SetCursorPos Size.Left + 110, Size.Top + 40
          mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
          mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
          SetCursorPos pt.x, pt.y
          Espera 300
          mostratela
          CopiaTelaCPF 1, idteladibdip
          numCPF = CapturaNumeroDetalhes
          If numCPF = "" Then
            GlobalRequerimentos(GlobalIDRequerimento).CPF = "-----------"
          Else
            GlobalRequerimentos(GlobalIDRequerimento).CPF = numCPF
          End If
          hThunderRT6CommandButtonDIBDIPGCONT = 0
          hThunderRT6CommandButtonFECHAR = 0
          Do While hThunderRT6CommandButtonDIBDIPGCONT = 0 Or hThunderRT6CommandButtonFECHAR = 0
            hThunderRT6CommandButtonDIBDIPGCONT = FindWindowEx(idteladibdip, 0, "ThunderRT6CommandButton", "DIB/DIP e Gcont")
            hThunderRT6CommandButtonFECHAR = FindWindowEx(idteladibdip, 0, "ThunderRT6CommandButton", "&Fechar")
            contador = contador + 1
            If contador > 5000 Then
              'fecha a tela Segunda Via de Marca��o de Exame
              Exit Do
            End If
          Loop
          If hThunderRT6CommandButtonDIBDIPGCONT > 0 Then
            Espera 300
            mostratela
            PostMessage hThunderRT6CommandButtonDIBDIPGCONT, BM_CLICK, 0, 0
            Espera 300
            mostratela
          End If
          If hThunderRT6CommandButtonFECHAR > 0 Then
            Espera 300
            mostratela
            PostMessage hThunderRT6CommandButtonFECHAR, BM_CLICK, 0, 0
          Else
            'fecha com clique
            Espera 1000
            mostratela
            MouseClique 1010, 630
          End If
    
          'espera tela "Detalhes Requerimento/Benef�cio" fechar
          titletmp = Space(256)
          nret = GetWindowText(GlobalIDTelaAtiva, titletmp, Len(titletmp))
          GlobalT�tulodaTelaAtiva = titletmp
          While InStr(1, GlobalT�tulodaTelaAtiva, "Detalhes Requerimento/Benef�cio") > 0
            GlobalIDTelaAtiva = GetForegroundWindow
            titletmp = Space(256)
            nret = GetWindowText(GlobalIDTelaAtiva, titletmp, Len(titletmp))
            GlobalT�tulodaTelaAtiva = titletmp
            SetForegroundWindow idteladibdip
            Espera 300
            mostratela
          Wend
                        
          'fim da rotina de DIB/DIP e Gcont
          If Val(GlobalRequerimentos(GlobalIDRequerimento).sequencia) >= Val(txttPrimeiro.Text) And Val(GlobalRequerimentos(GlobalIDRequerimento).sequencia) <= Val(txtUltimo.Text) Then
            If abrangencia <> "NENHUM" Then
              If abrangencia = "TODOS" Or GlobalRequerimentos(GlobalIDRequerimento).Tipo = abrangencia Then
                RequerimentoAtual = ImprimeSegundaViadoRequerimento(RequerimentoAtual.NIT, GlobalImpress�oAutom�tica)
                If RequerimentoAtual.Cr�tica <> "" Then RequerimentoAtual.Impresso = ""
                GlobalRequerimentos(GlobalIDRequerimento).Cr�tica = RequerimentoAtual.Cr�tica
                GlobalRequerimentos(GlobalIDRequerimento).Impresso = RequerimentoAtual.Impresso
                AtualizaListadeRequerimentos (GlobalIDRequerimento)
                If GlobalRequerimentos(GlobalIDRequerimento).Cr�tica <> "" Then
                  vermelho
                Else
                  If GlobalRequerimentos(GlobalIDRequerimento).Impresso = "SIM" Then desenhaimpressora 240, GlobalLinhaPicture - 14
                End If
              End If
            End If
          End If
        End If
      End If
    End If
  Next GlobalIDRequerimento
  Me.Top = Screen.Height - 760 - 3000
  pctProgressoFundo.Visible = False
  DoEvents
    
  pctCopiaPartedaTela.Visible = False
  LocalCopiar = False
  GlobalRelatorioPronto = True
  Me.Top = Screen.Height - 760 - 3000
  Me.Height = 2000
  Picture1.Top = 0

  ApresentaRelatorioFinal
End Sub

Private Sub cmdFechar2_Click()
  ClickMenu GlobalIDControleOperacional, 0, 6
  ApresentaRelatorioFinal
End Sub

Private Sub cmdIniciar_Click()
  Dim res As String
  
  lstClassificar.Visible = False
  lstMostrarRequerimentos.Visible = False
  If Dir(GlobalAreadeTrabalho & "\Agendamentos.txt") <> "" Then
    Kill GlobalAreadeTrabalho & "\Agendamentos.txt"
  End If
  lblversao.Visible = False
  cmdIniciar.Enabled = False

  GlobalIDControleOperacional = ObtemIDdaTelaPrincipalporTitulo("SABI - M�dulo de Controle Operacional")
  If GlobalIDControleOperacional <> 0 Then
    Timer1.Enabled = False
  Else
    res = SetWindowPos(Me.hWnd, -2, 0, 0, 0, 0, 3)
    Me.Left = 600
    MsgBox "Abra o m�dulo Controle Operacional do SABI.", vbCritical, "Agendamentos do SABI"
    res = SetWindowPos(Me.hWnd, -1, 0, 0, 0, 0, 3)
    mtempo1 = 60
    Exit Sub
  End If
  preparaSABI
  res = SetWindowPos(Me.hWnd, -1, 0, 0, 0, 0, 3)
End Sub

Private Sub cmdIniciar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Me.MousePointer = 0
End Sub

Private Sub desenhaimpressora(deslocamentox As Long, deslocamentoy As Long)
  Dim linha, coluna As Long
  
  For linha = 0 To pctImpressora.Height / 15 - 1
    For coluna = 0 To pctImpressora.Width / 15 - 1
      If pctImpressora.Point(coluna, linha) <> 0 Then Picture1.PSet (coluna + deslocamentox, linha + deslocamentoy), pctImpressora.Point(coluna, linha)
    Next coluna
  Next linha
  SavePicture Picture1.Image, GlobalPastadeTrabalho & "\" & GlobalDatadosRequerimentos & GlobalAgenciaEscolhida & "Todos.bmp"
End Sub

Private Sub mostralista()
  Dim RtnValue
  Dim win As Long
  Dim esquerda, altura, largura, dimensao As Long
  Dim lstdc As Long

  lstdc = GetWindowDC(lstMostrarRequerimentos.hWnd)
  esquerda = 900
  altura = 0
  largura = lstMostrarRequerimentos.Width
  dimensao = lstMostrarRequerimentos.Height
  win = GlobalhMDIClient
  cmdContinua.Visible = False
  cmdIniciar.Visible = False
  cmdFechar.Visible = False
  fraImprime.Visible = False
  lstMostrarRequerimentos.Top = 0
  Me.Height = Screen.Height
  Me.Refresh
  lstMostrarRequerimentos.Visible = True
  lstMostrarRequerimentos.Height = Me.Height - lstMostrarRequerimentos.Top - 40
  lstMostrarRequerimentos.Refresh
  pctEsteRequerimento.Refresh
  Me.Top = Screen.Height - 760 - 3000
  Me.Left = 600
  lstMostrarRequerimentos.Visible = True
  DoEvents
  RtnValue = BitBlt(GetDC(win), CLng(esquerda), _
  CLng(altura), CLng(largura), CLng(dimensao), lstdc, CLng(0), CLng(0), SRCCOPY)
  Me.Height = 2000
  lstMostrarRequerimentos.Visible = False
  lstMostrarRequerimentos.Top = 5000
  fraImprime.Left = fraOrdem.Left
  fraImprime.Top = Me.Height - fraImprime.Height - 120
  fraOrdem.Visible = False
  fraImprime.Visible = True
  cmdContinua.Visible = True
  cmdIniciar.Visible = True
  cmdFechar.Visible = True
  fraImprime.Visible = True
  mtempo2 = 0
  Timer2.Enabled = True
End Sub

Private Sub Form_Activate()
  Dim memo As String
  Dim conta As Long
    
  chkIniciais.Value = 1
  chkPP.Value = 0
  If GetSetting("AGENDAMENTODOSABI", "IMPRIMIR", "EXAMES", "") = "NENHUM" Then
    chkIniciais.Value = 0
    chkPP.Value = 0
  End If
  If GetSetting("AGENDAMENTODOSABI", "IMPRIMIR", "EXAMES", "") = "TODOS" Then
    chkIniciais.Value = 1
    chkPP.Value = 1
  End If
  If GetSetting("AGENDAMENTODOSABI", "IMPRIMIR", "EXAMES", "") = "INICIAIS" Then
    chkIniciais.Value = 1
    chkPP.Value = 0
  End If
  If GetSetting("AGENDAMENTODOSABI", "IMPRIMIR", "EXAMES", "") = "PP" Then
    chkIniciais.Value = 0
    chkPP.Value = 1
  End If
  Picture1.Width = 10890
  For conta = 0 To 10891
    Picture1.PSet (conta, 0), RGB(40, 40, 40)
    Picture1.PSet (conta, 1), RGB(40, 40, 40)
  Next conta
  If estaTemaAtivo = False Then
    MsgBox "Personalize a tela do seu computador com o tema 'Windows 7'", vbCritical, "Tema Aero"
    End
  End If
  GlobalInicio = GetTickCount
  If GlobalPrimeiraVez = False Then Exit Sub
  GlobalPrimeiraVez = False
  GlobalImpress�oAutom�tica = True
  If App.PrevInstance Then
    MsgBox "O aplicativo 'Agendamentos do SABI' j� est� aberto.", vbCritical, "Agendamentos do SABI"
    End
  End If

  'apaga todos  bmp de datas anteriores a atual
  memo = Dir(GlobalPastadeTrabalho & "\" & "*.bmp")
  While memo <> ""
    If Mid(memo, 1, 8) < Format(Date, "yyyymmdd") Then Kill GlobalPastadeTrabalho & "\" & memo
    memo = Dir()
  Wend
  
  'apaga todos  txt de datas anteriores a atual
  memo = Dir(GlobalPastadeTrabalho & "\" & "*.txt")
  While memo <> ""
    If Mid(memo, 1, 8) < Format(Date, "yyyymmdd") Then Kill GlobalPastadeTrabalho & "\" & memo
    memo = Dir()
  Wend
  memo = Dir(GlobalPastadeTrabalho & "\*" & ".txt")
  Do While memo <> ""
    If IsDate(Mid(memo, 7, 2) & "/" & Mid(memo, 5, 2) & "/" & Mid(memo, 1, 4)) Then
      Exit Do
    End If
    memo = Dir()
  Loop
End Sub

Private Sub Form_Load()
  Dim res As String
  Dim conta As Long
  Dim mPastaAgendamentosdoSABI As Boolean
  Dim sPath As String
  Dim lRet  As Long
  Dim nomeexecut�vel As String
  Dim pos As Long
  Dim pastaAppData As String
    
  GlobalMedidadeSeguranca = False
  mtempo1 = 0
  deslocalista = 0
  Picture1.BackColor = RGB(171, 171, 171)
  LocalCopiar = False
  GlobalLinhaPicture = 0
  GlobalModoSimulado = False
  GlobalPrimeiraVez = True
  GlobalRelatorioPronto = False
  pctCopiaPartedaTelaCPF.Top = -1000
  Get_User_Name
    
  'Se a pasta AppData da aplicacao nao existir, crie-a
  pastaAppData = getSpecialFolder(CSIDL_LOCAL_APPDATA) & "\"
  If Dir(pastaAppData & NomeAplicacao, vbDirectory) = "" Then
    MkDir pastaAppData & NomeAplicacao
  End If
    
  'WTF!!!
  If Dir(GlobalPastadeTrabalho & "\Teste.txt") = "" Then
    Open GlobalPastadeTrabalho & "\Teste.txt" For Output As #1
    Print #1, "This is a test"  ' Print text to file.
    Print #1,   ' Print blank line to file.
    Print #1, "Zone 1"; Tab; "Zone 2"
    Close #1    ' Close file.
  End If
    
  'Testa se o bloco de notas e o aplicativo padrao para abrir arquivos texto (PraQue?)
  sPath = String(255, 32)
  lRet = FindExecutable(GlobalPastadeTrabalho & "\Teste.txt", vbNullString, sPath)
  If InStr(1, UCase(Trim(sPath)), "NOTEPAD.EXE") = 0 Then
    nomeexecut�vel = Trim(sPath)
    pos = InStr(1, UCase(nomeexecut�vel), ".EXE")
    For conta = pos To 1 Step -1
      If Mid(nomeexecut�vel, conta, 1) = "\" Then Exit For
    Next conta
    nomeexecut�vel = Mid(nomeexecut�vel, conta + 1, pos - conta - 1)
    MsgBox "Aten��o: O Windows do seu computador est� configurado para abrir documentos com extens�o '.txt' com o '" & nomeexecut�vel & "'. Favor mudar o programa padr�o para 'Bloco de notas'.", vbCritical, "Abrir '.txt' com... Escolher programa padr�o."
  End If
  If GetSetting("AGENDAMENTODOSABI", "IMPRIMIR", "ORDEM", "") = "HORA" Then
    optOrdem(1).Value = True
  Else
    optOrdem(1).Value = False
  End If

  'Posicao da janela
  Me.Top = Screen.Height - 3760
  Me.Height = 2000
  GlobalT�tulodaTelaAtiva = ""
  GlobalMenuAtualizado = False
  GlobalIDControleOperacional = 0
  GlobalModoImprimeRequerimentos = False
  GlobalEscalaX = 256 / Screen.Width
  GlobalEscalaX = GlobalEscalaX * 256
  GlobalEscalay = 256 / Screen.Height
  GlobalEscalay = GlobalEscalay * 256
  GlobalTempodeEspera = Val(GetSetting("AgendamentosdoSA3BI", "Requerimento", "TempodeEsperadaResposta", ""))
  If GlobalTempodeEspera < 3 Or GlobalTempodeEspera > 10 Then GlobalTempodeEspera = 3
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim lngReturnValue As Long
  
  Me.MousePointer = 9
  Call ReleaseCapture
  lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Private Sub Form_Resize()
  lstMostrarRequerimentos.Top = imageIcone.Top + imageIcone.Height + 40
  Me.Top = Screen.Height - 3760
  If LocalCopiar Then
    Me.Width = 8145
    Me.Left = 600
    lblversao.Top = 0
    lblversao.Left = pctFundo.Width - lblversao.Width - 2600
  Else
    lblversao.Top = Me.Height - 1500
    lblversao.Left = (Me.Width - lblversao.Width) / 2
  End If
  pctFundo.Top = 0
  pctFundo.Left = 0
  pctFundo.Width = Me.Width
  pctFundo.Height = Me.Height
  lstMostrarRequerimentos.Left = 240
  lstMostrarRequerimentos.Width = Me.Width - 580
  lstMostrarRequerimentos.Height = Abs(pctFundo.Height - lstMostrarRequerimentos.Top - 600)
  fraOrdem.Left = 120
  fraOrdem.Top = Me.Height - fraOrdem.Height - 420
  fraImprime.Top = fraOrdem.Top
  fraImprime.Left = 360
  pctCopiaPartedaTela.Left = 0
  pctCopiaPartedaTela.Top = pctFundo.Height + 1000
  lblRelogio.Top = lblRequerimentodoSABI.Top + 40
  lblRelogio.Left = imageIcone.Left + imageIcone.Width + 40
  Picture1.Left = 0
  pctFundoCopias.Top = 3000
  pctFundoCopias.Left = 0
  pctFundoCopias.Width = pctFundo.Width
  pctFundoCopias.Height = Abs(pctFundo.Height - 1350)
  Picture1.Left = 0
  lblRequerimentodoSABI.Top = 120
  lblRequerimentodoSABI.Left = Me.Width / 2 - 1000
  lblRequerimentodoSABI.Width = Me.Width / 2 + 1000
  lblLocaleData.Left = Me.Width / 2 - 2000
  lblLocaleData.Width = Me.Width / 2 + 2000
  lblLocaleData.Top = 80
  lblLocaleData.AutoSize = True
End Sub

Private Sub lblImprimir_Click()
  Dim conta As Long
  Dim Size As RECT
  
  SetForegroundWindow (GlobalIDControleOperacional)
  For conta = 1 To GlobalQuantidadedeRequerimentos
    If GlobalRequerimentos(conta).NIT <> "" Then
      'Abre tela Segunda Via de Marca��o de Exame
      SetForegroundWindow (GlobalIDControleOperacional)
      ClickMenu GlobalIDControleOperacional, 4, 7
      Espera 1000
      SimulaSendKeys "<TAB>"
      Espera 100
      SimulaSendKeys "<TAB>"
      Espera 100
      SimulaSendKeys GlobalRequerimentos(conta).NIT
      Espera 100
      SimulaSendKeys Left$(GlobalRequerimentos(conta).NIT, 1)
      Espera 300
      MsgBox "imprime"
    End If
  Next conta
End Sub

Private Sub lblImprimir_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  lblImprimir.BorderStyle = 1
End Sub

Private Sub lblImprimir_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  lblImprimir.BorderStyle = 0
End Sub

Private Sub lblRelogio_Change()
  lblRelogio.Visible = lblRelogio.Caption <> "00"
  lblRelogio.Left = imageIcone.Left + imageIcone.Width + 40
  pctProgressoFundo.Left = lblRelogio.Left + lblRelogio.Width + 40
End Sub

Private Sub lblRelogio_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim lngReturnValue As Long
  
  Me.MousePointer = 9
  Call ReleaseCapture
  lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  Me.Top = Screen.Height - 760 - 3000
End Sub

Private Sub lblversao_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim lngReturnValue As Long
  
  Me.MousePointer = 9
  Call ReleaseCapture
  lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  Me.Top = Screen.Height - 760 - 3000
End Sub

Private Sub lblversao_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Me.MousePointer = 0
End Sub

Private Sub lstMostrarRequerimentos_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim lngReturnValue As Long
  
  imprimereq = True
  Me.MousePointer = 9
  Call ReleaseCapture
  lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  Me.Top = Screen.Height - 760 - 3000
End Sub

Private Sub lstMostrarRequerimentos_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  On Error Resume Next
  Dim memo As String
  Dim memorequerimento As String
  Dim numerorequerimento As Long
  Dim pos As Long
  
  If GlobalRelatorioPronto Then
    Me.MousePointer = 0
    lstMostrarRequerimentos.ListIndex = ItemUnderMouse(lstMostrarRequerimentos.hWnd, x, y)
    pos = ItemUnderMouse(lstMostrarRequerimentos.hWnd, x, y)
    If pos > 0 Then
      memorequerimento = lstMostrarRequerimentos.List(pos)
      memorequerimento = Trim(Mid(memorequerimento, 4))
      pos = InStr(1, memorequerimento, " ")
      If pos > 0 Then
        memorequerimento = Mid(memorequerimento, 1, pos - 1)
        If GlobalRequerimentoMostrado <> memorequerimento Then
          GlobalRequerimentoMostrado = memorequerimento
          memo = Dir(GlobalPastadeTrabalho & "\*" & memorequerimento & ".bmp")
          If memo <> "" Then
            pctApresentaPartedaTelaCopiada.Picture = LoadPicture(GlobalPastadeTrabalho & "\" & memo)
            pctApresentaPartedaTelaCopiada.Top = 0
            pctApresentaPartedaTelaCopiada.Left = 0
            pctApresentaPartedaTelaCopiada.Visible = True
            pctApresentaPartedaTelaCopiada.ZOrder
            Me.Top = Screen.Height - 760 - 3000
            DoEvents
          Else
            pctApresentaPartedaTelaCopiada.Visible = False
          End If
        End If
      End If
    End If
  End If
End Sub

Private Sub lstMostrarRequerimentos_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  imprimereq = False
End Sub

Private Sub MontaRela��odeRequerimentos()
  On Error Resume Next
  Dim COsize As RECT
  Dim Size As RECT
  Dim titletmp As String
  Dim nret As Long
  Dim TelaSize As RECT
  Dim arquivo As String
  Dim memo As String
  On Error Resume Next

  GlobalSe�ao = Format(Date, "YYYYMMDD") & Format(Time, "hhmmss")
  GlobalNomedoRelatorio = "requerimentos" & GlobalSe�ao
  
  'espera a proxima tela (a tela do crystal report n�o tem nome)
  While GlobalIDTelaImprimirAgendamento = GetForegroundWindow
    Espera 100
    DoEvents
  Wend
  arquivo = esperaCRYSTALREPORTeExporta
End Sub

Private Sub optOrdem_Click(Index As Integer)
  If Index = 1 Then
    SaveSetting "AGENDAMENTODOSABI", "IMPRIMIR", "ORDEM", "HORA"
  Else
    SaveSetting "AGENDAMENTODOSABI", "IMPRIMIR", "ORDEM", "NOME"
  End If
End Sub

Private Sub parabaixo_Click()
  Dim conta As Long

  mtempo2 = 0
  deslocalista = deslocalista - 5
  If deslocalista < 0 Then deslocalista = 0
  lstMostrarRequerimentos.Clear
  For conta = 1 To 50
    If (conta + deslocalista) < GlobalQuantidadedeRequerimentos + 1 Then
      lstMostrarRequerimentos.AddItem Format(conta + deslocalista, "000") & "         " & GlobalRequerimentos(conta + deslocalista).N�mero & "         " & GlobalRequerimentos(conta + deslocalista).Segurado
    End If
  Next conta
  lstMostrarRequerimentos.Height = Screen.Height
  mostralista
End Sub

Private Sub paracima_Click()
  Dim conta As Long

  mtempo2 = 0
  deslocalista = deslocalista + 5
  If deslocalista > GlobalQuantidadedeRequerimentos Then deslocalista = GlobalQuantidadedeRequerimentos
  lstMostrarRequerimentos.Visible = False
  lstMostrarRequerimentos.Clear
  For conta = 1 To 50
    If (conta + deslocalista) < GlobalQuantidadedeRequerimentos + 1 Then
      lstMostrarRequerimentos.AddItem Format(conta + deslocalista, "000") & "         " & GlobalRequerimentos(conta + deslocalista).N�mero & "         " & GlobalRequerimentos(conta + deslocalista).Segurado
    End If
  Next conta
  lstMostrarRequerimentos.Height = Screen.Height
  mostralista
End Sub

Private Sub pctCopiaPartedaTela_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim lngReturnValue As Long

  Me.MousePointer = 9
  Call ReleaseCapture
  lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  If Me.Height = 133 * 15 Or LocalCopiar = True Then Me.Top = Screen.Height - 760 - 3000
End Sub

Private Sub pctFundo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim lngReturnValue As Long
  
  Me.MousePointer = 9
  Call ReleaseCapture
  lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  If Me.Height = 133 * 15 Or LocalCopiar = True Then Me.Top = Screen.Height - 760 - 3000
End Sub

Private Sub pctFundo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Me.MousePointer = 0
End Sub

Private Sub pctCopiaPartedaTela_Change()
  GlobalTipo = "PP/PR"
  GlobalAlerta = False
  pctCopiaPartedaTela.Top = 0
  DoEvents
  If Point(pctCopiaPartedaTela.Left + 790 * 15, pctCopiaPartedaTela.Top + 7 * 15 - 45) = 16777215 Then GlobalAlerta = True
  If Point(pctCopiaPartedaTela.Left + 166 * 15 + 8, pctCopiaPartedaTela.Top + 53 * 15 - 45) <> 0 Then GlobalTipo = "INICIAL"
  pctCopiaPartedaTela.Top = pctFundo.Height + 1000
End Sub

Private Sub pctCopiaPartedaTela_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim lngReturnValue As Long

  Me.MousePointer = 5
  Call ReleaseCapture
  lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  Me.Top = Screen.Height - 760 - 3000
End Sub

Private Sub preparaSABI()
  On Error Resume Next
  Dim titletmp As String
  Dim nret As Long
  Dim localMenu As Long
  Dim Rela��odeMenu As String
  Dim newMenu As Long
  Dim hToolbar20WndClass As Long
  Dim hmsvb_lib_Tollbar As Long
  Dim hPrimeira As Long
  Dim hSegunda As Long
  Dim ClassedoControle As String
  Dim classlength As Long
  Dim IDtelasInternasdoSABI As Long
  Dim IDApagaTela As Long
  Dim contavezes As Long
  Dim IDPrimeiroToolBar As Long
  Dim DimensoesdatelaImprimir As RECT
  Dim res As String
  Dim dimensoesMDIClient As RECT
  Dim IDThunderRT6ComboBoxImprimirAgendamento As Long
  Dim IDMSMaskWndClass1ImprimirAgendamento As Long
  Dim IDMSMaskWndClass2ImprimirAgendamento As Long
  Dim IDToolbar20WndClassImprimirAgendamento As Long
  Dim Dimens�esdoCampoRequerimento As RECT
  Dim CentrodoCampoRequerimento As Long
  Dim contadordeloop As Long
  
  contadordeloop = 0
  
  'limpa todas telas internas do SABI
  'resta agora as telas externas
  verificaeapaga "Imprimir Agendamento"
  verificaeapaga "Imprimir Escala"
  verificaeapaga "Marca��o da Avalia��o Social"
  verificaeapaga "Segunda Via de Carta de Exig�ncia"
  verificaeapaga "Pesquisa de Requerente"
  GlobalhMDIClient = 0
  GlobalhMDIClient = FindWindowEx(GlobalIDControleOperacional, 0, "MDIClient", "")
  If GlobalhMDIClient = 0 Then
    MsgBox "N�o foi encontrada o indentificador da tela de fundo do Controle Operacional.", vbCritical, "Agendamentos do SABI"
    End
  End If
  
  IDtelasInternasdoSABI = 0
  IDtelasInternasdoSABI = FindWindowEx(GlobalhMDIClient, 0, vbNullString, vbNullString)
  While IDtelasInternasdoSABI <> 0
    If IDtelasInternasdoSABI <> 0 Then IDApagaTela = IDtelasInternasdoSABI
    DoEvents
    Espera 300
    IDtelasInternasdoSABI = FindWindowEx(GlobalhMDIClient, IDtelasInternasdoSABI, vbNullString, vbNullString)
    If IDApagaTela <> 0 Then SendMessage IDApagaTela, WM_CLOSE, 0, 0
    Wend
    
    'todas as telas internas foram limpas
    DoEvents
    Espera 300
    
    'acerta a tela
    ColocaTelaControleOperacionanoModoNormal
    ColocaTelaControleOperacionanoModoMaximizado

    'Abre tela Consulta Requerimento/Benef�cio
    ClickMenu GlobalIDControleOperacional, 2, 0
    Espera 3000
    IDtelasInternasdoSABI = 0
    contavezes = 0
    While IDtelasInternasdoSABI = 0
      IDtelasInternasdoSABI = FindWindowEx(GlobalhMDIClient, 0, "ThunderRT6FormDC", "Consulta Requerimento/Benef�cio")
      Espera 300
      DoEvents
      contavezes = contavezes + 1
      If contavezes > 400 Then
        MsgBox "O SABI est� muito lento. Tente outra hora.", vbCritical, "Consulta Requerimento/Benef�cio"
        End
      End If
    Wend
    DoEvents
    GlobalIDTelaConsultaRequerimentoBenef�cio = 0
    While GlobalIDTelaConsultaRequerimentoBenef�cio = 0
      Espera 300
      DoEvents
      GlobalIDTelaConsultaRequerimentoBenef�cio = achaTelaInternaAtiva("Consulta Requerimento/Benef�cio")
    Wend
    IDPrimeiroToolBar = FindWindowEx(IDtelasInternasdoSABI, 0, "Toolbar20WndClass", "")
    GlobalToolbarConsultaRequerimentoOCX = FindWindowEx(IDtelasInternasdoSABI, IDPrimeiroToolBar, "Toolbar20WndClass", "")
    Espera 300
    GlobalToolbarConsultaRequerimento = FindWindowEx(GlobalToolbarConsultaRequerimentoOCX, 0, "msvb_lib_toolbar", vbNullString)
    
    'move a tela de modo a aparecer apenas a borda superior
    res = GetWindowRect(GlobalhMDIClient, dimensoesMDIClient)
    res = GetWindowRect(GlobalIDTelaConsultaRequerimentoBenef�cio, DimensoesdatelaImprimir)
    res = SetWindowPos(GlobalIDTelaConsultaRequerimentoBenef�cio, 0, DimensoesdatelaImprimir.Left, DimensoesdatelaImprimir.Top + (dimensoesMDIClient.Bottom - dimensoesMDIClient.Top - 100), DimensoesdatelaImprimir.Right - DimensoesdatelaImprimir.Left, DimensoesdatelaImprimir.Bottom - DimensoesdatelaImprimir.Top, 0)
    DoEvents

    'abre Imprimir Agendamento
    Espera 300
    ClickMenu GlobalIDControleOperacional, 4, 0
    DoEvents
    Espera 300
    GlobalIDTelaImprimirAgendamento = 0
    contavezes = 0
    While GlobalIDTelaImprimirAgendamento = 0
      GlobalIDTelaImprimirAgendamento = ObtemIDdaTelaPrincipalporTitulo("Imprimir Agendamento")
      Espera 300
      DoEvents
      contavezes = contavezes + 1
      If contavezes > 400 Then
        MsgBox "O SABI est� muito lento. Tente outra hora.", vbCritical, "Imprimir Agendamento"
        End
      End If
    Wend

    'muda o titulo da tela
    SendMessageString GlobalIDTelaImprimirAgendamento, WM_SETTEXT, 0, "Escolha o dia dos agendamentos e clique em Visualizar"
    
    'move a tela para esquerda
    res = GetWindowRect(GlobalIDTelaImprimirAgendamento, DimensoesdatelaImprimir)
    res = SetWindowPos(GlobalIDTelaImprimirAgendamento, 0, DimensoesdatelaImprimir.Left + 100, DimensoesdatelaImprimir.Top, DimensoesdatelaImprimir.Right - DimensoesdatelaImprimir.Left, DimensoesdatelaImprimir.Bottom - DimensoesdatelaImprimir.Top, 0)
    
    'MontaRela��odeRequerimentos
    IDThunderRT6ComboBoxImprimirAgendamento = FindWindowEx(GlobalIDTelaImprimirAgendamento, 0, "ThunderRT6ComboBox", "")
    IDThunderRT6ComboBoxImprimirAgendamento = FindWindowEx(GlobalIDTelaImprimirAgendamento, IDThunderRT6ComboBoxImprimirAgendamento, "ThunderRT6ComboBox", "")
    IDMSMaskWndClass1ImprimirAgendamento = FindWindowEx(GlobalIDTelaImprimirAgendamento, 0, "MSMaskWndClass", "")
    IDMSMaskWndClass2ImprimirAgendamento = FindWindowEx(GlobalIDTelaImprimirAgendamento, IDMSMaskWndClass1ImprimirAgendamento, "MSMaskWndClass", "")
    res = SetWindowPos(GlobalIDTelaImprimirAgendamento, 0, DimensoesdatelaImprimir.Left + 120, DimensoesdatelaImprimir.Top, DimensoesdatelaImprimir.Right - DimensoesdatelaImprimir.Left - 60, DimensoesdatelaImprimir.Bottom - DimensoesdatelaImprimir.Top, 0)
    Do While GlobalIDTelaImprimirAgendamento > 0
      GlobalAgenciaEscolhida = ObtemTextodoControle(IDThunderRT6ComboBoxImprimirAgendamento)
      res = ObtemTextodoControle(IDMSMaskWndClass1ImprimirAgendamento)
      GlobalDataEscolhida = "01/01/1900"
      If Len(res) = 10 And res <> "  /  /" Then
        If IsDate(res) Then
          GlobalDataEscolhida = res
          If GlobalDataEscolhida >= Date And GlobalDataEscolhida < DateAdd("d", Date, 180) Then
            '------------------data final igual a data inicial
            res = ObtemTextodoControle(IDMSMaskWndClass2ImprimirAgendamento)
            If Len(res) = 10 And res <> "  /  /    " Then
              GlobalDataEscolhida2 = res
            Else
              GlobalDataEscolhida2 = "01/01/1900"
            End If
            If GlobalDataEscolhida2 <> GlobalDataEscolhida Then
              res = GetWindowRect(IDMSMaskWndClass2ImprimirAgendamento, Dimens�esdoCampoRequerimento)
              CentrodoCampoRequerimento = convlong(Dimens�esdoCampoRequerimento.Left + (Dimens�esdoCampoRequerimento.Right - Dimens�esdoCampoRequerimento.Left) / 2, Dimens�esdoCampoRequerimento.Top + (Dimens�esdoCampoRequerimento.Bottom - Dimens�esdoCampoRequerimento.Top) / 2)
              SendMessage IDMSMaskWndClass2ImprimirAgendamento, WM_LBUTTONDOWN, MK_LBUTTON, CentrodoCampoRequerimento
              SendMessage IDMSMaskWndClass2ImprimirAgendamento, WM_LBUTTONUP, MK_LBUTTON, CentrodoCampoRequerimento
              DoEvents
              Espera 100
              SimulaSendKeys "2"
              Espera 200
              SendMessage IDMSMaskWndClass2ImprimirAgendamento, WM_SETTEXT, 0, Format(GlobalDataEscolhida, "dd/mm/yyyy") & Chr$(0)
              DoEvents
              Espera 600
            End If
          '------------------
          Else
            GlobalDataEscolhida = "01/01/1900"
          End If
        End If
      End If
      If GlobalDataEscolhida = "01/01/1900" Then
        SendMessage IDMSMaskWndClass2ImprimirAgendamento, WM_SETTEXT, 0, "  /  /    " & Chr$(0)
        GlobalDataEscolhida2 = "01/01/1900"
      End If
      If GlobalAgenciaEscolhida <> "" And GlobalDataEscolhida <> "01/01/1900" And GlobalDataEscolhida2 <> "01/01/1900" And GlobalDataEscolhida = GlobalDataEscolhida2 Then
        res = SetWindowPos(GlobalIDTelaImprimirAgendamento, 0, DimensoesdatelaImprimir.Left + 120, DimensoesdatelaImprimir.Top, DimensoesdatelaImprimir.Right - DimensoesdatelaImprimir.Left, DimensoesdatelaImprimir.Bottom - DimensoesdatelaImprimir.Top, 0)
        Exit Do
      Else
        res = SetWindowPos(GlobalIDTelaImprimirAgendamento, 0, DimensoesdatelaImprimir.Left + 120, DimensoesdatelaImprimir.Top, DimensoesdatelaImprimir.Right - DimensoesdatelaImprimir.Left - 60, DimensoesdatelaImprimir.Bottom - DimensoesdatelaImprimir.Top, 0)
    End If
    contadordeloop = contadordeloop + 1
    If contadordeloop > 1000 Then
      lblRequerimentodoSABI.Caption = "Tempo esgotado para informar a data do agendamento"
      DoEvents
      Beep
      Espera 6000
      End
    End If
    Espera 100
    GlobalIDTelaImprimirAgendamento = ObtemIDdaTelaPrincipalporTitulo("Escolha o dia dos agendamentos e clique em Visualizar")
  Loop
  MontaRela��odeRequerimentos
End Sub

Private Sub decodeRequerimentos(texto As String)
  On Error Resume Next
  Dim memo As String
  Dim posmedico As Long
  Dim posRequerimento As Long
  Dim posproximapericia As Long
  Dim posdata As Long
  Dim posfimdata As Long
  Dim datamemo As String
  Dim pos As Long
  Dim FileNumber  As Long
  Dim mTexto, mLinha As String
  Dim fimcabecalho As Boolean
  Dim pos1, pos2, pos3, pos4, contador As Long
  Dim memonome As String
  Dim posnome As Long
  Dim contas As Long
  Dim conqta As Long
    
  contador = 0
  FileNumber = FreeFile
  fimcabecalho = False
  GlobalAgendamentosConsultaCabecalho = ""
  Open GlobalAreadeTrabalho & "\Agendamentos.txt" For Input As #FileNumber
  Do While Not EOF(FileNumber)
    Line Input #FileNumber, mLinha
    If InStr(1, mLinha, "Medico") Then fimcabecalho = True
    If fimcabecalho = False Then
      GlobalAgendamentosConsultaCabecalho = GlobalAgendamentosConsultaCabecalho & mLinha & Chr(13) & Chr(10)
    Else
      If InStr(1, mLinha, "Medico") Or InStr(1, mLinha, "Hor�rio") Then
      Else
        pos1 = InStr(1, mLinha, Chr(9))
        If pos1 > 0 Then
          contador = contador + 1
          GlobalAgendamentosQuandidade = contador
          GlobalAgendamentosConsulta(contador).Horario = Mid(mLinha, 1, pos1 - 1)
          pos2 = InStr(pos1 + 1, mLinha, Chr(9))
          If pos2 > 0 Then
            memonome = Mid(mLinha, pos1 + 2, pos2 - pos1 - 1)
            posnome = InStr(1, memonome, Chr(34))
            If posnome > 0 Then
              memonome = Mid(memonome, 1, posnome - 1)
            Else
              memonome = Mid(mLinha, pos1 + 2, pos2 - pos1 - 1)
            End If
            GlobalAgendamentosConsulta(contador).Segurado = memonome
            pos3 = InStr(pos2 + 1, mLinha, Chr(9))
            If pos3 > 0 Then
              GlobalAgendamentosConsulta(contador).Concluida = Mid(mLinha, pos2 + 1, pos3 - pos2 - 1)
              pos4 = InStr(pos3 + 1, mLinha, Chr(9))
              If pos4 > 0 Then
                GlobalAgendamentosConsulta(contador).Ordem = Mid(mLinha, pos3 + 1, pos4 - pos3 - 1)
                GlobalAgendamentosConsulta(contador).Requerimento = Val(Mid(mLinha, pos4 + 1))
              End If
            End If
          End If
        End If
      End If
    End If
  Loop
  Close #FileNumber
  If Dir(GlobalAreadeTrabalho & "\Agendamentos.txt") <> "" Then
    Kill GlobalAreadeTrabalho & "\Agendamentos.txt"
  End If
  pos1 = InStr(1, GlobalAgendamentosConsultaCabecalho, "Local:")
  If pos1 > 0 Then
    mTexto = Mid(GlobalAgendamentosConsultaCabecalho, pos1 + 8)
    For pos2 = 1 To Len(mTexto)
      If Mid(mTexto, pos2, 1) = Chr(13) Or Mid(mTexto, pos2, 1) = Chr(10) Then
        mTexto = Mid(mTexto, 1, pos2 - 1) & Mid(mTexto, pos2 + 1)
      End If
      If Asc(Mid(mTexto, pos2, 1)) = 34 Then
        mTexto = Mid(mTexto, 1, pos2 - 1) & " " & Mid(mTexto, pos2 + 1)
      End If
    Next pos2
    mTexto = Trim(mTexto)
    lblLocaleData.Caption = mTexto
    lblLocaleData.Visible = True
    lblRequerimentodoSABI.Visible = False
  End If
  MontaListadeRequerimentos (GlobalAgendamentosConsultaCabecalho)
End Sub

Private Sub pctFundoCopias_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim lngReturnValue As Long
  
  Me.MousePointer = 9
  Call ReleaseCapture
  lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  Me.Top = Screen.Height - 760 - 3000
End Sub

Private Sub pctFundoCopias_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Me.MousePointer = 0
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  On Error Resume Next
  
  Clipboard.Clear
  Clipboard.SetData Picture1.Picture
  LocalY = y
  If Button = 2 Then
    res = SetWindowPos(Me.hWnd, -2, 0, 0, 0, 0, 3)
    MsgBox "Esta rela��o de requerimentos foi salva na �rea de transfer�ncia. Para imprimir cole agora no 'Word' ou 'Paint'.", vbCritical, "Imprimir Rela��o de Requerimentos"
    res = SetWindowPos(Me.hWnd, -1, 0, 0, 0, 0, 3)
  End If
End Sub

Private Sub Timer1_Timer()
  On Error Resume Next
  
  mtempo1 = mtempo1 + 1
  If mtempo1 > 60 Then End
End Sub

Private Sub Timer2_Timer()
  On Error Resume Next
  
  mtempo2 = mtempo2 + 1
  If mtempo2 > 60 Then End
End Sub

Private Sub txttPrimeiro_Change()
  txttPrimeiro.Text = Val(txttPrimeiro.Text)
  If Val(txttPrimeiro.Text) < 1 Then txttPrimeiro.Text = 1
  If Val(txttPrimeiro.Text) > GlobalQuantidadedeRequerimentos Then txttPrimeiro.Text = GlobalQuantidadedeRequerimentos
  cmdContinua.Enabled = Val(txtUltimo.Text) >= Val(txttPrimeiro.Text)
End Sub

Private Sub txtUltimo_Change()
  txtUltimo.Text = Val(txtUltimo.Text)
  If Val(txtUltimo.Text) < 1 Then txtUltimo.Text = 1
  If Val(txtUltimo.Text) > GlobalQuantidadedeRequerimentos Then txtUltimo.Text = GlobalQuantidadedeRequerimentos
  cmdContinua.Enabled = Val(txtUltimo.Text) >= Val(txttPrimeiro.Text)
End Sub
