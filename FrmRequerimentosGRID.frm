VERSION 5.00
Begin VB.Form frmVer 
   Caption         =   "Requerimentos"
   ClientHeight    =   6120
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   7035
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pctFundo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   120
      ScaleHeight     =   5625
      ScaleWidth      =   6345
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.PictureBox pctFundoRequerimento 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   240
         ScaleHeight     =   2865
         ScaleWidth      =   4785
         TabIndex        =   1
         Top             =   0
         Width           =   4815
         Begin VB.PictureBox pctRequerimento 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   0
            Left            =   0
            ScaleHeight     =   375
            ScaleWidth      =   975
            TabIndex        =   2
            Top             =   0
            Visible         =   0   'False
            Width           =   975
         End
      End
      Begin VB.Label cmdCima 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         Caption         =   "  /\  "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   300
         Left            =   5880
         TabIndex        =   4
         Top             =   120
         Width           =   300
      End
      Begin VB.Label cmdBaixo 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         Caption         =   "  \/  "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   300
         Left            =   5880
         TabIndex        =   3
         Top             =   5160
         Width           =   300
      End
   End
End
Attribute VB_Name = "frmVer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBaixo_Click()
    On Error Resume Next
    pctFundoRequerimento.Top = pctFundoRequerimento.Top - 10 * pctRequerimento(1).Height
End Sub

Private Sub cmdBaixo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next

    cmdBaixo.BorderStyle = 1
End Sub

Private Sub cmdBaixo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next

    cmdBaixo.BorderStyle = 0
End Sub

Private Sub cmdCima_Click()
    On Error Resume Next

    pctFundoRequerimento.Top = pctFundoRequerimento.Top + 10 * pctRequerimento(1).Height
End Sub

Private Sub cmdCima_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next

    cmdCima.BorderStyle = 1
End Sub

Private Sub cmdCima_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next

    cmdCima.BorderStyle = 0
End Sub

Private Sub Form_Load()
    On Error Resume Next

    Dim memo As String
    Dim contador As Long
    On Error Resume Next
    res = SetWindowPos(Me.hWnd, -1, 0, 0, 0, 0, 3)
    Me.Height = Screen.Height - 2000
    contador = 0
    pasta = GlobalPastadeTrabalho
    memo = Dir(GlobalPastadeTrabalho & "\" & GlobalDatadosRequerimentos & "*.bmp")
    While memo <> ""
        contador = contador + 1
        
        Load pctRequerimento(contador)
        pctRequerimento(contador).Picture = LoadPicture(GlobalPastadeTrabalho & "\" & memo)
        pctRequerimento(contador).Top = contador * pctRequerimento(contador).Height
        pctRequerimento(contador).Visible = True
        pctFundoRequerimento.Height = pctRequerimento(contador).Top + pctRequerimento(contador).Height
        pctFundoRequerimento.Width = pctRequerimento(contador).Width
        pctFundo.Width = pctFundoRequerimento.Width + 360 + 500
        Me.Width = pctFundo.Width + 400
        memo = Dir()
        cmdCima.Left = pctFundo.Width - 400 - 40
        cmdBaixo.Left = pctFundo.Width - 400 - 40
    Wend
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    pctFundo.Height = Me.Height - 760
    cmdBaixo.Top = Me.Height - 1400
End Sub
