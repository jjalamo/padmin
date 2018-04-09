VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form inicio 
   Caption         =   "Menú Principal"
   ClientHeight    =   12000
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   14250
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   800
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   950
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command10 
      Caption         =   "SALIR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   9960
      Width           =   3135
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Modificar Parte de Admisión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   7560
      Width           =   2895
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Anular Parte de Admisión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   8400
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Consultar Parte de Admisión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6720
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nuevo Parte de Admisión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5880
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      Picture         =   "inicio.frx":0000
      ScaleHeight     =   1395
      ScaleWidth      =   11115
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1320
      Width           =   11175
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   11520
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1320
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StartOfWeek     =   50855938
      CurrentDate     =   39848
   End
   Begin VB.Line Line3 
      X1              =   944
      X2              =   8
      Y1              =   336
      Y2              =   336
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " MENU PRINCIPAL"
      Height          =   495
      Left            =   5760
      TabIndex        =   2
      Top             =   4200
      Width           =   2655
   End
   Begin VB.Line Line1 
      X1              =   944
      X2              =   8
      Y1              =   256
      Y2              =   256
   End
   Begin VB.Menu PA 
      Caption         =   "PARTES DE ADMISION"
      Begin VB.Menu npa 
         Caption         =   "Nuevo Parte de Admisión"
         Index           =   1
      End
      Begin VB.Menu npa 
         Caption         =   "Consultar Parte de Admisión"
         Index           =   2
      End
      Begin VB.Menu mpa 
         Caption         =   "Modificar Parte de Admisión"
      End
      Begin VB.Menu apa 
         Caption         =   "Anular Parte de Admisión"
      End
   End
   Begin VB.Menu blanco1 
      Caption         =   ""
   End
   Begin VB.Menu lpa 
      Caption         =   "LISTADOS DE PARTES DE ADMISION"
      Begin VB.Menu lnp 
         Caption         =   "Listado por Número de Parte"
      End
      Begin VB.Menu lfe 
         Caption         =   "Listado por Fecha de Entrada"
      End
      Begin VB.Menu lpc 
         Caption         =   "Listado por Clientes"
      End
      Begin VB.Menu lpn 
         Caption         =   "Listado de Partes Nulos"
      End
   End
   Begin VB.Menu blanco2 
      Caption         =   ""
   End
   Begin VB.Menu cerrar 
      Caption         =   "SALIR"
   End
End
Attribute VB_Name = "inicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DTPicker1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        DTPicker2.SetFocus
    End If
    
End Sub

Private Sub apa_Click()
    anularparte.Show
    
End Sub

Private Sub cerrar_Click()
    End
    
End Sub

Private Sub Command1_Click()
    nuevoparte.Show
    
End Sub

Private Sub Command10_Click()
    End
    
End Sub

Private Sub Command2_Click()
    consultarparte.Show
End Sub

Private Sub Command3_Click()
    listadonparte.Show
    
End Sub



Private Sub Command8_Click()
    anularparte.Show
    
End Sub

Private Sub Command9_Click()
    modificarparte.Show
End Sub

Private Sub lfe_Click()
    listadofecha.Show
    
End Sub

Private Sub lnp_Click()
    listadonparte.Show
    
End Sub

Private Sub lpc_Click()
    listadocliente.Show
End Sub

Private Sub lpn_Click()
    listadopartenulo.Show
    
End Sub

Private Sub mpa_Click()
    modificarparte.Show
End Sub

Private Sub npa_Click(Index As Integer)
    consultarparte.Show
    
End Sub
