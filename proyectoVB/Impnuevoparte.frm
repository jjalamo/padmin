VERSION 5.00
Begin VB.Form Impnuevoparte 
   BackColor       =   &H80000009&
   Caption         =   "PARTE DE ADMISION"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   11010
   ScaleWidth      =   8355
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000009&
      Caption         =   "AD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   36
      Top             =   8400
      Width           =   735
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H80000009&
      Caption         =   "MP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   35
      Top             =   8400
      Width           =   735
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H80000009&
      Caption         =   "PC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   34
      Top             =   8400
      Width           =   735
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H80000009&
      Height          =   360
      Left            =   6120
      TabIndex        =   27
      Top             =   7800
      Width           =   975
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H80000009&
      Height          =   360
      Left            =   4560
      TabIndex        =   26
      Top             =   7800
      Width           =   495
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H80000009&
      Height          =   360
      Left            =   1920
      TabIndex        =   25
      Top             =   7800
      Width           =   495
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000009&
      Caption         =   "Entrada / Salida"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   7080
      Width           =   1695
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H80000009&
      Height          =   360
      Left            =   5040
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   9000
      Width           =   1215
   End
   Begin VB.TextBox Text13 
      BackColor       =   &H80000009&
      Height          =   360
      Left            =   2640
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   9000
      Width           =   1215
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H80000009&
      Height          =   360
      Left            =   1440
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H80000009&
      Height          =   360
      Left            =   1320
      TabIndex        =   13
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000009&
      Height          =   360
      Left            =   1320
      TabIndex        =   12
      Top             =   5880
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000009&
      Height          =   360
      Left            =   5040
      TabIndex        =   11
      Top             =   5880
      Width           =   2895
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H80000009&
      Height          =   360
      Left            =   1320
      TabIndex        =   10
      Top             =   6480
      Width           =   3495
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H80000009&
      Height          =   360
      Left            =   6120
      TabIndex        =   9
      Top             =   6480
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   "Cliente"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4800
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   2
      Text            =   "N.I.R.T.A.:  H / JA / 00654"
      Top             =   2640
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   1
      Text            =   "    PARTE DE ADMISIÓN     "
      Top             =   240
      Width           =   4455
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      Height          =   1215
      Left            =   360
      Picture         =   "Impnuevoparte.frx":0000
      ScaleHeight     =   1155
      ScaleWidth      =   7515
      TabIndex        =   0
      Top             =   960
      Width           =   7575
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000009&
      Caption         =   "Salida"
      Height          =   255
      Left            =   4440
      TabIndex        =   33
      Top             =   9000
      Width           =   615
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000009&
      Caption         =   "Entrada"
      Height          =   255
      Left            =   1920
      TabIndex        =   32
      Top             =   9000
      Width           =   615
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000009&
      Caption         =   "Precio"
      Height          =   255
      Left            =   5520
      TabIndex        =   31
      Top             =   7800
      Width           =   495
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000009&
      Caption         =   "Número de personas"
      Height          =   255
      Left            =   2880
      TabIndex        =   30
      Top             =   7800
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000009&
      Caption         =   "Habitación"
      Height          =   255
      Left            =   960
      TabIndex        =   29
      Top             =   7800
      Width           =   855
   End
   Begin VB.Line Line10 
      X1              =   8160
      X2              =   120
      Y1              =   7320
      Y2              =   7320
   End
   Begin VB.Line Line8 
      X1              =   120
      X2              =   120
      Y1              =   9480
      Y2              =   7320
   End
   Begin VB.Line Line7 
      X1              =   120
      X2              =   8160
      Y1              =   9480
      Y2              =   9480
   End
   Begin VB.Line Line15 
      X1              =   8160
      X2              =   8160
      Y1              =   7320
      Y2              =   9480
   End
   Begin VB.Label Label18 
      BackColor       =   &H80000009&
      Caption         =   "€"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   28
      Top             =   7800
      Width           =   375
   End
   Begin VB.Line Line16 
      X1              =   0
      X2              =   8400
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label Label17 
      BackColor       =   &H80000009&
      Caption         =   "Firmado:______________________________"
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
      Left            =   240
      TabIndex        =   21
      Top             =   10200
      Width           =   3615
   End
   Begin VB.Line Line14 
      X1              =   8160
      X2              =   8160
      Y1              =   5040
      Y2              =   6960
   End
   Begin VB.Label Label16 
      BackColor       =   &H80000009&
      Caption         =   "Número de Parte"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label15 
      BackColor       =   &H80000009&
      Caption         =   "D.N.I."
      Height          =   255
      Left            =   720
      TabIndex        =   19
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label Label14 
      BackColor       =   &H80000009&
      Caption         =   "Nombre"
      Height          =   255
      Left            =   600
      TabIndex        =   18
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000009&
      Caption         =   "Apellidos"
      Height          =   255
      Left            =   4320
      TabIndex        =   17
      Top             =   5880
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000009&
      Caption         =   "Localidad"
      Height          =   255
      Left            =   480
      TabIndex        =   16
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000009&
      Caption         =   "Provincia"
      Height          =   255
      Left            =   5280
      TabIndex        =   15
      Top             =   6480
      Width           =   735
   End
   Begin VB.Line Line13 
      X1              =   120
      X2              =   8160
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Line Line12 
      X1              =   120
      X2              =   120
      Y1              =   6960
      Y2              =   5040
   End
   Begin VB.Line Line11 
      X1              =   8520
      X2              =   8520
      Y1              =   6720
      Y2              =   4800
   End
   Begin VB.Line Line6 
      X1              =   8160
      X2              =   240
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line9 
      X1              =   8520
      X2              =   8520
      Y1              =   9240
      Y2              =   7320
   End
   Begin VB.Line Line5 
      X1              =   0
      X2              =   8400
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line4 
      X1              =   4200
      X2              =   360
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line3 
      X1              =   4200
      X2              =   4200
      Y1              =   2280
      Y2              =   3480
   End
   Begin VB.Line Line2 
      X1              =   360
      X2              =   360
      Y1              =   2280
      Y2              =   3480
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   4200
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000009&
      Caption         =   "hotel@reysanchocuarto.com"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000009&
      Caption         =   "www.reysanchocuarto.com"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000009&
      Caption         =   "Tels. 953 402 301  -  953 402 732"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      Caption         =   "Santisteban Del Puerto, Jaén."
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "Avda, Andalucía 10."
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   2400
      Width           =   1575
   End
End
Attribute VB_Name = "Impnuevoparte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    Dim copias As String
    Dim ncopias As Integer
    Dim i As Integer
    
    Text11.Text = nuevoparte!Text1.Text
    Text10.Text = nuevoparte!Text2.Text
    Text3.Text = nuevoparte!Text3.Text
    Text4.Text = nuevoparte!Text4.Text
    Text5.Text = nuevoparte!Text5.Text
    Text6.Text = nuevoparte!Text6.Text
    Text7.Text = nuevoparte!Text7.Text
    Text8.Text = nuevoparte!Text8.Text
    Text9.Text = nuevoparte!Text9.Text
    Text13.Text = Format$(nuevoparte!DTPicker1.Value, "dd/mm/yyyy")
    Text12.Text = Format$(nuevoparte!DTPicker2.Value, "dd/mm/yyyy")
    Check1.Value = nuevoparte!Check1.Value
    Check2.Value = nuevoparte!Check2.Value
    Check3.Value = nuevoparte!Check3.Value
    ncopias = -1
    Do While ncopias < 0
        copias = InputBox$("Introduce el número de copias")
        ncopias = Val(copias)
    Loop
    
    For i = 1 To ncopias
        PrintForm
    Next i
    
    'borramos el formulario
    nuevoparte!Command4.Enabled = False
    nuevoparte!Command2.Enabled = False
    nuevoparte!Text1.Text = ""
    nuevoparte!Text2.Text = ""
    nuevoparte!Text3.Text = ""
    nuevoparte!Text4.Text = ""
    nuevoparte!Text5.Text = ""
    nuevoparte!Text6.Text = ""
    nuevoparte!Text7.Text = ""
    nuevoparte!Text8.Text = ""
    nuevoparte!Text9.Text = ""
    nuevoparte!Text10.Text = "si"
    nuevoparte!DTPicker1.Value = Format$(Now, "dd/mm/yyyy")
    nuevoparte!DTPicker2.Value = Format$(Now, "dd/mm/yyyy")
    nuevoparte!Check1.Value = 0
    nuevoparte!Check2.Value = 0
    nuevoparte!Check3.Value = 0
    nuevoparte!Text2.SetFocus
   
End Sub

