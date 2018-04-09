VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form consultarparte 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consultar Parte de Admisión"
   ClientHeight    =   12000
   ClientLeft      =   45
   ClientTop       =   435
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
   MinButton       =   0   'False
   ScaleHeight     =   800
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   950
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "AD"
      Enabled         =   0   'False
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
      Left            =   5760
      TabIndex        =   36
      Top             =   9600
      Width           =   735
   End
   Begin VB.CheckBox Check2 
      Caption         =   "MP"
      Enabled         =   0   'False
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
      Left            =   6600
      TabIndex        =   35
      Top             =   9600
      Width           =   735
   End
   Begin VB.CheckBox Check3 
      Caption         =   "PC"
      Enabled         =   0   'False
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
      Left            =   7440
      TabIndex        =   34
      Top             =   9600
      Width           =   735
   End
   Begin VB.TextBox Text13 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7920
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   10200
      Width           =   1455
   End
   Begin VB.TextBox Text12 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5400
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   10200
      Width           =   1575
   End
   Begin VB.TextBox Text11 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   12000
      TabIndex        =   30
      TabStop         =   0   'False
      Text            =   "1"
      Top             =   360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   "C:\Padmin\BD\padminbd.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "añoactivo"
      Top             =   480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   "C:\Padmin\BD\padminbd.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "partes"
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text10 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6120
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   5400
      Width           =   615
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\Padmin\BD\padminbd.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "cliente"
      Top             =   480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Padmin\BD\padminbd.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "numeracionpartes"
      Top             =   120
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Imprimir"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   11160
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Borrar Formulario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   11160
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   11160
      Width           =   1455
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   9000
      Width           =   975
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7560
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   9000
      Width           =   495
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4920
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   9000
      Width           =   495
   End
   Begin VB.Frame Frame2 
      Caption         =   "Entrada / Salida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   19
      Top             =   8280
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   18
      Top             =   6000
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9000
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   7680
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4200
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   7680
      Width           =   3495
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7920
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   7080
      Width           =   2895
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4200
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   7080
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4200
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6480
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3960
      TabIndex        =   1
      Top             =   5400
      Width           =   975
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
      Picture         =   "consultarparte.frx":0000
      ScaleHeight     =   1395
      ScaleWidth      =   11115
      TabIndex        =   10
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
      StartOfWeek     =   50790402
      CurrentDate     =   39848
   End
   Begin VB.Label Label14 
      Caption         =   "copias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11280
      TabIndex        =   31
      Top             =   360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label13 
      Caption         =   "valido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   29
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label12 
      Caption         =   "Salida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   24
      Top             =   10320
      Width           =   615
   End
   Begin VB.Label Label11 
      Caption         =   "Entrada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   23
      Top             =   10320
      Width           =   615
   End
   Begin VB.Label Label10 
      Caption         =   "Precio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   22
      Top             =   9000
      Width           =   495
   End
   Begin VB.Label Label9 
      Caption         =   "Número de personas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   21
      Top             =   9000
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "Habitación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   20
      Top             =   9000
      Width           =   855
   End
   Begin VB.Line Line10 
      X1              =   736
      X2              =   208
      Y1              =   568
      Y2              =   568
   End
   Begin VB.Line Line9 
      X1              =   736
      X2              =   736
      Y1              =   712
      Y2              =   568
   End
   Begin VB.Line Line8 
      X1              =   208
      X2              =   208
      Y1              =   712
      Y2              =   568
   End
   Begin VB.Line Line7 
      X1              =   208
      X2              =   736
      Y1              =   712
      Y2              =   712
   End
   Begin VB.Line Line6 
      X1              =   736
      X2              =   208
      Y1              =   416
      Y2              =   416
   End
   Begin VB.Line Line5 
      X1              =   736
      X2              =   736
      Y1              =   544
      Y2              =   416
   End
   Begin VB.Line Line4 
      X1              =   208
      X2              =   208
      Y1              =   544
      Y2              =   416
   End
   Begin VB.Line Line2 
      X1              =   208
      X2              =   736
      Y1              =   544
      Y2              =   544
   End
   Begin VB.Label Label7 
      Caption         =   "Provincia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8160
      TabIndex        =   17
      Top             =   7680
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Localidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   16
      Top             =   7680
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Apellidos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   15
      Top             =   7080
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   14
      Top             =   7080
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "D.N.I."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   13
      Top             =   6480
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Número de Parte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   12
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Line Line3 
      X1              =   944
      X2              =   8
      Y1              =   336
      Y2              =   336
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  CONSULTAR PARTE DE ADMISION"
      Height          =   495
      Left            =   4800
      TabIndex        =   11
      Top             =   4200
      Width           =   5055
   End
   Begin VB.Line Line1 
      X1              =   944
      X2              =   8
      Y1              =   256
      Y2              =   256
   End
End
Attribute VB_Name = "consultarparte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload consultarparte
    
End Sub

Private Sub Command2_Click()
    'actualizamos la numeracion de partes
    Data1.Recordset.MoveFirst
    Data1.Recordset.Edit
    Data1.Recordset("numero") = Val(Text1.Text) + 1
    Data1.UpdateRecord
    
    'Comprobamos si hay un nuevo cliente, y si es asi lo metemos en la base de datos
    If Text10.Text = "si" Then
        
        If Text2.Text = "" Then
            Text2.Text = "-"
        End If
        
        If Text3.Text = "" Then
            Text3.Text = "-"
        End If
        
        If Text4.Text = "" Then
            Text4.Text = "-"
        End If
        
        If Text5.Text = "" Then
            Text5.Text = "-"
        End If
        
        If Text6.Text = "" Then
            Text6.Text = "-"
        End If
        
        Data2.Recordset.AddNew
        Data2.Recordset("dni") = Text2.Text
        Data2.Recordset("nombre") = Text3.Text
        Data2.Recordset("apellidos") = Text4.Text
        Data2.Recordset("localidad") = Text5.Text
        Data2.Recordset("provincia") = Text6.Text
        Data2.Recordset.Update
        
        
    End If
        
    'introducimos el nuevo parte
    If Text2.Text = "" Then
        Text2.Text = "-"
    End If
        
    If Text3.Text = "" Then
        Text3.Text = "-"
    End If
     
    If Text4.Text = "" Then
        Text4.Text = "-"
    End If
        
    If Text5.Text = "" Then
        Text5.Text = "-"
    End If
       
    If Text6.Text = "" Then
        Text6.Text = "-"
    End If
       
    If Text7.Text = "" Then
        Text7.Text = "-"
    End If
       
    If Text8.Text = "" Then
        Text8.Text = "-"
    End If
       
    If Text9.Text = "" Then
        Text9.Text = "-"
    End If
    
    Data4.Recordset.MoveFirst
    Data3.Recordset.AddNew
    Data3.Recordset("año") = Data4.Recordset("año")
    Data3.Recordset("nparte") = Text1.Text
    Data3.Recordset("dni") = Text2.Text
    Data3.Recordset("nombre") = Text3.Text
    Data3.Recordset("apellidos") = Text4.Text
    Data3.Recordset("localidad") = Text5.Text
    Data3.Recordset("provincia") = Text6.Text
    Data3.Recordset("habitacion") = Text7.Text
    Data3.Recordset("personas") = Text8.Text
    Data3.Recordset("precio") = Text9.Text
    Data3.Recordset("entrada") = Format$(DTPicker1.Value, "dd,mm,yyyy")
    Data3.Recordset("salida") = Format$(DTPicker2.Value, "dd,mm,yyyy")
    Data3.Recordset("valido") = "si"
    Data3.Recordset.Update
        
    'borramos el formulario
    Command4.Enabled = False
    Command2.Enabled = False
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = "si"
    DTPicker1.Value = Format$(Now, "dd/mm/yyyy")
    DTPicker2.Value = Format$(Now, "dd/mm/yyyy")
    Text2.SetFocus
    
    
End Sub

Private Sub Command3_Click()
    Command4.Enabled = False
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Check1.Value = 0
    Check2.Value = 0
    Check3.Value = 0
    Text1.SetFocus
    
End Sub

Private Sub Command4_Click()

    Impconsultarparte.Show
End Sub


Private Sub cerrar_Click()
    End
    
End Sub

Private Sub Command10_Click()
    End
    
End Sub



Private Sub Text1_KeyPress(KeyAscii As Integer)
    Dim encontrado As Integer
    
    If KeyAscii = 13 Then
        'buscamos el numero de registro en la base de datos
        
        Data4.Recordset.MoveFirst
        Data3.Recordset.MoveFirst
        encontrado = 0
        Do While encontrado = 0 And Not Data3.Recordset.EOF
            If Data3.Recordset("nparte") = Text1.Text Then
                If Data3.Recordset("año") = Data4.Recordset("año") Then
                    encontrado = 1
                    Text2.Text = Data3.Recordset("dni")
                    Text3.Text = Data3.Recordset("nombre")
                    Text4.Text = Data3.Recordset("apellidos")
                    Text5.Text = Data3.Recordset("localidad")
                    Text6.Text = Data3.Recordset("provincia")
                    Text7.Text = Data3.Recordset("habitacion")
                    Text8.Text = Data3.Recordset("personas")
                    Text9.Text = Data3.Recordset("precio")
                    Text10.Text = Data3.Recordset("valido")
                    Text12.Text = Format$(Data3.Recordset("entrada"), "dd/mm/yyyy")
                    Text13.Text = Format$(Data3.Recordset("salida"), "dd/mm/yyyy")
                    
                    If UCase$(Data3.Recordset("regimen")) = "AD" Then
                        Check1.Value = 1
                        Check2.Value = 0
                        Check3.Value = 0
                    End If
                    
                    If UCase$(Data3.Recordset("regimen")) = "MP" Then
                        Check1.Value = 0
                        Check2.Value = 1
                        Check3.Value = 0
                    End If
                    
                    If UCase$(Data3.Recordset("regimen")) = "PC" Then
                        Check1.Value = 0
                        Check2.Value = 0
                        Check3.Value = 1
                    End If
                    
                    If UCase$(Data3.Recordset("regimen")) = "N" Then
                        Check1.Value = 0
                        Check2.Value = 0
                        Check3.Value = 0
                    End If
                    
                    Command4.Enabled = True
                    
                End If
            End If
            Data3.Recordset.MoveNext
        Loop
        
        ' si no se ha encontrado, borramos el campo numero de registro
        If encontrado = 0 Then
            Text1.Text = ""
            Text2.Text = ""
            Text3.Text = ""
            Text4.Text = ""
            Text5.Text = ""
            Text6.Text = ""
            Text7.Text = ""
            Text8.Text = ""
            Text9.Text = ""
            Text10.Text = ""
            Text12.Text = ""
            Text13.Text = ""
            Command4.Enabled = False
            
        End If
    End If
End Sub
