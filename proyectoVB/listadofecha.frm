VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{13592B48-653C-491D-ACB1-C3140AA12F33}#6.1#0"; "ubGrid.ocx"
Begin VB.Form listadofecha 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado por por Fecha de Entrada"
   ClientHeight    =   12000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   800
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   950
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   3360
      TabIndex        =   12
      Top             =   240
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   51314689
      CurrentDate     =   39854
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   4680
      TabIndex        =   10
      Top             =   5160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   51314689
      CurrentDate     =   39854
   End
   Begin ubGridControl.ubGrid ubGrid1 
      Height          =   4815
      Left            =   240
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5880
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   8493
      Rows            =   1
      Cols            =   5
      Redraw          =   -1  'True
      ShowGrid        =   -1  'True
      GridSolid       =   -1  'True
      GridLineColor   =   12632256
      BackColorFixed  =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontEdit {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Listar"
      Height          =   375
      Left            =   10200
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   "C:\Padmin\BD\padminbd.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "a�oactivo"
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
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\Padmin\BD\padminbd.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
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
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   11160
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Borrar Formulario"
      Height          =   375
      Left            =   6120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   11160
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   8040
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   11160
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   120
      Picture         =   "listadofecha.frx":0000
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
      StartOfWeek     =   51314690
      CurrentDate     =   39848
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   8280
      TabIndex        =   11
      Top             =   5160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   51314689
      CurrentDate     =   39854
   End
   Begin VB.Label Label3 
      Caption         =   "Hasta Fecha de entrada"
      Height          =   255
      Left            =   6360
      TabIndex        =   7
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Desde Fecha de Entrada"
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Line Line3 
      X1              =   944
      X2              =   8
      Y1              =   336
      Y2              =   336
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "   LISTADO POR FECHA DE ENTRADA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   2
      Top             =   4200
      Width           =   5415
   End
   Begin VB.Line Line1 
      X1              =   944
      X2              =   8
      Y1              =   256
      Y2              =   256
   End
End
Attribute VB_Name = "listadofecha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload listadofecha
    
End Sub

Private Sub Command2_Click()
    Dim parteinicial As Integer
    Dim partefinal As Integer
    Dim nparte As Integer
    Dim linea As Integer
    
    
    ubGrid1.Rows = 0
    ubGrid1.Rows = 1
    
    If DTPicker1.Value > DTPicker2.Value Then
        DTPicker1.Value = Format$(Now, "dd/mm/yyyy")
        DTPicker2.Value = Format$(Now, "dd/mm/yyyy")
        DTPicker1.SetFocus
    Else
        linea = 1
'        parteinicial = Val(Text1.Text)
'        partefinal = Val(Text2.Text)
    
        Data4.Recordset.MoveFirst
        Data3.Recordset.MoveFirst
        Do While Not Data3.Recordset.EOF
            If Data3.Recordset("valido") = "si" Then
                If Data3.Recordset("a�o") = Data4.Recordset("a�o") Then
                    DTPicker3.Value = Data3.Recordset("entrada")
'                    nparte = Val(Data3.Recordset("nparte"))
'                    If nparte >= parteinicial And nparte <= partefinal Then
                    If DTPicker3.Value >= DTPicker1.Value Then
                        If DTPicker3.Value <= DTPicker2.Value Then
                            ubGrid1.TextMatrix(linea, 1) = Data3.Recordset("nparte")
                            ubGrid1.TextMatrix(linea, 2) = Data3.Recordset("dni")
                            ubGrid1.TextMatrix(linea, 3) = Data3.Recordset("nombre")
                            ubGrid1.TextMatrix(linea, 4) = Data3.Recordset("apellidos")
                            ubGrid1.TextMatrix(linea, 5) = Data3.Recordset("localidad")
                            ubGrid1.TextMatrix(linea, 6) = Data3.Recordset("provincia")
                            ubGrid1.TextMatrix(linea, 7) = Data3.Recordset("habitacion")
                            ubGrid1.TextMatrix(linea, 8) = Data3.Recordset("personas")
                            ubGrid1.TextMatrix(linea, 9) = Data3.Recordset("precio")
                            ubGrid1.TextMatrix(linea, 10) = Format$(Data3.Recordset("entrada"), "dd/mm/yyyy")
                            ubGrid1.TextMatrix(linea, 11) = Format$(Data3.Recordset("salida"), "dd/mm/yyyy")
                        
                            If Data3.Recordset("regimen") <> "N" Then
                                ubGrid1.TextMatrix(linea, 12) = Data3.Recordset("regimen")
                            End If
                    
                            ubGrid1.AddItem ("")
                            linea = linea + 1
                            Command4.Enabled = True
                        End If
                    End If
                End If
            End If
            Data3.Recordset.MoveNext
        Loop
    End If
    
                    
                
                
            
End Sub

Private Sub Command3_Click()
    DTPicker1.Value = Format$(Now, "dd/mm/yyyy")
    DTPicker2.Value = Format$(Now, "dd/mm/yyyy")
    Command4.Enabled = False
    ubGrid1.Rows = 0
    ubGrid1.Rows = 1
    DTPicker1.SetFocus
End Sub

Private Sub Command4_Click()
    Dim cadtemp As String
    Dim cadena As String
    
    nlineas = ubGrid1.Rows
    linea = 8
    i = 1
'    j = 0
    nuevapag = 1
    terminado = 0
    
    
'    Printer.ColorMode = 2
'    Printer.ForeColor = RGB(100, 100, 100)
    Do While i <= nlineas
        If nuevapag = 1 Then
            Printer.Print " "
            Printer.Print " "
            Printer.Print "Listado de Partes de Admisi�n, por Fecha de Entrada. Desde: "; DTPicker1.Value; "   Hasta: "; DTPicker2.Value
            Printer.Print " "
            Printer.Print Now
            Printer.Print " "
            Printer.Print "PARTE"; Tab(10); "DNI"; Tab(23); "NOMBRE"; Tab(43); "APELLIDOS"; Tab(70); "POBLACION"; Tab(95); "PROVINCIA"; Tab(109); "Hab."; Tab(115); "Per."; Tab(120); "PRECIO"; Tab(130); "ENTRADA"; Tab(142); "SALIDA"; Tab(154); "Reg."
            Printer.Print "--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"


            linea = 8
            nuevapag = 0
        End If
        terminado = 0
        Printer.Print ubGrid1.TextMatrix(i, 1);
        
        cadtemp = ubGrid1.TextMatrix(i, 2)
        cadena = Left(cadtemp, 10)
        Printer.Print Tab(10); cadena;
        
        cadtemp = ubGrid1.TextMatrix(i, 3)
        cadena = Left(cadtemp, 11)
        Printer.Print Tab(23); cadena;
        
        cadtemp = ubGrid1.TextMatrix(i, 4)
        cadena = Left(cadtemp, 15)
        Printer.Print Tab(43); cadena;
        
        cadtemp = ubGrid1.TextMatrix(i, 5)
        cadena = Left(cadtemp, 17)
        Printer.Print Tab(70); cadena;
        
        cadtemp = ubGrid1.TextMatrix(i, 6)
        cadena = Left(cadtemp, 8)
        Printer.Print Tab(95); cadena; Tab(109); ubGrid1.TextMatrix(i, 7); Tab(115); ubGrid1.TextMatrix(i, 8); Tab(120); ubGrid1.TextMatrix(i, 9); Tab(130); ubGrid1.TextMatrix(i, 10); Tab(142); ubGrid1.TextMatrix(i, 11); Tab(154); ubGrid1.TextMatrix(i, 12)
        linea = linea + 1
        i = i + 1
        If linea = 80 Then
            nuevpag = 1
            Printer.EndDoc
            terminado = 1
        End If
    Loop
    Printer.EndDoc
       
    DTPicker1.Value = Format$(Now, "dd/mm/yyyy")
    DTPicker2.Value = Format$(Now, "dd/mm/yyyy")
    Command4.Enabled = False
    ubGrid1.Rows = 0
    ubGrid1.Rows = 1
    DTPicker1.SetFocus
            

End Sub

Private Sub DTPicker1_GotFocus()
    DTPicker1.Value = Format$(Now, "dd/mm/yyyy")
    DTPicker2.Value = Format$(Now, "dd/mm/yyyy")
End Sub

Private Sub Form_Load()
    ubGrid1.AutoSetup 1, 12, True, True, "Parte  |D.N.I.          |Nombre                  |Apellidos                           |Poblaci�n                                   |Provincia         |Habt.   |Pers.   |Precio      |Entrada       |Salida          |Reg."
    ubGrid1.Enabled = True
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    Dim nparte As Integer
    Dim nmaxparte As Integer
    
    If KeyAscii = 13 Then
        If Text1.Text <> "" Then
            Data1.Recordset.MoveFirst
            nmaxparte = Val(Data1.Recordset("numero"))
            nparte = Val(Text1.Text)
            If nparte < nmaxparte And nparte > 0 Then
                Text2.Enabled = True
                Text2.SetFocus
            Else
                Text1.Text = ""
            End If
        End If
    End If
            
            
    
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    Dim nparte As Integer
    Dim nmaxparte As Integer
    
    If KeyAscii = 13 Then
        If Text2.Text <> "" Then
            Data1.Recordset.MoveFirst
            nmaxparte = Val(Data1.Recordset("numero"))
            nparte = Val(Text2.Text)
            If nparte >= nmaxparte Then
                nparte = nmaxparte - 1
                Text2.Text = nparte
            End If
            
            If nparte < nmaxparte And nparte >= Val(Text1.Text) Then
                Command2.Enabled = True
                Command2.SetFocus
            Else
                Text2.Text = ""
            End If
        End If
    End If

End Sub
