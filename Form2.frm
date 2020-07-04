VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   8130
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12030
   LinkTopic       =   "Form2"
   ScaleHeight     =   8130
   ScaleWidth      =   12030
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "Foto"
      Height          =   375
      Left            =   7200
      TabIndex        =   25
      Top             =   7200
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command6 
      Caption         =   "ELIMINAR"
      Height          =   495
      Left            =   8880
      TabIndex        =   24
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "MODIFICAR"
      Height          =   495
      Left            =   8880
      TabIndex        =   23
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "GUARDAR"
      Height          =   495
      Left            =   8880
      TabIndex        =   22
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "NUEVO"
      Height          =   495
      Left            =   8880
      TabIndex        =   21
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   20
      Top             =   7200
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   19
      Top             =   7200
      Width           =   735
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   120
      Top             =   720
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\MerrickGT\Desktop\Proyecto1\Zoologico.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\MerrickGT\Desktop\Proyecto1\Zoologico.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Animales"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text8 
      DataField       =   "Lugar_Origen"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   3480
      TabIndex        =   17
      Top             =   7320
      Width           =   2415
   End
   Begin VB.TextBox Text7 
      DataField       =   "Edad"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   3480
      TabIndex        =   15
      Top             =   6960
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      DataField       =   "Tipo_Alimentacion"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   3480
      TabIndex        =   13
      Top             =   6600
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      DataField       =   "Peso"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   3480
      TabIndex        =   10
      Top             =   6240
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      DataField       =   "Cantidad"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   3480
      TabIndex        =   8
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      DataField       =   "Especies"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   3480
      TabIndex        =   6
      Top             =   5520
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      DataField       =   "Nombre"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   3480
      TabIndex        =   4
      Top             =   5160
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      DataField       =   "Codigo"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   3480
      TabIndex        =   2
      Top             =   4800
      Width           =   975
   End
   Begin VB.Image Image2 
      Height          =   2175
      Left            =   6360
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   2295
   End
   Begin VB.Label Label11 
      DataField       =   "Foto"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Lugar de Origen:"
      BeginProperty Font 
         Name            =   "Montserrat Medium"
         Size            =   8.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   16
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Edad:"
      BeginProperty Font 
         Name            =   "Montserrat Medium"
         Size            =   8.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   14
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Alimentación"
      BeginProperty Font 
         Name            =   "Montserrat Medium"
         Size            =   8.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   12
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Kg"
      BeginProperty Font 
         Name            =   "Montserrat"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   11
      Top             =   6240
      Width           =   255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Peso:"
      BeginProperty Font 
         Name            =   "Montserrat Medium"
         Size            =   8.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad:"
      BeginProperty Font 
         Name            =   "Montserrat Medium"
         Size            =   8.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Especie:"
      BeginProperty Font 
         Name            =   "Montserrat Medium"
         Size            =   8.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      BeginProperty Font 
         Name            =   "Montserrat Medium"
         Size            =   8.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
      BeginProperty Font 
         Name            =   "Montserrat Medium"
         Size            =   8.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ANIMALES"
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   12015
   End
   Begin VB.Image Image1 
      Height          =   8400
      Left            =   0
      Picture         =   "Form2.frx":0000
      Stretch         =   -1  'True
      Top             =   -240
      Width           =   12015
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Adodc1.Recordset.MovePrevious
    X = App.Path

    If Adodc1.Recordset.BOF Then
        Adodc1.Recordset.MoveLast
    End If
    
    Image2.Picture = LoadPicture(X & "\" & Label11.Caption)
   
End Sub

Private Sub Command2_Click()
    Adodc1.Recordset.MoveNext
    X = App.Path
    
    If Adodc1.Recordset.EOF Then
        Adodc1.Recordset.MoveFirst
    End If
    
    Image2.Picture = LoadPicture(X & "\" & Label11.Caption)

End Sub

Private Sub Command3_Click()
    Adodc1.Recordset.AddNew
    
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
    Text5.Enabled = True
    Text6.Enabled = True
    Text7.Enabled = True
    Text8.Enabled = True
    Command4.Enabled = True
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
    Command5.Enabled = False
    Command6.Enabled = False
    Command7.Enabled = True
    
    Text1.SetFocus
    ''' Elimina el contenido donde se muestra el nombre de la foto
    Label11.Caption = ""
    ''' Borra el contenido del Image al no encontrar una foto
    Image2.Picture = LoadPicture(Label11.Caption)
    
End Sub

Private Sub Command4_Click()
    FileCopy CommonDialog1.FileName, App.Path & "\\" & CommonDialog1.FileTitle
    Adodc1.Recordset.Update
    Adodc1.Recordset.MoveFirst
    X = App.Path
    Image2.Picture = LoadPicture(X & "\" & Label11.Caption)
    
    Text1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
    Text5.Enabled = False
    Text6.Enabled = False
    Text7.Enabled = False
    Text8.Enabled = False
    Command4.Enabled = False
    Command1.Enabled = True
    Command2.Enabled = True
    Command3.Enabled = True
    Command5.Enabled = True
    Command6.Enabled = True
    Command7.Enabled = False
    
End Sub

Private Sub Command5_Click()

    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
    Text5.Enabled = True
    Text6.Enabled = True
    Text7.Enabled = True
    Text8.Enabled = True
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = True
    Command5.Enabled = False
    Command6.Enabled = False


End Sub

Private Sub Command6_Click()
    Adodc1.Recordset.Delete
    Adodc1.Recordset.MoveFirst
    X = App.Path
    Image2.Picture = LoadPicture(X & "\" & Label11.Caption)
End Sub

Private Sub Command7_Click()
    CommonDialog1.ShowOpen
    Image2.Picture = LoadPicture(CommonDialog1.FileName)
    Label11.Caption = CommonDialog1.FileTitle
    
    If Label11.Caption = "" Then
        MsgBox ("Seleccione una imagen para continuar")
    Else
        Label11.Caption = CommonDialog1.FileTitle
    End If
End Sub

Private Sub Form_Load()
    X = App.Path
    Image2.Picture = LoadPicture(X & "\" & Label11.Caption)
    Text1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
    Text5.Enabled = False
    Text6.Enabled = False
    Text7.Enabled = False
    Text8.Enabled = False
    Command4.Enabled = False
    Command7.Enabled = False
End Sub

