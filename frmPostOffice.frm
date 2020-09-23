VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmPostOffice 
   BackColor       =   &H00808080&
   Caption         =   "Postal Service - Postal Order - Delivery of pre-adressed letters"
   ClientHeight    =   6795
   ClientLeft      =   180
   ClientTop       =   465
   ClientWidth     =   12270
   Icon            =   "frmPostOffice.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   12270
   StartUpPosition =   2  'CenterScreen
   Begin MSDBGrid.DBGrid Grid1 
      Bindings        =   "frmPostOffice.frx":0442
      Height          =   4335
      Left            =   120
      OleObjectBlob   =   "frmPostOffice.frx":045C
      TabIndex        =   21
      ToolTipText     =   "Dbl click to get letter weight"
      Top             =   2280
      Width           =   9015
   End
   Begin Project1.LaVolpeButton btnNew 
      Height          =   2055
      Left            =   5040
      TabIndex        =   19
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   3625
      BTNICON         =   "frmPostOffice.frx":0E2D
      BTYPE           =   3
      TX              =   "&New Post List"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   12648384
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmPostOffice.frx":14FF
      ALIGN           =   1
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   1
      IconSize        =   3
      SHOWF           =   -1  'True
      BSTYLE          =   0
      OPTVAL          =   0   'False
      OPTMOD          =   0   'False
      GStart          =   0
      GStop           =   16711680
      GStyle          =   0
   End
   Begin Project1.LaVolpeButton btnDelete 
      Height          =   2055
      Left            =   7920
      TabIndex        =   18
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   3625
      BTNICON         =   "frmPostOffice.frx":151B
      BTYPE           =   3
      TX              =   "&Delete Post List"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   12648384
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmPostOffice.frx":1675
      ALIGN           =   1
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   1
      IconSize        =   3
      SHOWF           =   -1  'True
      BSTYLE          =   0
      OPTVAL          =   0   'False
      OPTMOD          =   0   'False
      GStart          =   0
      GStop           =   16711680
      GStyle          =   0
   End
   Begin Project1.LaVolpeButton btnPrint 
      Height          =   735
      Left            =   6360
      TabIndex        =   17
      Top             =   1440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
      BTNICON         =   "frmPostOffice.frx":1691
      BTYPE           =   3
      TX              =   "&Print Post List"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   12648384
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmPostOffice.frx":17EB
      ALIGN           =   1
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   1
      IconSize        =   3
      SHOWF           =   -1  'True
      BSTYLE          =   0
      OPTVAL          =   0   'False
      OPTMOD          =   0   'False
      GStart          =   0
      GStop           =   16711680
      GStyle          =   0
   End
   Begin Project1.LaVolpeButton btnExit 
      Height          =   735
      Left            =   6360
      TabIndex        =   16
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
      BTNICON         =   "frmPostOffice.frx":1807
      BTYPE           =   3
      TX              =   "&Exit"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   12648384
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmPostOffice.frx":1961
      ALIGN           =   1
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   1
      IconSize        =   3
      SHOWF           =   -1  'True
      BSTYLE          =   0
      OPTVAL          =   0   'False
      OPTMOD          =   0   'False
      GStart          =   0
      GStop           =   16711680
      GStyle          =   0
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   5880
      Index           =   2
      Left            =   10200
      TabIndex        =   12
      Top             =   720
      Width           =   735
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   5880
      Index           =   1
      Left            =   10920
      TabIndex        =   11
      Top             =   720
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   2055
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4815
      Begin Project1.LaVolpeButton btnNewPartDelivery 
         Height          =   495
         Left            =   3000
         TabIndex        =   20
         Top             =   1320
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "New &Delivery"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   12648384
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmPostOffice.frx":197D
         ALIGN           =   1
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   1
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
         OPTVAL          =   0   'False
         OPTMOD          =   0   'False
         GStart          =   0
         GStop           =   16711680
         GStyle          =   0
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "Part Delivery ?"
         DataField       =   "PartDelivery"
         DataSource      =   "rsPostenNorge"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   2
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "PartDeliveryNo"
         DataSource      =   "rsPostenNorge"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   2040
         TabIndex        =   3
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "PostDate"
         DataSource      =   "rsPostenNorge"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   2040
         TabIndex        =   1
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "PostNumber"
         DataSource      =   "rsPostenNorge"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   2040
         TabIndex        =   0
         Top             =   360
         Width           =   1455
      End
      Begin VB.Image Image1 
         Height          =   495
         Left            =   3600
         Picture         =   "frmPostOffice.frx":1999
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Delivery Number:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   10
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Postal Date:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Post Number:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   60
      Left            =   7920
      Top             =   0
   End
   Begin VB.TextBox txtYear 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   6960
      TabIndex        =   5
      Top             =   960
      Width           =   855
   End
   Begin VB.Data rsPackets 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Programmering\Sandvik\Posten Norge\PostenNorge.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Parcels"
      Top             =   0
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data rsPostLines 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Programmering\Sandvik\Posten Norge\PostenNorge.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ParcelPostLines"
      Top             =   0
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data rsPostenNorge 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Programmering\Sandvik\Posten Norge\PostenNorge.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ParcelPostHead"
      Top             =   0
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   5880
      Index           =   0
      Left            =   9240
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   10920
      TabIndex        =   15
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Del."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   10200
      TabIndex        =   14
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Number"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   9240
      TabIndex        =   13
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Year:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   6240
      TabIndex        =   6
      Top             =   960
      Width           =   735
   End
   Begin VB.Menu mnuFiles 
      Caption         =   "&Files"
      Begin VB.Menu mnuSlutt 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuPack 
      Caption         =   "&Package"
   End
   Begin VB.Menu mnuSetUp 
      Caption         =   "&Set Up"
   End
End
Attribute VB_Name = "frmPostOffice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bookmarks1() As Variant
Dim iColNo As Integer, boolEdit As Boolean
Dim lPostNo As Long, lPart As Long
Dim bNewRecord As Boolean
Dim rsMyRec As Recordset
Dim rsExcelText As Recordset
Dim rsLanguage As Recordset
Dim mobjExcel As Excel.Application
Public Sub ShowText()
Dim strHelp As String
    'find YOUR Language text
    With rsLanguage
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = m_FileExt Then
                .Edit
                If IsNull(.Fields("Form")) Then
                    .Fields("Form") = Me.Caption
                Else
                    Me.Caption = .Fields("Form")
                End If
                If IsNull(.Fields("Label1(0)")) Then
                    .Fields("Label1(0)") = Label1(0).Caption
                Else
                    Label1(0).Caption = .Fields("Label1(0)")
                End If
                If IsNull(.Fields("Label1(1)")) Then
                    .Fields("Label1(1)") = Label1(1).Caption
                Else
                    Label1(1).Caption = .Fields("Label1(1)")
                End If
                If IsNull(.Fields("Label1(2)")) Then
                    .Fields("Label1(2)") = Label1(2).Caption
                Else
                    Label1(2).Caption = .Fields("Label1(2)")
                End If
                If IsNull(.Fields("Label1(3)")) Then
                    .Fields("Label1(3)") = Label1(3).Caption
                Else
                    Label1(3).Caption = .Fields("Label1(3)")
                End If
                If IsNull(.Fields("Label1(4)")) Then
                    .Fields("Label1(4)") = Label1(4).Caption
                Else
                    Label1(4).Caption = .Fields("Label1(4)")
                End If
                If IsNull(.Fields("Label1(5)")) Then
                    .Fields("Label1(5)") = Label1(5).Caption
                Else
                    Label1(5).Caption = .Fields("Label1(5)")
                End If
                If IsNull(.Fields("Label1(6)")) Then
                    .Fields("Label1(6)") = Label1(6).Caption
                Else
                    Label1(6).Caption = .Fields("Label1(6)")
                End If
                If IsNull(.Fields("Check1")) Then
                    .Fields("Check1") = Check1.Caption
                Else
                    Check1.Caption = .Fields("Check1")
                End If
                If IsNull(.Fields("btnNewPartDelivery")) Then
                    .Fields("btnNewPartDelivery") = btnNewPartDelivery.Caption
                Else
                    btnNewPartDelivery.Caption = .Fields("btnNewPartDelivery")
                End If
                If IsNull(.Fields("btnNew")) Then
                    .Fields("btnNew") = btnNew.Caption
                Else
                    btnNew.Caption = .Fields("btnNew")
                End If
                If IsNull(.Fields("btnPrint")) Then
                    .Fields("btnPrint") = btnPrint.Caption
                Else
                    btnPrint.Caption = .Fields("btnPrint")
                End If
                If IsNull(.Fields("btnDelete")) Then
                    .Fields("btnDelete") = btnDelete.Caption
                Else
                    btnDelete.Caption = .Fields("btnDelete")
                End If
                If IsNull(.Fields("btnExit")) Then
                    .Fields("btnExit") = btnExit.Caption
                Else
                    btnExit.Caption = .Fields("btnExit")
                End If
                If IsNull(.Fields("mnuFiles")) Then
                    .Fields("mnuFiles") = mnuFiles.Caption
                Else
                    mnuFiles.Caption = .Fields("mnuFiles")
                End If
                If IsNull(.Fields("mnuSlutt")) Then
                    .Fields("mnuSlutt") = mnuSlutt.Caption
                Else
                    mnuSlutt.Caption = .Fields("mnuSlutt")
                End If
                If IsNull(.Fields("mnuPack")) Then
                    .Fields("mnuPack") = mnuPack.Caption
                Else
                    mnuPack.Caption = .Fields("mnuPack")
                End If
                If IsNull(.Fields("mnuSetUp")) Then
                    .Fields("mnuSetUp") = mnuSetUp.Caption
                Else
                    mnuSetUp.Caption = .Fields("mnuSetUp")
                End If
                If IsNull(.Fields("Grid1")) Then
                    .Fields("Grid1") = Grid1.ToolTipText
                Else
                    Grid1.ToolTipText = .Fields("Grid1")
                End If
                .Update
                Exit Sub
            End If
        .MoveNext
        Loop
            
            'this language was not found, make it. Find the English text first
            strHelp = " "
            .MoveFirst
            Do While Not .EOF
                If .Fields("Language") = "ENG" Then
                    If Not IsNull(.Fields("Help")) Then
                        strHelp = .Fields("Help")
                        Exit Do
                    End If
                End If
            .MoveNext
            Loop
            
        .AddNew
        .Fields("Language") = m_FileExt
        .Fields("Form") = Me.Caption
        .Fields("Label1(0)") = Label1(0).Caption
        .Fields("Label1(1)") = Label1(1).Caption
        .Fields("Label1(2)") = Label1(2).Caption
        .Fields("Label1(3)") = Label1(3).Caption
        .Fields("Label1(4)") = Label1(4).Caption
        .Fields("Label1(5)") = Label1(5).Caption
        .Fields("Label1(6)") = Label1(6).Caption
        .Fields("Check1") = Check1.Caption
        .Fields("btnNewPartDelivery") = btnNewPartDelivery.Caption
        .Fields("btnNew") = btnNew.Caption
        .Fields("btnPrint") = btnPrint.Caption
        .Fields("btnDelete") = btnDelete.Caption
        .Fields("btnExit") = btnExit.Caption
        .Fields("mnuFiles") = mnuFiles.Caption
        .Fields("mnuSlutt") = mnuSlutt.Caption
        .Fields("mnuPack") = mnuPack.Caption
        .Fields("mnuSetUp") = mnuSetUp.Caption
        .Fields("grid1") = Grid1.ToolTipText
        .Fields("Help") = strHelp
        .Update
    End With
End Sub

Private Sub FindNextPartDeliveryNo()
Dim strSql As String
    On Error Resume Next
    For lPart = 0 To 20
        strSql = "SELECT * FROM ParcelPostHead WHERE CLng(PostNumber) ="
        strSql = strSql & Chr(34) & CLng(lPostNo) & Chr(34)
        strSql = strSql & "AND CLng(PartDeliveryNo) ="
        strSql = strSql & Chr(34) & CLng(lPart) & Chr(34)
        rsPostenNorge.RecordSource = strSql
        rsPostenNorge.Refresh
        If rsPostenNorge.Recordset.BOF And rsPostenNorge.Recordset.EOF Then Exit For
    Next
End Sub
Private Sub FindNextPostNo()
    On Error Resume Next
    With rsMyRec
        .MoveFirst
        lPostNo = CLng(.Fields("PostNumber"))
        .Edit
        .Fields("PostNumber") = CLng(.Fields("PostNumber")) + 1
        .Update
    End With
End Sub
Private Sub LoadList1()
    On Error Resume Next
    List1(0).Clear
    List1(1).Clear
    List1(2).Clear
    With rsPostenNorge.Recordset
        .MoveLast
        .MoveFirst
        ReDim bookmarks1(.RecordCount)
        Do While Not .EOF
            List1(0).AddItem .Fields("PostNumber")
            List1(0).ItemData(List1(0).NewIndex) = List1(0).ListCount - 1
            bookmarks1(List1(0).ListCount - 1) = .Bookmark
            List1(1).AddItem CDate(.Fields("PostDate"))
            List1(2).AddItem CLng(.Fields("PartDeliveryNo"))
        .MoveNext
        Loop
    End With
End Sub
Private Sub LoadList11()
Dim iIndex As Long
    On Error Resume Next
    SelectYear
    
    List1(0).Clear
    List1(1).Clear
    List1(2).Clear
    With rsPostenNorge.Recordset
        .MoveLast
        .MoveFirst
        ReDim bookmarks1(.RecordCount)
        Do While Not .EOF
            List1(0).AddItem .Fields("PostNumber")
            List1(0).ItemData(List1(0).NewIndex) = List1(0).ListCount - 1
            bookmarks1(List1(0).ListCount - 1) = .Bookmark
            List1(1).AddItem CDate(.Fields("PostDate"))
            List1(2).AddItem CLng(.Fields("PartDeliveryNo"))
            If CLng(.Fields("PostNumber")) = lPostNo Then
                If CLng(.Fields("PartDeliveryNo")) = lPart Then
                    iIndex = List1(0).ListCount - 1
                End If
            End If
        .MoveNext
        Loop
    End With
    List1(0).ListIndex = iIndex
End Sub

Private Sub PrintPostScheme()
Dim iRow As Integer
    On Error Resume Next
    With mobjExcel.Application
        'write the company spesific informmation
        .ActiveSheet.Range("A5").Select
        .ActiveCell.FormulaR1C1 = rsMyRec.Fields("PostalRefNo")
        .ActiveSheet.Range("A7").Select
        .ActiveCell.FormulaR1C1 = rsMyRec.Fields("Company")
        .ActiveSheet.Range("F7").Select
        .ActiveCell.FormulaR1C1 = rsMyRec.Fields("Telefon")
        .ActiveSheet.Range("G7").Select
        .ActiveCell.FormulaR1C1 = rsMyRec.Fields("Fax")
        .ActiveSheet.Range("H7").Select
        .ActiveCell.FormulaR1C1 = rsMyRec.Fields("E-MailAdress")
        .ActiveSheet.Range("A9").Select
        .ActiveCell.FormulaR1C1 = rsMyRec.Fields("Adress1")
        .ActiveSheet.Range("A11").Select
        .ActiveCell.FormulaR1C1 = rsMyRec.Fields("Zip")
        .ActiveSheet.Range("C11").Select
        .ActiveCell.FormulaR1C1 = rsMyRec.Fields("Town")
        .ActiveSheet.Range("A33").Select
        .ActiveCell.FormulaR1C1 = rsMyRec.Fields("Company")
        .ActiveSheet.Range("G33").Select
        .ActiveCell.FormulaR1C1 = rsMyRec.Fields("ContactPerson")
        .ActiveSheet.Range("J33").Select
        .ActiveCell.FormulaR1C1 = rsMyRec.Fields("Telefon")
        .ActiveSheet.Range("A36").Select
        .ActiveCell.FormulaR1C1 = rsMyRec.Fields("Town")
        .ActiveSheet.Range("A37").Select
        .ActiveCell.FormulaR1C1 = rsMyRec.Fields("ContactPerson")
        
        'write the post list spesific information
        .ActiveSheet.Range("F5:J5").Select
        .ActiveCell.FormulaR1C1 = rsPostenNorge.Recordset.Fields("PostNumber")
        If CBool(rsPostenNorge.Recordset.Fields("PartDelivery")) Then
            .ActiveSheet.Shapes("Text Box 1").Select
            .Selection.Characters.Text = "X"
            .ActiveSheet.Shapes("Text Box 4").Select
            .Selection.Characters.Text = "Delinnlevering nr. " & rsPostenNorge.Recordset.Fields("PartDeliveryNo")
        Else
            .ActiveSheet.Shapes("Text Box 2").Select
            .Selection.Characters.Text = "X"
        End If
        'read all postlines
        iRow = 17
        rsPostLines.Recordset.MoveFirst
        Do While Not rsPostLines.Recordset.EOF
            If CLng(rsPostLines.Recordset.Fields("Post Number")) = CLng(rsPostenNorge.Recordset.Fields("PostNumber")) Then
                If CLng(rsPostLines.Recordset.Fields("Line No")) = CLng(rsPostenNorge.Recordset.Fields("PartDeliveryNo")) Then
                    iRow = iRow + 1
                    If rsPostLines.Recordset.Fields("KatagoryA") = "X" Then
                        .ActiveSheet.Range("B" & iRow).Select
                    ElseIf rsPostLines.Recordset.Fields("KatagoryB") = "X" Then
                        .ActiveSheet.Range("C" & iRow).Select
                    Else
                        .ActiveSheet.Range("D" & iRow).Select
                    End If
                    .ActiveCell.FormulaR1C1 = "X"
                    .ActiveSheet.Range("E" & iRow).Select
                    .ActiveCell.FormulaR1C1 = CLng(rsPostLines.Recordset.Fields("WeightAPiece"))
                    If CLng(rsPostLines.Recordset.Fields("NumberSorted")) <> 0 Then
                        .ActiveSheet.Range("F" & iRow).Select
                        .ActiveCell.FormulaR1C1 = CLng(rsPostLines.Recordset.Fields("NumberSorted"))
                    Else
                        .ActiveSheet.Range("G" & iRow).Select
                        .ActiveCell.FormulaR1C1 = CLng(rsPostLines.Recordset.Fields("NumberUnsorted"))
                    End If
                    If CBool(rsPostLines.Recordset.Fields("ForMachine")) Then
                        .ActiveSheet.Range("H" & iRow).Select
                        .ActiveCell.FormulaR1C1 = "X"
                    End If
                    .ActiveSheet.Range("J" & iRow).Select
                    .ActiveCell.FormulaR1C1 = rsPostLines.Recordset.Fields("Ref")
                End If
            End If
        rsPostLines.Recordset.MoveNext
        Loop
        .ActiveSheet.Range("E36:F36").Select
        .ActiveCell.FormulaR1C1 = Format(CDate(rsPostenNorge.Recordset.Fields("PostDate")), "dd.mm.yyyy")
    End With
    
    'save this workbook
     mobjExcel.ActiveWorkbook.SaveAs FileName:= _
        (rsMyRec.Fields("PostFolder") & "\" & rsPostenNorge.Recordset.Fields("PostNumber") & "-" & rsPostenNorge.Recordset.Fields("PartDeliveryNo") & ".xls"), FileFormat:=xlNormal, _
        Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, _
        CreateBackup:=False
        
    'print the post sheet
    mobjExcel.ActiveWindow.SelectedSheets.PrintOut Copies:=CInt(rsMyRec.Fields("NoToPrint")), Collate:=True
End Sub

Private Sub SelectLines()
Dim strSql As String
    On Error Resume Next
    strSql = "SELECT * FROM ParcelPostLines WHERE CLng(PostNumber) ="
    strSql = strSql & Chr(34) & CLng(Text1(0).Text) & Chr(34)
    strSql = strSql & "AND CLng(LineNo) ="
    strSql = strSql & Chr(34) & CLng(Text1(2).Text) & Chr(34)
    strSql = strSql & "ORDER BY PostLineNo"
    rsPostLines.RecordSource = strSql
    rsPostLines.Refresh
End Sub


Private Sub SelectYear()
Dim strSql As String
    On Error Resume Next
    strSql = "SELECT * FROM ParcelPostHead WHERE Year(PostDate) ="
    strSql = strSql & Chr(34) & CInt(txtYear.Text) & Chr(34)
    strSql = strSql & "ORDER BY PostNumber, PartDeliveryNo"
    rsPostenNorge.RecordSource = strSql
    rsPostenNorge.Refresh
End Sub

Private Sub WriteHeadExcel()
    On Error Resume Next
    'open excel text-file and position to customer language
    Set rsExcelText = dbPosten.OpenRecordset("ExcelText")
    With rsExcelText
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = m_FileExt Then Exit Do
        .MoveNext
        Loop
    End With
    
    Set mobjExcel = New Excel.Application
    mobjExcel.Visible = True
    mobjExcel.Workbooks.Add Template:=App.Path & "\PostList.xlt"
    
    With mobjExcel.Application
        .ActiveSheet.Range("D2").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("TopHead")
        .ActiveSheet.Range("A3").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("CustomerInformation")
        .ActiveSheet.Range("A4").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("CustomerNumber")
        .ActiveSheet.Range("F4").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("CustomerRef")
        .ActiveSheet.Range("A6").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("CustomerName")
        .ActiveSheet.Range("F6").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("CustomerPhone")
        .ActiveSheet.Range("G6").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("CustomerFax")
        .ActiveSheet.Range("H6").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("CustomerMail")
        .ActiveSheet.Range("A8").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("CustomerAddress")
        .ActiveSheet.Range("F8").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("AtDelivery")
        .ActiveSheet.Range("A10").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("CustomerZip")
        .ActiveSheet.Range("C10").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("CustomerTown")
        .ActiveSheet.Range("B13").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("ShipSpecificatin")
        .ActiveSheet.Range("B14").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("ProductCatagory")
        .ActiveSheet.Range("E14").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("WeightPr")
        .ActiveSheet.Range("F14").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("MassMailNoShipment")
        .ActiveSheet.Range("I14").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("TotalWeight")
        .ActiveSheet.Range("J14").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("RefOnInvoice")
        .ActiveSheet.Range("B16").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("A")
        .ActiveSheet.Range("C16").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("B")
        .ActiveSheet.Range("D16").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("C")
        .ActiveSheet.Range("F16").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("Sorted")
        .ActiveSheet.Range("G16").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("Unsorted")
        .ActiveSheet.Range("H16").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("MachineSort")
        .ActiveSheet.Range("D28").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("Quantity")
        .ActiveSheet.Range("F28").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("Quantity")
        .ActiveSheet.Range("H28").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("Quantity")
        .ActiveSheet.Range("J28").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("Quantity")
        .ActiveSheet.Range("A29").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("FormatAdd")
        .ActiveSheet.Range("C29").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("n0-1000")
        .ActiveSheet.Range("E29").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("n1000-2000")
        .ActiveSheet.Range("G29").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("Registered")
        .ActiveSheet.Range("I29").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("CashOnDelivery")
        .ActiveSheet.Range("A31").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("DeliveryByPrinter")
        .ActiveSheet.Range("A32").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("Company")
        .ActiveSheet.Range("G32").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("CompanyContact")
        .ActiveSheet.Range("J32").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("CompanyPhone")
        .ActiveSheet.Range("A35").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("CompanyPlace")
        .ActiveSheet.Range("E35").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("CompanyDate")
        .ActiveSheet.Range("G35").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("Place")
        .ActiveSheet.Range("J35").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("DateAndTime")
        .ActiveSheet.Range("A41").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("CustomerSignature")
        .ActiveSheet.Range("A43").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("PostCommentary")
        .ActiveSheet.Range("A44").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("PostDemandSort")
        .ActiveSheet.Range("F44").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("PostDemandMachinSort")
        .ActiveSheet.Range("I44").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("ListChanged")
        .ActiveSheet.Range("A45").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("No")
        .ActiveSheet.Range("A46").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("SumWeightDelivery")
        .ActiveSheet.Range("F46").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("PostRemarks")
        .ActiveSheet.Range("E49").Select
        .ActiveCell.FormulaR1C1 = rsExcelText.Fields("Kg")
    End With
    
    PrintPostScheme
    
    rsExcelText.Close
    'release the excell object
    Set mobjExcel = Nothing
End Sub
Private Sub btnDelete_Click()
    On Error Resume Next
    'first delete all possible lines
    With rsPostLines.Recordset
        .MoveFirst
        Do While Not .EOF
            If CLng(.Fields("PostNumber")) = CLng(rsPostenNorge.Recordset.Fields("PostNumber")) Then
                If CLng(.Fields("LineNo")) = CLng(rsPostenNorge.Recordset.Fields("PartDeliveryNo")) Then
                    .Delete
                End If
            End If
        .MoveNext
        Loop
    End With
    'then delete the head-record itself
    rsPostenNorge.Recordset.Delete
    SelectYear
    LoadList1
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnNew_Click()
    On Error Resume Next
    rsPostenNorge.Recordset.AddNew
    boolNewRecord = True
    FindNextPostNo
    Text1(0).Text = lPostNo
    Text1(1).Text = Format(CDate(Now), "dd.mm.yyyy")
    Text1(2).Text = 0
    lPart = 0
    rsPostenNorge.Recordset.Update
    SelectYear
    LoadList11
    SelectLines
End Sub

Private Sub btnNewPartDelivery_Click()
Dim dDate As Date
    On Error Resume Next
    If Len(Text1(0).Text) = 0 Then
        Beep
        Text1(0).SetFocus
        Exit Sub
    End If
    With rsPostenNorge.Recordset
        lPostNo = CLng(.Fields("PostNumber"))
        dDate = CDate(.Fields("PostDate"))
        FindNextPartDeliveryNo  'find new running delivery number
        .AddNew
        .Fields("PostNumber") = lPostNo
        .Fields("PostDate") = dDate
        .Fields("PartDeliveryNo") = lPart
        .Fields("PartDelivery") = True
        .Update
        SelectYear
        LoadList11
    End With
    SelectLines
End Sub

Private Sub btnPrint_Click()
    WriteHeadExcel
End Sub
Private Sub Form_Activate()
    On Error Resume Next
    rsPackets.Refresh
    rsPostenNorge.Refresh
    rsPostLines.Refresh
    m_FileExt = rsMyRec.Fields("LanguageScreen")
    txtYear.Text = Year(Now)
    SelectYear
    LoadList1
    List1(0).ListIndex = (List1(0).ListCount - 1)
    ShowText
End Sub
Private Sub Form_Load()
    On Error GoTo errForm_Load
    sDirPath = App.Path & "\ThePostMan.mdb"
    rsPackets.DatabaseName = sDirPath
    rsPostenNorge.DatabaseName = sDirPath
    rsPostLines.DatabaseName = sDirPath
    
    Set dbPosten = OpenDatabase(sDirPath)
    Set rsMyRec = dbPosten.OpenRecordset("MyRec")
    Set rsLanguage = dbPosten.OpenRecordset("frmPostOffice")
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsPackets.Recordset.Close
    rsPostenNorge.Recordset.Close
    rsPostLines.Recordset.Close
    rsMyRec.Close
    rsLanguage.Close
    Set frmPostOffice = Nothing
End Sub

Private Sub Grid1_AfterColUpdate(ByVal ColIndex As Integer)
    'On Error Resume Next
    Select Case ColIndex
    Case 8
        Grid1.Columns(11).Value = (Grid1.Columns(8).Value * Grid1.Columns(7).Value) / 1000
    Case 9
        Grid1.Columns(11).Value = (Grid1.Columns(9).Value * Grid1.Columns(7).Value) / 1000
    Case Else
    End Select
End Sub

Private Sub Grid1_ColEdit(ByVal ColIndex As Integer)
    If ColIndex = 7 Then
    
    End If
End Sub

Private Sub Grid1_DblClick()
    On Error Resume Next
    If Grid1.Col = 7 Then
        boolFromMain = True
        frmPackage.Show 1
    End If
End Sub

Private Sub Grid1_OnAddNew()
    Grid1.Columns(0).Text = CLng(Text1(0).Text)
    Grid1.Columns(1).Text = CDate(Text1(1).Text)
    Grid1.Columns(2).Text = CLng(Text1(2).Text)
    Grid1.Columns(5).Text = "X"
End Sub
Private Sub List1_Click(Index As Integer)
    On Error Resume Next
    List1(0).ListIndex = List1(Index).ListIndex
    List1(1).ListIndex = List1(Index).ListIndex
    List1(2).ListIndex = List1(Index).ListIndex
    rsPostenNorge.Recordset.Bookmark = bookmarks1(List1(0).ItemData(List1(0).ListIndex))
    SelectLines
End Sub
Private Sub mnuPack_Click()
    frmPackage.Show 1
End Sub

Private Sub mnuSetUp_Click()
    frmSetUp.Show 1
End Sub
Private Sub mnuSlutt_Click()
    Unload Me
End Sub
Private Sub Timer1_Timer()
    On Error Resume Next
    List1(1).TopIndex = List1(0).TopIndex
    List1(2).TopIndex = List1(0).TopIndex
End Sub
Private Sub txtYear_Change()
    On Error Resume Next
    SelectYear
    LoadList1
End Sub


