VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSetUp 
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SetUp"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab Tab1 
      Height          =   4215
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   7435
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   8421504
      TabCaption(0)   =   "User Data"
      TabPicture(0)   =   "frmSetUp.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Company Information"
      TabPicture(1)   =   "frmSetUp.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Screen Text"
      TabPicture(2)   =   "frmSetUp.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Tab2"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame2 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   15
         Top             =   360
         Width           =   5895
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "E-MailAdress"
            DataSource      =   "rsMyRec"
            Height          =   285
            Index           =   10
            Left            =   2880
            TabIndex        =   35
            Top             =   3120
            Width           =   2895
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "ContactPerson"
            DataSource      =   "rsMyRec"
            Height          =   285
            Index           =   9
            Left            =   2880
            TabIndex        =   32
            Top             =   2880
            Width           =   2895
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "Fax"
            DataSource      =   "rsMyRec"
            Height          =   285
            Index           =   8
            Left            =   2880
            TabIndex        =   31
            Top             =   2520
            Width           =   1815
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "Telefon"
            DataSource      =   "rsMyRec"
            Height          =   285
            Index           =   7
            Left            =   2880
            TabIndex        =   28
            Top             =   2280
            Width           =   1815
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "Country"
            DataSource      =   "rsMyRec"
            Height          =   285
            Index           =   6
            Left            =   2880
            TabIndex        =   27
            Top             =   1920
            Width           =   2295
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "Town"
            DataSource      =   "rsMyRec"
            Height          =   285
            Index           =   5
            Left            =   2880
            TabIndex        =   22
            Top             =   1680
            Width           =   2295
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "Zip"
            DataSource      =   "rsMyRec"
            Height          =   285
            Index           =   4
            Left            =   2880
            TabIndex        =   21
            Top             =   1440
            Width           =   1215
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "Adress3"
            DataSource      =   "rsMyRec"
            Height          =   285
            Index           =   3
            Left            =   2880
            TabIndex        =   20
            Top             =   1080
            Width           =   2295
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "Adress2"
            DataSource      =   "rsMyRec"
            Height          =   285
            Index           =   2
            Left            =   2880
            TabIndex        =   19
            Top             =   840
            Width           =   2295
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "Adress1"
            DataSource      =   "rsMyRec"
            Height          =   285
            Index           =   1
            Left            =   2880
            TabIndex        =   18
            Top             =   600
            Width           =   2295
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "Company"
            DataSource      =   "rsMyRec"
            Height          =   285
            Index           =   0
            Left            =   2880
            TabIndex        =   16
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Person E-mail:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   34
            Top             =   3120
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Person:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   33
            Top             =   2880
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Fax:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   30
            Top             =   2520
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Phone:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   29
            Top             =   2280
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Country:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   26
            Top             =   1920
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Town:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   25
            Top             =   1680
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Zip Code:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   24
            Top             =   1440
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   23
            Top             =   840
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Company Name:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   2655
         End
      End
      Begin TabDlg.SSTab Tab2 
         Height          =   3015
         Left            =   -74760
         TabIndex        =   11
         Top             =   600
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   5318
         _Version        =   393216
         TabOrientation  =   1
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "Main Screen"
         TabPicture(0)   =   "frmSetUp.frx":0054
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "DBGrid1(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Data1"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Package"
         TabPicture(1)   =   "frmSetUp.frx":0070
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "DBGrid1(1)"
         Tab(1).Control(1)=   "Data2"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Set Up"
         TabPicture(2)   =   "frmSetUp.frx":008C
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "DBGrid1(2)"
         Tab(2).Control(1)=   "Data3"
         Tab(2).ControlCount=   2
         TabCaption(3)   =   "Excel Text"
         TabPicture(3)   =   "frmSetUp.frx":00A8
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "DBGrid1(3)"
         Tab(3).Control(1)=   "Data4"
         Tab(3).ControlCount=   2
         Begin VB.Data Data4 
            Caption         =   "Data1"
            Connect         =   "Access 2000;"
            DatabaseName    =   "H:\Egne Programmer\Posten Norge\PostenMultiLanguage\ThePostMan.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   -74040
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "ExcelText"
            Top             =   720
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.Data Data3 
            Caption         =   "Data1"
            Connect         =   "Access 2000;"
            DatabaseName    =   "H:\Egne Programmer\Posten Norge\PostenMultiLanguage\ThePostMan.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   -74760
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "frmSetUp"
            Top             =   1200
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.Data Data2 
            Caption         =   "Data1"
            Connect         =   "Access 2000;"
            DatabaseName    =   "H:\Egne Programmer\Posten Norge\PostenMultiLanguage\ThePostMan.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   -74760
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "frmPackage"
            Top             =   1200
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.Data Data1 
            Caption         =   "Data1"
            Connect         =   "Access 2000;"
            DatabaseName    =   "H:\Egne Programmer\Posten Norge\PostenMultiLanguage\ThePostMan.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   480
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "frmPostOffice"
            Top             =   1320
            Visible         =   0   'False
            Width           =   1140
         End
         Begin MSDBGrid.DBGrid DBGrid1 
            Bindings        =   "frmSetUp.frx":00C4
            Height          =   2415
            Index           =   0
            Left            =   120
            OleObjectBlob   =   "frmSetUp.frx":00D8
            TabIndex        =   12
            Top             =   120
            Width           =   5655
         End
         Begin MSDBGrid.DBGrid DBGrid1 
            Bindings        =   "frmSetUp.frx":0AAE
            Height          =   2295
            Index           =   1
            Left            =   -74880
            OleObjectBlob   =   "frmSetUp.frx":0AC2
            TabIndex        =   13
            Top             =   120
            Width           =   5655
         End
         Begin MSDBGrid.DBGrid DBGrid1 
            Bindings        =   "frmSetUp.frx":1498
            Height          =   2295
            Index           =   2
            Left            =   -74880
            OleObjectBlob   =   "frmSetUp.frx":14AC
            TabIndex        =   14
            Top             =   120
            Width           =   5655
         End
         Begin MSDBGrid.DBGrid DBGrid1 
            Bindings        =   "frmSetUp.frx":1E82
            Height          =   2295
            Index           =   3
            Left            =   -74880
            OleObjectBlob   =   "frmSetUp.frx":1E96
            TabIndex        =   38
            Top             =   120
            Width           =   5655
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3495
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   6015
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "PostalRefNo"
            DataSource      =   "rsMyRec"
            Height          =   285
            Left            =   2280
            TabIndex        =   37
            Top             =   2400
            Width           =   2175
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "NoToPrint"
            DataSource      =   "rsMyRec"
            Height          =   285
            Index           =   1
            Left            =   2280
            TabIndex        =   9
            Top             =   1920
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "PostFolder"
            DataSource      =   "rsMyRec"
            Height          =   285
            Left            =   2280
            TabIndex        =   4
            Top             =   720
            Width           =   3495
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "PostNumber"
            DataSource      =   "rsMyRec"
            Height          =   285
            Index           =   0
            Left            =   2280
            TabIndex        =   3
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Data rsMyRec 
            Caption         =   "Data1"
            Connect         =   "Access 2000;"
            DatabaseName    =   "H:\Egne Programmer\Posten Norge\PostenMultiLanguage\ThePostMan.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   0
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "MyRec"
            Top             =   0
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.ComboBox cboCountry 
            BackColor       =   &H00FFFFC0&
            DataField       =   "LanguageScreen"
            DataSource      =   "rsMyRec"
            Height          =   315
            Left            =   2280
            TabIndex        =   2
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Postal Ref. Number:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   36
            Top             =   2400
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "No. of copies to print:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   10
            Top             =   1920
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Post List Folder:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Next Free Post List No.:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   6
            Top             =   1080
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Language on screen:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   5
            Top             =   1440
            Width           =   2055
         End
      End
   End
   Begin Project1.LaVolpeButton btnExit 
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   4440
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   873
      BTNICON         =   "frmSetUp.frx":286C
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
      MICON           =   "frmSetUp.frx":29C6
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
End
Attribute VB_Name = "frmSetUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bookmarkCountry() As Variant
Dim rsCountry As Recordset
Dim rsLanguage As Recordset
Private Sub LoadCountry()
    cboCountry.Clear
    With rsCountry
        .MoveLast
        ReDim bookmarkCountry(.RecordCount)
        .MoveFirst
        Do While Not .EOF
            cboCountry.AddItem .Fields("Country")
            cboCountry.ItemData(cboCountry.NewIndex) = cboCountry.ListCount - 1
            bookmarkCountry(cboCountry.ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
End Sub
Private Sub ShowText()
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
                If IsNull(.Fields("btnExit")) Then
                    .Fields("btnExit") = btnExit.Caption
                Else
                    btnExit.Caption = .Fields("btnExit")
                End If
                Tab1.Tab = 0
                If IsNull(.Fields("Tab10")) Then
                    .Fields("Tab10") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab10")
                End If
                Tab1.Tab = 1
                If IsNull(.Fields("Tab11")) Then
                    .Fields("Tab11") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab11")
                End If
                Tab1.Tab = 2
                If IsNull(.Fields("Tab12")) Then
                    .Fields("Tab12") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab12")
                End If
                Tab1.Tab = 0
                Tab2.Tab = 0
                If IsNull(.Fields("Tab20")) Then
                    .Fields("Tab20") = Tab2.Caption
                Else
                    Tab2.Caption = .Fields("Tab20")
                End If
                Tab2.Tab = 1
                If IsNull(.Fields("Tab21")) Then
                    .Fields("Tab21") = Tab2.Caption
                Else
                    Tab2.Caption = .Fields("Tab21")
                End If
                Tab2.Tab = 2
                If IsNull(.Fields("Tab22")) Then
                    .Fields("Tab22") = Tab2.Caption
                Else
                    Tab2.Caption = .Fields("Tab22")
                End If
                Tab2.Tab = 0
                .Update
                Exit Sub
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
        .Fields("btnExit") = btnExit.Caption
        Tab1.Tab = 0
        .Fields("Tab10") = Tab1.Caption
        Tab1.Tab = 1
        .Fields("Tab11") = Tab1.Caption
        Tab1.Tab = 2
        .Fields("Tab12") = Tab1.Caption
        Tab1.Tab = 0
        Tab2.Tab = 0
        .Fields("Tab20") = Tab2.Caption
        Tab1.Tab = 1
        .Fields("Tab21") = Tab2.Caption
        Tab1.Tab = 2
        .Fields("Tab22") = Tab2.Caption
        Tab2.Tab = 0
        .Update
    End With
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub cboCountry_Click()
    On Error Resume Next
    rsCountry.Bookmark = bookmarkCountry(cboCountry.ItemData(cboCountry.ListIndex))
    cboCountry.Text = rsCountry.Fields("CountryFix")
    m_FileExt = Trim(cboCountry.Text)
    ShowText
    frmPostOffice.ShowText
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    LoadCountry
    rsMyRec.Refresh
    Data1.Refresh
    Data2.Refresh
    Data3.Refresh
    Data4.Refresh
    ShowText
End Sub
Private Sub Form_Load()
    'On Error Resume Next
    rsMyRec.DatabaseName = sDirPath
    Set rsCountry = dbPosten.OpenRecordset("Country")
    Set rsLanguage = dbPosten.OpenRecordset("frmSetUp")
    Data1.DatabaseName = sDirPath
    Data2.DatabaseName = sDirPath
    Data3.DatabaseName = sDirPath
    Data4.DatabaseName = sDirPath
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsMyRec.UpdateRecord
    rsMyRec.Recordset.Close
    Data1.Recordset.Close
    Data2.Recordset.Close
    Data3.Recordset.Close
    Data4.Recordset.Close
    rsCountry.Close
    rsLanguage.Close
    Set frmSetUp = Nothing
End Sub
