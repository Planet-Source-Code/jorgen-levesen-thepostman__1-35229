VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmPackage 
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Package text"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSDBGrid.DBGrid Grid1 
      Bindings        =   "frmPackage.frx":0000
      Height          =   4455
      Left            =   120
      OleObjectBlob   =   "frmPackage.frx":0018
      TabIndex        =   1
      Top             =   120
      Width           =   4695
   End
   Begin Project1.LaVolpeButton btnExit 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   4680
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   873
      BTNICON         =   "frmPackage.frx":09E9
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
      MICON           =   "frmPackage.frx":0B43
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
   Begin VB.Data rsPackage 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Programmering\Sandvik\Posten Norge\PostenNorge.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Parcels"
      Top             =   0
      Visible         =   0   'False
      Width           =   1140
   End
End
Attribute VB_Name = "frmPackage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLanguage As Recordset
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
                If IsNull(.Fields("btnExit")) Then
                    .Fields("btnExit") = btnExit.Caption
                Else
                    btnExit.Caption = .Fields("btnExit")
                End If
                .Update
                Exit Sub
            End If
        .MoveNext
        Loop
        
        .AddNew
        .Fields("Language") = m_FileExt
        .Fields("Form") = Me.Caption
        .Fields("btnExit") = btnExit.Caption
        .Update
    End With
End Sub

Private Sub btnExit_Click()
    rsPackage.Recordset.Bookmark = Grid1.Bookmark
    If boolFromMain Then    'call came from frmPostNorge.Grid1
        frmPostOffice.Grid1.Text = CLng(rsPackage.Recordset.Fields("ParcelWeight"))
        frmPostOffice.Grid1.Col = 12
        frmPostOffice.Grid1.Text = CStr(rsPackage.Recordset.Fields("ParcelName"))
        boolFromMain = False
    End If
    Unload Me
End Sub
Private Sub Form_Activate()
    On Error Resume Next
    rsPackage.Refresh
    ShowText
End Sub
Private Sub Form_Load()
    On Error Resume Next
    rsPackage.DatabaseName = sDirPath
    Set rsLanguage = dbPosten.OpenRecordset("frmPackage")
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsPackage.Recordset.Close
    rsLanguage.Close
    Set frmPackage = Nothing
End Sub
