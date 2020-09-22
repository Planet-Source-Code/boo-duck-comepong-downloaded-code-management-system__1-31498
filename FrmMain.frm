VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmMain 
   Caption         =   "PSC Management System 1.0"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10425
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   10425
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command10 
      Caption         =   "Exit"
      Height          =   375
      Left            =   8880
      TabIndex        =   38
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "+"
      Height          =   350
      Left            =   9960
      TabIndex        =   29
      Top             =   720
      Width           =   255
   End
   Begin VB.Frame Frame2 
      Height          =   3015
      Left            =   480
      TabIndex        =   20
      Top             =   2760
      Visible         =   0   'False
      Width           =   6375
      Begin VB.CommandButton Command8 
         Caption         =   "Save"
         Height          =   375
         Left            =   5040
         TabIndex        =   31
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Browse...."
         Height          =   345
         Left            =   5040
         TabIndex        =   30
         Top             =   840
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "FrmMain.frx":0442
         Height          =   315
         Left            =   1440
         TabIndex        =   28
         Top             =   2400
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "ProType"
         Text            =   "[All Group]"
      End
      Begin VB.TextBox Text7 
         Height          =   975
         Left            =   1440
         TabIndex        =   26
         Top             =   1320
         Width           =   4695
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   1440
         TabIndex        =   23
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   1440
         TabIndex        =   22
         Top             =   360
         Width           =   4695
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Group"
         Height          =   255
         Left            =   480
         TabIndex        =   27
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Description"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "File Path"
         Height          =   255
         Left            =   480
         TabIndex        =   24
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "File Name"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.TextBox Text3 
      Height          =   330
      Left            =   3120
      TabIndex        =   18
      Top             =   6720
      Width           =   2655
   End
   Begin VB.OptionButton Option2 
      Caption         =   "File Description"
      Height          =   255
      Left            =   1200
      TabIndex        =   17
      Top             =   6720
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "File Name"
      Height          =   255
      Left            =   1200
      TabIndex        =   16
      Top             =   6480
      Value           =   -1  'True
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc AdoType 
      Height          =   495
      Left            =   960
      Top             =   1200
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Psc\PSC.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Psc\PSC.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from Type"
      Caption         =   "AdoType"
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
   Begin VB.CommandButton Command2 
      Caption         =   "Add"
      Height          =   375
      Left            =   7560
      TabIndex        =   2
      Top             =   6720
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc AdoFile 
      Height          =   495
      Left            =   3240
      Top             =   1320
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Psc\PSC.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Psc\PSC.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from FileProfile"
      Caption         =   "AdoFile"
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
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   6720
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FrmMain.frx":0458
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   10610
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "FileName"
         Caption         =   "File Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "FileDesc"
         Caption         =   "File Description"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4500.284
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Description"
      Height          =   5055
      Left            =   7440
      TabIndex        =   5
      Top             =   1080
      Width           =   2775
      Begin VB.TextBox TxtIndex 
         Height          =   285
         Left            =   1800
         TabIndex        =   39
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Delete"
         Height          =   375
         Left            =   1440
         TabIndex        =   37
         Top             =   4560
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Save"
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   4560
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   1725
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   2760
         Width           =   2535
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   525
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   2535
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Browse...."
         Height          =   345
         Left            =   1560
         TabIndex        =   6
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "File Description"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label5 
         Height          =   255
         Left            =   1200
         TabIndex        =   12
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Date Load :"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "File Name"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "File Path (No Space Please)"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   2535
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   7440
      TabIndex        =   32
      Top             =   -120
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox Text8 
         Height          =   315
         Left            =   0
         TabIndex        =   34
         Top             =   480
         Width           =   2775
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Save"
         Height          =   350
         Left            =   1200
         TabIndex        =   33
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Add File Group Name"
         Height          =   375
         Left            =   0
         TabIndex        =   35
         Top             =   240
         Width           =   1815
      End
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "FrmMain.frx":046E
      Height          =   315
      Left            =   7440
      TabIndex        =   4
      Top             =   360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "ProType"
      Text            =   "[All Group]"
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "PSCMS1.0 by Nizam"
      Height          =   255
      Left            =   7560
      TabIndex        =   41
      Top             =   720
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Caption         =   "total files"
      Height          =   255
      Left            =   5040
      TabIndex        =   40
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Label Label8 
      Caption         =   "Enter Text"
      Height          =   255
      Left            =   3120
      TabIndex        =   19
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Search"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   6480
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Sort By Group"
      Height          =   255
      Left            =   7440
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim keytrue As Boolean
Public Sub TotalFiles()

With AdoFile
    .Refresh
    .Recordset.MoveFirst
    .Recordset.MoveLast
    Label14.Caption = .Recordset.RecordCount & " file(s)"

End With

End Sub


Private Sub Command1_Click()

Shell "VB6.EXE " + Trim(AdoFile.Recordset("FilePath")), vbMaximizedFocus

End Sub

Private Sub Command10_Click()
Unload Me
End
End Sub

Private Sub Command2_Click()
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Frame2.Visible = Not Frame2.Visible
End Sub

Private Sub Command3_Click()
Dim response, Msg, Style, Title

If Text2.Text = "" Then Exit Sub

With AdoFile
    .RecordSource = "Select * from FileProfile where FileIndex='" & Trim(TxtIndex.Text) & "'"
    .Refresh

    Msg = "Are You Sure to Delete File Name : " & Trim(Text2.Text)
    Style = vbYesNo + vbCritical
    Title = "Delete Confirmation"
    
     
    response = MsgBox(Msg, Style, Title)
    
    If response = vbYes Then
        .Recordset.Delete
        MsgBox ("Record deleted Successfully !")
    End If
    
    .RecordSource = "Select * from FileProfile"
    
    Call TotalFiles

Text2.Text = ""
Label5.Caption = ""
Text4.Text = ""
Text1.Text = ""
TxtIndex.Text = ""
    
End With
End Sub

Private Sub Command4_Click()
FlagBrowse = 1
FrmBrowse.Show 1
End Sub

Public Sub Command5_Click()
Dim i As Integer

If Text2.Text = "" Then Exit Sub

With AdoFile

    .RecordSource = "Select * from FileProfile where FileIndex='" & Trim(TxtIndex.Text) & "'"
    .Refresh
               
    .Recordset("FileName") = Trim(Text2.Text)
    .Recordset("FileDate") = Trim(Label5.Caption)
    .Recordset("FilePath") = Trim(Text4.Text)
    .Recordset("FileDesc") = Trim(Text1.Text)
    .Recordset("FileIndex") = Trim(TxtIndex)
    .Recordset.Update
    
    For i = 0 To 2
        .RecordSource = "Select * from FileProfile"
        .Refresh
    Next i

End With
    
End Sub

Private Sub Command6_Click()
Frame3.Visible = Not Frame3.Visible
End Sub

Private Sub Command7_Click()
FlagBrowse = 2
FrmBrowse.Show 1
End Sub

Private Sub Command8_Click()
Dim i As Integer
Dim TempIndex As Integer

If Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Then
    MsgBox "Please Fill In The Blank(s)", vbInformation, "PSC Management System"
    Exit Sub
End If

With AdoFile

    .RecordSource = "Select * from FileProfile order by FileIndex"
    .Refresh
    .Recordset.MoveLast
    TempIndex = .Recordset("FileIndex")
    .Refresh
    
    .Recordset.AddNew
    .Recordset("FileName") = Trim(Text5.Text)
    .Recordset("FileDate") = Format(Now, "dd/MM/yyyy")
    .Recordset("FilePath") = Trim(Text6.Text)
    .Recordset("FileDesc") = Trim(Text7.Text)
    .Recordset("FileGroup") = Trim(DataCombo2.Text)
    .Recordset("FileIndex") = Trim(CInt(TempIndex) + 1)
    .Recordset.Update
    
    For i = 0 To 2
        .RecordSource = "Select * from FileProfile"
        .Refresh
    Next i

End With
    

Frame2.Visible = False
End Sub

Private Sub Command9_Click()

If Text8.Text = "" Then
    MsgBox "Please Enter File Group Name", vbInformation, "PSC Management System"
    Exit Sub
End If

With AdoType
    .RecordSource = "Select * from Type"
    .Refresh
    With .Recordset
        .AddNew
        !ProType = Trim(Text8.Text)
        .Update
    End With

End With

Frame3.Visible = False

End Sub

Private Sub DataCombo1_Click(Area As Integer)

    AdoFile.RecordSource = "Select * from FileProfile where FileGroup='" & Trim(DataCombo1.Text) & "'"
    AdoFile.Refresh

If DataCombo1.Text = "[All Group]" Then

    AdoFile.RecordSource = "Select * from FileProfile"
    AdoFile.Refresh

End If

Call TotalFiles
    
End Sub

Private Sub DataGrid1_Click()

Text2.Text = AdoFile.Recordset("FileName")
Label5.Caption = AdoFile.Recordset("FileDate")
Text4.Text = AdoFile.Recordset("FilePath")
Text1.Text = AdoFile.Recordset("FileDesc")
TxtIndex.Text = AdoFile.Recordset("FileIndex")

End Sub

Private Sub Form_DblClick()
Label15.Visible = Not Label15.Visible
End Sub

Private Sub Form_Load()
FlagBrowse = 0
AdoFile.RecordSource = "SELECT * FROM FileProfile"
Call TotalFiles

End Sub

Private Sub Text3_Change()
On Error Resume Next
If keytrue = True Then Exit Sub

Dim temps$

temps = "'" & Trim(Text3.Text) & "%' "

If Option1.Value = True Then
    
    AdoFile.RecordSource = "SELECT * FROM FileProfile WHERE " _
                          & "FileName Like " _
                          & temps _
                          & "ORDER BY FileName;"
                         
Else
    AdoFile.RecordSource = "SELECT * FROM FileProfile WHERE " _
                          & "FileDesc Like " _
                          & temps _
                          & "ORDER BY FileDesc;"
                          
End If

Call TotalFiles

End Sub





