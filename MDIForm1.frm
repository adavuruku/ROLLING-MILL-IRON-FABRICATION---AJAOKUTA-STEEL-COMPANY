VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00404000&
   Caption         =   "ROLLIN STAND CALCULATION SYSTEM"
   ClientHeight    =   3330
   ClientLeft      =   120
   ClientTop       =   150
   ClientWidth     =   9105
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      AutoSize        =   -1  'True
      Height          =   15615
      Left            =   0
      Picture         =   "MDIForm1.frx":0000
      ScaleHeight     =   15555
      ScaleWidth      =   9045
      TabIndex        =   0
      Top             =   0
      Width           =   9105
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   7440
         Top             =   1800
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\MATHS\maths.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\MATHS\maths.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
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
      Begin VB.Timer Timer2 
         Interval        =   2000
         Left            =   2640
         Top             =   600
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1920
         Top             =   480
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   492
         ImageHeight     =   329
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   9
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":82075
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":90E00
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":94553
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":975CA
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":9AC55
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":9EA80
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":BFF45
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":141FCA
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":1682D5
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Timer Timer1 
         Interval        =   50
         Left            =   1440
         Top             =   600
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         Caption         =   "Project Design And Present  By:   Isezuo, Lawrence O."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   855
         Left            =   14520
         TabIndex        =   2
         Top             =   360
         Width           =   5295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   735
         Left            =   14520
         TabIndex        =   1
         Top             =   5160
         Width           =   5295
      End
      Begin VB.Image Image1 
         Height          =   3930
         Left            =   14520
         Picture         =   "MDIForm1.frx":1ABFFF
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   5295
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000080&
         Height          =   6015
         Left            =   14160
         TabIndex        =   3
         Top             =   120
         Width           =   5895
      End
   End
   Begin VB.Menu mnucalculate 
      Caption         =   "PERFORM CALCULATION"
   End
   Begin VB.Menu mnuadjustconstants 
      Caption         =   "ADJUST CONSTANTS"
   End
   Begin VB.Menu mnuprintreport 
      Caption         =   "PRINT REPORT"
      Begin VB.Menu rptone 
         Caption         =   "REPORT ONE"
      End
      Begin VB.Menu rpttwo 
         Caption         =   "REPORT TWO"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim ICOUNTER As Integer

Public Sub update_database()
Dim adoconn As New ADODB.Connection
Dim rs As New ADODB.Recordset
If adoconn.State = adStateOpen Then adoconn.Close
adoconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\MATHS\maths.mdb;Persist Security Info=False"
If rs.State = adStateOpen Then rs.Close
    
    str1 = "44e958cdd7ef3403022d16023ce0d1ef25ebcbbe"

    rs.Open "SELECT * FROM access", adoconn, adOpenDynamic, adLockOptimistic
    rs!user = "Program has Expired"
    rs.Update

adoconn.Close
Set adoconn = Nothing
    
    MsgBox "This program has already expired ...you have to Complete the Payment before Using it....Thanks !!", vbOKOnly + vbInformation
    
    'close the program
    Unload Me
End Sub
Private Sub MDIForm_Load()

'Dim adoconn As New ADODB.Connection
'Dim rs As New ADODB.Recordset
'Dim rs1 As New ADODB.Recordset
'Dim str1 As String

'If adoconn.State = adStateOpen Then adoconn.Close
'adoconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\MATHS\maths.mdb;Persist Security Info=False"
'If rs1.State = adStateOpen Then rs1.Close

''adoconn.CursorLocation = adUseClient
'str1 = "SELECT * FROM access"
 '   If rs1.State = adStateOpen Then rs1.Close
  '  rs1.Open str1, adoconn, adOpenDynamic, adLockOptimistic
    
   ' exp_date = rs1!Compare
    'computer_date = Format(Now, "dd mm yyyy")
    
    'ff = CDate(exp_date)
    'ff2 = CDate(computer_date)
    'If ff2 > ff Then
    
     '       'update db with a wrong password
      '      update_database
        
    'Else
        
     '   'check if no of usage is over
    
      '  no_of_usage = rs1!usage
       ' If no_of_usage > 38 Then
        
        '    'update db with a wrong password
         '   update_database
            
        'Else
         '   If rs1!user <> "44e958cdd7ef3403022d16023ce0d1ef25ebcbbe" Then
            
          '      'update db with a wrong password
           '     update_database
            
            'End If
        'End If
   ' End If
    
''increase the usage no
'If rs.State = adStateOpen Then rs.Close
'str1 = "44e958cdd7ef3403022d16023ce0d1ef25ebcbbe"
'rs.Open "SELECT * FROM access", adoconn, adOpenDynamic, adLockOptimistic
'no_used = rs!usage
'no_used = no_used + 1
'rs!usage = no_used
'rs.Update


 '   'Analysis of using date time
    
  '  'STATUS2 = Format(Now, "dd mm yyyy")
 ' '  ff2 = CDate(STATUS2)
    
   '' dd = "13 01 2015"
   ' 'ff = CDate(dd)
   ' 'If ff = ff2 Then
   '  '   MsgBox "equal"
   ' 'Else
   '  '   MsgBox "not equal"
   ' 'End If
    ''qq = DateDiff("d", start date, end date)
    ''qq = DateDiff("d", ff2, ff)
   ' 'MsgBox qq
    
    
    ''Set DataGrid1.DataSource = rs1



End Sub

Private Sub mnuadjustconstants_Click()
Form2.Show
End Sub

Private Sub mnucalculate_Click()
Form1.Show
End Sub

Private Sub rptone_Click()
'display report one
Dim ACCESSAPP As Access.Application
Set APPACCESS = New Access.Application
Set APPACCESS = CreateObject("ACCESS.APPLICATION")
APPACCESS.OpenCurrentDatabase ("C:\Maths\maths.mdb")
APPACCESS.DoCmd.OpenReport "Report_One", acViewPreview
APPACCESS.Visible = True
End Sub

Private Sub rpttwo_Click()
'display report two
Dim ACCESSAPP As Access.Application
Set APPACCESS = New Access.Application
Set APPACCESS = CreateObject("ACCESS.APPLICATION")
APPACCESS.OpenCurrentDatabase ("C:\Maths\maths.mdb")
APPACCESS.DoCmd.OpenReport "Report_Two", acViewPreview
APPACCESS.Visible = True
End Sub

Private Sub Timer1_Timer()
Label1.Caption = Format(Now, "ddd dd mmm, yyyy")
Label1.Caption = Label1.Caption & "  -  " & Format(Now, "hh:mm:ss: AM/PM")
End Sub

Private Sub Timer2_Timer()
ICOUNTER = ICOUNTER + 1
    If ICOUNTER = 1 Then
        Image1.Picture = ImageList1.ListImages((CInt(ICOUNTER))).Picture
    ElseIf ICOUNTER = 2 Then
        Image1.Picture = ImageList1.ListImages((CInt(ICOUNTER))).Picture
    ElseIf ICOUNTER = 3 Then
        Image1.Picture = ImageList1.ListImages((CInt(ICOUNTER))).Picture
    ElseIf ICOUNTER = 4 Then
        Image1.Picture = ImageList1.ListImages((CInt(ICOUNTER))).Picture
    ElseIf ICOUNTER = 5 Then
        Image1.Picture = ImageList1.ListImages((CInt(ICOUNTER))).Picture
    ElseIf ICOUNTER = 6 Then
        Image1.Picture = ImageList1.ListImages((CInt(ICOUNTER))).Picture
    ElseIf ICOUNTER = 8 Then
        Image1.Picture = ImageList1.ListImages((CInt(ICOUNTER))).Picture
    ElseIf ICOUNTER = 9 Then
        Image1.Picture = ImageList1.ListImages((CInt(ICOUNTER))).Picture
        ICOUNTER = 0
    End If
End Sub
