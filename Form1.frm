VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12060
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   12060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      Height          =   495
      Left            =   8520
      TabIndex        =   13
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtInvervall 
      Height          =   285
      Left            =   5760
      TabIndex        =   8
      Text            =   "1000"
      ToolTipText     =   "Sets the time intervalls for which to querry the server."
      Top             =   840
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   9960
      Top             =   600
   End
   Begin VB.TextBox txtDatabase 
      Height          =   285
      Left            =   8880
      TabIndex        =   6
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   10440
      TabIndex        =   7
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   7320
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   5760
      TabIndex        =   3
      Top             =   240
      Width           =   1455
   End
   Begin VB.ListBox List2 
      Height          =   1185
      Left            =   2880
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   0
      Width           =   2775
   End
   Begin VB.ListBox List1 
      Height          =   1185
      Left            =   0
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   0
      Width           =   2775
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   4048
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   7177
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   7177
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label lblRunning 
      Caption         =   "Running"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7440
      TabIndex        =   14
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Timer Intervall"
      Height          =   255
      Left            =   5760
      TabIndex        =   12
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Database"
      Height          =   255
      Left            =   8880
      TabIndex        =   11
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Server"
      Height          =   255
      Left            =   10440
      TabIndex        =   10
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   255
      Left            =   7320
      TabIndex        =   9
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "User Name"
      Height          =   255
      Left            =   5760
      TabIndex        =   4
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#####################################################################################
'## Author: Peter Marynowski                                                                                                                                                                                                     ##
'## Author: Peter Marynowski                                                                                                                                                                                                     ##
'## Author: Peter Marynowski                                                                                                                                                                                                     ##
'## Author: Peter Marynowski                                                                                                                                                                                                     ##
'## Author: Peter Marynowski                                                                                                                                                                                                     ##
'## Author: Peter Marynowski                                                                                                                                                                                                     ##
'#####################################################################################

Option Explicit
Dim dbCon As ADODB.Connection
Dim dbRs As ADODB.Recordset


Private Sub cmdLogin_Click()
On Error GoTo cmdError
Set dbRs = Nothing
Set dbCon = Nothing
Set dbCon = New ADODB.Connection
Set dbRs = New ADODB.Recordset

        'This connection string may change from server to version of database so you go ahead and change it.
        dbCon.ConnectionString = "Servername=" & txtServer & ";DB=" & txtDatabase & ";UID=" & txtUserName & ";PWD=" & txtPassword & ";driver=Sybase System 11"
        dbCon.Open
        dbRs.ActiveConnection = dbCon
        dbRs.CursorLocation = adUseClient
        dbRs.CursorType = adOpenDynamic
        dbRs.LockType = adLockOptimistic
        dbRs.Open "SELECT name FROM master..syslogins"
        List2.Clear
        While dbRs.EOF = False
            List2.AddItem Trim(dbRs.Fields(0).Value)
            dbRs.MoveNext
        Wend
        Timer1.Enabled = True
        Timer1.Interval = txtInvervall.Text
        Me.Caption = "DBScan"
Exit Sub

cmdError:

Me.Caption = "Error Loggin in........"
Exit Sub

End Sub

Private Sub Form_Load()
Dim iNumber As Integer
Dim sFields() As String
        
        txtUserName.Text = LCase(GetUserLogin)
        sFields = Split("spid, kpid, enginenum, status, suid, hostname, program_name, hostprocess, cmd, cpu, physical_io, memusage, blocked, dbid, uid, gid, tran_name, time_blocked, network_pktsz, fid, execlass, priority, affinity, id, stmtnum, linenum, origsuid,")
        For iNumber = 0 To UBound(sFields())
            List1.AddItem "b." & sFields(iNumber)
        Next iNumber
        ReDim sFields(0)
        
        sFields = Split("suid, status, accdate, totcpu, totio, spacelimit, timelimit, resultlimit, dbname, name, password, language, pwdate, audflags, fullname, srvname,")
        For iNumber = 0 To UBound(sFields())
            List1.AddItem "a." & sFields(iNumber)
        Next iNumber


End Sub

Private Sub Form_Resize()
On Error GoTo errForm
    DataGrid1.Height = Me.Height - DataGrid1.Top - 500
    DataGrid1.Width = Me.Width - 150
Exit Sub
errForm:
Exit Sub
End Sub

Private Sub Timer1_Timer()
On Error GoTo errTimer

Dim bNoNames As Boolean
Dim sSQL  As String
Dim iNumber As Integer
If dbRs.State = adStateOpen Then
    dbRs.Close
End If
bNoNames = False
    sSQL = "SELECT DISTINCT "
    For iNumber = 0 To List1.ListCount - 1
        If List1.Selected(iNumber) Then
            sSQL = sSQL & List1.List(iNumber)
        End If
    Next iNumber
    sSQL = Left(sSQL, Len(sSQL) - 1) & " "
    sSQL = sSQL & "from master..syslogins a, master..sysprocesses b "
    
    For iNumber = 0 To List2.ListCount - 1
        If List2.Selected(iNumber) Then
            If InStr(1, sSQL, "Where", vbTextCompare) = 0 Then
                sSQL = sSQL & "Where a.suid = b.suid "
                sSQL = sSQL & "and (a.name = '" & List2.List(iNumber) & "'"
                bNoNames = True
            Else
                sSQL = sSQL & " or a.name = '" & List2.List(iNumber) & "'"
            End If
        End If
    Next iNumber
    sSQL = sSQL & ")"
    If bNoNames = False Then
        sSQL = sSQL & "Where b.suid = a.suid"
    End If
    dbRs.Open sSQL
    DataGrid1.Refresh
    Set DataGrid1.DataSource = dbRs
    lblRunning.Visible = Not (lblRunning.Visible)
Exit Sub
errTimer:
Exit Sub

End Sub

Private Sub txtInvervall_Change()
On Error GoTo errorrrss
    Timer1.Interval = Val(txtInvervall.Text)
Exit Sub
errorrrss:
Exit Sub
End Sub




















