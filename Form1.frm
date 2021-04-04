VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Mail 
   Caption         =   "BulkMailer 1.0"
   ClientHeight    =   5490
   ClientLeft      =   2100
   ClientTop       =   1980
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   7860
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2880
      TabIndex        =   0
      Top             =   690
      Width           =   3495
   End
   Begin VB.TextBox msg 
      Height          =   1935
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   2760
      Width           =   6855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Attachment"
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   4920
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7200
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Select attach file"
      Filter          =   "*.*"
      InitDir         =   "c:\"
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   1575
      Left            =   480
      OleObjectBlob   =   "Form1.frx":0014
      TabIndex        =   1
      Top             =   1080
      Width           =   6855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   5400
      TabIndex        =   6
      Top             =   4920
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "S&end Mails"
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Session On"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   4920
      Width           =   975
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   720
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   120
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "MSG Subject"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1440
      TabIndex        =   8
      Top             =   720
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "BulkMailer 1.0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   2640
      TabIndex        =   7
      Top             =   0
      Width           =   2865
   End
   Begin VB.Menu mnu_exit 
      Caption         =   "E&xit"
   End
   Begin VB.Menu mnu_help 
      Caption         =   "He&lp"
      Begin VB.Menu mnu_ht 
         Caption         =   "Help Topics"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnu_abt 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Mail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ssid
Public cr
Public athfil
Public rs As Recordset

Private Sub Command1_Click()
MAPISession1.NewSession = True
MAPISession1.SignOn
ssid = MAPISession1.SessionID
End Sub

Private Sub Command2_Click()
infotxt = "This software is developed by SRUBS, C-126, 5th Cross, Thillai Nagar, Trichy - 620018, India. Phone. 91 431 765876 Email: srubs@vsnl.com. Visit: www.trichy-online.com"

rs.MoveFirst
MAPIMessages1.SessionID = ssid
Do While Not rs.EOF()

    If Trim(rs.Fields(2)) = "*" Then
        MAPIMessages1.Compose
        MAPIMessages1.RecipDisplayName = rs.Fields(0)
        MAPIMessages1.RecipAddress = rs.Fields(1)
        MAPIMessages1.MsgSubject = Text1.Text
        l = Len(athfil)
        If l > 0 Then
            lp = 1
            For k = 1 To l
                If Mid(athfil, k, 1) = "\" Then
                    lp = k
                End If
            Next
            filnm = Mid(athfil, lp + 1, l - lp)
            MAPIMessages1.AttachmentPathName = athfil
            MAPIMessages1.AttachmentName = filnm
        End If
        msgtxt = "Hi " + rs.Fields(0) + " ," + cr + cr
        msgtxt = msgtxt + msg.Text
        msgtxt = msgtxt + cr + cr + cr + cr + cr + cr + cr + cr + cr + infotxt
        MAPIMessages1.MsgNoteText = msgtxt
        MAPIMessages1.Send
    End If
    rs.MoveNext
Loop

End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command5_Click()
CommonDialog1.Action = 1
athfil = CommonDialog1.FileName
End Sub

Private Sub Form_Load()
KeyPreview = True
file = App.Path + "\contacts.mdb"
cr = Chr$(13) + Chr$(10)
Dim db As Database
'Dim rs As Recordset
Set db = OpenDatabase(file)
Set rs = db.OpenRecordset("select * from contacts")
Set Data1.Recordset = rs
DBGrid1.Refresh
DBGrid1.AllowUpdate = True
DBGrid1.AllowAddNew = True
DBGrid1.AllowDelete = True
DBGrid1.Columns(0).Width = 3000
DBGrid1.Columns(1).Width = 2830
DBGrid1.Columns(2).Width = 500
End Sub

Private Sub mnu_abt_Click()
abttxt = "This software is developed by" + cr
abttxt = abttxt + Space(66) + "SRUBS," + cr
abttxt = abttxt + Space(59) + "C-126, 5th Cross," + cr
abttxt = abttxt + Space(47) + "Thillai Nagar, Trichy - 620018. India" + cr
abttxt = abttxt + Space(54) + "Phone. 91 431 765876" + cr
abttxt = abttxt + Space(33) + "Email: srubs@vsnl.com. Visit: www.trichy-online.com" + cr + cr
abttxt = abttxt + Space(47) + "Send your comments - T. Rajesh"
msg.Text = abttxt

End Sub

Private Sub mnu_exit_Click()
End
End Sub

Private Sub mnu_ht_Click()
hlptxt = ""
hlptxt = hlptxt + Space(8) + "Thank you for using this software. - Double click or Press Esc to clear this text box" + cr
hlptxt = hlptxt + Space(8) + "---------------------------------------------------------------------------------------------------------------------------------" + cr
hlptxt = hlptxt + "To select the records, put * in the select column of the grid. " + "To unselect the records, delete" + cr
hlptxt = hlptxt + " * from the select column of grid and press space bar. Only the selected  rows are processed" + cr
hlptxt = hlptxt + "for the bulk mail. Type body message in the text box. Before sending bulk mails select" + cr
hlptxt = hlptxt + "attachment by Attachment button and it accepts only one attachment." + cr + cr
hlptxt = hlptxt + "To send bulk mails:  * Press Attachment button to select the attachment" + cr
hlptxt = hlptxt + Space(32) + "* Press the Session On button to activate the mailing session" + cr
hlptxt = hlptxt + Space(32) + "* Press Send Mails to send bulk mails" + cr + cr
hlptxt = hlptxt + Space(15) + "F1 - Help F2 - About F3 - Attachment F4 - Session on F5 - Send F6 - Exit" + cr + cr
hlptxt = hlptxt + Space(28) + "For comments and suggestions, contact srubs@vsnl.com"
msg.Text = hlptxt
End Sub

Private Sub msg_DblClick()
msg.Text = ""
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape:
            msg.Text = ""       ' Clear text box
        Case vbKeyF2:
            mnu_abt_Click       ' Display about
        Case vbKeyF3:
            Command5_Click      ' Attachment Select
        Case vbKeyF4:
            Command1_Click      ' Session on
        Case vbKeyF5:
            Command2_Click      ' Send mails
        Case vbKeyF6:
            End                 ' Exit the software
    End Select
End Sub
