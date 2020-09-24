VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H0063C7ED&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WIM"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9000
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   266
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   600
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtNumOnline 
      BackColor       =   &H80000003&
      DataSource      =   "dbs"
      Height          =   495
      Left            =   10680
      TabIndex        =   16
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox txtNumAccounts 
      BackColor       =   &H80000003&
      DataSource      =   "dbs"
      Height          =   495
      Left            =   10680
      TabIndex        =   15
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox txtHostName 
      BackColor       =   &H80000003&
      DataSource      =   "dbs"
      Height          =   495
      Left            =   10680
      TabIndex        =   14
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txtLocalPort 
      BackColor       =   &H80000003&
      DataSource      =   "dbs"
      Height          =   495
      Left            =   10680
      TabIndex        =   13
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtAddress 
      BackColor       =   &H80000003&
      DataSource      =   "dbs"
      Height          =   495
      Left            =   10680
      TabIndex        =   12
      Top             =   240
      Width           =   1335
   End
   Begin VB.Data dbs 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   10800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3840
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   10320
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Data db 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   9360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3840
      Width           =   1140
   End
   Begin VB.TextBox tWarn 
      BackColor       =   &H80000000&
      DataSource      =   "db"
      Height          =   495
      Left            =   9360
      TabIndex        =   11
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox tFro 
      BackColor       =   &H80000000&
      DataSource      =   "db"
      Height          =   495
      Left            =   9360
      TabIndex        =   10
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox tDate 
      BackColor       =   &H80000000&
      DataSource      =   "db"
      Height          =   495
      Left            =   9360
      TabIndex        =   9
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox tC 
      BackColor       =   &H80000000&
      DataSource      =   "db"
      Height          =   495
      Left            =   9360
      TabIndex        =   8
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox tPW 
      BackColor       =   &H80000000&
      DataSource      =   "db"
      Height          =   495
      Left            =   9360
      TabIndex        =   7
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox tID 
      BackColor       =   &H80000000&
      DataSource      =   "db"
      Height          =   495
      Left            =   9360
      TabIndex        =   6
      Top             =   240
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock server 
      Index           =   0
      Left            =   3720
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   9457
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   8520
      Picture         =   "frmMain.frx":08CA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   5040
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   9120
      Picture         =   "frmMain.frx":1594
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   5040
      Visible         =   0   'False
      Width           =   540
   End
   Begin MSComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   3615
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3969
            MinWidth        =   3969
            Text            =   "Server State here"
            TextSave        =   "Server State here"
            Key             =   "Server"
            Object.Tag             =   "Server"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3969
            MinWidth        =   3969
            Text            =   "Port"
            TextSave        =   "Port"
            Key             =   "Users"
            Object.Tag             =   "Users"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7673
            MinWidth        =   7673
            Text            =   "Last msg Sent"
            TextSave        =   "Last msg Sent"
            Key             =   "Message"
            Object.Tag             =   "Message"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList il1 
      Left            =   9720
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E5E
            Key             =   "Warn"
            Object.Tag             =   "Warn"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":273A
            Key             =   "Blue"
            Object.Tag             =   "Blue"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3016
            Key             =   "Red"
            Object.Tag             =   "Red"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":38F2
            Key             =   "White"
            Object.Tag             =   "White"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":41CE
            Key             =   "Stop"
            Object.Tag             =   "Stop"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4AAA
            Key             =   "Go"
            Object.Tag             =   "Go"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5786
            Key             =   "Fire"
            Object.Tag             =   "Fire"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6462
            Key             =   "hGlass"
            Object.Tag             =   "hGlass"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":685A
            Key             =   "pSound"
            Object.Tag             =   "pSound"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6C96
            Key             =   "oServer"
            Object.Tag             =   "oServer"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7232
            Key             =   "oDoor"
            Object.Tag             =   "oDoor"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7726
            Key             =   "cDoor"
            Object.Tag             =   "cDoor"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   3615
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   6376
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Time"
         Object.Tag             =   "Time"
         Text            =   "Time Stamp"
         Object.Width           =   3969
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Origin"
         Object.Tag             =   "Origin"
         Text            =   "Origin"
         Object.Width           =   3969
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Message"
         Object.Tag             =   "Message"
         Text            =   "Message / Command"
         Object.Width           =   7673
      EndProperty
   End
   Begin MSWinsockLib.Winsock sck 
      Index           =   0
      Left            =   8880
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   9456
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      ToolTipText     =   "Shut down BLAIM"
      Top             =   4680
      Width           =   600
   End
   Begin VB.CommandButton cmdSnd 
      Caption         =   "&Send"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Sends a Msg. To the client High-Lighted"
      Top             =   4680
      Width           =   600
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSend 
         Caption         =   "&Send File"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuServer 
      Caption         =   "&Server"
      Begin VB.Menu mnuStart 
         Caption         =   "&Start"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuStop 
         Caption         =   "&Stop"
         Enabled         =   0   'False
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuConnections 
         Caption         =   "&Connections"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuData 
         Caption         =   "&Database"
         Shortcut        =   ^D
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
                                    Option Explicit
' Twinb.MyFtp.org

'© Coded by William LeRoy       William_j_leroy@yahoo.com  send me a msg or email me
'**************************************************************************************
' this is far from being finished
' so far i have only spent about 30 minutes on this

    Dim cUser As Integer

    ' These are the declarations for Winsock
    Dim Data1 As String
    
    ' Flag for server
    Dim Isserver As Boolean
    



Private Sub cmdExit_Click()

    Unload frmIM1
    
    Unload Me
    
End Sub

Private Sub cmdSnd_Click()

    ' Open the Form we need to send Msg.
    frmIM1.Show
    
End Sub

Private Sub Form_Load()

    ' Add to the listview some startup text
    AddtoList lv1
    
    ' Add to the Status Bar some startup text
    sBar sb, server(0).LocalIP
    
    ' Set the flag to FALSE for Isserver(0)
    Isserver = False
    
    'Set the database name to point here
    db.DatabaseName = App.Path & "\env\server.mdb"
    
    'The source will be the Client field in the db
    db.RecordSource = "Client"
    
    'Associate the proper datafields with their respectable owners
    tID.DataField = "UserID" ' Sorry Gregg
    
    tPW.DataField = "uPassword"
    
    tC.DataField = "uConnected"
    
    tDate.DataField = "ccDate"
    
    tFro.DataField = "uFrozen"
    
    tWarn.DataField = "uWarnings"
    
    'load the second database
    dbs.DatabaseName = App.Path & "\env\server.mdb"
    
    dbs.RecordSource = "Server"
    
    txtAddress.DataField = "Address"
    
    txtLocalPort.DataField = "LocalPort"
    
    txtHostName.DataField = "HostName"
    
    txtNumAccounts.DataField = "NumAccounts"
    
    txtNumOnline.DataField = "NumOnline"
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Unload Me

End Sub

Private Sub mnuConnections_Click()
    
    frmClients.Show 'Show all connected clients
    
    
End Sub

Private Sub mnuData_Click()
    
    'Show the info stored in the databases
    
    
End Sub

Private Sub mnuStart_Click()
    
    If Isserver = False Then sck(0).Listen:: Isserver = True ' If the server(0) is off turn on
    
    mnuStart.Enabled = False ' So we don't try to ReStart when it is running
    
    mnuStop.Enabled = True ' Now enable the stop menu item
    
    AddtoList lv1, , "Server is On"
             
End Sub

Private Sub mnuStop_Click()
    
    If Isserver = True Then sck(0).Close:: 'sck(1).Close 'Close connections/server
    
    Isserver = False 'Reset flag so we could ReStart if wanted
    
    mnuStart.Enabled = True
    
    mnuStop.Enabled = False
    
    AddtoList lv1, , "Server is Off"
    
End Sub

Private Sub sb_PanelClick(ByVal Panel As MSComctlLib.Panel)
    
    'Just playing
    Panel = InputBox("Please enter the new information", _
                    "Status Bar Change", "Type Here")
                    

End Sub

Private Sub sck_ConnectionRequest(Index As Integer, _
                                  ByVal requestID As Long)

    ' first load new winsock into array
    cUser = cUser + 1
    
    Load sck(cUser)
    
    ' then set the port to 0 for random port
    sck(cUser).LocalPort = 0
    
    ' accept the connection
    sck(cUser).Accept requestID
    
    ' now store what #winsock user is on in .tag of txtIM
    frmIM1.txtIM.Tag = cUser

End Sub

Private Sub sck_DataArrival(Index As Integer, _
                            ByVal bytesTotal As Long)
                            
    ' Store the data as a String to Data1
    sck(cUser).GetData Data1, vbString
    
    'Check the message
    Check_Message Data1
    
    
End Sub



