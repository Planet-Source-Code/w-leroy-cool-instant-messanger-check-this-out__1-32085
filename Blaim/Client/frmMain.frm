VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVBUTTONS.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{60C994C2-D37E-48F1-891B-198A274B28E5}#1.0#0"; "ADM.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00B5A89E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cool Instant Messanger"
   ClientHeight    =   5430
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   3300
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   362
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   220
   StartUpPosition =   2  'CenterScreen
   Begin adm.MouseControl mc 
      Left            =   2160
      Top             =   2040
      _ExtentX        =   953
      _ExtentY        =   953
   End
   Begin LVbuttons.LaVolpeButton mnuB 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   661
      BTYPE           =   7
      TX              =   "&File"
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
      BCOL            =   13683389
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmMain.frx":08CA
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   2
      IconSize        =   2
      SHOWF           =   0   'False
      BSTYLE          =   0
   End
   Begin MSComctlLib.ImageList i1 
      Left            =   3600
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   15
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0B4F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0DB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":100A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1197
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1668
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B37
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1DA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F2F
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":20B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2322
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2592
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":27F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A53
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2CAD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TView 
      Height          =   3630
      Left            =   150
      TabIndex        =   2
      Top             =   1050
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   6403
      _Version        =   393217
      Style           =   1
      ImageList       =   "i1"
      Appearance      =   0
   End
   Begin LVbuttons.LaVolpeButton cmd 
      Height          =   375
      Index           =   1
      Left            =   1650
      TabIndex        =   1
      Top             =   4680
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Msg"
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
      BCOL            =   13683389
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   11905182
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmMain.frx":2EF1
      ALIGN           =   1
      IMGLST          =   "ImageList1"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   2
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton cmd 
      Height          =   375
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   4680
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Add"
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
      BCOL            =   13683389
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   11905182
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmMain.frx":2F0D
      ALIGN           =   1
      IMGLST          =   "ImageList1"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   2
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton mnuB 
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   6
      Top             =   0
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   661
      BTYPE           =   7
      TX              =   "&Actions"
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
      BCOL            =   13683389
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmMain.frx":2F29
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   2
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton mnuB 
      Height          =   375
      Index           =   2
      Left            =   1350
      TabIndex        =   7
      Top             =   0
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   661
      BTYPE           =   7
      TX              =   "&Tools"
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
      BCOL            =   13683389
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmMain.frx":2F45
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   2
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton mnuB 
      Height          =   375
      Index           =   3
      Left            =   2700
      TabIndex        =   8
      Top             =   0
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   661
      BTYPE           =   7
      TX              =   "&Help"
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
      BCOL            =   13683389
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmMain.frx":2F61
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   2
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin VB.Label Label1 
      BackColor       =   &H00D0CABD&
      Height          =   375
      Left            =   1950
      TabIndex        =   9
      Top             =   0
      Width           =   750
   End
   Begin VB.Label sb 
      Alignment       =   2  'Center
      BackColor       =   &H00B5A89E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D0CABD&
      Height          =   375
      Index           =   1
      Left            =   1650
      TabIndex        =   4
      Top             =   5055
      Width           =   1500
   End
   Begin VB.Label sb 
      Alignment       =   2  'Center
      BackColor       =   &H00B5A89E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D0CABD&
      Height          =   375
      Index           =   0
      Left            =   150
      TabIndex        =   3
      Top             =   5055
      Width           =   1500
   End
   Begin VB.Image Image1 
      Height          =   675
      Left            =   0
      Picture         =   "frmMain.frx":2F7D
      Stretch         =   -1  'True
      Top             =   375
      Width           =   3300
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
                                        Option Explicit
                                    




Private Sub cmd_Click(Index As Integer)

 sBar


End Sub



Private Sub cmd_MouseMove(Index As Integer, _
                          Button As Integer, _
                          Shift As Integer, _
                          X As Single, _
                          Y As Single)
    
    Select Case Index
            
            Case 0
                    
                    fColor
            
            Case 1
                    
                    fColor
            
    End Select
    
End Sub

Private Sub Form_Load()
    
    uOn = False 'Toggle the flag for the menu says no u are not signed on yet
    
    Add_Tree "Test", , 17
    Add_Tree "Testing", 1, 17
    Add_Tree "TEST What", 2, 16, True
    
   TView.Nodes(2).Expanded = True
   
   'refersh the form so it dont mess up
   
    TView.Refresh
    
    frmMain.Refresh
    
  
    
End Sub

Public Function Add_Tree(Name As String, _
                         Optional Section As Integer, _
                         Optional Image As Integer = 1, _
                         Optional IsChild As Boolean = False)
    
   If Section < 1 Then GoTo FirstNode:
   
   If Section >= 1 Then GoTo NextNode:
   
FirstNode:
    
    TView.Nodes.Add , , , Name, Image
    
    Exit Function
    
NextNode:

    If IsChild = True Then
        
        TView.Nodes.Add Section, tvwChild, , Name, Image
    
    Else
    
        TView.Nodes.Add Section, , , Name, Image
        
    End If
    
End Function

Private Sub Form_MouseMove(Button As Integer, _
                           Shift As Integer, _
                           X As Single, _
                           Y As Single)
    
    fColor
    
End Sub



Private Sub Form_Unload(Cancel As Integer)
    
    'Unload everything
    Unload frmMnu 'Just incase the menu does not close if the user opened then exited app
    Unload Me
    
End Sub

Private Sub mnuB_Click(Index As Integer)
    
    Select Case Index
    
           Case 0
                    
                    mFile = True
                    
                    mActions = False
                    mTools = False
                    mhelp = False
           
           Case 1
                    
                    mActions = True
                    
                    mFile = False
                    mTools = False
                    mhelp = False
                    
           Case 2
           
                    mTools = True
                    
                    mFile = False
                    mActions = False
                    mhelp = False
                    
           Case 3
           
                    mhelp = True
                    
                    mFile = False
                    mActions = False
                    mTools = False
                    
    End Select
    
    If IsMnu = False Then frmMnu.Show
                    
                    frmMnu.Top = mc.mouseY * 15 + 340
                    
                    frmMnu.Left = mc.mouseX * 15 - 300
                    
    
    IsMnu = True
    
    
End Sub

Private Sub sb_MouseMove(Index As Integer, _
                         Button As Integer, _
                         Shift As Integer, _
                         X As Single, _
                         Y As Single)
    
    Select Case Index
            
            Case 0
                    
                    sb(0).ForeColor = vbBlue
                    
                    sb(1).ForeColor = &HD0CABD
            
            Case 1
                    
                    sb(1).ForeColor = vbBlue
                    
                    sb(0).ForeColor = &HD0CABD
            
    End Select
    
    
End Sub



Private Sub TView_MouseMove(Button As Integer, _
                            Shift As Integer, _
                            X As Single, _
                            Y As Single)

    fColor

End Sub
