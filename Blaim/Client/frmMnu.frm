VERSION 5.00
Begin VB.Form frmMnu 
   BackColor       =   &H00B5A89E&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1500
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Line bl 
      Index           =   0
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   213
   End
   Begin VB.Label mItem 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "mnuItem"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1500
   End
End
Attribute VB_Name = "frmMnu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
                                        Option Explicit
                                        
                                        
                                        
Private Sub Form_Load()
    
    'Check to see which flag is up so we know what kind of menu to load
    If mFile = True Then cFile
    
    If mActions = True Then cActions
    
    If mTools = True Then cTools
    
    If mhelp = True Then cHelp
    
    
End Sub

Private Sub Form_LostFocus()
    
    IsMnu = False 'Toggle back to false so we can load if we have to
    
    Unload Me
    
    frmMain.Refresh

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Unload Me
    
End Sub

Private Function cFile()
    
    Dim i As Integer 'Where I will store the # of Items in the menu
    
    mItem(0).Caption = "Sign on" ' Set the first Menu Items' caption
    
    For i = 1 To 4 ' Obvious it's saying that i is a # can be between 1 and 4
        
    Load mItem(i) ' Now load the new Menu Item in the array
    
    Next ' Load the next Item in the Menu Array
    
    With mItem(1) ' Set the Menu Items' properties at run time
    'the reason for this is that we cannot set them at design time
    'they are created at run-time
            
            .Visible = True
            .Top = 20
            .Left = 0
            .Caption = "Sign out"
            .Width = 100
            .Height = 20
            
    End With
    
    With mItem(2)
            
            .Visible = True
            .Top = 40
            .Left = 0
            .Caption = "My Status"
            .Width = 100
            .Height = 20
            
    End With
    
    'Lets check to see if the User is signed on so we can Sort/enable the menu
    If uOn = False Then mItem(1).Enabled = False:: mItem(0).Enabled = True
    
    If uOn = True Then mItem(0).Enabled = False:: mItem(1).Enabled = True
    '## End of MenuItems
    
    '// Begin code for Setting up the Border
    
    ' Now that we have the look of the Menu lets set the height we already have the width
    Me.Height = 60 * 15 ' This simply says the height is 60 pixels X 15 = 900
    
    
    
    For i = 1 To 4 ' This is a For Next LOOP
    
        Load bl(i) 'This will load the new control into the Given Array
    
    Next '## End of For Next LOOP
    
    With bl(0) '// We are creating a border around the form with lines
    
            .Visible = True ' We just loaded this cntrl out of an array so Set visibility
            .X1 = 0 ' Represents the left to right axis of the line
            .X2 = 0
            .Y1 = 0 ' Represents the top to bottom axis of the line
            .Y2 = 60
    End With
    
    With bl(1)
    
            .Visible = True
            .X1 = 99
            .X2 = 99
            .Y1 = 0
            .Y2 = 60
    End With
    
    With bl(2)
            
            .Visible = True
            .X1 = 0
            .X2 = 100
            .Y1 = 0
            .Y2 = 0
    End With
    
    With bl(3)
    
            .Visible = True
            .X1 = 0
            .X2 = 100
            .Y1 = 59
            .Y2 = 59
            
    End With
        
    For i = 0 To 3
    
        bl(i).BorderWidth = 5
        
        bl(i).BorderColor = vbBlue
    
    Next
    
    
End Function

Private Function cActions()

End Function

Private Function cTools()

End Function

Private Function cHelp()

End Function

Private Sub mItem_Click(Index As Integer)
    
    'Check to see which flag is up so we have the commands/nsync with the menu
    If mFile = True Then GoTo oFile:
    
    If mActions = True Then GoTo oActions:
    
    If mTools = True Then GoTo oTools:
    
    If mhelp = True Then GoTo oHelp:
    
oFile: '// Code for Menu: File
        
        Select Case Index
        
                Case 0
                        
                        'Show the Sign-On form
                        'frmSign.Show
                        
                        'Close up the menu
                        frmMnu.Hide
                        
                        uOn = True 'Signal the flag now that we are on
                        
                        
                Case 1
                        
                        'Code for the sign off MenuItem
                        frmMnu.Hide
                        
                        uOn = False 'Signal the flag now that we are off
                        
                        
                Case 2
                        
                        'code for the MyStatus MenuItem
                        frmMnu.Hide
                        
        End Select
        
Exit Sub '## End Menu: File Code
        
oActions: '// Code for Menu: Actions


Exit Sub '## End Menu: Actions Code

oTools: '// Code for Menu: Tools


Exit Sub '## End Menu: Tools Code

oHelp: '// Code for Menu: Help


                        
Exit Sub '## End Menu: Help Code
    
End Sub

Private Sub mItem_MouseMove(Index As Integer, _
                            Button As Integer, _
                            Shift As Integer, _
                            X As Single, _
                            Y As Single)

'This code will highlight the text when the mouse is hover and
'Set the color back when mouse is not hover

    If mFile = True Then GoTo oFile:
    
oFile:
        
        Select Case Index 'Returns the Menu Item that the mouse is over
        
                    Case 0 'Menu Item Array # that Index Returned
                    
                            mItem(0).ForeColor = vbBlue ' When mouse is over
                            mItem(1).ForeColor = &H800000 'All others back to normal
                            mItem(2).ForeColor = &H800000
                            mItem(3).ForeColor = &H800000
                            
                    Case 1
                            
                            mItem(1).ForeColor = vbBlue
                            mItem(0).ForeColor = &H800000
                            mItem(2).ForeColor = &H800000
                            mItem(3).ForeColor = &H800000
                            
                    Case 2
                            
                            mItem(2).ForeColor = vbBlue
                            mItem(0).ForeColor = &H800000
                            mItem(1).ForeColor = &H800000
                            mItem(3).ForeColor = &H800000
                            
                    Case 3
                    
                            mItem(3).ForeColor = vbBlue
                            mItem(0).ForeColor = &H800000
                            mItem(1).ForeColor = &H800000
                            mItem(2).ForeColor = &H800000
                            
        End Select
            
                            
          
End Sub
