Attribute VB_Name = "modFuncs"
                                        Option Explicit



                                        
'   API call to play wav files
    
    Public Declare Function sndPlaySound Lib _
        "winmm.dll" Alias "sndPlaySoundA" _
        (ByVal lpszSoundName As String, _
        ByVal uFlags As Long) As Long            'sndPlaySound "\path to wav" , 1
'----------------------------------------------

'   To make use of the Encryption .dll
'   Astring = enCrypt ( MyString )
'   AnotherString = deCrypt ( Astring )

Public Sub sBar(Optional ByVal Pan1 As String = "^Available", _
                Optional ByVal Pan2 As String = "Offline")
                                 
    On Error GoTo errHandler:
                                 
    ' Handle the status bar
    frmMain.sb(0).Caption = Pan1 'Panel one
    
    frmMain.sb(1) = Pan2 'Panel two
    
    
    
errHandler:
    Exit Sub
End Sub

Public Function fColor() 'Just used to change colors in the Fake Status Bar When mouse is Over
    
    'Mouse hover events
    
    frmMain.sb(0).ForeColor = &HD0CABD
    
    frmMain.sb(1).ForeColor = &HD0CABD
    
    
End Function
