Attribute VB_Name = "modDb"
                                        Option Explicit

    'Functions for controlling the DataBase storing and movement Routines
Public Function Add_New(Data1 As Data)

    'This will open a new section of the database to store more information
    Data1.Recordset.AddNew

End Function

Public Function Pro_Delete(Data1 As Data)

    'This will delete the entire Clints profile stored on the database
    Data1.Recordset.Delete

    Data1.Recordset.MoveNext
    
End Function

Public Function Refresh_Db(Data1 As Data)
    
    'This will refresh the database in case there are multiple connections we want only the
    ' most recent data
    Data1.Refresh

End Function

Public Function Update_Db(Data1 As Data)
    
    'Update a record that has been changed/new or modified
    Data1.UpdateRecord
    
    'Not needed nor used here just for you to play with
    Data1.Recordset.Bookmark = Data1.Recordset.LastModified 'will show last entry you were in

End Function

Public Function Count_Users(Data1 As Data)
    
    'This is where we will store the data received
    Dim Incoming_Data As String
    
    'Count the database to see how many profiles/users we have
    Incoming_Data = Data1.Recordset.RecordCount
    
    
    
End Function

Public Function Search_DB(db As Data)
    
    'Searching a DataBase
    db.RecordSource = " Select UserID, uPassword, uConnected, ccDate, uFrozen, uWarnings From Client Where UserID = 'Asshole'"
    
    db.Refresh
    
    
End Function
