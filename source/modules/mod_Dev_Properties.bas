Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Dev_Properties
' Level:        Development module
' Version:      1.01
'
' Description:  Property related functions & procedures for version control
'
' Source/date:  Bonnie Campbell, November 22, 2017
' Revisions:    BLC - 11/22/2017 - 1.00 - initial version
'               BLC - 9/26/2018  - 1.01 - removed debug/test note (doesn't apply)
' =================================

' ---------------------------------
' FUNCTION:     AddDbProperty
' Description:  add custom properties to a database application
' Assumptions:  -
' Examples (from Immediate window):
'    ?AddDbProperty("Copyright Notice", "� 2017 B.Campbell for NCPN")
'    ?CurrentDb.Properties![Copyright Notice]
'       � 2017 B.Campbell for NCPN
'    ?AddDbProperty("Designed & Developed By", "B.Campbell")
'    ?CurrentDb.Properties![Designed & Developed By]
'       B.Campbell
' Parameters:   DbProperty - property to add (string)
'               DbPropertyValue - value for (string)
'               DbPropertyType - type values should be (optional, default DB_TEXT)
'               DbFilename - database to add property to (e.g. "C:\mydb.accdb", optional, default = Current aka CurrentDb)
' Returns:      -
' Throws:       none
' References:
'   Paul Murray, 6/14/1995
'   http://allenbrowne.com/ser-09.html
' Source/date:  Bonnie Campbell, November 22, 2017
' Adapted:      -
' Revisions:
'   BLC - 11/22/2017 - initial version
' ---------------------------------
Public Function AddDbProperty(DbProperty As String, _
                    DbPropertyValue As String, _
                    Optional DbPropertyType As Long = DB_TEXT, _
                    Optional DbFilename As String = "Current")
On Error GoTo Err_Handler

    Dim Db As DAO.Database
    Dim prop As Property
    
    If DbFilename = "Current" Then
        Set Db = DBEngine(0)(0)
    Else
        Set Db = OpenDatabase(DbFilename)
    End If

    'add the property
    Set prop = Db.CreateProperty(DbProperty, DbPropertyType, DbPropertyValue)
    Db.Properties.Append prop
    
Exit_Handler:
    Db.Close
    Set Db = Nothing
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.description, vbCritical, _
            "Error encountered (#" & Err.Number & " - AddDbProperty[mod_Dev_Debug])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     RemoveDbProperty
' Description:  remove custom properties from a database applciation
' Assumptions:  -
' Examples (from Immediate window):
'    ?RemoveDbProperty("Copyright Notice")
' Parameters:   DbProperty - property to add (string)
'               DbFilename - database to add property to (e.g. "C:\mydb.accdb", optional, default = Current aka CurrentDb)
' Returns:      -
' Throws:       none
' References:
'   Paul Murray, 6/14/1995
'   http://allenbrowne.com/ser-09.html
' Source/date:  Bonnie Campbell, November 22, 2017
' Adapted:      -
' Revisions:
'   BLC - 11/22/2017 - initial version
' ---------------------------------
Public Function RemoveDbProperty(DbProperty As String, _
                    Optional DbFilename As String = "Current")
On Error GoTo Err_Handler

    Dim Db As DAO.Database
    
    If DbFilename = "Current" Then
        Set Db = DBEngine(0)(0)
    Else
        Set Db = OpenDatabase(DbFilename)
    End If

    'remove the property
    Db.Properties.Delete DbProperty
    
Exit_Handler:
    Db.Close
    Set Db = Nothing
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.description, vbCritical, _
            "Error encountered (#" & Err.Number & " - RemoveDbProperty[mod_Dev_Debug])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     UpdateDbProperty
' Description:  add custom properties to a database application
' Assumptions:  -
' Examples (from Immediate window):
'    ?UpdateDbProperty("Developer", "B.Campbell for NCPN")
'    ?CurrentDb.Properties![Developer]
'       B.Campbell for NCPN
' Parameters:   DbProperty - property to add (string)
'               DbPropertyValue - value for (string)
'               DbFilename - database to add property to (e.g. "C:\mydb.accdb", optional, default = Current aka CurrentDb)
' Returns:      -
' Throws:       none
' References:
'   Paul Murray, 6/14/1995
'   http://allenbrowne.com/ser-09.html
' Source/date:  Bonnie Campbell, November 22, 2017
' Adapted:      -
' Revisions:
'   BLC - 11/22/2017 - initial version
' ---------------------------------
Public Function UpdateDbProperty(DbProperty As String, _
                    DbPropertyValue As String, _
                    Optional DbFilename As String = "Current")
On Error GoTo Err_Handler

    Dim Db As DAO.Database
    
    If DbFilename = "Current" Then
        Set Db = DBEngine(0)(0)
    Else
        Set Db = OpenDatabase(DbFilename)
    End If

    'add the property
    Db.Properties(DbProperty) = DbPropertyValue
    
Exit_Handler:
    Db.Close
    Set Db = Nothing
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.description, vbCritical, _
            "Error encountered (#" & Err.Number & " - UpdateDbProperty[mod_Dev_Debug])"
    End Select
    Resume Exit_Handler
End Function