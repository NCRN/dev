Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Dev_DbArchive
' Level:        Development module
' Version:      1.00
'
' Description:  Db archive related functions & procedures for version control
'
' Source/date:  Bonnie Campbell, September 26, 2018
' Revisions:    BLC - 9/26/2018 - 1.00 - initial version
' =================================

' ---------------------------------
' FUNCTION:     ExportTables
' Description:  export desired tables to CSV format
' Assumptions:  -
' Examples (from Immediate window):
'    ?ExportTables("Forest_Veg_be_9.6.1.9016_MASTER_WORKING.accdb","C:\Projects\TEST_FE\","C:\Projects\TEST_FE\")
' Parameters:   Db - database to export from (string)
'               DbPath - path of database to export from (string)
'               ExportPath - path to export to (string)
'               ExportFormat - format data should be exported as (string, default CSV)
'               IncludeData - whether data should be included (optional, default TRUE)
'               SubsetTables - only include certain tables (string, comma delimited list of tables to include, default = "")
'               IncludeData - whether first row of CSV should be field names (optional, default TRUE)
' Returns:      -
' Throws:       none
' References:
'   Galaxiom, 5/20/2012
'   https://access-programmers.co.uk/forums/showthread.php?t=226668
'   HansUp, 7/9/2013
'   https://stackoverflow.com/questions/17555174/how-to-loop-through-all-tables-in-an-ms-access-db
'   Jay Freedman, 1/26/2012
'   https://answers.microsoft.com/en-us/office/forum/office_2010-word/vba-code-to-access-an-external-database/4af21223-5cd3-4f2a-b773-ea6b354ba095?db=5
'   jsteph, 7/7/2005
'   https://www.tek-tips.com/viewthread.cfm?qid=1088240
'   ADezii, 8/5/2016
'   http://www.utteraccess.com/forum/index.php?showtopic=2038477
' Source/date:  Bonnie Campbell, September 26, 2018
' Adapted:      -
' Revisions:
'   BLC - 9/26/2018 - initial version
' ---------------------------------
Public Function ExportTables(DbName As String, _
                    DbPath As String, _
                    ExportPath As String, _
                    Optional ExportFormat As String = "CSV", _
                    Optional IncludeData As Boolean = True, _
                    Optional SubsetTables As String = "", _
                    Optional IncludeFieldNames As Boolean = True)
On Error GoTo Err_Handler

    Dim ac As Access.Application
    Dim Db As DAO.Database
    Dim DbFullPath As String
    Dim tbls As String
    Dim tdf As TableDef
    Dim tbl As Variant
    Dim tblsArray() As String
    Dim FileName As String
    
    'handle external databases
    'Set ac = New Access.Application
    Set ac = CreateObject("Access.Application")
    
    'target database
    DbFullPath = DbPath & DbName
    Set Db = OpenDatabase(DbFullPath)
    ac.OpenCurrentDatabase (DbFullPath)
    
    'include field names in export
    IncludeFieldNames = True
    
    'identify export tables
    Select Case Len(SubsetTables)
        Case Is = 0     ' all tables
            ' iterate through
            For Each tdf In Db.TableDefs
                If Not (tdf.Name Like "MSys*" Or tdf.Name Like "~*") Then _
                    tbls = tbls & "," & tdf.Name
            Next
           
            'remove starting comma
            tbls = Right(tbls, Len(tbls) - 1)
            
        Case Is > 0     ' subset of tables?
            tbls = SubsetTables
    End Select
    
 Debug.Print tbls
    
    'convert tbls to array
    tblsArray = Split(tbls, ",")
    
    'export
    For Each tbl In tblsArray
    
        'filename
        FileName = ExportPath & tbl & ".csv"
        
 Debug.Print FileName
    
        'transfer data
        ac.DoCmd.TransferText acExportDelim, , tbl, FileName, IncludeFieldNames
        
    Next
    
Exit_Handler:
    'cleanup
    Db.Close
    ac.CloseCurrentDatabase
    Set Db = Nothing
    Set ac = Nothing
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ExportTables[mod_Dev_DbArchive])"
    End Select
    Resume Exit_Handler
End Function