Option Compare Database
Option Explicit

Private Sub ImportDataButton_Click()

    On Error GoTo Err_ImportDataButton_Click

    Dim pathName, fileName, sheetName, amazonPathName, amazonFileName, amazonSheetName, dataSQL, amazonSQL As String
    Dim deleteData As Boolean

    'Get data from Main Form
    deleteData = Me.ReplaceDataCheck.Value
    fileName = Me.InputFileNameControl.Value
    sheetName = Me.InputSheetNameControl.Value

    'Drop the BaseData tables if requested
    'This is a faster solution than deleting rows
    If deleteData = True Then
    
        'Drop table
        '----------
        'Setup SQL string for table drop and run it
        dataSQL = "DROP TABLE BaseData;"
        DoCmd.RunSQL dataSQL, False

        'Create table
        '------------
        'Setup SQL string for table creation and run it
        dataSQL = "CREATE TABLE BaseData (" & _
                  "[ID] AUTOINCREMENT PRIMARY KEY, " & _
                  "[ConversationID] TEXT(100) NOT NULL, " & _
                  "[SkillID] TEXT(100) NOT NULL, " & _
                  "[UserID] TEXT(100) NOT NULL, " & _
                  "[ConversationDuration] LONG, " & _
                  "[Rating] SHORT, " & _
                  "[BetaUser] YESNO, " & _
                  "[DateValue] LONG, " & _
                  "[TimeValue] DOUBLE, " & _
                  "[Feedback] YESNO);"
        DoCmd.RunSQL dataSQL, False
        
        'Create Index
        '------------
        'Setup SQL string for index creation and run it
        dataSQL = "CREATE INDEX indexBaseUserID ON BaseData ([UserID]);"
        DoCmd.RunSQL dataSQL, False

    End If
  
    'Import data from Excel
    '----------------------
    pathName = CurrentProject.Path + "\" + fileName
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, "BaseData", pathName, True, sheetName & "!"

    'Update the user data. The sequence of these SQL statements is important
    '-----------------------------------------------------------------------
    'Append new Amazon users to the existing AmazonUser table
    'this table is used as an interim step to make the next SQL statements easier
    dataSQL = "INSERT INTO AmazonUsers ( UserID ) " & _
              "SELECT DISTINCT BaseData.UserID " & _
              "FROM BaseData LEFT JOIN AmazonUsers ON BaseData.UserID = AmazonUsers.UserId " & _
              "WHERE (((BaseData.BetaUser)=True) AND ((AmazonUsers.UserId) Is Null));"
     DoCmd.RunSQL dataSQL, True
   
    'Append new users to User table
    dataSQL = "INSERT INTO [User] ( UserID ) " & _
              "SELECT DISTINCT BaseData.UserID " & _
              "FROM BaseData LEFT JOIN [User] ON BaseData.UserID = User.UserID " & _
              "WHERE (((User.UserID) Is Null));"
    DoCmd.RunSQL dataSQL, True
   
    'Update User table to indicate which users are Amazon users
    dataSQL = "UPDATE [User] INNER JOIN AmazonUsers ON User.UserID = AmazonUsers.UserId SET [User].Amazon = True;"
    DoCmd.RunSQL dataSQL, True
   
    'Let user know data is loaded
    MsgBox "Data loading is complete. New Amazon users have been added to the existing table."

Exit_ImportDataButton_Click:
    Exit Sub

Err_ImportDataButton_Click:
    MsgBox Err.Description
    Resume Exit_ImportDataButton_Click
    
End Sub
