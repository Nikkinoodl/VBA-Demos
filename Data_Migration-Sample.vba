'''
'''Data Migration Script
'''
'''Used to add Item records on 4/24/19
'''
'''Item records from 2018 onwards were fully updated to include correct prices for each item
'''when multiple paintings were purchased in one transaction. Older records have the correct price on
'''the first item and subsequent items are priced at 0. Records for 2018 onwards were manually adjusted
'''based on sales books so they are correctly priced.
'''
'''Artwork sales/orders now has field for orderbook/web store reference number. These were populated for all transactions
'''from 2018 onwards. Older transaction records can be obtained from sales books.
'''
'''Excluded an item category for item records as it felt like too much data input.
'''
'''Charge lines were updated via a separate update query
'''
'''Execution plan:
'''
''' 1) Get all table naming and variable naming consistent
''' 2) Run with all edits and inserts commented out
''' 3) Run with counters that indicate numbers of records that would be affected/created
''' 4) Run with actual data and stop after 1 record
''' 5) Run with actual data and stop after 5 records
''' 6) Actual run
'''
Private Sub DataUpgradeButton_Click()

'Create recordset for artwork table
Dim PBSB As DAO.Database
Dim rst As DAO.Recordset
Dim rstCharges As DAO.Recordset
Dim rstItems As DAO.Recordset
 
 
Set PBSB = CurrentDb
Set rst = PBSB.OpenRecordset("Artwork", dbOpenDynaset)
Set rstCharges = PBSB.OpenRecordset("Charges", dbOpenDynaset)
Set rstItems = PBSB.OpenRecordset("Item", dbOpenDynaset)

'Open artwork records
With rst

.MoveFirst

Do Until .EOF

    'Initialize variables inside loop
    Dim thisPortraitID As Integer
    Dim thisTitle As Variant
    Dim thisSize As Variant
    Dim thisMedium As Variant
    Dim thisCompletedDate As Variant
    Dim thisActualTime As Integer
    Dim thisPrice As Currency
    Dim thisSaleType As Variant
    Dim thisServiceType As Variant
    Dim titleArray() As String
    Dim sizeArray() As String
    Dim i As Integer
    Dim noTitle As Integer
    Dim skips As Integer
    Dim numItems As Integer
    Dim numLoops As Integer
    Dim numProcessed As Integer
    Dim titleSize As Integer
    Dim sizeSize As Integer
    Dim equalLength As Boolean
    
    'Copy record values to variables
    thisPortraitID = .Fields("PortraitID")
    thisTitle = .Fields("Title")
    thisSize = .Fields("Size")
    thisMedium = .Fields("Medium") & " on " & .Fields("Grounds")
    thisCompletedDate = .Fields("CompletionDate")
    thisActualTime = .Fields("ActualTimeHrs")
    thisSaleType = .Fields("SaleType")
    thisServiceType = .Fields("ServiceType")

    'Skip any records that don't have artwork
    If (thisServiceType <> "Commission" And thisServiceType <> "Artwork sale") Then
        skips = skips + 1
        GoTo Next_Artwork
    End If
    
    With rstCharges
       
        'Find the ItemSubtotal line for this artwork in the charges table
        .FindFirst "[PortraitID] = " & thisPortraitID & " And [ChargeType] = 'ItemSubtotal'"
        
        'Skip processing this record if there is no BasePrice charge line -
        'data has already been cleaned and checked so we can assume that it's already been processed
        If .NoMatch Then
            'MsgBox ("No matching ItemSubtotal charge line for this artwork: " & thisPortraitID)
            numProcessed = numProcessed + 1
            GoTo Next_Artwork
        End If
        
        numLoops = numLoops + 1
        
        '************************************
        'Kill loop after processing n records
        'used to run incrementally
        '************************************
        'If numLoops > 5 Then
        '   GoTo End_Script
        'End If
             
        
        'read the amount
        thisPrice = .Fields("Amount")
                             
        'Open for editing and rename the BasePrice charges line to item subtotal
        .Edit
        .Fields("ChargeType") = "ItemSubtotal"
        .Update
        
    End With
    
    'Split title into an array - semicolon has been used consistently as a delimiter
    'Skip processing if title is NULL and count missing titles
    If IsNull(thisTitle) Then
        noTitle = noTitle + 1
        GoTo Next_Artwork
    Else
        titleArray = Split(thisTitle, ";")
    End If
    
    'Handle cases where size is NULL
    If IsNull(thisSize) Then
        thisSize = "N/A"
        ReDim sizeArray(1)
        sizeArray(0) = "N/A"
    Else
        sizeArray = Split(thisSize, ";")
    End If
    
    'If arrays are the same length then size info available for all records
    titleSize = UBound(titleArray) - LBound(titleArray) + 1
    sizeSize = UBound(sizeArray) - LBound(sizeArray) + 1
    
    If titleSize = sizeSize Then
        equalLength = True
    Else
        equalLength = False
    End If
           
    'Loop through array
    For i = LBound(titleArray) To UBound(titleArray)
    
        'Create a new item record for each array element
        Dim newTitle As String
        Dim newSize As String
        
        'Remove leading and trailing spaces
        newTitle = Trim(titleArray(i))
        
        'Use first array element for size if it doesnt have same number of elements as title array
        If equalLength = True Then
            newSize = Trim(sizeArray(i))
        Else
            newSize = Trim(sizeArray(0))
        End If
        
        'Set price to zero after the first item as we are unable to allocate to different records
        'unless we do it manually
        If i > LBound(titleArray) Then
            thisPrice = 0
        End If
        
        'Add a new record to the items table
        '****If we add category field do we have to populate it manually???
        With rstItems
            .AddNew
            !PortraitID = thisPortraitID
            !Title = newTitle
            !Size = newSize
            !Medium = thisMedium
            !completedDate = thisCompletedDate
            !actualTime = thisActualTime
            !Price = thisPrice
            .Update
        End With
    
        numItems = numItems + 1
     
    Next
    
Next_Artwork:
    
    'Goto next artwork record
    .MoveNext
 
 Loop
 
End_Script:
 
End With

'Close all recordsets and database
rst.Close
rstCharges.Close
rstItems.Close
PBSB.Close

MsgBox (skips & " records were skipped")
MsgBox (noTitle & " records were missing a title")
MsgBox (numProcessed & " records were already processed and were skipped")
MsgBox (numLoops & " original artwork records were processed")
MsgBox (numItems & " item records were created")

End Sub
