Option Explicit
Option Compare Text
'''
'''
Private Const REFERENCELENGTH As Integer = 6
'''
'''
''' Runs after form loaded
'''
Private Sub Form_Open(Cancel As Integer)
   DoCmd.Maximize
End Sub
'''
''' Ensure data is up-to-date when form is activated
'''
Private Sub Form_Activate()
   Me.Refresh
End Sub
'''
''' Runs when customer info is updated
'''
Private Sub CustomerID_Change()
    Me.Refresh
End Sub
'''
''' Runs when cursor leaves the Charges subform
'''
Private Sub Charges_Exit(Cancel As Integer)

    'Make sure that sale isn't already completed
    If [Completed] = True Then
        'No need for a message in this case
        GoTo Exit_ChargesExit
    Else
        'Update the totals field
        Call UpdateTotal
    End If

Exit_ChargesExit:
End Sub
'''
''' Update sales total when field is double clicked
'''
Private Sub FinalRetailPrice_DblClick(Cancel As Integer)

    'Make sure that sale isn't already completed
    If [Completed] = True Then
        MsgBox ("This sale has already been completed")
        GoTo Exit_FinalRetailPrice_DblClick
    End If
    
    Call UpdateTotal
    
Exit_FinalRetailPrice_DblClick:
End Sub
'''
''' Runs when active artwork record is changed
'''
Private Sub Form_Current()

    'A bunch of stuff we need to do for new records. It's easier to set the
    'default values for some fields this way.
    '
    'The goal is to make data entry as simple as possible especially when recording sales after doing
    'art fairs. We set defaults to be the same as the most recent sale.
    '
    If Me.NewRecord Then
    
        'Use DAO type recordsets
        Dim PBSB As Database
        Dim rst As Recordset
    
         Set PBSB = CurrentDb
    
        'Open Artwork table read only
        Set rst = PBSB.OpenRecordset("Artwork", dbOpenSnapshot)
        Dim oldRefNo As String
        Dim oldDate As Date
        Dim oldSaleType As String
        Dim oldSalesLocation As String
        Dim oldServiceType As String
        Dim thisTaxRate As Double
        Dim thisLocCode As String
    
        'Populate the Reference Number field with a default that's equal to reference number + 1 on the
        'last Artwork record (assuming reference number can be converted to integer)
        With rst
            .MoveLast
            oldDate = !OrderDate
            oldSaleType = !SaleType
            oldSalesLocation = !SalesLocation
            oldServiceType = !ServiceType
            oldRefNo = !ReferenceNo
        End With
     
        If Not IsNull(oldRefNo) Then
            [ReferenceNo] = NewDefaultReference(oldRefNo)
        End If
    
        If Not IsNull(oldDate) Then
            [OrderDate] = oldDate
        End If
            
        [SaleType] = oldSaleType
        [SalesLocation] = oldSalesLocation
        [ServiceType] = oldServiceType
            
       'The default tax rate and location code is provided by a lookup on SalesLocation table
        Set rst = PBSB.OpenRecordset("SalesLocation", dbOpenSnapshot)
        With rst
            If .recordCount > 0 Then
                
                .MoveFirst
                .FindFirst "[LocationName] = ' & oldSalesLocation & '"

                thisTaxRate = !TaxRate
                
                If Not IsNull(!LocCode) Then
                    thisLocCode = !LocCode
                Else
                    thisLocCode = ""
                End If
               
                'Return to first record
                .MoveFirst
        
             End If
        
        End With
            
        Me.[TaxRate] = thisTaxRate
        Me.[LocationCode] = thisLocCode
        
            
       ' Populate CustomerID if Customer Form is open and there is no existing CustomerID
        If CurrentProject.AllForms("Customer Management").IsLoaded Then
            
            If [CustomerID] = 0 And Forms![Customer Management]![CustomerID] > 0 Then
                [CustomerID] = Forms![Customer Management]![CustomerID]
            End If
            
        End If
    
    Else
    
        'Prevent changes to sales location for completed records
        Me.SalesLocation.Locked = Me.Completed
        
    End If

   ' Populate CustomerName if CustomerID is already known
   ' otherwise leave it blank
   If [CustomerID] > 0 Then
      [CustomerName] = DLookup("[CustomerName]", "[Customer]", "[CustomerID]=[Forms]![Artwork]![CustomerID]")
   Else
      [CustomerName] = ""
   End If
   
End Sub
'''
''' Opens payment form
'''
Private Sub PaymentCommand_Click()
On Error GoTo Err_PaymentCommand_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Payment Information"
    
    stLinkCriteria = "[PortraitID]=" & Me![PortraitID]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_PaymentCommand_Click:
    Exit Sub

Err_PaymentCommand_Click:
    MsgBox Err.Description
    Resume Exit_PaymentCommand_Click
    
End Sub
'''
''' Closes this form
'''
Private Sub CloseCommand_Click()
On Error GoTo Err_CloseCommand_Click

    DoCmd.Close

Exit_CloseCommand_Click:
    Exit Sub

Err_CloseCommand_Click:
    MsgBox Err.Description
    Resume Exit_CloseCommand_Click
    
End Sub
'''
''' Opens the invoice report
'''
Private Sub InvoiceCommand_Click()
On Error GoTo Err_InvoiceCommand_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Invoice"
    stLinkCriteria = "[PortraitID]=" & Me![PortraitID]
    DoCmd.OpenReport stDocName, acPreview, , stLinkCriteria

Exit_InvoiceCommand_Click:
    Exit Sub

Err_InvoiceCommand_Click:
    MsgBox Err.Description
    Resume Exit_InvoiceCommand_Click
    
End Sub
'''
''' Opens the order confirmation report
'''
Private Sub OrderConfirmationCommand_Click()
On Error GoTo Err_OrderConfirmationCommand_Click


    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "OrderConfirmation"
    stLinkCriteria = "[PortraitID]=" & Me![PortraitID]
    
    'Check that the totals are correct before order confirmation is created
    If [Completed] = False Then
        Call ValidateTotals
    End If
    
    'Open report
    DoCmd.OpenReport stDocName, acPreview, , stLinkCriteria

Exit_OrderConfirmationCommand_Click:
    Exit Sub

Err_OrderConfirmationCommand_Click:
    MsgBox Err.Description
    Resume Exit_OrderConfirmationCommand_Click
End Sub
'''
''' Creates an invoice for this sale
'''
Private Sub CreateInvoiceCommand_Click()
On Error GoTo Err_CreateInvoiceCommand_Click


    If GetMsgResponse("Do you want to create an invoice for this portrait ?") = vbYes Then
    
        
       'Warnings and alerts off
       DoCmd.SetWarnings False
    
        'Get the current record
        Dim record As Variant
        record = Screen.ActiveForm.CurrentRecord
    
       'Check that the totals are correct before invoice is created
        If [Completed] = False Then
            Call ValidateTotals
        End If
   
       'Save the artwork record and ensure we're working with latest data
       DoCmd.Save
       DoCmd.Requery
       DoCmd.GoToRecord acDataForm, "Artwork", acGoTo, record

       DoCmd.RunSQL "INSERT INTO Invoice ( InvoiceDate, CustomerID, PortraitID, CustomerName, Zip, " _
                 & "[Business Name], Address1, Address2, City, State, Title, Medium, Grounds, " _
                 & "FinalRetailPrice, SalesTax, TaxRate, OrderDate, CompletionDate, [Size], PortraitLandscape) " _
                 & "SELECT Now() AS Expr1, Customer.CustomerID, Artwork.PortraitID, Customer.CustomerName, " _
                 & "Customer.Zip, Customer.[Business Name], Customer.Address1, Customer.Address2, Customer.City, " _
                 & "Customer.State, Artwork.Title, Artwork.Medium, Artwork.Grounds, " _
                 & "Artwork.FinalRetailPrice, Artwork.SalesTax, Artwork.TaxRate, Artwork.OrderDate, Artwork.CompletionDate, Artwork.Size, " _
                 & "Artwork.PortraitLandscape " _
                 & "FROM Customer INNER JOIN Artwork ON Customer.CustomerID = Artwork.CustomerID " _
                 & "WHERE Artwork.PortraitID = Forms!Artwork![PortraitID];", False
                 
        DoCmd.RunSQL "INSERT INTO InvoiceItemLines ( InvoiceID, ItemID, Title, Size, Medium, Frame, Price ) " _
                 & "SELECT MaxInv.InvoiceId, Item.ItemID, Item.Title, Item.Size, Item.Medium, Item.Frame, Item.Price " _
                 & "FROM " _
                 & "(SELECT Max(Invoice.InvoiceID) AS InvoiceId, PortraitID FROM Invoice WHERE Invoice.PortraitID = Forms!Artwork![PortraitID] " _
                 & "GROUP BY PortraitID) " _
                 & "MaxInv INNER JOIN Item ON MaxInv.PortraitID = Item.PortraitID;", False
                 
       DoCmd.RunSQL "INSERT INTO InvoiceChargeLines ( InvoiceID, ChargeID, ChargeType, Amount ) " _
                 & "SELECT MaxInv.InvoiceId, Charges.ChargeID, Charges.ChargeType, Charges.Amount " _
                 & "FROM " _
                 & "(SELECT Max(Invoice.InvoiceID) AS InvoiceId, PortraitID FROM Invoice WHERE Invoice.PortraitID = Forms!Artwork![PortraitID] " _
                 & "GROUP BY PortraitID) " _
                 & "MaxInv INNER JOIN Charges ON MaxInv.PortraitID = Charges.PortraitID;", False

       DoCmd.RunSQL "INSERT INTO InvoiceLines ( InvoiceID, PaymentDate, PaymentAmt, PaymentMethod, PaymentType ) " _
                 & "SELECT MaxInv.InvoiceId, Payment.PaymentDate, Payment.PaymentAmt, Payment.PaymentMethod, Payment.PaymentType " _
                 & "FROM " _
                 & "(SELECT Max(Invoice.InvoiceID) AS InvoiceId, PortraitID FROM Invoice WHERE Invoice.PortraitID = Forms!Artwork![PortraitID] " _
                 & "GROUP BY PortraitID) " _
                 & "MaxInv INNER JOIN Payment ON MaxInv.PortraitID = Payment.PortraitID;", False


        DoCmd.RunSQL "UPDATE Artwork SET Invoiced=True WHERE PortraitID = Forms!Artwork![PortraitID];", False
        
        'Warnings and alerts back on
        DoCmd.SetWarnings True
        
        MsgBox "Invoice successfully created"

    Else    ' User chose No.
       MsgBox "Invoice creation cancelled"
       GoTo Exit_CreateInvoiceCommand_Click
    End If

DoCmd.Requery

Exit_CreateInvoiceCommand_Click:
    Exit Sub

Err_CreateInvoiceCommand_Click:
    MsgBox Err.Description
    Resume Exit_CreateInvoiceCommand_Click
    
End Sub
'''
''' Locks and unlocks some fields on the form to prevent overwriting business records
'''
Private Sub CompCheck_Click()

    Dim State As Boolean
    
    State = CompCheck.Value

    'Lock/unlock fields depending on state
    OrderDate.Locked = State
    SalesLocation.Locked = State
    ReferenceNo = State
    TaxRate.Locked = State
    LocCode.Locked = State
    ConfCode.Locked = State
    SaleType.Locked = State
    Items.Locked = State
    FinalRetailPrice.Locked = State
    SalesTax.Locked = State
    TotalPrice.Locked = State

End Sub
'''
''' Looks up and changes the default tax rate and location code when sales location changes
''' Note that sales location can't be altered when a completed record
'''
Private Sub SalesLocation_Click()

        'Use DAO type recordsets
        Dim PBSB As Database
        Dim rst As Recordset
    
        Set PBSB = CurrentDb
    
        Dim thisSalesLocation As String
        Dim thisTaxRate As Double
        Dim thisLocCode As Variant
          
        thisSalesLocation = [SalesLocation]

       'The default tax rate and location code is provided by a lookup on SalesLocation table
       '- this table was previously named Settings
        Set rst = PBSB.OpenRecordset("SalesLocation", dbOpenSnapshot)
        With rst
            If .recordCount > 0 Then
                
                .MoveFirst
                .FindFirst "[LocationName] = '" & thisSalesLocation & "'"
                thisTaxRate = !TaxRate
                thisLocCode = !LocCode
               
                'Return to first record
                .MoveFirst
        
             End If
        
        End With
            
        Me.[TaxRate] = thisTaxRate
        Me.[LocCode] = thisLocCode

Exit_SalesLocation_Click:
End Sub
'''
''' Calculates the sales tax amount when field is double clicked
'''
Private Sub SalesTax_DblClick(Cancel As Integer)
    
    'Make sure that sale isn't already completed
    If [Completed] = True Then
        MsgBox ("This sale has already been completed")
        GoTo Exit_SalesTax_DblClick
    End If
       
    Call UpdateSalesTax
    
Exit_SalesTax_DblClick:
End Sub
'''
''' Opens a popup form which provides a filtered list of customers
'''
Private Sub SearchCommand_Click()

    'Make sure that sale isn't already completed
    If [Completed] = True Then
        MsgBox ("This sale has already been completed")
        GoTo Exit_SearchCommand_Click
    End If

    'Open CustomerPopUp form
    'record filtering is done by dataset selection on form
    DoCmd.OpenForm "CustomerPopUp", acFormDS, , , acFormEdit, acWindowNormal
    
Exit_SearchCommand_Click:
End Sub
'''
''' Adds a new customer record to the database
'''
Private Sub AddCommand_Click()

    'Make sure that sale isn't already completed
    If [Completed] = True Then
        MsgBox ("This sale has already been completed")
        GoTo Exit_AddCommand_Click
    End If

    'Get the customer name
    Dim thisCustomerID As Variant
    Dim thisCustomerName As String
    
    'Make sure customer name is not empty
 
    If IsNull(Me![SearchCustomer]) Then
        MsgBox ("You must add a name in the search box")
        GoTo Exit_AddCommand_Click
    End If
    
    thisCustomerName = Me![SearchCustomer]
    
    'Make sure it's not empty
    'If thisCustomerName = "" Then
    '    MsgBox ("You must add a name in the search box")
    '    GoTo Exit_AddCommand_Click
    'End If
    
    'Create new customer using customer name provided
    Dim PBSB As Database
    Dim rst As Variant
    
    Set PBSB = CurrentDb
    Set rst = PBSB.OpenRecordset("Customer", dbOpenDynaset)
                        
    With rst
        .AddNew
        !CustomerName = thisCustomerName
        .Update
        .Bookmark = .LastModified
        thisCustomerID = !CustomerID
    End With
        
    'Update the form fields
    Me![SearchCustomer] = ""
    Me![CustomerName] = thisCustomerName
    Me![CustomerID] = thisCustomerID
        
Exit_AddCommand_Click:
End Sub
'''
''' Creates a payment line for the total order amount and sets the order to "Complete"
'''
Private Sub SquareBtn_Click()
On Error GoTo Err_SquareBtn_Click

    Dim PaymentType As String
    PaymentType = "Square"

    'Make sure that sale isn't already completed
    If [Completed] = True Then
        MsgBox ("This sale has already been completed")
        GoTo Exit_SquareBtn_Click
    End If
    
    Call ValidateTotals

    Call ProcessPayment(PaymentType)

Exit_SquareBtn_Click:
    Exit Sub

Err_SquareBtn_Click:
    MsgBox Err.Description
    Resume Exit_SquareBtn_Click
End Sub
'''
''' Creates a payment line for the total order amount and sets the order to "Complete"
'''
Private Sub CashBtn_Click()
On Error GoTo Err_CashBtn_Click

    Dim PaymentType As String
    PaymentType = "Cash"

    'Make sure that sale isn't already completed
    If [Completed] = True Then
        MsgBox ("This sale has already been completed")
        GoTo Exit_CashBtn_Click
    End If

    Call ValidateTotals

    Call ProcessPayment(PaymentType)
    
Exit_CashBtn_Click:
    Exit Sub

Err_CashBtn_Click:
    MsgBox Err.Description
    Resume Exit_CashBtn_Click
End Sub
'''
''' Creates a payment line for the total order amount and sets the order to "Complete"
'''
Private Sub CheckBtn_Click()
On Error GoTo Err_CheckBtn_Click

    Dim PaymentType As String
    PaymentType = "Check"

    'Make sure that sale isn't already completed
    If [Completed] = True Then
        MsgBox ("This sale has already been completed")
        GoTo Exit_CheckBtn_Click
    End If
    
    Call ValidateTotals

    Call ProcessPayment(PaymentType)
    
Exit_CheckBtn_Click:
    Exit Sub

Err_CheckBtn_Click:
    MsgBox Err.Description
    Resume Exit_CheckBtn_Click
End Sub
'''
''' Runs when subtotal button is clicked
'''
Private Sub SubtotalBtn_Click()
On Error GoTo Err_SubtotalBtn_Click
  
    If [Completed] = True Then
    
        'Don't make changes to completed records
        MsgBox ("This sale has already been completed")
        GoTo Exit_SubtotalBtn_Click
    Else
    
        'OK to update the item subtotal
        Call UpdateSubtotal
    End If

Exit_SubtotalBtn_Click:
    Exit Sub

Err_SubtotalBtn_Click:
    MsgBox Err.Description
    Resume Exit_SubtotalBtn_Click
End Sub
'''
''' Calculates the total price of the individual items on the items subform
''' then adds a line to the charges subform for that amount
'''
Private Sub UpdateSubtotal()

    If Not IsItem() Then
    
        'Abort if no Item records exist
        MsgBox ("There are no item records")
        End
    Else
    
        Dim total As Currency
        Dim itemSubtotalCharge As Currency
        Dim subTotalName As String
        Dim thisPortraitID As Integer
    
        subTotalName = "ItemSubtotal"
        thisPortraitID = Me![PortraitID]
        
        'Get the item price subtotal
        total = TotalItems()
        itemSubtotalCharge = ItemSubtotalChargeAmount()
        
        'If the item amounts exist
        If Not IsNull(total) Then
    
            'Check to see if there's an existing subtotal (exact matches are skipped)
            If IsSubtotal() And (total <> itemSubtotalCharge) Then
    
                If GetMsgResponse("It looks like the existing item subtotal doesn't match. Do you want to update it?") = vbYes Then
                       
                    'Do an update
                    Call UpdateItemSubtotalCharge(total, subTotalName)
                Else
                    
                    'Exit without making updates
                    GoTo Exit_CalculateSubtotal
                End If
            
            ElseIf Not IsSubtotal() Then
            
                'Subtotal does not already exist so add a new record
                Call AddItemSubtotalCharge(thisPortraitID, total, subTotalName)
            End If
        End If
    End If
    
Exit_CalculateSubtotal:
End Sub
'''
''' Calculates the order total by adding the order subtotal and the sales tax amounts
'''
Private Sub TotalPrice_DblClick(Cancel As Integer)

    'Make sure that sale isn't already completed
    If [Completed] = True Then
        MsgBox ("This sale has already been completed")
        Exit Sub
    End If

    Call UpdateTotalPrice
    
    Repaint

End Sub
'''
''' Calculates the item subtotal and places item titles into an array
'''
Private Function TotalItems() As Variant

    Dim rst As Variant
    Dim itemTitle() As String
    Dim i As Integer
       
    'Instantiate items recordset and move to first record
    Set rst = Me![Items].Form.Recordset

    ReDim itemTitle(0 To i)

    With rst
        If .recordCount > 0 Then
    
            .MoveFirst

            'Then loop through and accumulate total
            Do While Not .EOF
    
                 TotalItems = TotalItems + !Price
                 If IsNull(!Title) = False Then
                    ReDim Preserve itemTitle(0 To i)
                    itemTitle(i) = !Title
                    i = i + 1
                 End If
                .MoveNext

            Loop
               
            'Return to first record for appearance sake
            .MoveFirst
        
        Else
            'Return 0 if no item records??
            TotalItems = 0
        End If
        
    End With
    
    Call UpdateTitle(itemTitle())
    
End Function
'''
''' Get ItemSubtotal amount from Charges form
'''
Private Function ItemSubtotalChargeAmount() As Variant

    Dim rst As Variant
       
    'Instantiate items recordset and move to first record
    Set rst = Me![Charges].Form.Recordset

    With rst
        If .recordCount > 0 Then
                
            .MoveFirst
            
            .FindFirst "[ChargeType] = 'ItemSubtotal'"

            ItemSubtotalChargeAmount = !Amount
               
            'Return to first record for appearance sake
            .MoveFirst
        
        End If
        
    End With
    
End Function
'''
''' Check whether there are any Items
'''
Private Function IsItem() As Boolean

        Dim rst As Variant
        Set rst = Me![Items].Form.Recordset
         
        With rst
        
            If .recordCount > 0 Then
                IsItem = True
            End If
        
        End With
        
End Function
'''
''' Check for ItemSubtotal in Charges subform
'''
Private Function IsSubtotal() As Boolean

        Dim rst As Variant
        Dim subTotalName As String
        subTotalName = "ItemSubtotal"
        
        IsSubtotal = False
        
        Set rst = Me![Charges].Form.Recordset
    
        With rst
        
            If .recordCount > 0 Then
        
        
                'Check to see if subtotal has already been calculated
                Do While Not .EOF
    
                    'If found set to TRUE and exit
                    If !ChargeType = subTotalName Then
                        IsSubtotal = True
                        Exit Do
                    End If
            
                    .MoveNext
                Loop
             End If
        
        End With
        
End Function
'''
''' Updates the Charge that corresponds to ItemSubtotal with the correct amount
'''
Private Sub UpdateItemSubtotalCharge(ByVal total As Variant, ByVal subTotalName As String)
                Dim rst As Variant
                Set rst = Me![Charges].Form.Recordset
                
                With rst
                    
                    .MoveFirst
                    
                    Do While Not .EOF
                    
                        If !ChargeType = subTotalName Then       'make update and exit
                            .Edit
                            !Amount = total
                            .Update
                            .Bookmark = .LastModified
                            Exit Do
                         End If
                        
                        .MoveNext
                      
                    Loop
                    
                End With
                
Exit_UpdateItemSubtotalCharge:
End Sub
'''
''' Adds a new charge line for the ItemSubtotal
'''
Private Sub AddItemSubtotalCharge(ByVal thisPortraitID As Integer, ByVal total As Variant, ByVal subTotalName As String)
    Dim rst As Variant
    Set rst = Me![Charges].Form.Recordset
            
    With rst
        .AddNew
        !PortraitID = thisPortraitID
        !ChargeType = subTotalName
        !Amount = total
        !Taxable = True
        .Update
        .Bookmark = .LastModified
    End With
    
End Sub
'''
''' Utility to manage the following:
''' Check that item records are correctly subtotalled
''' Check that charges are correctly subtotalled
''' Check for correct sale tax information
''' Check for correct order total
'''
Private Sub ValidateTotals()
   
    Call UpdateSubtotal
    Call UpdateTotal
    Call UpdateSalesTax
    Call UpdateTotalPrice
    
Exit_ValidateTotals:
End Sub
'''
''' Updates the order subtotal by adding all amounts on the charges subform
'''
Private Sub UpdateTotal()
    Dim total As Double
    Dim rst As Variant
    
    'Instantiate recordset
    Set rst = Me![Charges].Form.Recordset
    
    With rst
    
        If .recordCount > 0 Then
    
            'Move to first record
            .MoveFirst
    
            'Then loop through it and accumulate total
             Do While Not .EOF
             
                'Screen out NULLs in amount field as they won't accumulate
                If Not IsNull(Me![Charges].Form.[Amount]) Then
                    total = total + Me![Charges].Form.[Amount]
                End If
                
                .MoveNext
            Loop

            'Go back to first record and publish total
            .MoveFirst

        Else
    
            'There are no charge records so total must be zero
            total = 0
    
        End If
    
    End With
    
    'Update the form field
    [FinalRetailPrice] = total
    
End Sub
'''
''' Method to calculate and update sales tax
''' (prevents code duplication)
'''
Private Sub UpdateSalesTax()

    'Handle cases where sales tax rate not populated
    If IsNull([TaxRate]) Then
        [SalesTax] = 0
    Else
        [SalesTax] = Round([TaxRate] * [FinalRetailPrice], 2)
    End If

End Sub
'''
''' Method to calculate the gross price of the sale
''' (prevents code duplication)
'''
Private Sub UpdateTotalPrice()
    
    'Update the form field
    [TotalPrice] = [SalesTax] + [FinalRetailPrice]

End Sub
'''
''' Method to create a default extended title for an artwork record
'''
Private Sub UpdateTitle(ByRef itemTitle() As String)

    Dim extendedTitle As String
    Dim i As Integer
    
    'Unpack the array and concatenante to a string
    For i = LBound(itemTitle) To UBound(itemTitle)
    
        'Note the special handling for the first record
        If i = LBound(itemTitle) Then
            extendedTitle = itemTitle(i)
        Else
            extendedTitle = extendedTitle & "; " & itemTitle(i)
        End If
    
    Next i

    'Update the form field if it is empty
    If IsNull([Title]) Then
        [Title] = [extendedTitle]
    End If
    
End Sub
'''
''' Method to pay a sale in full
'''
Private Sub ProcessPayment(ByVal PaymentType As String)
    
    'Get the current record
    Dim record As Variant
    record = Screen.ActiveForm.CurrentRecord

    'Warnings and alerts off
    DoCmd.SetWarnings False

   'Insert Record (paymentType is passed in as a variable)
    DoCmd.RunSQL "INSERT INTO Payment ( PortraitID, PaymentDate, PaymentAmt, PaymentType, PaymentMethod ) " _
       & "VALUES ( Forms!Artwork![PortraitID], Forms!Artwork![OrderDate],Forms!Artwork![TotalPrice], 'Final Payment',""" & PaymentType & """);", False
                 
    'Set Completed flag on Artwork record
    [Completed] = True

    'Save the artwork record, requery and redisplay the modified Artwork record
     DoCmd.Save
     DoCmd.Requery
     DoCmd.GoToRecord acDataForm, "Artwork", acGoTo, record
     
    'Warnings and alerts back on
    DoCmd.SetWarnings True

End Sub

'''
''' Utility gets a user response to a yes/no question
'''
Private Function GetMsgResponse(ByVal messageString As String) As Integer
            Dim Msg, Style, Title, Help, Ctxt
    
            Msg = messageString
            Style = vbYesNo + vbDefaultButton2    ' Define buttons.
            Title = "Confirm action"    ' Define title.

            GetMsgResponse = MsgBox(Msg, Style, Title, Help, Ctxt)
End Function
'''
''' Utility to create a new default reference number by adding 1 to last used
'''
Function NewDefaultReference(ByVal stValue As String) As String
    
    Dim oldRefNo As Long
    Dim newRefNo As Long
    
    On Error GoTo ConversionFailureHandler
    
    'Try to convert old reference number to a number
    oldRefNo = CLng(stValue)
    
    NewDefaultReference = CStr(oldRefNo + 1)
    
    Do Until Len(NewDefaultReference) = REFERENCELENGTH
        NewDefaultReference = ("0" & NewDefaultReference)
    Loop

    Exit Function
    
ConversionFailureHandler:
 'IF we've reached this point, then we did not succeed in conversion
 'If the error is type-mismatch, clear the error and return empty string
 'Otherwise, disable the error handler, and re-run the code to allow the system to
 'display the error
 If Err.Number = 13 Then 'error # 13 is Type mismatch
      Err.Clear
      NewDefaultReference = ""
      Exit Function
 Else
      On Error GoTo 0
      Resume
 End If
End Function
