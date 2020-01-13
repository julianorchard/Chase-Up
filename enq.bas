Sub CheckDates()

' Activated with an AHK Script
' that just clicks a button


' Strings
    Dim msgTo, _
        msgSub, _
        msgVar, _
        today _
    As String
    msgTo = "INPUT EMAIL" 				' The address to send to
    
' Ranges
    Sheet1.Activate
    Dim remindates As Range
    Dim name, company, email, tel, address, city, code, country, datetheycontacted, notesonthem As Range
    Set remindates = Range("J1") 		' This is the cell at the top of the table of dates you want to send the messages on
    
' Counters
    Dim overflow, _
        counter _
    As Integer
    overflow = 0
    counter = 0
    
' Get Today Date as a String
    today = Format(Now(), "dd/mm/yyyy")
    
    Do
    '  Reminders
        Set remindates = remindates.Offset(1, 0)
        If IsEmpty(remindates.Value) Then
            overflow = overflow + 1
        ElseIf remindates.Value = today Then
        
        ' Objects
            Dim oLook, _
                oMail _
            As Object
        ' Set Objects
            Set oLook = CreateObject("Outlook.Application")
            Set oMail = oLook.CreateItem(0)
            
        ' Message Variables
            Set name = remindates.Offset(0, -9)
            Set company = remindates.Offset(0, -8)
            Set email = remindates.Offset(0, -7)
            Set tel = remindates.Offset(0, -6)
            Set address = remindates.Offset(0, -5)
            Set city = remindates.Offset(0, -4)
            Set code = remindates.Offset(0, -3)
            Set country = remindates.Offset(0, -2)
            Set datetheycontacted = remindates.Offset(0, -1)
            Set notesonthem = remindates.Offset(0, 1)
            ' Subject (just in case there is no contact name)
            If name = "" And company = "" Then
                msgSub = "Chase up a reply from UNKNOWN [Autoreminder]"
            ElseIf name = "" Then
                msgSub = "Chase up a reply from " & company & "  [Autoreminder]"
            Else
                msgSub = "Chase up a reply from " & name & "  [Autoreminder]"
            End If
            
        ' Send the message
            With oMail
                .To = msgTo
                .Body = vbNewLine & vbNewLine & _
                        "======================================" & vbNewLine & vbNewLine & _
                        "  Excel Autoreminder from an Enquiry  " & vbNewLine & vbNewLine & _
                        "======================================" & vbNewLine & vbNewLine & _
                        vbNewLine & vbNewLine & _
                        "    The details we found on them... " & vbNewLine & vbNewLine & _
                        "      Name and Company:" & vbNewLine & _
                        "           " & name & vbNewLine & _
                        "           " & company & vbNewLine & vbNewLine & _
                        "      Contact Details:" & vbNewLine & _
                        "           " & email & vbNewLine & _
                        "           " & tel & vbNewLine & vbNewLine & _
                        "      Address:" & vbNewLine & _
                        "           " & address & vbNewLine & _
                        "           " & city & vbNewLine & _
                        "           " & code & vbNewLine & _
                        "           " & country & vbNewLine & vbNewLine & _
                        "      Date they contacted:" & vbNewLine & _
                        "           " & datetheycontacted & vbNewLine & vbNewLine & _
                        "      Notes:" & vbNewLine & _
                        "           " & notesonthem & _
                        vbNewLine & vbNewLine & vbNewLine & vbNewLine & _
                        "Note: This email was sent automatically, and not by me. " & vbNewLine & _
                        "      Please do not respond, but let me know if sent in error."
                .Subject = msgSub
            ' If you want attachments
                ' .Attachments.Add "ATTACHMENT LOCATION"
                .Send
            End With
            
        ' Important to unset these
            Set oLook = Nothing
            Set oMail = Nothing
        Else
            overflow = 0
        End If
        counter = counter + 1
    Loop Until overflow = 10 Or counter = 1000 

End Sub