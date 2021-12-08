Attribute VB_Name = "Module1"
Sub EmailAll()
Dim oApp As Object
Dim oMail As Object
Dim SendToName As String
Dim theSubject As String
Dim theBody As String

For Each c In Selection 'loop through (manually) selected records
'''For each row in selection, collect the key parts of
'''the email message from the Table
 
    SendToName = Range("E" & c.Row)
    Salutation = Range("K" & c.Row)
    MyName = Range("F" & c.Row)
    Company = Range("A" & c.Row)
    Sector = Range("C" & c.Row)
        
'''Compose emails for each selected record
  '''Set object variables.
    Set oApp = CreateObject("Outlook.Application")
    Set oMail = oApp.CreateItem(0)
    Set myAttachments = oMail.Attachments
    
  '''Compose the customized message
    With oMail
        .To = SendToName
        .Subject = "Wharton Undergraduate Finance Club: Partnership Opportunity"
        
        .Body = "Dear " & Salutation & "," & vbCrLf & vbCrLf & "My name is " & MyName & _
        ", and I am part of the Wharton Undergraduate Finance Club (WUFC), the largest finance club at the University of Pennsylvania, with 2800+ members." _
        & vbCrLf & vbCrLf & "We serve as the primary finance resource for students by hosting regular events throughout the year, including finance conferences, speaker events, networking sessions and more." _
        & vbCrLf & vbCrLf & "I am reaching out to see if " & Company & " would be interested in developing a partnership with WUFC by sponsoring us this academic year." _
        & vbCrLf & vbCrLf & "With our students' extensive interest in Finance and " & Sector & " in particular, we would love to have you as our sponsor. " & Company & "'s core values and mission really resonate with us, and as the main touchpoint for Penn students interested in Finance, WUFC wishes to serve as an effective liaison between you and thousands of our students." _
        & vbCrLf & vbCrLf & "Through a sponsorship, WUFC will support you with organizing exclusive corporate events on Penn's campus (creating online meetings, booking rooms, organizing catering and logistics, etc). We will also promote all of your materials on our website, bi-weekly listserv, and social media platforms (Facebook, Twitter, etc.), amongst other benefits." _
        & vbCrLf & vbCrLf & "This makes us the best resource to gain direct and personalized exposure to thousands of talented students interested in finance. You can also find more information about us on our website: https://whartonfinanceclub.com/" _
        & vbCrLf & vbCrLf & "If you are interested in our sponsorship, we would love to have a conversation with you over the phone! I attached our Sponsorship Package below if you would like to take a quick look." _
        & vbCrLf & vbCrLf & "Please let us know if you have any questions. Thank you, and we look forward to hearing from you." _
        & vbCrLf & vbCrLf & "Best regards," & vbCrLf & MyName & vbCrLf & "Corporate Relations, The Wharton Undergraduate Finance Club"

        
        
        .CC = "baptaud@wharton.upenn.edu; fab10@wharton.upenn.edu; doanng01@wharton.upenn.edu"
        myAttachments.Add "C:\Users\12489\Downloads\WUFCSponsorshipPackage.pdf", olByValue, 1, "2021-2022 WUFC Sponsorship Package"



    ''' If you want to send emails automatically, use the Send option.
    ''' If you want to generate draft emails and review before sending, use the Display option.
    ''' Do not use both!
    '''To activate your chosen option: Remove the single quote from the beginning of the code line, then
    '''add the single quote back to the option you didn't choose
    
     .Send
     '.Display
    End With
Next c
End Sub
