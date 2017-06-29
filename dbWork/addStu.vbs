Option Compare Database


Private Sub Command0_Click()

Dim idnum, SchCvt, SSNholder, active, msgHolder As String
Dim StuNameFirst, StuNameLast, StuNameMid, StuGrade, StuGender, SSN, School, Ethinicity, Race, StuHomeAdd, StuHomeAdd2, StuMailAdd, StuMailAdd2 As String
Dim Guardianship, ParFathName, ParFathDayPhone, ParFathHomePhone, ParFathEmp, ParMothName, ParMothDayPhone, ParMothHomePhone, ParMothEmp, GuardEmail As String

Dim StuBirthDate As String 'Because People enter it wrong into Powerschool...
Dim inPSerr, inDBerr As String
Dim DistID, exist, err As Integer
Dim rs, sped, rs2 As DAO.Recordset




'########Start to complete the table with data from stuData########
Dim rs2 As DAO.Recordset

Set rs2 = CurrentDb.OpenRecordset("SELECT * FROM SECC")

'Check to see if the recordset actually contains rows
If Not (rs2.EOF And rs2.BOF) Then
    rs2.MoveFirst 'Unnecessary in this case, but still a good habit
    
    Do Until rs.EOF = True 'Loop through all the records one at a time
        exist = 0
        'Save student district numder into a variable
        idnum = rs2!DistID
        active = rs2!StuStatus
        
        
             If Not (idnum Like "") Then
            
                Open filePath For Input As #1
                
                
                Do Until EOF(1) 'Loop through student spreadsheet for each file
                
                    Line Input #1, LineFromFile
                    
                    LineItems = Split(LineFromFile, vbTab)
                    
                    DistID = LineItems(0) 'id number in this case

                    
                
                    'if student number is the same import student
                    If (idnum Like DistID) Then
                        
                        rs2.Edit
                        
                        rs2!StuNameFirst = LineItems(1)
                        rs2!StuNameLast = LineItems(2)
                        rs2!StuNameMid = LineItems(3)
                        rs2!SSID = LineItems(4)
                        rs2!StuBirthDate = LineItems(5)
                        rs2!StuGrade = LineItems(6)
                        rs2!StuGender = LineItems(7)
                        
                        SSNholder = LineItems(8)
                        If Not (SSNholder Like "") Then
                            rs2!SSN = Right(SSNholder, 4)
                        Else
                            rs2!SSN = ""
                        End If
                        
                         rs2!Ethnicity = LineItems(10)
                        rs2!StuHomAdd = LineItems(12)
                        rs2!StuMailAdd = LineItems(16)
                        rs2!Guardianship = LineItems(20)
                        rs2!ParFathName = LineItems(21)
                        rs2!ParFathDayPhone = LineItems(22)
                        rs2!ParFathHomePhone = LineItems(23)
                        rs2!ParFathEmp = LineItems(24)
                        rs2!ParMothName = LineItems(25)
                        rs2!ParMothDayPhone = LineItems(26)
                        rs2!ParMothHomePhone = LineItems(27)
                        rs2!ParMothEmp = LineItems(28)
                        rs2!GuardEmail = LineItems(29)
                    
                        rs2.Update
                    
                        exist = 1
                    
                    End If
                
                Loop
                
                Close #1
                
                'if exist code is not modified by having a student number in the csv, throw a warning
                If (exist Like 0 And active Like "Active") Then
                    'rs!active = False
                    msgHolder = "Student " & idnum & " is no longer in PowerSchool " & Chr(13) & msgHolder
                    'MsgBox "Student " & idnum & " is no longer in PowerSchool"
                End If
                
                
            Else
                
                MsgBox "Student is missing an ID, contact IT"
                
            End If
        
        
        '####Perform an edit
        'rs.Edit
        'rs!VendorYN = True
        'rs("VendorYN") = True 'The other way to refer to a field
        'rs.Update


        'Move to the next record. Don't ever forget to do this.
        rs.MoveNext
    Loop
Else
    MsgBox "There are no records in the recordset."
End If

MsgBox "Finished looping through records. " & msgHolder

rs2.Close 'Close the recordset
Set rs2 = Nothing 'Clean up

'########End of filling table from stuData########
End Sub



