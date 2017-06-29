Option Compare Database


Sub importStudents()

Dim idnum, SchCvt, SSNholder, active, msgHolder As String
Dim StuNameFirst, StuNameLast, StuNameMid, StuGrade, StuGender, SSN, School, Ethinicity, Race, StuHomeAdd, StuHomeAdd2, StuMailAdd, StuMailAdd2 As String
Dim Guardianship, ParFathName, ParFathDayPhone, ParFathHomePhone, ParFathEmp, ParMothName, ParMothDayPhone, ParMothHomePhone, ParMothEmp, GuardEmail As String
Dim StuBirthDate As String 'Because People enter it wrong into Powerschool...
Dim DistID, exist As Integer

Dim rs As DAO.Recordset

Dim filePath As String

filePath = "\\hsd-fs-01\shared$\do\Sped Database\Access (under construction)\import\student.txt"


Set rs = CurrentDb.OpenRecordset("SELECT * FROM stuData")

'Check to see if the recordset actually contains rows
If Not (rs.EOF And rs.BOF) Then
    rs.MoveFirst 'Unnecessary in this case, but still a good habit
    
    Do Until rs.EOF = True 'Loop through all the records one at a time
        exist = 0
        'Save student district numder into a variable
        idnum = rs!DistID
        active = rs!StuStatus
        
        
             If Not (idnum Like "") Then
            
                Open filePath For Input As #1
                
                
                Do Until EOF(1) 'Loop through student spreadsheet for each file
                
                    Line Input #1, LineFromFile
                    
                    LineItems = Split(LineFromFile, vbTab)
                    
                    DistID = LineItems(0) 'id number in this case

                    
                
                    'if student number is the same import student
                    If (idnum Like DistID) Then
                        
                        rs.Edit
                        
                        rs!StuNameFirst = LineItems(1)
                        rs!StuNameLast = LineItems(2)
                        rs!StuNameMid = LineItems(3)
                        rs!SSID = LineItems(4)
                        rs!StuBirthDate = LineItems(5)
                        rs!StuGrade = LineItems(6)
                        rs!StuGender = LineItems(7)
                        
                        SSNholder = LineItems(8)
                        If Not (SSNholder Like "") Then
                            rs!SSN = Right(SSNholder, 4)
                        Else
                            rs!SSN = ""
                        End If
                        
                        SchCvt = LineItems(9)
                        Select Case SchCvt
                            Case "662"
                                rs!School = "ILC"
                            Case "512"
                                rs!School = "ALMS"
                            Case "353"
                                rs!School = "SMS"
                            Case "612"
                                rs!School = "HHS"
                            Case "142"
                                rs!School = "DVES"
                            Case "139"
                                rs!School = "HHES"
                            Case "158"
                                rs!School = "RHES"
                            Case "167"
                                rs!School = "SES"
                            Case "191"
                                rs!School = "WPES"
                            Case Else
                                rs!School = "Update PS"
                        End Select
                        
                        rs!Ethnicity = LineItems(10)
                        'Race = LineItems(11) '*****not correct output*****
                        rs!StuHomAdd = LineItems(12)
                        rs!StuHomAdd2 = LineItems(13) & ", " & LineItems(14) & " " & LineItems(15)
                        'StuHomeAdd3 = LineItems(14)
                        'StuHomeAdd4 = LineItems(15)
                        rs!StuMailAdd = LineItems(16)
                        rs!StuMailAdd2 = LineItems(17) & ", " & LineItems(18) & " " & LineItems(19)
                        'StuMailAdd3 = LineItems(18)
                        'StuMailAdd4 = LineItems(19)
                        rs!Guardianship = LineItems(20)
                        rs!ParFathName = LineItems(21)
                        rs!ParFathDayPhone = LineItems(22)
                        rs!ParFathHomePhone = LineItems(23)
                        rs!ParFathEmp = LineItems(24)
                        rs!ParMothName = LineItems(25)
                        rs!ParMothDayPhone = LineItems(26)
                        rs!ParMothHomePhone = LineItems(27)
                        rs!ParMothEmp = LineItems(28)
                        rs!GuardEmail = LineItems(29)
                    
                        rs.Update
                    
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

rs.Close 'Close the recordset
Set rs = Nothing 'Clean up


End Sub

