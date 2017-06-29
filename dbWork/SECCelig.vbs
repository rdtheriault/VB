Option Compare Database

Function YesNo(Info As String) As String
    'Info2 = CStr(Info)
    If Info Like 1 Then
        YesNo = "Y"
        Exit Function
    Else
        YesNo = "N"
        Exit Function
    End If
End Function

Private Sub Command0_Click()



'Dim idnum, SchCvt, SSNholder, active, msgHolder As String
'Dim StuNameFirst, StuNameLast, StuNameMid, StuGrade, StuGender, SSN, School, Ethinicity, Race, StuHomeAdd, StuHomeAdd2, StuMailAdd, StuMailAdd2 As String
'Dim Guardianship, ParFathName, ParFathDayPhone, ParFathHomePhone, ParFathEmp, ParMothName, ParMothDayPhone, ParMothHomePhone, ParMothEmp, GuardEmail As String
'Dim StuBirthDate As String 'Because People enter it wrong into Powerschool...

Dim inPSerr, inDBerr As String
Dim DistID, exist, err, testCnt As Integer
Dim rs, sped, rs2 As DAO.Recordset
Dim SECCRltdSvc(1 To 6) As String


inDBerr = "The follwing students are in the database but not marked in PS -"
inPSerr = "The follwing students are marked SPED in PS but are not in the database -"

CurrentDb.Execute "DELETE FROM SECCelig", dbFailOnError


'########This part of the code fills the SECC database with info from PowerSchool########
'date1 = Forms![SECC]![date1]
'date2 = Forms![SECC]![date2]
exist = 0

Dim filePath As String

filePath = "\\hsd-fs-01\shared$\do\Sped Database\Access (under construction)\import\student.txt"
'filePath = "students.csv"

    Open filePath For Input As #1
    
    
    Do Until EOF(1)
    
    testCnt = 0
    
    err = 0
    
        Line Input #1, LineFromFile
        
        LineItems = Split(LineFromFile, vbTab)
        
        test1 = LineItems(0) 'id number in this case
      
      
      '######loop through all records to find the students that need to be reported.######
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM stuData")
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst 'Unnecessary in this case, but still a good habit
        
        Do Until rs.EOF = True 'Loop through all the records one at a time
        
           If Not IsNull(rs!InitEligibDate) Then
            Dim ElibDate1 As Date
            ElibDate1 = CDate(rs!InitEligibDate)
                If test1 Like rs!DistID And ElibDate1 < CDate("06/30/2017") And ElibDate1 > CDate("07/01/2016") Then
                    err = 1
                End If
            End If
             rs.MoveNext
        Loop
    
    End If
      '######end loop move on to insert user if meets criteria.######

    
        'if student number is the same import student
        If (err = 1) Then
        
        '###Get data from stuData for current record###
        Set sped = CurrentDb.OpenRecordset("SELECT * FROM stuData WHERE DistID = " & test1)
            If Not (sped.EOF And sped.BOF) Then
                rs.MoveFirst 'Unnecessary in this case, but still a good habit
                
                Do Until sped.EOF = True 'Loop through all the records one at a time
                        SECCRltdSvc(1) = "00"
                        SECCRltdSvc(2) = "00"
                        SECCRltdSvc(3) = "00"
                        SECCRltdSvc(4) = "00"
                        SECCRltdSvc(5) = "00"
                        SECCRltdSvc(6) = "00"
                        
                        'ResdDistInstID = sped!ResDist
                        'ResdSchlInstID = sped!ResSchool
                        SECCPrimDsbltyCd = sped!PrimDisCode
                        'SECCSecDsblty1Cd = sped!SecDisCode
                        'SECCSecDsblty2Cd = sped!TerDisCode
                        
                        
                        Dim SECCRltdSvcCnt As Integer
                        SECCRltdSvcCnt = 1
 
                        If sped!Transp = True And SECCRltdSvcCnt < 7 Then
                            SECCRltdSvc(SECCRltdSvcCnt) = 25
                            SECCRltdSvcCnt = SECCRltdSvcCnt + 1
                        End If
                        If sped!Audiology = True And SECCRltdSvcCnt < 7 Then
                            SECCRltdSvc(SECCRltdSvcCnt) = 21
                            SECCRltdSvcCnt = SECCRltdSvcCnt + 1
                        End If
                        If (Not sped!OTRelated Like "") And SECCRltdSvcCnt < 7 Then
                            SECCRltdSvc(SECCRltdSvcCnt) = 19
                            SECCRltdSvcCnt = SECCRltdSvcCnt + 1
                        End If
                        If (Not sped!PTRelated Like "") And SECCRltdSvcCnt < 7 Then
                            SECCRltdSvc(SECCRltdSvcCnt) = 24
                            SECCRltdSvcCnt = SECCRltdSvcCnt + 1
                        End If
                        If (Not sped!SpeechRelated Like "") And SECCRltdSvcCnt < 7 Then
                            SECCRltdSvc(SECCRltdSvcCnt) = 20
                            SECCRltdSvcCnt = SECCRltdSvcCnt + 1
                        End If
                        If (Not sped!MentalHealthRelated Like "") And SECCRltdSvcCnt < 7 Then
                            SECCRltdSvc(SECCRltdSvcCnt) = 27
                            SECCRltdSvcCnt = SECCRltdSvcCnt + 1
                        End If
                        If (Not sped!AT Like "") And SECCRltdSvcCnt < 7 Then
                            SECCRltdSvc(SECCRltdSvcCnt) = 28
                            SECCRltdSvcCnt = SECCRltdSvcCnt + 1
                        End If
 
                        'SECCRltdSvc1 = SECCRltdSvc(1)
                        'SECCRltdSvc2 = SECCRltdSvc(2)
                        'SECCRltdSvc3 = SECCRltdSvc(3)
                        'SECCRltdSvc4 = SECCRltdSvc(4)
                        'SECCRltdSvc5 = SECCRltdSvc(5)
                        'SECCRltdSvc6 = SECCRltdSvc(6)
                        SECCFedPlCd = sped!FedPlaceCode
                        
                        If RegionStu = "" Then
                            SECCAgySrvCd = 30
                        Else
                            SECCAgySrvCd = 33
                        End If
                        
                        SECCAgySrvCd = sped!RegionStu
                        
                        'SECCEnrlTyp = sped!EnrollType
                        SECCEligDtTxt = sped!InitEligibDate
                        If Not SECCEligDtTxt = "" Then
                            SECCEligDtTxt = Format(CDate(SECCEligDtTxt), "MMDDYYYY")
                        End If
                        SECCLstIEPDtTxt = sped!LastIEP
                        If Not SECCLstIEPDtTxt = "" Then
                        SECCLstIEPDtTxt = Format(CDate(SECCLstIEPDtTxt), "MMDDYYYY")
                        End If
                        SECCSpEdExitDtTxt = sped!IEPupdateDate '83
                        If Not SECCSpEdExitDtTxt = "" Then
                        SECCSpEdExitDtTxt = Format(CDate(SECCSpEdExitDtTxt), "MMDDYYYY")
                        End If
                        'SECCRsnExtCd = sped!inactiveReason '84
                        
                        
                        EvlDtTxt = sped!EvalConsent
                        If Not EvlDtTxt = "" Then
                            EvlDtTxt = Format(CDate(EvlDtTxt), "MMDDYYYY")
                        End If
                        ElgDtTxt = sped!InitEligibDate
                        If Not ElgDtTxt = "" Then
                            ElgDtTxt = Format(CDate(ElgDtTxt), "MMDDYYYY")
                        End If
                        
 
                     sped.MoveNext
                Loop
    
            End If
                                            
                        DistStdntID = test1
                        LglFNm = LineItems(1)
                        LglFNm = Replace(LglFNm, "'", "''")
                        LglLNm = LineItems(2)
                        LglLNm = Replace(LglLNm, "'", "''")
                        LglMNm = LineItems(3)
                        LglMNm = Replace(LglMNm, "'", "''")
                        PrfrdFNm = LineItems(1)
                        PrfrdFNm = Replace(PrfrdFNm, "'", "''")
                        PrfrdLNm = LineItems(2)
                        PrfrdLNm = Replace(PrfrdLNm, "'", "''")
                        PrfrdMNm = LineItems(3)
                        PrfrdMNm = Replace(PrfrdMNm, "'", "''")
                        ChkDigitStdntID = LineItems(4)
                        
                        BirthDtTxt = LineItems(5) '
                        BirthDtTxt = Format(CDate(BirthDtTxt), "MMDDYYYY")
                        
                        GndrCd = LineItems(7)
                        HispEthnicFg = YesNo(CStr(LineItems(30)))
                        AmerIndianAlsknNtvRaceFg = YesNo(CStr(LineItems(31)))
                        AsianRaceFg = YesNo(CStr(LineItems(32)))
                        BlackRaceFg = YesNo(CStr(LineItems(33)))
                        WhiteRaceFg = YesNo(CStr(LineItems(34)))
                        PacIslndrRaceFg = YesNo(CStr(LineItems(35)))
                        LangOrgnCd = LineItems(36)
                        
                        EnrlGrdCdCvt = LineItems(6)
                        Select Case EnrlGrdCdCvt
                            Case "0"
                                EnrlGrdCd = "KG"
                            Case "1"
                                EnrlGrdCd = "01"
                            Case "2"
                                EnrlGrdCd = "02"
                            Case "3"
                                EnrlGrdCd = "03"
                            Case "4"
                                EnrlGrdCd = "04"
                            Case "5"
                                EnrlGrdCd = "05"
                            Case "6"
                                EnrlGrdCd = "06"
                            Case "7"
                                EnrlGrdCd = "07"
                            Case "8"
                                EnrlGrdCd = "08"
                            Case "9"
                                EnrlGrdCd = "09"
                            Case "10"
                                EnrlGrdCd = "10"
                            Case "11"
                                EnrlGrdCd = "11"
                            Case "12"
                                EnrlGrdCd = "12"
                            Case Else
                                EnrlGrdCd = "Update PS"
                        End Select
                        
                        If EnrlGrdCd Like "KG" Then
                            If SECCFedPlCd Like "30" Then
                                SECCSecFedPlCd = "M2"
                            End If
                            If SECCFedPlCd Like "33" Then
                                SECCSecFedPlCd = "L1"
                            End If
                        End If
                        
                        SSNholder = LineItems(8)
                        If Not SSNholder Like "" Then
                            SSN = Right(SSNholder, 4)
                        Else
                            SSN = ""
                        End If
                        
                        AttndDistInstID = "2206"
                        ResdDistInstID = "2206"
                        AttndSchlInstIDCvt = LineItems(9)
                        Select Case AttndSchlInstIDCvt
                            Case "662"
                                AttndSchlInstID = "4743"
                                ResdSchlInstID = "4743"
                            Case "512"
                                AttndSchlInstID = "1039"
                                ResdSchlInstID = "1039"
                            Case "353"
                                AttndSchlInstID = "1333"
                                ResdSchlInstID = "1333"
                            Case "612"
                                AttndSchlInstID = "1040"
                                ResdSchlInstID = "1040"
                            Case "142"
                                AttndSchlInstID = "3426"
                                ResdSchlInstID = "3426"
                            Case "139"
                                AttndSchlInstID = "1034"
                                ResdSchlInstID = "1034"
                            Case "158"
                                AttndSchlInstID = "1036"
                                ResdSchlInstID = "1036"
                            Case "167"
                                AttndSchlInstID = "1037"
                                ResdSchlInstID = "1037"
                            Case "191"
                                AttndSchlInstID = "1038"
                                ResdSchlInstID = "1038"
                            Case Else
                                AttndSchlInstID = "Update PS"
                                ResdSchlInstID = "Update PS"
                        End Select
                        
                        Addr = LineItems(12) 'street
                        Addr = Replace(Addr, ",", "-")
                        Addr = Replace(Addr, Chr(34), vbNullString)
                        City = LineItems(13) 'city
                        'StuHomeAdd3 = LineItems(14) 'state
                        ZipCd = LineItems(15) 'zip
                        ResdCntyCd = "30"
                        
                        Phn = LineItems(59)
                        Phn = Replace(Phn, "-", "")
                        
                        HSEntrySchlYr = LineItems(37)
                        EconDsvntgFg = YesNo(CStr(LineItems(38)))
                        Ttl1Fg = YesNo(CStr(LineItems(39)))
                        Sect504Fg = YesNo(CStr(LineItems(41)))
                        MigrntEdFg = YesNo(CStr(LineItems(42)))
                        IndianEdFg = YesNo(CStr(LineItems(43)))
                        LEPFg = YesNo(CStr(LineItems(44)))
                        DstncLrnFg = YesNo(CStr(LineItems(45)))
                        HomeSchlFg = YesNo(CStr(LineItems(46)))
                        TAGPtntTAGFg = YesNo(CStr(LineItems(47)))
                        TAGIntlctGiftFg = YesNo(CStr(LineItems(48)))
                        TAGAcdmTlntRdFg = YesNo(CStr(LineItems(49)))
                        TAGAcdmTlntMaFg = YesNo(CStr(LineItems(50)))
                        TAGCrtvAbltyFg = YesNo(CStr(LineItems(51)))
                        TAGLdrshpAbltyFg = YesNo(CStr(LineItems(52)))
                        TAGPrfmArtsAbltyFg = YesNo(CStr(LineItems(53)))
                        TrnstnProgFg = YesNo(CStr(LineItems(54)))
                        AltEdProgFg = YesNo(CStr(LineItems(55))) '53
                        AmerIndianTrbMbrshpCd = LineItems(56)
                        AmerIndianTrbEnrlmntNbr = LineItems(57)
                        'SECCRecTypCd = "A3" '57
                        'SECCSuppSvc1 = "00" '69
                        'SECCSuppSvc2 = "00" '70
                        'SECCSuppSvc3 = "00" '71
                        'SECCSuppSvc4 = "00" '72
                        'SECCSuppSvc5 = "00" '73
                        SECCResdDistInstID = ResdDistInstID '74
                        
                        SECCPrimLangCdCvt = LineItems(58) '85
                        Select Case SECCPrimLangCdCvt
                            Case "English"
                                SECCPrimLangCd = "1290"
                            Case "Spanish"
                                SECCPrimLangCd = "4260"
                            Case "Chinese"
                                SECCPrimLangCd = "3830"
                            Case Else
                                SECCPrimLangCd = "9999"
                        End Select
                        
                        
                        
                        SpEdFg = YesNo(CStr(LineItems(40)))
                        SECCEarlyEntryFg = "N" '89
                        
                        If Not SpEdFg Like "Y" Then
                            inDBerr = inDBerr & " " & test1
                        End If
                        
                        
        
            'Forms![Import Student]![infoImportStudent].Caption = StuNameFirst & " " & StuNameLast & " has been added to the database"

            stringSQL = "INSERT INTO SECCelig " _
            & "([ChkDigitStdntID],[DistStdntID],[AttndDistInstID],[AttndSchlInstID],[LglLNm],[LglFNm],[LglMNm],[GnrtnCd]," _
            & "[PrfrdLNm],[PrfrdFNm],[PrfrdMNm],[BirthDtTxt],[GndrCd],[HispEthnicFg],[AmerIndianAlsknNtvRaceFg],[AsianRaceFg]," _
            & "[BlackRaceFg],[WhiteRaceFg],[PacIslndrRaceFg],[LangOrgnCd],[SSN],[EnrlGrdCd],[Addr],[City]," _
            & "[ZipCd],[ResdCntyCd],[Phn],[HSEntrySchlYr],[EconDsvntgFg],[Ttl1Fg],[SpEdFg],[Sect504Fg]," _
            & "[MigrntEdFg],[IndianEdFg],[LEPFg],[DstncLrnFg],[HomeSchlFg],[TAGPtntTAGFg],[TAGIntlctGiftFg],[TAGAcdmTlntRdFg]," _
            & "[TAGAcdmTlntMaFg],[TAGCrtvAbltyFg],[TAGLdrshpAbltyFg],[TAGPrfmArtsAbltyFg],[TrnstnProgFg],[AltEdProgFg],[AmerIndianTrbMbrshpCd],[AmerIndianTrbEnrlmntNbr]," _
            & "[SECCResdDistInstID]," _
            & "[ResdDistInstID],[ResdSchlInstID],[SECCPrimDsbltyCd],[EvlDtTxt],[ElgDtTxt])" _
            & "VALUES (" _
            & "'" & ChkDigitStdntID & "','" & DistStdntID & "','" & AttndDistInstID & "','" & AttndSchlInstID & "','" & LglLNm & "','" & LglFNm & "','" & LglMNm & "','" & GnrtnCd & "'," _
            & "'" & PrfrdLNm & "','" & PrfrdFNm & "','" & PrfrdMNm & "','" & BirthDtTxt & "','" & GndrCd & "','" & HispEthnicFg & "','" & AmerIndianAlsknNtvRaceFg & "','" & AsianRaceFg & "'," _
            & "'" & BlackRaceFg & "','" & WhiteRaceFg & "','" & PacIslndrRaceFg & "','" & LangOrgnCd & "','" & SSN & "','" & EnrlGrdCd & "','" & Addr & "','" & City & "'," _
            & "'" & ZipCd & "','" & ResdCntyCd & "','" & Phn & "','" & HSEntrySchlYr & "','" & EconDsvntgFg & "','" & Ttl1Fg & "','" & SpEdFg & "','" & Sect504Fg & "'," _
            & "'" & MigrntEdFg & "','" & IndianEdFg & "','" & LEPFg & "','" & DstncLrnFg & "','" & HomeSchlFg & "','" & TAGPtntTAGFg & "','" & TAGIntlctGiftFg & "','" & TAGAcdmTlntRdFg & "'," _
            & "'" & TAGAcdmTlntMaFg & "','" & TAGCrtvAbltyFg & "','" & TAGLdrshpAbltyFg & "','" & TAGPrfmArtsAbltyFg & "','" & TrnstnProgFg & "','" & AltEdProgFg & "','" & AmerIndianTrbMbrshpCd & "','" & AmerIndianTrbEnrlmntNbr & "'," _
            & "'" & SECCResdDistInstID & "'," _
            & "'" & ResdDistInstID & "','" & ResdSchlInstID & "','" & SECCPrimDsbltyCd & "','" & EvlDtTxt & "','" & ElgDtTxt & "');  "
 
            If testCnt = 0 Then
                'MsgBox stringSQL
            End If
            
            'DoCmd.SetWarnings False
            'dbs.Execute stringSQL
            'dbs.Close
            'DoCmd.RunSQL (stringSQL)
            CurrentDb.Execute stringSQL, dbFailOnError
            
            exist = 1
            testCnt = 1
        End If
    
    Loop
    
    Close #1
    'end of inserting data
    msgHolder = inDBerr
'########End of filling table from PowerSchool########
MsgBox "Finished looping through records. " ' & msgHolder


End Sub



