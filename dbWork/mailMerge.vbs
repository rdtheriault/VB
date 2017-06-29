Option Compare Database

Private Sub Combo512_AfterUpdate()

'Dim goalID As Integer
'Dim sql As String

'goalID = Form!Combo512.Column(0)
'sql = "SELECT * FROM stuObj WHERE stuObj.goalID = " & goalID & ";"

'MsgBox (goalID)

'Populate Goal dropdown with current IEP goal areas

'Me.stuObj.Visible = True
'Me.stuObjSub.Visible = True

'Form!stuObjSub.SetFocus   stuObj.iepID = " & Me.ID & " AND
'Form!stuObjSub.Form.RecordSource = "SELECT stuObj.goalID, stuObj.area, stuObj.obj FROM stuObj WHERE stuGoals.iepID = " & Form![ID] & " AND stuObj.goalID = " & goalID & ";"
'Me.stuObjSub.SetFocus
'Me.stuObjSub.SourceObject = "Query.[stuData Query]"
'Me.stuObjSub.Form.RecordSource = sql
'Me.stuObjSub.Requery




End Sub

Private Sub Command839_Click()

    Dim recordnum, sqlStr As String
    'Dim


    DoCmd.OpenReport "stuSEP14", acViewNormal, , "[ID] =" & Me.ID

    recordnum = [ID].Value
    sqlStr = "SELECT * FROM [IEPs] WHERE IEPs.id = " & recordnum
    MsgBox ("The Access form will close and a Word document with the input will open")
    
    'DoEvents
    Form.Refresh
    DoCmd.Close
    'Form.
        
   Dim objWord As Word.Document
   ''Set objWord = GetObject("G:\do\Sped Database\Access (under construction)\IEP+\mailmerge.docx", "Word.Document")
   Set objWord = GetObject("\\hsd-data-01\sped2\Mailmerge\mailmerge1.docx", "Word.Document")
   '' Make Word visible.
   objWord.Application.Visible = True
   '' Set the mail merge data source as the Northwind database.
   objWord.MailMerge.OpenDataSource _
      Name:="\\hsd-data-01\sped2\IEP Db - Edit.accdb", _
      LinkToSource:=True, _
      Connection:="TABLE IEPs", _
      SQLStatement:=sqlStr
   ' Execute the mail merge.
   
    objWord.MailMerge.Destination = wdSendToNewDocument
    objWord.MailMerge.Execute
    objWord.Application.Options.PrintBackground = False
    objWord.Application.ActiveDocument.PrintOut
    objWord.Close SaveChanges:=wdDoNotSaveChanges
    'objWord.Quit

End Sub


Private Sub Form_Current()

'If [PSNParentCom] = True Then
    '[ParentComYes].Caption = "YES    X"
    '[ParentComNo].Caption = "NO"
'Else
    '[ParentComNo].Caption = "NO   X"
    '[ParentComYes].Caption = "YES"
'End If

'If [PSNStuCom] = True Then
    '[StuComYes].Caption = "YES    X"
    '[StuComNo].Caption = "NO"
'Else
    '[StuComNo].Caption = "NO   X"
    '[StuComYes].Caption = "YES"
'End If


'Populate Goal dropdown with current IEP goal areas
'Form!Combo512.RowSource = "SELECT stuGoals.goalID, stuGoals.area FROM stuGoals WHERE stuGoals.iepID = " & Form![ID] & " ORDER BY stuGoals.area;"


'hide Objectives on page 10 when area is not selected.
'Me.stuObj.Visible = False
'Me.stuObjSub.Visible = False

End Sub


Private Sub PrintSpec_Click()

    Dim recordnum, sqlStr As String
    'Dim


    DoCmd.OpenReport "stuGoalsRep10", acViewNormal, , "[iepID] =" & Me.ID
    DoCmd.OpenReport "stuServRep11&12", acViewNormal, , "[ID] =" & Me.ID
    DoCmd.OpenReport "stuESYall13", acViewNormal, , "[ID] =" & Me.ID
    DoCmd.OpenReport "stuSEP14", acViewNormal, , "[ID] =" & Me.ID

    recordnum = [ID].Value
    sqlStr = "SELECT * FROM [IEPs] WHERE IEPs.id = " & recordnum
    MsgBox ("The Access form will close and a Word document with the input will open")
    
    'DoEvents
    Form.Refresh
    DoCmd.Close
    'Form.
        
   Dim objWord As Word.Document
   ''Set objWord = GetObject("G:\do\Sped Database\Access (under construction)\IEP+\mailmerge.docx", "Word.Document")
   Set objWord = GetObject("\\hsd-data-01\sped2\Mailmerge\mailmerge.docx", "Word.Document")
   '' Make Word visible.
   objWord.Application.Visible = True
   '' Set the mail merge data source as the Northwind database.
   objWord.MailMerge.OpenDataSource _
      Name:="\\hsd-data-01\sped2\IEP Db - Edit.accdb", _
      LinkToSource:=True, _
      Connection:="TABLE IEPs", _
      SQLStatement:=sqlStr
   ' Execute the mail merge.
   
    objWord.MailMerge.Destination = wdSendToNewDocument
    objWord.MailMerge.Execute
    objWord.Application.Options.PrintBackground = False
    objWord.Application.ActiveDocument.PrintOut
    objWord.Close SaveChanges:=wdDoNotSaveChanges
    ''objWord.Quit


End Sub

Private Sub PSNParentCom_AfterUpdate()
    'If [PSNParentCom] = True Then
        '[ParentComYes].Caption = "YES    X"
        '[ParentComNo].Caption = "NO"
   'Else
        '[ParentComNo].Caption = "NO   X"
        '[ParentComYes].Caption = "YES"
    'End If
End Sub


Private Sub PSNStuCom_AfterUpdate()
    'If [PSNStuCom] = True Then
        '[StuComYes].Caption = "YES    X"
        '[StuComNo].Caption = "NO"
    'Else
        '[StuComNo].Caption = "NO   X"
        '[StuComYes].Caption = "YES"
    'End If
End Sub

'Private Sub ParCheck_AfterUpdate()
    'If [ParCheck] = True Then
        '[Parent Info].Caption = "Parent Info (X)"
    'Else
        '[Parent Info].Caption = "Parent Info"
    'End If
'End Sub

'Private Sub ProgPlacCheck_AfterUpdate()
    'If [ProgPlacCheck] = True Then
        '[Program/Placement].Caption = "Program/Placement (X)"
    'Else
        '[Program/Placement].Caption = "Program/Placement"
    'End If
'End Sub

