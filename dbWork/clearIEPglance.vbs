Option Compare Database

Sub updateIEPglnc()


Dim rs As DAO.Recordset
Set rs = CurrentDb.OpenRecordset("SELECT * FROM IEPglnc")

'Check to see if the recordset actually contains rows
If Not (rs.EOF And rs.BOF) Then
    rs.MoveFirst 'Unnecessary in this case, but still a good habit
    Do Until rs.EOF = True

        ''Save contact name into a variable
        'sContactName = rs!FirstName & " " & rs!LastName




        SQL = "UPDATE stuData SET ClassTeacher = '" & rs!ClassTeacher & "' , GlncDisAll = '" & rs!GlncDisAll & "' , GlncPresLvl = '" & rs!GlncPresLvl & "' , GlncCogStren = '" & rs!GlncCogStren & "' , GlncCogWeak = '" & rs!GlncCogWeak & "' , GlncSDI = '" & rs!GlncSDI & "' , GlncRelServ = '" & rs!GlncRelServ & "' , GlncSup = '" & rs!GlncSup & "' , GlncNotes = '" & rs!GlncNotes & "' , GlncSmrtBlnc = '" & rs!GlncSmrtBlnc & "' , GlncELPA = '" & rs!GlncELPA & "' WHERE ID like '" & rs!ID & "';"
        DoCmd.SetWarnings False
        
        DoCmd.RunSQL (SQL)

        'Move to the next record. Don't ever forget to do this.
        rs.MoveNext
    Loop
Else
    MsgBox "There are no records in the recordset."
End If

MsgBox "Finished looping through records."

rs.Close 'Close the recordset
Set rs = Nothing 'Clean up


End Sub
