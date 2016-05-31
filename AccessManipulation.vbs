Private Sub oneDollar_Click()

Dim RecordSt As DAO.Recordset
Dim dBase As DAO.Database
Dim stringSQL As String
Dim rCnt As Integer

Dim ItemNumber As Integer
Dim Amount As Integer
Dim Holder As Integer
Dim Bidder As Integer
Dim i As Integer

Holder = 1
i = 1

Bidder = Forms![Bid Sheets - Silent Auction]![Combo28]
ItemNumber = Forms![Bid Sheets - Silent Auction]![ItemNumber]

stringSQL = "SELECT * FROM BlankTableForBidSheet WHERE ItemNumber = " & ItemNumber & ";"

Set dBase = CurrentDb()

Set RecordSt = dBase.OpenRecordset(stringSQL)

Do While i <= RecordSt.RecordCount

   If RecordSt.Fields("BidNumber") > Holder Then
      Holder = RecordSt.Fields("BidNumber")
   End If
   i = i + 1
Loop


stringSQL = "SELECT * FROM BlankTableForBidSheet WHERE BidNumber = " & Holder & ";"

Set RecordSt = dBase.OpenRecordset(stringSQL)

Amount = RecordSt.Fields("Amount") + 1



stringSQL = "INSERT INTO BlankTableForBidSheet ([ItemNumber],[Amount],[BidderNumber]) VALUES (" & ItemNumber & "," & Amount & "," & Bidder & ");  "

DoCmd.SetWarnings False

DoCmd.RunSQL (stringSQL)


End Sub
