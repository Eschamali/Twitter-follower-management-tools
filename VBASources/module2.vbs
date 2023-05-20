sub openTwitter()
	Dim WSH
	Set WSH = CreateObject("Wscript.Shell")

	'先頭位置を取得
	Dim startRow
	startRow = Selection.Row

	'選択行数を取得
	Dim selectRow
	selectRow = Selection.Rows.Count

	'選択した範囲のtwitterIDを取得してそれでtwitterプロフィールページを開く
	Dim i
	Dim twitterID
	For i = startRow to startRow + selectRow - 1
		twitterID = Cells(i,2)
		WSH.Run "https://twitter.com/intent/user?user_id=" & twitterID
	Next
end sub