Option Explicit

'メッセージ用
public addMessage as string

sub csvRead()
	Dim FollowCSV
	Dim FollowerCSV
	Dim Before_FollowCSV
	Dim Before_FollowerCSV
	Dim BK_FollowCSV
	Dim BK_FollowerCSV
	Dim BK_Before_FollowCSV
	Dim BK_Before_FollowerCSV
	Dim WSH
	Dim CurrentDir
	Dim TwitterAccount
	Dim continue
	Dim pause

	'FF情報を取得するbatファイル名を指定
	const useGetFFInfoBat = "getFFinfo_API.bat"

	'bat終了時の待機フラグ
	pause = 0

	'不用意にアカウント名が変わらないように、ファイル名からFFリストを取得するアカウント名を特定する方式にしてます。
	TwitterAccount = Mid(ThisWorkbook.Name, 1, Len(ThisWorkbook.Name) - 5)

	'最初のシートにあるチェックBOXの状態を取得する
	If Sheets("compare_results").CheckBox1.Value = FALSE Then
		'確認メッセージが有効なので、bat終了時のpauseも有効にする
		pause = 1
		continue = MsgBox("Twitterアカウント：" & TwitterAccount & vbCrLf & "上記アカウントのFF情報を取得します。"& vbCrLf & vbCrLf & "よろしいですか？", vbYesNo + vbQuestion, "FFデータ取得")
		If continue = vbNo Then
			exit sub
		End If
	End If

	'bat起動用WshShellクラス
	Set WSH = CreateObject("Wscript.Shell")

	'CurrentDirを取得する(onedrive保存にも対応させる)
	'まずは普通にCurrentDirを取得する関数で
	CurrentDir = ThisWorkbook.Path

	'httpリンクで始まるものは「onedrive」環境変数を使って変換させる
	'そうじゃない場合は、そのままで。※ただし、SharePointでの挙動は未確認
	If Left(CurrentDir,4) = "http" Then
		CurrentDir = Environ("OneDrive") & Mid(CurrentDir, 41)
	End If

	'ファイルパスを読み込む
	FollowCSV = CurrentDir & "\input\" & TwitterAccount & "-follow.csv"
	FollowerCSV = CurrentDir & "\input\" & TwitterAccount & "-follower.csv"
	Before_FollowCSV = CurrentDir & "\input\" & TwitterAccount & "-before_follow.csv"
	Before_FollowerCSV = CurrentDir & "\input\" & TwitterAccount & "-before_follower.csv"
	BK_FollowCSV = CurrentDir & "\backup\" & TwitterAccount & "-follow.csv"
	BK_FollowerCSV = CurrentDir & "\backup\" & TwitterAccount & "-follower.csv"
	BK_Before_FollowCSV = CurrentDir & "\backup\" & TwitterAccount & "-before_follow.csv"
	BK_Before_FollowerCSV = CurrentDir & "\backup\" & TwitterAccount & "-before_follower.csv"

	'カレントディレクトリとTwitterアカウントを「setting」シートに貼り付ける
	'PowerQueryのCSV読み込み先に使います。
	With Sheets("setting").ListObjects(1)
		.DataBodyRange.AutoFilter 1, "FFリスト読み込み先"
		.ListColumns(2).DataBodyRange = CurrentDir & "\input\"

		.DataBodyRange.AutoFilter 1, "Twitterアカウント"
		.ListColumns(2).DataBodyRange = TwitterAccount

		.DataBodyRange.AutoFilter 1
	End With

	'ミスったときに備えて、inputフォルダにあるcsvファイルをbackupにコピー(1回分のみ)
	If Dir(Before_FollowCSV) <> "" Then
		FileCopy Before_FollowCSV, BK_Before_FollowCSV
	End If
	If Dir(Before_FollowerCSV) <> "" Then
		FileCopy Before_FollowerCSV, BK_Before_FollowerCSV
	End If

	If Dir(FollowCSV) <> "" Then
		FileCopy FollowCSV , BK_FollowCSV

		'コピー後、前回のリストcsvとして、beforeにリネーム
		FileCopy FollowCSV , Before_FollowCSV

		'コピーしたので削除する
		Kill FollowCSV
	End If

	If Dir(FollowerCSV) <> "" Then
		FileCopy FollowerCSV, BK_FollowerCSV

		'コピー後、前回のリストcsvとして、beforeにリネーム
		FileCopy FollowerCSV, Before_FollowerCSV

		'コピーしたので削除する
		Kill FollowerCSV
	End If

	'batファイルを起動してFF情報を取得します。
	'引数1:起動バッチファイルパス 2:そのbatファイルにあるディレクトリパスにcdする 3:FF情報を取得するアカウント 4:bat終了時のpauseコマンドを使うかフラグ
	WSH.Run CurrentDir & "\bat\" & useGetFFInfoBat & " " & CurrentDir & "\bat " & TwitterAccount & " " & pause, 1, True

	'ファイル存在確認
	'Pythonの処理の仕方を工夫し、FF情報を取得しきった後にファイルが作成されるので、簡易的に失敗判定する
	If Dir(FollowCSV) = "" OR Dir(FollowerCSV) = "" Then
		MsgBox "最新のFF情報の取得に失敗しました。" & vbCrLf & "書き込みの権限不足、リクエストの超過などが考えられます。", vbOKOnly + vbCritical, "FF情報の取得に失敗"
		exit sub
	End If

	'beforeファイルがない場合、取得したファイルをコピーし初回取得である旨のメッセージを追加する
	addMessage = ""
	If Dir(Before_FollowCSV) = "" AND Dir(Before_FollowerCSV) = "" Then
		FileCopy FollowCSV , Before_FollowCSV
		FileCopy FollowerCSV, Before_FollowerCSV
		addMessage = vbCrLf & vbCrLf & "※初回取得のため、コンペア結果、ログ記録は、次回データ取得時に" & vbCrLf & "FF状況に変化があるときに表示します。"
	End If

	'最初のシートにあるチェックBOXの状態を取得する
	If Sheets("compare_results").CheckBox1.Value = True Then
		dataRefresh
	else
		continue = MsgBox("Pythonのtweepyで、FF情報の取得に成功しました。" & vbCrLf & "続けて、データを更新しますがよろしいですか？" & addMessage, vbYesNo + vbQuestion, "FF情報の取得完了")
		If continue = vbYes Then
			dataRefresh
		End If
	End If
end sub


sub dataRefresh()
	'作成したクエリの数
	Const quaryNum = 7

	Dim refreshNum(quaryNum)
	refreshNum(0) = "クエリ - follow"
	refreshNum(1) = "クエリ - follower"
	refreshNum(2) = "クエリ - before_follow"
	refreshNum(3) = "クエリ - before_follower"
	refreshNum(4) = "クエリ - newFollow"
	refreshNum(5) = "クエリ - deleteFollow"
	refreshNum(6) = "クエリ - newFollower"
	refreshNum(7) = "クエリ - deleteFollower"

	'OLEDBとODBCでない場合エラーになるので回避のためのおなじない
	On Error Resume Next
	Dim i
	Dim bolOLEDB,bolODBC
	For i = 0 to quaryNum - 4'コネクションの全てをループ
		'BackgroundQueryの状態を変数に保存
		bolOLEDB = ActiveWorkbook.Connections.OLEDBConnection.BackgroundQuery
		bolODBC = ActiveWorkbook.Connections.ODBCConnection.BackgroundQuery
		ActiveWorkbook.Connections.OLEDBConnection.BackgroundQuery = False
		ActiveWorkbook.Connections.ODBCConnection.BackgroundQuery = False
		'更新実行
		ActiveWorkbook.Connections(refreshNum(i)).Refresh
		'BackgroundQueryの状態を元に戻す
		ActiveWorkbook.Connections.OLEDBConnection.BackgroundQuery = bolOLEDB
		ActiveWorkbook.Connections.ODBCConnection.BackgroundQuery = bolODBC
	Next

	'最初のシートにあるチェックBOXの状態を取得する
	If Sheets("compare_results").CheckBox1.Value = FALSE Then
		MsgBox "取得したFF情報からデータを更新しました。" & vbCrLf & "コンペア結果を表示します。" & addMessage, vbOKOnly + vbInformation, "反映完了"
	End If

	For i = 4 to quaryNum 'コネクションの全てをループ
		'BackgroundQueryの状態を変数に保存
		bolOLEDB = ActiveWorkbook.Connections.OLEDBConnection.BackgroundQuery
		bolODBC = ActiveWorkbook.Connections.ODBCConnection.BackgroundQuery
		ActiveWorkbook.Connections.OLEDBConnection.BackgroundQuery = False
		ActiveWorkbook.Connections.ODBCConnection.BackgroundQuery = False
		'更新実行
		ActiveWorkbook.Connections(refreshNum(i)).Refresh
		'BackgroundQueryの状態を元に戻す
		ActiveWorkbook.Connections.OLEDBConnection.BackgroundQuery = bolOLEDB
		ActiveWorkbook.Connections.ODBCConnection.BackgroundQuery = bolODBC
	Next

	On Error GoTo 0

	'指定のテーブルに関数を入れる
	With Sheets("compare_results")
		If .ListObjects(1).ListRows.Count > 0 Then
			.ListObjects(1).ListColumns(6).DataBodyRange.Formula = "=XLOOKUP([@ID],newFollower[ID],newFollower[前回とのコンペア],""No"")"
		End If

		If .ListObjects(2).ListRows.Count > 0 Then
			.ListObjects(2).ListColumns(6).DataBodyRange.Formula = "=XLOOKUP([@ID],deleteFollower[ID],deleteFollower[前回とのコンペア],""No"")"
		End If

		If .ListObjects(3).ListRows.Count > 0 Then
			.ListObjects(3).ListColumns(6).DataBodyRange.Formula = "=XLOOKUP([@ID],newFollow[ID],newFollow[前回とのコンペア],""No"")"
		End If

		If .ListObjects(4).ListRows.Count > 0 Then
			.ListObjects(4).ListColumns(6).DataBodyRange.Formula = "=XLOOKUP([@ID],deleteFollow[ID],deleteFollow[前回とのコンペア],""No"")"
		End If
	End With

	With Sheets("follow")
		If .ListObjects(1).ListRows.Count > 0 Then
			.ListObjects(1).ListColumns(17).DataBodyRange.Formula = "=SUBSTITUTE(XLOOKUP([@ID],before_follow[ID],before_follow[ID],""フォローした"",0,1),[@ID],""変化なし"")"
			.ListObjects(1).ListColumns(18).DataBodyRange.Formula = "=SUBSTITUTE(XLOOKUP([@ID],follower[ID],follower[ID],""片思い"",0,1),[@ID],""FF関係"")"
		End If
	End With

	With Sheets("follower")
		If .ListObjects(1).ListRows.Count > 0 Then
			.ListObjects(1).ListColumns(17).DataBodyRange.Formula = "=SUBSTITUTE(XLOOKUP([@ID],before_follower[ID],before_follower[ID],""フォローされた""),[@ID],""変化なし"")"
			.ListObjects(1).ListColumns(18).DataBodyRange.Formula = "=SUBSTITUTE(XLOOKUP([@ID],follow[ID],follow[ID],""片思われ""),[@ID],""FF関係"")"
		End If
	End With

	With Sheets("before_follow")
		If .ListObjects(1).ListRows.Count > 0 Then
			.ListObjects(1).ListColumns(17).DataBodyRange.Formula = "=SUBSTITUTE(XLOOKUP([@ID],follow[ID],follow[ID],""フォローを解除"",0,1),[@ID],""変化なし"")"
		End If
	End With

	With Sheets("before_follower")
		If .ListObjects(1).ListRows.Count > 0 Then
			.ListObjects(1).ListColumns(17).DataBodyRange.Formula = "=SUBSTITUTE(XLOOKUP([@ID],follower[ID],follower[ID],""フォローを解除された""),[@ID],""変化なし"")"
		End If
	End With

	'最初のシートにあるチェックBOXの状態を取得する
	If Sheets("compare_results").CheckBox1.Value = True Then
		Sheets("ReportLog").select
		compare
	else
		Dim continue

		continue = MsgBox("コンペア結果を表示しました。" & vbCrLf & "続けて、コンペア結果を記録しますがよろしいですか？" & addMessage, vbYesNo + vbQuestion, "更新完了")
		If continue = vbYes Then
			Sheets("ReportLog").select
			compare
		End If

	End If

End sub


Sub compare()
	'各ユニークIDを格納
	Dim NewFollow
	Dim DeleteFollow
	Dim NewFollower
	Dim DeleteFollower

	'日付を取得
	'Dim yearValue
	'Dim monthValue
	'Dim dayValue
	'Dim hourValue
	'Dim minitueValue
	'Dim secondValue
	Dim ActionNow
	'yearValue = Format(Year(Date()),"0000")
	'monthValue = Format(Month(Date()),"00")
	'dayValue = Format(Day(Date()),"00")
	'hourValue = Format((Time()),"00")
	'minitueValue = Format((Time()),"00")
	'secondValue = Format((Time()),"00")
	ActionNow = NOW()

	'コンペア結果時間を記録
	Sheets("compare_results").Cells(1,2).value = ActionNow

	'結果をまとめる
	Dim i
	for i = 1 to 4

		'ListObjectsを使ったテーブル操作
		With Sheets("compare_results").ListObjects(i)

			Select Case i
				Case 1
					if .ListRows.Count > 0 Then
						NewFollow = .DataBodyRange
					End if
				Case 2
					if .ListRows.Count > 0 Then
						DeleteFollow = .DataBodyRange
					End if
				Case 3
					if .ListRows.Count > 0 Then
						NewFollower = .DataBodyRange
					End if
				Case 4
					if .ListRows.Count > 0 Then
						DeleteFollower = .DataBodyRange
					End If
				Case Else
					'none
			End Select

		End With
	Next

	'結果を出力する
	for i = 1 to 4

		'ListObjectsを使ったテーブル操作
		With Sheets("ReportLog").ListObjects(1)

			Dim RowPotion, getNum,recordFlag
			Select Case i
				Case 1
					if IsEmpty(NewFollow) = FALSE Then
						For getNum = 1 to UBound(NewFollow)
							'書き込む行位置を特定する
							RowPotion = .ListRows.Count + 2

							'指定箇所の書式を設定
							.ListColumns(1).Range(RowPotion).NumberFormatLocal = "yyyy/mm/dd"
							.ListColumns(2).Range(RowPotion).NumberFormatLocal = "@"

							'配列内の値を出力
							.ListColumns(1).Range(RowPotion) = ActionNow
							.ListColumns(2).Range(RowPotion) = NewFollow(getNum,1)
							.ListColumns(3).Range(RowPotion) = NewFollow(getNum,2)
							.ListColumns(4).Range(RowPotion) = NewFollow(getNum,3)

								'フラグ設定
							.ListColumns(5).Range(RowPotion) = TRUE

							'もし記録時点で同時に相互関係になったら、同一レコードに記録する
							If NewFollow(getNum,6) = "フォローされた" Then
								.ListColumns(7).Range(RowPotion) = TRUE
							End If

							'フォロー、フォロワー数を設定
							.ListColumns(9).Range(RowPotion) = Sheets("follow").Range("B3")
							.ListColumns(10).Range(RowPotion) = Sheets("follower").Range("B3")

							'プロフィール画像を設定
							.ListColumns(11).Range(RowPotion) = NewFollow(getNum,4)
						Next
					End if

				Case 2
					if IsEmpty(DeleteFollow) = FALSE Then
						For getNum = 1 to UBound(DeleteFollow)
							'書き込む行位置を特定する
							RowPotion = .ListRows.Count + 2

							'指定箇所の書式を設定
							.ListColumns(1).Range(RowPotion).NumberFormatLocal = "yyyy/mm/dd"
							.ListColumns(2).Range(RowPotion).NumberFormatLocal = "@"

							'配列内の値を出力
							.ListColumns(1).Range(RowPotion) = ActionNow
							.ListColumns(2).Range(RowPotion) = DeleteFollow(getNum,1)
							.ListColumns(3).Range(RowPotion) = DeleteFollow(getNum,2)
							.ListColumns(4).Range(RowPotion) = DeleteFollow(getNum,3)

								'フラグ設定
							.ListColumns(6).Range(RowPotion) = TRUE

							'もし記録時点で凍結やブロ解の可能性がある判定になったら、同一レコードに記録する
							If DeleteFollow(getNum,6) = "フォローを解除された" Then
								.ListColumns(8).Range(RowPotion) = TRUE
							End If

								'フォロー、フォロワー数を設定
							.ListColumns(9).Range(RowPotion) = Sheets("follow").Range("B3")
							.ListColumns(10).Range(RowPotion) = Sheets("follower").Range("B3")

								'プロフィール画像を設定
							.ListColumns(11).Range(RowPotion) = DeleteFollow(getNum,4)
						Next
					End if

				Case 3
					if IsEmpty(NewFollower) = FALSE Then
						For getNum = 1 to UBound(NewFollower)
							'もし記録時点で相互関係になったら、すでにcase1で記録済みなので記録させない
							If NewFollower(getNum,6) = "フォローした" Then
								recordFlag = FALSE
							Else
								recordFlag = TRUE
							End If

							If recordFlag = TRUE Then
								'書き込む行位置を特定する
								RowPotion = .ListRows.Count + 2

								'指定箇所の書式を設定
								.ListColumns(1).Range(RowPotion).NumberFormatLocal = "yyyy/mm/dd"
								.ListColumns(2).Range(RowPotion).NumberFormatLocal = "@"

									'配列内の値を出力
								.ListColumns(1).Range(RowPotion) = ActionNow
								.ListColumns(2).Range(RowPotion) = NewFollower(getNum,1)
								.ListColumns(3).Range(RowPotion) = NewFollower(getNum,2)
								.ListColumns(4).Range(RowPotion) = NewFollower(getNum,3)

									'フラグ設定
								.ListColumns(7).Range(RowPotion) = TRUE

									'フォロー、フォロワー数を設定
								.ListColumns(9).Range(RowPotion) = Sheets("follow").Range("B3")
								.ListColumns(10).Range(RowPotion) =Sheets("follower").Range("B3")

									'プロフィール画像を設定
								.ListColumns(11).Range(RowPotion) = NewFollower(getNum,4)
							End If
						Next
					End if

				Case 4
					if IsEmpty(DeleteFollower) = FALSE Then
						For getNum = 1 to UBound(DeleteFollower)
							'もし記録時点で凍結やブロ解の可能性がある判定になったら、すでにcase2で記録済みなので記録させない
							If DeleteFollower(getNum,6) = "フォローを解除" Then
								recordFlag = FALSE
							Else
								recordFlag = TRUE
							End If

							If recordFlag = TRUE Then
								'書き込む行位置を特定する
								RowPotion = .ListRows.Count + 2

								'指定箇所の書式を設定
								.ListColumns(1).Range(RowPotion).NumberFormatLocal = "yyyy/mm/dd"
								.ListColumns(2).Range(RowPotion).NumberFormatLocal = "@"

								'配列内の値を出力
								.ListColumns(1).Range(RowPotion) = ActionNow
								.ListColumns(2).Range(RowPotion) = DeleteFollower(getNum,1)
								.ListColumns(3).Range(RowPotion) = DeleteFollower(getNum,2)
								.ListColumns(4).Range(RowPotion) = DeleteFollower(getNum,3)

									'フラグ設定
								.ListColumns(8).Range(RowPotion) = TRUE

									'フォロー、フォロワー数を設定
								.ListColumns(9).Range(RowPotion) = Sheets("follow").Range("B3")
								.ListColumns(10).Range(RowPotion) =Sheets("follower").Range("B3")

									'プロフィール画像を設定
								.ListColumns(11).Range(RowPotion) = DeleteFollower(getNum,4)
							End If
						Next
					End If

				Case Else
					'none

			End Select
		End With
	Next

	'指定のテーブルに関数を入れ、テーブルの高さも調整する
	With Sheets("ReportLog").ListObjects(1)
		If .ListRows.Count > 0 Then
			.ListColumns(12).DataBodyRange.Formula = "=IFERROR(IMAGE(SUBSTITUTE([@プロフィール画像URL],""normal"",""400x400""),[@名前]&""さんのtwitterプロフィール画像です。"",0),""削除されたか、"" & CHAR(10) & ""凍結されたアカウントです"")"
			.ListColumns(13).DataBodyRange.Formula = "=COUNTIF([ID],[@ID])-1"
			.ListColumns(14).DataBodyRange.Formula = "=HYPERLINK(""https://twitter.com/intent/user?user_id="" & [@ID],""OPEN"")"

			'最終レコードの行位置を特定
			RowPotion = .ListRows.Count + 1

			'ヘッダーを除くデータ部のテーブル高さを70に指定してアイコンを見やすくする
			.ListColumns(1).DataBodyRange.RowHeight = 70

			'指定箇所へアクティブセルを移動させる
			.ListColumns(1).Range(RowPotion).select
		End If
	End With

	MsgBox "コンペア結果を記録しました。", vbOKOnly + vbInformation, "記録完了"

End Sub