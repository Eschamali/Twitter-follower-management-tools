Option Explicit

'���b�Z�[�W�p
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

	'FF�����擾����bat�t�@�C�������w��
	const useGetFFInfoBat = "getFFinfo_API.bat"

	'bat�I�����̑ҋ@�t���O
	pause = 0

	'�s�p�ӂɃA�J�E���g�����ς��Ȃ��悤�ɁA�t�@�C��������FF���X�g���擾����A�J�E���g������肷������ɂ��Ă܂��B
	TwitterAccount = Mid(ThisWorkbook.Name, 1, Len(ThisWorkbook.Name) - 5)

	'�ŏ��̃V�[�g�ɂ���`�F�b�NBOX�̏�Ԃ��擾����
	If Sheets("compare_results").CheckBox1.Value = FALSE Then
		'�m�F���b�Z�[�W���L���Ȃ̂ŁAbat�I������pause���L���ɂ���
		pause = 1
		continue = MsgBox("Twitter�A�J�E���g�F" & TwitterAccount & vbCrLf & "��L�A�J�E���g��FF�����擾���܂��B"& vbCrLf & vbCrLf & "��낵���ł����H", vbYesNo + vbQuestion, "FF�f�[�^�擾")
		If continue = vbNo Then
			exit sub
		End If
	End If

	'bat�N���pWshShell�N���X
	Set WSH = CreateObject("Wscript.Shell")

	'CurrentDir���擾����(onedrive�ۑ��ɂ��Ή�������)
	'�܂��͕��ʂ�CurrentDir���擾����֐���
	CurrentDir = ThisWorkbook.Path

	'http�����N�Ŏn�܂���̂́uonedrive�v���ϐ����g���ĕϊ�������
	'��������Ȃ��ꍇ�́A���̂܂܂ŁB���������ASharePoint�ł̋����͖��m�F
	If Left(CurrentDir,4) = "http" Then
		CurrentDir = Environ("OneDrive") & Mid(CurrentDir, 41)
	End If

	'�t�@�C���p�X��ǂݍ���
	FollowCSV = CurrentDir & "\input\" & TwitterAccount & "-follow.csv"
	FollowerCSV = CurrentDir & "\input\" & TwitterAccount & "-follower.csv"
	Before_FollowCSV = CurrentDir & "\input\" & TwitterAccount & "-before_follow.csv"
	Before_FollowerCSV = CurrentDir & "\input\" & TwitterAccount & "-before_follower.csv"
	BK_FollowCSV = CurrentDir & "\backup\" & TwitterAccount & "-follow.csv"
	BK_FollowerCSV = CurrentDir & "\backup\" & TwitterAccount & "-follower.csv"
	BK_Before_FollowCSV = CurrentDir & "\backup\" & TwitterAccount & "-before_follow.csv"
	BK_Before_FollowerCSV = CurrentDir & "\backup\" & TwitterAccount & "-before_follower.csv"

	'�J�����g�f�B���N�g����Twitter�A�J�E���g���usetting�v�V�[�g�ɓ\��t����
	'PowerQuery��CSV�ǂݍ��ݐ�Ɏg���܂��B
	With Sheets("setting").ListObjects(1)
		.DataBodyRange.AutoFilter 1, "FF���X�g�ǂݍ��ݐ�"
		.ListColumns(2).DataBodyRange = CurrentDir & "\input\"

		.DataBodyRange.AutoFilter 1, "Twitter�A�J�E���g"
		.ListColumns(2).DataBodyRange = TwitterAccount

		.DataBodyRange.AutoFilter 1
	End With

	'�~�X�����Ƃ��ɔ����āAinput�t�H���_�ɂ���csv�t�@�C����backup�ɃR�s�[(1�񕪂̂�)
	If Dir(Before_FollowCSV) <> "" Then
		FileCopy Before_FollowCSV, BK_Before_FollowCSV
	End If
	If Dir(Before_FollowerCSV) <> "" Then
		FileCopy Before_FollowerCSV, BK_Before_FollowerCSV
	End If

	If Dir(FollowCSV) <> "" Then
		FileCopy FollowCSV , BK_FollowCSV

		'�R�s�[��A�O��̃��X�gcsv�Ƃ��āAbefore�Ƀ��l�[��
		FileCopy FollowCSV , Before_FollowCSV

		'�R�s�[�����̂ō폜����
		Kill FollowCSV
	End If

	If Dir(FollowerCSV) <> "" Then
		FileCopy FollowerCSV, BK_FollowerCSV

		'�R�s�[��A�O��̃��X�gcsv�Ƃ��āAbefore�Ƀ��l�[��
		FileCopy FollowerCSV, Before_FollowerCSV

		'�R�s�[�����̂ō폜����
		Kill FollowerCSV
	End If

	'bat�t�@�C�����N������FF�����擾���܂��B
	'����1:�N���o�b�`�t�@�C���p�X 2:����bat�t�@�C���ɂ���f�B���N�g���p�X��cd���� 3:FF�����擾����A�J�E���g 4:bat�I������pause�R�}���h���g�����t���O
	WSH.Run CurrentDir & "\bat\" & useGetFFInfoBat & " " & CurrentDir & "\bat " & TwitterAccount & " " & pause, 1, True

	'�t�@�C�����݊m�F
	'Python�̏����̎d�����H�v���AFF�����擾����������Ƀt�@�C�����쐬�����̂ŁA�ȈՓI�Ɏ��s���肷��
	If Dir(FollowCSV) = "" OR Dir(FollowerCSV) = "" Then
		MsgBox "�ŐV��FF���̎擾�Ɏ��s���܂����B" & vbCrLf & "�������݂̌����s���A���N�G�X�g�̒��߂Ȃǂ��l�����܂��B", vbOKOnly + vbCritical, "FF���̎擾�Ɏ��s"
		exit sub
	End If

	'before�t�@�C�����Ȃ��ꍇ�A�擾�����t�@�C�����R�s�[������擾�ł���|�̃��b�Z�[�W��ǉ�����
	addMessage = ""
	If Dir(Before_FollowCSV) = "" AND Dir(Before_FollowerCSV) = "" Then
		FileCopy FollowCSV , Before_FollowCSV
		FileCopy FollowerCSV, Before_FollowerCSV
		addMessage = vbCrLf & vbCrLf & "������擾�̂��߁A�R���y�A���ʁA���O�L�^�́A����f�[�^�擾����" & vbCrLf & "FF�󋵂ɕω�������Ƃ��ɕ\�����܂��B"
	End If

	'�ŏ��̃V�[�g�ɂ���`�F�b�NBOX�̏�Ԃ��擾����
	If Sheets("compare_results").CheckBox1.Value = True Then
		dataRefresh
	else
		continue = MsgBox("Python��tweepy�ŁAFF���̎擾�ɐ������܂����B" & vbCrLf & "�����āA�f�[�^���X�V���܂�����낵���ł����H" & addMessage, vbYesNo + vbQuestion, "FF���̎擾����")
		If continue = vbYes Then
			dataRefresh
		End If
	End If
end sub


sub dataRefresh()
	'�쐬�����N�G���̐�
	Const quaryNum = 7

	Dim refreshNum(quaryNum)
	refreshNum(0) = "�N�G�� - follow"
	refreshNum(1) = "�N�G�� - follower"
	refreshNum(2) = "�N�G�� - before_follow"
	refreshNum(3) = "�N�G�� - before_follower"
	refreshNum(4) = "�N�G�� - newFollow"
	refreshNum(5) = "�N�G�� - deleteFollow"
	refreshNum(6) = "�N�G�� - newFollower"
	refreshNum(7) = "�N�G�� - deleteFollower"

	'OLEDB��ODBC�łȂ��ꍇ�G���[�ɂȂ�̂ŉ���̂��߂̂��Ȃ��Ȃ�
	On Error Resume Next
	Dim i
	Dim bolOLEDB,bolODBC
	For i = 0 to quaryNum - 4'�R�l�N�V�����̑S�Ă����[�v
		'BackgroundQuery�̏�Ԃ�ϐ��ɕۑ�
		bolOLEDB = ActiveWorkbook.Connections.OLEDBConnection.BackgroundQuery
		bolODBC = ActiveWorkbook.Connections.ODBCConnection.BackgroundQuery
		ActiveWorkbook.Connections.OLEDBConnection.BackgroundQuery = False
		ActiveWorkbook.Connections.ODBCConnection.BackgroundQuery = False
		'�X�V���s
		ActiveWorkbook.Connections(refreshNum(i)).Refresh
		'BackgroundQuery�̏�Ԃ����ɖ߂�
		ActiveWorkbook.Connections.OLEDBConnection.BackgroundQuery = bolOLEDB
		ActiveWorkbook.Connections.ODBCConnection.BackgroundQuery = bolODBC
	Next

	'�ŏ��̃V�[�g�ɂ���`�F�b�NBOX�̏�Ԃ��擾����
	If Sheets("compare_results").CheckBox1.Value = FALSE Then
		MsgBox "�擾����FF��񂩂�f�[�^���X�V���܂����B" & vbCrLf & "�R���y�A���ʂ�\�����܂��B" & addMessage, vbOKOnly + vbInformation, "���f����"
	End If

	For i = 4 to quaryNum '�R�l�N�V�����̑S�Ă����[�v
		'BackgroundQuery�̏�Ԃ�ϐ��ɕۑ�
		bolOLEDB = ActiveWorkbook.Connections.OLEDBConnection.BackgroundQuery
		bolODBC = ActiveWorkbook.Connections.ODBCConnection.BackgroundQuery
		ActiveWorkbook.Connections.OLEDBConnection.BackgroundQuery = False
		ActiveWorkbook.Connections.ODBCConnection.BackgroundQuery = False
		'�X�V���s
		ActiveWorkbook.Connections(refreshNum(i)).Refresh
		'BackgroundQuery�̏�Ԃ����ɖ߂�
		ActiveWorkbook.Connections.OLEDBConnection.BackgroundQuery = bolOLEDB
		ActiveWorkbook.Connections.ODBCConnection.BackgroundQuery = bolODBC
	Next

	On Error GoTo 0

	'�w��̃e�[�u���Ɋ֐�������
	With Sheets("compare_results")
		If .ListObjects(1).ListRows.Count > 0 Then
			.ListObjects(1).ListColumns(6).DataBodyRange.Formula = "=XLOOKUP([@ID],newFollower[ID],newFollower[�O��Ƃ̃R���y�A],""No"")"
		End If

		If .ListObjects(2).ListRows.Count > 0 Then
			.ListObjects(2).ListColumns(6).DataBodyRange.Formula = "=XLOOKUP([@ID],deleteFollower[ID],deleteFollower[�O��Ƃ̃R���y�A],""No"")"
		End If

		If .ListObjects(3).ListRows.Count > 0 Then
			.ListObjects(3).ListColumns(6).DataBodyRange.Formula = "=XLOOKUP([@ID],newFollow[ID],newFollow[�O��Ƃ̃R���y�A],""No"")"
		End If

		If .ListObjects(4).ListRows.Count > 0 Then
			.ListObjects(4).ListColumns(6).DataBodyRange.Formula = "=XLOOKUP([@ID],deleteFollow[ID],deleteFollow[�O��Ƃ̃R���y�A],""No"")"
		End If
	End With

	With Sheets("follow")
		If .ListObjects(1).ListRows.Count > 0 Then
			.ListObjects(1).ListColumns(17).DataBodyRange.Formula = "=SUBSTITUTE(XLOOKUP([@ID],before_follow[ID],before_follow[ID],""�t�H���[����"",0,1),[@ID],""�ω��Ȃ�"")"
			.ListObjects(1).ListColumns(18).DataBodyRange.Formula = "=SUBSTITUTE(XLOOKUP([@ID],follower[ID],follower[ID],""�Ўv��"",0,1),[@ID],""FF�֌W"")"
		End If
	End With

	With Sheets("follower")
		If .ListObjects(1).ListRows.Count > 0 Then
			.ListObjects(1).ListColumns(17).DataBodyRange.Formula = "=SUBSTITUTE(XLOOKUP([@ID],before_follower[ID],before_follower[ID],""�t�H���[���ꂽ""),[@ID],""�ω��Ȃ�"")"
			.ListObjects(1).ListColumns(18).DataBodyRange.Formula = "=SUBSTITUTE(XLOOKUP([@ID],follow[ID],follow[ID],""�Ўv���""),[@ID],""FF�֌W"")"
		End If
	End With

	With Sheets("before_follow")
		If .ListObjects(1).ListRows.Count > 0 Then
			.ListObjects(1).ListColumns(17).DataBodyRange.Formula = "=SUBSTITUTE(XLOOKUP([@ID],follow[ID],follow[ID],""�t�H���[������"",0,1),[@ID],""�ω��Ȃ�"")"
		End If
	End With

	With Sheets("before_follower")
		If .ListObjects(1).ListRows.Count > 0 Then
			.ListObjects(1).ListColumns(17).DataBodyRange.Formula = "=SUBSTITUTE(XLOOKUP([@ID],follower[ID],follower[ID],""�t�H���[���������ꂽ""),[@ID],""�ω��Ȃ�"")"
		End If
	End With

	'�ŏ��̃V�[�g�ɂ���`�F�b�NBOX�̏�Ԃ��擾����
	If Sheets("compare_results").CheckBox1.Value = True Then
		Sheets("ReportLog").select
		compare
	else
		Dim continue

		continue = MsgBox("�R���y�A���ʂ�\�����܂����B" & vbCrLf & "�����āA�R���y�A���ʂ��L�^���܂�����낵���ł����H" & addMessage, vbYesNo + vbQuestion, "�X�V����")
		If continue = vbYes Then
			Sheets("ReportLog").select
			compare
		End If

	End If

End sub


Sub compare()
	'�e���j�[�NID���i�[
	Dim NewFollow
	Dim DeleteFollow
	Dim NewFollower
	Dim DeleteFollower

	'���t���擾
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

	'�R���y�A���ʎ��Ԃ��L�^
	Sheets("compare_results").Cells(1,2).value = ActionNow

	'���ʂ��܂Ƃ߂�
	Dim i
	for i = 1 to 4

		'ListObjects���g�����e�[�u������
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

	'���ʂ��o�͂���
	for i = 1 to 4

		'ListObjects���g�����e�[�u������
		With Sheets("ReportLog").ListObjects(1)

			Dim RowPotion, getNum,recordFlag
			Select Case i
				Case 1
					if IsEmpty(NewFollow) = FALSE Then
						For getNum = 1 to UBound(NewFollow)
							'�������ލs�ʒu����肷��
							RowPotion = .ListRows.Count + 2

							'�w��ӏ��̏�����ݒ�
							.ListColumns(1).Range(RowPotion).NumberFormatLocal = "yyyy/mm/dd"
							.ListColumns(2).Range(RowPotion).NumberFormatLocal = "@"

							'�z����̒l���o��
							.ListColumns(1).Range(RowPotion) = ActionNow
							.ListColumns(2).Range(RowPotion) = NewFollow(getNum,1)
							.ListColumns(3).Range(RowPotion) = NewFollow(getNum,2)
							.ListColumns(4).Range(RowPotion) = NewFollow(getNum,3)

								'�t���O�ݒ�
							.ListColumns(5).Range(RowPotion) = TRUE

							'�����L�^���_�œ����ɑ��݊֌W�ɂȂ�����A���ꃌ�R�[�h�ɋL�^����
							If NewFollow(getNum,6) = "�t�H���[���ꂽ" Then
								.ListColumns(7).Range(RowPotion) = TRUE
							End If

							'�t�H���[�A�t�H�����[����ݒ�
							.ListColumns(9).Range(RowPotion) = Sheets("follow").Range("B3")
							.ListColumns(10).Range(RowPotion) = Sheets("follower").Range("B3")

							'�v���t�B�[���摜��ݒ�
							.ListColumns(11).Range(RowPotion) = NewFollow(getNum,4)
						Next
					End if

				Case 2
					if IsEmpty(DeleteFollow) = FALSE Then
						For getNum = 1 to UBound(DeleteFollow)
							'�������ލs�ʒu����肷��
							RowPotion = .ListRows.Count + 2

							'�w��ӏ��̏�����ݒ�
							.ListColumns(1).Range(RowPotion).NumberFormatLocal = "yyyy/mm/dd"
							.ListColumns(2).Range(RowPotion).NumberFormatLocal = "@"

							'�z����̒l���o��
							.ListColumns(1).Range(RowPotion) = ActionNow
							.ListColumns(2).Range(RowPotion) = DeleteFollow(getNum,1)
							.ListColumns(3).Range(RowPotion) = DeleteFollow(getNum,2)
							.ListColumns(4).Range(RowPotion) = DeleteFollow(getNum,3)

								'�t���O�ݒ�
							.ListColumns(6).Range(RowPotion) = TRUE

							'�����L�^���_�œ�����u�����̉\�������锻��ɂȂ�����A���ꃌ�R�[�h�ɋL�^����
							If DeleteFollow(getNum,6) = "�t�H���[���������ꂽ" Then
								.ListColumns(8).Range(RowPotion) = TRUE
							End If

								'�t�H���[�A�t�H�����[����ݒ�
							.ListColumns(9).Range(RowPotion) = Sheets("follow").Range("B3")
							.ListColumns(10).Range(RowPotion) = Sheets("follower").Range("B3")

								'�v���t�B�[���摜��ݒ�
							.ListColumns(11).Range(RowPotion) = DeleteFollow(getNum,4)
						Next
					End if

				Case 3
					if IsEmpty(NewFollower) = FALSE Then
						For getNum = 1 to UBound(NewFollower)
							'�����L�^���_�ő��݊֌W�ɂȂ�����A���ł�case1�ŋL�^�ς݂Ȃ̂ŋL�^�����Ȃ�
							If NewFollower(getNum,6) = "�t�H���[����" Then
								recordFlag = FALSE
							Else
								recordFlag = TRUE
							End If

							If recordFlag = TRUE Then
								'�������ލs�ʒu����肷��
								RowPotion = .ListRows.Count + 2

								'�w��ӏ��̏�����ݒ�
								.ListColumns(1).Range(RowPotion).NumberFormatLocal = "yyyy/mm/dd"
								.ListColumns(2).Range(RowPotion).NumberFormatLocal = "@"

									'�z����̒l���o��
								.ListColumns(1).Range(RowPotion) = ActionNow
								.ListColumns(2).Range(RowPotion) = NewFollower(getNum,1)
								.ListColumns(3).Range(RowPotion) = NewFollower(getNum,2)
								.ListColumns(4).Range(RowPotion) = NewFollower(getNum,3)

									'�t���O�ݒ�
								.ListColumns(7).Range(RowPotion) = TRUE

									'�t�H���[�A�t�H�����[����ݒ�
								.ListColumns(9).Range(RowPotion) = Sheets("follow").Range("B3")
								.ListColumns(10).Range(RowPotion) =Sheets("follower").Range("B3")

									'�v���t�B�[���摜��ݒ�
								.ListColumns(11).Range(RowPotion) = NewFollower(getNum,4)
							End If
						Next
					End if

				Case 4
					if IsEmpty(DeleteFollower) = FALSE Then
						For getNum = 1 to UBound(DeleteFollower)
							'�����L�^���_�œ�����u�����̉\�������锻��ɂȂ�����A���ł�case2�ŋL�^�ς݂Ȃ̂ŋL�^�����Ȃ�
							If DeleteFollower(getNum,6) = "�t�H���[������" Then
								recordFlag = FALSE
							Else
								recordFlag = TRUE
							End If

							If recordFlag = TRUE Then
								'�������ލs�ʒu����肷��
								RowPotion = .ListRows.Count + 2

								'�w��ӏ��̏�����ݒ�
								.ListColumns(1).Range(RowPotion).NumberFormatLocal = "yyyy/mm/dd"
								.ListColumns(2).Range(RowPotion).NumberFormatLocal = "@"

								'�z����̒l���o��
								.ListColumns(1).Range(RowPotion) = ActionNow
								.ListColumns(2).Range(RowPotion) = DeleteFollower(getNum,1)
								.ListColumns(3).Range(RowPotion) = DeleteFollower(getNum,2)
								.ListColumns(4).Range(RowPotion) = DeleteFollower(getNum,3)

									'�t���O�ݒ�
								.ListColumns(8).Range(RowPotion) = TRUE

									'�t�H���[�A�t�H�����[����ݒ�
								.ListColumns(9).Range(RowPotion) = Sheets("follow").Range("B3")
								.ListColumns(10).Range(RowPotion) =Sheets("follower").Range("B3")

									'�v���t�B�[���摜��ݒ�
								.ListColumns(11).Range(RowPotion) = DeleteFollower(getNum,4)
							End If
						Next
					End If

				Case Else
					'none

			End Select
		End With
	Next

	'�w��̃e�[�u���Ɋ֐������A�e�[�u���̍�������������
	With Sheets("ReportLog").ListObjects(1)
		If .ListRows.Count > 0 Then
			.ListColumns(12).DataBodyRange.Formula = "=IFERROR(IMAGE(SUBSTITUTE([@�v���t�B�[���摜URL],""normal"",""400x400""),[@���O]&""�����twitter�v���t�B�[���摜�ł��B"",0),""�폜���ꂽ���A"" & CHAR(10) & ""�������ꂽ�A�J�E���g�ł�"")"
			.ListColumns(13).DataBodyRange.Formula = "=COUNTIF([ID],[@ID])-1"
			.ListColumns(14).DataBodyRange.Formula = "=HYPERLINK(""https://twitter.com/intent/user?user_id="" & [@ID],""OPEN"")"

			'�ŏI���R�[�h�̍s�ʒu�����
			RowPotion = .ListRows.Count + 1

			'�w�b�_�[�������f�[�^���̃e�[�u��������70�Ɏw�肵�ăA�C�R�������₷������
			.ListColumns(1).DataBodyRange.RowHeight = 70

			'�w��ӏ��փA�N�e�B�u�Z�����ړ�������
			.ListColumns(1).Range(RowPotion).select
		End If
	End With

	MsgBox "�R���y�A���ʂ��L�^���܂����B", vbOKOnly + vbInformation, "�L�^����"

End Sub