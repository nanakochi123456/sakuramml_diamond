' ���̂܂܎��s�ł��܂��B
'
main

Sub main()
	Dim lp
	Dim Bgm
	Dim title
	Dim vtitle
	Dim kasi

	title="�uLet's Mining�v�̎��W�F�l���[�^ (c)2014 NEET�������"
	vtitle="�_�C�����@�낤�I�uLet's Mining�vVOCALOID�p�̎�" & vbcrlf & "Copyright 2014 NEET�������" & vbcrlf

	Bgm=MsgBox("BGM���[�h�ł����H(�͂�) / �̕t�ł����H(������)",vbYesNo,title)
	If Bgm=vbYes Then
		lp=InputBox("���[�v���𐔎��œ��͂��ĉ�����",title)
	Else
		lp=2
	End If

	On Error Resume Next
	If lp=0 And Bgm=vbYes Then
		MsgBox "�l��0�Ȃ�", vbOKOnly + vbCritical,title
		Exit Sub
	End If

	If lp>99 Then
		MsgBox "�l��" & lp & "�Ƌ��傷���āA�������ɂ͏����ł��Ȃ��́B���������B�B�B", vbOKOnly + vbCritical,title
		Exit Sub
	End If

	On Error Goto 0

	kasi=makekashi(Bgm,lp)

	Dim filename
	filename="diamond_vocaloid_kashi.txt"
	Set objFso = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFso.OpenTextFile(filename, 2, True)

	If Err.Number > 0 Then
		MsgBox filename & "�ɏ������݂ł��Ȃ������Ȃ́G�G���������G�G", vbOKOnly + vbCritical ,title
	Else
		objFile.Write vtitle
		objFile.Write kasi
		objFile.Close
		Set objFile = Nothing
		Set objFso = Nothing
		MsgBox filename & "�ɏ������݂ł����Ȃ�", vbOKOnly + vbInformation,title
	End If
End Sub

Function skashi(str,count)
	Dim kashi
	For i=1 To count
		kashi=kashi+str
	Next
	skashi=kashi
End Function

Function kashititle(track)
	Dim str
	str="-----------------------------------------------" & vbcrlf
	str=str+"Track " & track & vbcrlf & vbcrlf
	kashititle=str
End Function

Function makekashi(bgm,lp)
	Dim kashi
	Dim mkashi


	If bgm=vbNo Then
		mkashi=mkashi+"���Ă�@������ �悱��"
		mkashi=mkashi+"��͂� �ق肷����"
		mkashi=mkashi+"�Ă点���@���܂@����"
		mkashi=mkashi+"��̂��� �ق肷����"
		mkashi=mkashi+"�����˂� ������"
		mkashi=mkashi+"�Ђ낦�� ����  ��"
		mkashi=mkashi+"�ق肷����"
		mkashi=mkashi+"���ĂɁ@������ �悱��"
		mkashi=mkashi+"��͂� �ق肷����"
		mkashi=mkashi+"�Ђ낦�� ���܂� ����"
		mkashi=mkashi+"��̂��� �ق肷����"
		mkashi=mkashi+"�R���{�� ���܂�"
		mkashi=mkashi+"���}���� �Ђ���"
		mkashi=mkashi+"�p�Y���݂��� �����������您"
		mkashi=mkashi+"�W���N��@�ǂ��ł�"
		mkashi=mkashi+"�N�C��@�ȂȂ�"
		mkashi=mkashi+"�L���O�� ��ʂ�"
		mkashi=mkashi+"�W���J�� �悢"
		mkashi=mkashi+"�������Ձ@����"
		mkashi=mkashi+"��͂��@����"
		mkashi=mkashi+"����΂� ��ʂ�"
		mkashi=mkashi+"�_�C����������"
		mkashi=mkashi+"�����̂����Ձ@�Ă�"
		mkashi=mkashi+"�������� �ق肷����"
		mkashi=mkashi+"����̂����Ձ@������"
		mkashi=mkashi+"�̂����Ձ@�ق肷����"
		mkashi=mkashi+"�����˂� ��͂�"
		mkashi=mkashi+"�����悽���܂�"
		mkashi=mkashi+"�ق肷����"
		mkashi=mkashi+"�����̂� ��͂��@�Ă�"
		mkashi=mkashi+"��͂� �ق肷����"
		mkashi=mkashi+"����̂@��͂��@������"
		mkashi=mkashi+"�̂�͂� �ق肷����"
		mkashi=mkashi+"�R���{�� ���߂�"
		mkashi=mkashi+"���}���� �Ђ��点"
		mkashi=mkashi+"�p�Y���݂��� �����������您"
		mkashi=mkashi+"�G�X�Ł@����"
		mkashi=mkashi+"�L���O���@�ق��"
		mkashi=mkashi+"�R���{�� ���߂�"
		mkashi=mkashi+"���Ⴍ�Ă�˂炦"
		mkashi=mkashi+"�X�R�b�v�@����"
		mkashi=mkashi+"��͂��@����"
		mkashi=mkashi+"����΂� ��ʂ�"
		mkashi=mkashi+"�_�C����������"

		kashi=kashi+kashititle(1)
		kashi=kashi+"�ق남"
		kashi=kashi+mkashi
		kashi=kashi+vbcrlf

		kashi=kashi+kashititle(2)
		kashi=kashi+"������"
		kashi=kashi+mkashi
		kashi=kashi+vbcrlf

		kashi=kashi+kashititle("3/4")
		kashi=kashi+"������"
		kashi=kashi+vbcrlf

	End If
	'Track 7/8/9
	kashi=kashi+kashititle("5/6/7")

	Dim Hatsu

	If bgm=vbNo Then
		Hatsu="����"
	Else
		Hatsu="��"
	End If

	'Intro
	kashi=kashi+skashi("����",3*4)
	kashi=kashi+skashi("��",22)

	For i=1 To lp
		If i<>1 Then
			'sub
			kashi=kashi+skashi(Hatsu,6)
			kashi=kashi+skashi("��",10)
		End If

		'melody1
		kashi=kashi+skashi(Hatsu,7*4)
		kashi=kashi+skashi("��",12)
		kashi=kashi+skashi(Hatsu,2)
		kashi=kashi+skashi("��",4)

		'melody2
		kashi=kashi+skashi(Hatsu,4)
		kashi=kashi+skashi("��",8)
		kashi=kashi+skashi(Hatsu,8)
		kashi=kashi+skashi("��",4)
	Next

	'end
	kashi=kashi+skashi("����",14)
	kashi=kashi+skashi("��",1)

	kashi=kashi+vbcrlf+vbcrlf

	'Track 8
	kashi=kashi+kashititle(8)
	kashi=kashi+skashi("��",1)
	kashi=kashi+vbcrlf+vbcrlf

	'Track 9
	kashi=kashi+kashititle(9)

	'Intro
	kashi=kashi+skashi("�Ƃ�",32)

	For i=1 To lp
		If i<>1 Then
			'sub
			kashi=kashi+skashi("�Ƃ�",16)
		End If

		'melody1
		kashi=kashi+skashi("�Ƃ�",15*4+12)

		'melody2
		kashi=kashi+skashi("�Ƃ�",5+4+4+4+5+5+4+4)
	Next

	'end
	kashi=kashi+skashi("�Ƃ�",6*4+2+8)
	kashi=kashi+skashi("��",1)
	kashi=kashi+vbcrlf+vbcrlf

	makekashi=kashi
End Function


' ���킲�ƁE�E�E
' SAKURAMML�ł��A���̋@�\�ł������Ȃ�ł����A
' �f�o�b�O�E�B���h�E�ł͂Ȃ��A�����ƃe�L�X�g�G�f�B�^��
' �J���������y�Ȃ̂ŁE�E�E
