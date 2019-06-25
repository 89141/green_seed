Attribute VB_Name = "Module2"
Sub exportCSV()
Attribute exportCSV.VB_ProcData.VB_Invoke_Func = " \n14"
'
' exportCSV Macro
'
 
Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
 
Dim csvFile As String
csvFile = ActiveWorkbook.Path & "\" & ws.Name & ".csv"
 
'ADODB.Stream�I�u�W�F�N�g�𐶐�
Dim adoSt As Object
Set adoSt = CreateObject("ADODB.Stream")
 
Dim strLine As String
Dim i As Long, j As Long, x As Long
i = 1

With adoSt
    .Charset = "UTF-8"
    .LineSeparator = -1
    .Open
    
    j = 1
    Do While ws.Cells(1, j + 1).Value <> ""
        x = j
        j = j + 1
    Loop
    
    Do While ws.Cells(i, 1).Value <> ""
 
        strLine = ""
 
        j = 1
        Do While j < x + 1
 
            strLine = strLine & ws.Cells(i, j).Value & ","
            j = j + 1
 
        Loop
 
        strLine = strLine & ws.Cells(i, j).Value
 
        .WriteText strLine, 1
 
        i = i + 1
 
    Loop
    
    .Position = 0 '�X�g���[���̈ʒu��0�ɂ���
    .Type = 1 '�f�[�^�̎�ނ��o�C�i���f�[�^�ɕύX
    .Position = 3 '�X�g���[���̈ʒu��3�ɂ���
 
    Dim byteData() As Byte '�ꎞ�i�[�p
    byteData = .Read '�X�g���[���̓��e���ꎞ�i�[�p�ϐ��ɕۑ�
    .Close '��U�X�g���[�������i���Z�b�g�j
 
    .Open '�X�g���[�����J��
    .Write byteData '�X�g���[���Ɉꎞ�i�[�����f�[�^�𗬂�����
    .SaveToFile csvFile, 2
    .Close
 
End With
Next
 
MsgBox "CSV�ɏ����o���܂���"
 
End Sub
