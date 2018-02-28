Attribute VB_Name = "ReCheckInputRule"
Option Explicit

Public Sub MainReCheckInputRule()
    
    ' ��ʕ`��X�V�̒�~
    Application.ScreenUpdating = False
    ' �m�F���b�Z�[�W���\��
    Application.DisplayAlerts = False
        
    Call mainLogic
    
    ' �m�F���b�Z�[�W��\��
    Application.DisplayAlerts = True
    ' ��ʕ`��X�V�̎��s
    Application.ScreenUpdating = True
    
End Sub

Private Function mainLogic()
    
    Dim tSheet As Worksheet ' �ΏۃV�[�g
    Dim ra As Range         ' �Ώۃ����W
    Dim target As String    ' �I��͈�
    Dim checkCnt As Integer ' �`�F�b�N��
    Dim errCnt As Integer   ' �G���[��
    
    ' �ΏۃV�[�g�擾
    Set tSheet = ActiveSheet
    ' �G���[��������
    errCnt = 0
    
    ' �I��͈͎擾
    target = Selection.Address
    
    ' �Ώ۔͈͂̃`�F�b�N���s��
    For Each ra In tSheet.Range(target)
        checkCnt = checkCnt + 1
        ' ���͋K���ɍ���Ȃ��ꍇ
        If Not ra.Validation.Value Then
            errCnt = errCnt + 1
            ra.Interior.ColorIndex = 3
        End If
    Next ra
    
    ' ���b�Z�[�W�o��
    If errCnt > 0 Then
        ' ��ʕ`��X�V�̎��s
        Application.ScreenUpdating = True
        MsgBox "���͋K���ᔽ�̃Z�������݂��܂��B" & vbCrLf & _
               "  �Ώې��F" & checkCnt & vbCrLf & _
               "  �ᔽ���F" & errCnt, vbCritical, "�G���[�ʒm"
    Else
        MsgBox "�`�F�b�N���������܂����B", vbInformation, "�����ʒm"
    End If

End Function
