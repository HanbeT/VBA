Attribute VB_Name = "MainAdjustMergeCell"

Public Sub MainAdjustMergeCellH()
    
    ' ��ʕ`��X�V�̒�~
    Application.ScreenUpdating = False
    ' �m�F���b�Z�[�W���\��
    Application.DisplayAlerts = Flase
    
    Dim laa As LogiAdjustMergeCellH
    Set laa = New LogiAdjustMergeCellH
    Call laa.mainLogic

    ' �m�F���b�Z�[�W��\��
    Application.DisplayAlerts = True
    ' ��ʕ`��X�V�̎��s
    Application.ScreenUpdating = True
    
End Sub
