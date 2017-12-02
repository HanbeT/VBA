VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogiAdjustMergeCellH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function mainLogic()
    
    ' �����J�n��
    Dim nowCol As Integer
    ' �����J�n�s
    Dim nowRow As Integer
    ' �����I����
    Dim endCol As Integer
    ' �����I���s
    Dim endRow As Integer
    ' �����ΏۃZ���I�u�W�F�N�g
    Dim ra As Range
    
    ' ���[�v�p�ϐ�
    Dim i As Integer
    Dim j As Integer
    
    ' �����J�n�s����擾
    nowCol = Selection(1).Column
    nowRow = Selection(1).Row
    
    ' �����I���s����擾
    endCol = Selection(Selection.Count).Column
    endRow = Selection(Selection.Count).Row
    
    For i = nowCol To endCol
        For j = nowRow To endRow
            ' �Z���I�u�W�F�N�g�擾
            Set ra = Cells(j, i)
            If ra.MergeArea.Address Like ra.Address & "*" Then
                Call adjust(ra)
                Application.StatusBar = ra.Address
            End If
            Set ra = Nothing
        Next j
    Next i
    ' �X�e�[�^�X�o�[������
    Application.StatusBar = False
    
End Function

'*********************************************************
' �֐����F���������֐�
' �T  �v�F�����̒������s��
' ��  ���FRange�I�u�W�F�N�g
' �߂�l�F�Ȃ�
'*********************************************************
Private Function adjust(aRa As Range)
    
    ' �������Z����
    Dim mrgCellW As Double
    ' �����͈̓A�h���X
    Dim mrgAdd As String
    ' �����������Z����
    Dim indCellW As Double
    ' �Z���s��
    Dim indCellH As Double
    ' �V�K�s��
    Dim newCellH As Double
    
    ' ���[�v�p�ϐ�
    Dim i As Integer
    Dim j As Integer
    
    ' �������Z���񕝂��擾
    mrgCellW = getMergeCellWidth(aRa)
    ' �����͈͂��擾
    mrgAdd = aRa.MergeArea.Address
    
    ' ����������
    aRa.UnMerge
    
    ' �Z���񕝂̏����l���擾
    indCellW = aRa.ColumnWidth
    ' �Z���s���̏����l���擾
    indCellH = aRa.RowHeight
    ' �������̃Z�����ɍ��킹��
    aRa.ColumnWidth = mrgCellW
    ' �܂�Ԃ���ݒ�
    aRa.WrapText = True
    ' �Z���s����������
    aRa.EntireRow.AutoFit
    ' �V�K�s�����擾
    newCellH = aRa.RowHeight
    
    ' �񕝂ɏ����l��ݒ�
    aRa.ColumnWidth = indCellW
    ' �Č���
    Range(mrgAdd).Merge
    
    If indCellH > newCellH Then
        ' �����s����ݒ�
        aRa.RowHeight = indCellH
    Else
        ' �V�K�s����ݒ�
        aRa.RowHeight = newCellH
    End If
    
End Function

'*********************************************************
' �֐����F�������Z�����擾�֐�
' �T  �v�F�������Ă���Z���̗񕝂̍��v���擾����B
' ��  ���F�ΏۃZ���I�u�W�F�N�g
' �߂�l�F������
'*********************************************************
Private Function getMergeCellWidth(aRa As Range) As Double
    
    Dim res As Double
    Dim indCell As Range
    For Each indCell In aRa.MergeArea.Cells
        res = res + indCell.ColumnWidth
    Next indCell
    getMergeCellWidth = res
    
End Function