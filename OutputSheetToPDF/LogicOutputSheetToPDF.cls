VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogicOutputSheetToPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function mainLogic()
On Error GoTo ErrRoute
    Dim tBook As Workbook
    Dim tSheet As Worksheet
    
    Dim targetFile As String
    
    Dim cu As ComUtil
    Set cu = New ComUtil
    
    targetFile = cu.selectFile("")
    
    
    ' �Ώۃu�b�N�I�[�v��
    Set tBook = Workbooks.Open(targetFile)
    
    For Each tSheet In tBook.Worksheets
        ' ����ݒ�
        With tSheet.PageSetup
            ' �g��/�k��
            .Zoom = 88
            ' �y�[�W���ɍ��킹�Ĉ������(�c)
            '.FitToPagesTall = 1
            ' �y�[�W���ɍ��킹�Ĉ������(��)
            '.FitToPagesWide = 1
        End With
    Next tSheet
    
    ' ���[�N�V�[�g�S�I��
    tBook.Worksheets.Select
    ' PDF�o��
    ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:=tBook.Path & "\" & cu.excludeExtension(tBook.Name) & ".pdf"
    
    ' �Ώۃu�b�N�N���[�Y
    Set tSheet = Nothing
    tBook.Close SaveChanges:=False
    Set tBook = Nothing
    
    Exit Function
    
' �G���[���[�g
ErrRoute:
    If Not tSheet Is Nothing Then
        Set tSheet = Nothing
    End If
    If Not tBook Is Nothing Then
        Set tBook = Nothing
    End If
End Function