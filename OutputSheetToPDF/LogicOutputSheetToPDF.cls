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
    
    
    ' 対象ブックオープン
    Set tBook = Workbooks.Open(targetFile)
    
    For Each tSheet In tBook.Worksheets
        ' 印刷設定
        With tSheet.PageSetup
            ' 拡大/縮小
            .Zoom = 88
            ' ページ数に合わせて印刷する(縦)
            '.FitToPagesTall = 1
            ' ページ数に合わせて印刷する(横)
            '.FitToPagesWide = 1
        End With
    Next tSheet
    
    ' ワークシート全選択
    tBook.Worksheets.Select
    ' PDF出力
    ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:=tBook.Path & "\" & cu.excludeExtension(tBook.Name) & ".pdf"
    
    ' 対象ブッククローズ
    Set tSheet = Nothing
    tBook.Close SaveChanges:=False
    Set tBook = Nothing
    
    Exit Function
    
' エラールート
ErrRoute:
    If Not tSheet Is Nothing Then
        Set tSheet = Nothing
    End If
    If Not tBook Is Nothing Then
        Set tBook = Nothing
    End If
End Function
