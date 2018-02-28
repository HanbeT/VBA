Attribute VB_Name = "ReCheckInputRule"
Option Explicit

Public Sub MainReCheckInputRule()
    
    ' 画面描画更新の停止
    Application.ScreenUpdating = False
    ' 確認メッセージを非表示
    Application.DisplayAlerts = False
        
    Call mainLogic
    
    ' 確認メッセージを表示
    Application.DisplayAlerts = True
    ' 画面描画更新の実行
    Application.ScreenUpdating = True
    
End Sub

Private Function mainLogic()
    
    Dim tSheet As Worksheet ' 対象シート
    Dim ra As Range         ' 対象レンジ
    Dim target As String    ' 選択範囲
    Dim checkCnt As Integer ' チェック数
    Dim errCnt As Integer   ' エラー数
    
    ' 対象シート取得
    Set tSheet = ActiveSheet
    ' エラー数初期化
    errCnt = 0
    
    ' 選択範囲取得
    target = Selection.Address
    
    ' 対象範囲のチェックを行う
    For Each ra In tSheet.Range(target)
        checkCnt = checkCnt + 1
        ' 入力規則に合わない場合
        If Not ra.Validation.Value Then
            errCnt = errCnt + 1
            ra.Interior.ColorIndex = 3
        End If
    Next ra
    
    ' メッセージ出力
    If errCnt > 0 Then
        ' 画面描画更新の実行
        Application.ScreenUpdating = True
        MsgBox "入力規則違反のセルが存在します。" & vbCrLf & _
               "  対象数：" & checkCnt & vbCrLf & _
               "  違反数：" & errCnt, vbCritical, "エラー通知"
    Else
        MsgBox "チェックが完了しました。", vbInformation, "完了通知"
    End If

End Function
