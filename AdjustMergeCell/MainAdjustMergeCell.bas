Attribute VB_Name = "MainAdjustMergeCell"

Public Sub MainAdjustMergeCellH()
    
    ' 画面描画更新の停止
    Application.ScreenUpdating = False
    ' 確認メッセージを非表示
    Application.DisplayAlerts = Flase
    
    Dim laa As LogiAdjustMergeCellH
    Set laa = New LogiAdjustMergeCellH
    Call laa.mainLogic

    ' 確認メッセージを表示
    Application.DisplayAlerts = True
    ' 画面描画更新の実行
    Application.ScreenUpdating = True
    
End Sub
