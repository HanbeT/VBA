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
    
    ' 処理開始列
    Dim nowCol As Integer
    ' 処理開始行
    Dim nowRow As Integer
    ' 処理終了列
    Dim endCol As Integer
    ' 処理終了行
    Dim endRow As Integer
    ' 処理対象セルオブジェクト
    Dim ra As Range
    
    ' ループ用変数
    Dim i As Integer
    Dim j As Integer
    
    ' 処理開始行列情報取得
    nowCol = Selection(1).Column
    nowRow = Selection(1).Row
    
    ' 処理終了行列情報取得
    endCol = Selection(Selection.Count).Column
    endRow = Selection(Selection.Count).Row
    
    For i = nowCol To endCol
        For j = nowRow To endRow
            ' セルオブジェクト取得
            Set ra = Cells(j, i)
            If ra.MergeArea.Address Like ra.Address & "*" Then
                Call adjust(ra)
                Application.StatusBar = ra.Address
            End If
            Set ra = Nothing
        Next j
    Next i
    ' ステータスバー初期化
    Application.StatusBar = False
    
End Function

'*********************************************************
' 関数名：高さ調整関数
' 概  要：高さの調整を行う
' 引  数：Rangeオブジェクト
' 戻り値：なし
'*********************************************************
Private Function adjust(aRa As Range)
    
    ' 結合時セル列幅
    Dim mrgCellW As Double
    ' 結合範囲アドレス
    Dim mrgAdd As String
    ' 結合解除時セル列幅
    Dim indCellW As Double
    ' セル行高
    Dim indCellH As Double
    ' 新規行高
    Dim newCellH As Double
    
    ' ループ用変数
    Dim i As Integer
    Dim j As Integer
    
    ' 結合時セル列幅を取得
    mrgCellW = getMergeCellWidth(aRa)
    ' 結合範囲を取得
    mrgAdd = aRa.MergeArea.Address
    
    ' 結合を解除
    aRa.UnMerge
    
    ' セル列幅の初期値を取得
    indCellW = aRa.ColumnWidth
    ' セル行高の初期値を取得
    indCellH = aRa.RowHeight
    ' 結合時のセル幅に合わせる
    aRa.ColumnWidth = mrgCellW
    ' 折り返しを設定
    aRa.WrapText = True
    ' セル行高自動調整
    aRa.EntireRow.AutoFit
    ' 新規行高を取得
    newCellH = aRa.RowHeight
    
    ' 列幅に初期値を設定
    aRa.ColumnWidth = indCellW
    ' 再結合
    Range(mrgAdd).Merge
    
    If indCellH > newCellH Then
        ' 既存行高を設定
        aRa.RowHeight = indCellH
    Else
        ' 新規行高を設定
        aRa.RowHeight = newCellH
    End If
    
End Function

'*********************************************************
' 関数名：結合時セル幅取得関数
' 概  要：結合しているセルの列幅の合計を取得する。
' 引  数：対象セルオブジェクト
' 戻り値：結合列幅
'*********************************************************
Private Function getMergeCellWidth(aRa As Range) As Double
    
    Dim res As Double
    Dim indCell As Range
    For Each indCell In aRa.MergeArea.Cells
        res = res + indCell.ColumnWidth
    Next indCell
    getMergeCellWidth = res
    
End Function
