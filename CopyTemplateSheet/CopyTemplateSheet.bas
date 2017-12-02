Attribute VB_Name = "CopyTemplateSheet"

Option Explicit
Option Private Module

Public Sub mainCopyTemplateSheet()
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    MenuCopyTemplateSheet.Show
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub

Public Function mainLogic(tSheet As String, sNum As Integer, enNum As Integer, zNum As Integer, fName As String, sName As String)
    
    Dim stepNum As Integer
    Dim padding As String
    Dim i As Integer
    Dim delFlg As Boolean
    delFlg = False
    
    padding = zeroPadding(zNum)
    
    If sNum < enNum Then
        ' インクリメントパターン
        stepNum = 1
    Else
        ' デクリメントパターン
        stepNum = -1
    End If
    
    For i = sNum To enNum Step stepNum
        If existsSheet(fName & Format(i, padding) & sName) Then
            If delFlg = False And MsgBox("同名シートが存在します。削除してもよろしいですか。" & vbCrLf & "「" & fName & Format(i, padding) & sName & "」", vbYesNo, "削除確認") = vbYes Then
                Worksheets(fName & Format(i, padding) & sName).Delete
                If MsgBox("以降、同名シートが存在する場合、" & vbCrLf & "未確認で削除してもよろしいでしょうか。", vbYesNo, "") = vbYes Then
                    delFlg = True
                End If
            ElseIf delFlg = True Then
                Worksheets(fName & Format(i, padding) & sName).Delete
            Else
                GoTo Continue1
            End If
        End If
        Worksheets(tSheet).Copy After:=Worksheets(Worksheets.Count)
        ActiveSheet.Name = fName & Format(i, padding) & sName
Continue1:
    Next i

End Function

Private Function existsSheet(aSName As String) As Boolean
    Dim i As Integer
    Dim res As Boolean
    res = False
    For i = 1 To Worksheets.Count
        If Worksheets(i).Name = aSName Then
            res = True
            Exit For
        End If
    Next i
    existsSheet = res
End Function

Private Function zeroPadding(aDigits As Integer)
    Dim i As Integer
    Dim res As String
    res = ""
    For i = 1 To aDigits
        res = res & "0"
    Next i
    zeroPadding = res
End Function

