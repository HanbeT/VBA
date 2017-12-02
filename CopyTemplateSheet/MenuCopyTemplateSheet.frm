VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MenuCopyTemplateSheet 
   Caption         =   "連番付シート複製くん Ver1.0.0"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8130
   OleObjectBlob   =   "MenuCopyTemplateSheet.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "MenuCopyTemplateSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub UserForm_Initialize()
    Dim sh As Worksheet
    For Each sh In Worksheets
        With ComboBox1
            .AddItem sh.Name
        End With
    Next
End Sub

Private Sub CommandButton1_Click()
    Dim TargetSheet As String
    Dim startNum As Integer
    Dim endNum As Integer
    Dim ZeroNum As Integer
    Dim FirstName As String
    Dim SecondName As String
    Dim errorFlg As Boolean
    errorFlg = True
    
    Label8.Caption = ""
    Label9.Caption = ""
    Label10.Caption = ""
    Label11.Caption = ""
    
    
    If ComboBox1.Value <> "" Then
        TargetSheet = ComboBox1.Value
    Else
        errorFlg = False
        Label8.Caption = "※シート名が未選択です。"
    End If
    
    If TextBox1.Value = "" Then
        errorFlg = False
        Label9.Caption = "※開始番号が未入力です。"
    ElseIf IsNumeric(TextBox1.Value) Then
        startNum = Val(StrConv(TextBox1.Value, vbNarrow))
    Else
        errorFlg = False
        Label9.Caption = "※開始番号が数字ではありません。"
    End If
        
    If TextBox2.Value = "" Then
        errorFlg = False
        Label10.Caption = "※開始番号が未入力です。"
    ElseIf IsNumeric(TextBox2.Value) Then
        endNum = Val(StrConv(TextBox2.Value, vbNarrow))
    Else
        errorFlg = False
        Label10.Caption = "※開始番号が数字ではありません。"
    End If
    
    If CheckBox1.Value = False Then
        ZeroNum = 1
    ElseIf CheckBox1.Value = True And TextBox5.Value = "" Then
        errorFlg = False
        Label11.Caption = "※桁数が未入力です。"
    ElseIf CheckBox1.Value = True And IsNumeric(TextBox5.Value) Then
        ZeroNum = Val(StrConv(TextBox5.Value, vbNarrow))
    Else
        errorFlg = False
        Label11.Caption = "※桁数が数字ではありません。"
    End If
    
    FirstName = TextBox3.Value
    SecondName = TextBox4.Value
    
    If errorFlg Then
        Unload Me
        Call mainLogic(TargetSheet, startNum, endNum, ZeroNum, FirstName, SecondName)
    End If
    
End Sub

Private Sub CommandButton2_Click()
    End
End Sub
