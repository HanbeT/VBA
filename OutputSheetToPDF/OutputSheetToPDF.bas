Attribute VB_Name = "OutputSheetToPDF"
Option Explicit
Option Private Module

Public Sub MainOutputSheetToPdf()
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Dim logic As LogicOutputSheetToPDF
    
    Set logic = New LogicOutputSheetToPDF
    Call logic.mainLogic
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
End Sub






