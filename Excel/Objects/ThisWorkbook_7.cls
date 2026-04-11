VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
  On Error Resume Next
  Application.ErrorCheckingOptions.BackgroundChecking = False
  Application.CommandBars("External Data").Visible = False
  Application.AutoRecover.Enabled = False
  ActiveWorkbook.EnableAutoRecover = False
  Application.AskToUpdateLinks = False
  Delete_Connection_XL2003
  ThisWorkbook.QueryTablesDelete "all"
  
  Application.MaxChange = 0.001
    With ActiveWorkbook
        .UpdateRemoteReferences = False
        .PrecisionAsDisplayed = False
        .SaveLinkValues = False
    End With

  With Application.AutoCorrect
        .TwoInitialCapitals = False
        .CorrectSentenceCap = False
        .CapitalizeNamesOfDays = False
        .CorrectCapsLock = False
        .ReplaceText = False
        .DisplayAutoCorrectOptions = False
  End With

End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
  ThisWorkbook.QueryTablesDelete "all"
End Sub
'_________埭ㄗ捄{式


Public Function QueryTablesDelete(ash As String)
  Dim ws As Worksheet
  Dim qt As QueryTable
  If ash = "all" Then
    For Each ws In ThisWorkbook.Worksheets
      For Each qt In ws.QueryTables
        qt.Delete
      Next
    Next
  Else
    For Each qt In Worksheets(ash).QueryTables
      qt.Delete
    Next
  End If
End Function
 '_________埭ㄗ捄{式
