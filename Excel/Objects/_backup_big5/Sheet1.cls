VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
Dim theID As String     '____代號
Dim theInfo As baseInfo
If Target.Columns.Count = Target.EntireRow.Columns.Count Then Exit Sub     '_______C則退出
If Target.Cells(1).Column > 1 Then Exit Sub
If Target.Cells(1).Row > 2 Then Exit Sub
Application.EnableEvents = False        '_________ぁ顗漱瑑o
[b14] = "__7_____..."y待..."
Application.ScreenUpdating = False      '________螢幕更新
Application.ScreenUpdating = True       '______幕更新
With Target.Cells(1)     '_______騔x存格
    theInfo = findAll(.Text)
    If theInfo.theCode = "" Then
    Application.StatusBar = "______..."中..."
    Call Stocklist
    Application.StatusBar = False
    theInfo = findAll(.Text)
    End If
    If theInfo.theCode = "" Then
        '[a2].ClearContents        '______A2____A2所有內容
        '[f1].ClearContents
        '[g1].ClearContents
       ' MsgBox "_______________"票代號或名稱！"
       ' [d14] = "":[a2].Select
    Else
        With .EntireRow
            [a2] = theInfo.theCode
            [a1] = theInfo.theName & " (" & theInfo.theCode & ")" & "_" & theIndustryy
            [a20] = theInfo.theName
        End With
    End If
End With
Application.ScreenUpdating = False      '________螢幕更新
'If theInfo.theCode <> "" Then Call Macro1
Call Macro1

Application.EnableEvents = True         '_______顗漱瑑o
Application.ScreenUpdating = True       '______幕更新

Call ProtectSheet(ActiveSheet)
Beep

End Sub

