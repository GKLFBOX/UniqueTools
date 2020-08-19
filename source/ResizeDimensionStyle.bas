Attribute VB_Name = "ResizeDimensionStyle"
Option Explicit

'------------------------------------------------------------------------------
' ## 寸法スタイル適用オブジェクトのサイズ変更   2020/08/19 G.O.
'
' 寸法スタイル適用オブジェクトの全体の寸法尺度を変更する
'------------------------------------------------------------------------------
Public Sub ResizeDimensionStyle()
    
    On Error GoTo Error_Handler
    
    Dim configData As String
    Dim targetSelectionSet As ZcadSelectionSet
    Dim targetEntity As ZcadEntity
    Dim sizeFactor As Long
    
    ' 設定値読み込み
    configData = CommitConfig.LoadConfig(FormDisplay.RESIZEDIMENSION_CONFIG)
    
    ThisDrawing.Utility.Prompt _
        "サイズ変更する寸法スタイルオブジェクトを選択してください。" & vbCrLf
    
    ' 変更対象を範囲選択
    Set targetSelectionSet = ThisDrawing.SelectionSets.Add("NewSelectionSet")
    targetSelectionSet.SelectOnScreen
    
    If targetSelectionSet.Count = 0 Then
        Call CommonSub.ReleaseSelectionSet(targetSelectionSet)
        Exit Sub
    End If
    
    For Each targetEntity In targetSelectionSet
        If isDimensionStyle(targetEntity) Then
            targetEntity.Highlight True
        End If
    Next
    
    sizeFactor = ThisDrawing.Utility.GetInteger _
        ("変更尺度を入力 または [" & configData & "]:")
    
    For Each targetEntity In targetSelectionSet
        If isDimensionStyle(targetEntity) Then
            targetEntity.ScaleFactor = sizeFactor
        End If
    Next
    
    Call CommonSub.ReleaseSelectionSet(targetSelectionSet)
    
    Exit Sub
    
Error_Handler:
    For Each targetEntity In targetSelectionSet
        Call CommonSub.ResetHighlight(targetEntity)
    Next
    Call CommonSub.ReleaseSelectionSet(targetSelectionSet)
    ThisDrawing.Utility.Prompt "なんらかのエラーです。" & vbCrLf
    
End Sub

'------------------------------------------------------------------------------
' ## 寸法スタイルオブジェクト判定
'------------------------------------------------------------------------------
Private Function isDimensionStyle _
    (ByVal target_entity As ZcadEntity) As Boolean
    
    If TypeOf target_entity Is ZcadDim3PointAngular _
    Or TypeOf target_entity Is ZcadDimAligned _
    Or TypeOf target_entity Is ZcadDimAngular _
    Or TypeOf target_entity Is ZcadDimArcLength _
    Or TypeOf target_entity Is ZcadDimRotated _
    Or TypeOf target_entity Is ZcadDimDiametric _
    Or TypeOf target_entity Is ZcadDimRadial _
    Or TypeOf target_entity Is ZcadLeader Then
        isDimensionStyle = True
    Else
        isDimensionStyle = False
    End If
    
End Function
