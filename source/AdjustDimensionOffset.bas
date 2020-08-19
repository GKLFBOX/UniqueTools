Attribute VB_Name = "AdjustDimensionOffset"
Option Explicit

'------------------------------------------------------------------------------
' ## 寸法線のオフセット量を変更する   2020/08/20 G.O.
'
' 寸法線の文字オフセット量(寸法線と文字の離れ量)を変更する
'------------------------------------------------------------------------------
Public Sub AdjustDimensionOffset()
    
    On Error GoTo Error_Handler
    
    Dim configData As String
    Dim targetSelectionSet As ZcadSelectionSet
    Dim targetEntity As ZcadEntity
    Dim offsetAmount As Single
    
    ' 設定値読み込み
    configData = CommitConfig.LoadConfig(FormDisplay.ADJUSTDIMENSION_CONFIG)
    
    ThisDrawing.Utility.Prompt _
        "寸法値オフセット量を変更する寸法線を選択してください。" & vbCrLf
    
    ' 変更対象を範囲選択
    Set targetSelectionSet = ThisDrawing.SelectionSets.Add("NewSelectionSet")
    targetSelectionSet.SelectOnScreen
    
    If targetSelectionSet.Count = 0 Then
        Call CommonSub.ReleaseSelectionSet(targetSelectionSet)
        Exit Sub
    End If
    
    For Each targetEntity In targetSelectionSet
        If isDimensionLine(targetEntity) Then
            targetEntity.Highlight True
        End If
    Next
    
    offsetAmount = ThisDrawing.Utility.GetInteger _
        ("変更オフセット量(x/10)を入力 または [" & configData & "]:")
    offsetAmount = offsetAmount * 0.1
    
    For Each targetEntity In targetSelectionSet
        If isDimensionLine(targetEntity) Then
            targetEntity.TextGap = offsetAmount
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
' ## 寸法線判定
'------------------------------------------------------------------------------
Private Function isDimensionLine _
    (ByVal target_entity As ZcadEntity) As Boolean
    
    If TypeOf target_entity Is ZcadDim3PointAngular _
    Or TypeOf target_entity Is ZcadDimAligned _
    Or TypeOf target_entity Is ZcadDimAngular _
    Or TypeOf target_entity Is ZcadDimArcLength _
    Or TypeOf target_entity Is ZcadDimRotated _
    Or TypeOf target_entity Is ZcadDimDiametric _
    Or TypeOf target_entity Is ZcadDimRadial Then
        isDimensionLine = True
    Else
        isDimensionLine = False
    End If
    
End Function

