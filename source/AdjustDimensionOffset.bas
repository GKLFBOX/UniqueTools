Attribute VB_Name = "AdjustDimensionOffset"
Option Explicit

'------------------------------------------------------------------------------
' ## ���@���̃I�t�Z�b�g�ʂ�ύX����   2020/08/20 G.O.
'
' ���@���̕����I�t�Z�b�g��(���@���ƕ����̗����)��ύX����
'------------------------------------------------------------------------------
Public Sub AdjustDimensionOffset()
    
    On Error GoTo Error_Handler
    
    Dim configData As String
    Dim targetSelectionSet As ZcadSelectionSet
    Dim targetEntity As ZcadEntity
    Dim offsetAmount As Single
    
    ' �ݒ�l�ǂݍ���
    configData = CommitConfig.LoadConfig(FormDisplay.ADJUSTDIMENSION_CONFIG)
    
    ThisDrawing.Utility.Prompt _
        "���@�l�I�t�Z�b�g�ʂ�ύX���鐡�@����I�����Ă��������B" & vbCrLf
    
    ' �ύX�Ώۂ�͈͑I��
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
        ("�ύX�I�t�Z�b�g��(x/10)����� �܂��� [" & configData & "]:")
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
    ThisDrawing.Utility.Prompt "�Ȃ�炩�̃G���[�ł��B" & vbCrLf
    
End Sub

'------------------------------------------------------------------------------
' ## ���@������
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

