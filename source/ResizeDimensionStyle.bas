Attribute VB_Name = "ResizeDimensionStyle"
Option Explicit

'------------------------------------------------------------------------------
' ## ���@�X�^�C���K�p�I�u�W�F�N�g�̃T�C�Y�ύX   2020/08/19 G.O.
'
' ���@�X�^�C���K�p�I�u�W�F�N�g�̑S�̂̐��@�ړx��ύX����
'------------------------------------------------------------------------------
Public Sub ResizeDimensionStyle()
    
    On Error GoTo Error_Handler
    
    Dim configData As String
    Dim targetSelectionSet As ZcadSelectionSet
    Dim targetEntity As ZcadEntity
    Dim sizeFactor As Long
    
    ' �ݒ�l�ǂݍ���
    configData = CommitConfig.LoadConfig(FormDisplay.RESIZEDIMENSION_CONFIG)
    
    ThisDrawing.Utility.Prompt _
        "�T�C�Y�ύX���鐡�@�X�^�C���I�u�W�F�N�g��I�����Ă��������B" & vbCrLf
    
    ' �ύX�Ώۂ�͈͑I��
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
        ("�ύX�ړx����� �܂��� [" & configData & "]:")
    
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
    ThisDrawing.Utility.Prompt "�Ȃ�炩�̃G���[�ł��B" & vbCrLf
    
End Sub

'------------------------------------------------------------------------------
' ## ���@�X�^�C���I�u�W�F�N�g����
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
