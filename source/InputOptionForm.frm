VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InputOptionForm 
   Caption         =   "���͑I�����ݒ�"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4665
   OleObjectBlob   =   "InputOptionForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "InputOptionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------------------------------
' ## �t�H�[��������
'------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    
    Dim configData As Variant
    
    ' �����n�I�u�W�F�N�g�ʒu�����ݒ�l�ǂݍ���
    configData = Split(CommitConfig.LoadConfig _
        (FormDisplay.ALIGNTEXT_CONFIG), vbCrLf)
    If UBound(configData) = 0 Then AlignTextBox.Value = configData(0)
    
    ' ���@�X�^�C���T�C�Y�����ݒ�l�ǂݍ���
    configData = Split(CommitConfig.LoadConfig _
        (FormDisplay.RESIZEDIMENSION_CONFIG), vbCrLf)
    If UBound(configData) = 0 Then ResizeDimensionBox.Value = configData(0)
    
    ' ���@�l�I�t�Z�b�g�ʒ����ݒ�l�ǂݍ���
    configData = Split(CommitConfig.LoadConfig _
        (FormDisplay.ADJUSTDIMENSION_CONFIG), vbCrLf)
    If UBound(configData) = 0 Then AdjustDimensionBox.Value = configData(0)
    
End Sub

'------------------------------------------------------------------------------
' ## �ݒ�l�ۑ�
'------------------------------------------------------------------------------
Private Sub InputOptionSaveButton_Click()
    
    Dim configData As Variant
    
    ' �ݒ�t�H���_�̏���
    Call CommitConfig.PrepareConfigFolder
    
    ' �����n�I�u�W�F�N�g�ʒu�����ݒ�l�ۑ�
    configData = AlignTextBox.Value
    Call CommitConfig.SaveConfig _
        (FormDisplay.ALIGNTEXT_CONFIG, configData)
    
    ' ���@�X�^�C���T�C�Y�����ݒ�l�ۑ�
    configData = ResizeDimensionBox.Value
    Call CommitConfig.SaveConfig _
        (FormDisplay.RESIZEDIMENSION_CONFIG, configData)
    
    ' ���@�l�I�t�Z�b�g�ʒ����ݒ�l�ۑ�
    configData = AdjustDimensionBox.Value
    Call CommitConfig.SaveConfig _
        (FormDisplay.ADJUSTDIMENSION_CONFIG, configData)
    
End Sub
