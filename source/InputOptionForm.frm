VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InputOptionForm 
   Caption         =   "入力選択肢設定"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4665
   OleObjectBlob   =   "InputOptionForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "InputOptionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------------------------------
' ## フォーム初期化
'------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    
    Dim configData As Variant
    
    ' 文字系オブジェクト位置調整設定値読み込み
    configData = Split(CommitConfig.LoadConfig _
        (FormDisplay.ALIGNTEXT_CONFIG), vbCrLf)
    If UBound(configData) = 0 Then AlignTextBox.Value = configData(0)
    
    ' 寸法スタイルサイズ調整設定値読み込み
    configData = Split(CommitConfig.LoadConfig _
        (FormDisplay.RESIZEDIMENSION_CONFIG), vbCrLf)
    If UBound(configData) = 0 Then ResizeDimensionBox.Value = configData(0)
    
    ' 寸法値オフセット量調整設定値読み込み
    configData = Split(CommitConfig.LoadConfig _
        (FormDisplay.ADJUSTDIMENSION_CONFIG), vbCrLf)
    If UBound(configData) = 0 Then AdjustDimensionBox.Value = configData(0)
    
End Sub

'------------------------------------------------------------------------------
' ## 設定値保存
'------------------------------------------------------------------------------
Private Sub InputOptionSaveButton_Click()
    
    Dim configData As Variant
    
    ' 設定フォルダの準備
    Call CommitConfig.PrepareConfigFolder
    
    ' 文字系オブジェクト位置調整設定値保存
    configData = AlignTextBox.Value
    Call CommitConfig.SaveConfig _
        (FormDisplay.ALIGNTEXT_CONFIG, configData)
    
    ' 寸法スタイルサイズ調整設定値保存
    configData = ResizeDimensionBox.Value
    Call CommitConfig.SaveConfig _
        (FormDisplay.RESIZEDIMENSION_CONFIG, configData)
    
    ' 寸法値オフセット量調整設定値保存
    configData = AdjustDimensionBox.Value
    Call CommitConfig.SaveConfig _
        (FormDisplay.ADJUSTDIMENSION_CONFIG, configData)
    
End Sub
