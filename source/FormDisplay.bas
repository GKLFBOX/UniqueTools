Attribute VB_Name = "FormDisplay"
Option Explicit

'------------------------------------------------------------------------------
' ## 文字装飾線設定ファイルのファイル名
'------------------------------------------------------------------------------
Public Const REFERENCELINE_CONFIG As String = "\ReferenceLine.config"
Public Const STRIKETHROUGH_CONFIG As String = "\Strikethrough.config"

'------------------------------------------------------------------------------
' ## 入力選択肢設定ファイルのファイル名
'------------------------------------------------------------------------------
Public Const ALIGNTEXT_CONFIG As String = "\AlignText.config"
Public Const RESIZEDIMENSION_CONFIG As String = "\ResizeDimension.config"
Public Const ADJUSTDIMENSION_CONFIG As String = "\AdjustDimension.config"

'------------------------------------------------------------------------------
' ## レイアウト編集フォーム表示
'------------------------------------------------------------------------------
Public Sub DisplayLayoutForm()
    
    ' モードレス表示はフォーカスが取れないため使用していない
    Load LayoutForm
    LayoutForm.Show
    
End Sub

'------------------------------------------------------------------------------
' ## 文字装飾設定フォーム表示
'------------------------------------------------------------------------------
Public Sub DisplayDecorationLineForm()
    
    Load DecorationLineForm
    DecorationLineForm.Show
    
End Sub

'------------------------------------------------------------------------------
' ## 用紙枠リストcsv出力フォーム表示
'------------------------------------------------------------------------------
Public Sub DisplayFrameListForm()
    
    Load FrameListForm
    FrameListForm.Show
    
End Sub

'------------------------------------------------------------------------------
' ## 入力選択肢設定フォーム表示
'------------------------------------------------------------------------------
Public Sub DisplayInputOptionForm()
    
    Load InputOptionForm
    InputOptionForm.Show
    
End Sub
