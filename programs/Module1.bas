Attribute VB_Name = "Module1"
Option Explicit
' *** パブリック変数設定 ***
Private WS_ERR As Worksheet
Private Const ERR_SHEET = "エラーリスト"
Private ERR_OptRow As Long
'メッセージ出力セルアドレス
Private Const OptMsgAddr As String = "A2"
'ファイルパス記載セルアドレス
Private Const IptPathAddr As String = "B4"
Private Const OptPathAddr As String = "B5"
'見出し行
Private Const IptStRow As Integer = 1
'削除用文字列
Private Const CstDelStr As String = "★削除行"
'言語判定用
Private Const CstJapanese As Integer = 0
Private Const CstEnglish As Integer = 1
'初期値、エラー値
Private Const CstInitVal As Integer = 0
Private Const CstErr As Integer = -999
'定数
Private CstStrSelBox As String '「セレクトボックス」
' *** ユーザー定義Type START ***
'列番号と列名を格納
Private Type ClmHead
    No As Integer
    Nm As String
End Type
'列名とClmHeadの要素番号を格納
Private Type CstClm
    Nm As String
    Idx As Integer
End Type
' *** ユーザー定義Type END ***
'試験名、シート名、エイリアスネームを格納
Private Head_T(2) As String
Private Const CstTestNmIdx As Integer = 0
Private Const CstSheetNmIdx As Integer = 1
Private Const CstAliasNmIdx As Integer = 2
Private Const CstHeadClm As Integer = 1
'列情報を格納　列名をキーにしてCstClmと結合
Private Clm_T() As ClmHead
'列名コンスタント格納用変数
Private CstPg As CstClm         ' "Page"
Private CstFldID As CstClm      ' "フィールドID"
Private CstType As CstClm       ' "種類"
Private CstDef As CstClm        ' "デフォルト値"
Private CstSel As CstClm        ' "選択肢"
Private CstValid As CstClm      ' "バリデーション"
Private CstPresence As CstClm   ' "必須チェック"
Private CstNum As CstClm        ' "数値チェック"
Private CstLogical As CstClm    ' "論理式チェック"
Private CstDay As CstClm        ' "日付チェック"
Private CstMst As CstClm        ' "保存先のマスタ/参照先"
Private CstCmt As CstClm        ' "コメント"
Private CstDeviation As CstClm  ' "逸脱判定"
Private CstPass As CstClm       ' "Pass/Fail"
Private CstDel As CstClm        ' "削除有無"
Private CstSeq As CstClm        ' "SEQ"
Private CstCDISC As CstClm      ' CDISC情報
Private CstRegExp As CstClm     ' 正規表現チェック
Private CstWCnt As CstClm       ' 文字数チェック

Private Sub sSelectFile()
'変換元ファイル名のパスを取得
Dim ObjWS       As Worksheet
Dim OpenFileName As String

    Set ObjWS = ActiveSheet

    Call sDispMsg(ObjWS, "")

    OpenFileName = Application.GetOpenFileName()
    If OpenFileName <> "False" Then
        ObjWS.Range(IptPathAddr).Value = OpenFileName
        Call sDispMsg(ObjWS, "入力元に" & ObjWS.Range(IptPathAddr).Value & "が指定されました")
    End If
    
    Set ObjWS = Nothing

End Sub

Private Sub sSelectFolder()
'変換先フォルダ名のパスを取得
Dim StrFldNm As String
Dim ObjWS       As Worksheet

    Set ObjWS = ActiveSheet

    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then
            StrFldNm = .SelectedItems(1)
        End If
    End With
    If StrFldNm <> "" Then
        ObjWS.Range(OptPathAddr).Value = StrFldNm
        Call sDispMsg(ObjWS, "出力先に" & ObjWS.Range(OptPathAddr).Value & "が指定されました")
    End If
        
End Sub

Private Sub sDispMsg(WS As Worksheet, strMsg As String)
'メッセージ出力
    WS.Activate
    WS.Range(OptMsgAddr).Value = strMsg
End Sub

Private Sub Convert()
'On Error GoTo ErrTrap
Dim WsThis  As Worksheet    'メインシート
Dim WbIn    As Workbook     '入力ファイル
Dim StrFileName As String
'出力ファイル
Dim OptWB As Workbook
Dim OptWS As Worksheet
'出力先パス
Dim OptPath As String
Dim SplPath As Variant
Dim OptFileName As String
Dim i As Integer
Dim Err_F As Boolean

    Set WsThis = ActiveSheet
    
    Call sDispMsg(WsThis, "")
    
    StrFileName = WsThis.Range(IptPathAddr).Value
    If StrFileName = "" Then
        MsgBox "ファイルを指定してください"
        Exit Sub
    End If
    If Dir(StrFileName) = "" Then
        MsgBox "ファイルが見つかりません"
        Exit Sub
    End If
    If FileLen(StrFileName) = 0 Then
        MsgBox "ファイルの中身が空です"
        Exit Sub
    End If
    '拡張子がxlsxのみ対象とする
    If Not StrConv(Right(StrFileName, 5), vbLowerCase) = ".xlsx" Then
        MsgBox "拡張子.xlsxのファイルのみ対象とします"
        Exit Sub
    End If
    
    '出力先を取得
    OptPath = Trim(WsThis.Range(OptPathAddr).Value)
    If OptPath = "" Then
        SplPath = Split(StrFileName, "