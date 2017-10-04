  Dim WSP As Workspace
 'トランザクション
    On Error GoTo Error
'open処理をする場合以下のコードも必要'
'Set dbX = WSP.OpenDatabase("e:\books\acc2007vba\myDB.accdb")

    Set WSP = DBEngine.Workspaces(0)
    WSP.BeginTrans
'ここに処理'
    WSP.CommitTrans
    dbs.Close: Set dbs = Nothing
    Exit Sub
Error:        '--error発生時にはこちらへ
    WSP.Rollback    'ロールバック処理、
    dbs.Close: Set dbs = Nothing'--閉じる＜解放
    MsgBox Error$'--error内容'