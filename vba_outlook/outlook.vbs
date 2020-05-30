Item.SendUsingAccount
↑上記の中にアカウントが入っている

http://tech.surviveplus.net/archives/136

Public WithEvents myInboxFolder As Folder
Public WithEvents myInboxFolder2 As Folder

Private Sub Application_Startup()

Set myInboxFolder = Me.Application.Session.GetDefaultFolder(olFolderInbox) '既定の受信トレイ
Set myInboxFolder2 =                                           Me.Application.Session.Folders("eggist77@kem.biglobe.ne.jp").Folders("受信トレイ")

End Sub

' 受信トレイからアイテムが移動される前に実行される処理
'
' 引数 Item - イベントが発生したオブジェクトです。
' 引数 MoveTo - 異動先のフォルダを表すオブジェクトです。
' 引数(参照) Cancel - True を設定すると、アイテムの移動を中止します。規定値は False です。
'
Private Sub myInboxFolder_BeforeItemMove(ByVal Item As Object, ByVal MoveTo As MAPIFolder, Cancel As Boolean)

    Select Case MsgBox("移動してもよいですか？", vbInformation Or vbYesNo)
        Case vbYes
            ' 移動実施

        Case vbNo
            ' 移動中止
            Cancel = True
            Exit Sub
    End Select
End Sub

Private Sub myInboxFolder2_BeforeItemMove(ByVal Item As Object, ByVal MoveTo As MAPIFolder, Cancel As Boolean)

    Select Case MsgBox("移動してもよいですか？", vbInformation Or vbYesNo)
        Case vbYes
            ' 移動実施

        Case vbNo
            ' 移動中止
            Cancel = True
            Exit Sub
    End Select
End Sub