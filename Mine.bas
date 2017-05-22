Attribute VB_Name = "Mine"
Sub Mine()

Dim nonce As Range
Dim hash As Range
Dim lastnonce As Range
Dim i As Integer


Set ThisRange = Selection
Set lastnonce = Application.InputBox("Select Cell with previous Nonce", "Target", ThisRange.Address, Type:=8)
Set nonce = Application.InputBox("Select Cell with new Nonce", "Target", ThisRange.Address, Type:=8)
Set hash = Application.InputBox("Select Cell with new Hash", "Target", ThisRange.Address, Type:=8)

i = lastnonce

Do Until x = "00"
    i = i + 1
    nonce.Cells.Value = i
    x = Left(hash.Value, 2)
Loop

MsgBox "You've just mined a block!"

End Sub


