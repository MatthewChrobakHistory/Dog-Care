Attribute VB_Name = "modData"
Public Sub LoadItems()
Dim i As Long
'1=food
'2=bed
'3=drink

For i = 1 To 9
    With Item(i)
        Select Case i
            Case 1
                .ItemType = 1
                .ActionValue = 10
                .Picture = 1
            Case 2
                .ItemType = 1
                .ActionValue = 25
                .Picture = 2
            Case 3
                .ItemType = 1
                .ActionValue = 40
                .Picture = 3
            Case 4
                .ItemType = 2
                .ActionValue = 10
                .Picture = 4
            Case 5
                .ItemType = 2
                .ActionValue = 25
                .Picture = 5
            Case 6
                .ItemType = 2
                .ActionValue = 40
                .Picture = 6
            Case 7
                .ItemType = 3
                .ActionValue = 10
                .Picture = 7
            Case 8
                .ItemType = 3
                .ActionValue = 25
                .Picture = 8
            Case 9
                .ItemType = 3
                .ActionValue = 40
                .Picture = 9
        End Select
    End With
Next

End Sub

Public Sub LoadGame()
Dim i As Byte

Call LoadItems

With frmMain
    .picMap.Picture = LoadPicture(App.Path & "\graphics\map\norm.bmp")
        
    '10 = blank
    '0 = noitem
    For i = 1 To 8
        .imgInventoryItem(i).Picture = LoadPicture(App.Path & "\graphics\items\10.bmp")
        Inventory(i).ItemType = 0
    Next
    
    .PBEnergy = 100
    .PBHunger = 100
    .PBThirst = 100
    Points = 0
    FinalPoints = 0
    frmMain.lblPoints = "Points: " & Points
    
    Call SaveDog

End With

For i = 1 To 9
    With ShopItem(i)
        Select Case i
            Case 1
                .Cost = 100
            Case 2
                .Cost = 400
            Case 3
                .Cost = 850
            Case 4
                .Cost = 1000
            Case 5
                .Cost = 2000
            Case 6
                .Cost = 3000
            Case 7
                .Cost = 100
            Case 8
                .Cost = 400
            Case 9
                .Cost = 850
        End Select
    End With
Next
    
End Sub
