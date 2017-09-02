Attribute VB_Name = "modLogic"
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Sub GameLoop()
Dim Tick As Long
Dim tmr2500 As Long
Dim Running As Boolean

Running = True

Do While Running
    Tick = GetTickCount
    
    If Tick > tmr2500 Then
        With frmMain
            If Dog.Hunger - 3 < 0 Then Call GameOver
            If Dog.Thirst - 4 < 0 Then Call GameOver
            If Dog.Energy - 2 < 0 Then Call GameOver
        
            .PBHunger.Value = Dog.Hunger - 3
            .PBThirst.Value = Dog.Thirst - 4
            .PBEnergy.Value = Dog.Energy - 2
            
            Call SaveDog
        End With
        
        If Dog.Frame = 2 Then
            frmMain.picMap.Picture = LoadPicture(App.Path & "\graphics\map\norm.bmp")
            Dog.InAction = False
            Dog.Frame = 0
        Else
            Dog.Frame = Dog.Frame + 1
        End If
        
        Points = Points + (Dog.Hunger / 3 + Dog.Thirst / 3 + Dog.Energy / 3)
        FinalPoints = FinalPoints + (Dog.Hunger / 3 + Dog.Thirst / 3 + Dog.Energy / 3)
        frmMain.lblPoints.Caption = "Points: " & Points
        If frmShop.Visible = True Then frmShop.lblPoints.Caption = frmMain.lblPoints.Caption
        
        tmr2500 = GetTickCount + 2500
    End If
    
    DoEvents
Loop

End Sub

Public Sub GameOver()

MsgBox "You lose. Your score was " & FinalPoints
End

End Sub

Public Sub UseItem(Index As Integer)
Dim ItemValue As Byte

ItemValue = Inventory(Index).ActionValue
'1=food
'2=bed
'3=drink
Select Case Inventory(Index).ItemType
    Case 1
        If frmMain.PBHunger.Value + ItemValue > 100 Then
            frmMain.PBHunger.Value = 100
        Else
            frmMain.PBHunger.Value = frmMain.PBHunger.Value + ItemValue
        End If
        frmMain.picMap.Picture = LoadPicture(App.Path & "\graphics\map\eating.bmp")
    Case 2
        If frmMain.PBEnergy.Value + ItemValue > 100 Then
            frmMain.PBEnergy.Value = 100
        Else
            frmMain.PBEnergy.Value = frmMain.PBEnergy.Value + ItemValue
        End If
        
        If Inventory(Index).ActionValue = 10 Then frmMain.picMap.Picture = LoadPicture(App.Path & "\graphics\map\sleepingGreen.bmp")
        If Inventory(Index).ActionValue = 25 Then frmMain.picMap.Picture = LoadPicture(App.Path & "\graphics\map\sleepingYellow.bmp")
        If Inventory(Index).ActionValue = 40 Then frmMain.picMap.Picture = LoadPicture(App.Path & "\graphics\map\sleepingPurple.bmp")
    Case 3
        If frmMain.PBThirst.Value + ItemValue > 100 Then
            frmMain.PBThirst.Value = 100
        Else
            frmMain.PBThirst.Value = frmMain.PBThirst.Value + ItemValue
        End If
        frmMain.picMap.Picture = LoadPicture(App.Path & "\graphics\map\drinking.bmp")
End Select

Call SaveDog
Dog.Frame = 1
Dog.InAction = True

If Inventory(Index).ItemType = 2 Then Exit Sub

Call RemoveItem(Index)

End Sub

Public Sub RemoveItem(ByVal Index As Integer)

With Inventory(Index)
    .ItemType = 0
    .ActionValue = 0
End With

frmMain.imgInventoryItem(Index).Picture = LoadPicture(App.Path & "\graphics\items\10.bmp")

End Sub

Public Sub SaveDog()

With Dog
    .Energy = frmMain.PBEnergy.Value
    .Hunger = frmMain.PBHunger.Value
    .Thirst = frmMain.PBThirst.Value

    If .Energy > 75 And .Hunger > 75 And .Thirst > 75 Then frmMain.picFace.Picture = LoadPicture(App.Path & "\graphics\faces\happy.bmp")
    If .Energy > 50 And .Energy < 75 Or .Hunger > 50 And .Hunger < 75 Or .Thirst > 50 And .Thirst < 75 Then frmMain.picFace.Picture = LoadPicture(App.Path & "\graphics\faces\norm.bmp")
    If .Energy < 50 Or .Hunger < 50 Or .Thirst < 50 Then frmMain.picFace.Picture = LoadPicture(App.Path & "\graphics\faces\sad.bmp")
End With

End Sub

Public Sub BuyItem(ByVal Index As Byte, ByVal price As Integer)
Dim i As Byte
Dim freeslot As Byte

freeslot = 0

If price > Points Then
    MsgBox "You don't have enough points to buy that!"
    Exit Sub
End If

For i = 1 To 8
    'find free space
    If Inventory(i).ItemType = 0 Then
        freeslot = i
    End If
Next

If freeslot = 0 Then Exit Sub
Inventory(freeslot).ActionValue = Item(Index).ActionValue
frmMain.imgInventoryItem(freeslot).Picture = LoadPicture(App.Path & "\graphics\items\" & Item(Index).Picture & ".bmp")
Inventory(freeslot).ItemType = Item(Index).ItemType

Call Economy(Index)

Points = Points - price
frmShop.lblPoints.Caption = "Points: " & Points
frmMain.lblPoints.Caption = "Points: " & Points

End Sub

Public Sub SellItem(ByVal Index As Byte, ByVal ActionValue As Integer, ByVal ItemType As Byte)
Dim i As Long
Dim RPoints As Integer

For i = 1 To 9
    If Item(i).ActionValue = ActionValue And Item(i).ItemType = ItemType Then
        RPoints = ShopItem(i).Cost / 2
        Points = Points + RPoints
        frmMain.lblPoints = "Points: " & Points
        If frmShop.Visible = True Then frmShop.lblPoints.Caption = "Points: " & Points
        RemoveItem (Index)
    End If
Next

End Sub

Public Sub Economy(ByVal Index As Byte)
Dim x As Byte
Dim y As Byte

x = 0
y = 0

Select Case Index
    Case Index = 1 To 3
        If Index = 1 Then
            x = 2
            y = 3
        End If
        If Index = 2 Then
            x = 3
            y = 1
        End If
        If Index = 3 Then
            x = 1
            y = 2
        End If
    Case Index = 4 To 6
        If Index = 4 Then
            x = 5
            y = 6
        End If
        If Index = 5 Then
            x = 6
            y = 4
        End If
        If Index = 6 Then
            x = 4
            y = 5
        End If
    Case Index = 7 To 9
        If Index = 7 Then
            x = 8
            y = 9
        End If
        If Index = 8 Then
            x = 9
            y = 7
        End If
        If Index = 9 Then
            x = 7
            y = 8
        End If
End Select

If x + y = 0 Then Exit Sub

If ShopItem(x).Cost - (ShopItem(x).Cost * 0.1) > 1 And ShopItem(y).Cost - (ShopItem(y).Cost * 0.1) > 1 Then
    ShopItem(x).Cost = ShopItem(x).Cost - (ShopItem(x).Cost * 0.1)
    ShopItem(y).Cost = ShopItem(y).Cost - (ShopItem(y).Cost * 0.1)
    ShopItem(Index).Cost = ShopItem(Index).Cost + (ShopItem(Index).Cost * 0.1)
End If

Call frmShop.ShowShop
    
End Sub
