Attribute VB_Name = "modCAT"
Option Explicit

Public Points As Long
Public FinalPoints As Long
Public Dog As DogRec
Public Inventory(1 To 8) As InvRec
Public Item(1 To 9) As ItemRec
Public ShopItem(1 To 9) As ShopRec
Public Globals As GlobalRec

Private Type GlobalRec
    AmountBought(1 To 9) As Long
End Type

Private Type DogRec
    DogName As String
    Hunger As Byte
    Thirst As Byte
    Energy As Byte
    InAction As Boolean
    Frame As Byte
    'sickness
    Sick As Boolean
    SicknessImmunity As Integer
End Type

Private Type InvRec
    ItemType As Byte
    ActionValue As Byte
    Picture As Byte
    Sickness As Integer
End Type

Private Type ItemRec
    ItemType As Byte
    ActionValue As Byte
    Picture As Byte
    Sickness As Integer
End Type

Private Type ShopRec
    Cost As Integer
End Type
