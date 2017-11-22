VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Overzicht_werkposten 
   Caption         =   "Optelling werkposten"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12000
   OleObjectBlob   =   "Overzicht_werkposten.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Overzicht_werkposten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Const ImgFileName As String = "C:ATest.bmp"

Private Sub CommandButton1_Click()
werkposten_sommeren
End Sub

Private Sub UserForm_Initialize()

Call werkposten_sommeren

End Sub

Public Sub setCommandButtonFaceId(faceId As Integer, _
 ByRef control As CommandButton)
 
    Application.CommandBars.FindControl(Type:=msoControlButton, ID:=faceId).CopyFace
    
    Dim ImgFileName As String
    ImgFileName = Environ("temp") & "\tmppicture.bmp"
     
    '// Save the img
    'SavePicture PastePicture(xlBitmap), ImgFileName

    '// Load the Img
    control.Picture = LoadPicture(ImgFileName)
    control.PicturePosition = 0
    
    '// Cleanup
    Kill ImgFileName
   
End Sub




