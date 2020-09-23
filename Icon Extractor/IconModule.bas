Attribute VB_Name = "IconModule"
Private Type PicBmp
   Size As Long
   tType As Long
   hBmp As Long
   hPal As Long
   Reserved As Long
End Type
Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hicon As Long) As Long

Public Function GetIconFromFile(FileName As String, IconIndex As Long, UseLargeIcon As Boolean) As Picture

'Parameters:
'FileName - File (EXE or DLL) containing icons
'IconIndex - Index of icon to extract, starting with 0
'UseLargeIcon-True for a large icon, False for a small icon
'Returns: Picture object, containing icon

Dim hlargeicon As Long
Dim hsmallicon As Long
Dim selhandle As Long

' IPicture requires a reference to "Standard OLE Types."
Dim pic As PicBmp
Dim IPic As IPicture
Dim IID_IDispatch As GUID

If ExtractIconEx(FileName, IconIndex, hlargeicon, _
hsmallicon, 1) > 0 Then

If UseLargeIcon Then
selhandle = hlargeicon
Else
selhandle = hsmallicon
End If

' Fill in with IDispatch Interface ID.
With IID_IDispatch
.Data1 = &H20400
.Data4(0) = &HC0
.Data4(7) = &H46
End With
' Fill Pic with necessary parts.
With pic
.Size = Len(pic) ' Length of structure.
.tType = vbPicTypeIcon ' Type of Picture (bitmap).
.hBmp = selhandle ' Handle to bitmap.
End With

' Create Picture object.
Call OleCreatePictureIndirect(pic, IID_IDispatch, 1, IPic)

' Return the new Picture object.
Set GetIconFromFile = IPic

DestroyIcon hsmallicon
DestroyIcon hlargeicon

End If
End Function

