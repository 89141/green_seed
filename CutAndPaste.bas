Attribute VB_Name = "Module1"
Sub CutAndPaste()
Attribute CutAndPaste.VB_ProcData.VB_Invoke_Func = "b\n14"
'
' CutAndPaste Macro
'
' Keyboard Shortcut: Ctrl+b
'
'�\��t��
    ActiveSheet.Paste
'�g���~���O
    Selection.ShapeRange.LockAspectRatio = msoFalse
    Selection.ShapeRange.IncrementLeft 0.7499212598
    Selection.ShapeRange.IncrementTop 127.5
    Selection.ShapeRange.ScaleWidth 0.9994791016, msoFalse, msoScaleFromTopLeft
    Selection.ShapeRange.ScaleHeight 0.8365180304, msoFalse, msoScaleFromTopLeft
    Selection.ShapeRange.PictureFormat.Crop.PictureWidth = 1439
    Selection.ShapeRange.PictureFormat.Crop.PictureHeight = 779
    Selection.ShapeRange.PictureFormat.Crop.PictureOffsetX = 0
    Selection.ShapeRange.PictureFormat.Crop.PictureOffsetY = -63
'�T�C�Y�ύX
    Selection.ShapeRange.ScaleWidth 0.4922490311, msoFalse, msoScaleFromTopLeft
    Selection.ShapeRange.ScaleHeight 0.4922489808, msoFalse, msoScaleFromTopLeft
'�ړ�
    Selection.ShapeRange.IncrementTop -93.75
'�I���Z��
    ActiveCell.Offset(20, 0).Activate

End Sub
