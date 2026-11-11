Attribute VB_Name = "PDImages"
Option Explicit

Private m_img As pdImage

Public Function GetActiveImage() As pdImage
    Set GetActiveImage = m_img
End Function

Public Function GetImageByID(ByVal imgID As Long) As pdImage
    Set GetImageByID = m_img
End Function
