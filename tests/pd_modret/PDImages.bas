Attribute VB_Name = "PDImages"
Option Explicit

Private m_img As pdImage

Public Function GetActiveImage() As pdImage
    Set GetActiveImage = m_img
End Function

Public Function GetImageByID(ByVal imgID As Long) As pdImage
    Set GetImageByID = m_img
End Function

' §56: **标准模块自己**用模块限定符调自己的函数（`PDImages.GetImageByID(id).成员`）。
' §43 那条「限定符是模块名」的推断只走 getClassMethodReturnType，而它的同模块闸是
' isClassModule_ ⇒ 标准模块恒答空 ⇒ 下游把 `.GetNumOfLayers` 当结构体字段发
' （整编实测 PDImages.c 一族 9 条 C2039）。这一句就是那条断链的最小夹具。
Public Function SelfModuleChain(ByVal imgID As Long) As Long
    SelfModuleChain = PDImages.GetImageByID(imgID).GetNumOfLayers()
End Function
