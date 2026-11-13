VERSION 5.00
Begin VB.Form Form1
   Caption         =   "Form1"
End
Attribute VB_Name = "Form1"
Option Explicit

Private m_SourceImageID As Long
Private m_SourceLayerID As Long
Private m_SampleMergedCopy As pdDIB

Private Sub Form_Load()
    Dim b As Boolean
    b = (PDImages.GetImageByID(m_SourceImageID).GetNumOfLayers = 1)
    If b Then b = Not PDImages.GetActiveImage.GetLayerByID(m_SourceLayerID).AffineTransformsActive(True)
    If b Then
        Set m_SampleMergedCopy = PDImages.GetImageByID(m_SourceImageID).GetLayerByID(m_SourceLayerID).GetLayerDIB
    Else
        PDImages.GetImageByID(m_SourceImageID).GetCompositedImage m_SampleMergedCopy, True
    End If
    Debug.Print m_SampleMergedCopy.GetDIBWidth, b
End Sub
