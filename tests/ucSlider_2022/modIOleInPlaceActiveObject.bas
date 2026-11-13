Attribute VB_Name = "modIOleInPlaceActiveObject"
Option Explicit

' ----==== Const ====----
Private Const S_OK As Long = &H0&
Private Const S_FALSE As Long = &H1&
Private Const WM_KEYUP As Long = &H101&
Private Const WM_KEYDOWN As Long = &H100&
Private Const CC_STDCALL As Long = &H4&
Private Const IID_Release As Long = &H8&
Private Const E_NOINTERFACE As Long = &H80004002
Private Const OLEIVERB_UIACTIVATE As Long = &HFFFFFFFC

Private Const IID_IOleObject As String = "{00000112-0000-0000-c000-000000000046}"
Private Const IID_IOleInPlaceSite As String = "{00000119-0000-0000-c000-000000000046}"
Private Const IID_IOleInPlaceActiveObject As String = "{00000117-0000-0000-c000-000000000046}"

Private Enum vtb_Interfaces

    ' IUnknown
    QueryInterface = 0
    
    ' IOleObject
    GetClientSite = 4
    DoVerb = 11
    
    ' IOleInPlaceSite
    GetWindowContext = 8
    
    ' IOleInPlaceFrame / IOleInPlaceUIWindow
    SetActiveObject = 8
    
    ' IOleInPlaceActiveObject
    TranslateAccelerator = 5
    
End Enum

Private Type POINT
    X As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type MSG
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINT
    lPrivate As Long
End Type

Private Type OLEINPLACEFRAMEINFO
    cb As Long
    fMDIApp As Long
    hwndFrame As Long
    haccel As Long
    cAccelEntries As Long
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type tIOleInPlaceActiveObject
    pVTable As Long
End Type

Private Type tIOleInPlaceActiveObjectVTable
    VTable(0 To 9) As Long
End Type

' ----==== Ole32.dll ====----
Private Declare Sub CoTaskMemFree Lib "Ole32.dll" ( _
                    ByVal hMem As Long)

Private Declare Function StringFromCLSID Lib "Ole32.dll" ( _
                         ByRef pCLSID As GUID, _
                         ByRef lpszProgID As Long) As Long

Private Declare Function CLSIDFromString Lib "Ole32.dll" ( _
                         ByVal pstring As Long, _
                         ByRef pCLSID As GUID) As Long

' ----==== User32.dll ====----
Private Declare Function SendMessageLong Lib "user32.dll" _
                         Alias "SendMessageA" ( _
                         ByVal hwnd As Long, _
                         ByVal wMsg As Long, _
                         ByVal wParam As Long, _
                         ByRef lParam As Long) As Long

' ----==== Kernel32.dll ====----
Private Declare Sub CopyMemory Lib "Kernel32.dll" _
                    Alias "RtlMoveMemory" ( _
                    ByRef Destination As Any, _
                    ByRef Source As Any, _
                    ByVal length As Long)

Private Declare Function lstrlenW Lib "Kernel32.dll" ( _
                         ByVal lpString As Long) As Long

' ----==== Oleaut32.dll Deklarationen ====----
Private Declare Function DispCallFunc Lib "Oleaut32.dll" ( _
                         ByVal pvInstance As Long, _
                         ByVal oVft As Long, _
                         ByVal cc As Long, _
                         ByVal vtReturn As VbVarType, _
                         ByVal cActuals As Long, _
                         ByVal prgvt As Long, _
                         ByVal prgpvarg As Long, _
                         ByRef pvargResult As Variant) As Long


' ----==== Variablen ====----
Private m_IPAO As Long
Private m_RefCount As Long
Private m_IOleInPlaceActiveObject As tIOleInPlaceActiveObject
Private m_IOleInPlaceActiveObjectVTable As tIOleInPlaceActiveObjectVTable

Public Sub InitIPAO()

    If m_IPAO = 0& Then

        ' ----==== Interface IOleInPlaceActiveObject  ====----
        With m_IOleInPlaceActiveObjectVTable
            .VTable(0) = ProcPtr(AddressOf QueryInterface_)
            .VTable(1) = ProcPtr(AddressOf AddRef_)
            .VTable(2) = ProcPtr(AddressOf Release_)
            .VTable(3) = ProcPtr(AddressOf GetWindow_)
            .VTable(4) = ProcPtr(AddressOf ContextSensitiveHelp_)
            .VTable(5) = ProcPtr(AddressOf TranslateAccelerator_)
            .VTable(6) = ProcPtr(AddressOf OnFrameWindowActivate_)
            .VTable(7) = ProcPtr(AddressOf OnDocWindowActivate_)
            .VTable(8) = ProcPtr(AddressOf ResizeBorder_)
            .VTable(9) = ProcPtr(AddressOf EnableModeless_)
        End With

        With m_IOleInPlaceActiveObject
            .pVTable = VarPtr(m_IOleInPlaceActiveObjectVTable)
        End With

        m_IPAO = VarPtr(m_IOleInPlaceActiveObject)
        
    End If

End Sub

Public Sub TerminateIPAO()
    
    Call Release(m_IPAO)
    
End Sub

Public Function SetIPAO(Ctrl As Object)

    Dim pIOleObject As Long
    Dim pIOleClientSite As Long
    Dim pIOleInPlaceSite As Long
    Dim pIOleInPlaceFrame As Long
    Dim pIOleInPlaceUIWindow As Long
    Dim rcPos As RECT
    Dim rcClip As RECT
    Dim uFrameInfo As OLEINPLACEFRAMEINFO
    
    If Invoke(ObjPtr(Ctrl), QueryInterface, VarPtr(Str2Guid(IID_IOleObject)), _
        VarPtr(pIOleObject)) = S_OK Then
        
        If Invoke(pIOleObject, GetClientSite, VarPtr(pIOleClientSite)) = S_OK Then
        
            If Invoke(pIOleClientSite, QueryInterface, VarPtr(Str2Guid(IID_IOleInPlaceSite)), _
                VarPtr(pIOleInPlaceSite)) = S_OK Then
                
                If Invoke(pIOleInPlaceSite, GetWindowContext, _
                    VarPtr(pIOleInPlaceFrame), VarPtr(pIOleInPlaceUIWindow), _
                    VarPtr(rcPos), VarPtr(rcClip), VarPtr(uFrameInfo)) = S_OK Then
                
                    If pIOleInPlaceFrame <> 0& Then
                    
                        If Invoke(pIOleInPlaceFrame, SetActiveObject, _
                            m_IPAO, 0&) = S_OK Then

                        End If
                        
                        Call Release(pIOleInPlaceFrame)
                        
                    End If
            
                    If pIOleInPlaceUIWindow <> 0& Then
                    
                        If Invoke(pIOleInPlaceUIWindow, SetActiveObject, _
                            m_IPAO, 0&) = S_OK Then
    
                        End If
                        
                        Call Release(pIOleInPlaceUIWindow)
                        
                    Else
                
                        If Invoke(pIOleObject, DoVerb, OLEIVERB_UIACTIVATE, 0&, _
                            pIOleInPlaceSite, 0&, 0&, VarPtr(rcPos)) = S_OK Then
                    
                        End If
                    
                    End If
            
            
                End If
            
                Call Release(pIOleInPlaceSite)
            
            End If

            Call Release(pIOleClientSite)

        End If
        
        Call Release(pIOleObject)
        
    End If

End Function

' ----==== Interface IUnknown ====----
Private Function QueryInterface_(ByVal this As Long, _
                                 ByRef riid As GUID, _
                                 ByRef pvObj As Long) As Long

    Dim lRet As Long
    
    Select Case UCase$(Guid2Str(riid))

    Case UCase$(IID_IOleInPlaceActiveObject)

        Call AddRef_(this)
        
        pvObj = m_IPAO
        
        lRet = S_OK

    Case Else
        
        pvObj = 0&
        
        lRet = E_NOINTERFACE
    
    End Select

    QueryInterface_ = lRet

End Function

Private Function AddRef_(ByVal this As Long) As Long

    m_RefCount = m_RefCount + 1

    AddRef_ = m_RefCount

End Function

Private Function Release_(ByVal this As Long) As Long
    
    m_RefCount = m_RefCount - 1

    If m_RefCount = 0 Then
        
        Call CopyMemory(m_IPAO, 0&, 4)
        
    End If
    
    Release_ = m_RefCount

End Function

' ----==== Interface IOleWindow ====----
Private Function GetWindow_(ByVal this As Long, _
                            ByRef phwnd As Long) As Long

End Function

Private Function ContextSensitiveHelp_(ByVal this As Long, _
                                       ByVal fEnterMode As Long) As Long

End Function

' ----==== Interface IOleInPlaceActiveObject ====----
Private Function TranslateAccelerator_(ByVal this As Long, _
                                       ByRef lpMsg As MSG) As Long
    Select Case lpMsg.message
    
        Case WM_KEYDOWN, WM_KEYUP
        
            Select Case lpMsg.wParam
            
                Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, _
                     vbKeyHome, vbKeyEnd, vbKeyPageDown, vbKeyPageUp
                    
                    Call SendMessageLong(lpMsg.hwnd, lpMsg.message, lpMsg.wParam, lpMsg.lParam)
                    
                    TranslateAccelerator_ = S_OK
                
                Case Else
                    
                    TranslateAccelerator_ = S_FALSE
                    
            End Select
            
        Case Else
        
           TranslateAccelerator_ = S_FALSE
        
    End Select

End Function

Private Function OnFrameWindowActivate_(ByVal this As Long, _
                                        ByVal fActivate As Long) As Long

End Function

Private Function OnDocWindowActivate_(ByVal this As Long, _
                                      ByVal fActivate As Long) As Long

End Function

Private Function ResizeBorder_(ByVal this As Long, _
                               prcBorder As RECT, _
                               ByVal puiWindow As Long, _
                               ByVal fFrameWindow As Long) As Long

End Function

Private Function EnableModeless_(ByVal this As Long, _
                                 ByVal fEnable As Long) As Long

End Function

' ----============================================----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal eInterfaceFunction As vtb_Interfaces, _
                        ParamArray arrParam()) As Variant

    If pInterface <> 0& Then
        
        Invoke = OleInvoke(pInterface, eInterfaceFunction, arrParam)

    End If

End Function

Private Sub Release(ByRef pInterface As Long)

    Dim varRet As Variant

    If pInterface <> 0& Then

        If DispCallFunc(pInterface, IID_Release, CC_STDCALL, _
            vbLong, 0&, 0&, 0&, varRet) = S_OK Then

            pInterface = 0&

        End If

    End If

End Sub

Private Function OleInvoke(ByVal pInterface As Long, _
                           ByVal lngCmd As Long, _
                           ParamArray arrParam()) As Variant
    
    Dim lngRet As Long
    Dim lngItem As Long
    Dim lngCount As Long
    Dim varRet As Variant
    Dim varParam As Variant
    Dim olePtr(10) As Long
    Dim oleTyp(10) As Integer

    If pInterface <> 0& Then

        If UBound(arrParam) >= 0 Then

            varParam = arrParam

            If IsArray(varParam) Then varParam = varParam(0)

            lngCount = UBound(varParam)

            For lngItem = 0 To lngCount

                oleTyp(lngItem) = VarType(varParam(lngItem))
                olePtr(lngItem) = VarPtr(varParam(lngItem))

            Next

        End If

        lngRet = DispCallFunc(pInterface, lngCmd * 4, CC_STDCALL, vbLong, _
            lngItem, VarPtr(oleTyp(0)), VarPtr(olePtr(0)), varRet)
            
        If lngRet <> S_OK Then
            
            'Debug.Print "Error: 0x" & Hex$(lngRet)
            
            varRet = S_FALSE
            
        End If
        
    End If

    OleInvoke = varRet

End Function

Private Function ProcPtr(ByVal ptr As Long) As Long

    ProcPtr = ptr

End Function

Private Function StrPtr2Str(ByVal lpStrPointer As Long) As String

    Dim lLen As Long
    Dim bBuffer() As Byte

    lLen = lstrlenW(lpStrPointer) * 2

    If lLen > 0 Then

        ReDim bBuffer(lLen - 1)

        Call CopyMemory(bBuffer(0), ByVal lpStrPointer, lLen)

        Call CoTaskMemFree(lpStrPointer)

        StrPtr2Str = bBuffer

    End If

End Function

Private Function Str2Guid(ByVal str As String) As GUID

    Call CLSIDFromString(StrPtr(str), Str2Guid)
    
End Function


Private Function Guid2Str(ByRef tguid As GUID) As String

    Dim lGuid As Long
    
    Call StringFromCLSID(tguid, lGuid)

    Guid2Str = StrPtr2Str(lGuid)

End Function


