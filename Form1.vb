Imports System.Runtime.InteropServices
Imports System.Security.Policy
Imports Direct2D
Imports DXGI
Imports DXGI.DXGITools
Imports GlobalStructures
Imports GlobalStructures.GlobalTools
Imports Sprite
Imports Sprite.CSprite
Imports WIC

Public Class Form1

    Public Const WM_SIZE = &H5
    Public Const WM_PAINT = &HF

    <DllImport("User32.dll", SetLastError:=True)>
    Public Shared Function BeginPaint(ByVal hWnd As IntPtr, ByRef lpPaint As PAINTSTRUCT) As IntPtr
    End Function

    <DllImport("User32.dll", SetLastError:=True)>
    Public Shared Function EndPaint(ByVal hWnd As IntPtr, ByRef lpPaint As PAINTSTRUCT) As Boolean
    End Function

    <StructLayout(LayoutKind.Sequential)>
    Public Structure PAINTSTRUCT
        Public hdc As IntPtr
        Public fErase As Boolean
        Public rcPaint_left As Integer
        Public rcPaint_top As Integer
        Public rcPaint_right As Integer
        Public rcPaint_bottom As Integer
        Public fRestore As Boolean
        Public fIncUpdate As Boolean
        Public reserved1 As Integer
        Public reserved2 As Integer
        Public reserved3 As Integer
        Public reserved4 As Integer
        Public reserved5 As Integer
        Public reserved6 As Integer
        Public reserved7 As Integer
        Public reserved8 As Integer
    End Structure

    <DllImport("User32.dll", SetLastError:=True)>
    Public Shared Function InvalidateRect(hWnd As IntPtr, ByRef lpRect As RECT, bErase As Boolean) As Boolean
    End Function

    <DllImport("User32.dll", SetLastError:=True)>
    Public Shared Function InvalidateRect(hWnd As IntPtr, lpRect As IntPtr, bErase As Boolean) As Boolean
    End Function

    <DllImport("User32.dll", SetLastError:=True)>
    Public Shared Function GetClientRect(hWnd As IntPtr, ByRef lpRect As RECT) As Boolean
    End Function

    <DllImport("User32.dll", SetLastError:=True)>
    Public Shared Function IsIconic(ByVal hWnd As IntPtr) As Boolean
    End Function

    <DllImport("User32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Public Shared Function ShowWindow(hWnd As IntPtr, nShowCmd As Integer) As Boolean
    End Function

    Public Const SW_HIDE As Integer = 0
    Public Const SW_SHOWNORMAL As Integer = 1
    Public Const SW_SHOWMINIMIZED As Integer = 2
    Public Const SW_SHOWMAXIMIZED As Integer = 3
    Public Const SW_SHOWNOACTIVATE As Integer = 4
    Public Const SW_SHOW As Integer = 5


    Dim m_pD2DFactory As ID2D1Factory = Nothing
    Dim m_pD2DFactory1 As ID2D1Factory1 = Nothing
    Dim m_pWICImagingFactory As IWICImagingFactory = Nothing

    Dim m_pD3D11DevicePtr As IntPtr = IntPtr.Zero
    Dim m_pD3D11DeviceContextPtr As IntPtr = IntPtr.Zero
    Dim m_pDXGIDevice As IDXGIDevice1 = Nothing

    Dim m_pD2DDevice As ID2D1Device = Nothing
    Dim m_pD2DDeviceContext As ID2D1DeviceContext = Nothing
    Dim m_pD2DDeviceContext3 As ID2D1DeviceContext3 = Nothing

    Dim m_pD2DBitmapBackground As ID2D1Bitmap = Nothing
    Dim m_pD2DBitmap As ID2D1Bitmap = Nothing
    Dim m_pD2DBitmap1 As ID2D1Bitmap = Nothing
    Dim m_pD2DBitmap2 As ID2D1Bitmap = Nothing
    Dim m_pD2DBitmapBackgroundMask As ID2D1Bitmap = Nothing

    Dim m_pD2DTargetBitmap As ID2D1Bitmap1 = Nothing
    Dim m_pDXGISwapChain1 As IDXGISwapChain1 = Nothing
    Dim m_pMainBrush As ID2D1SolidColorBrush = Nothing

    Private m_nSprite1Count As Integer = 15
    Private m_nSprite2Count As Integer = 8
    Private m_nBubblesCount As Integer = 5

    Private rand As Random = Nothing
    Private randColor As Random = Nothing
    Private CSprites As List(Of CSprite) = New List(Of CSprite)()
    Private CSprites2 As List(Of CSprite) = New List(Of CSprite)()
    Private CSprites3 As List(Of CSprite) = New List(Of CSprite)()

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim hr As HRESULT = HRESULT.S_OK
        'Me.ResizeRedraw = True
        Me.DoubleBuffered = True
        Me.ClientSize = New System.Drawing.Size(1280, 900)
        Me.Text = "VB - Direct2D - SpriteBatch"

        m_pWICImagingFactory = CType(Activator.CreateInstance(Type.GetTypeFromCLSID(WICTools.CLSID_WICImagingFactory)), IWICImagingFactory)
        hr = CreateD2D1Factory()
        hr = CreateD3D11Device()
        hr = CreateDeviceResources()
        hr = CreateSwapChain(Me.Handle)
        If SUCCEEDED(hr) Then hr = ConfigureSwapChain()

        ' To avoid background painting the first time the window is shown
        ShowWindow(Me.Handle, SW_HIDE)
        ShowWindow(Me.Handle, SW_SHOWMINIMIZED)
        ShowWindow(Me.Handle, SW_SHOWNORMAL)

        rand = New Random()
        randColor = New Random()
        Me.CenterToScreen()
    End Sub

    Protected Overrides Sub WndProc(ByRef m As Message)
        If (m.Msg = WM_PAINT) Then
            Dim hr As HRESULT = OnPaintProc(m.HWnd)

            If CSprites.Count < m_nSprite1Count Then
                If m_pD2DBitmap IsNot Nothing Then
                    'AddSprite(m_pD2DBitmap, 13, 7, 87)
                    'AddSprite(m_pD2DBitmap, 20, 10, 0)
                    AddSprite(1, m_pD2DBitmap, 11, 6, 0)
                End If
            End If

            If CSprites3.Count < m_nSprite2Count Then
                If m_pD2DBitmap2 IsNot Nothing Then
                    AddSprite(3, m_pD2DBitmap2, 8, 4, 0)
                End If
            End If

            If CSprites2.Count < m_nBubblesCount Then
                If m_pD2DBitmap1 IsNot Nothing Then
                    AddBubble(m_pD2DBitmap1, 4, 4, 0)
                End If
            End If

            m.Result = CType(hr, IntPtr)
        ElseIf (m.Msg = WM_SIZE) Then
            Dim hr As HRESULT = OnResizeProc(m)
            m.Result = IntPtr.Zero
        Else
            MyBase.WndProc(m)
        End If
    End Sub

    Public Function OnPaintProc(hWnd As IntPtr) As HRESULT
        Dim hr As HRESULT = HRESULT.S_OK
        Dim ps As PAINTSTRUCT = New PAINTSTRUCT
        If (BeginPaint(hWnd, ps) <> IntPtr.Zero) Then

            If (m_pD2DDeviceContext IsNot Nothing) Then
                m_pD2DDeviceContext.BeginDraw()

                Dim size As D2D1_SIZE_F
                m_pD2DDeviceContext.GetSize(size)
                If m_pD2DBitmapBackground IsNot Nothing Then
                    Dim sizeBmpBackground As D2D1_SIZE_F
                    m_pD2DBitmapBackground.GetSize(sizeBmpBackground)
                    Dim destRectBackground As D2D1_RECT_F = New D2D1_RECT_F(0.0F, 0.0F, size.width, size.height)
                    Dim sourceRectBackground As D2D1_RECT_F = New D2D1_RECT_F(0.0F, 0.0F, sizeBmpBackground.width, sizeBmpBackground.height)
                    m_pD2DDeviceContext.DrawBitmap(m_pD2DBitmapBackground, destRectBackground, 1.0F, D2D1_BITMAP_INTERPOLATION_MODE.D2D1_BITMAP_INTERPOLATION_MODE_LINEAR, sourceRectBackground)
                Else
                    m_pD2DDeviceContext.Clear(New ColorF(ColorF.Enum.Black))
                End If

                'If m_pMainBrush IsNot Nothing Then
                '    Dim rect = New D2D1_RECT_F(20, 20, size.width - 20, size.height - 20)
                '    m_pD2DDeviceContext.DrawRectangle(rect, m_pMainBrush, 2)
                'End If

                Dim nDPI As Single = GetDpiForWindow(hWnd)
                Dim nScale As Single = (nDPI / 96.0F)
                Dim rect As RECT
                GetClientRect(hWnd, rect)

                If (Not IsIconic(hWnd)) Then
                    If m_pD2DDeviceContext3 IsNot Nothing Then
                        Dim nOldAntialiasMode = Nothing
                        m_pD2DDeviceContext3.GetAntialiasMode(nOldAntialiasMode)
                        m_pD2DDeviceContext3.SetAntialiasMode(D2D1_ANTIALIAS_MODE.D2D1_ANTIALIAS_MODE_ALIASED)

                        For Each s As CSprite In CSprites
                            s.X += ((rand.NextDouble()) * s.StepX)
                            s.Y += ((rand.NextDouble()) * s.StepY)
                            s.Move(New D2D1_SIZE_F(rect.right / nScale, rect.bottom / nScale), m_pD2DDeviceContext3, HORIZONTALFLIP.LEFT, BOUNCE.BOTH)
                            s.Draw(m_pD2DDeviceContext3, s.CurrentIndex, 1, True)
                            s.CurrentIndex += 1
                        Next

                        For Each s As CSprite In CSprites3
                            s.X += ((rand.NextDouble()) * s.StepX)
                            s.Y += ((rand.NextDouble()) * s.StepY)
                            s.Move(New D2D1_SIZE_F(rect.right / nScale, rect.bottom / nScale), m_pD2DDeviceContext3, HORIZONTALFLIP.NONE, BOUNCE.BOTH)
                            s.Draw(m_pD2DDeviceContext3, s.CurrentIndex, 1, True)
                            s.CurrentIndex += 1
                        Next

                        Dim CSpritesDelete As List(Of CSprite) = New List(Of CSprite)()
                        For Each s As CSprite In CSprites2
                            's.X += ((rand.NextDouble()) * s.StepX)
                            s.Y -= ((rand.NextDouble()) * s.StepY)
                            s.Move(New D2D1_SIZE_F(rect.right / nScale, rect.bottom / nScale), m_pD2DDeviceContext3, HORIZONTALFLIP.NONE, BOUNCE.NONE)
                            s.Draw(m_pD2DDeviceContext3, s.CurrentIndex, 1, True)
                            s.CurrentIndex += 1
                            If (s.Y <= 0 - s.Height) Then
                                CSpritesDelete.Add(s)
                            End If
                        Next

                        For Each s As CSprite In CSpritesDelete
                            CSprites2.Remove(s)
                            AddBubble(m_pD2DBitmap1, 4, 4, 0)
                        Next
                        CSpritesDelete.Clear()

                        m_pD2DDeviceContext3.SetAntialiasMode(nOldAntialiasMode)
                    End If
                End If

                If m_pD2DBitmapBackground IsNot Nothing And m_pD2DBitmapBackgroundMask IsNot Nothing Then
                    Dim sizeBmpBackground As D2D1_SIZE_F
                    m_pD2DBitmapBackground.GetSize(sizeBmpBackground)

                    Dim vignetteEffect As ID2D1Effect = Nothing
                    hr = m_pD2DDeviceContext.CreateEffect(D2DTools.CLSID_D2D1Vignette, vignetteEffect)
                    'Dim vector4Color As New D2D1_VECTOR_4F(0.9F, 0.55F, 0.9F, 1.0F)
                    'Dim vector4Color As New D2D1_VECTOR_4F(0.0F, 1.0F, 1.0F, 1.0F) 'Cyan
                    Dim vector4Color As New D2D1_VECTOR_4F(0.0F, 0.0F, 1.0F, 1.0F) 'Blue
                    'Dim vector4Color As New D2D1_VECTOR_4F(0.0F, 0.0F, &H8B / 255.0F, 1.0F) ' DarkBlue
                    Dim pVector4Color As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(GetType(D2D1_VECTOR_4F)))
                    Marshal.StructureToPtr(vector4Color, pVector4Color, False)
                    hr = vignetteEffect.SetValue(D2D1_VIGNETTE_PROP.D2D1_VIGNETTE_PROP_COLOR, D2D1_PROPERTY_TYPE.D2D1_PROPERTY_TYPE_VECTOR4, pVector4Color, Marshal.SizeOf(GetType(D2D1_VECTOR_4F)))
                    Marshal.FreeHGlobal(pVector4Color)

                    Dim pTransitionsize = Marshal.AllocHGlobal(Marshal.SizeOf(GetType(Single)))
                    Dim f = {0.5F}
                    Marshal.Copy(f, 0, pTransitionsize, 1)
                    hr = vignetteEffect.SetValue(D2D1_VIGNETTE_PROP.D2D1_VIGNETTE_PROP_TRANSITION_SIZE, D2D1_PROPERTY_TYPE.D2D1_PROPERTY_TYPE_UNKNOWN, pTransitionsize, Marshal.SizeOf(GetType(Single)))
                    Marshal.FreeHGlobal(pTransitionsize)

                    Dim pStrength = Marshal.AllocHGlobal(Marshal.SizeOf(GetType(Single)))
                    Dim f2 = {0.5F}
                    Marshal.Copy(f2, 0, pStrength, 1)
                    hr = vignetteEffect.SetValue(D2D1_VIGNETTE_PROP.D2D1_VIGNETTE_PROP_STRENGTH, D2D1_PROPERTY_TYPE.D2D1_PROPERTY_TYPE_UNKNOWN, pStrength, Marshal.SizeOf(GetType(Single)))
                    Marshal.FreeHGlobal(pStrength)

                    Dim scaleEffect As ID2D1Effect = Nothing
                    hr = m_pD2DDeviceContext.CreateEffect(D2DTools.CLSID_D2D1Scale, scaleEffect)
                    Dim nScaleX As Single = size.width / sizeBmpBackground.width
                    Dim nScaleY As Single = size.height / sizeBmpBackground.height

                    Dim vector2 As D2D1_VECTOR_2F = New D2D1_VECTOR_2F(nScaleX, nScaleY)
                    Dim pVector2 As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(GetType(D2D1_VECTOR_2F)))
                    Marshal.StructureToPtr(vector2, pVector2, False)
                    hr = scaleEffect.SetValue(D2D1_SCALE_PROP.D2D1_SCALE_PROP_SCALE, D2D1_PROPERTY_TYPE.D2D1_PROPERTY_TYPE_VECTOR2, pVector2, Marshal.SizeOf(GetType(D2D1_VECTOR_2F)))
                    Marshal.FreeHGlobal(pVector2)

                    Dim outputScale As ID2D1Image = Nothing
                    scaleEffect.GetOutput(outputScale)
                    scaleEffect.SetInput(0, m_pD2DBitmapBackgroundMask)
                    vignetteEffect.SetInput(0, outputScale, True)
                    Dim outputVignette As ID2D1Image = Nothing
                    vignetteEffect.GetOutput(outputVignette)

                    Dim pt As D2D1_POINT_2F = New D2D1_POINT_2F(0, 0)
                    Dim sourceRectangle As D2D1_RECT_F = New D2D1_RECT_F(0, 0, size.width, size.height)
                    m_pD2DDeviceContext.DrawImage(vignetteEffect, pt, sourceRectangle, D2D1_INTERPOLATION_MODE.D2D1_INTERPOLATION_MODE_LINEAR, D2D1_COMPOSITE_MODE.D2D1_COMPOSITE_MODE_SOURCE_OVER)
                    SafeRelease(outputScale)
                    SafeRelease(scaleEffect)
                    SafeRelease(vignetteEffect)
                End If

                Dim tag1 As ULong, tag2 As ULong = 0
                hr = m_pD2DDeviceContext.EndDraw(tag1, tag2)
                ' 0x88990001 D2DERR_WRONG_STATE
                If hr = D2DTools.D2DERR_RECREATE_TARGET Then
                    m_pD2DDeviceContext.SetTarget(Nothing)
                    SafeRelease(m_pD2DDeviceContext)
                    hr = CreateD3D11Device()
                    hr = CreateDeviceResources()
                    hr = CreateSwapChain(Me.Handle)
                    If SUCCEEDED(hr) Then hr = ConfigureSwapChain()
                End If
                hr = m_pDXGISwapChain1.Present(1, 0)
            End If
            EndPaint(hWnd, ps)
        End If
        InvalidateRect(Me.Handle, IntPtr.Zero, False)
        Return (hr)
    End Function

    Public Function OnResizeProc(ByRef m As Message) As HRESULT
        Dim hr As HRESULT = HRESULT.S_OK

        If (m_pDXGISwapChain1 IsNot Nothing) Then

            If (m_pD2DDeviceContext IsNot Nothing) Then
                m_pD2DDeviceContext.SetTarget(Nothing)
            End If

            If (m_pD2DTargetBitmap IsNot Nothing) Then
                SafeRelease(m_pD2DTargetBitmap)
            End If

            '// 0, 0 => HRESULT 0x80070057 (E_INVALIDARG) if Not CreateSwapChainForHwnd
            hr = m_pDXGISwapChain1.ResizeBuffers(
             2,
             0,
             0,
             DXGI.DXGI_FORMAT.DXGI_FORMAT_B8G8R8A8_UNORM,
             0
             )
            'hr = m_pDXGISwapChain1.ResizeBuffers(
            '  2,
            '  (uint)sz.Width,
            '  (uint)sz.Height,
            '  DXGI_FORMAT.DXGI_FORMAT_B8G8R8A8_UNORM,
            '  0
            '  );

            ConfigureSwapChain()
        End If
        Return hr
    End Function

    Private Function CreateD2D1Factory() As HRESULT
        Dim hr As HRESULT = HRESULT.S_OK
        Dim options As D2D1_FACTORY_OPTIONS = New D2D1_FACTORY_OPTIONS()
#If DEBUG Then
        options.debugLevel = D2D1_DEBUG_LEVEL.D2D1_DEBUG_LEVEL_NONE
#End If
        hr = D2DTools.D2D1CreateFactory(D2D1_FACTORY_TYPE.D2D1_FACTORY_TYPE_SINGLE_THREADED, D2DTools.CLSID_D2D1Factory, options, m_pD2DFactory)
        m_pD2DFactory1 = CType(m_pD2DFactory, ID2D1Factory1)
        Return hr
    End Function

    Public Function CreateD3D11Device() As HRESULT
        Dim hr As HRESULT = HRESULT.S_OK
        Dim creationFlags = D3D11_CREATE_DEVICE_FLAG.D3D11_CREATE_DEVICE_BGRA_SUPPORT
#If DEBUG Then
        creationFlags = creationFlags Or D3D11_CREATE_DEVICE_FLAG.D3D11_CREATE_DEVICE_DEBUG
#End If

        Dim aD3D_FEATURE_LEVEL = New Integer() {D3D_FEATURE_LEVEL.D3D_FEATURE_LEVEL_11_1, D3D_FEATURE_LEVEL.D3D_FEATURE_LEVEL_11_0, D3D_FEATURE_LEVEL.D3D_FEATURE_LEVEL_10_1, D3D_FEATURE_LEVEL.D3D_FEATURE_LEVEL_10_0, D3D_FEATURE_LEVEL.D3D_FEATURE_LEVEL_9_3, D3D_FEATURE_LEVEL.D3D_FEATURE_LEVEL_9_2, D3D_FEATURE_LEVEL.D3D_FEATURE_LEVEL_9_1}

        Dim featureLevel As D3D_FEATURE_LEVEL
        hr = D3D11CreateDevice(Nothing, D3D_DRIVER_TYPE.D3D_DRIVER_TYPE_HARDWARE, IntPtr.Zero, creationFlags, aD3D_FEATURE_LEVEL, aD3D_FEATURE_LEVEL.Length, D3D11_SDK_VERSION, m_pD3D11DevicePtr, featureLevel, m_pD3D11DeviceContextPtr)    ' specify null to use the default adapter
        If SUCCEEDED(hr) Then
            'm_pD3D11DeviceContext = Marshal.GetObjectForIUnknown(pD3D11DeviceContextPtr) as ID3D11DeviceContext;             

            m_pDXGIDevice = TryCast(Marshal.GetObjectForIUnknown(m_pD3D11DevicePtr), IDXGIDevice1)
            hr = m_pD2DFactory1.CreateDevice(m_pDXGIDevice, m_pD2DDevice)
            If SUCCEEDED(hr) Then
                hr = m_pD2DDevice.CreateDeviceContext(D2D1_DEVICE_CONTEXT_OPTIONS.D2D1_DEVICE_CONTEXT_OPTIONS_NONE, m_pD2DDeviceContext)
                Marshal.ReleaseComObject(m_pD2DDevice)
            End If
            Marshal.Release(m_pD3D11DevicePtr)
            Marshal.Release(m_pD3D11DeviceContextPtr)
        End If
        Return hr
    End Function

    Private Function CreateDeviceResources() As HRESULT
        Dim hr As HRESULT = HRESULT.S_OK
        If m_pD2DDeviceContext IsNot Nothing Then
            If m_pMainBrush Is Nothing Then hr = m_pD2DDeviceContext.CreateSolidColorBrush(New ColorF(ColorF.Enum.Red), Nothing, m_pMainBrush)
            'If m_pD2DBitmapBackground Is Nothing Then hr = CreateD2DBitmapFromURL("https://i.ibb.co/2MtgC8C/clouds-country-daylight-371633.jpg", m_pD2DBitmapBackground)
            'If m_pD2DBitmapBackground Is Nothing Then hr = CreateD2DBitmapFromURL("https://i.ibb.co/34dTvWW/evening-landscape.jpg", m_pD2DBitmapBackground)
            If m_pD2DBitmapBackground Is Nothing Then hr = CreateD2DBitmapFromURL("https://i.ibb.co/VW8wxjnz/Corals.jpg", m_pD2DBitmapBackground)
            'If m_pD2DBitmap Is Nothing Then hr = CreateD2DBitmapFromURL("https://i.ibb.co/QCBKBjD/Flying-bird.png", m_pD2DBitmap)
            'If m_pD2DBitmap Is Nothing Then hr = CreateD2DBitmapFromURL("https://i.ibb.co/VgVp09Y/butterfly-sprite-sheet-blue.png", m_pD2DBitmap)
            'If m_pD2DBitmap Is Nothing Then hr = CreateD2DBitmapFromURL("https://i.ibb.co/7Jj8dc4F/jellyfish-20x10.png", m_pD2DBitmap)
            If m_pD2DBitmap Is Nothing Then hr = CreateD2DBitmapFromURL("https://i.ibb.co/5hZ1LD1r/Octopus-11x6.png", m_pD2DBitmap)
            If m_pD2DBitmap1 Is Nothing Then hr = CreateD2DBitmapFromURL("https://i.ibb.co/Vw8vzRT/Bubbles2.png", m_pD2DBitmap1)
            If m_pD2DBitmap2 Is Nothing Then hr = CreateD2DBitmapFromURL("https://i.ibb.co/FLBZM0H7/octopus8x4.png", m_pD2DBitmap2)

            If m_pD2DDeviceContext3 Is Nothing Then m_pD2DDeviceContext3 = CType(m_pD2DDeviceContext, ID2D1DeviceContext3)

            If (m_pD2DBitmapBackground IsNot Nothing) Then
                Dim sizeBmpBackground As D2D1_SIZE_F
                m_pD2DBitmapBackground.GetSize(sizeBmpBackground)
                Dim nPitch = sizeBmpBackground.width * 4
                Dim nSize = CInt(nPitch * sizeBmpBackground.height)
                Dim pBytesArray = New Byte(nSize - 1) {}
                ' Useless, just to test BGRA...
                For nY As Integer = 0 To sizeBmpBackground.height - 1
                    For nX As Integer = 0 To sizeBmpBackground.width - 1
                        Dim nPos = CInt(nY * nPitch + nX * 4)
                        Dim b As Byte = pBytesArray(nPos), g As Byte = pBytesArray(nPos + 1), r As Byte = pBytesArray(nPos + 2), a As Byte = pBytesArray(nPos + 3)
                        pBytesArray(nPos) = b
                        pBytesArray(nPos + 1) = g
                        pBytesArray(nPos + 2) = r
                        pBytesArray(nPos + 3) = 0
                    Next
                Next
                Dim pArray As IntPtr = Marshal.AllocCoTaskMem(Marshal.SizeOf(GetType(Byte)) * pBytesArray.Length)
                Marshal.Copy(pBytesArray, 0, pArray, pBytesArray.Length)
                Dim sizeU As D2D1_SIZE_U = New D2D1_SIZE_U(sizeBmpBackground.width, sizeBmpBackground.height)
                Dim bitmapProperties As D2D1_BITMAP_PROPERTIES1 = New D2D1_BITMAP_PROPERTIES1()
                bitmapProperties.pixelFormat = D2DTools.PixelFormat(DXGI_FORMAT.DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE.D2D1_ALPHA_MODE_IGNORE)
                Dim nDPI As UInteger = GetDpiForWindow(Me.Handle)
                bitmapProperties.dpiX = nDPI
                bitmapProperties.dpiY = nDPI
                hr = m_pD2DDeviceContext.CreateBitmap(sizeU, pArray, nPitch, bitmapProperties, m_pD2DBitmapBackgroundMask)
                Marshal.FreeHGlobal(pArray)
            End If
            If (m_pD2DBitmapBackground Is Nothing And m_pD2DBitmap Is Nothing And m_pD2DBitmap1 Is Nothing And m_pD2DBitmap2 Is Nothing) Then
                MessageBox.Show("Could not load bitmaps", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf (m_pD2DBitmap Is Nothing) Then
                MessageBox.Show("Could not load Sprite Sheet 1", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf (m_pD2DBitmap1 Is Nothing) Then
                MessageBox.Show("Could not load Sprite Sheet 2", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf (m_pD2DBitmap2 Is Nothing) Then
                MessageBox.Show("Could not load Sprite Sheet 3", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
        Return hr
    End Function

    Private Function CreateSwapChain(ByVal hWnd As IntPtr) As HRESULT
        Dim hr As HRESULT = HRESULT.S_OK
        Dim swapChainDesc As DXGI_SWAP_CHAIN_DESC1 = New DXGI_SWAP_CHAIN_DESC1()
        swapChainDesc.Width = 1
        swapChainDesc.Height = 1
        swapChainDesc.Format = DXGI.DXGI_FORMAT.DXGI_FORMAT_B8G8R8A8_UNORM ' this is the most common swapchain format
        swapChainDesc.Stereo = False
        swapChainDesc.SampleDesc.Count = 1                ' don't use multi-sampling
        swapChainDesc.SampleDesc.Quality = 0
        swapChainDesc.BufferUsage = DXGI_USAGE_RENDER_TARGET_OUTPUT
        swapChainDesc.BufferCount = 2                     ' use double buffering to enable flip
        swapChainDesc.Scaling = If(hWnd <> IntPtr.Zero, DXGI_SCALING.DXGI_SCALING_NONE, DXGI_SCALING.DXGI_SCALING_STRETCH)
        swapChainDesc.SwapEffect = DXGI_SWAP_EFFECT.DXGI_SWAP_EFFECT_FLIP_SEQUENTIAL ' all apps must use this SwapEffect       
        swapChainDesc.Flags = 0

        Dim pDXGIAdapter As IDXGIAdapter = Nothing
        hr = m_pDXGIDevice.GetAdapter(pDXGIAdapter)
        If SUCCEEDED(hr) Then
            Dim pDXGIFactory2Ptr As IntPtr
            hr = pDXGIAdapter.GetParent(GetType(IDXGIFactory2).GUID, pDXGIFactory2Ptr)
            If SUCCEEDED(hr) Then
                Dim pDXGIFactory2 As IDXGIFactory2 = TryCast(Marshal.GetObjectForIUnknown(pDXGIFactory2Ptr), IDXGIFactory2)
                If hWnd <> IntPtr.Zero Then
                    hr = pDXGIFactory2.CreateSwapChainForHwnd(m_pD3D11DevicePtr, hWnd, swapChainDesc, IntPtr.Zero, Nothing, m_pDXGISwapChain1)
                Else
                    ' For Composition SwapChain
                    swapChainDesc.AlphaMode = DXGI_ALPHA_MODE.DXGI_ALPHA_MODE_PREMULTIPLIED
                    hr = pDXGIFactory2.CreateSwapChainForComposition(m_pD3D11DevicePtr, swapChainDesc, Nothing, m_pDXGISwapChain1)
                End If
                hr = m_pDXGIDevice.SetMaximumFrameLatency(1)
                SafeRelease(pDXGIFactory2)
                Marshal.Release(pDXGIFactory2Ptr)
            End If
            SafeRelease(pDXGIAdapter)
        End If
        Return hr
    End Function

    Private Function ConfigureSwapChain() As HRESULT
        Dim hr As HRESULT = HRESULT.S_OK

        Dim bitmapProperties As D2D1_BITMAP_PROPERTIES1 = New D2D1_BITMAP_PROPERTIES1()
        bitmapProperties.bitmapOptions = D2D1_BITMAP_OPTIONS.D2D1_BITMAP_OPTIONS_TARGET Or D2D1_BITMAP_OPTIONS.D2D1_BITMAP_OPTIONS_CANNOT_DRAW
        bitmapProperties.pixelFormat = D2DTools.PixelFormat(DXGI_FORMAT.DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE.D2D1_ALPHA_MODE_IGNORE)

        Dim nDPI As UInteger = GetDpiForWindow(Me.Handle)
        bitmapProperties.dpiX = nDPI
        bitmapProperties.dpiY = nDPI

        Dim pDXGISurfacePtr = IntPtr.Zero
        hr = m_pDXGISwapChain1.GetBuffer(0, GetType(IDXGISurface).GUID, pDXGISurfacePtr)
        If SUCCEEDED(hr) Then
            Dim pDXGISurface As IDXGISurface = TryCast(Marshal.GetObjectForIUnknown(pDXGISurfacePtr), IDXGISurface)
            hr = m_pD2DDeviceContext.CreateBitmapFromDxgiSurface(pDXGISurface, bitmapProperties, m_pD2DTargetBitmap)
            If SUCCEEDED(hr) Then
                m_pD2DDeviceContext.SetTarget(m_pD2DTargetBitmap)
            End If
            SafeRelease(pDXGISurface)
            Marshal.Release(pDXGISurfacePtr)
        End If
        Return hr
    End Function

    Private Function CreateD2DBitmapFromURL(ByVal sURL As String, <Out> ByRef pD2DBitmap As ID2D1Bitmap) As HRESULT
        Dim hr As HRESULT = HRESULT.S_OK
        pD2DBitmap = Nothing
        Dim bytes As Byte() = Nothing
        Try
            Net.ServicePointManager.Expect100Continue = True
            Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.Tls12
            Dim webRequest As Net.HttpWebRequest = CType(Net.HttpWebRequest.Create(sURL), Net.HttpWebRequest)
            webRequest.AllowWriteStreamBuffering = True
            Using webResponse As Net.WebResponse = webRequest.GetResponse()
                Dim stream As IO.Stream = webResponse.GetResponseStream()
                Using ms As IO.MemoryStream = New IO.MemoryStream()
                    stream.CopyTo(ms)
                    bytes = ms.ToArray()
                End Using
            End Using
        Catch ex As Exception
            Return HRESULT.E_FAIL
        End Try

        Dim wicStream As IWICStream = Nothing
        hr = CType(m_pWICImagingFactory.CreateStream(wicStream), HRESULT)
        If SUCCEEDED(hr) Then
            hr = CType(wicStream.InitializeFromMemory(bytes, bytes.Length), HRESULT)
            If SUCCEEDED(hr) Then
                Dim pDecoder As IWICBitmapDecoder = Nothing
                hr = CType(m_pWICImagingFactory.CreateDecoderFromStream(wicStream, Guid.Empty, WICDecodeOptions.WICDecodeMetadataCacheOnDemand, pDecoder), HRESULT)
                If SUCCEEDED(hr) Then
                    Dim pFrame As IWICBitmapFrameDecode = Nothing
                    hr = CType(pDecoder.GetFrame(0, pFrame), HRESULT)
                    If SUCCEEDED(hr) Then
                        Dim pConvertedSourceBitmap As IWICFormatConverter = Nothing
                        hr = CType(m_pWICImagingFactory.CreateFormatConverter(pConvertedSourceBitmap), HRESULT)
                        If SUCCEEDED(hr) Then
                            hr = CType(pConvertedSourceBitmap.Initialize(CType(pFrame, IWICBitmapSource), WICTools.GUID_WICPixelFormat32bppPBGRA, WICBitmapDitherType.WICBitmapDitherTypeNone, Nothing, 0, WICBitmapPaletteType.WICBitmapPaletteTypeCustom), HRESULT)
                            If SUCCEEDED(hr) Then
                                Dim bitmapproperties As D2D1_BITMAP_PROPERTIES = New D2D1_BITMAP_PROPERTIES()
                                hr = m_pD2DDeviceContext.CreateBitmapFromWicBitmap(pConvertedSourceBitmap, bitmapproperties, pD2DBitmap)
                            End If
                            SafeRelease(pConvertedSourceBitmap)
                        End If
                        SafeRelease(pFrame)
                    End If
                    SafeRelease(pDecoder)
                End If
            End If
        End If
        Return hr
    End Function

    Private Sub AddSprite(nSprite As Integer, pBitmap As ID2D1Bitmap, nXSprite As Integer, nYSprite As Integer, nCountSprite As Integer)
        If pBitmap IsNot Nothing Then
            Dim size As D2D1_SIZE_F
            m_pD2DDeviceContext.GetSize(size)
            Dim nClientWidth = size.width
            Dim nClientHeight = size.height

            Dim nScale As Single = rand.NextDouble() * 1
            Dim scale As D2D1_MATRIX_3X2_F = New D2D1_MATRIX_3X2_F()
            scale._11 = nScale
            scale._22 = nScale
            Dim colors As Array = ColorF.Enum.GetValues(GetType(ColorF.Enum))
            Dim randomColor As ColorF.Enum
            randomColor = CType(colors.GetValue(randColor.[Next](colors.Length)), ColorF.Enum)
            Dim color = New ColorF(randomColor)
            'color.r *= 1.5
            'color.g *= 1.1
            'color.b /= 1.25
            Dim s = New CSprite(m_pD2DDeviceContext3, pBitmap, nXSprite, nYSprite, nCountSprite, rand.NextDouble() * 5, rand.NextDouble() * 5, color, scale)
            If (nSprite = 1) Then
                CSprites.Add(s)
            ElseIf (nSprite = 3) Then
                CSprites3.Add(s)
            End If

            Dim bmpSize As D2D1_SIZE_F
            pBitmap.GetSize(bmpSize)
            Dim nWidth As Single = bmpSize.width / nXSprite
            Dim nHeight As Single = bmpSize.width / nYSprite
            If scale._11 <> 0 Then
                nClientWidth *= 1 / scale._11
            End If
            If scale._22 <> 0 Then
                nClientHeight *= 1 / scale._22
            End If

            Dim nX As Single = rand.NextDouble() * nClientWidth
            Dim nY As Single = rand.NextDouble() * nClientHeight
            If nX + nWidth >= nClientWidth Then nX = nClientWidth - nWidth
            If nX <= 0 Then nX = 0
            If nY + nHeight >= nClientHeight Then nY = nClientHeight - nHeight
            If nY <= 0 Then nY = 0
            s.X = nX
            s.Y = nY
        End If
    End Sub

    Private Sub AddBubble(pBitmap As ID2D1Bitmap, nXSprite As Integer, nYSprite As Integer, nCountSprite As Integer)
        If pBitmap IsNot Nothing Then
            Dim size As D2D1_SIZE_F
            m_pD2DDeviceContext.GetSize(size)
            Dim nClientWidth = size.width
            Dim nClientHeight = size.height

            Dim nScale As Single = rand.NextDouble() * Math.Abs(1 - 0.5) + 0.5
            Dim scale As D2D1_MATRIX_3X2_F = New D2D1_MATRIX_3X2_F()
            scale._11 = nScale
            scale._22 = nScale
            Dim colors As Array = ColorF.Enum.GetValues(GetType(ColorF.Enum))
            Dim randomColor As ColorF.Enum
            randomColor = CType(colors.GetValue(randColor.[Next](colors.Length)), ColorF.Enum)
            Dim color = New ColorF(randomColor)
            'color.r *= 1.5
            'color.g *= 1.1
            'color.b /= 1.25
            Dim s = New CSprite(m_pD2DDeviceContext3, pBitmap, nXSprite, nYSprite, nCountSprite, rand.NextDouble() * 5, rand.NextDouble() * 5, color, scale)
            CSprites2.Add(s)

            Dim bmpSize As D2D1_SIZE_F
            pBitmap.GetSize(bmpSize)
            Dim nWidth As Single = bmpSize.width / nXSprite
            Dim nHeight As Single = bmpSize.width / nYSprite
            If scale._11 <> 0 Then
                nClientWidth *= 1 / scale._11
            End If
            If scale._22 <> 0 Then
                nClientHeight *= 1 / scale._22
            End If

            Dim nX As Single = rand.NextDouble() * nClientWidth
            'Dim nY As Single = rand.NextDouble() * nClientHeight
            Dim nY As Single = nClientHeight + nHeight
            If nX + nWidth >= nClientWidth Then nX = nClientWidth - nWidth
            If nX <= 0 Then nX = 0
            'If nY + nHeight >= nClientHeight Then nY = nClientHeight - nHeight
            If nY <= 0 Then nY = 0
            s.X = nX
            s.Y = nY
        End If
    End Sub

    Private Sub CleanDeviceResources()
        SafeRelease(m_pD2DBitmap)
        SafeRelease(m_pD2DBitmap1)
        SafeRelease(m_pD2DBitmap2)
        SafeRelease(m_pD2DBitmapBackground)
        SafeRelease(m_pD2DBitmapBackgroundMask)
        SafeRelease(m_pMainBrush)

        For Each s As CSprite In CSprites
            s.Dispose()
        Next
        For Each s As CSprite In CSprites2
            s.Dispose()
        Next
        For Each s As CSprite In CSprites3
            s.Dispose()
        Next
    End Sub

    Private Sub Clean()
        SafeRelease(m_pD2DDeviceContext)
        SafeRelease(m_pD2DDeviceContext3)

        CleanDeviceResources()

        SafeRelease(m_pD2DTargetBitmap)
        SafeRelease(m_pDXGISwapChain1)
        SafeRelease(m_pDXGIDevice)
        SafeRelease(m_pWICImagingFactory)
        SafeRelease(m_pD2DFactory1)
        SafeRelease(m_pD2DFactory)
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Clean()
    End Sub
End Class
