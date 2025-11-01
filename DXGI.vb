Imports System

Imports System.Runtime.InteropServices
Imports GlobalStructures

Namespace Global.DXGI
    <ComImport>
    <Guid("aec22fb8-76f3-4639-9be0-28eb43a67a2e")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDXGIObject
        Function SetPrivateData(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        Function SetPrivateDataInterface(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, pUnknown As IntPtr) As HRESULT
        Function GetPrivateData(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        Function GetParent(
<MarshalAs(UnmanagedType.LPStruct)> riid As Guid, <Out> ByRef ppParent As IntPtr) As HRESULT
    End Interface

    <ComImport>
    <Guid("3d3e0379-f9de-4d58-bb6c-18d62992f1a6")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDXGIDeviceSubObject
        Inherits IDXGIObject
#Region "<IDXGIObject>"
        Overloads Function SetPrivateData(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        Overloads Function SetPrivateDataInterface(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, pUnknown As IntPtr) As HRESULT
        Overloads Function GetPrivateData(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        Overloads Function GetParent(
<MarshalAs(UnmanagedType.LPStruct)> riid As Guid, <Out> ByRef ppParent As IntPtr) As HRESULT
#End Region

        Function GetDevice(
<MarshalAs(UnmanagedType.LPStruct)> riid As Guid, <Out> ByRef ppDevice As IntPtr) As HRESULT
    End Interface

    <ComImport>
    <Guid("2411e7e1-12ac-4ccf-bd14-9798e8534dc0")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDXGIAdapter
        Inherits IDXGIObject
#Region "IDXGIObject"
        Overloads Function SetPrivateData(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        Overloads Function SetPrivateDataInterface(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, pUnknown As IntPtr) As HRESULT
        Overloads Function GetPrivateData(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        Overloads Function GetParent(
<MarshalAs(UnmanagedType.LPStruct)> riid As Guid, <Out> ByRef ppParent As IntPtr) As HRESULT
#End Region

        <PreserveSig>
        Function EnumOutputs(Output As UInteger, ByRef ppOutput As IDXGIOutput) As HRESULT
        Function GetDesc(<Out> ByRef pDesc As DXGI_ADAPTER_DESC) As HRESULT
        Function CheckInterfaceSupport(
<MarshalAs(UnmanagedType.LPStruct)> InterfaceName As Guid, <Out> ByRef pUMDVersion As LARGE_INTEGER) As HRESULT
    End Interface

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Public Structure DXGI_ADAPTER_DESC
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=128)>
        Public Description As String
        Public VendorId As UInteger
        Public DeviceId As UInteger
        Public SubSysId As UInteger
        Public Revision As UInteger
        Public DedicatedVideoMemory As UInteger
        Public DedicatedSystemMemory As UInteger
        Public SharedSystemMemory As UInteger
        Public AdapterLuid As LUID
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure LUID
        Private LowPart As UInteger
        Private HighPart As Integer
    End Structure

    <ComImport>
    <Guid("ae02eedb-c735-4690-8d52-5a8dc20213aa")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDXGIOutput
        Inherits IDXGIObject
#Region "IDXGIObject"
        Overloads Function SetPrivateData(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        Overloads Function SetPrivateDataInterface(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, pUnknown As IntPtr) As HRESULT
        Overloads Function GetPrivateData(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        Overloads Function GetParent(
<MarshalAs(UnmanagedType.LPStruct)> riid As Guid, <Out> ByRef ppParent As IntPtr) As HRESULT
#End Region

        Function GetDesc(<Out> ByRef pDesc As DXGI_OUTPUT_DESC) As HRESULT
        Function GetDisplayModeList(EnumFormat As DXGI_FORMAT, Flags As UInteger, ByRef pNumModes As UInteger, pDesc As DXGI_MODE_DESC) As HRESULT
        'HRESULT FindClosestMatchingMode(DXGI_MODE_DESC pModeToMatch, out  DXGI_MODE_DESC pClosestMatch, IUnknown pConcernedDevice);
        Function FindClosestMatchingMode(pModeToMatch As DXGI_MODE_DESC, <Out> ByRef pClosestMatch As DXGI_MODE_DESC, pConcernedDevice As IntPtr) As HRESULT
        Function WaitForVBlank() As HRESULT
        'HRESULT TakeOwnership(IUnknown pDevice, bool Exclusive);
        Function TakeOwnership(pDevice As IntPtr, Exclusive As Boolean) As HRESULT
        Sub ReleaseOwnership()
        Function GetGammaControlCapabilities(<Out> ByRef pGammaCaps As DXGI_GAMMA_CONTROL_CAPABILITIES) As HRESULT
        Function SetGammaControl(pArray As DXGI_GAMMA_CONTROL) As HRESULT
        Function GetGammaControl(<Out> ByRef pArray As DXGI_GAMMA_CONTROL) As HRESULT
        Function SetDisplaySurface(pScanoutSurface As IDXGISurface) As HRESULT
        Function GetDisplaySurfaceData(pDestination As IDXGISurface) As HRESULT
        Function GetFrameStatistics(<Out> ByRef pStats As DXGI_FRAME_STATISTICS) As HRESULT
    End Interface

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Public Structure DXGI_OUTPUT_DESC
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=32)>
        Private DeviceName As String
        Public DesktopCoordinates As RECT
        Public AttachedToDesktop As Boolean
        Public Rotation As DXGI_MODE_ROTATION
        Public Monitor As IntPtr
    End Structure

    Public Enum DXGI_MODE_ROTATION
        DXGI_MODE_ROTATION_UNSPECIFIED = 0
        DXGI_MODE_ROTATION_IDENTITY = 1
        DXGI_MODE_ROTATION_ROTATE90 = 2
        DXGI_MODE_ROTATION_ROTATE180 = 3
        DXGI_MODE_ROTATION_ROTATE270 = 4
    End Enum

    Public Enum DXGI_FORMAT
        DXGI_FORMAT_UNKNOWN = 0
        DXGI_FORMAT_R32G32B32A32_TYPELESS = 1
        DXGI_FORMAT_R32G32B32A32_FLOAT = 2
        DXGI_FORMAT_R32G32B32A32_UINT = 3
        DXGI_FORMAT_R32G32B32A32_SINT = 4
        DXGI_FORMAT_R32G32B32_TYPELESS = 5
        DXGI_FORMAT_R32G32B32_FLOAT = 6
        DXGI_FORMAT_R32G32B32_UINT = 7
        DXGI_FORMAT_R32G32B32_SINT = 8
        DXGI_FORMAT_R16G16B16A16_TYPELESS = 9
        DXGI_FORMAT_R16G16B16A16_FLOAT = 10
        DXGI_FORMAT_R16G16B16A16_UNORM = 11
        DXGI_FORMAT_R16G16B16A16_UINT = 12
        DXGI_FORMAT_R16G16B16A16_SNORM = 13
        DXGI_FORMAT_R16G16B16A16_SINT = 14
        DXGI_FORMAT_R32G32_TYPELESS = 15
        DXGI_FORMAT_R32G32_FLOAT = 16
        DXGI_FORMAT_R32G32_UINT = 17
        DXGI_FORMAT_R32G32_SINT = 18
        DXGI_FORMAT_R32G8X24_TYPELESS = 19
        DXGI_FORMAT_D32_FLOAT_S8X24_UINT = 20
        DXGI_FORMAT_R32_FLOAT_X8X24_TYPELESS = 21
        DXGI_FORMAT_X32_TYPELESS_G8X24_UINT = 22
        DXGI_FORMAT_R10G10B10A2_TYPELESS = 23
        DXGI_FORMAT_R10G10B10A2_UNORM = 24
        DXGI_FORMAT_R10G10B10A2_UINT = 25
        DXGI_FORMAT_R11G11B10_FLOAT = 26
        DXGI_FORMAT_R8G8B8A8_TYPELESS = 27
        DXGI_FORMAT_R8G8B8A8_UNORM = 28
        DXGI_FORMAT_R8G8B8A8_UNORM_SRGB = 29
        DXGI_FORMAT_R8G8B8A8_UINT = 30
        DXGI_FORMAT_R8G8B8A8_SNORM = 31
        DXGI_FORMAT_R8G8B8A8_SINT = 32
        DXGI_FORMAT_R16G16_TYPELESS = 33
        DXGI_FORMAT_R16G16_FLOAT = 34
        DXGI_FORMAT_R16G16_UNORM = 35
        DXGI_FORMAT_R16G16_UINT = 36
        DXGI_FORMAT_R16G16_SNORM = 37
        DXGI_FORMAT_R16G16_SINT = 38
        DXGI_FORMAT_R32_TYPELESS = 39
        DXGI_FORMAT_D32_FLOAT = 40
        DXGI_FORMAT_R32_FLOAT = 41
        DXGI_FORMAT_R32_UINT = 42
        DXGI_FORMAT_R32_SINT = 43
        DXGI_FORMAT_R24G8_TYPELESS = 44
        DXGI_FORMAT_D24_UNORM_S8_UINT = 45
        DXGI_FORMAT_R24_UNORM_X8_TYPELESS = 46
        DXGI_FORMAT_X24_TYPELESS_G8_UINT = 47
        DXGI_FORMAT_R8G8_TYPELESS = 48
        DXGI_FORMAT_R8G8_UNORM = 49
        DXGI_FORMAT_R8G8_UINT = 50
        DXGI_FORMAT_R8G8_SNORM = 51
        DXGI_FORMAT_R8G8_SINT = 52
        DXGI_FORMAT_R16_TYPELESS = 53
        DXGI_FORMAT_R16_FLOAT = 54
        DXGI_FORMAT_D16_UNORM = 55
        DXGI_FORMAT_R16_UNORM = 56
        DXGI_FORMAT_R16_UINT = 57
        DXGI_FORMAT_R16_SNORM = 58
        DXGI_FORMAT_R16_SINT = 59
        DXGI_FORMAT_R8_TYPELESS = 60
        DXGI_FORMAT_R8_UNORM = 61
        DXGI_FORMAT_R8_UINT = 62
        DXGI_FORMAT_R8_SNORM = 63
        DXGI_FORMAT_R8_SINT = 64
        DXGI_FORMAT_A8_UNORM = 65
        DXGI_FORMAT_R1_UNORM = 66
        DXGI_FORMAT_R9G9B9E5_SHAREDEXP = 67
        DXGI_FORMAT_R8G8_B8G8_UNORM = 68
        DXGI_FORMAT_G8R8_G8B8_UNORM = 69
        DXGI_FORMAT_BC1_TYPELESS = 70
        DXGI_FORMAT_BC1_UNORM = 71
        DXGI_FORMAT_BC1_UNORM_SRGB = 72
        DXGI_FORMAT_BC2_TYPELESS = 73
        DXGI_FORMAT_BC2_UNORM = 74
        DXGI_FORMAT_BC2_UNORM_SRGB = 75
        DXGI_FORMAT_BC3_TYPELESS = 76
        DXGI_FORMAT_BC3_UNORM = 77
        DXGI_FORMAT_BC3_UNORM_SRGB = 78
        DXGI_FORMAT_BC4_TYPELESS = 79
        DXGI_FORMAT_BC4_UNORM = 80
        DXGI_FORMAT_BC4_SNORM = 81
        DXGI_FORMAT_BC5_TYPELESS = 82
        DXGI_FORMAT_BC5_UNORM = 83
        DXGI_FORMAT_BC5_SNORM = 84
        DXGI_FORMAT_B5G6R5_UNORM = 85
        DXGI_FORMAT_B5G5R5A1_UNORM = 86
        DXGI_FORMAT_B8G8R8A8_UNORM = 87
        DXGI_FORMAT_B8G8R8X8_UNORM = 88
        DXGI_FORMAT_R10G10B10_XR_BIAS_A2_UNORM = 89
        DXGI_FORMAT_B8G8R8A8_TYPELESS = 90
        DXGI_FORMAT_B8G8R8A8_UNORM_SRGB = 91
        DXGI_FORMAT_B8G8R8X8_TYPELESS = 92
        DXGI_FORMAT_B8G8R8X8_UNORM_SRGB = 93
        DXGI_FORMAT_BC6H_TYPELESS = 94
        DXGI_FORMAT_BC6H_UF16 = 95
        DXGI_FORMAT_BC6H_SF16 = 96
        DXGI_FORMAT_BC7_TYPELESS = 97
        DXGI_FORMAT_BC7_UNORM = 98
        DXGI_FORMAT_BC7_UNORM_SRGB = 99
        DXGI_FORMAT_AYUV = 100
        DXGI_FORMAT_Y410 = 101
        DXGI_FORMAT_Y416 = 102
        DXGI_FORMAT_NV12 = 103
        DXGI_FORMAT_P010 = 104
        DXGI_FORMAT_P016 = 105
        DXGI_FORMAT_420_OPAQUE = 106
        DXGI_FORMAT_YUY2 = 107
        DXGI_FORMAT_Y210 = 108
        DXGI_FORMAT_Y216 = 109
        DXGI_FORMAT_NV11 = 110
        DXGI_FORMAT_AI44 = 111
        DXGI_FORMAT_IA44 = 112
        DXGI_FORMAT_P8 = 113
        DXGI_FORMAT_A8P8 = 114
        DXGI_FORMAT_B4G4R4A4_UNORM = 115
        DXGI_FORMAT_FORCE_UINT = &HFFFFFFFF
    End Enum

    Public Enum DXGI_COLOR_SPACE_TYPE
        DXGI_COLOR_SPACE_RGB_FULL_G22_NONE_P709 = 0
        DXGI_COLOR_SPACE_RGB_FULL_G10_NONE_P709 = 1
        DXGI_COLOR_SPACE_RGB_STUDIO_G22_NONE_P709 = 2
        DXGI_COLOR_SPACE_RGB_STUDIO_G22_NONE_P2020 = 3
        DXGI_COLOR_SPACE_RESERVED = 4
        DXGI_COLOR_SPACE_YCBCR_FULL_G22_NONE_P709_X601 = 5
        DXGI_COLOR_SPACE_YCBCR_STUDIO_G22_LEFT_P601 = 6
        DXGI_COLOR_SPACE_YCBCR_FULL_G22_LEFT_P601 = 7
        DXGI_COLOR_SPACE_YCBCR_STUDIO_G22_LEFT_P709 = 8
        DXGI_COLOR_SPACE_YCBCR_FULL_G22_LEFT_P709 = 9
        DXGI_COLOR_SPACE_YCBCR_STUDIO_G22_LEFT_P2020 = 10
        DXGI_COLOR_SPACE_YCBCR_FULL_G22_LEFT_P2020 = 11
        DXGI_COLOR_SPACE_RGB_FULL_G2084_NONE_P2020 = 12
        DXGI_COLOR_SPACE_YCBCR_STUDIO_G2084_LEFT_P2020 = 13
        DXGI_COLOR_SPACE_RGB_STUDIO_G2084_NONE_P2020 = 14
        DXGI_COLOR_SPACE_YCBCR_STUDIO_G22_TOPLEFT_P2020 = 15
        DXGI_COLOR_SPACE_YCBCR_STUDIO_G2084_TOPLEFT_P2020 = 16
        DXGI_COLOR_SPACE_RGB_FULL_G22_NONE_P2020 = 17
        DXGI_COLOR_SPACE_YCBCR_STUDIO_GHLG_TOPLEFT_P2020 = 18
        DXGI_COLOR_SPACE_YCBCR_FULL_GHLG_TOPLEFT_P2020 = 19
        DXGI_COLOR_SPACE_RGB_STUDIO_G24_NONE_P709 = 20
        DXGI_COLOR_SPACE_RGB_STUDIO_G24_NONE_P2020 = 21
        DXGI_COLOR_SPACE_YCBCR_STUDIO_G24_LEFT_P709 = 22
        DXGI_COLOR_SPACE_YCBCR_STUDIO_G24_LEFT_P2020 = 23
        DXGI_COLOR_SPACE_YCBCR_STUDIO_G24_TOPLEFT_P2020 = 24
        DXGI_COLOR_SPACE_CUSTOM = &HFFFFFFFF
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DXGI_MODE_DESC
        Public Width As UInteger
        Public Height As UInteger
        Public RefreshRate As DXGI_RATIONAL
        Public Format As DXGI_FORMAT
        Public ScanlineOrdering As DXGI_MODE_SCANLINE_ORDER
        Public Scaling As DXGI_MODE_SCALING
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DXGI_RATIONAL
        Public Numerator As UInteger
        Public Denominator As UInteger
    End Structure

    Public Enum DXGI_MODE_SCANLINE_ORDER
        DXGI_MODE_SCANLINE_ORDER_UNSPECIFIED = 0
        DXGI_MODE_SCANLINE_ORDER_PROGRESSIVE = 1
        DXGI_MODE_SCANLINE_ORDER_UPPER_FIELD_FIRST = 2
        DXGI_MODE_SCANLINE_ORDER_LOWER_FIELD_FIRST = 3
    End Enum

    Public Enum DXGI_MODE_SCALING
        DXGI_MODE_SCALING_UNSPECIFIED = 0
        DXGI_MODE_SCALING_CENTERED = 1
        DXGI_MODE_SCALING_STRETCHED = 2
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DXGI_GAMMA_CONTROL_CAPABILITIES
        Public ScaleAndOffsetSupported As Boolean
        Public MaxConvertedValue As Single
        Public MinConvertedValue As Single
        Public NumGammaControlPoints As UInteger
        <MarshalAs(UnmanagedType.R4, SizeConst:=1025)>
        Private ControlPointPositions As Single
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DXGI_GAMMA_CONTROL
        Public Scale As DXGI_RGB
        Public Offset As DXGI_RGB
        <MarshalAs(UnmanagedType.Struct, SizeConst:=1025)>
        Private GammaCurve As DXGI_RGB
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DXGI_RGB
        Public Red As Single
        Public Green As Single
        Public Blue As Single
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DXGI_FRAME_STATISTICS
        Public PresentCount As UInteger
        Public PresentRefreshCount As UInteger
        Public SyncRefreshCount As UInteger
        Public SyncQPCTime As LARGE_INTEGER
        Public SyncGPUTime As LARGE_INTEGER
    End Structure

    <ComImport>
    <Guid("cafcb56c-6ac3-4889-bf47-9e23bbd260ec")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDXGISurface
        Inherits IDXGIDeviceSubObject
#Region "<IDXGIDeviceSubObject>"

#Region "<IDXGIObject>"
        Overloads Function SetPrivateData(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        Overloads Function SetPrivateDataInterface(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, pUnknown As IntPtr) As HRESULT
        Overloads Function GetPrivateData(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        Overloads Function GetParent(
<MarshalAs(UnmanagedType.LPStruct)> riid As Guid, <Out> ByRef ppParent As IntPtr) As HRESULT
#End Region

        Overloads Function GetDevice(
<MarshalAs(UnmanagedType.LPStruct)> riid As Guid, <Out> ByRef ppDevice As IntPtr) As HRESULT
#End Region

        Function GetDesc(<Out> ByRef pDesc As DXGI_SURFACE_DESC) As HRESULT
        <PreserveSig>
        Function Map(<Out> ByRef pLockedRect As DXGI_MAPPED_RECT, MapFlags As UInteger) As HRESULT
        <PreserveSig>
        Function Unmap() As HRESULT
    End Interface

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DXGI_SURFACE_DESC
        Public Width As UInteger
        Public Height As UInteger
        Public Format As DXGI_FORMAT
        Public SampleDesc As DXGI_SAMPLE_DESC
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DXGI_MAPPED_RECT
        Public Pitch As Integer
        Public pBits As IntPtr
    End Structure

    <ComImport>
    <Guid("7b7166ec-21c7-44ae-b21a-c9ae321ae369")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDXGIFactory
        Inherits IDXGIObject
#Region "<IDXGIObject>"
        Overloads Function SetPrivateData(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        Overloads Function SetPrivateDataInterface(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, pUnknown As IntPtr) As HRESULT
        Overloads Function GetPrivateData(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        Overloads Function GetParent(
<MarshalAs(UnmanagedType.LPStruct)> riid As Guid, <Out> ByRef ppParent As IntPtr) As HRESULT
#End Region

        Function EnumAdapters(Adapter As UInteger, <Out> ByRef ppAdapter As IDXGIAdapter) As HRESULT
        Function MakeWindowAssociation(WindowHandle As IntPtr, Flags As UInteger) As HRESULT
        Function GetWindowAssociation(<Out> ByRef pWindowHandle As IntPtr) As HRESULT
        <PreserveSig>
        Function CreateSwapChain(pDevice As IntPtr, pDesc As DXGI_SWAP_CHAIN_DESC, <Out> ByRef ppSwapChain As IDXGISwapChain) As HRESULT
        Function CreateSoftwareAdapter([Module] As IntPtr, <Out> ByRef ppAdapter As IDXGIAdapter) As HRESULT
    End Interface

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DXGI_SWAP_CHAIN_DESC
        Public BufferDesc As DXGI_MODE_DESC
        Public SampleDesc As DXGI_SAMPLE_DESC
        Public BufferUsage As UInteger
        Public BufferCount As UInteger
        Public OutputWindow As IntPtr
        Public Windowed As Boolean
        Public SwapEffect As DXGI_SWAP_EFFECT
        Public Flags As UInteger
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DXGI_SAMPLE_DESC
        Public Count As UInteger
        Public Quality As UInteger
    End Structure

    Public Enum DXGI_SWAP_EFFECT
        DXGI_SWAP_EFFECT_DISCARD = 0
        DXGI_SWAP_EFFECT_SEQUENTIAL = 1
        DXGI_SWAP_EFFECT_FLIP_SEQUENTIAL = 3
        DXGI_SWAP_EFFECT_FLIP_DISCARD = 4
    End Enum

    <ComImport>
    <Guid("310d36a0-d2e7-4c0a-aa04-6a9d23b8886a")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDXGISwapChain
        Inherits IDXGIDeviceSubObject
#Region "IDXGIDeviceSubObject"
#Region "IDXGIObject"
        Overloads Function SetPrivateData(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        Overloads Function SetPrivateDataInterface(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, pUnknown As IntPtr) As HRESULT
        Overloads Function GetPrivateData(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        Overloads Function GetParent(
<MarshalAs(UnmanagedType.LPStruct)> riid As Guid, <Out> ByRef ppParent As IntPtr) As HRESULT
#End Region

        Overloads Function GetDevice(
<MarshalAs(UnmanagedType.LPStruct)> riid As Guid, <Out> ByRef ppDevice As IntPtr) As HRESULT
#End Region

        <PreserveSig>
        Function Present(SyncInterval As UInteger, Flags As UInteger) As HRESULT
        Function GetBuffer(Buffer As UInteger, ByRef riid As Guid, <Out> ByRef ppSurface As IntPtr) As HRESULT
        Function SetFullscreenState(Fullscreen As Boolean, pTarget As IDXGIOutput) As HRESULT
        Function GetFullscreenState(<Out> ByRef pFullscreen As Boolean, <Out> ByRef ppTarget As IDXGIOutput) As HRESULT
        Function GetDesc(<Out> ByRef pDesc As DXGI_SWAP_CHAIN_DESC) As HRESULT
        Function ResizeBuffers(BufferCount As UInteger, Width As UInteger, Height As UInteger, NewFormat As DXGI_FORMAT, SwapChainFlags As UInteger) As HRESULT
        Function ResizeTarget(pNewTargetParameters As DXGI_MODE_DESC) As HRESULT
        Function GetContainingOutput(<Out> ByRef ppOutput As IDXGIOutput) As HRESULT
        Function GetFrameStatistics(<Out> ByRef pStats As DXGI_FRAME_STATISTICS) As HRESULT
        Function GetLastPresentCount(<Out> ByRef pLastPresentCount As UInteger) As HRESULT
    End Interface

    <ComImport>
    <Guid("770aae78-f26f-4dba-a829-253c83d1b387")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDXGIFactory1
        Inherits IDXGIFactory
#Region "<IDXGIObject>"
        Overloads Function SetPrivateData(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        Overloads Function SetPrivateDataInterface(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, pUnknown As IntPtr) As HRESULT
        Overloads Function GetPrivateData(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        Overloads Function GetParent(
<MarshalAs(UnmanagedType.LPStruct)> riid As Guid, <Out> ByRef ppParent As IntPtr) As HRESULT
#End Region

#Region "<IDXGIFactory>"
        Overloads Function EnumAdapters(Adapter As UInteger, <Out> ByRef ppAdapter As IDXGIAdapter) As HRESULT
        Overloads Function MakeWindowAssociation(WindowHandle As IntPtr, Flags As UInteger) As HRESULT
        Overloads Function GetWindowAssociation(<Out> ByRef pWindowHandle As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function CreateSwapChain(pDevice As IntPtr, pDesc As DXGI_SWAP_CHAIN_DESC, <Out> ByRef ppSwapChain As IDXGISwapChain) As HRESULT
        Overloads Function CreateSoftwareAdapter([Module] As IntPtr, <Out> ByRef ppAdapter As IDXGIAdapter) As HRESULT
#End Region

        Function EnumAdapters1(Adapter As UInteger, <Out> ByRef ppAdapter As IDXGIAdapter1) As HRESULT
        Function IsCurrent() As Boolean
    End Interface

    <ComImport>
    <Guid("29038f61-3839-4626-91fd-086879011a05")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDXGIAdapter1
        Inherits IDXGIAdapter
#Region "IDXGIAdapter"
#Region "IDXGIObject"
        Overloads Function SetPrivateData(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        Overloads Function SetPrivateDataInterface(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, pUnknown As IntPtr) As HRESULT
        Overloads Function GetPrivateData(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        Overloads Function GetParent(
<MarshalAs(UnmanagedType.LPStruct)> riid As Guid, <Out> ByRef ppParent As IntPtr) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function EnumOutputs(Output As UInteger, ByRef ppOutput As IDXGIOutput) As HRESULT
        Overloads Function GetDesc(<Out> ByRef pDesc As DXGI_ADAPTER_DESC) As HRESULT
        Overloads Function CheckInterfaceSupport(
<MarshalAs(UnmanagedType.LPStruct)> InterfaceName As Guid, <Out> ByRef pUMDVersion As LARGE_INTEGER) As HRESULT
#End Region

        Function GetDesc1(pDesc As DXGI_ADAPTER_DESC1) As HRESULT
    End Interface

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DXGI_ADAPTER_DESC1
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=128)>
        Public Description As String
        Public VendorId As UInteger
        Public DeviceId As UInteger
        Public SubSysId As UInteger
        Public Revision As UInteger
        Public DedicatedVideoMemory As UInteger
        Public DedicatedSystemMemory As UInteger
        Public SharedSystemMemory As UInteger
        Public AdapterLuid As LUID
        Public Flags As UInteger
    End Structure

    <ComImport>
    <Guid("50c83a1c-e072-4c48-87b0-3630fa36a6d0")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDXGIFactory2
        Inherits IDXGIFactory1
#Region "<IDXGIFactory1>"
#Region "<IDXGIObject>"
        Overloads Function SetPrivateData(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        Overloads Function SetPrivateDataInterface(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, pUnknown As IntPtr) As HRESULT
        Overloads Function GetPrivateData(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        Overloads Function GetParent(
<MarshalAs(UnmanagedType.LPStruct)> riid As Guid, <Out> ByRef ppParent As IntPtr) As HRESULT
#End Region

#Region "<IDXGIFactory>"
        Overloads Function EnumAdapters(Adapter As UInteger, <Out> ByRef ppAdapter As IDXGIAdapter) As HRESULT
        Overloads Function MakeWindowAssociation(WindowHandle As IntPtr, Flags As UInteger) As HRESULT
        Overloads Function GetWindowAssociation(<Out> ByRef pWindowHandle As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function CreateSwapChain(pDevice As IntPtr, pDesc As DXGI_SWAP_CHAIN_DESC, <Out> ByRef ppSwapChain As IDXGISwapChain) As HRESULT
        Overloads Function CreateSoftwareAdapter([Module] As IntPtr, <Out> ByRef ppAdapter As IDXGIAdapter) As HRESULT
#End Region

        Overloads Function EnumAdapters1(Adapter As UInteger, <Out> ByRef ppAdapter As IDXGIAdapter1) As HRESULT
        Overloads Function IsCurrent() As Boolean
#End Region

        Function IsWindowedStereoEnabled() As Boolean
        'HRESULT CreateSwapChainForHwnd(IntPtr pDevice, IntPtr hWnd, DXGI_SWAP_CHAIN_DESC1 pDesc, DXGI_SWAP_CHAIN_FULLSCREEN_DESC pFullscreenDesc, IDXGIOutput pRestrictToOutput, out IDXGISwapChain1 ppSwapChain);
        <PreserveSig>
        Function CreateSwapChainForHwnd(pDevice As IntPtr, hWnd As IntPtr, ByRef pDesc As DXGI_SWAP_CHAIN_DESC1, pFullscreenDesc As IntPtr, pRestrictToOutput As IDXGIOutput, <Out> ByRef ppSwapChain As IDXGISwapChain1) As HRESULT
        <PreserveSig>
        Function CreateSwapChainForCoreWindow(pDevice As IntPtr, pWindow As IntPtr, ByRef pDesc As DXGI_SWAP_CHAIN_DESC1, pRestrictToOutput As IDXGIOutput, <Out> ByRef ppSwapChain As IDXGISwapChain1) As HRESULT
        Function GetSharedResourceAdapterLuid(hResource As IntPtr, <Out> ByRef pLuid As LUID) As HRESULT
        Function RegisterStereoStatusWindow(WindowHandle As IntPtr, wMsg As UInteger, <Out> ByRef pdwCookie As UInteger) As HRESULT
        Function RegisterStereoStatusEvent(hEvent As IntPtr, <Out> ByRef pdwCookie As UInteger) As HRESULT
        Sub UnregisterStereoStatus(dwCookie As UInteger)
        Function RegisterOcclusionStatusWindow(WindowHandle As IntPtr, wMsg As UInteger, <Out> ByRef pdwCookie As UInteger) As HRESULT
        Function RegisterOcclusionStatusEvent(hEvent As IntPtr, <Out> ByRef pdwCookie As UInteger) As HRESULT
        Sub UnregisterOcclusionStatus(dwCookie As UInteger)
        <PreserveSig>
        Function CreateSwapChainForComposition(pDevice As IntPtr, ByRef pDesc As DXGI_SWAP_CHAIN_DESC1, pRestrictToOutput As IDXGIOutput, <Out> ByRef ppSwapChain As IDXGISwapChain1) As HRESULT
    End Interface

    <ComImport>
    <Guid("790a45f7-0d42-4876-983a-0a55cfe6f4aa")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDXGISwapChain1
        Inherits IDXGISwapChain
#Region "IDXGISwapChain"
#Region "IDXGIDeviceSubObject"
#Region "IDXGIObject"
        Overloads Function SetPrivateData(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        Overloads Function SetPrivateDataInterface(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, pUnknown As IntPtr) As HRESULT
        Overloads Function GetPrivateData(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        Overloads Function GetParent(
<MarshalAs(UnmanagedType.LPStruct)> riid As Guid, <Out> ByRef ppParent As IntPtr) As HRESULT
#End Region

        Overloads Function GetDevice(
<MarshalAs(UnmanagedType.LPStruct)> riid As Guid, <Out> ByRef ppDevice As IntPtr) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function Present(SyncInterval As UInteger, Flags As UInteger) As HRESULT
        Overloads Function GetBuffer(Buffer As UInteger, ByRef riid As Guid, <Out> ByRef ppSurface As IntPtr) As HRESULT
        Overloads Function SetFullscreenState(Fullscreen As Boolean, pTarget As IDXGIOutput) As HRESULT
        Overloads Function GetFullscreenState(<Out> ByRef pFullscreen As Boolean, <Out> ByRef ppTarget As IDXGIOutput) As HRESULT
        Overloads Function GetDesc(<Out> ByRef pDesc As DXGI_SWAP_CHAIN_DESC) As HRESULT

        <PreserveSig>
        Overloads Function ResizeBuffers(BufferCount As UInteger, Width As UInteger, Height As UInteger, NewFormat As DXGI_FORMAT, SwapChainFlags As UInteger) As HRESULT
        Overloads Function ResizeTarget(pNewTargetParameters As DXGI_MODE_DESC) As HRESULT
        Overloads Function GetContainingOutput(<Out> ByRef ppOutput As IDXGIOutput) As HRESULT
        <PreserveSig>
        Overloads Function GetFrameStatistics(<Out> ByRef pStats As DXGI_FRAME_STATISTICS) As HRESULT
        Overloads Function GetLastPresentCount(<Out> ByRef pLastPresentCount As UInteger) As HRESULT
#End Region

        Function GetDesc1(<Out> ByRef pDesc As DXGI_SWAP_CHAIN_DESC1) As HRESULT
        Function GetFullscreenDesc(<Out> ByRef pDesc As DXGI_SWAP_CHAIN_FULLSCREEN_DESC) As HRESULT
        Function GetIntPtr(<Out> ByRef pIntPtr As IntPtr) As HRESULT
        Function GetCoreWindow(ByRef refiid As Guid, <Out> ByRef ppUnk As IntPtr) As HRESULT
        <PreserveSig>
        Function Present1(SyncInterval As UInteger, PresentFlags As UInteger, pPresentParameters As DXGI_PRESENT_PARAMETERS) As HRESULT
        Function IsTemporaryMonoSupported() As Boolean
        Function GetRestrictToOutput(<Out> ByRef ppRestrictToOutput As IDXGIOutput) As HRESULT
        Function SetBackgroundColor(pColor As DXGI_RGBA) As HRESULT
        Function GetBackgroundColor(<Out> ByRef pColor As DXGI_RGBA) As HRESULT
        Function SetRotation(Rotation As DXGI_MODE_ROTATION) As HRESULT
        Function GetRotation(<Out> ByRef pRotation As DXGI_MODE_ROTATION) As HRESULT
    End Interface

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DXGI_SWAP_CHAIN_DESC1
        Public Width As UInteger
        Public Height As UInteger
        Public Format As DXGI_FORMAT
        Public Stereo As Boolean
        Public SampleDesc As DXGI_SAMPLE_DESC
        Public BufferUsage As UInteger
        Public BufferCount As UInteger
        Public Scaling As DXGI_SCALING
        Public SwapEffect As DXGI_SWAP_EFFECT
        Public AlphaMode As DXGI_ALPHA_MODE
        Public Flags As DXGI_SWAP_CHAIN_FLAG
    End Structure

    Public Enum DXGI_SCALING
        DXGI_SCALING_STRETCH = 0
        DXGI_SCALING_NONE = 1
        DXGI_SCALING_ASPECT_RATIO_STRETCH = 2
    End Enum

    Public Enum DXGI_ALPHA_MODE As UInteger
        DXGI_ALPHA_MODE_UNSPECIFIED = 0
        DXGI_ALPHA_MODE_PREMULTIPLIED = 1
        DXGI_ALPHA_MODE_STRAIGHT = 2
        DXGI_ALPHA_MODE_IGNORE = 3
        DXGI_ALPHA_MODE_FORCE_DWORD = &HFFFFFFFFUI
    End Enum

    Public Enum DXGI_SWAP_CHAIN_FLAG
        DXGI_SWAP_CHAIN_FLAG_NONPREROTATED = 1
        DXGI_SWAP_CHAIN_FLAG_ALLOW_MODE_SWITCH = 2
        DXGI_SWAP_CHAIN_FLAG_GDI_COMPATIBLE = 4
        DXGI_SWAP_CHAIN_FLAG_RESTRICTED_CONTENT = 8
        DXGI_SWAP_CHAIN_FLAG_RESTRICT_SHARED_RESOURCE_DRIVER = 16
        DXGI_SWAP_CHAIN_FLAG_DISPLAY_ONLY = 32
        DXGI_SWAP_CHAIN_FLAG_FRAME_LATENCY_WAITABLE_OBJECT = 64
        DXGI_SWAP_CHAIN_FLAG_FOREGROUND_LAYER = 128
        DXGI_SWAP_CHAIN_FLAG_FULLSCREEN_VIDEO = 256
        DXGI_SWAP_CHAIN_FLAG_YUV_VIDEO = 512
        DXGI_SWAP_CHAIN_FLAG_HW_PROTECTED = 1024
        DXGI_SWAP_CHAIN_FLAG_ALLOW_TEARING = 2048
        DXGI_SWAP_CHAIN_FLAG_RESTRICTED_TO_ALL_HOLOGRAPHIC_DISPLAYS = 4096
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DXGI_SWAP_CHAIN_FULLSCREEN_DESC
        Public RefreshRate As DXGI_RATIONAL
        Public ScanlineOrdering As DXGI_MODE_SCANLINE_ORDER
        Public Scaling As DXGI_MODE_SCALING
        Public Windowed As Boolean
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DXGI_PRESENT_PARAMETERS
        Public DirtyRectsCount As UInteger
        Public pDirtyRects As RECT
        Public pScrollRect As RECT
        Public pScrollOffset As Point
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DXGI_RGBA
        Public r As Single
        Public g As Single
        Public b As Single
        Public a As Single
    End Structure

    <ComImport>
    <Guid("54ec77fa-1377-44e6-8c32-88fd5f44c84c")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDXGIDevice
        Inherits IDXGIObject
#Region "IDXGIObject"
        Overloads Function SetPrivateData(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        Overloads Function SetPrivateDataInterface(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, pUnknown As IntPtr) As HRESULT
        Overloads Function GetPrivateData(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        Overloads Function GetParent(
<MarshalAs(UnmanagedType.LPStruct)> riid As Guid, <Out> ByRef ppParent As IntPtr) As HRESULT
#End Region

        Function GetAdapter(<Out> ByRef pAdapter As IDXGIAdapter) As HRESULT
        Function CreateSurface(pDesc As DXGI_SURFACE_DESC, NumSurfaces As UInteger, Usage As UInteger, ByRef pSharedResource As DXGI_SHARED_RESOURCE, <Out> ByRef ppSurface As IDXGISurface) As HRESULT
        Function QueryResourceResidency(ppResources As IntPtr, <Out> ByRef pResidencyStatus As DXGI_RESIDENCY, NumResources As UInteger) As HRESULT
        Function SetGPUThreadPriority(Priority As Integer) As HRESULT
        Function GetGPUThreadPriority(<Out> ByRef pPriority As Integer) As HRESULT
    End Interface

    <ComImport>
    <Guid("77db970f-6276-48ba-ba28-070143b4392c")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDXGIDevice1
        Inherits IDXGIDevice
#Region "IDXGIDevice"
#Region "IDXGIObject"
        Overloads Function SetPrivateData(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        Overloads Function SetPrivateDataInterface(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, pUnknown As IntPtr) As HRESULT
        Overloads Function GetPrivateData(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        Overloads Function GetParent(
<MarshalAs(UnmanagedType.LPStruct)> riid As Guid, <Out> ByRef ppParent As IntPtr) As HRESULT
#End Region

        Overloads Function GetAdapter(<Out> ByRef pAdapter As IDXGIAdapter) As HRESULT
        Overloads Function CreateSurface(pDesc As DXGI_SURFACE_DESC, NumSurfaces As UInteger, Usage As UInteger, ByRef pSharedResource As DXGI_SHARED_RESOURCE, <Out> ByRef ppSurface As IDXGISurface) As HRESULT
        Overloads Function QueryResourceResidency(ppResources As IntPtr, <Out> ByRef pResidencyStatus As DXGI_RESIDENCY, NumResources As UInteger) As HRESULT
        Overloads Function SetGPUThreadPriority(Priority As Integer) As HRESULT
        Overloads Function GetGPUThreadPriority(<Out> ByRef pPriority As Integer) As HRESULT
#End Region

        Function SetMaximumFrameLatency(MaxLatency As UInteger) As HRESULT
        Function GetMaximumFrameLatency(<Out> ByRef pMaxLatency As UInteger) As HRESULT
    End Interface

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DXGI_SHARED_RESOURCE
        Public Handle As IntPtr
    End Structure

    Public Enum DXGI_RESIDENCY
        DXGI_RESIDENCY_FULLY_RESIDENT = 1
        DXGI_RESIDENCY_RESIDENT_IN_SHARED_MEMORY = 2
        DXGI_RESIDENCY_EVICTED_TO_DISK = 3
    End Enum

    <ComImport>
    <Guid("eb533d5d-2db6-40f8-97a9-494692014f07")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IMFDXGIDeviceManager
        Function CloseDeviceIntPtr(hDevice As IntPtr) As HRESULT
        Function GetVideoService(hDevice As IntPtr, ByRef riid As Guid, <Out> ByRef ppService As IntPtr) As HRESULT
        Function LockDevice(hDevice As IntPtr, ByRef riid As Guid, <Out> ByRef ppUnkDevice As IntPtr, fBlock As Boolean) As HRESULT
        Function OpenDeviceHandle(<Out> ByRef phDevice As IntPtr) As HRESULT
        Function ResetDevice(pUnkDevice As IntPtr, resetToken As UInteger) As HRESULT
        Function TestDevice(hDevice As IntPtr) As HRESULT
        Function UnlockDevice(hDevice As IntPtr, fSaveState As Boolean) As HRESULT
    End Interface

    <ComImport>
    <Guid("9B7E4E00-342C-4106-A19F-4F2704F689F0")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID3D10Multithread
        Sub Enter()
        Sub Leave()
        Function SetMultithreadProtected(bMTProtect As Boolean) As Boolean
        Function GetMultithreadProtected() As Boolean
    End Interface

    <ComImport>
    <Guid("a8be2ac4-199f-4946-b331-79599fb98de7")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDXGISwapChain2
        Inherits IDXGISwapChain1
#Region "IDXGISwapChain1"
#Region "IDXGISwapChain"
#Region "IDXGIDeviceSubObject"
#Region "IDXGIObject"
        Overloads Function SetPrivateData(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        Overloads Function SetPrivateDataInterface(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, pUnknown As IntPtr) As HRESULT
        Overloads Function GetPrivateData(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        Overloads Function GetParent(
<MarshalAs(UnmanagedType.LPStruct)> riid As Guid, <Out> ByRef ppParent As IntPtr) As HRESULT
#End Region

        Overloads Function GetDevice(
<MarshalAs(UnmanagedType.LPStruct)> riid As Guid, <Out> ByRef ppDevice As IntPtr) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function Present(SyncInterval As UInteger, Flags As UInteger) As HRESULT
        Overloads Function GetBuffer(Buffer As UInteger, ByRef riid As Guid, <Out> ByRef ppSurface As IntPtr) As HRESULT
        Overloads Function SetFullscreenState(Fullscreen As Boolean, pTarget As IDXGIOutput) As HRESULT
        Overloads Function GetFullscreenState(<Out> ByRef pFullscreen As Boolean, <Out> ByRef ppTarget As IDXGIOutput) As HRESULT
        Overloads Function GetDesc(<Out> ByRef pDesc As DXGI_SWAP_CHAIN_DESC) As HRESULT

        <PreserveSig>
        Overloads Function ResizeBuffers(BufferCount As UInteger, Width As UInteger, Height As UInteger, NewFormat As DXGI_FORMAT, SwapChainFlags As UInteger) As HRESULT
        Overloads Function ResizeTarget(pNewTargetParameters As DXGI_MODE_DESC) As HRESULT
        Overloads Function GetContainingOutput(<Out> ByRef ppOutput As IDXGIOutput) As HRESULT
        Overloads Function GetFrameStatistics(<Out> ByRef pStats As DXGI_FRAME_STATISTICS) As HRESULT
        Overloads Function GetLastPresentCount(<Out> ByRef pLastPresentCount As UInteger) As HRESULT
#End Region

        Overloads Function GetDesc1(<Out> ByRef pDesc As DXGI_SWAP_CHAIN_DESC1) As HRESULT
        Overloads Function GetFullscreenDesc(<Out> ByRef pDesc As DXGI_SWAP_CHAIN_FULLSCREEN_DESC) As HRESULT
        Overloads Function GetIntPtr(<Out> ByRef pIntPtr As IntPtr) As HRESULT
        Overloads Function GetCoreWindow(ByRef refiid As Guid, <Out> ByRef ppUnk As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function Present1(SyncInterval As UInteger, PresentFlags As UInteger, pPresentParameters As DXGI_PRESENT_PARAMETERS) As HRESULT
        Overloads Function IsTemporaryMonoSupported() As Boolean
        Overloads Function GetRestrictToOutput(<Out> ByRef ppRestrictToOutput As IDXGIOutput) As HRESULT
        Overloads Function SetBackgroundColor(pColor As DXGI_RGBA) As HRESULT
        Overloads Function GetBackgroundColor(<Out> ByRef pColor As DXGI_RGBA) As HRESULT
        Overloads Function SetRotation(Rotation As DXGI_MODE_ROTATION) As HRESULT
        Overloads Function GetRotation(<Out> ByRef pRotation As DXGI_MODE_ROTATION) As HRESULT
#End Region

        <PreserveSig>
        Function SetSourceSize(Width As UInteger, Height As UInteger) As HRESULT
        Function GetSourceSize(<Out> ByRef pWidth As UInteger, <Out> ByRef pHeight As UInteger) As HRESULT
        Function SetMaximumFrameLatency(MaxLatency As UInteger) As HRESULT
        Function GetMaximumFrameLatency(<Out> ByRef pMaxLatency As UInteger) As HRESULT
        <PreserveSig>
        Function GetFrameLatencyWaitableObject() As IntPtr
        <PreserveSig>
        Function SetMatrixTransform(ByRef pMatrix As DXGI_MATRIX_3X2_F) As HRESULT
        Function GetMatrixTransform(<Out> ByRef pMatrix As DXGI_MATRIX_3X2_F) As HRESULT
    End Interface

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DXGI_MATRIX_3X2_F
        Public _11 As Single
        Public _12 As Single
        Public _21 As Single
        Public _22 As Single
        Public _31 As Single
        Public _32 As Single
    End Structure

    <ComImport>
    <Guid("4AE63092-6327-4c1b-80AE-BFE12EA32B86")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDXGISurface1
        Inherits IDXGISurface
#Region "<IDXGISurface>"
#Region "<IDXGIDeviceSubObject>"
#Region "<IDXGIObject>"
        Overloads Function SetPrivateData(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        Overloads Function SetPrivateDataInterface(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, pUnknown As IntPtr) As HRESULT
        Overloads Function GetPrivateData(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        Overloads Function GetParent(
<MarshalAs(UnmanagedType.LPStruct)> riid As Guid, <Out> ByRef ppParent As IntPtr) As HRESULT
#End Region
        Overloads Function GetDevice(
<MarshalAs(UnmanagedType.LPStruct)> riid As Guid, <Out> ByRef ppDevice As IntPtr) As HRESULT
#End Region
        Overloads Function GetDesc(<Out> ByRef pDesc As DXGI_SURFACE_DESC) As HRESULT
        Overloads Function Map(<Out> ByRef pLockedRect As DXGI_MAPPED_RECT, MapFlags As UInteger) As HRESULT
        Overloads Function Unmap() As HRESULT
#End Region

        <PreserveSig>
        Function GetDC(Discard As Boolean, <Out> ByRef phdc As IntPtr) As HRESULT
        'HRESULT ReleaseDC( ref RECT pDirtyRect);
        Function ReleaseDC(pDirtyRect As IntPtr) As HRESULT
    End Interface

    <ComImport>
    <Guid("aba496dd-b617-4cb8-a866-bc44d7eb1fa2")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDXGISurface2
        Inherits IDXGISurface1
#Region "<IDXGISurface1>"
#Region "<IDXGISurface>"
#Region "<IDXGIDeviceSubObject>"
#Region "<IDXGIObject>"
        Overloads Function SetPrivateData(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        Overloads Function SetPrivateDataInterface(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, pUnknown As IntPtr) As HRESULT
        Overloads Function GetPrivateData(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        Overloads Function GetParent(
<MarshalAs(UnmanagedType.LPStruct)> riid As Guid, <Out> ByRef ppParent As IntPtr) As HRESULT
#End Region
        Overloads Function GetDevice(
<MarshalAs(UnmanagedType.LPStruct)> riid As Guid, <Out> ByRef ppDevice As IntPtr) As HRESULT
#End Region
        Overloads Function GetDesc(<Out> ByRef pDesc As DXGI_SURFACE_DESC) As HRESULT
        Overloads Function Map(<Out> ByRef pLockedRect As DXGI_MAPPED_RECT, MapFlags As UInteger) As HRESULT
        Overloads Function Unmap() As HRESULT
#End Region

        <PreserveSig>
        Overloads Function GetDC(Discard As Boolean, <Out> ByRef phdc As IntPtr) As HRESULT
        'new HRESULT ReleaseDC(ref RECT pDirtyRect);
        Overloads Function ReleaseDC(pDirtyRect As IntPtr) As HRESULT
#End Region

        Function GetResource(ByRef riid As Guid, <Out> ByRef ppParentResource As IntPtr, <Out> ByRef pSubresourceIndex As UInteger) As HRESULT
    End Interface

    ' D3D11

    <ComImport>
    <Guid("1841e5c8-16b0-489b-bcc8-44cfb0d5deae")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID3D11DeviceChild
        'void GetDevice(out ID3D11Device ppDevice);
        Sub GetDevice(<Out> ByRef ppDevice As IntPtr)
        Function GetPrivateData(ByRef guid As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        Function SetPrivateData(ByRef guid As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        Function SetPrivateDataInterface(ByRef guid As Guid, pData As IntPtr) As HRESULT
    End Interface

    <ComImport>
    <Guid("dc8e63f3-d12b-4952-b47b-5e45026a862d")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID3D11Resource
        Inherits ID3D11DeviceChild
#Region "ID3D11DeviceChild"
        'new void GetDevice(out ID3D11Device ppDevice);
        Overloads Sub GetDevice(<Out> ByRef ppDevice As IntPtr)
        Overloads Function GetPrivateData(ByRef guid As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        Overloads Function SetPrivateData(ByRef guid As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        Overloads Function SetPrivateDataInterface(ByRef guid As Guid, pData As IntPtr) As HRESULT
#End Region

        Sub [GetType](<Out> ByRef pResourceDimension As D3D11_RESOURCE_DIMENSION)
        Sub SetEvictionPriority(EvictionPriority As UInteger)
        Function GetEvictionPriority() As UInteger
    End Interface

    Public Enum D3D11_RESOURCE_DIMENSION
        D3D11_RESOURCE_DIMENSION_UNKNOWN = 0
        D3D11_RESOURCE_DIMENSION_BUFFER = 1
        D3D11_RESOURCE_DIMENSION_TEXTURE1D = 2
        D3D11_RESOURCE_DIMENSION_TEXTURE2D = 3
        D3D11_RESOURCE_DIMENSION_TEXTURE3D = 4
    End Enum

    <ComImport>
    <Guid("6f15aaf2-d208-4e89-9ab4-489535d34f9c")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID3D11Texture2D
        Inherits ID3D11Resource
#Region "ID3D11Resource"
#Region "ID3D11DeviceChild"
        'new void GetDevice(out ID3D11Device ppDevice);
        Overloads Sub GetDevice(<Out> ByRef ppDevice As IntPtr)
        Overloads Function GetPrivateData(ByRef guid As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        Overloads Function SetPrivateData(ByRef guid As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        Overloads Function SetPrivateDataInterface(ByRef guid As Guid, pData As IntPtr) As HRESULT
#End Region

        Overloads Sub [GetType](<Out> ByRef pResourceDimension As D3D11_RESOURCE_DIMENSION)
        Overloads Sub SetEvictionPriority(EvictionPriority As UInteger)
        Overloads Function GetEvictionPriority() As UInteger
#End Region

        Sub GetDesc(<Out> ByRef pDesc As D3D11_TEXTURE2D_DESC)
    End Interface

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D3D11_TEXTURE2D_DESC
        Public Width As UInteger
        Public Height As UInteger
        Public MipLevels As UInteger
        Public ArraySize As UInteger
        Public Format As DXGI_FORMAT
        Public SampleDesc As DXGI_SAMPLE_DESC
        Public Usage As D3D11_USAGE
        Public BindFlags As UInteger
        Public CPUAccessFlags As UInteger
        Public MiscFlags As UInteger
    End Structure
    Public Enum D3D11_USAGE
        D3D11_USAGE_DEFAULT = 0
        D3D11_USAGE_IMMUTABLE = 1
        D3D11_USAGE_DYNAMIC = 2
        D3D11_USAGE_STAGING = 3
    End Enum

    <ComImport>
    <Guid("e7174cfa-1c9e-48b1-8866-626226bfc258")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IMFDXGIBuffer
        Function GetResource(ByRef riid As Guid, <Out> ByRef ppvObject As IntPtr) As HRESULT
        Function GetSubresourceIndex(<Out> ByRef puSubresource As UInteger) As HRESULT
        Function GetUnknown(ByRef guid As Guid, ByRef riid As Guid, <Out> ByRef ppvObject As IntPtr) As HRESULT
        Function SetUnknown(ByRef guid As Guid, pUnkData As IntPtr) As HRESULT
    End Interface

    <ComImport>
    <Guid("035f3ab4-482e-4e50-b41f-8a7f8bd8960b")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDXGIResource
        Inherits IDXGIDeviceSubObject
#Region "<IDXGIDeviceSubObject>"
#Region "<IDXGIObject>"
        Overloads Function SetPrivateData(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        Overloads Function SetPrivateDataInterface(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, pUnknown As IntPtr) As HRESULT
        Overloads Function GetPrivateData(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        Overloads Function GetParent(
<MarshalAs(UnmanagedType.LPStruct)> riid As Guid, <Out> ByRef ppParent As IntPtr) As HRESULT
#End Region

        Overloads Function GetDevice(
<MarshalAs(UnmanagedType.LPStruct)> riid As Guid, <Out> ByRef ppDevice As IntPtr) As HRESULT
#End Region

        <PreserveSig>
        Function GetSharedHandle(<Out> ByRef pSharedHandle As IntPtr) As HRESULT
        Function GetUsage(<Out> ByRef pUsage As UInteger) As HRESULT
        Function SetEvictionPriority(EvictionPriority As UInteger) As HRESULT
        Function GetEvictionPriority(<Out> ByRef pEvictionPriority As UInteger) As HRESULT
    End Interface

    <ComImport>
    <Guid("30961379-4609-4a41-998e-54fe567ee0c1")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDXGIResource1
        Inherits IDXGIResource
#Region "<IDXGIResource>"
#Region "<IDXGIDeviceSubObject>"

#Region "<IDXGIObject>"
        Overloads Function SetPrivateData(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        Overloads Function SetPrivateDataInterface(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, pUnknown As IntPtr) As HRESULT
        Overloads Function GetPrivateData(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        Overloads Function GetParent(
<MarshalAs(UnmanagedType.LPStruct)> riid As Guid, <Out> ByRef ppParent As IntPtr) As HRESULT
#End Region

        Overloads Function GetDevice(
<MarshalAs(UnmanagedType.LPStruct)> riid As Guid, <Out> ByRef ppDevice As IntPtr) As HRESULT
#End Region

        Overloads Function GetSharedHandle(<Out> ByRef pSharedHandle As IntPtr) As HRESULT
        Overloads Function GetUsage(<Out> ByRef pUsage As UInteger) As HRESULT
        Overloads Function SetEvictionPriority(EvictionPriority As UInteger) As HRESULT
        Overloads Function GetEvictionPriority(<Out> ByRef pEvictionPriority As UInteger) As HRESULT
#End Region

        Function CreateSubresourceSurface(index As UInteger, <Out> ByRef ppSurface As IDXGISurface2) As HRESULT
        'HRESULT CreateSharedHandle(SECURITY_ATTRIBUTES pAttributes, uint dwAccess, string lpName, out IntPtr pHandle);
        Function CreateSharedHandle(pAttributes As IntPtr, dwAccess As UInteger, lpName As String, <Out> ByRef pHandle As IntPtr) As HRESULT
    End Interface

    <ComImport>
    <Guid("9d8e1289-d7b3-465f-8126-250e349af85d")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDXGIKeyedMutex
        Inherits IDXGIDeviceSubObject
#Region "<IDXGIDeviceSubObject>"

#Region "<IDXGIObject>"
        Overloads Function SetPrivateData(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        Overloads Function SetPrivateDataInterface(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, pUnknown As IntPtr) As HRESULT
        Overloads Function GetPrivateData(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        Overloads Function GetParent(
<MarshalAs(UnmanagedType.LPStruct)> riid As Guid, <Out> ByRef ppParent As IntPtr) As HRESULT
#End Region

        Overloads Function GetDevice(
<MarshalAs(UnmanagedType.LPStruct)> riid As Guid, <Out> ByRef ppDevice As IntPtr) As HRESULT
#End Region

        Function AcquireSync(Key As ULong, dwMilliseconds As UInteger) As HRESULT
        Function ReleaseSync(Key As ULong) As HRESULT
    End Interface
    <ComImport>
    <Guid("41e7d1f2-a591-4f7b-a2e5-fa9c843e1c12")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDXGIFactoryMedia
        <PreserveSig>
        Function CreateSwapChainForCompositionSurfaceHandle(pDevice As IntPtr, hSurface As IntPtr, ByRef pDesc As DXGI_SWAP_CHAIN_DESC1, pRestrictToOutput As IDXGIOutput, <Out> ByRef ppSwapChain As IDXGISwapChain1) As HRESULT
        <PreserveSig>
        Function CreateDecodeSwapChainForCompositionSurfaceHandle(pDevice As IntPtr, hSurface As IntPtr, ByRef pDesc As DXGI_DECODE_SWAP_CHAIN_DESC, pYuvDecodeBuffers As IDXGIResource, pRestrictToOutput As IDXGIOutput, <Out> ByRef ppSwapChain As IDXGIDecodeSwapChain) As HRESULT
    End Interface
    <StructLayout(LayoutKind.Sequential)>
    Public Structure DXGI_DECODE_SWAP_CHAIN_DESC
        Public Flags As UInteger
    End Structure

    <ComImport>
    <Guid("2633066b-4514-4c7a-8fd8-12ea98059d18")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDXGIDecodeSwapChain
        Function PresentBuffer(BufferToPresent As UInteger, SyncInterval As UInteger, Flags As UInteger) As HRESULT
        Function SetSourceRect(pRect As RECT) As HRESULT
        Function SetTargetRect(pRect As RECT) As HRESULT
        Function SetDestSize(Width As UInteger, Height As UInteger) As HRESULT
        Function GetSourceRect(<Out> ByRef pRect As RECT) As HRESULT
        Function GetTargetRect(<Out> ByRef pRect As RECT) As HRESULT
        Function GetDestSize(<Out> ByRef pWidth As UInteger, <Out> ByRef pHeight As UInteger) As HRESULT
        Function SetColorSpace(ColorSpace As DXGI_MULTIPLANE_OVERLAY_YCbCr_FLAGS) As HRESULT
        <PreserveSig>
        Function GetColorSpace() As DXGI_MULTIPLANE_OVERLAY_YCbCr_FLAGS
    End Interface

    Public Enum DXGI_MULTIPLANE_OVERLAY_YCbCr_FLAGS
        DXGI_MULTIPLANE_OVERLAY_YCbCr_FLAG_NOMINAL_RANGE = &H1
        DXGI_MULTIPLANE_OVERLAY_YCbCr_FLAG_BT709 = &H2
        DXGI_MULTIPLANE_OVERLAY_YCbCr_FLAG_xvYCC = &H4
    End Enum

    Public Enum DXGI_FRAME_PRESENTATION_MODE
        DXGI_FRAME_PRESENTATION_MODE_COMPOSED = 0
        DXGI_FRAME_PRESENTATION_MODE_OVERLAY = 1
        DXGI_FRAME_PRESENTATION_MODE_NONE = 2
        DXGI_FRAME_PRESENTATION_MODE_COMPOSITION_FAILURE = 3
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DXGI_FRAME_STATISTICS_MEDIA
        Public PresentCount As UInteger
        Public PresentRefreshCount As UInteger
        Public SyncRefreshCount As UInteger
        Public SyncQPCTime As LARGE_INTEGER
        Public SyncGPUTime As LARGE_INTEGER
        Public CompositionMode As DXGI_FRAME_PRESENTATION_MODE
        Public ApprovedPresentDuration As UInteger
    End Structure

    <ComImport>
    <Guid("dd95b90b-f05f-4f6a-bd65-25bfb264bd84")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDXGISwapChainMedia
        Function GetFrameStatisticsMedia(<Out> ByRef pStats As DXGI_FRAME_STATISTICS_MEDIA) As HRESULT
        Function SetPresentDuration(Duration As UInteger) As HRESULT
        Function CheckPresentDurationSupport(DesiredPresentDuration As UInteger, <Out> ByRef pClosestSmallerPresentDuration As UInteger, <Out> ByRef pClosestLargerPresentDuration As UInteger) As HRESULT
    End Interface

    Public Enum DXGI_OVERLAY_SUPPORT_FLAG
        DXGI_OVERLAY_SUPPORT_FLAG_DIRECT = &H1
        DXGI_OVERLAY_SUPPORT_FLAG_SCALING = &H2
    End Enum


    <ComImport>
    <Guid("25483823-cd46-4c7d-86ca-47aa95b837bd")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDXGIFactory3
        Inherits IDXGIFactory2
#Region "<IDXGIFactory2>"
#Region "<IDXGIFactory1>"
#Region "<IDXGIObject>"
        Overloads Function SetPrivateData(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        Overloads Function SetPrivateDataInterface(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, pUnknown As IntPtr) As HRESULT
        Overloads Function GetPrivateData(
<MarshalAs(UnmanagedType.LPStruct)> Name As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        Overloads Function GetParent(
<MarshalAs(UnmanagedType.LPStruct)> riid As Guid, <Out> ByRef ppParent As IntPtr) As HRESULT
#End Region

#Region "<IDXGIFactory>"
        Overloads Function EnumAdapters(Adapter As UInteger, <Out> ByRef ppAdapter As IDXGIAdapter) As HRESULT
        Overloads Function MakeWindowAssociation(WindowHandle As IntPtr, Flags As UInteger) As HRESULT
        Overloads Function GetWindowAssociation(<Out> ByRef pWindowHandle As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function CreateSwapChain(pDevice As IntPtr, pDesc As DXGI_SWAP_CHAIN_DESC, <Out> ByRef ppSwapChain As IDXGISwapChain) As HRESULT
        Overloads Function CreateSoftwareAdapter([Module] As IntPtr, <Out> ByRef ppAdapter As IDXGIAdapter) As HRESULT
#End Region

        Overloads Function EnumAdapters1(Adapter As UInteger, <Out> ByRef ppAdapter As IDXGIAdapter1) As HRESULT
        Overloads Function IsCurrent() As Boolean
#End Region

        Overloads Function IsWindowedStereoEnabled() As Boolean
        'HRESULT CreateSwapChainForHwnd(IntPtr pDevice, IntPtr hWnd, DXGI_SWAP_CHAIN_DESC1 pDesc, DXGI_SWAP_CHAIN_FULLSCREEN_DESC pFullscreenDesc, IDXGIOutput pRestrictToOutput, out IDXGISwapChain1 ppSwapChain);
        <PreserveSig>
        Overloads Function CreateSwapChainForHwnd(pDevice As IntPtr, hWnd As IntPtr, ByRef pDesc As DXGI_SWAP_CHAIN_DESC1, pFullscreenDesc As IntPtr, pRestrictToOutput As IDXGIOutput, <Out> ByRef ppSwapChain As IDXGISwapChain1) As HRESULT
        <PreserveSig>
        Overloads Function CreateSwapChainForCoreWindow(pDevice As IntPtr, pWindow As IntPtr, ByRef pDesc As DXGI_SWAP_CHAIN_DESC1, pRestrictToOutput As IDXGIOutput, <Out> ByRef ppSwapChain As IDXGISwapChain1) As HRESULT
        Overloads Function GetSharedResourceAdapterLuid(hResource As IntPtr, <Out> ByRef pLuid As LUID) As HRESULT
        Overloads Function RegisterStereoStatusWindow(WindowHandle As IntPtr, wMsg As UInteger, <Out> ByRef pdwCookie As UInteger) As HRESULT
        Overloads Function RegisterStereoStatusEvent(hEvent As IntPtr, <Out> ByRef pdwCookie As UInteger) As HRESULT
        Overloads Sub UnregisterStereoStatus(dwCookie As UInteger)
        Overloads Function RegisterOcclusionStatusWindow(WindowHandle As IntPtr, wMsg As UInteger, <Out> ByRef pdwCookie As UInteger) As HRESULT
        Overloads Function RegisterOcclusionStatusEvent(hEvent As IntPtr, <Out> ByRef pdwCookie As UInteger) As HRESULT
        Overloads Sub UnregisterOcclusionStatus(dwCookie As UInteger)
        <PreserveSig>
        Overloads Function CreateSwapChainForComposition(pDevice As IntPtr, ByRef pDesc As DXGI_SWAP_CHAIN_DESC1, pRestrictToOutput As IDXGIOutput, <Out> ByRef ppSwapChain As IDXGISwapChain1) As HRESULT
#End Region

        <PreserveSig>
        Function GetCreationFlags() As UInteger
    End Interface

    Public Class DXGITools
        Public Shared IID_ID3D11Texture2D As Guid = New Guid("6f15aaf2-d208-4e89-9ab4-489535d34f9c")

        <DllImport("DXGI.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
        Public Shared Function CreateDXGIFactory(
        <[In], MarshalAs(UnmanagedType.LPStruct)> riid As Guid, <Out> ByRef ppFactory As IntPtr) As HRESULT
        End Function

        <DllImport("DXGI.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
        Public Shared Function CreateDXGIFactory1(
        <[In], MarshalAs(UnmanagedType.LPStruct)> riid As Guid, <Out> ByRef ppFactory As IDXGIFactory1) As HRESULT
        End Function

        <DllImport("DXGI.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
        Public Shared Function CreateDXGIFactory2(Flags As UInteger,
        <[In], MarshalAs(UnmanagedType.LPStruct)> riid As Guid, <Out> ByRef ppFactory As IDXGIFactory2) As HRESULT
        End Function

        Public Const DXGI_CREATE_FACTORY_DEBUG As Integer = &H1

        <DllImport("Mfplat.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
        Public Shared Function MFCreateDXGIDeviceManager(<Out> ByRef resetToken As UInteger, <Out> ByRef ppDeviceManager As IMFDXGIDeviceManager) As HRESULT
        End Function

        <DllImport("D3D11.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
        Public Shared Function D3D11CreateDevice(pAdapter As IDXGIAdapter, DriverType As D3D_DRIVER_TYPE, Software As IntPtr, Flags As UInteger,
            <MarshalAs(UnmanagedType.LPArray)> pFeatureLevels As Integer(), FeatureLevels As UInteger, SDKVersion As UInteger, <Out> ByRef ppDevice As IntPtr, <Out> ByRef pFeatureLevel As D3D_FEATURE_LEVEL, <Out> ByRef ppImmediateContext As IntPtr) As HRESULT
        End Function

        <ComImport, Guid("63aad0b8-7c24-40ff-85a8-640d944cc325"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
        Public Interface ISwapChainPanelNative
            <PreserveSig>
            Function SetSwapChain(swapChain As IDXGISwapChain) As HRESULT
        End Interface

        Public Const D3D11_SDK_VERSION As Integer = 7

        Public Const DXGI_MAP_READ As Integer = 1
        Public Const DXGI_MAP_WRITE As Integer = 2
        Public Const DXGI_MAP_DISCARD As Integer = 4

        Public Const DXGI_MAX_SWAP_CHAIN_BUFFERS As Integer = 16
        Public Const DXGI_PRESENT_TEST As Integer = &H1
        Public Const DXGI_PRESENT_DO_NOT_SEQUENCE As Integer = &H2
        Public Const DXGI_PRESENT_RESTART As Integer = &H4
        Public Const DXGI_PRESENT_DO_NOT_WAIT As Integer = &H8
        Public Const DXGI_PRESENT_STEREO_PREFER_RIGHT As Integer = &H10
        Public Const DXGI_PRESENT_STEREO_TEMPORARY_MONO As Integer = &H20
        Public Const DXGI_PRESENT_RESTRICT_TO_OUTPUT As Integer = &H40
        Public Const DXGI_PRESENT_USE_DURATION As Integer = &H100
        Public Const DXGI_PRESENT_ALLOW_TEARING As Integer = &H200

        Public Enum D3D_DRIVER_TYPE
            D3D_DRIVER_TYPE_UNKNOWN = 0
            D3D_DRIVER_TYPE_HARDWARE = D3D_DRIVER_TYPE_UNKNOWN + 1
            D3D_DRIVER_TYPE_REFERENCE = D3D_DRIVER_TYPE_HARDWARE + 1
            D3D_DRIVER_TYPE_NULL = D3D_DRIVER_TYPE_REFERENCE + 1
            D3D_DRIVER_TYPE_SOFTWARE = D3D_DRIVER_TYPE_NULL + 1
            D3D_DRIVER_TYPE_WARP = D3D_DRIVER_TYPE_SOFTWARE + 1
        End Enum

        Public Enum D3D11_CREATE_DEVICE_FLAG
            D3D11_CREATE_DEVICE_SINGLETHREADED = &H1
            D3D11_CREATE_DEVICE_DEBUG = &H2
            D3D11_CREATE_DEVICE_SWITCH_TO_REF = &H4
            D3D11_CREATE_DEVICE_PREVENT_INTERNAL_THREADING_OPTIMIZATIONS = &H8
            D3D11_CREATE_DEVICE_BGRA_SUPPORT = &H20
            D3D11_CREATE_DEVICE_DEBUGGABLE = &H40
            D3D11_CREATE_DEVICE_PREVENT_ALTERING_LAYER_SETTINGS_FROM_REGISTRY = &H80
            D3D11_CREATE_DEVICE_DISABLE_GPU_TIMEOUT = &H100
            D3D11_CREATE_DEVICE_VIDEO_SUPPORT = &H800
        End Enum

        Public Enum D3D_FEATURE_LEVEL
            D3D_FEATURE_LEVEL_1_0_CORE = &H1000
            D3D_FEATURE_LEVEL_9_1 = &H9100
            D3D_FEATURE_LEVEL_9_2 = &H9200
            D3D_FEATURE_LEVEL_9_3 = &H9300
            D3D_FEATURE_LEVEL_10_0 = &HA000
            D3D_FEATURE_LEVEL_10_1 = &HA100
            D3D_FEATURE_LEVEL_11_0 = &HB000
            D3D_FEATURE_LEVEL_11_1 = &HB100
            D3D_FEATURE_LEVEL_12_0 = &HC000
            D3D_FEATURE_LEVEL_12_1 = &HC100
        End Enum

        Public Const DXGI_USAGE_SHADER_INPUT As Integer = &H10
        Public Const DXGI_USAGE_RENDER_TARGET_OUTPUT As Integer = &H20
        Public Const DXGI_USAGE_BACK_BUFFER As Integer = &H40
        Public Const DXGI_USAGE_SHARED As Integer = &H80
        Public Const DXGI_USAGE_READ_ONLY As Integer = &H100
        Public Const DXGI_USAGE_DISCARD_ON_PRESENT As Integer = &H200
        Public Const DXGI_USAGE_UNORDERED_ACCESS As Integer = &H400

        '
        ' DXGI status (success) codes
        '

        '
        ' MessageId: DXGI_STATUS_OCCLUDED
        '
        ' MessageText:
        '
        ' The Present operation was invisible to the user.
        '
        Public Const DXGI_STATUS_OCCLUDED As HRESULT = CType(&H87A0001, HRESULT)

        '
        ' MessageId: DXGI_STATUS_CLIPPED
        '
        ' MessageText:
        '
        ' The Present operation was partially invisible to the user.
        '
        Public Const DXGI_STATUS_CLIPPED As HRESULT = CType(&H87A0002, HRESULT)

        '
        ' MessageId: DXGI_STATUS_NO_REDIRECTION
        '
        ' MessageText:
        '
        ' The driver is requesting that the DXGI runtime not use shared resources to communicate with the Desktop Window Manager.
        '
        Public Const DXGI_STATUS_NO_REDIRECTION As HRESULT = CType(&H87A0004, HRESULT)

        '
        ' MessageId: DXGI_STATUS_NO_DESKTOP_ACCESS
        '
        ' MessageText:
        '
        ' The Present operation was not visible because the Windows session has switched to another desktop (for example, ctrl-alt-de;.
        '
        Public Const DXGI_STATUS_NO_DESKTOP_ACCESS As HRESULT = CType(&H87A0005, HRESULT)

        '
        ' MessageId: DXGI_STATUS_GRAPHICS_VIDPN_SOURCE_IN_USE
        '
        ' MessageText:
        '
        ' The Present operation was not visible because the target monitor was being used for some other purpose.
        '
        Public Const DXGI_STATUS_GRAPHICS_VIDPN_SOURCE_IN_USE As HRESULT = CType(&H87A0006, HRESULT)

        '
        ' MessageId: DXGI_STATUS_MODE_CHANGED
        '
        ' MessageText:
        '
        ' The Present operation was not visible because the display mode changed. DXGI will have re-attempted the presentation.
        '
        Public Const DXGI_STATUS_MODE_CHANGED As HRESULT = CType(&H87A0007, HRESULT)

        '
        ' MessageId: DXGI_STATUS_MODE_CHANGE_IN_PROGRESS
        '
        ' MessageText:
        '
        ' The Present operation was not visible because another Direct3D device was attempting to take fullscreen mode at the time.
        '
        Public Const DXGI_STATUS_MODE_CHANGE_IN_PROGRESS As HRESULT = CType(&H87A0008, HRESULT)

        '
        ' DXGI error codes
        '

        '
        ' MessageId: DXGI_ERROR_INVALID_CALL
        '
        ' MessageText:
        '
        ' The application made a call that is invalid. Either the parameters of the call or the state of some object was incorrect.
        ' Enable the D3D debug layer in order to see details via debug messages.
        '
        Public Const DXGI_ERROR_INVALID_CALL As HRESULT = CType(&H887A0001, HRESULT)

        '
        ' MessageId: DXGI_ERROR_NOT_FOUND
        '
        ' MessageText:
        '
        ' The object was not found. If calling IDXGIFactory::EnumAdaptes, there is no adapter with the specified ordinal.
        '
        Public Const DXGI_ERROR_NOT_FOUND As HRESULT = CType(&H887A0002, HRESULT)

        '
        ' MessageId: DXGI_ERROR_MORE_DATA
        '
        ' MessageText:
        '
        ' The caller did not supply a sufficiently large buffer.
        '
        Public Const DXGI_ERROR_MORE_DATA As HRESULT = CType(&H887A0003, HRESULT)

        '
        ' MessageId: DXGI_ERROR_UNSUPPORTED
        '
        ' MessageText:
        '
        ' The specified device interface or feature level is not supported on this system.
        '
        Public Const DXGI_ERROR_UNSUPPORTED As HRESULT = CType(&H887A0004, HRESULT)

        '
        ' MessageId: DXGI_ERROR_DEVICE_REMOVED
        '
        ' MessageText:
        '
        ' The GPU device instance has been suspended. Use GetDeviceRemovedReason to determine the appropriate action.
        '
        Public Const DXGI_ERROR_DEVICE_REMOVED As HRESULT = CType(&H887A0005, HRESULT)

        '
        ' MessageId: DXGI_ERROR_DEVICE_HUNG
        '
        ' MessageText:
        '
        ' The GPU will not respond to more commands, most likely because of an invalid command passed by the calling application.
        '
        Public Const DXGI_ERROR_DEVICE_HUNG As HRESULT = CType(&H887A0006, HRESULT)

        '
        ' MessageId: DXGI_ERROR_DEVICE_RESET
        '
        ' MessageText:
        '
        ' The GPU will not respond to more commands, most likely because some other application submitted invalid commands.
        ' The calling application should re-create the device and continue.
        '
        Public Const DXGI_ERROR_DEVICE_RESET As HRESULT = CType(&H887A0007, HRESULT)

        '
        ' MessageId: DXGI_ERROR_WAS_STILL_DRAWING
        '
        ' MessageText:
        '
        ' The GPU was busy at the moment when the call was made, and the call was neither executed nor scheduled.
        '
        Public Const DXGI_ERROR_WAS_STILL_DRAWING As HRESULT = CType(&H887A000A, HRESULT)

        '
        ' MessageId: DXGI_ERROR_FRAME_STATISTICS_DISJOINT
        '
        ' MessageText:
        '
        ' An event (such as power cycle) interrupted the gathering of presentation statistics. Any previous statistics should be
        ' considered invalid.
        '
        Public Const DXGI_ERROR_FRAME_STATISTICS_DISJOINT As HRESULT = CType(&H887A000B, HRESULT)

        '
        ' MessageId: DXGI_ERROR_GRAPHICS_VIDPN_SOURCE_IN_USE
        '
        ' MessageText:
        '
        ' Fullscreen mode could not be achieved because the specified output was already in use.
        '
        Public Const DXGI_ERROR_GRAPHICS_VIDPN_SOURCE_IN_USE As HRESULT = CType(&H887A000C, HRESULT)

        '
        ' MessageId: DXGI_ERROR_DRIVER_INTERNAL_ERROR
        '
        ' MessageText:
        '
        ' An internal issue prevented the driver from carrying out the specified operation. The driver's state is probably suspect,
        ' and the application should not continue.
        '
        Public Const DXGI_ERROR_DRIVER_INTERNAL_ERROR As HRESULT = CType(&H887A0020, HRESULT)

        '
        ' MessageId: DXGI_ERROR_NONEXCLUSIVE
        '
        ' MessageText:
        '
        ' A global counter resource was in use, and the specified counter cannot be used by this Direct3D device at this time.
        '
        Public Const DXGI_ERROR_NONEXCLUSIVE As HRESULT = CType(&H887A0021, HRESULT)

        '
        ' MessageId: DXGI_ERROR_NOT_CURRENTLY_AVAILABLE
        '
        ' MessageText:
        '
        ' A resource is not available at the time of the call, but may become available later.
        '
        Public Const DXGI_ERROR_NOT_CURRENTLY_AVAILABLE As HRESULT = CType(&H887A0022, HRESULT)

        '
        ' MessageId: DXGI_ERROR_REMOTE_CLIENT_DISCONNECTED
        '
        ' MessageText:
        '
        ' The application's remote device has been removed due to session disconnect or network disconnect.
        ' The application should call IDXGIFactory1::IsCurrent to find out when the remote device becomes available again.
        '
        Public Const DXGI_ERROR_REMOTE_CLIENT_DISCONNECTED As HRESULT = CType(&H887A0023, HRESULT)

        '
        ' MessageId: DXGI_ERROR_REMOTE_OUTOFMEMORY
        '
        ' MessageText:
        '
        ' The device has been removed during a remote session because the remote computer ran out of memory.
        '
        Public Const DXGI_ERROR_REMOTE_OUTOFMEMORY As HRESULT = CType(&H887A0024, HRESULT)

        '
        ' MessageId: DXGI_ERROR_ACCESS_LOST
        '
        ' MessageText:
        '
        ' The keyed mutex was abandoned.
        '
        Public Const DXGI_ERROR_ACCESS_LOST As HRESULT = CType(&H887A0026, HRESULT)

        '
        ' MessageId: DXGI_ERROR_WAIT_TIMEOUT
        '
        ' MessageText:
        '
        ' The timeout value has elapsed and the resource is not yet available.
        '
        Public Const DXGI_ERROR_WAIT_TIMEOUT As HRESULT = CType(&H887A0027, HRESULT)

        '
        ' MessageId: DXGI_ERROR_SESSION_DISCONNECTED
        '
        ' MessageText:
        '
        ' The output duplication has been turned off because the Windows session ended or was disconnected.
        ' This happens when a remote user disconnects, or when "switch user" is used locally.
        '
        Public Const DXGI_ERROR_SESSION_DISCONNECTED As HRESULT = CType(&H887A0028, HRESULT)

        '
        ' MessageId: DXGI_ERROR_RESTRICT_TO_OUTPUT_STALE
        '
        ' MessageText:
        '
        ' The DXGI output (monitor) to which the swapchain content was restricted, has been disconnected or changed.
        '
        Public Const DXGI_ERROR_RESTRICT_TO_OUTPUT_STALE As HRESULT = CType(&H887A0029, HRESULT)

        '
        ' MessageId: DXGI_ERROR_CANNOT_PROTECT_CONTENT
        '
        ' MessageText:
        '
        ' DXGI is unable to provide content protection on the swapchain. This is typically caused by an older driver,
        ' or by the application using a swapchain that is incompatible with content protection.
        '
        Public Const DXGI_ERROR_CANNOT_PROTECT_CONTENT As HRESULT = CType(&H887A002A, HRESULT)

        '
        ' MessageId: DXGI_ERROR_ACCESS_DENIED
        '
        ' MessageText:
        '
        ' The application is trying to use a resource to which it does not have the required access privileges.
        ' This is most commonly caused by writing to a shared resource with read-only access.
        '
        Public Const DXGI_ERROR_ACCESS_DENIED As HRESULT = CType(&H887A002B, HRESULT)

        '
        ' MessageId: DXGI_ERROR_NAME_ALREADY_EXISTS
        '
        ' MessageText:
        '
        ' The application is trying to create a shared handle using a name that is already associated with some other resource.
        '
        Public Const DXGI_ERROR_NAME_ALREADY_EXISTS As HRESULT = CType(&H887A002C, HRESULT)

        '
        ' MessageId: DXGI_ERROR_SDK_COMPONENT_MISSING
        '
        ' MessageText:
        '
        ' The application requested an operation that depends on an SDK component that is missing or mismatched.
        '
        Public Const DXGI_ERROR_SDK_COMPONENT_MISSING As HRESULT = CType(&H887A002D, HRESULT)

        '
        ' MessageId: DXGI_ERROR_NOT_CURRENT
        '
        ' MessageText:
        '
        ' The DXGI objects that the application has created are no longer current & need to be recreated for this operation to be performed.
        '
        Public Const DXGI_ERROR_NOT_CURRENT As HRESULT = CType(&H887A002E, HRESULT)

        '
        ' MessageId: DXGI_ERROR_HW_PROTECTION_OUTOFMEMORY
        '
        ' MessageText:
        '
        ' Insufficient HW protected memory exits for proper function.
        '
        Public Const DXGI_ERROR_HW_PROTECTION_OUTOFMEMORY As HRESULT = CType(&H887A0030, HRESULT)

        '
        ' MessageId: DXGI_ERROR_DYNAMIC_CODE_POLICY_VIOLATION
        '
        ' MessageText:
        '
        ' Creating this device would violate the process's dynamic code policy.
        '
        Public Const DXGI_ERROR_DYNAMIC_CODE_POLICY_VIOLATION As HRESULT = CType(&H887A0031, HRESULT)

        '
        ' MessageId: DXGI_ERROR_NON_COMPOSITED_UI
        '
        ' MessageText:
        '
        ' The operation failed because the compositor is not in control of the output.
        '
        Public Const DXGI_ERROR_NON_COMPOSITED_UI As HRESULT = CType(&H887A0032, HRESULT)


        '
        ' DXCore error codes
        '

        '
        ' MessageId: DXCORE_ERROR_EVENT_NOT_UNREGISTERED
        '
        ' MessageText:
        '
        ' The application failed to unregister from an event it registered for.
        '
        Public Const DXCORE_ERROR_EVENT_NOT_UNREGISTERED As HRESULT = CType(&H88800001, HRESULT)


        '
        ' DXGI errors that are internal to the Desktop Window Manager
        '

        '
        ' MessageId: DXGI_STATUS_UNOCCLUDED
        '
        ' MessageText:
        '
        ' The swapchain has become unoccluded.
        '
        Public Const DXGI_STATUS_UNOCCLUDED As HRESULT = CType(&H87A0009, HRESULT)

        '
        ' MessageId: DXGI_STATUS_DDA_WAS_STILL_DRAWING
        '
        ' MessageText:
        '
        ' The adapter did not have access to the required resources to complete the Desktop Duplication Present() call, the Present() call needs to be made again
        '
        Public Const DXGI_STATUS_DDA_WAS_STILL_DRAWING As HRESULT = CType(&H87A000A, HRESULT)

        '
        ' MessageId: DXGI_ERROR_MODE_CHANGE_IN_PROGRESS
        '
        ' MessageText:
        '
        ' An on-going mode change prevented completion of the call. The call may succeed if attempted later.
        '
        Public Const DXGI_ERROR_MODE_CHANGE_IN_PROGRESS As HRESULT = CType(&H887A0025, HRESULT)

        '
        ' MessageId: DXGI_STATUS_PRESENT_REQUIRED
        '
        ' MessageText:
        '
        ' The present succeeded but the caller should present again on the next V-sync, even if there are no changes to the content.
        '
        Public Const DXGI_STATUS_PRESENT_REQUIRED As HRESULT = CType(&H87A002F, HRESULT)


        '
        ' DXGI errors that are produced by the D3D Shader Cache component
        '

        '
        ' MessageId: DXGI_ERROR_CACHE_CORRUPT
        '
        ' MessageText:
        '
        ' The cache is corrupt and either could not be opened or could not be reset.
        '
        Public Const DXGI_ERROR_CACHE_CORRUPT As HRESULT = CType(&H887A0033, HRESULT)

        '
        ' MessageId: DXGI_ERROR_CACHE_FULL
        '
        ' MessageText:
        '
        ' This entry would cause the cache to exceed its quota. On a load operation, this may indicate exceeding the maximum in-memory size.
        '
        Public Const DXGI_ERROR_CACHE_FULL As HRESULT = CType(&H887A0034, HRESULT)

        '
        ' MessageId: DXGI_ERROR_CACHE_HASH_COLLISION
        '
        ' MessageText:
        '
        ' A cache entry was found, but the key provided does not match the key stored in the entry.
        '
        Public Const DXGI_ERROR_CACHE_HASH_COLLISION As HRESULT = CType(&H887A0035, HRESULT)

        '
        ' MessageId: DXGI_ERROR_ALREADY_EXISTS
        '
        ' MessageText:
        '
        ' The desired element already exists.
        '
        Public Const DXGI_ERROR_ALREADY_EXISTS As HRESULT = CType(&H887A0036, HRESULT)


        '
        ' DXGI DDI
        '

        '
        ' MessageId: DXGI_DDI_ERR_WASSTILLDRAWING
        '
        ' MessageText:
        '
        ' The GPU was busy when the operation was requested.
        '
        Public Const DXGI_DDI_ERR_WASSTILLDRAWING As HRESULT = CType(&H887B0001, HRESULT)

        '
        ' MessageId: DXGI_DDI_ERR_UNSUPPORTED
        '
        ' MessageText:
        '
        ' The driver has rejected the creation of this resource.
        '
        Public Const DXGI_DDI_ERR_UNSUPPORTED As HRESULT = CType(&H887B0002, HRESULT)

        '
        ' MessageId: DXGI_DDI_ERR_NONEXCLUSIVE
        '
        ' MessageText:
        '
        ' The GPU counter was in use by another process or d3d device when application requested access to it.
        '
        Public Const DXGI_DDI_ERR_NONEXCLUSIVE As HRESULT = CType(&H887B0003, HRESULT)

    End Class
End Namespace
