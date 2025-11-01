Imports System
Imports System.Runtime.InteropServices
Imports System.Text
Imports Direct2D
Imports GlobalStructures

Namespace Global.WIC
    Friend Class WICTools
        Public Shared CLSID_WICImagingFactory As Guid = New Guid("{cacaf262-9370-4615-a13b-9f5539da4c0a}")
        Public Shared GUID_WICPixelFormat32bppBGR As Guid = New Guid("6fddc324-4e03-4bfe-b185-3d77768dc90e")
        Public Shared GUID_WICPixelFormat32bppBGRA As Guid = New Guid("6fddc324-4e03-4bfe-b185-3d77768dc90f")
        Public Shared GUID_WICPixelFormat32bppPBGRA As Guid = New Guid("6fddc324-4e03-4bfe-b185-3d77768dc910")
        Public Shared GUID_ContainerFormatBmp As Guid = New Guid("0af1d87e-fcfe-4188-bdeb-a7906471cbe3")
        Public Shared GUID_ContainerFormatJpeg As Guid = New Guid("19e4a5aa-5662-4fc5-a0c0-1758028e1057")
        Public Shared GUID_ContainerFormatPng As Guid = New Guid("1b7cfaf4-713f-473c-bbcd-6137425faeaf")
        Public Shared GUID_ContainerFormatGif As Guid = New Guid("1f8a5601-7d4d-4cbd-9c82-1bc8d4eeb9a5")
        Public Shared GUID_ContainerFormatTiff As Guid = New Guid("163bcc30-e2e9-4f0b-961d-a3e9fdb788a3")
    End Class

    <StructLayout(LayoutKind.Sequential)>
    Public Structure WICRect
        Public X As Integer
        Public Y As Integer
        Public Width As Integer
        Public Height As Integer
        Private v1 As Integer
        Private v2 As Integer
        Private v3 As Integer
        Private v4 As Integer

        Public Sub New(v1 As Integer, v2 As Integer, v3 As Integer, v4 As Integer)
            Me.New()
            Me.v1 = v1
            Me.v2 = v2
            Me.v3 = v3
            Me.v4 = v4
        End Sub
    End Structure

    <ComImport>
    <Guid("ec5ec8a9-c395-4314-9c77-54d7a935ff70")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IWICImagingFactory
        Function CreateDecoderFromFilename(wzFilename As String, ByRef pguidVendor As Guid, dwDesiredAccess As Integer, metadataOptions As WICDecodeOptions, <Out> ByRef ppIDecoder As IWICBitmapDecoder) As HRESULT
        Function CreateDecoderFromStream(pIStream As ComTypes.IStream, ByRef pguidVendor As Guid, metadataOptions As WICDecodeOptions, <Out> ByRef ppIDecoder As IWICBitmapDecoder) As HRESULT
        Function CreateDecoderFromFileHandle(hFile As IntPtr, ByRef pguidVendor As Guid, metadataOptions As WICDecodeOptions, <Out> ByRef ppIDecoder As IWICBitmapDecoder) As HRESULT
        Function CreateComponentInfo(ByRef clsidComponent As Guid, <Out> ByRef ppIInfo As IWICComponentInfo) As HRESULT
        Function CreateDecoder(ByRef guidContainerFormat As Guid, ByRef pguidVendor As Guid, <Out> ByRef ppIDecoder As IWICBitmapDecoder) As HRESULT
        Function CreateEncoder(ByRef guidContainerFormat As Guid, ByRef pguidVendor As Guid, <Out> ByRef ppIEncoder As IWICBitmapEncoder) As HRESULT
        Function CreatePalette(<Out> ByRef ppIPalette As IWICPalette) As HRESULT
        Function CreateFormatConverter(<Out> ByRef ppIFormatConverter As IWICFormatConverter) As HRESULT
        Function CreateBitmapScaler(<Out> ByRef ppIBitmapScaler As IWICBitmapScaler) As HRESULT
        Function CreateBitmapClipper(<Out> ByRef ppIBitmapClipper As IWICBitmapClipper) As HRESULT
        Function CreateBitmapFlipRotator(<Out> ByRef ppIBitmapFlipRotator As IWICBitmapFlipRotator) As HRESULT
        Function CreateStream(<Out> ByRef ppIWICStream As IWICStream) As HRESULT
        Function CreateColorContext(<Out> ByRef ppIWICColorContext As IWICColorContext) As HRESULT
        Function CreateColorTransformer(<Out> ByRef ppIWICColorTransform As IWICColorTransform) As HRESULT
        Function CreateBitmap(uiWidth As UInteger, uiHeight As UInteger, ByRef pixelFormat As Guid, [option] As WICBitmapCreateCacheOption, <Out> ByRef ppIBitmap As IWICBitmap) As HRESULT
        Function CreateBitmapFromSource(pIBitmapSource As IWICBitmapSource, [option] As WICBitmapCreateCacheOption, <Out> ByRef ppIBitmap As IWICBitmap) As HRESULT
        Function CreateBitmapFromSourceRect(pIBitmapSource As IWICBitmapSource, x As UInteger, y As UInteger, width As UInteger, height As UInteger, <Out> ByRef ppIBitmap As IWICBitmap) As HRESULT
        Function CreateBitmapFromMemory(uiWidth As UInteger, uiHeight As UInteger, ByRef pixelFormat As Guid, cbStride As UInteger, cbBufferSize As UInteger, pbBuffer As IntPtr, <Out> ByRef ppIBitmap As IWICBitmap) As HRESULT
        Function CreateBitmapFromHBITMAP(hBitmap As IntPtr, hPalette As IntPtr, options As WICBitmapAlphaChannelOption, <Out> ByRef ppIBitmap As IWICBitmap) As HRESULT
        Function CreateBitmapFromHICON(hIcon As IntPtr, <Out> ByRef ppIBitmap As IWICBitmap) As HRESULT
        Function CreateComponentEnumerator(componentTypes As WICComponentType, options As WICComponentEnumerateOptions, <Out> ByRef ppIEnumUnknown As IEnumUnknown) As HRESULT
        Function CreateFastMetadataEncoderFromDecoder(pIDecoder As IWICBitmapDecoder, <Out> ByRef ppIFastEncoder As IWICFastMetadataEncoder) As HRESULT
        Function CreateFastMetadataEncoderFromFrameDecode(pIFrameDecoder As IWICBitmapFrameDecode, <Out> ByRef ppIFastEncoder As IWICFastMetadataEncoder) As HRESULT
        Function CreateQueryWriter(ByRef guidMetadataFormat As Guid, ByRef pguidVendor As Guid, <Out> ByRef ppIQueryWriter As IWICMetadataQueryWriter) As HRESULT
        Function CreateQueryWriterFromReader(pIQueryReader As IWICMetadataQueryReader, ByRef pguidVendor As Guid, <Out> ByRef ppIQueryWriter As IWICMetadataQueryWriter) As HRESULT
    End Interface

    <Guid("00000100-0000-0000-C000-000000000046")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IEnumUnknown
        Function [Next](
<[In], MarshalAs(UnmanagedType.U4)> celt As UInteger,
            <Out, MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.IUnknown, SizeParamIndex:=0)> rgelt As Object(), <Out>
<MarshalAs(UnmanagedType.U4)> ByRef pceltFetched As UInteger) As HRESULT
        Function Skip(celt As UInteger) As HRESULT
        Function Reset() As HRESULT
        Function Clone(<Out> ByRef ppenum As IEnumUnknown) As HRESULT
    End Interface

    <ComImport>
    <Guid("7B816B45-1996-4476-B132-DE9E247C8AF0")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IWICImagingFactory2
        Inherits IWICImagingFactory
#Region "<IWICImagingFactory>"
        Overloads Function CreateDecoderFromFilename(wzFilename As String, ByRef pguidVendor As Guid, dwDesiredAccess As Integer, metadataOptions As WICDecodeOptions, <Out> ByRef ppIDecoder As IWICBitmapDecoder) As HRESULT
        Overloads Function CreateDecoderFromStream(pIStream As ComTypes.IStream, ByRef pguidVendor As Guid, metadataOptions As WICDecodeOptions, <Out> ByRef ppIDecoder As IWICBitmapDecoder) As HRESULT
        Overloads Function CreateDecoderFromFileHandle(hFile As IntPtr, ByRef pguidVendor As Guid, metadataOptions As WICDecodeOptions, <Out> ByRef ppIDecoder As IWICBitmapDecoder) As HRESULT
        Overloads Function CreateComponentInfo(ByRef clsidComponent As Guid, <Out> ByRef ppIInfo As IWICComponentInfo) As HRESULT
        Overloads Function CreateDecoder(ByRef guidContainerFormat As Guid, ByRef pguidVendor As Guid, <Out> ByRef ppIDecoder As IWICBitmapDecoder) As HRESULT
        Overloads Function CreateEncoder(ByRef guidContainerFormat As Guid, ByRef pguidVendor As Guid, <Out> ByRef ppIEncoder As IWICBitmapEncoder) As HRESULT
        Overloads Function CreatePalette(<Out> ByRef ppIPalette As IWICPalette) As HRESULT
        Overloads Function CreateFormatConverter(<Out> ByRef ppIFormatConverter As IWICFormatConverter) As HRESULT
        Overloads Function CreateBitmapScaler(<Out> ByRef ppIBitmapScaler As IWICBitmapScaler) As HRESULT
        Overloads Function CreateBitmapClipper(<Out> ByRef ppIBitmapClipper As IWICBitmapClipper) As HRESULT
        Overloads Function CreateBitmapFlipRotator(<Out> ByRef ppIBitmapFlipRotator As IWICBitmapFlipRotator) As HRESULT
        Overloads Function CreateStream(<Out> ByRef ppIWICStream As IWICStream) As HRESULT
        Overloads Function CreateColorContext(<Out> ByRef ppIWICColorContext As IWICColorContext) As HRESULT
        Overloads Function CreateColorTransformer(<Out> ByRef ppIWICColorTransform As IWICColorTransform) As HRESULT
        Overloads Function CreateBitmap(uiWidth As UInteger, uiHeight As UInteger, ByRef pixelFormat As Guid, [option] As WICBitmapCreateCacheOption, <Out> ByRef ppIBitmap As IWICBitmap) As HRESULT
        Overloads Function CreateBitmapFromSource(pIBitmapSource As IWICBitmapSource, [option] As WICBitmapCreateCacheOption, <Out> ByRef ppIBitmap As IWICBitmap) As HRESULT
        Overloads Function CreateBitmapFromSourceRect(pIBitmapSource As IWICBitmapSource, x As UInteger, y As UInteger, width As UInteger, height As UInteger, <Out> ByRef ppIBitmap As IWICBitmap) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapFromMemory(uiWidth As UInteger, uiHeight As UInteger, ByRef pixelFormat As Guid, cbStride As UInteger, cbBufferSize As UInteger, pbBuffer As IntPtr, <Out> ByRef ppIBitmap As IWICBitmap) As HRESULT
        Overloads Function CreateBitmapFromHBITMAP(hBitmap As IntPtr, hPalette As IntPtr, options As WICBitmapAlphaChannelOption, <Out> ByRef ppIBitmap As IWICBitmap) As HRESULT
        Overloads Function CreateBitmapFromHICON(hIcon As IntPtr, <Out> ByRef ppIBitmap As IWICBitmap) As HRESULT
        Overloads Function CreateComponentEnumerator(componentTypes As WICComponentType, options As WICComponentEnumerateOptions, <Out> ByRef ppIEnumUnknown As IEnumUnknown) As HRESULT
        Overloads Function CreateFastMetadataEncoderFromDecoder(pIDecoder As IWICBitmapDecoder, <Out> ByRef ppIFastEncoder As IWICFastMetadataEncoder) As HRESULT
        Overloads Function CreateFastMetadataEncoderFromFrameDecode(pIFrameDecoder As IWICBitmapFrameDecode, <Out> ByRef ppIFastEncoder As IWICFastMetadataEncoder) As HRESULT
        Overloads Function CreateQueryWriter(ByRef guidMetadataFormat As Guid, ByRef pguidVendor As Guid, <Out> ByRef ppIQueryWriter As IWICMetadataQueryWriter) As HRESULT
        Overloads Function CreateQueryWriterFromReader(pIQueryReader As IWICMetadataQueryReader, ByRef pguidVendor As Guid, <Out> ByRef ppIQueryWriter As IWICMetadataQueryWriter) As HRESULT
#End Region

        Function CreateImageEncoder(pD2DDevice As ID2D1Device, <Out> ByRef ppWICImageEncoder As IWICImageEncoder) As HRESULT
    End Interface

    Public Enum WICDecodeOptions
        WICDecodeMetadataCacheOnDemand = 0
        WICDecodeMetadataCacheOnLoad = 1
        WICMETADATACACHEOPTION_FORCE_DWORD = &H7FFFFFFF
    End Enum

    <ComImport>
    <Guid("23BC3F0A-698B-4357-886B-F24D50671334")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IWICComponentInfo
        Function GetComponentType(<Out> ByRef pType As WICComponentType) As HRESULT
        Function GetCLSID(<Out> ByRef pclsid As Guid) As HRESULT
        Function GetSigningStatus(<Out> ByRef pStatus As Integer) As HRESULT
        Function GetAuthor(cchAuthor As UInteger, wzAuthor As StringBuilder, <Out> ByRef pcchActual As UInteger) As HRESULT
        Function GetVendorGUID(<Out> ByRef pguidVendor As Guid) As HRESULT
        Function GetVersion(cchVersion As UInteger, wzVersion As StringBuilder, <Out> ByRef pcchActual As UInteger) As HRESULT
        Function GetSpecVersion(cchSpecVersion As UInteger, wzSpecVersion As StringBuilder, <Out> ByRef pcchActual As UInteger) As HRESULT
        Function GetFriendlyName(cchFriendlyName As UInteger, wzFriendlyName As StringBuilder, <Out> ByRef pcchActual As UInteger) As HRESULT
    End Interface

    <ComImport>
    <Guid("9EDDE9E7-8DEE-47ea-99DF-E6FAF2ED44BF")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IWICBitmapDecoder
        Function QueryCapability(pIStream As ComTypes.IStream, <Out> ByRef pdwCapability As Integer) As HRESULT
        Function Initialize(pIStream As ComTypes.IStream, cacheOptions As WICDecodeOptions) As HRESULT
        Function GetContainerFormat(<Out> ByRef pguidContainerFormat As Guid) As HRESULT
        Function GetDecoderInfo(<Out> ByRef ppIDecoderInfo As IWICBitmapDecoderInfo) As HRESULT
        Function CopyPalette(pIPalette As IWICPalette) As HRESULT
        Function GetMetadataQueryReader(<Out> ByRef ppIMetadataQueryReader As IWICMetadataQueryReader) As HRESULT
        Function GetPreview(<Out> ByRef ppIBitmapSource As IWICBitmapSource) As HRESULT
        Function GetColorContexts(cCount As UInteger,
<Out, [In]> ppIColorContexts As IWICColorContext, <Out> ByRef pcActualCount As UInteger) As HRESULT
        Function GetThumbnail(<Out> ByRef ppIThumbnail As IWICBitmapSource) As HRESULT
        Function GetFrameCount(<Out> ByRef pCount As UInteger) As HRESULT
        Function GetFrame(index As UInteger, <Out> ByRef ppIBitmapFrame As IWICBitmapFrameDecode) As HRESULT
    End Interface

    Public Enum WICComponentType
        WICDecoder = &H1
        WICEncoder = &H2
        WICPixelFormatConverter = &H4
        WICMetadataReader = &H8
        WICMetadataWriter = &H10
        WICPixelFormat = &H20
        WICAllComponents = &H3F
        WICCOMPONENTTYPE_FORCE_DWORD = &H7FFFFFFF
    End Enum

    Public Enum WICComponentEnumerateOptions
        WICComponentEnumerateDefault = 0
        WICComponentEnumerateRefresh = &H1
        WICComponentEnumerateDisabled = &H80000000
        WICComponentEnumerateUnsigned = &H40000000
        WICComponentEnumerateBuiltInOnly = &H20000000
        WICCOMPONENTENUMERATEOPTIONS_FORCE_DWORD = &H7FFFFFFF
    End Enum

    <ComImport>
    <Guid("D8CD007F-D08F-4191-9BFC-236EA7F0E4B5")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IWICBitmapDecoderInfo
        Inherits IWICBitmapCodecInfo
#Region "IWICBitmapCodecInfo"
#Region "IWICComponentInfo"
        Overloads Function GetComponentType(<Out> ByRef pType As WICComponentType) As HRESULT
        Overloads Function GetCLSID(<Out> ByRef pclsid As Guid) As HRESULT
        Overloads Function GetSigningStatus(<Out> ByRef pStatus As Integer) As HRESULT
        Overloads Function GetAuthor(cchAuthor As UInteger, wzAuthor As StringBuilder, <Out> ByRef pcchActual As UInteger) As HRESULT
        Overloads Function GetVendorGUID(<Out> ByRef pguidVendor As Guid) As HRESULT
        Overloads Function GetVersion(cchVersion As UInteger, wzVersion As StringBuilder, <Out> ByRef pcchActual As UInteger) As HRESULT
        Overloads Function GetSpecVersion(cchSpecVersion As UInteger, wzSpecVersion As StringBuilder, <Out> ByRef pcchActual As UInteger) As HRESULT
        Overloads Function GetFriendlyName(cchFriendlyName As UInteger, wzFriendlyName As StringBuilder, <Out> ByRef pcchActual As UInteger) As HRESULT
#End Region

        Overloads Function GetContainerFormat(<Out> ByRef pguidContainerFormat As Guid) As HRESULT
        Overloads Function GetPixelFormats(cFormats As UInteger, ByRef pguidPixelFormats As Guid, <Out> ByRef pcActual As UInteger) As HRESULT
        Overloads Function GetColorManagementVersion(cchColorManagementVersion As UInteger, wzColorManagementVersion As StringBuilder, <Out> ByRef pcchActual As UInteger) As HRESULT
        Overloads Function GetDeviceManufacturer(cchDeviceManufacturer As UInteger, wzDeviceManufacturer As StringBuilder, <Out> ByRef pcchActual As UInteger) As HRESULT
        Overloads Function GetDeviceModels(cchDeviceModels As UInteger, wzDeviceModels As StringBuilder, <Out> ByRef pcchActual As UInteger) As HRESULT
        Overloads Function GetMimeTypes(cchMimeTypes As UInteger, wzMimeTypes As StringBuilder, <Out> ByRef pcchActual As UInteger) As HRESULT
        Overloads Function GetFileExtensions(cchFileExtensions As UInteger, wzFileExtensions As StringBuilder, <Out> ByRef pcchActual As UInteger) As HRESULT
        Overloads Function DoesSupportAnimation(<Out> ByRef pfSupportAnimation As Boolean) As HRESULT
        Overloads Function DoesSupportChromakey(<Out> ByRef pfSupportChromakey As Boolean) As HRESULT
        Overloads Function DoesSupportLossless(<Out> ByRef pfSupportLossless As Boolean) As HRESULT
        Overloads Function DoesSupportMultiframe(<Out> ByRef pfSupportMultiframe As Boolean) As HRESULT
        Overloads Function MatchesMimeType(wzMimeType As String, <Out> ByRef pfMatches As Boolean) As HRESULT
#End Region

        Function GetPatterns(cbSizePatterns As UInteger, <Out> ByRef pPatterns As WICBitmapPattern, <Out> ByRef pcPatterns As UInteger, <Out> ByRef pcbPatternsActual As UInteger) As HRESULT
        Function MatchesPattern(pIStream As ComTypes.IStream, <Out> ByRef pfMatches As Boolean) As HRESULT
        Function CreateInstance(<Out> ByRef ppIBitmapDecoder As IWICBitmapDecoder) As HRESULT
    End Interface

    <ComImport>
    <Guid("E87A44C4-B76E-4c47-8B09-298EB12A2714")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IWICBitmapCodecInfo
        Inherits IWICComponentInfo
#Region "IWICComponentInfo"
        Overloads Function GetComponentType(<Out> ByRef pType As WICComponentType) As HRESULT
        Overloads Function GetCLSID(<Out> ByRef pclsid As Guid) As HRESULT
        Overloads Function GetSigningStatus(<Out> ByRef pStatus As Integer) As HRESULT
        Overloads Function GetAuthor(cchAuthor As UInteger, wzAuthor As StringBuilder, <Out> ByRef pcchActual As UInteger) As HRESULT
        Overloads Function GetVendorGUID(<Out> ByRef pguidVendor As Guid) As HRESULT
        Overloads Function GetVersion(cchVersion As UInteger, wzVersion As StringBuilder, <Out> ByRef pcchActual As UInteger) As HRESULT
        Overloads Function GetSpecVersion(cchSpecVersion As UInteger, wzSpecVersion As StringBuilder, <Out> ByRef pcchActual As UInteger) As HRESULT
        Overloads Function GetFriendlyName(cchFriendlyName As UInteger, wzFriendlyName As StringBuilder, <Out> ByRef pcchActual As UInteger) As HRESULT
#End Region

        Function GetContainerFormat(<Out> ByRef pguidContainerFormat As Guid) As HRESULT
        Function GetPixelFormats(cFormats As UInteger, ByRef pguidPixelFormats As Guid, <Out> ByRef pcActual As UInteger) As HRESULT
        Function GetColorManagementVersion(cchColorManagementVersion As UInteger, wzColorManagementVersion As StringBuilder, <Out> ByRef pcchActual As UInteger) As HRESULT
        Function GetDeviceManufacturer(cchDeviceManufacturer As UInteger, wzDeviceManufacturer As StringBuilder, <Out> ByRef pcchActual As UInteger) As HRESULT
        Function GetDeviceModels(cchDeviceModels As UInteger, wzDeviceModels As StringBuilder, <Out> ByRef pcchActual As UInteger) As HRESULT
        Function GetMimeTypes(cchMimeTypes As UInteger, wzMimeTypes As StringBuilder, <Out> ByRef pcchActual As UInteger) As HRESULT
        Function GetFileExtensions(cchFileExtensions As UInteger, wzFileExtensions As StringBuilder, <Out> ByRef pcchActual As UInteger) As HRESULT
        Function DoesSupportAnimation(<Out> ByRef pfSupportAnimation As Boolean) As HRESULT
        Function DoesSupportChromakey(<Out> ByRef pfSupportChromakey As Boolean) As HRESULT
        Function DoesSupportLossless(<Out> ByRef pfSupportLossless As Boolean) As HRESULT
        Function DoesSupportMultiframe(<Out> ByRef pfSupportMultiframe As Boolean) As HRESULT
        Function MatchesMimeType(wzMimeType As String, <Out> ByRef pfMatches As Boolean) As HRESULT
    End Interface

    <ComImport>
    <Guid("00000120-a8f2-4877-ba0a-fd2b6645fb94")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IWICBitmapSource
        Function GetSize(<Out> ByRef puiWidth As UInteger, <Out> ByRef puiHeight As UInteger) As HRESULT
        Function GetPixelFormat(<Out> ByRef pPixelFormat As Guid) As HRESULT
        Function GetResolution(<Out> ByRef pDpiX As Double, <Out> ByRef pDpiY As Double) As HRESULT
        Function CopyPalette(pIPalette As IWICPalette) As HRESULT
        'HRESULT CopyPixels(ref WICRect prc, uint cbStride, uint cbBufferSize, [Out, MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.U1)] byte[] pbBuffer);
        Function CopyPixels(ByRef prc As WICRect, cbStride As UInteger, cbBufferSize As UInteger, pbBuffer As IntPtr) As HRESULT
    End Interface

    <StructLayout(LayoutKind.Sequential)>
    Public Structure WICBitmapPattern
        Public Position As LARGE_INTEGER
        Public Length As UInteger
        ' public byte* Pattern;
        ' public byte* Mask;
        Public Pattern As IntPtr
        Public Mask As IntPtr
        Private EndOfStream As Boolean
    End Structure

    Public Enum WICBitmapEncoderCacheOption
        WICBitmapEncoderCacheInMemory = &H0
        WICBitmapEncoderCacheTempFile = &H1
        WICBitmapEncoderNoCache = &H2
        WICBITMAPENCODERCACHEOPTION_FORCE_DWORD = &H7FFFFFFF
    End Enum

    <ComImport>
    <Guid("94C9B4EE-A09F-4f92-8A1E-4A9BCE7E76FB")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IWICBitmapEncoderInfo
        Inherits IWICBitmapCodecInfo
#Region "<IWICBitmapCodecInfo>"
#Region "IWICComponentInfo"
        Overloads Function GetComponentType(<Out> ByRef pType As WICComponentType) As HRESULT
        Overloads Function GetCLSID(<Out> ByRef pclsid As Guid) As HRESULT
        Overloads Function GetSigningStatus(<Out> ByRef pStatus As Integer) As HRESULT
        Overloads Function GetAuthor(cchAuthor As UInteger, wzAuthor As StringBuilder, <Out> ByRef pcchActual As UInteger) As HRESULT
        Overloads Function GetVendorGUID(<Out> ByRef pguidVendor As Guid) As HRESULT
        Overloads Function GetVersion(cchVersion As UInteger, wzVersion As StringBuilder, <Out> ByRef pcchActual As UInteger) As HRESULT
        Overloads Function GetSpecVersion(cchSpecVersion As UInteger, wzSpecVersion As StringBuilder, <Out> ByRef pcchActual As UInteger) As HRESULT
        Overloads Function GetFriendlyName(cchFriendlyName As UInteger, wzFriendlyName As StringBuilder, <Out> ByRef pcchActual As UInteger) As HRESULT
#End Region

        Overloads Function GetContainerFormat(<Out> ByRef pguidContainerFormat As Guid) As HRESULT
        Overloads Function GetPixelFormats(cFormats As UInteger, ByRef pguidPixelFormats As Guid, <Out> ByRef pcActual As UInteger) As HRESULT
        Overloads Function GetColorManagementVersion(cchColorManagementVersion As UInteger, wzColorManagementVersion As StringBuilder, <Out> ByRef pcchActual As UInteger) As HRESULT
        Overloads Function GetDeviceManufacturer(cchDeviceManufacturer As UInteger, wzDeviceManufacturer As StringBuilder, <Out> ByRef pcchActual As UInteger) As HRESULT
        Overloads Function GetDeviceModels(cchDeviceModels As UInteger, wzDeviceModels As StringBuilder, <Out> ByRef pcchActual As UInteger) As HRESULT
        Overloads Function GetMimeTypes(cchMimeTypes As UInteger, wzMimeTypes As StringBuilder, <Out> ByRef pcchActual As UInteger) As HRESULT
        Overloads Function GetFileExtensions(cchFileExtensions As UInteger, wzFileExtensions As StringBuilder, <Out> ByRef pcchActual As UInteger) As HRESULT
        Overloads Function DoesSupportAnimation(<Out> ByRef pfSupportAnimation As Boolean) As HRESULT
        Overloads Function DoesSupportChromakey(<Out> ByRef pfSupportChromakey As Boolean) As HRESULT
        Overloads Function DoesSupportLossless(<Out> ByRef pfSupportLossless As Boolean) As HRESULT
        Overloads Function DoesSupportMultiframe(<Out> ByRef pfSupportMultiframe As Boolean) As HRESULT
        Overloads Function MatchesMimeType(wzMimeType As String, <Out> ByRef pfMatches As Boolean) As HRESULT

#End Region
        Function CreateInstance(<Out> ByRef ppIBitmapEncoder As IWICBitmapEncoder) As HRESULT
    End Interface

    <ComImport>
    <Guid("00000040-a8f2-4877-ba0a-fd2b6645fb94")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IWICPalette
        Function InitializePredefined(ePaletteType As WICBitmapPaletteType, fAddTransparentColor As Boolean) As HRESULT
        Function InitializeCustom(pColors As UInteger, cCount As UInteger) As HRESULT
        Function InitializeFromBitmap(pISurface As IWICBitmapSource, cCount As UInteger, fAddTransparentColor As Boolean) As HRESULT
        Function InitializeFromPalette(pIPalette As IWICPalette) As HRESULT
        Function [GetType](<Out> ByRef pePaletteType As WICBitmapPaletteType) As HRESULT
        Function GetColorCount(<Out> ByRef pcCount As UInteger) As HRESULT
        Function GetColors(cCount As UInteger, <Out> ByRef pColors As UInteger, <Out> ByRef pcActualColors As UInteger) As HRESULT
        Function IsBlackWhite(<Out> ByRef pfIsBlackWhite As Boolean) As HRESULT
        Function IsGrayscale(<Out> ByRef pfIsGrayscale As Boolean) As HRESULT
        Function HasAlpha(<Out> ByRef pfHasAlpha As Boolean) As HRESULT
    End Interface

    <ComImport>
    <Guid("00000103-a8f2-4877-ba0a-fd2b6645fb94")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IWICBitmapEncoder
        Function Initialize(pIStream As ComTypes.IStream, cacheOption As WICBitmapEncoderCacheOption) As HRESULT
        Function GetContainerFormat(<Out> ByRef pguidContainerFormat As Guid) As HRESULT
        Function GetEncoderInfo(<Out> ByRef ppIEncoderInfo As IWICBitmapEncoderInfo) As HRESULT
        Function SetColorContexts(cCount As UInteger, ppIColorContext As IWICColorContext) As HRESULT
        Function SetPalette(pIPalette As IWICPalette) As HRESULT
        Function SetThumbnail(pIThumbnail As IWICBitmapSource) As HRESULT
        Function SetPreview(pIPreview As IWICBitmapSource) As HRESULT
        Function CreateNewFrame(<Out> ByRef ppIFrameEncode As IWICBitmapFrameEncode,
<Out, [In]> ppIEncoderOptions As IPropertyBag2) As HRESULT
        Function Commit() As HRESULT
        Function GetMetadataQueryWriter(<Out> ByRef ppIMetadataQueryWriter As IWICMetadataQueryWriter) As HRESULT
    End Interface

    Public Enum WICBitmapDitherType
        WICBitmapDitherTypeNone = 0
        WICBitmapDitherTypeSolid = 0
        WICBitmapDitherTypeOrdered4x4 = &H1
        WICBitmapDitherTypeOrdered8x8 = &H2
        WICBitmapDitherTypeOrdered16x16 = &H3
        WICBitmapDitherTypeSpiral4x4 = &H4
        WICBitmapDitherTypeSpiral8x8 = &H5
        WICBitmapDitherTypeDualSpiral4x4 = &H6
        WICBitmapDitherTypeDualSpiral8x8 = &H7
        WICBitmapDitherTypeErrorDiffusion = &H8
        WICBITMAPDITHERTYPE_FORCE_DWORD = &H7FFFFFFF
    End Enum

    <ComImport>
    <Guid("00000301-a8f2-4877-ba0a-fd2b6645fb94")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IWICFormatConverter
        Inherits IWICBitmapSource
#Region "<IWICBitmapSource>"
        Overloads Function GetSize(<Out> ByRef puiWidth As UInteger, <Out> ByRef puiHeight As UInteger) As HRESULT
        Overloads Function GetPixelFormat(<Out> ByRef pPixelFormat As Guid) As HRESULT
        Overloads Function GetResolution(<Out> ByRef pDpiX As Double, <Out> ByRef pDpiY As Double) As HRESULT
        Overloads Function CopyPalette(pIPalette As IWICPalette) As HRESULT
        'new HRESULT CopyPixels(ref WICRect prc, uint cbStride, uint cbBufferSize, [Out, MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.U1)] byte[] pbBuffer);
        Overloads Function CopyPixels(ByRef prc As WICRect, cbStride As UInteger, cbBufferSize As UInteger, pbBuffer As IntPtr) As HRESULT
#End Region
        Function Initialize(pISource As IWICBitmapSource, ByRef dstFormat As Guid, dither As WICBitmapDitherType, pIPalette As IWICPalette, alphaThresholdPercent As Double, paletteTranslate As WICBitmapPaletteType) As HRESULT
        Function CanConvert(ByRef srcPixelFormat As Guid, ByRef dstPixelFormat As Guid, <Out> ByRef pfCanConvert As Boolean) As HRESULT
    End Interface

    <ComImport>
    <Guid("135FF860-22B7-4ddf-B0F6-218F4F299A43")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IWICStream
        Inherits ComTypes.IStream
#Region "IStream"
        Overloads Sub Read(pv As Byte(), cb As Integer, pcbRead As IntPtr)
        Overloads Sub Write(pv As Byte(), cb As Integer, pcbWritten As IntPtr)
        Overloads Sub Seek(dlibMove As Long, dwOrigin As Integer, plibNewPosition As IntPtr)
        Overloads Sub SetSize(libNewSize As Long)
        Overloads Sub CopyTo(pstm As ComTypes.IStream, cb As Long, pcbRead As IntPtr, pcbWritten As IntPtr)
        Overloads Sub Commit(grfCommitFlags As Integer)
        Overloads Sub Revert()
        Overloads Sub LockRegion(libOffset As Long, cb As Long, dwLockType As Integer)
        Overloads Sub UnlockRegion(libOffset As Long, cb As Long, dwLockType As Integer)
        Overloads Sub Stat(<Out> ByRef pstatstg As ComTypes.STATSTG, grfStatFlag As Integer)
        Overloads Sub Clone(<Out> ByRef ppstm As ComTypes.IStream)
#End Region

        Function InitializeFromIStream(pIStream As ComTypes.IStream) As HRESULT
        Function InitializeFromFilename(wzFileName As String, dwDesiredAccess As Integer) As HRESULT
        Function InitializeFromMemory(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.U1)> pbBuffer As Byte(), cbBufferSize As Integer) As HRESULT
        Function InitializeFromIStreamRegion(pIStream As ComTypes.IStream, ulOffset As LARGE_INTEGER, ulMaxSize As LARGE_INTEGER) As HRESULT
    End Interface

    Public Enum WICBitmapInterpolationMode
        WICBitmapInterpolationModeNearestNeighbor = 0
        WICBitmapInterpolationModeLinear = &H1
        WICBitmapInterpolationModeCubic = &H2
        WICBitmapInterpolationModeFant = &H3
        WICBITMAPINTERPOLATIONMODE_FORCE_DWORD = &H7FFFFFFF
    End Enum

    <ComImport>
    <Guid("00000302-a8f2-4877-ba0a-fd2b6645fb94")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IWICBitmapScaler
        Inherits IWICBitmapSource
#Region "<IWICBitmapSource>"
        Overloads Function GetSize(<Out> ByRef puiWidth As UInteger, <Out> ByRef puiHeight As UInteger) As HRESULT
        Overloads Function GetPixelFormat(<Out> ByRef pPixelFormat As Guid) As HRESULT
        Overloads Function GetResolution(<Out> ByRef pDpiX As Double, <Out> ByRef pDpiY As Double) As HRESULT
        Overloads Function CopyPalette(pIPalette As IWICPalette) As HRESULT
        Overloads Function CopyPixels(ByRef prc As WICRect, cbStride As UInteger, cbBufferSize As UInteger, pbBuffer As IntPtr) As HRESULT
#End Region
        Function Initialize(pISource As IWICBitmapSource, uiWidth As UInteger, uiHeight As UInteger, mode As WICBitmapInterpolationMode) As HRESULT
    End Interface

    <ComImport()>
    <Guid("E4FBCF03-223D-4e81-9333-D635556DD1B5")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IWICBitmapClipper
        Inherits IWICBitmapSource
#Region "IWICBitmapSource"
        Overloads Function GetSize(<Out> ByRef puiWidth As UInteger, <Out> ByRef puiHeight As UInteger) As HRESULT
        Overloads Function GetPixelFormat(<Out> ByRef pPixelFormat As Guid) As HRESULT
        Overloads Function GetResolution(<Out> ByRef pDpiX As Double, <Out> ByRef pDpiY As Double) As HRESULT
        Overloads Function CopyPalette(pIPalette As IWICPalette) As HRESULT
        'HRESULT CopyPixels(ref WICRect prc, uint cbStride, uint cbBufferSize, [Out, MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.U1)] byte[] pbBuffer);
        Overloads Function CopyPixels(ByRef prc As WICRect, cbStride As UInteger, cbBufferSize As UInteger, pbBuffer As IntPtr) As HRESULT
#End Region

        Function Initialize(pISource As IWICBitmapSource, prc As WICRect) As HRESULT
    End Interface

    Public Enum WICBitmapTransformOptions
        WICBitmapTransformRotate0 = 0
        WICBitmapTransformRotate90 = &H1
        WICBitmapTransformRotate180 = &H2
        WICBitmapTransformRotate270 = &H3
        WICBitmapTransformFlipHorizontal = &H8
        WICBitmapTransformFlipVertical = &H10
        WICBITMAPTRANSFORMOPTIONS_FORCE_DWORD = &H7FFFFFFF
    End Enum

    <ComImport>
    <Guid("5009834F-2D6A-41ce-9E1B-17C5AFF7A782")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IWICBitmapFlipRotator
        Inherits IWICBitmapSource
#Region "IWICBitmapSource"
        Overloads Function GetSize(<Out> ByRef puiWidth As UInteger, <Out> ByRef puiHeight As UInteger) As HRESULT
        Overloads Function GetPixelFormat(<Out> ByRef pPixelFormat As Guid) As HRESULT
        Overloads Function GetResolution(<Out> ByRef pDpiX As Double, <Out> ByRef pDpiY As Double) As HRESULT
        Overloads Function CopyPalette(pIPalette As IWICPalette) As HRESULT
        'HRESULT CopyPixels(ref WICRect prc, uint cbStride, uint cbBufferSize, [Out, MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.U1)] byte[] pbBuffer);
        Overloads Function CopyPixels(ByRef prc As WICRect, cbStride As UInteger, cbBufferSize As UInteger, pbBuffer As IntPtr) As HRESULT
#End Region

        Function Initialize(pISource As IWICBitmapSource, options As WICBitmapTransformOptions) As HRESULT
    End Interface

    <ComImport>
    <Guid("B66F034F-D0E2-40ab-B436-6DE39E321A94")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IWICColorTransform
        Inherits IWICBitmapSource
#Region "IWICBitmapSource"
        Overloads Function GetSize(<Out> ByRef puiWidth As UInteger, <Out> ByRef puiHeight As UInteger) As HRESULT
        Overloads Function GetPixelFormat(<Out> ByRef pPixelFormat As Guid) As HRESULT
        Overloads Function GetResolution(<Out> ByRef pDpiX As Double, <Out> ByRef pDpiY As Double) As HRESULT
        Overloads Function CopyPalette(pIPalette As IWICPalette) As HRESULT
        'HRESULT CopyPixels(ref WICRect prc, uint cbStride, uint cbBufferSize, [Out, MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.U1)] byte[] pbBuffer);
        Overloads Function CopyPixels(ByRef prc As WICRect, cbStride As UInteger, cbBufferSize As UInteger, pbBuffer As IntPtr) As HRESULT
#End Region

        Function Initialize(pIBitmapSource As IWICBitmapSource, pIContextSource As IWICColorContext, pIContextDest As IWICColorContext, ByRef pixelFmtDest As Guid) As HRESULT
    End Interface

    Public Enum WICBitmapCreateCacheOption
        WICBitmapNoCache = 0
        WICBitmapCacheOnDemand = &H1
        WICBitmapCacheOnLoad = &H2
        WICBITMAPCREATECACHEOPTION_FORCE_DWORD = &H7FFFFFFF
    End Enum

    Public Enum WICBitmapAlphaChannelOption
        WICBitmapUseAlpha = 0
        WICBitmapUsePremultipliedAlpha = &H1
        WICBitmapIgnoreAlpha = &H2
        WICBITMAPALPHACHANNELOPTIONS_FORCE_DWORD = &H7FFFFFFF
    End Enum

    <ComImport>
    <Guid("B84E2C09-78C9-4AC4-8BD3-524AE1663A2F")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IWICFastMetadataEncoder
        Function Commit() As HRESULT
        Function GetMetadataQueryWriter(<Out> ByRef ppIMetadataQueryWriter As IWICMetadataQueryWriter) As HRESULT
    End Interface

    Public Enum WICColorContextType
        WICColorContextUninitialized = 0
        WICColorContextProfile = 1
        WICColorContextExifColorSpace = 2
    End Enum

    <ComImport>
    <Guid("3C613A02-34B2-44ea-9A7C-45AEA9C6FD6D")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IWICColorContext
        Function InitializeFromFilename(wzFilename As String) As HRESULT
        Function InitializeFromMemory(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.U1)> pbBuffer As Byte(), cbBufferSize As Integer) As HRESULT
        Function InitializeFromExifColorSpace(value As UInteger) As HRESULT
        Function [GetType](<Out> ByRef pType As WICColorContextType) As HRESULT
        Function GetProfileBytes(cbBuffer As UInteger,
<Out, MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.U1)> pbBuffer As Byte(), <Out> ByRef pcbActual As UInteger) As HRESULT
        Function GetExifColorSpace(<Out> ByRef pValue As UInteger) As HRESULT
    End Interface

    <ComImport>
    <Guid("00000121-a8f2-4877-ba0a-fd2b6645fb94")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IWICBitmap
        Inherits IWICBitmapSource
#Region "IWICBitmapSource"
        Overloads Function GetSize(<Out> ByRef puiWidth As UInteger, <Out> ByRef puiHeight As UInteger) As HRESULT
        Overloads Function GetPixelFormat(<Out> ByRef pPixelFormat As Guid) As HRESULT
        Overloads Function GetResolution(<Out> ByRef pDpiX As Double, <Out> ByRef pDpiY As Double) As HRESULT
        Overloads Function CopyPalette(pIPalette As IWICPalette) As HRESULT
        'HRESULT CopyPixels(ref WICRect prc, uint cbStride, uint cbBufferSize, [Out, MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.U1)] byte[] pbBuffer);
        Overloads Function CopyPixels(ByRef prc As WICRect, cbStride As UInteger, cbBufferSize As UInteger, pbBuffer As IntPtr) As HRESULT
#End Region

        Function Lock(ByRef prcLock As WICRect, flags As WICBitmapLockFlags, <Out> ByRef ppILock As IWICBitmapLock) As HRESULT
        Function SetPalette(pIPalette As IWICPalette) As HRESULT
        Function SetResolution(dpiX As Double, dpiY As Double) As HRESULT
    End Interface

    <ComImport>
    <Guid("00000123-a8f2-4877-ba0a-fd2b6645fb94")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IWICBitmapLock
        Function GetSize(<Out> ByRef puiWidth As UInteger, <Out> ByRef puiHeight As UInteger) As HRESULT
        Function GetStride(<Out> ByRef pcbStride As UInteger) As HRESULT
        'HRESULT GetDataPointer(out uint pcbBufferSize, out WICInProcPointer ppbData);
        Function GetDataPointer(<Out> ByRef pcbBufferSize As UInteger, <Out> ByRef ppbData As IntPtr) As HRESULT
        Function GetPixelFormat(<Out> ByRef pPixelFormat As Guid) As HRESULT
    End Interface

    Public Enum WICBitmapLockFlags
        WICBitmapLockRead = 1
        WICBitmapLockWrite = 2
        WICBITMAPLOCKFLAGS_FORCE_DWORD = &H7FFFFFFF
    End Enum

    <ComImport>
    <Guid("3B16811B-6A43-4ec9-A813-3D930C13B940")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IWICBitmapFrameDecode
        Inherits IWICBitmapSource
#Region "<IWICBitmapSource>"
        Overloads Function GetSize(<Out> ByRef puiWidth As UInteger, <Out> ByRef puiHeight As UInteger) As HRESULT
        Overloads Function GetPixelFormat(<Out> ByRef pPixelFormat As Guid) As HRESULT
        Overloads Function GetResolution(<Out> ByRef pDpiX As Double, <Out> ByRef pDpiY As Double) As HRESULT
        Overloads Function CopyPalette(pIPalette As IWICPalette) As HRESULT
        'HRESULT CopyPixels(ref WICRect prc, uint cbStride, uint cbBufferSize, [Out, MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.U1)] byte[] pbBuffer);
        Overloads Function CopyPixels(ByRef prc As WICRect, cbStride As UInteger, cbBufferSize As UInteger, pbBuffer As IntPtr) As HRESULT
#End Region

        Function GetMetadataQueryReader(<Out> ByRef ppIMetadataQueryReader As IWICMetadataQueryReader) As HRESULT
        Function GetColorContexts(cCount As UInteger,
<Out, [In]> ppIColorContexts As IWICColorContext, <Out> ByRef pcActualCount As UInteger) As HRESULT
        Function GetThumbnail(<Out> ByRef ppIThumbnail As IWICBitmapSource) As HRESULT
    End Interface

    <ComImport>
    <Guid("A721791A-0DEF-4d06-BD91-2118BF1DB10B")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IWICMetadataQueryWriter
        Inherits IWICMetadataQueryReader
#Region "IWICMetadataQueryReader"
        Overloads Function GetContainerFormat(<Out> ByRef pguidContainerFormat As Guid) As HRESULT
        Overloads Function GetLocation(cchMaxLength As UInteger, wzNamespace As StringBuilder, <Out> ByRef pcchActualLength As UInteger) As HRESULT
        Overloads Function GetMetadataByName(wzName As String,
<Out, [In]> pvarValue As PROPVARIANT) As HRESULT
        ' new HRESULT GetEnumerator(out IEnumString ppIEnumString);
        Overloads Function GetEnumerator(<Out> ByRef ppIEnumString As IntPtr) As HRESULT
#End Region

        Function SetMetadataByName(wzName As String, pvarValue As PROPVARIANT) As HRESULT
        Function RemoveMetadataByName(wzName As String) As HRESULT
    End Interface

    <ComImport>
    <Guid("30989668-E1C9-4597-B395-458EEDB808DF")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IWICMetadataQueryReader
        Function GetContainerFormat(<Out> ByRef pguidContainerFormat As Guid) As HRESULT
        Function GetLocation(cchMaxLength As UInteger, wzNamespace As StringBuilder, <Out> ByRef pcchActualLength As UInteger) As HRESULT
        Function GetMetadataByName(wzName As String,
<Out, [In]> pvarValue As PROPVARIANT) As HRESULT
        'HRESULT GetEnumerator(out IEnumString ppIEnumString);
        Function GetEnumerator(<Out> ByRef ppIEnumString As IntPtr) As HRESULT
    End Interface

    <ComImport>
    <Guid("00000105-a8f2-4877-ba0a-fd2b6645fb94")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IWICBitmapFrameEncode
        Function Initialize(pIEncoderOptions As IPropertyBag2) As HRESULT
        Function SetSize(uiWidth As UInteger, uiHeight As UInteger) As HRESULT
        Function SetResolution(dpiX As Double, dpiY As Double) As HRESULT
        Function SetPixelFormat(
<Out, [In]> pPixelFormat As Guid) As HRESULT
        Function SetColorContexts(cCount As UInteger, ppIColorContext As IWICColorContext) As HRESULT
        Function SetPalette(pIPalette As IWICPalette) As HRESULT
        Function SetThumbnail(pIThumbnail As IWICBitmapSource) As HRESULT
        'HRESULT WritePixels(uint lineCount, uint cbStride, uint cbBufferSize, BYTE* pbPixels);
        Function WritePixels(lineCount As UInteger, cbStride As UInteger, cbBufferSize As UInteger, pbPixels As IntPtr) As HRESULT
        Function WriteSource(pIBitmapSource As IWICBitmapSource, ByRef prc As WICRect) As HRESULT
        Function Commit() As HRESULT
        Function GetMetadataQueryWriter(<Out> ByRef ppIMetadataQueryWriter As IWICMetadataQueryWriter) As HRESULT
    End Interface

    <StructLayout(LayoutKind.Sequential)>
    Public Structure PROPBAG2
        Public dwType As Integer
        Public vt As UShort
        Public cfType As UShort
        Public dwHint As Integer
        ' public LPOLESTR pstrName;
        Public pstrName As String
        Public clsid As Guid
    End Structure

    <ComImport>
    <Guid("22F55882-280B-11d0-A8A9-00A0C90C2004")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IPropertyBag2
        'HRESULT Read(uint cProperties, IErrorLog pErrLog, out VARIANT pvarValue,  [Out, In] HRESULT phrError);
        Function Read(cProperties As UInteger, pErrLog As IErrorLog, <Out> ByRef pvarValue As PROPVARIANT,
<Out, [In]> phrError As HRESULT) As HRESULT
        'HRESULT Write(uint cProperties, PROPBAG2 pPropBag, VARIANT pvarValue);
        Function Write(cProperties As UInteger, pPropBag As PROPBAG2, pvarValue As PROPVARIANT) As HRESULT
        Function CountProperties(<Out> ByRef pcProperties As UInteger) As HRESULT
        Function GetPropertyInfo(iProperty As UInteger, cProperties As UInteger, <Out> ByRef pPropBag As PROPBAG2, <Out> ByRef pcProperties As UInteger) As HRESULT
        'HRESULT LoadObject(LPCOLESTR pstrName, int dwHint, IUnknown pUnkObject,  IErrorLog pErrLog);
        Function LoadObject(
<[In], MarshalAs(UnmanagedType.LPWStr)> pstrName As String, dwHint As Integer, pUnkObject As IntPtr, pErrLog As IErrorLog) As HRESULT
    End Interface

    <ComImport>
    <Guid("3127CA40-446E-11CE-8135-00AA004BB851")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IErrorLog
        'HRESULT AddError(LPCOLESTR pszPropName, System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo);
        Function AddError(
<[In], MarshalAs(UnmanagedType.LPWStr)> pszPropName As String, pExcepInfo As ComTypes.EXCEPINFO) As HRESULT
    End Interface

    Public Enum WICBitmapPaletteType
        WICBitmapPaletteTypeCustom = 0
        WICBitmapPaletteTypeMedianCut = 1
        WICBitmapPaletteTypeFixedBW = 2
        WICBitmapPaletteTypeFixedHalftone8 = 3
        WICBitmapPaletteTypeFixedHalftone27 = 4
        WICBitmapPaletteTypeFixedHalftone64 = 5
        WICBitmapPaletteTypeFixedHalftone125 = 6
        WICBitmapPaletteTypeFixedHalftone216 = 7
        WICBitmapPaletteTypeFixedWebPalette = WICBitmapPaletteTypeFixedHalftone216
        WICBitmapPaletteTypeFixedHalftone252 = 8
        WICBitmapPaletteTypeFixedHalftone256 = 9
        WICBitmapPaletteTypeFixedGray4 = 10
        WICBitmapPaletteTypeFixedGray16 = 11
        WICBitmapPaletteTypeFixedGray256 = 12
        WICBITMAPPALETTETYPE_FORCE_DWORD = &H7FFFFFFF
    End Enum

    <ComImport>
    <Guid("04C75BF8-3CE1-473B-ACC5-3CC4F5E94999")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IWICImageEncoder
        'HRESULT WriteFrame(ID2D1Image pImage, IWICBitmapFrameEncode pFrameEncode, WICImageParameters pImageParameters);
        Function WriteFrame(pImage As ID2D1Image, pFrameEncode As IWICBitmapFrameEncode, pImageParameters As IntPtr) As HRESULT
        'HRESULT WriteFrameThumbnail(ID2D1Image pImage, IWICBitmapFrameEncode pFrameEncode, WICImageParameters pImageParameters);
        Function WriteFrameThumbnail(pImage As ID2D1Image, pFrameEncode As IWICBitmapFrameEncode, pImageParameters As IntPtr) As HRESULT
        'HRESULT WriteThumbnail(ID2D1Image pImage, IWICBitmapEncoder pEncoder, WICImageParameters pImageParameters);
        Function WriteThumbnail(pImage As ID2D1Image, pEncoder As IWICBitmapEncoder, pImageParameters As IntPtr) As HRESULT
    End Interface

    <StructLayout(LayoutKind.Sequential)>
    Public Structure WICImageParameters
        Public PixelFormat As D2D1_PIXEL_FORMAT
        Public DpiX As Single
        Public DpiY As Single
        Public Top As Single
        Public Left As Single
        Public PixelWidth As UInteger
        Public PixelHeight As UInteger
    End Structure
End Namespace
