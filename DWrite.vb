Imports System
Imports System.Runtime.InteropServices
Imports GlobalStructures

'C: \Users\Christian\.nuget\packages\microsoft.windowsappsdk\1.5.240627000\include
'dwrite_core.h

Namespace Global.DWrite
    Friend Class DWriteTools
        ' C:\Program Files\WindowsApps\Microsoft.WindowsAppRuntime.1.5_5001.178.1908.0_x86__8wekyb3d8bbwe
        ' C:\Program Files\WindowsApps\Microsoft.WindowsAppRuntime.1.5_5001.178.1908.0_x64__8wekyb3d8bbwe
        <DllImport("DWriteCore.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
        Public Shared Function DWriteCoreCreateFactory(factoryType As DWRITE_FACTORY_TYPE, ByRef iid As Guid, <Out> ByRef factory As IntPtr) As HRESULT
        End Function

        <DllImport("DWrite.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
        Public Shared Function DWriteCreateFactory(factoryType As DWRITE_FACTORY_TYPE, ByRef iid As Guid, <Out> ByRef factory As IntPtr) As HRESULT
        End Function

        Public Shared CLSID_DWriteFactory As Guid = New Guid("B859EE5A-D838-4B5B-A2E8-1ADC7D93DB48")
        Public Shared CLSID_DWriteFactory1 As Guid = New Guid("30572f99-dac6-41db-a16e-0486307e606a")
        Public Shared CLSID_DWriteFactory2 As Guid = New Guid("0439fc60-ca44-4994-8dee-3a9af7b732ec")
        Public Shared CLSID_DWriteFactory3 As Guid = New Guid("9A1B41C3-D3BB-466A-87FC-FE67556A3B65")
        Public Shared CLSID_DWriteFactory4 As Guid = New Guid("4B0B5BD3-0797-4549-8AC5-FE915CC53856")
        Public Shared CLSID_DWriteFactory5 As Guid = New Guid("958DB99A-BE2A-4F09-AF7D-65189803D1D3")
        Public Shared CLSID_DWriteFactory6 As Guid = New Guid("F3744D80-21F7-42EB-B35D-995BC72FC223")
        Public Shared CLSID_DWriteFactory7 As Guid = New Guid("35D0E0B3-9076-4D2E-A016-A91B568A06B4")

    End Class

    <StructLayout(LayoutKind.Sequential)>
    Public Structure SIZE
        Public cx As Integer
        Public cy As Integer
        Public Sub New(cx As Integer, cy As Integer)
            Me.cx = cx
            Me.cy = cy
        End Sub
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Public Class LOGFONT
        Public lfHeight As Integer = 0
        Public lfWidth As Integer = 0
        Public lfEscapement As Integer = 0
        Public lfOrientation As Integer = 0
        Public lfWeight As Integer = 0
        Public lfItalic As Byte = 0
        Public lfUnderline As Byte = 0
        Public lfStrikeOut As Byte = 0
        Public lfCharSet As Byte = 0
        Public lfOutPrecision As Byte = 0
        Public lfClipPrecision As Byte = 0
        Public lfQuality As Byte = 0
        Public lfPitchAndFamily As Byte = 0
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=32)>
        Public lfFaceName As String = String.Empty
    End Class

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_POINT_2F
        Public x As Single
        Public y As Single

        Public Sub New(x As Single, y As Single)
            Me.x = x
            Me.y = y
        End Sub
    End Structure

    Public Enum DWRITE_FACTORY_TYPE
        ''' <summary>
        ''' Shared factory allow for re-use of cached font data across multiple in process components.
        ''' Such factories also take advantage of cross process font caching components for better performance.
        ''' </summary>
        DWRITE_FACTORY_TYPE_SHARED
        ''' <summary>
        ''' Objects created from the isolated factory do not interact with internal DirectWrite state from other components.
        ''' </summary>
        DWRITE_FACTORY_TYPE_ISOLATED
    End Enum

    <ComImport>
    <Guid("b859ee5a-d838-4b5b-a2e8-1adc7d93db48")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFactory
        Function GetSystemFontCollection(<Out> ByRef fontCollection As IDWriteFontCollection, Optional checkForUpdates As Boolean = False) As HRESULT
        <PreserveSig>
        Function CreateCustomFontCollection(collectionLoader As IDWriteFontCollectionLoader, collectionKey As IntPtr, collectionKeySize As Integer, <Out> ByRef fontCollection As IDWriteFontCollection) As HRESULT
        <PreserveSig>
        Function RegisterFontCollectionLoader(fontCollectionLoader As IDWriteFontCollectionLoader) As HRESULT
        <PreserveSig>
        Function UnregisterFontCollectionLoader(fontCollectionLoader As IDWriteFontCollectionLoader) As HRESULT
        'HRESULT CreateFontFileReference(string filePath, System.Runtime.InteropServices.ComTypes.FILETIME lastWriteTime, out IDWriteFontFile fontFile);
        <PreserveSig>
        Function CreateFontFileReference(
        <MarshalAs(UnmanagedType.LPWStr)> filePath As String, lastWriteTime As IntPtr, <Out> ByRef fontFile As IDWriteFontFile) As HRESULT
        <PreserveSig>
        Function CreateCustomFontFileReference(fontFileReferenceKey As IntPtr, fontFileReferenceKeySize As UInteger, fontFileLoader As IDWriteFontFileLoader, <Out> ByRef fontFile As IDWriteFontFile) As HRESULT
        'HRESULT CreateCustomFontFileReference(IntPtr fontFileReferenceKey, uint fontFileReferenceKeySize,
        '[In()][MarshalAs(UnmanagedType.IUnknown)] object fontFileLoader,
        '[Out, MarshalAs(UnmanagedType.IUnknown)] out IDWriteFontFile fontFile);
        <PreserveSig>
        Function CreateFontFace(fontFaceType As DWRITE_FONT_FACE_TYPE, numberOfFiles As Integer,
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> fontFiles As IDWriteFontFile(), faceIndex As Integer, fontFaceSimulationFlags As DWRITE_FONT_SIMULATIONS, <Out> ByRef fontFace As IDWriteFontFace) As HRESULT
        <PreserveSig>
        Function CreateRenderingParams(<Out> ByRef renderingParams As IDWriteRenderingParams) As HRESULT
        <PreserveSig>
        Function CreateMonitorRenderingParams(monitor As IntPtr, <Out> ByRef renderingParams As IDWriteRenderingParams) As HRESULT
        <PreserveSig>
        Function CreateCustomRenderingParams(gamma As Single, enhancedContrast As Single, clearTypeLevel As Single, pixelGeometry As DWRITE_PIXEL_GEOMETRY, renderingMode As DWRITE_RENDERING_MODE, <Out> ByRef renderingParams As IDWriteRenderingParams) As HRESULT
        <PreserveSig>
        Function RegisterFontFileLoader(fontFileLoader As IDWriteFontFileLoader) As HRESULT
        <PreserveSig>
        Function UnregisterFontFileLoader(fontFileLoader As IDWriteFontFileLoader) As HRESULT
        <PreserveSig>
        Function CreateTextFormat(
        <MarshalAs(UnmanagedType.LPWStr)> fontFamilyName As String, fontCollection As IDWriteFontCollection, fontWeight As DWRITE_FONT_WEIGHT, fontStyle As DWRITE_FONT_STYLE, fontStretch As DWRITE_FONT_STRETCH, fontSize As Single,
            <MarshalAs(UnmanagedType.LPWStr)> localeName As String, <Out> ByRef textFormat As IDWriteTextFormat) As HRESULT
        <PreserveSig>
        Function CreateTypography(<Out> ByRef typography As IDWriteTypography) As HRESULT
        <PreserveSig>
        Function GetGdiInterop(<Out> ByRef gdiInterop As IDWriteGdiInterop) As HRESULT
        <PreserveSig>
        Function CreateTextLayout(
        <MarshalAs(UnmanagedType.LPWStr)> str As String, stringLength As Integer, textFormat As IDWriteTextFormat, maxWidth As Single, maxHeight As Single, <Out> ByRef textLayout As IDWriteTextLayout) As HRESULT
        <PreserveSig>
        Function CreateGdiCompatibleTextLayout(
            <MarshalAs(UnmanagedType.LPWStr)> str As String, stringLength As Integer, textFormat As IDWriteTextFormat, layoutWidth As Single, layoutHeight As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, useGdiNatural As Boolean, <Out> ByRef textLayout As IDWriteTextLayout) As HRESULT
        <PreserveSig>
        Function CreateEllipsisTrimmingSign(textFormat As IDWriteTextFormat, <Out> ByRef trimmingSign As IDWriteInlineObject) As HRESULT
        <PreserveSig>
        Function CreateTextAnalyzer(<Out> ByRef textAnalyzer As IDWriteTextAnalyzer) As HRESULT
        <PreserveSig>
        Function CreateNumberSubstitution(substitutionMethod As DWRITE_NUMBER_SUBSTITUTION_METHOD, localeName As String, ignoreUserOverride As Boolean, <Out> ByRef numberSubstitution As IDWriteNumberSubstitution) As HRESULT
        <PreserveSig>
        Function CreateGlyphRunAnalysis(ByRef glyphRun As DWRITE_GLYPH_RUN, pixelsPerDip As Single, transform As DWRITE_MATRIX, renderingMode As DWRITE_RENDERING_MODE, measuringMode As DWRITE_MEASURING_MODE, baselineOriginX As Single, baselineOriginY As Single, <Out> ByRef glyphRunAnalysis As IDWriteGlyphRunAnalysis) As HRESULT
    End Interface

    <ComImport>
    <Guid("30572f99-dac6-41db-a16e-0486307e606a")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFactory1
        Inherits IDWriteFactory
#Region "IDWriteFactory"
        Overloads Function GetSystemFontCollection(<Out> ByRef fontCollection As IDWriteFontCollection, Optional checkForUpdates As Boolean = False) As HRESULT
        <PreserveSig>
        Overloads Function CreateCustomFontCollection(collectionLoader As IDWriteFontCollectionLoader, collectionKey As IntPtr, collectionKeySize As Integer, <Out> ByRef fontCollection As IDWriteFontCollection) As HRESULT
        <PreserveSig>
        Overloads Function RegisterFontCollectionLoader(fontCollectionLoader As IDWriteFontCollectionLoader) As HRESULT
        <PreserveSig>
        Overloads Function UnregisterFontCollectionLoader(fontCollectionLoader As IDWriteFontCollectionLoader) As HRESULT
        'new HRESULT CreateFontFileReference(string filePath, System.Runtime.InteropServices.ComTypes.FILETIME lastWriteTime, out IDWriteFontFile fontFile);
        <PreserveSig>
        Overloads Function CreateFontFileReference(
        <MarshalAs(UnmanagedType.LPWStr)> filePath As String, lastWriteTime As IntPtr, <Out> ByRef fontFile As IDWriteFontFile) As HRESULT
        <PreserveSig>
        Overloads Function CreateCustomFontFileReference(fontFileReferenceKey As IntPtr, fontFileReferenceKeySize As UInteger, fontFileLoader As IDWriteFontFileLoader, <Out> ByRef fontFile As IDWriteFontFile) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontFace(fontFaceType As DWRITE_FONT_FACE_TYPE, numberOfFiles As Integer,
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> fontFiles As IDWriteFontFile(), faceIndex As Integer, fontFaceSimulationFlags As DWRITE_FONT_SIMULATIONS, <Out> ByRef fontFace As IDWriteFontFace) As HRESULT
        <PreserveSig>
        Overloads Function CreateRenderingParams(<Out> ByRef renderingParams As IDWriteRenderingParams) As HRESULT
        <PreserveSig>
        Overloads Function CreateMonitorRenderingParams(monitor As IntPtr, <Out> ByRef renderingParams As IDWriteRenderingParams) As HRESULT
        <PreserveSig>
        Overloads Function CreateCustomRenderingParams(gamma As Single, enhancedContrast As Single, clearTypeLevel As Single, pixelGeometry As DWRITE_PIXEL_GEOMETRY, renderingMode As DWRITE_RENDERING_MODE, <Out> ByRef renderingParams As IDWriteRenderingParams) As HRESULT
        <PreserveSig>
        Overloads Function RegisterFontFileLoader(fontFileLoader As IDWriteFontFileLoader) As HRESULT
        <PreserveSig>
        Overloads Function UnregisterFontFileLoader(fontFileLoader As IDWriteFontFileLoader) As HRESULT
        <PreserveSig>
        Overloads Function CreateTextFormat(
        <MarshalAs(UnmanagedType.LPWStr)> fontFamilyName As String, fontCollection As IDWriteFontCollection, fontWeight As DWRITE_FONT_WEIGHT, fontStyle As DWRITE_FONT_STYLE, fontStretch As DWRITE_FONT_STRETCH, fontSize As Single,
            <MarshalAs(UnmanagedType.LPWStr)> localeName As String, <Out> ByRef textFormat As IDWriteTextFormat) As HRESULT
        <PreserveSig>
        Overloads Function CreateTypography(<Out> ByRef typography As IDWriteTypography) As HRESULT
        <PreserveSig>
        Overloads Function GetGdiInterop(<Out> ByRef gdiInterop As IDWriteGdiInterop) As HRESULT
        <PreserveSig>
        Overloads Function CreateTextLayout(
        <MarshalAs(UnmanagedType.LPWStr)> str As String, stringLength As Integer, textFormat As IDWriteTextFormat, maxWidth As Single, maxHeight As Single, <Out> ByRef textLayout As IDWriteTextLayout) As HRESULT
        <PreserveSig>
        Overloads Function CreateGdiCompatibleTextLayout(
            <MarshalAs(UnmanagedType.LPWStr)> str As String, stringLength As Integer, textFormat As IDWriteTextFormat, layoutWidth As Single, layoutHeight As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, useGdiNatural As Boolean, <Out> ByRef textLayout As IDWriteTextLayout) As HRESULT
        <PreserveSig>
        Overloads Function CreateEllipsisTrimmingSign(textFormat As IDWriteTextFormat, <Out> ByRef trimmingSign As IDWriteInlineObject) As HRESULT
        <PreserveSig>
        Overloads Function CreateTextAnalyzer(<Out> ByRef textAnalyzer As IDWriteTextAnalyzer) As HRESULT
        <PreserveSig>
        Overloads Function CreateNumberSubstitution(substitutionMethod As DWRITE_NUMBER_SUBSTITUTION_METHOD, localeName As String, ignoreUserOverride As Boolean, <Out> ByRef numberSubstitution As IDWriteNumberSubstitution) As HRESULT
        <PreserveSig>
        Overloads Function CreateGlyphRunAnalysis(ByRef glyphRun As DWRITE_GLYPH_RUN, pixelsPerDip As Single, transform As DWRITE_MATRIX, renderingMode As DWRITE_RENDERING_MODE, measuringMode As DWRITE_MEASURING_MODE, baselineOriginX As Single, baselineOriginY As Single, <Out> ByRef glyphRunAnalysis As IDWriteGlyphRunAnalysis) As HRESULT
#End Region

        <PreserveSig>
        Function GetEudcFontCollection(<Out> ByRef fontCollection As IDWriteFontCollection, Optional checkForUpdates As Boolean = False) As HRESULT
        <PreserveSig>
        Function CreateCustomRenderingParams1(gamma As Single, enhancedContrast As Single, enhancedContrastGrayscale As Single, clearTypeLevel As Single, pixelGeometry As DWRITE_PIXEL_GEOMETRY, renderingMode As DWRITE_RENDERING_MODE, <Out> ByRef renderingParams As IDWriteRenderingParams1) As HRESULT
    End Interface

    <ComImport>
    <Guid("0439fc60-ca44-4994-8dee-3a9af7b732ec")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFactory2
        Inherits IDWriteFactory1
#Region "IDWriteFactory1"
#Region "IDWriteFactory"
        <PreserveSig>
        Overloads Function GetSystemFontCollection(<Out> ByRef fontCollection As IDWriteFontCollection, Optional checkForUpdates As Boolean = False) As HRESULT
        <PreserveSig>
        Overloads Function CreateCustomFontCollection(collectionLoader As IDWriteFontCollectionLoader, collectionKey As IntPtr, collectionKeySize As Integer, <Out> ByRef fontCollection As IDWriteFontCollection) As HRESULT
        <PreserveSig>
        Overloads Function RegisterFontCollectionLoader(fontCollectionLoader As IDWriteFontCollectionLoader) As HRESULT
        <PreserveSig>
        Overloads Function UnregisterFontCollectionLoader(fontCollectionLoader As IDWriteFontCollectionLoader) As HRESULT
        'new HRESULT CreateFontFileReference(string filePath, System.Runtime.InteropServices.ComTypes.FILETIME lastWriteTime, out IDWriteFontFile fontFile);
        <PreserveSig>
        Overloads Function CreateFontFileReference(
        <MarshalAs(UnmanagedType.LPWStr)> filePath As String, lastWriteTime As IntPtr, <Out> ByRef fontFile As IDWriteFontFile) As HRESULT
        <PreserveSig>
        Overloads Function CreateCustomFontFileReference(fontFileReferenceKey As IntPtr, fontFileReferenceKeySize As UInteger, fontFileLoader As IDWriteFontFileLoader, <Out> ByRef fontFile As IDWriteFontFile) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontFace(fontFaceType As DWRITE_FONT_FACE_TYPE, numberOfFiles As Integer,
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> fontFiles As IDWriteFontFile(), faceIndex As Integer, fontFaceSimulationFlags As DWRITE_FONT_SIMULATIONS, <Out> ByRef fontFace As IDWriteFontFace) As HRESULT
        <PreserveSig>
        Overloads Function CreateRenderingParams(<Out> ByRef renderingParams As IDWriteRenderingParams) As HRESULT
        <PreserveSig>
        Overloads Function CreateMonitorRenderingParams(monitor As IntPtr, <Out> ByRef renderingParams As IDWriteRenderingParams) As HRESULT
        <PreserveSig>
        Overloads Function CreateCustomRenderingParams(gamma As Single, enhancedContrast As Single, clearTypeLevel As Single, pixelGeometry As DWRITE_PIXEL_GEOMETRY, renderingMode As DWRITE_RENDERING_MODE, <Out> ByRef renderingParams As IDWriteRenderingParams) As HRESULT
        <PreserveSig>
        Overloads Function RegisterFontFileLoader(fontFileLoader As IDWriteFontFileLoader) As HRESULT
        Overloads Function UnregisterFontFileLoader(fontFileLoader As IDWriteFontFileLoader) As HRESULT
        <PreserveSig>
        Overloads Function CreateTextFormat(
        <MarshalAs(UnmanagedType.LPWStr)> fontFamilyName As String, fontCollection As IDWriteFontCollection, fontWeight As DWRITE_FONT_WEIGHT, fontStyle As DWRITE_FONT_STYLE, fontStretch As DWRITE_FONT_STRETCH, fontSize As Single,
            <MarshalAs(UnmanagedType.LPWStr)> localeName As String, <Out> ByRef textFormat As IDWriteTextFormat) As HRESULT
        <PreserveSig>
        Overloads Function CreateTypography(<Out> ByRef typography As IDWriteTypography) As HRESULT
        <PreserveSig>
        Overloads Function GetGdiInterop(<Out> ByRef gdiInterop As IDWriteGdiInterop) As HRESULT
        <PreserveSig>
        Overloads Function CreateTextLayout(
        <MarshalAs(UnmanagedType.LPWStr)> str As String, stringLength As Integer, textFormat As IDWriteTextFormat, maxWidth As Single, maxHeight As Single, <Out> ByRef textLayout As IDWriteTextLayout) As HRESULT
        <PreserveSig>
        Overloads Function CreateGdiCompatibleTextLayout(
            <MarshalAs(UnmanagedType.LPWStr)> str As String, stringLength As Integer, textFormat As IDWriteTextFormat, layoutWidth As Single, layoutHeight As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, useGdiNatural As Boolean, <Out> ByRef textLayout As IDWriteTextLayout) As HRESULT
        <PreserveSig>
        Overloads Function CreateEllipsisTrimmingSign(textFormat As IDWriteTextFormat, <Out> ByRef trimmingSign As IDWriteInlineObject) As HRESULT
        <PreserveSig>
        Overloads Function CreateTextAnalyzer(<Out> ByRef textAnalyzer As IDWriteTextAnalyzer) As HRESULT
        <PreserveSig>
        Overloads Function CreateNumberSubstitution(substitutionMethod As DWRITE_NUMBER_SUBSTITUTION_METHOD, localeName As String, ignoreUserOverride As Boolean, <Out> ByRef numberSubstitution As IDWriteNumberSubstitution) As HRESULT
        <PreserveSig>
        Overloads Function CreateGlyphRunAnalysis(ByRef glyphRun As DWRITE_GLYPH_RUN, pixelsPerDip As Single, transform As DWRITE_MATRIX, renderingMode As DWRITE_RENDERING_MODE, measuringMode As DWRITE_MEASURING_MODE, baselineOriginX As Single, baselineOriginY As Single, <Out> ByRef glyphRunAnalysis As IDWriteGlyphRunAnalysis) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function GetEudcFontCollection(<Out> ByRef fontCollection As IDWriteFontCollection, Optional checkForUpdates As Boolean = False) As HRESULT
        <PreserveSig>
        Overloads Function CreateCustomRenderingParams1(gamma As Single, enhancedContrast As Single, enhancedContrastGrayscale As Single, clearTypeLevel As Single, pixelGeometry As DWRITE_PIXEL_GEOMETRY, renderingMode As DWRITE_RENDERING_MODE, <Out> ByRef renderingParams As IDWriteRenderingParams1) As HRESULT
#End Region

        <PreserveSig>
        Function GetSystemFontFallback(<Out> ByRef fontFallback As IDWriteFontFallback) As HRESULT
        <PreserveSig>
        Function CreateFontFallbackBuilder(<Out> ByRef fontFallbackBuilder As IDWriteFontFallbackBuilder) As HRESULT
        <PreserveSig>
        Function TranslateColorGlyphRun(baselineOriginX As D2D1_POINT_2F, baselineOriginY As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, glyphRunDescription As IntPtr, measuringMode As DWRITE_MEASURING_MODE, worldToDeviceTransform As IntPtr, colorPaletteIndex As UInteger, <Out> ByRef colorLayers As IDWriteColorGlyphRunEnumerator) As HRESULT
        <PreserveSig>
        Function CreateCustomRenderingParams2(gamma As Single, enhancedContrast As Single, grayscaleEnhancedContrast As Single, clearTypeLevel As Single, pixelGeometry As DWRITE_PIXEL_GEOMETRY, renderingMode As DWRITE_RENDERING_MODE, gridFitMode As DWRITE_GRID_FIT_MODE, <Out> ByRef renderingParams As IDWriteRenderingParams2) As HRESULT
        <PreserveSig>
        Function CreateGlyphRunAnalysis2(ByRef glyphRun As DWRITE_GLYPH_RUN, transform As DWRITE_MATRIX, renderingMode As DWRITE_RENDERING_MODE, measuringMode As DWRITE_MEASURING_MODE, gridFitMode As DWRITE_GRID_FIT_MODE, antialiasMode As DWRITE_TEXT_ANTIALIAS_MODE, baselineOriginX As Single, baselineOriginY As Single, <Out> ByRef glyphRunAnalysis As IDWriteGlyphRunAnalysis) As HRESULT
    End Interface

    <ComImport>
    <Guid("9A1B41C3-D3BB-466A-87FC-FE67556A3B65")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFactory3
        Inherits IDWriteFactory2
#Region "IDWriteFactory2"
#Region "IDWriteFactory1"
#Region "IDWriteFactory"
        Overloads Function GetSystemFontCollection(<Out> ByRef fontCollection As IDWriteFontCollection, Optional checkForUpdates As Boolean = False) As HRESULT
        <PreserveSig>
        Overloads Function CreateCustomFontCollection(collectionLoader As IDWriteFontCollectionLoader, collectionKey As IntPtr, collectionKeySize As Integer, <Out> ByRef fontCollection As IDWriteFontCollection) As HRESULT
        <PreserveSig>
        Overloads Function RegisterFontCollectionLoader(fontCollectionLoader As IDWriteFontCollectionLoader) As HRESULT
        <PreserveSig>
        Overloads Function UnregisterFontCollectionLoader(fontCollectionLoader As IDWriteFontCollectionLoader) As HRESULT
        'new HRESULT CreateFontFileReference(string filePath, System.Runtime.InteropServices.ComTypes.FILETIME lastWriteTime, out IDWriteFontFile fontFile);
        <PreserveSig>
        Overloads Function CreateFontFileReference(
        <MarshalAs(UnmanagedType.LPWStr)> filePath As String, lastWriteTime As IntPtr, <Out> ByRef fontFile As IDWriteFontFile) As HRESULT
        <PreserveSig>
        Overloads Function CreateCustomFontFileReference(fontFileReferenceKey As IntPtr, fontFileReferenceKeySize As UInteger, fontFileLoader As IDWriteFontFileLoader, <Out> ByRef fontFile As IDWriteFontFile) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontFace(fontFaceType As DWRITE_FONT_FACE_TYPE, numberOfFiles As Integer,
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> fontFiles As IDWriteFontFile(), faceIndex As Integer, fontFaceSimulationFlags As DWRITE_FONT_SIMULATIONS, <Out> ByRef fontFace As IDWriteFontFace) As HRESULT
        <PreserveSig>
        Overloads Function CreateRenderingParams(<Out> ByRef renderingParams As IDWriteRenderingParams) As HRESULT
        <PreserveSig>
        Overloads Function CreateMonitorRenderingParams(monitor As IntPtr, <Out> ByRef renderingParams As IDWriteRenderingParams) As HRESULT
        <PreserveSig>
        Overloads Function CreateCustomRenderingParams(gamma As Single, enhancedContrast As Single, clearTypeLevel As Single, pixelGeometry As DWRITE_PIXEL_GEOMETRY, renderingMode As DWRITE_RENDERING_MODE, <Out> ByRef renderingParams As IDWriteRenderingParams) As HRESULT
        <PreserveSig>
        Overloads Function RegisterFontFileLoader(fontFileLoader As IDWriteFontFileLoader) As HRESULT
        <PreserveSig>
        Overloads Function UnregisterFontFileLoader(fontFileLoader As IDWriteFontFileLoader) As HRESULT
        <PreserveSig>
        Overloads Function CreateTextFormat(
        <MarshalAs(UnmanagedType.LPWStr)> fontFamilyName As String, fontCollection As IDWriteFontCollection, fontWeight As DWRITE_FONT_WEIGHT, fontStyle As DWRITE_FONT_STYLE, fontStretch As DWRITE_FONT_STRETCH, fontSize As Single,
            <MarshalAs(UnmanagedType.LPWStr)> localeName As String, <Out> ByRef textFormat As IDWriteTextFormat) As HRESULT
        <PreserveSig>
        Overloads Function CreateTypography(<Out> ByRef typography As IDWriteTypography) As HRESULT
        <PreserveSig>
        Overloads Function GetGdiInterop(<Out> ByRef gdiInterop As IDWriteGdiInterop) As HRESULT
        <PreserveSig>
        Overloads Function CreateTextLayout(
        <MarshalAs(UnmanagedType.LPWStr)> str As String, stringLength As Integer, textFormat As IDWriteTextFormat, maxWidth As Single, maxHeight As Single, <Out> ByRef textLayout As IDWriteTextLayout) As HRESULT
        <PreserveSig>
        Overloads Function CreateGdiCompatibleTextLayout(
            <MarshalAs(UnmanagedType.LPWStr)> str As String, stringLength As Integer, textFormat As IDWriteTextFormat, layoutWidth As Single, layoutHeight As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, useGdiNatural As Boolean, <Out> ByRef textLayout As IDWriteTextLayout) As HRESULT
        <PreserveSig>
        Overloads Function CreateEllipsisTrimmingSign(textFormat As IDWriteTextFormat, <Out> ByRef trimmingSign As IDWriteInlineObject) As HRESULT
        <PreserveSig>
        Overloads Function CreateTextAnalyzer(<Out> ByRef textAnalyzer As IDWriteTextAnalyzer) As HRESULT
        <PreserveSig>
        Overloads Function CreateNumberSubstitution(substitutionMethod As DWRITE_NUMBER_SUBSTITUTION_METHOD, localeName As String, ignoreUserOverride As Boolean, <Out> ByRef numberSubstitution As IDWriteNumberSubstitution) As HRESULT
        <PreserveSig>
        Overloads Function CreateGlyphRunAnalysis(ByRef glyphRun As DWRITE_GLYPH_RUN, pixelsPerDip As Single, transform As DWRITE_MATRIX, renderingMode As DWRITE_RENDERING_MODE, measuringMode As DWRITE_MEASURING_MODE, baselineOriginX As Single, baselineOriginY As Single, <Out> ByRef glyphRunAnalysis As IDWriteGlyphRunAnalysis) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function GetEudcFontCollection(<Out> ByRef fontCollection As IDWriteFontCollection, Optional checkForUpdates As Boolean = False) As HRESULT
        <PreserveSig>
        Overloads Function CreateCustomRenderingParams1(gamma As Single, enhancedContrast As Single, enhancedContrastGrayscale As Single, clearTypeLevel As Single, pixelGeometry As DWRITE_PIXEL_GEOMETRY, renderingMode As DWRITE_RENDERING_MODE, <Out> ByRef renderingParams As IDWriteRenderingParams1) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function GetSystemFontFallback(<Out> ByRef fontFallback As IDWriteFontFallback) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontFallbackBuilder(<Out> ByRef fontFallbackBuilder As IDWriteFontFallbackBuilder) As HRESULT
        <PreserveSig>
        Overloads Function TranslateColorGlyphRun(baselineOriginX As D2D1_POINT_2F, baselineOriginY As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, glyphRunDescription As IntPtr, measuringMode As DWRITE_MEASURING_MODE, worldToDeviceTransform As IntPtr, colorPaletteIndex As UInteger, <Out> ByRef colorLayers As IDWriteColorGlyphRunEnumerator) As HRESULT
        <PreserveSig>
        Overloads Function CreateCustomRenderingParams2(gamma As Single, enhancedContrast As Single, grayscaleEnhancedContrast As Single, clearTypeLevel As Single, pixelGeometry As DWRITE_PIXEL_GEOMETRY, renderingMode As DWRITE_RENDERING_MODE, gridFitMode As DWRITE_GRID_FIT_MODE, <Out> ByRef renderingParams As IDWriteRenderingParams2) As HRESULT
        <PreserveSig>
        Overloads Function CreateGlyphRunAnalysis2(ByRef glyphRun As DWRITE_GLYPH_RUN, transform As DWRITE_MATRIX, renderingMode As DWRITE_RENDERING_MODE, measuringMode As DWRITE_MEASURING_MODE, gridFitMode As DWRITE_GRID_FIT_MODE, antialiasMode As DWRITE_TEXT_ANTIALIAS_MODE, baselineOriginX As Single, baselineOriginY As Single, <Out> ByRef glyphRunAnalysis As IDWriteGlyphRunAnalysis) As HRESULT
#End Region

        <PreserveSig>
        Function CreateGlyphRunAnalysis3(ByRef glyphRun As DWRITE_GLYPH_RUN, transform As DWRITE_MATRIX, renderingMode As DWRITE_RENDERING_MODE1, measuringMode As DWRITE_MEASURING_MODE, gridFitMode As DWRITE_GRID_FIT_MODE, antialiasMode As DWRITE_TEXT_ANTIALIAS_MODE, baselineOriginX As Single, baselineOriginY As Single, <Out> ByRef glyphRunAnalysis As IDWriteGlyphRunAnalysis) As HRESULT
        <PreserveSig>
        Function CreateCustomRenderingParams3(gamma As Single, enhancedContrast As Single, grayscaleEnhancedContrast As Single, clearTypeLevel As Single, pixelGeometry As DWRITE_PIXEL_GEOMETRY, renderingMode As DWRITE_RENDERING_MODE1, gridFitMode As DWRITE_GRID_FIT_MODE, <Out> ByRef renderingParams As IDWriteRenderingParams3) As HRESULT
        'HRESULT CreateFontFaceReference3(string filePath, System.Runtime.InteropServices.ComTypes.FILETIME lastWriteTime, uint faceIndex, DWRITE_FONT_SIMULATIONS fontSimulations, out IDWriteFontFaceReference fontFaceReference);
        <PreserveSig>
        Function CreateFontFaceReference3(filePath As String, lastWriteTime As IntPtr, faceIndex As UInteger, fontSimulations As DWRITE_FONT_SIMULATIONS, <Out> ByRef fontFaceReference As IDWriteFontFaceReference) As HRESULT
        <PreserveSig>
        Function CreateFontFaceReference3(fontFile As IDWriteFontFile, faceIndex As UInteger, fontSimulations As DWRITE_FONT_SIMULATIONS, <Out> ByRef fontFaceReference As IDWriteFontFaceReference) As HRESULT
        <PreserveSig>
        Function GetSystemFontSet3(<Out> ByRef fontSet As IDWriteFontSet) As HRESULT
        <PreserveSig>
        Function CreateFontSetBuilder3(<Out> ByRef fontSetBuilder As IDWriteFontSetBuilder) As HRESULT
        <PreserveSig>
        Function CreateFontCollectionFromFontSet3(fontSet As IDWriteFontSet, <Out> ByRef fontCollection As IDWriteFontCollection1) As HRESULT
        <PreserveSig>
        Function GetSystemFontCollection3(includeDownloadableFonts As Boolean, <Out> ByRef fontCollection As IDWriteFontCollection1, Optional checkForUpdates As Boolean = False) As HRESULT
        <PreserveSig>
        Function GetFontDownloadQueue(<Out> ByRef fontDownloadQueue As IDWriteFontDownloadQueue) As HRESULT
    End Interface

    <ComImport>
    <Guid("4B0B5BD3-0797-4549-8AC5-FE915CC53856")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFactory4
        Inherits IDWriteFactory3
#Region "IDWriteFactory3"
#Region "IDWriteFactory2"
#Region "IDWriteFactory1"
#Region "IDWriteFactory"
        Overloads Function GetSystemFontCollection(<Out> ByRef fontCollection As IDWriteFontCollection, Optional checkForUpdates As Boolean = False) As HRESULT
        <PreserveSig>
        Overloads Function CreateCustomFontCollection(collectionLoader As IDWriteFontCollectionLoader, collectionKey As IntPtr, collectionKeySize As Integer, <Out> ByRef fontCollection As IDWriteFontCollection) As HRESULT
        <PreserveSig>
        Overloads Function RegisterFontCollectionLoader(fontCollectionLoader As IDWriteFontCollectionLoader) As HRESULT
        <PreserveSig>
        Overloads Function UnregisterFontCollectionLoader(fontCollectionLoader As IDWriteFontCollectionLoader) As HRESULT
        'new HRESULT CreateFontFileReference(string filePath, System.Runtime.InteropServices.ComTypes.FILETIME lastWriteTime, out IDWriteFontFile fontFile);
        <PreserveSig>
        Overloads Function CreateFontFileReference(
        <MarshalAs(UnmanagedType.LPWStr)> filePath As String, lastWriteTime As IntPtr, <Out> ByRef fontFile As IDWriteFontFile) As HRESULT
        <PreserveSig>
        Overloads Function CreateCustomFontFileReference(fontFileReferenceKey As IntPtr, fontFileReferenceKeySize As UInteger, fontFileLoader As IDWriteFontFileLoader, <Out> ByRef fontFile As IDWriteFontFile) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontFace(fontFaceType As DWRITE_FONT_FACE_TYPE, numberOfFiles As Integer,
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> fontFiles As IDWriteFontFile(), faceIndex As Integer, fontFaceSimulationFlags As DWRITE_FONT_SIMULATIONS, <Out> ByRef fontFace As IDWriteFontFace) As HRESULT
        <PreserveSig>
        Overloads Function CreateRenderingParams(<Out> ByRef renderingParams As IDWriteRenderingParams) As HRESULT
        <PreserveSig>
        Overloads Function CreateMonitorRenderingParams(monitor As IntPtr, <Out> ByRef renderingParams As IDWriteRenderingParams) As HRESULT
        <PreserveSig>
        Overloads Function CreateCustomRenderingParams(gamma As Single, enhancedContrast As Single, clearTypeLevel As Single, pixelGeometry As DWRITE_PIXEL_GEOMETRY, renderingMode As DWRITE_RENDERING_MODE, <Out> ByRef renderingParams As IDWriteRenderingParams) As HRESULT
        <PreserveSig>
        Overloads Function RegisterFontFileLoader(fontFileLoader As IDWriteFontFileLoader) As HRESULT
        <PreserveSig>
        Overloads Function UnregisterFontFileLoader(fontFileLoader As IDWriteFontFileLoader) As HRESULT
        <PreserveSig>
        Overloads Function CreateTextFormat(
        <MarshalAs(UnmanagedType.LPWStr)> fontFamilyName As String, fontCollection As IDWriteFontCollection, fontWeight As DWRITE_FONT_WEIGHT, fontStyle As DWRITE_FONT_STYLE, fontStretch As DWRITE_FONT_STRETCH, fontSize As Single,
            <MarshalAs(UnmanagedType.LPWStr)> localeName As String, <Out> ByRef textFormat As IDWriteTextFormat) As HRESULT
        <PreserveSig>
        Overloads Function CreateTypography(<Out> ByRef typography As IDWriteTypography) As HRESULT
        <PreserveSig>
        Overloads Function GetGdiInterop(<Out> ByRef gdiInterop As IDWriteGdiInterop) As HRESULT
        <PreserveSig>
        Overloads Function CreateTextLayout(
        <MarshalAs(UnmanagedType.LPWStr)> str As String, stringLength As Integer, textFormat As IDWriteTextFormat, maxWidth As Single, maxHeight As Single, <Out> ByRef textLayout As IDWriteTextLayout) As HRESULT
        <PreserveSig>
        Overloads Function CreateGdiCompatibleTextLayout(
            <MarshalAs(UnmanagedType.LPWStr)> str As String, stringLength As Integer, textFormat As IDWriteTextFormat, layoutWidth As Single, layoutHeight As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, useGdiNatural As Boolean, <Out> ByRef textLayout As IDWriteTextLayout) As HRESULT
        <PreserveSig>
        Overloads Function CreateEllipsisTrimmingSign(textFormat As IDWriteTextFormat, <Out> ByRef trimmingSign As IDWriteInlineObject) As HRESULT
        <PreserveSig>
        Overloads Function CreateTextAnalyzer(<Out> ByRef textAnalyzer As IDWriteTextAnalyzer) As HRESULT
        <PreserveSig>
        Overloads Function CreateNumberSubstitution(substitutionMethod As DWRITE_NUMBER_SUBSTITUTION_METHOD, localeName As String, ignoreUserOverride As Boolean, <Out> ByRef numberSubstitution As IDWriteNumberSubstitution) As HRESULT
        <PreserveSig>
        Overloads Function CreateGlyphRunAnalysis(ByRef glyphRun As DWRITE_GLYPH_RUN, pixelsPerDip As Single, transform As DWRITE_MATRIX, renderingMode As DWRITE_RENDERING_MODE, measuringMode As DWRITE_MEASURING_MODE, baselineOriginX As Single, baselineOriginY As Single, <Out> ByRef glyphRunAnalysis As IDWriteGlyphRunAnalysis) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function GetEudcFontCollection(<Out> ByRef fontCollection As IDWriteFontCollection, Optional checkForUpdates As Boolean = False) As HRESULT
        <PreserveSig>
        Overloads Function CreateCustomRenderingParams1(gamma As Single, enhancedContrast As Single, enhancedContrastGrayscale As Single, clearTypeLevel As Single, pixelGeometry As DWRITE_PIXEL_GEOMETRY, renderingMode As DWRITE_RENDERING_MODE, <Out> ByRef renderingParams As IDWriteRenderingParams1) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function GetSystemFontFallback(<Out> ByRef fontFallback As IDWriteFontFallback) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontFallbackBuilder(<Out> ByRef fontFallbackBuilder As IDWriteFontFallbackBuilder) As HRESULT
        <PreserveSig>
        Overloads Function TranslateColorGlyphRun(baselineOriginX As D2D1_POINT_2F, baselineOriginY As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, glyphRunDescription As IntPtr, measuringMode As DWRITE_MEASURING_MODE, worldToDeviceTransform As IntPtr, colorPaletteIndex As UInteger, <Out> ByRef colorLayers As IDWriteColorGlyphRunEnumerator) As HRESULT
        <PreserveSig>
        Overloads Function CreateCustomRenderingParams2(gamma As Single, enhancedContrast As Single, grayscaleEnhancedContrast As Single, clearTypeLevel As Single, pixelGeometry As DWRITE_PIXEL_GEOMETRY, renderingMode As DWRITE_RENDERING_MODE, gridFitMode As DWRITE_GRID_FIT_MODE, <Out> ByRef renderingParams As IDWriteRenderingParams2) As HRESULT
        <PreserveSig>
        Overloads Function CreateGlyphRunAnalysis2(ByRef glyphRun As DWRITE_GLYPH_RUN, transform As DWRITE_MATRIX, renderingMode As DWRITE_RENDERING_MODE, measuringMode As DWRITE_MEASURING_MODE, gridFitMode As DWRITE_GRID_FIT_MODE, antialiasMode As DWRITE_TEXT_ANTIALIAS_MODE, baselineOriginX As Single, baselineOriginY As Single, <Out> ByRef glyphRunAnalysis As IDWriteGlyphRunAnalysis) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateGlyphRunAnalysis3(ByRef glyphRun As DWRITE_GLYPH_RUN, transform As DWRITE_MATRIX, renderingMode As DWRITE_RENDERING_MODE1, measuringMode As DWRITE_MEASURING_MODE, gridFitMode As DWRITE_GRID_FIT_MODE, antialiasMode As DWRITE_TEXT_ANTIALIAS_MODE, baselineOriginX As Single, baselineOriginY As Single, <Out> ByRef glyphRunAnalysis As IDWriteGlyphRunAnalysis) As HRESULT
        <PreserveSig>
        Overloads Function CreateCustomRenderingParams3(gamma As Single, enhancedContrast As Single, grayscaleEnhancedContrast As Single, clearTypeLevel As Single, pixelGeometry As DWRITE_PIXEL_GEOMETRY, renderingMode As DWRITE_RENDERING_MODE1, gridFitMode As DWRITE_GRID_FIT_MODE, <Out> ByRef renderingParams As IDWriteRenderingParams3) As HRESULT
        'new HRESULT CreateFontFaceReference3(string filePath, System.Runtime.InteropServices.ComTypes.FILETIME lastWriteTime, uint faceIndex, DWRITE_FONT_SIMULATIONS fontSimulations, out IDWriteFontFaceReference fontFaceReference);
        <PreserveSig>
        Overloads Function CreateFontFaceReference3(filePath As String, lastWriteTime As IntPtr, faceIndex As UInteger, fontSimulations As DWRITE_FONT_SIMULATIONS, <Out> ByRef fontFaceReference As IDWriteFontFaceReference) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontFaceReference3(fontFile As IDWriteFontFile, faceIndex As UInteger, fontSimulations As DWRITE_FONT_SIMULATIONS, <Out> ByRef fontFaceReference As IDWriteFontFaceReference) As HRESULT
        Overloads Function GetSystemFontSet3(<Out> ByRef fontSet As IDWriteFontSet) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontSetBuilder3(<Out> ByRef fontSetBuilder As IDWriteFontSetBuilder) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontCollectionFromFontSet3(fontSet As IDWriteFontSet, <Out> ByRef fontCollection As IDWriteFontCollection1) As HRESULT
        <PreserveSig>
        Overloads Function GetSystemFontCollection3(includeDownloadableFonts As Boolean, <Out> ByRef fontCollection As IDWriteFontCollection1, Optional checkForUpdates As Boolean = False) As HRESULT
        <PreserveSig>
        Overloads Function GetFontDownloadQueue(<Out> ByRef fontDownloadQueue As IDWriteFontDownloadQueue) As HRESULT
#End Region

        <PreserveSig>
        Function TranslateColorGlyphRun4(baselineOrigin As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, glyphRunDescription As IntPtr, desiredGlyphImageFormats As DWRITE_GLYPH_IMAGE_FORMATS, measuringMode As DWRITE_MEASURING_MODE, worldAndDpiTransform As IntPtr, colorPaletteIndex As UInteger, <Out> ByRef colorLayers As IDWriteColorGlyphRunEnumerator1) As HRESULT

        <PreserveSig>
        Function ComputeGlyphOrigins(ByRef glyphRun As DWRITE_GLYPH_RUN, measuringMode As DWRITE_MEASURING_MODE, baselineOrigin As D2D1_POINT_2F, worldAndDpiTransform As DWRITE_MATRIX, <Out> ByRef glyphOrigins As D2D1_POINT_2F) As HRESULT
        <PreserveSig>
        Function ComputeGlyphOrigins(ByRef glyphRun As DWRITE_GLYPH_RUN, baselineOrigin As D2D1_POINT_2F, <Out> ByRef glyphOrigins As D2D1_POINT_2F) As HRESULT
    End Interface

    <ComImport>
    <Guid("958DB99A-BE2A-4F09-AF7D-65189803D1D3")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFactory5
        Inherits IDWriteFactory4
#Region "IDWriteFactory4"
#Region "IDWriteFactory3"
#Region "IDWriteFactory2"
#Region "IDWriteFactory1"
#Region "IDWriteFactory"
        Overloads Function GetSystemFontCollection(<Out> ByRef fontCollection As IDWriteFontCollection, Optional checkForUpdates As Boolean = False) As HRESULT
        <PreserveSig>
        Overloads Function CreateCustomFontCollection(collectionLoader As IDWriteFontCollectionLoader, collectionKey As IntPtr, collectionKeySize As Integer, <Out> ByRef fontCollection As IDWriteFontCollection) As HRESULT
        <PreserveSig>
        Overloads Function RegisterFontCollectionLoader(fontCollectionLoader As IDWriteFontCollectionLoader) As HRESULT
        Overloads Function UnregisterFontCollectionLoader(fontCollectionLoader As IDWriteFontCollectionLoader) As HRESULT
        'new HRESULT CreateFontFileReference(string filePath, System.Runtime.InteropServices.ComTypes.FILETIME lastWriteTime, out IDWriteFontFile fontFile);
        <PreserveSig>
        Overloads Function CreateFontFileReference(
        <MarshalAs(UnmanagedType.LPWStr)> filePath As String, lastWriteTime As IntPtr, <Out> ByRef fontFile As IDWriteFontFile) As HRESULT
        <PreserveSig>
        Overloads Function CreateCustomFontFileReference(fontFileReferenceKey As IntPtr, fontFileReferenceKeySize As UInteger, fontFileLoader As IDWriteFontFileLoader, <Out> ByRef fontFile As IDWriteFontFile) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontFace(fontFaceType As DWRITE_FONT_FACE_TYPE, numberOfFiles As Integer,
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> fontFiles As IDWriteFontFile(), faceIndex As Integer, fontFaceSimulationFlags As DWRITE_FONT_SIMULATIONS, <Out> ByRef fontFace As IDWriteFontFace) As HRESULT
        <PreserveSig>
        Overloads Function CreateRenderingParams(<Out> ByRef renderingParams As IDWriteRenderingParams) As HRESULT
        <PreserveSig>
        Overloads Function CreateMonitorRenderingParams(monitor As IntPtr, <Out> ByRef renderingParams As IDWriteRenderingParams) As HRESULT
        <PreserveSig>
        Overloads Function CreateCustomRenderingParams(gamma As Single, enhancedContrast As Single, clearTypeLevel As Single, pixelGeometry As DWRITE_PIXEL_GEOMETRY, renderingMode As DWRITE_RENDERING_MODE, <Out> ByRef renderingParams As IDWriteRenderingParams) As HRESULT
        <PreserveSig>
        Overloads Function RegisterFontFileLoader(fontFileLoader As IDWriteFontFileLoader) As HRESULT
        Overloads Function UnregisterFontFileLoader(fontFileLoader As IDWriteFontFileLoader) As HRESULT
        <PreserveSig>
        Overloads Function CreateTextFormat(
        <MarshalAs(UnmanagedType.LPWStr)> fontFamilyName As String, fontCollection As IDWriteFontCollection, fontWeight As DWRITE_FONT_WEIGHT, fontStyle As DWRITE_FONT_STYLE, fontStretch As DWRITE_FONT_STRETCH, fontSize As Single,
            <MarshalAs(UnmanagedType.LPWStr)> localeName As String, <Out> ByRef textFormat As IDWriteTextFormat) As HRESULT
        <PreserveSig>
        Overloads Function CreateTypography(<Out> ByRef typography As IDWriteTypography) As HRESULT
        <PreserveSig>
        Overloads Function GetGdiInterop(<Out> ByRef gdiInterop As IDWriteGdiInterop) As HRESULT
        <PreserveSig>
        Overloads Function CreateTextLayout(
        <MarshalAs(UnmanagedType.LPWStr)> str As String, stringLength As Integer, textFormat As IDWriteTextFormat, maxWidth As Single, maxHeight As Single, <Out> ByRef textLayout As IDWriteTextLayout) As HRESULT
        <PreserveSig>
        Overloads Function CreateGdiCompatibleTextLayout(
            <MarshalAs(UnmanagedType.LPWStr)> str As String, stringLength As Integer, textFormat As IDWriteTextFormat, layoutWidth As Single, layoutHeight As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, useGdiNatural As Boolean, <Out> ByRef textLayout As IDWriteTextLayout) As HRESULT
        <PreserveSig>
        Overloads Function CreateEllipsisTrimmingSign(textFormat As IDWriteTextFormat, <Out> ByRef trimmingSign As IDWriteInlineObject) As HRESULT
        <PreserveSig>
        Overloads Function CreateTextAnalyzer(<Out> ByRef textAnalyzer As IDWriteTextAnalyzer) As HRESULT
        <PreserveSig>
        Overloads Function CreateNumberSubstitution(substitutionMethod As DWRITE_NUMBER_SUBSTITUTION_METHOD, localeName As String, ignoreUserOverride As Boolean, <Out> ByRef numberSubstitution As IDWriteNumberSubstitution) As HRESULT
        <PreserveSig>
        Overloads Function CreateGlyphRunAnalysis(ByRef glyphRun As DWRITE_GLYPH_RUN, pixelsPerDip As Single, transform As DWRITE_MATRIX, renderingMode As DWRITE_RENDERING_MODE, measuringMode As DWRITE_MEASURING_MODE, baselineOriginX As Single, baselineOriginY As Single, <Out> ByRef glyphRunAnalysis As IDWriteGlyphRunAnalysis) As HRESULT
#End Region

        Overloads Function GetEudcFontCollection(<Out> ByRef fontCollection As IDWriteFontCollection, Optional checkForUpdates As Boolean = False) As HRESULT
        <PreserveSig>
        Overloads Function CreateCustomRenderingParams1(gamma As Single, enhancedContrast As Single, enhancedContrastGrayscale As Single, clearTypeLevel As Single, pixelGeometry As DWRITE_PIXEL_GEOMETRY, renderingMode As DWRITE_RENDERING_MODE, <Out> ByRef renderingParams As IDWriteRenderingParams1) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function GetSystemFontFallback(<Out> ByRef fontFallback As IDWriteFontFallback) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontFallbackBuilder(<Out> ByRef fontFallbackBuilder As IDWriteFontFallbackBuilder) As HRESULT
        <PreserveSig>
        Overloads Function TranslateColorGlyphRun(baselineOriginX As D2D1_POINT_2F, baselineOriginY As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, glyphRunDescription As IntPtr, measuringMode As DWRITE_MEASURING_MODE, worldToDeviceTransform As IntPtr, colorPaletteIndex As UInteger, <Out> ByRef colorLayers As IDWriteColorGlyphRunEnumerator) As HRESULT
        <PreserveSig>
        Overloads Function CreateCustomRenderingParams2(gamma As Single, enhancedContrast As Single, grayscaleEnhancedContrast As Single, clearTypeLevel As Single, pixelGeometry As DWRITE_PIXEL_GEOMETRY, renderingMode As DWRITE_RENDERING_MODE, gridFitMode As DWRITE_GRID_FIT_MODE, <Out> ByRef renderingParams As IDWriteRenderingParams2) As HRESULT
        <PreserveSig>
        Overloads Function CreateGlyphRunAnalysis2(ByRef glyphRun As DWRITE_GLYPH_RUN, transform As DWRITE_MATRIX, renderingMode As DWRITE_RENDERING_MODE, measuringMode As DWRITE_MEASURING_MODE, gridFitMode As DWRITE_GRID_FIT_MODE, antialiasMode As DWRITE_TEXT_ANTIALIAS_MODE, baselineOriginX As Single, baselineOriginY As Single, <Out> ByRef glyphRunAnalysis As IDWriteGlyphRunAnalysis) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateGlyphRunAnalysis3(ByRef glyphRun As DWRITE_GLYPH_RUN, transform As DWRITE_MATRIX, renderingMode As DWRITE_RENDERING_MODE1, measuringMode As DWRITE_MEASURING_MODE, gridFitMode As DWRITE_GRID_FIT_MODE, antialiasMode As DWRITE_TEXT_ANTIALIAS_MODE, baselineOriginX As Single, baselineOriginY As Single, <Out> ByRef glyphRunAnalysis As IDWriteGlyphRunAnalysis) As HRESULT
        <PreserveSig>
        Overloads Function CreateCustomRenderingParams3(gamma As Single, enhancedContrast As Single, grayscaleEnhancedContrast As Single, clearTypeLevel As Single, pixelGeometry As DWRITE_PIXEL_GEOMETRY, renderingMode As DWRITE_RENDERING_MODE1, gridFitMode As DWRITE_GRID_FIT_MODE, <Out> ByRef renderingParams As IDWriteRenderingParams3) As HRESULT
        'new HRESULT CreateFontFaceReference3(string filePath, System.Runtime.InteropServices.ComTypes.FILETIME lastWriteTime, uint faceIndex, DWRITE_FONT_SIMULATIONS fontSimulations, out IDWriteFontFaceReference fontFaceReference);
        <PreserveSig>
        Overloads Function CreateFontFaceReference3(filePath As String, lastWriteTime As IntPtr, faceIndex As UInteger, fontSimulations As DWRITE_FONT_SIMULATIONS, <Out> ByRef fontFaceReference As IDWriteFontFaceReference) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontFaceReference3(fontFile As IDWriteFontFile, faceIndex As UInteger, fontSimulations As DWRITE_FONT_SIMULATIONS, <Out> ByRef fontFaceReference As IDWriteFontFaceReference) As HRESULT
        <PreserveSig>
        Overloads Function GetSystemFontSet3(<Out> ByRef fontSet As IDWriteFontSet) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontSetBuilder3(<Out> ByRef fontSetBuilder As IDWriteFontSetBuilder) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontCollectionFromFontSet3(fontSet As IDWriteFontSet, <Out> ByRef fontCollection As IDWriteFontCollection1) As HRESULT
        <PreserveSig>
        Overloads Function GetSystemFontCollection3(includeDownloadableFonts As Boolean, <Out> ByRef fontCollection As IDWriteFontCollection1, Optional checkForUpdates As Boolean = False) As HRESULT
        <PreserveSig>
        Overloads Function GetFontDownloadQueue(<Out> ByRef fontDownloadQueue As IDWriteFontDownloadQueue) As HRESULT
#End Region

        Overloads Function TranslateColorGlyphRun4(baselineOrigin As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, glyphRunDescription As IntPtr, desiredGlyphImageFormats As DWRITE_GLYPH_IMAGE_FORMATS, measuringMode As DWRITE_MEASURING_MODE, worldAndDpiTransform As IntPtr, colorPaletteIndex As UInteger, <Out> ByRef colorLayers As IDWriteColorGlyphRunEnumerator1) As HRESULT
        Overloads Function ComputeGlyphOrigins(ByRef glyphRun As DWRITE_GLYPH_RUN, measuringMode As DWRITE_MEASURING_MODE, baselineOrigin As D2D1_POINT_2F, worldAndDpiTransform As DWRITE_MATRIX, <Out> ByRef glyphOrigins As D2D1_POINT_2F) As HRESULT
        Overloads Function ComputeGlyphOrigins(ByRef glyphRun As DWRITE_GLYPH_RUN, baselineOrigin As D2D1_POINT_2F, <Out> ByRef glyphOrigins As D2D1_POINT_2F) As HRESULT
#End Region

        <PreserveSig>
        Function CreateFontSetBuilder5(<Out> ByRef fontSetBuilder As IDWriteFontSetBuilder1) As HRESULT
        <PreserveSig>
        Function CreateInMemoryFontFileLoader(<Out> ByRef newLoader As IDWriteInMemoryFontFileLoader) As HRESULT
        <PreserveSig>
        Function CreateHttpFontFileLoader(referrerUrl As String, extraHeaders As String, <Out> ByRef newLoader As IDWriteRemoteFontFileLoader) As HRESULT
        Function AnalyzeContainerType(fileData As IntPtr, fileDataSize As UInteger) As DWRITE_CONTAINER_TYPE
        <PreserveSig>
        Function UnpackFontFile(containerType As DWRITE_CONTAINER_TYPE, fileData As IntPtr, fileDataSize As UInteger, <Out> ByRef unpackedFontStream As IDWriteFontFileStream) As HRESULT
    End Interface

    <ComImport>
    <Guid("F3744D80-21F7-42EB-B35D-995BC72FC223")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFactory6
        Inherits IDWriteFactory5
#Region "IDWriteFactory5"
#Region "IDWriteFactory4"
#Region "IDWriteFactory3"
#Region "IDWriteFactory2"
#Region "IDWriteFactory1"
#Region "IDWriteFactory"
        Overloads Function GetSystemFontCollection(<Out> ByRef fontCollection As IDWriteFontCollection, Optional checkForUpdates As Boolean = False) As HRESULT
        <PreserveSig>
        Overloads Function CreateCustomFontCollection(collectionLoader As IDWriteFontCollectionLoader, collectionKey As IntPtr, collectionKeySize As Integer, <Out> ByRef fontCollection As IDWriteFontCollection) As HRESULT
        <PreserveSig>
        Overloads Function RegisterFontCollectionLoader(fontCollectionLoader As IDWriteFontCollectionLoader) As HRESULT
        Overloads Function UnregisterFontCollectionLoader(fontCollectionLoader As IDWriteFontCollectionLoader) As HRESULT
        'new HRESULT CreateFontFileReference(string filePath, System.Runtime.InteropServices.ComTypes.FILETIME lastWriteTime, out IDWriteFontFile fontFile);
        <PreserveSig>
        Overloads Function CreateFontFileReference(
        <MarshalAs(UnmanagedType.LPWStr)> filePath As String, lastWriteTime As IntPtr, <Out> ByRef fontFile As IDWriteFontFile) As HRESULT
        <PreserveSig>
        Overloads Function CreateCustomFontFileReference(fontFileReferenceKey As IntPtr, fontFileReferenceKeySize As UInteger, fontFileLoader As IDWriteFontFileLoader, <Out> ByRef fontFile As IDWriteFontFile) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontFace(fontFaceType As DWRITE_FONT_FACE_TYPE, numberOfFiles As Integer,
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> fontFiles As IDWriteFontFile(), faceIndex As Integer, fontFaceSimulationFlags As DWRITE_FONT_SIMULATIONS, <Out> ByRef fontFace As IDWriteFontFace) As HRESULT
        <PreserveSig>
        Overloads Function CreateRenderingParams(<Out> ByRef renderingParams As IDWriteRenderingParams) As HRESULT
        <PreserveSig>
        Overloads Function CreateMonitorRenderingParams(monitor As IntPtr, <Out> ByRef renderingParams As IDWriteRenderingParams) As HRESULT
        <PreserveSig>
        Overloads Function CreateCustomRenderingParams(gamma As Single, enhancedContrast As Single, clearTypeLevel As Single, pixelGeometry As DWRITE_PIXEL_GEOMETRY, renderingMode As DWRITE_RENDERING_MODE, <Out> ByRef renderingParams As IDWriteRenderingParams) As HRESULT
        <PreserveSig>
        Overloads Function RegisterFontFileLoader(fontFileLoader As IDWriteFontFileLoader) As HRESULT
        Overloads Function UnregisterFontFileLoader(fontFileLoader As IDWriteFontFileLoader) As HRESULT
        <PreserveSig>
        Overloads Function CreateTextFormat(
        <MarshalAs(UnmanagedType.LPWStr)> fontFamilyName As String, fontCollection As IDWriteFontCollection, fontWeight As DWRITE_FONT_WEIGHT, fontStyle As DWRITE_FONT_STYLE, fontStretch As DWRITE_FONT_STRETCH, fontSize As Single,
            <MarshalAs(UnmanagedType.LPWStr)> localeName As String, <Out> ByRef textFormat As IDWriteTextFormat) As HRESULT
        <PreserveSig>
        Overloads Function CreateTypography(<Out> ByRef typography As IDWriteTypography) As HRESULT
        <PreserveSig>
        Overloads Function GetGdiInterop(<Out> ByRef gdiInterop As IDWriteGdiInterop) As HRESULT
        <PreserveSig>
        Overloads Function CreateTextLayout(
        <MarshalAs(UnmanagedType.LPWStr)> str As String, stringLength As Integer, textFormat As IDWriteTextFormat, maxWidth As Single, maxHeight As Single, <Out> ByRef textLayout As IDWriteTextLayout) As HRESULT
        <PreserveSig>
        Overloads Function CreateGdiCompatibleTextLayout(
            <MarshalAs(UnmanagedType.LPWStr)> str As String, stringLength As Integer, textFormat As IDWriteTextFormat, layoutWidth As Single, layoutHeight As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, useGdiNatural As Boolean, <Out> ByRef textLayout As IDWriteTextLayout) As HRESULT
        <PreserveSig>
        Overloads Function CreateEllipsisTrimmingSign(textFormat As IDWriteTextFormat, <Out> ByRef trimmingSign As IDWriteInlineObject) As HRESULT
        <PreserveSig>
        Overloads Function CreateTextAnalyzer(<Out> ByRef textAnalyzer As IDWriteTextAnalyzer) As HRESULT
        <PreserveSig>
        Overloads Function CreateNumberSubstitution(substitutionMethod As DWRITE_NUMBER_SUBSTITUTION_METHOD, localeName As String, ignoreUserOverride As Boolean, <Out> ByRef numberSubstitution As IDWriteNumberSubstitution) As HRESULT
        <PreserveSig>
        Overloads Function CreateGlyphRunAnalysis(ByRef glyphRun As DWRITE_GLYPH_RUN, pixelsPerDip As Single, transform As DWRITE_MATRIX, renderingMode As DWRITE_RENDERING_MODE, measuringMode As DWRITE_MEASURING_MODE, baselineOriginX As Single, baselineOriginY As Single, <Out> ByRef glyphRunAnalysis As IDWriteGlyphRunAnalysis) As HRESULT
#End Region

        Overloads Function GetEudcFontCollection(<Out> ByRef fontCollection As IDWriteFontCollection, Optional checkForUpdates As Boolean = False) As HRESULT
        <PreserveSig>
        Overloads Function CreateCustomRenderingParams1(gamma As Single, enhancedContrast As Single, enhancedContrastGrayscale As Single, clearTypeLevel As Single, pixelGeometry As DWRITE_PIXEL_GEOMETRY, renderingMode As DWRITE_RENDERING_MODE, <Out> ByRef renderingParams As IDWriteRenderingParams1) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function GetSystemFontFallback(<Out> ByRef fontFallback As IDWriteFontFallback) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontFallbackBuilder(<Out> ByRef fontFallbackBuilder As IDWriteFontFallbackBuilder) As HRESULT
        <PreserveSig>
        Overloads Function TranslateColorGlyphRun(baselineOriginX As D2D1_POINT_2F, baselineOriginY As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, glyphRunDescription As IntPtr, measuringMode As DWRITE_MEASURING_MODE, worldToDeviceTransform As IntPtr, colorPaletteIndex As UInteger, <Out> ByRef colorLayers As IDWriteColorGlyphRunEnumerator) As HRESULT
        <PreserveSig>
        Overloads Function CreateCustomRenderingParams2(gamma As Single, enhancedContrast As Single, grayscaleEnhancedContrast As Single, clearTypeLevel As Single, pixelGeometry As DWRITE_PIXEL_GEOMETRY, renderingMode As DWRITE_RENDERING_MODE, gridFitMode As DWRITE_GRID_FIT_MODE, <Out> ByRef renderingParams As IDWriteRenderingParams2) As HRESULT
        <PreserveSig>
        Overloads Function CreateGlyphRunAnalysis2(ByRef glyphRun As DWRITE_GLYPH_RUN, transform As DWRITE_MATRIX, renderingMode As DWRITE_RENDERING_MODE, measuringMode As DWRITE_MEASURING_MODE, gridFitMode As DWRITE_GRID_FIT_MODE, antialiasMode As DWRITE_TEXT_ANTIALIAS_MODE, baselineOriginX As Single, baselineOriginY As Single, <Out> ByRef glyphRunAnalysis As IDWriteGlyphRunAnalysis) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateGlyphRunAnalysis3(ByRef glyphRun As DWRITE_GLYPH_RUN, transform As DWRITE_MATRIX, renderingMode As DWRITE_RENDERING_MODE1, measuringMode As DWRITE_MEASURING_MODE, gridFitMode As DWRITE_GRID_FIT_MODE, antialiasMode As DWRITE_TEXT_ANTIALIAS_MODE, baselineOriginX As Single, baselineOriginY As Single, <Out> ByRef glyphRunAnalysis As IDWriteGlyphRunAnalysis) As HRESULT
        <PreserveSig>
        Overloads Function CreateCustomRenderingParams3(gamma As Single, enhancedContrast As Single, grayscaleEnhancedContrast As Single, clearTypeLevel As Single, pixelGeometry As DWRITE_PIXEL_GEOMETRY, renderingMode As DWRITE_RENDERING_MODE1, gridFitMode As DWRITE_GRID_FIT_MODE, <Out> ByRef renderingParams As IDWriteRenderingParams3) As HRESULT
        'new HRESULT CreateFontFaceReference3(string filePath, System.Runtime.InteropServices.ComTypes.FILETIME lastWriteTime, uint faceIndex, DWRITE_FONT_SIMULATIONS fontSimulations, out IDWriteFontFaceReference fontFaceReference);
        <PreserveSig>
        Overloads Function CreateFontFaceReference3(filePath As String, lastWriteTime As IntPtr, faceIndex As UInteger, fontSimulations As DWRITE_FONT_SIMULATIONS, <Out> ByRef fontFaceReference As IDWriteFontFaceReference) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontFaceReference3(fontFile As IDWriteFontFile, faceIndex As UInteger, fontSimulations As DWRITE_FONT_SIMULATIONS, <Out> ByRef fontFaceReference As IDWriteFontFaceReference) As HRESULT
        <PreserveSig>
        Overloads Function GetSystemFontSet3(<Out> ByRef fontSet As IDWriteFontSet) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontSetBuilder3(<Out> ByRef fontSetBuilder As IDWriteFontSetBuilder) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontCollectionFromFontSet3(fontSet As IDWriteFontSet, <Out> ByRef fontCollection As IDWriteFontCollection1) As HRESULT
        <PreserveSig>
        Overloads Function GetSystemFontCollection3(includeDownloadableFonts As Boolean, <Out> ByRef fontCollection As IDWriteFontCollection1, Optional checkForUpdates As Boolean = False) As HRESULT
        <PreserveSig>
        Overloads Function GetFontDownloadQueue(<Out> ByRef fontDownloadQueue As IDWriteFontDownloadQueue) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function TranslateColorGlyphRun4(baselineOrigin As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, glyphRunDescription As IntPtr, desiredGlyphImageFormats As DWRITE_GLYPH_IMAGE_FORMATS, measuringMode As DWRITE_MEASURING_MODE, worldAndDpiTransform As IntPtr, colorPaletteIndex As UInteger, <Out> ByRef colorLayers As IDWriteColorGlyphRunEnumerator1) As HRESULT
        <PreserveSig>
        Overloads Function ComputeGlyphOrigins(ByRef glyphRun As DWRITE_GLYPH_RUN, measuringMode As DWRITE_MEASURING_MODE, baselineOrigin As D2D1_POINT_2F, worldAndDpiTransform As DWRITE_MATRIX, <Out> ByRef glyphOrigins As D2D1_POINT_2F) As HRESULT
        <PreserveSig>
        Overloads Function ComputeGlyphOrigins(ByRef glyphRun As DWRITE_GLYPH_RUN, baselineOrigin As D2D1_POINT_2F, <Out> ByRef glyphOrigins As D2D1_POINT_2F) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateFontSetBuilder5(<Out> ByRef fontSetBuilder As IDWriteFontSetBuilder1) As HRESULT
        <PreserveSig>
        Overloads Function CreateInMemoryFontFileLoader(<Out> ByRef newLoader As IDWriteInMemoryFontFileLoader) As HRESULT
        <PreserveSig>
        Overloads Function CreateHttpFontFileLoader(referrerUrl As String, extraHeaders As String, <Out> ByRef newLoader As IDWriteRemoteFontFileLoader) As HRESULT
        <PreserveSig>
        Overloads Function AnalyzeContainerType(fileData As IntPtr, fileDataSize As UInteger) As DWRITE_CONTAINER_TYPE
        <PreserveSig>
        Overloads Function UnpackFontFile(containerType As DWRITE_CONTAINER_TYPE, fileData As IntPtr, fileDataSize As UInteger, <Out> ByRef unpackedFontStream As IDWriteFontFileStream) As HRESULT
#End Region

        <PreserveSig>
        Function CreateFontFaceReference6(fontFile As IDWriteFontFile, faceIndex As UInteger, fontSimulations As DWRITE_FONT_SIMULATIONS, fontAxisValues As DWRITE_FONT_AXIS_VALUE, fontAxisValueCount As UInteger, <Out> ByRef fontFaceReference As IDWriteFontFaceReference1) As HRESULT
        <PreserveSig>
        Function CreateFontResource(fontFile As IDWriteFontFile, faceIndex As UInteger, <Out> ByRef fontResource As IDWriteFontResource) As HRESULT
        <PreserveSig>
        Function GetSystemFontSet6(includeDownloadableFonts As Boolean, <Out> ByRef fontSet As IDWriteFontSet1) As HRESULT
        <PreserveSig>
        Function GetSystemFontCollection6(includeDownloadableFonts As Boolean, fontFamilyModel As DWRITE_FONT_FAMILY_MODEL, <Out> ByRef fontCollection As IDWriteFontCollection2) As HRESULT
        <PreserveSig>
        Function CreateFontCollectionFromFontSet6(fontSet As IDWriteFontSet, fontFamilyModel As DWRITE_FONT_FAMILY_MODEL, <Out> ByRef fontCollection As IDWriteFontCollection2) As HRESULT
        <PreserveSig>
        Function CreateFontSetBuilder6(<Out> ByRef fontSetBuilder As IDWriteFontSetBuilder2) As HRESULT
        <PreserveSig>
        Function CreateTextFormat6(fontFamilyName As String, fontCollection As IDWriteFontCollection, fontAxisValues As DWRITE_FONT_AXIS_VALUE, fontAxisValueCount As UInteger, fontSize As Single, localeName As String, <Out> ByRef textFormat As IDWriteTextFormat3) As HRESULT
    End Interface

    <ComImport>
    <Guid("35D0E0B3-9076-4D2E-A016-A91B568A06B4")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFactory7
        Inherits IDWriteFactory6
#Region "IDWriteFactory6"
#Region "IDWriteFactory5"
#Region "IDWriteFactory4"
#Region "IDWriteFactory3"
#Region "IDWriteFactory2"
#Region "IDWriteFactory1"
#Region "IDWriteFactory"
        Overloads Function GetSystemFontCollection(<Out> ByRef fontCollection As IDWriteFontCollection, Optional checkForUpdates As Boolean = False) As HRESULT
        <PreserveSig>
        Overloads Function CreateCustomFontCollection(collectionLoader As IDWriteFontCollectionLoader, collectionKey As IntPtr, collectionKeySize As Integer, <Out> ByRef fontCollection As IDWriteFontCollection) As HRESULT
        <PreserveSig>
        Overloads Function RegisterFontCollectionLoader(fontCollectionLoader As IDWriteFontCollectionLoader) As HRESULT
        Overloads Function UnregisterFontCollectionLoader(fontCollectionLoader As IDWriteFontCollectionLoader) As HRESULT
        'new HRESULT CreateFontFileReference(string filePath, System.Runtime.InteropServices.ComTypes.FILETIME lastWriteTime, out IDWriteFontFile fontFile);
        <PreserveSig>
        Overloads Function CreateFontFileReference(
        <MarshalAs(UnmanagedType.LPWStr)> filePath As String, lastWriteTime As IntPtr, <Out> ByRef fontFile As IDWriteFontFile) As HRESULT
        <PreserveSig>
        Overloads Function CreateCustomFontFileReference(fontFileReferenceKey As IntPtr, fontFileReferenceKeySize As UInteger, fontFileLoader As IDWriteFontFileLoader, <Out> ByRef fontFile As IDWriteFontFile) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontFace(fontFaceType As DWRITE_FONT_FACE_TYPE, numberOfFiles As Integer,
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> fontFiles As IDWriteFontFile(), faceIndex As Integer, fontFaceSimulationFlags As DWRITE_FONT_SIMULATIONS, <Out> ByRef fontFace As IDWriteFontFace) As HRESULT
        <PreserveSig>
        Overloads Function CreateRenderingParams(<Out> ByRef renderingParams As IDWriteRenderingParams) As HRESULT
        <PreserveSig>
        Overloads Function CreateMonitorRenderingParams(monitor As IntPtr, <Out> ByRef renderingParams As IDWriteRenderingParams) As HRESULT
        <PreserveSig>
        Overloads Function CreateCustomRenderingParams(gamma As Single, enhancedContrast As Single, clearTypeLevel As Single, pixelGeometry As DWRITE_PIXEL_GEOMETRY, renderingMode As DWRITE_RENDERING_MODE, <Out> ByRef renderingParams As IDWriteRenderingParams) As HRESULT
        <PreserveSig>
        Overloads Function RegisterFontFileLoader(fontFileLoader As IDWriteFontFileLoader) As HRESULT
        Overloads Function UnregisterFontFileLoader(fontFileLoader As IDWriteFontFileLoader) As HRESULT
        <PreserveSig>
        Overloads Function CreateTextFormat(
        <MarshalAs(UnmanagedType.LPWStr)> fontFamilyName As String, fontCollection As IDWriteFontCollection, fontWeight As DWRITE_FONT_WEIGHT, fontStyle As DWRITE_FONT_STYLE, fontStretch As DWRITE_FONT_STRETCH, fontSize As Single,
            <MarshalAs(UnmanagedType.LPWStr)> localeName As String, <Out> ByRef textFormat As IDWriteTextFormat) As HRESULT
        <PreserveSig>
        Overloads Function CreateTypography(<Out> ByRef typography As IDWriteTypography) As HRESULT
        <PreserveSig>
        Overloads Function GetGdiInterop(<Out> ByRef gdiInterop As IDWriteGdiInterop) As HRESULT
        <PreserveSig>
        Overloads Function CreateTextLayout(
        <MarshalAs(UnmanagedType.LPWStr)> str As String, stringLength As Integer, textFormat As IDWriteTextFormat, maxWidth As Single, maxHeight As Single, <Out> ByRef textLayout As IDWriteTextLayout) As HRESULT
        <PreserveSig>
        Overloads Function CreateGdiCompatibleTextLayout(
            <MarshalAs(UnmanagedType.LPWStr)> str As String, stringLength As Integer, textFormat As IDWriteTextFormat, layoutWidth As Single, layoutHeight As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, useGdiNatural As Boolean, <Out> ByRef textLayout As IDWriteTextLayout) As HRESULT
        <PreserveSig>
        Overloads Function CreateEllipsisTrimmingSign(textFormat As IDWriteTextFormat, <Out> ByRef trimmingSign As IDWriteInlineObject) As HRESULT
        <PreserveSig>
        Overloads Function CreateTextAnalyzer(<Out> ByRef textAnalyzer As IDWriteTextAnalyzer) As HRESULT
        <PreserveSig>
        Overloads Function CreateNumberSubstitution(substitutionMethod As DWRITE_NUMBER_SUBSTITUTION_METHOD, localeName As String, ignoreUserOverride As Boolean, <Out> ByRef numberSubstitution As IDWriteNumberSubstitution) As HRESULT
        <PreserveSig>
        Overloads Function CreateGlyphRunAnalysis(ByRef glyphRun As DWRITE_GLYPH_RUN, pixelsPerDip As Single, transform As DWRITE_MATRIX, renderingMode As DWRITE_RENDERING_MODE, measuringMode As DWRITE_MEASURING_MODE, baselineOriginX As Single, baselineOriginY As Single, <Out> ByRef glyphRunAnalysis As IDWriteGlyphRunAnalysis) As HRESULT
#End Region

        Overloads Function GetEudcFontCollection(<Out> ByRef fontCollection As IDWriteFontCollection, Optional checkForUpdates As Boolean = False) As HRESULT
        <PreserveSig>
        Overloads Function CreateCustomRenderingParams1(gamma As Single, enhancedContrast As Single, enhancedContrastGrayscale As Single, clearTypeLevel As Single, pixelGeometry As DWRITE_PIXEL_GEOMETRY, renderingMode As DWRITE_RENDERING_MODE, <Out> ByRef renderingParams As IDWriteRenderingParams1) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function GetSystemFontFallback(<Out> ByRef fontFallback As IDWriteFontFallback) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontFallbackBuilder(<Out> ByRef fontFallbackBuilder As IDWriteFontFallbackBuilder) As HRESULT
        <PreserveSig>
        Overloads Function TranslateColorGlyphRun(baselineOriginX As D2D1_POINT_2F, baselineOriginY As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, glyphRunDescription As IntPtr, measuringMode As DWRITE_MEASURING_MODE, worldToDeviceTransform As IntPtr, colorPaletteIndex As UInteger, <Out> ByRef colorLayers As IDWriteColorGlyphRunEnumerator) As HRESULT
        <PreserveSig>
        Overloads Function CreateCustomRenderingParams2(gamma As Single, enhancedContrast As Single, grayscaleEnhancedContrast As Single, clearTypeLevel As Single, pixelGeometry As DWRITE_PIXEL_GEOMETRY, renderingMode As DWRITE_RENDERING_MODE, gridFitMode As DWRITE_GRID_FIT_MODE, <Out> ByRef renderingParams As IDWriteRenderingParams2) As HRESULT
        <PreserveSig>
        Overloads Function CreateGlyphRunAnalysis2(ByRef glyphRun As DWRITE_GLYPH_RUN, transform As DWRITE_MATRIX, renderingMode As DWRITE_RENDERING_MODE, measuringMode As DWRITE_MEASURING_MODE, gridFitMode As DWRITE_GRID_FIT_MODE, antialiasMode As DWRITE_TEXT_ANTIALIAS_MODE, baselineOriginX As Single, baselineOriginY As Single, <Out> ByRef glyphRunAnalysis As IDWriteGlyphRunAnalysis) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateGlyphRunAnalysis3(ByRef glyphRun As DWRITE_GLYPH_RUN, transform As DWRITE_MATRIX, renderingMode As DWRITE_RENDERING_MODE1, measuringMode As DWRITE_MEASURING_MODE, gridFitMode As DWRITE_GRID_FIT_MODE, antialiasMode As DWRITE_TEXT_ANTIALIAS_MODE, baselineOriginX As Single, baselineOriginY As Single, <Out> ByRef glyphRunAnalysis As IDWriteGlyphRunAnalysis) As HRESULT
        <PreserveSig>
        Overloads Function CreateCustomRenderingParams3(gamma As Single, enhancedContrast As Single, grayscaleEnhancedContrast As Single, clearTypeLevel As Single, pixelGeometry As DWRITE_PIXEL_GEOMETRY, renderingMode As DWRITE_RENDERING_MODE1, gridFitMode As DWRITE_GRID_FIT_MODE, <Out> ByRef renderingParams As IDWriteRenderingParams3) As HRESULT
        'new HRESULT CreateFontFaceReference3(string filePath, System.Runtime.InteropServices.ComTypes.FILETIME lastWriteTime, uint faceIndex, DWRITE_FONT_SIMULATIONS fontSimulations, out IDWriteFontFaceReference fontFaceReference);
        <PreserveSig>
        Overloads Function CreateFontFaceReference3(filePath As String, lastWriteTime As IntPtr, faceIndex As UInteger, fontSimulations As DWRITE_FONT_SIMULATIONS, <Out> ByRef fontFaceReference As IDWriteFontFaceReference) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontFaceReference3(fontFile As IDWriteFontFile, faceIndex As UInteger, fontSimulations As DWRITE_FONT_SIMULATIONS, <Out> ByRef fontFaceReference As IDWriteFontFaceReference) As HRESULT
        <PreserveSig>
        Overloads Function GetSystemFontSet3(<Out> ByRef fontSet As IDWriteFontSet) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontSetBuilder3(<Out> ByRef fontSetBuilder As IDWriteFontSetBuilder) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontCollectionFromFontSet3(fontSet As IDWriteFontSet, <Out> ByRef fontCollection As IDWriteFontCollection1) As HRESULT
        <PreserveSig>
        Overloads Function GetSystemFontCollection3(includeDownloadableFonts As Boolean, <Out> ByRef fontCollection As IDWriteFontCollection1, Optional checkForUpdates As Boolean = False) As HRESULT
        <PreserveSig>
        Overloads Function GetFontDownloadQueue(<Out> ByRef fontDownloadQueue As IDWriteFontDownloadQueue) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function TranslateColorGlyphRun4(baselineOrigin As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, glyphRunDescription As IntPtr, desiredGlyphImageFormats As DWRITE_GLYPH_IMAGE_FORMATS, measuringMode As DWRITE_MEASURING_MODE, worldAndDpiTransform As IntPtr, colorPaletteIndex As UInteger, <Out> ByRef colorLayers As IDWriteColorGlyphRunEnumerator1) As HRESULT
        <PreserveSig>
        Overloads Function ComputeGlyphOrigins(ByRef glyphRun As DWRITE_GLYPH_RUN, measuringMode As DWRITE_MEASURING_MODE, baselineOrigin As D2D1_POINT_2F, worldAndDpiTransform As DWRITE_MATRIX, <Out> ByRef glyphOrigins As D2D1_POINT_2F) As HRESULT
        <PreserveSig>
        Overloads Function ComputeGlyphOrigins(ByRef glyphRun As DWRITE_GLYPH_RUN, baselineOrigin As D2D1_POINT_2F, <Out> ByRef glyphOrigins As D2D1_POINT_2F) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateFontSetBuilder5(<Out> ByRef fontSetBuilder As IDWriteFontSetBuilder1) As HRESULT
        <PreserveSig>
        Overloads Function CreateInMemoryFontFileLoader(<Out> ByRef newLoader As IDWriteInMemoryFontFileLoader) As HRESULT
        <PreserveSig>
        Overloads Function CreateHttpFontFileLoader(referrerUrl As String, extraHeaders As String, <Out> ByRef newLoader As IDWriteRemoteFontFileLoader) As HRESULT
        <PreserveSig>
        Overloads Function AnalyzeContainerType(fileData As IntPtr, fileDataSize As UInteger) As DWRITE_CONTAINER_TYPE
        <PreserveSig>
        Overloads Function UnpackFontFile(containerType As DWRITE_CONTAINER_TYPE, fileData As IntPtr, fileDataSize As UInteger, <Out> ByRef unpackedFontStream As IDWriteFontFileStream) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateFontFaceReference6(fontFile As IDWriteFontFile, faceIndex As UInteger, fontSimulations As DWRITE_FONT_SIMULATIONS, fontAxisValues As DWRITE_FONT_AXIS_VALUE, fontAxisValueCount As UInteger, <Out> ByRef fontFaceReference As IDWriteFontFaceReference1) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontResource(fontFile As IDWriteFontFile, faceIndex As UInteger, <Out> ByRef fontResource As IDWriteFontResource) As HRESULT
        <PreserveSig>
        Overloads Function GetSystemFontSet6(includeDownloadableFonts As Boolean, <Out> ByRef fontSet As IDWriteFontSet1) As HRESULT
        <PreserveSig>
        Overloads Function GetSystemFontCollection6(includeDownloadableFonts As Boolean, fontFamilyModel As DWRITE_FONT_FAMILY_MODEL, <Out> ByRef fontCollection As IDWriteFontCollection2) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontCollectionFromFontSet6(fontSet As IDWriteFontSet, fontFamilyModel As DWRITE_FONT_FAMILY_MODEL, <Out> ByRef fontCollection As IDWriteFontCollection2) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontSetBuilder6(<Out> ByRef fontSetBuilder As IDWriteFontSetBuilder2) As HRESULT
        <PreserveSig>
        Overloads Function CreateTextFormat6(fontFamilyName As String, fontCollection As IDWriteFontCollection, fontAxisValues As DWRITE_FONT_AXIS_VALUE, fontAxisValueCount As UInteger, fontSize As Single, localeName As String, <Out> ByRef textFormat As IDWriteTextFormat3) As HRESULT
#End Region

        <PreserveSig>
        Function GetSystemFontSet7(includeDownloadableFonts As Boolean, <Out> ByRef fontSet As IDWriteFontSet2) As HRESULT
        <PreserveSig>
        Function GetSystemFontCollection7(includeDownloadableFonts As Boolean, fontFamilyModel As DWRITE_FONT_FAMILY_MODEL, <Out> ByRef fontCollection As IDWriteFontCollection3) As HRESULT
    End Interface

    <ComImport>
    <Guid("EE0A7FB5-DEF4-4C23-A454-C9C7DC878398")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFactory8
        Inherits IDWriteFactory7
#Region "IDWriteFactory6"
#Region "IDWriteFactory5"
#Region "IDWriteFactory4"
#Region "IDWriteFactory3"
#Region "IDWriteFactory2"
#Region "IDWriteFactory1"
#Region "IDWriteFactory"
        Overloads Function GetSystemFontCollection(<Out> ByRef fontCollection As IDWriteFontCollection, Optional checkForUpdates As Boolean = False) As HRESULT
        <PreserveSig>
        Overloads Function CreateCustomFontCollection(collectionLoader As IDWriteFontCollectionLoader, collectionKey As IntPtr, collectionKeySize As Integer, <Out> ByRef fontCollection As IDWriteFontCollection) As HRESULT
        <PreserveSig>
        Overloads Function RegisterFontCollectionLoader(fontCollectionLoader As IDWriteFontCollectionLoader) As HRESULT
        Overloads Function UnregisterFontCollectionLoader(fontCollectionLoader As IDWriteFontCollectionLoader) As HRESULT
        'new HRESULT CreateFontFileReference(string filePath, System.Runtime.InteropServices.ComTypes.FILETIME lastWriteTime, out IDWriteFontFile fontFile);
        <PreserveSig>
        Overloads Function CreateFontFileReference(
        <MarshalAs(UnmanagedType.LPWStr)> filePath As String, lastWriteTime As IntPtr, <Out> ByRef fontFile As IDWriteFontFile) As HRESULT
        <PreserveSig>
        Overloads Function CreateCustomFontFileReference(fontFileReferenceKey As IntPtr, fontFileReferenceKeySize As UInteger, fontFileLoader As IDWriteFontFileLoader, <Out> ByRef fontFile As IDWriteFontFile) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontFace(fontFaceType As DWRITE_FONT_FACE_TYPE, numberOfFiles As Integer,
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> fontFiles As IDWriteFontFile(), faceIndex As Integer, fontFaceSimulationFlags As DWRITE_FONT_SIMULATIONS, <Out> ByRef fontFace As IDWriteFontFace) As HRESULT
        <PreserveSig>
        Overloads Function CreateRenderingParams(<Out> ByRef renderingParams As IDWriteRenderingParams) As HRESULT
        <PreserveSig>
        Overloads Function CreateMonitorRenderingParams(monitor As IntPtr, <Out> ByRef renderingParams As IDWriteRenderingParams) As HRESULT
        <PreserveSig>
        Overloads Function CreateCustomRenderingParams(gamma As Single, enhancedContrast As Single, clearTypeLevel As Single, pixelGeometry As DWRITE_PIXEL_GEOMETRY, renderingMode As DWRITE_RENDERING_MODE, <Out> ByRef renderingParams As IDWriteRenderingParams) As HRESULT
        <PreserveSig>
        Overloads Function RegisterFontFileLoader(fontFileLoader As IDWriteFontFileLoader) As HRESULT
        Overloads Function UnregisterFontFileLoader(fontFileLoader As IDWriteFontFileLoader) As HRESULT
        <PreserveSig>
        Overloads Function CreateTextFormat(
        <MarshalAs(UnmanagedType.LPWStr)> fontFamilyName As String, fontCollection As IDWriteFontCollection, fontWeight As DWRITE_FONT_WEIGHT, fontStyle As DWRITE_FONT_STYLE, fontStretch As DWRITE_FONT_STRETCH, fontSize As Single,
            <MarshalAs(UnmanagedType.LPWStr)> localeName As String, <Out> ByRef textFormat As IDWriteTextFormat) As HRESULT
        <PreserveSig>
        Overloads Function CreateTypography(<Out> ByRef typography As IDWriteTypography) As HRESULT
        <PreserveSig>
        Overloads Function GetGdiInterop(<Out> ByRef gdiInterop As IDWriteGdiInterop) As HRESULT
        <PreserveSig>
        Overloads Function CreateTextLayout(
        <MarshalAs(UnmanagedType.LPWStr)> str As String, stringLength As Integer, textFormat As IDWriteTextFormat, maxWidth As Single, maxHeight As Single, <Out> ByRef textLayout As IDWriteTextLayout) As HRESULT
        <PreserveSig>
        Overloads Function CreateGdiCompatibleTextLayout(
            <MarshalAs(UnmanagedType.LPWStr)> str As String, stringLength As Integer, textFormat As IDWriteTextFormat, layoutWidth As Single, layoutHeight As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, useGdiNatural As Boolean, <Out> ByRef textLayout As IDWriteTextLayout) As HRESULT
        <PreserveSig>
        Overloads Function CreateEllipsisTrimmingSign(textFormat As IDWriteTextFormat, <Out> ByRef trimmingSign As IDWriteInlineObject) As HRESULT
        <PreserveSig>
        Overloads Function CreateTextAnalyzer(<Out> ByRef textAnalyzer As IDWriteTextAnalyzer) As HRESULT
        <PreserveSig>
        Overloads Function CreateNumberSubstitution(substitutionMethod As DWRITE_NUMBER_SUBSTITUTION_METHOD, localeName As String, ignoreUserOverride As Boolean, <Out> ByRef numberSubstitution As IDWriteNumberSubstitution) As HRESULT
        <PreserveSig>
        Overloads Function CreateGlyphRunAnalysis(ByRef glyphRun As DWRITE_GLYPH_RUN, pixelsPerDip As Single, transform As DWRITE_MATRIX, renderingMode As DWRITE_RENDERING_MODE, measuringMode As DWRITE_MEASURING_MODE, baselineOriginX As Single, baselineOriginY As Single, <Out> ByRef glyphRunAnalysis As IDWriteGlyphRunAnalysis) As HRESULT
#End Region

        Overloads Function GetEudcFontCollection(<Out> ByRef fontCollection As IDWriteFontCollection, Optional checkForUpdates As Boolean = False) As HRESULT
        <PreserveSig>
        Overloads Function CreateCustomRenderingParams1(gamma As Single, enhancedContrast As Single, enhancedContrastGrayscale As Single, clearTypeLevel As Single, pixelGeometry As DWRITE_PIXEL_GEOMETRY, renderingMode As DWRITE_RENDERING_MODE, <Out> ByRef renderingParams As IDWriteRenderingParams1) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function GetSystemFontFallback(<Out> ByRef fontFallback As IDWriteFontFallback) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontFallbackBuilder(<Out> ByRef fontFallbackBuilder As IDWriteFontFallbackBuilder) As HRESULT
        <PreserveSig>
        Overloads Function TranslateColorGlyphRun(baselineOriginX As D2D1_POINT_2F, baselineOriginY As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, glyphRunDescription As IntPtr, measuringMode As DWRITE_MEASURING_MODE, worldToDeviceTransform As IntPtr, colorPaletteIndex As UInteger, <Out> ByRef colorLayers As IDWriteColorGlyphRunEnumerator) As HRESULT
        <PreserveSig>
        Overloads Function CreateCustomRenderingParams2(gamma As Single, enhancedContrast As Single, grayscaleEnhancedContrast As Single, clearTypeLevel As Single, pixelGeometry As DWRITE_PIXEL_GEOMETRY, renderingMode As DWRITE_RENDERING_MODE, gridFitMode As DWRITE_GRID_FIT_MODE, <Out> ByRef renderingParams As IDWriteRenderingParams2) As HRESULT
        <PreserveSig>
        Overloads Function CreateGlyphRunAnalysis2(ByRef glyphRun As DWRITE_GLYPH_RUN, transform As DWRITE_MATRIX, renderingMode As DWRITE_RENDERING_MODE, measuringMode As DWRITE_MEASURING_MODE, gridFitMode As DWRITE_GRID_FIT_MODE, antialiasMode As DWRITE_TEXT_ANTIALIAS_MODE, baselineOriginX As Single, baselineOriginY As Single, <Out> ByRef glyphRunAnalysis As IDWriteGlyphRunAnalysis) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateGlyphRunAnalysis3(ByRef glyphRun As DWRITE_GLYPH_RUN, transform As DWRITE_MATRIX, renderingMode As DWRITE_RENDERING_MODE1, measuringMode As DWRITE_MEASURING_MODE, gridFitMode As DWRITE_GRID_FIT_MODE, antialiasMode As DWRITE_TEXT_ANTIALIAS_MODE, baselineOriginX As Single, baselineOriginY As Single, <Out> ByRef glyphRunAnalysis As IDWriteGlyphRunAnalysis) As HRESULT
        <PreserveSig>
        Overloads Function CreateCustomRenderingParams3(gamma As Single, enhancedContrast As Single, grayscaleEnhancedContrast As Single, clearTypeLevel As Single, pixelGeometry As DWRITE_PIXEL_GEOMETRY, renderingMode As DWRITE_RENDERING_MODE1, gridFitMode As DWRITE_GRID_FIT_MODE, <Out> ByRef renderingParams As IDWriteRenderingParams3) As HRESULT
        'new HRESULT CreateFontFaceReference3(string filePath, System.Runtime.InteropServices.ComTypes.FILETIME lastWriteTime, uint faceIndex, DWRITE_FONT_SIMULATIONS fontSimulations, out IDWriteFontFaceReference fontFaceReference);
        <PreserveSig>
        Overloads Function CreateFontFaceReference3(filePath As String, lastWriteTime As IntPtr, faceIndex As UInteger, fontSimulations As DWRITE_FONT_SIMULATIONS, <Out> ByRef fontFaceReference As IDWriteFontFaceReference) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontFaceReference3(fontFile As IDWriteFontFile, faceIndex As UInteger, fontSimulations As DWRITE_FONT_SIMULATIONS, <Out> ByRef fontFaceReference As IDWriteFontFaceReference) As HRESULT
        <PreserveSig>
        Overloads Function GetSystemFontSet3(<Out> ByRef fontSet As IDWriteFontSet) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontSetBuilder3(<Out> ByRef fontSetBuilder As IDWriteFontSetBuilder) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontCollectionFromFontSet3(fontSet As IDWriteFontSet, <Out> ByRef fontCollection As IDWriteFontCollection1) As HRESULT
        <PreserveSig>
        Overloads Function GetSystemFontCollection3(includeDownloadableFonts As Boolean, <Out> ByRef fontCollection As IDWriteFontCollection1, Optional checkForUpdates As Boolean = False) As HRESULT
        <PreserveSig>
        Overloads Function GetFontDownloadQueue(<Out> ByRef fontDownloadQueue As IDWriteFontDownloadQueue) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function TranslateColorGlyphRun4(baselineOrigin As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, glyphRunDescription As IntPtr, desiredGlyphImageFormats As DWRITE_GLYPH_IMAGE_FORMATS, measuringMode As DWRITE_MEASURING_MODE, worldAndDpiTransform As IntPtr, colorPaletteIndex As UInteger, <Out> ByRef colorLayers As IDWriteColorGlyphRunEnumerator1) As HRESULT
        <PreserveSig>
        Overloads Function ComputeGlyphOrigins(ByRef glyphRun As DWRITE_GLYPH_RUN, measuringMode As DWRITE_MEASURING_MODE, baselineOrigin As D2D1_POINT_2F, worldAndDpiTransform As DWRITE_MATRIX, <Out> ByRef glyphOrigins As D2D1_POINT_2F) As HRESULT
        <PreserveSig>
        Overloads Function ComputeGlyphOrigins(ByRef glyphRun As DWRITE_GLYPH_RUN, baselineOrigin As D2D1_POINT_2F, <Out> ByRef glyphOrigins As D2D1_POINT_2F) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateFontSetBuilder5(<Out> ByRef fontSetBuilder As IDWriteFontSetBuilder1) As HRESULT
        <PreserveSig>
        Overloads Function CreateInMemoryFontFileLoader(<Out> ByRef newLoader As IDWriteInMemoryFontFileLoader) As HRESULT
        <PreserveSig>
        Overloads Function CreateHttpFontFileLoader(referrerUrl As String, extraHeaders As String, <Out> ByRef newLoader As IDWriteRemoteFontFileLoader) As HRESULT
        <PreserveSig>
        Overloads Function AnalyzeContainerType(fileData As IntPtr, fileDataSize As UInteger) As DWRITE_CONTAINER_TYPE
        <PreserveSig>
        Overloads Function UnpackFontFile(containerType As DWRITE_CONTAINER_TYPE, fileData As IntPtr, fileDataSize As UInteger, <Out> ByRef unpackedFontStream As IDWriteFontFileStream) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateFontFaceReference6(fontFile As IDWriteFontFile, faceIndex As UInteger, fontSimulations As DWRITE_FONT_SIMULATIONS, fontAxisValues As DWRITE_FONT_AXIS_VALUE, fontAxisValueCount As UInteger, <Out> ByRef fontFaceReference As IDWriteFontFaceReference1) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontResource(fontFile As IDWriteFontFile, faceIndex As UInteger, <Out> ByRef fontResource As IDWriteFontResource) As HRESULT
        <PreserveSig>
        Overloads Function GetSystemFontSet6(includeDownloadableFonts As Boolean, <Out> ByRef fontSet As IDWriteFontSet1) As HRESULT
        <PreserveSig>
        Overloads Function GetSystemFontCollection6(includeDownloadableFonts As Boolean, fontFamilyModel As DWRITE_FONT_FAMILY_MODEL, <Out> ByRef fontCollection As IDWriteFontCollection2) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontCollectionFromFontSet6(fontSet As IDWriteFontSet, fontFamilyModel As DWRITE_FONT_FAMILY_MODEL, <Out> ByRef fontCollection As IDWriteFontCollection2) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontSetBuilder6(<Out> ByRef fontSetBuilder As IDWriteFontSetBuilder2) As HRESULT
        <PreserveSig>
        Overloads Function CreateTextFormat6(fontFamilyName As String, fontCollection As IDWriteFontCollection, fontAxisValues As DWRITE_FONT_AXIS_VALUE, fontAxisValueCount As UInteger, fontSize As Single, localeName As String, <Out> ByRef textFormat As IDWriteTextFormat3) As HRESULT
#End Region

#Region "<IDWriteFactory7>"
        <PreserveSig>
        Overloads Function GetSystemFontSet7(includeDownloadableFonts As Boolean, <Out> ByRef fontSet As IDWriteFontSet2) As HRESULT
        <PreserveSig>
        Overloads Function GetSystemFontCollection7(includeDownloadableFonts As Boolean, fontFamilyModel As DWRITE_FONT_FAMILY_MODEL, <Out> ByRef fontCollection As IDWriteFontCollection3) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function TranslateColorGlyphRun(baselineOrigin As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, glyphRunDescription As IntPtr, desiredGlyphImageFormats As DWRITE_GLYPH_IMAGE_FORMATS, paintFeatureLevel As DWRITE_PAINT_FEATURE_LEVEL, measuringMode As DWRITE_MEASURING_MODE, worldAndDpiTransform As IntPtr, colorPaletteIndex As UInteger, <Out> ByRef colorEnumerator As IDWriteColorGlyphRunEnumerator1) As HRESULT
    End Interface

    <ComImport>
    <Guid("53585141-D9F8-4095-8321-D73CF6BD116B")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontSet
        <PreserveSig>
        Function GetFontCount() As UInteger
        Function GetFontFaceReference(listIndex As UInteger, <Out> ByRef fontFaceReference As IDWriteFontFaceReference) As HRESULT
        Function FindFontFaceReference(fontFaceReference As IDWriteFontFaceReference, <Out> ByRef listIndex As UInteger, <Out> ByRef exists As Boolean) As HRESULT
        Function FindFontFace(fontFace As IDWriteFontFace, <Out> ByRef listIndex As UInteger, <Out> ByRef exists As Boolean) As HRESULT

        Function GetPropertyValues(propertyID As DWRITE_FONT_PROPERTY_ID, <Out> ByRef values As IDWriteStringList) As HRESULT
        Function GetPropertyValues(propertyID As DWRITE_FONT_PROPERTY_ID,
<MarshalAs(UnmanagedType.LPWStr)> preferredLocaleNames As String, <Out> ByRef values As IDWriteStringList) As HRESULT
        Function GetPropertyValues(listIndex As UInteger, propertyId As DWRITE_FONT_PROPERTY_ID, <Out> ByRef exists As Boolean, <Out> ByRef values As IDWriteLocalizedStrings) As HRESULT

        Function GetPropertyOccurrenceCount([property] As DWRITE_FONT_PROPERTY, <Out> ByRef propertyOccurrenceCount As UInteger) As HRESULT
        Function GetMatchingFonts(properties As DWRITE_FONT_PROPERTY, propertyCount As UInteger, <Out> ByRef filteredSet As IDWriteFontSet) As HRESULT
        Function GetMatchingFonts(familyName As String, fontWeight As DWRITE_FONT_WEIGHT, fontStretch As DWRITE_FONT_STRETCH, fontStyle As DWRITE_FONT_STYLE, <Out> ByRef filteredSet As IDWriteFontSet) As HRESULT
    End Interface

    <ComImport>
    <Guid("7E9FDA85-6C92-4053-BC47-7AE3530DB4D3")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontSet1
        Inherits IDWriteFontSet
#Region "IDWriteFontSet"
        <PreserveSig>
        Overloads Function GetFontCount() As UInteger
        Overloads Function GetFontFaceReference(listIndex As UInteger, <Out> ByRef fontFaceReference As IDWriteFontFaceReference) As HRESULT
        Overloads Function FindFontFaceReference(fontFaceReference As IDWriteFontFaceReference, <Out> ByRef listIndex As UInteger, <Out> ByRef exists As Boolean) As HRESULT
        Overloads Function FindFontFace(fontFace As IDWriteFontFace, <Out> ByRef listIndex As UInteger, <Out> ByRef exists As Boolean) As HRESULT

        Overloads Function GetPropertyValues(propertyID As DWRITE_FONT_PROPERTY_ID, <Out> ByRef values As IDWriteStringList) As HRESULT
        Overloads Function GetPropertyValues(propertyID As DWRITE_FONT_PROPERTY_ID,
<MarshalAs(UnmanagedType.LPWStr)> preferredLocaleNames As String, <Out> ByRef values As IDWriteStringList) As HRESULT
        Overloads Function GetPropertyValues(listIndex As UInteger, propertyId As DWRITE_FONT_PROPERTY_ID, <Out> ByRef exists As Boolean, <Out> ByRef values As IDWriteLocalizedStrings) As HRESULT

        Overloads Function GetPropertyOccurrenceCount([property] As DWRITE_FONT_PROPERTY, <Out> ByRef propertyOccurrenceCount As UInteger) As HRESULT
        Overloads Function GetMatchingFonts(properties As DWRITE_FONT_PROPERTY, propertyCount As UInteger, <Out> ByRef filteredSet As IDWriteFontSet) As HRESULT
        Overloads Function GetMatchingFonts(familyName As String, fontWeight As DWRITE_FONT_WEIGHT, fontStretch As DWRITE_FONT_STRETCH, fontStyle As DWRITE_FONT_STYLE, <Out> ByRef filteredSet As IDWriteFontSet) As HRESULT
#End Region

        Overloads Function GetMatchingFonts(fontProperty As DWRITE_FONT_PROPERTY, fontAxisValues As DWRITE_FONT_AXIS_VALUE, fontAxisValueCount As UInteger, <Out> ByRef matchingFonts As IDWriteFontSet1) As HRESULT
        Function GetFirstFontResources(<Out> ByRef filteredFontSet As IDWriteFontSet1) As HRESULT
        Function GetFilteredFonts(properties As DWRITE_FONT_PROPERTY, propertyCount As UInteger, selectAnyProperty As Boolean, <Out> ByRef filteredFontSet As IDWriteFontSet1) As HRESULT
        Function GetFilteredFonts(fontAxisRanges As DWRITE_FONT_AXIS_RANGE, fontAxisRangeCount As UInteger, selectAnyRange As Boolean, <Out> ByRef filteredFontSet As IDWriteFontSet1) As HRESULT
        Function GetFilteredFonts(indices As UInteger, indexCount As UInteger, <Out> ByRef filteredFontSet As IDWriteFontSet1) As HRESULT
        Function GetFilteredFontIndices(properties As DWRITE_FONT_PROPERTY, propertyCount As UInteger, selectAnyProperty As Boolean, <Out> ByRef indices As UInteger, maxIndexCount As UInteger, <Out> ByRef actualIndexCount As UInteger) As HRESULT
        Function GetFilteredFontIndices(fontAxisRanges As DWRITE_FONT_AXIS_RANGE, fontAxisRangeCount As UInteger, selectAnyRange As Boolean, <Out> ByRef indices As UInteger, maxIndexCount As UInteger, <Out> ByRef actualIndexCount As UInteger) As HRESULT
        Function GetFontAxisRanges(<Out> ByRef fontAxisRanges As DWRITE_FONT_AXIS_RANGE, maxFontAxisRangeCount As UInteger, <Out> ByRef actualFontAxisRangeCount As UInteger) As HRESULT
        Function GetFontAxisRanges(listIndex As UInteger, <Out> ByRef fontAxisRanges As DWRITE_FONT_AXIS_RANGE, maxFontAxisRangeCount As UInteger, <Out> ByRef actualFontAxisRangeCount As UInteger) As HRESULT
        Overloads Function GetFontFaceReference(listIndex As UInteger, <Out> ByRef fontFaceReference As IDWriteFontFaceReference1) As HRESULT
        <PreserveSig>
        Function CreateFontResource(listIndex As UInteger, <Out> ByRef fontResource As IDWriteFontResource) As HRESULT
        <PreserveSig>
        Function CreateFontFace(listIndex As UInteger, <Out> ByRef fontFace As IDWriteFontFace5) As HRESULT
        Function GetFontLocality(listIndex As UInteger) As DWRITE_LOCALITY
    End Interface

    <ComImport>
    <Guid("DC7EAD19-E54C-43AF-B2DA-4E2B79BA3F7F")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontSet2
        Inherits IDWriteFontSet1
#Region "IDWriteFontSet1"
#Region "IDWriteFontSet"
        <PreserveSig>
        Overloads Function GetFontCount() As UInteger
        Overloads Function GetFontFaceReference(listIndex As UInteger, <Out> ByRef fontFaceReference As IDWriteFontFaceReference) As HRESULT
        Overloads Function FindFontFaceReference(fontFaceReference As IDWriteFontFaceReference, <Out> ByRef listIndex As UInteger, <Out> ByRef exists As Boolean) As HRESULT
        Overloads Function FindFontFace(fontFace As IDWriteFontFace, <Out> ByRef listIndex As UInteger, <Out> ByRef exists As Boolean) As HRESULT

        Overloads Function GetPropertyValues(propertyID As DWRITE_FONT_PROPERTY_ID, <Out> ByRef values As IDWriteStringList) As HRESULT
        Overloads Function GetPropertyValues(propertyID As DWRITE_FONT_PROPERTY_ID,
<MarshalAs(UnmanagedType.LPWStr)> preferredLocaleNames As String, <Out> ByRef values As IDWriteStringList) As HRESULT
        Overloads Function GetPropertyValues(listIndex As UInteger, propertyId As DWRITE_FONT_PROPERTY_ID, <Out> ByRef exists As Boolean, <Out> ByRef values As IDWriteLocalizedStrings) As HRESULT

        Overloads Function GetPropertyOccurrenceCount([property] As DWRITE_FONT_PROPERTY, <Out> ByRef propertyOccurrenceCount As UInteger) As HRESULT
        Overloads Function GetMatchingFonts(properties As DWRITE_FONT_PROPERTY, propertyCount As UInteger, <Out> ByRef filteredSet As IDWriteFontSet) As HRESULT
        Overloads Function GetMatchingFonts(familyName As String, fontWeight As DWRITE_FONT_WEIGHT, fontStretch As DWRITE_FONT_STRETCH, fontStyle As DWRITE_FONT_STYLE, <Out> ByRef filteredSet As IDWriteFontSet) As HRESULT
#End Region

        Overloads Function GetMatchingFonts(fontProperty As DWRITE_FONT_PROPERTY, fontAxisValues As DWRITE_FONT_AXIS_VALUE, fontAxisValueCount As UInteger, <Out> ByRef matchingFonts As IDWriteFontSet1) As HRESULT
        Overloads Function GetFirstFontResources(<Out> ByRef filteredFontSet As IDWriteFontSet1) As HRESULT
        Overloads Function GetFilteredFonts(properties As DWRITE_FONT_PROPERTY, propertyCount As UInteger, selectAnyProperty As Boolean, <Out> ByRef filteredFontSet As IDWriteFontSet1) As HRESULT
        Overloads Function GetFilteredFonts(fontAxisRanges As DWRITE_FONT_AXIS_RANGE, fontAxisRangeCount As UInteger, selectAnyRange As Boolean, <Out> ByRef filteredFontSet As IDWriteFontSet1) As HRESULT
        Overloads Function GetFilteredFonts(indices As UInteger, indexCount As UInteger, <Out> ByRef filteredFontSet As IDWriteFontSet1) As HRESULT
        Overloads Function GetFilteredFontIndices(properties As DWRITE_FONT_PROPERTY, propertyCount As UInteger, selectAnyProperty As Boolean, <Out> ByRef indices As UInteger, maxIndexCount As UInteger, <Out> ByRef actualIndexCount As UInteger) As HRESULT
        Overloads Function GetFilteredFontIndices(fontAxisRanges As DWRITE_FONT_AXIS_RANGE, fontAxisRangeCount As UInteger, selectAnyRange As Boolean, <Out> ByRef indices As UInteger, maxIndexCount As UInteger, <Out> ByRef actualIndexCount As UInteger) As HRESULT
        Overloads Function GetFontAxisRanges(<Out> ByRef fontAxisRanges As DWRITE_FONT_AXIS_RANGE, maxFontAxisRangeCount As UInteger, <Out> ByRef actualFontAxisRangeCount As UInteger) As HRESULT
        Overloads Function GetFontAxisRanges(listIndex As UInteger, <Out> ByRef fontAxisRanges As DWRITE_FONT_AXIS_RANGE, maxFontAxisRangeCount As UInteger, <Out> ByRef actualFontAxisRangeCount As UInteger) As HRESULT
        Overloads Function GetFontFaceReference(listIndex As UInteger, <Out> ByRef fontFaceReference As IDWriteFontFaceReference1) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontResource(listIndex As UInteger, <Out> ByRef fontResource As IDWriteFontResource) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontFace(listIndex As UInteger, <Out> ByRef fontFace As IDWriteFontFace5) As HRESULT
        Overloads Function GetFontLocality(listIndex As UInteger) As DWRITE_LOCALITY
#End Region

        Function GetExpirationEvent() As IntPtr
    End Interface

    <ComImport>
    <Guid("7C073EF2-A7F4-4045-8C32-8AB8AE640F90")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontSet3
        Inherits IDWriteFontSet2
#Region "IDWriteFontSet2"
#Region "IDWriteFontSet1"
#Region "IDWriteFontSet"
        <PreserveSig>
        Overloads Function GetFontCount() As UInteger
        Overloads Function GetFontFaceReference(listIndex As UInteger, <Out> ByRef fontFaceReference As IDWriteFontFaceReference) As HRESULT
        Overloads Function FindFontFaceReference(fontFaceReference As IDWriteFontFaceReference, <Out> ByRef listIndex As UInteger, <Out> ByRef exists As Boolean) As HRESULT
        Overloads Function FindFontFace(fontFace As IDWriteFontFace, <Out> ByRef listIndex As UInteger, <Out> ByRef exists As Boolean) As HRESULT

        Overloads Function GetPropertyValues(propertyID As DWRITE_FONT_PROPERTY_ID, <Out> ByRef values As IDWriteStringList) As HRESULT
        Overloads Function GetPropertyValues(propertyID As DWRITE_FONT_PROPERTY_ID,
<MarshalAs(UnmanagedType.LPWStr)> preferredLocaleNames As String, <Out> ByRef values As IDWriteStringList) As HRESULT
        Overloads Function GetPropertyValues(listIndex As UInteger, propertyId As DWRITE_FONT_PROPERTY_ID, <Out> ByRef exists As Boolean, <Out> ByRef values As IDWriteLocalizedStrings) As HRESULT

        Overloads Function GetPropertyOccurrenceCount([property] As DWRITE_FONT_PROPERTY, <Out> ByRef propertyOccurrenceCount As UInteger) As HRESULT
        Overloads Function GetMatchingFonts(properties As DWRITE_FONT_PROPERTY, propertyCount As UInteger, <Out> ByRef filteredSet As IDWriteFontSet) As HRESULT
        Overloads Function GetMatchingFonts(familyName As String, fontWeight As DWRITE_FONT_WEIGHT, fontStretch As DWRITE_FONT_STRETCH, fontStyle As DWRITE_FONT_STYLE, <Out> ByRef filteredSet As IDWriteFontSet) As HRESULT
#End Region

        Overloads Function GetMatchingFonts(fontProperty As DWRITE_FONT_PROPERTY, fontAxisValues As DWRITE_FONT_AXIS_VALUE, fontAxisValueCount As UInteger, <Out> ByRef matchingFonts As IDWriteFontSet1) As HRESULT
        Overloads Function GetFirstFontResources(<Out> ByRef filteredFontSet As IDWriteFontSet1) As HRESULT
        Overloads Function GetFilteredFonts(properties As DWRITE_FONT_PROPERTY, propertyCount As UInteger, selectAnyProperty As Boolean, <Out> ByRef filteredFontSet As IDWriteFontSet1) As HRESULT
        Overloads Function GetFilteredFonts(fontAxisRanges As DWRITE_FONT_AXIS_RANGE, fontAxisRangeCount As UInteger, selectAnyRange As Boolean, <Out> ByRef filteredFontSet As IDWriteFontSet1) As HRESULT
        Overloads Function GetFilteredFonts(indices As UInteger, indexCount As UInteger, <Out> ByRef filteredFontSet As IDWriteFontSet1) As HRESULT
        Overloads Function GetFilteredFontIndices(properties As DWRITE_FONT_PROPERTY, propertyCount As UInteger, selectAnyProperty As Boolean, <Out> ByRef indices As UInteger, maxIndexCount As UInteger, <Out> ByRef actualIndexCount As UInteger) As HRESULT
        Overloads Function GetFilteredFontIndices(fontAxisRanges As DWRITE_FONT_AXIS_RANGE, fontAxisRangeCount As UInteger, selectAnyRange As Boolean, <Out> ByRef indices As UInteger, maxIndexCount As UInteger, <Out> ByRef actualIndexCount As UInteger) As HRESULT
        Overloads Function GetFontAxisRanges(<Out> ByRef fontAxisRanges As DWRITE_FONT_AXIS_RANGE, maxFontAxisRangeCount As UInteger, <Out> ByRef actualFontAxisRangeCount As UInteger) As HRESULT
        Overloads Function GetFontAxisRanges(listIndex As UInteger, <Out> ByRef fontAxisRanges As DWRITE_FONT_AXIS_RANGE, maxFontAxisRangeCount As UInteger, <Out> ByRef actualFontAxisRangeCount As UInteger) As HRESULT
        Overloads Function GetFontFaceReference(listIndex As UInteger, <Out> ByRef fontFaceReference As IDWriteFontFaceReference1) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontResource(listIndex As UInteger, <Out> ByRef fontResource As IDWriteFontResource) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontFace(listIndex As UInteger, <Out> ByRef fontFace As IDWriteFontFace5) As HRESULT
        Overloads Function GetFontLocality(listIndex As UInteger) As DWRITE_LOCALITY
#End Region

        Overloads Function GetExpirationEvent() As IntPtr
#End Region

        <PreserveSig>
        Function GetFontSourceType(fontIndex As UInteger) As DWRITE_FONT_SOURCE_TYPE
        <PreserveSig>
        Function GetFontSourceNameLength(listIndex As UInteger) As UInteger
        Function GetFontSourceName(listIndex As UInteger, stringBuffer As Text.StringBuilder, stringBufferSize As UInteger) As HRESULT
    End Interface

    <ComImport>
    <Guid("EEC175FC-BEA9-4C86-8B53-CCBDD7DF0C82")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontSet4
        Inherits IDWriteFontSet3
#Region "IDWriteFontSet2"
#Region "IDWriteFontSet1"
#Region "IDWriteFontSet"
        <PreserveSig>
        Overloads Function GetFontCount() As UInteger
        Overloads Function GetFontFaceReference(listIndex As UInteger, <Out> ByRef fontFaceReference As IDWriteFontFaceReference) As HRESULT
        Overloads Function FindFontFaceReference(fontFaceReference As IDWriteFontFaceReference, <Out> ByRef listIndex As UInteger, <Out> ByRef exists As Boolean) As HRESULT
        Overloads Function FindFontFace(fontFace As IDWriteFontFace, <Out> ByRef listIndex As UInteger, <Out> ByRef exists As Boolean) As HRESULT

        Overloads Function GetPropertyValues(propertyID As DWRITE_FONT_PROPERTY_ID, <Out> ByRef values As IDWriteStringList) As HRESULT
        Overloads Function GetPropertyValues(propertyID As DWRITE_FONT_PROPERTY_ID,
<MarshalAs(UnmanagedType.LPWStr)> preferredLocaleNames As String, <Out> ByRef values As IDWriteStringList) As HRESULT
        Overloads Function GetPropertyValues(listIndex As UInteger, propertyId As DWRITE_FONT_PROPERTY_ID, <Out> ByRef exists As Boolean, <Out> ByRef values As IDWriteLocalizedStrings) As HRESULT

        Overloads Function GetPropertyOccurrenceCount([property] As DWRITE_FONT_PROPERTY, <Out> ByRef propertyOccurrenceCount As UInteger) As HRESULT
        Overloads Function GetMatchingFonts(properties As DWRITE_FONT_PROPERTY, propertyCount As UInteger, <Out> ByRef filteredSet As IDWriteFontSet) As HRESULT
        Overloads Function GetMatchingFonts(familyName As String, fontWeight As DWRITE_FONT_WEIGHT, fontStretch As DWRITE_FONT_STRETCH, fontStyle As DWRITE_FONT_STYLE, <Out> ByRef filteredSet As IDWriteFontSet) As HRESULT
#End Region

        Overloads Function GetMatchingFonts(fontProperty As DWRITE_FONT_PROPERTY, fontAxisValues As DWRITE_FONT_AXIS_VALUE, fontAxisValueCount As UInteger, <Out> ByRef matchingFonts As IDWriteFontSet1) As HRESULT
        Overloads Function GetFirstFontResources(<Out> ByRef filteredFontSet As IDWriteFontSet1) As HRESULT
        Overloads Function GetFilteredFonts(properties As DWRITE_FONT_PROPERTY, propertyCount As UInteger, selectAnyProperty As Boolean, <Out> ByRef filteredFontSet As IDWriteFontSet1) As HRESULT
        Overloads Function GetFilteredFonts(fontAxisRanges As DWRITE_FONT_AXIS_RANGE, fontAxisRangeCount As UInteger, selectAnyRange As Boolean, <Out> ByRef filteredFontSet As IDWriteFontSet1) As HRESULT
        Overloads Function GetFilteredFonts(indices As UInteger, indexCount As UInteger, <Out> ByRef filteredFontSet As IDWriteFontSet1) As HRESULT
        Overloads Function GetFilteredFontIndices(properties As DWRITE_FONT_PROPERTY, propertyCount As UInteger, selectAnyProperty As Boolean, <Out> ByRef indices As UInteger, maxIndexCount As UInteger, <Out> ByRef actualIndexCount As UInteger) As HRESULT
        Overloads Function GetFilteredFontIndices(fontAxisRanges As DWRITE_FONT_AXIS_RANGE, fontAxisRangeCount As UInteger, selectAnyRange As Boolean, <Out> ByRef indices As UInteger, maxIndexCount As UInteger, <Out> ByRef actualIndexCount As UInteger) As HRESULT
        Overloads Function GetFontAxisRanges(<Out> ByRef fontAxisRanges As DWRITE_FONT_AXIS_RANGE, maxFontAxisRangeCount As UInteger, <Out> ByRef actualFontAxisRangeCount As UInteger) As HRESULT
        Overloads Function GetFontAxisRanges(listIndex As UInteger, <Out> ByRef fontAxisRanges As DWRITE_FONT_AXIS_RANGE, maxFontAxisRangeCount As UInteger, <Out> ByRef actualFontAxisRangeCount As UInteger) As HRESULT
        Overloads Function GetFontFaceReference(listIndex As UInteger, <Out> ByRef fontFaceReference As IDWriteFontFaceReference1) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontResource(listIndex As UInteger, <Out> ByRef fontResource As IDWriteFontResource) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontFace(listIndex As UInteger, <Out> ByRef fontFace As IDWriteFontFace5) As HRESULT
        Overloads Function GetFontLocality(listIndex As UInteger) As DWRITE_LOCALITY
#End Region

        Overloads Function GetExpirationEvent() As IntPtr
#End Region

#Region "<IDWriteFontSet3>"
        <PreserveSig>
        Overloads Function GetFontSourceType(fontIndex As UInteger) As DWRITE_FONT_SOURCE_TYPE
        <PreserveSig>
        Overloads Function GetFontSourceNameLength(listIndex As UInteger) As UInteger
        <PreserveSig>
        Overloads Function GetFontSourceName(listIndex As UInteger, stringBuffer As Text.StringBuilder, stringBufferSize As UInteger) As HRESULT
#End Region

        <PreserveSig>
        Function ConvertWeightStretchStyleToFontAxisValues(
                        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> inputAxisValues As DWRITE_FONT_AXIS_VALUE(), inputAxisCount As UInteger, fontWeight As DWRITE_FONT_WEIGHT, fontStretch As DWRITE_FONT_STRETCH, fontStyle As DWRITE_FONT_STYLE, fontSize As Single,
        <Out, MarshalAs(UnmanagedType.LPArray)> outputAxisValues As DWRITE_FONT_AXIS_VALUE()) As UInteger

        <PreserveSig>
        Overloads Function GetMatchingFonts(
        <MarshalAs(UnmanagedType.LPWStr)> familyName As String,
                        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=2)> fontAxisValues As DWRITE_FONT_AXIS_VALUE(), fontAxisValueCount As Integer, allowedSimulations As DWRITE_FONT_SIMULATIONS, <Out> ByRef matchingFonts As IDWriteFontSet4) As HRESULT

    End Interface

    <ComImport>
    <Guid("1F803A76-6871-48E8-987F-B975551C50F2")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontResource
        Function GetFontFile(<Out> ByRef fontFile As IDWriteFontFile) As HRESULT
        <PreserveSig>
        Function GetFontFaceIndex() As UInteger
        <PreserveSig>
        Function GetFontAxisCount() As UInteger
        <PreserveSig>
        Function GetDefaultFontAxisValues(<Out> ByRef fontAxisValues As DWRITE_FONT_AXIS_VALUE, fontAxisValueCount As UInteger) As HRESULT
        <PreserveSig>
        Function GetFontAxisRanges(<Out> ByRef fontAxisRanges As DWRITE_FONT_AXIS_RANGE, fontAxisRangeCount As UInteger) As HRESULT
        <PreserveSig>
        Function GetFontAxisAttributes(axisIndex As UInteger) As DWRITE_FONT_AXIS_ATTRIBUTES
        <PreserveSig>
        Function GetAxisNames(axisIndex As UInteger, <Out> ByRef names As IDWriteLocalizedStrings) As HRESULT
        <PreserveSig>
        Function GetAxisValueNameCount(axisIndex As UInteger) As UInteger
        <PreserveSig>
        Function GetAxisValueNames(axisIndex As UInteger, axisValueIndex As UInteger, <Out> ByRef fontAxisRange As DWRITE_FONT_AXIS_RANGE, <Out> ByRef names As IDWriteLocalizedStrings) As HRESULT
        <PreserveSig>
        Function CreateFontFace(fontSimulations As DWRITE_FONT_SIMULATIONS, fontAxisValues As DWRITE_FONT_AXIS_VALUE, fontAxisValueCount As UInteger, <Out> ByRef fontFace As IDWriteFontFace5) As HRESULT
        <PreserveSig>
        Function CreateFontFaceReference(fontSimulations As DWRITE_FONT_SIMULATIONS, fontAxisValues As DWRITE_FONT_AXIS_VALUE, fontAxisValueCount As UInteger, <Out> ByRef fontFaceReference As IDWriteFontFaceReference1) As HRESULT
    End Interface

    <ComImport>
    <Guid("B71E6052-5AEA-4FA3-832E-F60D431F7E91")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontDownloadQueue
        Function AddListener(listener As IDWriteFontDownloadListener, <Out> ByRef token As UInteger) As HRESULT
        Function RemoveListener(token As UInteger) As HRESULT
        Function IsEmpty() As Boolean
        Function BeginDownload(context As IntPtr) As HRESULT
        Function CancelDownload() As HRESULT
        <PreserveSig>
        Function GetGenerationCount() As ULong
    End Interface

    <ComImport>
    <Guid("B06FE5B9-43EC-4393-881B-DBE4DC72FDA7")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontDownloadListener
        Sub DownloadCompleted(downloadQueue As IDWriteFontDownloadQueue, context As IntPtr, downloadResult As HRESULT)
    End Interface

    <ComImport>
    <Guid("CFEE3140-1157-47CA-8B85-31BFCF3F2D0E")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteStringList
        <PreserveSig>
        Function GetCount() As UInteger
        Function GetLocaleNameLength(listIndex As UInteger, <Out> ByRef length As UInteger) As HRESULT
        Function GetLocaleName(listIndex As UInteger, localeName As Text.StringBuilder, size As UInteger) As HRESULT
        Function GetStringLength(listIndex As UInteger, <Out> ByRef length As UInteger) As HRESULT
        Function GetString(listIndex As UInteger, stringBuffer As Text.StringBuilder, stringBufferSize As UInteger) As HRESULT
    End Interface

    <ComImport>
    <Guid("5E7FA7CA-DDE3-424C-89F0-9FCD6FED58CD")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontFaceReference
        <PreserveSig>
        Function CreateFontFace(<Out> ByRef fontFace As IDWriteFontFace3) As HRESULT
        <PreserveSig>
        Function CreateFontFaceWithSimulations(fontFaceSimulationFlags As DWRITE_FONT_SIMULATIONS, <Out> ByRef fontFace As IDWriteFontFace3) As HRESULT
        <PreserveSig>
        Function Equals(fontFaceReference As IDWriteFontFaceReference) As Boolean
        <PreserveSig>
        Function GetFontFaceIndex() As UInteger
        <PreserveSig>
        Function GetSimulations() As DWRITE_FONT_SIMULATIONS
        <PreserveSig>
        Function GetFontFile(<Out> ByRef fontFile As IDWriteFontFile) As HRESULT
        <PreserveSig>
        Function GetLocalFileSize() As ULong
        <PreserveSig>
        Function GetFileSize() As ULong
        <PreserveSig>
        Function GetFileTime(<Out> ByRef lastWriteTime As ComTypes.FILETIME) As HRESULT
        <PreserveSig>
        Function GetLocality() As DWRITE_LOCALITY
        <PreserveSig>
        Function EnqueueFontDownloadRequest() As HRESULT
        <PreserveSig>
        Function EnqueueCharacterDownloadRequest(characters As String, characterCount As UInteger) As HRESULT
        <PreserveSig>
        Function EnqueueGlyphDownloadRequest(glyphIndices As UShort, glyphCount As UInteger) As HRESULT
        <PreserveSig>
        Function EnqueueFileFragmentDownloadRequest(fileOffset As ULong, fragmentSize As ULong) As HRESULT
    End Interface

    <ComImport>
    <Guid("C081FE77-2FD1-41AC-A5A3-34983C4BA61A")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontFaceReference1
        Inherits IDWriteFontFaceReference
#Region "IDWriteFontFaceReference"
        <PreserveSig>
        Overloads Function CreateFontFace(<Out> ByRef fontFace As IDWriteFontFace3) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontFaceWithSimulations(fontFaceSimulationFlags As DWRITE_FONT_SIMULATIONS, <Out> ByRef fontFace As IDWriteFontFace3) As HRESULT
        <PreserveSig>
        Overloads Function Equals(fontFaceReference As IDWriteFontFaceReference) As Boolean
        <PreserveSig>
        Overloads Function GetFontFaceIndex() As UInteger
        <PreserveSig>
        Overloads Function GetSimulations() As DWRITE_FONT_SIMULATIONS
        <PreserveSig>
        Overloads Function GetFontFile(<Out> ByRef fontFile As IDWriteFontFile) As HRESULT
        <PreserveSig>
        Overloads Function GetLocalFileSize() As ULong
        <PreserveSig>
        Overloads Function GetFileSize() As ULong
        <PreserveSig>
        Overloads Function GetFileTime(<Out> ByRef lastWriteTime As ComTypes.FILETIME) As HRESULT
        <PreserveSig>
        Overloads Function GetLocality() As DWRITE_LOCALITY
        <PreserveSig>
        Overloads Function EnqueueFontDownloadRequest() As HRESULT
        <PreserveSig>
        Overloads Function EnqueueCharacterDownloadRequest(characters As String, characterCount As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function EnqueueGlyphDownloadRequest(glyphIndices As UShort, glyphCount As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function EnqueueFileFragmentDownloadRequest(fileOffset As ULong, fragmentSize As ULong) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateFontFace(<Out> ByRef fontFace As IDWriteFontFace5) As HRESULT
        <PreserveSig>
        Function GetFontAxisValueCount() As UInteger
        <PreserveSig>
        Function GetFontAxisValues(<Out> ByRef fontAxisValues As DWRITE_FONT_AXIS_VALUE, fontAxisValueCount As UInteger) As HRESULT
    End Interface

    <ComImport>
    <Guid("2F642AFE-9C68-4F40-B8BE-457401AFCB3D")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontSetBuilder
        Function AddFontFaceReference(fontFaceReference As IDWriteFontFaceReference) As HRESULT
        Function AddFontFaceReference(fontFaceReference As IDWriteFontFaceReference, properties As DWRITE_FONT_PROPERTY, propertyCount As UInteger) As HRESULT
        Function AddFontSet(fontSet As IDWriteFontSet) As HRESULT
        <PreserveSig>
        Function CreateFontSet(<Out> ByRef fontSet As IDWriteFontSet) As HRESULT
    End Interface

    <ComImport>
    <Guid("3FF7715F-3CDC-4DC6-9B72-EC5621DCCAFD")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontSetBuilder1
        Inherits IDWriteFontSetBuilder
#Region "IDWriteFontSetBuilder"
        Overloads Function AddFontFaceReference(fontFaceReference As IDWriteFontFaceReference) As HRESULT
        Overloads Function AddFontFaceReference(fontFaceReference As IDWriteFontFaceReference, properties As DWRITE_FONT_PROPERTY, propertyCount As UInteger) As HRESULT
        Overloads Function AddFontSet(fontSet As IDWriteFontSet) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontSet(<Out> ByRef fontSet As IDWriteFontSet) As HRESULT
#End Region

        Function AddFontFile(fontFile As IDWriteFontFile) As HRESULT
    End Interface

    <ComImport>
    <Guid("EE5BA612-B131-463C-8F4F-3189B9401E45")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontSetBuilder2
        Inherits IDWriteFontSetBuilder1
#Region "IDWriteFontSetBuilder1"
#Region "IDWriteFontSetBuilder"
        Overloads Function AddFontFaceReference(fontFaceReference As IDWriteFontFaceReference) As HRESULT
        Overloads Function AddFontFaceReference(fontFaceReference As IDWriteFontFaceReference, properties As DWRITE_FONT_PROPERTY, propertyCount As UInteger) As HRESULT
        Overloads Function AddFontSet(fontSet As IDWriteFontSet) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontSet(<Out> ByRef fontSet As IDWriteFontSet) As HRESULT
#End Region

        Overloads Function AddFontFile(fontFile As IDWriteFontFile) As HRESULT
#End Region

        Function AddFont(fontFile As IDWriteFontFile, fontFaceIndex As UInteger, fontSimulations As DWRITE_FONT_SIMULATIONS, fontAxisValues As DWRITE_FONT_AXIS_VALUE, fontAxisValueCount As UInteger, fontAxisRanges As DWRITE_FONT_AXIS_RANGE, fontAxisRangeCount As UInteger, properties As DWRITE_FONT_PROPERTY, propertyCount As UInteger) As HRESULT
        Overloads Function AddFontFile(filePath As String) As HRESULT
    End Interface

    <ComImport>
    <Guid("CE25F8FD-863B-4D13-9651-C1F88DC73FE2")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteAsyncResult
        <PreserveSig>
        Function GetWaitHandle() As IntPtr
        <PreserveSig>
        Function GetResult() As HRESULT
    End Interface

    <ComImport>
    <Guid("d31fbe17-f157-41a2-8d24-cb779e0560e8")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteColorGlyphRunEnumerator
        <PreserveSig>
        Function MoveNext(<Out> ByRef hasRun As Boolean) As HRESULT
        <PreserveSig>
        Function GetCurrentRun(<Out> ByRef colorGlyphRun As IntPtr) As HRESULT
    End Interface

    <ComImport>
    <Guid("7C5F86DA-C7A1-4F05-B8E1-55A179FE5A35")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteColorGlyphRunEnumerator1
        Inherits IDWriteColorGlyphRunEnumerator
#Region "IDWriteColorGlyphRunEnumerator"
        <PreserveSig>
        Overloads Function MoveNext(<Out> ByRef hasRun As Boolean) As HRESULT
        <PreserveSig>
        Overloads Function GetCurrentRun(<Out> ByRef colorGlyphRun As IntPtr) As HRESULT
#End Region

        <PreserveSig>
        Function GetCurrentRun1(<Out> ByRef colorGlyphRun As IntPtr) As HRESULT
    End Interface

    <ComImport>
    <Guid("a84cee02-3eea-4eee-a827-87c1a02a0fcc")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontCollection
        <PreserveSig>
        Function GetFontFamilyCount() As UInteger
        Function GetFontFamily(index As UInteger, <Out> ByRef fontFamily As IDWriteFontFamily) As HRESULT
        Function FindFamilyName(familyName As String, <Out> ByRef index As UInteger, <Out> ByRef exists As Boolean) As HRESULT
        Function GetFontFromFontFace(fontFace As IDWriteFontFace, <Out> ByRef font As IDWriteFont) As HRESULT
    End Interface

    <ComImport>
    <Guid("53585141-D9F8-4095-8321-D73CF6BD116C")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontCollection1
        Inherits IDWriteFontCollection
#Region "IDWriteFontCollection"
        <PreserveSig>
        Overloads Function GetFontFamilyCount() As UInteger
        Overloads Function GetFontFamily(index As UInteger, <Out> ByRef fontFamily As IDWriteFontFamily) As HRESULT
        Overloads Function FindFamilyName(familyName As String, <Out> ByRef index As UInteger, <Out> ByRef exists As Boolean) As HRESULT
        Overloads Function GetFontFromFontFace(fontFace As IDWriteFontFace, <Out> ByRef font As IDWriteFont) As HRESULT
#End Region

        Function GetFontSet(<Out> ByRef fontSet As IDWriteFontSet) As HRESULT
        Overloads Function GetFontFamily(index As UInteger, <Out> ByRef fontFamily As IDWriteFontFamily1) As HRESULT
    End Interface

    <ComImport>
    <Guid("514039C6-4617-4064-BF8B-92EA83E506E0")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontCollection2
        Inherits IDWriteFontCollection1
#Region "IDWriteFontCollection1"
#Region "IDWriteFontCollection"
        <PreserveSig>
        Overloads Function GetFontFamilyCount() As UInteger
        <PreserveSig>
        Overloads Function GetFontFamily(index As UInteger, <Out> ByRef fontFamily As IDWriteFontFamily) As HRESULT
        <PreserveSig>
        Overloads Function FindFamilyName(familyName As String, <Out> ByRef index As UInteger, <Out> ByRef exists As Boolean) As HRESULT
        <PreserveSig>
        Overloads Function GetFontFromFontFace(fontFace As IDWriteFontFace, <Out> ByRef font As IDWriteFont) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function GetFontSet(<Out> ByRef fontSet As IDWriteFontSet) As HRESULT
        <PreserveSig>
        Overloads Function GetFontFamily(index As UInteger, <Out> ByRef fontFamily As IDWriteFontFamily1) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function GetFontFamily(index As UInteger, <Out> ByRef fontFamily As IDWriteFontFamily2) As HRESULT
        <PreserveSig>
        Function GetMatchingFonts(familyName As String, fontAxisValues As DWRITE_FONT_AXIS_VALUE, fontAxisValueCount As UInteger, <Out> ByRef fontList As IDWriteFontList2) As HRESULT
        <PreserveSig>
        Function GetFontFamilyModel() As DWRITE_FONT_FAMILY_MODEL
        <PreserveSig>
        Overloads Function GetFontSet(<Out> ByRef fontSet As IDWriteFontSet1) As HRESULT
    End Interface

    <ComImport>
    <Guid("A4D055A6-F9E3-4E25-93B7-9E309F3AF8E9")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontCollection3
        Inherits IDWriteFontCollection2
#Region "IDWriteFontCollection2"
#Region "IDWriteFontCollection1"
#Region "IDWriteFontCollection"
        <PreserveSig>
        Overloads Function GetFontFamilyCount() As UInteger
        <PreserveSig>
        Overloads Function GetFontFamily(index As UInteger, <Out> ByRef fontFamily As IDWriteFontFamily) As HRESULT
        <PreserveSig>
        Overloads Function FindFamilyName(familyName As String, <Out> ByRef index As UInteger, <Out> ByRef exists As Boolean) As HRESULT
        <PreserveSig>
        Overloads Function GetFontFromFontFace(fontFace As IDWriteFontFace, <Out> ByRef font As IDWriteFont) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function GetFontSet(<Out> ByRef fontSet As IDWriteFontSet) As HRESULT
        <PreserveSig>
        Overloads Function GetFontFamily(index As UInteger, <Out> ByRef fontFamily As IDWriteFontFamily1) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function GetFontFamily(index As UInteger, <Out> ByRef fontFamily As IDWriteFontFamily2) As HRESULT
        <PreserveSig>
        Overloads Function GetMatchingFonts(familyName As String, fontAxisValues As DWRITE_FONT_AXIS_VALUE, fontAxisValueCount As UInteger, <Out> ByRef fontList As IDWriteFontList2) As HRESULT
        <PreserveSig>
        Overloads Function GetFontFamilyModel() As DWRITE_FONT_FAMILY_MODEL
        <PreserveSig>
        Overloads Function GetFontSet(<Out> ByRef fontSet As IDWriteFontSet1) As HRESULT
#End Region

        <PreserveSig>
        Function GetExpirationEvent() As IntPtr
    End Interface

    <ComImport>
    <Guid("cca920e4-52f0-492b-bfa8-29c72ee0a468")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontCollectionLoader
        <PreserveSig>
        Function CreateEnumeratorFromKey(factory As IDWriteFactory, collectionKey As IntPtr, collectionKeySize As UInteger, <Out> ByRef fontFileEnumerator As IDWriteFontFileEnumerator) As HRESULT
        'HRESULT CreateEnumeratorFromKey(IntPtr factory, IntPtr collectionKey, uint collectionKeySize, out IntPtr fontFileEnumerator);
    End Interface

    <ComImport>
    <Guid("72755049-5ff7-435d-8348-4be97cfa6c7c")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontFileEnumerator
        <PreserveSig>
        Function MoveNext(<Out>
        <MarshalAs(UnmanagedType.Bool)> ByRef hasCurrentFile As Boolean) As HRESULT
        <PreserveSig>
        Function GetCurrentFontFile(<Out> ByRef fontFile As IDWriteFontFile) As HRESULT
    End Interface

    <ComImport>
    <Guid("739d886a-cef5-47dc-8769-1a8b41bebbb0")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontFile
        Function GetReferenceKey(<Out> ByRef fontFileReferenceKey As IntPtr, <Out> ByRef fontFileReferenceKeySize As Integer) As HRESULT
        Function GetLoader(<Out> ByRef fontFileLoader As IDWriteFontFileLoader) As HRESULT
        Function Analyze(<Out> ByRef isSupportedFontType As Boolean, <Out> ByRef fontFileType As DWRITE_FONT_FILE_TYPE, <Out> ByRef fontFaceType As DWRITE_FONT_FACE_TYPE, <Out> ByRef numberOfFaces As Integer) As HRESULT
    End Interface

    <ComImport>
    <Guid("727cad4e-d6af-4c9e-8a08-d695b11caa49")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontFileLoader
        <PreserveSig>
        Function CreateStreamFromKey(fontFileReferenceKey As IntPtr, fontFileReferenceKeySize As UInteger, <Out> ByRef fontFileStream As IDWriteFontFileStream) As HRESULT
    End Interface

    <ComImport>
    <Guid("b2d9f3ec-c9fe-4a11-a2ec-d86208f7c0a2")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteLocalFontFileLoader
        Inherits IDWriteFontFileLoader
#Region "IDWriteFontFileLoader"
        <PreserveSig>
        Overloads Function CreateStreamFromKey(fontFileReferenceKey As IntPtr, fontFileReferenceKeySize As UInteger, <Out> ByRef fontFileStream As IDWriteFontFileStream) As HRESULT
#End Region

        Function GetFilePathLengthFromKey(fontFileReferenceKey As IntPtr, fontFileReferenceKeySize As UInteger, <Out> ByRef filePathLength As UInteger) As HRESULT
        Function GetFilePathFromKey(fontFileReferenceKey As IntPtr, fontFileReferenceKeySize As UInteger, filePath As Text.StringBuilder, filePathSize As UInteger) As HRESULT
        Function GetLastWriteTimeFromKey(fontFileReferenceKey As IntPtr, fontFileReferenceKeySize As UInteger, <Out> ByRef lastWriteTime As IntPtr) As HRESULT
    End Interface

    <ComImport>
    <Guid("68648C83-6EDE-46C0-AB46-20083A887FDE")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteRemoteFontFileLoader
        Inherits IDWriteFontFileLoader
#Region "IDWriteFontFileLoader"
        <PreserveSig>
        Overloads Function CreateStreamFromKey(fontFileReferenceKey As IntPtr, fontFileReferenceKeySize As UInteger, <Out> ByRef fontFileStream As IDWriteFontFileStream) As HRESULT
#End Region

        <PreserveSig>
        Function CreateRemoteStreamFromKey(fontFileReferenceKey As IntPtr, fontFileReferenceKeySize As UInteger, <Out> ByRef fontFileStream As IDWriteRemoteFontFileStream) As HRESULT
        Function GetLocalityFromKey(fontFileReferenceKey As IntPtr, fontFileReferenceKeySize As UInteger, <Out> ByRef locality As DWRITE_LOCALITY) As HRESULT
        <PreserveSig>
        Function CreateFontFileReferenceFromUrl(factory As IDWriteFactory,
        <MarshalAs(UnmanagedType.LPWStr)> baseUrl As String,
        <MarshalAs(UnmanagedType.LPWStr)> fontFileUrl As String, <Out> ByRef fontFile As IDWriteFontFile) As HRESULT
    End Interface

    <ComImport>
    <Guid("DC102F47-A12D-4B1C-822D-9E117E33043F")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteInMemoryFontFileLoader
        Inherits IDWriteFontFileLoader
#Region "IDWriteFontFileLoader"
        <PreserveSig>
        Overloads Function CreateStreamFromKey(fontFileReferenceKey As IntPtr, fontFileReferenceKeySize As UInteger, <Out> ByRef fontFileStream As IDWriteFontFileStream) As HRESULT
#End Region

        <PreserveSig>
        Function CreateInMemoryFontFileReference(factory As IDWriteFactory, fontData As IntPtr, fontDataSize As UInteger, ownerObject As IntPtr, <Out> ByRef fontFile As IDWriteFontFile) As HRESULT
        <PreserveSig>
        Function GetFileCount() As UInteger
    End Interface

    <ComImport>
    <Guid("6d4865fe-0ab8-4d91-8f62-5dd6be34a3e0")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontFileStream
        Function ReadFileFragment(<Out> ByRef fragmentStart As IntPtr, fileOffset As ULong, fragmentSize As ULong, <Out> ByRef fragmentContext As IntPtr) As HRESULT
        Sub ReleaseFileFragment(fragmentContext As IntPtr)
        Function GetFileSize(<Out> ByRef fileSize As ULong) As HRESULT
        Function GetLastWriteTime(<Out> ByRef lastWriteTime As ULong) As HRESULT
    End Interface

    <ComImport>
    <Guid("4DB3757A-2C72-4ED9-B2B6-1ABABE1AFF9C")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteRemoteFontFileStream
        Inherits IDWriteFontFileStream
#Region "IDWriteFontFileStream"
        Overloads Function ReadFileFragment(<Out> ByRef fragmentStart As IntPtr, fileOffset As ULong, fragmentSize As ULong, <Out> ByRef fragmentContext As IntPtr) As HRESULT
        Overloads Sub ReleaseFileFragment(fragmentContext As IntPtr)
        Overloads Function GetFileSize(<Out> ByRef fileSize As ULong) As HRESULT
        Overloads Function GetLastWriteTime(<Out> ByRef lastWriteTime As ULong) As HRESULT
#End Region

        Function GetLocalFileSize(<Out> ByRef localFileSize As ULong) As HRESULT
        Function GetFileFragmentLocality(fileOffset As ULong, fragmentSize As ULong, <Out> ByRef isLocal As Boolean, <Out> ByRef partialSize As ULong) As HRESULT
        Function GetLocality() As DWRITE_LOCALITY
        Function BeginDownload(ByRef downloadOperationID As Guid, fileFragments As DWRITE_FILE_FRAGMENT, fragmentCount As UInteger, <Out> ByRef asyncResult As IDWriteAsyncResult) As HRESULT
    End Interface

    <ComImport>
    <Guid("2f0da53a-2add-47cd-82ee-d9ec34688e75")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteRenderingParams
        <PreserveSig>
        Function GetGamma() As Single
        <PreserveSig>
        Function GetEnhancedContrast() As Single
        <PreserveSig>
        Function GetClearTypeLevel() As Single
        <PreserveSig>
        Function GetPixelGeometry() As DWRITE_PIXEL_GEOMETRY
        <PreserveSig>
        Function GetRenderingMode() As DWRITE_RENDERING_MODE
    End Interface

    <ComImport>
    <Guid("94413cf4-a6fc-4248-8b50-6674348fcad3")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteRenderingParams1
        Inherits IDWriteRenderingParams
#Region "IDWriteRenderingParams"

        <PreserveSig>
        Overloads Function GetGamma() As Single
        <PreserveSig>
        Overloads Function GetEnhancedContrast() As Single
        <PreserveSig>
        Overloads Function GetClearTypeLevel() As Single
        <PreserveSig>
        Overloads Function GetPixelGeometry() As DWRITE_PIXEL_GEOMETRY
        <PreserveSig>
        Overloads Function GetRenderingMode() As DWRITE_RENDERING_MODE
#End Region

        <PreserveSig>
        Function GetGrayscaleEnhancedContrast() As Single
    End Interface

    <ComImport>
    <Guid("F9D711C3-9777-40AE-87E8-3E5AF9BF0948")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteRenderingParams2
        Inherits IDWriteRenderingParams1
#Region "IDWriteRenderingParams1"
#Region "IDWriteRenderingParams"

        <PreserveSig>
        Overloads Function GetGamma() As Single
        <PreserveSig>
        Overloads Function GetEnhancedContrast() As Single
        <PreserveSig>
        Overloads Function GetClearTypeLevel() As Single
        <PreserveSig>
        Overloads Function GetPixelGeometry() As DWRITE_PIXEL_GEOMETRY
        <PreserveSig>
        Overloads Function GetRenderingMode() As DWRITE_RENDERING_MODE
#End Region

        <PreserveSig>
        Overloads Function GetGrayscaleEnhancedContrast() As Single
#End Region

        <PreserveSig>
        Function GetGridFitMode() As DWRITE_GRID_FIT_MODE
    End Interface

    <ComImport>
    <Guid("B7924BAA-391B-412A-8C5C-E44CC2D867DC")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteRenderingParams3
        Inherits IDWriteRenderingParams2
#Region "IDWriteRenderingParams2"
#Region "IDWriteRenderingParams1"
#Region "IDWriteRenderingParams"

        <PreserveSig>
        Overloads Function GetGamma() As Single
        <PreserveSig>
        Overloads Function GetEnhancedContrast() As Single
        <PreserveSig>
        Overloads Function GetClearTypeLevel() As Single
        <PreserveSig>
        Overloads Function GetPixelGeometry() As DWRITE_PIXEL_GEOMETRY
        <PreserveSig>
        Overloads Function GetRenderingMode() As DWRITE_RENDERING_MODE
#End Region

        <PreserveSig>
        Overloads Function GetGrayscaleEnhancedContrast() As Single
#End Region

        <PreserveSig>
        Overloads Function GetGridFitMode() As DWRITE_GRID_FIT_MODE
#End Region

        <PreserveSig>
        Function GetRenderingMode1() As DWRITE_RENDERING_MODE1
    End Interface

    <ComImport>
    <Guid("9c906818-31d7-4fd3-a151-7c5e225db55a")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteTextFormat
        Function SetTextAlignment(textAlignment As DWRITE_TEXT_ALIGNMENT) As HRESULT
        Function SetParagraphAlignment(paragraphAlignment As DWRITE_PARAGRAPH_ALIGNMENT) As HRESULT
        Function SetWordWrapping(wordWrapping As DWRITE_WORD_WRAPPING) As HRESULT
        Function SetReadingDirection(readingDirection As DWRITE_READING_DIRECTION) As HRESULT
        Function SetFlowDirection(flowDirection As DWRITE_FLOW_DIRECTION) As HRESULT
        Function SetIncrementalTabStop(incrementalTabStop As Single) As HRESULT
        Function SetTrimming(trimmingOptions As DWRITE_TRIMMING, trimmingSign As IDWriteInlineObject) As HRESULT
        Function SetLineSpacing(lineSpacingMethod As DWRITE_LINE_SPACING_METHOD, lineSpacing As Single, baseline As Single) As HRESULT
        <PreserveSig>
        Function GetTextAlignment() As DWRITE_TEXT_ALIGNMENT
        <PreserveSig>
        Function GetParagraphAlignment() As DWRITE_PARAGRAPH_ALIGNMENT
        <PreserveSig>
        Function GetWordWrapping() As DWRITE_WORD_WRAPPING
        <PreserveSig>
        Function GetReadingDirection() As DWRITE_READING_DIRECTION
        <PreserveSig>
        Function GetFlowDirection() As DWRITE_FLOW_DIRECTION
        <PreserveSig>
        Function GetIncrementalTabStop() As Single
        Function GetTrimming(<Out> ByRef trimmingOptions As DWRITE_TRIMMING, <Out> ByRef trimmingSign As IDWriteInlineObject) As HRESULT
        Function GetLineSpacing(<Out> ByRef lineSpacingMethod As DWRITE_LINE_SPACING_METHOD, <Out> ByRef lineSpacing As Single, <Out> ByRef baseline As Single) As HRESULT
        Function GetFontCollection(<Out> ByRef fontCollection As IDWriteFontCollection) As HRESULT
        <PreserveSig>
        Function GetFontFamilyNameLength() As UInteger
        Function GetFontFamilyName(<Out> ByRef fontFamilyName As String, nameSize As UInteger) As HRESULT
        <PreserveSig>
        Function GetFontWeight() As DWRITE_FONT_WEIGHT
        <PreserveSig>
        Function GetFontStyle() As DWRITE_FONT_STYLE
        <PreserveSig>
        Function GetFontStretch() As DWRITE_FONT_STRETCH
        <PreserveSig>
        Function GetFontSize() As Single
        <PreserveSig>
        Function GetLocaleNameLength() As UInteger
        Function GetLocaleName(<Out> ByRef localeName As String, nameSize As UInteger) As HRESULT
    End Interface

    <ComImport>
    <Guid("5F174B49-0D8B-4CFB-8BCA-F1CCE9D06C67")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteTextFormat1
        Inherits IDWriteTextFormat
#Region "IDWriteTextFormat"
        Overloads Function SetTextAlignment(textAlignment As DWRITE_TEXT_ALIGNMENT) As HRESULT
        Overloads Function SetParagraphAlignment(paragraphAlignment As DWRITE_PARAGRAPH_ALIGNMENT) As HRESULT
        Overloads Function SetWordWrapping(wordWrapping As DWRITE_WORD_WRAPPING) As HRESULT
        Overloads Function SetReadingDirection(readingDirection As DWRITE_READING_DIRECTION) As HRESULT
        Overloads Function SetFlowDirection(flowDirection As DWRITE_FLOW_DIRECTION) As HRESULT
        Overloads Function SetIncrementalTabStop(incrementalTabStop As Single) As HRESULT
        Overloads Function SetTrimming(trimmingOptions As DWRITE_TRIMMING, trimmingSign As IDWriteInlineObject) As HRESULT
        Overloads Function SetLineSpacing(lineSpacingMethod As DWRITE_LINE_SPACING_METHOD, lineSpacing As Single, baseline As Single) As HRESULT
        <PreserveSig>
        Overloads Function GetTextAlignment() As DWRITE_TEXT_ALIGNMENT
        <PreserveSig>
        Overloads Function GetParagraphAlignment() As DWRITE_PARAGRAPH_ALIGNMENT
        <PreserveSig>
        Overloads Function GetWordWrapping() As DWRITE_WORD_WRAPPING
        <PreserveSig>
        Overloads Function GetReadingDirection() As DWRITE_READING_DIRECTION
        <PreserveSig>
        Overloads Function GetFlowDirection() As DWRITE_FLOW_DIRECTION
        <PreserveSig>
        Overloads Function GetIncrementalTabStop() As Single
        Overloads Function GetTrimming(<Out> ByRef trimmingOptions As DWRITE_TRIMMING, <Out> ByRef trimmingSign As IDWriteInlineObject) As HRESULT
        Overloads Function GetLineSpacing(<Out> ByRef lineSpacingMethod As DWRITE_LINE_SPACING_METHOD, <Out> ByRef lineSpacing As Single, <Out> ByRef baseline As Single) As HRESULT
        Overloads Function GetFontCollection(<Out> ByRef fontCollection As IDWriteFontCollection) As HRESULT
        <PreserveSig>
        Overloads Function GetFontFamilyNameLength() As UInteger
        Overloads Function GetFontFamilyName(<Out> ByRef fontFamilyName As String, nameSize As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function GetFontWeight() As DWRITE_FONT_WEIGHT
        <PreserveSig>
        Overloads Function GetFontStyle() As DWRITE_FONT_STYLE
        <PreserveSig>
        Overloads Function GetFontStretch() As DWRITE_FONT_STRETCH
        <PreserveSig>
        Overloads Function GetFontSize() As Single
        <PreserveSig>
        Overloads Function GetLocaleNameLength() As UInteger
        Overloads Function GetLocaleName(<Out> ByRef localeName As String, nameSize As UInteger) As HRESULT
#End Region

        Function SetVerticalGlyphOrientation(glyphOrientation As DWRITE_VERTICAL_GLYPH_ORIENTATION) As HRESULT
        <PreserveSig>
        Function GetVerticalGlyphOrientation() As DWRITE_VERTICAL_GLYPH_ORIENTATION
        Function SetLastLineWrapping(isLastLineWrappingEnabled As Boolean) As HRESULT
        <PreserveSig>
        Function GetLastLineWrapping() As Boolean
        Function SetOpticalAlignment(opticalAlignment As DWRITE_OPTICAL_ALIGNMENT) As HRESULT
        <PreserveSig>
        Function GetOpticalAlignment() As DWRITE_OPTICAL_ALIGNMENT
        Function SetFontFallback(fontFallback As IDWriteFontFallback) As HRESULT
        Function GetFontFallback(<Out> ByRef fontFallback As IDWriteFontFallback) As HRESULT
    End Interface

    <ComImport>
    <Guid("F67E0EDD-9E3D-4ECC-8C32-4183253DFE70")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteTextFormat2
        Inherits IDWriteTextFormat1
#Region "IDWriteTextForma1"
#Region "IDWriteTextFormat"
        Overloads Function SetTextAlignment(textAlignment As DWRITE_TEXT_ALIGNMENT) As HRESULT
        Overloads Function SetParagraphAlignment(paragraphAlignment As DWRITE_PARAGRAPH_ALIGNMENT) As HRESULT
        Overloads Function SetWordWrapping(wordWrapping As DWRITE_WORD_WRAPPING) As HRESULT
        Overloads Function SetReadingDirection(readingDirection As DWRITE_READING_DIRECTION) As HRESULT
        Overloads Function SetFlowDirection(flowDirection As DWRITE_FLOW_DIRECTION) As HRESULT
        Overloads Function SetIncrementalTabStop(incrementalTabStop As Single) As HRESULT
        Overloads Function SetTrimming(trimmingOptions As DWRITE_TRIMMING, trimmingSign As IDWriteInlineObject) As HRESULT
        Overloads Function SetLineSpacing(lineSpacingMethod As DWRITE_LINE_SPACING_METHOD, lineSpacing As Single, baseline As Single) As HRESULT
        <PreserveSig>
        Overloads Function GetTextAlignment() As DWRITE_TEXT_ALIGNMENT
        <PreserveSig>
        Overloads Function GetParagraphAlignment() As DWRITE_PARAGRAPH_ALIGNMENT
        <PreserveSig>
        Overloads Function GetWordWrapping() As DWRITE_WORD_WRAPPING
        <PreserveSig>
        Overloads Function GetReadingDirection() As DWRITE_READING_DIRECTION
        <PreserveSig>
        Overloads Function GetFlowDirection() As DWRITE_FLOW_DIRECTION
        <PreserveSig>
        Overloads Function GetIncrementalTabStop() As Single
        Overloads Function GetTrimming(<Out> ByRef trimmingOptions As DWRITE_TRIMMING, <Out> ByRef trimmingSign As IDWriteInlineObject) As HRESULT
        Overloads Function GetLineSpacing(<Out> ByRef lineSpacingMethod As DWRITE_LINE_SPACING_METHOD, <Out> ByRef lineSpacing As Single, <Out> ByRef baseline As Single) As HRESULT
        Overloads Function GetFontCollection(<Out> ByRef fontCollection As IDWriteFontCollection) As HRESULT
        <PreserveSig>
        Overloads Function GetFontFamilyNameLength() As UInteger
        Overloads Function GetFontFamilyName(<Out> ByRef fontFamilyName As String, nameSize As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function GetFontWeight() As DWRITE_FONT_WEIGHT
        <PreserveSig>
        Overloads Function GetFontStyle() As DWRITE_FONT_STYLE
        <PreserveSig>
        Overloads Function GetFontStretch() As DWRITE_FONT_STRETCH
        <PreserveSig>
        Overloads Function GetFontSize() As Single
        <PreserveSig>
        Overloads Function GetLocaleNameLength() As UInteger
        Overloads Function GetLocaleName(<Out> ByRef localeName As String, nameSize As UInteger) As HRESULT
#End Region

        Overloads Function SetVerticalGlyphOrientation(glyphOrientation As DWRITE_VERTICAL_GLYPH_ORIENTATION) As HRESULT
        <PreserveSig>
        Overloads Function GetVerticalGlyphOrientation() As DWRITE_VERTICAL_GLYPH_ORIENTATION
        Overloads Function SetLastLineWrapping(isLastLineWrappingEnabled As Boolean) As HRESULT
        <PreserveSig>
        Overloads Function GetLastLineWrapping() As Boolean
        Overloads Function SetOpticalAlignment(opticalAlignment As DWRITE_OPTICAL_ALIGNMENT) As HRESULT
        <PreserveSig>
        Overloads Function GetOpticalAlignment() As DWRITE_OPTICAL_ALIGNMENT
        Overloads Function SetFontFallback(fontFallback As IDWriteFontFallback) As HRESULT
        Overloads Function GetFontFallback(<Out> ByRef fontFallback As IDWriteFontFallback) As HRESULT
#End Region

        Overloads Function SetLineSpacing(lineSpacingOptions As DWRITE_LINE_SPACING) As HRESULT
        Overloads Function GetLineSpacing(<Out> ByRef lineSpacingOptions As DWRITE_LINE_SPACING) As HRESULT
    End Interface

    <ComImport>
    <Guid("6D3B5641-E550-430D-A85B-B7BF48A93427")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteTextFormat3
        Inherits IDWriteTextFormat2
#Region "IDWriteTextFormat2"
#Region "IDWriteTextForma1"
#Region "IDWriteTextFormat"
        Overloads Function SetTextAlignment(textAlignment As DWRITE_TEXT_ALIGNMENT) As HRESULT
        Overloads Function SetParagraphAlignment(paragraphAlignment As DWRITE_PARAGRAPH_ALIGNMENT) As HRESULT
        Overloads Function SetWordWrapping(wordWrapping As DWRITE_WORD_WRAPPING) As HRESULT
        Overloads Function SetReadingDirection(readingDirection As DWRITE_READING_DIRECTION) As HRESULT
        Overloads Function SetFlowDirection(flowDirection As DWRITE_FLOW_DIRECTION) As HRESULT
        Overloads Function SetIncrementalTabStop(incrementalTabStop As Single) As HRESULT
        Overloads Function SetTrimming(trimmingOptions As DWRITE_TRIMMING, trimmingSign As IDWriteInlineObject) As HRESULT
        Overloads Function SetLineSpacing(lineSpacingMethod As DWRITE_LINE_SPACING_METHOD, lineSpacing As Single, baseline As Single) As HRESULT
        <PreserveSig>
        Overloads Function GetTextAlignment() As DWRITE_TEXT_ALIGNMENT
        <PreserveSig>
        Overloads Function GetParagraphAlignment() As DWRITE_PARAGRAPH_ALIGNMENT
        <PreserveSig>
        Overloads Function GetWordWrapping() As DWRITE_WORD_WRAPPING
        <PreserveSig>
        Overloads Function GetReadingDirection() As DWRITE_READING_DIRECTION
        <PreserveSig>
        Overloads Function GetFlowDirection() As DWRITE_FLOW_DIRECTION
        <PreserveSig>
        Overloads Function GetIncrementalTabStop() As Single
        Overloads Function GetTrimming(<Out> ByRef trimmingOptions As DWRITE_TRIMMING, <Out> ByRef trimmingSign As IDWriteInlineObject) As HRESULT
        Overloads Function GetLineSpacing(<Out> ByRef lineSpacingMethod As DWRITE_LINE_SPACING_METHOD, <Out> ByRef lineSpacing As Single, <Out> ByRef baseline As Single) As HRESULT
        Overloads Function GetFontCollection(<Out> ByRef fontCollection As IDWriteFontCollection) As HRESULT
        <PreserveSig>
        Overloads Function GetFontFamilyNameLength() As UInteger
        Overloads Function GetFontFamilyName(<Out> ByRef fontFamilyName As String, nameSize As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function GetFontWeight() As DWRITE_FONT_WEIGHT
        <PreserveSig>
        Overloads Function GetFontStyle() As DWRITE_FONT_STYLE
        <PreserveSig>
        Overloads Function GetFontStretch() As DWRITE_FONT_STRETCH
        <PreserveSig>
        Overloads Function GetFontSize() As Single
        <PreserveSig>
        Overloads Function GetLocaleNameLength() As UInteger
        Overloads Function GetLocaleName(<Out> ByRef localeName As String, nameSize As UInteger) As HRESULT
#End Region

        Overloads Function SetVerticalGlyphOrientation(glyphOrientation As DWRITE_VERTICAL_GLYPH_ORIENTATION) As HRESULT
        <PreserveSig>
        Overloads Function GetVerticalGlyphOrientation() As DWRITE_VERTICAL_GLYPH_ORIENTATION
        Overloads Function SetLastLineWrapping(isLastLineWrappingEnabled As Boolean) As HRESULT
        <PreserveSig>
        Overloads Function GetLastLineWrapping() As Boolean
        Overloads Function SetOpticalAlignment(opticalAlignment As DWRITE_OPTICAL_ALIGNMENT) As HRESULT
        <PreserveSig>
        Overloads Function GetOpticalAlignment() As DWRITE_OPTICAL_ALIGNMENT
        Overloads Function SetFontFallback(fontFallback As IDWriteFontFallback) As HRESULT
        Overloads Function GetFontFallback(<Out> ByRef fontFallback As IDWriteFontFallback) As HRESULT
#End Region

        Overloads Function SetLineSpacing(lineSpacingOptions As DWRITE_LINE_SPACING) As HRESULT
        Overloads Function GetLineSpacing(<Out> ByRef lineSpacingOptions As DWRITE_LINE_SPACING) As HRESULT
#End Region

        Function SetFontAxisValues(fontAxisValues As DWRITE_FONT_AXIS_VALUE, fontAxisValueCount As UInteger) As HRESULT
        <PreserveSig>
        Function GetFontAxisValueCount() As UInteger
        Function GetFontAxisValues(<Out> ByRef fontAxisValues As DWRITE_FONT_AXIS_VALUE, fontAxisValueCount As UInteger) As HRESULT
        Function GetAutomaticFontAxes() As DWRITE_AUTOMATIC_FONT_AXES
        Function SetAutomaticFontAxes(automaticFontAxes As DWRITE_AUTOMATIC_FONT_AXES) As HRESULT
    End Interface

    <ComImport>
    <Guid("EFA008F9-F7A1-48BF-B05C-F224713CC0FF")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontFallback
        Function MapCharacters(analysisSource As IDWriteTextAnalysisSource, textPosition As UInteger, textLength As UInteger, baseFontCollection As IDWriteFontCollection, baseFamilyName As String, baseWeight As DWRITE_FONT_WEIGHT, baseStyle As DWRITE_FONT_STYLE, baseStretch As DWRITE_FONT_STRETCH, <Out> ByRef mappedLength As UInteger, <Out> ByRef mappedFont As IDWriteFont, <Out> ByRef scale As Single) As HRESULT
    End Interface

    <ComImport>
    <Guid("FD882D06-8ABA-4FB8-B849-8BE8B73E14DE")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontFallbackBuilder
        Function AddMapping(ranges As DWRITE_UNICODE_RANGE, rangesCount As UInteger, targetFamilyNames As String, targetFamilyNamesCount As UInteger, Optional fontCollection As IDWriteFontCollection = Nothing, Optional localeName As String = Nothing, Optional baseFamilyName As String = Nothing, Optional scale As Single = 1.0F) As HRESULT
        Function AddMappings(fontFallback As IDWriteFontFallback) As HRESULT
        <PreserveSig>
        Function CreateFontFallback(<Out> ByRef fontFallback As IDWriteFontFallback) As HRESULT
    End Interface

    <ComImport>
    <Guid("55f1112b-1dc2-4b3c-9541-f46894ed85b6")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteTypography
        Function AddFontFeature(fontFeature As DWRITE_FONT_FEATURE) As HRESULT
        <PreserveSig>
        Function GetFontFeatureCount() As UInteger
        Function GetFontFeature(fontFeatureIndex As UInteger, <Out> ByRef fontFeature As DWRITE_FONT_FEATURE) As HRESULT
    End Interface

    <ComImport>
    <Guid("1edd9491-9853-4299-898f-6432983b6f3a")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteGdiInterop
        <PreserveSig>
        Function CreateFontFromLOGFONT(logFont As LOGFONT, <Out> ByRef font As IDWriteFont) As HRESULT
        Function ConvertFontToLOGFONT(font As IDWriteFont, <Out> ByRef logFont As LOGFONT, <Out> ByRef isSystemFont As Boolean) As HRESULT
        Function ConvertFontFaceToLOGFONT(font As IDWriteFontFace, <Out> ByRef logFont As LOGFONT) As HRESULT
        <PreserveSig>
        Function CreateFontFaceFromHdc(hdc As IntPtr, <Out> ByRef fontFace As IDWriteFontFace) As HRESULT
        <PreserveSig>
        Function CreateBitmapRenderTarget(hdc As IntPtr, width As Integer, height As Integer, <Out> ByRef renderTarget As IDWriteBitmapRenderTarget) As HRESULT
    End Interface

    <ComImport>
    <Guid("53737037-6d14-410b-9bfe-0b182bb70961")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteTextLayout
        Inherits IDWriteTextFormat
#Region "IDWriteTextFormat"
        Overloads Function SetTextAlignment(textAlignment As DWRITE_TEXT_ALIGNMENT) As HRESULT
        Overloads Function SetParagraphAlignment(paragraphAlignment As DWRITE_PARAGRAPH_ALIGNMENT) As HRESULT
        Overloads Function SetWordWrapping(wordWrapping As DWRITE_WORD_WRAPPING) As HRESULT
        Overloads Function SetReadingDirection(readingDirection As DWRITE_READING_DIRECTION) As HRESULT
        Overloads Function SetFlowDirection(flowDirection As DWRITE_FLOW_DIRECTION) As HRESULT
        Overloads Function SetIncrementalTabStop(incrementalTabStop As Single) As HRESULT
        Overloads Function SetTrimming(trimmingOptions As DWRITE_TRIMMING, trimmingSign As IDWriteInlineObject) As HRESULT
        Overloads Function SetLineSpacing(lineSpacingMethod As DWRITE_LINE_SPACING_METHOD, lineSpacing As Single, baseline As Single) As HRESULT
        <PreserveSig>
        Overloads Function GetTextAlignment() As DWRITE_TEXT_ALIGNMENT
        <PreserveSig>
        Overloads Function GetParagraphAlignment() As DWRITE_PARAGRAPH_ALIGNMENT
        <PreserveSig>
        Overloads Function GetWordWrapping() As DWRITE_WORD_WRAPPING
        <PreserveSig>
        Overloads Function GetReadingDirection() As DWRITE_READING_DIRECTION
        <PreserveSig>
        Overloads Function GetFlowDirection() As DWRITE_FLOW_DIRECTION
        <PreserveSig>
        Overloads Function GetIncrementalTabStop() As Single
        Overloads Function GetTrimming(<Out> ByRef trimmingOptions As DWRITE_TRIMMING, <Out> ByRef trimmingSign As IDWriteInlineObject) As HRESULT
        Overloads Function GetLineSpacing(<Out> ByRef lineSpacingMethod As DWRITE_LINE_SPACING_METHOD, <Out> ByRef lineSpacing As Single, <Out> ByRef baseline As Single) As HRESULT
        Overloads Function GetFontCollection(<Out> ByRef fontCollection As IDWriteFontCollection) As HRESULT
        <PreserveSig>
        Overloads Function GetFontFamilyNameLength() As UInteger
        Overloads Function GetFontFamilyName(<Out> ByRef fontFamilyName As String, nameSize As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function GetFontWeight() As DWRITE_FONT_WEIGHT
        <PreserveSig>
        Overloads Function GetFontStyle() As DWRITE_FONT_STYLE
        <PreserveSig>
        Overloads Function GetFontStretch() As DWRITE_FONT_STRETCH
        <PreserveSig>
        Overloads Function GetFontSize() As Single
        <PreserveSig>
        Overloads Function GetLocaleNameLength() As UInteger
        Overloads Function GetLocaleName(<Out> ByRef localeName As String, nameSize As UInteger) As HRESULT
#End Region

        Function SetMaxWidth(maxWidth As Single) As HRESULT
        Function SetMaxHeight(maxHeight As Single) As HRESULT
        Function SetFontCollection(fontCollection As IDWriteFontCollection, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Function SetFontFamilyName(fontFamilyName As String, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Function SetFontWeight(fontWeight As DWRITE_FONT_WEIGHT, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Function SetFontStyle(fontStyle As DWRITE_FONT_STYLE, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Function SetFontStretch(fontStretch As DWRITE_FONT_STRETCH, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Function SetFontSize(fontSize As Single, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Function SetUnderline(hasUnderline As Boolean, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Function SetStrikethrough(hasStrikethrough As Boolean, textRange As DWRITE_TEXT_RANGE) As HRESULT
        'HRESULT SetDrawingEffect(IUnknown drawingEffect, DWRITE_TEXT_RANGE textRange);
        Function SetDrawingEffect(drawingEffect As IntPtr, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Function SetInlineObject(inlineObject As IDWriteInlineObject, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Function SetTypography(typography As IDWriteTypography, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Function SetLocaleName(localeName As String, textRange As DWRITE_TEXT_RANGE) As HRESULT
        <PreserveSig>
        Function GetMaxWidth() As Single
        <PreserveSig>
        Function GetMaxHeight() As Single
        Overloads Function GetFontCollection(currentPosition As UInteger, <Out> ByRef fontCollection As IDWriteFontCollection, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetFontFamilyNameLength(currentPosition As UInteger, <Out> ByRef nameLength As UInteger, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetFontFamilyName(currentPosition As UInteger, <Out> ByRef fontFamilyName As String, nameSize As UInteger, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetFontWeight(currentPosition As UInteger, <Out> ByRef fontWeight As DWRITE_FONT_WEIGHT, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetFontStyle(currentPosition As UInteger, <Out> ByRef fontStyle As DWRITE_FONT_STYLE, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetFontStretch(currentPosition As UInteger, <Out> ByRef fontStretch As DWRITE_FONT_STRETCH, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetFontSize(currentPosition As UInteger, <Out> ByRef fontSize As Single, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Function GetUnderline(currentPosition As UInteger, <Out> ByRef hasUnderline As Boolean, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Function GetStrikethrough(currentPosition As UInteger, <Out> ByRef hasStrikethrough As Boolean, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        'HRESULT GetDrawingEffect(uint currentPosition, out IUnknown drawingEffect, out DWRITE_TEXT_RANGE textRange);
        Function GetDrawingEffect(currentPosition As UInteger, <Out> ByRef drawingEffect As IntPtr, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Function GetInlineObject(currentPosition As UInteger, <Out> ByRef inlineObject As IDWriteInlineObject, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Function GetTypography(currentPosition As UInteger, <Out> ByRef typography As IDWriteTypography, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetLocaleNameLength(currentPosition As UInteger, <Out> ByRef nameLength As UInteger, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetLocaleName(currentPosition As UInteger, <Out> ByRef localeName As String, nameSize As UInteger, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        <PreserveSig>
        Function Draw(clientDrawingContext As IntPtr, renderer As IDWriteTextRenderer, originX As Single, originY As Single) As HRESULT
        Function GetLineMetrics(<Out> ByRef lineMetrics As DWRITE_LINE_METRICS, maxLineCount As UInteger, <Out> ByRef actualLineCount As UInteger) As HRESULT
        Function GetMetrics(<Out> ByRef textMetrics As DWRITE_TEXT_METRICS) As HRESULT
        Function GetOverhangMetrics(<Out> ByRef overhangs As DWRITE_OVERHANG_METRICS) As HRESULT
        Function GetClusterMetrics(<Out> ByRef clusterMetrics As DWRITE_CLUSTER_METRICS, maxClusterCount As UInteger, <Out> ByRef actualClusterCount As UInteger) As HRESULT
        Function DetermineMinWidth(<Out> ByRef minWidth As Single) As HRESULT
        Function HitTestPoint(pointX As Single, pointY As Single, <Out> ByRef isTrailingHit As Boolean, <Out> ByRef isInside As Boolean, <Out> ByRef hitTestMetrics As DWRITE_HIT_TEST_METRICS) As HRESULT
        Function HitTestTextPosition(textPosition As UInteger, isTrailingHit As Boolean, <Out> ByRef pointX As Single, <Out> ByRef pointY As Single, <Out> ByRef hitTestMetrics As DWRITE_HIT_TEST_METRICS) As HRESULT
        Function HitTestTextRange(textPosition As UInteger, textLength As UInteger, originX As Single, originY As Single, <Out> ByRef hitTestMetrics As DWRITE_HIT_TEST_METRICS, maxHitTestMetricsCount As UInteger, <Out> ByRef actualHitTestMetricsCount As UInteger) As HRESULT
    End Interface

    <ComImport()>
    <Guid("9064D822-80A7-465C-A986-DF65F78B8FEB")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteTextLayout1
        Inherits IDWriteTextLayout
#Region "IDWriteTextLayout"
        Overloads Function SetTextAlignment(textAlignment As DWRITE_TEXT_ALIGNMENT) As HRESULT
        Overloads Function SetParagraphAlignment(paragraphAlignment As DWRITE_PARAGRAPH_ALIGNMENT) As HRESULT
        Overloads Function SetWordWrapping(wordWrapping As DWRITE_WORD_WRAPPING) As HRESULT
        Overloads Function SetReadingDirection(readingDirection As DWRITE_READING_DIRECTION) As HRESULT
        Overloads Function SetFlowDirection(flowDirection As DWRITE_FLOW_DIRECTION) As HRESULT
        Overloads Function SetIncrementalTabStop(incrementalTabStop As Single) As HRESULT
        Overloads Function SetTrimming(trimmingOptions As DWRITE_TRIMMING, trimmingSign As IDWriteInlineObject) As HRESULT
        Overloads Function SetLineSpacing(lineSpacingMethod As DWRITE_LINE_SPACING_METHOD, lineSpacing As Single, baseline As Single) As HRESULT
        <PreserveSig>
        Overloads Function GetTextAlignment() As DWRITE_TEXT_ALIGNMENT
        <PreserveSig>
        Overloads Function GetParagraphAlignment() As DWRITE_PARAGRAPH_ALIGNMENT
        <PreserveSig>
        Overloads Function GetWordWrapping() As DWRITE_WORD_WRAPPING
        <PreserveSig>
        Overloads Function GetReadingDirection() As DWRITE_READING_DIRECTION
        <PreserveSig>
        Overloads Function GetFlowDirection() As DWRITE_FLOW_DIRECTION
        <PreserveSig>
        Overloads Function GetIncrementalTabStop() As Single
        Overloads Function GetTrimming(<Out> ByRef trimmingOptions As DWRITE_TRIMMING, <Out> ByRef trimmingSign As IDWriteInlineObject) As HRESULT
        Overloads Function GetLineSpacing(<Out> ByRef lineSpacingMethod As DWRITE_LINE_SPACING_METHOD, <Out> ByRef lineSpacing As Single, <Out> ByRef baseline As Single) As HRESULT
        Overloads Function GetFontCollection(<Out> ByRef fontCollection As IDWriteFontCollection) As HRESULT
        <PreserveSig>
        Overloads Function GetFontFamilyNameLength() As UInteger
        Overloads Function GetFontFamilyName(<Out> ByRef fontFamilyName As String, nameSize As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function GetFontWeight() As DWRITE_FONT_WEIGHT
        <PreserveSig>
        Overloads Function GetFontStyle() As DWRITE_FONT_STYLE
        <PreserveSig>
        Overloads Function GetFontStretch() As DWRITE_FONT_STRETCH
        <PreserveSig>
        Overloads Function GetFontSize() As Single
        <PreserveSig>
        Overloads Function GetLocaleNameLength() As UInteger
        Overloads Function GetLocaleName(<Out> ByRef localeName As String, nameSize As UInteger) As HRESULT
        Overloads Function SetMaxWidth(maxWidth As Single) As HRESULT
        Overloads Function SetMaxHeight(maxHeight As Single) As HRESULT
        Overloads Function SetFontCollection(fontCollection As IDWriteFontCollection, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetFontFamilyName(fontFamilyName As String, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetFontWeight(fontWeight As DWRITE_FONT_WEIGHT, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetFontStyle(fontStyle As DWRITE_FONT_STYLE, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetFontStretch(fontStretch As DWRITE_FONT_STRETCH, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetFontSize(fontSize As Single, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetUnderline(hasUnderline As Boolean, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetStrikethrough(hasStrikethrough As Boolean, textRange As DWRITE_TEXT_RANGE) As HRESULT

        Overloads Function SetDrawingEffect(drawingEffect As IntPtr, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetInlineObject(inlineObject As IDWriteInlineObject, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetTypography(typography As IDWriteTypography, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetLocaleName(localeName As String, textRange As DWRITE_TEXT_RANGE) As HRESULT
        <PreserveSig>
        Overloads Function GetMaxWidth() As Single
        <PreserveSig>
        Overloads Function GetMaxHeight() As Single
        Overloads Function GetFontCollection(currentPosition As UInteger, <Out> ByRef fontCollection As IDWriteFontCollection, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetFontFamilyNameLength(currentPosition As UInteger, <Out> ByRef nameLength As UInteger, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetFontFamilyName(currentPosition As UInteger, <Out> ByRef fontFamilyName As String, nameSize As UInteger, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetFontWeight(currentPosition As UInteger, <Out> ByRef fontWeight As DWRITE_FONT_WEIGHT, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetFontStyle(currentPosition As UInteger, <Out> ByRef fontStyle As DWRITE_FONT_STYLE, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetFontStretch(currentPosition As UInteger, <Out> ByRef fontStretch As DWRITE_FONT_STRETCH, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetFontSize(currentPosition As UInteger, <Out> ByRef fontSize As Single, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetUnderline(currentPosition As UInteger, <Out> ByRef hasUnderline As Boolean, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetStrikethrough(currentPosition As UInteger, <Out> ByRef hasStrikethrough As Boolean, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT

        Overloads Function GetDrawingEffect(currentPosition As UInteger, <Out> ByRef drawingEffect As IntPtr, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetInlineObject(currentPosition As UInteger, <Out> ByRef inlineObject As IDWriteInlineObject, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetTypography(currentPosition As UInteger, <Out> ByRef typography As IDWriteTypography, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetLocaleNameLength(currentPosition As UInteger, <Out> ByRef nameLength As UInteger, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetLocaleName(currentPosition As UInteger, <Out> ByRef localeName As String, nameSize As UInteger, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function Draw(clientDrawingContext As IntPtr, renderer As IDWriteTextRenderer, originX As Single, originY As Single) As HRESULT
        Overloads Function GetLineMetrics(<Out> ByRef lineMetrics As DWRITE_LINE_METRICS, maxLineCount As UInteger, <Out> ByRef actualLineCount As UInteger) As HRESULT
        Overloads Function GetMetrics(<Out> ByRef textMetrics As DWRITE_TEXT_METRICS) As HRESULT
        Overloads Function GetOverhangMetrics(<Out> ByRef overhangs As DWRITE_OVERHANG_METRICS) As HRESULT
        Overloads Function GetClusterMetrics(<Out> ByRef clusterMetrics As DWRITE_CLUSTER_METRICS, maxClusterCount As UInteger, <Out> ByRef actualClusterCount As UInteger) As HRESULT
        Overloads Function DetermineMinWidth(<Out> ByRef minWidth As Single) As HRESULT
        Overloads Function HitTestPoint(pointX As Single, pointY As Single, <Out> ByRef isTrailingHit As Boolean, <Out> ByRef isInside As Boolean, <Out> ByRef hitTestMetrics As DWRITE_HIT_TEST_METRICS) As HRESULT
        Overloads Function HitTestTextPosition(textPosition As UInteger, isTrailingHit As Boolean, <Out> ByRef pointX As Single, <Out> ByRef pointY As Single, <Out> ByRef hitTestMetrics As DWRITE_HIT_TEST_METRICS) As HRESULT
        Overloads Function HitTestTextRange(textPosition As UInteger, textLength As UInteger, originX As Single, originY As Single, <Out> ByRef hitTestMetrics As DWRITE_HIT_TEST_METRICS, maxHitTestMetricsCount As UInteger, <Out> ByRef actualHitTestMetricsCount As UInteger) As HRESULT
#End Region

        Function SetPairKerning(isPairKerningEnabled As Boolean, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Function GetPairKerning(currentPosition As UInteger, ByRef isPairKerningEnabled As Boolean, ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Function SetCharacterSpacing(leadingSpacing As Single, trailingSpacing As Single, minimumAdvanceWidth As Single, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Function GetCharacterSpacing(currentPosition As UInteger, ByRef leadingSpacing As Single, <Out> ByRef trailingSpacing As Single, <Out> ByRef minimumAdvanceWidth As Single, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
    End Interface

    <ComImport()>
    <Guid("1093C18F-8D5E-43F0-B064-0917311B525E")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteTextLayout2
        Inherits IDWriteTextLayout1
#Region "IDWriteTextLayou1"
#Region "IDWriteTextLayout"
        Overloads Function SetTextAlignment(textAlignment As DWRITE_TEXT_ALIGNMENT) As HRESULT
        Overloads Function SetParagraphAlignment(paragraphAlignment As DWRITE_PARAGRAPH_ALIGNMENT) As HRESULT
        Overloads Function SetWordWrapping(wordWrapping As DWRITE_WORD_WRAPPING) As HRESULT
        Overloads Function SetReadingDirection(readingDirection As DWRITE_READING_DIRECTION) As HRESULT
        Overloads Function SetFlowDirection(flowDirection As DWRITE_FLOW_DIRECTION) As HRESULT
        Overloads Function SetIncrementalTabStop(incrementalTabStop As Single) As HRESULT
        Overloads Function SetTrimming(trimmingOptions As DWRITE_TRIMMING, trimmingSign As IDWriteInlineObject) As HRESULT
        Overloads Function SetLineSpacing(lineSpacingMethod As DWRITE_LINE_SPACING_METHOD, lineSpacing As Single, baseline As Single) As HRESULT
        <PreserveSig>
        Overloads Function GetTextAlignment() As DWRITE_TEXT_ALIGNMENT
        <PreserveSig>
        Overloads Function GetParagraphAlignment() As DWRITE_PARAGRAPH_ALIGNMENT
        <PreserveSig>
        Overloads Function GetWordWrapping() As DWRITE_WORD_WRAPPING
        <PreserveSig>
        Overloads Function GetReadingDirection() As DWRITE_READING_DIRECTION
        <PreserveSig>
        Overloads Function GetFlowDirection() As DWRITE_FLOW_DIRECTION
        <PreserveSig>
        Overloads Function GetIncrementalTabStop() As Single
        Overloads Function GetTrimming(<Out> ByRef trimmingOptions As DWRITE_TRIMMING, <Out> ByRef trimmingSign As IDWriteInlineObject) As HRESULT
        Overloads Function GetLineSpacing(<Out> ByRef lineSpacingMethod As DWRITE_LINE_SPACING_METHOD, <Out> ByRef lineSpacing As Single, <Out> ByRef baseline As Single) As HRESULT
        Overloads Function GetFontCollection(<Out> ByRef fontCollection As IDWriteFontCollection) As HRESULT
        <PreserveSig>
        Overloads Function GetFontFamilyNameLength() As UInteger
        Overloads Function GetFontFamilyName(<Out> ByRef fontFamilyName As String, nameSize As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function GetFontWeight() As DWRITE_FONT_WEIGHT
        <PreserveSig>
        Overloads Function GetFontStyle() As DWRITE_FONT_STYLE
        <PreserveSig>
        Overloads Function GetFontStretch() As DWRITE_FONT_STRETCH
        <PreserveSig>
        Overloads Function GetFontSize() As Single
        <PreserveSig>
        Overloads Function GetLocaleNameLength() As UInteger
        Overloads Function GetLocaleName(<Out> ByRef localeName As String, nameSize As UInteger) As HRESULT
        Overloads Function SetMaxWidth(maxWidth As Single) As HRESULT
        Overloads Function SetMaxHeight(maxHeight As Single) As HRESULT
        Overloads Function SetFontCollection(fontCollection As IDWriteFontCollection, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetFontFamilyName(fontFamilyName As String, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetFontWeight(fontWeight As DWRITE_FONT_WEIGHT, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetFontStyle(fontStyle As DWRITE_FONT_STYLE, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetFontStretch(fontStretch As DWRITE_FONT_STRETCH, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetFontSize(fontSize As Single, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetUnderline(hasUnderline As Boolean, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetStrikethrough(hasStrikethrough As Boolean, textRange As DWRITE_TEXT_RANGE) As HRESULT

        Overloads Function SetDrawingEffect(drawingEffect As IntPtr, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetInlineObject(inlineObject As IDWriteInlineObject, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetTypography(typography As IDWriteTypography, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetLocaleName(localeName As String, textRange As DWRITE_TEXT_RANGE) As HRESULT
        <PreserveSig>
        Overloads Function GetMaxWidth() As Single
        <PreserveSig>
        Overloads Function GetMaxHeight() As Single
        Overloads Function GetFontCollection(currentPosition As UInteger, <Out> ByRef fontCollection As IDWriteFontCollection, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetFontFamilyNameLength(currentPosition As UInteger, <Out> ByRef nameLength As UInteger, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetFontFamilyName(currentPosition As UInteger, <Out> ByRef fontFamilyName As String, nameSize As UInteger, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetFontWeight(currentPosition As UInteger, <Out> ByRef fontWeight As DWRITE_FONT_WEIGHT, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetFontStyle(currentPosition As UInteger, <Out> ByRef fontStyle As DWRITE_FONT_STYLE, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetFontStretch(currentPosition As UInteger, <Out> ByRef fontStretch As DWRITE_FONT_STRETCH, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetFontSize(currentPosition As UInteger, <Out> ByRef fontSize As Single, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetUnderline(currentPosition As UInteger, <Out> ByRef hasUnderline As Boolean, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetStrikethrough(currentPosition As UInteger, <Out> ByRef hasStrikethrough As Boolean, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT

        Overloads Function GetDrawingEffect(currentPosition As UInteger, <Out> ByRef drawingEffect As IntPtr, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetInlineObject(currentPosition As UInteger, <Out> ByRef inlineObject As IDWriteInlineObject, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetTypography(currentPosition As UInteger, <Out> ByRef typography As IDWriteTypography, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetLocaleNameLength(currentPosition As UInteger, <Out> ByRef nameLength As UInteger, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetLocaleName(currentPosition As UInteger, <Out> ByRef localeName As String, nameSize As UInteger, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function Draw(clientDrawingContext As IntPtr, renderer As IDWriteTextRenderer, originX As Single, originY As Single) As HRESULT
        Overloads Function GetLineMetrics(<Out> ByRef lineMetrics As DWRITE_LINE_METRICS, maxLineCount As UInteger, <Out> ByRef actualLineCount As UInteger) As HRESULT
        Overloads Function GetMetrics(<Out> ByRef textMetrics As DWRITE_TEXT_METRICS) As HRESULT
        Overloads Function GetOverhangMetrics(<Out> ByRef overhangs As DWRITE_OVERHANG_METRICS) As HRESULT
        Overloads Function GetClusterMetrics(<Out> ByRef clusterMetrics As DWRITE_CLUSTER_METRICS, maxClusterCount As UInteger, <Out> ByRef actualClusterCount As UInteger) As HRESULT
        Overloads Function DetermineMinWidth(<Out> ByRef minWidth As Single) As HRESULT
        Overloads Function HitTestPoint(pointX As Single, pointY As Single, <Out> ByRef isTrailingHit As Boolean, <Out> ByRef isInside As Boolean, <Out> ByRef hitTestMetrics As DWRITE_HIT_TEST_METRICS) As HRESULT
        Overloads Function HitTestTextPosition(textPosition As UInteger, isTrailingHit As Boolean, <Out> ByRef pointX As Single, <Out> ByRef pointY As Single, <Out> ByRef hitTestMetrics As DWRITE_HIT_TEST_METRICS) As HRESULT
        Overloads Function HitTestTextRange(textPosition As UInteger, textLength As UInteger, originX As Single, originY As Single, <Out> ByRef hitTestMetrics As DWRITE_HIT_TEST_METRICS, maxHitTestMetricsCount As UInteger, <Out> ByRef actualHitTestMetricsCount As UInteger) As HRESULT
#End Region

        Overloads Function SetPairKerning(isPairKerningEnabled As Boolean, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetPairKerning(currentPosition As UInteger, ByRef isPairKerningEnabled As Boolean, ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetCharacterSpacing(leadingSpacing As Single, trailingSpacing As Single, minimumAdvanceWidth As Single, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetCharacterSpacing(currentPosition As UInteger, ByRef leadingSpacing As Single, <Out> ByRef trailingSpacing As Single, <Out> ByRef minimumAdvanceWidth As Single, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
#End Region

        Overloads Function GetMetrics(<Out> ByRef textMetrics As DWRITE_TEXT_METRICS1) As HRESULT
        Function SetVerticalGlyphOrientation(glyphOrientation As DWRITE_VERTICAL_GLYPH_ORIENTATION) As HRESULT
        <PreserveSig>
        Function GetVerticalGlyphOrientation() As DWRITE_VERTICAL_GLYPH_ORIENTATION
        Function SetLastLineWrapping(isLastLineWrappingEnabled As Boolean) As HRESULT
        <PreserveSig>
        Function GetLastLineWrapping() As Boolean
        Function SetOpticalAlignment(opticalAlignment As DWRITE_OPTICAL_ALIGNMENT) As HRESULT
        <PreserveSig>
        Function GetOpticalAlignment() As DWRITE_OPTICAL_ALIGNMENT
        Function SetFontFallback(fontFallback As IDWriteFontFallback) As HRESULT
        Function GetFontFallback(<Out> ByRef fontFallback As IDWriteFontFallback) As HRESULT
    End Interface

    <ComImport()>
    <Guid("07DDCD52-020E-4DE8-AC33-6C953D83F92D")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteTextLayout3
        Inherits IDWriteTextLayout2
#Region "IDWriteTextLayout1"
#Region "IDWriteTextLayout        "
        Overloads Function SetTextAlignment(textAlignment As DWRITE_TEXT_ALIGNMENT) As HRESULT
        Overloads Function SetParagraphAlignment(paragraphAlignment As DWRITE_PARAGRAPH_ALIGNMENT) As HRESULT
        Overloads Function SetWordWrapping(wordWrapping As DWRITE_WORD_WRAPPING) As HRESULT
        Overloads Function SetReadingDirection(readingDirection As DWRITE_READING_DIRECTION) As HRESULT
        Overloads Function SetFlowDirection(flowDirection As DWRITE_FLOW_DIRECTION) As HRESULT
        Overloads Function SetIncrementalTabStop(incrementalTabStop As Single) As HRESULT
        Overloads Function SetTrimming(trimmingOptions As DWRITE_TRIMMING, trimmingSign As IDWriteInlineObject) As HRESULT
        Overloads Function SetLineSpacing(lineSpacingMethod As DWRITE_LINE_SPACING_METHOD, lineSpacing As Single, baseline As Single) As HRESULT
        <PreserveSig>
        Overloads Function GetTextAlignment() As DWRITE_TEXT_ALIGNMENT
        <PreserveSig>
        Overloads Function GetParagraphAlignment() As DWRITE_PARAGRAPH_ALIGNMENT
        <PreserveSig>
        Overloads Function GetWordWrapping() As DWRITE_WORD_WRAPPING
        <PreserveSig>
        Overloads Function GetReadingDirection() As DWRITE_READING_DIRECTION
        <PreserveSig>
        Overloads Function GetFlowDirection() As DWRITE_FLOW_DIRECTION
        <PreserveSig>
        Overloads Function GetIncrementalTabStop() As Single
        Overloads Function GetTrimming(<Out> ByRef trimmingOptions As DWRITE_TRIMMING, <Out> ByRef trimmingSign As IDWriteInlineObject) As HRESULT
        Overloads Function GetLineSpacing(<Out> ByRef lineSpacingMethod As DWRITE_LINE_SPACING_METHOD, <Out> ByRef lineSpacing As Single, <Out> ByRef baseline As Single) As HRESULT
        Overloads Function GetFontCollection(<Out> ByRef fontCollection As IDWriteFontCollection) As HRESULT
        <PreserveSig>
        Overloads Function GetFontFamilyNameLength() As UInteger
        Overloads Function GetFontFamilyName(<Out> ByRef fontFamilyName As String, nameSize As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function GetFontWeight() As DWRITE_FONT_WEIGHT
        <PreserveSig>
        Overloads Function GetFontStyle() As DWRITE_FONT_STYLE
        <PreserveSig>
        Overloads Function GetFontStretch() As DWRITE_FONT_STRETCH
        <PreserveSig>
        Overloads Function GetFontSize() As Single
        <PreserveSig>
        Overloads Function GetLocaleNameLength() As UInteger
        Overloads Function GetLocaleName(<Out> ByRef localeName As String, nameSize As UInteger) As HRESULT
        Overloads Function SetMaxWidth(maxWidth As Single) As HRESULT
        Overloads Function SetMaxHeight(maxHeight As Single) As HRESULT
        Overloads Function SetFontCollection(fontCollection As IDWriteFontCollection, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetFontFamilyName(fontFamilyName As String, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetFontWeight(fontWeight As DWRITE_FONT_WEIGHT, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetFontStyle(fontStyle As DWRITE_FONT_STYLE, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetFontStretch(fontStretch As DWRITE_FONT_STRETCH, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetFontSize(fontSize As Single, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetUnderline(hasUnderline As Boolean, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetStrikethrough(hasStrikethrough As Boolean, textRange As DWRITE_TEXT_RANGE) As HRESULT

        Overloads Function SetDrawingEffect(drawingEffect As IntPtr, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetInlineObject(inlineObject As IDWriteInlineObject, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetTypography(typography As IDWriteTypography, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetLocaleName(localeName As String, textRange As DWRITE_TEXT_RANGE) As HRESULT
        <PreserveSig>
        Overloads Function GetMaxWidth() As Single
        <PreserveSig>
        Overloads Function GetMaxHeight() As Single
        Overloads Function GetFontCollection(currentPosition As UInteger, <Out> ByRef fontCollection As IDWriteFontCollection, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetFontFamilyNameLength(currentPosition As UInteger, <Out> ByRef nameLength As UInteger, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetFontFamilyName(currentPosition As UInteger, <Out> ByRef fontFamilyName As String, nameSize As UInteger, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetFontWeight(currentPosition As UInteger, <Out> ByRef fontWeight As DWRITE_FONT_WEIGHT, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetFontStyle(currentPosition As UInteger, <Out> ByRef fontStyle As DWRITE_FONT_STYLE, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetFontStretch(currentPosition As UInteger, <Out> ByRef fontStretch As DWRITE_FONT_STRETCH, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetFontSize(currentPosition As UInteger, <Out> ByRef fontSize As Single, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetUnderline(currentPosition As UInteger, <Out> ByRef hasUnderline As Boolean, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetStrikethrough(currentPosition As UInteger, <Out> ByRef hasStrikethrough As Boolean, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT

        Overloads Function GetDrawingEffect(currentPosition As UInteger, <Out> ByRef drawingEffect As IntPtr, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetInlineObject(currentPosition As UInteger, <Out> ByRef inlineObject As IDWriteInlineObject, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetTypography(currentPosition As UInteger, <Out> ByRef typography As IDWriteTypography, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetLocaleNameLength(currentPosition As UInteger, <Out> ByRef nameLength As UInteger, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetLocaleName(currentPosition As UInteger, <Out> ByRef localeName As String, nameSize As UInteger, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function Draw(clientDrawingContext As IntPtr, renderer As IDWriteTextRenderer, originX As Single, originY As Single) As HRESULT
        Overloads Function GetLineMetrics(<Out> ByRef lineMetrics As DWRITE_LINE_METRICS, maxLineCount As UInteger, <Out> ByRef actualLineCount As UInteger) As HRESULT
        Overloads Function GetMetrics(<Out> ByRef textMetrics As DWRITE_TEXT_METRICS) As HRESULT
        Overloads Function GetOverhangMetrics(<Out> ByRef overhangs As DWRITE_OVERHANG_METRICS) As HRESULT
        Overloads Function GetClusterMetrics(<Out> ByRef clusterMetrics As DWRITE_CLUSTER_METRICS, maxClusterCount As UInteger, <Out> ByRef actualClusterCount As UInteger) As HRESULT
        Overloads Function DetermineMinWidth(<Out> ByRef minWidth As Single) As HRESULT
        Overloads Function HitTestPoint(pointX As Single, pointY As Single, <Out> ByRef isTrailingHit As Boolean, <Out> ByRef isInside As Boolean, <Out> ByRef hitTestMetrics As DWRITE_HIT_TEST_METRICS) As HRESULT
        Overloads Function HitTestTextPosition(textPosition As UInteger, isTrailingHit As Boolean, <Out> ByRef pointX As Single, <Out> ByRef pointY As Single, <Out> ByRef hitTestMetrics As DWRITE_HIT_TEST_METRICS) As HRESULT
        Overloads Function HitTestTextRange(textPosition As UInteger, textLength As UInteger, originX As Single, originY As Single, <Out> ByRef hitTestMetrics As DWRITE_HIT_TEST_METRICS, maxHitTestMetricsCount As UInteger, <Out> ByRef actualHitTestMetricsCount As UInteger) As HRESULT
#End Region

        Overloads Function SetPairKerning(isPairKerningEnabled As Boolean, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetPairKerning(currentPosition As UInteger, ByRef isPairKerningEnabled As Boolean, ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetCharacterSpacing(leadingSpacing As Single, trailingSpacing As Single, minimumAdvanceWidth As Single, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetCharacterSpacing(currentPosition As UInteger, ByRef leadingSpacing As Single, <Out> ByRef trailingSpacing As Single, <Out> ByRef minimumAdvanceWidth As Single, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
#End Region

#Region "<IDWriteTextLayout2>"
        Overloads Function GetMetrics(<Out> ByRef textMetrics As DWRITE_TEXT_METRICS1) As HRESULT
        Overloads Function SetVerticalGlyphOrientation(glyphOrientation As DWRITE_VERTICAL_GLYPH_ORIENTATION) As HRESULT
        <PreserveSig>
        Overloads Function GetVerticalGlyphOrientation() As DWRITE_VERTICAL_GLYPH_ORIENTATION
        Overloads Function SetLastLineWrapping(isLastLineWrappingEnabled As Boolean) As HRESULT
        <PreserveSig>
        Overloads Function GetLastLineWrapping() As Boolean
        Overloads Function SetOpticalAlignment(opticalAlignment As DWRITE_OPTICAL_ALIGNMENT) As HRESULT
        <PreserveSig>
        Overloads Function GetOpticalAlignment() As DWRITE_OPTICAL_ALIGNMENT
        Overloads Function SetFontFallback(fontFallback As IDWriteFontFallback) As HRESULT
        Overloads Function GetFontFallback(<Out> ByRef fontFallback As IDWriteFontFallback) As HRESULT
#End Region

        <PreserveSig>
        Function InvalidateLayout() As HRESULT
        <PreserveSig>
        Overloads Function SetLineSpacing(ByRef lineSpacingOptions As DWRITE_LINE_SPACING) As HRESULT
        <PreserveSig>
        Overloads Function GetLineSpacing(<Out> ByRef lineSpacingOptions As DWRITE_LINE_SPACING) As HRESULT
        <PreserveSig>
        Overloads Function GetLineMetrics(
            <Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> lineMetrics As DWRITE_LINE_METRICS(), maxLineCount As UInteger, <Out> ByRef actualLineCount As UInteger) As HRESULT
    End Interface

    <ComImport()>
    <Guid("05A9BF42-223F-4441-B5FB-8263685F55E9")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteTextLayout4
        Inherits IDWriteTextLayout3
#Region "IDWriteTextLayout1"
#Region "IDWriteTextLayout        "
        Overloads Function SetTextAlignment(textAlignment As DWRITE_TEXT_ALIGNMENT) As HRESULT
        Overloads Function SetParagraphAlignment(paragraphAlignment As DWRITE_PARAGRAPH_ALIGNMENT) As HRESULT
        Overloads Function SetWordWrapping(wordWrapping As DWRITE_WORD_WRAPPING) As HRESULT
        Overloads Function SetReadingDirection(readingDirection As DWRITE_READING_DIRECTION) As HRESULT
        Overloads Function SetFlowDirection(flowDirection As DWRITE_FLOW_DIRECTION) As HRESULT
        Overloads Function SetIncrementalTabStop(incrementalTabStop As Single) As HRESULT
        Overloads Function SetTrimming(trimmingOptions As DWRITE_TRIMMING, trimmingSign As IDWriteInlineObject) As HRESULT
        Overloads Function SetLineSpacing(lineSpacingMethod As DWRITE_LINE_SPACING_METHOD, lineSpacing As Single, baseline As Single) As HRESULT
        <PreserveSig>
        Overloads Function GetTextAlignment() As DWRITE_TEXT_ALIGNMENT
        <PreserveSig>
        Overloads Function GetParagraphAlignment() As DWRITE_PARAGRAPH_ALIGNMENT
        <PreserveSig>
        Overloads Function GetWordWrapping() As DWRITE_WORD_WRAPPING
        <PreserveSig>
        Overloads Function GetReadingDirection() As DWRITE_READING_DIRECTION
        <PreserveSig>
        Overloads Function GetFlowDirection() As DWRITE_FLOW_DIRECTION
        <PreserveSig>
        Overloads Function GetIncrementalTabStop() As Single
        Overloads Function GetTrimming(<Out> ByRef trimmingOptions As DWRITE_TRIMMING, <Out> ByRef trimmingSign As IDWriteInlineObject) As HRESULT
        Overloads Function GetLineSpacing(<Out> ByRef lineSpacingMethod As DWRITE_LINE_SPACING_METHOD, <Out> ByRef lineSpacing As Single, <Out> ByRef baseline As Single) As HRESULT
        Overloads Function GetFontCollection(<Out> ByRef fontCollection As IDWriteFontCollection) As HRESULT
        <PreserveSig>
        Overloads Function GetFontFamilyNameLength() As UInteger
        Overloads Function GetFontFamilyName(<Out> ByRef fontFamilyName As String, nameSize As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function GetFontWeight() As DWRITE_FONT_WEIGHT
        <PreserveSig>
        Overloads Function GetFontStyle() As DWRITE_FONT_STYLE
        <PreserveSig>
        Overloads Function GetFontStretch() As DWRITE_FONT_STRETCH
        <PreserveSig>
        Overloads Function GetFontSize() As Single
        <PreserveSig>
        Overloads Function GetLocaleNameLength() As UInteger
        Overloads Function GetLocaleName(<Out> ByRef localeName As String, nameSize As UInteger) As HRESULT
        Overloads Function SetMaxWidth(maxWidth As Single) As HRESULT
        Overloads Function SetMaxHeight(maxHeight As Single) As HRESULT
        Overloads Function SetFontCollection(fontCollection As IDWriteFontCollection, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetFontFamilyName(fontFamilyName As String, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetFontWeight(fontWeight As DWRITE_FONT_WEIGHT, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetFontStyle(fontStyle As DWRITE_FONT_STYLE, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetFontStretch(fontStretch As DWRITE_FONT_STRETCH, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetFontSize(fontSize As Single, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetUnderline(hasUnderline As Boolean, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetStrikethrough(hasStrikethrough As Boolean, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetDrawingEffect(drawingEffect As IntPtr, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetInlineObject(inlineObject As IDWriteInlineObject, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetTypography(typography As IDWriteTypography, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetLocaleName(localeName As String, textRange As DWRITE_TEXT_RANGE) As HRESULT
        <PreserveSig>
        Overloads Function GetMaxWidth() As Single
        <PreserveSig>
        Overloads Function GetMaxHeight() As Single
        Overloads Function GetFontCollection(currentPosition As UInteger, <Out> ByRef fontCollection As IDWriteFontCollection, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetFontFamilyNameLength(currentPosition As UInteger, <Out> ByRef nameLength As UInteger, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetFontFamilyName(currentPosition As UInteger, <Out> ByRef fontFamilyName As String, nameSize As UInteger, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetFontWeight(currentPosition As UInteger, <Out> ByRef fontWeight As DWRITE_FONT_WEIGHT, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetFontStyle(currentPosition As UInteger, <Out> ByRef fontStyle As DWRITE_FONT_STYLE, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetFontStretch(currentPosition As UInteger, <Out> ByRef fontStretch As DWRITE_FONT_STRETCH, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetFontSize(currentPosition As UInteger, <Out> ByRef fontSize As Single, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetUnderline(currentPosition As UInteger, <Out> ByRef hasUnderline As Boolean, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetStrikethrough(currentPosition As UInteger, <Out> ByRef hasStrikethrough As Boolean, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT

        Overloads Function GetDrawingEffect(currentPosition As UInteger, <Out> ByRef drawingEffect As IntPtr, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetInlineObject(currentPosition As UInteger, <Out> ByRef inlineObject As IDWriteInlineObject, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetTypography(currentPosition As UInteger, <Out> ByRef typography As IDWriteTypography, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetLocaleNameLength(currentPosition As UInteger, <Out> ByRef nameLength As UInteger, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetLocaleName(currentPosition As UInteger, <Out> ByRef localeName As String, nameSize As UInteger, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function Draw(clientDrawingContext As IntPtr, renderer As IDWriteTextRenderer, originX As Single, originY As Single) As HRESULT
        Overloads Function GetLineMetrics(<Out> ByRef lineMetrics As DWRITE_LINE_METRICS, maxLineCount As UInteger, <Out> ByRef actualLineCount As UInteger) As HRESULT
        Overloads Function GetMetrics(<Out> ByRef textMetrics As DWRITE_TEXT_METRICS) As HRESULT
        Overloads Function GetOverhangMetrics(<Out> ByRef overhangs As DWRITE_OVERHANG_METRICS) As HRESULT
        Overloads Function GetClusterMetrics(<Out> ByRef clusterMetrics As DWRITE_CLUSTER_METRICS, maxClusterCount As UInteger, <Out> ByRef actualClusterCount As UInteger) As HRESULT
        Overloads Function DetermineMinWidth(<Out> ByRef minWidth As Single) As HRESULT
        Overloads Function HitTestPoint(pointX As Single, pointY As Single, <Out> ByRef isTrailingHit As Boolean, <Out> ByRef isInside As Boolean, <Out> ByRef hitTestMetrics As DWRITE_HIT_TEST_METRICS) As HRESULT
        Overloads Function HitTestTextPosition(textPosition As UInteger, isTrailingHit As Boolean, <Out> ByRef pointX As Single, <Out> ByRef pointY As Single, <Out> ByRef hitTestMetrics As DWRITE_HIT_TEST_METRICS) As HRESULT
        Overloads Function HitTestTextRange(textPosition As UInteger, textLength As UInteger, originX As Single, originY As Single, <Out> ByRef hitTestMetrics As DWRITE_HIT_TEST_METRICS, maxHitTestMetricsCount As UInteger, <Out> ByRef actualHitTestMetricsCount As UInteger) As HRESULT
#End Region

        Overloads Function SetPairKerning(isPairKerningEnabled As Boolean, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetPairKerning(currentPosition As UInteger, ByRef isPairKerningEnabled As Boolean, ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function SetCharacterSpacing(leadingSpacing As Single, trailingSpacing As Single, minimumAdvanceWidth As Single, textRange As DWRITE_TEXT_RANGE) As HRESULT
        Overloads Function GetCharacterSpacing(currentPosition As UInteger, ByRef leadingSpacing As Single, <Out> ByRef trailingSpacing As Single, <Out> ByRef minimumAdvanceWidth As Single, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
#End Region

#Region "<IDWriteTextLayout2>"
        Overloads Function GetMetrics(<Out> ByRef textMetrics As DWRITE_TEXT_METRICS1) As HRESULT
        Overloads Function SetVerticalGlyphOrientation(glyphOrientation As DWRITE_VERTICAL_GLYPH_ORIENTATION) As HRESULT
        <PreserveSig>
        Overloads Function GetVerticalGlyphOrientation() As DWRITE_VERTICAL_GLYPH_ORIENTATION
        Overloads Function SetLastLineWrapping(isLastLineWrappingEnabled As Boolean) As HRESULT
        <PreserveSig>
        Overloads Function GetLastLineWrapping() As Boolean
        Overloads Function SetOpticalAlignment(opticalAlignment As DWRITE_OPTICAL_ALIGNMENT) As HRESULT
        <PreserveSig>
        Overloads Function GetOpticalAlignment() As DWRITE_OPTICAL_ALIGNMENT
        Overloads Function SetFontFallback(fontFallback As IDWriteFontFallback) As HRESULT
        Overloads Function GetFontFallback(<Out> ByRef fontFallback As IDWriteFontFallback) As HRESULT
#End Region

#Region "<IDWriteTextLayout3>"
        <PreserveSig>
        Overloads Function InvalidateLayout() As HRESULT
        <PreserveSig>
        Overloads Function SetLineSpacing(ByRef lineSpacingOptions As DWRITE_LINE_SPACING) As HRESULT
        <PreserveSig>
        Overloads Function GetLineSpacing(<Out> ByRef lineSpacingOptions As DWRITE_LINE_SPACING) As HRESULT
        <PreserveSig>
        Overloads Function GetLineMetrics(
            <Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> lineMetrics As DWRITE_LINE_METRICS(), maxLineCount As UInteger, <Out> ByRef actualLineCount As UInteger) As HRESULT
#End Region

        <PreserveSig>
        Function SetFontAxisValues(
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> fontAxisValues As DWRITE_FONT_AXIS_VALUE(), fontAxisValueCount As UInteger, textRange As DWRITE_TEXT_RANGE) As HRESULT
        <PreserveSig>
        Function GetFontAxisValueCount(currentPosition As UInteger) As UInteger
        <PreserveSig>
        Function GetFontAxisValues(currentPosition As UInteger,
        <Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=2)> fontAxisValues As DWRITE_FONT_AXIS_VALUE(), fontAxisValueCount As UInteger, textRange As IntPtr) As HRESULT
        <PreserveSig>
        Function GetAutomaticFontAxes() As DWRITE_AUTOMATIC_FONT_AXES
        <PreserveSig>
        Function SetAutomaticFontAxes(automaticFontAxes As DWRITE_AUTOMATIC_FONT_AXES) As HRESULT
    End Interface

    <ComImport>
    <Guid("8339FDE3-106F-47ab-8373-1C6295EB10B3")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteInlineObject
        'HRESULT Draw(IntPtr clientDrawingContext, IDWriteTextRenderer renderer, float originX, float originY, bool isSideways, bool isRightToLeft, IUnknown* clientDrawingEffect);
        Function Draw(clientDrawingContext As IntPtr, renderer As IDWriteTextRenderer, originX As Single, originY As Single, isSideways As Boolean, isRightToLeft As Boolean, clientDrawingEffect As IntPtr) As HRESULT
        Function GetMetrics(<Out> ByRef metrics As DWRITE_INLINE_OBJECT_METRICS) As HRESULT
        Function GetOverhangMetrics(<Out> ByRef overhangs As DWRITE_OVERHANG_METRICS) As HRESULT
        Function GetBreakConditions(<Out> ByRef breakConditionBefore As DWRITE_BREAK_CONDITION, <Out> ByRef breakConditionAfter As DWRITE_BREAK_CONDITION) As HRESULT
    End Interface

    <ComImport>
    <Guid("b7e6163e-7f46-43b4-84b3-e4e6249c365d")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteTextAnalyzer
        Function AnalyzeScript(analysisSource As IDWriteTextAnalysisSource, textPosition As Integer, textLength As Integer, analysisSink As IDWriteTextAnalysisSink) As HRESULT
        Function AnalyzeBidi(analysisSource As IDWriteTextAnalysisSource, textPosition As Integer, textLength As Integer, analysisSink As IDWriteTextAnalysisSink) As HRESULT
        Function AnalyzeNumberSubstitution(analysisSource As IDWriteTextAnalysisSource, textPosition As Integer, textLength As Integer, analysisSink As IDWriteTextAnalysisSink) As HRESULT
        Function AnalyzeLineBreakpoints(analysisSource As IDWriteTextAnalysisSource, textPosition As Integer, textLength As Integer, analysisSink As IDWriteTextAnalysisSink) As HRESULT
        Function GetGlyphs(textString As String, textLength As Integer, fontFace As IDWriteFontFace, isSideways As Boolean, isRightToLeft As Boolean, scriptAnalysis As DWRITE_SCRIPT_ANALYSIS, localeName As String, numberSubstitution As IDWriteNumberSubstitution, features As DWRITE_TYPOGRAPHIC_FEATURES, featureRangeLengths As Integer, featureRanges As Integer, maxGlyphCount As Integer, <Out> ByRef clusterMap As UShort, <Out> ByRef textProps As DWRITE_SHAPING_TEXT_PROPERTIES, <Out> ByRef glyphIndices As UShort, <Out> ByRef glyphProps As DWRITE_SHAPING_GLYPH_PROPERTIES, <Out> ByRef actualGlyphCount As Integer) As HRESULT
        Function GetGlyphPlacements(textString As String, clusterMap As UShort, textProps As DWRITE_SHAPING_TEXT_PROPERTIES, textLength As Integer, glyphIndices As UShort, glyphProps As DWRITE_SHAPING_GLYPH_PROPERTIES, glyphCount As Integer, fontFace As IDWriteFontFace, fontEmSize As Single, isSideways As Boolean, isRightToLeft As Boolean, scriptAnalysis As DWRITE_SCRIPT_ANALYSIS, localeName As String, features As DWRITE_TYPOGRAPHIC_FEATURES, featureRangeLengths As Integer, featureRanges As Integer, <Out> ByRef glyphAdvances As Single, <Out> ByRef glyphOffsets As DWRITE_GLYPH_OFFSET) As HRESULT
        Function GetGdiCompatibleGlyphPlacements(textString As String, clusterMap As UShort, textProps As DWRITE_SHAPING_TEXT_PROPERTIES, textLength As Integer, glyphIndices As UShort, glyphProps As DWRITE_SHAPING_GLYPH_PROPERTIES, glyphCount As Integer, fontFace As IDWriteFontFace, fontEmSize As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, useGdiNatural As Boolean, isSideways As Boolean, isRightToLeft As Boolean, scriptAnalysis As DWRITE_SCRIPT_ANALYSIS, localeName As String, features As DWRITE_TYPOGRAPHIC_FEATURES, featureRangeLengths As Integer, featureRanges As Integer, <Out> ByRef glyphAdvances As Single, <Out> ByRef glyphOffsets As DWRITE_GLYPH_OFFSET) As HRESULT
    End Interface

    <ComImport>
    <Guid("80DAD800-E21F-4E83-96CE-BFCCE500DB7C")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteTextAnalyzer1
        Inherits IDWriteTextAnalyzer
#Region "IDWriteTextAnalyzer"
        Overloads Function AnalyzeScript(analysisSource As IDWriteTextAnalysisSource, textPosition As Integer, textLength As Integer, analysisSink As IDWriteTextAnalysisSink) As HRESULT
        Overloads Function AnalyzeBidi(analysisSource As IDWriteTextAnalysisSource, textPosition As Integer, textLength As Integer, analysisSink As IDWriteTextAnalysisSink) As HRESULT
        Overloads Function AnalyzeNumberSubstitution(analysisSource As IDWriteTextAnalysisSource, textPosition As Integer, textLength As Integer, analysisSink As IDWriteTextAnalysisSink) As HRESULT
        Overloads Function AnalyzeLineBreakpoints(analysisSource As IDWriteTextAnalysisSource, textPosition As Integer, textLength As Integer, analysisSink As IDWriteTextAnalysisSink) As HRESULT
        Overloads Function GetGlyphs(textString As String, textLength As Integer, fontFace As IDWriteFontFace, isSideways As Boolean, isRightToLeft As Boolean, scriptAnalysis As DWRITE_SCRIPT_ANALYSIS, localeName As String, numberSubstitution As IDWriteNumberSubstitution, features As DWRITE_TYPOGRAPHIC_FEATURES, featureRangeLengths As Integer, featureRanges As Integer, maxGlyphCount As Integer, <Out> ByRef clusterMap As UShort, <Out> ByRef textProps As DWRITE_SHAPING_TEXT_PROPERTIES, <Out> ByRef glyphIndices As UShort, <Out> ByRef glyphProps As DWRITE_SHAPING_GLYPH_PROPERTIES, <Out> ByRef actualGlyphCount As Integer) As HRESULT
        Overloads Function GetGlyphPlacements(textString As String, clusterMap As UShort, textProps As DWRITE_SHAPING_TEXT_PROPERTIES, textLength As Integer, glyphIndices As UShort, glyphProps As DWRITE_SHAPING_GLYPH_PROPERTIES, glyphCount As Integer, fontFace As IDWriteFontFace, fontEmSize As Single, isSideways As Boolean, isRightToLeft As Boolean, scriptAnalysis As DWRITE_SCRIPT_ANALYSIS, localeName As String, features As DWRITE_TYPOGRAPHIC_FEATURES, featureRangeLengths As Integer, featureRanges As Integer, <Out> ByRef glyphAdvances As Single, <Out> ByRef glyphOffsets As DWRITE_GLYPH_OFFSET) As HRESULT
        Overloads Function GetGdiCompatibleGlyphPlacements(textString As String, clusterMap As UShort, textProps As DWRITE_SHAPING_TEXT_PROPERTIES, textLength As Integer, glyphIndices As UShort, glyphProps As DWRITE_SHAPING_GLYPH_PROPERTIES, glyphCount As Integer, fontFace As IDWriteFontFace, fontEmSize As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, useGdiNatural As Boolean, isSideways As Boolean, isRightToLeft As Boolean, scriptAnalysis As DWRITE_SCRIPT_ANALYSIS, localeName As String, features As DWRITE_TYPOGRAPHIC_FEATURES, featureRangeLengths As Integer, featureRanges As Integer, <Out> ByRef glyphAdvances As Single, <Out> ByRef glyphOffsets As DWRITE_GLYPH_OFFSET) As HRESULT
#End Region

        Function ApplyCharacterSpacing(leadingSpacing As Single, trailingSpacing As Single, minimumAdvanceWidth As Single, textLength As UInteger, glyphCount As UInteger, clusterMap As UShort, glyphAdvances As Single, glyphOffsets As DWRITE_GLYPH_OFFSET, glyphProperties As DWRITE_SHAPING_GLYPH_PROPERTIES, <Out> ByRef modifiedGlyphAdvances As Single, <Out> ByRef modifiedGlyphOffsets As DWRITE_GLYPH_OFFSET) As HRESULT
        Function GetBaseline(fontFace As IDWriteFontFace, baseline As DWRITE_BASELINE, isVertical As Boolean, isSimulationAllowed As Boolean, scriptAnalysis As DWRITE_SCRIPT_ANALYSIS, localeName As String, <Out> ByRef baselineCoordinate As Integer, <Out> ByRef exists As Boolean) As HRESULT
        Function AnalyzeVerticalGlyphOrientation(analysisSource As IDWriteTextAnalysisSource1, textPosition As UInteger, textLength As UInteger, analysisSink As IDWriteTextAnalysisSink1) As HRESULT
        Function GetGlyphOrientationTransform(glyphOrientationAngle As DWRITE_GLYPH_ORIENTATION_ANGLE, isSideways As Boolean, <Out> ByRef transform As DWRITE_MATRIX) As HRESULT
        Function GetScriptProperties(scriptAnalysis As DWRITE_SCRIPT_ANALYSIS, <Out> ByRef scriptProperties As DWRITE_SCRIPT_PROPERTIES) As HRESULT
        Function GetTextComplexity(textString As String, textLength As UInteger, fontFace As IDWriteFontFace, <Out> ByRef isTextSimple As Boolean, <Out> ByRef textLengthRead As UInteger, <Out> ByRef glyphIndices As UShort) As HRESULT
        Function GetJustificationOpportunities(fontFace As IDWriteFontFace, fontEmSize As Single, scriptAnalysis As DWRITE_SCRIPT_ANALYSIS, textLength As UInteger, glyphCount As UInteger, textString As String, clusterMap As UShort, glyphProperties As DWRITE_SHAPING_GLYPH_PROPERTIES, <Out> ByRef justificationOpportunities As DWRITE_JUSTIFICATION_OPPORTUNITY) As HRESULT
        Function JustifyGlyphAdvances(lineWidth As Single, glyphCount As UInteger, justificationOpportunities As DWRITE_JUSTIFICATION_OPPORTUNITY, glyphAdvances As Single, glyphOffsets As DWRITE_GLYPH_OFFSET, <Out> ByRef justifiedGlyphAdvances As Single, <Out> ByRef justifiedGlyphOffsets As DWRITE_GLYPH_OFFSET) As HRESULT
        Function GetJustifiedGlyphs(fontFace As IDWriteFontFace, fontEmSize As Single, scriptAnalysis As DWRITE_SCRIPT_ANALYSIS, textLength As UInteger, glyphCount As UInteger, maxGlyphCount As UInteger, clusterMap As UShort, glyphIndices As UShort, glyphAdvances As Single, justifiedGlyphAdvances As Single, justifiedGlyphOffsets As DWRITE_GLYPH_OFFSET, glyphProperties As DWRITE_SHAPING_GLYPH_PROPERTIES, <Out> ByRef actualGlyphCount As UInteger, <Out> ByRef modifiedClusterMap As UShort, <Out> ByRef modifiedGlyphIndices As UShort, <Out> ByRef modifiedGlyphAdvances As Single, <Out> ByRef modifiedGlyphOffsets As DWRITE_GLYPH_OFFSET) As HRESULT
    End Interface

    <ComImport>
    <Guid("553A9FF3-5693-4DF7-B52B-74806F7F2EB9")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteTextAnalyzer2
        Inherits IDWriteTextAnalyzer1
#Region "IDWriteTextAnalyzer1"
#Region "IDWriteTextAnalyzer"
        Overloads Function AnalyzeScript(analysisSource As IDWriteTextAnalysisSource, textPosition As Integer, textLength As Integer, analysisSink As IDWriteTextAnalysisSink) As HRESULT
        Overloads Function AnalyzeBidi(analysisSource As IDWriteTextAnalysisSource, textPosition As Integer, textLength As Integer, analysisSink As IDWriteTextAnalysisSink) As HRESULT
        Overloads Function AnalyzeNumberSubstitution(analysisSource As IDWriteTextAnalysisSource, textPosition As Integer, textLength As Integer, analysisSink As IDWriteTextAnalysisSink) As HRESULT
        Overloads Function AnalyzeLineBreakpoints(analysisSource As IDWriteTextAnalysisSource, textPosition As Integer, textLength As Integer, analysisSink As IDWriteTextAnalysisSink) As HRESULT
        Overloads Function GetGlyphs(textString As String, textLength As Integer, fontFace As IDWriteFontFace, isSideways As Boolean, isRightToLeft As Boolean, scriptAnalysis As DWRITE_SCRIPT_ANALYSIS, localeName As String, numberSubstitution As IDWriteNumberSubstitution, features As DWRITE_TYPOGRAPHIC_FEATURES, featureRangeLengths As Integer, featureRanges As Integer, maxGlyphCount As Integer, <Out> ByRef clusterMap As UShort, <Out> ByRef textProps As DWRITE_SHAPING_TEXT_PROPERTIES, <Out> ByRef glyphIndices As UShort, <Out> ByRef glyphProps As DWRITE_SHAPING_GLYPH_PROPERTIES, <Out> ByRef actualGlyphCount As Integer) As HRESULT
        Overloads Function GetGlyphPlacements(textString As String, clusterMap As UShort, textProps As DWRITE_SHAPING_TEXT_PROPERTIES, textLength As Integer, glyphIndices As UShort, glyphProps As DWRITE_SHAPING_GLYPH_PROPERTIES, glyphCount As Integer, fontFace As IDWriteFontFace, fontEmSize As Single, isSideways As Boolean, isRightToLeft As Boolean, scriptAnalysis As DWRITE_SCRIPT_ANALYSIS, localeName As String, features As DWRITE_TYPOGRAPHIC_FEATURES, featureRangeLengths As Integer, featureRanges As Integer, <Out> ByRef glyphAdvances As Single, <Out> ByRef glyphOffsets As DWRITE_GLYPH_OFFSET) As HRESULT
        Overloads Function GetGdiCompatibleGlyphPlacements(textString As String, clusterMap As UShort, textProps As DWRITE_SHAPING_TEXT_PROPERTIES, textLength As Integer, glyphIndices As UShort, glyphProps As DWRITE_SHAPING_GLYPH_PROPERTIES, glyphCount As Integer, fontFace As IDWriteFontFace, fontEmSize As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, useGdiNatural As Boolean, isSideways As Boolean, isRightToLeft As Boolean, scriptAnalysis As DWRITE_SCRIPT_ANALYSIS, localeName As String, features As DWRITE_TYPOGRAPHIC_FEATURES, featureRangeLengths As Integer, featureRanges As Integer, <Out> ByRef glyphAdvances As Single, <Out> ByRef glyphOffsets As DWRITE_GLYPH_OFFSET) As HRESULT
#End Region

        Overloads Function ApplyCharacterSpacing(leadingSpacing As Single, trailingSpacing As Single, minimumAdvanceWidth As Single, textLength As UInteger, glyphCount As UInteger, clusterMap As UShort, glyphAdvances As Single, glyphOffsets As DWRITE_GLYPH_OFFSET, glyphProperties As DWRITE_SHAPING_GLYPH_PROPERTIES, <Out> ByRef modifiedGlyphAdvances As Single, <Out> ByRef modifiedGlyphOffsets As DWRITE_GLYPH_OFFSET) As HRESULT
        Overloads Function GetBaseline(fontFace As IDWriteFontFace, baseline As DWRITE_BASELINE, isVertical As Boolean, isSimulationAllowed As Boolean, scriptAnalysis As DWRITE_SCRIPT_ANALYSIS, localeName As String, <Out> ByRef baselineCoordinate As Integer, <Out> ByRef exists As Boolean) As HRESULT
        Overloads Function AnalyzeVerticalGlyphOrientation(analysisSource As IDWriteTextAnalysisSource1, textPosition As UInteger, textLength As UInteger, analysisSink As IDWriteTextAnalysisSink1) As HRESULT
        Overloads Function GetGlyphOrientationTransform(glyphOrientationAngle As DWRITE_GLYPH_ORIENTATION_ANGLE, isSideways As Boolean, <Out> ByRef transform As DWRITE_MATRIX) As HRESULT
        Overloads Function GetScriptProperties(scriptAnalysis As DWRITE_SCRIPT_ANALYSIS, <Out> ByRef scriptProperties As DWRITE_SCRIPT_PROPERTIES) As HRESULT
        Overloads Function GetTextComplexity(textString As String, textLength As UInteger, fontFace As IDWriteFontFace, <Out> ByRef isTextSimple As Boolean, <Out> ByRef textLengthRead As UInteger, <Out> ByRef glyphIndices As UShort) As HRESULT
        Overloads Function GetJustificationOpportunities(fontFace As IDWriteFontFace, fontEmSize As Single, scriptAnalysis As DWRITE_SCRIPT_ANALYSIS, textLength As UInteger, glyphCount As UInteger, textString As String, clusterMap As UShort, glyphProperties As DWRITE_SHAPING_GLYPH_PROPERTIES, <Out> ByRef justificationOpportunities As DWRITE_JUSTIFICATION_OPPORTUNITY) As HRESULT
        Overloads Function JustifyGlyphAdvances(lineWidth As Single, glyphCount As UInteger, justificationOpportunities As DWRITE_JUSTIFICATION_OPPORTUNITY, glyphAdvances As Single, glyphOffsets As DWRITE_GLYPH_OFFSET, <Out> ByRef justifiedGlyphAdvances As Single, <Out> ByRef justifiedGlyphOffsets As DWRITE_GLYPH_OFFSET) As HRESULT
        Overloads Function GetJustifiedGlyphs(fontFace As IDWriteFontFace, fontEmSize As Single, scriptAnalysis As DWRITE_SCRIPT_ANALYSIS, textLength As UInteger, glyphCount As UInteger, maxGlyphCount As UInteger, clusterMap As UShort, glyphIndices As UShort, glyphAdvances As Single, justifiedGlyphAdvances As Single, justifiedGlyphOffsets As DWRITE_GLYPH_OFFSET, glyphProperties As DWRITE_SHAPING_GLYPH_PROPERTIES, <Out> ByRef actualGlyphCount As UInteger, <Out> ByRef modifiedClusterMap As UShort, <Out> ByRef modifiedGlyphIndices As UShort, <Out> ByRef modifiedGlyphAdvances As Single, <Out> ByRef modifiedGlyphOffsets As DWRITE_GLYPH_OFFSET) As HRESULT
#End Region

        Overloads Function GetGlyphOrientationTransform(glyphOrientationAngle As DWRITE_GLYPH_ORIENTATION_ANGLE, isSideways As Boolean, originX As Single, originY As Single, <Out> ByRef transform As DWRITE_MATRIX) As HRESULT
        Function GetTypographicFeatures(fontFace As IDWriteFontFace, scriptAnalysis As DWRITE_SCRIPT_ANALYSIS, localeName As String, maxTagCount As UInteger, <Out> ByRef actualTagCount As UInteger, <Out> ByRef tags As DWRITE_FONT_FEATURE_TAG) As HRESULT
        Function CheckTypographicFeature(fontFace As IDWriteFontFace, scriptAnalysis As DWRITE_SCRIPT_ANALYSIS, localeName As String, featureTag As DWRITE_FONT_FEATURE_TAG, glyphCount As UInteger, glyphIndices As UShort, <Out> ByRef featureApplies As Byte) As HRESULT
    End Interface

    <ComImport>
    <Guid("acd16696-8c14-4f5d-877e-fe3fc1d32737")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFont
        Function GetFontFamily(<Out> ByRef fontFamily As IDWriteFontFamily) As HRESULT
        <PreserveSig>
        Function GetWeight() As DWRITE_FONT_WEIGHT
        <PreserveSig>
        Function GetStretch() As DWRITE_FONT_STRETCH
        <PreserveSig>
        Function GetStyle() As DWRITE_FONT_STYLE
        <PreserveSig>
        Function IsSymbolFont() As Boolean
        Function GetFaceNames(<Out> ByRef names As IDWriteLocalizedStrings) As HRESULT
        Function GetInformationalStrings(informationalStringID As DWRITE_INFORMATIONAL_STRING_ID, <Out> ByRef informationalStrings As IDWriteLocalizedStrings, <Out> ByRef exists As Boolean) As HRESULT
        <PreserveSig>
        Function GetSimulations() As DWRITE_FONT_SIMULATIONS
        Sub GetMetrics(<Out> ByRef fontMetrics As DWRITE_FONT_METRICS)
        Function HasCharacter(unicodeValue As Integer, <Out> ByRef exists As Boolean) As HRESULT
        <PreserveSig>
        Function CreateFontFace(<Out> ByRef fontFace As IDWriteFontFace) As HRESULT
    End Interface

    <ComImport>
    <Guid("acd16696-8c14-4f5d-877e-fe3fc1d32738")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFont1
        Inherits IDWriteFont
#Region "IDWriteFont"
        Overloads Function GetFontFamily(<Out> ByRef fontFamily As IDWriteFontFamily) As HRESULT
        <PreserveSig>
        Overloads Function GetWeight() As DWRITE_FONT_WEIGHT
        <PreserveSig>
        Overloads Function GetStretch() As DWRITE_FONT_STRETCH
        <PreserveSig>
        Overloads Function GetStyle() As DWRITE_FONT_STYLE
        <PreserveSig>
        Overloads Function IsSymbolFont() As Boolean
        Overloads Function GetFaceNames(<Out> ByRef names As IDWriteLocalizedStrings) As HRESULT
        Overloads Function GetInformationalStrings(informationalStringID As DWRITE_INFORMATIONAL_STRING_ID, <Out> ByRef informationalStrings As IDWriteLocalizedStrings, <Out> ByRef exists As Boolean) As HRESULT
        <PreserveSig>
        Overloads Function GetSimulations() As DWRITE_FONT_SIMULATIONS
        Overloads Sub GetMetrics(<Out> ByRef fontMetrics As DWRITE_FONT_METRICS)
        Overloads Function HasCharacter(unicodeValue As Integer, <Out> ByRef exists As Boolean) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontFace(<Out> ByRef fontFace As IDWriteFontFace) As HRESULT
#End Region

        Overloads Sub GetMetrics(<Out> ByRef fontMetrics As DWRITE_FONT_METRICS1)
        'void GetPanose(out DWRITE_PANOSE panose);
        Sub GetPanose(<Out> ByRef panose As IntPtr)
        Function GetUnicodeRanges(maxRangeCount As UInteger, <Out> ByRef unicodeRanges As DWRITE_UNICODE_RANGE, <Out> ByRef actualRangeCount As UInteger) As HRESULT
        <PreserveSig>
        Function IsMonospacedFont() As Boolean
    End Interface

    Public Enum DWRITE_PANOSE_FAMILY
        DWRITE_PANOSE_FAMILY_ANY = 0
        DWRITE_PANOSE_FAMILY_NO_FIT = 1
        DWRITE_PANOSE_FAMILY_TEXT_DISPLAY = 2
        DWRITE_PANOSE_FAMILY_SCRIPT = 3 ' or hand written
        DWRITE_PANOSE_FAMILY_DECORATIVE = 4
        DWRITE_PANOSE_FAMILY_SYMBOL = 5 ' or symbol
        DWRITE_PANOSE_FAMILY_PICTORIAL = DWRITE_PANOSE_FAMILY_SYMBOL
    End Enum

    Public Enum DWRITE_PANOSE_SERIF_STYLE
        DWRITE_PANOSE_SERIF_STYLE_ANY = 0
        DWRITE_PANOSE_SERIF_STYLE_NO_FIT = 1
        DWRITE_PANOSE_SERIF_STYLE_COVE = 2
        DWRITE_PANOSE_SERIF_STYLE_OBTUSE_COVE = 3
        DWRITE_PANOSE_SERIF_STYLE_SQUARE_COVE = 4
        DWRITE_PANOSE_SERIF_STYLE_OBTUSE_SQUARE_COVE = 5
        DWRITE_PANOSE_SERIF_STYLE_SQUARE = 6
        DWRITE_PANOSE_SERIF_STYLE_THIN = 7
        DWRITE_PANOSE_SERIF_STYLE_OVAL = 8
        DWRITE_PANOSE_SERIF_STYLE_EXAGGERATED = 9
        DWRITE_PANOSE_SERIF_STYLE_TRIANGLE = 10
        DWRITE_PANOSE_SERIF_STYLE_NORMAL_SANS = 11
        DWRITE_PANOSE_SERIF_STYLE_OBTUSE_SANS = 12
        DWRITE_PANOSE_SERIF_STYLE_PERPENDICULAR_SANS = 13
        DWRITE_PANOSE_SERIF_STYLE_FLARED = 14
        DWRITE_PANOSE_SERIF_STYLE_ROUNDED = 15
        DWRITE_PANOSE_SERIF_STYLE_SCRIPT = 16
        DWRITE_PANOSE_SERIF_STYLE_PERP_SANS = DWRITE_PANOSE_SERIF_STYLE_PERPENDICULAR_SANS
        DWRITE_PANOSE_SERIF_STYLE_BONE = DWRITE_PANOSE_SERIF_STYLE_OVAL
    End Enum

    Public Enum DWRITE_PANOSE_WEIGHT
        DWRITE_PANOSE_WEIGHT_ANY = 0
        DWRITE_PANOSE_WEIGHT_NO_FIT = 1
        DWRITE_PANOSE_WEIGHT_VERY_LIGHT = 2
        DWRITE_PANOSE_WEIGHT_LIGHT = 3
        DWRITE_PANOSE_WEIGHT_THIN = 4
        DWRITE_PANOSE_WEIGHT_BOOK = 5
        DWRITE_PANOSE_WEIGHT_MEDIUM = 6
        DWRITE_PANOSE_WEIGHT_DEMI = 7
        DWRITE_PANOSE_WEIGHT_BOLD = 8
        DWRITE_PANOSE_WEIGHT_HEAVY = 9
        DWRITE_PANOSE_WEIGHT_BLACK = 10
        DWRITE_PANOSE_WEIGHT_EXTRA_BLACK = 11
        DWRITE_PANOSE_WEIGHT_NORD = DWRITE_PANOSE_WEIGHT_EXTRA_BLACK
    End Enum

    Public Enum DWRITE_PANOSE_PROPORTION
        DWRITE_PANOSE_PROPORTION_ANY = 0
        DWRITE_PANOSE_PROPORTION_NO_FIT = 1
        DWRITE_PANOSE_PROPORTION_OLD_STYLE = 2
        DWRITE_PANOSE_PROPORTION_MODERN = 3
        DWRITE_PANOSE_PROPORTION_EVEN_WIDTH = 4
        DWRITE_PANOSE_PROPORTION_EXPANDED = 5
        DWRITE_PANOSE_PROPORTION_CONDENSED = 6
        DWRITE_PANOSE_PROPORTION_VERY_EXPANDED = 7
        DWRITE_PANOSE_PROPORTION_VERY_CONDENSED = 8
        DWRITE_PANOSE_PROPORTION_MONOSPACED = 9
    End Enum

    Public Enum DWRITE_PANOSE_CONTRAST
        DWRITE_PANOSE_CONTRAST_ANY = 0
        DWRITE_PANOSE_CONTRAST_NO_FIT = 1
        DWRITE_PANOSE_CONTRAST_NONE = 2
        DWRITE_PANOSE_CONTRAST_VERY_LOW = 3
        DWRITE_PANOSE_CONTRAST_LOW = 4
        DWRITE_PANOSE_CONTRAST_MEDIUM_LOW = 5
        DWRITE_PANOSE_CONTRAST_MEDIUM = 6
        DWRITE_PANOSE_CONTRAST_MEDIUM_HIGH = 7
        DWRITE_PANOSE_CONTRAST_HIGH = 8
        DWRITE_PANOSE_CONTRAST_VERY_HIGH = 9
        DWRITE_PANOSE_CONTRAST_HORIZONTAL_LOW = 10
        DWRITE_PANOSE_CONTRAST_HORIZONTAL_MEDIUM = 11
        DWRITE_PANOSE_CONTRAST_HORIZONTAL_HIGH = 12
        DWRITE_PANOSE_CONTRAST_BROKEN = 13
    End Enum

    Public Enum DWRITE_PANOSE_STROKE_VARIATION
        DWRITE_PANOSE_STROKE_VARIATION_ANY = 0
        DWRITE_PANOSE_STROKE_VARIATION_NO_FIT = 1
        DWRITE_PANOSE_STROKE_VARIATION_NO_VARIATION = 2
        DWRITE_PANOSE_STROKE_VARIATION_GRADUAL_DIAGONAL = 3
        DWRITE_PANOSE_STROKE_VARIATION_GRADUAL_TRANSITIONAL = 4
        DWRITE_PANOSE_STROKE_VARIATION_GRADUAL_VERTICAL = 5
        DWRITE_PANOSE_STROKE_VARIATION_GRADUAL_HORIZONTAL = 6
        DWRITE_PANOSE_STROKE_VARIATION_RAPID_VERTICAL = 7
        DWRITE_PANOSE_STROKE_VARIATION_RAPID_HORIZONTAL = 8
        DWRITE_PANOSE_STROKE_VARIATION_INSTANT_VERTICAL = 9
        DWRITE_PANOSE_STROKE_VARIATION_INSTANT_HORIZONTAL = 10
    End Enum

    Public Enum DWRITE_PANOSE_ARM_STYLE
        DWRITE_PANOSE_ARM_STYLE_ANY = 0
        DWRITE_PANOSE_ARM_STYLE_NO_FIT = 1
        DWRITE_PANOSE_ARM_STYLE_STRAIGHT_ARMS_HORIZONTAL = 2
        DWRITE_PANOSE_ARM_STYLE_STRAIGHT_ARMS_WEDGE = 3
        DWRITE_PANOSE_ARM_STYLE_STRAIGHT_ARMS_VERTICAL = 4
        DWRITE_PANOSE_ARM_STYLE_STRAIGHT_ARMS_SINGLE_SERIF = 5
        DWRITE_PANOSE_ARM_STYLE_STRAIGHT_ARMS_DOUBLE_SERIF = 6
        DWRITE_PANOSE_ARM_STYLE_NONSTRAIGHT_ARMS_HORIZONTAL = 7
        DWRITE_PANOSE_ARM_STYLE_NONSTRAIGHT_ARMS_WEDGE = 8
        DWRITE_PANOSE_ARM_STYLE_NONSTRAIGHT_ARMS_VERTICAL = 9
        DWRITE_PANOSE_ARM_STYLE_NONSTRAIGHT_ARMS_SINGLE_SERIF = 10
        DWRITE_PANOSE_ARM_STYLE_NONSTRAIGHT_ARMS_DOUBLE_SERIF = 11
        DWRITE_PANOSE_ARM_STYLE_STRAIGHT_ARMS_HORZ = DWRITE_PANOSE_ARM_STYLE_STRAIGHT_ARMS_HORIZONTAL
        DWRITE_PANOSE_ARM_STYLE_STRAIGHT_ARMS_VERT = DWRITE_PANOSE_ARM_STYLE_STRAIGHT_ARMS_VERTICAL
        DWRITE_PANOSE_ARM_STYLE_BENT_ARMS_HORZ = DWRITE_PANOSE_ARM_STYLE_NONSTRAIGHT_ARMS_HORIZONTAL
        DWRITE_PANOSE_ARM_STYLE_BENT_ARMS_WEDGE = DWRITE_PANOSE_ARM_STYLE_NONSTRAIGHT_ARMS_WEDGE
        DWRITE_PANOSE_ARM_STYLE_BENT_ARMS_VERT = DWRITE_PANOSE_ARM_STYLE_NONSTRAIGHT_ARMS_VERTICAL
        DWRITE_PANOSE_ARM_STYLE_BENT_ARMS_SINGLE_SERIF = DWRITE_PANOSE_ARM_STYLE_NONSTRAIGHT_ARMS_SINGLE_SERIF
        DWRITE_PANOSE_ARM_STYLE_BENT_ARMS_DOUBLE_SERIF = DWRITE_PANOSE_ARM_STYLE_NONSTRAIGHT_ARMS_DOUBLE_SERIF
    End Enum

    Public Enum DWRITE_PANOSE_LETTERFORM
        DWRITE_PANOSE_LETTERFORM_ANY = 0
        DWRITE_PANOSE_LETTERFORM_NO_FIT = 1
        DWRITE_PANOSE_LETTERFORM_NORMAL_CONTACT = 2
        DWRITE_PANOSE_LETTERFORM_NORMAL_WEIGHTED = 3
        DWRITE_PANOSE_LETTERFORM_NORMAL_BOXED = 4
        DWRITE_PANOSE_LETTERFORM_NORMAL_FLATTENED = 5
        DWRITE_PANOSE_LETTERFORM_NORMAL_ROUNDED = 6
        DWRITE_PANOSE_LETTERFORM_NORMAL_OFF_CENTER = 7
        DWRITE_PANOSE_LETTERFORM_NORMAL_SQUARE = 8
        DWRITE_PANOSE_LETTERFORM_OBLIQUE_CONTACT = 9
        DWRITE_PANOSE_LETTERFORM_OBLIQUE_WEIGHTED = 10
        DWRITE_PANOSE_LETTERFORM_OBLIQUE_BOXED = 11
        DWRITE_PANOSE_LETTERFORM_OBLIQUE_FLATTENED = 12
        DWRITE_PANOSE_LETTERFORM_OBLIQUE_ROUNDED = 13
        DWRITE_PANOSE_LETTERFORM_OBLIQUE_OFF_CENTER = 14
        DWRITE_PANOSE_LETTERFORM_OBLIQUE_SQUARE = 15
    End Enum

    Public Enum DWRITE_PANOSE_MIDLINE
        DWRITE_PANOSE_MIDLINE_ANY = 0
        DWRITE_PANOSE_MIDLINE_NO_FIT = 1
        DWRITE_PANOSE_MIDLINE_STANDARD_TRIMMED = 2
        DWRITE_PANOSE_MIDLINE_STANDARD_POINTED = 3
        DWRITE_PANOSE_MIDLINE_STANDARD_SERIFED = 4
        DWRITE_PANOSE_MIDLINE_HIGH_TRIMMED = 5
        DWRITE_PANOSE_MIDLINE_HIGH_POINTED = 6
        DWRITE_PANOSE_MIDLINE_HIGH_SERIFED = 7
        DWRITE_PANOSE_MIDLINE_CONSTANT_TRIMMED = 8
        DWRITE_PANOSE_MIDLINE_CONSTANT_POINTED = 9
        DWRITE_PANOSE_MIDLINE_CONSTANT_SERIFED = 10
        DWRITE_PANOSE_MIDLINE_LOW_TRIMMED = 11
        DWRITE_PANOSE_MIDLINE_LOW_POINTED = 12
        DWRITE_PANOSE_MIDLINE_LOW_SERIFED = 13
    End Enum

    Public Enum DWRITE_PANOSE_XHEIGHT
        DWRITE_PANOSE_XHEIGHT_ANY = 0
        DWRITE_PANOSE_XHEIGHT_NO_FIT = 1
        DWRITE_PANOSE_XHEIGHT_CONSTANT_SMALL = 2
        DWRITE_PANOSE_XHEIGHT_CONSTANT_STANDARD = 3
        DWRITE_PANOSE_XHEIGHT_CONSTANT_LARGE = 4
        DWRITE_PANOSE_XHEIGHT_DUCKING_SMALL = 5
        DWRITE_PANOSE_XHEIGHT_DUCKING_STANDARD = 6
        DWRITE_PANOSE_XHEIGHT_DUCKING_LARGE = 7
        DWRITE_PANOSE_XHEIGHT_CONSTANT_STD = DWRITE_PANOSE_XHEIGHT_CONSTANT_STANDARD
        DWRITE_PANOSE_XHEIGHT_DUCKING_STD = DWRITE_PANOSE_XHEIGHT_DUCKING_STANDARD
    End Enum

    Public Enum DWRITE_PANOSE_TOOL_KIND
        DWRITE_PANOSE_TOOL_KIND_ANY = 0
        DWRITE_PANOSE_TOOL_KIND_NO_FIT = 1
        DWRITE_PANOSE_TOOL_KIND_FLAT_NIB = 2
        DWRITE_PANOSE_TOOL_KIND_PRESSURE_POINT = 3
        DWRITE_PANOSE_TOOL_KIND_ENGRAVED = 4
        DWRITE_PANOSE_TOOL_KIND_BALL = 5
        DWRITE_PANOSE_TOOL_KIND_BRUSH = 6
        DWRITE_PANOSE_TOOL_KIND_ROUGH = 7
        DWRITE_PANOSE_TOOL_KIND_FELT_PEN_BRUSH_TIP = 8
        DWRITE_PANOSE_TOOL_KIND_WILD_BRUSH = 9
    End Enum

    Public Enum DWRITE_PANOSE_SPACING
        DWRITE_PANOSE_SPACING_ANY = 0
        DWRITE_PANOSE_SPACING_NO_FIT = 1
        DWRITE_PANOSE_SPACING_PROPORTIONAL_SPACED = 2
        DWRITE_PANOSE_SPACING_MONOSPACED = 3
    End Enum

    Public Enum DWRITE_PANOSE_ASPECT_RATIO
        DWRITE_PANOSE_ASPECT_RATIO_ANY = 0
        DWRITE_PANOSE_ASPECT_RATIO_NO_FIT = 1
        DWRITE_PANOSE_ASPECT_RATIO_VERY_CONDENSED = 2
        DWRITE_PANOSE_ASPECT_RATIO_CONDENSED = 3
        DWRITE_PANOSE_ASPECT_RATIO_NORMAL = 4
        DWRITE_PANOSE_ASPECT_RATIO_EXPANDED = 5
        DWRITE_PANOSE_ASPECT_RATIO_VERY_EXPANDED = 6
    End Enum

    Public Enum DWRITE_PANOSE_SCRIPT_TOPOLOGY
        DWRITE_PANOSE_SCRIPT_TOPOLOGY_ANY = 0
        DWRITE_PANOSE_SCRIPT_TOPOLOGY_NO_FIT = 1
        DWRITE_PANOSE_SCRIPT_TOPOLOGY_ROMAN_DISCONNECTED = 2
        DWRITE_PANOSE_SCRIPT_TOPOLOGY_ROMAN_TRAILING = 3
        DWRITE_PANOSE_SCRIPT_TOPOLOGY_ROMAN_CONNECTED = 4
        DWRITE_PANOSE_SCRIPT_TOPOLOGY_CURSIVE_DISCONNECTED = 5
        DWRITE_PANOSE_SCRIPT_TOPOLOGY_CURSIVE_TRAILING = 6
        DWRITE_PANOSE_SCRIPT_TOPOLOGY_CURSIVE_CONNECTED = 7
        DWRITE_PANOSE_SCRIPT_TOPOLOGY_BLACKLETTER_DISCONNECTED = 8
        DWRITE_PANOSE_SCRIPT_TOPOLOGY_BLACKLETTER_TRAILING = 9
        DWRITE_PANOSE_SCRIPT_TOPOLOGY_BLACKLETTER_CONNECTED = 10
    End Enum

    Public Enum DWRITE_PANOSE_SCRIPT_FORM
        DWRITE_PANOSE_SCRIPT_FORM_ANY = 0
        DWRITE_PANOSE_SCRIPT_FORM_NO_FIT = 1
        DWRITE_PANOSE_SCRIPT_FORM_UPRIGHT_NO_WRAPPING = 2
        DWRITE_PANOSE_SCRIPT_FORM_UPRIGHT_SOME_WRAPPING = 3
        DWRITE_PANOSE_SCRIPT_FORM_UPRIGHT_MORE_WRAPPING = 4
        DWRITE_PANOSE_SCRIPT_FORM_UPRIGHT_EXTREME_WRAPPING = 5
        DWRITE_PANOSE_SCRIPT_FORM_OBLIQUE_NO_WRAPPING = 6
        DWRITE_PANOSE_SCRIPT_FORM_OBLIQUE_SOME_WRAPPING = 7
        DWRITE_PANOSE_SCRIPT_FORM_OBLIQUE_MORE_WRAPPING = 8
        DWRITE_PANOSE_SCRIPT_FORM_OBLIQUE_EXTREME_WRAPPING = 9
        DWRITE_PANOSE_SCRIPT_FORM_EXAGGERATED_NO_WRAPPING = 10
        DWRITE_PANOSE_SCRIPT_FORM_EXAGGERATED_SOME_WRAPPING = 11
        DWRITE_PANOSE_SCRIPT_FORM_EXAGGERATED_MORE_WRAPPING = 12
        DWRITE_PANOSE_SCRIPT_FORM_EXAGGERATED_EXTREME_WRAPPING = 13
    End Enum

    Public Enum DWRITE_PANOSE_FINIALS
        DWRITE_PANOSE_FINIALS_ANY = 0
        DWRITE_PANOSE_FINIALS_NO_FIT = 1
        DWRITE_PANOSE_FINIALS_NONE_NO_LOOPS = 2
        DWRITE_PANOSE_FINIALS_NONE_CLOSED_LOOPS = 3
        DWRITE_PANOSE_FINIALS_NONE_OPEN_LOOPS = 4
        DWRITE_PANOSE_FINIALS_SHARP_NO_LOOPS = 5
        DWRITE_PANOSE_FINIALS_SHARP_CLOSED_LOOPS = 6
        DWRITE_PANOSE_FINIALS_SHARP_OPEN_LOOPS = 7
        DWRITE_PANOSE_FINIALS_TAPERED_NO_LOOPS = 8
        DWRITE_PANOSE_FINIALS_TAPERED_CLOSED_LOOPS = 9
        DWRITE_PANOSE_FINIALS_TAPERED_OPEN_LOOPS = 10
        DWRITE_PANOSE_FINIALS_ROUND_NO_LOOPS = 11
        DWRITE_PANOSE_FINIALS_ROUND_CLOSED_LOOPS = 12
        DWRITE_PANOSE_FINIALS_ROUND_OPEN_LOOPS = 13
    End Enum

    Public Enum DWRITE_PANOSE_XASCENT
        DWRITE_PANOSE_XASCENT_ANY = 0
        DWRITE_PANOSE_XASCENT_NO_FIT = 1
        DWRITE_PANOSE_XASCENT_VERY_LOW = 2
        DWRITE_PANOSE_XASCENT_LOW = 3
        DWRITE_PANOSE_XASCENT_MEDIUM = 4
        DWRITE_PANOSE_XASCENT_HIGH = 5
        DWRITE_PANOSE_XASCENT_VERY_HIGH = 6
    End Enum

    Public Enum DWRITE_PANOSE_DECORATIVE_CLASS
        DWRITE_PANOSE_DECORATIVE_CLASS_ANY = 0
        DWRITE_PANOSE_DECORATIVE_CLASS_NO_FIT = 1
        DWRITE_PANOSE_DECORATIVE_CLASS_DERIVATIVE = 2
        DWRITE_PANOSE_DECORATIVE_CLASS_NONSTANDARD_TOPOLOGY = 3
        DWRITE_PANOSE_DECORATIVE_CLASS_NONSTANDARD_ELEMENTS = 4
        DWRITE_PANOSE_DECORATIVE_CLASS_NONSTANDARD_ASPECT = 5
        DWRITE_PANOSE_DECORATIVE_CLASS_INITIALS = 6
        DWRITE_PANOSE_DECORATIVE_CLASS_CARTOON = 7
        DWRITE_PANOSE_DECORATIVE_CLASS_PICTURE_STEMS = 8
        DWRITE_PANOSE_DECORATIVE_CLASS_ORNAMENTED = 9
        DWRITE_PANOSE_DECORATIVE_CLASS_TEXT_AND_BACKGROUND = 10
        DWRITE_PANOSE_DECORATIVE_CLASS_COLLAGE = 11
        DWRITE_PANOSE_DECORATIVE_CLASS_MONTAGE = 12
    End Enum

    Public Enum DWRITE_PANOSE_ASPECT
        DWRITE_PANOSE_ASPECT_ANY = 0
        DWRITE_PANOSE_ASPECT_NO_FIT = 1
        DWRITE_PANOSE_ASPECT_SUPER_CONDENSED = 2
        DWRITE_PANOSE_ASPECT_VERY_CONDENSED = 3
        DWRITE_PANOSE_ASPECT_CONDENSED = 4
        DWRITE_PANOSE_ASPECT_NORMAL = 5
        DWRITE_PANOSE_ASPECT_EXTENDED = 6
        DWRITE_PANOSE_ASPECT_VERY_EXTENDED = 7
        DWRITE_PANOSE_ASPECT_SUPER_EXTENDED = 8
        DWRITE_PANOSE_ASPECT_MONOSPACED = 9
    End Enum

    Public Enum DWRITE_PANOSE_FILL
        DWRITE_PANOSE_FILL_ANY = 0
        DWRITE_PANOSE_FILL_NO_FIT = 1
        DWRITE_PANOSE_FILL_STANDARD_SOLID_FILL = 2
        DWRITE_PANOSE_FILL_NO_FILL = 3
        DWRITE_PANOSE_FILL_PATTERNED_FILL = 4
        DWRITE_PANOSE_FILL_COMPLEX_FILL = 5
        DWRITE_PANOSE_FILL_SHAPED_FILL = 6
        DWRITE_PANOSE_FILL_DRAWN_DISTRESSED = 7
    End Enum

    Public Enum DWRITE_PANOSE_LINING
        DWRITE_PANOSE_LINING_ANY = 0
        DWRITE_PANOSE_LINING_NO_FIT = 1
        DWRITE_PANOSE_LINING_NONE = 2
        DWRITE_PANOSE_LINING_INLINE = 3
        DWRITE_PANOSE_LINING_OUTLINE = 4
        DWRITE_PANOSE_LINING_ENGRAVED = 5
        DWRITE_PANOSE_LINING_SHADOW = 6
        DWRITE_PANOSE_LINING_RELIEF = 7
        DWRITE_PANOSE_LINING_BACKDROP = 8
    End Enum

    Public Enum DWRITE_PANOSE_DECORATIVE_TOPOLOGY
        DWRITE_PANOSE_DECORATIVE_TOPOLOGY_ANY = 0
        DWRITE_PANOSE_DECORATIVE_TOPOLOGY_NO_FIT = 1
        DWRITE_PANOSE_DECORATIVE_TOPOLOGY_STANDARD = 2
        DWRITE_PANOSE_DECORATIVE_TOPOLOGY_SQUARE = 3
        DWRITE_PANOSE_DECORATIVE_TOPOLOGY_MULTIPLE_SEGMENT = 4
        DWRITE_PANOSE_DECORATIVE_TOPOLOGY_ART_DECO = 5
        DWRITE_PANOSE_DECORATIVE_TOPOLOGY_UNEVEN_WEIGHTING = 6
        DWRITE_PANOSE_DECORATIVE_TOPOLOGY_DIVERSE_ARMS = 7
        DWRITE_PANOSE_DECORATIVE_TOPOLOGY_DIVERSE_FORMS = 8
        DWRITE_PANOSE_DECORATIVE_TOPOLOGY_LOMBARDIC_FORMS = 9
        DWRITE_PANOSE_DECORATIVE_TOPOLOGY_UPPER_CASE_IN_LOWER_CASE = 10
        DWRITE_PANOSE_DECORATIVE_TOPOLOGY_IMPLIED_TOPOLOGY = 11
        DWRITE_PANOSE_DECORATIVE_TOPOLOGY_HORSESHOE_E_AND_A = 12
        DWRITE_PANOSE_DECORATIVE_TOPOLOGY_CURSIVE = 13
        DWRITE_PANOSE_DECORATIVE_TOPOLOGY_BLACKLETTER = 14
        DWRITE_PANOSE_DECORATIVE_TOPOLOGY_SWASH_VARIANCE = 15
    End Enum

    Public Enum DWRITE_PANOSE_CHARACTER_RANGES
        DWRITE_PANOSE_CHARACTER_RANGES_ANY = 0
        DWRITE_PANOSE_CHARACTER_RANGES_NO_FIT = 1
        DWRITE_PANOSE_CHARACTER_RANGES_EXTENDED_COLLECTION = 2
        DWRITE_PANOSE_CHARACTER_RANGES_LITERALS = 3
        DWRITE_PANOSE_CHARACTER_RANGES_NO_LOWER_CASE = 4
        DWRITE_PANOSE_CHARACTER_RANGES_SMALL_CAPS = 5
    End Enum

    Public Enum DWRITE_PANOSE_SYMBOL_KIND
        DWRITE_PANOSE_SYMBOL_KIND_ANY = 0
        DWRITE_PANOSE_SYMBOL_KIND_NO_FIT = 1
        DWRITE_PANOSE_SYMBOL_KIND_MONTAGES = 2
        DWRITE_PANOSE_SYMBOL_KIND_PICTURES = 3
        DWRITE_PANOSE_SYMBOL_KIND_SHAPES = 4
        DWRITE_PANOSE_SYMBOL_KIND_SCIENTIFIC = 5
        DWRITE_PANOSE_SYMBOL_KIND_MUSIC = 6
        DWRITE_PANOSE_SYMBOL_KIND_EXPERT = 7
        DWRITE_PANOSE_SYMBOL_KIND_PATTERNS = 8
        DWRITE_PANOSE_SYMBOL_KIND_BOARDERS = 9
        DWRITE_PANOSE_SYMBOL_KIND_ICONS = 10
        DWRITE_PANOSE_SYMBOL_KIND_LOGOS = 11
        DWRITE_PANOSE_SYMBOL_KIND_INDUSTRY_SPECIFIC = 12
    End Enum

    Public Enum DWRITE_PANOSE_SYMBOL_ASPECT_RATIO
        DWRITE_PANOSE_SYMBOL_ASPECT_RATIO_ANY = 0
        DWRITE_PANOSE_SYMBOL_ASPECT_RATIO_NO_FIT = 1
        DWRITE_PANOSE_SYMBOL_ASPECT_RATIO_NO_WIDTH = 2
        DWRITE_PANOSE_SYMBOL_ASPECT_RATIO_EXCEPTIONALLY_WIDE = 3
        DWRITE_PANOSE_SYMBOL_ASPECT_RATIO_SUPER_WIDE = 4
        DWRITE_PANOSE_SYMBOL_ASPECT_RATIO_VERY_WIDE = 5
        DWRITE_PANOSE_SYMBOL_ASPECT_RATIO_WIDE = 6
        DWRITE_PANOSE_SYMBOL_ASPECT_RATIO_NORMAL = 7
        DWRITE_PANOSE_SYMBOL_ASPECT_RATIO_NARROW = 8
        DWRITE_PANOSE_SYMBOL_ASPECT_RATIO_VERY_NARROW = 9
    End Enum

    <ComImport>
    <Guid("29748ed6-8c9c-4a6a-be0b-d912e8538944")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFont2
        Inherits IDWriteFont1
#Region "IDWriteFont1"
#Region "IDWriteFont"
        Overloads Function GetFontFamily(<Out> ByRef fontFamily As IDWriteFontFamily) As HRESULT
        <PreserveSig>
        Overloads Function GetWeight() As DWRITE_FONT_WEIGHT
        <PreserveSig>
        Overloads Function GetStretch() As DWRITE_FONT_STRETCH
        <PreserveSig>
        Overloads Function GetStyle() As DWRITE_FONT_STYLE
        <PreserveSig>
        Overloads Function IsSymbolFont() As Boolean
        Overloads Function GetFaceNames(<Out> ByRef names As IDWriteLocalizedStrings) As HRESULT
        Overloads Function GetInformationalStrings(informationalStringID As DWRITE_INFORMATIONAL_STRING_ID, <Out> ByRef informationalStrings As IDWriteLocalizedStrings, <Out> ByRef exists As Boolean) As HRESULT
        Overloads Function GetSimulations() As DWRITE_FONT_SIMULATIONS
        Overloads Sub GetMetrics(<Out> ByRef fontMetrics As DWRITE_FONT_METRICS)
        Overloads Function HasCharacter(unicodeValue As Integer, <Out> ByRef exists As Boolean) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontFace(<Out> ByRef fontFace As IDWriteFontFace) As HRESULT
#End Region

        Overloads Sub GetMetrics(<Out> ByRef fontMetrics As DWRITE_FONT_METRICS1)
        'new void GetPanose(out DWRITE_PANOSE panose);
        Overloads Sub GetPanose(<Out> ByRef panose As IntPtr)
        Overloads Function GetUnicodeRanges(maxRangeCount As UInteger, <Out> ByRef unicodeRanges As DWRITE_UNICODE_RANGE, <Out> ByRef actualRangeCount As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function IsMonospacedFont() As Boolean
#End Region

        <PreserveSig>
        Function IsColorFont() As Boolean
    End Interface

    <ComImport>
    <Guid("29748ED6-8C9C-4A6A-BE0B-D912E8538944")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFont3
        Inherits IDWriteFont2
#Region "IDWriteFont2"
#Region "IDWriteFont1"
#Region "IDWriteFont"
        Overloads Function GetFontFamily(<Out> ByRef fontFamily As IDWriteFontFamily) As HRESULT
        <PreserveSig>
        Overloads Function GetWeight() As DWRITE_FONT_WEIGHT
        <PreserveSig>
        Overloads Function GetStretch() As DWRITE_FONT_STRETCH
        <PreserveSig>
        Overloads Function GetStyle() As DWRITE_FONT_STYLE
        <PreserveSig>
        Overloads Function IsSymbolFont() As Boolean
        Overloads Function GetFaceNames(<Out> ByRef names As IDWriteLocalizedStrings) As HRESULT
        Overloads Function GetInformationalStrings(informationalStringID As DWRITE_INFORMATIONAL_STRING_ID, <Out> ByRef informationalStrings As IDWriteLocalizedStrings, <Out> ByRef exists As Boolean) As HRESULT
        Overloads Function GetSimulations() As DWRITE_FONT_SIMULATIONS
        Overloads Sub GetMetrics(<Out> ByRef fontMetrics As DWRITE_FONT_METRICS)
        Overloads Function HasCharacter(unicodeValue As Integer, <Out> ByRef exists As Boolean) As HRESULT
        <PreserveSig>
        Overloads Function CreateFontFace(<Out> ByRef fontFace As IDWriteFontFace) As HRESULT
#End Region

        Overloads Sub GetMetrics(<Out> ByRef fontMetrics As DWRITE_FONT_METRICS1)
        'new void GetPanose(out DWRITE_PANOSE panose);
        Overloads Sub GetPanose(<Out> ByRef panose As IntPtr)
        Overloads Function GetUnicodeRanges(maxRangeCount As UInteger, <Out> ByRef unicodeRanges As DWRITE_UNICODE_RANGE, <Out> ByRef actualRangeCount As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function IsMonospacedFont() As Boolean
#End Region

        <PreserveSig>
        Overloads Function IsColorFont() As Boolean
#End Region

        <PreserveSig>
        Overloads Function CreateFontFace(<Out> ByRef fontFace As IDWriteFontFace3) As HRESULT
        <PreserveSig>
        Function Equals(font As IDWriteFont) As Boolean
        Function GetFontFaceReference(<Out> ByRef fontFaceReference As IDWriteFontFaceReference) As HRESULT
        <PreserveSig>
        Overloads Function HasCharacter(unicodeValue As UInteger) As Boolean
        <PreserveSig>
        Function GetLocality() As DWRITE_LOCALITY
    End Interface

    <ComImport>
    <Guid("08256209-099a-4b34-b86d-c22b110e7771")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteLocalizedStrings
        <PreserveSig>
        Function GetCount() As Integer
        Function FindLocaleName(localeName As String, <Out> ByRef index As UInteger, <Out> ByRef exists As Boolean) As HRESULT
        Function GetLocaleNameLength(index As UInteger, <Out> ByRef length As UInteger) As HRESULT
        Function GetLocaleName(index As UInteger, <Out> ByRef localeName As String, size As UInteger) As HRESULT
        Function GetStringLength(index As UInteger, <Out> ByRef length As UInteger) As HRESULT
        Function GetString(index As UInteger, stringBuffer As Text.StringBuilder, size As UInteger) As HRESULT
    End Interface

    <ComImport>
    <Guid("da20d8ef-812a-4c43-9802-62ec4abd7add")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontFamily
        Inherits IDWriteFontList
#Region "IDWriteFontList"
        Overloads Function GetFontCollection(<Out> ByRef fontCollection As IDWriteFontCollection) As HRESULT
        <PreserveSig>
        Overloads Function GetFontCount() As Integer
        Overloads Function GetFont(index As Integer, <Out> ByRef font As IDWriteFont) As HRESULT
#End Region

        Function GetFamilyNames(<Out> ByRef names As IDWriteLocalizedStrings) As HRESULT
        Function GetFirstMatchingFont(weight As DWRITE_FONT_WEIGHT, stretch As DWRITE_FONT_STRETCH, style As DWRITE_FONT_STYLE, <Out> ByRef matchingFont As IDWriteFont) As HRESULT
        Function GetMatchingFonts(weight As DWRITE_FONT_WEIGHT, stretch As DWRITE_FONT_STRETCH, style As DWRITE_FONT_STYLE, <Out> ByRef matchingFonts As IDWriteFontList) As HRESULT
    End Interface

    <ComImport>
    <Guid("DA20D8EF-812A-4C43-9802-62EC4ABD7ADF")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontFamily1
        Inherits IDWriteFontFamily
#Region "IDWriteFontFamily"
#Region "IDWriteFontList"
        Overloads Function GetFontCollection(<Out> ByRef fontCollection As IDWriteFontCollection) As HRESULT
        <PreserveSig>
        Overloads Function GetFontCount() As Integer
        Overloads Function GetFont(index As Integer, <Out> ByRef font As IDWriteFont) As HRESULT
#End Region

        Overloads Function GetFamilyNames(<Out> ByRef names As IDWriteLocalizedStrings) As HRESULT
        Overloads Function GetFirstMatchingFont(weight As DWRITE_FONT_WEIGHT, stretch As DWRITE_FONT_STRETCH, style As DWRITE_FONT_STYLE, <Out> ByRef matchingFont As IDWriteFont) As HRESULT
        Overloads Function GetMatchingFonts(weight As DWRITE_FONT_WEIGHT, stretch As DWRITE_FONT_STRETCH, style As DWRITE_FONT_STYLE, <Out> ByRef matchingFonts As IDWriteFontList) As HRESULT
#End Region

        Function GetFontLocality(listIndex As UInteger) As DWRITE_LOCALITY
        Overloads Function GetFont(listIndex As UInteger, <Out> ByRef font As IDWriteFont3) As HRESULT
        Function GetFontFaceReference(listIndex As UInteger, <Out> ByRef fontFaceReference As IDWriteFontFaceReference) As HRESULT
    End Interface

    <ComImport>
    <Guid("3ED49E77-A398-4261-B9CF-C126C2131EF3")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontFamily2
        Inherits IDWriteFontFamily1
#Region "IDWriteFontFamily1"
#Region "IDWriteFontFamily"
#Region "IDWriteFontList"
        Overloads Function GetFontCollection(<Out> ByRef fontCollection As IDWriteFontCollection) As HRESULT
        <PreserveSig>
        Overloads Function GetFontCount() As Integer
        Overloads Function GetFont(index As Integer, <Out> ByRef font As IDWriteFont) As HRESULT
#End Region

        Overloads Function GetFamilyNames(<Out> ByRef names As IDWriteLocalizedStrings) As HRESULT
        Overloads Function GetFirstMatchingFont(weight As DWRITE_FONT_WEIGHT, stretch As DWRITE_FONT_STRETCH, style As DWRITE_FONT_STYLE, <Out> ByRef matchingFont As IDWriteFont) As HRESULT
        Overloads Function GetMatchingFonts(weight As DWRITE_FONT_WEIGHT, stretch As DWRITE_FONT_STRETCH, style As DWRITE_FONT_STYLE, <Out> ByRef matchingFonts As IDWriteFontList) As HRESULT
#End Region

        Overloads Function GetFontLocality(listIndex As UInteger) As DWRITE_LOCALITY
        Overloads Function GetFont(listIndex As UInteger, <Out> ByRef font As IDWriteFont3) As HRESULT
        Overloads Function GetFontFaceReference(listIndex As UInteger, <Out> ByRef fontFaceReference As IDWriteFontFaceReference) As HRESULT
#End Region

        Overloads Function GetMatchingFonts(fontAxisValues As DWRITE_FONT_AXIS_VALUE, fontAxisValueCount As UInteger, <Out> ByRef matchingFonts As IDWriteFontList2) As HRESULT
        Function GetFontSet(<Out> ByRef fontSet As IDWriteFontSet1) As HRESULT
    End Interface

    <ComImport>
    <Guid("1a0d8438-1d97-4ec1-aef9-a2fb86ed6acb")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontList
        Function GetFontCollection(<Out> ByRef fontCollection As IDWriteFontCollection) As HRESULT
        <PreserveSig>
        Function GetFontCount() As Integer
        Function GetFont(index As Integer, <Out> ByRef font As IDWriteFont) As HRESULT
    End Interface

    <ComImport>
    <Guid("DA20D8EF-812A-4C43-9802-62EC4ABD7ADE")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontList1
        Inherits IDWriteFontList
#Region "IDWriteFontList"
        Overloads Function GetFontCollection(<Out> ByRef fontCollection As IDWriteFontCollection) As HRESULT
        <PreserveSig>
        Overloads Function GetFontCount() As Integer
        Overloads Function GetFont(index As Integer, <Out> ByRef font As IDWriteFont) As HRESULT
#End Region

        <PreserveSig>
        Function GetFontLocality(listIndex As UInteger) As DWRITE_LOCALITY
        Overloads Function GetFont(listIndex As UInteger, <Out> ByRef font As IDWriteFont3) As HRESULT
        Function GetFontFaceReference(listIndex As UInteger, <Out> ByRef fontFaceReference As IDWriteFontFaceReference) As HRESULT
    End Interface

    <ComImport>
    <Guid("C0763A34-77AF-445A-B735-08C37B0A5BF5")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontList2
        Inherits IDWriteFontList1
#Region "IDWriteFontList1"
#Region "IDWriteFontList"
        Overloads Function GetFontCollection(<Out> ByRef fontCollection As IDWriteFontCollection) As HRESULT
        <PreserveSig>
        Overloads Function GetFontCount() As Integer
        Overloads Function GetFont(index As Integer, <Out> ByRef font As IDWriteFont) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function GetFontLocality(listIndex As UInteger) As DWRITE_LOCALITY
        Overloads Function GetFont(listIndex As UInteger, <Out> ByRef font As IDWriteFont3) As HRESULT
        Overloads Function GetFontFaceReference(listIndex As UInteger, <Out> ByRef fontFaceReference As IDWriteFontFaceReference) As HRESULT
#End Region

        Function GetFontSet(<Out> ByRef fontSet As IDWriteFontSet1) As HRESULT
    End Interface

    <ComImport>
    <Guid("688e1a58-5094-47c8-adc8-fbcea60ae92b")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteTextAnalysisSource
        Function GetTextAtPosition(textPosition As Integer, <Out> ByRef textString As String, <Out> ByRef textLength As Integer) As HRESULT
        Function GetTextBeforePosition(textPosition As Integer, <Out> ByRef textString As String, <Out> ByRef textLength As Integer) As HRESULT
        <PreserveSig>
        Function GetParagraphReadingDirection() As DWRITE_READING_DIRECTION
        Function GetLocaleName(textPosition As Integer, <Out> ByRef textLength As Integer, <Out> ByRef localeName As String) As HRESULT
        Function GetNumberSubstitution(textPosition As Integer, <Out> ByRef textLength As Integer, <Out> ByRef numberSubstitution As IDWriteNumberSubstitution) As HRESULT
    End Interface

    <ComImport>
    <Guid("639CFAD8-0FB4-4B21-A58A-067920120009")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteTextAnalysisSource1
        Inherits IDWriteTextAnalysisSource
#Region "IDWriteTextAnalysisSource"
        Overloads Function GetTextAtPosition(textPosition As Integer, <Out> ByRef textString As String, <Out> ByRef textLength As Integer) As HRESULT
        Overloads Function GetTextBeforePosition(textPosition As Integer, <Out> ByRef textString As String, <Out> ByRef textLength As Integer) As HRESULT
        <PreserveSig>
        Overloads Function GetParagraphReadingDirection() As DWRITE_READING_DIRECTION
        Overloads Function GetLocaleName(textPosition As Integer, <Out> ByRef textLength As Integer, <Out> ByRef localeName As String) As HRESULT
        Overloads Function GetNumberSubstitution(textPosition As Integer, <Out> ByRef textLength As Integer, <Out> ByRef numberSubstitution As IDWriteNumberSubstitution) As HRESULT
#End Region

        Function GetVerticalGlyphOrientation(textPosition As UInteger, <Out> ByRef textLength As UInteger, <Out> ByRef glyphOrientation As DWRITE_VERTICAL_GLYPH_ORIENTATION, <Out> ByRef bidiLevel As Byte) As HRESULT
    End Interface

    <ComImport>
    <Guid("14885CC9-BAB0-4f90-B6ED-5C366A2CD03D")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteNumberSubstitution

    End Interface

    <ComImport>
    <Guid("5f49804d-7024-4d43-bfa9-d25984f53849")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontFace
        <PreserveSig>
        Function [GetType]() As DWRITE_FONT_FACE_TYPE
        Function GetFiles(
<[In], Out> ByRef numberOfFiles As UInteger,
<Out, MarshalAs(UnmanagedType.LPArray)> fontFiles As IDWriteFontFile()) As HRESULT
        <PreserveSig>
        Function GetIndex() As Integer
        <PreserveSig>
        Function GetSimulations() As DWRITE_FONT_SIMULATIONS
        <PreserveSig>
        Function IsSymbolFont() As Boolean
        Sub GetMetrics(<Out> ByRef fontFaceMetrics As DWRITE_FONT_METRICS)
        <PreserveSig>
        Function GetGlyphCount() As UShort
        Function GetDesignGlyphMetrics(
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> glyphIndices As UShort(), glyphCount As Integer,
<Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> glyphMetrics As DWRITE_GLYPH_METRICS(), Optional isSideways As Boolean = False) As HRESULT
        <PreserveSig>
        Function GetGlyphIndices(
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> codePoints As UInteger(), codePointCount As Integer,
        <Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> glyphIndices As UShort()) As HRESULT
        Function TryGetFontTable(openTypeTableTag As Integer, <Out> ByRef tableData As IntPtr, <Out> ByRef tableSize As Integer, <Out> ByRef tableContext As IntPtr, <Out> ByRef exists As Boolean) As HRESULT
        Sub ReleaseFontTable(tableContext As IntPtr)
        Function GetGlyphRunOutline(emSize As Single,
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=4)> glyphIndices As UShort(),
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=4)> glyphAdvances As Single(),
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=4)> glyphOffsets As DWRITE_GLYPH_OFFSET(), glyphCount As Integer, isSideways As Boolean, isRightToLeft As Boolean, geometrySink As ID2D1SimplifiedGeometrySink) As HRESULT
        Function GetRecommendedRenderingMode(emSize As Single, pixelsPerDip As Single, measuringMode As DWRITE_MEASURING_MODE, renderingParams As IDWriteRenderingParams, <Out> ByRef renderingMode As DWRITE_RENDERING_MODE) As HRESULT
        Function GetGdiCompatibleMetrics(emSize As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, <Out> ByRef fontFaceMetrics As DWRITE_FONT_METRICS) As HRESULT
        Function GetGdiCompatibleGlyphMetrics(emSize As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, useGdiNatural As Boolean, glyphIndices As UShort, glyphCount As Integer, <Out> ByRef glyphMetrics As DWRITE_GLYPH_METRICS, Optional isSideways As Boolean = False) As HRESULT
    End Interface

    <ComImport>
    <Guid("a71efdb4-9fdb-4838-ad90-cfc3be8c3daf")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontFace1
        Inherits IDWriteFontFace
#Region "IDWriteFontFace"
        <PreserveSig>
        Overloads Function [GetType]() As DWRITE_FONT_FACE_TYPE
        Overloads Function GetFiles(
<[In], Out> ByRef numberOfFiles As UInteger,
<Out, MarshalAs(UnmanagedType.LPArray)> fontFiles As IDWriteFontFile()) As HRESULT
        <PreserveSig>
        Overloads Function GetIndex() As Integer
        <PreserveSig>
        Overloads Function GetSimulations() As DWRITE_FONT_SIMULATIONS
        <PreserveSig>
        Overloads Function IsSymbolFont() As Boolean
        Overloads Sub GetMetrics(<Out> ByRef fontFaceMetrics As DWRITE_FONT_METRICS)
        <PreserveSig>
        Overloads Function GetGlyphCount() As UShort
        Overloads Function GetDesignGlyphMetrics(
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> glyphIndices As UShort(), glyphCount As Integer,
<Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> glyphMetrics As DWRITE_GLYPH_METRICS(), Optional isSideways As Boolean = False) As HRESULT
        <PreserveSig>
        Overloads Function GetGlyphIndices(
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> codePoints As UInteger(), codePointCount As Integer,
        <Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> glyphIndices As UShort()) As HRESULT
        Overloads Function TryGetFontTable(openTypeTableTag As Integer, <Out> ByRef tableData As IntPtr, <Out> ByRef tableSize As Integer, <Out> ByRef tableContext As IntPtr, <Out> ByRef exists As Boolean) As HRESULT
        Overloads Sub ReleaseFontTable(tableContext As IntPtr)
        Overloads Function GetGlyphRunOutline(emSize As Single,
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=4)> glyphIndices As UShort(),
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=4)> glyphAdvances As Single(),
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=4)> glyphOffsets As DWRITE_GLYPH_OFFSET(), glyphCount As Integer, isSideways As Boolean, isRightToLeft As Boolean, geometrySink As ID2D1SimplifiedGeometrySink) As HRESULT
        Overloads Function GetRecommendedRenderingMode(emSize As Single, pixelsPerDip As Single, measuringMode As DWRITE_MEASURING_MODE, renderingParams As IDWriteRenderingParams, <Out> ByRef renderingMode As DWRITE_RENDERING_MODE) As HRESULT
        Overloads Function GetGdiCompatibleMetrics(emSize As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, <Out> ByRef fontFaceMetrics As DWRITE_FONT_METRICS) As HRESULT
        Overloads Function GetGdiCompatibleGlyphMetrics(emSize As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, useGdiNatural As Boolean, glyphIndices As UShort, glyphCount As Integer, <Out> ByRef glyphMetrics As DWRITE_GLYPH_METRICS, Optional isSideways As Boolean = False) As HRESULT
#End Region

        Overloads Sub GetMetrics(<Out> ByRef fontMetrics As DWRITE_FONT_METRICS1)
        Overloads Function GetGdiCompatibleMetrics(emSize As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, <Out> ByRef fontMetrics As DWRITE_FONT_METRICS1) As HRESULT
        Sub GetCaretMetrics(<Out> ByRef caretMetrics As DWRITE_CARET_METRICS)
        Function GetUnicodeRanges(maxRangeCount As UInteger, <Out> ByRef unicodeRanges As DWRITE_UNICODE_RANGE, <Out> ByRef actualRangeCount As UInteger) As HRESULT
        <PreserveSig>
        Function IsMonospacedFont() As Boolean
        Function GetDesignGlyphAdvances(glyphCount As UInteger, glyphIndices As UShort, <Out> ByRef glyphAdvances As Integer, Optional isSideways As Boolean = False) As HRESULT
        Function GetGdiCompatibleGlyphAdvances(emSize As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, useGdiNatural As Boolean, isSideways As Boolean, glyphCount As UInteger, glyphIndices As UShort, <Out> ByRef glyphAdvances As Integer) As HRESULT
        Function GetKerningPairAdjustments(glyphCount As UInteger, glyphIndices As UShort, <Out> ByRef glyphAdvanceAdjustments As Integer) As HRESULT
        <PreserveSig>
        Function HasKerningPairs() As Boolean
        Overloads Function GetRecommendedRenderingMode(fontEmSize As Single, dpiX As Single, dpiY As Single, transform As DWRITE_MATRIX, isSideways As Boolean, outlineThreshold As DWRITE_OUTLINE_THRESHOLD, measuringMode As DWRITE_MEASURING_MODE, <Out> ByRef renderingMode As DWRITE_RENDERING_MODE) As HRESULT
        Function GetVerticalGlyphVariants(glyphCount As UInteger, nominalGlyphIndices As UShort, <Out> ByRef verticalGlyphIndices As UShort) As HRESULT
        <PreserveSig>
        Function HasVerticalGlyphVariants() As Boolean
    End Interface

    <ComImport>
    <Guid("d8b768ff-64bc-4e66-982b-ec8e87f693f7")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontFace2
        Inherits IDWriteFontFace1
#Region "IDWriteFontFace1"
#Region "IDWriteFontFace"
        <PreserveSig>
        Overloads Function [GetType]() As DWRITE_FONT_FACE_TYPE
        Overloads Function GetFiles(
<[In], Out> ByRef numberOfFiles As UInteger,
<Out, MarshalAs(UnmanagedType.LPArray)> fontFiles As IDWriteFontFile()) As HRESULT
        <PreserveSig>
        Overloads Function GetIndex() As Integer
        <PreserveSig>
        Overloads Function GetSimulations() As DWRITE_FONT_SIMULATIONS
        <PreserveSig>
        Overloads Function IsSymbolFont() As Boolean
        Overloads Sub GetMetrics(<Out> ByRef fontFaceMetrics As DWRITE_FONT_METRICS)
        <PreserveSig>
        Overloads Function GetGlyphCount() As UShort
        Overloads Function GetDesignGlyphMetrics(
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> glyphIndices As UShort(), glyphCount As Integer,
<Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> glyphMetrics As DWRITE_GLYPH_METRICS(), Optional isSideways As Boolean = False) As HRESULT
        <PreserveSig>
        Overloads Function GetGlyphIndices(
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> codePoints As UInteger(), codePointCount As Integer,
        <Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> glyphIndices As UShort()) As HRESULT
        Overloads Function TryGetFontTable(openTypeTableTag As Integer, <Out> ByRef tableData As IntPtr, <Out> ByRef tableSize As Integer, <Out> ByRef tableContext As IntPtr, <Out> ByRef exists As Boolean) As HRESULT
        Overloads Sub ReleaseFontTable(tableContext As IntPtr)
        Overloads Function GetGlyphRunOutline(emSize As Single,
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=4)> glyphIndices As UShort(),
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=4)> glyphAdvances As Single(),
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=4)> glyphOffsets As DWRITE_GLYPH_OFFSET(), glyphCount As Integer, isSideways As Boolean, isRightToLeft As Boolean, geometrySink As ID2D1SimplifiedGeometrySink) As HRESULT
        Overloads Function GetRecommendedRenderingMode(emSize As Single, pixelsPerDip As Single, measuringMode As DWRITE_MEASURING_MODE, renderingParams As IDWriteRenderingParams, <Out> ByRef renderingMode As DWRITE_RENDERING_MODE) As HRESULT
        Overloads Function GetGdiCompatibleMetrics(emSize As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, <Out> ByRef fontFaceMetrics As DWRITE_FONT_METRICS) As HRESULT
        Overloads Function GetGdiCompatibleGlyphMetrics(emSize As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, useGdiNatural As Boolean, glyphIndices As UShort, glyphCount As Integer, <Out> ByRef glyphMetrics As DWRITE_GLYPH_METRICS, Optional isSideways As Boolean = False) As HRESULT
#End Region

        Overloads Sub GetMetrics(<Out> ByRef fontMetrics As DWRITE_FONT_METRICS1)
        Overloads Function GetGdiCompatibleMetrics(emSize As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, <Out> ByRef fontMetrics As DWRITE_FONT_METRICS1) As HRESULT
        Overloads Sub GetCaretMetrics(<Out> ByRef caretMetrics As DWRITE_CARET_METRICS)
        Overloads Function GetUnicodeRanges(maxRangeCount As UInteger, <Out> ByRef unicodeRanges As DWRITE_UNICODE_RANGE, <Out> ByRef actualRangeCount As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function IsMonospacedFont() As Boolean
        Overloads Function GetDesignGlyphAdvances(glyphCount As UInteger, glyphIndices As UShort, <Out> ByRef glyphAdvances As Integer, Optional isSideways As Boolean = False) As HRESULT
        Overloads Function GetGdiCompatibleGlyphAdvances(emSize As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, useGdiNatural As Boolean, isSideways As Boolean, glyphCount As UInteger, glyphIndices As UShort, <Out> ByRef glyphAdvances As Integer) As HRESULT
        Overloads Function GetKerningPairAdjustments(glyphCount As UInteger, glyphIndices As UShort, <Out> ByRef glyphAdvanceAdjustments As Integer) As HRESULT
        <PreserveSig>
        Overloads Function HasKerningPairs() As Boolean
        Overloads Function GetRecommendedRenderingMode(fontEmSize As Single, dpiX As Single, dpiY As Single, transform As DWRITE_MATRIX, isSideways As Boolean, outlineThreshold As DWRITE_OUTLINE_THRESHOLD, measuringMode As DWRITE_MEASURING_MODE, <Out> ByRef renderingMode As DWRITE_RENDERING_MODE) As HRESULT
        Overloads Function GetVerticalGlyphVariants(glyphCount As UInteger, nominalGlyphIndices As UShort, <Out> ByRef verticalGlyphIndices As UShort) As HRESULT
        <PreserveSig>
        Overloads Function HasVerticalGlyphVariants() As Boolean
#End Region

        <PreserveSig>
        Function IsColorFont() As Boolean
        <PreserveSig>
        Function GetColorPaletteCount() As UInteger
        <PreserveSig>
        Function GetPaletteEntryCount() As UInteger
        Function GetPaletteEntries(colorPaletteIndex As UInteger, firstEntryIndex As UInteger, entryCount As UInteger, <Out> ByRef paletteEntries As DWRITE_COLOR_F) As HRESULT
        Overloads Function GetRecommendedRenderingMode(fontEmSize As Single, dpiX As Single, dpiY As Single, transform As DWRITE_MATRIX, isSideways As Boolean, outlineThreshold As DWRITE_OUTLINE_THRESHOLD, measuringMode As DWRITE_MEASURING_MODE, renderingParams As IDWriteRenderingParams, <Out> ByRef renderingMode As DWRITE_RENDERING_MODE, <Out> ByRef gridFitMode As DWRITE_GRID_FIT_MODE) As HRESULT
    End Interface

    <ComImport>
    <Guid("D37D7598-09BE-4222-A236-2081341CC1F2")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontFace3
        Inherits IDWriteFontFace2
#Region "IDWriteFontFace2"
#Region "IDWriteFontFace1"
#Region "IDWriteFontFace"
        <PreserveSig>
        Overloads Function [GetType]() As DWRITE_FONT_FACE_TYPE
        Overloads Function GetFiles(
<[In], Out> ByRef numberOfFiles As UInteger,
<Out, MarshalAs(UnmanagedType.LPArray)> fontFiles As IDWriteFontFile()) As HRESULT
        <PreserveSig>
        Overloads Function GetIndex() As Integer
        <PreserveSig>
        Overloads Function GetSimulations() As DWRITE_FONT_SIMULATIONS
        <PreserveSig>
        Overloads Function IsSymbolFont() As Boolean
        Overloads Sub GetMetrics(<Out> ByRef fontFaceMetrics As DWRITE_FONT_METRICS)
        <PreserveSig>
        Overloads Function GetGlyphCount() As UShort
        Overloads Function GetDesignGlyphMetrics(
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> glyphIndices As UShort(), glyphCount As Integer,
<Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> glyphMetrics As DWRITE_GLYPH_METRICS(), Optional isSideways As Boolean = False) As HRESULT
        <PreserveSig>
        Overloads Function GetGlyphIndices(
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> codePoints As UInteger(), codePointCount As Integer,
        <Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> glyphIndices As UShort()) As HRESULT
        Overloads Function TryGetFontTable(openTypeTableTag As Integer, <Out> ByRef tableData As IntPtr, <Out> ByRef tableSize As Integer, <Out> ByRef tableContext As IntPtr, <Out> ByRef exists As Boolean) As HRESULT
        Overloads Sub ReleaseFontTable(tableContext As IntPtr)
        Overloads Function GetGlyphRunOutline(emSize As Single,
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=4)> glyphIndices As UShort(),
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=4)> glyphAdvances As Single(),
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=4)> glyphOffsets As DWRITE_GLYPH_OFFSET(), glyphCount As Integer, isSideways As Boolean, isRightToLeft As Boolean, geometrySink As ID2D1SimplifiedGeometrySink) As HRESULT
        Overloads Function GetRecommendedRenderingMode(emSize As Single, pixelsPerDip As Single, measuringMode As DWRITE_MEASURING_MODE, renderingParams As IDWriteRenderingParams, <Out> ByRef renderingMode As DWRITE_RENDERING_MODE) As HRESULT
        Overloads Function GetGdiCompatibleMetrics(emSize As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, <Out> ByRef fontFaceMetrics As DWRITE_FONT_METRICS) As HRESULT
        Overloads Function GetGdiCompatibleGlyphMetrics(emSize As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, useGdiNatural As Boolean, glyphIndices As UShort, glyphCount As Integer, <Out> ByRef glyphMetrics As DWRITE_GLYPH_METRICS, Optional isSideways As Boolean = False) As HRESULT
#End Region

        Overloads Sub GetMetrics(<Out> ByRef fontMetrics As DWRITE_FONT_METRICS1)
        Overloads Function GetGdiCompatibleMetrics(emSize As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, <Out> ByRef fontMetrics As DWRITE_FONT_METRICS1) As HRESULT
        Overloads Sub GetCaretMetrics(<Out> ByRef caretMetrics As DWRITE_CARET_METRICS)
        Overloads Function GetUnicodeRanges(maxRangeCount As UInteger, <Out> ByRef unicodeRanges As DWRITE_UNICODE_RANGE, <Out> ByRef actualRangeCount As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function IsMonospacedFont() As Boolean
        Overloads Function GetDesignGlyphAdvances(glyphCount As UInteger, glyphIndices As UShort, <Out> ByRef glyphAdvances As Integer, Optional isSideways As Boolean = False) As HRESULT
        Overloads Function GetGdiCompatibleGlyphAdvances(emSize As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, useGdiNatural As Boolean, isSideways As Boolean, glyphCount As UInteger, glyphIndices As UShort, <Out> ByRef glyphAdvances As Integer) As HRESULT
        Overloads Function GetKerningPairAdjustments(glyphCount As UInteger, glyphIndices As UShort, <Out> ByRef glyphAdvanceAdjustments As Integer) As HRESULT
        <PreserveSig>
        Overloads Function HasKerningPairs() As Boolean
        Overloads Function GetRecommendedRenderingMode(fontEmSize As Single, dpiX As Single, dpiY As Single, transform As DWRITE_MATRIX, isSideways As Boolean, outlineThreshold As DWRITE_OUTLINE_THRESHOLD, measuringMode As DWRITE_MEASURING_MODE, <Out> ByRef renderingMode As DWRITE_RENDERING_MODE) As HRESULT
        Overloads Function GetVerticalGlyphVariants(glyphCount As UInteger, nominalGlyphIndices As UShort, <Out> ByRef verticalGlyphIndices As UShort) As HRESULT
        <PreserveSig>
        Overloads Function HasVerticalGlyphVariants() As Boolean
#End Region

        <PreserveSig>
        Overloads Function IsColorFont() As Boolean
        <PreserveSig>
        Overloads Function GetColorPaletteCount() As UInteger
        <PreserveSig>
        Overloads Function GetPaletteEntryCount() As UInteger
        Overloads Function GetPaletteEntries(colorPaletteIndex As UInteger, firstEntryIndex As UInteger, entryCount As UInteger, <Out> ByRef paletteEntries As DWRITE_COLOR_F) As HRESULT
        Overloads Function GetRecommendedRenderingMode(fontEmSize As Single, dpiX As Single, dpiY As Single, transform As DWRITE_MATRIX, isSideways As Boolean, outlineThreshold As DWRITE_OUTLINE_THRESHOLD, measuringMode As DWRITE_MEASURING_MODE, renderingParams As IDWriteRenderingParams, <Out> ByRef renderingMode As DWRITE_RENDERING_MODE, <Out> ByRef gridFitMode As DWRITE_GRID_FIT_MODE) As HRESULT
#End Region

        Function GetFontFaceReference(<Out> ByRef fontFaceReference As IDWriteFontFaceReference) As HRESULT
        'void GetPanose(out DWRITE_PANOSE panose);
        Sub GetPanose(<Out> ByRef panose As IntPtr)
        <PreserveSig>
        Function GetWeight() As DWRITE_FONT_WEIGHT
        <PreserveSig>
        Function GetStretch() As DWRITE_FONT_STRETCH
        <PreserveSig>
        Function GetStyle() As DWRITE_FONT_STYLE
        Function GetFamilyNames(<Out> ByRef names As IDWriteLocalizedStrings) As HRESULT
        Function GetFaceNames(<Out> ByRef names As IDWriteLocalizedStrings) As HRESULT
        Function GetInformationalStrings(informationalStringID As DWRITE_INFORMATIONAL_STRING_ID, <Out> ByRef informationalStrings As IDWriteLocalizedStrings, <Out> ByRef exists As Boolean) As HRESULT
        <PreserveSig>
        Function HasCharacter(unicodeValue As UInteger) As Boolean
        Overloads Function GetRecommendedRenderingMode(fontEmSize As Single, dpiX As Single, dpiY As Single, transform As DWRITE_MATRIX, isSideways As Boolean, outlineThreshold As DWRITE_OUTLINE_THRESHOLD, measuringMode As DWRITE_MEASURING_MODE, renderingParams As IDWriteRenderingParams, <Out> ByRef renderingMode As DWRITE_RENDERING_MODE1, <Out> ByRef gridFitMode As DWRITE_GRID_FIT_MODE) As HRESULT
        <PreserveSig>
        Function IsCharacterLocal(unicodeValue As UInteger) As Boolean
        <PreserveSig>
        Function IsGlyphLocal(glyphId As UShort) As Boolean
        Function AreCharactersLocal(characters As String, characterCount As UInteger, enqueueIfNotLocal As Boolean, <Out> ByRef isLocal As Boolean) As HRESULT
        Function AreGlyphsLocal(glyphIndices As UShort, glyphCount As UInteger, enqueueIfNotLocal As Boolean, <Out> ByRef isLocal As Boolean) As HRESULT
    End Interface

    <ComImport>
    <Guid("27F2A904-4EB8-441D-9678-0563F53E3E2F")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontFace4
        Inherits IDWriteFontFace3
#Region "IDWriteFontFace3"
#Region "IDWriteFontFace2"
#Region "IDWriteFontFace1"
#Region "IDWriteFontFace"
        <PreserveSig>
        Overloads Function [GetType]() As DWRITE_FONT_FACE_TYPE
        Overloads Function GetFiles(
<[In], Out> ByRef numberOfFiles As UInteger,
<Out, MarshalAs(UnmanagedType.LPArray)> fontFiles As IDWriteFontFile()) As HRESULT
        <PreserveSig>
        Overloads Function GetIndex() As Integer
        <PreserveSig>
        Overloads Function GetSimulations() As DWRITE_FONT_SIMULATIONS
        <PreserveSig>
        Overloads Function IsSymbolFont() As Boolean
        Overloads Sub GetMetrics(<Out> ByRef fontFaceMetrics As DWRITE_FONT_METRICS)
        <PreserveSig>
        Overloads Function GetGlyphCount() As UShort
        Overloads Function GetDesignGlyphMetrics(
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> glyphIndices As UShort(), glyphCount As Integer,
<Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> glyphMetrics As DWRITE_GLYPH_METRICS(), Optional isSideways As Boolean = False) As HRESULT
        <PreserveSig>
        Overloads Function GetGlyphIndices(
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> codePoints As UInteger(), codePointCount As Integer,
        <Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> glyphIndices As UShort()) As HRESULT
        Overloads Function TryGetFontTable(openTypeTableTag As Integer, <Out> ByRef tableData As IntPtr, <Out> ByRef tableSize As Integer, <Out> ByRef tableContext As IntPtr, <Out> ByRef exists As Boolean) As HRESULT
        Overloads Sub ReleaseFontTable(tableContext As IntPtr)
        Overloads Function GetGlyphRunOutline(emSize As Single,
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=4)> glyphIndices As UShort(),
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=4)> glyphAdvances As Single(),
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=4)> glyphOffsets As DWRITE_GLYPH_OFFSET(), glyphCount As Integer, isSideways As Boolean, isRightToLeft As Boolean, geometrySink As ID2D1SimplifiedGeometrySink) As HRESULT
        Overloads Function GetRecommendedRenderingMode(emSize As Single, pixelsPerDip As Single, measuringMode As DWRITE_MEASURING_MODE, renderingParams As IDWriteRenderingParams, <Out> ByRef renderingMode As DWRITE_RENDERING_MODE) As HRESULT
        Overloads Function GetGdiCompatibleMetrics(emSize As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, <Out> ByRef fontFaceMetrics As DWRITE_FONT_METRICS) As HRESULT
        Overloads Function GetGdiCompatibleGlyphMetrics(emSize As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, useGdiNatural As Boolean, glyphIndices As UShort, glyphCount As Integer, <Out> ByRef glyphMetrics As DWRITE_GLYPH_METRICS, Optional isSideways As Boolean = False) As HRESULT
#End Region

        Overloads Sub GetMetrics(<Out> ByRef fontMetrics As DWRITE_FONT_METRICS1)
        Overloads Function GetGdiCompatibleMetrics(emSize As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, <Out> ByRef fontMetrics As DWRITE_FONT_METRICS1) As HRESULT
        Overloads Sub GetCaretMetrics(<Out> ByRef caretMetrics As DWRITE_CARET_METRICS)
        Overloads Function GetUnicodeRanges(maxRangeCount As UInteger, <Out> ByRef unicodeRanges As DWRITE_UNICODE_RANGE, <Out> ByRef actualRangeCount As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function IsMonospacedFont() As Boolean
        Overloads Function GetDesignGlyphAdvances(glyphCount As UInteger, glyphIndices As UShort, <Out> ByRef glyphAdvances As Integer, Optional isSideways As Boolean = False) As HRESULT
        Overloads Function GetGdiCompatibleGlyphAdvances(emSize As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, useGdiNatural As Boolean, isSideways As Boolean, glyphCount As UInteger, glyphIndices As UShort, <Out> ByRef glyphAdvances As Integer) As HRESULT
        Overloads Function GetKerningPairAdjustments(glyphCount As UInteger, glyphIndices As UShort, <Out> ByRef glyphAdvanceAdjustments As Integer) As HRESULT
        <PreserveSig>
        Overloads Function HasKerningPairs() As Boolean
        Overloads Function GetRecommendedRenderingMode(fontEmSize As Single, dpiX As Single, dpiY As Single, transform As DWRITE_MATRIX, isSideways As Boolean, outlineThreshold As DWRITE_OUTLINE_THRESHOLD, measuringMode As DWRITE_MEASURING_MODE, <Out> ByRef renderingMode As DWRITE_RENDERING_MODE) As HRESULT
        Overloads Function GetVerticalGlyphVariants(glyphCount As UInteger, nominalGlyphIndices As UShort, <Out> ByRef verticalGlyphIndices As UShort) As HRESULT
        <PreserveSig>
        Overloads Function HasVerticalGlyphVariants() As Boolean
#End Region

        <PreserveSig>
        Overloads Function IsColorFont() As Boolean
        <PreserveSig>
        Overloads Function GetColorPaletteCount() As UInteger
        <PreserveSig>
        Overloads Function GetPaletteEntryCount() As UInteger
        Overloads Function GetPaletteEntries(colorPaletteIndex As UInteger, firstEntryIndex As UInteger, entryCount As UInteger, <Out> ByRef paletteEntries As DWRITE_COLOR_F) As HRESULT
        Overloads Function GetRecommendedRenderingMode(fontEmSize As Single, dpiX As Single, dpiY As Single, transform As DWRITE_MATRIX, isSideways As Boolean, outlineThreshold As DWRITE_OUTLINE_THRESHOLD, measuringMode As DWRITE_MEASURING_MODE, renderingParams As IDWriteRenderingParams, <Out> ByRef renderingMode As DWRITE_RENDERING_MODE, <Out> ByRef gridFitMode As DWRITE_GRID_FIT_MODE) As HRESULT
#End Region

        Overloads Function GetFontFaceReference(<Out> ByRef fontFaceReference As IDWriteFontFaceReference) As HRESULT
        'new void GetPanose(out DWRITE_PANOSE panose);
        Overloads Sub GetPanose(<Out> ByRef panose As IntPtr)
        <PreserveSig>
        Overloads Function GetWeight() As DWRITE_FONT_WEIGHT
        <PreserveSig>
        Overloads Function GetStretch() As DWRITE_FONT_STRETCH
        <PreserveSig>
        Overloads Function GetStyle() As DWRITE_FONT_STYLE
        Overloads Function GetFamilyNames(<Out> ByRef names As IDWriteLocalizedStrings) As HRESULT
        Overloads Function GetFaceNames(<Out> ByRef names As IDWriteLocalizedStrings) As HRESULT
        Overloads Function GetInformationalStrings(informationalStringID As DWRITE_INFORMATIONAL_STRING_ID, <Out> ByRef informationalStrings As IDWriteLocalizedStrings, <Out> ByRef exists As Boolean) As HRESULT
        <PreserveSig>
        Overloads Function HasCharacter(unicodeValue As UInteger) As Boolean
        Overloads Function GetRecommendedRenderingMode(fontEmSize As Single, dpiX As Single, dpiY As Single, transform As DWRITE_MATRIX, isSideways As Boolean, outlineThreshold As DWRITE_OUTLINE_THRESHOLD, measuringMode As DWRITE_MEASURING_MODE, renderingParams As IDWriteRenderingParams, <Out> ByRef renderingMode As DWRITE_RENDERING_MODE1, <Out> ByRef gridFitMode As DWRITE_GRID_FIT_MODE) As HRESULT
        <PreserveSig>
        Overloads Function IsCharacterLocal(unicodeValue As UInteger) As Boolean
        <PreserveSig>
        Overloads Function IsGlyphLocal(glyphId As UShort) As Boolean
        Overloads Function AreCharactersLocal(characters As String, characterCount As UInteger, enqueueIfNotLocal As Boolean, <Out> ByRef isLocal As Boolean) As HRESULT
        Overloads Function AreGlyphsLocal(glyphIndices As UShort, glyphCount As UInteger, enqueueIfNotLocal As Boolean, <Out> ByRef isLocal As Boolean) As HRESULT
#End Region

        <PreserveSig>
        Function GetGlyphImageFormats() As DWRITE_GLYPH_IMAGE_FORMATS
        Function GetGlyphImageFormats(glyphId As UShort, pixelsPerEmFirst As UInteger, pixelsPerEmLast As UInteger, <Out> ByRef glyphImageFormats As DWRITE_GLYPH_IMAGE_FORMATS) As HRESULT
        Function GetGlyphImageData(glyphId As UShort, pixelsPerEm As UInteger, glyphImageFormat As DWRITE_GLYPH_IMAGE_FORMATS, <Out> ByRef glyphData As DWRITE_GLYPH_IMAGE_DATA, <Out> ByRef glyphDataContext As IntPtr) As HRESULT
        Sub ReleaseGlyphImageData(glyphDataContext As IntPtr)
    End Interface

    <ComImport>
    <Guid("98EFF3A5-B667-479A-B145-E2FA5B9FDC29")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontFace5
        Inherits IDWriteFontFace4
#Region "IDWriteFontFace4"
#Region "IDWriteFontFace3"
#Region "IDWriteFontFace2"
#Region "IDWriteFontFace1"
#Region "IDWriteFontFace"
        <PreserveSig>
        Overloads Function [GetType]() As DWRITE_FONT_FACE_TYPE
        Overloads Function GetFiles(
<[In], Out> ByRef numberOfFiles As UInteger,
<Out, MarshalAs(UnmanagedType.LPArray)> fontFiles As IDWriteFontFile()) As HRESULT
        <PreserveSig>
        Overloads Function GetIndex() As Integer
        <PreserveSig>
        Overloads Function GetSimulations() As DWRITE_FONT_SIMULATIONS
        <PreserveSig>
        Overloads Function IsSymbolFont() As Boolean
        Overloads Sub GetMetrics(<Out> ByRef fontFaceMetrics As DWRITE_FONT_METRICS)
        <PreserveSig>
        Overloads Function GetGlyphCount() As UShort
        Overloads Function GetDesignGlyphMetrics(
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> glyphIndices As UShort(), glyphCount As Integer,
<Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> glyphMetrics As DWRITE_GLYPH_METRICS(), Optional isSideways As Boolean = False) As HRESULT
        <PreserveSig>
        Overloads Function GetGlyphIndices(
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> codePoints As UInteger(), codePointCount As Integer,
        <Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> glyphIndices As UShort()) As HRESULT
        Overloads Function TryGetFontTable(openTypeTableTag As Integer, <Out> ByRef tableData As IntPtr, <Out> ByRef tableSize As Integer, <Out> ByRef tableContext As IntPtr, <Out> ByRef exists As Boolean) As HRESULT
        Overloads Sub ReleaseFontTable(tableContext As IntPtr)
        Overloads Function GetGlyphRunOutline(emSize As Single,
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=4)> glyphIndices As UShort(),
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=4)> glyphAdvances As Single(),
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=4)> glyphOffsets As DWRITE_GLYPH_OFFSET(), glyphCount As Integer, isSideways As Boolean, isRightToLeft As Boolean, geometrySink As ID2D1SimplifiedGeometrySink) As HRESULT
        Overloads Function GetRecommendedRenderingMode(emSize As Single, pixelsPerDip As Single, measuringMode As DWRITE_MEASURING_MODE, renderingParams As IDWriteRenderingParams, <Out> ByRef renderingMode As DWRITE_RENDERING_MODE) As HRESULT
        Overloads Function GetGdiCompatibleMetrics(emSize As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, <Out> ByRef fontFaceMetrics As DWRITE_FONT_METRICS) As HRESULT
        Overloads Function GetGdiCompatibleGlyphMetrics(emSize As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, useGdiNatural As Boolean, glyphIndices As UShort, glyphCount As Integer, <Out> ByRef glyphMetrics As DWRITE_GLYPH_METRICS, Optional isSideways As Boolean = False) As HRESULT
#End Region

        Overloads Sub GetMetrics(<Out> ByRef fontMetrics As DWRITE_FONT_METRICS1)
        Overloads Function GetGdiCompatibleMetrics(emSize As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, <Out> ByRef fontMetrics As DWRITE_FONT_METRICS1) As HRESULT
        Overloads Sub GetCaretMetrics(<Out> ByRef caretMetrics As DWRITE_CARET_METRICS)
        Overloads Function GetUnicodeRanges(maxRangeCount As UInteger, <Out> ByRef unicodeRanges As DWRITE_UNICODE_RANGE, <Out> ByRef actualRangeCount As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function IsMonospacedFont() As Boolean
        Overloads Function GetDesignGlyphAdvances(glyphCount As UInteger, glyphIndices As UShort, <Out> ByRef glyphAdvances As Integer, Optional isSideways As Boolean = False) As HRESULT
        Overloads Function GetGdiCompatibleGlyphAdvances(emSize As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, useGdiNatural As Boolean, isSideways As Boolean, glyphCount As UInteger, glyphIndices As UShort, <Out> ByRef glyphAdvances As Integer) As HRESULT
        Overloads Function GetKerningPairAdjustments(glyphCount As UInteger, glyphIndices As UShort, <Out> ByRef glyphAdvanceAdjustments As Integer) As HRESULT
        <PreserveSig>
        Overloads Function HasKerningPairs() As Boolean
        Overloads Function GetRecommendedRenderingMode(fontEmSize As Single, dpiX As Single, dpiY As Single, transform As DWRITE_MATRIX, isSideways As Boolean, outlineThreshold As DWRITE_OUTLINE_THRESHOLD, measuringMode As DWRITE_MEASURING_MODE, <Out> ByRef renderingMode As DWRITE_RENDERING_MODE) As HRESULT
        Overloads Function GetVerticalGlyphVariants(glyphCount As UInteger, nominalGlyphIndices As UShort, <Out> ByRef verticalGlyphIndices As UShort) As HRESULT
        <PreserveSig>
        Overloads Function HasVerticalGlyphVariants() As Boolean
#End Region

        <PreserveSig>
        Overloads Function IsColorFont() As Boolean
        <PreserveSig>
        Overloads Function GetColorPaletteCount() As UInteger
        <PreserveSig>
        Overloads Function GetPaletteEntryCount() As UInteger
        Overloads Function GetPaletteEntries(colorPaletteIndex As UInteger, firstEntryIndex As UInteger, entryCount As UInteger, <Out> ByRef paletteEntries As DWRITE_COLOR_F) As HRESULT
        Overloads Function GetRecommendedRenderingMode(fontEmSize As Single, dpiX As Single, dpiY As Single, transform As DWRITE_MATRIX, isSideways As Boolean, outlineThreshold As DWRITE_OUTLINE_THRESHOLD, measuringMode As DWRITE_MEASURING_MODE, renderingParams As IDWriteRenderingParams, <Out> ByRef renderingMode As DWRITE_RENDERING_MODE, <Out> ByRef gridFitMode As DWRITE_GRID_FIT_MODE) As HRESULT
#End Region

        Overloads Function GetFontFaceReference(<Out> ByRef fontFaceReference As IDWriteFontFaceReference) As HRESULT
        'new void GetPanose(out DWRITE_PANOSE panose);
        Overloads Sub GetPanose(<Out> ByRef panose As IntPtr)
        <PreserveSig>
        Overloads Function GetWeight() As DWRITE_FONT_WEIGHT
        <PreserveSig>
        Overloads Function GetStretch() As DWRITE_FONT_STRETCH
        <PreserveSig>
        Overloads Function GetStyle() As DWRITE_FONT_STYLE
        Overloads Function GetFamilyNames(<Out> ByRef names As IDWriteLocalizedStrings) As HRESULT
        Overloads Function GetFaceNames(<Out> ByRef names As IDWriteLocalizedStrings) As HRESULT
        Overloads Function GetInformationalStrings(informationalStringID As DWRITE_INFORMATIONAL_STRING_ID, <Out> ByRef informationalStrings As IDWriteLocalizedStrings, <Out> ByRef exists As Boolean) As HRESULT
        <PreserveSig>
        Overloads Function HasCharacter(unicodeValue As UInteger) As Boolean
        Overloads Function GetRecommendedRenderingMode(fontEmSize As Single, dpiX As Single, dpiY As Single, transform As DWRITE_MATRIX, isSideways As Boolean, outlineThreshold As DWRITE_OUTLINE_THRESHOLD, measuringMode As DWRITE_MEASURING_MODE, renderingParams As IDWriteRenderingParams, <Out> ByRef renderingMode As DWRITE_RENDERING_MODE1, <Out> ByRef gridFitMode As DWRITE_GRID_FIT_MODE) As HRESULT
        <PreserveSig>
        Overloads Function IsCharacterLocal(unicodeValue As UInteger) As Boolean
        <PreserveSig>
        Overloads Function IsGlyphLocal(glyphId As UShort) As Boolean
        Overloads Function AreCharactersLocal(characters As String, characterCount As UInteger, enqueueIfNotLocal As Boolean, <Out> ByRef isLocal As Boolean) As HRESULT
        Overloads Function AreGlyphsLocal(glyphIndices As UShort, glyphCount As UInteger, enqueueIfNotLocal As Boolean, <Out> ByRef isLocal As Boolean) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function GetGlyphImageFormats() As DWRITE_GLYPH_IMAGE_FORMATS
        Overloads Function GetGlyphImageFormats(glyphId As UShort, pixelsPerEmFirst As UInteger, pixelsPerEmLast As UInteger, <Out> ByRef glyphImageFormats As DWRITE_GLYPH_IMAGE_FORMATS) As HRESULT
        Overloads Function GetGlyphImageData(glyphId As UShort, pixelsPerEm As UInteger, glyphImageFormat As DWRITE_GLYPH_IMAGE_FORMATS, <Out> ByRef glyphData As DWRITE_GLYPH_IMAGE_DATA, <Out> ByRef glyphDataContext As IntPtr) As HRESULT
        Overloads Sub ReleaseGlyphImageData(glyphDataContext As IntPtr)
#End Region

        <PreserveSig>
        Function GetFontAxisValueCount() As UInteger
        <PreserveSig>
        Function GetFontAxisValues(<Out> ByRef fontAxisValues As DWRITE_FONT_AXIS_VALUE, fontAxisValueCount As UInteger) As HRESULT
        <PreserveSig>
        Function HasVariations() As Boolean
        <PreserveSig>
        Function GetFontResource(<Out> ByRef fontResource As IDWriteFontResource) As HRESULT
        <PreserveSig>
        Function Equals(fontFace As IDWriteFontFace) As Boolean
    End Interface

    <ComImport>
    <Guid("C4B1FE1B-6E84-47D5-B54C-A597981B06AD")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontFace6
        Inherits IDWriteFontFace5
#Region "IDWriteFontFace4"
#Region "IDWriteFontFace3"
#Region "IDWriteFontFace2"
#Region "IDWriteFontFace1"
#Region "IDWriteFontFace"
        <PreserveSig>
        Overloads Function [GetType]() As DWRITE_FONT_FACE_TYPE
        Overloads Function GetFiles(
<[In], Out> ByRef numberOfFiles As UInteger,
<Out, MarshalAs(UnmanagedType.LPArray)> fontFiles As IDWriteFontFile()) As HRESULT
        <PreserveSig>
        Overloads Function GetIndex() As Integer
        <PreserveSig>
        Overloads Function GetSimulations() As DWRITE_FONT_SIMULATIONS
        <PreserveSig>
        Overloads Function IsSymbolFont() As Boolean
        Overloads Sub GetMetrics(<Out> ByRef fontFaceMetrics As DWRITE_FONT_METRICS)
        <PreserveSig>
        Overloads Function GetGlyphCount() As UShort
        Overloads Function GetDesignGlyphMetrics(
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> glyphIndices As UShort(), glyphCount As Integer,
<Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> glyphMetrics As DWRITE_GLYPH_METRICS(), Optional isSideways As Boolean = False) As HRESULT
        <PreserveSig>
        Overloads Function GetGlyphIndices(
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> codePoints As UInteger(), codePointCount As Integer,
        <Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> glyphIndices As UShort()) As HRESULT
        Overloads Function TryGetFontTable(openTypeTableTag As Integer, <Out> ByRef tableData As IntPtr, <Out> ByRef tableSize As Integer, <Out> ByRef tableContext As IntPtr, <Out> ByRef exists As Boolean) As HRESULT
        Overloads Sub ReleaseFontTable(tableContext As IntPtr)
        Overloads Function GetGlyphRunOutline(emSize As Single,
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=4)> glyphIndices As UShort(),
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=4)> glyphAdvances As Single(),
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=4)> glyphOffsets As DWRITE_GLYPH_OFFSET(), glyphCount As Integer, isSideways As Boolean, isRightToLeft As Boolean, geometrySink As ID2D1SimplifiedGeometrySink) As HRESULT
        Overloads Function GetRecommendedRenderingMode(emSize As Single, pixelsPerDip As Single, measuringMode As DWRITE_MEASURING_MODE, renderingParams As IDWriteRenderingParams, <Out> ByRef renderingMode As DWRITE_RENDERING_MODE) As HRESULT
        Overloads Function GetGdiCompatibleMetrics(emSize As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, <Out> ByRef fontFaceMetrics As DWRITE_FONT_METRICS) As HRESULT
        Overloads Function GetGdiCompatibleGlyphMetrics(emSize As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, useGdiNatural As Boolean, glyphIndices As UShort, glyphCount As Integer, <Out> ByRef glyphMetrics As DWRITE_GLYPH_METRICS, Optional isSideways As Boolean = False) As HRESULT
#End Region

        Overloads Sub GetMetrics(<Out> ByRef fontMetrics As DWRITE_FONT_METRICS1)
        Overloads Function GetGdiCompatibleMetrics(emSize As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, <Out> ByRef fontMetrics As DWRITE_FONT_METRICS1) As HRESULT
        Overloads Sub GetCaretMetrics(<Out> ByRef caretMetrics As DWRITE_CARET_METRICS)
        Overloads Function GetUnicodeRanges(maxRangeCount As UInteger, <Out> ByRef unicodeRanges As DWRITE_UNICODE_RANGE, <Out> ByRef actualRangeCount As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function IsMonospacedFont() As Boolean
        Overloads Function GetDesignGlyphAdvances(glyphCount As UInteger, glyphIndices As UShort, <Out> ByRef glyphAdvances As Integer, Optional isSideways As Boolean = False) As HRESULT
        Overloads Function GetGdiCompatibleGlyphAdvances(emSize As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, useGdiNatural As Boolean, isSideways As Boolean, glyphCount As UInteger, glyphIndices As UShort, <Out> ByRef glyphAdvances As Integer) As HRESULT
        Overloads Function GetKerningPairAdjustments(glyphCount As UInteger, glyphIndices As UShort, <Out> ByRef glyphAdvanceAdjustments As Integer) As HRESULT
        <PreserveSig>
        Overloads Function HasKerningPairs() As Boolean
        Overloads Function GetRecommendedRenderingMode(fontEmSize As Single, dpiX As Single, dpiY As Single, transform As DWRITE_MATRIX, isSideways As Boolean, outlineThreshold As DWRITE_OUTLINE_THRESHOLD, measuringMode As DWRITE_MEASURING_MODE, <Out> ByRef renderingMode As DWRITE_RENDERING_MODE) As HRESULT
        Overloads Function GetVerticalGlyphVariants(glyphCount As UInteger, nominalGlyphIndices As UShort, <Out> ByRef verticalGlyphIndices As UShort) As HRESULT
        <PreserveSig>
        Overloads Function HasVerticalGlyphVariants() As Boolean
#End Region

        <PreserveSig>
        Overloads Function IsColorFont() As Boolean
        <PreserveSig>
        Overloads Function GetColorPaletteCount() As UInteger
        <PreserveSig>
        Overloads Function GetPaletteEntryCount() As UInteger
        Overloads Function GetPaletteEntries(colorPaletteIndex As UInteger, firstEntryIndex As UInteger, entryCount As UInteger, <Out> ByRef paletteEntries As DWRITE_COLOR_F) As HRESULT
        Overloads Function GetRecommendedRenderingMode(fontEmSize As Single, dpiX As Single, dpiY As Single, transform As DWRITE_MATRIX, isSideways As Boolean, outlineThreshold As DWRITE_OUTLINE_THRESHOLD, measuringMode As DWRITE_MEASURING_MODE, renderingParams As IDWriteRenderingParams, <Out> ByRef renderingMode As DWRITE_RENDERING_MODE, <Out> ByRef gridFitMode As DWRITE_GRID_FIT_MODE) As HRESULT
#End Region

        Overloads Function GetFontFaceReference(<Out> ByRef fontFaceReference As IDWriteFontFaceReference) As HRESULT
        'new void GetPanose(out DWRITE_PANOSE panose);
        Overloads Sub GetPanose(<Out> ByRef panose As IntPtr)
        <PreserveSig>
        Overloads Function GetWeight() As DWRITE_FONT_WEIGHT
        <PreserveSig>
        Overloads Function GetStretch() As DWRITE_FONT_STRETCH
        <PreserveSig>
        Overloads Function GetStyle() As DWRITE_FONT_STYLE
        Overloads Function GetFamilyNames(<Out> ByRef names As IDWriteLocalizedStrings) As HRESULT
        Overloads Function GetFaceNames(<Out> ByRef names As IDWriteLocalizedStrings) As HRESULT
        Overloads Function GetInformationalStrings(informationalStringID As DWRITE_INFORMATIONAL_STRING_ID, <Out> ByRef informationalStrings As IDWriteLocalizedStrings, <Out> ByRef exists As Boolean) As HRESULT
        <PreserveSig>
        Overloads Function HasCharacter(unicodeValue As UInteger) As Boolean
        Overloads Function GetRecommendedRenderingMode(fontEmSize As Single, dpiX As Single, dpiY As Single, transform As DWRITE_MATRIX, isSideways As Boolean, outlineThreshold As DWRITE_OUTLINE_THRESHOLD, measuringMode As DWRITE_MEASURING_MODE, renderingParams As IDWriteRenderingParams, <Out> ByRef renderingMode As DWRITE_RENDERING_MODE1, <Out> ByRef gridFitMode As DWRITE_GRID_FIT_MODE) As HRESULT
        <PreserveSig>
        Overloads Function IsCharacterLocal(unicodeValue As UInteger) As Boolean
        <PreserveSig>
        Overloads Function IsGlyphLocal(glyphId As UShort) As Boolean
        Overloads Function AreCharactersLocal(characters As String, characterCount As UInteger, enqueueIfNotLocal As Boolean, <Out> ByRef isLocal As Boolean) As HRESULT
        Overloads Function AreGlyphsLocal(glyphIndices As UShort, glyphCount As UInteger, enqueueIfNotLocal As Boolean, <Out> ByRef isLocal As Boolean) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function GetGlyphImageFormats() As DWRITE_GLYPH_IMAGE_FORMATS
        Overloads Function GetGlyphImageFormats(glyphId As UShort, pixelsPerEmFirst As UInteger, pixelsPerEmLast As UInteger, <Out> ByRef glyphImageFormats As DWRITE_GLYPH_IMAGE_FORMATS) As HRESULT
        Overloads Function GetGlyphImageData(glyphId As UShort, pixelsPerEm As UInteger, glyphImageFormat As DWRITE_GLYPH_IMAGE_FORMATS, <Out> ByRef glyphData As DWRITE_GLYPH_IMAGE_DATA, <Out> ByRef glyphDataContext As IntPtr) As HRESULT
        Overloads Sub ReleaseGlyphImageData(glyphDataContext As IntPtr)
#End Region

#Region "<IDWriteFontFace5>"
        <PreserveSig>
        Overloads Function GetFontAxisValueCount() As UInteger
        <PreserveSig>
        Overloads Function GetFontAxisValues(<Out> ByRef fontAxisValues As DWRITE_FONT_AXIS_VALUE, fontAxisValueCount As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function HasVariations() As Boolean
        <PreserveSig>
        Overloads Function GetFontResource(<Out> ByRef fontResource As IDWriteFontResource) As HRESULT
        <PreserveSig>
        Overloads Function Equals(fontFace As IDWriteFontFace) As Boolean
#End Region

        <PreserveSig>
        Overloads Function GetFamilyNames(fontFamilyModel As DWRITE_FONT_FAMILY_MODEL, <Out> ByRef names As IDWriteLocalizedStrings) As HRESULT
        <PreserveSig>
        Overloads Function GetFaceNames(fontFamilyModel As DWRITE_FONT_FAMILY_MODEL, <Out> ByRef names As IDWriteLocalizedStrings) As HRESULT
    End Interface

    <ComImport>
    <Guid("3945B85B-BC95-40F7-B72C-8B73BFC7E13B")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontFace7
        Inherits IDWriteFontFace6
#Region "IDWriteFontFace4"
#Region "IDWriteFontFace3"
#Region "IDWriteFontFace2"
#Region "IDWriteFontFace1"
#Region "IDWriteFontFace"
        <PreserveSig>
        Overloads Function [GetType]() As DWRITE_FONT_FACE_TYPE
        Overloads Function GetFiles(
<[In], Out> ByRef numberOfFiles As UInteger,
<Out, MarshalAs(UnmanagedType.LPArray)> fontFiles As IDWriteFontFile()) As HRESULT
        <PreserveSig>
        Overloads Function GetIndex() As Integer
        <PreserveSig>
        Overloads Function GetSimulations() As DWRITE_FONT_SIMULATIONS
        <PreserveSig>
        Overloads Function IsSymbolFont() As Boolean
        Overloads Sub GetMetrics(<Out> ByRef fontFaceMetrics As DWRITE_FONT_METRICS)
        <PreserveSig>
        Overloads Function GetGlyphCount() As UShort
        Overloads Function GetDesignGlyphMetrics(
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> glyphIndices As UShort(), glyphCount As Integer,
<Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> glyphMetrics As DWRITE_GLYPH_METRICS(), Optional isSideways As Boolean = False) As HRESULT
        <PreserveSig>
        Overloads Function GetGlyphIndices(
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> codePoints As UInteger(), codePointCount As Integer,
        <Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> glyphIndices As UShort()) As HRESULT
        Overloads Function TryGetFontTable(openTypeTableTag As Integer, <Out> ByRef tableData As IntPtr, <Out> ByRef tableSize As Integer, <Out> ByRef tableContext As IntPtr, <Out> ByRef exists As Boolean) As HRESULT
        Overloads Sub ReleaseFontTable(tableContext As IntPtr)
        Overloads Function GetGlyphRunOutline(emSize As Single,
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=4)> glyphIndices As UShort(),
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=4)> glyphAdvances As Single(),
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=4)> glyphOffsets As DWRITE_GLYPH_OFFSET(), glyphCount As Integer, isSideways As Boolean, isRightToLeft As Boolean, geometrySink As ID2D1SimplifiedGeometrySink) As HRESULT
        Overloads Function GetRecommendedRenderingMode(emSize As Single, pixelsPerDip As Single, measuringMode As DWRITE_MEASURING_MODE, renderingParams As IDWriteRenderingParams, <Out> ByRef renderingMode As DWRITE_RENDERING_MODE) As HRESULT
        Overloads Function GetGdiCompatibleMetrics(emSize As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, <Out> ByRef fontFaceMetrics As DWRITE_FONT_METRICS) As HRESULT
        Overloads Function GetGdiCompatibleGlyphMetrics(emSize As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, useGdiNatural As Boolean, glyphIndices As UShort, glyphCount As Integer, <Out> ByRef glyphMetrics As DWRITE_GLYPH_METRICS, Optional isSideways As Boolean = False) As HRESULT
#End Region

        Overloads Sub GetMetrics(<Out> ByRef fontMetrics As DWRITE_FONT_METRICS1)
        Overloads Function GetGdiCompatibleMetrics(emSize As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, <Out> ByRef fontMetrics As DWRITE_FONT_METRICS1) As HRESULT
        Overloads Sub GetCaretMetrics(<Out> ByRef caretMetrics As DWRITE_CARET_METRICS)
        Overloads Function GetUnicodeRanges(maxRangeCount As UInteger, <Out> ByRef unicodeRanges As DWRITE_UNICODE_RANGE, <Out> ByRef actualRangeCount As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function IsMonospacedFont() As Boolean
        Overloads Function GetDesignGlyphAdvances(glyphCount As UInteger, glyphIndices As UShort, <Out> ByRef glyphAdvances As Integer, Optional isSideways As Boolean = False) As HRESULT
        Overloads Function GetGdiCompatibleGlyphAdvances(emSize As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, useGdiNatural As Boolean, isSideways As Boolean, glyphCount As UInteger, glyphIndices As UShort, <Out> ByRef glyphAdvances As Integer) As HRESULT
        Overloads Function GetKerningPairAdjustments(glyphCount As UInteger, glyphIndices As UShort, <Out> ByRef glyphAdvanceAdjustments As Integer) As HRESULT
        <PreserveSig>
        Overloads Function HasKerningPairs() As Boolean
        Overloads Function GetRecommendedRenderingMode(fontEmSize As Single, dpiX As Single, dpiY As Single, transform As DWRITE_MATRIX, isSideways As Boolean, outlineThreshold As DWRITE_OUTLINE_THRESHOLD, measuringMode As DWRITE_MEASURING_MODE, <Out> ByRef renderingMode As DWRITE_RENDERING_MODE) As HRESULT
        Overloads Function GetVerticalGlyphVariants(glyphCount As UInteger, nominalGlyphIndices As UShort, <Out> ByRef verticalGlyphIndices As UShort) As HRESULT
        <PreserveSig>
        Overloads Function HasVerticalGlyphVariants() As Boolean
#End Region

        <PreserveSig>
        Overloads Function IsColorFont() As Boolean
        <PreserveSig>
        Overloads Function GetColorPaletteCount() As UInteger
        <PreserveSig>
        Overloads Function GetPaletteEntryCount() As UInteger
        Overloads Function GetPaletteEntries(colorPaletteIndex As UInteger, firstEntryIndex As UInteger, entryCount As UInteger, <Out> ByRef paletteEntries As DWRITE_COLOR_F) As HRESULT
        Overloads Function GetRecommendedRenderingMode(fontEmSize As Single, dpiX As Single, dpiY As Single, transform As DWRITE_MATRIX, isSideways As Boolean, outlineThreshold As DWRITE_OUTLINE_THRESHOLD, measuringMode As DWRITE_MEASURING_MODE, renderingParams As IDWriteRenderingParams, <Out> ByRef renderingMode As DWRITE_RENDERING_MODE, <Out> ByRef gridFitMode As DWRITE_GRID_FIT_MODE) As HRESULT
#End Region

        Overloads Function GetFontFaceReference(<Out> ByRef fontFaceReference As IDWriteFontFaceReference) As HRESULT
        'new void GetPanose(out DWRITE_PANOSE panose);
        Overloads Sub GetPanose(<Out> ByRef panose As IntPtr)
        <PreserveSig>
        Overloads Function GetWeight() As DWRITE_FONT_WEIGHT
        <PreserveSig>
        Overloads Function GetStretch() As DWRITE_FONT_STRETCH
        <PreserveSig>
        Overloads Function GetStyle() As DWRITE_FONT_STYLE
        Overloads Function GetFamilyNames(<Out> ByRef names As IDWriteLocalizedStrings) As HRESULT
        Overloads Function GetFaceNames(<Out> ByRef names As IDWriteLocalizedStrings) As HRESULT
        Overloads Function GetInformationalStrings(informationalStringID As DWRITE_INFORMATIONAL_STRING_ID, <Out> ByRef informationalStrings As IDWriteLocalizedStrings, <Out> ByRef exists As Boolean) As HRESULT
        <PreserveSig>
        Overloads Function HasCharacter(unicodeValue As UInteger) As Boolean
        Overloads Function GetRecommendedRenderingMode(fontEmSize As Single, dpiX As Single, dpiY As Single, transform As DWRITE_MATRIX, isSideways As Boolean, outlineThreshold As DWRITE_OUTLINE_THRESHOLD, measuringMode As DWRITE_MEASURING_MODE, renderingParams As IDWriteRenderingParams, <Out> ByRef renderingMode As DWRITE_RENDERING_MODE1, <Out> ByRef gridFitMode As DWRITE_GRID_FIT_MODE) As HRESULT
        <PreserveSig>
        Overloads Function IsCharacterLocal(unicodeValue As UInteger) As Boolean
        <PreserveSig>
        Overloads Function IsGlyphLocal(glyphId As UShort) As Boolean
        Overloads Function AreCharactersLocal(characters As String, characterCount As UInteger, enqueueIfNotLocal As Boolean, <Out> ByRef isLocal As Boolean) As HRESULT
        Overloads Function AreGlyphsLocal(glyphIndices As UShort, glyphCount As UInteger, enqueueIfNotLocal As Boolean, <Out> ByRef isLocal As Boolean) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function GetGlyphImageFormats() As DWRITE_GLYPH_IMAGE_FORMATS
        Overloads Function GetGlyphImageFormats(glyphId As UShort, pixelsPerEmFirst As UInteger, pixelsPerEmLast As UInteger, <Out> ByRef glyphImageFormats As DWRITE_GLYPH_IMAGE_FORMATS) As HRESULT
        Overloads Function GetGlyphImageData(glyphId As UShort, pixelsPerEm As UInteger, glyphImageFormat As DWRITE_GLYPH_IMAGE_FORMATS, <Out> ByRef glyphData As DWRITE_GLYPH_IMAGE_DATA, <Out> ByRef glyphDataContext As IntPtr) As HRESULT
        Overloads Sub ReleaseGlyphImageData(glyphDataContext As IntPtr)
#End Region

#Region "<IDWriteFontFace5>"
        <PreserveSig>
        Overloads Function GetFontAxisValueCount() As UInteger
        <PreserveSig>
        Overloads Function GetFontAxisValues(<Out> ByRef fontAxisValues As DWRITE_FONT_AXIS_VALUE, fontAxisValueCount As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function HasVariations() As Boolean
        <PreserveSig>
        Overloads Function GetFontResource(<Out> ByRef fontResource As IDWriteFontResource) As HRESULT
        <PreserveSig>
        Overloads Function Equals(fontFace As IDWriteFontFace) As Boolean
#End Region

#Region "<IDWriteFontFace6>"
        <PreserveSig>
        Overloads Function GetFamilyNames(fontFamilyModel As DWRITE_FONT_FAMILY_MODEL, <Out> ByRef names As IDWriteLocalizedStrings) As HRESULT
        <PreserveSig>
        Overloads Function GetFaceNames(fontFamilyModel As DWRITE_FONT_FAMILY_MODEL, <Out> ByRef names As IDWriteLocalizedStrings) As HRESULT
#End Region

        <PreserveSig>
        Function GetPaintFeatureLevel(glyphImageFormat As DWRITE_GLYPH_IMAGE_FORMATS) As DWRITE_PAINT_FEATURE_LEVEL
        <PreserveSig>
        Function CreatePaintReader(glyphImageFormat As DWRITE_GLYPH_IMAGE_FORMATS, paintFeatureLevel As DWRITE_PAINT_FEATURE_LEVEL, <Out> ByRef paintReader As IDWritePaintReader) As HRESULT
    End Interface

    <ComImport>
    <Guid("8128E912-3B97-42A5-AB6C-24AAD3A86E54")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWritePaintReader
        Function SetCurrentGlyph(glyphIndex As UInteger, paintElement As IntPtr, structSize As UInteger, <Out> ByRef clipBox As Direct2D.D2D1_RECT_F, glyphAttributes As IntPtr) As HRESULT
        Function SetTextColor(ByRef textColor As DWRITE_COLOR_F) As HRESULT
        Function SetColorPaletteIndex(colorPaletteIndex As UInteger) As HRESULT
        Function SetCustomColorPalette(
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> paletteEntries As DWRITE_COLOR_F(), paletteEntryCount As UInteger) As HRESULT
        Function MoveToFirstChild(paintElement As IntPtr, structSize As UInteger) As HRESULT
        Function MoveToNextSibling(paintElement As IntPtr, structSize As UInteger) As HRESULT
        Function MoveToParent() As HRESULT
        Function GetGradientStops(firstGradientStopIndex As UInteger, gradientStopCount As UInteger,
            <Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> gradientStops As Direct2D.D2D1_GRADIENT_STOP()) As HRESULT
        Function GetGradientStopColors(firstGradientStopIndex As UInteger, gradientStopCount As UInteger,
            <Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> gradientStopColors As DWRITE_PAINT_COLOR()) As HRESULT
    End Interface

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DWRITE_PAINT_COLOR
        Public value As DWRITE_COLOR_F
        Public paletteEntryIndex As UShort
        Public alphaMultiplier As Single
        Public colorAttributes As DWRITE_PAINT_ATTRIBUTES
    End Structure

    Public Enum DWRITE_PAINT_ATTRIBUTES
        DWRITE_PAINT_ATTRIBUTES_NONE = 0

        ''' <summary>
        ''' Specifies that the color value (or any color value in the glyph) comes from the font's
        ''' color palette. This means the appearance may depend on the current palette index, which
        ''' may be important to clients that cache color glyphs.
        ''' </summary>
        DWRITE_PAINT_ATTRIBUTES_USES_PALETTE = &H01

        ''' <summary>
        ''' Specifies that the color value (or any color value in the glyph) comes from the client-specified
        ''' text color. This means the appearance may depend on the text color, which may be important to
        ''' clients that cache color glyphs.
        ''' </summary>
        DWRITE_PAINT_ATTRIBUTES_USES_TEXT_COLOR = &H02
    End Enum

    Public Enum DWRITE_COLOR_COMPOSITE_MODE
        ' Porter-Duff modes.
        DWRITE_COLOR_COMPOSITE_CLEAR
        DWRITE_COLOR_COMPOSITE_SRC
        DWRITE_COLOR_COMPOSITE_DEST
        DWRITE_COLOR_COMPOSITE_SRC_OVER
        DWRITE_COLOR_COMPOSITE_DEST_OVER
        DWRITE_COLOR_COMPOSITE_SRC_IN
        DWRITE_COLOR_COMPOSITE_DEST_IN
        DWRITE_COLOR_COMPOSITE_SRC_OUT
        DWRITE_COLOR_COMPOSITE_DEST_OUT
        DWRITE_COLOR_COMPOSITE_SRC_ATOP
        DWRITE_COLOR_COMPOSITE_DEST_ATOP
        DWRITE_COLOR_COMPOSITE_XOR
        DWRITE_COLOR_COMPOSITE_PLUS

        ' Separable color blend modes.
        DWRITE_COLOR_COMPOSITE_SCREEN
        DWRITE_COLOR_COMPOSITE_OVERLAY
        DWRITE_COLOR_COMPOSITE_DARKEN
        DWRITE_COLOR_COMPOSITE_LIGHTEN
        DWRITE_COLOR_COMPOSITE_COLOR_DODGE
        DWRITE_COLOR_COMPOSITE_COLOR_BURN
        DWRITE_COLOR_COMPOSITE_HARD_LIGHT
        DWRITE_COLOR_COMPOSITE_SOFT_LIGHT
        DWRITE_COLOR_COMPOSITE_DIFFERENCE
        DWRITE_COLOR_COMPOSITE_EXCLUSION
        DWRITE_COLOR_COMPOSITE_MULTIPLY

        ' Non-separable color blend modes.
        DWRITE_COLOR_COMPOSITE_HSL_HUE
        DWRITE_COLOR_COMPOSITE_HSL_SATURATION
        DWRITE_COLOR_COMPOSITE_HSL_COLOR
        DWRITE_COLOR_COMPOSITE_HSL_LUMINOSITY
    End Enum

    Public Partial Structure DWRITE_PAINT_ELEMENT
        Public paintType As DWRITE_PAINT_TYPE
        Public paint As PAINT_UNION
        <StructLayout(LayoutKind.Explicit)>
        Public Partial Structure PAINT_UNION
            <FieldOffset(0)>
            Public layers As PAINT_LAYERS
            <FieldOffset(0)>
            Public solidGlyph As PAINT_SOLID_GLYPH
            <FieldOffset(0)>
            Public solid As DWRITE_PAINT_COLOR
            <FieldOffset(0)>
            Public linearGradient As PAINT_LINEAR_GRADIENT
            <FieldOffset(0)>
            Public radialGradient As PAINT_RADIAL_GRADIENT
            <FieldOffset(0)>
            Public sweepGradient As PAINT_SWEEP_GRADIENT
            <FieldOffset(0)>
            Public glyph As PAINT_GLYPH
            <FieldOffset(0)>
            Public colorGlyph As PAINT_COLOR_GLYPH
            <FieldOffset(0)>
            Public transform As DWRITE_MATRIX
            <FieldOffset(0)>
            Public composite As PAINT_COMPOSITE
            Public Partial Structure PAINT_LAYERS
                Public childCount As UInteger
            End Structure
            Public Partial Structure PAINT_SOLID_GLYPH
                Public glyphIndex As UInteger
                Public color As DWRITE_PAINT_COLOR
            End Structure
            Public Partial Structure PAINT_LINEAR_GRADIENT
                Public extendMode As UInteger
                Public gradientStopCount As UInteger
                Public x0 As Single
                Public y0 As Single
                Public x1 As Single
                Public y1 As Single
                Public x2 As Single
                Public y2 As Single
            End Structure
            Public Partial Structure PAINT_RADIAL_GRADIENT
                Public extendMode As UInteger
                Public gradientStopCount As UInteger
                Public x0 As Single
                Public y0 As Single
                Public radius0 As Single
                Public x1 As Single
                Public y1 As Single
                Public radius1 As Single
            End Structure
            Public Partial Structure PAINT_SWEEP_GRADIENT
                Public extendMode As UInteger
                Public gradientStopCount As UInteger
                Public centerX As Single
                Public centerY As Single
                Public startAngle As Single
                Public endAngle As Single
            End Structure
            Public Partial Structure PAINT_GLYPH
                Public glyphIndex As UInteger
            End Structure
            Public Partial Structure PAINT_COLOR_GLYPH
                Public glyphIndex As UInteger
                Public clipBox As Direct2D.D2D1_RECT_F
            End Structure
            Public Partial Structure PAINT_COMPOSITE
                Public mode As DWRITE_COLOR_COMPOSITE_MODE
            End Structure
        End Structure
    End Structure

    Public Enum DWRITE_PAINT_TYPE
        ' The following paint types may be returned for color feature levels greater than
        ' or equal to DWRITE_PAINT_FEATURE_LEVEL_COLR_V0.
        DWRITE_PAINT_TYPE_NONE
        DWRITE_PAINT_TYPE_LAYERS
        DWRITE_PAINT_TYPE_SOLID_GLYPH

        ' The following paint types may be returned for color feature levels greater than
        ' or equal to DWRITE_PAINT_FEATURE_LEVEL_COLR_V1.
        DWRITE_PAINT_TYPE_SOLID
        DWRITE_PAINT_TYPE_LINEAR_GRADIENT
        DWRITE_PAINT_TYPE_RADIAL_GRADIENT
        DWRITE_PAINT_TYPE_SWEEP_GRADIENT
        DWRITE_PAINT_TYPE_GLYPH
        DWRITE_PAINT_TYPE_COLOR_GLYPH
        DWRITE_PAINT_TYPE_TRANSFORM
        DWRITE_PAINT_TYPE_COMPOSITE
    End Enum

    <ComImport>
    <Guid("7d97dbf7-e085-42d4-81e3-6a883bded118")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteGlyphRunAnalysis
        <PreserveSig>
        Function GetAlphaTextureBounds(textureType As DWRITE_TEXTURE_TYPE, <Out> ByRef textureBounds As RECT) As HRESULT
        <PreserveSig>
        Function CreateAlphaTexture(textureType As DWRITE_TEXTURE_TYPE, textureBounds As RECT, <Out> ByRef alphaValues As IntPtr, bufferSize As Integer) As HRESULT
        <PreserveSig>
        Function GetAlphaBlendParams(renderingParams As IDWriteRenderingParams, <Out> ByRef blendGamma As Single, <Out> ByRef blendEnhancedContrast As Single, <Out> ByRef blendClearTypeLevel As Single) As HRESULT
    End Interface

    <ComImport>
    <Guid("5e5a32a3-8dff-4773-9ff6-0696eab77267")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteBitmapRenderTarget
        <PreserveSig>
        Function DrawGlyphRun(baselineOriginX As Single, baselineOriginY As Single, measuringMode As DWRITE_MEASURING_MODE, ByRef glyphRun As DWRITE_GLYPH_RUN, renderingParams As IDWriteRenderingParams, textColor As UInteger, <Out> ByRef blackBoxRect As RECT) As HRESULT
        <PreserveSig>
        Function GetMemoryDC() As IntPtr
        <PreserveSig>
        Function GetPixelsPerDip() As Single
        <PreserveSig>
        Function SetPixelsPerDip(pixelsPerDip As Single) As HRESULT
        <PreserveSig>
        Function GetCurrentTransform(<Out> ByRef transform As DWRITE_MATRIX) As HRESULT
        <PreserveSig>
        Function SetCurrentTransform(transform As DWRITE_MATRIX) As HRESULT
        <PreserveSig>
        Function GetSize(<Out> ByRef size As SIZE) As HRESULT
        <PreserveSig>
        Function Resize(width As Integer, height As Integer) As HRESULT
    End Interface

    <ComImport>
    <Guid("791e8298-3ef3-4230-9880-c9bdecc42064")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteBitmapRenderTarget1
        Inherits IDWriteBitmapRenderTarget
#Region "<IDWriteBitmapRenderTarget>"
        <PreserveSig>
        Overloads Function DrawGlyphRun(baselineOriginX As Single, baselineOriginY As Single, measuringMode As DWRITE_MEASURING_MODE, ByRef glyphRun As DWRITE_GLYPH_RUN, renderingParams As IDWriteRenderingParams, textColor As UInteger, <Out> ByRef blackBoxRect As RECT) As HRESULT
        <PreserveSig>
        Overloads Function GetMemoryDC() As IntPtr
        <PreserveSig>
        Overloads Function GetPixelsPerDip() As Single
        <PreserveSig>
        Overloads Function SetPixelsPerDip(pixelsPerDip As Single) As HRESULT
        <PreserveSig>
        Overloads Function GetCurrentTransform(<Out> ByRef transform As DWRITE_MATRIX) As HRESULT
        <PreserveSig>
        Overloads Function SetCurrentTransform(transform As DWRITE_MATRIX) As HRESULT
        <PreserveSig>
        Overloads Function GetSize(<Out> ByRef size As SIZE) As HRESULT
        <PreserveSig>
        Overloads Function Resize(width As Integer, height As Integer) As HRESULT
#End Region

        <PreserveSig>
        Function GetTextAntialiasMode() As DWRITE_TEXT_ANTIALIAS_MODE
        <PreserveSig>
        Function SetTextAntialiasMode(antialiasMode As DWRITE_TEXT_ANTIALIAS_MODE) As HRESULT
    End Interface

    <ComImport>
    <Guid("C553A742-FC01-44DA-A66E-B8B9ED6C3995")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteBitmapRenderTarget2
        Inherits IDWriteBitmapRenderTarget1
#Region "<IDWriteBitmapRenderTarget>"
        <PreserveSig>
        Overloads Function DrawGlyphRun(baselineOriginX As Single, baselineOriginY As Single, measuringMode As DWRITE_MEASURING_MODE, ByRef glyphRun As DWRITE_GLYPH_RUN, renderingParams As IDWriteRenderingParams, textColor As UInteger, <Out> ByRef blackBoxRect As RECT) As HRESULT
        <PreserveSig>
        Overloads Function GetMemoryDC() As IntPtr
        <PreserveSig>
        Overloads Function GetPixelsPerDip() As Single
        <PreserveSig>
        Overloads Function SetPixelsPerDip(pixelsPerDip As Single) As HRESULT
        <PreserveSig>
        Overloads Function GetCurrentTransform(<Out> ByRef transform As DWRITE_MATRIX) As HRESULT
        <PreserveSig>
        Overloads Function SetCurrentTransform(transform As DWRITE_MATRIX) As HRESULT
        <PreserveSig>
        Overloads Function GetSize(<Out> ByRef size As SIZE) As HRESULT
        <PreserveSig>
        Overloads Function Resize(width As Integer, height As Integer) As HRESULT
#End Region

#Region "<IDWriteBitmapRenderTarget1>"
        <PreserveSig>
        Overloads Function GetTextAntialiasMode() As DWRITE_TEXT_ANTIALIAS_MODE
        <PreserveSig>
        Overloads Function SetTextAntialiasMode(antialiasMode As DWRITE_TEXT_ANTIALIAS_MODE) As HRESULT
#End Region

        <PreserveSig>
        Function GetBitmapData(<Out> ByRef bitmapData As DWRITE_BITMAP_DATA_BGRA32) As HRESULT
    End Interface

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DWRITE_BITMAP_DATA_BGRA32
        Public width As UInteger
        Public height As UInteger
        Public pixels As IntPtr
    End Structure

    <ComImport>
    <Guid("AEEC37DB-C337-40F1-8E2A-9A41B167B238")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteBitmapRenderTarget3
        Inherits IDWriteBitmapRenderTarget2
#Region "<IDWriteBitmapRenderTarget>"
        <PreserveSig>
        Overloads Function DrawGlyphRun(baselineOriginX As Single, baselineOriginY As Single, measuringMode As DWRITE_MEASURING_MODE, ByRef glyphRun As DWRITE_GLYPH_RUN, renderingParams As IDWriteRenderingParams, textColor As UInteger, <Out> ByRef blackBoxRect As RECT) As HRESULT
        <PreserveSig>
        Overloads Function GetMemoryDC() As IntPtr
        <PreserveSig>
        Overloads Function GetPixelsPerDip() As Single
        <PreserveSig>
        Overloads Function SetPixelsPerDip(pixelsPerDip As Single) As HRESULT
        <PreserveSig>
        Overloads Function GetCurrentTransform(<Out> ByRef transform As DWRITE_MATRIX) As HRESULT
        <PreserveSig>
        Overloads Function SetCurrentTransform(transform As DWRITE_MATRIX) As HRESULT
        <PreserveSig>
        Overloads Function GetSize(<Out> ByRef size As SIZE) As HRESULT
        <PreserveSig>
        Overloads Function Resize(width As Integer, height As Integer) As HRESULT
#End Region

#Region "<IDWriteBitmapRenderTarget1>"
        <PreserveSig>
        Overloads Function GetTextAntialiasMode() As DWRITE_TEXT_ANTIALIAS_MODE
        <PreserveSig>
        Overloads Function SetTextAntialiasMode(antialiasMode As DWRITE_TEXT_ANTIALIAS_MODE) As HRESULT
#End Region

#Region "<IDWriteBitmapRenderTarget2>"
        <PreserveSig>
        Overloads Function GetBitmapData(<Out> ByRef bitmapData As DWRITE_BITMAP_DATA_BGRA32) As HRESULT
#End Region

        <PreserveSig>
        Function GetPaintFeatureLevel() As DWRITE_PAINT_FEATURE_LEVEL
        <PreserveSig>
        Function DrawPaintGlyphRun(baselineOriginX As Single, baselineOriginY As Single, measuringMode As DWRITE_MEASURING_MODE, ByRef glyphRun As DWRITE_GLYPH_RUN, glyphImageFormat As DWRITE_GLYPH_IMAGE_FORMATS, textColor As UInteger, colorPaletteIndex As UInteger, <Out> ByRef blackBoxRect As RECT) As HRESULT
        <PreserveSig>
        Function DrawGlyphRunWithColorSupport(baselineOriginX As Single, baselineOriginY As Single, measuringMode As DWRITE_MEASURING_MODE, ByRef glyphRun As DWRITE_GLYPH_RUN, renderingParams As IDWriteRenderingParams, textColor As UInteger, colorPaletteIndex As UInteger, <Out> ByRef blackBoxRect As RECT) As HRESULT
    End Interface

    Public Enum DWRITE_PAINT_FEATURE_LEVEL As Integer
        ''' <summary>
        ''' No paint API support.
        ''' </summary>
        DWRITE_PAINT_FEATURE_LEVEL_NONE = 0

        ''' <summary>
        ''' Specifies a level of functionality corresponding to OpenType COLR version 0.
        ''' </summary>
        DWRITE_PAINT_FEATURE_LEVEL_COLR_V0 = 1

        ''' <summary>
        ''' Specifies a level of functionality corresponding to OpenType COLR version 1.
        ''' </summary>
        DWRITE_PAINT_FEATURE_LEVEL_COLR_V1 = 2
    End Enum

    <ComImport>
    <Guid("5810cd44-0ca0-4701-b3fa-bec5182ae4f6")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteTextAnalysisSink
        Function SetScriptAnalysis(textPosition As Integer, textLength As Integer, scriptAnalysis As DWRITE_SCRIPT_ANALYSIS) As HRESULT
        Function SetLineBreakpoints(textPosition As Integer, textLength As Integer, lineBreakpoints As DWRITE_LINE_BREAKPOINT) As HRESULT
        Function SetBidiLevel(textPosition As Integer, textLength As Integer, explicitLevel As Byte, resolvedLevel As Byte) As HRESULT
        Function SetNumberSubstitution(textPosition As Integer, textLength As Integer, numberSubstitution As IDWriteNumberSubstitution) As HRESULT
    End Interface

    <ComImport>
    <Guid("B0D941A0-85E7-4D8B-9FD3-5CED9934482A")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteTextAnalysisSink1
        Inherits IDWriteTextAnalysisSink
#Region "IDWriteTextAnalysisSink"
        Overloads Function SetScriptAnalysis(textPosition As Integer, textLength As Integer, scriptAnalysis As DWRITE_SCRIPT_ANALYSIS) As HRESULT
        Overloads Function SetLineBreakpoints(textPosition As Integer, textLength As Integer, lineBreakpoints As DWRITE_LINE_BREAKPOINT) As HRESULT
        Overloads Function SetBidiLevel(textPosition As Integer, textLength As Integer, explicitLevel As Byte, resolvedLevel As Byte) As HRESULT
        Overloads Function SetNumberSubstitution(textPosition As Integer, textLength As Integer, numberSubstitution As IDWriteNumberSubstitution) As HRESULT
#End Region

        Function SetGlyphOrientation(textPosition As UInteger, textLength As UInteger, glyphOrientationAngle As DWRITE_GLYPH_ORIENTATION_ANGLE, adjustedBidiLevel As Byte, isSideways As Boolean, isRightToLeft As Boolean) As HRESULT
    End Interface

    <ComImport>
    <Guid("ef8a8135-5cc6-45fe-8825-c5a0724eb819")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteTextRenderer
        Inherits IDWritePixelSnapping
#Region "IDWritePixelSnapping"
        <PreserveSig>
        Overloads Function IsPixelSnappingDisabled(clientDrawingContext As IntPtr, <Out> ByRef isDisabled As Boolean) As HRESULT
        <PreserveSig>
        Overloads Function GetCurrentTransform(clientDrawingContext As IntPtr, <Out> ByRef transform As DWRITE_MATRIX) As HRESULT
        <PreserveSig>
        Overloads Function GetPixelsPerDip(clientDrawingContext As IntPtr, <Out>
        <MarshalAs(UnmanagedType.R4)> ByRef pixelsPerDip As Single) As HRESULT
#End Region

        <PreserveSig>
        Function DrawGlyphRun(clientDrawingContext As IntPtr, baselineOriginX As Single, baselineOriginY As Single, measuringMode As DWRITE_MEASURING_MODE, ByRef glyphRun As DWRITE_GLYPH_RUN, glyphRunDescription As IntPtr, clientDrawingEffect As IntPtr) As HRESULT
        <PreserveSig>
        Function DrawUnderline(clientDrawingContext As IntPtr, baselineOriginX As Single, baselineOriginY As Single, underline As DWRITE_UNDERLINE, clientDrawingEffect As IntPtr) As HRESULT
        <PreserveSig>
        Function DrawStrikethrough(clientDrawingContext As IntPtr, baselineOriginX As Single, baselineOriginY As Single, strikethrough As DWRITE_STRIKETHROUGH, clientDrawingEffect As IntPtr) As HRESULT
        <PreserveSig>
        Function DrawInlineObject(clientDrawingContext As IntPtr, originX As Single, originY As Single, inlineObject As IDWriteInlineObject, isSideways As Boolean, isRightToLeft As Boolean, clientDrawingEffect As IntPtr) As HRESULT
    End Interface

    <ComImport>
    <Guid("D3E0E934-22A0-427E-AAE4-7D9574B59DB1")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteTextRenderer1
        Inherits IDWriteTextRenderer
#Region "IDWriteTextRenderer"
#Region "IDWritePixelSnapping"
        <PreserveSig>
        Overloads Function IsPixelSnappingDisabled(clientDrawingContext As IntPtr, <Out> ByRef isDisabled As Boolean) As HRESULT
        <PreserveSig>
        Overloads Function GetCurrentTransform(clientDrawingContext As IntPtr, <Out> ByRef transform As DWRITE_MATRIX) As HRESULT
        <PreserveSig>
        Overloads Function GetPixelsPerDip(clientDrawingContext As IntPtr, <Out>
        <MarshalAs(UnmanagedType.R4)> ByRef pixelsPerDip As Single) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function DrawGlyphRun(clientDrawingContext As IntPtr, baselineOriginX As Single, baselineOriginY As Single, measuringMode As DWRITE_MEASURING_MODE, ByRef glyphRun As DWRITE_GLYPH_RUN, glyphRunDescription As IntPtr, clientDrawingEffect As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function DrawUnderline(clientDrawingContext As IntPtr, baselineOriginX As Single, baselineOriginY As Single, underline As DWRITE_UNDERLINE, clientDrawingEffect As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function DrawStrikethrough(clientDrawingContext As IntPtr, baselineOriginX As Single, baselineOriginY As Single, strikethrough As DWRITE_STRIKETHROUGH, clientDrawingEffect As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function DrawInlineObject(clientDrawingContext As IntPtr, originX As Single, originY As Single, inlineObject As IDWriteInlineObject, isSideways As Boolean, isRightToLeft As Boolean, clientDrawingEffect As IntPtr) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function DrawGlyphRun(clientDrawingContext As IntPtr, baselineOriginX As Single, baselineOriginY As Single, orientationAngle As DWRITE_GLYPH_ORIENTATION_ANGLE, measuringMode As DWRITE_MEASURING_MODE, ByRef glyphRun As DWRITE_GLYPH_RUN, glyphRunDescription As IntPtr, clientDrawingEffect As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function DrawUnderline(clientDrawingContext As IntPtr, baselineOriginX As Single, baselineOriginY As Single, orientationAngle As DWRITE_GLYPH_ORIENTATION_ANGLE, underline As DWRITE_UNDERLINE, clientDrawingEffect As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function DrawStrikethrough(clientDrawingContext As IntPtr, baselineOriginX As Single, baselineOriginY As Single, orientationAngle As DWRITE_GLYPH_ORIENTATION_ANGLE, strikethrough As DWRITE_STRIKETHROUGH, clientDrawingEffect As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function DrawInlineObject(clientDrawingContext As IntPtr, originX As Single, originY As Single, orientationAngle As DWRITE_GLYPH_ORIENTATION_ANGLE, inlineObject As IDWriteInlineObject, isSideways As Boolean, isRightToLeft As Boolean, clientDrawingEffect As IntPtr) As HRESULT
    End Interface

    <ComImport>
    <Guid("eaf3a2da-ecf4-4d24-b644-b34f6842024b")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWritePixelSnapping
        <PreserveSig>
        Function IsPixelSnappingDisabled(clientDrawingContext As IntPtr, <Out> ByRef isDisabled As Boolean) As HRESULT
        <PreserveSig>
        Function GetCurrentTransform(clientDrawingContext As IntPtr, <Out> ByRef transform As DWRITE_MATRIX) As HRESULT
        <PreserveSig>
        Function GetPixelsPerDip(clientDrawingContext As IntPtr, <Out>
        <MarshalAs(UnmanagedType.R4)> ByRef pixelsPerDip As Single) As HRESULT
    End Interface

    <ComImport>
    <Guid("2cd9069e-12e2-11dc-9fed-001143a055f9")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1SimplifiedGeometrySink
        Function SetFillMode(fillMode As D2D1_FILL_MODE) As HRESULT
        Function SetSegmentFlags(vertexFlags As D2D1_PATH_SEGMENT) As HRESULT
        Function BeginFigure(startPoint As D2D1_POINT_2F, figureBegin As D2D1_FIGURE_BEGIN) As HRESULT
        Function AddLines(
<MarshalAs(UnmanagedType.LPArray)> points As D2D1_POINT_2F(), pointsCount As UInteger) As HRESULT
        Function AddBeziers(beziers As D2D1_BEZIER_SEGMENT, beziersCount As UInteger) As HRESULT
        Function EndFigure(figureEnd As D2D1_FIGURE_END) As HRESULT
        Function Close() As HRESULT
    End Interface

    Public Enum DWRITE_TEXTURE_TYPE
        ''' <summary>
        ''' Specifies an alpha texture for aliased text rendering (i.e., bi-level, where each pixel is either fully opaque or fully transparent),
        ''' with one byte per pixel.
        ''' </summary>
        DWRITE_TEXTURE_ALIASED_1x1
        ''' <summary>
        ''' Specifies an alpha texture for ClearType text rendering, with three bytes per pixel in the horizontal dimension and 
        ''' one byte per pixel in the vertical dimension.
        ''' </summary>
        DWRITE_TEXTURE_CLEARTYPE_3x1
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DWRITE_TYPOGRAPHIC_FEATURES
        ''' <summary>
        ''' Array of font features.
        ''' </summary>
        Public features As DWRITE_FONT_FEATURE

        ''' <summary>
        ''' The number of features.
        ''' </summary>
        Public featureCount As Integer
    End Structure

    Public Enum DWRITE_RENDERING_MODE
        ''' <summary>
        ''' Specifies that the rendering mode is determined automatically based on the font and size.
        ''' </summary>
        DWRITE_RENDERING_MODE_DEFAULT
        ''' <summary>
        ''' Specifies that no antialiasing is performed. Each pixel is either set to the foreground 
        ''' color of the text or retains the color of the background.
        ''' </summary>
        DWRITE_RENDERING_MODE_ALIASED
        ''' <summary>
        ''' Specifies that antialiasing is performed in the horizontal direction and the appearance
        ''' of glyphs is layout-compatible with GDI using CLEARTYPE_QUALITY. Use DWRITE_MEASURING_MODE_GDI_CLASSIC 
        ''' to get glyph advances. The antialiasing may be either ClearType or grayscale depending on
        ''' the text antialiasing mode.
        ''' </summary>
        DWRITE_RENDERING_MODE_GDI_CLASSIC
        ''' <summary>
        ''' Specifies that antialiasing is performed in the horizontal direction and the appearance
        ''' of glyphs is layout-compatible with GDI using CLEARTYPE_NATURAL_QUALITY. Glyph advances
        ''' are close to the font design advances, but are still rounded to whole pixels. Use
        ''' DWRITE_MEASURING_MODE_GDI_NATURAL to get glyph advances. The antialiasing may be either
        ''' ClearType or grayscale depending on the text antialiasing mode.
        ''' </summary>
        DWRITE_RENDERING_MODE_GDI_NATURAL
        ''' <summary>
        ''' Specifies that antialiasing is performed in the horizontal direction. This rendering
        ''' mode allows glyphs to be positioned with subpixel precision and is therefore suitable
        ''' for natural (i.e., resolution-independent) layout. The antialiasing may be either
        ''' ClearType or grayscale depending on the text antialiasing mode.
        ''' </summary>
        DWRITE_RENDERING_MODE_NATURAL
        ''' <summary>
        ''' Similar to natural mode except that antialiasing is performed in both the horizontal
        ''' and vertical directions. This is typically used at larger sizes to make curves and
        ''' diagonal lines look smoother. The antialiasing may be either ClearType or grayscale
        ''' depending on the text antialiasing mode.
        ''' </summary>
        DWRITE_RENDERING_MODE_NATURAL_SYMMETRIC
        ''' <summary>
        ''' Specifies that rendering should bypass the rasterizer and use the outlines directly. 
        ''' This is typically used at very large sizes.
        ''' </summary>
        DWRITE_RENDERING_MODE_OUTLINE
        ' The following names are obsolete, but are kept as aliases to avoid breaking existing code.
        ' Each of these rendering modes may result in either ClearType or grayscale antialiasing 
        ' depending on the DWRITE_TEXT_ANTIALIASING_MODE.
        DWRITE_RENDERING_MODE_CLEARTYPE_GDI_CLASSIC = DWRITE_RENDERING_MODE_GDI_CLASSIC
        DWRITE_RENDERING_MODE_CLEARTYPE_GDI_NATURAL = DWRITE_RENDERING_MODE_GDI_NATURAL
        DWRITE_RENDERING_MODE_CLEARTYPE_NATURAL = DWRITE_RENDERING_MODE_NATURAL
        DWRITE_RENDERING_MODE_CLEARTYPE_NATURAL_SYMMETRIC = DWRITE_RENDERING_MODE_NATURAL_SYMMETRIC
    End Enum

    Public Enum DWRITE_PIXEL_GEOMETRY
        ''' <summary>
        ''' The red, green, and blue color components of each pixel are assumed to occupy the same point.
        ''' </summary>
        DWRITE_PIXEL_GEOMETRY_FLAT
        ''' <summary>
        ''' Each pixel comprises three vertical stripes, with red on the left, green in the center, and 
        ''' blue on the right. This is the most common pixel geometry for LCD monitors.
        ''' </summary>
        DWRITE_PIXEL_GEOMETRY_RGB
        ''' <summary>
        ''' Each pixel comprises three vertical stripes, with blue on the left, green in the center, and 
        ''' red on the right.
        ''' </summary>
        DWRITE_PIXEL_GEOMETRY_BGR
    End Enum

    Public Enum DWRITE_FONT_FILE_TYPE
        ''' <summary>
        ''' Font type is not recognized by the DirectWrite font system.
        ''' </summary>
        DWRITE_FONT_FILE_TYPE_UNKNOWN
        ''' <summary>
        ''' OpenType font with CFF outlines.
        ''' </summary>
        DWRITE_FONT_FILE_TYPE_CFF
        ''' <summary>
        ''' OpenType font with TrueType outlines.
        ''' </summary>
        DWRITE_FONT_FILE_TYPE_TRUETYPE
        ''' <summary>
        ''' OpenType font that contains a TrueType collection.
        ''' </summary>
        DWRITE_FONT_FILE_TYPE_OPENTYPE_COLLECTION
        ''' <summary>
        ''' Type 1 PFM font.
        ''' </summary>
        DWRITE_FONT_FILE_TYPE_TYPE1_PFM
        ''' <summary>
        ''' Type 1 PFB font.
        ''' </summary>
        DWRITE_FONT_FILE_TYPE_TYPE1_PFB
        ''' <summary>
        ''' Vector .FON font.
        ''' </summary>
        DWRITE_FONT_FILE_TYPE_VECTOR
        ''' <summary>
        ''' Bitmap .FON font.
        ''' </summary>
        DWRITE_FONT_FILE_TYPE_BITMAP
        ' The following name is obsolete, but kept as an alias to avoid breaking existing code.
        DWRITE_FONT_FILE_TYPE_TRUETYPE_COLLECTION = DWRITE_FONT_FILE_TYPE_OPENTYPE_COLLECTION
    End Enum

    Public Enum DWRITE_FONT_FACE_TYPE
        ''' <summary>
        ''' OpenType font face with CFF outlines.
        ''' </summary>
        DWRITE_FONT_FACE_TYPE_CFF
        ''' <summary>
        ''' OpenType font face with TrueType outlines.
        ''' </summary>
        DWRITE_FONT_FACE_TYPE_TRUETYPE
        ''' <summary>
        ''' OpenType font face that is a part of a TrueType or CFF collection.
        ''' </summary>
        DWRITE_FONT_FACE_TYPE_OPENTYPE_COLLECTION
        ''' <summary>
        ''' A Type 1 font face.
        ''' </summary>
        DWRITE_FONT_FACE_TYPE_TYPE1
        ''' <summary>
        ''' A vector .FON format font face.
        ''' </summary>
        DWRITE_FONT_FACE_TYPE_VECTOR
        ''' <summary>
        ''' A bitmap .FON format font face.
        ''' </summary>
        DWRITE_FONT_FACE_TYPE_BITMAP
        ''' <summary>
        ''' Font face type is not recognized by the DirectWrite font system.
        ''' </summary>
        DWRITE_FONT_FACE_TYPE_UNKNOWN
        ''' <summary>
        ''' The font data includes only the CFF table from an OpenType CFF font.
        ''' This font face type can be used only for embedded fonts (i.e., custom
        ''' font file loaders) and the resulting font face object supports only the
        ''' minimum functionality necessary to render glyphs.
        ''' </summary>
        DWRITE_FONT_FACE_TYPE_RAW_CFF
        ' The following name is obsolete, but kept as an alias to avoid breaking existing code.
        DWRITE_FONT_FACE_TYPE_TRUETYPE_COLLECTION = DWRITE_FONT_FACE_TYPE_OPENTYPE_COLLECTION
    End Enum

    Public Enum DWRITE_FONT_SIMULATIONS
        ''' <summary>
        ''' No simulations are performed.
        ''' </summary>
        DWRITE_FONT_SIMULATIONS_NONE = &H0000
        ''' <summary>
        ''' Algorithmic emboldening is performed.
        ''' </summary>
        DWRITE_FONT_SIMULATIONS_BOLD = &H0001
        ''' <summary>
        ''' Algorithmic italicization is performed.
        ''' </summary>
        DWRITE_FONT_SIMULATIONS_OBLIQUE = &H0002
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DWRITE_FONT_METRICS
        ''' <summary>
        ''' The number of font design units per em unit.
        ''' Font files use their own coordinate system of font design units.
        ''' A font design unit is the smallest measurable unit in the em square,
        ''' an imaginary square that is used to size and align glyphs.
        ''' The concept of em square is used as a reference scale factor when defining font size and device transformation semantics.
        ''' The size of one em square is also commonly used to compute the paragraph indentation value.
        ''' </summary>
        Public designUnitsPerEm As UShort
        ''' <summary>
        ''' Ascent value of the font face in font design units.
        ''' Ascent is the distance from the top of font character alignment box to English baseline.
        ''' </summary>
        Public ascent As UShort
        ''' <summary>
        ''' Descent value of the font face in font design units.
        ''' Descent is the distance from the bottom of font character alignment box to English baseline.
        ''' </summary>
        Public descent As UShort
        ''' <summary>
        ''' Line gap in font design units.
        ''' Recommended additional white space to add between lines to improve legibility. The recommended line spacing 
        ''' (baseline-to-baseline distance) is thus the sum of ascent, descent, and lineGap. The line gap is usually 
        ''' positive or zero but can be negative, in which case the recommended line spacing is less than the height
        ''' of the character alignment box.
        ''' </summary>
        Public lineGap As UShort
        ''' <summary>
        ''' Cap height value of the font face in font design units.
        ''' Cap height is the distance from English baseline to the top of a typical English capital.
        ''' Capital "H" is often used as a reference character for the purpose of calculating the cap height value.
        ''' </summary>
        Public capHeight As UShort
        ''' <summary>
        ''' x-height value of the font face in font design units.
        ''' x-height is the distance from English baseline to the top of lowercase letter "x", or a similar lowercase character.
        ''' </summary>
        Public xHeight As UShort
        ''' <summary>
        ''' The underline position value of the font face in font design units.
        ''' Underline position is the position of underline relative to the English baseline.
        ''' The value is usually made negative in order to place the underline below the baseline.
        ''' </summary>
        Public underlinePosition As UShort
        ''' <summary>
        ''' The suggested underline thickness value of the font face in font design units.
        ''' </summary>
        Public underlineThickness As UShort
        ''' <summary>
        ''' The strikethrough position value of the font face in font design units.
        ''' Strikethrough position is the position of strikethrough relative to the English baseline.
        ''' The value is usually made positive in order to place the strikethrough above the baseline.
        ''' </summary>
        Public strikethroughPosition As UShort
        ''' <summary>
        ''' The suggested strikethrough thickness value of the font face in font design units.
        ''' </summary>
        Public strikethroughThickness As UShort
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DWRITE_GLYPH_METRICS
        ''' <summary>
        ''' Specifies the X offset from the glyph origin to the left edge of the black box.
        ''' The glyph origin is the current horizontal writing position.
        ''' A negative value means the black box extends to the left of the origin (often true for lowercase italic 'f').
        ''' </summary>
        Public leftSideBearing As Integer
        ''' <summary>
        ''' Specifies the X offset from the origin of the current glyph to the origin of the next glyph when writing horizontally.
        ''' </summary>
        Public advanceWidth As Integer
        ''' <summary>
        ''' Specifies the X offset from the right edge of the black box to the origin of the next glyph when writing horizontally.
        ''' The value is negative when the right edge of the black box overhangs the layout box.
        ''' </summary>
        Public rightSideBearing As Integer
        ''' <summary>
        ''' Specifies the vertical offset from the vertical origin to the top of the black box.
        ''' Thus, a positive value adds whitespace whereas a negative value means the glyph overhangs the top of the layout box.
        ''' </summary>
        Public topSideBearing As Integer
        ''' <summary>
        ''' Specifies the Y offset from the vertical origin of the current glyph to the vertical origin of the next glyph when writing vertically.
        ''' (Note that the term "origin" by itself denotes the horizontal origin. The vertical origin is different.
        ''' Its Y coordinate is specified by verticalOriginY value,
        ''' and its X coordinate is half the advanceWidth to the right of the horizontal origin).
        ''' </summary>
        Public advanceHeight As Integer
        ''' <summary>
        ''' Specifies the vertical distance from the black box's bottom edge to the advance height.
        ''' Positive when the bottom edge of the black box is within the layout box.
        ''' Negative when the bottom edge of black box overhangs the layout box.
        ''' </summary>
        Public bottomSideBearing As Integer
        ''' <summary>
        ''' Specifies the Y coordinate of a glyph's vertical origin, in the font's design coordinate system.
        ''' The y coordinate of a glyph's vertical origin is the sum of the glyph's top side bearing
        ''' and the top (i.e. yMax) of the glyph's bounding box.
        ''' </summary>
        Public verticalOriginY As Integer
    End Structure

    Public Enum DWRITE_TEXT_ALIGNMENT
        ''' <summary>
        ''' The leading edge of the paragraph text is aligned to the layout box's leading edge.
        ''' </summary>
        DWRITE_TEXT_ALIGNMENT_LEADING
        ''' <summary>
        ''' The trailing edge of the paragraph text is aligned to the layout box's trailing edge.
        ''' </summary>
        DWRITE_TEXT_ALIGNMENT_TRAILING
        ''' <summary>
        ''' The center of the paragraph text is aligned to the center of the layout box.
        ''' </summary>
        DWRITE_TEXT_ALIGNMENT_CENTER
        ''' <summary>
        ''' Align text to the leading side, and also justify text to fill the lines.
        ''' </summary>
        DWRITE_TEXT_ALIGNMENT_JUSTIFIED
    End Enum

    Public Enum DWRITE_PARAGRAPH_ALIGNMENT
        ''' <summary>
        ''' The first line of paragraph is aligned to the flow's beginning edge of the layout box.
        ''' </summary>
        DWRITE_PARAGRAPH_ALIGNMENT_NEAR
        ''' <summary>
        ''' The last line of paragraph is aligned to the flow's ending edge of the layout box.
        ''' </summary>
        DWRITE_PARAGRAPH_ALIGNMENT_FAR
        ''' <summary>
        ''' The center of the paragraph is aligned to the center of the flow of the layout box.
        ''' </summary>
        DWRITE_PARAGRAPH_ALIGNMENT_CENTER
    End Enum

    Public Enum DWRITE_WORD_WRAPPING
        ''' <summary>
        ''' Words are broken across lines to avoid text overflowing the layout box.
        ''' </summary>
        DWRITE_WORD_WRAPPING_WRAP = 0
        ''' <summary>
        ''' Words are kept within the same line even when it overflows the layout box.
        ''' This option is often used with scrolling to reveal overflow text. 
        ''' </summary>
        DWRITE_WORD_WRAPPING_NO_WRAP = 1
        ''' <summary>
        ''' Words are broken across lines to avoid text overflowing the layout box.
        ''' Emergency wrapping occurs if the word is larger than the maximum width.
        ''' </summary>
        DWRITE_WORD_WRAPPING_EMERGENCY_BREAK = 2
        ''' <summary>
        ''' Only wrap whole words, never breaking words (emergency wrapping) when the
        ''' layout width is too small for even a single word.
        ''' </summary>
        DWRITE_WORD_WRAPPING_WHOLE_WORD = 3
        ''' <summary>
        ''' Wrap between any valid characters clusters.
        ''' </summary>
        DWRITE_WORD_WRAPPING_CHARACTER = 4
    End Enum

    Public Enum DWRITE_LINE_SPACING_METHOD
        ''' <summary>
        ''' Line spacing depends solely on the content, growing to accommodate the size of fonts and inline objects.
        ''' </summary>
        DWRITE_LINE_SPACING_METHOD_DEFAULT
        ''' <summary>
        ''' Lines are explicitly set to uniform spacing, regardless of contained font sizes.
        ''' This can be useful to avoid the uneven appearance that can occur from font fallback.
        ''' </summary>
        DWRITE_LINE_SPACING_METHOD_UNIFORM
        ''' <summary>
        ''' Line spacing and baseline distances are proportional to the computed values based on the content, the size of the fonts and inline objects.
        ''' </summary>
        DWRITE_LINE_SPACING_METHOD_PROPORTIONAL
    End Enum

    Public Enum DWRITE_READING_DIRECTION
        ''' <summary>
        ''' Reading progresses from left to right.
        ''' </summary>
        DWRITE_READING_DIRECTION_LEFT_TO_RIGHT = 0
        ''' <summary>
        ''' Reading progresses from right to left.
        ''' </summary>
        DWRITE_READING_DIRECTION_RIGHT_TO_LEFT = 1
        ''' <summary>
        ''' Reading progresses from top to bottom.
        ''' </summary>
        DWRITE_READING_DIRECTION_TOP_TO_BOTTOM = 2
        ''' <summary>
        ''' Reading progresses from bottom to top.
        ''' </summary>
        DWRITE_READING_DIRECTION_BOTTOM_TO_TOP = 3
    End Enum

    Public Enum DWRITE_FLOW_DIRECTION
        ''' <summary>
        ''' Text lines are placed from top to bottom.
        ''' </summary>
        DWRITE_FLOW_DIRECTION_TOP_TO_BOTTOM = 0
        ''' <summary>
        ''' Text lines are placed from bottom to top.
        ''' </summary>
        DWRITE_FLOW_DIRECTION_BOTTOM_TO_TOP = 1
        ''' <summary>
        ''' Text lines are placed from left to right.
        ''' </summary>
        DWRITE_FLOW_DIRECTION_LEFT_TO_RIGHT = 2
        ''' <summary>
        ''' Text lines are placed from right to left.
        ''' </summary>
        DWRITE_FLOW_DIRECTION_RIGHT_TO_LEFT = 3
    End Enum

    Public Enum DWRITE_TRIMMING_GRANULARITY
        ''' <summary>
        ''' No trimming occurs. Text flows beyond the layout width.
        ''' </summary>
        DWRITE_TRIMMING_GRANULARITY_NONE
        ''' <summary>
        ''' Trimming occurs at character cluster boundary.
        ''' </summary>
        DWRITE_TRIMMING_GRANULARITY_CHARACTER
        ''' <summary>
        ''' Trimming occurs at word boundary.
        ''' </summary>
        DWRITE_TRIMMING_GRANULARITY_WORD
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DWRITE_TRIMMING
        ''' <summary>
        ''' Text granularity of which trimming applies.
        ''' </summary>
        Public granularity As DWRITE_TRIMMING_GRANULARITY
        ''' <summary>
        ''' Character code used as the delimiter signaling the beginning of the portion of text to be preserved,
        ''' most useful for path ellipsis, where the delimiter would be a slash. Leave this zero if there is no
        ''' delimiter.
        ''' </summary>
        Public delimiter As UInteger
        ''' <summary>
        ''' How many occurrences of the delimiter to step back. Leave this zero if there is no delimiter.
        ''' </summary>
        Public delimiterCount As UInteger
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DWRITE_INLINE_OBJECT_METRICS
        ''' <summary>
        ''' Width of the inline object.
        ''' </summary>
        Public width As Single
        ''' <summary>
        ''' Height of the inline object as measured from top to bottom.
        ''' </summary>
        Public height As Single
        ''' <summary>
        ''' Distance from the top of the object to the baseline where it is lined up with the adjacent text.
        ''' If the baseline is at the bottom, baseline simply equals height.
        ''' </summary>
        Public baseline As Single
        ''' <summary>
        ''' Flag indicating whether the object is to be placed upright or alongside the text baseline
        ''' for vertical text.
        ''' </summary>
        Public supportsSideways As Boolean
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DWRITE_OVERHANG_METRICS
        ''' <summary>
        ''' The distance from the left-most visible DIP to its left alignment edge.
        ''' </summary>
        Public left As Single
        ''' <summary>
        ''' The distance from the top-most visible DIP to its top alignment edge.
        ''' </summary>
        Public top As Single
        ''' <summary>
        ''' The distance from the right-most visible DIP to its right alignment edge.
        ''' </summary>
        Public right As Single
        ''' <summary>
        ''' The distance from the bottom-most visible DIP to its bottom alignment edge.
        ''' </summary>
        Public bottom As Single
    End Structure

    Public Enum DWRITE_BREAK_CONDITION
        ''' <summary>
        ''' Whether a break is allowed is determined by the condition of the
        ''' neighboring text span or inline object.
        ''' </summary>
        DWRITE_BREAK_CONDITION_NEUTRAL
        ''' <summary>
        ''' A break is allowed, unless overruled by the condition of the
        ''' neighboring text span or inline object, either prohibited by a
        ''' May Not or forced by a Must.
        ''' </summary>
        DWRITE_BREAK_CONDITION_CAN_BREAK
        ''' <summary>
        ''' There should be no break, unless overruled by a Must condition from
        ''' the neighboring text span or inline object.
        ''' </summary>
        DWRITE_BREAK_CONDITION_MAY_NOT_BREAK
        ''' <summary>
        ''' The break must happen, regardless of the condition of the adjacent
        ''' text span or inline object.
        ''' </summary>
        DWRITE_BREAK_CONDITION_MUST_BREAK
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DWRITE_GLYPH_OFFSET
        ''' <summary>
        ''' Offset in the advance direction of the run. A positive advance offset moves the glyph to the right
        ''' (in pre-transform coordinates) if the run is left-to-right or to the left if the run is right-to-left.
        ''' </summary>
        Public advanceOffset As Single
        ''' <summary>
        ''' Offset in the ascent direction, i.e., the direction ascenders point. A positive ascender offset moves
        ''' the glyph up (in pre-transform coordinates).
        ''' </summary>
        Public ascenderOffset As Single
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DWRITE_GLYPH_RUN
        ''' <summary>
        ''' The physical font face to draw with.
        ''' </summary>
        Public fontFace As IntPtr
        ''' <summary>
        ''' Logical size of the font in DIPs, not points (equals 1/96 inch).
        ''' </summary>
        Public fontEmSize As Single
        ''' <summary>
        ''' The number of glyphs.
        ''' </summary>
        Public glyphCount As UInteger
        ''' <summary>
        ''' The indices to render.
        ''' </summary>
        Public glyphIndices As IntPtr
        ''' <summary>
        ''' Glyph advance widths.
        ''' </summary>
        Public glyphAdvances As IntPtr
        ''' <summary>
        ''' Glyph offsets.
        ''' </summary>
        Public glyphOffsets As IntPtr
        ''' <summary>
        ''' If true, specifies that glyphs are rotated 90 degrees to the left and vertical metrics are used. Vertical writing is achieved by specifying isSideways = true and rotating the entire run 90 degrees to the right via a rotate transform.
        ''' </summary>
        Public isSideways As Boolean
        ''' <summary>
        ''' The implicit resolved bidi level of the run. Odd levels indicate right-to-left languages like Hebrew and Arabic, while even levels indicate left-to-right languages like English and Japanese (when written horizontally). For right-to-left languages, the text origin is on the right, and text should be drawn to the left.
        ''' </summary>
        Public bidiLevel As UInteger
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DWRITE_GLYPH_RUN_DESCRIPTION
        ''' <summary>
        ''' The locale name associated with this run.
        ''' </summary>
        Public localeName As String
        ''' <summary>
        ''' The text associated with the glyphs.
        ''' </summary>
        Public str As String
        ''' <summary>
        ''' The number of characters (UTF16 code-units).
        ''' Note that this may be different than the number of glyphs.
        ''' </summary>
        Public stringLength As UInteger

        Public clusterMap As UShort

        ''' <summary>
        ''' Corresponding text position in the original string
        ''' this glyph run came from.
        ''' </summary>
        Public textPosition As UInteger
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DWRITE_UNDERLINE
        ''' <summary>
        ''' Width of the underline, measured parallel to the baseline.
        ''' </summary>
        Public width As Single
        ''' <summary>
        ''' Thickness of the underline, measured perpendicular to the
        ''' baseline.
        ''' </summary>
        Public thickness As Single
        ''' <summary>
        ''' Offset of the underline from the baseline.
        ''' A positive offset represents a position below the baseline and
        ''' a negative offset is above.
        ''' </summary>
        Public offset As Single
        ''' <summary>
        ''' Height of the tallest run where the underline applies.
        ''' </summary>
        Public runHeight As Single
        ''' <summary>
        ''' Reading direction of the text associated with the underline.  This 
        ''' value is used to interpret whether the width value runs horizontally 
        ''' or vertically.
        ''' </summary>
        Public readingDirection As DWRITE_READING_DIRECTION
        ''' <summary>
        ''' Flow direction of the text associated with the underline.  This value
        ''' is used to interpret whether the thickness value advances top to 
        ''' bottom, left to right, or right to left.
        ''' </summary>
        Public flowDirection As DWRITE_FLOW_DIRECTION
        ''' <summary>
        ''' Locale of the text the underline is being drawn under. Can be
        ''' pertinent where the locale affects how the underline is drawn.
        ''' For example, in vertical text, the underline belongs on the
        ''' left for Chinese but on the right for Japanese.
        ''' This choice is completely left up to higher levels.
        ''' </summary>
        Public localeName As String
        ''' <summary>
        ''' The measuring mode can be useful to the renderer to determine how
        ''' underlines are rendered, e.g. rounding the thickness to a whole pixel
        ''' in GDI-compatible modes.
        ''' </summary>
        Public measuringMode As DWRITE_MEASURING_MODE
    End Structure

    Public Enum DWRITE_MEASURING_MODE
        '
        ' Text is measured using glyph ideal metrics whose values are independent to the current display resolution.
        '
        DWRITE_MEASURING_MODE_NATURAL
        '
        ' Text is measured using glyph display compatible metrics whose values tuned for the current display resolution.
        '
        DWRITE_MEASURING_MODE_GDI_CLASSIC
        '
        ' Text is measured using the same glyph display metrics as text measured by GDI using a font
        ' created with CLEARTYPE_NATURAL_QUALITY.
        '
        DWRITE_MEASURING_MODE_GDI_NATURAL
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DWRITE_STRIKETHROUGH
        ''' <summary>
        ''' Width of the strikethrough, measured parallel to the baseline.
        ''' </summary>
        Public width As Single
        ''' <summary>
        ''' Thickness of the strikethrough, measured perpendicular to the
        ''' baseline.
        ''' </summary>
        Public thickness As Single
        ''' <summary>
        ''' Offset of the strikethrough from the baseline.
        ''' A positive offset represents a position below the baseline and
        ''' a negative offset is above.
        ''' </summary>
        Public offset As Single
        ''' <summary>
        ''' Reading direction of the text associated with the strikethrough.  This
        ''' value is used to interpret whether the width value runs horizontally 
        ''' or vertically.
        ''' </summary>
        Public readingDirection As DWRITE_READING_DIRECTION
        ''' <summary>
        ''' Flow direction of the text associated with the strikethrough.  This 
        ''' value is used to interpret whether the thickness value advances top to
        ''' bottom, left to right, or right to left.
        ''' </summary>
        Public flowDirection As DWRITE_FLOW_DIRECTION
        ''' <summary>
        ''' Locale of the range. Can be pertinent where the locale affects the style.
        ''' </summary>
        Public localeName As String
        ''' <summary>
        ''' The measuring mode can be useful to the renderer to determine how
        ''' underlines are rendered, e.g. rounding the thickness to a whole pixel
        ''' in GDI-compatible modes.
        ''' </summary>
        Public measuringMode As DWRITE_MEASURING_MODE
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DWRITE_MATRIX
        ''' <summary>
        ''' Horizontal scaling / cosine of rotation
        ''' </summary>
        Public m11 As Single
        ''' <summary>
        ''' Vertical shear / sine of rotation
        ''' </summary>
        Public m12 As Single
        ''' <summary>
        ''' Horizontal shear / negative sine of rotation
        ''' </summary>
        Public m21 As Single
        ''' <summary>
        ''' Vertical scaling / cosine of rotation
        ''' </summary>
        Public m22 As Single
        ''' <summary>
        ''' Horizontal shift (always orthogonal regardless of rotation)
        ''' </summary>
        Public dx As Single
        ''' <summary>
        ''' Vertical shift (always orthogonal regardless of rotation)
        ''' </summary>
        Public dy As Single
    End Structure

    Public Enum DWRITE_FONT_WEIGHT
        ''' <summary>
        ''' Predefined font weight : Thin (100).
        ''' </summary>
        DWRITE_FONT_WEIGHT_THIN = 100
        ''' <summary>
        ''' Predefined font weight : Extra-light (200).
        ''' </summary>
        DWRITE_FONT_WEIGHT_EXTRA_LIGHT = 200
        ''' <summary>
        ''' Predefined font weight : Ultra-light (200).
        ''' </summary>
        DWRITE_FONT_WEIGHT_ULTRA_LIGHT = 200
        ''' <summary>
        ''' Predefined font weight : Light (300).
        ''' </summary>
        DWRITE_FONT_WEIGHT_LIGHT = 300
        ''' <summary>
        ''' Predefined font weight : Semi-light (350).
        ''' </summary>
        DWRITE_FONT_WEIGHT_SEMI_LIGHT = 350
        ''' <summary>
        ''' Predefined font weight : Normal (400).
        ''' </summary>
        DWRITE_FONT_WEIGHT_NORMAL = 400
        ''' <summary>
        ''' Predefined font weight : Regular (400).
        ''' </summary>
        DWRITE_FONT_WEIGHT_REGULAR = 400
        ''' <summary>
        ''' Predefined font weight : Medium (500).
        ''' </summary>
        DWRITE_FONT_WEIGHT_MEDIUM = 500
        ''' <summary>
        ''' Predefined font weight : Demi-bold (600).
        ''' </summary>
        DWRITE_FONT_WEIGHT_DEMI_BOLD = 600
        ''' <summary>
        ''' Predefined font weight : Semi-bold (600).
        ''' </summary>
        DWRITE_FONT_WEIGHT_SEMI_BOLD = 600
        ''' <summary>
        ''' Predefined font weight : Bold (700).
        ''' </summary>
        DWRITE_FONT_WEIGHT_BOLD = 700
        ''' <summary>
        ''' Predefined font weight : Extra-bold (800).
        ''' </summary>
        DWRITE_FONT_WEIGHT_EXTRA_BOLD = 800
        ''' <summary>
        ''' Predefined font weight : Ultra-bold (800).
        ''' </summary>
        DWRITE_FONT_WEIGHT_ULTRA_BOLD = 800
        ''' <summary>
        ''' Predefined font weight : Black (900).
        ''' </summary>
        DWRITE_FONT_WEIGHT_BLACK = 900
        ''' <summary>
        ''' Predefined font weight : Heavy (900).
        ''' </summary>
        DWRITE_FONT_WEIGHT_HEAVY = 900
        ''' <summary>
        ''' Predefined font weight : Extra-black (950).
        ''' </summary>
        DWRITE_FONT_WEIGHT_EXTRA_BLACK = 950
        ''' <summary>
        ''' Predefined font weight : Ultra-black (950).
        ''' </summary>
        DWRITE_FONT_WEIGHT_ULTRA_BLACK = 950
    End Enum

    ''' <summary>
    ''' The font stretch enumeration describes relative change from the normal aspect ratio
    ''' as specified by a font designer for the glyphs in a font.
    ''' Values less than 1 or greater than 9 are considered to be invalid, and they are rejected by font API functions.
    ''' </summary>
    Public Enum DWRITE_FONT_STRETCH
        ''' <summary>
        ''' Predefined font stretch : Not known (0).
        ''' </summary>
        DWRITE_FONT_STRETCH_UNDEFINED = 0
        ''' <summary>
        ''' Predefined font stretch : Ultra-condensed (1).
        ''' </summary>
        DWRITE_FONT_STRETCH_ULTRA_CONDENSED = 1
        ''' <summary>
        ''' Predefined font stretch : Extra-condensed (2).
        ''' </summary>
        DWRITE_FONT_STRETCH_EXTRA_CONDENSED = 2
        ''' <summary>
        ''' Predefined font stretch : Condensed (3).
        ''' </summary>
        DWRITE_FONT_STRETCH_CONDENSED = 3
        ''' <summary>
        ''' Predefined font stretch : Semi-condensed (4).
        ''' </summary>
        DWRITE_FONT_STRETCH_SEMI_CONDENSED = 4
        ''' <summary>
        ''' Predefined font stretch : Normal (5).
        ''' </summary>
        DWRITE_FONT_STRETCH_NORMAL = 5
        ''' <summary>
        ''' Predefined font stretch : Medium (5).
        ''' </summary>
        DWRITE_FONT_STRETCH_MEDIUM = 5
        ''' <summary>
        ''' Predefined font stretch : Semi-expanded (6).
        ''' </summary>
        DWRITE_FONT_STRETCH_SEMI_EXPANDED = 6
        ''' <summary>
        ''' Predefined font stretch : Expanded (7).
        ''' </summary>
        DWRITE_FONT_STRETCH_EXPANDED = 7
        ''' <summary>
        ''' Predefined font stretch : Extra-expanded (8).
        ''' </summary>
        DWRITE_FONT_STRETCH_EXTRA_EXPANDED = 8
        ''' <summary>
        ''' Predefined font stretch : Ultra-expanded (9).
        ''' </summary>
        DWRITE_FONT_STRETCH_ULTRA_EXPANDED = 9
    End Enum

    Public Enum DWRITE_FONT_STYLE
        ''' <summary>
        ''' Font slope style : Normal.
        ''' </summary>
        DWRITE_FONT_STYLE_NORMAL
        ''' <summary>
        ''' Font slope style : Oblique.
        ''' </summary>
        DWRITE_FONT_STYLE_OBLIQUE
        ''' <summary>
        ''' Font slope style : Italic.
        ''' </summary>
        DWRITE_FONT_STYLE_ITALIC
    End Enum

    Public Enum DWRITE_FONT_FEATURE_TAG
        DWRITE_FONT_FEATURE_TAG_ALTERNATIVE_FRACTIONS = &H63726661 ' 'afrc'
        DWRITE_FONT_FEATURE_TAG_PETITE_CAPITALS_FROM_CAPITALS = &H63703263 ' 'c2pc'
        DWRITE_FONT_FEATURE_TAG_SMALL_CAPITALS_FROM_CAPITALS = &H63733263 ' 'c2sc'
        DWRITE_FONT_FEATURE_TAG_CONTEXTUAL_ALTERNATES = &H746c6163 ' 'calt'
        DWRITE_FONT_FEATURE_TAG_CASE_SENSITIVE_FORMS = &H65736163 ' 'case'
        DWRITE_FONT_FEATURE_TAG_GLYPH_COMPOSITION_DECOMPOSITION = &H706d6363 ' 'ccmp'
        DWRITE_FONT_FEATURE_TAG_CONTEXTUAL_LIGATURES = &H67696c63 ' 'clig'
        DWRITE_FONT_FEATURE_TAG_CAPITAL_SPACING = &H70737063 ' 'cpsp'
        DWRITE_FONT_FEATURE_TAG_CONTEXTUAL_SWASH = &H68777363 ' 'cswh'
        DWRITE_FONT_FEATURE_TAG_CURSIVE_POSITIONING = &H73727563 ' 'curs'
        DWRITE_FONT_FEATURE_TAG_DEFAULT = &H746c6664 ' 'dflt'
        DWRITE_FONT_FEATURE_TAG_DISCRETIONARY_LIGATURES = &H67696c64 ' 'dlig'
        DWRITE_FONT_FEATURE_TAG_EXPERT_FORMS = &H74707865 ' 'expt'
        DWRITE_FONT_FEATURE_TAG_FRACTIONS = &H63617266 ' 'frac'
        DWRITE_FONT_FEATURE_TAG_FULL_WIDTH = &H64697766 ' 'fwid'
        DWRITE_FONT_FEATURE_TAG_HALF_FORMS = &H666c6168 ' 'half'
        DWRITE_FONT_FEATURE_TAG_HALANT_FORMS = &H6e6c6168 ' 'haln'
        DWRITE_FONT_FEATURE_TAG_ALTERNATE_HALF_WIDTH = &H746c6168 ' 'halt'
        DWRITE_FONT_FEATURE_TAG_HISTORICAL_FORMS = &H74736968 ' 'hist'
        DWRITE_FONT_FEATURE_TAG_HORIZONTAL_KANA_ALTERNATES = &H616e6b68 ' 'hkna'
        DWRITE_FONT_FEATURE_TAG_HISTORICAL_LIGATURES = &H67696c68 ' 'hlig'
        DWRITE_FONT_FEATURE_TAG_HALF_WIDTH = &H64697768 ' 'hwid'
        DWRITE_FONT_FEATURE_TAG_HOJO_KANJI_FORMS = &H6f6a6f68 ' 'hojo'
        DWRITE_FONT_FEATURE_TAG_JIS04_FORMS = &H3430706a ' 'jp04'
        DWRITE_FONT_FEATURE_TAG_JIS78_FORMS = &H3837706a ' 'jp78'
        DWRITE_FONT_FEATURE_TAG_JIS83_FORMS = &H3338706a ' 'jp83'
        DWRITE_FONT_FEATURE_TAG_JIS90_FORMS = &H3039706a ' 'jp90'
        DWRITE_FONT_FEATURE_TAG_KERNING = &H6e72656b ' 'kern'
        DWRITE_FONT_FEATURE_TAG_STANDARD_LIGATURES = &H6167696c ' 'liga'
        DWRITE_FONT_FEATURE_TAG_LINING_FIGURES = &H6d756e6c ' 'lnum'
        DWRITE_FONT_FEATURE_TAG_LOCALIZED_FORMS = &H6c636f6c ' 'locl'
        DWRITE_FONT_FEATURE_TAG_MARK_POSITIONING = &H6b72616D ' 'mark'
        DWRITE_FONT_FEATURE_TAG_MATHEMATICAL_GREEK = &H6b72676D ' 'mgrk'
        DWRITE_FONT_FEATURE_TAG_MARK_TO_MARK_POSITIONING = &H6b6d6b6D ' 'mkmk'
        DWRITE_FONT_FEATURE_TAG_ALTERNATE_ANNOTATION_FORMS = &H746c616e ' 'nalt'
        DWRITE_FONT_FEATURE_TAG_NLC_KANJI_FORMS = &H6b636c6e ' 'nlck'
        DWRITE_FONT_FEATURE_TAG_OLD_STYLE_FIGURES = &H6d756e6F ' 'onum'
        DWRITE_FONT_FEATURE_TAG_ORDINALS = &H6e64726F ' 'ordn'
        DWRITE_FONT_FEATURE_TAG_PROPORTIONAL_ALTERNATE_WIDTH = &H746c6170 ' 'palt'
        DWRITE_FONT_FEATURE_TAG_PETITE_CAPITALS = &H70616370 ' 'pcap'
        DWRITE_FONT_FEATURE_TAG_PROPORTIONAL_FIGURES = &H6d756e70 ' 'pnum'
        DWRITE_FONT_FEATURE_TAG_PROPORTIONAL_WIDTHS = &H64697770 ' 'pwid'
        DWRITE_FONT_FEATURE_TAG_QUARTER_WIDTHS = &H64697771 ' 'qwid'
        DWRITE_FONT_FEATURE_TAG_REQUIRED_LIGATURES = &H67696c72 ' 'rlig'
        DWRITE_FONT_FEATURE_TAG_RUBY_NOTATION_FORMS = &H79627572 ' 'ruby'
        DWRITE_FONT_FEATURE_TAG_STYLISTIC_ALTERNATES = &H746c6173 ' 'salt'
        DWRITE_FONT_FEATURE_TAG_SCIENTIFIC_INFERIORS = &H666e6973 ' 'sinf'
        DWRITE_FONT_FEATURE_TAG_SMALL_CAPITALS = &H70636d73 ' 'smcp'
        DWRITE_FONT_FEATURE_TAG_SIMPLIFIED_FORMS = &H6c706d73 ' 'smpl'
        DWRITE_FONT_FEATURE_TAG_STYLISTIC_SET_1 = &H31307373 ' 'ss01'
        DWRITE_FONT_FEATURE_TAG_STYLISTIC_SET_2 = &H32307373 ' 'ss02'
        DWRITE_FONT_FEATURE_TAG_STYLISTIC_SET_3 = &H33307373 ' 'ss03'
        DWRITE_FONT_FEATURE_TAG_STYLISTIC_SET_4 = &H34307373 ' 'ss04'
        DWRITE_FONT_FEATURE_TAG_STYLISTIC_SET_5 = &H35307373 ' 'ss05'
        DWRITE_FONT_FEATURE_TAG_STYLISTIC_SET_6 = &H36307373 ' 'ss06'
        DWRITE_FONT_FEATURE_TAG_STYLISTIC_SET_7 = &H37307373 ' 'ss07'
        DWRITE_FONT_FEATURE_TAG_STYLISTIC_SET_8 = &H38307373 ' 'ss08'
        DWRITE_FONT_FEATURE_TAG_STYLISTIC_SET_9 = &H39307373 ' 'ss09'
        DWRITE_FONT_FEATURE_TAG_STYLISTIC_SET_10 = &H30317373 ' 'ss10'
        DWRITE_FONT_FEATURE_TAG_STYLISTIC_SET_11 = &H31317373 ' 'ss11'
        DWRITE_FONT_FEATURE_TAG_STYLISTIC_SET_12 = &H32317373 ' 'ss12'
        DWRITE_FONT_FEATURE_TAG_STYLISTIC_SET_13 = &H33317373 ' 'ss13'
        DWRITE_FONT_FEATURE_TAG_STYLISTIC_SET_14 = &H34317373 ' 'ss14'
        DWRITE_FONT_FEATURE_TAG_STYLISTIC_SET_15 = &H35317373 ' 'ss15'
        DWRITE_FONT_FEATURE_TAG_STYLISTIC_SET_16 = &H36317373 ' 'ss16'
        DWRITE_FONT_FEATURE_TAG_STYLISTIC_SET_17 = &H37317373 ' 'ss17'
        DWRITE_FONT_FEATURE_TAG_STYLISTIC_SET_18 = &H38317373 ' 'ss18'
        DWRITE_FONT_FEATURE_TAG_STYLISTIC_SET_19 = &H39317373 ' 'ss19'
        DWRITE_FONT_FEATURE_TAG_STYLISTIC_SET_20 = &H30327373 ' 'ss20'
        DWRITE_FONT_FEATURE_TAG_SUBSCRIPT = &H73627573 ' 'subs'
        DWRITE_FONT_FEATURE_TAG_SUPERSCRIPT = &H73707573 ' 'sups'
        DWRITE_FONT_FEATURE_TAG_SWASH = &H68737773 ' 'swsh'
        DWRITE_FONT_FEATURE_TAG_TITLING = &H6c746974 ' 'titl'
        DWRITE_FONT_FEATURE_TAG_TRADITIONAL_NAME_FORMS = &H6d616e74 ' 'tnam'
        DWRITE_FONT_FEATURE_TAG_TABULAR_FIGURES = &H6d756e74 ' 'tnum'
        DWRITE_FONT_FEATURE_TAG_TRADITIONAL_FORMS = &H64617274 ' 'trad'
        DWRITE_FONT_FEATURE_TAG_THIRD_WIDTHS = &H64697774 ' 'twid'
        DWRITE_FONT_FEATURE_TAG_UNICASE = &H63696e75 ' 'unic'
        DWRITE_FONT_FEATURE_TAG_VERTICAL_WRITING = &H74726576 ' 'vert'
        DWRITE_FONT_FEATURE_TAG_VERTICAL_ALTERNATES_AND_ROTATION = &H32747276 ' 'vrt2'
        DWRITE_FONT_FEATURE_TAG_SLASHED_ZERO = &H6f72657a ' 'zero'
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DWRITE_FONT_FEATURE
        ''' <summary>
        ''' The feature OpenType name identifier.
        ''' </summary>
        Public nameTag As DWRITE_FONT_FEATURE_TAG
        ''' <summary>
        ''' Execution parameter of the feature.
        ''' </summary>
        ''' <remarks>
        ''' The parameter should be non-zero to enable the feature.  Once enabled, a feature can't be disabled again within
        ''' the same range.  Features requiring a selector use this value to indicate the selector index. 
        ''' </remarks>
        Public parameter As UInteger

        Public Sub New(nameTag As DWRITE_FONT_FEATURE_TAG, parameter As UInteger)
            Me.nameTag = nameTag
            Me.parameter = parameter
        End Sub
    End Structure

    Public Enum DWRITE_NUMBER_SUBSTITUTION_METHOD
        ''' <summary>
        ''' Specifies that the substitution method should be determined based
        ''' on LOCALE_IDIGITSUBSTITUTION value of the specified text culture.
        ''' </summary>
        DWRITE_NUMBER_SUBSTITUTION_METHOD_FROM_CULTURE
        ''' <summary>
        ''' If the culture is Arabic or Farsi, specifies that the number shape
        ''' depend on the context. Either traditional or nominal number shape
        ''' are used depending on the nearest preceding strong character or (if
        ''' there is none) the reading direction of the paragraph.
        ''' </summary>
        DWRITE_NUMBER_SUBSTITUTION_METHOD_CONTEXTUAL
        ''' <summary>
        ''' Specifies that code points 0x30-0x39 are always rendered as nominal numeral 
        ''' shapes (ones of the European number), i.e., no substitution is performed.
        ''' </summary>
        DWRITE_NUMBER_SUBSTITUTION_METHOD_NONE
        ''' <summary>
        ''' Specifies that number are rendered using the national number shape 
        ''' as specified by the LOCALE_SNATIVEDIGITS value of the specified text culture.
        ''' </summary>
        DWRITE_NUMBER_SUBSTITUTION_METHOD_NATIONAL
        ''' <summary>
        ''' Specifies that number are rendered using the traditional shape
        ''' for the specified culture. For most cultures, this is the same as
        ''' NativeNational. However, NativeNational results in Latin number
        ''' for some Arabic cultures, whereas this value results in Arabic
        ''' number for all Arabic cultures.
        ''' </summary>
        DWRITE_NUMBER_SUBSTITUTION_METHOD_TRADITIONAL
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DWRITE_TEXT_RANGE
        ''' <summary>
        '''         ''' The start text position of the range.
        '''         ''' </summary>
        Public startPosition As UInteger
        ''' <summary>
        '''         ''' The number of text positions in the range.
        '''         ''' </summary>
        Public length As UInteger
        Public Sub New(startPosition As UInteger, length As UInteger)
            Me.startPosition = startPosition
            Me.length = length
        End Sub
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DWRITE_SCRIPT_ANALYSIS
        ''' <summary>
        ''' Zero-based index representation of writing system script.
        ''' </summary>
        Public script As UShort
        ''' <summary>
        ''' Additional shaping requirement of text.
        ''' </summary>
        Public shapes As DWRITE_SCRIPT_SHAPES
    End Structure

    Public Enum DWRITE_SCRIPT_SHAPES
        ''' <summary>
        ''' No additional shaping requirement. Text is shaped with the writing system default behavior.
        ''' </summary>
        DWRITE_SCRIPT_SHAPES_DEFAULT = 0
        ''' <summary>
        ''' Text should leave no visual on display i.e. control or format control characters.
        ''' </summary>
        DWRITE_SCRIPT_SHAPES_NO_VISUAL = 1
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DWRITE_SHAPING_TEXT_PROPERTIES
        ''' <summary>
        ''' This character can be shaped independently from the others
        ''' (usually set for the space character).
        ''' </summary>
        <MarshalAs(UnmanagedType.U2, SizeConst:=1)>
        Public isShapedAlone As UShort
        ''' <summary>
        ''' Reserved for use by shaping engine.
        ''' </summary>
        <MarshalAs(UnmanagedType.U2, SizeConst:=15)>
        Public reserved As UShort
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DWRITE_SHAPING_GLYPH_PROPERTIES
        ''' <summary>
        ''' Justification class, whether to use spacing, kashidas, or
        ''' another method. This exists for backwards compatibility
        ''' with Uniscribe's SCRIPT_JUSTIFY enum.
        ''' </summary>
        <MarshalAs(UnmanagedType.U2, SizeConst:=4)>
        Public justification As UShort
        ''' <summary>
        ''' Indicates glyph is the first of a cluster.
        ''' </summary>
        <MarshalAs(UnmanagedType.U2, SizeConst:=1)>
        Public isClusterStart1 As UShort
        ''' <summary>
        ''' Glyph is a diacritic.
        ''' </summary>
        <MarshalAs(UnmanagedType.U2, SizeConst:=1)>
        Public isDiacritic As UShort
        ''' <summary>
        ''' Glyph has no width, blank, ZWJ, ZWNJ etc.
        ''' </summary>
        <MarshalAs(UnmanagedType.U2, SizeConst:=1)>
        Public isZeroWidthSpace As UShort
        ''' <summary>
        ''' Reserved for use by shaping engine.
        ''' </summary>
        <MarshalAs(UnmanagedType.U2, SizeConst:=9)>
        Public reserved As UShort
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DWRITE_LINE_METRICS
        ''' <summary>
        ''' The number of total text positions in the line.
        ''' This includes any trailing whitespace and newline characters.
        ''' </summary>
        Public length As Integer
        ''' <summary>
        ''' The number of whitespace positions at the end of the line.  Newline
        ''' sequences are considered whitespace.
        ''' </summary>
        Public trailingWhitespaceLength As Integer
        ''' <summary>
        ''' The number of characters in the newline sequence at the end of the line.
        ''' If the count is zero, then the line was either wrapped or it is the
        ''' end of the text.
        ''' </summary>
        Public newlineLength As Integer
        ''' <summary>
        ''' Height of the line as measured from top to bottom.
        ''' </summary>
        Public height As Single
        ''' <summary>
        ''' Distance from the top of the line to its baseline.
        ''' </summary>
        Public baseline As Single
        ''' <summary>
        ''' The line is trimmed.
        ''' </summary>
        Public isTrimmed As Boolean
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DWRITE_TEXT_METRICS
        ''' <summary>
        ''' Left-most point of formatted text relative to layout box
        ''' (excluding any glyph overhang).
        ''' </summary>
        Public left As Single
        ''' <summary>
        ''' Top-most point of formatted text relative to layout box
        ''' (excluding any glyph overhang).
        ''' </summary>
        Public top As Single
        ''' <summary>
        ''' The width of the formatted text ignoring trailing whitespace
        ''' at the end of each line.
        ''' </summary>
        Public width As Single
        ''' <summary>
        ''' The width of the formatted text taking into account the
        ''' trailing whitespace at the end of each line.
        ''' </summary>
        Public widthIncludingTrailingWhitespace As Single
        ''' <summary>
        ''' The height of the formatted text. The height of an empty string
        ''' is determined by the size of the default font's line height.
        ''' </summary>
        Public height As Single
        ''' <summary>
        ''' Initial width given to the layout. Depending on whether the text
        ''' was wrapped or not, it can be either larger or smaller than the
        ''' text content width.
        ''' </summary>
        Public layoutWidth As Single
        ''' <summary>
        ''' Initial height given to the layout. Depending on the length of the
        ''' text, it may be larger or smaller than the text content height.
        ''' </summary>
        Public layoutHeight As Single
        ''' <summary>
        ''' The maximum reordering count of any line of text, used
        ''' to calculate the most number of hit-testing boxes needed.
        ''' If the layout has no bidirectional text or no text at all,
        ''' the minimum level is 1.
        ''' </summary>
        Public maxBidiReorderingDepth As Integer
        ''' <summary>
        ''' Total number of lines.
        ''' </summary>
        Public lineCount As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DWRITE_CLUSTER_METRICS
        ''' <summary>
        ''' The total advance width of all glyphs in the cluster.
        ''' </summary>
        Public width As Single
        ''' <summary>
        ''' The number of text positions in the cluster.
        ''' </summary>
        Public length As UShort
        ''' <summary>
        ''' Indicate whether line can be broken right after the cluster.
        ''' </summary>
        <MarshalAs(UnmanagedType.U2, SizeConst:=1)>
        Public canWrapLineAfter As UShort
        ''' <summary>
        ''' Indicate whether the cluster corresponds to whitespace character.
        ''' </summary>
        <MarshalAs(UnmanagedType.U2, SizeConst:=1)>
        Public isWhitespace As UShort
        ''' <summary>
        ''' Indicate whether the cluster corresponds to a newline character.
        ''' </summary>
        <MarshalAs(UnmanagedType.U2, SizeConst:=1)>
        Public isNewline As UShort
        ''' <summary>
        ''' Indicate whether the cluster corresponds to soft hyphen character.
        ''' </summary>
        <MarshalAs(UnmanagedType.U2, SizeConst:=1)>
        Public isSoftHyphen As UShort
        ''' <summary>
        ''' Indicate whether the cluster is read from right to left.
        ''' </summary>
        Public isRightToLeft As UShort
        <MarshalAs(UnmanagedType.U2, SizeConst:=11)>
        Public padding As UShort
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DWRITE_HIT_TEST_METRICS
        ''' <summary>
        ''' First text position within the geometry.
        ''' </summary>
        Public textPosition As Integer
        ''' <summary>
        ''' Number of text positions within the geometry.
        ''' </summary>
        Public length As Integer
        ''' <summary>
        ''' Left position of the top-left coordinate of the geometry.
        ''' </summary>
        Public left As Single
        ''' <summary>
        ''' Top position of the top-left coordinate of the geometry.
        ''' </summary>
        Public top As Single
        ''' <summary>
        ''' Geometry's width.
        ''' </summary>
        Public width As Single
        ''' <summary>
        ''' Geometry's height.
        ''' </summary>
        Public height As Single
        ''' <summary>
        ''' Bidi level of text positions enclosed within the geometry.
        ''' </summary>
        Public bidiLevel As Integer
        ''' <summary>
        ''' Geometry encloses text?
        ''' </summary>
        Public isText As Boolean
        ''' <summary>
        ''' Range is trimmed.
        ''' </summary>
        Public isTrimmed As Boolean
    End Structure

    Public Enum DWRITE_INFORMATIONAL_STRING_ID
        ''' <summary>
        ''' Unspecified name ID.
        ''' </summary>
        DWRITE_INFORMATIONAL_STRING_NONE
        ''' <summary>
        ''' Copyright notice provided by the font.
        ''' </summary>
        DWRITE_INFORMATIONAL_STRING_COPYRIGHT_NOTICE
        ''' <summary>
        ''' String containing a version number.
        ''' </summary>
        DWRITE_INFORMATIONAL_STRING_VERSION_STRINGS
        ''' <summary>
        ''' Trademark information provided by the font.
        ''' </summary>
        DWRITE_INFORMATIONAL_STRING_TRADEMARK
        ''' <summary>
        ''' Name of the font manufacturer.
        ''' </summary>
        DWRITE_INFORMATIONAL_STRING_MANUFACTURER
        ''' <summary>
        ''' Name of the font designer.
        ''' </summary>
        DWRITE_INFORMATIONAL_STRING_DESIGNER
        ''' <summary>
        ''' URL of font designer (with protocol, e.g., http://, ftp://).
        ''' </summary>
        DWRITE_INFORMATIONAL_STRING_DESIGNER_URL
        ''' <summary>
        ''' Description of the font. Can contain revision information, usage recommendations, history, features, etc.
        ''' </summary>
        DWRITE_INFORMATIONAL_STRING_DESCRIPTION
        ''' <summary>
        ''' URL of font vendor (with protocol, e.g., http://, ftp://). If a unique serial number is embedded in the URL, it can be used to register the font.
        ''' </summary>
        DWRITE_INFORMATIONAL_STRING_FONT_VENDOR_URL
        ''' <summary>
        ''' Description of how the font may be legally used, or different example scenarios for licensed use. This field should be written in plain language, not legalese.
        ''' </summary>
        DWRITE_INFORMATIONAL_STRING_LICENSE_DESCRIPTION
        ''' <summary>
        ''' URL where additional licensing information can be found.
        ''' </summary>
        DWRITE_INFORMATIONAL_STRING_LICENSE_INFO_URL
        ''' <summary>
        ''' GDI-compatible family name. Because GDI allows a maximum of four fonts per family, fonts in the same family may have different GDI-compatible family names
        ''' (e.g., "Arial", "Arial Narrow", "Arial Black").
        ''' </summary>
        DWRITE_INFORMATIONAL_STRING_WIN32_FAMILY_NAMES
        ''' <summary>
        ''' GDI-compatible subfamily name.
        ''' </summary>
        DWRITE_INFORMATIONAL_STRING_WIN32_SUBFAMILY_NAMES
        ''' <summary>
        ''' Family name preferred by the designer. This enables font designers to group more than four fonts in a single family without losing compatibility with
        ''' GDI. This name is typically only present if it differs from the GDI-compatible family name.
        ''' </summary>
        DWRITE_INFORMATIONAL_STRING_PREFERRED_FAMILY_NAMES
        ''' <summary>
        ''' Subfamily name preferred by the designer. This name is typically only present if it differs from the GDI-compatible subfamily name. 
        ''' </summary>
        DWRITE_INFORMATIONAL_STRING_PREFERRED_SUBFAMILY_NAMES
        ''' <summary>
        ''' Sample text. This can be the font name or any other text that the designer thinks is the best example to display the font in.
        ''' </summary>
        DWRITE_INFORMATIONAL_STRING_SAMPLE_TEXT
        ''' <summary>
        ''' The full name of the font, e.g. "Arial Bold", from name id 4 in the name table.
        ''' </summary>
        DWRITE_INFORMATIONAL_STRING_FULL_NAME
        ''' <summary>
        ''' The postscript name of the font, e.g. "GillSans-Bold" from name id 6 in the name table.
        ''' </summary>
        DWRITE_INFORMATIONAL_STRING_POSTSCRIPT_NAME
        ''' <summary>
        ''' The postscript CID findfont name, from name id 20 in the name table.
        ''' </summary>
        DWRITE_INFORMATIONAL_STRING_POSTSCRIPT_CID_NAME
        ''' <summary>
        ''' Family name for the weight-width-slope model.
        ''' </summary>
        DWRITE_INFORMATIONAL_STRING_WWS_FAMILY_NAME
        ''' <summary>
        ''' Script/language tag to identify the scripts or languages that the font was
        ''' primarily designed to support. See DWRITE_FONT_PROPERTY_ID_DESIGN_SCRIPT_LANGUAGE_TAG
        ''' for a longer description.
        ''' </summary>
        DWRITE_INFORMATIONAL_STRING_DESIGN_SCRIPT_LANGUAGE_TAG
        ''' <summary>
        ''' Script/language tag to identify the scripts or languages that the font declares
        ''' it is able to support.
        ''' </summary>
        DWRITE_INFORMATIONAL_STRING_SUPPORTED_SCRIPT_LANGUAGE_TAG
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DWRITE_LINE_BREAKPOINT
        ''' <summary>
        ''' Breaking condition before the character.
        ''' </summary>
        <MarshalAs(UnmanagedType.U1, SizeConst:=2)>
        Public breakConditionBefore As Byte
        ''' <summary>
        ''' Breaking condition after the character.
        ''' </summary>
        <MarshalAs(UnmanagedType.U1, SizeConst:=2)>
        Public breakConditionAfter As Byte
        ''' <summary>
        ''' The character is some form of whitespace, which may be meaningful
        ''' for justification.
        ''' </summary>
        <MarshalAs(UnmanagedType.U1, SizeConst:=1)>
        Public isWhitespace As Byte
        ''' <summary>
        ''' The character is a soft hyphen, often used to indicate hyphenation
        ''' points inside words.
        ''' </summary>
        <MarshalAs(UnmanagedType.U1, SizeConst:=1)>
        Public isSoftHyphen As Byte

        <MarshalAs(UnmanagedType.U1, SizeConst:=2)>
        Public padding As Byte
    End Structure

    Public Enum D2D1_FILL_MODE
        D2D1_FILL_MODE_ALTERNATE = 0
        D2D1_FILL_MODE_WINDING = 1
        D2D1_FILL_MODE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_PATH_SEGMENT
        D2D1_PATH_SEGMENT_NONE = &H00000000
        D2D1_PATH_SEGMENT_FORCE_UNSTROKED = &H00000001
        D2D1_PATH_SEGMENT_FORCE_ROUND_LINE_JOIN = &H2
        D2D1_PATH_SEGMENT_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_FIGURE_BEGIN
        D2D1_FIGURE_BEGIN_FILLED = 0
        D2D1_FIGURE_BEGIN_HOLLOW = 1
        D2D1_FIGURE_BEGIN_FORCE_DWORD = &HFFFFFFFF
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_BEZIER_SEGMENT
        Public point1 As D2D1_POINT_2F
        Public point2 As D2D1_POINT_2F
        Public point3 As D2D1_POINT_2F
    End Structure

    Public Enum D2D1_FIGURE_END
        D2D1_FIGURE_END_OPEN = 0
        D2D1_FIGURE_END_CLOSED = 1
        D2D1_FIGURE_END_FORCE_DWORD = &HFFFFFFFF
    End Enum
    Public Enum DWRITE_GLYPH_ORIENTATION_ANGLE
        ''' <summary>
        ''' Glyph orientation is upright.
        ''' </summary>
        DWRITE_GLYPH_ORIENTATION_ANGLE_0_DEGREES

        ''' <summary>
        ''' Glyph orientation is rotated 90 clockwise.
        ''' </summary>
        DWRITE_GLYPH_ORIENTATION_ANGLE_90_DEGREES

        ''' <summary>
        ''' Glyph orientation is upside-down.
        ''' </summary>
        DWRITE_GLYPH_ORIENTATION_ANGLE_180_DEGREES

        ''' <summary>
        ''' Glyph orientation is rotated 270 clockwise.
        ''' </summary>
        DWRITE_GLYPH_ORIENTATION_ANGLE_270_DEGREES
    End Enum

    Public Enum DWRITE_VERTICAL_GLYPH_ORIENTATION
        ''' <summary>
        ''' In vertical layout, naturally horizontal scripts (Latin, Thai, Arabic,
        ''' Devanagari) rotate 90 degrees clockwise, while ideographic scripts
        ''' (Chinese, Japanese, Korean) remain upright, 0 degrees.
        ''' </summary>
        DWRITE_VERTICAL_GLYPH_ORIENTATION_DEFAULT

        ''' <summary>
        ''' Ideographic scripts and scripts that permit stacking
        ''' (Latin, Hebrew) are stacked in vertical reading layout.
        ''' Connected scripts (Arabic, Syriac, 'Phags-pa, Ogham),
        ''' which would otherwise look broken if glyphs were kept
        ''' at 0 degrees, remain connected and rotate.
        ''' </summary>
        DWRITE_VERTICAL_GLYPH_ORIENTATION_STACKED
    End Enum

    Public Enum DWRITE_OPTICAL_ALIGNMENT
        ''' <summary>
        ''' Align to the default metrics of the glyph.
        ''' </summary>
        DWRITE_OPTICAL_ALIGNMENT_NONE

        ''' <summary>
        ''' Align glyphs to the margins. Without this, some small whitespace
        ''' may be present between the text and the margin from the glyph's side
        ''' bearing values. Note that glyphs may still overhang outside the
        ''' margin, such as flourishes or italic slants.
        ''' </summary>
        DWRITE_OPTICAL_ALIGNMENT_NO_SIDE_BEARINGS
    End Enum

    Public Enum DWRITE_GRID_FIT_MODE
        ''' <summary>
        ''' Choose grid fitting base on the font's gasp table information.
        ''' </summary>
        DWRITE_GRID_FIT_MODE_DEFAULT

        ''' <summary>
        ''' Always disable grid fitting, using the ideal glyph outlines.
        ''' </summary>
        DWRITE_GRID_FIT_MODE_DISABLED

        ''' <summary>
        ''' Enable grid fitting, adjusting glyph outlines for device pixel display.
        ''' </summary>
        DWRITE_GRID_FIT_MODE_ENABLED
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DWRITE_TEXT_METRICS1 ' : DWRITE_TEXT_METRICS
        ''' <summary>
        ''' Left-most point of formatted text relative to layout box
        ''' (excluding any glyph overhang).
        ''' </summary>
        Public left As Single
        ''' <summary>
        ''' Top-most point of formatted text relative to layout box
        ''' (excluding any glyph overhang).
        ''' </summary>
        Public top As Single
        ''' <summary>
        ''' The width of the formatted text ignoring trailing whitespace
        ''' at the end of each line.
        ''' </summary>
        Public width As Single
        ''' <summary>
        ''' The width of the formatted text taking into account the
        ''' trailing whitespace at the end of each line.
        ''' </summary>
        Public widthIncludingTrailingWhitespace As Single
        ''' <summary>
        ''' The height of the formatted text. The height of an empty string
        ''' is determined by the size of the default font's line height.
        ''' </summary>
        Public height As Single
        ''' <summary>
        ''' Initial width given to the layout. Depending on whether the text
        ''' was wrapped or not, it can be either larger or smaller than the
        ''' text content width.
        ''' </summary>
        Public layoutWidth As Single
        ''' <summary>
        ''' Initial height given to the layout. Depending on the length of the
        ''' text, it may be larger or smaller than the text content height.
        ''' </summary>
        Public layoutHeight As Single
        ''' <summary>
        ''' The maximum reordering count of any line of text, used
        ''' to calculate the most number of hit-testing boxes needed.
        ''' If the layout has no bidirectional text or no text at all,
        ''' the minimum level is 1.
        ''' </summary>
        Public maxBidiReorderingDepth As Integer
        ''' <summary>
        ''' Total number of lines.
        ''' </summary>
        Public lineCount As Integer
        ''' <summary>
        ''' The height of the formatted text taking into account the
        ''' trailing whitespace at the end of each line, which will
        ''' matter for vertical reading directions.
        ''' </summary>
        Public heightIncludingTrailingWhitespace As Single
    End Structure

    Public Enum DWRITE_BASELINE
        ''' <summary>
        ''' The Roman baseline for horizontal, Central baseline for vertical.
        ''' </summary>
        DWRITE_BASELINE_DEFAULT

        ''' <summary>
        ''' The baseline used by alphabetic scripts such as Latin, Greek, Cyrillic.
        ''' </summary>
        DWRITE_BASELINE_ROMAN

        ''' <summary>
        ''' Central baseline, generally used for vertical text.
        ''' </summary>
        DWRITE_BASELINE_CENTRAL

        ''' <summary>
        ''' Mathematical baseline which math characters are centered on.
        ''' </summary>
        DWRITE_BASELINE_MATH

        ''' <summary>
        ''' Hanging baseline, used in scripts like Devanagari.
        ''' </summary>
        DWRITE_BASELINE_HANGING

        ''' <summary>
        ''' Ideographic bottom baseline for CJK, left in vertical.
        ''' </summary>
        DWRITE_BASELINE_IDEOGRAPHIC_BOTTOM

        ''' <summary>
        ''' Ideographic top baseline for CJK, right in vertical.
        ''' </summary>
        DWRITE_BASELINE_IDEOGRAPHIC_TOP

        ''' <summary>
        ''' The bottom-most extent in horizontal, left-most in vertical.
        ''' </summary>
        DWRITE_BASELINE_MINIMUM

        ''' <summary>
        ''' The top-most extent in horizontal, right-most in vertical.
        ''' </summary>
        DWRITE_BASELINE_MAXIMUM
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DWRITE_SCRIPT_PROPERTIES
        ''' <summary>
        ''' The standardized four character code for the given script.
        ''' Note these only include the general Unicode scripts, not any
        ''' additional ISO 15924 scripts for bibliographic distinction
        ''' (for example, Fraktur Latin vs Gaelic Latin).
        ''' http://unicode.org/iso15924/iso15924-codes.html
        ''' </summary>
        Public isoScriptCode As UInteger

        ''' <summary>
        ''' The standardized numeric code, ranging 0-999.
        ''' http://unicode.org/iso15924/iso15924-codes.html
        ''' </summary>
        Public isoScriptNumber As UInteger

        ''' <summary>
        ''' Number of characters to estimate look-ahead for complex scripts.
        ''' Latin and all Kana are generally 1. Indic scripts are up to 15,
        ''' and most others are 8. Note that combining marks and variation
        ''' selectors can produce clusters longer than these look-aheads,
        ''' so this estimate is considered typical language use. Diacritics
        ''' must be tested explicitly separately.
        ''' </summary>
        Public clusterLookahead As UInteger

        ''' <summary>
        ''' Appropriate character to elongate the given script for justification.
        '''
        ''' Examples:
        '''   Arabic    - U+0640 Tatweel
        '''   Ogham     - U+1680 Ogham Space Mark
        ''' </summary>
        Public justificationCharacter As UInteger

        ''' <summary>
        ''' Restrict the caret to whole clusters, like Thai and Devanagari. Scripts
        ''' such as Arabic by default allow navigation between clusters. Others
        ''' like Thai always navigate across whole clusters.
        ''' </summary>
        ''' 
        <MarshalAs(UnmanagedType.SysUInt, SizeConst:=1)>
        Public restrictCaretToClusters As UInteger

        ''' <summary>
        ''' The language uses dividers between words, such as spaces between Latin
        ''' or the Ethiopic wordspace.
        '''
        ''' Examples: Latin, Greek, Devanagari, Ethiopic
        ''' Excludes: Chinese, Korean, Thai.
        ''' </summary>
        '''
        <MarshalAs(UnmanagedType.SysUInt, SizeConst:=1)>
        Public usesWordDividers As UInteger

        ''' <summary>
        ''' The characters are discrete units from each other. This includes both
        ''' block scripts and clustered scripts.
        '''
        ''' Examples: Latin, Greek, Cyrillic, Hebrew, Chinese, Thai
        ''' </summary>
        ''' 
        <MarshalAs(UnmanagedType.SysUInt, SizeConst:=1)>
        Public isDiscreteWriting As UInteger

        ''' <summary>
        ''' The language is a block script, expanding between characters.
        '''
        ''' Examples: Chinese, Japanese, Korean, Bopomofo.
        ''' </summary>
        ''' 
        <MarshalAs(UnmanagedType.SysUInt, SizeConst:=1)>
        Public isBlockWriting As UInteger

        ''' <summary>
        ''' The language is justified within glyph clusters, not just between glyph
        ''' clusters. One such as the character sequence is Thai Lu and Sara Am
        ''' (U+E026, U+E033) which form a single cluster but still expand between
        ''' them.
        '''
        ''' Examples: Thai, Lao, Khmer
        ''' </summary>
        ''' 
        <MarshalAs(UnmanagedType.SysUInt, SizeConst:=1)>
        Public isDistributedWithinCluster As UInteger

        ''' <summary>
        ''' The script's clusters are connected to each other (such as the
        ''' baseline-linked Devanagari), and no separation should be added
        ''' between characters. Note that cursively linked scripts like Arabic
        ''' are also connected (but not all connected scripts are
        ''' cursive).
        ''' 
        ''' Examples: Devanagari, Arabic, Syriac, Bengali, Gurmukhi, Ogham
        ''' Excludes: Latin, Chinese, Thaana
        ''' </summary>
        ''' 
        <MarshalAs(UnmanagedType.SysUInt, SizeConst:=1)>
        Public isConnectedWriting1 As UInteger

        ''' <summary>
        ''' The script is naturally cursive (Arabic/Syriac), meaning it uses other
        ''' justification methods like kashida extension rather than intercharacter
        ''' spacing. Note that although other scripts like Latin and Japanese may
        ''' actually support handwritten cursive forms, they are not considered
        ''' cursive scripts.
        ''' 
        ''' Examples: Arabic, Syriac, Mongolian
        ''' Excludes: Thaana, Devanagari, Latin, Chinese
        ''' </summary>
        ''' 
        <MarshalAs(UnmanagedType.SysUInt, SizeConst:=1)>
        Public isCursiveWriting1 As UInteger

        <MarshalAs(UnmanagedType.SysUInt, SizeConst:=25)>
        Public reserved As UInteger
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DWRITE_JUSTIFICATION_OPPORTUNITY
        ''' <summary>
        ''' Minimum amount of expansion to apply to the side of the glyph.
        ''' This may vary from 0 to infinity, typically being zero except
        ''' for kashida.
        ''' </summary>
        Public expansionMinimum As Single

        ''' <summary>
        ''' Maximum amount of expansion to apply to the side of the glyph.
        ''' This may vary from 0 to infinity, being zero for fixed-size characters
        ''' and connected scripts, and non-zero for discrete scripts, and non-zero
        ''' for cursive scripts at expansion points.
        ''' </summary>
        Public expansionMaximum As Single

        ''' <summary>
        ''' Maximum amount of compression to apply to the side of the glyph.
        ''' This may vary from 0 up to the glyph cluster size.
        ''' </summary>
        Public compressionMaximum As Single

        ''' <summary>
        ''' Priority of this expansion point. Larger priorities are applied later,
        ''' while priority zero does nothing.
        ''' </summary>
        ''' 
        <MarshalAs(UnmanagedType.SysUInt, SizeConst:=8)>
        Public expansionPriority As UInteger

        ''' <summary>
        ''' Priority of this compression point. Larger priorities are applied later,
        ''' while priority zero does nothing.
        ''' </summary>
        ''' 
        <MarshalAs(UnmanagedType.SysUInt, SizeConst:=8)>
        Public compressionPriority As UInteger

        ''' <summary>
        ''' Allow this expansion point to use up any remaining slack space even
        ''' after all expansion priorities have been used up.
        ''' </summary>
        ''' 
        <MarshalAs(UnmanagedType.SysUInt, SizeConst:=1)>
        Public allowResidualExpansion As UInteger

        ''' <summary>
        ''' Allow this compression point to use up any remaining space even after
        ''' all compression priorities have been used up.
        ''' </summary>
        ''' 
        <MarshalAs(UnmanagedType.SysUInt, SizeConst:=1)>
        Public allowResidualCompression As UInteger

        ''' <summary>
        ''' Apply expansion/compression to the leading edge of the glyph. This will
        ''' be false for connected scripts, fixed-size characters, and diacritics.
        ''' It is generally false within a multi-glyph cluster, unless the script
        ''' allows expansion of glyphs within a cluster, like Thai.
        ''' </summary>
        ''' 
        <MarshalAs(UnmanagedType.SysUInt, SizeConst:=1)>
        Public applyToLeadingEdge As UInteger

        ''' <summary>
        ''' Apply expansion/compression to the trailing edge of the glyph. This will
        ''' be false for connected scripts, fixed-size characters, and diacritics.
        ''' It is generally false within a multi-glyph cluster, unless the script
        ''' allows expansion of glyphs within a cluster, like Thai.
        ''' </summary>
        ''' 
        <MarshalAs(UnmanagedType.SysUInt, SizeConst:=1)>
        Public applyToTrailingEdge As UInteger

        <MarshalAs(UnmanagedType.SysUInt, SizeConst:=12)>
        Public reserved As UInteger
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DWRITE_UNICODE_RANGE
        ''' <summary>
        ''' The first codepoint in the Unicode range.
        ''' </summary>
        Public first As UInteger

        ''' <summary>
        ''' The last codepoint in the Unicode range.
        ''' </summary>
        Public last As UInteger
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DWRITE_FONT_METRICS1 ' : public DWRITE_FONT_METRICS
        ''' <summary>
        ''' The number of font design units per em unit.
        ''' Font files use their own coordinate system of font design units.
        ''' A font design unit is the smallest measurable unit in the em square,
        ''' an imaginary square that is used to size and align glyphs.
        ''' The concept of em square is used as a reference scale factor when defining font size and device transformation semantics.
        ''' The size of one em square is also commonly used to compute the paragraph indentation value.
        ''' </summary>
        Public designUnitsPerEm As UShort
        ''' <summary>
        ''' Ascent value of the font face in font design units.
        ''' Ascent is the distance from the top of font character alignment box to English baseline.
        ''' </summary>
        Public ascent As UShort
        ''' <summary>
        ''' Descent value of the font face in font design units.
        ''' Descent is the distance from the bottom of font character alignment box to English baseline.
        ''' </summary>
        Public descent As UShort
        ''' <summary>
        ''' Line gap in font design units.
        ''' Recommended additional white space to add between lines to improve legibility. The recommended line spacing 
        ''' (baseline-to-baseline distance) is thus the sum of ascent, descent, and lineGap. The line gap is usually 
        ''' positive or zero but can be negative, in which case the recommended line spacing is less than the height
        ''' of the character alignment box.
        ''' </summary>
        Public lineGap As UShort
        ''' <summary>
        ''' Cap height value of the font face in font design units.
        ''' Cap height is the distance from English baseline to the top of a typical English capital.
        ''' Capital "H" is often used as a reference character for the purpose of calculating the cap height value.
        ''' </summary>
        Public capHeight As UShort
        ''' <summary>
        ''' x-height value of the font face in font design units.
        ''' x-height is the distance from English baseline to the top of lowercase letter "x", or a similar lowercase character.
        ''' </summary>
        Public xHeight As UShort
        ''' <summary>
        ''' The underline position value of the font face in font design units.
        ''' Underline position is the position of underline relative to the English baseline.
        ''' The value is usually made negative in order to place the underline below the baseline.
        ''' </summary>
        Public underlinePosition As UShort
        ''' <summary>
        ''' The suggested underline thickness value of the font face in font design units.
        ''' </summary>
        Public underlineThickness As UShort
        ''' <summary>
        ''' The strikethrough position value of the font face in font design units.
        ''' Strikethrough position is the position of strikethrough relative to the English baseline.
        ''' The value is usually made positive in order to place the strikethrough above the baseline.
        ''' </summary>
        Public strikethroughPosition As UShort
        ''' <summary>
        ''' The suggested strikethrough thickness value of the font face in font design units.
        ''' </summary>
        Public strikethroughThickness As UShort
        ''' <summary>
        ''' Left edge of accumulated bounding blackbox of all glyphs in the font.
        ''' </summary>
        Public glyphBoxLeft As Short

        ''' <summary>
        ''' Top edge of accumulated bounding blackbox of all glyphs in the font.
        ''' </summary>
        Public glyphBoxTop As Short

        ''' <summary>
        ''' Right edge of accumulated bounding blackbox of all glyphs in the font.
        ''' </summary>
        Public glyphBoxRight As Short

        ''' <summary>
        ''' Bottom edge of accumulated bounding blackbox of all glyphs in the font.
        ''' </summary>
        Public glyphBoxBottom As Short

        ''' <summary>
        ''' Horizontal position of the subscript relative to the baseline origin.
        ''' This is typically negative (to the left) in italic/oblique fonts, and
        ''' zero in regular fonts.
        ''' </summary>
        Public subscriptPositionX As Short

        ''' <summary>
        ''' Vertical position of the subscript relative to the baseline.
        ''' This is typically negative.
        ''' </summary>
        Public subscriptPositionY As Short

        ''' <summary>
        ''' Horizontal size of the subscript em box in design units, used to
        ''' scale the simulated subscript relative to the full em box size.
        ''' This the numerator of the scaling ratio where denominator is the
        ''' design units per em. If this member is zero, the font does not specify
        ''' a scale factor, and the client should use its own policy.
        ''' </summary>
        Public subscriptSizeX As Short

        ''' <summary>
        ''' Vertical size of the subscript em box in design units, used to
        ''' scale the simulated subscript relative to the full em box size.
        ''' This the numerator of the scaling ratio where denominator is the
        ''' design units per em. If this member is zero, the font does not specify
        ''' a scale factor, and the client should use its own policy.
        ''' </summary>
        Public subscriptSizeY As Short

        ''' <summary>
        ''' Horizontal position of the superscript relative to the baseline origin.
        ''' This is typically positive (to the right) in italic/oblique fonts, and
        ''' zero in regular fonts.
        ''' </summary>
        Public superscriptPositionX As Short

        ''' <summary>
        ''' Vertical position of the superscript relative to the baseline.
        ''' This is typically positive.
        ''' </summary>
        Public superscriptPositionY As Short

        ''' <summary>
        ''' Horizontal size of the superscript em box in design units, used to
        ''' scale the simulated superscript relative to the full em box size.
        ''' This the numerator of the scaling ratio where denominator is the
        ''' design units per em. If this member is zero, the font does not specify
        ''' a scale factor, and the client should use its own policy.
        ''' </summary>
        Public superscriptSizeX As Short

        ''' <summary>
        ''' Vertical size of the superscript em box in design units, used to
        ''' scale the simulated superscript relative to the full em box size.
        ''' This the numerator of the scaling ratio where denominator is the
        ''' design units per em. If this member is zero, the font does not specify
        ''' a scale factor, and the client should use its own policy.
        ''' </summary>
        Public superscriptSizeY As Short

        ''' <summary>
        ''' Indicates that the ascent, descent, and lineGap are based on newer 
        ''' 'typographic' values in the font, rather than legacy values.
        ''' </summary>
        Public hasTypographicMetrics As Boolean
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DWRITE_CARET_METRICS
        ''' <summary>
        ''' Vertical rise of the caret. Rise / Run yields the caret angle.
        ''' Rise = 1 for perfectly upright fonts (non-italic).
        ''' </summary>
        Public slopeRise As Short

        ''' <summary>
        ''' Horizontal run of th caret. Rise / Run yields the caret angle.
        ''' Run = 0 for perfectly upright fonts (non-italic).
        ''' </summary>
        Public slopeRun As Short

        ''' <summary>
        ''' Horizontal offset of the caret along the baseline for good appearance.
        ''' Offset = 0 for perfectly upright fonts (non-italic).
        ''' </summary>
        Public offset As Short
    End Structure ' = 2 for text
    ' = 3 for script
    ' = 4 for decorative
    ' = 5 for symbol

    <StructLayout(LayoutKind.Explicit)>
    Public Structure DWRITE_PANOSE
        <FieldOffset(0)>
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=10)>
        Public Values() As Byte

        <FieldOffset(0)>
        Public FamilyKind As Byte

        <FieldOffset(0)>
        Public Text As PanoseText

        <FieldOffset(0)>
        Public Script As PanoseScript

        <FieldOffset(0)>
        Public Decorative As PanoseDecorative

        <FieldOffset(0)>
        Public Symbol As PanoseSymbol
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure PanoseText
        Public familyKind As Byte   ' = 2 for text
        Public serifStyle As Byte
        Public weight As Byte
        Public proportion As Byte
        Public contrast As Byte
        Public strokeVariation As Byte
        Public armStyle As Byte
        Public letterform As Byte
        Public midline As Byte
        Public xHeight As Byte
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure PanoseScript
        Public familyKind As Byte   ' = 3 for script
        Public toolKind As Byte
        Public weight As Byte
        Public spacing As Byte
        Public aspectRatio As Byte
        Public contrast As Byte
        Public scriptTopology As Byte
        Public scriptForm As Byte
        Public finials As Byte
        Public xAscent As Byte
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure PanoseDecorative
        Public familyKind As Byte   ' = 4 for decorative
        Public decorativeClass As Byte
        Public weight As Byte
        Public aspect As Byte
        Public contrast As Byte
        Public serifVariant As Byte
        Public fill As Byte
        Public lining As Byte
        Public decorativeTopology As Byte
        Public characterRange As Byte
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure PanoseSymbol
        Public familyKind As Byte   ' = 5 for symbol
        Public symbolKind As Byte
        Public weight As Byte
        Public spacing As Byte
        Public aspectRatioAndContrast As Byte
        Public aspectRatio94 As Byte
        Public aspectRatio119 As Byte
        Public aspectRatio157 As Byte
        Public aspectRatio163 As Byte
        Public aspectRatio211 As Byte
    End Structure




    Public Enum DWRITE_OUTLINE_THRESHOLD
        DWRITE_OUTLINE_THRESHOLD_ANTIALIASED
        DWRITE_OUTLINE_THRESHOLD_ALIASED
    End Enum

    <StructLayout(LayoutKind.Sequential, Pack:=4)>
    Public Structure D3DCOLORVALUE
        Public r As Single
        Public g As Single
        Public b As Single
        Public a As Single
    End Structure

    <StructLayout(LayoutKind.Sequential, Pack:=4)>
    Public Structure DWRITE_COLOR_F
        Public r As Single
        Public g As Single
        Public b As Single
        Public a As Single
    End Structure

    Public Enum DWRITE_TEXT_ANTIALIAS_MODE
        ''' <summary>
        ''' ClearType antialiasing computes coverage independently for the red, green, and blue
        ''' color elements of each pixel. This allows for more detail than conventional antialiasing.
        ''' However, because there is no one alpha value for each pixel, ClearType is not suitable
        ''' rendering text onto a transparent intermediate bitmap.
        ''' </summary>
        DWRITE_TEXT_ANTIALIAS_MODE_CLEARTYPE

        ''' <summary>
        ''' Grayscale antialiasing computes one coverage value for each pixel. Because the alpha
        ''' value of each pixel is well-defined, text can be rendered onto a transparent bitmap, 
        ''' which can then be composited with other content. Note that grayscale rendering with
        ''' IDWriteBitmapRenderTarget1 uses premultiplied alpha.
        ''' </summary>
        DWRITE_TEXT_ANTIALIAS_MODE_GRAYSCALE
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DWRITE_COLOR_GLYPH_RUN
        ''' <summary>
        ''' Glyph run to render.
        ''' </summary>
        Public glyphRun As DWRITE_GLYPH_RUN

        ''' <summary>
        ''' Optional glyph run description.
        ''' </summary>
        Public glyphRunDescription As DWRITE_GLYPH_RUN_DESCRIPTION

        ''' <summary>
        ''' Location at which to draw this glyph run.
        ''' </summary>
        Public baselineOriginX As Single
        Public baselineOriginY As Single

        ''' <summary>
        ''' Color to use for this layer, if any. This is the same color that
        ''' IDWriteFontFace2::GetPaletteEntries would return for the current
        ''' palette index if the paletteIndex member is less than 0xFFFF. If
        ''' the paletteIndex member is 0xFFFF then there is no associated
        ''' palette entry, this member is set to { 0, 0, 0, 0 }, and the client
        ''' should use the current foreground brush.
        ''' </summary>
        ''' 
        Public runColor As DWRITE_COLOR_F

        ''' <summary>
        ''' Zero-based index of this layer's color entry in the current color
        ''' palette, or 0xFFFF if this layer is to be rendered using 
        ''' the current foreground brush.
        ''' </summary>
        Public paletteIndex As UShort
    End Structure

    Public Enum DWRITE_FONT_PROPERTY_ID
        ''' <summary>
        ''' Unspecified font property identifier.
        ''' </summary>
        DWRITE_FONT_PROPERTY_ID_NONE

        ''' <summary>
        ''' Family name for the weight-stretch-style model.
        ''' </summary>
        DWRITE_FONT_PROPERTY_ID_WEIGHT_STRETCH_STYLE_FAMILY_NAME

        ''' <summary>
        ''' Family name preferred by the designer. This enables font designers to group more than four fonts in a single family without losing compatibility with
        ''' GDI. This name is typically only present if it differs from the GDI-compatible family name.
        ''' </summary>
        DWRITE_FONT_PROPERTY_ID_TYPOGRAPHIC_FAMILY_NAME

        ''' <summary>
        ''' Face name of the for the weight-stretch-style (e.g., Regular or Bold).
        ''' </summary>
        DWRITE_FONT_PROPERTY_ID_WEIGHT_STRETCH_STYLE_FACE_NAME

        ''' <summary>
        ''' The full name of the font, e.g. "Arial Bold", from name id 4 in the name table.
        ''' </summary>
        DWRITE_FONT_PROPERTY_ID_FULL_NAME

        ''' <summary>
        ''' GDI-compatible family name. Because GDI allows a maximum of four fonts per family, fonts in the same family may have different GDI-compatible family names
        ''' (e.g., "Arial", "Arial Narrow", "Arial Black").
        ''' </summary>
        DWRITE_FONT_PROPERTY_ID_WIN32_FAMILY_NAME

        ''' <summary>
        ''' The postscript name of the font, e.g. "GillSans-Bold" from name id 6 in the name table.
        ''' </summary>
        DWRITE_FONT_PROPERTY_ID_POSTSCRIPT_NAME

        ''' <summary>
        ''' Script/language tag to identify the scripts or languages that the font was
        ''' primarily designed to support.
        ''' </summary>
        ''' <remarks>
        ''' The design script/language tag is meant to be understood from the perspective of
        ''' users. For example, a font is considered designed for English if it is considered
        ''' useful for English users. Note that this is different from what a font might be
        ''' capable of supporting. For example, the Meiryo font was primarily designed for
        ''' Japanese users. While it is capable of displaying English well, it was not
        ''' meant to be offered for the benefit of non-Japanese-speaking English users.
        '''
        ''' As another example, a font designed for Chinese may be capable of displaying
        ''' Japanese text, but would likely look incorrect to Japanese users.
        ''' 
        ''' The valid values for this property are "ScriptLangTag" values. These are adapted
        ''' from the IETF BCP 47 specification, "Tags for Identifying Languages" (see
        ''' http://tools.ietf.org/html/bcp47). In a BCP 47 language tag, a language subtag
        ''' element is mandatory and other subtags are optional. In a ScriptLangTag, a
        ''' script subtag is mandatory and other subtags are option. The following
        ''' augmented BNF syntax, adapted from BCP 47, is used:
        ''' 
        '''     ScriptLangTag = [language "-"]
        '''                     script
        '''                     ["-" region]
        '''                     *("-" variant)
        '''                     *("-" extension)
        '''                     ["-" privateuse]
        ''' 
        ''' The expansion of the elements and the intended semantics associated with each
        ''' are as defined in BCP 47. Script subtags are taken from ISO 15924. At present,
        ''' no extensions are defined, and any extension should be ignored. Private use
        ''' subtags are defined by private agreement between the source and recipient and
        ''' may be ignored.
        ''' 
        ''' Subtags must be valid for use in BCP 47 and contained in the Language Subtag
        ''' Registry maintained by IANA. (See
        ''' http://www.iana.org/assignments/language-subtag-registry/language-subtag-registry
        ''' and section 3 of BCP 47 for details.
        ''' 
        ''' Any ScriptLangTag value not conforming to these specifications is ignored.
        ''' 
        ''' Examples:
        '''   "Latn" denotes Latin script (and any language or writing system using Latin)
        '''   "Cyrl" denotes Cyrillic script
        '''   "sr-Cyrl" denotes Cyrillic script as used for writing the Serbian language;
        '''       a font that has this property value may not be suitable for displaying
        '''       text in Russian or other languages written using Cyrillic script
        '''   "Jpan" denotes Japanese writing (Han + Hiragana + Katakana)
        '''
        ''' When passing this property to GetPropertyValues, use the overload which does
        ''' not take a language parameter, since this property has no specific language.
        ''' </remarks>
        DWRITE_FONT_PROPERTY_ID_DESIGN_SCRIPT_LANGUAGE_TAG

        ''' <summary>
        ''' Script/language tag to identify the scripts or languages that the font declares
        ''' it is able to support.
        ''' </summary>
        DWRITE_FONT_PROPERTY_ID_SUPPORTED_SCRIPT_LANGUAGE_TAG

        ''' <summary>
        ''' Semantic tag to describe the font (e.g. Fancy, Decorative, Handmade, Sans-serif, Swiss, Pixel, Futuristic).
        ''' </summary>
        DWRITE_FONT_PROPERTY_ID_SEMANTIC_TAG

        ''' <summary>
        ''' Weight of the font represented as a decimal string in the range 1-999.
        ''' </summary>
        ''' <remark>
        ''' This enum is discouraged for use with IDWriteFontSetBuilder2 in favor of the more generic font axis
        ''' DWRITE_FONT_AXIS_TAG_WEIGHT which supports higher precision and range.
        ''' </remark>
        DWRITE_FONT_PROPERTY_ID_WEIGHT

        ''' <summary>
        ''' Stretch of the font represented as a decimal string in the range 1-9.
        ''' </summary>
        ''' <remark>
        ''' This enum is discouraged for use with IDWriteFontSetBuilder2 in favor of the more generic font axis
        ''' DWRITE_FONT_AXIS_TAG_WIDTH which supports higher precision and range.
        ''' </remark>
        DWRITE_FONT_PROPERTY_ID_STRETCH

        ''' <summary>
        ''' Style of the font represented as a decimal string in the range 0-2.
        ''' </summary>
        ''' <remark>
        ''' This enum is discouraged for use with IDWriteFontSetBuilder2 in favor of the more generic font axes
        ''' DWRITE_FONT_AXIS_TAG_SLANT and DWRITE_FONT_AXIS_TAG_ITAL.
        ''' </remark>
        DWRITE_FONT_PROPERTY_ID_STYLE

        ''' <summary>
        ''' Face name preferred by the designer. This enables font designers to group more than four fonts in a single
        ''' family without losing compatibility with GDI.
        ''' </summary>
        DWRITE_FONT_PROPERTY_ID_TYPOGRAPHIC_FACE_NAME

        ''' <summary>
        ''' Total number of properties for NTDDI_WIN10 (IDWriteFontSet).
        ''' </summary>
        ''' <remarks>
        ''' DWRITE_FONT_PROPERTY_ID_TOTAL cannot be used as a property ID.
        ''' </remarks>
        DWRITE_FONT_PROPERTY_ID_TOTAL = DWRITE_FONT_PROPERTY_ID_STYLE + 1

        ''' <summary>
        ''' Total number of properties for NTDDI_WIN10_RS3 (IDWriteFontSet1).
        ''' </summary>
        DWRITE_FONT_PROPERTY_ID_TOTAL_RS3 = DWRITE_FONT_PROPERTY_ID_TYPOGRAPHIC_FACE_NAME + 1

        ' Obsolete aliases kept to avoid breaking existing code.
        DWRITE_FONT_PROPERTY_ID_PREFERRED_FAMILY_NAME = DWRITE_FONT_PROPERTY_ID_TYPOGRAPHIC_FAMILY_NAME
        DWRITE_FONT_PROPERTY_ID_FAMILY_NAME = DWRITE_FONT_PROPERTY_ID_WEIGHT_STRETCH_STYLE_FAMILY_NAME
        DWRITE_FONT_PROPERTY_ID_FACE_NAME = DWRITE_FONT_PROPERTY_ID_WEIGHT_STRETCH_STYLE_FACE_NAME
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DWRITE_FONT_PROPERTY
        ''' <summary>
        ''' Specifies the requested font property, such as DWRITE_FONT_PROPERTY_ID_FAMILY_NAME.
        ''' </summary>
        Public propertyId As DWRITE_FONT_PROPERTY_ID

        ''' <summary>
        ''' Specifies the property value, such as "Segoe UI".
        ''' </summary>
        Public propertyValue As String

        ''' <summary>
        ''' Specifies the language / locale to use, such as "en-US". 
        ''' </summary>
        ''' <remarks>
        ''' When passing property information to AddFontFaceReference, localeName indicates
        ''' the language of the property value. BCP 47 language tags should be used. If a
        ''' property value is inherently non-linguistic, this can be left empty.
        '''
        ''' When used for font set filtering, leave this empty: a match will be found
        ''' regardless of language associated with property values.
        ''' </remarks>
        Public localeName As String
    End Structure

    Public Enum DWRITE_LOCALITY
        ''' <summary>
        ''' The resource is remote, and information is unknown yet, including the file size and date.
        ''' Attempting to create a font or file stream will fail until locality becomes at least partial.
        ''' </summary>
        DWRITE_LOCALITY_REMOTE

        ''' <summary>
        ''' The resource is partially local, meaning you can query the size and date of the file
        ''' stream, and you may be able to create a font face and retrieve the particular glyphs
        ''' for metrics and drawing, but not all the glyphs will be present.
        ''' </summary>
        DWRITE_LOCALITY_PARTIAL

        ''' <summary>
        ''' The resource is completely local, and all font functions can be called
        ''' without concern of missing data or errors related to network connectivity.
        ''' </summary>
        DWRITE_LOCALITY_LOCAL
    End Enum

    Public Enum DWRITE_RENDERING_MODE1
        ''' <summary>
        ''' Specifies that the rendering mode is determined automatically based on the font and size.
        ''' </summary>
        DWRITE_RENDERING_MODE1_DEFAULT = DWRITE_RENDERING_MODE.DWRITE_RENDERING_MODE_DEFAULT

        ''' <summary>
        ''' Specifies that no antialiasing is performed. Each pixel is either set to the foreground 
        ''' color of the text or retains the color of the background.
        ''' </summary>
        DWRITE_RENDERING_MODE1_ALIASED = DWRITE_RENDERING_MODE.DWRITE_RENDERING_MODE_ALIASED

        ''' <summary>
        ''' Specifies that antialiasing is performed in the horizontal direction and the appearance
        ''' of glyphs is layout-compatible with GDI using CLEARTYPE_QUALITY. Use DWRITE_MEASURING_MODE_GDI_CLASSIC 
        ''' to get glyph advances. The antialiasing may be either ClearType or grayscale depending on
        ''' the text antialiasing mode.
        ''' </summary>
        DWRITE_RENDERING_MODE1_GDI_CLASSIC = DWRITE_RENDERING_MODE.DWRITE_RENDERING_MODE_GDI_CLASSIC

        ''' <summary>
        ''' Specifies that antialiasing is performed in the horizontal direction and the appearance
        ''' of glyphs is layout-compatible with GDI using CLEARTYPE_NATURAL_QUALITY. Glyph advances
        ''' are close to the font design advances, but are still rounded to whole pixels. Use
        ''' DWRITE_MEASURING_MODE_GDI_NATURAL to get glyph advances. The antialiasing may be either
        ''' ClearType or grayscale depending on the text antialiasing mode.
        ''' </summary>
        DWRITE_RENDERING_MODE1_GDI_NATURAL = DWRITE_RENDERING_MODE.DWRITE_RENDERING_MODE_GDI_NATURAL

        ''' <summary>
        ''' Specifies that antialiasing is performed in the horizontal direction. This rendering
        ''' mode allows glyphs to be positioned with subpixel precision and is therefore suitable
        ''' for natural (i.e., resolution-independent) layout. The antialiasing may be either
        ''' ClearType or grayscale depending on the text antialiasing mode.
        ''' </summary>
        DWRITE_RENDERING_MODE1_NATURAL = DWRITE_RENDERING_MODE.DWRITE_RENDERING_MODE_NATURAL

        ''' <summary>
        ''' Similar to natural mode except that antialiasing is performed in both the horizontal
        ''' and vertical directions. This is typically used at larger sizes to make curves and
        ''' diagonal lines look smoother. The antialiasing may be either ClearType or grayscale
        ''' depending on the text antialiasing mode.
        ''' </summary>
        DWRITE_RENDERING_MODE1_NATURAL_SYMMETRIC = DWRITE_RENDERING_MODE.DWRITE_RENDERING_MODE_NATURAL_SYMMETRIC

        ''' <summary>
        ''' Specifies that rendering should bypass the rasterizer and use the outlines directly. 
        ''' This is typically used at very large sizes.
        ''' </summary>
        DWRITE_RENDERING_MODE1_OUTLINE = DWRITE_RENDERING_MODE.DWRITE_RENDERING_MODE_OUTLINE

        ''' <summary>
        ''' Similar to natural symmetric mode except that when possible, text should be rasterized
        ''' in a downsampled form.
        ''' </summary>
        DWRITE_RENDERING_MODE1_NATURAL_SYMMETRIC_DOWNSAMPLED
    End Enum

    Public Enum DWRITE_GLYPH_IMAGE_FORMATS
        ''' <summary>
        ''' Indicates no data is available for this glyph.
        ''' </summary>
        DWRITE_GLYPH_IMAGE_FORMATS_NONE = &H00000000

        ''' <summary>
        ''' The glyph has TrueType outlines.
        ''' </summary>
        DWRITE_GLYPH_IMAGE_FORMATS_TRUETYPE = &H00000001

        ''' <summary>
        ''' The glyph has CFF outlines.
        ''' </summary>
        DWRITE_GLYPH_IMAGE_FORMATS_CFF = &H00000002

        ''' <summary>
        ''' The glyph has multilayered COLR data.
        ''' </summary>
        DWRITE_GLYPH_IMAGE_FORMATS_COLR = &H00000004

        ''' <summary>
        ''' The glyph has SVG outlines as standard XML.
        ''' </summary>
        ''' <remarks>
        ''' Fonts may store the content gzip'd rather than plain text,
        ''' indicated by the first two bytes as gzip header {0x1F 0x8B}.
        ''' </remarks>
        DWRITE_GLYPH_IMAGE_FORMATS_SVG = &H00000008

        ''' <summary>
        ''' The glyph has PNG image data, with standard PNG IHDR.
        ''' </summary>
        DWRITE_GLYPH_IMAGE_FORMATS_PNG = &H00000010

        ''' <summary>
        ''' The glyph has JPEG image data, with standard JIFF SOI header.
        ''' </summary>
        DWRITE_GLYPH_IMAGE_FORMATS_JPEG = &H00000020

        ''' <summary>
        ''' The glyph has TIFF image data.
        ''' </summary>
        DWRITE_GLYPH_IMAGE_FORMATS_TIFF = &H00000040

        ''' <summary>
        ''' The glyph has raw 32-bit premultiplied BGRA data.
        ''' </summary>
        DWRITE_GLYPH_IMAGE_FORMATS_PREMULTIPLIED_B8G8R8A8 = &H00000080
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DWRITE_GLYPH_IMAGE_DATA
        ''' <summary>
        ''' Pointer to the glyph data, be it SVG, PNG, JPEG, TIFF.
        ''' </summary>
        Public imageData As IntPtr

        ''' <summary>
        ''' Size of glyph data in bytes.
        ''' </summary>
        Public imageDataSize As UInteger

        ''' <summary>
        ''' Unique identifier for the glyph data. Clients may use this to cache a parsed/decompressed
        ''' version and tell whether a repeated call to the same font returns the same data.
        ''' </summary>
        Public uniqueDataId As UInteger

        ''' <summary>
        ''' Pixels per em of the returned data. For non-scalable raster data (PNG/TIFF/JPG), this can be larger
        ''' or smaller than requested from GetGlyphImageData when there isn't an exact match.
        ''' For scaling intermediate sizes, use: desired pixels per em * font em size / actual pixels per em.
        ''' </summary>
        Public pixelsPerEm As UInteger

        ''' <summary>
        ''' Size of image when the format is pixel data.
        ''' </summary>
        Public pixelSize As D2D1_SIZE_U

        ''' <summary>
        ''' Left origin along the horizontal Roman baseline.
        ''' </summary>
        Public horizontalLeftOrigin As D2D1_POINT_2L

        ''' <summary>
        ''' Right origin along the horizontal Roman baseline.
        ''' </summary>
        Public horizontalRightOrigin As D2D1_POINT_2L

        ''' <summary>
        ''' Top origin along the vertical central baseline.
        ''' </summary>
        Public verticalTopOrigin As D2D1_POINT_2L

        ''' <summary>
        ''' Bottom origin along vertical central baseline.
        ''' </summary>
        Public verticalBottomOrigin As D2D1_POINT_2L
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_SIZE_U
        Public width As UInteger
        Public height As UInteger
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_POINT_2L
        Public x As Integer
        Public y As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DWRITE_COLOR_GLYPH_RUN1 ': DWRITE_COLOR_GLYPH_RUN
        ''' <summary>
        ''' Glyph run to render.
        ''' </summary>
        Public glyphRun As DWRITE_GLYPH_RUN
        'public IntPtr glyphRun;

        ''' <summary>
        ''' Optional glyph run description.
        ''' </summary>
        'public DWRITE_GLYPH_RUN_DESCRIPTION glyphRunDescription;
        Public glyphRunDescription As IntPtr

        ''' <summary>
        ''' Location at which to draw this glyph run.
        ''' </summary>
        Public baselineOriginX As Single
        Public baselineOriginY As Single

        ''' <summary>
        ''' Color to use for this layer, if any. This is the same color that
        ''' IDWriteFontFace2::GetPaletteEntries would return for the current
        ''' palette index if the paletteIndex member is less than 0xFFFF. If
        ''' the paletteIndex member is 0xFFFF then there is no associated
        ''' palette entry, this member is set to { 0, 0, 0, 0 }, and the client
        ''' should use the current foreground brush.
        ''' </summary>
        ''' 
        Public runColor As DWRITE_COLOR_F

        ''' <summary>
        ''' Zero-based index of this layer's color entry in the current color
        ''' palette, or 0xFFFF if this layer is to be rendered using 
        ''' the current foreground brush.
        ''' </summary>
        Public paletteIndex As UShort

        ''' <summary>
        ''' Type of glyph image format for this color run. Exactly one type will be set since
        ''' TranslateColorGlyphRun has already broken down the run into separate parts.
        ''' </summary>
        Public glyphImageFormat As DWRITE_GLYPH_IMAGE_FORMATS

        ''' <summary>
        ''' Measuring mode to use for this glyph run.
        ''' </summary>
        Public measuringMode As DWRITE_MEASURING_MODE
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DWRITE_FILE_FRAGMENT
        ''' <summary>
        ''' Starting offset of the fragment from the beginning of the file.
        ''' </summary>
        Public fileOffset As ULong

        ''' <summary>
        ''' Size of the file fragment, in bytes.
        ''' </summary>
        Public fragmentSize As ULong
    End Structure

    Public Enum DWRITE_CONTAINER_TYPE
        DWRITE_CONTAINER_TYPE_UNKNOWN
        DWRITE_CONTAINER_TYPE_WOFF
        DWRITE_CONTAINER_TYPE_WOFF2
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DWRITE_FONT_AXIS_VALUE
        ''' <summary>
        ''' Four character identifier of the font axis (weight, width, slant, italic...).
        ''' </summary>
        Public axisTag As DWRITE_FONT_AXIS_TAG

        ''' <summary>
        ''' Value for the given axis, with the meaning and range depending on the axis semantics.
        ''' Certain well known axes have standard ranges and defaults, such as weight (1..1000, default=400),
        ''' width (>0, default=100), slant (-90..90, default=-20), and italic (0 or 1).
        ''' </summary>
        Public value As Single
    End Structure

    Public Enum DWRITE_FONT_AXIS_TAG As UInteger
        DWRITE_FONT_AXIS_TAG_WEIGHT = CUInt(Microsoft.VisualBasic.AscW("t"c)) << 24 Or CUInt(Microsoft.VisualBasic.AscW("h"c)) << 16 Or CUInt(Microsoft.VisualBasic.AscW("g"c)) << 8 Or CUInt(Microsoft.VisualBasic.AscW("w"c))
        DWRITE_FONT_AXIS_TAG_WIDTH = CUInt(Microsoft.VisualBasic.AscW("h"c)) << 24 Or CUInt(Microsoft.VisualBasic.AscW("t"c)) << 16 Or CUInt(Microsoft.VisualBasic.AscW("d"c)) << 8 Or CUInt(Microsoft.VisualBasic.AscW("w"c))
        DWRITE_FONT_AXIS_TAG_SLANT = CUInt(Microsoft.VisualBasic.AscW("t"c)) << 24 Or CUInt(Microsoft.VisualBasic.AscW("n"c)) << 16 Or CUInt(Microsoft.VisualBasic.AscW("l"c)) << 8 Or CUInt(Microsoft.VisualBasic.AscW("s"c))
        DWRITE_FONT_AXIS_TAG_OPTICAL_SIZE = CUInt(Microsoft.VisualBasic.AscW("z"c)) << 24 Or CUInt(Microsoft.VisualBasic.AscW("s"c)) << 16 Or CUInt(Microsoft.VisualBasic.AscW("p"c)) << 8 Or CUInt(Microsoft.VisualBasic.AscW("o"c))
        DWRITE_FONT_AXIS_TAG_ITALIC = CUInt(Microsoft.VisualBasic.AscW("l"c)) << 24 Or CUInt(Microsoft.VisualBasic.AscW("a"c)) << 16 Or CUInt(Microsoft.VisualBasic.AscW("t"c)) << 8 Or CUInt(Microsoft.VisualBasic.AscW("i"c))
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DWRITE_FONT_AXIS_RANGE
        ''' <summary>
        ''' Four character identifier of the font axis (weight, width, slant, italic...).
        ''' </summary>
        Public axisTag As DWRITE_FONT_AXIS_TAG

        ''' <summary>
        ''' Minimum value supported by this axis.
        ''' </summary>
        Public minValue As Single

        ''' <summary>
        ''' Maximum value supported by this axis. The maximum can equal the minimum.
        ''' </summary>
        Public maxValue As Single
    End Structure

    Public Enum DWRITE_FONT_AXIS_ATTRIBUTES
        ''' <summary>
        ''' No attributes.
        ''' </summary>
        DWRITE_FONT_AXIS_ATTRIBUTES_NONE = &H0000

        ''' <summary>
        ''' This axis is implemented as a variation axis in a variable font, with a continuous range of
        ''' values, such as a range of weights from 100..900. Otherwise it is either a static axis that
        ''' holds a single point, or it has a range but doesn't vary, such as optical size in the Skia
        ''' Heading font which covers a range of points but doesn't interpolate any new glyph outlines.
        ''' </summary>
        DWRITE_FONT_AXIS_ATTRIBUTES_VARIABLE = &H0001

        ''' <summary>
        ''' This axis is recommended to be remain hidden in user interfaces. The font developer may
        ''' recommend this if an axis is intended to be accessed only programmatically, or is meant for
        ''' font-internal or font-developer use only. The axis may be exposed in lower-level font
        ''' inspection utilities, but should not be exposed in common or even advanced-mode user
        ''' interfaces in content-authoring apps.
        ''' </summary>
        DWRITE_FONT_AXIS_ATTRIBUTES_HIDDEN = &H0002
    End Enum

    Public Enum DWRITE_FONT_FAMILY_MODEL
        ''' <summary>
        ''' Families are grouped by the typographic family name preferred by the font author. The family can contain as
        ''' many face as the font author wants.
        ''' This corresponds to the DWRITE_FONT_PROPERTY_ID_TYPOGRAPHIC_FAMILY_NAME.
        ''' </summary>
        DWRITE_FONT_FAMILY_MODEL_TYPOGRAPHIC

        ''' <summary>
        ''' Families are grouped by the weight-stretch-style family name, where all faces that differ only by those three
        ''' axes are grouped into the same family, but any other axes go into a distinct family. For example, the Sitka
        ''' family with six different optical sizes yields six separate families (Sitka Caption, Display, Text, Subheading,
        ''' Heading, Banner...). This corresponds to the DWRITE_FONT_PROPERTY_ID_WEIGHT_STRETCH_STYLE_FAMILY_NAME.
        ''' </summary>
        DWRITE_FONT_FAMILY_MODEL_WEIGHT_STRETCH_STYLE
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DWRITE_LINE_SPACING
        ''' <summary>
        ''' Method used to determine line spacing.
        ''' </summary>
        Public method As DWRITE_LINE_SPACING_METHOD

        ''' <summary>
        ''' Spacing between lines.
        ''' The interpretation of this parameter depends upon the line spacing method, as follows:
        ''' - default line spacing: ignored
        ''' - uniform line spacing: explicit distance in DIPs between lines
        ''' - proportional line spacing: a scaling factor to be applied to the computed line height; 
        '''   for each line, the height of the line is computed as for default line spacing, and the scaling factor is applied to that value.
        ''' </summary>
        Public height As Single

        ''' <summary>
        ''' Distance from top of line to baseline. 
        ''' The interpretation of this parameter depends upon the line spacing method, as follows:
        ''' - default line spacing: ignored
        ''' - uniform line spacing: explicit distance in DIPs from the top of the line to the baseline
        ''' - proportional line spacing: a scaling factor applied to the computed baseline; for each line, 
        '''   the baseline distance is computed as for default line spacing, and the scaling factor is applied to that value.
        ''' </summary>
        Public baseline As Single

        ''' <summary>
        ''' Proportion of the entire leading distributed before the line. The allowed value is between 0 and 1.0. The remaining
        ''' leading is distributed after the line. It is ignored for the default and uniform line spacing methods.
        ''' The leading that is available to distribute before or after the line depends on the values of the height and
        ''' baseline parameters.
        ''' </summary>
        Public leadingBefore As Single

        ''' <summary>
        ''' Specify whether DWRITE_FONT_METRICS::lineGap value should be part of the line metrics.
        ''' </summary>
        Public fontLineGapUsage As DWRITE_FONT_LINE_GAP_USAGE
    End Structure

    Public Enum DWRITE_FONT_LINE_GAP_USAGE
        ''' <summary>
        ''' The usage of the font line gap depends on the method used for text layout.
        ''' </summary>
        DWRITE_FONT_LINE_GAP_USAGE_DEFAULT

        ''' <summary>
        ''' The font line gap is excluded from line spacing
        ''' </summary>
        DWRITE_FONT_LINE_GAP_USAGE_DISABLED

        ''' <summary>
        ''' The font line gap is included in line spacing
        ''' </summary>
        DWRITE_FONT_LINE_GAP_USAGE_ENABLED
    End Enum

    Public Enum DWRITE_AUTOMATIC_FONT_AXES
        ''' <summary>
        ''' No axes are automatically applied.
        ''' </summary>
        DWRITE_AUTOMATIC_FONT_AXES_NONE = &H0000

        ''' <summary>
        ''' Automatically pick an appropriate optical value based on the font size (via SetFontSize) when no value is
        ''' specified via DWRITE_FONT_AXIS_TAG_OPTICAL_SIZE. Callers can still explicitly apply the 'opsz' value over
        ''' text ranges via SetFontAxisValues, which take priority.
        ''' </summary>
        DWRITE_AUTOMATIC_FONT_AXES_OPTICAL_SIZE = &H0001
    End Enum

    Public Enum DWRITE_FONT_SOURCE_TYPE
        ''' <summary>
        ''' The font source is unknown or is not any of the other defined font source types.
        ''' </summary>
        DWRITE_FONT_SOURCE_TYPE_UNKNOWN

        ''' <summary>
        ''' The font source is a font file, which is installed for all users on the device.
        ''' </summary>
        DWRITE_FONT_SOURCE_TYPE_PER_MACHINE

        ''' <summary>
        ''' The font source is a font file, which is installed for the current user.
        ''' </summary>
        DWRITE_FONT_SOURCE_TYPE_PER_USER

        ''' <summary>
        ''' The font source is an APPX package, which includes one or more font files.
        ''' The font source name is the full name of the package.
        ''' </summary>
        DWRITE_FONT_SOURCE_TYPE_APPX_PACKAGE

        ''' <summary>
        ''' The font source is a font provider for downloadable fonts.
        ''' </summary>
        DWRITE_FONT_SOURCE_TYPE_REMOTE_FONT_PROVIDER
    End Enum

    Public Enum DWRITE_HRESULT As Integer
        '
        ' MessageId: DWRITE_E_FILEFORMAT
        '
        ' MessageText:
        '
        ' Indicates an error in an input file such as a font file.
        '
        DWRITE_E_FILEFORMAT = &H88985000

        '
        ' MessageId: DWRITE_E_UNEXPECTED
        '
        ' MessageText:
        '
        ' Indicates an error originating in DirectWrite code, which is not expected to occur but is safe to recover from.
        '
        DWRITE_E_UNEXPECTED = &H88985001

        '
        ' MessageId: DWRITE_E_NOFONT
        '
        ' MessageText:
        '
        ' Indicates the specified font does not exist.
        '
        DWRITE_E_NOFONT = &H88985002

        '
        ' MessageId: DWRITE_E_FILENOTFOUND
        '
        ' MessageText:
        '
        ' A font file could not be opened because the file, directory, network location, drive, or other storage location does not exist or is unavailable.
        '
        DWRITE_E_FILENOTFOUND = &H88985003

        '
        ' MessageId: DWRITE_E_FILEACCESS
        '
        ' MessageText:
        '
        ' A font file exists but could not be opened due to access denied, sharing violation, or similar error.
        '
        DWRITE_E_FILEACCESS = &H88985004

        '
        ' MessageId: DWRITE_E_FONTCOLLECTIONOBSOLETE
        '
        ' MessageText:
        '
        ' A font collection is obsolete due to changes in the system.
        '
        DWRITE_E_FONTCOLLECTIONOBSOLETE = &H88985005

        '
        ' MessageId: DWRITE_E_ALREADYREGISTERED
        '
        ' MessageText:
        '
        ' The given interface is already registered.
        '
        DWRITE_E_ALREADYREGISTERED = &H88985006

        '
        ' MessageId: DWRITE_E_CACHEFORMAT
        '
        ' MessageText:
        '
        ' The font cache contains invalid data.
        '
        DWRITE_E_CACHEFORMAT = &H88985007

        '
        ' MessageId: DWRITE_E_CACHEVERSION
        '
        ' MessageText:
        '
        ' A font cache file corresponds to a different version of DirectWrite.
        '
        DWRITE_E_CACHEVERSION = &H88985008

        '
        ' MessageId: DWRITE_E_UNSUPPORTEDOPERATION
        '
        ' MessageText:
        '
        ' The operation is not supported for this type of font.
        '
        DWRITE_E_UNSUPPORTEDOPERATION = &H88985009

        '
        ' MessageId: DWRITE_E_TEXTRENDERERINCOMPATIBLE
        '
        ' MessageText:
        '
        ' The version of the text renderer interface is not compatible.
        '
        DWRITE_E_TEXTRENDERERINCOMPATIBLE = &H8898500A

        '
        ' MessageId: DWRITE_E_FLOWDIRECTIONCONFLICTS
        '
        ' MessageText:
        '
        ' The flow direction conflicts with the reading direction. They must be perpendicular to each other.
        '
        DWRITE_E_FLOWDIRECTIONCONFLICTS = &H8898500B

        '
        ' MessageId: DWRITE_E_NOCOLOR
        '
        ' MessageText:
        '
        ' The font or glyph run does not contain any colored glyphs.
        '
        DWRITE_E_NOCOLOR = &H8898500C

        '
        ' MessageId: DWRITE_E_REMOTEFONT
        '
        ' MessageText:
        '
        ' A font resource could not be accessed because it is remote.
        '
        DWRITE_E_REMOTEFONT = &H8898500D

        '
        ' MessageId: DWRITE_E_DOWNLOADCANCELLED
        '
        ' MessageText:
        '
        ' A font download was canceled.
        '
        DWRITE_E_DOWNLOADCANCELLED = &H8898500E

        '
        ' MessageId: DWRITE_E_DOWNLOADFAILED
        '
        ' MessageText:
        '
        ' A font download failed.
        '
        DWRITE_E_DOWNLOADFAILED = &H8898500F

        '
        ' MessageId: DWRITE_E_TOOMANYDOWNLOADS
        '
        ' MessageText:
        '
        ' A font download request was not added or a download failed because there are too many active downloads.
        '
        DWRITE_E_TOOMANYDOWNLOADS = &H88985010
    End Enum
End Namespace
