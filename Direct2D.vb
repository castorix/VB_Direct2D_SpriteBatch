Imports System
Imports System.Runtime.InteropServices
Imports DXGI
Imports WIC
Imports GlobalStructures
Imports DXGI.DXGITools
Imports Direct2D.D2DTools

#If DWRITE
Imports DWrite
#End If

Namespace Global.Direct2D
    Public Enum D2D1_FACTORY_TYPE As UInteger
        '
        ' The resulting factory and derived resources may only be invoked serially.
        ' Reference counts on resources are interlocked, however, resource and render
        ' target state is not protected from multi-threaded access.
        '
        D2D1_FACTORY_TYPE_SINGLE_THREADED = 0
        '
        ' The resulting factory may be invoked from multiple threads. Returned resources
        ' use interlocked reference counting and their state is protected.
        '
        D2D1_FACTORY_TYPE_MULTI_THREADED = 1
        D2D1_FACTORY_TYPE_FORCE_DWORD = &HFFFFFFFFUI
    End Enum

    Public Enum D2D1_DEBUG_LEVEL As UInteger
        D2D1_DEBUG_LEVEL_NONE = 0
        D2D1_DEBUG_LEVEL_ERROR = 1
        D2D1_DEBUG_LEVEL_WARNING = 2
        D2D1_DEBUG_LEVEL_INFORMATION = 3
        D2D1_DEBUG_LEVEL_FORCE_DWORD = &HFFFFFFFFUI
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_FACTORY_OPTIONS
        Public debugLevel As D2D1_DEBUG_LEVEL
    End Structure

    Public Enum D2D1_RENDER_TARGET_TYPE
        '
        ' D2D is free to choose the render target type for the caller.
        '
        D2D1_RENDER_TARGET_TYPE_DEFAULT = 0
        '
        ' The render target will render using the CPU.
        '
        D2D1_RENDER_TARGET_TYPE_SOFTWARE = 1
        '
        ' The render target will render using the GPU.
        '
        D2D1_RENDER_TARGET_TYPE_HARDWARE = 2
        D2D1_RENDER_TARGET_TYPE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_ALPHA_MODE
        '
        ' Alpha mode should be determined implicitly. Some target surfaces do not supply
        ' or imply this information in which case alpha must be specified.
        '
        D2D1_ALPHA_MODE_UNKNOWN = 0
        '
        ' Treat the alpha as premultipled.
        '
        D2D1_ALPHA_MODE_PREMULTIPLIED = 1
        '
        ' Opacity is in the 'A' component only.
        '
        D2D1_ALPHA_MODE_STRAIGHT = 2
        '
        ' Ignore any alpha channel information.
        '
        D2D1_ALPHA_MODE_IGNORE = 3

        D2D1_ALPHA_MODE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_PIXEL_FORMAT
        Public format As DXGI_FORMAT
        Public alphaMode As D2D1_ALPHA_MODE
    End Structure

    Public Enum D2D1_RENDER_TARGET_USAGE
        D2D1_RENDER_TARGET_USAGE_NONE = &H0
        '
        ' Rendering will occur locally, if a terminal-services session is established, the
        ' bitmap updates will be sent to the terminal services client.
        '
        D2D1_RENDER_TARGET_USAGE_FORCE_BITMAP_REMOTING = &H1
        '
        ' The render target will allow a call to GetDC on the ID2D1GdiInteropRenderTarget
        ' interface. Rendering will also occur locally.
        '
        D2D1_RENDER_TARGET_USAGE_GDI_COMPATIBLE = &H2
        D2D1_RENDER_TARGET_USAGE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_FEATURE_LEVEL
        '
        ' The caller does not require a particular underlying D3D device level.
        '
        D2D1_FEATURE_LEVEL_DEFAULT = 0

        '
        ' The D3D device level is DX9 compatible.
        '
        D2D1_FEATURE_LEVEL_9 = D3D_FEATURE_LEVEL.D3D_FEATURE_LEVEL_9_1

        '
        ' The D3D device level is DX10 compatible.
        '
        D2D1_FEATURE_LEVEL_10 = D3D_FEATURE_LEVEL.D3D_FEATURE_LEVEL_10_0
        D2D1_FEATURE_LEVEL_FORCE_DWORD = &HFFFFFFFF
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_RENDER_TARGET_PROPERTIES
        Public type As D2D1_RENDER_TARGET_TYPE
        Public pixelFormat As D2D1_PIXEL_FORMAT
        Public dpiX As Single
        Public dpiY As Single
        Public usage As D2D1_RENDER_TARGET_USAGE
        Public minLevel As D2D1_FEATURE_LEVEL
    End Structure

    Public Enum D2D1_WINDOW_STATE
        D2D1_WINDOW_STATE_NONE = &H0
        D2D1_WINDOW_STATE_OCCLUDED = &H1
        D2D1_WINDOW_STATE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    <ComImport>
    <Guid("2cd906a8-12e2-11dc-9fed-001143a055f9")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1Brush
        Inherits ID2D1Resource
#Region "ID2D1Resource"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region

        <PreserveSig>
        Sub SetOpacity(opacity As Single)
        <PreserveSig>
        Sub SetTransform(transform As D2D1_MATRIX_3X2_F)
        <PreserveSig>
        Function GetOpacity() As Single
        <PreserveSig>
        Sub GetTransform(<Out> ByRef transform As D2D1_MATRIX_3X2_F_STRUCT)
    End Interface

    <ComImport>
    <Guid("2cd906aa-12e2-11dc-9fed-001143a055f9")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1BitmapBrush
        Inherits ID2D1Brush
#Region "ID2D1Brush"
#Region "ID2D1Resource"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region
        <PreserveSig>
        Overloads Sub SetOpacity(opacity As Single)
        <PreserveSig>
        Overloads Sub SetTransform(transform As D2D1_MATRIX_3X2_F)
        <PreserveSig>
        Overloads Function GetOpacity() As Single
        <PreserveSig>
        Overloads Sub GetTransform(<Out> ByRef transform As D2D1_MATRIX_3X2_F_STRUCT)
#End Region

        <PreserveSig>
        Sub SetExtendModeX(extendModeX As D2D1_EXTEND_MODE)
        <PreserveSig>
        Sub SetExtendModeY(extendModeY As D2D1_EXTEND_MODE)
        <PreserveSig>
        Sub SetInterpolationMode(interpolationMode As D2D1_BITMAP_INTERPOLATION_MODE)
        <PreserveSig>
        Sub SetBitmap(bitmap As ID2D1Bitmap)
        <PreserveSig>
        Sub GetExtendModeX(<Out> ByRef extendedModeX As D2D1_EXTEND_MODE)
        <PreserveSig>
        Sub GetExtendModeY(<Out> ByRef extendedModeY As D2D1_EXTEND_MODE)
        <PreserveSig>
        Sub GetInterpolationMode(<Out> ByRef interpolationMode As D2D1_BITMAP_INTERPOLATION_MODE)
        <PreserveSig>
        Sub GetBitmap(<Out> ByRef bitmap As ID2D1Bitmap)
    End Interface

    <ComImport>
    <Guid("2cd906ab-12e2-11dc-9fed-001143a055f9")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1LinearGradientBrush
        Inherits ID2D1Brush
#Region "ID2D1Brush"
#Region "ID2D1Resource"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region
        <PreserveSig>
        Overloads Sub SetOpacity(opacity As Single)
        <PreserveSig>
        Overloads Sub SetTransform(transform As D2D1_MATRIX_3X2_F)
        <PreserveSig>
        Overloads Function GetOpacity() As Single
        <PreserveSig>
        Overloads Sub GetTransform(<Out> ByRef transform As D2D1_MATRIX_3X2_F_STRUCT)
#End Region

        <PreserveSig>
        Sub SetStartPoint(startPoint As D2D1_POINT_2F)
        <PreserveSig>
        Sub SetEndPoint(endPoint As D2D1_POINT_2F)
        <PreserveSig>
        Sub GetStartPoint(<Out> ByRef startPoint As D2D1_POINT_2F)
        <PreserveSig>
        Sub GetEndPoint(<Out> ByRef endPoint As D2D1_POINT_2F)
        <PreserveSig>
        Sub GetGradientStopCollection(<Out> ByRef gradientStopCollection As ID2D1GradientStopCollection)
    End Interface

    <ComImport>
    <Guid("2cd906ac-12e2-11dc-9fed-001143a055f9")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1RadialGradientBrush
        Inherits ID2D1Brush
#Region "ID2D1Brush"
#Region "ID2D1Resource"
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region
        <PreserveSig>
        Overloads Sub SetOpacity(opacity As Single)
        <PreserveSig>
        Overloads Sub SetTransform(transform As D2D1_MATRIX_3X2_F)
        <PreserveSig>
        Overloads Function GetOpacity() As Single
        <PreserveSig>
        Overloads Sub GetTransform(<Out> ByRef transform As D2D1_MATRIX_3X2_F_STRUCT)
#End Region

        <PreserveSig>
        Sub SetCenter(ByRef center As D2D1_POINT_2F)
        <PreserveSig>
        Sub SetGradientOriginOffset(ByRef gradientOriginOffset As D2D1_POINT_2F)
        <PreserveSig>
        Sub SetRadiusX(radiusX As Single)
        <PreserveSig>
        Sub SetRadiusY(radiusY As Single)
        <PreserveSig>
        Sub GetCenter(<Out> ByRef center As D2D1_POINT_2F)
        <PreserveSig>
        Sub GetGradientOriginOffset(<Out> ByRef gradientOriginOffset As D2D1_POINT_2F)
        <PreserveSig>
        Function GetRadiusX() As Single
        <PreserveSig>
        Function GetRadiusY() As Single
        <PreserveSig>
        Sub GetGradientStopCollection(<Out> ByRef gradientStopCollection As ID2D1GradientStopCollection)
    End Interface

    <ComImport>
    <Guid("2cd9069d-12e2-11dc-9fed-001143a055f9")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1StrokeStyle
        Inherits ID2D1Resource
#Region "ID2D1Resource"
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region

        <PreserveSig>
        Sub GetStartCap(<Out> ByRef capStyle As D2D1_CAP_STYLE)
        <PreserveSig>
        Sub GetEndCap(<Out> ByRef capStyle As D2D1_CAP_STYLE)
        <PreserveSig>
        Sub GetDashCap(<Out> ByRef capStyle As D2D1_CAP_STYLE)
        <PreserveSig>
        Function GetMiterLimit() As Single
        <PreserveSig>
        Sub GetLineJoin(<Out> ByRef lineJoin As D2D1_LINE_JOIN)
        <PreserveSig>
        Function GetDashOffset() As Single
        <PreserveSig>
        Sub GetDashStyle(<Out> ByRef dashStyle As D2D1_DASH_STYLE)
        <PreserveSig>
        Function GetDashesCount() As UInteger
        <PreserveSig>
        Sub GetDashes(<Out> ByRef dashes As Single, dashesCount As UInteger)
    End Interface

    Public Enum D2D1_GEOMETRY_RELATION
        '
        ' The relation between the geometries couldn't be determined. This value is never
        ' returned by any D2D method.
        '
        D2D1_GEOMETRY_RELATION_UNKNOWN = 0
        '
        ' The two geometries do not intersect at all.
        '
        D2D1_GEOMETRY_RELATION_DISJOINT = 1
        '
        ' The passed in geometry is entirely contained by the object.            //
        D2D1_GEOMETRY_RELATION_IS_CONTAINED = 2
        '
        ' The object entirely contains the passed in geometry.
        '
        D2D1_GEOMETRY_RELATION_CONTAINS = 3
        '
        ' The two geometries overlap but neither completely contains the other.
        '
        D2D1_GEOMETRY_RELATION_OVERLAP = 4
        D2D1_GEOMETRY_RELATION_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_GEOMETRY_SIMPLIFICATION_OPTION
        D2D1_GEOMETRY_SIMPLIFICATION_OPTION_CUBICS_AND_LINES = 0
        D2D1_GEOMETRY_SIMPLIFICATION_OPTION_LINES = 1
        D2D1_GEOMETRY_SIMPLIFICATION_OPTION_FORCE_DWORD = &HFFFFFFFF
    End Enum
    Public Enum D2D1_COMBINE_MODE
        '
        ' Produce a geometry representing the set of points contained in either
        ' the first or the second geometry.
        '
        D2D1_COMBINE_MODE_UNION = 0
        '
        ' Produce a geometry representing the set of points common to the first
        ' and the second geometries.
        '
        D2D1_COMBINE_MODE_INTERSECT = 1
        '
        ' Produce a geometry representing the set of points contained in the
        ' first geometry or the second geometry, but not both.
        '
        D2D1_COMBINE_MODE_XOR = 2
        '
        ' Produce a geometry representing the set of points contained in the
        ' first geometry but not the second geometry.
        '
        D2D1_COMBINE_MODE_EXCLUDE = 3
        D2D1_COMBINE_MODE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_PATH_SEGMENT
        D2D1_PATH_SEGMENT_NONE = &H0
        D2D1_PATH_SEGMENT_FORCE_UNSTROKED = &H1
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

    <ComImport>
    <Guid("2cd9069e-12e2-11dc-9fed-001143a055f9")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1SimplifiedGeometrySink
        <PreserveSig>
        Function SetFillMode(fillMode As D2D1_FILL_MODE) As HRESULT
        <PreserveSig>
        Function SetSegmentFlags(vertexFlags As D2D1_PATH_SEGMENT) As HRESULT
        <PreserveSig>
        Function BeginFigure(startPoint As D2D1_POINT_2F, figureBegin As D2D1_FIGURE_BEGIN) As HRESULT
        <PreserveSig>
        Function AddLines(
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> points As D2D1_POINT_2F(), pointsCount As Integer) As HRESULT
        <PreserveSig>
        Function AddBeziers(
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> beziers As D2D1_BEZIER_SEGMENT(), beziersCount As Integer) As HRESULT
        <PreserveSig>
        Function EndFigure(figureEnd As D2D1_FIGURE_END) As HRESULT
        <PreserveSig>
        Function Close() As HRESULT
    End Interface

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_QUADRATIC_BEZIER_SEGMENT
        Public point1 As D2D1_POINT_2F
        Public point2 As D2D1_POINT_2F
    End Structure

    Public Enum D2D1_SWEEP_DIRECTION
        D2D1_SWEEP_DIRECTION_COUNTER_CLOCKWISE = 0
        D2D1_SWEEP_DIRECTION_CLOCKWISE = 1
        D2D1_SWEEP_DIRECTION_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_ARC_SIZE
        D2D1_ARC_SIZE_SMALL = 0
        D2D1_ARC_SIZE_LARGE = 1
        D2D1_ARC_SIZE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_ARC_SEGMENT
        Public point As D2D1_POINT_2F
        Public size As D2D1_SIZE_F
        Public rotationAngle As Single
        Public sweepDirection As D2D1_SWEEP_DIRECTION
        Public arcSize As D2D1_ARC_SIZE
    End Structure

    <ComImport>
    <Guid("e0db51c3-6f77-4bae-b3d5-e47509b35838")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1GdiInteropRenderTarget
        <PreserveSig>
        Function GetDC(mode As D2D1_DC_INITIALIZE_MODE, <Out> ByRef hdc As IntPtr) As HRESULT
        <PreserveSig>
        Function ReleaseDC(ByRef update As RECT) As HRESULT
    End Interface

    Public Enum D2D1_DC_INITIALIZE_MODE
        ''' <summary>
        ''' The contents of the D2D render target will be copied to the DC.
        ''' </summary>
        D2D1_DC_INITIALIZE_MODE_COPY = 0

        ''' <summary>
        ''' The contents of the DC will be cleared.
        ''' </summary>
        D2D1_DC_INITIALIZE_MODE_CLEAR = 1
        D2D1_DC_INITIALIZE_MODE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    <ComImport>
    <Guid("2cd9069f-12e2-11dc-9fed-001143a055f9")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1GeometrySink
        Inherits ID2D1SimplifiedGeometrySink
#Region "ID2D1SimplifiedGeometrySink"
        <PreserveSig>
        Overloads Function SetFillMode(fillMode As D2D1_FILL_MODE) As HRESULT
        <PreserveSig>
        Overloads Function SetSegmentFlags(vertexFlags As D2D1_PATH_SEGMENT) As HRESULT
        <PreserveSig>
        Overloads Function BeginFigure(startPoint As D2D1_POINT_2F, figureBegin As D2D1_FIGURE_BEGIN) As HRESULT
        <PreserveSig>
        Overloads Function AddLines(
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> points As D2D1_POINT_2F(), pointsCount As Integer) As HRESULT
        <PreserveSig>
        Overloads Function AddBeziers(
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> beziers As D2D1_BEZIER_SEGMENT(), beziersCount As Integer) As HRESULT
        <PreserveSig>
        Overloads Function EndFigure(figureEnd As D2D1_FIGURE_END) As HRESULT
        <PreserveSig>
        Overloads Function Close() As HRESULT
#End Region

        <PreserveSig>
        Function AddLine(point As D2D1_POINT_2F) As HRESULT
        <PreserveSig>
        Function AddBezier(ByRef bezier As D2D1_BEZIER_SEGMENT) As HRESULT
        <PreserveSig>
        Function AddQuadraticBezier(ByRef bezier As D2D1_QUADRATIC_BEZIER_SEGMENT) As HRESULT
        <PreserveSig>
        Function AddQuadraticBeziers(
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> beziers As D2D1_QUADRATIC_BEZIER_SEGMENT(), beziersCount As Integer) As HRESULT
        <PreserveSig>
        Function AddArc(ByRef arc As D2D1_ARC_SEGMENT) As HRESULT
    End Interface

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_TRIANGLE
        Public point1 As D2D1_POINT_2F
        Public point2 As D2D1_POINT_2F
        Public point3 As D2D1_POINT_2F
    End Structure

    <ComImport>
    <Guid("2cd906c1-12e2-11dc-9fed-001143a055f9")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1TessellationSink
        <PreserveSig>
        Sub AddTriangles(triangles As D2D1_TRIANGLE, trianglesCount As UInteger)
        <PreserveSig>
        Sub Close()
    End Interface

    <ComImport>
    <Guid("2cd906a1-12e2-11dc-9fed-001143a055f9")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1Geometry
        Inherits ID2D1Resource
#Region "ID2D1Resource"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region

        <PreserveSig>
        Function GetBounds(worldTransform As D2D1_MATRIX_3X2_F, <Out> ByRef bounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Function GetWidenedBounds(strokeWidth As Single, strokeStyle As ID2D1StrokeStyle, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef bounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Function StrokeContainsPoint(ByRef point As D2D1_POINT_2F, strokeWidth As Single, strokeStyle As ID2D1StrokeStyle, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef contains As Boolean) As HRESULT
        <PreserveSig>
        Function FillContainsPoint(ByRef point As D2D1_POINT_2F, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef contains As Boolean) As HRESULT
        <PreserveSig>
        Function CompareWithGeometry(inputGeometry As ID2D1Geometry, inputGeometryTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef relation As D2D1_GEOMETRY_RELATION) As HRESULT
        <PreserveSig>
        Function Simplify(simplificationOption As D2D1_GEOMETRY_SIMPLIFICATION_OPTION, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, geometrySink As ID2D1SimplifiedGeometrySink) As HRESULT
        <PreserveSig>
        Function Tessellate(worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, tessellationSink As ID2D1TessellationSink) As HRESULT
        <PreserveSig>
        Function CombineWithGeometry(inputGeometry As ID2D1Geometry, combineMode As D2D1_COMBINE_MODE, inputGeometryTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, geometrySink As ID2D1SimplifiedGeometrySink) As HRESULT
        <PreserveSig>
        Function Outline(worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, geometrySink As ID2D1SimplifiedGeometrySink) As HRESULT
        <PreserveSig>
        Function ComputeArea(worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef area As Single) As HRESULT
        <PreserveSig>
        Function ComputeLength(worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef length As Single) As HRESULT
        <PreserveSig>
        Function ComputePointAtLength(length As Single, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef point As D2D1_POINT_2F, <Out> ByRef unitTangentVector As D2D1_POINT_2F) As HRESULT
        <PreserveSig>
        Function Widen(strokeWidth As Single, strokeStyle As ID2D1StrokeStyle, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, geometrySink As ID2D1SimplifiedGeometrySink) As HRESULT
    End Interface

    <ComImport>
    <Guid("2cd906a3-12e2-11dc-9fed-001143a055f9")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1RoundedRectangleGeometry
        Inherits ID2D1Geometry
#Region "ID2D1Geometry"
#Region "ID2D1Resource"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region

        <PreserveSig>
        Overloads Function GetBounds(worldTransform As D2D1_MATRIX_3X2_F, <Out> ByRef bounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Function GetWidenedBounds(strokeWidth As Single, strokeStyle As ID2D1StrokeStyle, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef bounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Function StrokeContainsPoint(ByRef point As D2D1_POINT_2F, strokeWidth As Single, strokeStyle As ID2D1StrokeStyle, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef contains As Boolean) As HRESULT
        <PreserveSig>
        Overloads Function FillContainsPoint(ByRef point As D2D1_POINT_2F, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef contains As Boolean) As HRESULT
        <PreserveSig>
        Overloads Function CompareWithGeometry(inputGeometry As ID2D1Geometry, inputGeometryTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef relation As D2D1_GEOMETRY_RELATION) As HRESULT
        <PreserveSig>
        Overloads Function Simplify(simplificationOption As D2D1_GEOMETRY_SIMPLIFICATION_OPTION, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, geometrySink As ID2D1SimplifiedGeometrySink) As HRESULT
        <PreserveSig>
        Overloads Function Tessellate(worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, tessellationSink As ID2D1TessellationSink) As HRESULT
        <PreserveSig>
        Overloads Function CombineWithGeometry(inputGeometry As ID2D1Geometry, combineMode As D2D1_COMBINE_MODE, inputGeometryTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, geometrySink As ID2D1SimplifiedGeometrySink) As HRESULT
        <PreserveSig>
        Overloads Function Outline(worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, geometrySink As ID2D1SimplifiedGeometrySink) As HRESULT
        <PreserveSig>
        Overloads Function ComputeArea(worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef area As Single) As HRESULT
        <PreserveSig>
        Overloads Function ComputeLength(worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef length As Single) As HRESULT
        <PreserveSig>
        Overloads Function ComputePointAtLength(length As Single, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef point As D2D1_POINT_2F, <Out> ByRef unitTangentVector As D2D1_POINT_2F) As HRESULT
        <PreserveSig>
        Overloads Function Widen(strokeWidth As Single, strokeStyle As ID2D1StrokeStyle, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, geometrySink As ID2D1SimplifiedGeometrySink) As HRESULT

#End Region

        <PreserveSig>
        Sub GetRoundedRect(<Out> ByRef roundedRect As D2D1_ROUNDED_RECT)
    End Interface

    <ComImport>
    <Guid("2cd906a4-12e2-11dc-9fed-001143a055f9")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1EllipseGeometry
        Inherits ID2D1Geometry
#Region "ID2D1Geometry"
#Region "ID2D1Resource"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region

        <PreserveSig>
        Overloads Function GetBounds(worldTransform As D2D1_MATRIX_3X2_F, <Out> ByRef bounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Function GetWidenedBounds(strokeWidth As Single, strokeStyle As ID2D1StrokeStyle, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef bounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Function StrokeContainsPoint(ByRef point As D2D1_POINT_2F, strokeWidth As Single, strokeStyle As ID2D1StrokeStyle, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef contains As Boolean) As HRESULT
        <PreserveSig>
        Overloads Function FillContainsPoint(ByRef point As D2D1_POINT_2F, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef contains As Boolean) As HRESULT
        <PreserveSig>
        Overloads Function CompareWithGeometry(inputGeometry As ID2D1Geometry, inputGeometryTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef relation As D2D1_GEOMETRY_RELATION) As HRESULT
        <PreserveSig>
        Overloads Function Simplify(simplificationOption As D2D1_GEOMETRY_SIMPLIFICATION_OPTION, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, geometrySink As ID2D1SimplifiedGeometrySink) As HRESULT
        <PreserveSig>
        Overloads Function Tessellate(worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, tessellationSink As ID2D1TessellationSink) As HRESULT
        <PreserveSig>
        Overloads Function CombineWithGeometry(inputGeometry As ID2D1Geometry, combineMode As D2D1_COMBINE_MODE, inputGeometryTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, geometrySink As ID2D1SimplifiedGeometrySink) As HRESULT
        <PreserveSig>
        Overloads Function Outline(worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, geometrySink As ID2D1SimplifiedGeometrySink) As HRESULT
        <PreserveSig>
        Overloads Function ComputeArea(worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef area As Single) As HRESULT
        <PreserveSig>
        Overloads Function ComputeLength(worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef length As Single) As HRESULT
        <PreserveSig>
        Overloads Function ComputePointAtLength(length As Single, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef point As D2D1_POINT_2F, <Out> ByRef unitTangentVector As D2D1_POINT_2F) As HRESULT
        <PreserveSig>
        Overloads Function Widen(strokeWidth As Single, strokeStyle As ID2D1StrokeStyle, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, geometrySink As ID2D1SimplifiedGeometrySink) As HRESULT

#End Region

        <PreserveSig>
        Sub GetEllipse(<Out> ByRef ellipse As D2D1_ELLIPSE)
    End Interface

    <ComImport>
    <Guid("2cd906a6-12e2-11dc-9fed-001143a055f9")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1GeometryGroup
        Inherits ID2D1Geometry
#Region "ID2D1Geometry"
#Region "ID2D1Resource"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region

        <PreserveSig>
        Overloads Function GetBounds(worldTransform As D2D1_MATRIX_3X2_F, <Out> ByRef bounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Function GetWidenedBounds(strokeWidth As Single, strokeStyle As ID2D1StrokeStyle, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef bounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Function StrokeContainsPoint(ByRef point As D2D1_POINT_2F, strokeWidth As Single, strokeStyle As ID2D1StrokeStyle, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef contains As Boolean) As HRESULT
        <PreserveSig>
        Overloads Function FillContainsPoint(ByRef point As D2D1_POINT_2F, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef contains As Boolean) As HRESULT
        <PreserveSig>
        Overloads Function CompareWithGeometry(inputGeometry As ID2D1Geometry, inputGeometryTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef relation As D2D1_GEOMETRY_RELATION) As HRESULT
        <PreserveSig>
        Overloads Function Simplify(simplificationOption As D2D1_GEOMETRY_SIMPLIFICATION_OPTION, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, geometrySink As ID2D1SimplifiedGeometrySink) As HRESULT
        <PreserveSig>
        Overloads Function Tessellate(worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, tessellationSink As ID2D1TessellationSink) As HRESULT
        <PreserveSig>
        Overloads Function CombineWithGeometry(inputGeometry As ID2D1Geometry, combineMode As D2D1_COMBINE_MODE, inputGeometryTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, geometrySink As ID2D1SimplifiedGeometrySink) As HRESULT
        <PreserveSig>
        Overloads Function Outline(worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, geometrySink As ID2D1SimplifiedGeometrySink) As HRESULT
        <PreserveSig>
        Overloads Function ComputeArea(worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef area As Single) As HRESULT
        <PreserveSig>
        Overloads Function ComputeLength(worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef length As Single) As HRESULT
        <PreserveSig>
        Overloads Function ComputePointAtLength(length As Single, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef point As D2D1_POINT_2F, <Out> ByRef unitTangentVector As D2D1_POINT_2F) As HRESULT
        <PreserveSig>
        Overloads Function Widen(strokeWidth As Single, strokeStyle As ID2D1StrokeStyle, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, geometrySink As ID2D1SimplifiedGeometrySink) As HRESULT

#End Region

        <PreserveSig>
        Sub GetFillMode(<Out> ByRef fillMode As D2D1_FILL_MODE)
        <PreserveSig>
        Function GetSourceGeometryCount() As Integer
        <PreserveSig>
        Sub GetSourceGeometries(<Out> ByRef geometries As ID2D1Geometry, geometriesCount As Integer)
    End Interface

    <ComImport>
    <Guid("2cd906bb-12e2-11dc-9fed-001143a055f9")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1TransformedGeometry
        Inherits ID2D1Geometry
#Region "ID2D1Geometry"
#Region "ID2D1Resource"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region

        <PreserveSig>
        Overloads Function GetBounds(worldTransform As D2D1_MATRIX_3X2_F, <Out> ByRef bounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Function GetWidenedBounds(strokeWidth As Single, strokeStyle As ID2D1StrokeStyle, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef bounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Function StrokeContainsPoint(ByRef point As D2D1_POINT_2F, strokeWidth As Single, strokeStyle As ID2D1StrokeStyle, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef contains As Boolean) As HRESULT
        <PreserveSig>
        Overloads Function FillContainsPoint(ByRef point As D2D1_POINT_2F, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef contains As Boolean) As HRESULT
        <PreserveSig>
        Overloads Function CompareWithGeometry(inputGeometry As ID2D1Geometry, inputGeometryTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef relation As D2D1_GEOMETRY_RELATION) As HRESULT
        <PreserveSig>
        Overloads Function Simplify(simplificationOption As D2D1_GEOMETRY_SIMPLIFICATION_OPTION, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, geometrySink As ID2D1SimplifiedGeometrySink) As HRESULT
        <PreserveSig>
        Overloads Function Tessellate(worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, tessellationSink As ID2D1TessellationSink) As HRESULT
        <PreserveSig>
        Overloads Function CombineWithGeometry(inputGeometry As ID2D1Geometry, combineMode As D2D1_COMBINE_MODE, inputGeometryTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, geometrySink As ID2D1SimplifiedGeometrySink) As HRESULT
        <PreserveSig>
        Overloads Function Outline(worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, geometrySink As ID2D1SimplifiedGeometrySink) As HRESULT
        <PreserveSig>
        Overloads Function ComputeArea(worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef area As Single) As HRESULT
        <PreserveSig>
        Overloads Function ComputeLength(worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef length As Single) As HRESULT
        <PreserveSig>
        Overloads Function ComputePointAtLength(length As Single, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef point As D2D1_POINT_2F, <Out> ByRef unitTangentVector As D2D1_POINT_2F) As HRESULT
        <PreserveSig>
        Overloads Function Widen(strokeWidth As Single, strokeStyle As ID2D1StrokeStyle, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, geometrySink As ID2D1SimplifiedGeometrySink) As HRESULT

#End Region

        <PreserveSig>
        Sub GetSourceGeometry(<Out> ByRef sourceGeometry As ID2D1Geometry)
        <PreserveSig>
        Sub GetTransform(<Out> ByRef transform As D2D1_MATRIX_3X2_F_STRUCT)
    End Interface

    <ComImport>
    <Guid("2cd906a5-12e2-11dc-9fed-001143a055f9")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1PathGeometry
        Inherits ID2D1Geometry
#Region "ID2D1Geometry"
#Region "ID2D1Resource"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region

        <PreserveSig>
        Overloads Function GetBounds(worldTransform As D2D1_MATRIX_3X2_F, <Out> ByRef bounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Function GetWidenedBounds(strokeWidth As Single, strokeStyle As ID2D1StrokeStyle, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef bounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Function StrokeContainsPoint(ByRef point As D2D1_POINT_2F, strokeWidth As Single, strokeStyle As ID2D1StrokeStyle, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef contains As Boolean) As HRESULT
        <PreserveSig>
        Overloads Function FillContainsPoint(ByRef point As D2D1_POINT_2F, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef contains As Boolean) As HRESULT
        <PreserveSig>
        Overloads Function CompareWithGeometry(inputGeometry As ID2D1Geometry, inputGeometryTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef relation As D2D1_GEOMETRY_RELATION) As HRESULT
        <PreserveSig>
        Overloads Function Simplify(simplificationOption As D2D1_GEOMETRY_SIMPLIFICATION_OPTION, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, geometrySink As ID2D1SimplifiedGeometrySink) As HRESULT
        <PreserveSig>
        Overloads Function Tessellate(worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, tessellationSink As ID2D1TessellationSink) As HRESULT
        <PreserveSig>
        Overloads Function CombineWithGeometry(inputGeometry As ID2D1Geometry, combineMode As D2D1_COMBINE_MODE, inputGeometryTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, geometrySink As ID2D1SimplifiedGeometrySink) As HRESULT
        <PreserveSig>
        Overloads Function Outline(worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, geometrySink As ID2D1SimplifiedGeometrySink) As HRESULT
        <PreserveSig>
        Overloads Function ComputeArea(worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef area As Single) As HRESULT
        <PreserveSig>
        Overloads Function ComputeLength(worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef length As Single) As HRESULT
        <PreserveSig>
        Overloads Function ComputePointAtLength(length As Single, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef point As D2D1_POINT_2F, <Out> ByRef unitTangentVector As D2D1_POINT_2F) As HRESULT
        <PreserveSig>
        Overloads Function Widen(strokeWidth As Single, strokeStyle As ID2D1StrokeStyle, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, geometrySink As ID2D1SimplifiedGeometrySink) As HRESULT

#End Region

        <PreserveSig>
        Function Open(<Out> ByRef geometrySink As ID2D1GeometrySink) As HRESULT
        <PreserveSig>
        Function Stream(geometrySink As ID2D1GeometrySink) As HRESULT
        <PreserveSig>
        Function GetSegmentCount(<Out> ByRef count As Integer) As HRESULT
        <PreserveSig>
        Function GetFigureCount(<Out> ByRef count As Integer) As HRESULT
    End Interface

    <ComImport>
    <Guid("28506e39-ebf6-46a1-bb47-fd85565ab957")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1DrawingStateBlock
        Inherits ID2D1Resource
#Region "ID2D1Resource"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region

        <PreserveSig>
        Sub GetDescription(<Out> ByRef stateDescription As D2D1_DRAWING_STATE_DESCRIPTION)
        <PreserveSig>
        Sub SetDescription(stateDescription As D2D1_DRAWING_STATE_DESCRIPTION)
        <PreserveSig>
        Sub SetTextRenderingParams(Optional textRenderingParams As IDWriteRenderingParams = Nothing)
        <PreserveSig>
        Sub GetTextRenderingParams(<Out> ByRef textRenderingParams As IDWriteRenderingParams)
    End Interface

    <ComImport>
    <Guid("2cd90698-12e2-11dc-9fed-001143a055f9")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1HwndRenderTarget
        Inherits ID2D1RenderTarget
#Region "<ID2D1RenderTarget>"

#Region "<ID2D1Resource>"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)

#End Region
        <PreserveSig>
        Overloads Function CreateBitmap(size As D2D1_SIZE_U, srcData As IntPtr, pitch As UInteger, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES, <Out> ByRef bitmap As ID2D1Bitmap) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapFromWicBitmap(wicBitmapSource As IWICBitmapSource, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES, <Out> ByRef bitmap As ID2D1Bitmap) As HRESULT
        <PreserveSig>
        Overloads Function CreateSharedBitmap(ByRef riid As Guid,
        <[In], Out> data As IntPtr, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES, <Out> ByRef bitmap As ID2D1Bitmap) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapBrush(bitmap As ID2D1Bitmap, ByRef bitmapBrushProperties As D2D1_BITMAP_BRUSH_PROPERTIES, brushProperties As D2D1_BRUSH_PROPERTIES, <Out> ByRef bitmapBrush As ID2D1BitmapBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateSolidColorBrush(color As D2D1_COLOR_F, brushProperties As D2D1_BRUSH_PROPERTIES, <Out> ByRef solidColorBrush As ID2D1SolidColorBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateGradientStopCollection(
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> gradientStops As D2D1_GRADIENT_STOP(), gradientStopsCount As UInteger, colorInterpolationGamma As D2D1_GAMMA, extendMode As D2D1_EXTEND_MODE, <Out> ByRef gradientStopCollection As ID2D1GradientStopCollection) As HRESULT
        <PreserveSig>
        Overloads Function CreateLinearGradientBrush(ByRef linearGradientBrushProperties As D2D1_LINEAR_GRADIENT_BRUSH_PROPERTIES, brushProperties As IntPtr, gradientStopCollection As ID2D1GradientStopCollection, <Out> ByRef linearGradientBrush As ID2D1LinearGradientBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateRadialGradientBrush(ByRef radialGradientBrushProperties As D2D1_RADIAL_GRADIENT_BRUSH_PROPERTIES, brushProperties As IntPtr, gradientStopCollection As ID2D1GradientStopCollection, <Out> ByRef radialGradientBrush As ID2D1RadialGradientBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateCompatibleRenderTarget(ByRef desiredSize As D2D1_SIZE_F, ByRef desiredPixelSize As D2D1_SIZE_U, ByRef desiredFormat As D2D1_PIXEL_FORMAT, options As D2D1_COMPATIBLE_RENDER_TARGET_OPTIONS, <Out> ByRef bitmapRenderTarget As ID2D1BitmapRenderTarget) As HRESULT
        <PreserveSig>
        Overloads Function CreateLayer(ByRef size As D2D1_SIZE_F, <Out> ByRef layer As ID2D1Layer) As HRESULT
        <PreserveSig>
        Overloads Sub CreateMesh(<Out> ByRef mesh As ID2D1Mesh)
        <PreserveSig>
        Overloads Sub DrawLine(point0 As D2D1_POINT_2F, point1 As D2D1_POINT_2F, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub DrawRectangle(ByRef rect As D2D1_RECT_F, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillRectangle(ByRef rect As D2D1_RECT_F, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub DrawRoundedRectangle(ByRef roundedRect As D2D1_ROUNDED_RECT, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillRoundedRectangle(ByRef roundedRect As D2D1_ROUNDED_RECT, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub DrawEllipse(ByRef ellipse As D2D1_ELLIPSE, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillEllipse(ByRef ellipse As D2D1_ELLIPSE, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub DrawGeometry(geometry As ID2D1Geometry, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillGeometry(geometry As ID2D1Geometry, brush As ID2D1Brush, Optional opacityBrush As ID2D1Brush = Nothing)
        <PreserveSig>
        Overloads Sub FillMesh(mesh As ID2D1Mesh, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub FillOpacityMask(opacityMask As ID2D1Bitmap, brush As ID2D1Brush, content As D2D1_OPACITY_MASK_CONTENT, destinationRectangle As D2D1_RECT_F, sourceRectangle As D2D1_RECT_F)
        <PreserveSig>
        Overloads Sub DrawBitmap(bitmap As ID2D1Bitmap, ByRef destinationRectangle As D2D1_RECT_F, opacity As Single, interpolationMode As D2D1_BITMAP_INTERPOLATION_MODE, ByRef sourceRectangle As D2D1_RECT_F)
        <PreserveSig>
        Overloads Sub DrawText(str As String, stringLength As UInteger, textFormat As IDWriteTextFormat, ByRef layoutRect As D2D1_RECT_F, defaultForegroundBrush As ID2D1Brush, options As D2D1_DRAW_TEXT_OPTIONS, measuringMode As DWRITE_MEASURING_MODE)
        <PreserveSig>
        Overloads Sub DrawTextLayout(origin As D2D1_POINT_2F, textLayout As IDWriteTextLayout, defaultForegroundBrush As ID2D1Brush, options As D2D1_DRAW_TEXT_OPTIONS)
        <PreserveSig>
        Overloads Sub DrawGlyphRun(baselineOrigin As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, foregroundBrush As ID2D1Brush, measuringMode As DWRITE_MEASURING_MODE)
        <PreserveSig>
        Overloads Sub SetTransform(transform As D2D1_MATRIX_3X2_F)
        <PreserveSig>
        Overloads Sub GetTransform(<Out> ByRef transform As D2D1_MATRIX_3X2_F_STRUCT)
        <PreserveSig>
        Overloads Sub SetAntialiasMode(antialiasMode As D2D1_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub GetAntialiasMode(<Out> ByRef antialiasMode As D2D1_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub SetTextAntialiasMode(textAntialiasMode As D2D1_TEXT_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub GetTextAntialiasMode(<Out> ByRef textAntialiasMode As D2D1_TEXT_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub SetTextRenderingParams(Optional textRenderingParams As IDWriteRenderingParams = Nothing)
        <PreserveSig>
        Overloads Sub GetTextRenderingParams(<Out> ByRef textRenderingParams As IDWriteRenderingParams)
        <PreserveSig>
        Overloads Sub SetTags(tag1 As ULong, tag2 As ULong)
        <PreserveSig>
        Overloads Sub GetTags(<Out> ByRef tag1 As ULong, <Out> ByRef tag2 As ULong)
        <PreserveSig>
        Overloads Sub PushLayer(ByRef layerParameters As D2D1_LAYER_PARAMETERS, layer As ID2D1Layer)
        <PreserveSig>
        Overloads Sub PopLayer()
        <PreserveSig>
        Overloads Sub Flush(<Out> ByRef tag1 As ULong, <Out> ByRef tag2 As ULong)
        <PreserveSig>
        Overloads Sub SaveDrawingState(
        <[In], Out> drawingStateBlock As ID2D1DrawingStateBlock)
        <PreserveSig>
        Overloads Sub RestoreDrawingState(drawingStateBlock As ID2D1DrawingStateBlock)
        <PreserveSig>
        Overloads Sub PushAxisAlignedClip(clipRect As D2D1_RECT_F, antialiasMode As D2D1_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub PopAxisAlignedClip()
        <PreserveSig>
        Overloads Sub Clear(clearColor As D2D1_COLOR_F)
        <PreserveSig>
        Overloads Sub BeginDraw()
        <PreserveSig>
        Overloads Function EndDraw(<Out> ByRef tag1 As ULong, <Out> ByRef tag2 As ULong) As HRESULT
        <PreserveSig>
        Overloads Sub GetPixelFormat(<Out> ByRef format As D2D1_PIXEL_FORMAT)
        <PreserveSig>
        Overloads Sub SetDpi(dpiX As Single, dpiY As Single)
        <PreserveSig>
        Overloads Sub GetDpi(<Out> ByRef dpiX As Single, <Out> ByRef dpiY As Single)
        <PreserveSig>
        Overloads Sub GetSize(<Out> ByRef size As D2D1_SIZE_F)
        <PreserveSig>
        Overloads Sub GetPixelSize(<Out> ByRef size As D2D1_SIZE_U)
        <PreserveSig>
        Overloads Function GetMaximumBitmapSize() As UInteger
        <PreserveSig>
        Overloads Function IsSupported(renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES) As Boolean

        'new HRESULT Function1();
        'new HRESULT Function2();
        'new HRESULT Function3();
        'new HRESULT Function4();
        'new HRESULT Function5();
        'new HRESULT Function6();
        'new HRESULT Function7();
        'new HRESULT Function8();
        'new HRESULT Function9();
        'new HRESULT Function10();
        'new HRESULT Function11();
        'new HRESULT Function12();
        'new HRESULT Function13();
        'new HRESULT Function14();
        'new HRESULT Function15();
        'new HRESULT Function16();
        'new HRESULT Function17();
        'new HRESULT Function18();
        'new HRESULT Function19();
        'new HRESULT Function20();
        'new HRESULT Function21();
        'new void Function22();
        'new void Function23();
        'new void Function24();
        'new void Function25();
        'new void Function26();
        'new void Function27();
        'new void Function28();
        'new void Function29();
        'new void Function30();
        'new void Function31();
        'new void Function32();
        'new void Function33();
        'new void Function34();
        'new void Function35();
        'new bool Function36();

#End Region

        <PreserveSig>
        Function CheckWindowState() As D2D1_WINDOW_STATE
        <PreserveSig>
        Function Resize(ByRef pixelSize As D2D1_SIZE_U) As HRESULT

        '[return: MarshalAs(UnmanagedType.SysInt)]
        <PreserveSig>
        Function GetHwnd() As IntPtr
    End Interface

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_SIZE_U
        Public width As UInteger
        Public height As UInteger
        Public Sub New(width As UInteger, height As UInteger)
            Me.width = width
            Me.height = height
        End Sub
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_VECTOR_2F
        Public x As Single
        Public y As Single
        Public Sub New(x As Single, y As Single)
            Me.x = x
            Me.y = y
        End Sub
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_VECTOR_3F
        Public x As Single
        Public y As Single
        Public z As Single
        Public Sub New(x As Single, y As Single, z As Single)
            Me.x = x
            Me.y = y
            Me.z = z
        End Sub
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_VECTOR_4F
        Public x As Single
        Public y As Single
        Public z As Single
        Public w As Single
        Public Sub New(x As Single, y As Single, z As Single, w As Single)
            Me.x = x
            Me.y = y
            Me.z = z
            Me.w = w
        End Sub
    End Structure

    'D2D1_SIZE_U SizeU(UInt32 width = 0, UInt32 height = 0)
    '{
    '    return Size<UInt32>(width, height);
    '}

    '  public D2D1_RENDER_TARGET_PROPERTIES RenderTargetProperties(
    'D2D1_RENDER_TARGET_TYPE type = D2D1_RENDER_TARGET_TYPE.D2D1_RENDER_TARGET_TYPE_DEFAULT,
    'D2D1_PIXEL_FORMAT pixelFormat = PixelFormat(),
    'float dpiX = 0.0f,
    'float dpiY = 0.0f,
    'D2D1_RENDER_TARGET_USAGE usage = D2D1_RENDER_TARGET_USAGE.D2D1_RENDER_TARGET_USAGE_NONE,
    'D2D1_FEATURE_LEVEL minLevel = D2D1_FEATURE_LEVEL.D2D1_FEATURE_LEVEL_DEFAULT
    ')
    '  {
    '      D2D1_RENDER_TARGET_PROPERTIES renderTargetProperties;

    '      renderTargetProperties.type = type;
    '      renderTargetProperties.pixelFormat = pixelFormat;
    '      renderTargetProperties.dpiX = dpiX;
    '      renderTargetProperties.dpiY = dpiY;
    '      renderTargetProperties.usage = usage;
    '      renderTargetProperties.minLevel = minLevel;

    '      return renderTargetProperties;
    '  }

    '[StructLayout(LayoutKind.Sequential, Pack = 1)]
    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_RECT_F
        Public left As Single
        Public top As Single
        Public right As Single
        Public bottom As Single
        Public Sub New(left As Single, top As Single, right As Single, bottom As Single)
            Me.left = left
            Me.top = top
            Me.right = right
            Me.bottom = bottom
        End Sub
    End Structure

    <ComImport>
    <Guid("2cd906a2-12e2-11dc-9fed-001143a055f9")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1RectangleGeometry
        Inherits ID2D1Geometry
#Region "<ID2D1Geometry>"

#Region "ID2D1Resource"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region

        <PreserveSig>
        Overloads Function GetBounds(worldTransform As D2D1_MATRIX_3X2_F, <Out> ByRef bounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Function GetWidenedBounds(strokeWidth As Single, strokeStyle As ID2D1StrokeStyle, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef bounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Function StrokeContainsPoint(ByRef point As D2D1_POINT_2F, strokeWidth As Single, strokeStyle As ID2D1StrokeStyle, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef contains As Boolean) As HRESULT
        <PreserveSig>
        Overloads Function FillContainsPoint(ByRef point As D2D1_POINT_2F, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef contains As Boolean) As HRESULT
        <PreserveSig>
        Overloads Function CompareWithGeometry(inputGeometry As ID2D1Geometry, inputGeometryTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef relation As D2D1_GEOMETRY_RELATION) As HRESULT
        <PreserveSig>
        Overloads Function Simplify(simplificationOption As D2D1_GEOMETRY_SIMPLIFICATION_OPTION, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, geometrySink As ID2D1SimplifiedGeometrySink) As HRESULT
        <PreserveSig>
        Overloads Function Tessellate(worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, tessellationSink As ID2D1TessellationSink) As HRESULT
        <PreserveSig>
        Overloads Function CombineWithGeometry(inputGeometry As ID2D1Geometry, combineMode As D2D1_COMBINE_MODE, inputGeometryTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, geometrySink As ID2D1SimplifiedGeometrySink) As HRESULT
        <PreserveSig>
        Overloads Function Outline(worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, geometrySink As ID2D1SimplifiedGeometrySink) As HRESULT
        <PreserveSig>
        Overloads Function ComputeArea(worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef area As Single) As HRESULT
        <PreserveSig>
        Overloads Function ComputeLength(worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef length As Single) As HRESULT
        <PreserveSig>
        Overloads Function ComputePointAtLength(length As Single, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef point As D2D1_POINT_2F, <Out> ByRef unitTangentVector As D2D1_POINT_2F) As HRESULT
        <PreserveSig>
        Overloads Function Widen(strokeWidth As Single, strokeStyle As ID2D1StrokeStyle, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, geometrySink As ID2D1SimplifiedGeometrySink) As HRESULT

#End Region

        <PreserveSig>
        Sub GetRect(<Out> ByRef rect As D2D1_RECT_F)
    End Interface

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_ROUNDED_RECT
        Public rect As D2D1_RECT_F
        Public radiusX As Single
        Public radiusY As Single
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_POINT_2F
        Public x As Single
        Public y As Single

        Public Sub New(x As Single, y As Single)
            Me.x = x
            Me.y = y
        End Sub
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_ELLIPSE
        Public point As D2D1_POINT_2F
        Public radiusX As Single
        Public radiusY As Single
    End Structure

    Public Enum D2D1_FILL_MODE
        D2D1_FILL_MODE_ALTERNATE = 0
        D2D1_FILL_MODE_WINDING = 1
        D2D1_FILL_MODE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    <StructLayout(LayoutKind.Explicit, Size:=24)>
    Public Class D2D1_MATRIX_3X2_F
        <FieldOffset(0)>
        Public _11 As Single
        <FieldOffset(4)>
        Public _12 As Single
        <FieldOffset(8)>
        Public _21 As Single
        <FieldOffset(12)>
        Public _22 As Single
        <FieldOffset(16)>
        Public _31 As Single
        <FieldOffset(20)>
        Public _32 As Single

        'public D2D1_MATRIX_3X2_F(float _11, float _12, float _21, float _22, float _31, float _32)
        '{
        '    this._11 = _11;
        '    this._12 = _12;
        '    this._21 = _21;
        '    this._22 = _22;
        '    this._31 = _31;
        '    this._32 = _32;
        '}
    End Class

    <StructLayout(LayoutKind.Explicit, Size:=24)>
    Public Structure D2D1_MATRIX_3X2_F_STRUCT
        <FieldOffset(0)>
        Public _11 As Single
        <FieldOffset(4)>
        Public _12 As Single
        <FieldOffset(8)>
        Public _21 As Single
        <FieldOffset(12)>
        Public _22 As Single
        <FieldOffset(16)>
        Public _31 As Single
        <FieldOffset(20)>
        Public _32 As Single
    End Structure

    <StructLayout(LayoutKind.Explicit, Size:=64)>
    Public Class D2D1_MATRIX_4X4_F
        <FieldOffset(0)>
        Public _11 As Single
        <FieldOffset(4)>
        Public _12 As Single
        <FieldOffset(8)>
        Public _13 As Single
        <FieldOffset(12)>
        Public _14 As Single

        <FieldOffset(16)>
        Public _21 As Single
        <FieldOffset(20)>
        Public _22 As Single
        <FieldOffset(24)>
        Public _23 As Single
        <FieldOffset(28)>
        Public _24 As Single

        <FieldOffset(32)>
        Public _31 As Single
        <FieldOffset(36)>
        Public _32 As Single
        <FieldOffset(40)>
        Public _33 As Single
        <FieldOffset(44)>
        Public _34 As Single

        <FieldOffset(48)>
        Public _41 As Single
        <FieldOffset(52)>
        Public _42 As Single
        <FieldOffset(56)>
        Public _43 As Single
        <FieldOffset(60)>
        Public _44 As Single
    End Class

    Public Enum D2D1_DASH_STYLE
        D2D1_DASH_STYLE_SOLID = 0
        D2D1_DASH_STYLE_DASH = 1
        D2D1_DASH_STYLE_DOT = 2
        D2D1_DASH_STYLE_DASH_DOT = 3
        D2D1_DASH_STYLE_DASH_DOT_DOT = 4
        D2D1_DASH_STYLE_CUSTOM = 5
        D2D1_DASH_STYLE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_LINE_JOIN
        '
        ' Miter join.
        '
        D2D1_LINE_JOIN_MITER = 0
        '
        ' Bevel join.
        '
        D2D1_LINE_JOIN_BEVEL = 1
        '
        ' Round join.
        '
        D2D1_LINE_JOIN_ROUND = 2
        '
        ' Miter/Bevel join.
        '
        D2D1_LINE_JOIN_MITER_OR_BEVEL = 3
        D2D1_LINE_JOIN_FORCE_DWORD = &HFFFFFFFF
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Class D2D1_STROKE_STYLE_PROPERTIES
        Public startCap As D2D1_CAP_STYLE
        Public endCap As D2D1_CAP_STYLE
        Public dashCap As D2D1_CAP_STYLE
        Public lineJoin As D2D1_LINE_JOIN
        Public miterLimit As Single
        Public dashStyle As D2D1_DASH_STYLE
        Public dashOffset As Single
        Public Sub New(startCap As D2D1_CAP_STYLE, endCap As D2D1_CAP_STYLE, dashCap As D2D1_CAP_STYLE, lineJoin As D2D1_LINE_JOIN, miterLimit As Single, dashStyle As D2D1_DASH_STYLE, dashOffset As Single)
            Me.startCap = startCap
            Me.endCap = endCap
            Me.dashCap = dashCap
            Me.lineJoin = lineJoin
            Me.miterLimit = miterLimit
            Me.dashStyle = dashStyle
            Me.dashOffset = dashOffset
        End Sub
    End Class

    Public Enum D2D1_ANTIALIAS_MODE
        '
        ' The edges of each primitive are antialiased sequentially.
        '
        D2D1_ANTIALIAS_MODE_PER_PRIMITIVE = 0

        '
        ' Each pixel is rendered if its pixel center is contained by the geometry.
        '
        D2D1_ANTIALIAS_MODE_ALIASED = 1
        D2D1_ANTIALIAS_MODE_FORCE_DWORD = &HFFFFFFFF
    End Enum
    Public Enum D2D1_TEXT_ANTIALIAS_MODE
        '
        ' Render text using the current system setting.
        '
        D2D1_TEXT_ANTIALIAS_MODE_DEFAULT = 0
        '
        ' Render text using ClearType.
        '
        D2D1_TEXT_ANTIALIAS_MODE_CLEARTYPE = 1
        '
        ' Render text using gray-scale.
        '
        D2D1_TEXT_ANTIALIAS_MODE_GRAYSCALE = 2
        '
        ' Render text aliased.
        '
        D2D1_TEXT_ANTIALIAS_MODE_ALIASED = 3
        D2D1_TEXT_ANTIALIAS_MODE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_DRAWING_STATE_DESCRIPTION
        Public antialiasMode As D2D1_ANTIALIAS_MODE
        Public textAntialiasMode As D2D1_TEXT_ANTIALIAS_MODE
        Public tag1 As ULong
        Public tag2 As ULong
        Public transform As D2D1_MATRIX_3X2_F
    End Structure

    Public Enum D2D1_PRESENT_OPTIONS
        D2D1_PRESENT_OPTIONS_NONE = &H0
        '
        ' Keep the target contents intact through present.
        '
        D2D1_PRESENT_OPTIONS_RETAIN_CONTENTS = &H1
        '
        ' Do not wait for display refresh to commit changes to display.
        '
        D2D1_PRESENT_OPTIONS_IMMEDIATELY = &H2
        D2D1_PRESENT_OPTIONS_FORCE_DWORD = &HFFFFFFFF
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_HWND_RENDER_TARGET_PROPERTIES
        Public hwnd As IntPtr
        Public pixelSize As D2D1_SIZE_U
        Public presentOptions As D2D1_PRESENT_OPTIONS
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_BITMAP_PROPERTIES
        Public pixelFormat As D2D1_PIXEL_FORMAT
        Public dpiX As Single
        Public dpiY As Single
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_SIZE_F
        Public width As Single
        Public height As Single

        Public Sub New(width As Single, height As Single)
            Me.width = width
            Me.height = height
        End Sub
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_POINT_2U
        Public x As UInteger
        Public y As UInteger

        Public Sub New(x As UInteger, y As UInteger)
            Me.x = x
            Me.y = y
        End Sub
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_RECT_U
        Public left As UInteger
        Public top As UInteger
        Public right As UInteger
        Public bottom As UInteger
        Public Sub New(left As UInteger, top As UInteger, right As UInteger, bottom As UInteger)
            Me.left = left
            Me.top = top
            Me.right = right
            Me.bottom = bottom
        End Sub
    End Structure

    <ComImport>
    <Guid("65019f75-8da2-497c-b32c-dfa34e48ede6")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1Image
        Inherits ID2D1Resource
#Region "<ID2D1Resource>"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region
    End Interface

    <ComImport>
    <Guid("a2296057-ea42-4099-983b-539fb6505426")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1Bitmap
        Inherits ID2D1Image
#Region "<ID2D1Image>"

#Region "<ID2D1Resource>"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)

#End Region

#End Region

        <PreserveSig>
        Sub GetSize(<Out> ByRef size As D2D1_SIZE_F)
        <PreserveSig>
        Sub GetPixelSize(<Out> ByRef size As D2D1_SIZE_U)
        <PreserveSig>
        Sub GetPixelFormat(<Out> ByRef format As D2D1_PIXEL_FORMAT)
        <PreserveSig>
        Sub GetDpi(<Out> ByRef dpiX As Single, <Out> ByRef dpiY As Single)
        <PreserveSig>
        Sub CopyFromBitmap(ByRef destPoint As D2D1_POINT_2U, bitmap As ID2D1Bitmap, srcRect As D2D1_RECT_U)
        <PreserveSig>
        Sub CopyFromRenderTarget(ByRef destPoint As D2D1_POINT_2U, renderTarget As ID2D1RenderTarget, srcRect As D2D1_RECT_U)
        <PreserveSig>
        Sub CopyFromMemory(ByRef dstRect As D2D1_RECT_U, srcData As IntPtr, pitch As UInteger)
    End Interface

    Public Enum D2D1_EXTEND_MODE
        '
        ' Extend the edges of the source out by clamping sample points outside the source
        ' to the edges.
        '
        D2D1_EXTEND_MODE_CLAMP = 0
        '
        ' The base tile is drawn untransformed and the remainder are filled by repeating
        ' the base tile.
        '
        D2D1_EXTEND_MODE_WRAP = 1
        '
        ' The same as wrap, but alternate tiles are flipped  The base tile is drawn
        ' untransformed.
        '
        D2D1_EXTEND_MODE_MIRROR = 2
        D2D1_EXTEND_MODE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_INTERPOLATION_MODE_DEFINITION
        D2D1_INTERPOLATION_MODE_DEFINITION_NEAREST_NEIGHBOR = 0
        D2D1_INTERPOLATION_MODE_DEFINITION_LINEAR = 1
        D2D1_INTERPOLATION_MODE_DEFINITION_CUBIC = 2
        D2D1_INTERPOLATION_MODE_DEFINITION_MULTI_SAMPLE_LINEAR = 3
        D2D1_INTERPOLATION_MODE_DEFINITION_ANISOTROPIC = 4
        D2D1_INTERPOLATION_MODE_DEFINITION_HIGH_QUALITY_CUBIC = 5
        D2D1_INTERPOLATION_MODE_DEFINITION_FANT = 6
        D2D1_INTERPOLATION_MODE_DEFINITION_MIPMAP_LINEAR = 7
    End Enum

    Public Enum D2D1_BITMAP_INTERPOLATION_MODE
        '
        ' Nearest Neighbor filtering. Also known as nearest pixel or nearest point
        ' sampling.
        '
        D2D1_BITMAP_INTERPOLATION_MODE_NEAREST_NEIGHBOR = D2D1_INTERPOLATION_MODE_DEFINITION.D2D1_INTERPOLATION_MODE_DEFINITION_NEAREST_NEIGHBOR
        '
        ' Linear filtering.
        '
        D2D1_BITMAP_INTERPOLATION_MODE_LINEAR = D2D1_INTERPOLATION_MODE_DEFINITION.D2D1_INTERPOLATION_MODE_DEFINITION_LINEAR
        D2D1_BITMAP_INTERPOLATION_MODE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_BITMAP_BRUSH_PROPERTIES
        Public extendModeX As D2D1_EXTEND_MODE
        Public extendModeY As D2D1_EXTEND_MODE
        Public interpolationMode As D2D1_BITMAP_INTERPOLATION_MODE
    End Structure

    '<StructLayout(LayoutKind.Sequential)>
    'Public Structure D2D1_BRUSH_PROPERTIES
    '    Public opacity As Single
    '    Public transform As D2D1_MATRIX_3X2_F
    'End Structure

    Public Class D2D1_BRUSH_PROPERTIES
        Public opacity As Single
        Public transform As D2D1_MATRIX_3X2_F
    End Class

    <StructLayout(LayoutKind.Sequential)>
    Public Class D2D1_COLOR_F
        Public r As Single
        Public g As Single
        Public b As Single
        Public a As Single
    End Class

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_COLOR_F_STRUCT
        Public r As Single
        Public g As Single
        Public b As Single
        Public a As Single

        Public Sub New(r As Single, g As Single, b As Single, a As Single)
            Me.r = r
            Me.g = g
            Me.b = b
            Me.a = a
        End Sub
    End Structure

    '[StructLayout(LayoutKind.Sequential)]
    'public struct D2D1_COLOR_F_STRUCT
    '{
    '    public float r;
    '    public float g;
    '    public float b;
    '    public float a;

    '    public D2D1_COLOR_F_STRUCT(float r, float g, float b, float a)
    '    {
    '        this.r = r;
    '        this.g = g;
    '        this.b = b;       
    '        this.a = a;
    '    }

    '    //public static implicit operator ColorF(D2D1_COLOR_F_STRUCT d)
    '    //{
    '    //    return d.r;
    '    //}

    '    //public static explicit operator D2D1_COLOR_F_STRUCT(ColorF c)
    '    //{
    '    //    return new D2D1_COLOR_F_STRUCT(c.r, c.g, c.b, c.a);
    '    //}

    '    //public static implicit operator D2D1_COLOR_F_STRUCT(ColorF c)
    '    //{

    '    //   return new D2D1_COLOR_F_STRUCT(c.r, c.g, c.b, c.a);
    '    //}
    '}

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_GRADIENT_STOP
        Public position As Single
        Public color As D2D1_COLOR_F

        Public Sub New(position As Single, color As D2D1_COLOR_F)
            Me.position = position
            Me.color = color
        End Sub
    End Structure

    Public Enum D2D1_GAMMA
        '
        ' Colors are manipulated in 2.2 gamma color space.
        '
        D2D1_GAMMA_2_2 = 0
        '
        ' Colors are manipulated in 1.0 gamma color space.
        '
        D2D1_GAMMA_1_0 = 1
        D2D1_GAMMA_FORCE_DWORD = &HFFFFFFFF
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_LINEAR_GRADIENT_BRUSH_PROPERTIES
        Public startPoint As D2D1_POINT_2F
        Public endPoint As D2D1_POINT_2F

        Public Sub New(startPoint As D2D1_POINT_2F, endPoint As D2D1_POINT_2F)
            Me.startPoint = startPoint
            Me.endPoint = endPoint
        End Sub
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_RADIAL_GRADIENT_BRUSH_PROPERTIES
        Public center As D2D1_POINT_2F
        Public gradientOriginOffset As D2D1_POINT_2F
        Public radiusX As Single
        Public radiusY As Single
    End Structure

    Public Enum D2D1_COMPATIBLE_RENDER_TARGET_OPTIONS
        D2D1_COMPATIBLE_RENDER_TARGET_OPTIONS_NONE = &H0
        '
        ' The compatible render target will allow a call to GetDC on the
        ' ID2D1GdiInteropRenderTarget interface. This can be specified even if the parent
        ' render target is not GDI compatible.
        '
        D2D1_COMPATIBLE_RENDER_TARGET_OPTIONS_GDI_COMPATIBLE = &H1
        D2D1_COMPATIBLE_RENDER_TARGET_OPTIONS_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_OPACITY_MASK_CONTENT
        '
        ' The mask contains geometries or bitmaps.
        '
        D2D1_OPACITY_MASK_CONTENT_GRAPHICS = 0
        '
        ' The mask contains text rendered using one of the natural text modes.
        '
        D2D1_OPACITY_MASK_CONTENT_TEXT_NATURAL = 1
        '
        ' The mask contains text rendered using one of the GDI compatible text modes.
        '
        D2D1_OPACITY_MASK_CONTENT_TEXT_GDI_COMPATIBLE = 2
        D2D1_OPACITY_MASK_CONTENT_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_DRAW_TEXT_OPTIONS
        '
        ' Do not snap the baseline of the text vertically.
        '
        D2D1_DRAW_TEXT_OPTIONS_NO_SNAP = &H1
        '
        ' Clip the text to the content bounds.
        '
        D2D1_DRAW_TEXT_OPTIONS_CLIP = &H2
        '
        ' Render color versions of glyphs if defined by the font.
        '
        D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT = &H4
        D2D1_DRAW_TEXT_OPTIONS_NONE = &H0
        D2D1_DRAW_TEXT_OPTIONS_FORCE_DWORD = &HFFFFFFFF
    End Enum

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

    Public Enum D2D1_LAYER_OPTIONS
        D2D1_LAYER_OPTIONS_NONE = &H0
        '
        ' The layer will render correctly for ClearType text. If the render target was set
        ' to ClearType previously, the layer will continue to render ClearType. If the
        ' render target was set to ClearType and this option is not specified, the render
        ' target will be set to render gray-scale until the layer is popped. The caller
        ' can override this default by calling SetTextAntialiasMode while within the
        ' layer. This flag is slightly slower than the default.
        '
        D2D1_LAYER_OPTIONS_INITIALIZE_FOR_CLEARTYPE = &H1
        D2D1_LAYER_OPTIONS_FORCE_DWORD = &HFFFFFFFF
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_LAYER_PARAMETERS
        Public contentBounds As D2D1_RECT_F
        '_Field_size_opt_(1) ID2D1Geometry* geometricMask;
        'public IntPtr geometricMask;
        Public geometricMask As ID2D1Geometry
        Public maskAntialiasMode As D2D1_ANTIALIAS_MODE
        Public maskTransform As D2D1_MATRIX_3X2_F
        Public opacity As Single
        '_Field_size_opt_(1) ID2D1Brush* opacityBrush;
        Public opacityBrush As IntPtr
        Public layerOptions As D2D1_LAYER_OPTIONS
    End Structure

    <ComImport>
    <Guid("2cd90691-12e2-11dc-9fed-001143a055f9")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1Resource
        <PreserveSig>
        Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
    End Interface

    <ComImport>
    <Guid("2cd906a9-12e2-11dc-9fed-001143a055f9")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1SolidColorBrush
        Inherits ID2D1Brush
#Region "ID2D1Brush"
#Region "ID2D1Resource"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region
        <PreserveSig>
        Overloads Sub SetOpacity(opacity As Single)
        <PreserveSig>
        Overloads Sub SetTransform(transform As D2D1_MATRIX_3X2_F)
        <PreserveSig>
        Overloads Function GetOpacity() As Single
        <PreserveSig>
        Overloads Sub GetTransform(<Out> ByRef transform As D2D1_MATRIX_3X2_F_STRUCT)
#End Region

        <PreserveSig>
        Sub SetColor(color As D2D1_COLOR_F)
        <PreserveSig>
        Sub GetColor(<Out> ByRef color As D2D1_COLOR_F_STRUCT)
    End Interface

    <ComImport>
    <Guid("2cd906a7-12e2-11dc-9fed-001143a055f9")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1GradientStopCollection
        Inherits ID2D1Resource
#Region "<ID2D1Resource>"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)

#End Region
        <PreserveSig>
        Function GetGradientStopCount() As UInteger
        <PreserveSig>
        Sub GetGradientStops(<Out> ByRef gradientStops As D2D1_GRADIENT_STOP, gradientStopsCount As UInteger)
        <PreserveSig>
        Sub GetColorInterpolationGamma(<Out> ByRef colorInterpolationGamma As D2D1_GAMMA)
        <PreserveSig>
        Sub GetExtendMode(<Out> ByRef extendedMode As D2D1_EXTEND_MODE)
    End Interface

    <ComImport>
    <Guid("2cd90695-12e2-11dc-9fed-001143a055f9")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1BitmapRenderTarget
        Inherits ID2D1RenderTarget
#Region "<ID2D1RenderTarget>"

#Region "<ID2D1Resource>"

        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)

#End Region

        <PreserveSig>
        Overloads Function CreateBitmap(size As D2D1_SIZE_U, srcData As IntPtr, pitch As UInteger, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES, <Out> ByRef bitmap As ID2D1Bitmap) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapFromWicBitmap(wicBitmapSource As IWICBitmapSource, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES, <Out> ByRef bitmap As ID2D1Bitmap) As HRESULT
        <PreserveSig>
        Overloads Function CreateSharedBitmap(ByRef riid As Guid,
        <[In], Out> data As IntPtr, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES, <Out> ByRef bitmap As ID2D1Bitmap) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapBrush(bitmap As ID2D1Bitmap, ByRef bitmapBrushProperties As D2D1_BITMAP_BRUSH_PROPERTIES, brushProperties As D2D1_BRUSH_PROPERTIES, <Out> ByRef bitmapBrush As ID2D1BitmapBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateSolidColorBrush(color As D2D1_COLOR_F, brushProperties As D2D1_BRUSH_PROPERTIES, <Out> ByRef solidColorBrush As ID2D1SolidColorBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateGradientStopCollection(
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> gradientStops As D2D1_GRADIENT_STOP(), gradientStopsCount As UInteger, colorInterpolationGamma As D2D1_GAMMA, extendMode As D2D1_EXTEND_MODE, <Out> ByRef gradientStopCollection As ID2D1GradientStopCollection) As HRESULT
        <PreserveSig>
        Overloads Function CreateLinearGradientBrush(ByRef linearGradientBrushProperties As D2D1_LINEAR_GRADIENT_BRUSH_PROPERTIES, brushProperties As IntPtr, gradientStopCollection As ID2D1GradientStopCollection, <Out> ByRef linearGradientBrush As ID2D1LinearGradientBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateRadialGradientBrush(ByRef radialGradientBrushProperties As D2D1_RADIAL_GRADIENT_BRUSH_PROPERTIES, brushProperties As IntPtr, gradientStopCollection As ID2D1GradientStopCollection, <Out> ByRef radialGradientBrush As ID2D1RadialGradientBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateCompatibleRenderTarget(ByRef desiredSize As D2D1_SIZE_F, ByRef desiredPixelSize As D2D1_SIZE_U, ByRef desiredFormat As D2D1_PIXEL_FORMAT, options As D2D1_COMPATIBLE_RENDER_TARGET_OPTIONS, <Out> ByRef bitmapRenderTarget As ID2D1BitmapRenderTarget) As HRESULT
        <PreserveSig>
        Overloads Function CreateLayer(ByRef size As D2D1_SIZE_F, <Out> ByRef layer As ID2D1Layer) As HRESULT
        <PreserveSig>
        Overloads Sub CreateMesh(<Out> ByRef mesh As ID2D1Mesh)
        <PreserveSig>
        Overloads Sub DrawLine(point0 As D2D1_POINT_2F, point1 As D2D1_POINT_2F, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub DrawRectangle(ByRef rect As D2D1_RECT_F, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillRectangle(ByRef rect As D2D1_RECT_F, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub DrawRoundedRectangle(ByRef roundedRect As D2D1_ROUNDED_RECT, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillRoundedRectangle(ByRef roundedRect As D2D1_ROUNDED_RECT, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub DrawEllipse(ByRef ellipse As D2D1_ELLIPSE, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillEllipse(ByRef ellipse As D2D1_ELLIPSE, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub DrawGeometry(geometry As ID2D1Geometry, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillGeometry(geometry As ID2D1Geometry, brush As ID2D1Brush, Optional opacityBrush As ID2D1Brush = Nothing)
        <PreserveSig>
        Overloads Sub FillMesh(mesh As ID2D1Mesh, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub FillOpacityMask(opacityMask As ID2D1Bitmap, brush As ID2D1Brush, content As D2D1_OPACITY_MASK_CONTENT, destinationRectangle As D2D1_RECT_F, sourceRectangle As D2D1_RECT_F)
        <PreserveSig>
        Overloads Sub DrawBitmap(bitmap As ID2D1Bitmap, ByRef destinationRectangle As D2D1_RECT_F, opacity As Single, interpolationMode As D2D1_BITMAP_INTERPOLATION_MODE, ByRef sourceRectangle As D2D1_RECT_F)
        <PreserveSig>
        Overloads Sub DrawText(str As String, stringLength As UInteger, textFormat As IDWriteTextFormat, ByRef layoutRect As D2D1_RECT_F, defaultForegroundBrush As ID2D1Brush, options As D2D1_DRAW_TEXT_OPTIONS, measuringMode As DWRITE_MEASURING_MODE)
        <PreserveSig>
        Overloads Sub DrawTextLayout(origin As D2D1_POINT_2F, textLayout As IDWriteTextLayout, defaultForegroundBrush As ID2D1Brush, options As D2D1_DRAW_TEXT_OPTIONS)
        <PreserveSig>
        Overloads Sub DrawGlyphRun(baselineOrigin As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, foregroundBrush As ID2D1Brush, measuringMode As DWRITE_MEASURING_MODE)
        <PreserveSig>
        Overloads Sub SetTransform(transform As D2D1_MATRIX_3X2_F)
        <PreserveSig>
        Overloads Sub GetTransform(<Out> ByRef transform As D2D1_MATRIX_3X2_F_STRUCT)
        <PreserveSig>
        Overloads Sub SetAntialiasMode(antialiasMode As D2D1_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub GetAntialiasMode(<Out> ByRef antialiasMode As D2D1_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub SetTextAntialiasMode(textAntialiasMode As D2D1_TEXT_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub GetTextAntialiasMode(<Out> ByRef textAntialiasMode As D2D1_TEXT_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub SetTextRenderingParams(Optional textRenderingParams As IDWriteRenderingParams = Nothing)
        <PreserveSig>
        Overloads Sub GetTextRenderingParams(<Out> ByRef textRenderingParams As IDWriteRenderingParams)
        <PreserveSig>
        Overloads Sub SetTags(tag1 As ULong, tag2 As ULong)
        <PreserveSig>
        Overloads Sub GetTags(<Out> ByRef tag1 As ULong, <Out> ByRef tag2 As ULong)
        <PreserveSig>
        Overloads Sub PushLayer(ByRef layerParameters As D2D1_LAYER_PARAMETERS, layer As ID2D1Layer)
        <PreserveSig>
        Overloads Sub PopLayer()
        <PreserveSig>
        Overloads Sub Flush(<Out> ByRef tag1 As ULong, <Out> ByRef tag2 As ULong)
        <PreserveSig>
        Overloads Sub SaveDrawingState(
        <[In], Out> drawingStateBlock As ID2D1DrawingStateBlock)
        <PreserveSig>
        Overloads Sub RestoreDrawingState(drawingStateBlock As ID2D1DrawingStateBlock)
        <PreserveSig>
        Overloads Sub PushAxisAlignedClip(clipRect As D2D1_RECT_F, antialiasMode As D2D1_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub PopAxisAlignedClip()
        <PreserveSig>
        Overloads Sub Clear(clearColor As D2D1_COLOR_F)
        <PreserveSig>
        Overloads Sub BeginDraw()
        <PreserveSig>
        Overloads Function EndDraw(<Out> ByRef tag1 As ULong, <Out> ByRef tag2 As ULong) As HRESULT
        <PreserveSig>
        Overloads Sub GetPixelFormat(<Out> ByRef pixelFormat As D2D1_PIXEL_FORMAT)
        <PreserveSig>
        Overloads Sub SetDpi(dpiX As Single, dpiY As Single)
        <PreserveSig>
        Overloads Sub GetDpi(<Out> ByRef dpiX As Single, <Out> ByRef dpiY As Single)
        <PreserveSig>
        Overloads Sub GetSize(<Out> ByRef size As D2D1_SIZE_F)
        <PreserveSig>
        Overloads Sub GetPixelSize(<Out> ByRef size As D2D1_SIZE_U)
        <PreserveSig>
        Overloads Function GetMaximumBitmapSize() As UInteger
        <PreserveSig>
        Overloads Function IsSupported(renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES) As Boolean
#End Region

        <PreserveSig>
        Function GetBitmap(<Out> ByRef bitmap As ID2D1Bitmap) As HRESULT
    End Interface

    <ComImport>
    <Guid("2cd9069b-12e2-11dc-9fed-001143a055f9")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1Layer
        Inherits ID2D1Resource
#Region "<ID2D1Resource>"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region

        <PreserveSig>
        Sub GetSize(<Out> ByRef size As D2D1_SIZE_F)
    End Interface

    <ComImport>
    <Guid("2cd906c2-12e2-11dc-9fed-001143a055f9")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1Mesh
        Inherits ID2D1Resource
#Region "<ID2D1Resource>"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region
        <PreserveSig>
        Sub Open(<Out> ByRef tessellationSink As IntPtr)
    End Interface

    <ComImport>
    <Guid("2cd90694-12e2-11dc-9fed-001143a055f9")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1RenderTarget
        Inherits ID2D1Resource
#Region "<ID2D1Resource>"

        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)

#End Region

        <PreserveSig>
        Function CreateBitmap(size As D2D1_SIZE_U, srcData As IntPtr, pitch As UInteger, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES, <Out> ByRef bitmap As ID2D1Bitmap) As HRESULT
        <PreserveSig>
        Function CreateBitmapFromWicBitmap(wicBitmapSource As IWICBitmapSource, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES, <Out> ByRef bitmap As ID2D1Bitmap) As HRESULT
        <PreserveSig>
        Function CreateSharedBitmap(ByRef riid As Guid,
        <[In], Out> data As IntPtr, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES, <Out> ByRef bitmap As ID2D1Bitmap) As HRESULT
        <PreserveSig>
        Function CreateBitmapBrush(bitmap As ID2D1Bitmap, ByRef bitmapBrushProperties As D2D1_BITMAP_BRUSH_PROPERTIES, brushProperties As D2D1_BRUSH_PROPERTIES, <Out> ByRef bitmapBrush As ID2D1BitmapBrush) As HRESULT
        <PreserveSig>
        Function CreateSolidColorBrush(color As D2D1_COLOR_F, brushProperties As D2D1_BRUSH_PROPERTIES, <Out> ByRef solidColorBrush As ID2D1SolidColorBrush) As HRESULT
        <PreserveSig>
        Function CreateGradientStopCollection(
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> gradientStops As D2D1_GRADIENT_STOP(), gradientStopsCount As UInteger, colorInterpolationGamma As D2D1_GAMMA, extendMode As D2D1_EXTEND_MODE, <Out> ByRef gradientStopCollection As ID2D1GradientStopCollection) As HRESULT
        <PreserveSig>
        Function CreateLinearGradientBrush(ByRef linearGradientBrushProperties As D2D1_LINEAR_GRADIENT_BRUSH_PROPERTIES, brushProperties As IntPtr, gradientStopCollection As ID2D1GradientStopCollection, <Out> ByRef linearGradientBrush As ID2D1LinearGradientBrush) As HRESULT
        <PreserveSig>
        Function CreateRadialGradientBrush(ByRef radialGradientBrushProperties As D2D1_RADIAL_GRADIENT_BRUSH_PROPERTIES, brushProperties As IntPtr, gradientStopCollection As ID2D1GradientStopCollection, <Out> ByRef radialGradientBrush As ID2D1RadialGradientBrush) As HRESULT
        <PreserveSig>
        Function CreateCompatibleRenderTarget(ByRef desiredSize As D2D1_SIZE_F, ByRef desiredPixelSize As D2D1_SIZE_U, ByRef desiredFormat As D2D1_PIXEL_FORMAT, options As D2D1_COMPATIBLE_RENDER_TARGET_OPTIONS, <Out> ByRef bitmapRenderTarget As ID2D1BitmapRenderTarget) As HRESULT
        <PreserveSig>
        Function CreateLayer(ByRef size As D2D1_SIZE_F, <Out> ByRef layer As ID2D1Layer) As HRESULT
        <PreserveSig>
        Sub CreateMesh(<Out> ByRef mesh As ID2D1Mesh)
        <PreserveSig>
        Sub DrawLine(point0 As D2D1_POINT_2F, point1 As D2D1_POINT_2F, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Sub DrawRectangle(ByRef rect As D2D1_RECT_F, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Sub FillRectangle(ByRef rect As D2D1_RECT_F, brush As ID2D1Brush)
        <PreserveSig>
        Sub DrawRoundedRectangle(ByRef roundedRect As D2D1_ROUNDED_RECT, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Sub FillRoundedRectangle(ByRef roundedRect As D2D1_ROUNDED_RECT, brush As ID2D1Brush)
        <PreserveSig>
        Sub DrawEllipse(ByRef ellipse As D2D1_ELLIPSE, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Sub FillEllipse(ByRef ellipse As D2D1_ELLIPSE, brush As ID2D1Brush)
        <PreserveSig>
        Sub DrawGeometry(geometry As ID2D1Geometry, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Sub FillGeometry(geometry As ID2D1Geometry, brush As ID2D1Brush, Optional opacityBrush As ID2D1Brush = Nothing)
        <PreserveSig>
        Sub FillMesh(mesh As ID2D1Mesh, brush As ID2D1Brush)
        <PreserveSig>
        Sub FillOpacityMask(opacityMask As ID2D1Bitmap, brush As ID2D1Brush, content As D2D1_OPACITY_MASK_CONTENT, destinationRectangle As D2D1_RECT_F, sourceRectangle As D2D1_RECT_F)
        <PreserveSig>
        Sub DrawBitmap(bitmap As ID2D1Bitmap, ByRef destinationRectangle As D2D1_RECT_F, opacity As Single, interpolationMode As D2D1_BITMAP_INTERPOLATION_MODE, ByRef sourceRectangle As D2D1_RECT_F)
        <PreserveSig>
        Sub DrawText(str As String, stringLength As UInteger, textFormat As IDWriteTextFormat, ByRef layoutRect As D2D1_RECT_F, defaultForegroundBrush As ID2D1Brush, options As D2D1_DRAW_TEXT_OPTIONS, measuringMode As DWRITE_MEASURING_MODE)
        <PreserveSig>
        Sub DrawTextLayout(origin As D2D1_POINT_2F, textLayout As IDWriteTextLayout, defaultForegroundBrush As ID2D1Brush, options As D2D1_DRAW_TEXT_OPTIONS)
        <PreserveSig>
        Sub DrawGlyphRun(baselineOrigin As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, foregroundBrush As ID2D1Brush, measuringMode As DWRITE_MEASURING_MODE)
        <PreserveSig>
        Sub SetTransform(transform As D2D1_MATRIX_3X2_F)
        <PreserveSig>
        Sub GetTransform(<Out> ByRef transform As D2D1_MATRIX_3X2_F_STRUCT)
        <PreserveSig>
        Sub SetAntialiasMode(antialiasMode As D2D1_ANTIALIAS_MODE)
        <PreserveSig>
        Sub GetAntialiasMode(<Out> ByRef antialiasMode As D2D1_ANTIALIAS_MODE)
        <PreserveSig>
        Sub SetTextAntialiasMode(textAntialiasMode As D2D1_TEXT_ANTIALIAS_MODE)
        <PreserveSig>
        Sub GetTextAntialiasMode(<Out> ByRef textAntialiasMode As D2D1_TEXT_ANTIALIAS_MODE)
        <PreserveSig>
        Sub SetTextRenderingParams(Optional textRenderingParams As IDWriteRenderingParams = Nothing)
        <PreserveSig>
        Sub GetTextRenderingParams(<Out> ByRef textRenderingParams As IDWriteRenderingParams)
        <PreserveSig>
        Sub SetTags(tag1 As ULong, tag2 As ULong)
        <PreserveSig>
        Sub GetTags(<Out> ByRef tag1 As ULong, <Out> ByRef tag2 As ULong)
        <PreserveSig>
        Sub PushLayer(ByRef layerParameters As D2D1_LAYER_PARAMETERS, layer As ID2D1Layer)
        <PreserveSig>
        Sub PopLayer()
        <PreserveSig>
        Sub Flush(<Out> ByRef tag1 As ULong, <Out> ByRef tag2 As ULong)
        <PreserveSig>
        Sub SaveDrawingState(
        <[In], Out> drawingStateBlock As ID2D1DrawingStateBlock)
        <PreserveSig>
        Sub RestoreDrawingState(drawingStateBlock As ID2D1DrawingStateBlock)
        <PreserveSig>
        Sub PushAxisAlignedClip(clipRect As D2D1_RECT_F, antialiasMode As D2D1_ANTIALIAS_MODE)
        <PreserveSig>
        Sub PopAxisAlignedClip()
        <PreserveSig>
        Sub Clear(clearColor As D2D1_COLOR_F)
        <PreserveSig>
        Sub BeginDraw()
        <PreserveSig>
        Function EndDraw(<Out> ByRef tag1 As ULong, <Out> ByRef tag2 As ULong) As HRESULT
        <PreserveSig>
        Sub GetPixelFormat(<Out> ByRef pixelFormat As D2D1_PIXEL_FORMAT)
        <PreserveSig>
        Sub SetDpi(dpiX As Single, dpiY As Single)
        <PreserveSig>
        Sub GetDpi(<Out> ByRef dpiX As Single, <Out> ByRef dpiY As Single)
        <PreserveSig>
        Sub GetSize(<Out> ByRef size As D2D1_SIZE_F)
        <PreserveSig>
        Sub GetPixelSize(<Out> ByRef size As D2D1_SIZE_U)
        <PreserveSig>
        Function GetMaximumBitmapSize() As UInteger
        <PreserveSig>
        Function IsSupported(renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES) As Boolean

        'HRESULT Function1();
        'HRESULT Function2();
        'HRESULT Function3();
        'HRESULT Function4();
        'HRESULT Function5();
        'HRESULT Function6();
        'HRESULT Function7();
        'HRESULT Function8();
        'HRESULT Function9();
        'HRESULT Function10();
        'HRESULT Function11();
        'HRESULT Function12();
        'HRESULT Function13();
        'HRESULT Function14();
        'HRESULT Function15();
        'HRESULT Function16();
        'HRESULT Function17();
        'HRESULT Function18();
        'HRESULT Function19();
        'HRESULT Function20();
        'HRESULT Function21();
        'void Function22();
        'void Function23();
        'void Function24();
        'void Function25();
        'void Function26();
        'void Function27();
        'void Function28();
        'void Function29();
        'void Function30();
        'void Function31();
        'void Function32();
        'void Function33();
        'void Function34();
        'void Function35();
        'bool Function36();
    End Interface

    ' Incomplete
    <ComImport>
    <Guid("06152247-6f50-465a-9245-118bfd3b6007")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1Factory
        <PreserveSig>
        Function ReloadSystemMetrics() As HRESULT
        <PreserveSig>
        Function GetDesktopDpi(<Out> ByRef dpiX As Single, <Out> ByRef dpiY As Single) As HRESULT
        <PreserveSig>
        Function CreateRectangleGeometry(ByRef rectangle As D2D1_RECT_F, <Out> ByRef rectangleGeometry As ID2D1RectangleGeometry) As HRESULT
        <PreserveSig>
        Function CreateRoundedRectangleGeometry(ByRef roundedRectangle As D2D1_ROUNDED_RECT, <Out> ByRef roundedRectangleGeometry As ID2D1RoundedRectangleGeometry) As HRESULT
        <PreserveSig>
        Function CreateEllipseGeometry(ByRef ellipse As D2D1_ELLIPSE, <Out> ByRef ellipseGeometry As ID2D1EllipseGeometry) As HRESULT
        <PreserveSig>
        Function CreateGeometryGroup(fillMode As D2D1_FILL_MODE, geometries As ID2D1Geometry, geometriesCount As UInteger, <Out> ByRef geometryGroup As ID2D1GeometryGroup) As HRESULT
        <PreserveSig>
        Function CreateTransformedGeometry(sourceGeometry As ID2D1Geometry, transform As D2D1_MATRIX_3X2_F, <Out> ByRef transformedGeometry As ID2D1TransformedGeometry) As HRESULT
        <PreserveSig>
        Function CreatePathGeometry(<Out> ByRef pathGeometry As ID2D1PathGeometry) As HRESULT
        <PreserveSig>
        Function CreateStrokeStyle(ByRef strokeStyleProperties As D2D1_STROKE_STYLE_PROPERTIES,
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=2)> dashes As Single(), Optional dashesCount As UInteger = 0) As ID2D1StrokeStyle
        <PreserveSig>
        Function CreateDrawingStateBlock(ByRef drawingStateDescription As D2D1_DRAWING_STATE_DESCRIPTION, textRenderingParams As IDWriteRenderingParams, <Out> ByRef drawingStateBlock As ID2D1DrawingStateBlock) As HRESULT
        <PreserveSig>
        Function CreateWicBitmapRenderTarget(target As IWICBitmap, renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES, <Out> ByRef renderTarget As ID2D1RenderTarget) As HRESULT
        <PreserveSig>
        Function CreateHwndRenderTarget(ByRef renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES, ByRef hwndRenderTargetProperties As D2D1_HWND_RENDER_TARGET_PROPERTIES, <Out> ByRef hwndRenderTarget As ID2D1HwndRenderTarget) As HRESULT
        <PreserveSig>
        Function CreateDxgiSurfaceRenderTarget(dxgiSurface As IntPtr, ByRef renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES, ByRef renderTarget As ID2D1RenderTarget) As HRESULT
        <PreserveSig>
        Function CreateDCRenderTarget(ByRef renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES, ByRef dcRenderTarget As ID2D1DCRenderTarget) As HRESULT
    End Interface


    <ComImport>
    <Guid("bb12d362-daee-4b9a-aa1d-14ba401cfa1f")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1Factory1
        Inherits ID2D1Factory
#Region "<ID2D1Factory>"
        Overloads Function ReloadSystemMetrics() As HRESULT
        <PreserveSig>
        Overloads Function GetDesktopDpi(<Out> ByRef dpiX As Single, <Out> ByRef dpiY As Single) As HRESULT
        <PreserveSig>
        Overloads Function CreateRectangleGeometry(ByRef rectangle As D2D1_RECT_F, <Out> ByRef rectangleGeometry As ID2D1RectangleGeometry) As HRESULT
        <PreserveSig>
        Overloads Function CreateRoundedRectangleGeometry(ByRef roundedRectangle As D2D1_ROUNDED_RECT, <Out> ByRef roundedRectangleGeometry As ID2D1RoundedRectangleGeometry) As HRESULT
        <PreserveSig>
        Overloads Function CreateEllipseGeometry(ByRef ellipse As D2D1_ELLIPSE, <Out> ByRef ellipseGeometry As ID2D1EllipseGeometry) As HRESULT
        <PreserveSig>
        Overloads Function CreateGeometryGroup(fillMode As D2D1_FILL_MODE, geometries As ID2D1Geometry, geometriesCount As UInteger, <Out> ByRef geometryGroup As ID2D1GeometryGroup) As HRESULT
        <PreserveSig>
        Overloads Function CreateTransformedGeometry(sourceGeometry As ID2D1Geometry, transform As D2D1_MATRIX_3X2_F, <Out> ByRef transformedGeometry As ID2D1TransformedGeometry) As HRESULT
        <PreserveSig>
        Overloads Function CreatePathGeometry(<Out> ByRef pathGeometry As ID2D1PathGeometry) As HRESULT
        <PreserveSig>
        Overloads Function CreateStrokeStyle(ByRef strokeStyleProperties As D2D1_STROKE_STYLE_PROPERTIES,
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=2)> dashes As Single(), Optional dashesCount As UInteger = 0) As ID2D1StrokeStyle
        <PreserveSig>
        Overloads Function CreateDrawingStateBlock(ByRef drawingStateDescription As D2D1_DRAWING_STATE_DESCRIPTION, textRenderingParams As IDWriteRenderingParams, <Out> ByRef drawingStateBlock As ID2D1DrawingStateBlock) As HRESULT
        <PreserveSig>
        Overloads Function CreateWicBitmapRenderTarget(target As IWICBitmap, renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES, <Out> ByRef renderTarget As ID2D1RenderTarget) As HRESULT
        <PreserveSig>
        Overloads Function CreateHwndRenderTarget(ByRef renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES, ByRef hwndRenderTargetProperties As D2D1_HWND_RENDER_TARGET_PROPERTIES, <Out> ByRef hwndRenderTarget As ID2D1HwndRenderTarget) As HRESULT
        <PreserveSig>
        Overloads Function CreateDxgiSurfaceRenderTarget(dxgiSurface As IntPtr, ByRef renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES, ByRef renderTarget As ID2D1RenderTarget) As HRESULT
        <PreserveSig>
        Overloads Function CreateDCRenderTarget(ByRef renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES, ByRef dcRenderTarget As ID2D1DCRenderTarget) As HRESULT
#End Region

        Function CreateDevice(dxgiDevice As IDXGIDevice, <Out> ByRef d2dDevice As ID2D1Device) As HRESULT
        <PreserveSig>
        Overloads Function CreateStrokeStyle(ByRef strokeStyleProperties As D2D1_STROKE_STYLE_PROPERTIES1,
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=2)> dashes As Single(), dashesCount As UInteger, <Out> ByRef strokeStyle As ID2D1StrokeStyle1) As HRESULT
        <PreserveSig>
        Overloads Function CreatePathGeometry(<Out> ByRef pathGeometry As ID2D1PathGeometry1) As HRESULT
        <PreserveSig>
        Overloads Function CreateDrawingStateBlock(ByRef drawingStateDescription As D2D1_DRAWING_STATE_DESCRIPTION1, ByRef textRenderingParams As IDWriteRenderingParams, <Out> ByRef drawingStateBlock As ID2D1DrawingStateBlock1) As HRESULT
        <PreserveSig>
        Function CreateGdiMetafile(metafileStream As ComTypes.IStream, <Out> ByRef metafile As ID2D1GdiMetafile) As HRESULT
        <PreserveSig>
        Function RegisterEffectFromStream(
        <MarshalAs(UnmanagedType.LPStruct)> classId As Guid, propertyXml As ComTypes.IStream,
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=3)> bindings As D2D1_PROPERTY_BINDING(), bindingsCount As Integer, effectFactory As PD2D1_EFFECT_FACTORY) As HRESULT
        <PreserveSig>
        Function RegisterEffectFromString(
        <MarshalAs(UnmanagedType.LPStruct)> classId As Guid,
        <MarshalAs(UnmanagedType.LPWStr)> propertyXml As String,
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=3)> bindings As D2D1_PROPERTY_BINDING(), bindingsCount As Integer, effectFactory As PD2D1_EFFECT_FACTORY) As HRESULT
        <PreserveSig>
        Function UnregisterEffect(
        <MarshalAs(UnmanagedType.LPStruct)> classId As Guid) As HRESULT
        <PreserveSig>
        Function GetRegisteredEffects(
        <Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> effects As Guid(), effectsCount As Integer, effectsReturned As IntPtr, effectsRegistered As IntPtr) As HRESULT
        <PreserveSig>
        Function GetEffectProperties(
        <MarshalAs(UnmanagedType.LPStruct)> effectId As Guid, <Out> ByRef properties As ID2D1Properties) As HRESULT
    End Interface

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_DRAWING_STATE_DESCRIPTION1
        Public antialiasMode As D2D1_ANTIALIAS_MODE
        Public textAntialiasMode As D2D1_TEXT_ANTIALIAS_MODE
        Public tag1 As ULong
        Public tag2 As ULong
        Public transform As D2D1_MATRIX_3X2_F
        Public primitiveBlend As D2D1_PRIMITIVE_BLEND
        Public unitMode As D2D1_UNIT_MODE
    End Structure

    <ComImport>
    <Guid("689f1f85-c72e-4e33-8f19-85754efd5ace")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1DrawingStateBlock1
        Inherits ID2D1DrawingStateBlock
#Region "ID2D1DrawingStateBlock"
#Region "ID2D1Resource"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region

        Overloads Sub GetDescription(<Out> ByRef stateDescription As D2D1_DRAWING_STATE_DESCRIPTION)
        Overloads Sub SetDescription(stateDescription As D2D1_DRAWING_STATE_DESCRIPTION)
        Overloads Sub SetTextRenderingParams(Optional textRenderingParams As IDWriteRenderingParams = Nothing)
        Overloads Sub GetTextRenderingParams(<Out> ByRef textRenderingParams As IDWriteRenderingParams)
#End Region

        <PreserveSig>
        Overloads Sub GetDescription(<Out> ByRef stateDescription As D2D1_DRAWING_STATE_DESCRIPTION1)
        <PreserveSig>
        Overloads Sub SetDescription(stateDescription As D2D1_DRAWING_STATE_DESCRIPTION1)
    End Interface

    <ComImport>
    <Guid("10a72a66-e91c-43f4-993f-ddf4b82b0b4a")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1StrokeStyle1
        Inherits ID2D1StrokeStyle
#Region "<ID2D1StrokeStyle>"
#Region "<ID2D1Resource>"
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region
        <PreserveSig>
        Overloads Sub GetStartCap(<Out> ByRef capStyle As D2D1_CAP_STYLE)
        <PreserveSig>
        Overloads Sub GetEndCap(<Out> ByRef capStyle As D2D1_CAP_STYLE)
        <PreserveSig>
        Overloads Sub GetDashCap(<Out> ByRef capStyle As D2D1_CAP_STYLE)
        <PreserveSig>
        Overloads Function GetMiterLimit() As Single
        <PreserveSig>
        Overloads Sub GetLineJoin(<Out> ByRef lineJoin As D2D1_LINE_JOIN)
        <PreserveSig>
        Overloads Function GetDashOffset() As Single
        <PreserveSig>
        Overloads Sub GetDashStyle(<Out> ByRef dashStyle As D2D1_DASH_STYLE)
        <PreserveSig>
        Overloads Function GetDashesCount() As UInteger
        <PreserveSig>
        Overloads Sub GetDashes(<Out> ByRef dashes As Single, dashesCount As UInteger)
#End Region

        <PreserveSig>
        Sub GetStrokeTransformType(<Out> ByRef strokeTransformType As D2D1_STROKE_TRANSFORM_TYPE)
    End Interface

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_STROKE_STYLE_PROPERTIES1
        Public startCap As D2D1_CAP_STYLE
        Public endCap As D2D1_CAP_STYLE
        Public dashCap As D2D1_CAP_STYLE
        Public lineJoin As D2D1_LINE_JOIN
        Public miterLimit As Single
        Public dashStyle As D2D1_DASH_STYLE
        Public dashOffset As Single
        ''' <summary>
        ''' How the nib of the stroke is influenced by the context properties.
        ''' </summary>
        Public transformType As D2D1_STROKE_TRANSFORM_TYPE
    End Structure

    Public Enum D2D1_STROKE_TRANSFORM_TYPE
        ''' <summary>
        ''' The stroke respects the world transform, the DPI, and the stroke width.
        ''' </summary>
        D2D1_STROKE_TRANSFORM_TYPE_NORMAL = 0

        ''' <summary>
        ''' The stroke does not respect the world transform, but it does respect the DPI and
        ''' the stroke width.
        ''' </summary>
        D2D1_STROKE_TRANSFORM_TYPE_FIXED = 1

        ''' <summary>
        ''' The stroke is forced to one pixel wide.
        ''' </summary>
        D2D1_STROKE_TRANSFORM_TYPE_HAIRLINE = 2
        D2D1_STROKE_TRANSFORM_TYPE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    <ComImport()>
    <Guid("1c51bc64-de61-46fd-9899-63a5d8f03950")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1DCRenderTarget
        Inherits ID2D1RenderTarget
#Region "<ID2D1Resource>"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region

        <PreserveSig>
        Overloads Function CreateBitmap(size As D2D1_SIZE_U, srcData As IntPtr, pitch As UInteger, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES, <Out> ByRef bitmap As ID2D1Bitmap) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapFromWicBitmap(wicBitmapSource As IWICBitmapSource, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES, <Out> ByRef bitmap As ID2D1Bitmap) As HRESULT
        <PreserveSig>
        Overloads Function CreateSharedBitmap(ByRef riid As Guid,
        <[In], Out> data As IntPtr, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES, <Out> ByRef bitmap As ID2D1Bitmap) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapBrush(bitmap As ID2D1Bitmap, ByRef bitmapBrushProperties As D2D1_BITMAP_BRUSH_PROPERTIES, brushProperties As D2D1_BRUSH_PROPERTIES, <Out> ByRef bitmapBrush As ID2D1BitmapBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateSolidColorBrush(color As D2D1_COLOR_F, brushProperties As D2D1_BRUSH_PROPERTIES, <Out> ByRef solidColorBrush As ID2D1SolidColorBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateGradientStopCollection(
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> gradientStops As D2D1_GRADIENT_STOP(), gradientStopsCount As UInteger, colorInterpolationGamma As D2D1_GAMMA, extendMode As D2D1_EXTEND_MODE, <Out> ByRef gradientStopCollection As ID2D1GradientStopCollection) As HRESULT
        <PreserveSig>
        Overloads Function CreateLinearGradientBrush(ByRef linearGradientBrushProperties As D2D1_LINEAR_GRADIENT_BRUSH_PROPERTIES, brushProperties As IntPtr, gradientStopCollection As ID2D1GradientStopCollection, <Out> ByRef linearGradientBrush As ID2D1LinearGradientBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateRadialGradientBrush(ByRef radialGradientBrushProperties As D2D1_RADIAL_GRADIENT_BRUSH_PROPERTIES, brushProperties As IntPtr, gradientStopCollection As ID2D1GradientStopCollection, <Out> ByRef radialGradientBrush As ID2D1RadialGradientBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateCompatibleRenderTarget(ByRef desiredSize As D2D1_SIZE_F, ByRef desiredPixelSize As D2D1_SIZE_U, ByRef desiredFormat As D2D1_PIXEL_FORMAT, options As D2D1_COMPATIBLE_RENDER_TARGET_OPTIONS, <Out> ByRef bitmapRenderTarget As ID2D1BitmapRenderTarget) As HRESULT
        <PreserveSig>
        Overloads Function CreateLayer(ByRef size As D2D1_SIZE_F, <Out> ByRef layer As ID2D1Layer) As HRESULT
        <PreserveSig>
        Overloads Sub CreateMesh(<Out> ByRef mesh As ID2D1Mesh)
        <PreserveSig>
        Overloads Sub DrawLine(point0 As D2D1_POINT_2F, point1 As D2D1_POINT_2F, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub DrawRectangle(ByRef rect As D2D1_RECT_F, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillRectangle(ByRef rect As D2D1_RECT_F, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub DrawRoundedRectangle(ByRef roundedRect As D2D1_ROUNDED_RECT, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillRoundedRectangle(ByRef roundedRect As D2D1_ROUNDED_RECT, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub DrawEllipse(ByRef ellipse As D2D1_ELLIPSE, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillEllipse(ByRef ellipse As D2D1_ELLIPSE, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub DrawGeometry(geometry As ID2D1Geometry, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillGeometry(geometry As ID2D1Geometry, brush As ID2D1Brush, Optional opacityBrush As ID2D1Brush = Nothing)
        <PreserveSig>
        Overloads Sub FillMesh(mesh As ID2D1Mesh, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub FillOpacityMask(opacityMask As ID2D1Bitmap, brush As ID2D1Brush, content As D2D1_OPACITY_MASK_CONTENT, destinationRectangle As D2D1_RECT_F, sourceRectangle As D2D1_RECT_F)
        <PreserveSig>
        Overloads Sub DrawBitmap(bitmap As ID2D1Bitmap, ByRef destinationRectangle As D2D1_RECT_F, opacity As Single, interpolationMode As D2D1_BITMAP_INTERPOLATION_MODE, ByRef sourceRectangle As D2D1_RECT_F)
        <PreserveSig>
        Overloads Sub DrawText(str As String, stringLength As UInteger, textFormat As IDWriteTextFormat, ByRef layoutRect As D2D1_RECT_F, defaultForegroundBrush As ID2D1Brush, options As D2D1_DRAW_TEXT_OPTIONS, measuringMode As DWRITE_MEASURING_MODE)
        <PreserveSig>
        Overloads Sub DrawTextLayout(origin As D2D1_POINT_2F, textLayout As IDWriteTextLayout, defaultForegroundBrush As ID2D1Brush, options As D2D1_DRAW_TEXT_OPTIONS)
        <PreserveSig>
        Overloads Sub DrawGlyphRun(baselineOrigin As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, foregroundBrush As ID2D1Brush, measuringMode As DWRITE_MEASURING_MODE)
        <PreserveSig>
        Overloads Sub SetTransform(transform As D2D1_MATRIX_3X2_F)
        <PreserveSig>
        Overloads Sub GetTransform(<Out> ByRef transform As D2D1_MATRIX_3X2_F_STRUCT)
        <PreserveSig>
        Overloads Sub SetAntialiasMode(antialiasMode As D2D1_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub GetAntialiasMode(<Out> ByRef antialiasMode As D2D1_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub SetTextAntialiasMode(textAntialiasMode As D2D1_TEXT_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub GetTextAntialiasMode(<Out> ByRef textAntialiasMode As D2D1_TEXT_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub SetTextRenderingParams(Optional textRenderingParams As IDWriteRenderingParams = Nothing)
        <PreserveSig>
        Overloads Sub GetTextRenderingParams(<Out> ByRef textRenderingParams As IDWriteRenderingParams)
        <PreserveSig>
        Overloads Sub SetTags(tag1 As ULong, tag2 As ULong)
        <PreserveSig>
        Overloads Sub GetTags(<Out> ByRef tag1 As ULong, <Out> ByRef tag2 As ULong)
        <PreserveSig>
        Overloads Sub PushLayer(ByRef layerParameters As D2D1_LAYER_PARAMETERS, layer As ID2D1Layer)
        <PreserveSig>
        Overloads Sub PopLayer()
        <PreserveSig>
        Overloads Sub Flush(<Out> ByRef tag1 As ULong, <Out> ByRef tag2 As ULong)
        <PreserveSig>
        Overloads Sub SaveDrawingState(drawingStateBlock As ID2D1DrawingStateBlock)
        <PreserveSig>
        Overloads Sub RestoreDrawingState(drawingStateBlock As ID2D1DrawingStateBlock)
        <PreserveSig>
        Overloads Sub PushAxisAlignedClip(clipRect As D2D1_RECT_F, antialiasMode As D2D1_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub PopAxisAlignedClip()
        <PreserveSig>
        Overloads Sub Clear(clearColor As D2D1_COLOR_F)
        <PreserveSig>
        Overloads Sub BeginDraw()
        <PreserveSig>
        Overloads Function EndDraw(<Out> ByRef tag1 As ULong, <Out> ByRef tag2 As ULong) As HRESULT
        <PreserveSig>
        Overloads Sub GetPixelFormat(<Out> ByRef pixelFormat As D2D1_PIXEL_FORMAT)
        <PreserveSig>
        Overloads Sub SetDpi(dpiX As Single, dpiY As Single)
        <PreserveSig>
        Overloads Sub GetDpi(<Out> ByRef dpiX As Single, <Out> ByRef dpiY As Single)
        <PreserveSig>
        Overloads Sub GetSize(<Out> ByRef size As D2D1_SIZE_F)
        <PreserveSig>
        Overloads Sub GetPixelSize(<Out> ByRef size As D2D1_SIZE_U)
        <PreserveSig>
        Overloads Function GetMaximumBitmapSize() As UInteger
        <PreserveSig>
        Overloads Function IsSupported(renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES) As Boolean
        'new HRESULT Function1();
        'new HRESULT Function2();
        'new HRESULT Function3();
        'new HRESULT Function4();
        'new HRESULT Function5();
        'new HRESULT Function6();
        'new HRESULT Function7();
        'new HRESULT Function8();
        'new HRESULT Function9();
        'new HRESULT Function10();
        'new HRESULT Function11();
        'new HRESULT Function12();
        'new HRESULT Function13();
        'new HRESULT Function14();
        'new HRESULT Function15();
        'new HRESULT Function16();
        'new HRESULT Function17();
        'new HRESULT Function18();
        'new HRESULT Function19();
        'new HRESULT Function20();
        'new HRESULT Function21();
        'new void Function22();
        'new void Function23();
        'new void Function24();
        'new void Function25();
        'new void Function26();
        'new void Function27();
        'new void Function28();
        'new void Function29();
        'new void Function30();
        'new void Function31();
        'new void Function32();
        'new void Function33();
        'new void Function34();
        'new void Function35();
        'new bool Function36();

        <PreserveSig>
        Function BindDC(hDC As IntPtr, ByRef pSubRect As RECT) As HRESULT
    End Interface

    <ComImport>
    <Guid("28211a43-7d89-476f-8181-2d6159b220ad")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1Effect
        Inherits ID2D1Properties
#Region "<ID2D1Properties>"
        <PreserveSig>
        Overloads Function GetPropertyCount() As UInteger
        <PreserveSig>
        Overloads Function GetPropertyName(index As UInteger, <Out> ByRef name As String, nameCount As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function GetPropertyNameLength(index As UInteger) As UInteger
        <PreserveSig>
        Overloads Function [GetType](index As UInteger) As D2D1_PROPERTY_TYPE
        <PreserveSig>
        Overloads Function GetPropertyIndex(name As String) As UInteger
        <PreserveSig>
        Overloads Function SetValueByName(name As String, type As D2D1_PROPERTY_TYPE, data As IntPtr, dataSize As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function SetValue(index As UInteger, type As D2D1_PROPERTY_TYPE, data As IntPtr, dataSize As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function GetValueByName(name As String, type As D2D1_PROPERTY_TYPE, <Out> ByRef data As IntPtr, dataSize As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function GetValue(index As UInteger, type As D2D1_PROPERTY_TYPE, <Out> ByRef data As IntPtr, dataSize As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function GetValueSize(index As UInteger) As UInteger
        <PreserveSig>
        Overloads Function GetSubProperties(index As UInteger, <Out> ByRef subProperties As ID2D1Properties) As HRESULT
#End Region

        <PreserveSig>
        Sub SetInput(index As UInteger, input As ID2D1Image, Optional invalidate As Boolean = True)
        <PreserveSig>
        Function SetInputCount(inputCount As UInteger) As HRESULT
        <PreserveSig>
        Sub GetInput(index As UInteger, <Out> ByRef input As ID2D1Image)
        <PreserveSig>
        Function GetInputCount() As UInteger
        <PreserveSig>
        Sub GetOutput(<Out> ByRef outputImage As ID2D1Image)
    End Interface

    <ComImport>
    <Guid("483473d7-cd46-4f9d-9d3a-3112aa80159d")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1Properties
        <PreserveSig>
        Function GetPropertyCount() As UInteger
        <PreserveSig>
        Function GetPropertyName(index As UInteger, <Out> ByRef name As String, nameCount As UInteger) As HRESULT
        <PreserveSig>
        Function GetPropertyNameLength(index As UInteger) As UInteger
        <PreserveSig>
        Function [GetType](index As UInteger) As D2D1_PROPERTY_TYPE
        <PreserveSig>
        Function GetPropertyIndex(name As String) As UInteger
        <PreserveSig>
        Function SetValueByName(name As String, type As D2D1_PROPERTY_TYPE, data As IntPtr, dataSize As UInteger) As HRESULT
        <PreserveSig>
        Function SetValue(index As UInteger, type As D2D1_PROPERTY_TYPE, data As IntPtr, dataSize As UInteger) As HRESULT
        <PreserveSig>
        Function GetValueByName(name As String, type As D2D1_PROPERTY_TYPE, <Out> ByRef data As IntPtr, dataSize As UInteger) As HRESULT
        <PreserveSig>
        Function GetValue(index As UInteger, type As D2D1_PROPERTY_TYPE, <Out> ByRef data As IntPtr, dataSize As UInteger) As HRESULT
        <PreserveSig>
        Function GetValueSize(index As UInteger) As UInteger
        <PreserveSig>
        Function GetSubProperties(index As UInteger, <Out> ByRef subProperties As ID2D1Properties) As HRESULT
    End Interface

    Public Enum D2D1_PROPERTY_TYPE
        D2D1_PROPERTY_TYPE_UNKNOWN = 0
        D2D1_PROPERTY_TYPE_STRING = 1
        D2D1_PROPERTY_TYPE_BOOL = 2
        D2D1_PROPERTY_TYPE_UINT32 = 3
        D2D1_PROPERTY_TYPE_INT32 = 4
        D2D1_PROPERTY_TYPE_FLOAT = 5
        D2D1_PROPERTY_TYPE_VECTOR2 = 6
        D2D1_PROPERTY_TYPE_VECTOR3 = 7
        D2D1_PROPERTY_TYPE_VECTOR4 = 8
        D2D1_PROPERTY_TYPE_BLOB = 9
        D2D1_PROPERTY_TYPE_IUNKNOWN = 10
        D2D1_PROPERTY_TYPE_ENUM = 11
        D2D1_PROPERTY_TYPE_ARRAY = 12
        D2D1_PROPERTY_TYPE_CLSID = 13
        D2D1_PROPERTY_TYPE_MATRIX_3X2 = 14
        D2D1_PROPERTY_TYPE_MATRIX_4X3 = 15
        D2D1_PROPERTY_TYPE_MATRIX_4X4 = 16
        D2D1_PROPERTY_TYPE_MATRIX_5X4 = 17
        D2D1_PROPERTY_TYPE_COLOR_CONTEXT = 18
        D2D1_PROPERTY_TYPE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_PROPERTY
        D2D1_PROPERTY_CLSID = &H80000000
        D2D1_PROPERTY_DISPLAYNAME = &H80000001
        D2D1_PROPERTY_AUTHOR = &H80000002
        D2D1_PROPERTY_CATEGORY = &H80000003
        D2D1_PROPERTY_DESCRIPTION = &H80000004
        D2D1_PROPERTY_INPUTS = &H80000005
        D2D1_PROPERTY_CACHED = &H80000006
        D2D1_PROPERTY_PRECISION = &H80000007
        D2D1_PROPERTY_MIN_INPUTS = &H80000008
        D2D1_PROPERTY_MAX_INPUTS = &H80000009
        D2D1_PROPERTY_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_SUBPROPERTY
        D2D1_SUBPROPERTY_DISPLAYNAME = &H80000000
        D2D1_SUBPROPERTY_ISREADONLY = &H80000001
        D2D1_SUBPROPERTY_MIN = &H80000002
        D2D1_SUBPROPERTY_MAX = &H80000003
        D2D1_SUBPROPERTY_DEFAULT = &H80000004
        D2D1_SUBPROPERTY_FIELDS = &H80000005
        D2D1_SUBPROPERTY_INDEX = &H80000006
        D2D1_SUBPROPERTY_FORCE_DWORD = &HFFFFFFFF
    End Enum


    <ComImport>
    <Guid("e8f7fe7a-191c-466d-ad95-975678bda998")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1DeviceContext
        Inherits ID2D1RenderTarget
#Region "<ID2D1RenderTarget>"

#Region "<ID2D1Resource>"

        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)

#End Region
        <PreserveSig>
        Overloads Function CreateBitmap(size As D2D1_SIZE_U, srcData As IntPtr, pitch As UInteger, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES, <Out> ByRef bitmap As ID2D1Bitmap) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapFromWicBitmap(wicBitmapSource As IWICBitmapSource, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES, <Out> ByRef bitmap As ID2D1Bitmap) As HRESULT
        <PreserveSig>
        Overloads Function CreateSharedBitmap(ByRef riid As Guid,
        <[In], Out> data As IntPtr, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES, <Out> ByRef bitmap As ID2D1Bitmap) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapBrush(bitmap As ID2D1Bitmap, ByRef bitmapBrushProperties As D2D1_BITMAP_BRUSH_PROPERTIES, brushProperties As D2D1_BRUSH_PROPERTIES, <Out> ByRef bitmapBrush As ID2D1BitmapBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateSolidColorBrush(color As D2D1_COLOR_F, brushProperties As D2D1_BRUSH_PROPERTIES, <Out> ByRef solidColorBrush As ID2D1SolidColorBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateGradientStopCollection(
            <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> gradientStops As D2D1_GRADIENT_STOP(), gradientStopsCount As UInteger, colorInterpolationGamma As D2D1_GAMMA, extendMode As D2D1_EXTEND_MODE, <Out> ByRef gradientStopCollection As ID2D1GradientStopCollection) As HRESULT
        'new HRESULT CreateLinearGradientBrush(ref D2D1_LINEAR_GRADIENT_BRUSH_PROPERTIES linearGradientBrushProperties, D2D1_BRUSH_PROPERTIES brushProperties, ID2D1GradientStopCollection gradientStopCollection, out ID2D1LinearGradientBrush linearGradientBrush);
        <PreserveSig>
        Overloads Function CreateLinearGradientBrush(ByRef linearGradientBrushProperties As D2D1_LINEAR_GRADIENT_BRUSH_PROPERTIES, brushProperties As IntPtr, gradientStopCollection As ID2D1GradientStopCollection, <Out> ByRef linearGradientBrush As ID2D1LinearGradientBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateRadialGradientBrush(ByRef radialGradientBrushProperties As D2D1_RADIAL_GRADIENT_BRUSH_PROPERTIES, brushProperties As IntPtr, gradientStopCollection As ID2D1GradientStopCollection, <Out> ByRef radialGradientBrush As ID2D1RadialGradientBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateCompatibleRenderTarget(ByRef desiredSize As D2D1_SIZE_F, ByRef desiredPixelSize As D2D1_SIZE_U, ByRef desiredFormat As D2D1_PIXEL_FORMAT, options As D2D1_COMPATIBLE_RENDER_TARGET_OPTIONS, <Out> ByRef bitmapRenderTarget As ID2D1BitmapRenderTarget) As HRESULT
        <PreserveSig>
        Overloads Function CreateLayer(ByRef size As D2D1_SIZE_F, <Out> ByRef layer As ID2D1Layer) As HRESULT
        Overloads Sub CreateMesh(<Out> ByRef mesh As ID2D1Mesh)
        <PreserveSig>
        Overloads Sub DrawLine(point0 As D2D1_POINT_2F, point1 As D2D1_POINT_2F, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub DrawRectangle(ByRef rect As D2D1_RECT_F, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillRectangle(ByRef rect As D2D1_RECT_F, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub DrawRoundedRectangle(ByRef roundedRect As D2D1_ROUNDED_RECT, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillRoundedRectangle(ByRef roundedRect As D2D1_ROUNDED_RECT, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub DrawEllipse(ByRef ellipse As D2D1_ELLIPSE, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillEllipse(ByRef ellipse As D2D1_ELLIPSE, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub DrawGeometry(geometry As ID2D1Geometry, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillGeometry(geometry As ID2D1Geometry, brush As ID2D1Brush, Optional opacityBrush As ID2D1Brush = Nothing)
        <PreserveSig>
        Overloads Sub FillMesh(mesh As ID2D1Mesh, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub FillOpacityMask(opacityMask As ID2D1Bitmap, brush As ID2D1Brush, content As D2D1_OPACITY_MASK_CONTENT, destinationRectangle As D2D1_RECT_F, sourceRectangle As D2D1_RECT_F)
        <PreserveSig>
        Overloads Sub DrawBitmap(bitmap As ID2D1Bitmap, ByRef destinationRectangle As D2D1_RECT_F, opacity As Single, interpolationMode As D2D1_BITMAP_INTERPOLATION_MODE, ByRef sourceRectangle As D2D1_RECT_F)
        <PreserveSig>
        Overloads Sub DrawText(str As String, stringLength As UInteger, textFormat As IDWriteTextFormat, ByRef layoutRect As D2D1_RECT_F, defaultForegroundBrush As ID2D1Brush, options As D2D1_DRAW_TEXT_OPTIONS, measuringMode As DWRITE_MEASURING_MODE)
        <PreserveSig>
        Overloads Sub DrawTextLayout(origin As D2D1_POINT_2F, textLayout As IDWriteTextLayout, defaultForegroundBrush As ID2D1Brush, options As D2D1_DRAW_TEXT_OPTIONS)
        <PreserveSig>
        Overloads Sub DrawGlyphRun(baselineOrigin As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, foregroundBrush As ID2D1Brush, measuringMode As DWRITE_MEASURING_MODE)
        <PreserveSig>
        Overloads Sub SetTransform(transform As D2D1_MATRIX_3X2_F)
        <PreserveSig>
        Overloads Sub GetTransform(<Out> ByRef transform As D2D1_MATRIX_3X2_F_STRUCT)
        <PreserveSig>
        Overloads Sub SetAntialiasMode(antialiasMode As D2D1_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub GetAntialiasMode(<Out> ByRef antialiasMode As D2D1_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub SetTextAntialiasMode(textAntialiasMode As D2D1_TEXT_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub GetTextAntialiasMode(<Out> ByRef textAntialiasMode As D2D1_TEXT_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub SetTextRenderingParams(Optional textRenderingParams As IDWriteRenderingParams = Nothing)
        <PreserveSig>
        Overloads Sub GetTextRenderingParams(<Out> ByRef textRenderingParams As IDWriteRenderingParams)
        <PreserveSig>
        Overloads Sub SetTags(tag1 As ULong, tag2 As ULong)
        <PreserveSig>
        Overloads Sub GetTags(<Out> ByRef tag1 As ULong, <Out> ByRef tag2 As ULong)
        <PreserveSig>
        Overloads Sub PushLayer(ByRef layerParameters As D2D1_LAYER_PARAMETERS, Optional layer As ID2D1Layer = Nothing)
        <PreserveSig>
        Overloads Sub PopLayer()
        <PreserveSig>
        Overloads Sub Flush(<Out> ByRef tag1 As ULong, <Out> ByRef tag2 As ULong)
        <PreserveSig>
        Overloads Sub SaveDrawingState(
        <[In], Out> drawingStateBlock As ID2D1DrawingStateBlock)
        <PreserveSig>
        Overloads Sub RestoreDrawingState(drawingStateBlock As ID2D1DrawingStateBlock)
        <PreserveSig>
        Overloads Sub PushAxisAlignedClip(clipRect As D2D1_RECT_F, antialiasMode As D2D1_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub PopAxisAlignedClip()
        <PreserveSig>
        Overloads Sub Clear(clearColor As D2D1_COLOR_F)
        <PreserveSig>
        Overloads Sub BeginDraw()
        <PreserveSig>
        Overloads Function EndDraw(<Out> ByRef tag1 As ULong, <Out> ByRef tag2 As ULong) As HRESULT
        <PreserveSig>
        Overloads Sub GetPixelFormat(<Out> ByRef pixelFormat As D2D1_PIXEL_FORMAT)
        <PreserveSig>
        Overloads Sub SetDpi(dpiX As Single, dpiY As Single)
        <PreserveSig>
        Overloads Sub GetDpi(<Out> ByRef dpiX As Single, <Out> ByRef dpiY As Single)
        <PreserveSig>
        Overloads Sub GetSize(<Out> ByRef size As D2D1_SIZE_F)
        <PreserveSig>
        Overloads Sub GetPixelSize(<Out> ByRef size As D2D1_SIZE_U)
        <PreserveSig>
        Overloads Function GetMaximumBitmapSize() As UInteger
        <PreserveSig>
        Overloads Function IsSupported(renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES) As Boolean
#End Region
        <PreserveSig>
        Overloads Function CreateBitmap(size As D2D1_SIZE_U, sourceData As IntPtr, pitch As UInteger, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES1, <Out> ByRef bitmap As ID2D1Bitmap1) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapFromWicBitmap(wicBitmapSource As IWICBitmapSource, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES1, <Out> ByRef bitmap As ID2D1Bitmap1) As HRESULT
        <PreserveSig>
        Function CreateColorContext(space As D2D1_COLOR_SPACE, profile As IntPtr, profileSize As UInteger, <Out> ByRef colorContext As ID2D1ColorContext) As HRESULT
        <PreserveSig>
        Function CreateColorContextFromFilename(filename As String, <Out> ByRef colorContext As ID2D1ColorContext) As HRESULT
        <PreserveSig>
        Function CreateColorContextFromWicColorContext(wicColorContext As IWICColorContext, <Out> ByRef colorContext As ID2D1ColorContext) As HRESULT
        <PreserveSig>
        Function CreateBitmapFromDxgiSurface(surface As IDXGISurface, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES1, <Out> ByRef bitmap As ID2D1Bitmap1) As HRESULT
        <PreserveSig>
        Function CreateEffect(ByRef effectId As Guid, <Out> ByRef effect As ID2D1Effect) As HRESULT
        <PreserveSig>
        Overloads Function CreateGradientStopCollection(
                        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> straightAlphaGradientStops As D2D1_GRADIENT_STOP(), straightAlphaGradientStopsCount As UInteger, preInterpolationSpace As D2D1_COLOR_SPACE, postInterpolationSpace As D2D1_COLOR_SPACE, bufferPrecision As D2D1_BUFFER_PRECISION, extendMode As D2D1_EXTEND_MODE, colorInterpolationMode As D2D1_COLOR_INTERPOLATION_MODE, <Out> ByRef gradientStopCollection1 As ID2D1GradientStopCollection1) As HRESULT
        <PreserveSig>
        Function CreateImageBrush(image As ID2D1Image, ByRef imageBrushProperties As D2D1_IMAGE_BRUSH_PROPERTIES, brushProperties As D2D1_BRUSH_PROPERTIES, <Out> ByRef imageBrush As ID2D1ImageBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapBrush(bitmap As ID2D1Bitmap, ByRef bitmapBrushProperties As D2D1_BITMAP_BRUSH_PROPERTIES1, brushProperties As D2D1_BRUSH_PROPERTIES, <Out> ByRef bitmapBrush As ID2D1BitmapBrush1) As HRESULT
        <PreserveSig>
        Function CreateCommandList(<Out> ByRef commandList As ID2D1CommandList) As HRESULT
        <PreserveSig>
        Function IsDxgiFormatSupported(format As DXGI_FORMAT) As Boolean
        <PreserveSig>
        Function IsBufferPrecisionSupported(bufferPrecision As D2D1_BUFFER_PRECISION) As Boolean
        <PreserveSig>
        Function GetImageLocalBounds(image As ID2D1Image, <Out> ByRef localBounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Function GetImageWorldBounds(image As ID2D1Image, <Out> ByRef worldBounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Function GetGlyphRunWorldBounds(ByRef baselineOrigin As D2D1_POINT_2F, glyphRun As DWRITE_GLYPH_RUN, measuringMode As DWRITE_MEASURING_MODE, <Out> ByRef bounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Sub GetDevice(<Out> ByRef device As ID2D1Device)
        <PreserveSig>
        Sub SetTarget(image As ID2D1Image)
        <PreserveSig>
        Sub GetTarget(<Out> ByRef image As ID2D1Image)
        <PreserveSig>
        Sub SetRenderingControls(renderingControls As D2D1_RENDERING_CONTROLS)
        <PreserveSig>
        Sub GetRenderingControls(<Out> ByRef renderingControls As D2D1_RENDERING_CONTROLS)
        <PreserveSig>
        Sub SetPrimitiveBlend(primitiveBlend As D2D1_PRIMITIVE_BLEND)
        <PreserveSig>
        Sub GetPrimitiveBlend(<Out> ByRef primitiveBlend As D2D1_PRIMITIVE_BLEND)
        <PreserveSig>
        Sub SetUnitMode(unitMode As D2D1_UNIT_MODE)
        <PreserveSig>
        Sub GetUnitMode(<Out> ByRef unitMode As D2D1_UNIT_MODE)
        <PreserveSig>
        Overloads Sub DrawGlyphRun(baselineOrigin As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, glyphRunDescription As IntPtr, foregroundBrush As ID2D1Brush, measuringMode As DWRITE_MEASURING_MODE)
        <PreserveSig>
        Sub DrawImage(image As ID2D1Image, ByRef targetOffset As D2D1_POINT_2F, ByRef imageRectangle As D2D1_RECT_F, interpolationMode As D2D1_INTERPOLATION_MODE, compositeMode As D2D1_COMPOSITE_MODE)
        <PreserveSig>
        Sub DrawGdiMetafile(gdiMetafile As ID2D1GdiMetafile, ByRef targetOffset As D2D1_POINT_2F)
        <PreserveSig>
        Overloads Sub DrawBitmap(bitmap As ID2D1Bitmap, ByRef destinationRectangle As D2D1_RECT_F, opacity As Single, interpolationMode As D2D1_INTERPOLATION_MODE, ByRef sourceRectangle As D2D1_RECT_F, Optional perspectiveTransform As D2D1_MATRIX_4X4_F = Nothing)
        <PreserveSig>
        Overloads Sub PushLayer(ByRef layerParameters As D2D1_LAYER_PARAMETERS1, layer As ID2D1Layer)
        Function InvalidateEffectInputRectangle(effect As ID2D1Effect, input As UInteger, inputRectangle As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Function GetEffectInvalidRectangleCount(effect As ID2D1Effect, <Out> ByRef rectangleCount As UInteger) As HRESULT
        <PreserveSig>
        Function GetEffectInvalidRectangles(effect As ID2D1Effect, <Out> ByRef rectangles As IntPtr, rectanglesCount As UInteger) As HRESULT
        <PreserveSig>
        Function GetEffectRequiredInputRectangles(renderEffect As ID2D1Effect, renderImageRectangle As D2D1_RECT_F, inputDescriptions As D2D1_EFFECT_INPUT_DESCRIPTION, <Out> ByRef requiredInputRects As IntPtr, inputCount As UInteger) As HRESULT
        <PreserveSig>
        Overloads Sub FillOpacityMask(opacityMask As ID2D1Bitmap, brush As ID2D1Brush, ByRef destinationRectangle As D2D1_RECT_F, sourceRectangle As D2D1_RECT_F)
    End Interface

    <ComImport>
    <Guid("d37f57e4-6908-459f-a199-e72f24f79987")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1DeviceContext1
        Inherits ID2D1DeviceContext
#Region "<ID2D1DeviceContext>"
#Region "<ID2D1RenderTarget>"

#Region "<ID2D1Resource>"

        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)

#End Region
        <PreserveSig>
        Overloads Function CreateBitmap(size As D2D1_SIZE_U, srcData As IntPtr, pitch As UInteger, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES, <Out> ByRef bitmap As ID2D1Bitmap) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapFromWicBitmap(wicBitmapSource As IWICBitmapSource, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES, <Out> ByRef bitmap As ID2D1Bitmap) As HRESULT
        <PreserveSig>
        Overloads Function CreateSharedBitmap(ByRef riid As Guid,
        <[In], Out> data As IntPtr, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES, <Out> ByRef bitmap As ID2D1Bitmap) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapBrush(bitmap As ID2D1Bitmap, ByRef bitmapBrushProperties As D2D1_BITMAP_BRUSH_PROPERTIES, brushProperties As D2D1_BRUSH_PROPERTIES, <Out> ByRef bitmapBrush As ID2D1BitmapBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateSolidColorBrush(color As D2D1_COLOR_F, brushProperties As D2D1_BRUSH_PROPERTIES, <Out> ByRef solidColorBrush As ID2D1SolidColorBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateGradientStopCollection(
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> gradientStops As D2D1_GRADIENT_STOP(), gradientStopsCount As UInteger, colorInterpolationGamma As D2D1_GAMMA, extendMode As D2D1_EXTEND_MODE, <Out> ByRef gradientStopCollection As ID2D1GradientStopCollection) As HRESULT
        <PreserveSig>
        Overloads Function CreateLinearGradientBrush(ByRef linearGradientBrushProperties As D2D1_LINEAR_GRADIENT_BRUSH_PROPERTIES, brushProperties As IntPtr, gradientStopCollection As ID2D1GradientStopCollection, <Out> ByRef linearGradientBrush As ID2D1LinearGradientBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateRadialGradientBrush(ByRef radialGradientBrushProperties As D2D1_RADIAL_GRADIENT_BRUSH_PROPERTIES, brushProperties As IntPtr, gradientStopCollection As ID2D1GradientStopCollection, <Out> ByRef radialGradientBrush As ID2D1RadialGradientBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateCompatibleRenderTarget(ByRef desiredSize As D2D1_SIZE_F, ByRef desiredPixelSize As D2D1_SIZE_U, ByRef desiredFormat As D2D1_PIXEL_FORMAT, options As D2D1_COMPATIBLE_RENDER_TARGET_OPTIONS, <Out> ByRef bitmapRenderTarget As ID2D1BitmapRenderTarget) As HRESULT
        <PreserveSig>
        Overloads Function CreateLayer(ByRef size As D2D1_SIZE_F, <Out> ByRef layer As ID2D1Layer) As HRESULT
        <PreserveSig>
        Overloads Sub CreateMesh(<Out> ByRef mesh As ID2D1Mesh)
        <PreserveSig>
        Overloads Sub DrawLine(point0 As D2D1_POINT_2F, point1 As D2D1_POINT_2F, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub DrawRectangle(ByRef rect As D2D1_RECT_F, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillRectangle(ByRef rect As D2D1_RECT_F, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub DrawRoundedRectangle(ByRef roundedRect As D2D1_ROUNDED_RECT, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillRoundedRectangle(ByRef roundedRect As D2D1_ROUNDED_RECT, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub DrawEllipse(ByRef ellipse As D2D1_ELLIPSE, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillEllipse(ByRef ellipse As D2D1_ELLIPSE, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub DrawGeometry(geometry As ID2D1Geometry, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillGeometry(geometry As ID2D1Geometry, brush As ID2D1Brush, Optional opacityBrush As ID2D1Brush = Nothing)
        <PreserveSig>
        Overloads Sub FillMesh(mesh As ID2D1Mesh, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub FillOpacityMask(opacityMask As ID2D1Bitmap, brush As ID2D1Brush, content As D2D1_OPACITY_MASK_CONTENT, destinationRectangle As D2D1_RECT_F, sourceRectangle As D2D1_RECT_F)
        <PreserveSig>
        Overloads Sub DrawBitmap(bitmap As ID2D1Bitmap, ByRef destinationRectangle As D2D1_RECT_F, opacity As Single, interpolationMode As D2D1_BITMAP_INTERPOLATION_MODE, ByRef sourceRectangle As D2D1_RECT_F)
        <PreserveSig>
        Overloads Sub DrawText(str As String, stringLength As UInteger, textFormat As IDWriteTextFormat, ByRef layoutRect As D2D1_RECT_F, defaultForegroundBrush As ID2D1Brush, options As D2D1_DRAW_TEXT_OPTIONS, measuringMode As DWRITE_MEASURING_MODE)
        <PreserveSig>
        Overloads Sub DrawTextLayout(origin As D2D1_POINT_2F, textLayout As IDWriteTextLayout, defaultForegroundBrush As ID2D1Brush, options As D2D1_DRAW_TEXT_OPTIONS)
        <PreserveSig>
        Overloads Sub DrawGlyphRun(baselineOrigin As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, foregroundBrush As ID2D1Brush, measuringMode As DWRITE_MEASURING_MODE)
        <PreserveSig>
        Overloads Sub SetTransform(transform As D2D1_MATRIX_3X2_F)
        <PreserveSig>
        Overloads Sub GetTransform(<Out> ByRef transform As D2D1_MATRIX_3X2_F_STRUCT)
        <PreserveSig>
        Overloads Sub SetAntialiasMode(antialiasMode As D2D1_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub GetAntialiasMode(<Out> ByRef antialiasMode As D2D1_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub SetTextAntialiasMode(textAntialiasMode As D2D1_TEXT_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub GetTextAntialiasMode(<Out> ByRef textAntialiasMode As D2D1_TEXT_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub SetTextRenderingParams(Optional textRenderingParams As IDWriteRenderingParams = Nothing)
        <PreserveSig>
        Overloads Sub GetTextRenderingParams(<Out> ByRef textRenderingParams As IDWriteRenderingParams)
        <PreserveSig>
        Overloads Sub SetTags(tag1 As ULong, tag2 As ULong)
        <PreserveSig>
        Overloads Sub GetTags(<Out> ByRef tag1 As ULong, <Out> ByRef tag2 As ULong)
        <PreserveSig>
        Overloads Sub PushLayer(ByRef layerParameters As D2D1_LAYER_PARAMETERS, layer As ID2D1Layer)
        <PreserveSig>
        Overloads Sub PopLayer()
        <PreserveSig>
        Overloads Sub Flush(<Out> ByRef tag1 As ULong, <Out> ByRef tag2 As ULong)
        <PreserveSig>
        Overloads Sub SaveDrawingState(
        <[In], Out> drawingStateBlock As ID2D1DrawingStateBlock)
        <PreserveSig>
        Overloads Sub RestoreDrawingState(drawingStateBlock As ID2D1DrawingStateBlock)
        <PreserveSig>
        Overloads Sub PushAxisAlignedClip(clipRect As D2D1_RECT_F, antialiasMode As D2D1_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub PopAxisAlignedClip()
        <PreserveSig>
        Overloads Sub Clear(clearColor As D2D1_COLOR_F)
        <PreserveSig>
        Overloads Sub BeginDraw()
        <PreserveSig>
        Overloads Function EndDraw(<Out> ByRef tag1 As ULong, <Out> ByRef tag2 As ULong) As HRESULT
        <PreserveSig>
        Overloads Sub GetPixelFormat(<Out> ByRef pixelFormat As D2D1_PIXEL_FORMAT)
        <PreserveSig>
        Overloads Sub SetDpi(dpiX As Single, dpiY As Single)
        <PreserveSig>
        Overloads Sub GetDpi(<Out> ByRef dpiX As Single, <Out> ByRef dpiY As Single)
        <PreserveSig>
        Overloads Sub GetSize(<Out> ByRef size As D2D1_SIZE_F)
        <PreserveSig>
        Overloads Sub GetPixelSize(<Out> ByRef size As D2D1_SIZE_U)
        <PreserveSig>
        Overloads Function GetMaximumBitmapSize() As UInteger
        <PreserveSig>
        Overloads Function IsSupported(renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES) As Boolean
#End Region

        <PreserveSig>
        Overloads Function CreateBitmap(size As D2D1_SIZE_U, sourceData As IntPtr, pitch As UInteger, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES1, <Out> ByRef bitmap As ID2D1Bitmap1) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapFromWicBitmap(wicBitmapSource As IWICBitmapSource, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES1, <Out> ByRef bitmap As ID2D1Bitmap1) As HRESULT
        <PreserveSig>
        Overloads Function CreateColorContext(space As D2D1_COLOR_SPACE, profile As IntPtr, profileSize As UInteger, <Out> ByRef colorContext As ID2D1ColorContext) As HRESULT
        <PreserveSig>
        Overloads Function CreateColorContextFromFilename(filename As String, <Out> ByRef colorContext As ID2D1ColorContext) As HRESULT
        <PreserveSig>
        Overloads Function CreateColorContextFromWicColorContext(wicColorContext As IWICColorContext, <Out> ByRef colorContext As ID2D1ColorContext) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapFromDxgiSurface(surface As IDXGISurface, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES1, <Out> ByRef bitmap As ID2D1Bitmap1) As HRESULT
        <PreserveSig>
        Overloads Function CreateEffect(ByRef effectId As Guid, <Out> ByRef effect As ID2D1Effect) As HRESULT
        <PreserveSig>
        Overloads Function CreateGradientStopCollection(
                          <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> straightAlphaGradientStops As D2D1_GRADIENT_STOP(), straightAlphaGradientStopsCount As UInteger, preInterpolationSpace As D2D1_COLOR_SPACE, postInterpolationSpace As D2D1_COLOR_SPACE, bufferPrecision As D2D1_BUFFER_PRECISION, extendMode As D2D1_EXTEND_MODE, colorInterpolationMode As D2D1_COLOR_INTERPOLATION_MODE, <Out> ByRef gradientStopCollection1 As ID2D1GradientStopCollection1) As HRESULT
        <PreserveSig>
        Overloads Function CreateImageBrush(image As ID2D1Image, ByRef imageBrushProperties As D2D1_IMAGE_BRUSH_PROPERTIES, brushProperties As D2D1_BRUSH_PROPERTIES, <Out> ByRef imageBrush As ID2D1ImageBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapBrush(bitmap As ID2D1Bitmap, ByRef bitmapBrushProperties As D2D1_BITMAP_BRUSH_PROPERTIES1, brushProperties As D2D1_BRUSH_PROPERTIES, <Out> ByRef bitmapBrush As ID2D1BitmapBrush1) As HRESULT
        <PreserveSig>
        Overloads Function CreateCommandList(<Out> ByRef commandList As ID2D1CommandList) As HRESULT
        <PreserveSig>
        Overloads Function IsDxgiFormatSupported(format As DXGI_FORMAT) As Boolean
        <PreserveSig>
        Overloads Function IsBufferPrecisionSupported(bufferPrecision As D2D1_BUFFER_PRECISION) As Boolean
        <PreserveSig>
        Overloads Function GetImageLocalBounds(image As ID2D1Image, <Out> ByRef localBounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Function GetImageWorldBounds(image As ID2D1Image, <Out> ByRef worldBounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Function GetGlyphRunWorldBounds(ByRef baselineOrigin As D2D1_POINT_2F, glyphRun As DWRITE_GLYPH_RUN, measuringMode As DWRITE_MEASURING_MODE, <Out> ByRef bounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Sub GetDevice(<Out> ByRef device As ID2D1Device)
        <PreserveSig>
        Overloads Sub SetTarget(image As ID2D1Image)
        <PreserveSig>
        Overloads Sub GetTarget(<Out> ByRef image As ID2D1Image)
        <PreserveSig>
        Overloads Sub SetRenderingControls(renderingControls As D2D1_RENDERING_CONTROLS)
        <PreserveSig>
        Overloads Sub GetRenderingControls(<Out> ByRef renderingControls As D2D1_RENDERING_CONTROLS)
        <PreserveSig>
        Overloads Sub SetPrimitiveBlend(primitiveBlend As D2D1_PRIMITIVE_BLEND)
        <PreserveSig>
        Overloads Sub GetPrimitiveBlend(<Out> ByRef primitiveBlend As D2D1_PRIMITIVE_BLEND)
        <PreserveSig>
        Overloads Sub SetUnitMode(unitMode As D2D1_UNIT_MODE)
        <PreserveSig>
        Overloads Sub GetUnitMode(<Out> ByRef unitMode As D2D1_UNIT_MODE)
        <PreserveSig>
        Overloads Sub DrawGlyphRun(baselineOrigin As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, glyphRunDescription As IntPtr, foregroundBrush As ID2D1Brush, measuringMode As DWRITE_MEASURING_MODE)
        <PreserveSig>
        Overloads Sub DrawImage(image As ID2D1Image, ByRef targetOffset As D2D1_POINT_2F, ByRef imageRectangle As D2D1_RECT_F, interpolationMode As D2D1_INTERPOLATION_MODE, compositeMode As D2D1_COMPOSITE_MODE)
        <PreserveSig>
        Overloads Sub DrawGdiMetafile(gdiMetafile As ID2D1GdiMetafile, ByRef targetOffset As D2D1_POINT_2F)
        <PreserveSig>
        Overloads Sub DrawBitmap(bitmap As ID2D1Bitmap, ByRef destinationRectangle As D2D1_RECT_F, opacity As Single, interpolationMode As D2D1_INTERPOLATION_MODE, ByRef sourceRectangle As D2D1_RECT_F, perspectiveTransform As D2D1_MATRIX_4X4_F)
        <PreserveSig>
        Overloads Sub PushLayer(ByRef layerParameters As D2D1_LAYER_PARAMETERS1, layer As ID2D1Layer)
        <PreserveSig>
        Overloads Function InvalidateEffectInputRectangle(effect As ID2D1Effect, input As UInteger, inputRectangle As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Function GetEffectInvalidRectangleCount(effect As ID2D1Effect, <Out> ByRef rectangleCount As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function GetEffectInvalidRectangles(effect As ID2D1Effect, <Out> ByRef rectangles As IntPtr, rectanglesCount As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function GetEffectRequiredInputRectangles(renderEffect As ID2D1Effect, renderImageRectangle As D2D1_RECT_F, inputDescriptions As D2D1_EFFECT_INPUT_DESCRIPTION, <Out> ByRef requiredInputRects As IntPtr, inputCount As UInteger) As HRESULT
        <PreserveSig>
        Overloads Sub FillOpacityMask(opacityMask As ID2D1Bitmap, brush As ID2D1Brush, ByRef destinationRectangle As D2D1_RECT_F, sourceRectangle As D2D1_RECT_F)
#End Region

        <PreserveSig>
        Function CreateFilledGeometryRealization(geometry As ID2D1Geometry, flatteningTolerance As Single, <Out> ByRef geometryRealization As ID2D1GeometryRealization) As HRESULT
        <PreserveSig>
        Function CreateStrokedGeometryRealization(geometry As ID2D1Geometry, flatteningTolerance As Single, strokeWidth As Single, strokeStyle As ID2D1StrokeStyle, <Out> ByRef geometryRealization As ID2D1GeometryRealization) As HRESULT
        <PreserveSig>
        Sub DrawGeometryRealization(geometryRealization As ID2D1GeometryRealization, brush As ID2D1Brush)
    End Interface

    <ComImport>
    <Guid("a16907d7-bc02-4801-99e8-8cf7f485f774")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1GeometryRealization
        Inherits ID2D1Resource
#Region "<ID2D1Resource>"
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region
    End Interface

    <ComImport>
    <Guid("394ea6a3-0c34-4321-950b-6ca20f0be6c7")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1DeviceContext2
        Inherits ID2D1DeviceContext1
#Region "<ID2D1DeviceContext1>"
#Region "<ID2D1DeviceContext>"
#Region "<ID2D1RenderTarget>"

#Region "<ID2D1Resource>"

        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)

#End Region
        <PreserveSig>
        Overloads Function CreateBitmap(size As D2D1_SIZE_U, srcData As IntPtr, pitch As UInteger, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES, <Out> ByRef bitmap As ID2D1Bitmap) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapFromWicBitmap(wicBitmapSource As IWICBitmapSource, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES, <Out> ByRef bitmap As ID2D1Bitmap) As HRESULT
        <PreserveSig>
        Overloads Function CreateSharedBitmap(ByRef riid As Guid,
        <[In], Out> data As IntPtr, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES, <Out> ByRef bitmap As ID2D1Bitmap) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapBrush(bitmap As ID2D1Bitmap, ByRef bitmapBrushProperties As D2D1_BITMAP_BRUSH_PROPERTIES, brushProperties As D2D1_BRUSH_PROPERTIES, <Out> ByRef bitmapBrush As ID2D1BitmapBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateSolidColorBrush(color As D2D1_COLOR_F, brushProperties As D2D1_BRUSH_PROPERTIES, <Out> ByRef solidColorBrush As ID2D1SolidColorBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateGradientStopCollection(
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> gradientStops As D2D1_GRADIENT_STOP(), gradientStopsCount As UInteger, colorInterpolationGamma As D2D1_GAMMA, extendMode As D2D1_EXTEND_MODE, <Out> ByRef gradientStopCollection As ID2D1GradientStopCollection) As HRESULT
        <PreserveSig>
        Overloads Function CreateLinearGradientBrush(ByRef linearGradientBrushProperties As D2D1_LINEAR_GRADIENT_BRUSH_PROPERTIES, brushProperties As IntPtr, gradientStopCollection As ID2D1GradientStopCollection, <Out> ByRef linearGradientBrush As ID2D1LinearGradientBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateRadialGradientBrush(ByRef radialGradientBrushProperties As D2D1_RADIAL_GRADIENT_BRUSH_PROPERTIES, brushProperties As IntPtr, gradientStopCollection As ID2D1GradientStopCollection, <Out> ByRef radialGradientBrush As ID2D1RadialGradientBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateCompatibleRenderTarget(ByRef desiredSize As D2D1_SIZE_F, ByRef desiredPixelSize As D2D1_SIZE_U, ByRef desiredFormat As D2D1_PIXEL_FORMAT, options As D2D1_COMPATIBLE_RENDER_TARGET_OPTIONS, <Out> ByRef bitmapRenderTarget As ID2D1BitmapRenderTarget) As HRESULT
        <PreserveSig>
        Overloads Function CreateLayer(ByRef size As D2D1_SIZE_F, <Out> ByRef layer As ID2D1Layer) As HRESULT
        <PreserveSig>
        Overloads Sub CreateMesh(<Out> ByRef mesh As ID2D1Mesh)
        <PreserveSig>
        Overloads Sub DrawLine(point0 As D2D1_POINT_2F, point1 As D2D1_POINT_2F, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub DrawRectangle(ByRef rect As D2D1_RECT_F, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillRectangle(ByRef rect As D2D1_RECT_F, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub DrawRoundedRectangle(ByRef roundedRect As D2D1_ROUNDED_RECT, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillRoundedRectangle(ByRef roundedRect As D2D1_ROUNDED_RECT, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub DrawEllipse(ByRef ellipse As D2D1_ELLIPSE, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillEllipse(ByRef ellipse As D2D1_ELLIPSE, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub DrawGeometry(geometry As ID2D1Geometry, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillGeometry(geometry As ID2D1Geometry, brush As ID2D1Brush, Optional opacityBrush As ID2D1Brush = Nothing)
        <PreserveSig>
        Overloads Sub FillMesh(mesh As ID2D1Mesh, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub FillOpacityMask(opacityMask As ID2D1Bitmap, brush As ID2D1Brush, content As D2D1_OPACITY_MASK_CONTENT, destinationRectangle As D2D1_RECT_F, sourceRectangle As D2D1_RECT_F)
        <PreserveSig>
        Overloads Sub DrawBitmap(bitmap As ID2D1Bitmap, ByRef destinationRectangle As D2D1_RECT_F, opacity As Single, interpolationMode As D2D1_BITMAP_INTERPOLATION_MODE, ByRef sourceRectangle As D2D1_RECT_F)
        <PreserveSig>
        Overloads Sub DrawText(str As String, stringLength As UInteger, textFormat As IDWriteTextFormat, ByRef layoutRect As D2D1_RECT_F, defaultForegroundBrush As ID2D1Brush, options As D2D1_DRAW_TEXT_OPTIONS, measuringMode As DWRITE_MEASURING_MODE)
        <PreserveSig>
        Overloads Sub DrawTextLayout(origin As D2D1_POINT_2F, textLayout As IDWriteTextLayout, defaultForegroundBrush As ID2D1Brush, options As D2D1_DRAW_TEXT_OPTIONS)
        <PreserveSig>
        Overloads Sub DrawGlyphRun(baselineOrigin As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, foregroundBrush As ID2D1Brush, measuringMode As DWRITE_MEASURING_MODE)
        <PreserveSig>
        Overloads Sub SetTransform(transform As D2D1_MATRIX_3X2_F)
        <PreserveSig>
        Overloads Sub GetTransform(<Out> ByRef transform As D2D1_MATRIX_3X2_F_STRUCT)
        <PreserveSig>
        Overloads Sub SetAntialiasMode(antialiasMode As D2D1_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub GetAntialiasMode(<Out> ByRef antialiasMode As D2D1_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub SetTextAntialiasMode(textAntialiasMode As D2D1_TEXT_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub GetTextAntialiasMode(<Out> ByRef textAntialiasMode As D2D1_TEXT_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub SetTextRenderingParams(Optional textRenderingParams As IDWriteRenderingParams = Nothing)
        <PreserveSig>
        Overloads Sub GetTextRenderingParams(<Out> ByRef textRenderingParams As IDWriteRenderingParams)
        <PreserveSig>
        Overloads Sub SetTags(tag1 As ULong, tag2 As ULong)
        <PreserveSig>
        Overloads Sub GetTags(<Out> ByRef tag1 As ULong, <Out> ByRef tag2 As ULong)
        <PreserveSig>
        Overloads Sub PushLayer(ByRef layerParameters As D2D1_LAYER_PARAMETERS, layer As ID2D1Layer)
        <PreserveSig>
        Overloads Sub PopLayer()
        <PreserveSig>
        Overloads Sub Flush(<Out> ByRef tag1 As ULong, <Out> ByRef tag2 As ULong)
        <PreserveSig>
        Overloads Sub SaveDrawingState(
        <[In], Out> drawingStateBlock As ID2D1DrawingStateBlock)
        <PreserveSig>
        Overloads Sub RestoreDrawingState(drawingStateBlock As ID2D1DrawingStateBlock)
        <PreserveSig>
        Overloads Sub PushAxisAlignedClip(clipRect As D2D1_RECT_F, antialiasMode As D2D1_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub PopAxisAlignedClip()
        <PreserveSig>
        Overloads Sub Clear(clearColor As D2D1_COLOR_F)
        <PreserveSig>
        Overloads Sub BeginDraw()
        <PreserveSig>
        Overloads Function EndDraw(<Out> ByRef tag1 As ULong, <Out> ByRef tag2 As ULong) As HRESULT
        <PreserveSig>
        Overloads Sub GetPixelFormat(<Out> ByRef pixelFormat As D2D1_PIXEL_FORMAT)
        <PreserveSig>
        Overloads Sub SetDpi(dpiX As Single, dpiY As Single)
        <PreserveSig>
        Overloads Sub GetDpi(<Out> ByRef dpiX As Single, <Out> ByRef dpiY As Single)
        <PreserveSig>
        Overloads Sub GetSize(<Out> ByRef size As D2D1_SIZE_F)
        <PreserveSig>
        Overloads Sub GetPixelSize(<Out> ByRef size As D2D1_SIZE_U)
        <PreserveSig>
        Overloads Function GetMaximumBitmapSize() As UInteger
        <PreserveSig>
        Overloads Function IsSupported(renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES) As Boolean
#End Region

        <PreserveSig>
        Overloads Function CreateBitmap(size As D2D1_SIZE_U, sourceData As IntPtr, pitch As UInteger, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES1, <Out> ByRef bitmap As ID2D1Bitmap1) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapFromWicBitmap(wicBitmapSource As IWICBitmapSource, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES1, <Out> ByRef bitmap As ID2D1Bitmap1) As HRESULT
        <PreserveSig>
        Overloads Function CreateColorContext(space As D2D1_COLOR_SPACE, profile As IntPtr, profileSize As UInteger, <Out> ByRef colorContext As ID2D1ColorContext) As HRESULT
        <PreserveSig>
        Overloads Function CreateColorContextFromFilename(filename As String, <Out> ByRef colorContext As ID2D1ColorContext) As HRESULT
        <PreserveSig>
        Overloads Function CreateColorContextFromWicColorContext(wicColorContext As IWICColorContext, <Out> ByRef colorContext As ID2D1ColorContext) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapFromDxgiSurface(surface As IDXGISurface, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES1, <Out> ByRef bitmap As ID2D1Bitmap1) As HRESULT
        <PreserveSig>
        Overloads Function CreateEffect(ByRef effectId As Guid, <Out> ByRef effect As ID2D1Effect) As HRESULT
        <PreserveSig>
        Overloads Function CreateGradientStopCollection(
                    <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> straightAlphaGradientStops As D2D1_GRADIENT_STOP(), straightAlphaGradientStopsCount As UInteger, preInterpolationSpace As D2D1_COLOR_SPACE, postInterpolationSpace As D2D1_COLOR_SPACE, bufferPrecision As D2D1_BUFFER_PRECISION, extendMode As D2D1_EXTEND_MODE, colorInterpolationMode As D2D1_COLOR_INTERPOLATION_MODE, <Out> ByRef gradientStopCollection1 As ID2D1GradientStopCollection1) As HRESULT
        <PreserveSig>
        Overloads Function CreateImageBrush(image As ID2D1Image, ByRef imageBrushProperties As D2D1_IMAGE_BRUSH_PROPERTIES, brushProperties As D2D1_BRUSH_PROPERTIES, <Out> ByRef imageBrush As ID2D1ImageBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapBrush(bitmap As ID2D1Bitmap, ByRef bitmapBrushProperties As D2D1_BITMAP_BRUSH_PROPERTIES1, brushProperties As D2D1_BRUSH_PROPERTIES, <Out> ByRef bitmapBrush As ID2D1BitmapBrush1) As HRESULT
        <PreserveSig>
        Overloads Function CreateCommandList(<Out> ByRef commandList As ID2D1CommandList) As HRESULT
        <PreserveSig>
        Overloads Function IsDxgiFormatSupported(format As DXGI_FORMAT) As Boolean
        <PreserveSig>
        Overloads Function IsBufferPrecisionSupported(bufferPrecision As D2D1_BUFFER_PRECISION) As Boolean
        <PreserveSig>
        Overloads Function GetImageLocalBounds(image As ID2D1Image, <Out> ByRef localBounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Function GetImageWorldBounds(image As ID2D1Image, <Out> ByRef worldBounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Function GetGlyphRunWorldBounds(ByRef baselineOrigin As D2D1_POINT_2F, glyphRun As DWRITE_GLYPH_RUN, measuringMode As DWRITE_MEASURING_MODE, <Out> ByRef bounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Sub GetDevice(<Out> ByRef device As ID2D1Device)
        <PreserveSig>
        Overloads Sub SetTarget(image As ID2D1Image)
        <PreserveSig>
        Overloads Sub GetTarget(<Out> ByRef image As ID2D1Image)
        <PreserveSig>
        Overloads Sub SetRenderingControls(renderingControls As D2D1_RENDERING_CONTROLS)
        <PreserveSig>
        Overloads Sub GetRenderingControls(<Out> ByRef renderingControls As D2D1_RENDERING_CONTROLS)
        <PreserveSig>
        Overloads Sub SetPrimitiveBlend(primitiveBlend As D2D1_PRIMITIVE_BLEND)
        <PreserveSig>
        Overloads Sub GetPrimitiveBlend(<Out> ByRef primitiveBlend As D2D1_PRIMITIVE_BLEND)
        <PreserveSig>
        Overloads Sub SetUnitMode(unitMode As D2D1_UNIT_MODE)
        <PreserveSig>
        Overloads Sub GetUnitMode(<Out> ByRef unitMode As D2D1_UNIT_MODE)
        <PreserveSig>
        Overloads Sub DrawGlyphRun(baselineOrigin As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, glyphRunDescription As IntPtr, foregroundBrush As ID2D1Brush, measuringMode As DWRITE_MEASURING_MODE)
        <PreserveSig>
        Overloads Sub DrawImage(image As ID2D1Image, ByRef targetOffset As D2D1_POINT_2F, ByRef imageRectangle As D2D1_RECT_F, interpolationMode As D2D1_INTERPOLATION_MODE, compositeMode As D2D1_COMPOSITE_MODE)
        <PreserveSig>
        Overloads Sub DrawGdiMetafile(gdiMetafile As ID2D1GdiMetafile, ByRef targetOffset As D2D1_POINT_2F)
        <PreserveSig>
        Overloads Sub DrawBitmap(bitmap As ID2D1Bitmap, ByRef destinationRectangle As D2D1_RECT_F, opacity As Single, interpolationMode As D2D1_INTERPOLATION_MODE, ByRef sourceRectangle As D2D1_RECT_F, perspectiveTransform As D2D1_MATRIX_4X4_F)
        <PreserveSig>
        Overloads Sub PushLayer(ByRef layerParameters As D2D1_LAYER_PARAMETERS1, layer As ID2D1Layer)
        <PreserveSig>
        Overloads Function InvalidateEffectInputRectangle(effect As ID2D1Effect, input As UInteger, inputRectangle As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Function GetEffectInvalidRectangleCount(effect As ID2D1Effect, <Out> ByRef rectangleCount As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function GetEffectInvalidRectangles(effect As ID2D1Effect, <Out> ByRef rectangles As IntPtr, rectanglesCount As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function GetEffectRequiredInputRectangles(renderEffect As ID2D1Effect, renderImageRectangle As D2D1_RECT_F, inputDescriptions As D2D1_EFFECT_INPUT_DESCRIPTION, <Out> ByRef requiredInputRects As IntPtr, inputCount As UInteger) As HRESULT
        <PreserveSig>
        Overloads Sub FillOpacityMask(opacityMask As ID2D1Bitmap, brush As ID2D1Brush, ByRef destinationRectangle As D2D1_RECT_F, sourceRectangle As D2D1_RECT_F)
#End Region

        <PreserveSig>
        Overloads Function CreateFilledGeometryRealization(geometry As ID2D1Geometry, flatteningTolerance As Single, <Out> ByRef geometryRealization As ID2D1GeometryRealization) As HRESULT
        <PreserveSig>
        Overloads Function CreateStrokedGeometryRealization(geometry As ID2D1Geometry, flatteningTolerance As Single, strokeWidth As Single, strokeStyle As ID2D1StrokeStyle, <Out> ByRef geometryRealization As ID2D1GeometryRealization) As HRESULT
        <PreserveSig>
        Overloads Sub DrawGeometryRealization(geometryRealization As ID2D1GeometryRealization, brush As ID2D1Brush)
#End Region

        <PreserveSig>
        Function CreateInk(startPoint As D2D1_INK_POINT, <Out> ByRef ink As ID2D1Ink) As HRESULT
        <PreserveSig>
        Function CreateInkStyle(inkStyleProperties As D2D1_INK_STYLE_PROPERTIES, <Out> ByRef inkStyle As ID2D1InkStyle) As HRESULT
        <PreserveSig>
        Function CreateGradientMesh(patches As D2D1_GRADIENT_MESH_PATCH, patchesCount As UInteger, <Out> ByRef gradientMesh As ID2D1GradientMesh) As HRESULT
        <PreserveSig>
        Function CreateImageSourceFromWic(wicBitmapSource As IWICBitmapSource, loadingOptions As D2D1_IMAGE_SOURCE_LOADING_OPTIONS, alphaMode As D2D1_ALPHA_MODE, <Out> ByRef imageSource As ID2D1ImageSourceFromWic) As HRESULT
        <PreserveSig>
        Function CreateLookupTable3D(precision As D2D1_BUFFER_PRECISION, extents As UInteger, data As IntPtr, dataCount As UInteger, strides As UInteger, <Out> ByRef lookupTable As ID2D1LookupTable3D) As HRESULT
        <PreserveSig>
        Function CreateImageSourceFromDxgi(surfaces As IntPtr, surfaceCount As UInteger, colorSpace As DXGI_COLOR_SPACE_TYPE, options As D2D1_IMAGE_SOURCE_FROM_DXGI_OPTIONS, <Out> ByRef imageSource As ID2D1ImageSource) As HRESULT
        <PreserveSig>
        Function GetGradientMeshWorldBounds(gradientMesh As ID2D1GradientMesh, <Out> ByRef pBounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Sub DrawInk(ink As ID2D1Ink, brush As ID2D1Brush, inkStyle As ID2D1InkStyle)
        <PreserveSig>
        Sub DrawGradientMesh(gradientMesh As ID2D1GradientMesh)
        <PreserveSig>
        Overloads Sub DrawGdiMetafile(gdiMetafile As ID2D1GdiMetafile, ByRef destinationRectangle As D2D1_RECT_F, ByRef sourceRectangle As D2D1_RECT_F)
        <PreserveSig>
        Function CreateTransformedImageSource(imageSource As ID2D1ImageSource, properties As D2D1_TRANSFORMED_IMAGE_SOURCE_PROPERTIES, <Out> ByRef transformedImageSource As ID2D1TransformedImageSource) As HRESULT
    End Interface

    <ComImport>
    <Guid("235a7496-8351-414c-bcd4-6672ab2d8e00")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1DeviceContext3
        Inherits ID2D1DeviceContext2
#Region "<ID2D1DeviceContext2>"

#Region "<ID2D1DeviceContext1>"
#Region "<ID2D1DeviceContext>"
#Region "<ID2D1RenderTarget>"

#Region "<ID2D1Resource>"

        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)

#End Region
        <PreserveSig>
        Overloads Function CreateBitmap(size As D2D1_SIZE_U, srcData As IntPtr, pitch As UInteger, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES, <Out> ByRef bitmap As ID2D1Bitmap) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapFromWicBitmap(wicBitmapSource As IWICBitmapSource, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES, <Out> ByRef bitmap As ID2D1Bitmap) As HRESULT
        <PreserveSig>
        Overloads Function CreateSharedBitmap(ByRef riid As Guid,
        <[In], Out> data As IntPtr, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES, <Out> ByRef bitmap As ID2D1Bitmap) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapBrush(bitmap As ID2D1Bitmap, ByRef bitmapBrushProperties As D2D1_BITMAP_BRUSH_PROPERTIES, brushProperties As D2D1_BRUSH_PROPERTIES, <Out> ByRef bitmapBrush As ID2D1BitmapBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateSolidColorBrush(color As D2D1_COLOR_F, brushProperties As D2D1_BRUSH_PROPERTIES, <Out> ByRef solidColorBrush As ID2D1SolidColorBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateGradientStopCollection(
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> gradientStops As D2D1_GRADIENT_STOP(), gradientStopsCount As UInteger, colorInterpolationGamma As D2D1_GAMMA, extendMode As D2D1_EXTEND_MODE, <Out> ByRef gradientStopCollection As ID2D1GradientStopCollection) As HRESULT
        <PreserveSig>
        Overloads Function CreateLinearGradientBrush(ByRef linearGradientBrushProperties As D2D1_LINEAR_GRADIENT_BRUSH_PROPERTIES, brushProperties As IntPtr, gradientStopCollection As ID2D1GradientStopCollection, <Out> ByRef linearGradientBrush As ID2D1LinearGradientBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateRadialGradientBrush(ByRef radialGradientBrushProperties As D2D1_RADIAL_GRADIENT_BRUSH_PROPERTIES, brushProperties As IntPtr, gradientStopCollection As ID2D1GradientStopCollection, <Out> ByRef radialGradientBrush As ID2D1RadialGradientBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateCompatibleRenderTarget(ByRef desiredSize As D2D1_SIZE_F, ByRef desiredPixelSize As D2D1_SIZE_U, ByRef desiredFormat As D2D1_PIXEL_FORMAT, options As D2D1_COMPATIBLE_RENDER_TARGET_OPTIONS, <Out> ByRef bitmapRenderTarget As ID2D1BitmapRenderTarget) As HRESULT
        <PreserveSig>
        Overloads Function CreateLayer(ByRef size As D2D1_SIZE_F, <Out> ByRef layer As ID2D1Layer) As HRESULT
        <PreserveSig>
        Overloads Sub CreateMesh(<Out> ByRef mesh As ID2D1Mesh)
        <PreserveSig>
        Overloads Sub DrawLine(point0 As D2D1_POINT_2F, point1 As D2D1_POINT_2F, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub DrawRectangle(ByRef rect As D2D1_RECT_F, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillRectangle(ByRef rect As D2D1_RECT_F, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub DrawRoundedRectangle(ByRef roundedRect As D2D1_ROUNDED_RECT, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillRoundedRectangle(ByRef roundedRect As D2D1_ROUNDED_RECT, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub DrawEllipse(ByRef ellipse As D2D1_ELLIPSE, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillEllipse(ByRef ellipse As D2D1_ELLIPSE, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub DrawGeometry(geometry As ID2D1Geometry, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillGeometry(geometry As ID2D1Geometry, brush As ID2D1Brush, Optional opacityBrush As ID2D1Brush = Nothing)
        <PreserveSig>
        Overloads Sub FillMesh(mesh As ID2D1Mesh, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub FillOpacityMask(opacityMask As ID2D1Bitmap, brush As ID2D1Brush, content As D2D1_OPACITY_MASK_CONTENT, destinationRectangle As D2D1_RECT_F, sourceRectangle As D2D1_RECT_F)
        <PreserveSig>
        Overloads Sub DrawBitmap(bitmap As ID2D1Bitmap, ByRef destinationRectangle As D2D1_RECT_F, opacity As Single, interpolationMode As D2D1_BITMAP_INTERPOLATION_MODE, ByRef sourceRectangle As D2D1_RECT_F)
        <PreserveSig>
        Overloads Sub DrawText(str As String, stringLength As UInteger, textFormat As IDWriteTextFormat, ByRef layoutRect As D2D1_RECT_F, defaultForegroundBrush As ID2D1Brush, options As D2D1_DRAW_TEXT_OPTIONS, measuringMode As DWRITE_MEASURING_MODE)
        <PreserveSig>
        Overloads Sub DrawTextLayout(origin As D2D1_POINT_2F, textLayout As IDWriteTextLayout, defaultForegroundBrush As ID2D1Brush, options As D2D1_DRAW_TEXT_OPTIONS)
        <PreserveSig>
        Overloads Sub DrawGlyphRun(baselineOrigin As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, foregroundBrush As ID2D1Brush, measuringMode As DWRITE_MEASURING_MODE)
        <PreserveSig>
        Overloads Sub SetTransform(transform As D2D1_MATRIX_3X2_F)
        <PreserveSig>
        Overloads Sub GetTransform(<Out> ByRef transform As D2D1_MATRIX_3X2_F_STRUCT)
        <PreserveSig>
        Overloads Sub SetAntialiasMode(antialiasMode As D2D1_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub GetAntialiasMode(<Out> ByRef antialiasMode As D2D1_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub SetTextAntialiasMode(textAntialiasMode As D2D1_TEXT_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub GetTextAntialiasMode(<Out> ByRef textAntialiasMode As D2D1_TEXT_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub SetTextRenderingParams(Optional textRenderingParams As IDWriteRenderingParams = Nothing)
        <PreserveSig>
        Overloads Sub GetTextRenderingParams(<Out> ByRef textRenderingParams As IDWriteRenderingParams)
        <PreserveSig>
        Overloads Sub SetTags(tag1 As ULong, tag2 As ULong)
        <PreserveSig>
        Overloads Sub GetTags(<Out> ByRef tag1 As ULong, <Out> ByRef tag2 As ULong)
        <PreserveSig>
        Overloads Sub PushLayer(ByRef layerParameters As D2D1_LAYER_PARAMETERS, layer As ID2D1Layer)
        <PreserveSig>
        Overloads Sub PopLayer()
        <PreserveSig>
        Overloads Sub Flush(<Out> ByRef tag1 As ULong, <Out> ByRef tag2 As ULong)
        <PreserveSig>
        Overloads Sub SaveDrawingState(
        <[In], Out> drawingStateBlock As ID2D1DrawingStateBlock)
        <PreserveSig>
        Overloads Sub RestoreDrawingState(drawingStateBlock As ID2D1DrawingStateBlock)
        <PreserveSig>
        Overloads Sub PushAxisAlignedClip(clipRect As D2D1_RECT_F, antialiasMode As D2D1_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub PopAxisAlignedClip()
        <PreserveSig>
        Overloads Sub Clear(clearColor As D2D1_COLOR_F)
        <PreserveSig>
        Overloads Sub BeginDraw()
        <PreserveSig>
        Overloads Function EndDraw(<Out> ByRef tag1 As ULong, <Out> ByRef tag2 As ULong) As HRESULT
        <PreserveSig>
        Overloads Sub GetPixelFormat(<Out> ByRef pixelFormat As D2D1_PIXEL_FORMAT)
        <PreserveSig>
        Overloads Sub SetDpi(dpiX As Single, dpiY As Single)
        <PreserveSig>
        Overloads Sub GetDpi(<Out> ByRef dpiX As Single, <Out> ByRef dpiY As Single)
        <PreserveSig>
        Overloads Sub GetSize(<Out> ByRef size As D2D1_SIZE_F)
        <PreserveSig>
        Overloads Sub GetPixelSize(<Out> ByRef size As D2D1_SIZE_U)
        <PreserveSig>
        Overloads Function GetMaximumBitmapSize() As UInteger
        <PreserveSig>
        Overloads Function IsSupported(renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES) As Boolean
#End Region

        <PreserveSig>
        Overloads Function CreateBitmap(size As D2D1_SIZE_U, sourceData As IntPtr, pitch As UInteger, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES1, <Out> ByRef bitmap As ID2D1Bitmap1) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapFromWicBitmap(wicBitmapSource As IWICBitmapSource, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES1, <Out> ByRef bitmap As ID2D1Bitmap1) As HRESULT
        <PreserveSig>
        Overloads Function CreateColorContext(space As D2D1_COLOR_SPACE, profile As IntPtr, profileSize As UInteger, <Out> ByRef colorContext As ID2D1ColorContext) As HRESULT
        <PreserveSig>
        Overloads Function CreateColorContextFromFilename(filename As String, <Out> ByRef colorContext As ID2D1ColorContext) As HRESULT
        <PreserveSig>
        Overloads Function CreateColorContextFromWicColorContext(wicColorContext As IWICColorContext, <Out> ByRef colorContext As ID2D1ColorContext) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapFromDxgiSurface(surface As IDXGISurface, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES1, <Out> ByRef bitmap As ID2D1Bitmap1) As HRESULT
        <PreserveSig>
        Overloads Function CreateEffect(ByRef effectId As Guid, <Out> ByRef effect As ID2D1Effect) As HRESULT
        <PreserveSig>
        Overloads Function CreateGradientStopCollection(
                    <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> straightAlphaGradientStops As D2D1_GRADIENT_STOP(), straightAlphaGradientStopsCount As UInteger, preInterpolationSpace As D2D1_COLOR_SPACE, postInterpolationSpace As D2D1_COLOR_SPACE, bufferPrecision As D2D1_BUFFER_PRECISION, extendMode As D2D1_EXTEND_MODE, colorInterpolationMode As D2D1_COLOR_INTERPOLATION_MODE, <Out> ByRef gradientStopCollection1 As ID2D1GradientStopCollection1) As HRESULT
        <PreserveSig>
        Overloads Function CreateImageBrush(image As ID2D1Image, ByRef imageBrushProperties As D2D1_IMAGE_BRUSH_PROPERTIES, brushProperties As D2D1_BRUSH_PROPERTIES, <Out> ByRef imageBrush As ID2D1ImageBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapBrush(bitmap As ID2D1Bitmap, ByRef bitmapBrushProperties As D2D1_BITMAP_BRUSH_PROPERTIES1, brushProperties As D2D1_BRUSH_PROPERTIES, <Out> ByRef bitmapBrush As ID2D1BitmapBrush1) As HRESULT
        <PreserveSig>
        Overloads Function CreateCommandList(<Out> ByRef commandList As ID2D1CommandList) As HRESULT
        <PreserveSig>
        Overloads Function IsDxgiFormatSupported(format As DXGI_FORMAT) As Boolean
        <PreserveSig>
        Overloads Function IsBufferPrecisionSupported(bufferPrecision As D2D1_BUFFER_PRECISION) As Boolean
        <PreserveSig>
        Overloads Function GetImageLocalBounds(image As ID2D1Image, <Out> ByRef localBounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Function GetImageWorldBounds(image As ID2D1Image, <Out> ByRef worldBounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Function GetGlyphRunWorldBounds(ByRef baselineOrigin As D2D1_POINT_2F, glyphRun As DWRITE_GLYPH_RUN, measuringMode As DWRITE_MEASURING_MODE, <Out> ByRef bounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Sub GetDevice(<Out> ByRef device As ID2D1Device)
        <PreserveSig>
        Overloads Sub SetTarget(image As ID2D1Image)
        <PreserveSig>
        Overloads Sub GetTarget(<Out> ByRef image As ID2D1Image)
        <PreserveSig>
        Overloads Sub SetRenderingControls(renderingControls As D2D1_RENDERING_CONTROLS)
        <PreserveSig>
        Overloads Sub GetRenderingControls(<Out> ByRef renderingControls As D2D1_RENDERING_CONTROLS)
        <PreserveSig>
        Overloads Sub SetPrimitiveBlend(primitiveBlend As D2D1_PRIMITIVE_BLEND)
        <PreserveSig>
        Overloads Sub GetPrimitiveBlend(<Out> ByRef primitiveBlend As D2D1_PRIMITIVE_BLEND)
        <PreserveSig>
        Overloads Sub SetUnitMode(unitMode As D2D1_UNIT_MODE)
        <PreserveSig>
        Overloads Sub GetUnitMode(<Out> ByRef unitMode As D2D1_UNIT_MODE)
        <PreserveSig>
        Overloads Sub DrawGlyphRun(baselineOrigin As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, glyphRunDescription As IntPtr, foregroundBrush As ID2D1Brush, measuringMode As DWRITE_MEASURING_MODE)
        <PreserveSig>
        Overloads Sub DrawImage(image As ID2D1Image, ByRef targetOffset As D2D1_POINT_2F, ByRef imageRectangle As D2D1_RECT_F, interpolationMode As D2D1_INTERPOLATION_MODE, compositeMode As D2D1_COMPOSITE_MODE)
        <PreserveSig>
        Overloads Sub DrawGdiMetafile(gdiMetafile As ID2D1GdiMetafile, ByRef targetOffset As D2D1_POINT_2F)
        <PreserveSig>
        Overloads Sub DrawBitmap(bitmap As ID2D1Bitmap, ByRef destinationRectangle As D2D1_RECT_F, opacity As Single, interpolationMode As D2D1_INTERPOLATION_MODE, ByRef sourceRectangle As D2D1_RECT_F, perspectiveTransform As D2D1_MATRIX_4X4_F)
        <PreserveSig>
        Overloads Sub PushLayer(ByRef layerParameters As D2D1_LAYER_PARAMETERS1, layer As ID2D1Layer)
        <PreserveSig>
        Overloads Function InvalidateEffectInputRectangle(effect As ID2D1Effect, input As UInteger, inputRectangle As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Function GetEffectInvalidRectangleCount(effect As ID2D1Effect, <Out> ByRef rectangleCount As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function GetEffectInvalidRectangles(effect As ID2D1Effect, <Out> ByRef rectangles As IntPtr, rectanglesCount As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function GetEffectRequiredInputRectangles(renderEffect As ID2D1Effect, renderImageRectangle As D2D1_RECT_F, inputDescriptions As D2D1_EFFECT_INPUT_DESCRIPTION, <Out> ByRef requiredInputRects As IntPtr, inputCount As UInteger) As HRESULT
        <PreserveSig>
        Overloads Sub FillOpacityMask(opacityMask As ID2D1Bitmap, brush As ID2D1Brush, ByRef destinationRectangle As D2D1_RECT_F, sourceRectangle As D2D1_RECT_F)
#End Region
        <PreserveSig>
        Overloads Function CreateFilledGeometryRealization(geometry As ID2D1Geometry, flatteningTolerance As Single, <Out> ByRef geometryRealization As ID2D1GeometryRealization) As HRESULT
        <PreserveSig>
        Overloads Function CreateStrokedGeometryRealization(geometry As ID2D1Geometry, flatteningTolerance As Single, strokeWidth As Single, strokeStyle As ID2D1StrokeStyle, <Out> ByRef geometryRealization As ID2D1GeometryRealization) As HRESULT
        <PreserveSig>
        Overloads Sub DrawGeometryRealization(geometryRealization As ID2D1GeometryRealization, brush As ID2D1Brush)
#End Region

        <PreserveSig>
        Overloads Function CreateInk(startPoint As D2D1_INK_POINT, <Out> ByRef ink As ID2D1Ink) As HRESULT
        <PreserveSig>
        Overloads Function CreateInkStyle(inkStyleProperties As D2D1_INK_STYLE_PROPERTIES, <Out> ByRef inkStyle As ID2D1InkStyle) As HRESULT
        <PreserveSig>
        Overloads Function CreateGradientMesh(patches As D2D1_GRADIENT_MESH_PATCH, patchesCount As UInteger, <Out> ByRef gradientMesh As ID2D1GradientMesh) As HRESULT
        <PreserveSig>
        Overloads Function CreateImageSourceFromWic(wicBitmapSource As IWICBitmapSource, loadingOptions As D2D1_IMAGE_SOURCE_LOADING_OPTIONS, alphaMode As D2D1_ALPHA_MODE, <Out> ByRef imageSource As ID2D1ImageSourceFromWic) As HRESULT
        <PreserveSig>
        Overloads Function CreateLookupTable3D(precision As D2D1_BUFFER_PRECISION, extents As UInteger, data As IntPtr, dataCount As UInteger, strides As UInteger, <Out> ByRef lookupTable As ID2D1LookupTable3D) As HRESULT
        'new HRESULT CreateImageSourceFromDxgi(IDXGISurface** surfaces, uint surfaceCount, DXGI_COLOR_SPACE_TYPE colorSpace, D2D1_IMAGE_SOURCE_FROM_DXGI_OPTIONS options, out ID2D1ImageSource** imageSource);
        <PreserveSig>
        Overloads Function CreateImageSourceFromDxgi(surfaces As IntPtr, surfaceCount As UInteger, colorSpace As DXGI_COLOR_SPACE_TYPE, options As D2D1_IMAGE_SOURCE_FROM_DXGI_OPTIONS, <Out> ByRef imageSource As ID2D1ImageSource) As HRESULT
        <PreserveSig>
        Overloads Function GetGradientMeshWorldBounds(gradientMesh As ID2D1GradientMesh, <Out> ByRef pBounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Sub DrawInk(ink As ID2D1Ink, brush As ID2D1Brush, inkStyle As ID2D1InkStyle)
        <PreserveSig>
        Overloads Sub DrawGradientMesh(gradientMesh As ID2D1GradientMesh)
        <PreserveSig>
        Overloads Sub DrawGdiMetafile(gdiMetafile As ID2D1GdiMetafile, ByRef destinationRectangle As D2D1_RECT_F, ByRef sourceRectangle As D2D1_RECT_F)
        <PreserveSig>
        Overloads Function CreateTransformedImageSource(imageSource As ID2D1ImageSource, properties As D2D1_TRANSFORMED_IMAGE_SOURCE_PROPERTIES, <Out> ByRef transformedImageSource As ID2D1TransformedImageSource) As HRESULT

#End Region

        <PreserveSig>
        Function CreateSpriteBatch(<Out> ByRef spriteBatch As ID2D1SpriteBatch) As HRESULT
        <PreserveSig>
        Sub DrawSpriteBatch(spriteBatch As ID2D1SpriteBatch, startIndex As UInteger, spriteCount As UInteger, bitmap As ID2D1Bitmap, Optional interpolationMode As D2D1_BITMAP_INTERPOLATION_MODE = D2D1_BITMAP_INTERPOLATION_MODE.D2D1_BITMAP_INTERPOLATION_MODE_LINEAR, Optional spriteOptions As D2D1_SPRITE_OPTIONS = D2D1_SPRITE_OPTIONS.D2D1_SPRITE_OPTIONS_NONE)
    End Interface

    Public Enum D2D1_SPRITE_OPTIONS
        ''' <summary>
        ''' Use default sprite rendering behavior.
        ''' </summary>
        D2D1_SPRITE_OPTIONS_NONE = 0

        ''' <summary>
        ''' Bitmap interpolation will be clamped to the sprite's source rectangle.
        ''' </summary>
        D2D1_SPRITE_OPTIONS_CLAMP_TO_SOURCE_RECTANGLE = 1
        D2D1_SPRITE_OPTIONS_FORCE_DWORD = &HFFFFFFFF
    End Enum

    <ComImport>
    <Guid("4dc583bf-3a10-438a-8722-e9765224f1f1")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1SpriteBatch
        Inherits ID2D1Resource
#Region "<ID2D1Resource>"
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region

        'unsafe
        '       HRESULT AddSprites(uint spriteCount,  D2D1_RECT_F* destinationRectangles, void* sourceRectangles, void* colors, void* transforms,
        '    uint destinationRectanglesStride, uint sourceRectanglesStride, uint colorsStride, uint transformsStride);

        <PreserveSig>
        Function AddSprites(spriteCount As UInteger,
        <MarshalAs(UnmanagedType.LPArray)> destinationRectangles As D2D1_RECT_F(),
        <MarshalAs(UnmanagedType.LPArray)> sourceRectangles As D2D1_RECT_U(),
         <MarshalAs(UnmanagedType.LPArray)> colors As D2D1_COLOR_F_STRUCT(),
        <MarshalAs(UnmanagedType.LPArray)> transforms As D2D1_MATRIX_3X2_F_STRUCT(), destinationRectanglesStride As UInteger, sourceRectanglesStride As UInteger, colorsStride As UInteger, transformsStride As UInteger) As HRESULT
        <PreserveSig>
        Function SetSprites(startIndex As UInteger, spriteCount As UInteger,
           <MarshalAs(UnmanagedType.LPArray)> destinationRectangles As D2D1_RECT_F(),
        <MarshalAs(UnmanagedType.LPArray)> sourceRectangles As D2D1_RECT_U(),
            <MarshalAs(UnmanagedType.LPArray)> colors As D2D1_COLOR_F_STRUCT(),
           <MarshalAs(UnmanagedType.LPArray)> transforms As D2D1_MATRIX_3X2_F_STRUCT(), destinationRectanglesStride As UInteger, sourceRectanglesStride As UInteger, colorsStride As UInteger, transformsStride As UInteger) As HRESULT
        <PreserveSig>
        Function GetSprites(startIndex As UInteger, spriteCount As UInteger, <Out> ByRef destinationRectangles As IntPtr, <Out> ByRef sourceRectangles As IntPtr, <Out> ByRef colors As IntPtr, <Out> ByRef transforms As IntPtr) As HRESULT
        <PreserveSig>
        Function GetSpriteCount() As UInteger
        <PreserveSig>
        Sub Clear()
    End Interface

    <ComImport>
    <Guid("8c427831-3d90-4476-b647-c4fae349e4db")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1DeviceContext4
        Inherits ID2D1DeviceContext3
#Region "<ID2D1DeviceContext3>"

#Region "<ID2D1DeviceContext2>"

#Region "<ID2D1DeviceContext1>"
#Region "<ID2D1DeviceContext>"
#Region "<ID2D1RenderTarget>"

#Region "<ID2D1Resource>"

        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)

#End Region
        <PreserveSig>
        Overloads Function CreateBitmap(size As D2D1_SIZE_U, srcData As IntPtr, pitch As UInteger, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES, <Out> ByRef bitmap As ID2D1Bitmap) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapFromWicBitmap(wicBitmapSource As IWICBitmapSource, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES, <Out> ByRef bitmap As ID2D1Bitmap) As HRESULT
        <PreserveSig>
        Overloads Function CreateSharedBitmap(ByRef riid As Guid,
        <[In], Out> data As IntPtr, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES, <Out> ByRef bitmap As ID2D1Bitmap) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapBrush(bitmap As ID2D1Bitmap, ByRef bitmapBrushProperties As D2D1_BITMAP_BRUSH_PROPERTIES, brushProperties As D2D1_BRUSH_PROPERTIES, <Out> ByRef bitmapBrush As ID2D1BitmapBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateSolidColorBrush(color As D2D1_COLOR_F, brushProperties As D2D1_BRUSH_PROPERTIES, <Out> ByRef solidColorBrush As ID2D1SolidColorBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateGradientStopCollection(
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> gradientStops As D2D1_GRADIENT_STOP(), gradientStopsCount As UInteger, colorInterpolationGamma As D2D1_GAMMA, extendMode As D2D1_EXTEND_MODE, <Out> ByRef gradientStopCollection As ID2D1GradientStopCollection) As HRESULT
        <PreserveSig>
        Overloads Function CreateLinearGradientBrush(ByRef linearGradientBrushProperties As D2D1_LINEAR_GRADIENT_BRUSH_PROPERTIES, brushProperties As IntPtr, gradientStopCollection As ID2D1GradientStopCollection, <Out> ByRef linearGradientBrush As ID2D1LinearGradientBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateRadialGradientBrush(ByRef radialGradientBrushProperties As D2D1_RADIAL_GRADIENT_BRUSH_PROPERTIES, brushProperties As IntPtr, gradientStopCollection As ID2D1GradientStopCollection, <Out> ByRef radialGradientBrush As ID2D1RadialGradientBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateCompatibleRenderTarget(ByRef desiredSize As D2D1_SIZE_F, ByRef desiredPixelSize As D2D1_SIZE_U, ByRef desiredFormat As D2D1_PIXEL_FORMAT, options As D2D1_COMPATIBLE_RENDER_TARGET_OPTIONS, <Out> ByRef bitmapRenderTarget As ID2D1BitmapRenderTarget) As HRESULT
        <PreserveSig>
        Overloads Function CreateLayer(ByRef size As D2D1_SIZE_F, <Out> ByRef layer As ID2D1Layer) As HRESULT
        <PreserveSig>
        Overloads Sub CreateMesh(<Out> ByRef mesh As ID2D1Mesh)
        <PreserveSig>
        Overloads Sub DrawLine(point0 As D2D1_POINT_2F, point1 As D2D1_POINT_2F, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub DrawRectangle(ByRef rect As D2D1_RECT_F, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillRectangle(ByRef rect As D2D1_RECT_F, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub DrawRoundedRectangle(ByRef roundedRect As D2D1_ROUNDED_RECT, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillRoundedRectangle(ByRef roundedRect As D2D1_ROUNDED_RECT, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub DrawEllipse(ByRef ellipse As D2D1_ELLIPSE, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillEllipse(ByRef ellipse As D2D1_ELLIPSE, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub DrawGeometry(geometry As ID2D1Geometry, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillGeometry(geometry As ID2D1Geometry, brush As ID2D1Brush, Optional opacityBrush As ID2D1Brush = Nothing)
        <PreserveSig>
        Overloads Sub FillMesh(mesh As ID2D1Mesh, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub FillOpacityMask(opacityMask As ID2D1Bitmap, brush As ID2D1Brush, content As D2D1_OPACITY_MASK_CONTENT, destinationRectangle As D2D1_RECT_F, sourceRectangle As D2D1_RECT_F)
        <PreserveSig>
        Overloads Sub DrawBitmap(bitmap As ID2D1Bitmap, ByRef destinationRectangle As D2D1_RECT_F, opacity As Single, interpolationMode As D2D1_BITMAP_INTERPOLATION_MODE, ByRef sourceRectangle As D2D1_RECT_F)
        <PreserveSig>
        Overloads Sub DrawText(str As String, stringLength As UInteger, textFormat As IDWriteTextFormat, ByRef layoutRect As D2D1_RECT_F, defaultForegroundBrush As ID2D1Brush, options As D2D1_DRAW_TEXT_OPTIONS, measuringMode As DWRITE_MEASURING_MODE)
        <PreserveSig>
        Overloads Sub DrawTextLayout(origin As D2D1_POINT_2F, textLayout As IDWriteTextLayout, defaultForegroundBrush As ID2D1Brush, options As D2D1_DRAW_TEXT_OPTIONS)
        <PreserveSig>
        Overloads Sub DrawGlyphRun(baselineOrigin As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, foregroundBrush As ID2D1Brush, measuringMode As DWRITE_MEASURING_MODE)
        'new void DrawGlyphRun(D2D1_POINT_2F baselineOrigin, IntPtr glyphRun, ID2D1Brush foregroundBrush, DWRITE_MEASURING_MODE measuringMode);
        <PreserveSig>
        Overloads Sub SetTransform(transform As D2D1_MATRIX_3X2_F)
        <PreserveSig>
        Overloads Sub GetTransform(<Out> ByRef transform As D2D1_MATRIX_3X2_F_STRUCT)
        <PreserveSig>
        Overloads Sub SetAntialiasMode(antialiasMode As D2D1_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub GetAntialiasMode(<Out> ByRef antialiasMode As D2D1_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub SetTextAntialiasMode(textAntialiasMode As D2D1_TEXT_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub GetTextAntialiasMode(<Out> ByRef textAntialiasMode As D2D1_TEXT_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub SetTextRenderingParams(Optional textRenderingParams As IDWriteRenderingParams = Nothing)
        <PreserveSig>
        Overloads Sub GetTextRenderingParams(<Out> ByRef textRenderingParams As IDWriteRenderingParams)
        <PreserveSig>
        Overloads Sub SetTags(tag1 As ULong, tag2 As ULong)
        <PreserveSig>
        Overloads Sub GetTags(<Out> ByRef tag1 As ULong, <Out> ByRef tag2 As ULong)
        <PreserveSig>
        Overloads Sub PushLayer(ByRef layerParameters As D2D1_LAYER_PARAMETERS, layer As ID2D1Layer)
        <PreserveSig>
        Overloads Sub PopLayer()
        <PreserveSig>
        Overloads Sub Flush(<Out> ByRef tag1 As ULong, <Out> ByRef tag2 As ULong)
        <PreserveSig>
        Overloads Sub SaveDrawingState(
        <[In], Out> drawingStateBlock As ID2D1DrawingStateBlock)
        <PreserveSig>
        Overloads Sub RestoreDrawingState(drawingStateBlock As ID2D1DrawingStateBlock)
        <PreserveSig>
        Overloads Sub PushAxisAlignedClip(clipRect As D2D1_RECT_F, antialiasMode As D2D1_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub PopAxisAlignedClip()
        <PreserveSig>
        Overloads Sub Clear(clearColor As D2D1_COLOR_F)
        <PreserveSig>
        Overloads Sub BeginDraw()
        <PreserveSig>
        Overloads Function EndDraw(<Out> ByRef tag1 As ULong, <Out> ByRef tag2 As ULong) As HRESULT
        <PreserveSig>
        Overloads Sub GetPixelFormat(<Out> ByRef pixelFormat As D2D1_PIXEL_FORMAT)
        <PreserveSig>
        Overloads Sub SetDpi(dpiX As Single, dpiY As Single)
        <PreserveSig>
        Overloads Sub GetDpi(<Out> ByRef dpiX As Single, <Out> ByRef dpiY As Single)
        <PreserveSig>
        Overloads Sub GetSize(<Out> ByRef size As D2D1_SIZE_F)
        <PreserveSig>
        Overloads Sub GetPixelSize(<Out> ByRef size As D2D1_SIZE_U)
        <PreserveSig>
        Overloads Function GetMaximumBitmapSize() As UInteger
        <PreserveSig>
        Overloads Function IsSupported(renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES) As Boolean
#End Region

        <PreserveSig>
        Overloads Function CreateBitmap(size As D2D1_SIZE_U, sourceData As IntPtr, pitch As UInteger, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES1, <Out> ByRef bitmap As ID2D1Bitmap1) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapFromWicBitmap(wicBitmapSource As IWICBitmapSource, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES1, <Out> ByRef bitmap As ID2D1Bitmap1) As HRESULT
        <PreserveSig>
        Overloads Function CreateColorContext(space As D2D1_COLOR_SPACE, profile As IntPtr, profileSize As UInteger, <Out> ByRef colorContext As ID2D1ColorContext) As HRESULT
        <PreserveSig>
        Overloads Function CreateColorContextFromFilename(filename As String, <Out> ByRef colorContext As ID2D1ColorContext) As HRESULT
        <PreserveSig>
        Overloads Function CreateColorContextFromWicColorContext(wicColorContext As IWICColorContext, <Out> ByRef colorContext As ID2D1ColorContext) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapFromDxgiSurface(surface As IDXGISurface, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES1, <Out> ByRef bitmap As ID2D1Bitmap1) As HRESULT
        <PreserveSig>
        Overloads Function CreateEffect(ByRef effectId As Guid, <Out> ByRef effect As ID2D1Effect) As HRESULT
        <PreserveSig>
        Overloads Function CreateGradientStopCollection(
                    <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> straightAlphaGradientStops As D2D1_GRADIENT_STOP(), straightAlphaGradientStopsCount As UInteger, preInterpolationSpace As D2D1_COLOR_SPACE, postInterpolationSpace As D2D1_COLOR_SPACE, bufferPrecision As D2D1_BUFFER_PRECISION, extendMode As D2D1_EXTEND_MODE, colorInterpolationMode As D2D1_COLOR_INTERPOLATION_MODE, <Out> ByRef gradientStopCollection1 As ID2D1GradientStopCollection1) As HRESULT
        <PreserveSig>
        Overloads Function CreateImageBrush(image As ID2D1Image, ByRef imageBrushProperties As D2D1_IMAGE_BRUSH_PROPERTIES, brushProperties As D2D1_BRUSH_PROPERTIES, <Out> ByRef imageBrush As ID2D1ImageBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapBrush(bitmap As ID2D1Bitmap, ByRef bitmapBrushProperties As D2D1_BITMAP_BRUSH_PROPERTIES1, brushProperties As D2D1_BRUSH_PROPERTIES, <Out> ByRef bitmapBrush As ID2D1BitmapBrush1) As HRESULT
        <PreserveSig>
        Overloads Function CreateCommandList(<Out> ByRef commandList As ID2D1CommandList) As HRESULT
        <PreserveSig>
        Overloads Function IsDxgiFormatSupported(format As DXGI_FORMAT) As Boolean
        <PreserveSig>
        Overloads Function IsBufferPrecisionSupported(bufferPrecision As D2D1_BUFFER_PRECISION) As Boolean
        <PreserveSig>
        Overloads Function GetImageLocalBounds(image As ID2D1Image, <Out> ByRef localBounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Function GetImageWorldBounds(image As ID2D1Image, <Out> ByRef worldBounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Function GetGlyphRunWorldBounds(ByRef baselineOrigin As D2D1_POINT_2F, glyphRun As DWRITE_GLYPH_RUN, measuringMode As DWRITE_MEASURING_MODE, <Out> ByRef bounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Sub GetDevice(<Out> ByRef device As ID2D1Device)
        <PreserveSig>
        Overloads Sub SetTarget(image As ID2D1Image)
        <PreserveSig>
        Overloads Sub GetTarget(<Out> ByRef image As ID2D1Image)
        <PreserveSig>
        Overloads Sub SetRenderingControls(renderingControls As D2D1_RENDERING_CONTROLS)
        <PreserveSig>
        Overloads Sub GetRenderingControls(<Out> ByRef renderingControls As D2D1_RENDERING_CONTROLS)
        <PreserveSig>
        Overloads Sub SetPrimitiveBlend(primitiveBlend As D2D1_PRIMITIVE_BLEND)
        <PreserveSig>
        Overloads Sub GetPrimitiveBlend(<Out> ByRef primitiveBlend As D2D1_PRIMITIVE_BLEND)
        <PreserveSig>
        Overloads Sub SetUnitMode(unitMode As D2D1_UNIT_MODE)
        <PreserveSig>
        Overloads Sub GetUnitMode(<Out> ByRef unitMode As D2D1_UNIT_MODE)
        <PreserveSig>
        Overloads Sub DrawGlyphRun(baselineOrigin As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, glyphRunDescription As IntPtr, foregroundBrush As ID2D1Brush, measuringMode As DWRITE_MEASURING_MODE)
        <PreserveSig>
        Overloads Sub DrawImage(image As ID2D1Image, ByRef targetOffset As D2D1_POINT_2F, ByRef imageRectangle As D2D1_RECT_F, interpolationMode As D2D1_INTERPOLATION_MODE, compositeMode As D2D1_COMPOSITE_MODE)
        <PreserveSig>
        Overloads Sub DrawGdiMetafile(gdiMetafile As ID2D1GdiMetafile, ByRef targetOffset As D2D1_POINT_2F)
        <PreserveSig>
        Overloads Sub DrawBitmap(bitmap As ID2D1Bitmap, ByRef destinationRectangle As D2D1_RECT_F, opacity As Single, interpolationMode As D2D1_INTERPOLATION_MODE, ByRef sourceRectangle As D2D1_RECT_F, perspectiveTransform As D2D1_MATRIX_4X4_F)
        <PreserveSig>
        Overloads Sub PushLayer(ByRef layerParameters As D2D1_LAYER_PARAMETERS1, layer As ID2D1Layer)
        <PreserveSig>
        Overloads Function InvalidateEffectInputRectangle(effect As ID2D1Effect, input As UInteger, inputRectangle As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Function GetEffectInvalidRectangleCount(effect As ID2D1Effect, <Out> ByRef rectangleCount As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function GetEffectInvalidRectangles(effect As ID2D1Effect, <Out> ByRef rectangles As IntPtr, rectanglesCount As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function GetEffectRequiredInputRectangles(renderEffect As ID2D1Effect, renderImageRectangle As D2D1_RECT_F, inputDescriptions As D2D1_EFFECT_INPUT_DESCRIPTION, <Out> ByRef requiredInputRects As IntPtr, inputCount As UInteger) As HRESULT
        <PreserveSig>
        Overloads Sub FillOpacityMask(opacityMask As ID2D1Bitmap, brush As ID2D1Brush, ByRef destinationRectangle As D2D1_RECT_F, sourceRectangle As D2D1_RECT_F)
#End Region

        <PreserveSig>
        Overloads Function CreateFilledGeometryRealization(geometry As ID2D1Geometry, flatteningTolerance As Single, <Out> ByRef geometryRealization As ID2D1GeometryRealization) As HRESULT
        <PreserveSig>
        Overloads Function CreateStrokedGeometryRealization(geometry As ID2D1Geometry, flatteningTolerance As Single, strokeWidth As Single, strokeStyle As ID2D1StrokeStyle, <Out> ByRef geometryRealization As ID2D1GeometryRealization) As HRESULT
        <PreserveSig>
        Overloads Sub DrawGeometryRealization(geometryRealization As ID2D1GeometryRealization, brush As ID2D1Brush)
#End Region

        <PreserveSig>
        Overloads Function CreateInk(startPoint As D2D1_INK_POINT, <Out> ByRef ink As ID2D1Ink) As HRESULT
        <PreserveSig>
        Overloads Function CreateInkStyle(inkStyleProperties As D2D1_INK_STYLE_PROPERTIES, <Out> ByRef inkStyle As ID2D1InkStyle) As HRESULT
        <PreserveSig>
        Overloads Function CreateGradientMesh(patches As D2D1_GRADIENT_MESH_PATCH, patchesCount As UInteger, <Out> ByRef gradientMesh As ID2D1GradientMesh) As HRESULT
        <PreserveSig>
        Overloads Function CreateImageSourceFromWic(wicBitmapSource As IWICBitmapSource, loadingOptions As D2D1_IMAGE_SOURCE_LOADING_OPTIONS, alphaMode As D2D1_ALPHA_MODE, <Out> ByRef imageSource As ID2D1ImageSourceFromWic) As HRESULT
        <PreserveSig>
        Overloads Function CreateLookupTable3D(precision As D2D1_BUFFER_PRECISION, extents As UInteger, data As IntPtr, dataCount As UInteger, strides As UInteger, <Out> ByRef lookupTable As ID2D1LookupTable3D) As HRESULT
        <PreserveSig>
        Overloads Function CreateImageSourceFromDxgi(surfaces As IntPtr, surfaceCount As UInteger, colorSpace As DXGI_COLOR_SPACE_TYPE, options As D2D1_IMAGE_SOURCE_FROM_DXGI_OPTIONS, <Out> ByRef imageSource As ID2D1ImageSource) As HRESULT
        <PreserveSig>
        Overloads Function GetGradientMeshWorldBounds(gradientMesh As ID2D1GradientMesh, <Out> ByRef pBounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Sub DrawInk(ink As ID2D1Ink, brush As ID2D1Brush, inkStyle As ID2D1InkStyle)
        <PreserveSig>
        Overloads Sub DrawGradientMesh(gradientMesh As ID2D1GradientMesh)
        <PreserveSig>
        Overloads Sub DrawGdiMetafile(gdiMetafile As ID2D1GdiMetafile, ByRef destinationRectangle As D2D1_RECT_F, ByRef sourceRectangle As D2D1_RECT_F)
        <PreserveSig>
        Overloads Function CreateTransformedImageSource(imageSource As ID2D1ImageSource, properties As D2D1_TRANSFORMED_IMAGE_SOURCE_PROPERTIES, <Out> ByRef transformedImageSource As ID2D1TransformedImageSource) As HRESULT

#End Region

        <PreserveSig>
        Overloads Function CreateSpriteBatch(<Out> ByRef spriteBatch As ID2D1SpriteBatch) As HRESULT
        <PreserveSig>
        Overloads Sub DrawSpriteBatch(spriteBatch As ID2D1SpriteBatch, startIndex As UInteger, spriteCount As UInteger, bitmap As ID2D1Bitmap, Optional interpolationMode As D2D1_BITMAP_INTERPOLATION_MODE = D2D1_BITMAP_INTERPOLATION_MODE.D2D1_BITMAP_INTERPOLATION_MODE_LINEAR, Optional spriteOptions As D2D1_SPRITE_OPTIONS = D2D1_SPRITE_OPTIONS.D2D1_SPRITE_OPTIONS_NONE)

#End Region

        <PreserveSig>
        Function CreateSvgGlyphStyle(<Out> ByRef svgGlyphStyle As ID2D1SvgGlyphStyle) As HRESULT
        <PreserveSig>
        Overloads Sub DrawText(sString As String, stringLength As UInteger, textFormat As IDWriteTextFormat, ByRef layoutRect As D2D1_RECT_F, defaultFillBrush As ID2D1Brush, svgGlyphStyle As ID2D1SvgGlyphStyle, Optional colorPaletteIndex As UInteger = 0, Optional options As D2D1_DRAW_TEXT_OPTIONS = D2D1_DRAW_TEXT_OPTIONS.D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT, Optional measuringMode As DWRITE_MEASURING_MODE = DWRITE_MEASURING_MODE.DWRITE_MEASURING_MODE_NATURAL)
        <PreserveSig>
        Overloads Sub DrawTextLayout(origin As D2D1_POINT_2F, textLayout As IDWriteTextLayout, defaultFillBrush As ID2D1Brush, svgGlyphStyle As ID2D1SvgGlyphStyle, Optional colorPaletteIndex As UInteger = 0, Optional options As D2D1_DRAW_TEXT_OPTIONS = D2D1_DRAW_TEXT_OPTIONS.D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT)
        <PreserveSig>
        Sub DrawColorBitmapGlyphRun(glyphImageFormat As DWRITE_GLYPH_IMAGE_FORMATS, baselineOrigin As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, Optional measuringMode As DWRITE_MEASURING_MODE = DWRITE_MEASURING_MODE.DWRITE_MEASURING_MODE_NATURAL, Optional bitmapSnapOption As D2D1_COLOR_BITMAP_GLYPH_SNAP_OPTION = D2D1_COLOR_BITMAP_GLYPH_SNAP_OPTION.D2D1_COLOR_BITMAP_GLYPH_SNAP_OPTION_DEFAULT)
        <PreserveSig>
        Sub DrawSvgGlyphRun(baselineOrigin As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, Optional defaultFillBrush As ID2D1Brush = Nothing, Optional svgGlyphStyle As ID2D1SvgGlyphStyle = Nothing, Optional colorPaletteIndex As UInteger = 0, Optional measuringMode As DWRITE_MEASURING_MODE = DWRITE_MEASURING_MODE.DWRITE_MEASURING_MODE_NATURAL)
        <PreserveSig>
        Function GetColorBitmapGlyphImage(glyphImageFormat As DWRITE_GLYPH_IMAGE_FORMATS, glyphOrigin As D2D1_POINT_2F, fontFace As IDWriteFontFace, fontEmSize As Single, glyphIndex As UShort, isSideways As Boolean, worldTransform As D2D1_MATRIX_3X2_F, dpiX As Single, dpiY As Single, <Out> ByRef glyphTransform As D2D1_MATRIX_3X2_F, <Out> ByRef glyphImage As ID2D1Image) As HRESULT
        <PreserveSig>
        Function GetSvgGlyphImage(ByRef glyphOrigin As D2D1_POINT_2F, fontFace As IDWriteFontFace, fontEmSize As Single, glyphIndex As UShort, isSideways As Boolean, worldTransform As D2D1_MATRIX_3X2_F, defaultFillBrush As ID2D1Brush, svgGlyphStyle As ID2D1SvgGlyphStyle, colorPaletteIndex As UInteger, <Out> ByRef glyphTransform As D2D1_MATRIX_3X2_F, <Out> ByRef glyphImage As ID2D1CommandList) As HRESULT
    End Interface

    <ComImport>
    <Guid("af671749-d241-4db8-8e41-dcc2e5c1a438")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1SvgGlyphStyle
        Inherits ID2D1Resource
#Region "<ID2D1Resource>"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region

        <PreserveSig>
        Function SetFill(brush As ID2D1Brush) As HRESULT
        <PreserveSig>
        Sub GetFill(<Out> ByRef brush As ID2D1Brush)
        <PreserveSig>
        Function SetStroke(brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional dashes As Single = 0.0F, Optional dashesCount As UInteger = 0, Optional dashOffset As Single = 1.0F) As HRESULT
        <PreserveSig>
        Function GetStrokeDashesCount() As UInteger
        <PreserveSig>
        Sub GetStroke(<Out> ByRef brush As ID2D1Brush, <Out> ByRef strokeWidth As Single, <Out> ByRef dashes As Single, dashesCount As UInteger, <Out> ByRef dashOffset As Single)
    End Interface

    ''' <summary>
    ''' Specifies the pixel snapping policy when rendering color bitmap glyphs.
    ''' </summary>
    Public Enum D2D1_COLOR_BITMAP_GLYPH_SNAP_OPTION
        ''' <summary>
        ''' Color bitmap glyph positions are snapped to the nearest pixel if the bitmap
        ''' resolution matches that of the device context.
        ''' </summary>
        D2D1_COLOR_BITMAP_GLYPH_SNAP_OPTION_DEFAULT = 0

        ''' <summary>
        ''' Color bitmap glyph positions are not snapped.
        ''' </summary>
        D2D1_COLOR_BITMAP_GLYPH_SNAP_OPTION_DISABLE = 1
        D2D1_COLOR_BITMAP_GLYPH_SNAP_OPTION_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' Fonts may contain multiple drawable data formats for glyphs. These flags specify which formats
    ''' are supported in the font, either at a font-wide level or per glyph, and the app may use them
    ''' to tell DWrite which formats to return when splitting a color glyph run.
    ''' </summary>
    Public Enum DWRITE_GLYPH_IMAGE_FORMATS
        ''' <summary>
        ''' Indicates no data is available for this glyph.
        ''' </summary>
        DWRITE_GLYPH_IMAGE_FORMATS_NONE = &H0

        ''' <summary>
        ''' The glyph has TrueType outlines.
        ''' </summary>
        DWRITE_GLYPH_IMAGE_FORMATS_TRUETYPE = &H1

        ''' <summary>
        ''' The glyph has CFF outlines.
        ''' </summary>
        DWRITE_GLYPH_IMAGE_FORMATS_CFF = &H2

        ''' <summary>
        ''' The glyph has multilayered COLR data.
        ''' </summary>
        DWRITE_GLYPH_IMAGE_FORMATS_COLR = &H4

        ''' <summary>
        ''' The glyph has SVG outlines as standard XML.
        ''' </summary>
        ''' <remarks>
        ''' Fonts may store the content gzip'd rather than plain text,
        ''' indicated by the first two bytes as gzip header {0x1F 0x8B}.
        ''' </remarks>
        DWRITE_GLYPH_IMAGE_FORMATS_SVG = &H8

        ''' <summary>
        ''' The glyph has PNG image data, with standard PNG IHDR.
        ''' </summary>
        DWRITE_GLYPH_IMAGE_FORMATS_PNG = &H10

        ''' <summary>
        ''' The glyph has JPEG image data, with standard JIFF SOI header.
        ''' </summary>
        DWRITE_GLYPH_IMAGE_FORMATS_JPEG = &H20

        ''' <summary>
        ''' The glyph has TIFF image data.
        ''' </summary>
        DWRITE_GLYPH_IMAGE_FORMATS_TIFF = &H40

        ''' <summary>
        ''' The glyph has raw 32-bit premultiplied BGRA data.
        ''' </summary>
        DWRITE_GLYPH_IMAGE_FORMATS_PREMULTIPLIED_B8G8R8A8 = &H80
    End Enum


    <ComImport>
    <Guid("7836d248-68cc-4df6-b9e8-de991bf62eb7")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1DeviceContext5
        Inherits ID2D1DeviceContext4
#Region "<ID2D1DeviceContext4>"

#Region "<ID2D1DeviceContext3>"

#Region "<ID2D1DeviceContext2>"

#Region "<ID2D1DeviceContext1>"
#Region "<ID2D1DeviceContext>"
#Region "<ID2D1RenderTarget>"

#Region "<ID2D1Resource>"

        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)

#End Region

        <PreserveSig>
        Overloads Function CreateBitmap(size As D2D1_SIZE_U, srcData As IntPtr, pitch As UInteger, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES, <Out> ByRef bitmap As ID2D1Bitmap) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapFromWicBitmap(wicBitmapSource As IWICBitmapSource, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES, <Out> ByRef bitmap As ID2D1Bitmap) As HRESULT
        <PreserveSig>
        Overloads Function CreateSharedBitmap(ByRef riid As Guid,
        <[In], Out> data As IntPtr, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES, <Out> ByRef bitmap As ID2D1Bitmap) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapBrush(bitmap As ID2D1Bitmap, ByRef bitmapBrushProperties As D2D1_BITMAP_BRUSH_PROPERTIES, brushProperties As D2D1_BRUSH_PROPERTIES, <Out> ByRef bitmapBrush As ID2D1BitmapBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateSolidColorBrush(color As D2D1_COLOR_F, brushProperties As D2D1_BRUSH_PROPERTIES, <Out> ByRef solidColorBrush As ID2D1SolidColorBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateGradientStopCollection(
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> gradientStops As D2D1_GRADIENT_STOP(), gradientStopsCount As UInteger, colorInterpolationGamma As D2D1_GAMMA, extendMode As D2D1_EXTEND_MODE, <Out> ByRef gradientStopCollection As ID2D1GradientStopCollection) As HRESULT
        <PreserveSig>
        Overloads Function CreateLinearGradientBrush(ByRef linearGradientBrushProperties As D2D1_LINEAR_GRADIENT_BRUSH_PROPERTIES, brushProperties As IntPtr, gradientStopCollection As ID2D1GradientStopCollection, <Out> ByRef linearGradientBrush As ID2D1LinearGradientBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateRadialGradientBrush(ByRef radialGradientBrushProperties As D2D1_RADIAL_GRADIENT_BRUSH_PROPERTIES, brushProperties As IntPtr, gradientStopCollection As ID2D1GradientStopCollection, <Out> ByRef radialGradientBrush As ID2D1RadialGradientBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateCompatibleRenderTarget(ByRef desiredSize As D2D1_SIZE_F, ByRef desiredPixelSize As D2D1_SIZE_U, ByRef desiredFormat As D2D1_PIXEL_FORMAT, options As D2D1_COMPATIBLE_RENDER_TARGET_OPTIONS, <Out> ByRef bitmapRenderTarget As ID2D1BitmapRenderTarget) As HRESULT
        <PreserveSig>
        Overloads Function CreateLayer(ByRef size As D2D1_SIZE_F, <Out> ByRef layer As ID2D1Layer) As HRESULT
        <PreserveSig>
        Overloads Sub CreateMesh(<Out> ByRef mesh As ID2D1Mesh)
        <PreserveSig>
        Overloads Sub DrawLine(point0 As D2D1_POINT_2F, point1 As D2D1_POINT_2F, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub DrawRectangle(ByRef rect As D2D1_RECT_F, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillRectangle(ByRef rect As D2D1_RECT_F, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub DrawRoundedRectangle(ByRef roundedRect As D2D1_ROUNDED_RECT, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillRoundedRectangle(ByRef roundedRect As D2D1_ROUNDED_RECT, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub DrawEllipse(ByRef ellipse As D2D1_ELLIPSE, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillEllipse(ByRef ellipse As D2D1_ELLIPSE, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub DrawGeometry(geometry As ID2D1Geometry, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillGeometry(geometry As ID2D1Geometry, brush As ID2D1Brush, Optional opacityBrush As ID2D1Brush = Nothing)
        <PreserveSig>
        Overloads Sub FillMesh(mesh As ID2D1Mesh, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub FillOpacityMask(opacityMask As ID2D1Bitmap, brush As ID2D1Brush, content As D2D1_OPACITY_MASK_CONTENT, destinationRectangle As D2D1_RECT_F, sourceRectangle As D2D1_RECT_F)
        <PreserveSig>
        Overloads Sub DrawBitmap(bitmap As ID2D1Bitmap, ByRef destinationRectangle As D2D1_RECT_F, opacity As Single, interpolationMode As D2D1_BITMAP_INTERPOLATION_MODE, ByRef sourceRectangle As D2D1_RECT_F)
        <PreserveSig>
        Overloads Sub DrawText(str As String, stringLength As UInteger, textFormat As IDWriteTextFormat, ByRef layoutRect As D2D1_RECT_F, defaultForegroundBrush As ID2D1Brush, options As D2D1_DRAW_TEXT_OPTIONS, measuringMode As DWRITE_MEASURING_MODE)
        <PreserveSig>
        Overloads Sub DrawTextLayout(origin As D2D1_POINT_2F, textLayout As IDWriteTextLayout, defaultForegroundBrush As ID2D1Brush, options As D2D1_DRAW_TEXT_OPTIONS)
        <PreserveSig>
        Overloads Sub DrawGlyphRun(baselineOrigin As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, foregroundBrush As ID2D1Brush, measuringMode As DWRITE_MEASURING_MODE)
        <PreserveSig>
        Overloads Sub SetTransform(transform As D2D1_MATRIX_3X2_F)
        <PreserveSig>
        Overloads Sub GetTransform(<Out> ByRef transform As D2D1_MATRIX_3X2_F_STRUCT)
        <PreserveSig>
        Overloads Sub SetAntialiasMode(antialiasMode As D2D1_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub GetAntialiasMode(<Out> ByRef antialiasMode As D2D1_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub SetTextAntialiasMode(textAntialiasMode As D2D1_TEXT_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub GetTextAntialiasMode(<Out> ByRef textAntialiasMode As D2D1_TEXT_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub SetTextRenderingParams(Optional textRenderingParams As IDWriteRenderingParams = Nothing)
        <PreserveSig>
        Overloads Sub GetTextRenderingParams(<Out> ByRef textRenderingParams As IDWriteRenderingParams)
        <PreserveSig>
        Overloads Sub SetTags(tag1 As ULong, tag2 As ULong)
        <PreserveSig>
        Overloads Sub GetTags(<Out> ByRef tag1 As ULong, <Out> ByRef tag2 As ULong)
        <PreserveSig>
        Overloads Sub PushLayer(ByRef layerParameters As D2D1_LAYER_PARAMETERS, layer As ID2D1Layer)
        <PreserveSig>
        Overloads Sub PopLayer()
        <PreserveSig>
        Overloads Sub Flush(<Out> ByRef tag1 As ULong, <Out> ByRef tag2 As ULong)
        <PreserveSig>
        Overloads Sub SaveDrawingState(
        <[In], Out> drawingStateBlock As ID2D1DrawingStateBlock)
        <PreserveSig>
        Overloads Sub RestoreDrawingState(drawingStateBlock As ID2D1DrawingStateBlock)
        <PreserveSig>
        Overloads Sub PushAxisAlignedClip(clipRect As D2D1_RECT_F, antialiasMode As D2D1_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub PopAxisAlignedClip()
        <PreserveSig>
        Overloads Sub Clear(clearColor As D2D1_COLOR_F)
        <PreserveSig>
        Overloads Sub BeginDraw()
        <PreserveSig>
        Overloads Function EndDraw(<Out> ByRef tag1 As ULong, <Out> ByRef tag2 As ULong) As HRESULT
        <PreserveSig>
        Overloads Sub GetPixelFormat(<Out> ByRef pixelFormat As D2D1_PIXEL_FORMAT)
        <PreserveSig>
        Overloads Sub SetDpi(dpiX As Single, dpiY As Single)
        <PreserveSig>
        Overloads Sub GetDpi(<Out> ByRef dpiX As Single, <Out> ByRef dpiY As Single)
        <PreserveSig>
        Overloads Sub GetSize(<Out> ByRef size As D2D1_SIZE_F)
        <PreserveSig>
        Overloads Sub GetPixelSize(<Out> ByRef size As D2D1_SIZE_U)
        <PreserveSig>
        Overloads Function GetMaximumBitmapSize() As UInteger
        <PreserveSig>
        Overloads Function IsSupported(renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES) As Boolean
#End Region

        <PreserveSig>
        Overloads Function CreateBitmap(size As D2D1_SIZE_U, sourceData As IntPtr, pitch As UInteger, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES1, <Out> ByRef bitmap As ID2D1Bitmap1) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapFromWicBitmap(wicBitmapSource As IWICBitmapSource, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES1, <Out> ByRef bitmap As ID2D1Bitmap1) As HRESULT
        <PreserveSig>
        Overloads Function CreateColorContext(space As D2D1_COLOR_SPACE, profile As IntPtr, profileSize As UInteger, <Out> ByRef colorContext As ID2D1ColorContext) As HRESULT
        <PreserveSig>
        Overloads Function CreateColorContextFromFilename(filename As String, <Out> ByRef colorContext As ID2D1ColorContext) As HRESULT
        <PreserveSig>
        Overloads Function CreateColorContextFromWicColorContext(wicColorContext As IWICColorContext, <Out> ByRef colorContext As ID2D1ColorContext) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapFromDxgiSurface(surface As IDXGISurface, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES1, <Out> ByRef bitmap As ID2D1Bitmap1) As HRESULT
        <PreserveSig>
        Overloads Function CreateEffect(ByRef effectId As Guid, <Out> ByRef effect As ID2D1Effect) As HRESULT
        <PreserveSig>
        Overloads Function CreateGradientStopCollection(
                        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> straightAlphaGradientStops As D2D1_GRADIENT_STOP(), straightAlphaGradientStopsCount As UInteger, preInterpolationSpace As D2D1_COLOR_SPACE, postInterpolationSpace As D2D1_COLOR_SPACE, bufferPrecision As D2D1_BUFFER_PRECISION, extendMode As D2D1_EXTEND_MODE, colorInterpolationMode As D2D1_COLOR_INTERPOLATION_MODE, <Out> ByRef gradientStopCollection1 As ID2D1GradientStopCollection1) As HRESULT
        <PreserveSig>
        Overloads Function CreateImageBrush(image As ID2D1Image, ByRef imageBrushProperties As D2D1_IMAGE_BRUSH_PROPERTIES, brushProperties As D2D1_BRUSH_PROPERTIES, <Out> ByRef imageBrush As ID2D1ImageBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapBrush(bitmap As ID2D1Bitmap, ByRef bitmapBrushProperties As D2D1_BITMAP_BRUSH_PROPERTIES1, brushProperties As D2D1_BRUSH_PROPERTIES, <Out> ByRef bitmapBrush As ID2D1BitmapBrush1) As HRESULT
        <PreserveSig>
        Overloads Function CreateCommandList(<Out> ByRef commandList As ID2D1CommandList) As HRESULT
        <PreserveSig>
        Overloads Function IsDxgiFormatSupported(format As DXGI_FORMAT) As Boolean
        <PreserveSig>
        Overloads Function IsBufferPrecisionSupported(bufferPrecision As D2D1_BUFFER_PRECISION) As Boolean
        <PreserveSig>
        Overloads Function GetImageLocalBounds(image As ID2D1Image, <Out> ByRef localBounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Function GetImageWorldBounds(image As ID2D1Image, <Out> ByRef worldBounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Function GetGlyphRunWorldBounds(ByRef baselineOrigin As D2D1_POINT_2F, glyphRun As DWRITE_GLYPH_RUN, measuringMode As DWRITE_MEASURING_MODE, <Out> ByRef bounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Sub GetDevice(<Out> ByRef device As ID2D1Device)
        <PreserveSig>
        Overloads Sub SetTarget(image As ID2D1Image)
        <PreserveSig>
        Overloads Sub GetTarget(<Out> ByRef image As ID2D1Image)
        <PreserveSig>
        Overloads Sub SetRenderingControls(renderingControls As D2D1_RENDERING_CONTROLS)
        <PreserveSig>
        Overloads Sub GetRenderingControls(<Out> ByRef renderingControls As D2D1_RENDERING_CONTROLS)
        <PreserveSig>
        Overloads Sub SetPrimitiveBlend(primitiveBlend As D2D1_PRIMITIVE_BLEND)
        <PreserveSig>
        Overloads Sub GetPrimitiveBlend(<Out> ByRef primitiveBlend As D2D1_PRIMITIVE_BLEND)
        <PreserveSig>
        Overloads Sub SetUnitMode(unitMode As D2D1_UNIT_MODE)
        <PreserveSig>
        Overloads Sub GetUnitMode(<Out> ByRef unitMode As D2D1_UNIT_MODE)
        <PreserveSig>
        Overloads Sub DrawGlyphRun(baselineOrigin As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, glyphRunDescription As IntPtr, foregroundBrush As ID2D1Brush, measuringMode As DWRITE_MEASURING_MODE)
        <PreserveSig>
        Overloads Sub DrawImage(image As ID2D1Image, ByRef targetOffset As D2D1_POINT_2F, ByRef imageRectangle As D2D1_RECT_F, interpolationMode As D2D1_INTERPOLATION_MODE, compositeMode As D2D1_COMPOSITE_MODE)
        <PreserveSig>
        Overloads Sub DrawGdiMetafile(gdiMetafile As ID2D1GdiMetafile, ByRef targetOffset As D2D1_POINT_2F)
        <PreserveSig>
        Overloads Sub DrawBitmap(bitmap As ID2D1Bitmap, ByRef destinationRectangle As D2D1_RECT_F, opacity As Single, interpolationMode As D2D1_INTERPOLATION_MODE, ByRef sourceRectangle As D2D1_RECT_F, perspectiveTransform As D2D1_MATRIX_4X4_F)
        <PreserveSig>
        Overloads Sub PushLayer(ByRef layerParameters As D2D1_LAYER_PARAMETERS1, layer As ID2D1Layer)
        <PreserveSig>
        Overloads Function InvalidateEffectInputRectangle(effect As ID2D1Effect, input As UInteger, inputRectangle As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Function GetEffectInvalidRectangleCount(effect As ID2D1Effect, <Out> ByRef rectangleCount As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function GetEffectInvalidRectangles(effect As ID2D1Effect, <Out> ByRef rectangles As IntPtr, rectanglesCount As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function GetEffectRequiredInputRectangles(renderEffect As ID2D1Effect, renderImageRectangle As D2D1_RECT_F, inputDescriptions As D2D1_EFFECT_INPUT_DESCRIPTION, <Out> ByRef requiredInputRects As IntPtr, inputCount As UInteger) As HRESULT
        <PreserveSig>
        Overloads Sub FillOpacityMask(opacityMask As ID2D1Bitmap, brush As ID2D1Brush, ByRef destinationRectangle As D2D1_RECT_F, sourceRectangle As D2D1_RECT_F)
#End Region

        <PreserveSig>
        Overloads Function CreateFilledGeometryRealization(geometry As ID2D1Geometry, flatteningTolerance As Single, <Out> ByRef geometryRealization As ID2D1GeometryRealization) As HRESULT
        <PreserveSig>
        Overloads Function CreateStrokedGeometryRealization(geometry As ID2D1Geometry, flatteningTolerance As Single, strokeWidth As Single, strokeStyle As ID2D1StrokeStyle, <Out> ByRef geometryRealization As ID2D1GeometryRealization) As HRESULT
        <PreserveSig>
        Overloads Sub DrawGeometryRealization(geometryRealization As ID2D1GeometryRealization, brush As ID2D1Brush)
#End Region

        <PreserveSig>
        Overloads Function CreateInk(startPoint As D2D1_INK_POINT, <Out> ByRef ink As ID2D1Ink) As HRESULT
        <PreserveSig>
        Overloads Function CreateInkStyle(inkStyleProperties As D2D1_INK_STYLE_PROPERTIES, <Out> ByRef inkStyle As ID2D1InkStyle) As HRESULT
        <PreserveSig>
        Overloads Function CreateGradientMesh(patches As D2D1_GRADIENT_MESH_PATCH, patchesCount As UInteger, <Out> ByRef gradientMesh As ID2D1GradientMesh) As HRESULT
        <PreserveSig>
        Overloads Function CreateImageSourceFromWic(wicBitmapSource As IWICBitmapSource, loadingOptions As D2D1_IMAGE_SOURCE_LOADING_OPTIONS, alphaMode As D2D1_ALPHA_MODE, <Out> ByRef imageSource As ID2D1ImageSourceFromWic) As HRESULT
        <PreserveSig>
        Overloads Function CreateLookupTable3D(precision As D2D1_BUFFER_PRECISION, extents As UInteger, data As IntPtr, dataCount As UInteger, strides As UInteger, <Out> ByRef lookupTable As ID2D1LookupTable3D) As HRESULT
        <PreserveSig>
        Overloads Function CreateImageSourceFromDxgi(surfaces As IntPtr, surfaceCount As UInteger, colorSpace As DXGI_COLOR_SPACE_TYPE, options As D2D1_IMAGE_SOURCE_FROM_DXGI_OPTIONS, <Out> ByRef imageSource As ID2D1ImageSource) As HRESULT
        <PreserveSig>
        Overloads Function GetGradientMeshWorldBounds(gradientMesh As ID2D1GradientMesh, <Out> ByRef pBounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Sub DrawInk(ink As ID2D1Ink, brush As ID2D1Brush, inkStyle As ID2D1InkStyle)
        <PreserveSig>
        Overloads Sub DrawGradientMesh(gradientMesh As ID2D1GradientMesh)
        <PreserveSig>
        Overloads Sub DrawGdiMetafile(gdiMetafile As ID2D1GdiMetafile, ByRef destinationRectangle As D2D1_RECT_F, ByRef sourceRectangle As D2D1_RECT_F)
        <PreserveSig>
        Overloads Function CreateTransformedImageSource(imageSource As ID2D1ImageSource, properties As D2D1_TRANSFORMED_IMAGE_SOURCE_PROPERTIES, <Out> ByRef transformedImageSource As ID2D1TransformedImageSource) As HRESULT

#End Region

        <PreserveSig>
        Overloads Function CreateSpriteBatch(<Out> ByRef spriteBatch As ID2D1SpriteBatch) As HRESULT
        <PreserveSig>
        Overloads Sub DrawSpriteBatch(spriteBatch As ID2D1SpriteBatch, startIndex As UInteger, spriteCount As UInteger, bitmap As ID2D1Bitmap, Optional interpolationMode As D2D1_BITMAP_INTERPOLATION_MODE = D2D1_BITMAP_INTERPOLATION_MODE.D2D1_BITMAP_INTERPOLATION_MODE_LINEAR, Optional spriteOptions As D2D1_SPRITE_OPTIONS = D2D1_SPRITE_OPTIONS.D2D1_SPRITE_OPTIONS_NONE)

#End Region

        <PreserveSig>
        Overloads Function CreateSvgGlyphStyle(<Out> ByRef svgGlyphStyle As ID2D1SvgGlyphStyle) As HRESULT
        <PreserveSig>
        Overloads Sub DrawText(sString As String, stringLength As UInteger, textFormat As IDWriteTextFormat, ByRef layoutRect As D2D1_RECT_F, defaultFillBrush As ID2D1Brush, svgGlyphStyle As ID2D1SvgGlyphStyle, Optional colorPaletteIndex As UInteger = 0, Optional options As D2D1_DRAW_TEXT_OPTIONS = D2D1_DRAW_TEXT_OPTIONS.D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT, Optional measuringMode As DWRITE_MEASURING_MODE = DWRITE_MEASURING_MODE.DWRITE_MEASURING_MODE_NATURAL)
        <PreserveSig>
        Overloads Sub DrawTextLayout(origin As D2D1_POINT_2F, textLayout As IDWriteTextLayout, defaultFillBrush As ID2D1Brush, svgGlyphStyle As ID2D1SvgGlyphStyle, Optional colorPaletteIndex As UInteger = 0, Optional options As D2D1_DRAW_TEXT_OPTIONS = D2D1_DRAW_TEXT_OPTIONS.D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT)
        <PreserveSig>
        Overloads Sub DrawColorBitmapGlyphRun(glyphImageFormat As DWRITE_GLYPH_IMAGE_FORMATS, baselineOrigin As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, Optional measuringMode As DWRITE_MEASURING_MODE = DWRITE_MEASURING_MODE.DWRITE_MEASURING_MODE_NATURAL, Optional bitmapSnapOption As D2D1_COLOR_BITMAP_GLYPH_SNAP_OPTION = D2D1_COLOR_BITMAP_GLYPH_SNAP_OPTION.D2D1_COLOR_BITMAP_GLYPH_SNAP_OPTION_DEFAULT)
        <PreserveSig>
        Overloads Sub DrawSvgGlyphRun(baselineOrigin As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, Optional defaultFillBrush As ID2D1Brush = Nothing, Optional svgGlyphStyle As ID2D1SvgGlyphStyle = Nothing, Optional colorPaletteIndex As UInteger = 0, Optional measuringMode As DWRITE_MEASURING_MODE = DWRITE_MEASURING_MODE.DWRITE_MEASURING_MODE_NATURAL)
        <PreserveSig>
        Overloads Function GetColorBitmapGlyphImage(glyphImageFormat As DWRITE_GLYPH_IMAGE_FORMATS, glyphOrigin As D2D1_POINT_2F, fontFace As IDWriteFontFace, fontEmSize As Single, glyphIndex As UShort, isSideways As Boolean, worldTransform As D2D1_MATRIX_3X2_F, dpiX As Single, dpiY As Single, <Out> ByRef glyphTransform As D2D1_MATRIX_3X2_F, <Out> ByRef glyphImage As ID2D1Image) As HRESULT
        <PreserveSig>
        Overloads Function GetSvgGlyphImage(ByRef glyphOrigin As D2D1_POINT_2F, fontFace As IDWriteFontFace, fontEmSize As Single, glyphIndex As UShort, isSideways As Boolean, worldTransform As D2D1_MATRIX_3X2_F, defaultFillBrush As ID2D1Brush, svgGlyphStyle As ID2D1SvgGlyphStyle, colorPaletteIndex As UInteger, <Out> ByRef glyphTransform As D2D1_MATRIX_3X2_F, <Out> ByRef glyphImage As ID2D1CommandList) As HRESULT
#End Region

        <PreserveSig>
        Function CreateSvgDocument(inputXmlStream As ComTypes.IStream, viewportSize As D2D1_SIZE_F, <Out> ByRef svgDocument As ID2D1SvgDocument) As HRESULT
        <PreserveSig>
        Sub DrawSvgDocument(svgDocument As ID2D1SvgDocument)
        <PreserveSig>
        Function CreateColorContextFromDxgiColorSpace(colorSpace As DXGI_COLOR_SPACE_TYPE, <Out> ByRef colorContext As ID2D1ColorContext1) As HRESULT
        <PreserveSig>
        Function CreateColorContextFromSimpleColorProfile(ByRef simpleProfile As D2D1_SIMPLE_COLOR_PROFILE, <Out> ByRef colorContext As ID2D1ColorContext1) As HRESULT
    End Interface

    <ComImport>
    <Guid("985f7e37-4ed0-4a19-98a3-15b0edfde306")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1DeviceContext6
        Inherits ID2D1DeviceContext5
#Region "<ID2D1DeviceContext5>"
#Region "<ID2D1DeviceContext4>"
#Region "<ID2D1DeviceContext3>"
#Region "<ID2D1DeviceContext2>"
#Region "<ID2D1DeviceContext1>"
#Region "<ID2D1DeviceContext>"
#Region "<ID2D1RenderTarget>"

#Region "<ID2D1Resource>"

        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)

#End Region

        <PreserveSig>
        Overloads Function CreateBitmap(size As D2D1_SIZE_U, srcData As IntPtr, pitch As UInteger, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES, <Out> ByRef bitmap As ID2D1Bitmap) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapFromWicBitmap(wicBitmapSource As IWICBitmapSource, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES, <Out> ByRef bitmap As ID2D1Bitmap) As HRESULT
        <PreserveSig>
        Overloads Function CreateSharedBitmap(ByRef riid As Guid,
        <[In], Out> data As IntPtr, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES, <Out> ByRef bitmap As ID2D1Bitmap) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapBrush(bitmap As ID2D1Bitmap, ByRef bitmapBrushProperties As D2D1_BITMAP_BRUSH_PROPERTIES, brushProperties As D2D1_BRUSH_PROPERTIES, <Out> ByRef bitmapBrush As ID2D1BitmapBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateSolidColorBrush(color As D2D1_COLOR_F, brushProperties As D2D1_BRUSH_PROPERTIES, <Out> ByRef solidColorBrush As ID2D1SolidColorBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateGradientStopCollection(
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> gradientStops As D2D1_GRADIENT_STOP(), gradientStopsCount As UInteger, colorInterpolationGamma As D2D1_GAMMA, extendMode As D2D1_EXTEND_MODE, <Out> ByRef gradientStopCollection As ID2D1GradientStopCollection) As HRESULT
        <PreserveSig>
        Overloads Function CreateLinearGradientBrush(ByRef linearGradientBrushProperties As D2D1_LINEAR_GRADIENT_BRUSH_PROPERTIES, brushProperties As IntPtr, gradientStopCollection As ID2D1GradientStopCollection, <Out> ByRef linearGradientBrush As ID2D1LinearGradientBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateRadialGradientBrush(ByRef radialGradientBrushProperties As D2D1_RADIAL_GRADIENT_BRUSH_PROPERTIES, brushProperties As IntPtr, gradientStopCollection As ID2D1GradientStopCollection, <Out> ByRef radialGradientBrush As ID2D1RadialGradientBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateCompatibleRenderTarget(ByRef desiredSize As D2D1_SIZE_F, ByRef desiredPixelSize As D2D1_SIZE_U, ByRef desiredFormat As D2D1_PIXEL_FORMAT, options As D2D1_COMPATIBLE_RENDER_TARGET_OPTIONS, <Out> ByRef bitmapRenderTarget As ID2D1BitmapRenderTarget) As HRESULT
        <PreserveSig>
        Overloads Function CreateLayer(ByRef size As D2D1_SIZE_F, <Out> ByRef layer As ID2D1Layer) As HRESULT
        <PreserveSig>
        Overloads Sub CreateMesh(<Out> ByRef mesh As ID2D1Mesh)
        <PreserveSig>
        Overloads Sub DrawLine(point0 As D2D1_POINT_2F, point1 As D2D1_POINT_2F, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub DrawRectangle(ByRef rect As D2D1_RECT_F, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillRectangle(ByRef rect As D2D1_RECT_F, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub DrawRoundedRectangle(ByRef roundedRect As D2D1_ROUNDED_RECT, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillRoundedRectangle(ByRef roundedRect As D2D1_ROUNDED_RECT, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub DrawEllipse(ByRef ellipse As D2D1_ELLIPSE, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillEllipse(ByRef ellipse As D2D1_ELLIPSE, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub DrawGeometry(geometry As ID2D1Geometry, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillGeometry(geometry As ID2D1Geometry, brush As ID2D1Brush, Optional opacityBrush As ID2D1Brush = Nothing)
        <PreserveSig>
        Overloads Sub FillMesh(mesh As ID2D1Mesh, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub FillOpacityMask(opacityMask As ID2D1Bitmap, brush As ID2D1Brush, content As D2D1_OPACITY_MASK_CONTENT, destinationRectangle As D2D1_RECT_F, sourceRectangle As D2D1_RECT_F)
        <PreserveSig>
        Overloads Sub DrawBitmap(bitmap As ID2D1Bitmap, ByRef destinationRectangle As D2D1_RECT_F, opacity As Single, interpolationMode As D2D1_BITMAP_INTERPOLATION_MODE, ByRef sourceRectangle As D2D1_RECT_F)
        <PreserveSig>
        Overloads Sub DrawText(str As String, stringLength As UInteger, textFormat As IDWriteTextFormat, ByRef layoutRect As D2D1_RECT_F, defaultForegroundBrush As ID2D1Brush, options As D2D1_DRAW_TEXT_OPTIONS, measuringMode As DWRITE_MEASURING_MODE)
        <PreserveSig>
        Overloads Sub DrawTextLayout(origin As D2D1_POINT_2F, textLayout As IDWriteTextLayout, defaultForegroundBrush As ID2D1Brush, options As D2D1_DRAW_TEXT_OPTIONS)
        <PreserveSig>
        Overloads Sub DrawGlyphRun(baselineOrigin As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, foregroundBrush As ID2D1Brush, measuringMode As DWRITE_MEASURING_MODE)
        <PreserveSig>
        Overloads Sub SetTransform(transform As D2D1_MATRIX_3X2_F)
        <PreserveSig>
        Overloads Sub GetTransform(<Out> ByRef transform As D2D1_MATRIX_3X2_F_STRUCT)
        <PreserveSig>
        Overloads Sub SetAntialiasMode(antialiasMode As D2D1_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub GetAntialiasMode(<Out> ByRef antialiasMode As D2D1_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub SetTextAntialiasMode(textAntialiasMode As D2D1_TEXT_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub GetTextAntialiasMode(<Out> ByRef textAntialiasMode As D2D1_TEXT_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub SetTextRenderingParams(Optional textRenderingParams As IDWriteRenderingParams = Nothing)
        <PreserveSig>
        Overloads Sub GetTextRenderingParams(<Out> ByRef textRenderingParams As IDWriteRenderingParams)
        <PreserveSig>
        Overloads Sub SetTags(tag1 As ULong, tag2 As ULong)
        <PreserveSig>
        Overloads Sub GetTags(<Out> ByRef tag1 As ULong, <Out> ByRef tag2 As ULong)
        <PreserveSig>
        Overloads Sub PushLayer(ByRef layerParameters As D2D1_LAYER_PARAMETERS, layer As ID2D1Layer)
        <PreserveSig>
        Overloads Sub PopLayer()
        <PreserveSig>
        Overloads Sub Flush(<Out> ByRef tag1 As ULong, <Out> ByRef tag2 As ULong)
        <PreserveSig>
        Overloads Sub SaveDrawingState(
        <[In], Out> drawingStateBlock As ID2D1DrawingStateBlock)
        <PreserveSig>
        Overloads Sub RestoreDrawingState(drawingStateBlock As ID2D1DrawingStateBlock)
        <PreserveSig>
        Overloads Sub PushAxisAlignedClip(clipRect As D2D1_RECT_F, antialiasMode As D2D1_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub PopAxisAlignedClip()
        <PreserveSig>
        Overloads Sub Clear(clearColor As D2D1_COLOR_F)
        <PreserveSig>
        Overloads Sub BeginDraw()
        <PreserveSig>
        Overloads Function EndDraw(<Out> ByRef tag1 As ULong, <Out> ByRef tag2 As ULong) As HRESULT
        <PreserveSig>
        Overloads Sub GetPixelFormat(<Out> ByRef pixelFormat As D2D1_PIXEL_FORMAT)
        <PreserveSig>
        Overloads Sub SetDpi(dpiX As Single, dpiY As Single)
        <PreserveSig>
        Overloads Sub GetDpi(<Out> ByRef dpiX As Single, <Out> ByRef dpiY As Single)
        <PreserveSig>
        Overloads Sub GetSize(<Out> ByRef size As D2D1_SIZE_F)
        <PreserveSig>
        Overloads Sub GetPixelSize(<Out> ByRef size As D2D1_SIZE_U)
        <PreserveSig>
        Overloads Function GetMaximumBitmapSize() As UInteger
        <PreserveSig>
        Overloads Function IsSupported(renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES) As Boolean
#End Region

        <PreserveSig>
        Overloads Function CreateBitmap(size As D2D1_SIZE_U, sourceData As IntPtr, pitch As UInteger, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES1, <Out> ByRef bitmap As ID2D1Bitmap1) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapFromWicBitmap(wicBitmapSource As IWICBitmapSource, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES1, <Out> ByRef bitmap As ID2D1Bitmap1) As HRESULT
        <PreserveSig>
        Overloads Function CreateColorContext(space As D2D1_COLOR_SPACE, profile As IntPtr, profileSize As UInteger, <Out> ByRef colorContext As ID2D1ColorContext) As HRESULT
        <PreserveSig>
        Overloads Function CreateColorContextFromFilename(filename As String, <Out> ByRef colorContext As ID2D1ColorContext) As HRESULT
        <PreserveSig>
        Overloads Function CreateColorContextFromWicColorContext(wicColorContext As IWICColorContext, <Out> ByRef colorContext As ID2D1ColorContext) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapFromDxgiSurface(surface As IDXGISurface, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES1, <Out> ByRef bitmap As ID2D1Bitmap1) As HRESULT
        <PreserveSig>
        Overloads Function CreateEffect(ByRef effectId As Guid, <Out> ByRef effect As ID2D1Effect) As HRESULT
        <PreserveSig>
        Overloads Function CreateGradientStopCollection(
                        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> straightAlphaGradientStops As D2D1_GRADIENT_STOP(), straightAlphaGradientStopsCount As UInteger, preInterpolationSpace As D2D1_COLOR_SPACE, postInterpolationSpace As D2D1_COLOR_SPACE, bufferPrecision As D2D1_BUFFER_PRECISION, extendMode As D2D1_EXTEND_MODE, colorInterpolationMode As D2D1_COLOR_INTERPOLATION_MODE, <Out> ByRef gradientStopCollection1 As ID2D1GradientStopCollection1) As HRESULT
        <PreserveSig>
        Overloads Function CreateImageBrush(image As ID2D1Image, ByRef imageBrushProperties As D2D1_IMAGE_BRUSH_PROPERTIES, brushProperties As D2D1_BRUSH_PROPERTIES, <Out> ByRef imageBrush As ID2D1ImageBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapBrush(bitmap As ID2D1Bitmap, ByRef bitmapBrushProperties As D2D1_BITMAP_BRUSH_PROPERTIES1, brushProperties As D2D1_BRUSH_PROPERTIES, <Out> ByRef bitmapBrush As ID2D1BitmapBrush1) As HRESULT
        <PreserveSig>
        Overloads Function CreateCommandList(<Out> ByRef commandList As ID2D1CommandList) As HRESULT
        <PreserveSig>
        Overloads Function IsDxgiFormatSupported(format As DXGI_FORMAT) As Boolean
        <PreserveSig>
        Overloads Function IsBufferPrecisionSupported(bufferPrecision As D2D1_BUFFER_PRECISION) As Boolean
        <PreserveSig>
        Overloads Function GetImageLocalBounds(image As ID2D1Image, <Out> ByRef localBounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Function GetImageWorldBounds(image As ID2D1Image, <Out> ByRef worldBounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Function GetGlyphRunWorldBounds(ByRef baselineOrigin As D2D1_POINT_2F, glyphRun As DWRITE_GLYPH_RUN, measuringMode As DWRITE_MEASURING_MODE, <Out> ByRef bounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Sub GetDevice(<Out> ByRef device As ID2D1Device)
        <PreserveSig>
        Overloads Sub SetTarget(image As ID2D1Image)
        <PreserveSig>
        Overloads Sub GetTarget(<Out> ByRef image As ID2D1Image)
        <PreserveSig>
        Overloads Sub SetRenderingControls(renderingControls As D2D1_RENDERING_CONTROLS)
        <PreserveSig>
        Overloads Sub GetRenderingControls(<Out> ByRef renderingControls As D2D1_RENDERING_CONTROLS)
        <PreserveSig>
        Overloads Sub SetPrimitiveBlend(primitiveBlend As D2D1_PRIMITIVE_BLEND)
        <PreserveSig>
        Overloads Sub GetPrimitiveBlend(<Out> ByRef primitiveBlend As D2D1_PRIMITIVE_BLEND)
        <PreserveSig>
        Overloads Sub SetUnitMode(unitMode As D2D1_UNIT_MODE)
        <PreserveSig>
        Overloads Sub GetUnitMode(<Out> ByRef unitMode As D2D1_UNIT_MODE)
        <PreserveSig>
        Overloads Sub DrawGlyphRun(baselineOrigin As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, glyphRunDescription As IntPtr, foregroundBrush As ID2D1Brush, measuringMode As DWRITE_MEASURING_MODE)
        <PreserveSig>
        Overloads Sub DrawImage(image As ID2D1Image, ByRef targetOffset As D2D1_POINT_2F, ByRef imageRectangle As D2D1_RECT_F, interpolationMode As D2D1_INTERPOLATION_MODE, compositeMode As D2D1_COMPOSITE_MODE)
        <PreserveSig>
        Overloads Sub DrawGdiMetafile(gdiMetafile As ID2D1GdiMetafile, ByRef targetOffset As D2D1_POINT_2F)
        <PreserveSig>
        Overloads Sub DrawBitmap(bitmap As ID2D1Bitmap, ByRef destinationRectangle As D2D1_RECT_F, opacity As Single, interpolationMode As D2D1_INTERPOLATION_MODE, ByRef sourceRectangle As D2D1_RECT_F, perspectiveTransform As D2D1_MATRIX_4X4_F)
        <PreserveSig>
        Overloads Sub PushLayer(ByRef layerParameters As D2D1_LAYER_PARAMETERS1, layer As ID2D1Layer)
        <PreserveSig>
        Overloads Function InvalidateEffectInputRectangle(effect As ID2D1Effect, input As UInteger, inputRectangle As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Function GetEffectInvalidRectangleCount(effect As ID2D1Effect, <Out> ByRef rectangleCount As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function GetEffectInvalidRectangles(effect As ID2D1Effect, <Out> ByRef rectangles As IntPtr, rectanglesCount As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function GetEffectRequiredInputRectangles(renderEffect As ID2D1Effect, renderImageRectangle As D2D1_RECT_F, inputDescriptions As D2D1_EFFECT_INPUT_DESCRIPTION, <Out> ByRef requiredInputRects As IntPtr, inputCount As UInteger) As HRESULT
        <PreserveSig>
        Overloads Sub FillOpacityMask(opacityMask As ID2D1Bitmap, brush As ID2D1Brush, ByRef destinationRectangle As D2D1_RECT_F, sourceRectangle As D2D1_RECT_F)
#End Region

        <PreserveSig>
        Overloads Function CreateFilledGeometryRealization(geometry As ID2D1Geometry, flatteningTolerance As Single, <Out> ByRef geometryRealization As ID2D1GeometryRealization) As HRESULT
        <PreserveSig>
        Overloads Function CreateStrokedGeometryRealization(geometry As ID2D1Geometry, flatteningTolerance As Single, strokeWidth As Single, strokeStyle As ID2D1StrokeStyle, <Out> ByRef geometryRealization As ID2D1GeometryRealization) As HRESULT
        <PreserveSig>
        Overloads Sub DrawGeometryRealization(geometryRealization As ID2D1GeometryRealization, brush As ID2D1Brush)
#End Region

        <PreserveSig>
        Overloads Function CreateInk(startPoint As D2D1_INK_POINT, <Out> ByRef ink As ID2D1Ink) As HRESULT
        <PreserveSig>
        Overloads Function CreateInkStyle(inkStyleProperties As D2D1_INK_STYLE_PROPERTIES, <Out> ByRef inkStyle As ID2D1InkStyle) As HRESULT
        <PreserveSig>
        Overloads Function CreateGradientMesh(patches As D2D1_GRADIENT_MESH_PATCH, patchesCount As UInteger, <Out> ByRef gradientMesh As ID2D1GradientMesh) As HRESULT
        <PreserveSig>
        Overloads Function CreateImageSourceFromWic(wicBitmapSource As IWICBitmapSource, loadingOptions As D2D1_IMAGE_SOURCE_LOADING_OPTIONS, alphaMode As D2D1_ALPHA_MODE, <Out> ByRef imageSource As ID2D1ImageSourceFromWic) As HRESULT
        <PreserveSig>
        Overloads Function CreateLookupTable3D(precision As D2D1_BUFFER_PRECISION, extents As UInteger, data As IntPtr, dataCount As UInteger, strides As UInteger, <Out> ByRef lookupTable As ID2D1LookupTable3D) As HRESULT
        <PreserveSig>
        Overloads Function CreateImageSourceFromDxgi(surfaces As IntPtr, surfaceCount As UInteger, colorSpace As DXGI_COLOR_SPACE_TYPE, options As D2D1_IMAGE_SOURCE_FROM_DXGI_OPTIONS, <Out> ByRef imageSource As ID2D1ImageSource) As HRESULT
        <PreserveSig>
        Overloads Function GetGradientMeshWorldBounds(gradientMesh As ID2D1GradientMesh, <Out> ByRef pBounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Sub DrawInk(ink As ID2D1Ink, brush As ID2D1Brush, inkStyle As ID2D1InkStyle)
        <PreserveSig>
        Overloads Sub DrawGradientMesh(gradientMesh As ID2D1GradientMesh)
        <PreserveSig>
        Overloads Sub DrawGdiMetafile(gdiMetafile As ID2D1GdiMetafile, ByRef destinationRectangle As D2D1_RECT_F, ByRef sourceRectangle As D2D1_RECT_F)
        <PreserveSig>
        Overloads Function CreateTransformedImageSource(imageSource As ID2D1ImageSource, properties As D2D1_TRANSFORMED_IMAGE_SOURCE_PROPERTIES, <Out> ByRef transformedImageSource As ID2D1TransformedImageSource) As HRESULT

#End Region

        <PreserveSig>
        Overloads Function CreateSpriteBatch(<Out> ByRef spriteBatch As ID2D1SpriteBatch) As HRESULT
        <PreserveSig>
        Overloads Sub DrawSpriteBatch(spriteBatch As ID2D1SpriteBatch, startIndex As UInteger, spriteCount As UInteger, bitmap As ID2D1Bitmap, Optional interpolationMode As D2D1_BITMAP_INTERPOLATION_MODE = D2D1_BITMAP_INTERPOLATION_MODE.D2D1_BITMAP_INTERPOLATION_MODE_LINEAR, Optional spriteOptions As D2D1_SPRITE_OPTIONS = D2D1_SPRITE_OPTIONS.D2D1_SPRITE_OPTIONS_NONE)

#End Region

        <PreserveSig>
        Overloads Function CreateSvgGlyphStyle(<Out> ByRef svgGlyphStyle As ID2D1SvgGlyphStyle) As HRESULT
        <PreserveSig>
        Overloads Sub DrawText(sString As String, stringLength As UInteger, textFormat As IDWriteTextFormat, ByRef layoutRect As D2D1_RECT_F, defaultFillBrush As ID2D1Brush, svgGlyphStyle As ID2D1SvgGlyphStyle, Optional colorPaletteIndex As UInteger = 0, Optional options As D2D1_DRAW_TEXT_OPTIONS = D2D1_DRAW_TEXT_OPTIONS.D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT, Optional measuringMode As DWRITE_MEASURING_MODE = DWRITE_MEASURING_MODE.DWRITE_MEASURING_MODE_NATURAL)
        <PreserveSig>
        Overloads Sub DrawTextLayout(origin As D2D1_POINT_2F, textLayout As IDWriteTextLayout, defaultFillBrush As ID2D1Brush, svgGlyphStyle As ID2D1SvgGlyphStyle, Optional colorPaletteIndex As UInteger = 0, Optional options As D2D1_DRAW_TEXT_OPTIONS = D2D1_DRAW_TEXT_OPTIONS.D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT)
        <PreserveSig>
        Overloads Sub DrawColorBitmapGlyphRun(glyphImageFormat As DWRITE_GLYPH_IMAGE_FORMATS, baselineOrigin As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, Optional measuringMode As DWRITE_MEASURING_MODE = DWRITE_MEASURING_MODE.DWRITE_MEASURING_MODE_NATURAL, Optional bitmapSnapOption As D2D1_COLOR_BITMAP_GLYPH_SNAP_OPTION = D2D1_COLOR_BITMAP_GLYPH_SNAP_OPTION.D2D1_COLOR_BITMAP_GLYPH_SNAP_OPTION_DEFAULT)
        <PreserveSig>
        Overloads Sub DrawSvgGlyphRun(baselineOrigin As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, Optional defaultFillBrush As ID2D1Brush = Nothing, Optional svgGlyphStyle As ID2D1SvgGlyphStyle = Nothing, Optional colorPaletteIndex As UInteger = 0, Optional measuringMode As DWRITE_MEASURING_MODE = DWRITE_MEASURING_MODE.DWRITE_MEASURING_MODE_NATURAL)
        <PreserveSig>
        Overloads Function GetColorBitmapGlyphImage(glyphImageFormat As DWRITE_GLYPH_IMAGE_FORMATS, glyphOrigin As D2D1_POINT_2F, fontFace As IDWriteFontFace, fontEmSize As Single, glyphIndex As UShort, isSideways As Boolean, worldTransform As D2D1_MATRIX_3X2_F, dpiX As Single, dpiY As Single, <Out> ByRef glyphTransform As D2D1_MATRIX_3X2_F, <Out> ByRef glyphImage As ID2D1Image) As HRESULT
        <PreserveSig>
        Overloads Function GetSvgGlyphImage(ByRef glyphOrigin As D2D1_POINT_2F, fontFace As IDWriteFontFace, fontEmSize As Single, glyphIndex As UShort, isSideways As Boolean, worldTransform As D2D1_MATRIX_3X2_F, defaultFillBrush As ID2D1Brush, svgGlyphStyle As ID2D1SvgGlyphStyle, colorPaletteIndex As UInteger, <Out> ByRef glyphTransform As D2D1_MATRIX_3X2_F, <Out> ByRef glyphImage As ID2D1CommandList) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateSvgDocument(inputXmlStream As ComTypes.IStream, viewportSize As D2D1_SIZE_F, <Out> ByRef svgDocument As ID2D1SvgDocument) As HRESULT
        <PreserveSig>
        Overloads Sub DrawSvgDocument(svgDocument As ID2D1SvgDocument)
        <PreserveSig>
        Overloads Function CreateColorContextFromDxgiColorSpace(colorSpace As DXGI_COLOR_SPACE_TYPE, <Out> ByRef colorContext As ID2D1ColorContext1) As HRESULT
        <PreserveSig>
        Overloads Function CreateColorContextFromSimpleColorProfile(ByRef simpleProfile As D2D1_SIMPLE_COLOR_PROFILE, <Out> ByRef colorContext As ID2D1ColorContext1) As HRESULT
#End Region

        <PreserveSig>
        Sub BlendImage(image As ID2D1Image, blendMode As D2D1_BLEND_MODE, ByRef targetOffset As D2D1_POINT_2F, ByRef imageRectangle As D2D1_RECT_F, Optional interpolationMode As D2D1_INTERPOLATION_MODE = D2D1_INTERPOLATION_MODE.D2D1_INTERPOLATION_MODE_LINEAR)
    End Interface

    <ComImport>
    <Guid("ec891cf7-9b69-4851-9def-4e0915771e62")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1DeviceContext7
        Inherits ID2D1DeviceContext6
#Region "<ID2D1DeviceContext5>"
#Region "<ID2D1DeviceContext4>"
#Region "<ID2D1DeviceContext3>"
#Region "<ID2D1DeviceContext2>"
#Region "<ID2D1DeviceContext1>"
#Region "<ID2D1DeviceContext>"
#Region "<ID2D1RenderTarget>"

#Region "<ID2D1Resource>"

        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)

#End Region

        <PreserveSig>
        Overloads Function CreateBitmap(size As D2D1_SIZE_U, srcData As IntPtr, pitch As UInteger, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES, <Out> ByRef bitmap As ID2D1Bitmap) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapFromWicBitmap(wicBitmapSource As IWICBitmapSource, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES, <Out> ByRef bitmap As ID2D1Bitmap) As HRESULT
        <PreserveSig>
        Overloads Function CreateSharedBitmap(ByRef riid As Guid,
        <[In], Out> data As IntPtr, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES, <Out> ByRef bitmap As ID2D1Bitmap) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapBrush(bitmap As ID2D1Bitmap, ByRef bitmapBrushProperties As D2D1_BITMAP_BRUSH_PROPERTIES, brushProperties As D2D1_BRUSH_PROPERTIES, <Out> ByRef bitmapBrush As ID2D1BitmapBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateSolidColorBrush(color As D2D1_COLOR_F, brushProperties As D2D1_BRUSH_PROPERTIES, <Out> ByRef solidColorBrush As ID2D1SolidColorBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateGradientStopCollection(
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> gradientStops As D2D1_GRADIENT_STOP(), gradientStopsCount As UInteger, colorInterpolationGamma As D2D1_GAMMA, extendMode As D2D1_EXTEND_MODE, <Out> ByRef gradientStopCollection As ID2D1GradientStopCollection) As HRESULT
        <PreserveSig>
        Overloads Function CreateLinearGradientBrush(ByRef linearGradientBrushProperties As D2D1_LINEAR_GRADIENT_BRUSH_PROPERTIES, brushProperties As IntPtr, gradientStopCollection As ID2D1GradientStopCollection, <Out> ByRef linearGradientBrush As ID2D1LinearGradientBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateRadialGradientBrush(ByRef radialGradientBrushProperties As D2D1_RADIAL_GRADIENT_BRUSH_PROPERTIES, brushProperties As IntPtr, gradientStopCollection As ID2D1GradientStopCollection, <Out> ByRef radialGradientBrush As ID2D1RadialGradientBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateCompatibleRenderTarget(ByRef desiredSize As D2D1_SIZE_F, ByRef desiredPixelSize As D2D1_SIZE_U, ByRef desiredFormat As D2D1_PIXEL_FORMAT, options As D2D1_COMPATIBLE_RENDER_TARGET_OPTIONS, <Out> ByRef bitmapRenderTarget As ID2D1BitmapRenderTarget) As HRESULT
        <PreserveSig>
        Overloads Function CreateLayer(ByRef size As D2D1_SIZE_F, <Out> ByRef layer As ID2D1Layer) As HRESULT
        <PreserveSig>
        Overloads Sub CreateMesh(<Out> ByRef mesh As ID2D1Mesh)
        <PreserveSig>
        Overloads Sub DrawLine(point0 As D2D1_POINT_2F, point1 As D2D1_POINT_2F, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub DrawRectangle(ByRef rect As D2D1_RECT_F, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillRectangle(ByRef rect As D2D1_RECT_F, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub DrawRoundedRectangle(ByRef roundedRect As D2D1_ROUNDED_RECT, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillRoundedRectangle(ByRef roundedRect As D2D1_ROUNDED_RECT, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub DrawEllipse(ByRef ellipse As D2D1_ELLIPSE, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillEllipse(ByRef ellipse As D2D1_ELLIPSE, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub DrawGeometry(geometry As ID2D1Geometry, brush As ID2D1Brush, Optional strokeWidth As Single = 1.0F, Optional strokeStyle As ID2D1StrokeStyle = Nothing)
        <PreserveSig>
        Overloads Sub FillGeometry(geometry As ID2D1Geometry, brush As ID2D1Brush, Optional opacityBrush As ID2D1Brush = Nothing)
        <PreserveSig>
        Overloads Sub FillMesh(mesh As ID2D1Mesh, brush As ID2D1Brush)
        <PreserveSig>
        Overloads Sub FillOpacityMask(opacityMask As ID2D1Bitmap, brush As ID2D1Brush, content As D2D1_OPACITY_MASK_CONTENT, destinationRectangle As D2D1_RECT_F, sourceRectangle As D2D1_RECT_F)
        <PreserveSig>
        Overloads Sub DrawBitmap(bitmap As ID2D1Bitmap, ByRef destinationRectangle As D2D1_RECT_F, opacity As Single, interpolationMode As D2D1_BITMAP_INTERPOLATION_MODE, ByRef sourceRectangle As D2D1_RECT_F)
        <PreserveSig>
        Overloads Sub DrawText(str As String, stringLength As UInteger, textFormat As IDWriteTextFormat, ByRef layoutRect As D2D1_RECT_F, defaultForegroundBrush As ID2D1Brush, options As D2D1_DRAW_TEXT_OPTIONS, measuringMode As DWRITE_MEASURING_MODE)
        <PreserveSig>
        Overloads Sub DrawTextLayout(origin As D2D1_POINT_2F, textLayout As IDWriteTextLayout, defaultForegroundBrush As ID2D1Brush, options As D2D1_DRAW_TEXT_OPTIONS)
        <PreserveSig>
        Overloads Sub DrawGlyphRun(baselineOrigin As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, foregroundBrush As ID2D1Brush, measuringMode As DWRITE_MEASURING_MODE)
        <PreserveSig>
        Overloads Sub SetTransform(transform As D2D1_MATRIX_3X2_F)
        <PreserveSig>
        Overloads Sub GetTransform(<Out> ByRef transform As D2D1_MATRIX_3X2_F_STRUCT)
        <PreserveSig>
        Overloads Sub SetAntialiasMode(antialiasMode As D2D1_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub GetAntialiasMode(<Out> ByRef antialiasMode As D2D1_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub SetTextAntialiasMode(textAntialiasMode As D2D1_TEXT_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub GetTextAntialiasMode(<Out> ByRef textAntialiasMode As D2D1_TEXT_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub SetTextRenderingParams(Optional textRenderingParams As IDWriteRenderingParams = Nothing)
        <PreserveSig>
        Overloads Sub GetTextRenderingParams(<Out> ByRef textRenderingParams As IDWriteRenderingParams)
        <PreserveSig>
        Overloads Sub SetTags(tag1 As ULong, tag2 As ULong)
        <PreserveSig>
        Overloads Sub GetTags(<Out> ByRef tag1 As ULong, <Out> ByRef tag2 As ULong)
        <PreserveSig>
        Overloads Sub PushLayer(ByRef layerParameters As D2D1_LAYER_PARAMETERS, layer As ID2D1Layer)
        <PreserveSig>
        Overloads Sub PopLayer()
        <PreserveSig>
        Overloads Sub Flush(<Out> ByRef tag1 As ULong, <Out> ByRef tag2 As ULong)
        <PreserveSig>
        Overloads Sub SaveDrawingState(
        <[In], Out> drawingStateBlock As ID2D1DrawingStateBlock)
        <PreserveSig>
        Overloads Sub RestoreDrawingState(drawingStateBlock As ID2D1DrawingStateBlock)
        <PreserveSig>
        Overloads Sub PushAxisAlignedClip(clipRect As D2D1_RECT_F, antialiasMode As D2D1_ANTIALIAS_MODE)
        <PreserveSig>
        Overloads Sub PopAxisAlignedClip()
        <PreserveSig>
        Overloads Sub Clear(clearColor As D2D1_COLOR_F)
        <PreserveSig>
        Overloads Sub BeginDraw()
        <PreserveSig>
        Overloads Function EndDraw(<Out> ByRef tag1 As ULong, <Out> ByRef tag2 As ULong) As HRESULT
        <PreserveSig>
        Overloads Sub GetPixelFormat(<Out> ByRef pixelFormat As D2D1_PIXEL_FORMAT)
        <PreserveSig>
        Overloads Sub SetDpi(dpiX As Single, dpiY As Single)
        <PreserveSig>
        Overloads Sub GetDpi(<Out> ByRef dpiX As Single, <Out> ByRef dpiY As Single)
        <PreserveSig>
        Overloads Sub GetSize(<Out> ByRef size As D2D1_SIZE_F)
        <PreserveSig>
        Overloads Sub GetPixelSize(<Out> ByRef size As D2D1_SIZE_U)
        <PreserveSig>
        Overloads Function GetMaximumBitmapSize() As UInteger
        <PreserveSig>
        Overloads Function IsSupported(renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES) As Boolean
#End Region

        <PreserveSig>
        Overloads Function CreateBitmap(size As D2D1_SIZE_U, sourceData As IntPtr, pitch As UInteger, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES1, <Out> ByRef bitmap As ID2D1Bitmap1) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapFromWicBitmap(wicBitmapSource As IWICBitmapSource, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES1, <Out> ByRef bitmap As ID2D1Bitmap1) As HRESULT
        <PreserveSig>
        Overloads Function CreateColorContext(space As D2D1_COLOR_SPACE, profile As IntPtr, profileSize As UInteger, <Out> ByRef colorContext As ID2D1ColorContext) As HRESULT
        <PreserveSig>
        Overloads Function CreateColorContextFromFilename(filename As String, <Out> ByRef colorContext As ID2D1ColorContext) As HRESULT
        <PreserveSig>
        Overloads Function CreateColorContextFromWicColorContext(wicColorContext As IWICColorContext, <Out> ByRef colorContext As ID2D1ColorContext) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapFromDxgiSurface(surface As IDXGISurface, ByRef bitmapProperties As D2D1_BITMAP_PROPERTIES1, <Out> ByRef bitmap As ID2D1Bitmap1) As HRESULT
        <PreserveSig>
        Overloads Function CreateEffect(ByRef effectId As Guid, <Out> ByRef effect As ID2D1Effect) As HRESULT
        <PreserveSig>
        Overloads Function CreateGradientStopCollection(
                        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> straightAlphaGradientStops As D2D1_GRADIENT_STOP(), straightAlphaGradientStopsCount As UInteger, preInterpolationSpace As D2D1_COLOR_SPACE, postInterpolationSpace As D2D1_COLOR_SPACE, bufferPrecision As D2D1_BUFFER_PRECISION, extendMode As D2D1_EXTEND_MODE, colorInterpolationMode As D2D1_COLOR_INTERPOLATION_MODE, <Out> ByRef gradientStopCollection1 As ID2D1GradientStopCollection1) As HRESULT
        <PreserveSig>
        Overloads Function CreateImageBrush(image As ID2D1Image, ByRef imageBrushProperties As D2D1_IMAGE_BRUSH_PROPERTIES, brushProperties As D2D1_BRUSH_PROPERTIES, <Out> ByRef imageBrush As ID2D1ImageBrush) As HRESULT
        <PreserveSig>
        Overloads Function CreateBitmapBrush(bitmap As ID2D1Bitmap, ByRef bitmapBrushProperties As D2D1_BITMAP_BRUSH_PROPERTIES1, brushProperties As D2D1_BRUSH_PROPERTIES, <Out> ByRef bitmapBrush As ID2D1BitmapBrush1) As HRESULT
        <PreserveSig>
        Overloads Function CreateCommandList(<Out> ByRef commandList As ID2D1CommandList) As HRESULT
        <PreserveSig>
        Overloads Function IsDxgiFormatSupported(format As DXGI_FORMAT) As Boolean
        <PreserveSig>
        Overloads Function IsBufferPrecisionSupported(bufferPrecision As D2D1_BUFFER_PRECISION) As Boolean
        <PreserveSig>
        Overloads Function GetImageLocalBounds(image As ID2D1Image, <Out> ByRef localBounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Function GetImageWorldBounds(image As ID2D1Image, <Out> ByRef worldBounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Function GetGlyphRunWorldBounds(ByRef baselineOrigin As D2D1_POINT_2F, glyphRun As DWRITE_GLYPH_RUN, measuringMode As DWRITE_MEASURING_MODE, <Out> ByRef bounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Sub GetDevice(<Out> ByRef device As ID2D1Device)
        <PreserveSig>
        Overloads Sub SetTarget(image As ID2D1Image)
        <PreserveSig>
        Overloads Sub GetTarget(<Out> ByRef image As ID2D1Image)
        <PreserveSig>
        Overloads Sub SetRenderingControls(renderingControls As D2D1_RENDERING_CONTROLS)
        <PreserveSig>
        Overloads Sub GetRenderingControls(<Out> ByRef renderingControls As D2D1_RENDERING_CONTROLS)
        <PreserveSig>
        Overloads Sub SetPrimitiveBlend(primitiveBlend As D2D1_PRIMITIVE_BLEND)
        <PreserveSig>
        Overloads Sub GetPrimitiveBlend(<Out> ByRef primitiveBlend As D2D1_PRIMITIVE_BLEND)
        <PreserveSig>
        Overloads Sub SetUnitMode(unitMode As D2D1_UNIT_MODE)
        <PreserveSig>
        Overloads Sub GetUnitMode(<Out> ByRef unitMode As D2D1_UNIT_MODE)
        <PreserveSig>
        Overloads Sub DrawGlyphRun(baselineOrigin As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, glyphRunDescription As IntPtr, foregroundBrush As ID2D1Brush, measuringMode As DWRITE_MEASURING_MODE)
        <PreserveSig>
        Overloads Sub DrawImage(image As ID2D1Image, ByRef targetOffset As D2D1_POINT_2F, ByRef imageRectangle As D2D1_RECT_F, interpolationMode As D2D1_INTERPOLATION_MODE, compositeMode As D2D1_COMPOSITE_MODE)
        <PreserveSig>
        Overloads Sub DrawGdiMetafile(gdiMetafile As ID2D1GdiMetafile, ByRef targetOffset As D2D1_POINT_2F)
        <PreserveSig>
        Overloads Sub DrawBitmap(bitmap As ID2D1Bitmap, ByRef destinationRectangle As D2D1_RECT_F, opacity As Single, interpolationMode As D2D1_INTERPOLATION_MODE, ByRef sourceRectangle As D2D1_RECT_F, perspectiveTransform As D2D1_MATRIX_4X4_F)
        <PreserveSig>
        Overloads Sub PushLayer(ByRef layerParameters As D2D1_LAYER_PARAMETERS1, layer As ID2D1Layer)
        <PreserveSig>
        Overloads Function InvalidateEffectInputRectangle(effect As ID2D1Effect, input As UInteger, inputRectangle As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Function GetEffectInvalidRectangleCount(effect As ID2D1Effect, <Out> ByRef rectangleCount As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function GetEffectInvalidRectangles(effect As ID2D1Effect, <Out> ByRef rectangles As IntPtr, rectanglesCount As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function GetEffectRequiredInputRectangles(renderEffect As ID2D1Effect, renderImageRectangle As D2D1_RECT_F, inputDescriptions As D2D1_EFFECT_INPUT_DESCRIPTION, <Out> ByRef requiredInputRects As IntPtr, inputCount As UInteger) As HRESULT
        <PreserveSig>
        Overloads Sub FillOpacityMask(opacityMask As ID2D1Bitmap, brush As ID2D1Brush, ByRef destinationRectangle As D2D1_RECT_F, sourceRectangle As D2D1_RECT_F)
#End Region

        <PreserveSig>
        Overloads Function CreateFilledGeometryRealization(geometry As ID2D1Geometry, flatteningTolerance As Single, <Out> ByRef geometryRealization As ID2D1GeometryRealization) As HRESULT
        <PreserveSig>
        Overloads Function CreateStrokedGeometryRealization(geometry As ID2D1Geometry, flatteningTolerance As Single, strokeWidth As Single, strokeStyle As ID2D1StrokeStyle, <Out> ByRef geometryRealization As ID2D1GeometryRealization) As HRESULT
        <PreserveSig>
        Overloads Sub DrawGeometryRealization(geometryRealization As ID2D1GeometryRealization, brush As ID2D1Brush)
#End Region

        <PreserveSig>
        Overloads Function CreateInk(startPoint As D2D1_INK_POINT, <Out> ByRef ink As ID2D1Ink) As HRESULT
        <PreserveSig>
        Overloads Function CreateInkStyle(inkStyleProperties As D2D1_INK_STYLE_PROPERTIES, <Out> ByRef inkStyle As ID2D1InkStyle) As HRESULT
        <PreserveSig>
        Overloads Function CreateGradientMesh(patches As D2D1_GRADIENT_MESH_PATCH, patchesCount As UInteger, <Out> ByRef gradientMesh As ID2D1GradientMesh) As HRESULT
        <PreserveSig>
        Overloads Function CreateImageSourceFromWic(wicBitmapSource As IWICBitmapSource, loadingOptions As D2D1_IMAGE_SOURCE_LOADING_OPTIONS, alphaMode As D2D1_ALPHA_MODE, <Out> ByRef imageSource As ID2D1ImageSourceFromWic) As HRESULT
        <PreserveSig>
        Overloads Function CreateLookupTable3D(precision As D2D1_BUFFER_PRECISION, extents As UInteger, data As IntPtr, dataCount As UInteger, strides As UInteger, <Out> ByRef lookupTable As ID2D1LookupTable3D) As HRESULT
        <PreserveSig>
        Overloads Function CreateImageSourceFromDxgi(surfaces As IntPtr, surfaceCount As UInteger, colorSpace As DXGI_COLOR_SPACE_TYPE, options As D2D1_IMAGE_SOURCE_FROM_DXGI_OPTIONS, <Out> ByRef imageSource As ID2D1ImageSource) As HRESULT
        <PreserveSig>
        Overloads Function GetGradientMeshWorldBounds(gradientMesh As ID2D1GradientMesh, <Out> ByRef pBounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Sub DrawInk(ink As ID2D1Ink, brush As ID2D1Brush, inkStyle As ID2D1InkStyle)
        <PreserveSig>
        Overloads Sub DrawGradientMesh(gradientMesh As ID2D1GradientMesh)
        <PreserveSig>
        Overloads Sub DrawGdiMetafile(gdiMetafile As ID2D1GdiMetafile, ByRef destinationRectangle As D2D1_RECT_F, ByRef sourceRectangle As D2D1_RECT_F)
        <PreserveSig>
        Overloads Function CreateTransformedImageSource(imageSource As ID2D1ImageSource, properties As D2D1_TRANSFORMED_IMAGE_SOURCE_PROPERTIES, <Out> ByRef transformedImageSource As ID2D1TransformedImageSource) As HRESULT

#End Region

        <PreserveSig>
        Overloads Function CreateSpriteBatch(<Out> ByRef spriteBatch As ID2D1SpriteBatch) As HRESULT
        <PreserveSig>
        Overloads Sub DrawSpriteBatch(spriteBatch As ID2D1SpriteBatch, startIndex As UInteger, spriteCount As UInteger, bitmap As ID2D1Bitmap, Optional interpolationMode As D2D1_BITMAP_INTERPOLATION_MODE = D2D1_BITMAP_INTERPOLATION_MODE.D2D1_BITMAP_INTERPOLATION_MODE_LINEAR, Optional spriteOptions As D2D1_SPRITE_OPTIONS = D2D1_SPRITE_OPTIONS.D2D1_SPRITE_OPTIONS_NONE)

#End Region

        <PreserveSig>
        Overloads Function CreateSvgGlyphStyle(<Out> ByRef svgGlyphStyle As ID2D1SvgGlyphStyle) As HRESULT
        <PreserveSig>
        Overloads Sub DrawText(sString As String, stringLength As UInteger, textFormat As IDWriteTextFormat, ByRef layoutRect As D2D1_RECT_F, defaultFillBrush As ID2D1Brush, svgGlyphStyle As ID2D1SvgGlyphStyle, Optional colorPaletteIndex As UInteger = 0, Optional options As D2D1_DRAW_TEXT_OPTIONS = D2D1_DRAW_TEXT_OPTIONS.D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT, Optional measuringMode As DWRITE_MEASURING_MODE = DWRITE_MEASURING_MODE.DWRITE_MEASURING_MODE_NATURAL)
        <PreserveSig>
        Overloads Sub DrawTextLayout(origin As D2D1_POINT_2F, textLayout As IDWriteTextLayout, defaultFillBrush As ID2D1Brush, svgGlyphStyle As ID2D1SvgGlyphStyle, Optional colorPaletteIndex As UInteger = 0, Optional options As D2D1_DRAW_TEXT_OPTIONS = D2D1_DRAW_TEXT_OPTIONS.D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT)
        <PreserveSig>
        Overloads Sub DrawColorBitmapGlyphRun(glyphImageFormat As DWRITE_GLYPH_IMAGE_FORMATS, baselineOrigin As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, Optional measuringMode As DWRITE_MEASURING_MODE = DWRITE_MEASURING_MODE.DWRITE_MEASURING_MODE_NATURAL, Optional bitmapSnapOption As D2D1_COLOR_BITMAP_GLYPH_SNAP_OPTION = D2D1_COLOR_BITMAP_GLYPH_SNAP_OPTION.D2D1_COLOR_BITMAP_GLYPH_SNAP_OPTION_DEFAULT)
        <PreserveSig>
        Overloads Sub DrawSvgGlyphRun(baselineOrigin As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, Optional defaultFillBrush As ID2D1Brush = Nothing, Optional svgGlyphStyle As ID2D1SvgGlyphStyle = Nothing, Optional colorPaletteIndex As UInteger = 0, Optional measuringMode As DWRITE_MEASURING_MODE = DWRITE_MEASURING_MODE.DWRITE_MEASURING_MODE_NATURAL)
        <PreserveSig>
        Overloads Function GetColorBitmapGlyphImage(glyphImageFormat As DWRITE_GLYPH_IMAGE_FORMATS, glyphOrigin As D2D1_POINT_2F, fontFace As IDWriteFontFace, fontEmSize As Single, glyphIndex As UShort, isSideways As Boolean, worldTransform As D2D1_MATRIX_3X2_F, dpiX As Single, dpiY As Single, <Out> ByRef glyphTransform As D2D1_MATRIX_3X2_F, <Out> ByRef glyphImage As ID2D1Image) As HRESULT
        <PreserveSig>
        Overloads Function GetSvgGlyphImage(ByRef glyphOrigin As D2D1_POINT_2F, fontFace As IDWriteFontFace, fontEmSize As Single, glyphIndex As UShort, isSideways As Boolean, worldTransform As D2D1_MATRIX_3X2_F, defaultFillBrush As ID2D1Brush, svgGlyphStyle As ID2D1SvgGlyphStyle, colorPaletteIndex As UInteger, <Out> ByRef glyphTransform As D2D1_MATRIX_3X2_F, <Out> ByRef glyphImage As ID2D1CommandList) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateSvgDocument(inputXmlStream As ComTypes.IStream, viewportSize As D2D1_SIZE_F, <Out> ByRef svgDocument As ID2D1SvgDocument) As HRESULT
        <PreserveSig>
        Overloads Sub DrawSvgDocument(svgDocument As ID2D1SvgDocument)
        <PreserveSig>
        Overloads Function CreateColorContextFromDxgiColorSpace(colorSpace As DXGI_COLOR_SPACE_TYPE, <Out> ByRef colorContext As ID2D1ColorContext1) As HRESULT
        <PreserveSig>
        Overloads Function CreateColorContextFromSimpleColorProfile(ByRef simpleProfile As D2D1_SIMPLE_COLOR_PROFILE, <Out> ByRef colorContext As ID2D1ColorContext1) As HRESULT
#End Region

#Region "<ID2D1DeviceContext6>"
        <PreserveSig>
        Overloads Sub BlendImage(image As ID2D1Image, blendMode As D2D1_BLEND_MODE, ByRef targetOffset As D2D1_POINT_2F, ByRef imageRectangle As D2D1_RECT_F, Optional interpolationMode As D2D1_INTERPOLATION_MODE = D2D1_INTERPOLATION_MODE.D2D1_INTERPOLATION_MODE_LINEAR)
#End Region

        <PreserveSig>
        Function GetPaintFeatureLevel() As DWRITE_PAINT_FEATURE_LEVEL

        ''' <summary>
        ''' Draws a color glyph run that has the format of
        ''' DWRITE_GLYPH_IMAGE_FORMATS_COLR_PAINT_TREE.
        ''' </summary>
        ''' <param name="colorPaletteIndex"> The index used to select a color palette within
        ''' a color font. Note that this not the same as the paletteIndex in the
        ''' DWRITE_COLOR_GLYPH_RUN struct, which is not relevant for paint glyphs.</param>
        <PreserveSig>
        Sub DrawPaintGlyphRun(baselineOrigin As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, Optional defaultFillBrush As ID2D1Brush = Nothing, Optional colorPaletteIndex As UInteger = 0, Optional measuringMode As DWRITE_MEASURING_MODE = DWRITE_MEASURING_MODE.DWRITE_MEASURING_MODE_NATURAL)

        ''' <summary>
        ''' Draws a glyph run, using color representations of glyphs if available.
        ''' </summary>
        <PreserveSig>
        Sub DrawGlyphRunWithColorSupport(baselineOrigin As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, glyphRunDescription As IntPtr, foregroundBrush As ID2D1Brush, svgGlyphStyle As ID2D1SvgGlyphStyle, Optional colorPaletteIndex As UInteger = 0, Optional measuringMode As DWRITE_MEASURING_MODE = DWRITE_MEASURING_MODE.DWRITE_MEASURING_MODE_NATURAL, Optional bitmapSnapOption As D2D1_COLOR_BITMAP_GLYPH_SNAP_OPTION = D2D1_COLOR_BITMAP_GLYPH_SNAP_OPTION.D2D1_COLOR_BITMAP_GLYPH_SNAP_OPTION_DEFAULT)
    End Interface

    <ComImport>
    <Guid("1ab42875-c57f-4be9-bd85-9cd78d6f55ee")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1ColorContext1
        Inherits ID2D1ColorContext
#Region "<ID2D1ColorContext>"

#Region "<ID2D1Resource>"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region

        <PreserveSig>
        Overloads Sub GetColorSpace(<Out> ByRef colorSpace As D2D1_COLOR_SPACE)
        <PreserveSig>
        Overloads Function GetProfileSize() As UInteger
        <PreserveSig>
        Overloads Function GetProfile(<Out> ByRef profile As IntPtr, profileSize As UInteger) As HRESULT
#End Region

        <PreserveSig>
        Function GetColorContextType() As D2D1_COLOR_CONTEXT_TYPE
        <PreserveSig>
        Function GetDXGIColorSpace() As DXGI_COLOR_SPACE_TYPE
        <PreserveSig>
        Function GetSimpleColorProfile(<Out> ByRef simpleProfile As D2D1_SIMPLE_COLOR_PROFILE) As HRESULT
    End Interface

    ''' <summary>
    ''' Specifies which way a color profile is defined.
    ''' </summary>
    Public Enum D2D1_COLOR_CONTEXT_TYPE
        D2D1_COLOR_CONTEXT_TYPE_ICC = 0
        D2D1_COLOR_CONTEXT_TYPE_SIMPLE = 1
        D2D1_COLOR_CONTEXT_TYPE_DXGI = 2
        D2D1_COLOR_CONTEXT_TYPE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    <ComImport>
    <Guid("86b88e4d-afa4-4d7b-88e4-68a51c4a0aec")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1SvgDocument
        Inherits ID2D1Resource
#Region "<ID2D1Resource>"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region

        <PreserveSig>
        Function SetViewportSize(viewportSize As D2D1_SIZE_F) As HRESULT
        <PreserveSig>
        Sub GetViewportSize(<Out> ByRef viewportSize As D2D1_SIZE_F)
        <PreserveSig>
        Function SetRoot(root As ID2D1SvgElement) As HRESULT
        <PreserveSig>
        Sub GetRoot(<Out> ByRef root As ID2D1SvgElement)
        <PreserveSig>
        Function FindElementById(id As String, <Out> ByRef svgElement As ID2D1SvgElement) As HRESULT
        <PreserveSig>
        Function Serialize(outputXmlStream As ComTypes.IStream, Optional subtree As ID2D1SvgElement = Nothing) As HRESULT
        <PreserveSig>
        Function Deserialize(inputXmlStream As ComTypes.IStream, <Out> ByRef subtree As ID2D1SvgElement) As HRESULT
        <PreserveSig>
        Function CreatePaint(paintType As D2D1_SVG_PAINT_TYPE, color As D2D1_COLOR_F, id As String, <Out> ByRef paint As ID2D1SvgPaint) As HRESULT
        <PreserveSig>
        Function CreateStrokeDashArray(dashes As D2D1_SVG_LENGTH, dashesCount As UInteger, <Out> ByRef strokeDashArray As ID2D1SvgStrokeDashArray) As HRESULT
        <PreserveSig>
        Function CreatePointCollection(ByRef points As D2D1_POINT_2F, pointsCount As UInteger, <Out> ByRef pointCollection As ID2D1SvgPointCollection) As HRESULT
        <PreserveSig>
        Function CreatePathData(segmentData As Single, segmentDataCount As UInteger, commands As D2D1_SVG_PATH_COMMAND, commandsCount As UInteger, <Out> ByRef pathData As ID2D1SvgPathData) As HRESULT
    End Interface

    <ComImport>
    <Guid("ac7b67a6-183e-49c1-a823-0ebe40b0db29")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1SvgElement
        Inherits ID2D1Resource
#Region "<ID2D1Resource>"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region

        <PreserveSig>
        Sub GetDocument(<Out> ByRef document As ID2D1SvgDocument)
        <PreserveSig>
        Function GetTagName(<Out> ByRef name As String, nameCount As UInteger) As HRESULT
        <PreserveSig>
        Function GetTagNameLength() As UInteger
        <PreserveSig>
        Function IsTextContent() As Boolean
        <PreserveSig>
        Sub GetParent(<Out> ByRef parent As ID2D1SvgElement)
        <PreserveSig>
        Function HasChildren() As Boolean
        <PreserveSig>
        Sub GetFirstChild(<Out> ByRef child As ID2D1SvgElement)
        <PreserveSig>
        Sub GetLastChild(<Out> ByRef child As ID2D1SvgElement)
        <PreserveSig>
        Function GetPreviousChild(referenceChild As ID2D1SvgElement, <Out> ByRef previousChild As ID2D1SvgElement) As HRESULT
        <PreserveSig>
        Function GetNextChild(referenceChild As ID2D1SvgElement, <Out> ByRef nextChild As ID2D1SvgElement) As HRESULT
        <PreserveSig>
        Function InsertChildBefore(newChild As ID2D1SvgElement, Optional referenceChild As ID2D1SvgElement = Nothing) As HRESULT
        <PreserveSig>
        Function AppendChild(newChild As ID2D1SvgElement) As HRESULT
        <PreserveSig>
        Function ReplaceChild(newChild As ID2D1SvgElement, oldChild As ID2D1SvgElement) As HRESULT
        <PreserveSig>
        Function RemoveChild(oldChild As ID2D1SvgElement) As HRESULT
        <PreserveSig>
        Function CreateChild(tagName As String, <Out> ByRef newChild As ID2D1SvgElement) As HRESULT
        <PreserveSig>
        Function IsAttributeSpecified(name As String, <Out> ByRef inherited As Boolean) As Boolean
        <PreserveSig>
        Function GetSpecifiedAttributeCount() As UInteger
        <PreserveSig>
        Function GetSpecifiedAttributeName(index As UInteger, <Out> ByRef name As String, nameCount As UInteger, <Out> ByRef inherited As Boolean) As HRESULT
        <PreserveSig>
        Function GetSpecifiedAttributeNameLength(index As UInteger, <Out> ByRef nameLength As UInteger, <Out> ByRef inherited As Boolean) As HRESULT
        <PreserveSig>
        Function RemoveAttribute(name As String) As HRESULT
        <PreserveSig>
        Function SetTextValue(name As String, nameCount As UInteger) As HRESULT
        <PreserveSig>
        Function GetTextValue(<Out> ByRef name As String, nameCount As UInteger) As HRESULT
        <PreserveSig>
        Function GetTextValueLength() As UInteger
        <PreserveSig>
        Function SetAttributeValue(name As String, type As D2D1_SVG_ATTRIBUTE_STRING_TYPE, value As String) As HRESULT
        <PreserveSig>
        Function GetAttributeValue(
        <MarshalAs(UnmanagedType.LPWStr)> name As String, type As D2D1_SVG_ATTRIBUTE_STRING_TYPE, value As IntPtr, valueCount As UInteger) As HRESULT
        <PreserveSig>
        Function GetAttributeValueLength(name As String, type As D2D1_SVG_ATTRIBUTE_STRING_TYPE, <Out> ByRef valueLength As UInteger) As HRESULT
        <PreserveSig>
        Function SetAttributeValue(name As String, type As D2D1_SVG_ATTRIBUTE_POD_TYPE, value As IntPtr, valueSizeInBytes As UInteger) As HRESULT
        <PreserveSig>
        Function GetAttributeValue(name As String, type As D2D1_SVG_ATTRIBUTE_POD_TYPE, <Out> ByRef value As IntPtr, valueSizeInBytes As UInteger) As HRESULT
        <PreserveSig>
        Function SetAttributeValue(name As String, value As ID2D1SvgAttribute) As HRESULT
        <PreserveSig>
        Function GetAttributeValue(name As String, ByRef riid As Guid, <Out> ByRef value As IntPtr) As HRESULT
    End Interface

    <ComImport>
    <Guid("c9cdb0dd-f8c9-4e70-b7c2-301c80292c5e")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1SvgAttribute
        Inherits ID2D1Resource
#Region "<ID2D1Resource>"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region

        <PreserveSig>
        Sub GetElement(<Out> ByRef element As ID2D1SvgElement)
        <PreserveSig>
        Function Clone(<Out> ByRef attribute As ID2D1SvgAttribute) As HRESULT
    End Interface

    <ComImport>
    <Guid("d59bab0a-68a2-455b-a5dc-9eb2854e2490")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1SvgPaint
        Inherits ID2D1SvgAttribute
#Region "<ID2D1SvgAttribute>"
#Region "<ID2D1Resource>"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region

        <PreserveSig>
        Overloads Sub GetElement(<Out> ByRef element As ID2D1SvgElement)
        <PreserveSig>
        Overloads Function Clone(<Out> ByRef attribute As ID2D1SvgAttribute) As HRESULT
#End Region

        <PreserveSig>
        Function SetPaintType(paintType As D2D1_SVG_PAINT_TYPE) As HRESULT
        <PreserveSig>
        Function GetPaintType() As D2D1_SVG_PAINT_TYPE
        <PreserveSig>
        Function SetColor(color As D2D1_COLOR_F) As HRESULT
        <PreserveSig>
        Sub GetColor(<Out> ByRef color As D2D1_COLOR_F)
        <PreserveSig>
        Function SetId(id As String) As HRESULT
        <PreserveSig>
        Function GetId(<Out> ByRef id As String, idCount As UInteger) As HRESULT
        <PreserveSig>
        Function GetIdLength() As UInteger
    End Interface

    <ComImport>
    <Guid("f1c0ca52-92a3-4f00-b4ce-f35691efd9d9")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1SvgStrokeDashArray
        Inherits ID2D1SvgAttribute
#Region "<ID2D1SvgAttribute>"
#Region "<ID2D1Resource>"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region

        <PreserveSig>
        Overloads Sub GetElement(<Out> ByRef element As ID2D1SvgElement)
        <PreserveSig>
        Overloads Function Clone(<Out> ByRef attribute As ID2D1SvgAttribute) As HRESULT
#End Region

        <PreserveSig>
        Function RemoveDashesAtEnd(dashesCount As UInteger) As HRESULT
        <PreserveSig>
        Function UpdateDashes(dashes As Single, dashesCount As UInteger, Optional startIndex As UInteger = 0) As HRESULT
        <PreserveSig>
        Function UpdateDashes(dashes As D2D1_SVG_LENGTH, dashesCount As UInteger, Optional startIndex As UInteger = 0) As HRESULT
        <PreserveSig>
        Function GetDashes(<Out> ByRef dashes As Single, dashesCount As UInteger, Optional startIndex As UInteger = 0) As HRESULT
        <PreserveSig>
        Function GetDashes(<Out> ByRef dashes As D2D1_SVG_LENGTH, dashesCount As UInteger, Optional startIndex As UInteger = 0) As HRESULT
        <PreserveSig>
        Function GetDashesCount() As UInteger
    End Interface

    <ComImport>
    <Guid("9dbe4c0d-3572-4dd9-9825-5530813bb712")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1SvgPointCollection
        Inherits ID2D1SvgAttribute
#Region "<ID2D1SvgAttribute>"
#Region "<ID2D1Resource>"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region

        <PreserveSig>
        Overloads Sub GetElement(<Out> ByRef element As ID2D1SvgElement)
        <PreserveSig>
        Overloads Function Clone(<Out> ByRef attribute As ID2D1SvgAttribute) As HRESULT
#End Region

        <PreserveSig>
        Function RemovePointsAtEnd(pointsCount As UInteger) As HRESULT
        <PreserveSig>
        Function UpdatePoints(ByRef points As D2D1_POINT_2F, pointsCount As UInteger, Optional startIndex As UInteger = 0) As HRESULT
        <PreserveSig>
        Function GetPoints(<Out> ByRef points As D2D1_POINT_2F, pointsCount As UInteger, Optional startIndex As UInteger = 0) As HRESULT
        <PreserveSig>
        Function GetPointsCount() As UInteger
    End Interface

    <ComImport>
    <Guid("c095e4f4-bb98-43d6-9745-4d1b84ec9888")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1SvgPathData
        Inherits ID2D1SvgAttribute
#Region "<ID2D1SvgAttribute>"
#Region "<ID2D1Resource>"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region

        <PreserveSig>
        Overloads Sub GetElement(<Out> ByRef element As ID2D1SvgElement)
        <PreserveSig>
        Overloads Function Clone(<Out> ByRef attribute As ID2D1SvgAttribute) As HRESULT
#End Region

        <PreserveSig>
        Function RemoveSegmentDataAtEnd(dataCount As UInteger) As HRESULT
        <PreserveSig>
        Function UpdateSegmentData(data As Single, dataCount As UInteger, Optional startIndex As UInteger = 0) As HRESULT
        <PreserveSig>
        Function GetSegmentData(<Out> ByRef data As Single, dataCount As UInteger, Optional startIndex As UInteger = 0) As HRESULT
        <PreserveSig>
        Function GetSegmentDataCount() As UInteger
        <PreserveSig>
        Function RemoveCommandsAtEnd(commandsCount As UInteger) As HRESULT
        <PreserveSig>
        Function UpdateCommands(commands As D2D1_SVG_PATH_COMMAND, commandsCount As UInteger, Optional startIndex As UInteger = 0) As HRESULT
        <PreserveSig>
        Function GetCommands(<Out> ByRef commands As D2D1_SVG_PATH_COMMAND, commandsCount As UInteger, Optional startIndex As UInteger = 0) As HRESULT
        <PreserveSig>
        Function GetCommandsCount() As UInteger
        <PreserveSig>
        Function CreatePathGeometry(fillMode As D2D1_FILL_MODE, <Out> ByRef pathGeometry As ID2D1PathGeometry1) As HRESULT
    End Interface

    <ComImport>
    <Guid("62baa2d2-ab54-41b7-b872-787e0106a421")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1PathGeometry1
        Inherits ID2D1PathGeometry
#Region "<ID2D1PathGeometry>"

#Region "ID2D1Geometry"
#Region "ID2D1Resource"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region

        <PreserveSig>
        Overloads Function GetBounds(worldTransform As D2D1_MATRIX_3X2_F, <Out> ByRef bounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Function GetWidenedBounds(strokeWidth As Single, strokeStyle As ID2D1StrokeStyle, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef bounds As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Overloads Function StrokeContainsPoint(ByRef point As D2D1_POINT_2F, strokeWidth As Single, strokeStyle As ID2D1StrokeStyle, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef contains As Boolean) As HRESULT
        <PreserveSig>
        Overloads Function FillContainsPoint(ByRef point As D2D1_POINT_2F, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef contains As Boolean) As HRESULT
        <PreserveSig>
        Overloads Function CompareWithGeometry(inputGeometry As ID2D1Geometry, inputGeometryTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef relation As D2D1_GEOMETRY_RELATION) As HRESULT
        <PreserveSig>
        Overloads Function Simplify(simplificationOption As D2D1_GEOMETRY_SIMPLIFICATION_OPTION, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, geometrySink As ID2D1SimplifiedGeometrySink) As HRESULT
        <PreserveSig>
        Overloads Function Tessellate(worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, tessellationSink As ID2D1TessellationSink) As HRESULT
        <PreserveSig>
        Overloads Function CombineWithGeometry(inputGeometry As ID2D1Geometry, combineMode As D2D1_COMBINE_MODE, inputGeometryTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, geometrySink As ID2D1SimplifiedGeometrySink) As HRESULT
        <PreserveSig>
        Overloads Function Outline(worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, geometrySink As ID2D1SimplifiedGeometrySink) As HRESULT
        <PreserveSig>
        Overloads Function ComputeArea(worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef area As Single) As HRESULT
        <PreserveSig>
        Overloads Function ComputeLength(worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef length As Single) As HRESULT
        <PreserveSig>
        Overloads Function ComputePointAtLength(length As Single, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef point As D2D1_POINT_2F, <Out> ByRef unitTangentVector As D2D1_POINT_2F) As HRESULT
        <PreserveSig>
        Overloads Function Widen(strokeWidth As Single, strokeStyle As ID2D1StrokeStyle, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, geometrySink As ID2D1SimplifiedGeometrySink) As HRESULT

#End Region

        <PreserveSig>
        Overloads Function Open(<Out> ByRef geometrySink As ID2D1GeometrySink) As HRESULT
        <PreserveSig>
        Overloads Function Stream(geometrySink As ID2D1GeometrySink) As HRESULT
        <PreserveSig>
        Overloads Function GetSegmentCount(<Out> ByRef count As Integer) As HRESULT
        <PreserveSig>
        Overloads Function GetFigureCount(<Out> ByRef count As Integer) As HRESULT
#End Region

        <PreserveSig>
        Function ComputePointAndSegmentAtLength(length As Single, startSegment As UInteger, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, <Out> ByRef pointDescription As D2D1_POINT_DESCRIPTION) As HRESULT
    End Interface

    ''' <summary>
    ''' Describes a point along a path.
    ''' </summary>
    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_POINT_DESCRIPTION
        Public point As D2D1_POINT_2F
        Public unitTangentVector As D2D1_POINT_2F
        Public endSegment As UInteger
        Public endFigure As UInteger
        Public lengthToEndSegment As Single
    End Structure

    ''' <summary>
    ''' Simple description of a color space.
    ''' </summary>
    ''' 
    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_SIMPLE_COLOR_PROFILE
        ''' <summary>
        ''' The XY coordinates of the red primary in CIEXYZ space.
        ''' </summary>
        Public redPrimary As D2D1_POINT_2F

        ''' <summary>
        ''' The XY coordinates of the green primary in CIEXYZ space.
        ''' </summary>
        Public greenPrimary As D2D1_POINT_2F

        ''' <summary>
        ''' The XY coordinates of the blue primary in CIEXYZ space.
        ''' </summary>
        Public bluePrimary As D2D1_POINT_2F

        ''' <summary>
        ''' The X/Z tristimulus values for the whitepoint, normalized for relative
        ''' luminance.
        ''' </summary>
        Public whitePointXZ As D2D1_POINT_2F

        ''' <summary>
        ''' The gamma encoding to use for this color space.
        ''' </summary>
        Public gamma As D2D1_GAMMA1
    End Structure

    ''' <summary>
    ''' This determines what gamma is used for interpolation/blending.
    ''' </summary>
    Public Enum D2D1_GAMMA1
        ''' <summary>
        ''' Colors are manipulated in 2.2 gamma color space.
        ''' </summary>
        D2D1_GAMMA1_G22 = D2D1_GAMMA.D2D1_GAMMA_2_2

        ''' <summary>
        ''' Colors are manipulated in 1.0 gamma color space.
        ''' </summary>
        D2D1_GAMMA1_G10 = D2D1_GAMMA.D2D1_GAMMA_1_0

        ''' <summary>
        ''' Colors are manipulated in ST.2084 PQ gamma color space.
        ''' </summary>
        D2D1_GAMMA1_G2084 = 2
        D2D1_GAMMA1_FORCE_DWORD = &HFFFFFFFF
    End Enum

    <ComImport>
    <Guid("b499923b-7029-478f-a8b3-432c7c5f5312")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1Ink
        Inherits ID2D1Resource
#Region "<ID2D1Resource>"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region

        <PreserveSig>
        Sub SetStartPoint(startPoint As D2D1_INK_POINT)
        <PreserveSig>
        Sub GetStartPoint(<Out> ByRef startPoint As D2D1_INK_POINT)
        <PreserveSig>
        Function AddSegments(segments As D2D1_INK_BEZIER_SEGMENT, segmentsCount As UInteger) As HRESULT
        <PreserveSig>
        Function RemoveSegmentsAtEnd(segmentsCount As UInteger) As HRESULT
        <PreserveSig>
        Function SetSegments(startSegment As UInteger, segments As D2D1_INK_BEZIER_SEGMENT, segmentsCount As UInteger) As HRESULT
        <PreserveSig>
        Function SetSegmentAtEnd(segment As D2D1_INK_BEZIER_SEGMENT) As HRESULT
        <PreserveSig>
        Function GetSegmentCount() As UInteger
        <PreserveSig>
        Function GetSegments(startSegment As UInteger, <Out> ByRef segments As D2D1_INK_BEZIER_SEGMENT, segmentsCount As UInteger) As HRESULT
        <PreserveSig>
        Function StreamAsGeometry(inkStyle As ID2D1InkStyle, worldTransform As D2D1_MATRIX_3X2_F, flatteningTolerance As Single, geometrySink As ID2D1SimplifiedGeometrySink) As HRESULT
        <PreserveSig>
        Function GetBounds(inkStyle As ID2D1InkStyle, worldTransform As D2D1_MATRIX_3X2_F, <Out> ByRef bounds As D2D1_RECT_F) As HRESULT
    End Interface

    <ComImport>
    <Guid("bae8b344-23fc-4071-8cb5-d05d6f073848")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1InkStyle
        Inherits ID2D1Resource
#Region "<ID2D1Resource>"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region

        <PreserveSig>
        Sub SetNibTransform(transform As D2D1_MATRIX_3X2_F)
        <PreserveSig>
        Sub GetNibTransform(<Out> ByRef transform As D2D1_MATRIX_3X2_F)
        <PreserveSig>
        Sub SetNibShape(nibShape As D2D1_INK_NIB_SHAPE)
        <PreserveSig>
        Sub GetNibShape(<Out> ByRef nibShape As D2D1_INK_NIB_SHAPE)
    End Interface

    <ComImport>
    <Guid("f292e401-c050-4cde-83d7-04962d3b23c2")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1GradientMesh
        Inherits ID2D1Resource
#Region "<ID2D1Resource>"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region

        <PreserveSig>
        Function GetPatchCount() As UInteger
        <PreserveSig>
        Function GetPatches(startIndex As UInteger, <Out> ByRef patches As D2D1_GRADIENT_MESH_PATCH, patchesCount As UInteger) As HRESULT
    End Interface

    <ComImport>
    <Guid("c9b664e5-74a1-4378-9ac2-eefc37a3f4d8")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1ImageSource
        Inherits ID2D1Image
#Region "<ID2D1Image>"
#Region "<ID2D1Resource>"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)

#End Region
#End Region

        <PreserveSig>
        Function OfferResources() As HRESULT
        <PreserveSig>
        Function TryReclaimResources(<Out> ByRef resourcesDiscarded As Boolean) As HRESULT
    End Interface

    <ComImport>
    <Guid("7f1f79e5-2796-416c-8f55-700f911445e5")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1TransformedImageSource
        Inherits ID2D1Image
#Region "<ID2D1Image>"
#Region "<ID2D1Resource>"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)

#End Region
#End Region

        <PreserveSig>
        Sub GetSource(<Out> ByRef imageSource As ID2D1ImageSource)
        <PreserveSig>
        Sub GetProperties(<Out> ByRef properties As D2D1_TRANSFORMED_IMAGE_SOURCE_PROPERTIES)
    End Interface

    <ComImport>
    <Guid("77395441-1c8f-4555-8683-f50dab0fe792")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1ImageSourceFromWic
        Inherits ID2D1ImageSource
#Region "<ID2D1ImageSource>"

#Region "<ID2D1Image>"
#Region "<ID2D1Resource>"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)

#End Region
#End Region

        <PreserveSig>
        Overloads Function OfferResources() As HRESULT
        <PreserveSig>
        Overloads Function TryReclaimResources(<Out> ByRef resourcesDiscarded As Boolean) As HRESULT
#End Region

        <PreserveSig>
        Function EnsureCached(rectangleToFill As D2D1_RECT_U) As HRESULT
        <PreserveSig>
        Function TrimCache(rectangleToPreserve As D2D1_RECT_U) As HRESULT
        <PreserveSig>
        Sub GetSource(<Out> ByRef wicBitmapSource As IWICBitmapSource)
    End Interface

    <ComImport>
    <Guid("53dd9855-a3b0-4d5b-82e1-26e25c5e5797")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1LookupTable3D
        Inherits ID2D1Resource
#Region "<ID2D1Resource>"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region
    End Interface

    ''' <summary>
    '''  Represents a Bezier segment to be used in the creation of an ID2D1Ink object.
    ''' This structure differs from D2D1_BEZIER_SEGMENT in that it is composed of
    ''' D2D1_INK_POINT s, which contain a radius in addition to x- and y-coordinates.
    ''' </summary>
    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_INK_BEZIER_SEGMENT
        Public point1 As D2D1_INK_POINT
        Public point2 As D2D1_INK_POINT
        Public point3 As D2D1_INK_POINT
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_INK_POINT
        Public x As Single
        Public y As Single
        Public radius As Single
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_INK_STYLE_PROPERTIES
        ''' <summary>
        ''' The general shape of the nib used to draw a given ink object.
        ''' </summary>
        Public nibShape As D2D1_INK_NIB_SHAPE

        ''' <summary>
        ''' The transform applied to shape of the nib. _31 and _32 are ignored.
        ''' </summary>
        Public nibTransform As D2D1_MATRIX_3X2_F
    End Structure

    ''' <summary>
    ''' Specifies the appearance of the ink nib (pen tip) as part of an
    ''' D2D1_INK_STYLE_PROPERTIES structure.
    ''' </summary>
    Public Enum D2D1_INK_NIB_SHAPE
        D2D1_INK_NIB_SHAPE_ROUND = 0
        D2D1_INK_NIB_SHAPE_SQUARE = 1
        D2D1_INK_NIB_SHAPE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' Represents a tensor patch with 16 control points, 4 corner colors, and boundary
    ''' flags. An ID2D1GradientMesh is made up of 1 or more gradient mesh patches. Use
    ''' the GradientMeshPatch function or the GradientMeshPatchFromCoonsPatch function
    ''' to create one.
    ''' </summary>
    ''' 
    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_GRADIENT_MESH_PATCH

        ''' <summary>
        ''' The gradient mesh patch control point at position 00.
        ''' </summary>
        Public point00 As D2D1_POINT_2F

        ''' <summary>
        ''' The gradient mesh patch control point at position 01.
        ''' </summary>
        Public point01 As D2D1_POINT_2F

        ''' <summary>
        ''' The gradient mesh patch control point at position 02.
        ''' </summary>
        Public point02 As D2D1_POINT_2F

        ''' <summary>
        ''' The gradient mesh patch control point at position 03.
        ''' </summary>
        Public point03 As D2D1_POINT_2F

        ''' <summary>
        ''' The gradient mesh patch control point at position 10.
        ''' </summary>
        Public point10 As D2D1_POINT_2F

        ''' <summary>
        ''' The gradient mesh patch control point at position 11.
        ''' </summary>
        Public point11 As D2D1_POINT_2F

        ''' <summary>
        ''' The gradient mesh patch control point at position 12.
        ''' </summary>
        Public point12 As D2D1_POINT_2F

        ''' <summary>
        ''' The gradient mesh patch control point at position 13.
        ''' </summary>
        Public point13 As D2D1_POINT_2F

        ''' <summary>
        ''' The gradient mesh patch control point at position 20.
        ''' </summary>
        Public point20 As D2D1_POINT_2F

        ''' <summary>
        ''' The gradient mesh patch control point at position 21.
        ''' </summary>
        Public point21 As D2D1_POINT_2F

        ''' <summary>
        ''' The gradient mesh patch control point at position 22.
        ''' </summary>
        Public point22 As D2D1_POINT_2F

        ''' <summary>
        ''' The gradient mesh patch control point at position 23.
        ''' </summary>
        Public point23 As D2D1_POINT_2F

        ''' <summary>
        ''' The gradient mesh patch control point at position 30.
        ''' </summary>
        Public point30 As D2D1_POINT_2F

        ''' <summary>
        ''' The gradient mesh patch control point at position 31.
        ''' </summary>
        Public point31 As D2D1_POINT_2F

        ''' <summary>
        ''' The gradient mesh patch control point at position 32.
        ''' </summary>
        Public point32 As D2D1_POINT_2F

        ''' <summary>
        ''' The gradient mesh patch control point at position 33.
        ''' </summary>
        Public point33 As D2D1_POINT_2F

        ''' <summary>
        ''' The color associated with control point at position 00.
        ''' </summary>
        Public color00 As D2D1_COLOR_F

        ''' <summary>
        ''' The color associated with control point at position 03.
        ''' </summary>
        Public color03 As D2D1_COLOR_F

        ''' <summary>
        ''' The color associated with control point at position 30.
        ''' </summary>
        Public color30 As D2D1_COLOR_F

        ''' <summary>
        ''' The color associated with control point at position 33.
        ''' </summary>
        Public color33 As D2D1_COLOR_F

        ''' <summary>
        ''' The edge mode for the top edge of the patch.
        ''' </summary>
        Public topEdgeMode As D2D1_PATCH_EDGE_MODE

        ''' <summary>
        ''' The edge mode for the left edge of the patch.
        ''' </summary>
        Public leftEdgeMode As D2D1_PATCH_EDGE_MODE

        ''' <summary>
        ''' The edge mode for the bottom edge of the patch.
        ''' </summary>
        Public bottomEdgeMode As D2D1_PATCH_EDGE_MODE

        ''' <summary>
        ''' The edge mode for the right edge of the patch.
        ''' </summary>
        Public rightEdgeMode As D2D1_PATCH_EDGE_MODE
    End Structure

    ''' <summary>
    ''' Specifies how to render gradient mesh edges.
    ''' </summary>
    Public Enum D2D1_PATCH_EDGE_MODE
        ''' <summary>
        ''' Render this edge aliased.
        ''' </summary>
        D2D1_PATCH_EDGE_MODE_ALIASED = 0

        ''' <summary>
        ''' Render this edge antialiased.
        ''' </summary>
        D2D1_PATCH_EDGE_MODE_ANTIALIASED = 1

        ''' <summary>
        ''' Render this edge aliased and inflated out slightly.
        ''' </summary>
        D2D1_PATCH_EDGE_MODE_ALIASED_INFLATED = 2
        D2D1_PATCH_EDGE_MODE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' Option flags controlling how images sources are loaded during
    ''' CreateImageSourceFromWic.
    ''' </summary>
    Public Enum D2D1_IMAGE_SOURCE_LOADING_OPTIONS
        D2D1_IMAGE_SOURCE_LOADING_OPTIONS_NONE = 0
        D2D1_IMAGE_SOURCE_LOADING_OPTIONS_RELEASE_SOURCE = 1
        D2D1_IMAGE_SOURCE_LOADING_OPTIONS_CACHE_ON_DEMAND = 2
        D2D1_IMAGE_SOURCE_LOADING_OPTIONS_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' Option flags controlling primary conversion performed by
    ''' CreateImageSourceFromDxgi, if any.
    ''' </summary>
    Public Enum D2D1_IMAGE_SOURCE_FROM_DXGI_OPTIONS
        D2D1_IMAGE_SOURCE_FROM_DXGI_OPTIONS_NONE = 0
        D2D1_IMAGE_SOURCE_FROM_DXGI_OPTIONS_LOW_QUALITY_PRIMARY_CONVERSION = 1
        D2D1_IMAGE_SOURCE_FROM_DXGI_OPTIONS_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' Properties of a transformed image source.
    ''' </summary>
    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_TRANSFORMED_IMAGE_SOURCE_PROPERTIES
        ''' <summary>
        ''' The orientation at which the image source is drawn.
        ''' </summary>
        Public orientation As D2D1_ORIENTATION

        ''' <summary>
        ''' The horizontal scale factor at which the image source is drawn.
        ''' </summary>
        Public scaleX As Single

        ''' <summary>
        ''' The vertical scale factor at which the image source is drawn.
        ''' </summary>
        Public scaleY As Single

        ''' <summary>
        ''' The interpolation mode used when the image source is drawn.  This is ignored if
        ''' the image source is drawn using the DrawImage method, or using an image brush.
        ''' </summary>
        Public interpolationMode As D2D1_INTERPOLATION_MODE

        ''' <summary>
        ''' Option flags.
        ''' </summary>
        Public options As D2D1_TRANSFORMED_IMAGE_SOURCE_OPTIONS
    End Structure

    ''' <summary>
    ''' Specifies the orientation of an image.
    ''' </summary>
    Public Enum D2D1_ORIENTATION
        D2D1_ORIENTATION_DEFAULT = 1
        D2D1_ORIENTATION_FLIP_HORIZONTAL = 2
        D2D1_ORIENTATION_ROTATE_CLOCKWISE180 = 3
        D2D1_ORIENTATION_ROTATE_CLOCKWISE180_FLIP_HORIZONTAL = 4
        D2D1_ORIENTATION_ROTATE_CLOCKWISE90_FLIP_HORIZONTAL = 5
        D2D1_ORIENTATION_ROTATE_CLOCKWISE270 = 6
        D2D1_ORIENTATION_ROTATE_CLOCKWISE270_FLIP_HORIZONTAL = 7
        D2D1_ORIENTATION_ROTATE_CLOCKWISE90 = 8
        D2D1_ORIENTATION_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' Option flags for transformed image sources.
    ''' </summary>
    Public Enum D2D1_TRANSFORMED_IMAGE_SOURCE_OPTIONS
        D2D1_TRANSFORMED_IMAGE_SOURCE_OPTIONS_NONE = 0

        ''' <summary>
        ''' Prevents the image source from being automatically scaled (by a ratio of the
        ''' context DPI divided by 96) while drawn.
        ''' </summary>
        D2D1_TRANSFORMED_IMAGE_SOURCE_OPTIONS_DISABLE_DPI_SCALE = 1
        D2D1_TRANSFORMED_IMAGE_SOURCE_OPTIONS_FORCE_DWORD = &HFFFFFFFF
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_EFFECT_INPUT_DESCRIPTION
        ''' <summary>
        ''' The effect whose input connection is being specified.
        ''' </summary>
        Public effect As ID2D1Effect

        ''' <summary>
        ''' The index of the input connection into the specified effect.
        ''' </summary>
        Private inputIndex As UInteger

        ''' <summary>
        ''' The rectangle which would be available on the specified input connection during
        ''' render operations.
        ''' </summary>
        Private inputRectangle As D2D1_RECT_F
    End Structure

    <ComImport>
    <Guid("2c1d867d-c290-41c8-ae7e-34a98702e9a5")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1PrintControl
        Function AddPage(commandList As ID2D1CommandList, pageSize As D2D_SIZE_F, pagePrintTicketStream As ComTypes.IStream, <Out> ByRef tag1 As ULong, <Out> ByRef tag2 As ULong) As HRESULT
        Function Close() As HRESULT
    End Interface

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D_SIZE_F
        Public width As Single
        Public height As Single

        Public Sub New(width As Single, height As Single)
            Me.width = width
            Me.height = height
        End Sub
    End Structure

    <ComImport>
    <Guid("31e6e7bc-e0ff-4d46-8c64-a0a8c41c15d3")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1Multithread
        <PreserveSig>
        Function GetMultithreadProtected() As Boolean
        <PreserveSig>
        Sub Enter()
        <PreserveSig>
        Sub Leave()
    End Interface

    <ComImport>
    <Guid("47dd575d-ac05-4cdd-8049-9b02cd16f44c")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1Device
        Inherits ID2D1Resource
#Region "<ID2D1Resource>"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region

        <PreserveSig>
        Function CreateDeviceContext(options As D2D1_DEVICE_CONTEXT_OPTIONS, <Out> ByRef deviceContext As ID2D1DeviceContext) As HRESULT
        <PreserveSig>
        Function CreatePrintControl(wicFactory As IWICImagingFactory, documentTarget As IntPtr, ByRef printControlProperties As D2D1_PRINT_CONTROL_PROPERTIES, <Out> ByRef printControl As ID2D1PrintControl) As HRESULT
        <PreserveSig>
        Sub SetMaximumTextureMemory(maximumInBytes As ULong)
        <PreserveSig>
        Function GetMaximumTextureMemory() As ULong
        <PreserveSig>
        Sub ClearResources(Optional millisecondsSinceUse As UInteger = 0)
    End Interface

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_PRINT_CONTROL_PROPERTIES
        Public fontSubset As D2D1_PRINT_FONT_SUBSET_MODE

        ''' <summary>
        ''' DPI for rasterization of all unsupported D2D commands or options, defaults to
        ''' 150.0
        ''' </summary>
        Public rasterDPI As Single

        ''' <summary>
        ''' Color space for vector graphics in XPS package
        ''' </summary>
        Public colorSpace As D2D1_COLOR_SPACE
    End Structure

    Public Enum D2D1_PRINT_FONT_SUBSET_MODE
        ''' <summary>
        ''' Subset for used glyphs, send and discard font resource after every five pages
        ''' </summary>
        D2D1_PRINT_FONT_SUBSET_MODE_DEFAULT = 0

        ''' <summary>
        ''' Subset for used glyphs, send and discard font resource after each page
        ''' </summary>
        D2D1_PRINT_FONT_SUBSET_MODE_EACHPAGE = 1

        ''' <summary>
        ''' Do not subset, reuse font for all pages, send it after first page
        ''' </summary>
        D2D1_PRINT_FONT_SUBSET_MODE_NONE = 2
        D2D1_PRINT_FONT_SUBSET_MODE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_DEVICE_CONTEXT_OPTIONS
        D2D1_DEVICE_CONTEXT_OPTIONS_NONE = 0
        ''' <summary>
        ''' Geometry rendering will be performed on many threads in parallel, a single
        ''' thread is the default.
        ''' </summary>
        D2D1_DEVICE_CONTEXT_OPTIONS_ENABLE_MULTITHREADED_OPTIMIZATIONS = 1
        D2D1_DEVICE_CONTEXT_OPTIONS_FORCE_DWORD = &HFFFFFFFF
    End Enum

    <ComImport>
    <Guid("54d7898a-a061-40a7-bec7-e465bcba2c4f")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1CommandSink
        <PreserveSig>
        Function BeginDraw() As HRESULT
        <PreserveSig>
        Function EndDraw() As HRESULT
        <PreserveSig>
        Function SetAntialiasMode(antialiasMode As D2D1_ANTIALIAS_MODE) As HRESULT
        <PreserveSig>
        Function SetTags(tag1 As ULong, tag2 As ULong) As HRESULT
        <PreserveSig>
        Function SetTextAntialiasMode(textAntialiasMode As D2D1_TEXT_ANTIALIAS_MODE) As HRESULT
        <PreserveSig>
        Function SetTextRenderingParams(textRenderingParams As IDWriteRenderingParams) As HRESULT
        <PreserveSig>
        Function SetTransform(transform As D2D1_MATRIX_3X2_F) As HRESULT
        <PreserveSig>
        Function SetPrimitiveBlend(primitiveBlend As D2D1_PRIMITIVE_BLEND) As HRESULT
        <PreserveSig>
        Function SetUnitMode(unitMode As D2D1_UNIT_MODE) As HRESULT
        <PreserveSig>
        Function Clear(color As D2D1_COLOR_F) As HRESULT
        <PreserveSig>
        Function DrawGlyphRun(baselineOrigin As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, glyphRunDescription As IntPtr, foregroundBrush As ID2D1Brush, measuringMode As DWRITE_MEASURING_MODE) As HRESULT
        <PreserveSig>
        Function DrawLine(point0 As D2D1_POINT_2F, point1 As D2D1_POINT_2F, brush As ID2D1Brush, strokeWidth As Single, strokeStyle As ID2D1StrokeStyle) As HRESULT
        <PreserveSig>
        Function DrawGeometry(geometry As ID2D1Geometry, brush As ID2D1Brush, strokeWidth As Single, strokeStyle As ID2D1StrokeStyle) As HRESULT
        <PreserveSig>
        Function DrawRectangle(rect As D2D1_RECT_F, brush As ID2D1Brush, strokeWidth As Single, strokeStyle As ID2D1StrokeStyle) As HRESULT
        <PreserveSig>
        Function DrawBitmap(bitmap As ID2D1Bitmap, destinationRectangle As D2D1_RECT_F, opacity As Single, interpolationMode As D2D1_INTERPOLATION_MODE, sourceRectangle As D2D1_RECT_F, perspectiveTransform As D2D1_MATRIX_4X4_F) As HRESULT
        <PreserveSig>
        Function DrawImage(image As ID2D1Image, ByRef targetOffset As D2D1_POINT_2F, imageRectangle As D2D1_RECT_F, interpolationMode As D2D1_INTERPOLATION_MODE, compositeMode As D2D1_COMPOSITE_MODE) As HRESULT
        <PreserveSig>
        Function DrawGdiMetafile(gdiMetafile As ID2D1GdiMetafile, ByRef targetOffset As D2D1_POINT_2F) As HRESULT
        <PreserveSig>
        Function FillMesh(mesh As ID2D1Mesh, brush As ID2D1Brush) As HRESULT
        <PreserveSig>
        Function FillOpacityMask(opacityMask As ID2D1Bitmap, brush As ID2D1Brush, destinationRectangle As D2D1_RECT_F, sourceRectangle As D2D1_RECT_F) As HRESULT
        <PreserveSig>
        Function FillGeometry(geometry As ID2D1Geometry, brush As ID2D1Brush, opacityBrush As ID2D1Brush) As HRESULT
        <PreserveSig>
        Function FillRectangle(rect As D2D1_RECT_F, brush As ID2D1Brush) As HRESULT
        <PreserveSig>
        Function PushAxisAlignedClip(clipRect As D2D1_RECT_F, antialiasMode As D2D1_ANTIALIAS_MODE) As HRESULT
        <PreserveSig>
        Function PushLayer(ByRef layerParameters1 As D2D1_LAYER_PARAMETERS1, layer As ID2D1Layer) As HRESULT
        <PreserveSig>
        Function PopAxisAlignedClip() As HRESULT
        <PreserveSig>
        Function PopLayer() As HRESULT
    End Interface

    <ComImport>
    <Guid("82237326-8111-4f7c-bcf4-b5c1175564fe")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1GdiMetafileSink
        <PreserveSig>
        Function ProcessRecord(recordType As UInteger, recordData As IntPtr, recordDataSize As UInteger) As HRESULT
    End Interface

    <ComImport>
    <Guid("fd0ecb6b-91e6-411e-8655-395e760f91b4")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1GdiMetafileSink1
        Inherits ID2D1GdiMetafileSink
#Region "<ID2D1GdiMetafileSink>"
        <PreserveSig>
        Overloads Function ProcessRecord(recordType As UInteger, recordData As IntPtr, recordDataSize As UInteger) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function ProcessRecord(recordType As UInteger, recordData As IntPtr, recordDataSize As UInteger, flags As UInteger) As HRESULT
    End Interface

    <ComImport>
    <Guid("2f543dc3-cfc1-4211-864f-cfd91c6f3395")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1GdiMetafile
        Inherits ID2D1Resource
#Region "<ID2D1Resource>"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region

        <PreserveSig>
        Function Stream(sink As ID2D1GdiMetafileSink) As HRESULT
        <PreserveSig>
        Function GetBounds(<Out> ByRef bounds As D2D1_RECT_F) As HRESULT
    End Interface

    <ComImport>
    <Guid("2e69f9e8-dd3f-4bf9-95ba-c04f49d788df")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1GdiMetafile1
        Inherits ID2D1GdiMetafile
#Region "<ID2D1GdiMetafile>"
#Region "<ID2D1Resource>"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region

        <PreserveSig>
        Overloads Function Stream(sink As ID2D1GdiMetafileSink) As HRESULT
        <PreserveSig>
        Overloads Function GetBounds(<Out> ByRef bounds As D2D1_RECT_F) As HRESULT
#End Region

        <PreserveSig>
        Function GetDpi(<Out> ByRef dpiX As Single, <Out> ByRef dpiY As Single) As HRESULT
        <PreserveSig>
        Function GetSourceBounds(<Out> ByRef bounds As D2D1_RECT_F) As HRESULT
    End Interface

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_LAYER_PARAMETERS1
        Public contentBounds As D2D1_RECT_F
        Public geometricMask As ID2D1Geometry
        Public maskAntialiasMode As D2D1_ANTIALIAS_MODE
        Public maskTransform As D2D1_MATRIX_3X2_F
        Public opacity As Single
        Public opacityBrush As ID2D1Brush
        Public layerOptions As D2D1_LAYER_OPTIONS1
    End Structure

    Public Enum D2D1_LAYER_OPTIONS1
        D2D1_LAYER_OPTIONS1_NONE = 0
        D2D1_LAYER_OPTIONS1_INITIALIZE_FROM_BACKGROUND = 1
        D2D1_LAYER_OPTIONS1_IGNORE_ALPHA = 2
        D2D1_LAYER_OPTIONS1_FORCE_DWORD = &HFFFFFFFF
    End Enum

    <ComImport>
    <Guid("b4f34a19-2383-4d76-94f6-ec343657c3dc")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1CommandList
        Inherits ID2D1Image
#Region "<ID2D1Image>"
#Region "<ID2D1Resource>"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region
#End Region

        <PreserveSig>
        Function Stream(sink As ID2D1CommandSink) As HRESULT
        <PreserveSig>
        Function Close() As HRESULT
    End Interface

    <ComImport>
    <Guid("41343a53-e41a-49a2-91cd-21793bbb62e5")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1BitmapBrush1
        Inherits ID2D1BitmapBrush
#Region "<ID2D1BitmapBrush>"

#Region "ID2D1Brush"
        <PreserveSig>
        Overloads Sub SetOpacity(opacity As Single)
        <PreserveSig>
        Overloads Sub SetTransform(transform As D2D1_MATRIX_3X2_F)
        <PreserveSig>
        Overloads Function GetOpacity() As Single
        <PreserveSig>
        Overloads Sub GetTransform(<Out> ByRef transform As D2D1_MATRIX_3X2_F_STRUCT)
#End Region
        <PreserveSig>
        Overloads Sub SetExtendModeX(extendModeX As D2D1_EXTEND_MODE)
        <PreserveSig>
        Overloads Sub SetExtendModeY(extendModeY As D2D1_EXTEND_MODE)
        <PreserveSig>
        Overloads Sub SetInterpolationMode(interpolationMode As D2D1_BITMAP_INTERPOLATION_MODE)
        <PreserveSig>
        Overloads Sub SetBitmap(bitmap As ID2D1Bitmap)
        <PreserveSig>
        Overloads Sub GetExtendModeX(<Out> ByRef extendedModeX As D2D1_EXTEND_MODE)
        <PreserveSig>
        Overloads Sub GetExtendModeY(<Out> ByRef extendedModeY As D2D1_EXTEND_MODE)
        <PreserveSig>
        Overloads Sub GetInterpolationMode(<Out> ByRef interpolationMode As D2D1_BITMAP_INTERPOLATION_MODE)
        <PreserveSig>
        Overloads Sub GetBitmap(<Out> ByRef bitmap As ID2D1Bitmap)
#End Region

        <PreserveSig>
        Sub SetInterpolationMode1(interpolationMode As D2D1_INTERPOLATION_MODE)
        <PreserveSig>
        Sub GetInterpolationMode1(<Out> ByRef interpolationMode1 As D2D1_INTERPOLATION_MODE)
    End Interface

    <ComImport>
    <Guid("fe9e984d-3f95-407c-b5db-cb94d4e8f87c")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1ImageBrush
        Inherits ID2D1Brush
#Region "ID2D1Brush"
#Region "ID2D1Resource"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region
        <PreserveSig>
        Overloads Sub SetOpacity(opacity As Single)
        <PreserveSig>
        Overloads Sub SetTransform(transform As D2D1_MATRIX_3X2_F)
        <PreserveSig>
        Overloads Function GetOpacity() As Single
        <PreserveSig>
        Overloads Sub GetTransform(<Out> ByRef transform As D2D1_MATRIX_3X2_F_STRUCT)
#End Region

        <PreserveSig>
        Sub SetImage(image As ID2D1Image)
        <PreserveSig>
        Sub SetExtendModeX(extendModeX As D2D1_EXTEND_MODE)
        <PreserveSig>
        Sub SetExtendModeY(extendModeY As D2D1_EXTEND_MODE)
        <PreserveSig>
        Sub SetInterpolationMode(interpolationMode As D2D1_INTERPOLATION_MODE)
        <PreserveSig>
        Sub SetSourceRectangle(sourceRectangle As D2D1_RECT_F)
        <PreserveSig>
        Sub GetImage(<Out> ByRef image As ID2D1Image)
        <PreserveSig>
        Sub GetExtendModeX(<Out> ByRef extendedModeX As D2D1_EXTEND_MODE)
        <PreserveSig>
        Sub GetExtendModeY(<Out> ByRef extendedModeY As D2D1_EXTEND_MODE)
        <PreserveSig>
        Sub GetInterpolationMode(<Out> ByRef interpolationMode As D2D1_BITMAP_INTERPOLATION_MODE)
        <PreserveSig>
        Sub GetSourceRectangle(<Out> ByRef sourceRectangle As D2D1_RECT_F)
    End Interface

    <ComImport>
    <Guid("ae1572f4-5dd0-4777-998b-9279472ae63b")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1GradientStopCollection1
        Inherits ID2D1GradientStopCollection
#Region "<ID2D1GradientStopCollection>"
#Region "<ID2D1Resource>"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region
        <PreserveSig>
        Overloads Function GetGradientStopCount() As UInteger
        <PreserveSig>
        Overloads Sub GetGradientStops(<Out> ByRef gradientStops As D2D1_GRADIENT_STOP, gradientStopsCount As UInteger)
        <PreserveSig>
        Overloads Sub GetColorInterpolationGamma(<Out> ByRef colorInterpolationGamma As D2D1_GAMMA)
        <PreserveSig>
        Overloads Sub GetExtendMode(<Out> ByRef extendedMode As D2D1_EXTEND_MODE)
#End Region

        Sub GetGradientStops1(<Out> ByRef gradientStops As D2D1_GRADIENT_STOP, gradientStopsCount As UInteger)
        <PreserveSig>
        Sub GetPreInterpolationSpace(<Out> ByRef preInterpolationSpace As D2D1_COLOR_SPACE)
        <PreserveSig>
        Sub GetPostInterpolationSpace(<Out> ByRef postInterpolationSpace As D2D1_COLOR_SPACE)
        <PreserveSig>
        Sub GetBufferPrecision(<Out> ByRef bufferPrecision As D2D1_BUFFER_PRECISION)
        <PreserveSig>
        Sub GetColorInterpolationMode(<Out> ByRef colorInterpolationMode As D2D1_COLOR_INTERPOLATION_MODE)
    End Interface


    ''' <summary>
    ''' Specifies how the Crop effect handles the crop rectangle falling on fractional
    ''' pixel coordinates.
    ''' </summary>
    Public Enum D2D1_BORDER_MODE
        D2D1_BORDER_MODE_SOFT = 0
        D2D1_BORDER_MODE_HARD = 1
        D2D1_BORDER_MODE_FORCE_DWORD = &HFFFFFFFF
    End Enum


    ''' <summary>
    ''' Specifies the color channel the Displacement map effect extracts the intensity
    ''' from and uses it to spatially displace the image in the X or Y direction.
    ''' </summary>
    Public Enum D2D1_CHANNEL_SELECTOR
        D2D1_CHANNEL_SELECTOR_R = 0
        D2D1_CHANNEL_SELECTOR_G = 1
        D2D1_CHANNEL_SELECTOR_B = 2
        D2D1_CHANNEL_SELECTOR_A = 3
        D2D1_CHANNEL_SELECTOR_FORCE_DWORD = &HFFFFFFFF
    End Enum


    ''' <summary>
    ''' Speficies whether a flip and/or rotation operation should be performed by the
    ''' Bitmap source effect
    ''' </summary>
    Public Enum D2D1_BITMAPSOURCE_ORIENTATION
        D2D1_BITMAPSOURCE_ORIENTATION_DEFAULT = 1
        D2D1_BITMAPSOURCE_ORIENTATION_FLIP_HORIZONTAL = 2
        D2D1_BITMAPSOURCE_ORIENTATION_ROTATE_CLOCKWISE180 = 3
        D2D1_BITMAPSOURCE_ORIENTATION_ROTATE_CLOCKWISE180_FLIP_HORIZONTAL = 4
        D2D1_BITMAPSOURCE_ORIENTATION_ROTATE_CLOCKWISE270_FLIP_HORIZONTAL = 5
        D2D1_BITMAPSOURCE_ORIENTATION_ROTATE_CLOCKWISE90 = 6
        D2D1_BITMAPSOURCE_ORIENTATION_ROTATE_CLOCKWISE90_FLIP_HORIZONTAL = 7
        D2D1_BITMAPSOURCE_ORIENTATION_ROTATE_CLOCKWISE270 = 8
        D2D1_BITMAPSOURCE_ORIENTATION_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Gaussian Blur effect's top level properties.
    ''' Effect description: Applies a gaussian blur to a bitmap with the specified blur
    ''' radius and angle.
    ''' </summary>
    Public Enum D2D1_GAUSSIANBLUR_PROP
        ''' <summary>
        ''' Property Name: "StandardDeviation"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_GAUSSIANBLUR_PROP_STANDARD_DEVIATION = 0

        ''' <summary>
        ''' Property Name: "Optimization"
        ''' Property Type: D2D1_GAUSSIANBLUR_OPTIMIZATION
        ''' </summary>
        D2D1_GAUSSIANBLUR_PROP_OPTIMIZATION = 1

        ''' <summary>
        ''' Property Name: "BorderMode"
        ''' Property Type: D2D1_BORDER_MODE
        ''' </summary>
        D2D1_GAUSSIANBLUR_PROP_BORDER_MODE = 2
        D2D1_GAUSSIANBLUR_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_GAUSSIANBLUR_OPTIMIZATION
        D2D1_GAUSSIANBLUR_OPTIMIZATION_SPEED = 0
        D2D1_GAUSSIANBLUR_OPTIMIZATION_BALANCED = 1
        D2D1_GAUSSIANBLUR_OPTIMIZATION_QUALITY = 2
        D2D1_GAUSSIANBLUR_OPTIMIZATION_FORCE_DWORD = &HFFFFFFFF
    End Enum


    ''' <summary>
    ''' The enumeration of the Directional Blur effect's top level properties.
    ''' Effect description: Applies a directional blur to a bitmap with the specified
    ''' blur radius and angle.
    ''' </summary>
    Public Enum D2D1_DIRECTIONALBLUR_PROP
        ''' <summary>
        ''' Property Name: "StandardDeviation"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_DIRECTIONALBLUR_PROP_STANDARD_DEVIATION = 0

        ''' <summary>
        ''' Property Name: "Angle"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_DIRECTIONALBLUR_PROP_ANGLE = 1

        ''' <summary>
        ''' Property Name: "Optimization"
        ''' Property Type: D2D1_DIRECTIONALBLUR_OPTIMIZATION
        ''' </summary>
        D2D1_DIRECTIONALBLUR_PROP_OPTIMIZATION = 2

        ''' <summary>
        ''' Property Name: "BorderMode"
        ''' Property Type: D2D1_BORDER_MODE
        ''' </summary>
        D2D1_DIRECTIONALBLUR_PROP_BORDER_MODE = 3
        D2D1_DIRECTIONALBLUR_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_DIRECTIONALBLUR_OPTIMIZATION
        D2D1_DIRECTIONALBLUR_OPTIMIZATION_SPEED = 0
        D2D1_DIRECTIONALBLUR_OPTIMIZATION_BALANCED = 1
        D2D1_DIRECTIONALBLUR_OPTIMIZATION_QUALITY = 2
        D2D1_DIRECTIONALBLUR_OPTIMIZATION_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Shadow effect's top level properties.
    ''' Effect description: Applies a shadow to a bitmap based on its alpha channel.
    ''' </summary>
    Public Enum D2D1_SHADOW_PROP
        ''' <summary>
        ''' Property Name: "BlurStandardDeviation"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_SHADOW_PROP_BLUR_STANDARD_DEVIATION = 0

        ''' <summary>
        ''' Property Name: "Color"
        ''' Property Type: D2D1_VECTOR_4F
        ''' </summary>
        D2D1_SHADOW_PROP_COLOR = 1

        ''' <summary>
        ''' Property Name: "Optimization"
        ''' Property Type: D2D1_SHADOW_OPTIMIZATION
        ''' </summary>
        D2D1_SHADOW_PROP_OPTIMIZATION = 2
        D2D1_SHADOW_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_SHADOW_OPTIMIZATION
        D2D1_SHADOW_OPTIMIZATION_SPEED = 0
        D2D1_SHADOW_OPTIMIZATION_BALANCED = 1
        D2D1_SHADOW_OPTIMIZATION_QUALITY = 2
        D2D1_SHADOW_OPTIMIZATION_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Blend effect's top level properties.
    ''' Effect description: Blends a foreground and background using a pre-defined blend
    ''' mode.
    ''' </summary>
    Public Enum D2D1_BLEND_PROP
        ''' <summary>
        ''' Property Name: "Mode"
        ''' Property Type: D2D1_BLEND_MODE
        ''' </summary>
        D2D1_BLEND_PROP_MODE = 0
        D2D1_BLEND_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_BLEND_MODE
        D2D1_BLEND_MODE_MULTIPLY = 0
        D2D1_BLEND_MODE_SCREEN = 1
        D2D1_BLEND_MODE_DARKEN = 2
        D2D1_BLEND_MODE_LIGHTEN = 3
        D2D1_BLEND_MODE_DISSOLVE = 4
        D2D1_BLEND_MODE_COLOR_BURN = 5
        D2D1_BLEND_MODE_LINEAR_BURN = 6
        D2D1_BLEND_MODE_DARKER_COLOR = 7
        D2D1_BLEND_MODE_LIGHTER_COLOR = 8
        D2D1_BLEND_MODE_COLOR_DODGE = 9
        D2D1_BLEND_MODE_LINEAR_DODGE = 10
        D2D1_BLEND_MODE_OVERLAY = 11
        D2D1_BLEND_MODE_SOFT_LIGHT = 12
        D2D1_BLEND_MODE_HARD_LIGHT = 13
        D2D1_BLEND_MODE_VIVID_LIGHT = 14
        D2D1_BLEND_MODE_LINEAR_LIGHT = 15
        D2D1_BLEND_MODE_PIN_LIGHT = 16
        D2D1_BLEND_MODE_HARD_MIX = 17
        D2D1_BLEND_MODE_DIFFERENCE = 18
        D2D1_BLEND_MODE_EXCLUSION = 19
        D2D1_BLEND_MODE_HUE = 20
        D2D1_BLEND_MODE_SATURATION = 21
        D2D1_BLEND_MODE_COLOR = 22
        D2D1_BLEND_MODE_LUMINOSITY = 23
        D2D1_BLEND_MODE_SUBTRACT = 24
        D2D1_BLEND_MODE_DIVISION = 25
        D2D1_BLEND_MODE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Saturation effect's top level properties.
    ''' Effect description: Alters the saturation of the bitmap based on the user
    ''' specified saturation value.
    ''' </summary>
    Public Enum D2D1_SATURATION_PROP
        ''' <summary>
        ''' Property Name: "Saturation"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_SATURATION_PROP_SATURATION = 0
        D2D1_SATURATION_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Hue Rotation effect's top level properties.
    ''' Effect description: Changes the Hue of a bitmap based on a user specified Hue
    ''' Rotation angle.
    ''' </summary>
    Public Enum D2D1_HUEROTATION_PROP
        ''' <summary>
        ''' Property Name: "Angle"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_HUEROTATION_PROP_ANGLE = 0
        D2D1_HUEROTATION_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum


    ''' <summary>
    ''' The enumeration of the Color Matrix effect's top level properties.
    ''' Effect description: Applies a user specified color matrix to each channel of the
    ''' input bitmap.
    ''' </summary>
    Public Enum D2D1_COLORMATRIX_PROP
        ''' <summary>
        ''' Property Name: "ColorMatrix"
        ''' Property Type: D2D1_MATRIX_5X4_F
        ''' </summary>
        D2D1_COLORMATRIX_PROP_COLOR_MATRIX = 0

        ''' <summary>
        ''' Property Name: "AlphaMode"
        ''' Property Type: D2D1_COLORMATRIX_ALPHA_MODE
        ''' </summary>
        D2D1_COLORMATRIX_PROP_ALPHA_MODE = 1

        ''' <summary>
        ''' Property Name: "ClampOutput"
        ''' Property Type: BOOL
        ''' </summary>
        D2D1_COLORMATRIX_PROP_CLAMP_OUTPUT = 2
        D2D1_COLORMATRIX_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_COLORMATRIX_ALPHA_MODE
        D2D1_COLORMATRIX_ALPHA_MODE_PREMULTIPLIED = 1
        D2D1_COLORMATRIX_ALPHA_MODE_STRAIGHT = 2
        D2D1_COLORMATRIX_ALPHA_MODE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Bitmap Source effect's top level properties.
    ''' Effect description: Provides an image source.
    ''' </summary>
    Public Enum D2D1_BITMAPSOURCE_PROP
        ''' <summary>
        ''' Property Name: "WicBitmapSource"
        ''' Property Type: IUnknown *
        ''' </summary>
        D2D1_BITMAPSOURCE_PROP_WIC_BITMAP_SOURCE = 0

        ''' <summary>
        ''' Property Name: "Scale"
        ''' Property Type: D2D1_VECTOR_2F
        ''' </summary>
        D2D1_BITMAPSOURCE_PROP_SCALE = 1

        ''' <summary>
        ''' Property Name: "InterpolationMode"
        ''' Property Type: D2D1_BITMAPSOURCE_INTERPOLATION_MODE
        ''' </summary>
        D2D1_BITMAPSOURCE_PROP_INTERPOLATION_MODE = 2

        ''' <summary>
        ''' Property Name: "EnableDPICorrection"
        ''' Property Type: BOOL
        ''' </summary>
        D2D1_BITMAPSOURCE_PROP_ENABLE_DPI_CORRECTION = 3

        ''' <summary>
        ''' Property Name: "AlphaMode"
        ''' Property Type: D2D1_BITMAPSOURCE_ALPHA_MODE
        ''' </summary>
        D2D1_BITMAPSOURCE_PROP_ALPHA_MODE = 4

        ''' <summary>
        ''' Property Name: "Orientation"
        ''' Property Type: D2D1_BITMAPSOURCE_ORIENTATION
        ''' </summary>
        D2D1_BITMAPSOURCE_PROP_ORIENTATION = 5
        D2D1_BITMAPSOURCE_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_BITMAPSOURCE_INTERPOLATION_MODE
        D2D1_BITMAPSOURCE_INTERPOLATION_MODE_NEAREST_NEIGHBOR = 0
        D2D1_BITMAPSOURCE_INTERPOLATION_MODE_LINEAR = 1
        D2D1_BITMAPSOURCE_INTERPOLATION_MODE_CUBIC = 2
        D2D1_BITMAPSOURCE_INTERPOLATION_MODE_FANT = 6
        D2D1_BITMAPSOURCE_INTERPOLATION_MODE_MIPMAP_LINEAR = 7
        D2D1_BITMAPSOURCE_INTERPOLATION_MODE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_BITMAPSOURCE_ALPHA_MODE
        D2D1_BITMAPSOURCE_ALPHA_MODE_PREMULTIPLIED = 1
        D2D1_BITMAPSOURCE_ALPHA_MODE_STRAIGHT = 2
        D2D1_BITMAPSOURCE_ALPHA_MODE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Composite effect's top level properties.
    ''' Effect description: Composites foreground and background images using the
    ''' selected composition mode.
    ''' </summary>
    Public Enum D2D1_COMPOSITE_PROP
        ''' <summary>
        ''' Property Name: "Mode"
        ''' Property Type: D2D1_COMPOSITE_MODE
        ''' </summary>
        D2D1_COMPOSITE_PROP_MODE = 0
        D2D1_COMPOSITE_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the 3D Transform effect's top level properties.
    ''' Effect description: Applies a 3D transform to a bitmap.
    ''' </summary>
    Public Enum D2D1_3DTRANSFORM_PROP
        ''' <summary>
        ''' Property Name: "InterpolationMode"
        ''' Property Type: D2D1_3DTRANSFORM_INTERPOLATION_MODE
        ''' </summary>
        D2D1_3DTRANSFORM_PROP_INTERPOLATION_MODE = 0

        ''' <summary>
        ''' Property Name: "BorderMode"
        ''' Property Type: D2D1_BORDER_MODE
        ''' </summary>
        D2D1_3DTRANSFORM_PROP_BORDER_MODE = 1

        ''' <summary>
        ''' Property Name: "TransformMatrix"
        ''' Property Type: D2D1_MATRIX_4X4_F
        ''' </summary>
        D2D1_3DTRANSFORM_PROP_TRANSFORM_MATRIX = 2
        D2D1_3DTRANSFORM_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_3DTRANSFORM_INTERPOLATION_MODE
        D2D1_3DTRANSFORM_INTERPOLATION_MODE_NEAREST_NEIGHBOR = 0
        D2D1_3DTRANSFORM_INTERPOLATION_MODE_LINEAR = 1
        D2D1_3DTRANSFORM_INTERPOLATION_MODE_CUBIC = 2
        D2D1_3DTRANSFORM_INTERPOLATION_MODE_MULTI_SAMPLE_LINEAR = 3
        D2D1_3DTRANSFORM_INTERPOLATION_MODE_ANISOTROPIC = 4
        D2D1_3DTRANSFORM_INTERPOLATION_MODE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the 3D Perspective Transform effect's top level properties.
    ''' Effect description: Applies a 3D perspective transform to a bitmap.
    ''' </summary>
    Public Enum D2D1_3DPERSPECTIVETRANSFORM_PROP
        ''' <summary>
        ''' Property Name: "InterpolationMode"
        ''' Property Type: D2D1_3DPERSPECTIVETRANSFORM_INTERPOLATION_MODE
        ''' </summary>
        D2D1_3DPERSPECTIVETRANSFORM_PROP_INTERPOLATION_MODE = 0

        ''' <summary>
        ''' Property Name: "BorderMode"
        ''' Property Type: D2D1_BORDER_MODE
        ''' </summary>
        D2D1_3DPERSPECTIVETRANSFORM_PROP_BORDER_MODE = 1

        ''' <summary>
        ''' Property Name: "Depth"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_3DPERSPECTIVETRANSFORM_PROP_DEPTH = 2

        ''' <summary>
        ''' Property Name: "PerspectiveOrigin"
        ''' Property Type: D2D1_VECTOR_2F
        ''' </summary>
        D2D1_3DPERSPECTIVETRANSFORM_PROP_PERSPECTIVE_ORIGIN = 3

        ''' <summary>
        ''' Property Name: "LocalOffset"
        ''' Property Type: D2D1_VECTOR_3F
        ''' </summary>
        D2D1_3DPERSPECTIVETRANSFORM_PROP_LOCAL_OFFSET = 4

        ''' <summary>
        ''' Property Name: "GlobalOffset"
        ''' Property Type: D2D1_VECTOR_3F
        ''' </summary>
        D2D1_3DPERSPECTIVETRANSFORM_PROP_GLOBAL_OFFSET = 5

        ''' <summary>
        ''' Property Name: "RotationOrigin"
        ''' Property Type: D2D1_VECTOR_3F
        ''' </summary>
        D2D1_3DPERSPECTIVETRANSFORM_PROP_ROTATION_ORIGIN = 6

        ''' <summary>
        ''' Property Name: "Rotation"
        ''' Property Type: D2D1_VECTOR_3F
        ''' </summary>
        D2D1_3DPERSPECTIVETRANSFORM_PROP_ROTATION = 7
        D2D1_3DPERSPECTIVETRANSFORM_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_3DPERSPECTIVETRANSFORM_INTERPOLATION_MODE
        D2D1_3DPERSPECTIVETRANSFORM_INTERPOLATION_MODE_NEAREST_NEIGHBOR = 0
        D2D1_3DPERSPECTIVETRANSFORM_INTERPOLATION_MODE_LINEAR = 1
        D2D1_3DPERSPECTIVETRANSFORM_INTERPOLATION_MODE_CUBIC = 2
        D2D1_3DPERSPECTIVETRANSFORM_INTERPOLATION_MODE_MULTI_SAMPLE_LINEAR = 3
        D2D1_3DPERSPECTIVETRANSFORM_INTERPOLATION_MODE_ANISOTROPIC = 4
        D2D1_3DPERSPECTIVETRANSFORM_INTERPOLATION_MODE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the 2D Affine Transform effect's top level properties.
    ''' Effect description: Applies a 2D affine transform to a bitmap.
    ''' </summary>
    Public Enum D2D1_2DAFFINETRANSFORM_PROP
        ''' <summary>
        ''' Property Name: "InterpolationMode"
        ''' Property Type: D2D1_2DAFFINETRANSFORM_INTERPOLATION_MODE
        ''' </summary>
        D2D1_2DAFFINETRANSFORM_PROP_INTERPOLATION_MODE = 0

        ''' <summary>
        ''' Property Name: "BorderMode"
        ''' Property Type: D2D1_BORDER_MODE
        ''' </summary>
        D2D1_2DAFFINETRANSFORM_PROP_BORDER_MODE = 1

        ''' <summary>
        ''' Property Name: "TransformMatrix"
        ''' Property Type: D2D1_MATRIX_3X2_F
        ''' </summary>
        D2D1_2DAFFINETRANSFORM_PROP_TRANSFORM_MATRIX = 2

        ''' <summary>
        ''' Property Name: "Sharpness"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_2DAFFINETRANSFORM_PROP_SHARPNESS = 3
        D2D1_2DAFFINETRANSFORM_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_2DAFFINETRANSFORM_INTERPOLATION_MODE
        D2D1_2DAFFINETRANSFORM_INTERPOLATION_MODE_NEAREST_NEIGHBOR = 0
        D2D1_2DAFFINETRANSFORM_INTERPOLATION_MODE_LINEAR = 1
        D2D1_2DAFFINETRANSFORM_INTERPOLATION_MODE_CUBIC = 2
        D2D1_2DAFFINETRANSFORM_INTERPOLATION_MODE_MULTI_SAMPLE_LINEAR = 3
        D2D1_2DAFFINETRANSFORM_INTERPOLATION_MODE_ANISOTROPIC = 4
        D2D1_2DAFFINETRANSFORM_INTERPOLATION_MODE_HIGH_QUALITY_CUBIC = 5
        D2D1_2DAFFINETRANSFORM_INTERPOLATION_MODE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the DPI Compensation effect's top level properties.
    ''' Effect description: Scales according to the input DPI and the current context
    ''' DPI
    ''' </summary>
    Public Enum D2D1_DPICOMPENSATION_PROP
        ''' <summary>
        ''' Property Name: "InterpolationMode"
        ''' Property Type: D2D1_DPICOMPENSATION_INTERPOLATION_MODE
        ''' </summary>
        D2D1_DPICOMPENSATION_PROP_INTERPOLATION_MODE = 0

        ''' <summary>
        ''' Property Name: "BorderMode"
        ''' Property Type: D2D1_BORDER_MODE
        ''' </summary>
        D2D1_DPICOMPENSATION_PROP_BORDER_MODE = 1

        ''' <summary>
        ''' Property Name: "InputDpi"
        ''' Property Type: D2D1_VECTOR_2F
        ''' </summary>
        D2D1_DPICOMPENSATION_PROP_INPUT_DPI = 2
        D2D1_DPICOMPENSATION_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_DPICOMPENSATION_INTERPOLATION_MODE
        D2D1_DPICOMPENSATION_INTERPOLATION_MODE_NEAREST_NEIGHBOR = 0
        D2D1_DPICOMPENSATION_INTERPOLATION_MODE_LINEAR = 1
        D2D1_DPICOMPENSATION_INTERPOLATION_MODE_CUBIC = 2
        D2D1_DPICOMPENSATION_INTERPOLATION_MODE_MULTI_SAMPLE_LINEAR = 3
        D2D1_DPICOMPENSATION_INTERPOLATION_MODE_ANISOTROPIC = 4
        D2D1_DPICOMPENSATION_INTERPOLATION_MODE_HIGH_QUALITY_CUBIC = 5
        D2D1_DPICOMPENSATION_INTERPOLATION_MODE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Scale effect's top level properties.
    ''' Effect description: Applies scaling operation to the bitmap.
    ''' </summary>
    Public Enum D2D1_SCALE_PROP
        ''' <summary>
        ''' Property Name: "Scale"
        ''' Property Type: D2D1_VECTOR_2F
        ''' </summary>
        D2D1_SCALE_PROP_SCALE = 0

        ''' <summary>
        ''' Property Name: "CenterPoint"
        ''' Property Type: D2D1_VECTOR_2F
        ''' </summary>
        D2D1_SCALE_PROP_CENTER_POINT = 1

        ''' <summary>
        ''' Property Name: "InterpolationMode"
        ''' Property Type: D2D1_SCALE_INTERPOLATION_MODE
        ''' </summary>
        D2D1_SCALE_PROP_INTERPOLATION_MODE = 2

        ''' <summary>
        ''' Property Name: "BorderMode"
        ''' Property Type: D2D1_BORDER_MODE
        ''' </summary>
        D2D1_SCALE_PROP_BORDER_MODE = 3

        ''' <summary>
        ''' Property Name: "Sharpness"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_SCALE_PROP_SHARPNESS = 4
        D2D1_SCALE_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_SCALE_INTERPOLATION_MODE
        D2D1_SCALE_INTERPOLATION_MODE_NEAREST_NEIGHBOR = 0
        D2D1_SCALE_INTERPOLATION_MODE_LINEAR = 1
        D2D1_SCALE_INTERPOLATION_MODE_CUBIC = 2
        D2D1_SCALE_INTERPOLATION_MODE_MULTI_SAMPLE_LINEAR = 3
        D2D1_SCALE_INTERPOLATION_MODE_ANISOTROPIC = 4
        D2D1_SCALE_INTERPOLATION_MODE_HIGH_QUALITY_CUBIC = 5
        D2D1_SCALE_INTERPOLATION_MODE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Turbulence effect's top level properties.
    ''' Effect description: Generates a bitmap based on the Perlin noise turbulence
    ''' function.
    ''' </summary>
    Public Enum D2D1_TURBULENCE_PROP
        ''' <summary>
        ''' Property Name: "Offset"
        ''' Property Type: D2D1_VECTOR_2F
        ''' </summary>
        D2D1_TURBULENCE_PROP_OFFSET = 0

        ''' <summary>
        ''' Property Name: "Size"
        ''' Property Type: D2D1_VECTOR_2F
        ''' </summary>
        D2D1_TURBULENCE_PROP_SIZE = 1

        ''' <summary>
        ''' Property Name: "BaseFrequency"
        ''' Property Type: D2D1_VECTOR_2F
        ''' </summary>
        D2D1_TURBULENCE_PROP_BASE_FREQUENCY = 2

        ''' <summary>
        ''' Property Name: "NumOctaves"
        ''' Property Type: UINT32
        ''' </summary>
        D2D1_TURBULENCE_PROP_NUM_OCTAVES = 3

        ''' <summary>
        ''' Property Name: "Seed"
        ''' Property Type: INT32
        ''' </summary>
        D2D1_TURBULENCE_PROP_SEED = 4

        ''' <summary>
        ''' Property Name: "Noise"
        ''' Property Type: D2D1_TURBULENCE_NOISE
        ''' </summary>
        D2D1_TURBULENCE_PROP_NOISE = 5

        ''' <summary>
        ''' Property Name: "Stitchable"
        ''' Property Type: BOOL
        ''' </summary>
        D2D1_TURBULENCE_PROP_STITCHABLE = 6
        D2D1_TURBULENCE_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_TURBULENCE_NOISE
        D2D1_TURBULENCE_NOISE_FRACTAL_SUM = 0
        D2D1_TURBULENCE_NOISE_TURBULENCE = 1
        D2D1_TURBULENCE_NOISE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Displacement Map effect's top level properties.
    ''' Effect description: Displaces a bitmap based on user specified setting and
    ''' another bitmap.
    ''' </summary>
    Public Enum D2D1_DISPLACEMENTMAP_PROP
        ''' <summary>
        ''' Property Name: "Scale"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_DISPLACEMENTMAP_PROP_SCALE = 0

        ''' <summary>
        ''' Property Name: "XChannelSelect"
        ''' Property Type: D2D1_CHANNEL_SELECTOR
        ''' </summary>
        D2D1_DISPLACEMENTMAP_PROP_X_CHANNEL_SELECT = 1

        ''' <summary>
        ''' Property Name: "YChannelSelect"
        ''' Property Type: D2D1_CHANNEL_SELECTOR
        ''' </summary>
        D2D1_DISPLACEMENTMAP_PROP_Y_CHANNEL_SELECT = 2
        D2D1_DISPLACEMENTMAP_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Color Management effect's top level properties.
    ''' Effect description: Changes colors based on user provided color contexts.
    ''' </summary>
    Public Enum D2D1_COLORMANAGEMENT_PROP
        ''' <summary>
        ''' Property Name: "SourceColorContext"
        ''' Property Type: ID2D1ColorContext *
        ''' </summary>
        D2D1_COLORMANAGEMENT_PROP_SOURCE_COLOR_CONTEXT = 0

        ''' <summary>
        ''' Property Name: "SourceRenderingIntent"
        ''' Property Type: D2D1_RENDERING_INTENT
        ''' </summary>
        D2D1_COLORMANAGEMENT_PROP_SOURCE_RENDERING_INTENT = 1

        ''' <summary>
        ''' Property Name: "DestinationColorContext"
        ''' Property Type: ID2D1ColorContext *
        ''' </summary>
        D2D1_COLORMANAGEMENT_PROP_DESTINATION_COLOR_CONTEXT = 2

        ''' <summary>
        ''' Property Name: "DestinationRenderingIntent"
        ''' Property Type: D2D1_RENDERING_INTENT
        ''' </summary>
        D2D1_COLORMANAGEMENT_PROP_DESTINATION_RENDERING_INTENT = 3

        ''' <summary>
        ''' Property Name: "AlphaMode"
        ''' Property Type: D2D1_COLORMANAGEMENT_ALPHA_MODE
        ''' </summary>
        D2D1_COLORMANAGEMENT_PROP_ALPHA_MODE = 4

        ''' <summary>
        ''' Property Name: "Quality"
        ''' Property Type: D2D1_COLORMANAGEMENT_QUALITY
        ''' </summary>
        D2D1_COLORMANAGEMENT_PROP_QUALITY = 5
        D2D1_COLORMANAGEMENT_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_COLORMANAGEMENT_ALPHA_MODE
        D2D1_COLORMANAGEMENT_ALPHA_MODE_PREMULTIPLIED = 1
        D2D1_COLORMANAGEMENT_ALPHA_MODE_STRAIGHT = 2
        D2D1_COLORMANAGEMENT_ALPHA_MODE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_COLORMANAGEMENT_QUALITY
        D2D1_COLORMANAGEMENT_QUALITY_PROOF = 0
        D2D1_COLORMANAGEMENT_QUALITY_NORMAL = 1
        D2D1_COLORMANAGEMENT_QUALITY_BEST = 2
        D2D1_COLORMANAGEMENT_QUALITY_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' Specifies which ICC rendering intent the Color management effect should use.
    ''' </summary>
    Public Enum D2D1_COLORMANAGEMENT_RENDERING_INTENT
        D2D1_COLORMANAGEMENT_RENDERING_INTENT_PERCEPTUAL = 0
        D2D1_COLORMANAGEMENT_RENDERING_INTENT_RELATIVE_COLORIMETRIC = 1
        D2D1_COLORMANAGEMENT_RENDERING_INTENT_SATURATION = 2
        D2D1_COLORMANAGEMENT_RENDERING_INTENT_ABSOLUTE_COLORIMETRIC = 3
        D2D1_COLORMANAGEMENT_RENDERING_INTENT_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Histogram effect's top level properties.
    ''' Effect description: Computes the histogram of an image.
    ''' </summary>
    Public Enum D2D1_HISTOGRAM_PROP
        ''' <summary>
        ''' Property Name: "NumBins"
        ''' Property Type: UINT32
        ''' </summary>
        D2D1_HISTOGRAM_PROP_NUM_BINS = 0

        ''' <summary>
        ''' Property Name: "ChannelSelect"
        ''' Property Type: D2D1_CHANNEL_SELECTOR
        ''' </summary>
        D2D1_HISTOGRAM_PROP_CHANNEL_SELECT = 1

        ''' <summary>
        ''' Property Name: "HistogramOutput"
        ''' Property Type: (blob)
        ''' </summary>
        D2D1_HISTOGRAM_PROP_HISTOGRAM_OUTPUT = 2
        D2D1_HISTOGRAM_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Point-Specular effect's top level properties.
    ''' Effect description: Creates a specular lighting effect with a point light
    ''' source.
    ''' </summary>
    Public Enum D2D1_POINTSPECULAR_PROP
        ''' <summary>
        ''' Property Name: "LightPosition"
        ''' Property Type: D2D1_VECTOR_3F
        ''' </summary>
        D2D1_POINTSPECULAR_PROP_LIGHT_POSITION = 0

        ''' <summary>
        ''' Property Name: "SpecularExponent"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_POINTSPECULAR_PROP_SPECULAR_EXPONENT = 1

        ''' <summary>
        ''' Property Name: "SpecularConstant"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_POINTSPECULAR_PROP_SPECULAR_CONSTANT = 2

        ''' <summary>
        ''' Property Name: "SurfaceScale"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_POINTSPECULAR_PROP_SURFACE_SCALE = 3

        ''' <summary>
        ''' Property Name: "Color"
        ''' Property Type: D2D1_VECTOR_3F
        ''' </summary>
        D2D1_POINTSPECULAR_PROP_COLOR = 4

        ''' <summary>
        ''' Property Name: "KernelUnitLength"
        ''' Property Type: D2D1_VECTOR_2F
        ''' </summary>
        D2D1_POINTSPECULAR_PROP_KERNEL_UNIT_LENGTH = 5

        ''' <summary>
        ''' Property Name: "ScaleMode"
        ''' Property Type: D2D1_POINTSPECULAR_SCALE_MODE
        ''' </summary>
        D2D1_POINTSPECULAR_PROP_SCALE_MODE = 6
        D2D1_POINTSPECULAR_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_POINTSPECULAR_SCALE_MODE
        D2D1_POINTSPECULAR_SCALE_MODE_NEAREST_NEIGHBOR = 0
        D2D1_POINTSPECULAR_SCALE_MODE_LINEAR = 1
        D2D1_POINTSPECULAR_SCALE_MODE_CUBIC = 2
        D2D1_POINTSPECULAR_SCALE_MODE_MULTI_SAMPLE_LINEAR = 3
        D2D1_POINTSPECULAR_SCALE_MODE_ANISOTROPIC = 4
        D2D1_POINTSPECULAR_SCALE_MODE_HIGH_QUALITY_CUBIC = 5
        D2D1_POINTSPECULAR_SCALE_MODE_FORCE_DWORD = &HFFFFFFFF
    End Enum


    ''' <summary>
    ''' The enumeration of the Spot-Specular effect's top level properties.
    ''' Effect description: Creates a specular lighting effect with a spot light source.
    ''' </summary>
    Public Enum D2D1_SPOTSPECULAR_PROP
        ''' <summary>
        ''' Property Name: "LightPosition"
        ''' Property Type: D2D1_VECTOR_3F
        ''' </summary>
        D2D1_SPOTSPECULAR_PROP_LIGHT_POSITION = 0

        ''' <summary>
        ''' Property Name: "PointsAt"
        ''' Property Type: D2D1_VECTOR_3F
        ''' </summary>
        D2D1_SPOTSPECULAR_PROP_POINTS_AT = 1

        ''' <summary>
        ''' Property Name: "Focus"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_SPOTSPECULAR_PROP_FOCUS = 2

        ''' <summary>
        ''' Property Name: "LimitingConeAngle"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_SPOTSPECULAR_PROP_LIMITING_CONE_ANGLE = 3

        ''' <summary>
        ''' Property Name: "SpecularExponent"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_SPOTSPECULAR_PROP_SPECULAR_EXPONENT = 4

        ''' <summary>
        ''' Property Name: "SpecularConstant"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_SPOTSPECULAR_PROP_SPECULAR_CONSTANT = 5

        ''' <summary>
        ''' Property Name: "SurfaceScale"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_SPOTSPECULAR_PROP_SURFACE_SCALE = 6

        ''' <summary>
        ''' Property Name: "Color"
        ''' Property Type: D2D1_VECTOR_3F
        ''' </summary>
        D2D1_SPOTSPECULAR_PROP_COLOR = 7

        ''' <summary>
        ''' Property Name: "KernelUnitLength"
        ''' Property Type: D2D1_VECTOR_2F
        ''' </summary>
        D2D1_SPOTSPECULAR_PROP_KERNEL_UNIT_LENGTH = 8

        ''' <summary>
        ''' Property Name: "ScaleMode"
        ''' Property Type: D2D1_SPOTSPECULAR_SCALE_MODE
        ''' </summary>
        D2D1_SPOTSPECULAR_PROP_SCALE_MODE = 9
        D2D1_SPOTSPECULAR_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_SPOTSPECULAR_SCALE_MODE
        D2D1_SPOTSPECULAR_SCALE_MODE_NEAREST_NEIGHBOR = 0
        D2D1_SPOTSPECULAR_SCALE_MODE_LINEAR = 1
        D2D1_SPOTSPECULAR_SCALE_MODE_CUBIC = 2
        D2D1_SPOTSPECULAR_SCALE_MODE_MULTI_SAMPLE_LINEAR = 3
        D2D1_SPOTSPECULAR_SCALE_MODE_ANISOTROPIC = 4
        D2D1_SPOTSPECULAR_SCALE_MODE_HIGH_QUALITY_CUBIC = 5
        D2D1_SPOTSPECULAR_SCALE_MODE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Distant-Specular effect's top level properties.
    ''' Effect description: Creates a specular lighting effect with a distant light
    ''' source.
    ''' </summary>
    Public Enum D2D1_DISTANTSPECULAR_PROP
        ''' <summary>
        ''' Property Name: "Azimuth"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_DISTANTSPECULAR_PROP_AZIMUTH = 0

        ''' <summary>
        ''' Property Name: "Elevation"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_DISTANTSPECULAR_PROP_ELEVATION = 1

        ''' <summary>
        ''' Property Name: "SpecularExponent"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_DISTANTSPECULAR_PROP_SPECULAR_EXPONENT = 2

        ''' <summary>
        ''' Property Name: "SpecularConstant"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_DISTANTSPECULAR_PROP_SPECULAR_CONSTANT = 3

        ''' <summary>
        ''' Property Name: "SurfaceScale"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_DISTANTSPECULAR_PROP_SURFACE_SCALE = 4

        ''' <summary>
        ''' Property Name: "Color"
        ''' Property Type: D2D1_VECTOR_3F
        ''' </summary>
        D2D1_DISTANTSPECULAR_PROP_COLOR = 5

        ''' <summary>
        ''' Property Name: "KernelUnitLength"
        ''' Property Type: D2D1_VECTOR_2F
        ''' </summary>
        D2D1_DISTANTSPECULAR_PROP_KERNEL_UNIT_LENGTH = 6

        ''' <summary>
        ''' Property Name: "ScaleMode"
        ''' Property Type: D2D1_DISTANTSPECULAR_SCALE_MODE
        ''' </summary>
        D2D1_DISTANTSPECULAR_PROP_SCALE_MODE = 7
        D2D1_DISTANTSPECULAR_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_DISTANTSPECULAR_SCALE_MODE
        D2D1_DISTANTSPECULAR_SCALE_MODE_NEAREST_NEIGHBOR = 0
        D2D1_DISTANTSPECULAR_SCALE_MODE_LINEAR = 1
        D2D1_DISTANTSPECULAR_SCALE_MODE_CUBIC = 2
        D2D1_DISTANTSPECULAR_SCALE_MODE_MULTI_SAMPLE_LINEAR = 3
        D2D1_DISTANTSPECULAR_SCALE_MODE_ANISOTROPIC = 4
        D2D1_DISTANTSPECULAR_SCALE_MODE_HIGH_QUALITY_CUBIC = 5
        D2D1_DISTANTSPECULAR_SCALE_MODE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Point-Diffuse effect's top level properties.
    ''' Effect description: Creates a diffuse lighting effect with a point light source.
    ''' </summary>
    Public Enum D2D1_POINTDIFFUSE_PROP
        ''' <summary>
        ''' Property Name: "LightPosition"
        ''' Property Type: D2D1_VECTOR_3F
        ''' </summary>
        D2D1_POINTDIFFUSE_PROP_LIGHT_POSITION = 0

        ''' <summary>
        ''' Property Name: "DiffuseConstant"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_POINTDIFFUSE_PROP_DIFFUSE_CONSTANT = 1

        ''' <summary>
        ''' Property Name: "SurfaceScale"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_POINTDIFFUSE_PROP_SURFACE_SCALE = 2

        ''' <summary>
        ''' Property Name: "Color"
        ''' Property Type: D2D1_VECTOR_3F
        ''' </summary>
        D2D1_POINTDIFFUSE_PROP_COLOR = 3

        ''' <summary>
        ''' Property Name: "KernelUnitLength"
        ''' Property Type: D2D1_VECTOR_2F
        ''' </summary>
        D2D1_POINTDIFFUSE_PROP_KERNEL_UNIT_LENGTH = 4

        ''' <summary>
        ''' Property Name: "ScaleMode"
        ''' Property Type: D2D1_POINTDIFFUSE_SCALE_MODE
        ''' </summary>
        D2D1_POINTDIFFUSE_PROP_SCALE_MODE = 5
        D2D1_POINTDIFFUSE_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_POINTDIFFUSE_SCALE_MODE
        D2D1_POINTDIFFUSE_SCALE_MODE_NEAREST_NEIGHBOR = 0
        D2D1_POINTDIFFUSE_SCALE_MODE_LINEAR = 1
        D2D1_POINTDIFFUSE_SCALE_MODE_CUBIC = 2
        D2D1_POINTDIFFUSE_SCALE_MODE_MULTI_SAMPLE_LINEAR = 3
        D2D1_POINTDIFFUSE_SCALE_MODE_ANISOTROPIC = 4
        D2D1_POINTDIFFUSE_SCALE_MODE_HIGH_QUALITY_CUBIC = 5
        D2D1_POINTDIFFUSE_SCALE_MODE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Spot-Diffuse effect's top level properties.
    ''' Effect description: Creates a diffuse lighting effect with a spot light source.
    ''' </summary>
    Public Enum D2D1_SPOTDIFFUSE_PROP
        ''' <summary>
        ''' Property Name: "LightPosition"
        ''' Property Type: D2D1_VECTOR_3F
        ''' </summary>
        D2D1_SPOTDIFFUSE_PROP_LIGHT_POSITION = 0

        ''' <summary>
        ''' Property Name: "PointsAt"
        ''' Property Type: D2D1_VECTOR_3F
        ''' </summary>
        D2D1_SPOTDIFFUSE_PROP_POINTS_AT = 1

        ''' <summary>
        ''' Property Name: "Focus"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_SPOTDIFFUSE_PROP_FOCUS = 2

        ''' <summary>
        ''' Property Name: "LimitingConeAngle"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_SPOTDIFFUSE_PROP_LIMITING_CONE_ANGLE = 3

        ''' <summary>
        ''' Property Name: "DiffuseConstant"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_SPOTDIFFUSE_PROP_DIFFUSE_CONSTANT = 4

        ''' <summary>
        ''' Property Name: "SurfaceScale"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_SPOTDIFFUSE_PROP_SURFACE_SCALE = 5

        ''' <summary>
        ''' Property Name: "Color"
        ''' Property Type: D2D1_VECTOR_3F
        ''' </summary>
        D2D1_SPOTDIFFUSE_PROP_COLOR = 6

        ''' <summary>
        ''' Property Name: "KernelUnitLength"
        ''' Property Type: D2D1_VECTOR_2F
        ''' </summary>
        D2D1_SPOTDIFFUSE_PROP_KERNEL_UNIT_LENGTH = 7

        ''' <summary>
        ''' Property Name: "ScaleMode"
        ''' Property Type: D2D1_SPOTDIFFUSE_SCALE_MODE
        ''' </summary>
        D2D1_SPOTDIFFUSE_PROP_SCALE_MODE = 8
        D2D1_SPOTDIFFUSE_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_SPOTDIFFUSE_SCALE_MODE
        D2D1_SPOTDIFFUSE_SCALE_MODE_NEAREST_NEIGHBOR = 0
        D2D1_SPOTDIFFUSE_SCALE_MODE_LINEAR = 1
        D2D1_SPOTDIFFUSE_SCALE_MODE_CUBIC = 2
        D2D1_SPOTDIFFUSE_SCALE_MODE_MULTI_SAMPLE_LINEAR = 3
        D2D1_SPOTDIFFUSE_SCALE_MODE_ANISOTROPIC = 4
        D2D1_SPOTDIFFUSE_SCALE_MODE_HIGH_QUALITY_CUBIC = 5
        D2D1_SPOTDIFFUSE_SCALE_MODE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Distant-Diffuse effect's top level properties.
    ''' Effect description: Creates a diffuse lighting effect with a distant light
    ''' source.
    ''' </summary>
    Public Enum D2D1_DISTANTDIFFUSE_PROP
        ''' <summary>
        ''' Property Name: "Azimuth"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_DISTANTDIFFUSE_PROP_AZIMUTH = 0

        ''' <summary>
        ''' Property Name: "Elevation"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_DISTANTDIFFUSE_PROP_ELEVATION = 1

        ''' <summary>
        ''' Property Name: "DiffuseConstant"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_DISTANTDIFFUSE_PROP_DIFFUSE_CONSTANT = 2

        ''' <summary>
        ''' Property Name: "SurfaceScale"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_DISTANTDIFFUSE_PROP_SURFACE_SCALE = 3

        ''' <summary>
        ''' Property Name: "Color"
        ''' Property Type: D2D1_VECTOR_3F
        ''' </summary>
        D2D1_DISTANTDIFFUSE_PROP_COLOR = 4

        ''' <summary>
        ''' Property Name: "KernelUnitLength"
        ''' Property Type: D2D1_VECTOR_2F
        ''' </summary>
        D2D1_DISTANTDIFFUSE_PROP_KERNEL_UNIT_LENGTH = 5

        ''' <summary>
        ''' Property Name: "ScaleMode"
        ''' Property Type: D2D1_DISTANTDIFFUSE_SCALE_MODE
        ''' </summary>
        D2D1_DISTANTDIFFUSE_PROP_SCALE_MODE = 6
        D2D1_DISTANTDIFFUSE_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_DISTANTDIFFUSE_SCALE_MODE
        D2D1_DISTANTDIFFUSE_SCALE_MODE_NEAREST_NEIGHBOR = 0
        D2D1_DISTANTDIFFUSE_SCALE_MODE_LINEAR = 1
        D2D1_DISTANTDIFFUSE_SCALE_MODE_CUBIC = 2
        D2D1_DISTANTDIFFUSE_SCALE_MODE_MULTI_SAMPLE_LINEAR = 3
        D2D1_DISTANTDIFFUSE_SCALE_MODE_ANISOTROPIC = 4
        D2D1_DISTANTDIFFUSE_SCALE_MODE_HIGH_QUALITY_CUBIC = 5
        D2D1_DISTANTDIFFUSE_SCALE_MODE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Flood effect's top level properties.
    ''' Effect description: Renders an infinite sized floodfill of the given color.
    ''' </summary>
    Public Enum D2D1_FLOOD_PROP
        ''' <summary>
        ''' Property Name: "Color"
        ''' Property Type: D2D1_VECTOR_4F
        ''' </summary>
        D2D1_FLOOD_PROP_COLOR = 0
        D2D1_FLOOD_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Linear Transfer effect's top level properties.
    ''' Effect description: Remaps the color intensities of the input bitmap based on a
    ''' user specified linear transfer function for each RGBA channel.
    ''' </summary>
    Public Enum D2D1_LINEARTRANSFER_PROP
        ''' <summary>
        ''' Property Name: "RedYIntercept"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_LINEARTRANSFER_PROP_RED_Y_INTERCEPT = 0

        ''' <summary>
        ''' Property Name: "RedSlope"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_LINEARTRANSFER_PROP_RED_SLOPE = 1

        ''' <summary>
        ''' Property Name: "RedDisable"
        ''' Property Type: BOOL
        ''' </summary>
        D2D1_LINEARTRANSFER_PROP_RED_DISABLE = 2

        ''' <summary>
        ''' Property Name: "GreenYIntercept"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_LINEARTRANSFER_PROP_GREEN_Y_INTERCEPT = 3

        ''' <summary>
        ''' Property Name: "GreenSlope"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_LINEARTRANSFER_PROP_GREEN_SLOPE = 4

        ''' <summary>
        ''' Property Name: "GreenDisable"
        ''' Property Type: BOOL
        ''' </summary>
        D2D1_LINEARTRANSFER_PROP_GREEN_DISABLE = 5

        ''' <summary>
        ''' Property Name: "BlueYIntercept"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_LINEARTRANSFER_PROP_BLUE_Y_INTERCEPT = 6

        ''' <summary>
        ''' Property Name: "BlueSlope"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_LINEARTRANSFER_PROP_BLUE_SLOPE = 7

        ''' <summary>
        ''' Property Name: "BlueDisable"
        ''' Property Type: BOOL
        ''' </summary>
        D2D1_LINEARTRANSFER_PROP_BLUE_DISABLE = 8

        ''' <summary>
        ''' Property Name: "AlphaYIntercept"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_LINEARTRANSFER_PROP_ALPHA_Y_INTERCEPT = 9

        ''' <summary>
        ''' Property Name: "AlphaSlope"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_LINEARTRANSFER_PROP_ALPHA_SLOPE = 10

        ''' <summary>
        ''' Property Name: "AlphaDisable"
        ''' Property Type: BOOL
        ''' </summary>
        D2D1_LINEARTRANSFER_PROP_ALPHA_DISABLE = 11

        ''' <summary>
        ''' Property Name: "ClampOutput"
        ''' Property Type: BOOL
        ''' </summary>
        D2D1_LINEARTRANSFER_PROP_CLAMP_OUTPUT = 12
        D2D1_LINEARTRANSFER_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Gamma Transfer effect's top level properties.
    ''' Effect description: Remaps the color intensities of the input bitmap based on a
    ''' user specified gamma transfer function for each RGBA channel.
    ''' </summary>
    Public Enum D2D1_GAMMATRANSFER_PROP
        ''' <summary>
        ''' Property Name: "RedAmplitude"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_GAMMATRANSFER_PROP_RED_AMPLITUDE = 0

        ''' <summary>
        ''' Property Name: "RedExponent"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_GAMMATRANSFER_PROP_RED_EXPONENT = 1

        ''' <summary>
        ''' Property Name: "RedOffset"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_GAMMATRANSFER_PROP_RED_OFFSET = 2

        ''' <summary>
        ''' Property Name: "RedDisable"
        ''' Property Type: BOOL
        ''' </summary>
        D2D1_GAMMATRANSFER_PROP_RED_DISABLE = 3

        ''' <summary>
        ''' Property Name: "GreenAmplitude"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_GAMMATRANSFER_PROP_GREEN_AMPLITUDE = 4

        ''' <summary>
        ''' Property Name: "GreenExponent"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_GAMMATRANSFER_PROP_GREEN_EXPONENT = 5

        ''' <summary>
        ''' Property Name: "GreenOffset"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_GAMMATRANSFER_PROP_GREEN_OFFSET = 6

        ''' <summary>
        ''' Property Name: "GreenDisable"
        ''' Property Type: BOOL
        ''' </summary>
        D2D1_GAMMATRANSFER_PROP_GREEN_DISABLE = 7

        ''' <summary>
        ''' Property Name: "BlueAmplitude"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_GAMMATRANSFER_PROP_BLUE_AMPLITUDE = 8

        ''' <summary>
        ''' Property Name: "BlueExponent"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_GAMMATRANSFER_PROP_BLUE_EXPONENT = 9

        ''' <summary>
        ''' Property Name: "BlueOffset"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_GAMMATRANSFER_PROP_BLUE_OFFSET = 10

        ''' <summary>
        ''' Property Name: "BlueDisable"
        ''' Property Type: BOOL
        ''' </summary>
        D2D1_GAMMATRANSFER_PROP_BLUE_DISABLE = 11

        ''' <summary>
        ''' Property Name: "AlphaAmplitude"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_GAMMATRANSFER_PROP_ALPHA_AMPLITUDE = 12

        ''' <summary>
        ''' Property Name: "AlphaExponent"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_GAMMATRANSFER_PROP_ALPHA_EXPONENT = 13

        ''' <summary>
        ''' Property Name: "AlphaOffset"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_GAMMATRANSFER_PROP_ALPHA_OFFSET = 14

        ''' <summary>
        ''' Property Name: "AlphaDisable"
        ''' Property Type: BOOL
        ''' </summary>
        D2D1_GAMMATRANSFER_PROP_ALPHA_DISABLE = 15

        ''' <summary>
        ''' Property Name: "ClampOutput"
        ''' Property Type: BOOL
        ''' </summary>
        D2D1_GAMMATRANSFER_PROP_CLAMP_OUTPUT = 16
        D2D1_GAMMATRANSFER_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Table Transfer effect's top level properties.
    ''' Effect description: Remaps the color intensities of the input bitmap based on a
    ''' transfer function generated by a user specified list of values for each RGBA
    ''' channel.
    ''' </summary>
    Public Enum D2D1_TABLETRANSFER_PROP
        ''' <summary>
        ''' Property Name: "RedTable"
        ''' Property Type: (blob)
        ''' </summary>
        D2D1_TABLETRANSFER_PROP_RED_TABLE = 0

        ''' <summary>
        ''' Property Name: "RedDisable"
        ''' Property Type: BOOL
        ''' </summary>
        D2D1_TABLETRANSFER_PROP_RED_DISABLE = 1

        ''' <summary>
        ''' Property Name: "GreenTable"
        ''' Property Type: (blob)
        ''' </summary>
        D2D1_TABLETRANSFER_PROP_GREEN_TABLE = 2

        ''' <summary>
        ''' Property Name: "GreenDisable"
        ''' Property Type: BOOL
        ''' </summary>
        D2D1_TABLETRANSFER_PROP_GREEN_DISABLE = 3

        ''' <summary>
        ''' Property Name: "BlueTable"
        ''' Property Type: (blob)
        ''' </summary>
        D2D1_TABLETRANSFER_PROP_BLUE_TABLE = 4

        ''' <summary>
        ''' Property Name: "BlueDisable"
        ''' Property Type: BOOL
        ''' </summary>
        D2D1_TABLETRANSFER_PROP_BLUE_DISABLE = 5

        ''' <summary>
        ''' Property Name: "AlphaTable"
        ''' Property Type: (blob)
        ''' </summary>
        D2D1_TABLETRANSFER_PROP_ALPHA_TABLE = 6

        ''' <summary>
        ''' Property Name: "AlphaDisable"
        ''' Property Type: BOOL
        ''' </summary>
        D2D1_TABLETRANSFER_PROP_ALPHA_DISABLE = 7

        ''' <summary>
        ''' Property Name: "ClampOutput"
        ''' Property Type: BOOL
        ''' </summary>
        D2D1_TABLETRANSFER_PROP_CLAMP_OUTPUT = 8
        D2D1_TABLETRANSFER_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Discrete Transfer effect's top level properties.
    ''' Effect description: Remaps the color intensities of the input bitmap based on a
    ''' discrete function generated by a user specified list of values for each RGBA
    ''' channel.
    ''' </summary>
    Public Enum D2D1_DISCRETETRANSFER_PROP
        ''' <summary>
        ''' Property Name: "RedTable"
        ''' Property Type: (blob)
        ''' </summary>
        D2D1_DISCRETETRANSFER_PROP_RED_TABLE = 0

        ''' <summary>
        ''' Property Name: "RedDisable"
        ''' Property Type: BOOL
        ''' </summary>
        D2D1_DISCRETETRANSFER_PROP_RED_DISABLE = 1

        ''' <summary>
        ''' Property Name: "GreenTable"
        ''' Property Type: (blob)
        ''' </summary>
        D2D1_DISCRETETRANSFER_PROP_GREEN_TABLE = 2

        ''' <summary>
        ''' Property Name: "GreenDisable"
        ''' Property Type: BOOL
        ''' </summary>
        D2D1_DISCRETETRANSFER_PROP_GREEN_DISABLE = 3

        ''' <summary>
        ''' Property Name: "BlueTable"
        ''' Property Type: (blob)
        ''' </summary>
        D2D1_DISCRETETRANSFER_PROP_BLUE_TABLE = 4

        ''' <summary>
        ''' Property Name: "BlueDisable"
        ''' Property Type: BOOL
        ''' </summary>
        D2D1_DISCRETETRANSFER_PROP_BLUE_DISABLE = 5

        ''' <summary>
        ''' Property Name: "AlphaTable"
        ''' Property Type: (blob)
        ''' </summary>
        D2D1_DISCRETETRANSFER_PROP_ALPHA_TABLE = 6

        ''' <summary>
        ''' Property Name: "AlphaDisable"
        ''' Property Type: BOOL
        ''' </summary>
        D2D1_DISCRETETRANSFER_PROP_ALPHA_DISABLE = 7

        ''' <summary>
        ''' Property Name: "ClampOutput"
        ''' Property Type: BOOL
        ''' </summary>
        D2D1_DISCRETETRANSFER_PROP_CLAMP_OUTPUT = 8
        D2D1_DISCRETETRANSFER_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Convolve Matrix effect's top level properties.
    ''' Effect description: Applies a user specified convolution kernel to a bitmap.
    ''' </summary>
    Public Enum D2D1_CONVOLVEMATRIX_PROP
        ''' <summary>
        ''' Property Name: "KernelUnitLength"
        ''' Property Type: D2D1_VECTOR_2F
        ''' </summary>
        D2D1_CONVOLVEMATRIX_PROP_KERNEL_UNIT_LENGTH = 0

        ''' <summary>
        ''' Property Name: "ScaleMode"
        ''' Property Type: D2D1_CONVOLVEMATRIX_SCALE_MODE
        ''' </summary>
        D2D1_CONVOLVEMATRIX_PROP_SCALE_MODE = 1

        ''' <summary>
        ''' Property Name: "KernelSizeX"
        ''' Property Type: UINT32
        ''' </summary>
        D2D1_CONVOLVEMATRIX_PROP_KERNEL_SIZE_X = 2

        ''' <summary>
        ''' Property Name: "KernelSizeY"
        ''' Property Type: UINT32
        ''' </summary>
        D2D1_CONVOLVEMATRIX_PROP_KERNEL_SIZE_Y = 3

        ''' <summary>
        ''' Property Name: "KernelMatrix"
        ''' Property Type: (blob)
        ''' </summary>
        D2D1_CONVOLVEMATRIX_PROP_KERNEL_MATRIX = 4

        ''' <summary>
        ''' Property Name: "Divisor"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_CONVOLVEMATRIX_PROP_DIVISOR = 5

        ''' <summary>
        ''' Property Name: "Bias"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_CONVOLVEMATRIX_PROP_BIAS = 6

        ''' <summary>
        ''' Property Name: "KernelOffset"
        ''' Property Type: D2D1_VECTOR_2F
        ''' </summary>
        D2D1_CONVOLVEMATRIX_PROP_KERNEL_OFFSET = 7

        ''' <summary>
        ''' Property Name: "PreserveAlpha"
        ''' Property Type: BOOL
        ''' </summary>
        D2D1_CONVOLVEMATRIX_PROP_PRESERVE_ALPHA = 8

        ''' <summary>
        ''' Property Name: "BorderMode"
        ''' Property Type: D2D1_BORDER_MODE
        ''' </summary>
        D2D1_CONVOLVEMATRIX_PROP_BORDER_MODE = 9

        ''' <summary>
        ''' Property Name: "ClampOutput"
        ''' Property Type: BOOL
        ''' </summary>
        D2D1_CONVOLVEMATRIX_PROP_CLAMP_OUTPUT = 10
        D2D1_CONVOLVEMATRIX_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_CONVOLVEMATRIX_SCALE_MODE
        D2D1_CONVOLVEMATRIX_SCALE_MODE_NEAREST_NEIGHBOR = 0
        D2D1_CONVOLVEMATRIX_SCALE_MODE_LINEAR = 1
        D2D1_CONVOLVEMATRIX_SCALE_MODE_CUBIC = 2
        D2D1_CONVOLVEMATRIX_SCALE_MODE_MULTI_SAMPLE_LINEAR = 3
        D2D1_CONVOLVEMATRIX_SCALE_MODE_ANISOTROPIC = 4
        D2D1_CONVOLVEMATRIX_SCALE_MODE_HIGH_QUALITY_CUBIC = 5
        D2D1_CONVOLVEMATRIX_SCALE_MODE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Brightness effect's top level properties.
    ''' Effect description: Adjusts the brightness of the image based on the specified
    ''' white and black point.
    ''' </summary>
    Public Enum D2D1_BRIGHTNESS_PROP
        ''' <summary>
        ''' Property Name: "WhitePoint"
        ''' Property Type: D2D1_VECTOR_2F
        ''' </summary>
        D2D1_BRIGHTNESS_PROP_WHITE_POINT = 0

        ''' <summary>
        ''' Property Name: "BlackPoint"
        ''' Property Type: D2D1_VECTOR_2F
        ''' </summary>
        D2D1_BRIGHTNESS_PROP_BLACK_POINT = 1
        D2D1_BRIGHTNESS_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Arithmetic Composite effect's top level properties.
    ''' Effect description: Composites two bitmaps based on the following algorithm:
    ''' Output = Coefficients_1 * Source * Destination + Coefficients_2 * Source+
    ''' Coefficients_3 * Destination + Coefficients_4.
    ''' </summary>
    Public Enum D2D1_ARITHMETICCOMPOSITE_PROP
        ''' <summary>
        ''' Property Name: "Coefficients"
        ''' Property Type: D2D1_VECTOR_4F
        ''' </summary>
        D2D1_ARITHMETICCOMPOSITE_PROP_COEFFICIENTS = 0

        ''' <summary>
        ''' Property Name: "ClampOutput"
        ''' Property Type: BOOL
        ''' </summary>
        D2D1_ARITHMETICCOMPOSITE_PROP_CLAMP_OUTPUT = 1
        D2D1_ARITHMETICCOMPOSITE_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Crop effect's top level properties.
    ''' Effect description: Crops the input bitmap according to the specified
    ''' parameters.
    ''' </summary>
    Public Enum D2D1_CROP_PROP
        ''' <summary>
        ''' Property Name: "Rect"
        ''' Property Type: D2D1_VECTOR_4F
        ''' </summary>
        D2D1_CROP_PROP_RECT = 0

        ''' <summary>
        ''' Property Name: "BorderMode"
        ''' Property Type: D2D1_BORDER_MODE
        ''' </summary>
        D2D1_CROP_PROP_BORDER_MODE = 1
        D2D1_CROP_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Border effect's top level properties.
    ''' Effect description: Extends the region of the bitmap based on the selected
    ''' border mode.
    ''' </summary>
    Public Enum D2D1_BORDER_PROP
        ''' <summary>
        ''' Property Name: "EdgeModeX"
        ''' Property Type: D2D1_BORDER_EDGE_MODE
        ''' </summary>
        D2D1_BORDER_PROP_EDGE_MODE_X = 0

        ''' <summary>
        ''' Property Name: "EdgeModeY"
        ''' Property Type: D2D1_BORDER_EDGE_MODE
        ''' </summary>
        D2D1_BORDER_PROP_EDGE_MODE_Y = 1
        D2D1_BORDER_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The edge mode for the Border effect.
    ''' </summary>
    Public Enum D2D1_BORDER_EDGE_MODE
        D2D1_BORDER_EDGE_MODE_CLAMP = 0
        D2D1_BORDER_EDGE_MODE_WRAP = 1
        D2D1_BORDER_EDGE_MODE_MIRROR = 2
        D2D1_BORDER_EDGE_MODE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Morphology effect's top level properties.
    ''' Effect description: Erodes or dilates a bitmap by the given radius.
    ''' </summary>
    Public Enum D2D1_MORPHOLOGY_PROP
        ''' <summary>
        ''' Property Name: "Mode"
        ''' Property Type: D2D1_MORPHOLOGY_MODE
        ''' </summary>
        D2D1_MORPHOLOGY_PROP_MODE = 0

        ''' <summary>
        ''' Property Name: "Width"
        ''' Property Type: UINT32
        ''' </summary>
        D2D1_MORPHOLOGY_PROP_WIDTH = 1

        ''' <summary>
        ''' Property Name: "Height"
        ''' Property Type: UINT32
        ''' </summary>
        D2D1_MORPHOLOGY_PROP_HEIGHT = 2
        D2D1_MORPHOLOGY_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_MORPHOLOGY_MODE
        D2D1_MORPHOLOGY_MODE_ERODE = 0
        D2D1_MORPHOLOGY_MODE_DILATE = 1
        D2D1_MORPHOLOGY_MODE_FORCE_DWORD = &HFFFFFFFF
    End Enum


    ''' <summary>
    ''' The enumeration of the Tile effect's top level properties.
    ''' Effect description: Tiles the specified region of the input bitmap.
    ''' </summary>
    Public Enum D2D1_TILE_PROP
        ''' <summary>
        ''' Property Name: "Rect"
        ''' Property Type: D2D1_VECTOR_4F
        ''' </summary>
        D2D1_TILE_PROP_RECT = 0
        D2D1_TILE_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Atlas effect's top level properties.
    ''' Effect description: Changes the available area of the input to the specified
    ''' rectangle. Provides an optimization for scenarios where a bitmap is used as an
    ''' atlas.
    ''' </summary>
    Public Enum D2D1_ATLAS_PROP
        ''' <summary>
        ''' Property Name: "InputRect"
        ''' Property Type: D2D1_VECTOR_4F
        ''' </summary>
        D2D1_ATLAS_PROP_INPUT_RECT = 0

        ''' <summary>
        ''' Property Name: "InputPaddingRect"
        ''' Property Type: D2D1_VECTOR_4F
        ''' </summary>
        D2D1_ATLAS_PROP_INPUT_PADDING_RECT = 1
        D2D1_ATLAS_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Opacity Metadata effect's top level properties.
    ''' Effect description: Changes the rectangle which is assumed to be opaque.
    ''' Provides optimizations in certain scenarios.
    ''' </summary>
    Public Enum D2D1_OPACITYMETADATA_PROP
        ''' <summary>
        ''' Property Name: "InputOpaqueRect"
        ''' Property Type: D2D1_VECTOR_4F
        ''' </summary>
        D2D1_OPACITYMETADATA_PROP_INPUT_OPAQUE_RECT = 0
        D2D1_OPACITYMETADATA_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum


    ''' <summary>
    ''' The enumeration of the Contrast effect's top level properties.
    ''' Effect description: Adjusts the contrast of an image.
    ''' </summary>
    Public Enum D2D1_CONTRAST_PROP
        ''' <summary>
        ''' Property Name: "Contrast"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_CONTRAST_PROP_CONTRAST = 0

        ''' <summary>
        ''' Property Name: "ClampInput"
        ''' Property Type: BOOL
        ''' </summary>
        D2D1_CONTRAST_PROP_CLAMP_INPUT = 1
        D2D1_CONTRAST_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the RgbToHue effect's top level properties.
    ''' Effect description: Converts an RGB bitmap to HSV or HSL.
    ''' </summary>
    Public Enum D2D1_RGBTOHUE_PROP
        ''' <summary>
        ''' Property Name: "OutputColorSpace"
        ''' Property Type: D2D1_RGBTOHUE_OUTPUT_COLOR_SPACE
        ''' </summary>
        D2D1_RGBTOHUE_PROP_OUTPUT_COLOR_SPACE = 0
        D2D1_RGBTOHUE_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_RGBTOHUE_OUTPUT_COLOR_SPACE
        D2D1_RGBTOHUE_OUTPUT_COLOR_SPACE_HUE_SATURATION_VALUE = 0
        D2D1_RGBTOHUE_OUTPUT_COLOR_SPACE_HUE_SATURATION_LIGHTNESS = 1
        D2D1_RGBTOHUE_OUTPUT_COLOR_SPACE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the HueToRgb effect's top level properties.
    ''' Effect description: Converts an HSV or HSL bitmap into an RGB bitmap.
    ''' </summary>
    Public Enum D2D1_HUETORGB_PROP
        ''' <summary>
        ''' Property Name: "InputColorSpace"
        ''' Property Type: D2D1_HUETORGB_INPUT_COLOR_SPACE
        ''' </summary>
        D2D1_HUETORGB_PROP_INPUT_COLOR_SPACE = 0
        D2D1_HUETORGB_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_HUETORGB_INPUT_COLOR_SPACE
        D2D1_HUETORGB_INPUT_COLOR_SPACE_HUE_SATURATION_VALUE = 0
        D2D1_HUETORGB_INPUT_COLOR_SPACE_HUE_SATURATION_LIGHTNESS = 1
        D2D1_HUETORGB_INPUT_COLOR_SPACE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Chroma Key effect's top level properties.
    ''' Effect description: Converts a specified color to transparent.
    ''' </summary>
    Public Enum D2D1_CHROMAKEY_PROP
        ''' <summary>
        ''' Property Name: "Color"
        ''' Property Type: D2D1_VECTOR_3F
        ''' </summary>
        D2D1_CHROMAKEY_PROP_COLOR = 0

        ''' <summary>
        ''' Property Name: "Tolerance"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_CHROMAKEY_PROP_TOLERANCE = 1

        ''' <summary>
        ''' Property Name: "InvertAlpha"
        ''' Property Type: BOOL
        ''' </summary>
        D2D1_CHROMAKEY_PROP_INVERT_ALPHA = 2

        ''' <summary>
        ''' Property Name: "Feather"
        ''' Property Type: BOOL
        ''' </summary>
        D2D1_CHROMAKEY_PROP_FEATHER = 3
        D2D1_CHROMAKEY_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Emboss effect's top level properties.
    ''' Effect description: Applies an embossing effect to an image.
    ''' </summary>
    Public Enum D2D1_EMBOSS_PROP
        ''' <summary>
        ''' Property Name: "Height"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_EMBOSS_PROP_HEIGHT = 0

        ''' <summary>
        ''' Property Name: "Direction"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_EMBOSS_PROP_DIRECTION = 1
        D2D1_EMBOSS_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Exposure effect's top level properties.
    ''' Effect description: Simulates camera exposure adjustment.
    ''' </summary>
    Public Enum D2D1_EXPOSURE_PROP
        ''' <summary>
        ''' Property Name: "ExposureValue"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_EXPOSURE_PROP_EXPOSURE_VALUE = 0
        D2D1_EXPOSURE_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Posterize effect's top level properties.
    ''' Effect description: Reduces the number of colors in an image.
    ''' </summary>
    Public Enum D2D1_POSTERIZE_PROP
        ''' <summary>
        ''' Property Name: "RedValueCount"
        ''' Property Type: UINT32
        ''' </summary>
        D2D1_POSTERIZE_PROP_RED_VALUE_COUNT = 0

        ''' <summary>
        ''' Property Name: "GreenValueCount"
        ''' Property Type: UINT32
        ''' </summary>
        D2D1_POSTERIZE_PROP_GREEN_VALUE_COUNT = 1

        ''' <summary>
        ''' Property Name: "BlueValueCount"
        ''' Property Type: UINT32
        ''' </summary>
        D2D1_POSTERIZE_PROP_BLUE_VALUE_COUNT = 2
        D2D1_POSTERIZE_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Sepia effect's top level properties.
    ''' Effect description: Applies a Sepia tone to an image.
    ''' </summary>
    Public Enum D2D1_SEPIA_PROP
        ''' <summary>
        ''' Property Name: "Intensity"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_SEPIA_PROP_INTENSITY = 0

        ''' <summary>
        ''' Property Name: "AlphaMode"
        ''' Property Type: D2D1_ALPHA_MODE
        ''' </summary>
        D2D1_SEPIA_PROP_ALPHA_MODE = 1
        D2D1_SEPIA_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Sharpen effect's top level properties.
    ''' Effect description: Performs sharpening adjustment
    ''' </summary>
    Public Enum D2D1_SHARPEN_PROP
        ''' <summary>
        ''' Property Name: "Sharpness"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_SHARPEN_PROP_SHARPNESS = 0

        ''' <summary>
        ''' Property Name: "Threshold"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_SHARPEN_PROP_THRESHOLD = 1
        D2D1_SHARPEN_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Straighten effect's top level properties.
    ''' Effect description: Straightens an image.
    ''' </summary>
    Public Enum D2D1_STRAIGHTEN_PROP
        ''' <summary>
        ''' Property Name: "Angle"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_STRAIGHTEN_PROP_ANGLE = 0

        ''' <summary>
        ''' Property Name: "MaintainSize"
        ''' Property Type: BOOL
        ''' </summary>
        D2D1_STRAIGHTEN_PROP_MAINTAIN_SIZE = 1

        ''' <summary>
        ''' Property Name: "ScaleMode"
        ''' Property Type: D2D1_STRAIGHTEN_SCALE_MODE
        ''' </summary>
        D2D1_STRAIGHTEN_PROP_SCALE_MODE = 2
        D2D1_STRAIGHTEN_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_STRAIGHTEN_SCALE_MODE
        D2D1_STRAIGHTEN_SCALE_MODE_NEAREST_NEIGHBOR = 0
        D2D1_STRAIGHTEN_SCALE_MODE_LINEAR = 1
        D2D1_STRAIGHTEN_SCALE_MODE_CUBIC = 2
        D2D1_STRAIGHTEN_SCALE_MODE_MULTI_SAMPLE_LINEAR = 3
        D2D1_STRAIGHTEN_SCALE_MODE_ANISOTROPIC = 4
        D2D1_STRAIGHTEN_SCALE_MODE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Temperature And Tint effect's top level properties.
    ''' Effect description: Adjusts the temperature and tint of an image.
    ''' </summary>
    Public Enum D2D1_TEMPERATUREANDTINT_PROP
        ''' <summary>
        ''' Property Name: "Temperature"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_TEMPERATUREANDTINT_PROP_TEMPERATURE = 0

        ''' <summary>
        ''' Property Name: "Tint"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_TEMPERATUREANDTINT_PROP_TINT = 1
        D2D1_TEMPERATUREANDTINT_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Vignette effect's top level properties.
    ''' Effect description: Fades the edges of an image to the specified color.
    ''' </summary>
    Public Enum D2D1_VIGNETTE_PROP
        ''' <summary>
        ''' Property Name: "Color"
        ''' Property Type: D2D1_VECTOR_4F
        ''' </summary>
        D2D1_VIGNETTE_PROP_COLOR = 0

        ''' <summary>
        ''' Property Name: "TransitionSize"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_VIGNETTE_PROP_TRANSITION_SIZE = 1

        ''' <summary>
        ''' Property Name: "Strength"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_VIGNETTE_PROP_STRENGTH = 2
        D2D1_VIGNETTE_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Edge Detection effect's top level properties.
    ''' Effect description: Detects edges of an image.
    ''' </summary>
    Public Enum D2D1_EDGEDETECTION_PROP

        ''' <summary>
        ''' Property Name: "Strength"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_EDGEDETECTION_PROP_STRENGTH = 0

        ''' <summary>
        ''' Property Name: "BlurRadius"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_EDGEDETECTION_PROP_BLUR_RADIUS = 1

        ''' <summary>
        ''' Property Name: "Mode"
        ''' Property Type: D2D1_EDGEDETECTION_MODE
        ''' </summary>
        D2D1_EDGEDETECTION_PROP_MODE = 2

        ''' <summary>
        ''' Property Name: "OverlayEdges"
        ''' Property Type: BOOL
        ''' </summary>
        D2D1_EDGEDETECTION_PROP_OVERLAY_EDGES = 3

        ''' <summary>
        ''' Property Name: "AlphaMode"
        ''' Property Type: D2D1_ALPHA_MODE
        ''' </summary>
        D2D1_EDGEDETECTION_PROP_ALPHA_MODE = 4
        D2D1_EDGEDETECTION_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_EDGEDETECTION_MODE
        D2D1_EDGEDETECTION_MODE_SOBEL = 0
        D2D1_EDGEDETECTION_MODE_PREWITT = 1
        D2D1_EDGEDETECTION_MODE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Highlights and Shadows effect's top level properties.
    ''' Effect description: Adjusts the highlight and shadow strength of an image.
    ''' </summary>
    Public Enum D2D1_HIGHLIGHTSANDSHADOWS_PROP
        ''' <summary>
        ''' Property Name: "Highlights"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_HIGHLIGHTSANDSHADOWS_PROP_HIGHLIGHTS = 0

        ''' <summary>
        ''' Property Name: "Shadows"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_HIGHLIGHTSANDSHADOWS_PROP_SHADOWS = 1

        ''' <summary>
        ''' Property Name: "Clarity"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_HIGHLIGHTSANDSHADOWS_PROP_CLARITY = 2

        ''' <summary>
        ''' Property Name: "InputGamma"
        ''' Property Type: D2D1_HIGHLIGHTSANDSHADOWS_INPUT_GAMMA
        ''' </summary>
        D2D1_HIGHLIGHTSANDSHADOWS_PROP_INPUT_GAMMA = 3

        ''' <summary>
        ''' Property Name: "MaskBlurRadius"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_HIGHLIGHTSANDSHADOWS_PROP_MASK_BLUR_RADIUS = 4
        D2D1_HIGHLIGHTSANDSHADOWS_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_HIGHLIGHTSANDSHADOWS_INPUT_GAMMA
        D2D1_HIGHLIGHTSANDSHADOWS_INPUT_GAMMA_LINEAR = 0
        D2D1_HIGHLIGHTSANDSHADOWS_INPUT_GAMMA_SRGB = 1
        D2D1_HIGHLIGHTSANDSHADOWS_INPUT_GAMMA_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Lookup Table 3D effect's top level properties.
    ''' Effect description: Remaps colors in an image via a 3D lookup table.
    ''' </summary>
    Public Enum D2D1_LOOKUPTABLE3D_PROP
        ''' <summary>
        ''' Property Name: "Lut"
        ''' Property Type: IUnknown *
        ''' </summary>
        D2D1_LOOKUPTABLE3D_PROP_LUT = 0

        ''' <summary>
        ''' Property Name: "AlphaMode"
        ''' Property Type: D2D1_ALPHA_MODE
        ''' </summary>
        D2D1_LOOKUPTABLE3D_PROP_ALPHA_MODE = 1
        D2D1_LOOKUPTABLE3D_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Opacity effect's top level properties.
    ''' Effect description: Adjusts the opacity of an image by multiplying the alpha
    ''' channel by the specified opacity.
    ''' </summary>
    Public Enum D2D1_OPACITY_PROP
        ''' <summary>
        ''' Property Name: "Opacity"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_OPACITY_PROP_OPACITY = 0
        D2D1_OPACITY_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Cross Fade effect's top level properties.
    ''' Effect description: This effect combines two images by adding weighted pixels
    ''' from input images. The formula can be expressed as output = weight * Destination
    ''' + (1 - weight) * Source
    ''' </summary>
    Public Enum D2D1_CROSSFADE_PROP
        ''' <summary>
        ''' Property Name: "Weight"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_CROSSFADE_PROP_WEIGHT = 0
        D2D1_CROSSFADE_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the Tint effect's top level properties.
    ''' Effect description: This effect tints the source image by multiplying the
    ''' specified color by the source image.
    ''' </summary>
    Public Enum D2D1_TINT_PROP
        ''' <summary>
        ''' Property Name: "Color"
        ''' Property Type: D2D1_VECTOR_4F
        ''' </summary>
        D2D1_TINT_PROP_COLOR = 0

        ''' <summary>
        ''' Property Name: "ClampOutput"
        ''' Property Type: BOOL
        ''' </summary>
        D2D1_TINT_PROP_CLAMP_OUTPUT = 1
        D2D1_TINT_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the White Level Adjustment effect's top level properties.
    ''' Effect description: This effect adjusts the white level of the source image by
    ''' multiplying the source image color by the ratio of the input and output white
    ''' levels. Input and output white levels are specified in nits.
    ''' </summary>
    Public Enum D2D1_WHITELEVELADJUSTMENT_PROP
        ''' <summary>
        ''' Property Name: "InputWhiteLevel"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_WHITELEVELADJUSTMENT_PROP_INPUT_WHITE_LEVEL = 0

        ''' <summary>
        ''' Property Name: "OutputWhiteLevel"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_WHITELEVELADJUSTMENT_PROP_OUTPUT_WHITE_LEVEL = 1
        D2D1_WHITELEVELADJUSTMENT_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    ''' <summary>
    ''' The enumeration of the HDR Tone Map effect's top level properties.
    ''' Effect description: Adjusts the maximum luminance of the source image to fit
    ''' within the maximum output luminance supported. Input and output luminance values
    ''' are specified in nits. Note that the color space of the image is assumed to be
    ''' scRGB.
    ''' </summary>
    Public Enum D2D1_HDRTONEMAP_PROP
        ''' <summary>
        ''' Property Name: "InputMaxLuminance"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_HDRTONEMAP_PROP_INPUT_MAX_LUMINANCE = 0

        ''' <summary>
        ''' Property Name: "OutputMaxLuminance"
        ''' Property Type: FLOAT
        ''' </summary>
        D2D1_HDRTONEMAP_PROP_OUTPUT_MAX_LUMINANCE = 1

        ''' <summary>
        ''' Property Name: "DisplayMode"
        ''' Property Type: D2D1_HDRTONEMAP_DISPLAY_MODE
        ''' </summary>
        D2D1_HDRTONEMAP_PROP_DISPLAY_MODE = 2
        D2D1_HDRTONEMAP_PROP_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_HDRTONEMAP_DISPLAY_MODE
        D2D1_HDRTONEMAP_DISPLAY_MODE_SDR = 0
        D2D1_HDRTONEMAP_DISPLAY_MODE_HDR = 1
        D2D1_HDRTONEMAP_DISPLAY_MODE_FORCE_DWORD = &HFFFFFFFF
    End Enum




    Public Enum D2D1_UNIT_MODE
        D2D1_UNIT_MODE_DIPS = 0
        D2D1_UNIT_MODE_PIXELS = 1
        D2D1_UNIT_MODE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_RENDERING_CONTROLS
        ''' <summary>
        ''' The default buffer precision, used if the precision isn't otherwise specified.
        ''' </summary>
        Public bufferPrecision As D2D1_BUFFER_PRECISION

        ''' <summary>
        ''' The size of allocated tiles used to render imaging effects.
        ''' </summary>
        Public tileSize As D2D1_SIZE_U
    End Structure

    Public Enum D2D1_COMPOSITE_MODE
        D2D1_COMPOSITE_MODE_SOURCE_OVER = 0
        D2D1_COMPOSITE_MODE_DESTINATION_OVER = 1
        D2D1_COMPOSITE_MODE_SOURCE_IN = 2
        D2D1_COMPOSITE_MODE_DESTINATION_IN = 3
        D2D1_COMPOSITE_MODE_SOURCE_OUT = 4
        D2D1_COMPOSITE_MODE_DESTINATION_OUT = 5
        D2D1_COMPOSITE_MODE_SOURCE_ATOP = 6
        D2D1_COMPOSITE_MODE_DESTINATION_ATOP = 7
        D2D1_COMPOSITE_MODE_XOR = 8
        D2D1_COMPOSITE_MODE_PLUS = 9
        D2D1_COMPOSITE_MODE_SOURCE_COPY = 10
        D2D1_COMPOSITE_MODE_BOUNDED_SOURCE_COPY = 11
        D2D1_COMPOSITE_MODE_MASK_INVERT = 12
        D2D1_COMPOSITE_MODE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_PRIMITIVE_BLEND
        D2D1_PRIMITIVE_BLEND_SOURCE_OVER = 0
        D2D1_PRIMITIVE_BLEND_COPY = 1
        D2D1_PRIMITIVE_BLEND_MIN = 2
        D2D1_PRIMITIVE_BLEND_ADD = 3
        D2D1_PRIMITIVE_BLEND_MAX = 4
        D2D1_PRIMITIVE_BLEND_FORCE_DWORD = &HFFFFFFFF
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_IMAGE_BRUSH_PROPERTIES
        Public sourceRectangle As D2D1_RECT_F
        Public extendModeX As D2D1_EXTEND_MODE
        Public extendModeY As D2D1_EXTEND_MODE
        Public interpolationMode As D2D1_INTERPOLATION_MODE
    End Structure

    Public Enum D2D1_INTERPOLATION_MODE
        D2D1_INTERPOLATION_MODE_NEAREST_NEIGHBOR = D2D1_INTERPOLATION_MODE_DEFINITION.D2D1_INTERPOLATION_MODE_DEFINITION_NEAREST_NEIGHBOR
        D2D1_INTERPOLATION_MODE_LINEAR = D2D1_INTERPOLATION_MODE_DEFINITION.D2D1_INTERPOLATION_MODE_DEFINITION_LINEAR
        D2D1_INTERPOLATION_MODE_CUBIC = D2D1_INTERPOLATION_MODE_DEFINITION.D2D1_INTERPOLATION_MODE_DEFINITION_CUBIC
        D2D1_INTERPOLATION_MODE_MULTI_SAMPLE_LINEAR = D2D1_INTERPOLATION_MODE_DEFINITION.D2D1_INTERPOLATION_MODE_DEFINITION_MULTI_SAMPLE_LINEAR
        D2D1_INTERPOLATION_MODE_ANISOTROPIC = D2D1_INTERPOLATION_MODE_DEFINITION.D2D1_INTERPOLATION_MODE_DEFINITION_ANISOTROPIC
        D2D1_INTERPOLATION_MODE_HIGH_QUALITY_CUBIC = D2D1_INTERPOLATION_MODE_DEFINITION.D2D1_INTERPOLATION_MODE_DEFINITION_HIGH_QUALITY_CUBIC
        D2D1_INTERPOLATION_MODE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_BITMAP_BRUSH_PROPERTIES1
        Public extendModeX As D2D1_EXTEND_MODE
        Public extendModeY As D2D1_EXTEND_MODE
        Public interpolationMode As D2D1_INTERPOLATION_MODE
    End Structure

    Public Enum D2D1_BUFFER_PRECISION
        D2D1_BUFFER_PRECISION_UNKNOWN = 0
        D2D1_BUFFER_PRECISION_8BPC_UNORM = 1
        D2D1_BUFFER_PRECISION_8BPC_UNORM_SRGB = 2
        D2D1_BUFFER_PRECISION_16BPC_UNORM = 3
        D2D1_BUFFER_PRECISION_16BPC_FLOAT = 4
        D2D1_BUFFER_PRECISION_32BPC_FLOAT = 5
        D2D1_BUFFER_PRECISION_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_COLOR_INTERPOLATION_MODE
        ''' <summary>
        ''' Colors will be interpolated in straight alpha space.
        ''' </summary>
        D2D1_COLOR_INTERPOLATION_MODE_STRAIGHT = 0

        ''' <summary>
        ''' Colors will be interpolated in premultiplied alpha space.
        ''' </summary>
        D2D1_COLOR_INTERPOLATION_MODE_PREMULTIPLIED = 1
        D2D1_COLOR_INTERPOLATION_MODE_FORCE_DWORD = &HFFFFFFFF
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_BITMAP_PROPERTIES1
        Public pixelFormat As D2D1_PIXEL_FORMAT
        Public dpiX As Single
        Public dpiY As Single
        Public bitmapOptions As D2D1_BITMAP_OPTIONS

        Public colorContext As ID2D1ColorContext
    End Structure

    <ComImport>
    <Guid("a898a84c-3873-4588-b08b-ebbf978df041")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1Bitmap1
        Inherits ID2D1Bitmap
#Region "<ID2D1Bitmap>"

#Region "<ID2D1Image>"

#Region "<ID2D1Resource>"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)

#End Region

#End Region

        <PreserveSig>
        Overloads Sub GetSize(<Out> ByRef size As D2D1_SIZE_F)
        <PreserveSig>
        Overloads Sub GetPixelSize(<Out> ByRef size As D2D1_SIZE_U)
        <PreserveSig>
        Overloads Sub GetPixelFormat(<Out> ByRef format As D2D1_PIXEL_FORMAT)
        <PreserveSig>
        Overloads Sub GetDpi(<Out> ByRef dpiX As Single, <Out> ByRef dpiY As Single)
        <PreserveSig>
        Overloads Sub CopyFromBitmap(ByRef destPoint As D2D1_POINT_2U, bitmap As ID2D1Bitmap, srcRect As D2D1_RECT_U)
        <PreserveSig>
        Overloads Sub CopyFromRenderTarget(ByRef destPoint As D2D1_POINT_2U, renderTarget As ID2D1RenderTarget, srcRect As D2D1_RECT_U)
        <PreserveSig>
        Overloads Sub CopyFromMemory(ByRef dstRect As D2D1_RECT_U, srcData As IntPtr, pitch As UInteger)
#End Region

        <PreserveSig>
        Sub GetColorContext(<Out> ByRef colorContext As ID2D1ColorContext)
        <PreserveSig>
        Sub GetOptions(<Out> ByRef options As D2D1_BITMAP_OPTIONS)
        <PreserveSig>
        Function GetSurface(<Out> ByRef dxgiSurface As IDXGISurface) As HRESULT
        <PreserveSig>
        Function Map(options As D2D1_MAP_OPTIONS, <Out> ByRef mappedRect As D2D1_MAPPED_RECT) As HRESULT
        <PreserveSig>
        Function Unmap() As HRESULT
    End Interface

    Public Enum D2D1_BITMAP_OPTIONS
        ''' <summary>
        ''' The bitmap is created with default properties.
        ''' </summary>
        D2D1_BITMAP_OPTIONS_NONE = &H0

        ''' <summary>
        ''' The bitmap can be specified as a target in ID2D1DeviceContext::SetTarget
        ''' </summary>
        D2D1_BITMAP_OPTIONS_TARGET = &H1

        ''' <summary>
        ''' The bitmap cannot be used as an input to DrawBitmap, DrawImage, in a bitmap
        ''' brush or as an input to an effect.
        ''' </summary>
        D2D1_BITMAP_OPTIONS_CANNOT_DRAW = &H2

        ''' <summary>
        ''' The bitmap can be read from the CPU.
        ''' </summary>
        D2D1_BITMAP_OPTIONS_CPU_READ = &H4

        ''' <summary>
        ''' The bitmap works with the ID2D1GdiInteropRenderTarget::GetDC API.
        ''' </summary>
        D2D1_BITMAP_OPTIONS_GDI_COMPATIBLE = &H8
        D2D1_BITMAP_OPTIONS_FORCE_DWORD = &HFFFFFFFF
    End Enum

    Public Enum D2D1_MAP_OPTIONS
        ''' <summary>
        ''' The mapped pointer has undefined behavior.
        ''' </summary>
        D2D1_MAP_OPTIONS_NONE = 0

        ''' <summary>
        ''' The mapped pointer can be read from.
        ''' </summary>
        D2D1_MAP_OPTIONS_READ = 1

        ''' <summary>
        ''' The mapped pointer can be written to.
        ''' </summary>
        D2D1_MAP_OPTIONS_WRITE = 2

        ''' <summary>
        ''' The previous contents of the bitmap are discarded when it is mapped.
        ''' </summary>
        D2D1_MAP_OPTIONS_DISCARD = 4
        D2D1_MAP_OPTIONS_FORCE_DWORD = &HFFFFFFFF
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_MAPPED_RECT
        Public pitch As UInteger
        Public bits As IntPtr
    End Structure

    <ComImport>
    <Guid("1c4820bb-5771-4518-a581-2fe4dd0ec657")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1ColorContext
        Inherits ID2D1Resource
#Region "<ID2D1Resource>"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region

        <PreserveSig>
        Sub GetColorSpace(<Out> ByRef colorSpace As D2D1_COLOR_SPACE)
        <PreserveSig>
        Function GetProfileSize() As UInteger
        <PreserveSig>
        Function GetProfile(<Out> ByRef profile As IntPtr, profileSize As UInteger) As HRESULT
    End Interface

    Public Enum D2D1_COLOR_SPACE
        ''' <summary>
        ''' The color space is described by accompanying data, such as a color profile.
        ''' </summary>
        D2D1_COLOR_SPACE_CUSTOM = 0

        ''' <summary>
        ''' The sRGB color space.
        ''' </summary>
        D2D1_COLOR_SPACE_SRGB = 1

        ''' <summary>
        ''' The scRGB color space.
        ''' </summary>
        D2D1_COLOR_SPACE_SCRGB = 2
        D2D1_COLOR_SPACE_FORCE_DWORD = &HFFFFFFFF
    End Enum


    Public Class D2DTools
        Public Shared CLSID_D2D1Factory As Guid = New Guid("06152247-6f50-465a-9245-118bfd3b6007")

        Public Shared CLSID_D2D12DAffineTransform As Guid = New Guid("{6AA97485-6354-4CFC-908C-E4A74F62C96C}")
        Public Shared CLSID_D2D13DPerspectiveTransform As Guid = New Guid("{C2844D0B-3D86-46E7-85BA-526C9240F3FB}")
        Public Shared CLSID_D2D13DTransform As Guid = New Guid("{E8467B04-EC61-4B8A-B5DE-D4D73DEBEA5A}")
        Public Shared CLSID_D2D1ArithmeticComposite As Guid = New Guid("{FC151437-049A-4784-A24A-F1C4DAF20987}")
        Public Shared CLSID_D2D1Atlas As Guid = New Guid("{913E2BE4-FDCF-4FE2-A5F0-2454F14FF408}")
        Public Shared CLSID_D2D1BitmapSource As Guid = New Guid("{5FB6C24D-C6DD-4231-9404-50F4D5C3252D}")
        Public Shared CLSID_D2D1Blend As Guid = New Guid("{81C5B77B-13F8-4CDD-AD20-C890547AC65D}")
        Public Shared CLSID_D2D1Border As Guid = New Guid("{2A2D49C0-4ACF-43C7-8C6A-7C4A27874D27}")
        Public Shared CLSID_D2D1Brightness As Guid = New Guid("{8CEA8D1E-77B0-4986-B3B9-2F0C0EAE7887}")
        Public Shared CLSID_D2D1ColorManagement As Guid = New Guid("{1A28524C-FDD6-4AA4-AE8F-837EB8267B37}")
        Public Shared CLSID_D2D1ColorMatrix As Guid = New Guid("{921F03D6-641C-47DF-852D-B4BB6153AE11}")
        Public Shared CLSID_D2D1Composite As Guid = New Guid("{48FC9F51-F6AC-48F1-8B58-3B28AC46F76D}")
        Public Shared CLSID_D2D1ConvolveMatrix As Guid = New Guid("{407F8C08-5533-4331-A341-23CC3877843E}")
        Public Shared CLSID_D2D1Crop As Guid = New Guid("{E23F7110-0E9A-4324-AF47-6A2C0C46F35B}")
        Public Shared CLSID_D2D1DirectionalBlur As Guid = New Guid("{174319A6-58E9-49B2-BB63-CAF2C811A3DB}")
        Public Shared CLSID_D2D1DiscreteTransfer As Guid = New Guid("{90866FCD-488E-454B-AF06-E5041B66C36C}")
        Public Shared CLSID_D2D1DisplacementMap As Guid = New Guid("{EDC48364-0417-4111-9450-43845FA9F890}")
        Public Shared CLSID_D2D1DistantDiffuse As Guid = New Guid("{3E7EFD62-A32D-46D4-A83C-5278889AC954}")
        Public Shared CLSID_D2D1DistantSpecular As Guid = New Guid("{428C1EE5-77B8-4450-8AB5-72219C21ABDA}")
        Public Shared CLSID_D2D1DpiCompensation As Guid = New Guid("{6C26C5C7-34E0-46FC-9CFD-E5823706E228}")
        Public Shared CLSID_D2D1Flood As Guid = New Guid("{61C23C20-AE69-4D8E-94CF-50078DF638F2}")
        Public Shared CLSID_D2D1GammaTransfer As Guid = New Guid("{409444C4-C419-41A0-B0C1-8CD0C0A18E42}")
        Public Shared CLSID_D2D1GaussianBlur As Guid = New Guid("{1FEB6D69-2FE6-4AC9-8C58-1D7F93E7A6A5}")
        Public Shared CLSID_D2D1Scale As Guid = New Guid("{9DAF9369-3846-4D0E-A44E-0C607934A5D7}")
        Public Shared CLSID_D2D1Histogram As Guid = New Guid("{881DB7D0-F7EE-4D4D-A6D2-4697ACC66EE8}")
        Public Shared CLSID_D2D1HueRotation As Guid = New Guid("{0F4458EC-4B32-491B-9E85-BD73F44D3EB6}")
        Public Shared CLSID_D2D1LinearTransfer As Guid = New Guid("{AD47C8FD-63EF-4ACC-9B51-67979C036C06}")
        Public Shared CLSID_D2D1LuminanceToAlpha As Guid = New Guid("{41251AB7-0BEB-46F8-9DA7-59E93FCCE5DE}")
        Public Shared CLSID_D2D1Morphology As Guid = New Guid("{EAE6C40D-626A-4C2D-BFCB-391001ABE202}")
        Public Shared CLSID_D2D1OpacityMetadata As Guid = New Guid("{6C53006A-4450-4199-AA5B-AD1656FECE5E}")
        Public Shared CLSID_D2D1PointDiffuse As Guid = New Guid("{B9E303C3-C08C-4F91-8B7B-38656BC48C20}")
        Public Shared CLSID_D2D1PointSpecular As Guid = New Guid("{09C3CA26-3AE2-4F09-9EBC-ED3865D53F22}")
        Public Shared CLSID_D2D1Premultiply As Guid = New Guid("{06EAB419-DEED-4018-80D2-3E1D471ADEB2}")
        Public Shared CLSID_D2D1Saturation As Guid = New Guid("{5CB2D9CF-327D-459F-A0CE-40C0B2086BF7}")
        Public Shared CLSID_D2D1Shadow As Guid = New Guid("{C67EA361-1863-4E69-89DB-695D3E9A5B6B}")
        Public Shared CLSID_D2D1SpotDiffuse As Guid = New Guid("{818A1105-7932-44F4-AA86-08AE7B2F2C93}")
        Public Shared CLSID_D2D1SpotSpecular As Guid = New Guid("{EDAE421E-7654-4A37-9DB8-71ACC1BEB3C1}")
        Public Shared CLSID_D2D1TableTransfer As Guid = New Guid("{5BF818C3-5E43-48CB-B631-868396D6A1D4}")
        Public Shared CLSID_D2D1Tile As Guid = New Guid("{B0784138-3B76-4BC5-B13B-0FA2AD02659F}")
        Public Shared CLSID_D2D1Turbulence As Guid = New Guid("{CF2BB6AE-889A-4AD7-BA29-A2FD732C9FC9}")
        Public Shared CLSID_D2D1UnPremultiply As Guid = New Guid("{FB9AC489-AD8D-41ED-9999-BB6347D110F7}")

        Public Shared CLSID_D2D1Contrast As Guid = New Guid("{B648A78A-0ED5-4F80-A94A-8E825ACA6B77}")
        Public Shared CLSID_D2D1RgbToHue As Guid = New Guid("{23F3E5EC-91E8-4D3D-AD0A-AFADC1004AA1}")
        Public Shared CLSID_D2D1HueToRgb As Guid = New Guid("{7B78A6BD-0141-4DEF-8A52-6356EE0CBDD5}")
        Public Shared CLSID_D2D1ChromaKey As Guid = New Guid("{74C01F5B-2A0D-408C-88E2-C7A3C7197742}")
        Public Shared CLSID_D2D1Emboss As Guid = New Guid("{B1C5EB2B-0348-43F0-8107-4957CACBA2AE}")
        Public Shared CLSID_D2D1Exposure As Guid = New Guid("{B56C8CFA-F634-41EE-BEE0-FFA617106004}")
        Public Shared CLSID_D2D1Grayscale As Guid = New Guid("{36DDE0EB-3725-42E0-836D-52FB20AEE644}")
        Public Shared CLSID_D2D1Invert As Guid = New Guid("{E0C3784D-CB39-4E84-B6FD-6B72F0810263}")
        Public Shared CLSID_D2D1Posterize As Guid = New Guid("{2188945E-33A3-4366-B7BC-086BD02D0884}")
        Public Shared CLSID_D2D1Sepia As Guid = New Guid("{3A1AF410-5F1D-4DBE-84DF-915DA79B7153}")
        Public Shared CLSID_D2D1Sharpen As Guid = New Guid("{C9B887CB-C5FF-4DC5-9779-273DCF417C7D}")
        Public Shared CLSID_D2D1Straighten As Guid = New Guid("{4DA47B12-79A3-4FB0-8237-BBC3B2A4DE08}")
        Public Shared CLSID_D2D1TemperatureTint As Guid = New Guid("{89176087-8AF9-4A08-AEB1-895F38DB1766}")
        Public Shared CLSID_D2D1Vignette As Guid = New Guid("{C00C40BE-5E67-4CA3-95B4-F4B02C115135}")
        Public Shared CLSID_D2D1EdgeDetection As Guid = New Guid("{EFF583CA-CB07-4AA9-AC5D-2CC44C76460F}")
        Public Shared CLSID_D2D1HighlightsShadows As Guid = New Guid("{CADC8384-323F-4C7E-A361-2E2B24DF6EE4}")
        Public Shared CLSID_D2D1LookupTable3D As Guid = New Guid("{349E0EDA-0088-4A79-9CA3-C7E300202020}")
        Public Shared CLSID_D2D1Opacity As Guid = New Guid("{811D79A4-DE28-4454-8094-C64685F8BD4C}")
        Public Shared CLSID_D2D1AlphaMask As Guid = New Guid("{C80ECFF0-3FD5-4F05-8328-C5D1724B4F0A}")
        Public Shared CLSID_D2D1CrossFade As Guid = New Guid("{12F575E8-4DB1-485F-9A84-03A07DD3829F}")
        Public Shared CLSID_D2D1Tint As Guid = New Guid("{36312B17-F7DD-4014-915D-FFCA768CF211}")
        Public Shared CLSID_D2D1WhiteLevelAdjustment As Guid = New Guid("{44A1CADB-6CDD-4818-8FF4-26C1CFE95BDB}")
        Public Shared CLSID_D2D1HdrToneMap As Guid = New Guid("{7B0B748D-4610-4486-A90C-999D9A2E2B11}")

        <DllImport("D2D1.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
        Public Shared Function D2D1CreateFactory(factoryType As D2D1_FACTORY_TYPE, ByRef riid As Guid, ByRef pFactoryOptions As D2D1_FACTORY_OPTIONS, <Out> ByRef ppIFactory As ID2D1Factory) As HRESULT
        End Function

        Public Declare Sub D2D1MakeRotateMatrix_ Lib "D2D1.dll" Alias "D2D1MakeRotateMatrix" (ByVal arg0 As Single, ByVal arg1 As D2D1_POINT_2F, ByVal arg2 As IntPtr)

        <DllImport("D2D1.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
        Public Shared Function D2D1CreateDevice(dxgiDevice As IDXGIDevice, ByRef creationProperties As D2D1_CREATION_PROPERTIES, <Out> ByRef d2dDevice As ID2D1Device) As HRESULT
        End Function

        <DllImport("D2D1.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
        Public Shared Function D2D1CreateDeviceContext(dxgiSurface As IDXGISurface, ByRef creationProperties As D2D1_CREATION_PROPERTIES, <Out> ByRef d2dDeviceContext As ID2D1DeviceContext) As HRESULT
        End Function

        <DllImport("D2D1.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
        Public Shared Sub D2D1SinCos(angle As Single, <Out> ByRef s As Single, <Out> ByRef c As Single)
        End Sub

        <DllImport("D2D1.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
        Public Shared Function D2D1Tan(angle As Single) As Single
        End Function

        <DllImport("D2D1.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
        Public Shared Function D2D1Vec3Length(x As Single, y As Single, z As Single) As Single
        End Function

        Public Delegate Function PD2D1_EFFECT_FACTORY(<Out> ByRef effectImpl As IntPtr) As HRESULT

        Public Const D2DERR_RECREATE_TARGET As Long = &H8899000CL

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

        Public Shared Function PixelFormat(Optional dxgiFormat As DXGI_FORMAT = DXGI_FORMAT.DXGI_FORMAT_UNKNOWN, Optional alphaMode As D2D1_ALPHA_MODE = D2D1_ALPHA_MODE.D2D1_ALPHA_MODE_UNKNOWN) As D2D1_PIXEL_FORMAT
            Dim lPixelFormat As D2D1_PIXEL_FORMAT
            lPixelFormat.format = dxgiFormat
            lPixelFormat.alphaMode = alphaMode
            Return lPixelFormat
        End Function

        'public D2D1_PIXEL_FORMAT PixelFormat()
        '{
        '    return this.PixelFormat(DXGI_FORMAT.DXGI_FORMAT_UNKNOWN, D2D1_ALPHA_MODE.D2D1_ALPHA_MODE_UNKNOWN);
        '}

        'public D2D1_PIXEL_FORMAT PixelFormat(DXGI_FORMAT dxgiFormat, D2D1_ALPHA_MODE alphaMode)
        '{
        '    D2D1_PIXEL_FORMAT pixelFormat;
        '     pixelFormat.format = dxgiFormat;
        '      pixelFormat.alphaMode = alphaMode;
        '       return pixelFormat;
        '}

        'public D2D1_RENDER_TARGET_PROPERTIES RenderTargetProperties()
        '{
        '    return this.RenderTargetProperties(D2D1_RENDER_TARGET_TYPE.D2D1_RENDER_TARGET_TYPE_DEFAULT, PixelFormat(),
        '        0.0f,
        '        0.0f,
        '        D2D1_RENDER_TARGET_USAGE.D2D1_RENDER_TARGET_USAGE_NONE,
        '        D2D1_FEATURE_LEVEL.D2D1_FEATURE_LEVEL_DEFAULT);
        '}

        'public D2D1_RENDER_TARGET_PROPERTIES RenderTargetProperties(
        '    D2D1_RENDER_TARGET_TYPE type,
        '    D2D1_PIXEL_FORMAT pixelFormat,
        '    float dpiX,
        '    float dpiY,
        '    D2D1_RENDER_TARGET_USAGE usage,
        '    D2D1_FEATURE_LEVEL minLevel
        '    )
        '{
        '    D2D1_RENDER_TARGET_PROPERTIES renderTargetProperties;

        '    renderTargetProperties.type = type;
        '    renderTargetProperties.pixelFormat = pixelFormat;
        '    renderTargetProperties.dpiX = dpiX;
        '    renderTargetProperties.dpiY = dpiY;
        '    renderTargetProperties.usage = usage;
        '    renderTargetProperties.minLevel = minLevel;

        '    return renderTargetProperties;
        '}

        Public Shared Function RenderTargetProperties(type As D2D1_RENDER_TARGET_TYPE, pixelFormat As D2D1_PIXEL_FORMAT, Optional dpiX As Single = 0.0F, Optional dpiY As Single = 0.0F, Optional usage As D2D1_RENDER_TARGET_USAGE = D2D1_RENDER_TARGET_USAGE.D2D1_RENDER_TARGET_USAGE_NONE, Optional minLevel As D2D1_FEATURE_LEVEL = D2D1_FEATURE_LEVEL.D2D1_FEATURE_LEVEL_DEFAULT) As D2D1_RENDER_TARGET_PROPERTIES
            Dim lRenderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES
            lRenderTargetProperties.type = type
            lRenderTargetProperties.pixelFormat = pixelFormat
            lRenderTargetProperties.dpiX = dpiX
            lRenderTargetProperties.dpiY = dpiY
            lRenderTargetProperties.usage = usage
            lRenderTargetProperties.minLevel = minLevel
            Return lRenderTargetProperties
        End Function
        Public Shared Function HwndRenderTargetProperties(hwnd As IntPtr, pixelSize As D2D1_SIZE_U, Optional presentOptions As D2D1_PRESENT_OPTIONS = D2D1_PRESENT_OPTIONS.D2D1_PRESENT_OPTIONS_NONE) As D2D1_HWND_RENDER_TARGET_PROPERTIES
            Dim lHwndRenderTargetProperties As D2D1_HWND_RENDER_TARGET_PROPERTIES
            lHwndRenderTargetProperties.hwnd = hwnd
            lHwndRenderTargetProperties.pixelSize = pixelSize
            lHwndRenderTargetProperties.presentOptions = presentOptions
            Return lHwndRenderTargetProperties
        End Function

        Public Const D3D11_SDK_VERSION As Integer = 7

        <DllImport("D3D11.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
        Public Shared Function D3D11CreateDevice(pAdapter As IDXGIAdapter, DriverType As D3D_DRIVER_TYPE, Software As IntPtr, Flags As UInteger,
              <MarshalAs(UnmanagedType.LPArray)> pFeatureLevels As Integer(), FeatureLevels As UInteger, SDKVersion As UInteger, <Out> ByRef ppDevice As IntPtr, <Out> ByRef pFeatureLevel As D3D_FEATURE_LEVEL, <Out> ByRef ppImmediateContext As ID3D11DeviceContext) As HRESULT
        End Function

        <DllImport("D3D11.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
        Public Shared Function D3D11CreateDeviceAndSwapChain(pAdapter As IDXGIAdapter, DriverType As D3D_DRIVER_TYPE, Software As IntPtr, Flags As UInteger, pFeatureLevels As IntPtr, FeatureLevels As UInteger, SDKVersion As UInteger, ByRef pSwapChainDesc As DXGI_SWAP_CHAIN_DESC, <Out> ByRef ppSwapChain As IDXGISwapChain, <Out> ByRef ppDevice As IntPtr, <Out> ByRef pFeatureLevel As D3D_FEATURE_LEVEL, <Out> ByRef ppImmediateContext As IntPtr) As HRESULT
        End Function

        Public Const DXGI_USAGE_SHADER_INPUT As Integer = &H10
        Public Const DXGI_USAGE_RENDER_TARGET_OUTPUT As Integer = &H20
        Public Const DXGI_USAGE_BACK_BUFFER As Integer = &H40
        Public Const DXGI_USAGE_SHARED As Integer = &H80
        Public Const DXGI_USAGE_READ_ONLY As Integer = &H100
        Public Const DXGI_USAGE_DISCARD_ON_PRESENT As Integer = &H200
        Public Const DXGI_USAGE_UNORDERED_ACCESS As Integer = &H400

        <StructLayout(LayoutKind.Sequential)>
        Public Structure D2D1_CREATION_PROPERTIES
            Private threadingMode As D2D1_THREADING_MODE
            Private debugLevel As D2D1_DEBUG_LEVEL
            Private options As D2D1_DEVICE_CONTEXT_OPTIONS
        End Structure

        Public Enum D2D1_THREADING_MODE As UInteger
            ''' <summary>
            ''' Resources may only be invoked serially.  Reference counts on resources are
            ''' interlocked, however, resource and render target state is not protected from
            ''' multi-threaded access
            ''' </summary>
            D2D1_THREADING_MODE_SINGLE_THREADED = D2D1_FACTORY_TYPE.D2D1_FACTORY_TYPE_SINGLE_THREADED
            ''' <summary>
            ''' Resources may be invoked from multiple threads. Resources use interlocked
            ''' reference counting and their state is protected.
            ''' </summary>
            D2D1_THREADING_MODE_MULTI_THREADED = D2D1_FACTORY_TYPE.D2D1_FACTORY_TYPE_MULTI_THREADED
            D2D1_THREADING_MODE_FORCE_DWORD = &HFFFFFFFFUI
        End Enum

        <DllImport("D2D1.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
        Public Shared Function D2D1ComputeMaximumScaleFactor(ByRef matrix As D2D1_MATRIX_3X2_F) As Single
        End Function

        ' 
        ' Set to alignedByteOffset within D2D1_INPUT_ELEMENT_DESC for elements that 
        ' immediately follow preceding elements in memory
        '
        Public Const D2D1_APPEND_ALIGNED_ELEMENT As UInteger = &HFFFFFFFFUI

        Public Shared Sub SetInputEffect(inputEffectOrig As ID2D1Effect, index As UInteger, inputEffect As ID2D1Effect, Optional invalidate As Boolean = True)
            Dim output As ID2D1Image = Nothing
            If inputEffect IsNot Nothing Then
                inputEffect.GetOutput(output)
            End If
            inputEffectOrig.SetInput(index, output, invalidate)
            If output IsNot Nothing Then
                Marshal.ReleaseComObject(output)
            End If
        End Sub

        Public Shared Function FloatMax() As Single
            Return Single.MaxValue
        End Function

        Public Shared Function Point2F(Optional x As Single = 0.0F, Optional y As Single = 0.0F) As D2D1_POINT_2F
            Return New D2D1_POINT_2F With {
                .x = x,
                .y = y
            }
        End Function

        Public Shared Function SizeF(Optional width As Single = 0.0F, Optional height As Single = 0.0F) As D2D1_SIZE_F
            Return New D2D1_SIZE_F With {
                .width = width,
                .height = height
            }
        End Function

        Public Shared Function SizeU(Optional width As UInteger = 0, Optional height As UInteger = 0) As D2D1_SIZE_U
            Return New D2D1_SIZE_U With {
                .width = width,
                .height = height
            }
        End Function

        Public Shared Function RectF(Optional left As Single = 0.0F, Optional top As Single = 0.0F, Optional right As Single = 0.0F, Optional bottom As Single = 0.0F) As D2D1_RECT_F
            Return New D2D1_RECT_F With {
                .left = left,
                .top = top,
                .right = right,
                .bottom = bottom
            }
        End Function

        Public Shared Function RectU(Optional left As UInteger = 0, Optional top As UInteger = 0, Optional right As UInteger = 0, Optional bottom As UInteger = 0) As D2D1_RECT_U
            Return New D2D1_RECT_U With {
                .left = left,
                .top = top,
                .right = right,
                .bottom = bottom
            }
        End Function

        Public Shared Function RoundedRect(rect As D2D1_RECT_F, radiusX As Single, radiusY As Single) As D2D1_ROUNDED_RECT
            Return New D2D1_ROUNDED_RECT With {
                .rect = rect,
                .radiusX = radiusX,
                .radiusY = radiusY
            }
        End Function

        Public Shared Function IdentityMatrix() As D2D1_MATRIX_3X2_F
            Return New Matrix3x2F(1, 0, 0, 1, 0, 0)
        End Function

        Public Shared Function InfiniteRect() As D2D1_RECT_F
            Return New D2D1_RECT_F With {
                .left = -FloatMax(),
                .top = -FloatMax(),
                .right = FloatMax(),
                .bottom = FloatMax()
            }
        End Function

        Public Shared Function LayerParameters(Optional contentBounds As D2D1_RECT_F = Nothing, Optional geometricMask As ID2D1Geometry = Nothing, Optional maskAntialiasMode As D2D1_ANTIALIAS_MODE = D2D1_ANTIALIAS_MODE.D2D1_ANTIALIAS_MODE_PER_PRIMITIVE, Optional maskTransform As D2D1_MATRIX_3X2_F = Nothing, Optional opacity As Single = 1.0F, Optional opacityBrush As IntPtr = Nothing, Optional layerOptions As D2D1_LAYER_OPTIONS = D2D1_LAYER_OPTIONS.D2D1_LAYER_OPTIONS_NONE) As D2D1_LAYER_PARAMETERS
            If contentBounds.Equals(Nothing) Then
                contentBounds = InfiniteRect()
            End If

            'if (maskTransform.Equals(default(Matrix3x2F)))
            If maskTransform Is Nothing Then
                maskTransform = IdentityMatrix()
            End If

            Return New D2D1_LAYER_PARAMETERS With {
    .contentBounds = contentBounds,
    .geometricMask = geometricMask,
    .maskAntialiasMode = maskAntialiasMode,
    .maskTransform = maskTransform,
    .opacity = opacity,
    .opacityBrush = opacityBrush,
    .layerOptions = layerOptions
}
        End Function

        Public Shared Function BitmapBrushProperties(Optional extendModeX As D2D1_EXTEND_MODE = D2D1_EXTEND_MODE.D2D1_EXTEND_MODE_CLAMP, Optional extendModeY As D2D1_EXTEND_MODE = D2D1_EXTEND_MODE.D2D1_EXTEND_MODE_CLAMP, Optional interpolationMode As D2D1_BITMAP_INTERPOLATION_MODE = D2D1_BITMAP_INTERPOLATION_MODE.D2D1_BITMAP_INTERPOLATION_MODE_LINEAR) As D2D1_BITMAP_BRUSH_PROPERTIES
            Return New D2D1_BITMAP_BRUSH_PROPERTIES With {
    .extendModeX = extendModeX,
    .extendModeY = extendModeY,
    .interpolationMode = interpolationMode
}
        End Function

        Public Shared Function BrushProperties(Optional opacity As Single = 1.0F, Optional transform As D2D1_MATRIX_3X2_F = Nothing) As D2D1_BRUSH_PROPERTIES
            If transform Is Nothing Then
                transform = IdentityMatrix()
            End If

            Return New D2D1_BRUSH_PROPERTIES With {
                .opacity = opacity,
                .transform = transform
            }
        End Function

        Public Shared Function Ellipse(center As D2D1_POINT_2F, radiusX As Single, radiusY As Single) As D2D1_ELLIPSE
            Return New D2D1_ELLIPSE With {
                .point = center,
                .radiusX = radiusX,
                .radiusY = radiusY
            }
        End Function

        Public Shared Function GradientStop(position As Single, color As D2D1_COLOR_F) As D2D1_GRADIENT_STOP
            Return New D2D1_GRADIENT_STOP With {
                .position = position,
                .color = color
            }
        End Function

        Public Shared Function QuadraticBezierSegment(point1 As D2D1_POINT_2F, point2 As D2D1_POINT_2F) As D2D1_QUADRATIC_BEZIER_SEGMENT
            Return New D2D1_QUADRATIC_BEZIER_SEGMENT With {
                .point1 = point1,
                .point2 = point2
            }
        End Function
    End Class

    Public Class ColorF
        Inherits D2D1_COLOR_F
        Public Enum [Enum]
            AliceBlue = &HF0F8FF
            AntiqueWhite = &HFAEBD7
            Aqua = &HFFFF
            Aquamarine = &H7FFFD4
            Azure = &HF0FFFF
            Beige = &HF5F5DC
            Bisque = &HFFE4C4
            Black = &H0
            BlanchedAlmond = &HFFEBCD
            Blue = &HFF
            BlueViolet = &H8A2BE2
            Brown = &HA52A2A
            BurlyWood = &HDEB887
            CadetBlue = &H5F9EA0
            Chartreuse = &H7FFF00
            Chocolate = &HD2691E
            Coral = &HFF7F50
            CornflowerBlue = &H6495ED
            Cornsilk = &HFFF8DC
            Crimson = &HDC143C
            Cyan = &HFFFF
            DarkBlue = &H8B
            DarkCyan = &H8B8B
            DarkGoldenrod = &HB8860B
            DarkGray = &HA9A9A9
            DarkGreen = &H6400
            DarkKhaki = &HBDB76B
            DarkMagenta = &H8B008B
            DarkOliveGreen = &H556B2F
            DarkOrange = &HFF8C00
            DarkOrchid = &H9932CC
            DarkRed = &H8B0000
            DarkSalmon = &HE9967A
            DarkSeaGreen = &H8FBC8F
            DarkSlateBlue = &H483D8B
            DarkSlateGray = &H2F4F4F
            DarkTurquoise = &HCED1
            DarkViolet = &H9400D3
            DeepPink = &HFF1493
            DeepSkyBlue = &HBFFF
            DimGray = &H696969
            DodgerBlue = &H1E90FF
            Firebrick = &HB22222
            FloralWhite = &HFFFAF0
            ForestGreen = &H228B22
            Fuchsia = &HFF00FF
            Gainsboro = &HDCDCDC
            GhostWhite = &HF8F8FF
            Gold = &HFFD700
            Goldenrod = &HDAA520
            Gray = &H808080
            Green = &H8000
            GreenYellow = &HADFF2F
            Honeydew = &HF0FFF0
            HotPink = &HFF69B4
            IndianRed = &HCD5C5C
            Indigo = &H4B0082
            Ivory = &HFFFFF0
            Khaki = &HF0E68C
            Lavender = &HE6E6FA
            LavenderBlush = &HFFF0F5
            LawnGreen = &H7CFC00
            LemonChiffon = &HFFFACD
            LightBlue = &HADD8E6
            LightCoral = &HF08080
            LightCyan = &HE0FFFF
            LightGoldenrodYellow = &HFAFAD2
            LightGreen = &H90EE90
            LightGray = &HD3D3D3
            LightPink = &HFFB6C1
            LightSalmon = &HFFA07A
            LightSeaGreen = &H20B2AA
            LightSkyBlue = &H87CEFA
            LightSlateGray = &H778899
            LightSteelBlue = &HB0C4DE
            LightYellow = &HFFFFE0
            Lime = &HFF00
            LimeGreen = &H32CD32
            Linen = &HFAF0E6
            Magenta = &HFF00FF
            Maroon = &H800000
            MediumAquamarine = &H66CDAA
            MediumBlue = &HCD
            MediumOrchid = &HBA55D3
            MediumPurple = &H9370DB
            MediumSeaGreen = &H3CB371
            MediumSlateBlue = &H7B68EE
            MediumSpringGreen = &HFA9A
            MediumTurquoise = &H48D1CC
            MediumVioletRed = &HC71585
            MidnightBlue = &H191970
            MintCream = &HF5FFFA
            MistyRose = &HFFE4E1
            Moccasin = &HFFE4B5
            NavajoWhite = &HFFDEAD
            Navy = &H80
            OldLace = &HFDF5E6
            Olive = &H808000
            OliveDrab = &H6B8E23
            Orange = &HFFA500
            OrangeRed = &HFF4500
            Orchid = &HDA70D6
            PaleGoldenrod = &HEEE8AA
            PaleGreen = &H98FB98
            PaleTurquoise = &HAFEEEE
            PaleVioletRed = &HDB7093
            PapayaWhip = &HFFEFD5
            PeachPuff = &HFFDAB9
            Peru = &HCD853F
            Pink = &HFFC0CB
            Plum = &HDDA0DD
            PowderBlue = &HB0E0E6
            Purple = &H800080
            Red = &HFF0000
            RosyBrown = &HBC8F8F
            RoyalBlue = &H4169E1
            SaddleBrown = &H8B4513
            Salmon = &HFA8072
            SandyBrown = &HF4A460
            SeaGreen = &H2E8B57
            SeaShell = &HFFF5EE
            Sienna = &HA0522D
            Silver = &HC0C0C0
            SkyBlue = &H87CEEB
            SlateBlue = &H6A5ACD
            SlateGray = &H708090
            Snow = &HFFFAFA
            SpringGreen = &HFF7F
            SteelBlue = &H4682B4
            Tan = &HD2B48C
            Teal = &H8080
            Thistle = &HD8BFD8
            Tomato = &HFF6347
            Turquoise = &H40E0D0
            Violet = &HEE82EE
            Wheat = &HF5DEB3
            White = &HFFFFFF
            WhiteSmoke = &HF5F5F5
            Yellow = &HFFFF00
            YellowGreen = &H9ACD32
        End Enum

        '
        ' Construct a color, note that the alpha value from the "rgb" component
        ' is never used.
        ' 

        Public Sub New(rgb As UInteger, Optional a As Single = 1.0F)
            Init(rgb, a)
        End Sub

        Public Sub New(knownColor As [Enum], Optional a As Single = 1.0F)
            Init(knownColor, a)
        End Sub
        Public Sub New(red As Single, green As Single, blue As Single, Optional alpha As Single = 1.0F)
            r = red
            g = green
            b = blue
            a = alpha
        End Sub
        Private Sub Init(rgb As UInteger, alpha As Single)
            r = ((rgb And sc_redMask) >> sc_redShift) / 255.0F
            g = ((rgb And sc_greenMask) >> sc_greenShift) / 255.0F
            b = ((rgb And sc_blueMask) >> sc_blueShift) / 255.0F
            a = alpha
        End Sub
        Const sc_redShift As UInteger = 16
        Const sc_greenShift As UInteger = 8
        Const sc_blueShift As UInteger = 0

        Const sc_redMask As UInteger = &HFF << sc_redShift
        Const sc_greenMask As UInteger = &HFF << sc_greenShift
        Const sc_blueMask As UInteger = &HFF << sc_blueShift
    End Class


    Public Class Matrix3x2F
        Inherits D2D1_MATRIX_3X2_F
        Public Sub New(m11 As Single, m12 As Single, m21 As Single, m22 As Single, m31 As Single, m32 As Single)
            _11 = m11
            _12 = m12
            _21 = m21
            _22 = m22
            _31 = m31
            _32 = m32
        End Sub

        Public Sub New()
        End Sub

        Public Shared Function Identity() As Matrix3x2F
            Return New Matrix3x2F(1.0F, 0F, 0F, 1.0F, 0F, 0F)
        End Function

        <StructLayout(LayoutKind.Sequential, Pack:=4)>
        Public Structure Matrix3x2
            ''' <summary>
            ''' Gets the identity matrix.
            ''' </summary>
            Public Shared ReadOnly Identity As Matrix3x2 = New Matrix3x2(1, 0, 0, 1, 0, 0)

            ''' <summary>
            ''' Element (1,1)
            ''' </summary>
            Public M11 As Single

            ''' <summary>
            ''' Element (1,2)
            ''' </summary>
            Public M12 As Single

            ''' <summary>
            ''' Element (2,1)
            ''' </summary>
            Public M21 As Single

            ''' <summary>
            ''' Element (2,2)
            ''' </summary>
            Public M22 As Single

            ''' <summary>
            ''' Element (3,1)
            ''' </summary>
            Public M31 As Single

            ''' <summary>
            ''' Element (3,2)
            ''' </summary>
            Public M32 As Single

            ''' <summary>
            ''' Initializes a new instance of the <see cref="Matrix3x2"/> struct.
            ''' </summary>
            ''' <param name="value">The value that will be assigned to all components.</param>
            Public Sub New(ByVal value As Single)
                Me.M32 = value
                Me.M31 = value
                Me.M22 = value
                Me.M21 = value
                Me.M12 = value
                Me.M11 = value
            End Sub

            ''' <summary>
            ''' Initializes a new instance of the <see cref="Matrix3x2"/> struct.
            ''' </summary>
            ''' <param name="M11">The value to assign at row 1 column 1 of the matrix.</param>
            ''' <param name="M12">The value to assign at row 1 column 2 of the matrix.</param>
            ''' <param name="M21">The value to assign at row 2 column 1 of the matrix.</param>
            ''' <param name="M22">The value to assign at row 2 column 2 of the matrix.</param>
            ''' <param name="M31">The value to assign at row 3 column 1 of the matrix.</param>
            ''' <param name="M32">The value to assign at row 3 column 2 of the matrix.</param>
            Public Sub New(M11 As Single, M12 As Single, M21 As Single, M22 As Single, M31 As Single, M32 As Single)
                Me.M11 = M11
                Me.M12 = M12
                Me.M21 = M21
                Me.M22 = M22
                Me.M31 = M31
                Me.M32 = M32
            End Sub

            ''' <summary>
            ''' Initializes a new instance of the <see cref="Matrix3x2"/> struct.
            ''' </summary>
            ''' <param name="values">The values to assign to the components of the matrix. This must be an array with six elements.</param>
            ''' <exception cref="ArgumentNullException">Thrown when <paramref name="values"/> is <c>null</c>.</exception>
            ''' <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="values"/> contains more or less than six elements.</exception>
            Public Sub New(values As Single())
                If values Is Nothing Then Throw New ArgumentNullException("values")
                If values.Length <> 6 Then Throw New ArgumentOutOfRangeException("values", "There must be six input values for Matrix3x2.")

                M11 = values(0)
                M12 = values(1)

                M21 = values(2)
                M22 = values(3)

                M31 = values(4)
                M32 = values(5)
            End Sub

            Private Class CSharpImpl
                <Obsolete("Please refactor calling code to use normal Visual Basic assignment")>
                Shared Function __Assign(Of T)(ByRef target As T, value As T) As T
                    target = value
                    Return value
                End Function
            End Class
        End Structure

        'public static void MakeRotateMatrix(float angle, D2D1_POINT_2F center, out Matrix3x2 matrix)
        '{
        '    unsafe
        '    {
        '        IntPtr pRotation = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(D2D1_MATRIX_3X2_F)));
        '        D2D1_MATRIX_3X2_F rotation = new D2D1_MATRIX_3X2_F();
        '        D2DTools.D2D1MakeRotateMatrix_(angle, center,(void*) pRotation);
        '        rotation = (D2D1_MATRIX_3X2_F)Marshal.PtrToStructure(pRotation, typeof(D2D1_MATRIX_3X2_F));

        '        matrix = new Matrix3x2();
        '        fixed (void* matrix_ = &matrix)
        '            D2DTools.D2D1MakeRotateMatrix_(angle, center, matrix_);               
        '    }
        '}

        Public Shared Function MakeRotateMatrix(ByVal angle As Single, Optional ByVal center As D2D1_POINT_2F = Nothing) As Matrix3x2
            Dim pRotation As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(GetType(D2D1_MATRIX_3X2_F)))
            Dim rotation As D2D1_MATRIX_3X2_F = New D2D1_MATRIX_3X2_F
            D2DTools.D2D1MakeRotateMatrix_(angle, center, pRotation)
            rotation = CType(Marshal.PtrToStructure(pRotation, GetType(D2D1_MATRIX_3X2_F)), D2D1_MATRIX_3X2_F)
            Dim matrix As Matrix3x2 = New Matrix3x2()
            Dim matrix_ As IntPtr = IntPtr.Zero
            Marshal.StructureToPtr(matrix, matrix_, True)
            D2DTools.D2D1MakeRotateMatrix_(angle, center, matrix_)
            Return matrix
        End Function

        Public Shared Function Rotation(ByVal angle As Single, Optional ByVal center As D2D1_POINT_2F = Nothing) As Matrix3x2F
            Dim rotationMatrix As D2D1_MATRIX_3X2_F = New D2D1_MATRIX_3X2_F()
            Dim pRotation As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(GetType(D2D1_MATRIX_3X2_F)))
            'Marshal.StructureToPtr(rotation, pRotation, false);
            'D2D1_MATRIX_3X2_F_BIS rot2 = new D2D1_MATRIX_3X2_F_BIS();
            D2DTools.D2D1MakeRotateMatrix_(angle, center, pRotation)
            'IntPtr p = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(D2D1_MATRIX_3X2_F_BIS)));
            'D2DTools.D2D1MakeRotateMatrix(angle, center, out pRotation);
            ' D2DTools.D2D1MakeRotateMatrix(angle, center, out p);
            'rot2 = (D2D1_MATRIX_3X2_F_BIS)Marshal.PtrToStructure(p, typeof(D2D1_MATRIX_3X2_F_BIS));
            'var ItemIDList = (D2D1_MATRIX_3X2_F)Marshal.PtrToStructure(pRotation, typeof(D2D1_MATRIX_3X2_F));
            rotationMatrix = CType(Marshal.PtrToStructure(pRotation, GetType(D2D1_MATRIX_3X2_F)), D2D1_MATRIX_3X2_F)
            Dim m As Matrix3x2F = New Matrix3x2F(rotationMatrix._11, rotationMatrix._12, rotationMatrix._21, rotationMatrix._22, rotationMatrix._31, rotationMatrix._32)
            Marshal.FreeHGlobal(pRotation)
            'D2DTools.D2D1MakeRotateMatrix(angle, center, out rotation);    
            Return m
        End Function

        Public Shared Function Skew(angleX As Single, angleY As Single, Optional center As D2D1_POINT_2F = Nothing) As Matrix3x2F
            Dim tanX As Single = Math.Tan(angleX * CSng(Math.PI) / 180.0F)
            Dim tanY As Single = Math.Tan(angleY * CSng(Math.PI) / 180.0F)

            Return New Matrix3x2F(1.0F, tanY, tanX, 1.0F, -tanX * center.y, -tanY * center.x)
        End Function

        Public Overloads Shared Function Scale(ByVal size As D2D1_SIZE_F, Optional ByVal center As D2D1_POINT_2F = Nothing) As Matrix3x2F
            Dim scaleMatrix As Matrix3x2F = New Matrix3x2F()
            'Matrix3x2F scale = null;
            scaleMatrix._11 = size.width
            scaleMatrix._12 = 0!
            scaleMatrix._21 = 0!
            scaleMatrix._22 = size.height
            scaleMatrix._31 = (center.x - (size.width * center.x))
            scaleMatrix._32 = (center.y - (size.height * center.y))
            Return scaleMatrix
        End Function

        Public Overloads Shared Function Scale(ByVal x As Single, ByVal y As Single, Optional ByVal center As D2D1_POINT_2F = Nothing) As Matrix3x2F
            Return Matrix3x2F.Scale(New D2D1_SIZE_F(x, y), center)
        End Function

        Public Shared Function Translation(size As D2D1_SIZE_F) As Matrix3x2F
            Dim lTranslation As Matrix3x2F = New Matrix3x2F()
            lTranslation._11 = 1.0F
            lTranslation._12 = 0.0F
            lTranslation._21 = 0.0F
            lTranslation._22 = 1.0F
            lTranslation._31 = size.width
            lTranslation._32 = size.height
            Return lTranslation
        End Function

        Public Shared Function Translation(x As Single, y As Single) As Matrix3x2F
            Return Translation(New D2D1_SIZE_F(x, y))
        End Function

        Public Function Determinant() As Single
            Return _11 * _22 - _12 * _21
        End Function

        Public Function IsInvertible() As Boolean
            Return Determinant() <> 0
        End Function

        Public Function Invert() As Boolean
            Dim det As Single = Determinant()
            If det = 0 Then
                Return False
            End If

            Dim invDet = 1.0F / det
            Dim m11 = _22 * invDet
            Dim m12 = -_12 * invDet
            Dim m21 = -_21 * invDet
            Dim m22 = _11 * invDet
            Dim m31 = (_21 * _32 - _22 * _31) * invDet
            Dim m32 = (_12 * _31 - _11 * _32) * invDet

            _11 = m11
            _12 = m12
            _21 = m21
            _22 = m22
            _31 = m31
            _32 = m32

            Return True
        End Function

        Public Function IsIdentity() As Boolean
            Return _11 = 1.0F AndAlso _12 = 0F AndAlso _21 = 0F AndAlso _22 = 1.0F AndAlso _31 = 0F AndAlso _32 = 0F
        End Function

        Public Sub SetProduct(a As Matrix3x2F, b As Matrix3x2F)
            _11 = a._11 * b._11 + a._12 * b._21
            _12 = a._11 * b._12 + a._12 * b._22
            _21 = a._21 * b._11 + a._22 * b._21
            _22 = a._21 * b._12 + a._22 * b._22
            _31 = a._31 * b._11 + a._32 * b._21 + b._31
            _32 = a._31 * b._12 + a._32 * b._22 + b._32
        End Sub

        Public Shared Operator *(b As Matrix3x2F, c As Matrix3x2F) As Matrix3x2F
            Dim result As Matrix3x2F = New Matrix3x2F()
            result.SetProduct(b, c)
            Return result
        End Operator

        Public Function TransformPoint(point As D2D1_POINT_2F) As D2D1_POINT_2F
            Return New D2D1_POINT_2F(point.x * _11 + point.y * _21 + _31, point.x * _12 + point.y * _22 + _32)
        End Function
    End Class

    <ComImport>
    <Guid("6f15aaf2-d208-4e89-9ab4-489535d34f9c")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID3D11Texture2D
        Inherits ID3D11Resource
#Region "ID3D11Resource"
#Region "ID3D11DeviceChild"
        <PreserveSig>
        Overloads Sub GetDevice(<Out> ByRef ppDevice As IntPtr)
        <PreserveSig>
        Overloads Function GetPrivateData(ByRef guid As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateData(ByRef guid As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateDataInterface(ByRef guid As Guid, pData As IntPtr) As HRESULT
#End Region

        <PreserveSig>
        Overloads Sub [GetType](<Out> ByRef pResourceDimension As D3D11_RESOURCE_DIMENSION)
        <PreserveSig>
        Overloads Sub SetEvictionPriority(EvictionPriority As UInteger)
        <PreserveSig>
        Overloads Function GetEvictionPriority() As UInteger
#End Region

        <PreserveSig>
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
    <Guid("dc8e63f3-d12b-4952-b47b-5e45026a862d")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID3D11Resource
        Inherits ID3D11DeviceChild
#Region "ID3D11DeviceChild"
        <PreserveSig>
        Overloads Sub GetDevice(<Out> ByRef ppDevice As IntPtr)
        <PreserveSig>
        Overloads Function GetPrivateData(ByRef guid As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateData(ByRef guid As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateDataInterface(ByRef guid As Guid, pData As IntPtr) As HRESULT
#End Region

        <PreserveSig>
        Sub [GetType](<Out> ByRef pResourceDimension As D3D11_RESOURCE_DIMENSION)
        <PreserveSig>
        Sub SetEvictionPriority(EvictionPriority As UInteger)
        <PreserveSig>
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
    <Guid("1841e5c8-16b0-489b-bcc8-44cfb0d5deae")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID3D11DeviceChild
        <PreserveSig>
        Sub GetDevice(<Out> ByRef ppDevice As IntPtr)
        <PreserveSig>
        Function GetPrivateData(ByRef guid As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        <PreserveSig>
        Function SetPrivateData(ByRef guid As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        <PreserveSig>
        Function SetPrivateDataInterface(ByRef guid As Guid, pData As IntPtr) As HRESULT
    End Interface

    <ComImport>
    <Guid("c0bfa96c-e089-44fb-8eaf-26f8796190da")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID3D11DeviceContext
        Inherits ID3D11DeviceChild
#Region "ID3D11DeviceChild"
        <PreserveSig>
        Overloads Sub GetDevice(<Out> ByRef ppDevice As IntPtr)
        <PreserveSig>
        Overloads Function GetPrivateData(ByRef guid As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateData(ByRef guid As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateDataInterface(ByRef guid As Guid, pData As IntPtr) As HRESULT
#End Region

        <PreserveSig>
        Sub VSSetConstantBuffers(StartSlot As UInteger, NumBuffers As UInteger, ppConstantBuffers As ID3D11Buffer)
        <PreserveSig>
        Sub PSSetShaderResources(StartSlot As UInteger, NumViews As UInteger, ppShaderResourceViews As ID3D11ShaderResourceView)
        <PreserveSig>
        Sub PSSetShader(pPixelShader As ID3D11PixelShader, ppClassInstances As ID3D11ClassInstance, NumClassInstances As UInteger)
        <PreserveSig>
        Sub PSSetSamplers(StartSlot As UInteger, NumSamplers As UInteger, ppSamplers As ID3D11SamplerState)
        <PreserveSig>
        Sub VSSetShader(pVertexShader As ID3D11VertexShader, ppClassInstances As ID3D11ClassInstance, NumClassInstances As UInteger)
        <PreserveSig>
        Sub DrawIndexed(IndexCount As UInteger, StartIndexLocation As UInteger, BaseVertexLocation As Integer)
        <PreserveSig>
        Sub Draw(VertexCount As UInteger, StartVertexLocation As UInteger)
        <PreserveSig>
        Function Map(pResource As ID3D11Resource, Subresource As UInteger, MapType As D3D11_MAP, MapFlags As UInteger, <Out> ByRef pMappedResource As D3D11_MAPPED_SUBRESOURCE) As HRESULT
        <PreserveSig>
        Sub Unmap(pResource As ID3D11Resource, Subresource As UInteger)
        <PreserveSig>
        Sub PSSetConstantBuffers(StartSlot As UInteger, NumBuffers As UInteger, ppConstantBuffers As ID3D11Buffer)
        <PreserveSig>
        Sub IASetInputLayout(pInputLayout As ID3D11InputLayout)
        <PreserveSig>
        Sub IASetVertexBuffers(StartSlot As UInteger, NumBuffers As UInteger, ppVertexBuffers As ID3D11Buffer, pStrides As UInteger, pOffsets As UInteger)
        <PreserveSig>
        Sub IASetIndexBuffer(pIndexBuffer As ID3D11Buffer, Format As DXGI_FORMAT, Offset As UInteger)
        <PreserveSig>
        Sub DrawIndexedInstanced(IndexCountPerInstance As UInteger, InstanceCount As UInteger, StartIndexLocation As UInteger, BaseVertexLocation As Integer, StartInstanceLocation As UInteger)
        <PreserveSig>
        Sub DrawInstanced(VertexCountPerInstance As UInteger, InstanceCount As UInteger, StartVertexLocation As UInteger, StartInstanceLocation As UInteger)
        <PreserveSig>
        Sub GSSetConstantBuffers(StartSlot As UInteger, NumBuffers As UInteger, ppConstantBuffers As ID3D11Buffer)
        <PreserveSig>
        Sub GSSetShader(pShader As ID3D11GeometryShader, ppClassInstances As ID3D11ClassInstance, NumClassInstances As UInteger)
        <PreserveSig>
        Sub IASetPrimitiveTopology(Topology As D3D_PRIMITIVE_TOPOLOGY)
        <PreserveSig>
        Sub VSSetShaderResources(StartSlot As UInteger, NumViews As UInteger, ppShaderResourceViews As ID3D11ShaderResourceView)
        <PreserveSig>
        Sub VSSetSamplers(StartSlot As UInteger, NumSamplers As UInteger, ppSamplers As ID3D11SamplerState)
        <PreserveSig>
        Sub Begin(pAsync As ID3D11Asynchronous)
        <PreserveSig>
        Sub [End](pAsync As ID3D11Asynchronous)
        <PreserveSig>
        Function GetData(pAsync As ID3D11Asynchronous, <Out> ByRef pData As IntPtr, DataSize As UInteger, GetDataFlags As UInteger) As HRESULT
        <PreserveSig>
        Sub SetPredication(pPredicate As ID3D11Predicate, PredicateValue As Boolean)
        <PreserveSig>
        Sub GSSetShaderResources(StartSlot As UInteger, NumViews As UInteger, ppShaderResourceViews As ID3D11ShaderResourceView)
        <PreserveSig>
        Sub GSSetSamplers(StartSlot As UInteger, NumSamplers As UInteger, ppSamplers As ID3D11SamplerState)
        <PreserveSig>
        Sub OMSetRenderTargets(NumViews As UInteger, ppRenderTargetViews As ID3D11RenderTargetView, pDepthStencilView As ID3D11DepthStencilView)
        <PreserveSig>
        Sub OMSetRenderTargetsAndUnorderedAccessViews(NumRTVs As UInteger, ppRenderTargetViews As ID3D11RenderTargetView, pDepthStencilView As ID3D11DepthStencilView, UAVStartSlot As UInteger, NumUAVs As UInteger, ppUnorderedAccessViews As ID3D11UnorderedAccessView, pUAVInitialCounts As UInteger)
        <PreserveSig>
        Sub OMSetBlendState(pBlendState As ID3D11BlendState, BlendFactor As Single(), SampleMask As UInteger)
        <PreserveSig>
        Sub OMSetDepthStencilState(pDepthStencilState As ID3D11DepthStencilState, StencilRef As UInteger)
        <PreserveSig>
        Sub SOSetTargets(NumBuffers As UInteger, ppSOTargets As ID3D11Buffer, pOffsets As UInteger)
        <PreserveSig>
        Sub DrawAuto()
        <PreserveSig>
        Sub DrawIndexedInstancedIndirect(pBufferForArgs As ID3D11Buffer, AlignedByteOffsetForArgs As UInteger)
        <PreserveSig>
        Sub DrawInstancedIndirect(pBufferForArgs As ID3D11Buffer, AlignedByteOffsetForArgs As UInteger)
        <PreserveSig>
        Sub Dispatch(ThreadGroupCountX As UInteger, ThreadGroupCountY As UInteger, ThreadGroupCountZ As UInteger)
        <PreserveSig>
        Sub DispatchIndirect(pBufferForArgs As ID3D11Buffer, AlignedByteOffsetForArgs As UInteger)
        <PreserveSig>
        Sub RSSetState(pRasterizerState As ID3D11RasterizerState)
        <PreserveSig>
        Sub RSSetViewports(NumViewports As UInteger, pViewports As D3D11_VIEWPORT)
        <PreserveSig>
        Sub RSSetScissorRects(NumRects As UInteger, pRects As RECT)
        <PreserveSig>
        Sub CopySubresourceRegion(pDstResource As ID3D11Resource, DstSubresource As UInteger, DstX As UInteger, DstY As UInteger, DstZ As UInteger, pSrcResource As ID3D11Resource, SrcSubresource As UInteger, pSrcBox As D3D11_BOX)
        <PreserveSig>
        Sub CopyResource(pDstResource As ID3D11Resource, pSrcResource As ID3D11Resource)
        <PreserveSig>
        Sub UpdateSubresource(pDstResource As ID3D11Resource, DstSubresource As UInteger, pDstBox As D3D11_BOX, pSrcData As IntPtr, SrcRowPitch As UInteger, SrcDepthPitch As UInteger)
        <PreserveSig>
        Sub CopyStructureCount(pDstBuffer As ID3D11Buffer, DstAlignedByteOffset As UInteger, pSrcView As ID3D11UnorderedAccessView)
        <PreserveSig>
        Sub ClearRenderTargetView(pRenderTargetView As ID3D11RenderTargetView, ColorRGBA As Single())
        <PreserveSig>
        Sub ClearUnorderedAccessViewuint(pUnorderedAccessView As ID3D11UnorderedAccessView, Values As UInteger())
        <PreserveSig>
        Sub ClearUnorderedAccessViewfloat(pUnorderedAccessView As ID3D11UnorderedAccessView, Values As Single())
        <PreserveSig>
        Sub ClearDepthStencilView(pDepthStencilView As ID3D11DepthStencilView, ClearFlags As UInteger, Depth As Single, Stencil As Byte)
        <PreserveSig>
        Sub GenerateMips(pShaderResourceView As ID3D11ShaderResourceView)
        <PreserveSig>
        Sub SetResourceMinLOD(pResource As ID3D11Resource, MinLOD As Single)
        <PreserveSig>
        Function GetResourceMinLOD(pResource As ID3D11Resource) As Single
        <PreserveSig>
        Sub ResolveSubresource(pDstResource As ID3D11Resource, DstSubresource As UInteger, pSrcResource As ID3D11Resource, SrcSubresource As UInteger, Format As DXGI_FORMAT)
        <PreserveSig>
        Sub ExecuteCommandList(pCommandList As ID3D11CommandList, RestoreContextState As Boolean)
        <PreserveSig>
        Sub HSSetShaderResources(StartSlot As UInteger, NumViews As UInteger, ppShaderResourceViews As ID3D11ShaderResourceView)
        <PreserveSig>
        Sub HSSetShader(pHullShader As ID3D11HullShader, ppClassInstances As ID3D11ClassInstance, NumClassInstances As UInteger)
        <PreserveSig>
        Sub HSSetSamplers(StartSlot As UInteger, NumSamplers As UInteger, ppSamplers As ID3D11SamplerState)
        <PreserveSig>
        Sub HSSetConstantBuffers(StartSlot As UInteger, NumBuffers As UInteger, ppConstantBuffers As ID3D11Buffer)
        <PreserveSig>
        Sub DSSetShaderResources(StartSlot As UInteger, NumViews As UInteger, ppShaderResourceViews As ID3D11ShaderResourceView)
        <PreserveSig>
        Sub DSSetShader(pDomainShader As ID3D11DomainShader, ppClassInstances As ID3D11ClassInstance, NumClassInstances As UInteger)
        <PreserveSig>
        Sub DSSetSamplers(StartSlot As UInteger, NumSamplers As UInteger, ppSamplers As ID3D11SamplerState)
        <PreserveSig>
        Sub DSSetConstantBuffers(StartSlot As UInteger, NumBuffers As UInteger, ppConstantBuffers As ID3D11Buffer)
        <PreserveSig>
        Sub CSSetShaderResources(StartSlot As UInteger, NumViews As UInteger, ppShaderResourceViews As ID3D11ShaderResourceView)
        <PreserveSig>
        Sub CSSetUnorderedAccessViews(StartSlot As UInteger, NumUAVs As UInteger, ppUnorderedAccessViews As ID3D11UnorderedAccessView, pUAVInitialCounts As UInteger)
        <PreserveSig>
        Sub CSSetShader(pComputeShader As ID3D11ComputeShader, ppClassInstances As ID3D11ClassInstance, NumClassInstances As UInteger)
        <PreserveSig>
        Sub CSSetSamplers(StartSlot As UInteger, NumSamplers As UInteger, ppSamplers As ID3D11SamplerState)
        <PreserveSig>
        Sub CSSetConstantBuffers(StartSlot As UInteger, NumBuffers As UInteger, ppConstantBuffers As ID3D11Buffer)
        <PreserveSig>
        Sub VSGetConstantBuffers(StartSlot As UInteger, NumBuffers As UInteger, <Out> ByRef ppConstantBuffers As ID3D11Buffer)
        <PreserveSig>
        Sub PSGetShaderResources(StartSlot As UInteger, NumViews As UInteger, <Out> ByRef ppShaderResourceViews As ID3D11ShaderResourceView)
        <PreserveSig>
        Sub PSGetShader(<Out> ByRef ppPixelShader As ID3D11PixelShader, <Out> ByRef ppClassInstances As ID3D11ClassInstance, ByRef pNumClassInstances As UInteger)
        <PreserveSig>
        Sub PSGetSamplers(StartSlot As UInteger, NumSamplers As UInteger, <Out> ByRef ppSamplers As ID3D11SamplerState)
        <PreserveSig>
        Sub VSGetShader(<Out> ByRef ppVertexShader As ID3D11VertexShader, <Out> ByRef ppClassInstances As ID3D11ClassInstance, ByRef pNumClassInstances As UInteger)
        <PreserveSig>
        Sub PSGetConstantBuffers(StartSlot As UInteger, NumBuffers As UInteger, <Out> ByRef ppConstantBuffers As ID3D11Buffer)
        <PreserveSig>
        Sub IAGetInputLayout(<Out> ByRef ppInputLayout As ID3D11InputLayout)
        <PreserveSig>
        Sub IAGetVertexBuffers(StartSlot As UInteger, NumBuffers As UInteger, <Out> ByRef ppVertexBuffers As ID3D11Buffer, <Out> ByRef pStrides As UInteger, <Out> ByRef pOffsets As UInteger)
        <PreserveSig>
        Sub IAGetIndexBuffer(<Out> ByRef pIndexBuffer As ID3D11Buffer, <Out> ByRef Format As DXGI_FORMAT, <Out> ByRef Offset As UInteger)
        <PreserveSig>
        Sub GSGetConstantBuffers(StartSlot As UInteger, NumBuffers As UInteger, <Out> ByRef ppConstantBuffers As ID3D11Buffer)
        <PreserveSig>
        Sub GSGetShader(<Out> ByRef ppGeometryShader As ID3D11GeometryShader, <Out> ByRef ppClassInstances As ID3D11ClassInstance, ByRef pNumClassInstances As UInteger)
        <PreserveSig>
        Sub IAGetPrimitiveTopology(<Out> ByRef pTopology As D3D_PRIMITIVE_TOPOLOGY)
        <PreserveSig>
        Sub VSGetShaderResources(StartSlot As UInteger, NumViews As UInteger, <Out> ByRef ppShaderResourceViews As ID3D11ShaderResourceView)
        <PreserveSig>
        Sub VSGetSamplers(StartSlot As UInteger, NumSamplers As UInteger, <Out> ByRef ppSamplers As ID3D11SamplerState)
        <PreserveSig>
        Sub GetPredication(<Out> ByRef ppPredicate As ID3D11Predicate, <Out> ByRef pPredicateValue As Boolean)
        <PreserveSig>
        Sub GSGetShaderResources(StartSlot As UInteger, NumViews As UInteger, <Out> ByRef ppShaderResourceViews As ID3D11ShaderResourceView)
        <PreserveSig>
        Sub GSGetSamplers(StartSlot As UInteger, NumSamplers As UInteger, <Out> ByRef ppSamplers As ID3D11SamplerState)
        <PreserveSig>
        Sub OMGetRenderTargets(NumViews As UInteger, <Out> ByRef ppRenderTargetViews As ID3D11RenderTargetView, <Out> ByRef ppDepthStencilView As ID3D11DepthStencilView)
        <PreserveSig>
        Sub OMGetRenderTargetsAndUnorderedAccessViews(NumRTVs As UInteger, <Out> ByRef ppRenderTargetViews As ID3D11RenderTargetView, <Out> ByRef ppDepthStencilView As ID3D11DepthStencilView, UAVStartSlot As UInteger, NumUAVs As UInteger, <Out> ByRef ppUnorderedAccessViews As ID3D11UnorderedAccessView)
        <PreserveSig>
        Sub OMGetBlendState(<Out> ByRef ppBlendState As ID3D11BlendState, <Out> ByRef BlendFactor As Single(), <Out> ByRef pSampleMask As UInteger)
        <PreserveSig>
        Sub OMGetDepthStencilState(<Out> ByRef ppDepthStencilState As ID3D11DepthStencilState, <Out> ByRef pStencilRef As UInteger)
        <PreserveSig>
        Sub SOGetTargets(NumBuffers As UInteger, <Out> ByRef ppSOTargets As ID3D11Buffer)
        <PreserveSig>
        Sub RSGetState(<Out> ByRef ppRasterizerState As ID3D11RasterizerState)
        <PreserveSig>
        Sub RSGetViewports(ByRef pNumViewports As UInteger, <Out> ByRef pViewports As D3D11_VIEWPORT)
        <PreserveSig>
        Sub RSGetScissorRects(ByRef pNumRects As UInteger, <Out> ByRef pRects As RECT)
        <PreserveSig>
        Sub HSGetShaderResources(StartSlot As UInteger, NumViews As UInteger, <Out> ByRef ppShaderResourceViews As ID3D11ShaderResourceView)
        <PreserveSig>
        Sub HSGetShader(<Out> ByRef ppHullShader As ID3D11HullShader, <Out> ByRef ppClassInstances As ID3D11ClassInstance, ByRef pNumClassInstances As UInteger)
        <PreserveSig>
        Sub HSGetSamplers(StartSlot As UInteger, NumSamplers As UInteger, <Out> ByRef ppSamplers As ID3D11SamplerState)
        <PreserveSig>
        Sub HSGetConstantBuffers(StartSlot As UInteger, NumBuffers As UInteger, <Out> ByRef ppConstantBuffers As ID3D11Buffer)
        <PreserveSig>
        Sub DSGetShaderResources(StartSlot As UInteger, NumViews As UInteger, <Out> ByRef ppShaderResourceViews As ID3D11ShaderResourceView)
        <PreserveSig>
        Sub DSGetShader(<Out> ByRef ppDomainShader As ID3D11DomainShader, <Out> ByRef ppClassInstances As ID3D11ClassInstance, ByRef pNumClassInstances As UInteger)
        <PreserveSig>
        Sub DSGetSamplers(StartSlot As UInteger, NumSamplers As UInteger, <Out> ByRef ppSamplers As ID3D11SamplerState)
        <PreserveSig>
        Sub DSGetConstantBuffers(StartSlot As UInteger, NumBuffers As UInteger, <Out> ByRef ppConstantBuffers As ID3D11Buffer)
        <PreserveSig>
        Sub CSGetShaderResources(StartSlot As UInteger, NumViews As UInteger, <Out> ByRef ppShaderResourceViews As ID3D11ShaderResourceView)
        <PreserveSig>
        Sub CSGetUnorderedAccessViews(StartSlot As UInteger, NumUAVs As UInteger, <Out> ByRef ppUnorderedAccessViews As ID3D11UnorderedAccessView)
        <PreserveSig>
        Sub CSGetShader(<Out> ByRef ppComputeShader As ID3D11ComputeShader, <Out> ByRef ppClassInstances As ID3D11ClassInstance, ByRef pNumClassInstances As UInteger)
        <PreserveSig>
        Sub CSGetSamplers(StartSlot As UInteger, NumSamplers As UInteger, <Out> ByRef ppSamplers As ID3D11SamplerState)
        <PreserveSig>
        Sub CSGetConstantBuffers(StartSlot As UInteger, NumBuffers As UInteger, <Out> ByRef ppConstantBuffers As ID3D11Buffer)
        <PreserveSig>
        Sub ClearState()
        <PreserveSig>
        Sub Flush()
        <PreserveSig>
        Sub [GetType](<Out> ByRef deviceContextType As D3D11_DEVICE_CONTEXT_TYPE)
        <PreserveSig>
        Function GetContextFlags() As UInteger
        <PreserveSig>
        Function FinishCommandList(RestoreDeferredContextState As Boolean, <Out> ByRef ppCommandList As ID3D11CommandList) As HRESULT
    End Interface

    <ComImport>
    <Guid("48570b85-d1ee-4fcd-a250-eb350722b037")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID3D11Buffer
        Inherits ID3D11Resource

#Region "ID3D11Resource"
#Region "ID3D11DeviceChild"
        <PreserveSig>
        Overloads Sub GetDevice(<Out> ByRef ppDevice As IntPtr)
        <PreserveSig>
        Overloads Function GetPrivateData(ByRef guid As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateData(ByRef guid As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateDataInterface(ByRef guid As Guid, pData As IntPtr) As HRESULT
#End Region

        <PreserveSig>
        Overloads Sub [GetType](<Out> ByRef pResourceDimension As D3D11_RESOURCE_DIMENSION)
        <PreserveSig>
        Overloads Sub SetEvictionPriority(EvictionPriority As UInteger)
        <PreserveSig>
        Overloads Function GetEvictionPriority() As UInteger
#End Region

        <PreserveSig>
        Sub GetDesc(<Out> ByRef pDesc As D3D11_BUFFER_DESC)
    End Interface

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D3D11_BUFFER_DESC
        Public ByteWidth As UInteger
        Public Usage As D3D11_USAGE
        Public BindFlags As UInteger
        Public CPUAccessFlags As UInteger
        Public MiscFlags As UInteger
        Public StructureByteStride As UInteger
    End Structure

    <ComImport>
    <Guid("839d1216-bb2e-412b-b7f4-a9dbebe08ed1")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID3D11View
        Inherits ID3D11DeviceChild
#Region "ID3D11DeviceChild"
        <PreserveSig>
        Overloads Sub GetDevice(<Out> ByRef ppDevice As IntPtr)
        <PreserveSig>
        Overloads Function GetPrivateData(ByRef guid As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateData(ByRef guid As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateDataInterface(ByRef guid As Guid, pData As IntPtr) As HRESULT
#End Region

        <PreserveSig>
        Sub GetResource(<Out> ByRef ppResource As ID3D11Resource)
    End Interface

    <ComImport>
    <Guid("b0e06fe0-8192-4e1a-b1ca-36d7414710b2")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID3D11ShaderResourceView
        Inherits ID3D11View
#Region "ID3D11View"
#Region "ID3D11DeviceChild"
        <PreserveSig>
        Overloads Sub GetDevice(<Out> ByRef ppDevice As IntPtr)
        <PreserveSig>
        Overloads Function GetPrivateData(ByRef guid As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateData(ByRef guid As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateDataInterface(ByRef guid As Guid, pData As IntPtr) As HRESULT
#End Region

        <PreserveSig>
        Overloads Sub GetResource(<Out> ByRef ppResource As ID3D11Resource)
#End Region

        <PreserveSig>
        Sub GetDesc(<Out> ByRef pDesc As D3D11_SHADER_RESOURCE_VIEW_DESC)
    End Interface

    Public Enum D3D_SRV_DIMENSION
        D3D_SRV_DIMENSION_UNKNOWN = 0
        D3D_SRV_DIMENSION_BUFFER = 1
        D3D_SRV_DIMENSION_TEXTURE1D = 2
        D3D_SRV_DIMENSION_TEXTURE1DARRAY = 3
        D3D_SRV_DIMENSION_TEXTURE2D = 4
        D3D_SRV_DIMENSION_TEXTURE2DARRAY = 5
        D3D_SRV_DIMENSION_TEXTURE2DMS = 6
        D3D_SRV_DIMENSION_TEXTURE2DMSARRAY = 7
        D3D_SRV_DIMENSION_TEXTURE3D = 8
        D3D_SRV_DIMENSION_TEXTURECUBE = 9
        D3D_SRV_DIMENSION_TEXTURECUBEARRAY = 10
        D3D_SRV_DIMENSION_BUFFEREX = 11
        D3D10_SRV_DIMENSION_UNKNOWN = D3D_SRV_DIMENSION_UNKNOWN
        D3D10_SRV_DIMENSION_BUFFER = D3D_SRV_DIMENSION_BUFFER
        D3D10_SRV_DIMENSION_TEXTURE1D = D3D_SRV_DIMENSION_TEXTURE1D
        D3D10_SRV_DIMENSION_TEXTURE1DARRAY = D3D_SRV_DIMENSION_TEXTURE1DARRAY
        D3D10_SRV_DIMENSION_TEXTURE2D = D3D_SRV_DIMENSION_TEXTURE2D
        D3D10_SRV_DIMENSION_TEXTURE2DARRAY = D3D_SRV_DIMENSION_TEXTURE2DARRAY
        D3D10_SRV_DIMENSION_TEXTURE2DMS = D3D_SRV_DIMENSION_TEXTURE2DMS
        D3D10_SRV_DIMENSION_TEXTURE2DMSARRAY = D3D_SRV_DIMENSION_TEXTURE2DMSARRAY
        D3D10_SRV_DIMENSION_TEXTURE3D = D3D_SRV_DIMENSION_TEXTURE3D
        D3D10_SRV_DIMENSION_TEXTURECUBE = D3D_SRV_DIMENSION_TEXTURECUBE
        D3D10_1_SRV_DIMENSION_UNKNOWN = D3D_SRV_DIMENSION_UNKNOWN
        D3D10_1_SRV_DIMENSION_BUFFER = D3D_SRV_DIMENSION_BUFFER
        D3D10_1_SRV_DIMENSION_TEXTURE1D = D3D_SRV_DIMENSION_TEXTURE1D
        D3D10_1_SRV_DIMENSION_TEXTURE1DARRAY = D3D_SRV_DIMENSION_TEXTURE1DARRAY
        D3D10_1_SRV_DIMENSION_TEXTURE2D = D3D_SRV_DIMENSION_TEXTURE2D
        D3D10_1_SRV_DIMENSION_TEXTURE2DARRAY = D3D_SRV_DIMENSION_TEXTURE2DARRAY
        D3D10_1_SRV_DIMENSION_TEXTURE2DMS = D3D_SRV_DIMENSION_TEXTURE2DMS
        D3D10_1_SRV_DIMENSION_TEXTURE2DMSARRAY = D3D_SRV_DIMENSION_TEXTURE2DMSARRAY
        D3D10_1_SRV_DIMENSION_TEXTURE3D = D3D_SRV_DIMENSION_TEXTURE3D
        D3D10_1_SRV_DIMENSION_TEXTURECUBE = D3D_SRV_DIMENSION_TEXTURECUBE
        D3D10_1_SRV_DIMENSION_TEXTURECUBEARRAY = D3D_SRV_DIMENSION_TEXTURECUBEARRAY
        D3D11_SRV_DIMENSION_UNKNOWN = D3D_SRV_DIMENSION_UNKNOWN
        D3D11_SRV_DIMENSION_BUFFER = D3D_SRV_DIMENSION_BUFFER
        D3D11_SRV_DIMENSION_TEXTURE1D = D3D_SRV_DIMENSION_TEXTURE1D
        D3D11_SRV_DIMENSION_TEXTURE1DARRAY = D3D_SRV_DIMENSION_TEXTURE1DARRAY
        D3D11_SRV_DIMENSION_TEXTURE2D = D3D_SRV_DIMENSION_TEXTURE2D
        D3D11_SRV_DIMENSION_TEXTURE2DARRAY = D3D_SRV_DIMENSION_TEXTURE2DARRAY
        D3D11_SRV_DIMENSION_TEXTURE2DMS = D3D_SRV_DIMENSION_TEXTURE2DMS
        D3D11_SRV_DIMENSION_TEXTURE2DMSARRAY = D3D_SRV_DIMENSION_TEXTURE2DMSARRAY
        D3D11_SRV_DIMENSION_TEXTURE3D = D3D_SRV_DIMENSION_TEXTURE3D
        D3D11_SRV_DIMENSION_TEXTURECUBE = D3D_SRV_DIMENSION_TEXTURECUBE
        D3D11_SRV_DIMENSION_TEXTURECUBEARRAY = D3D_SRV_DIMENSION_TEXTURECUBEARRAY
        D3D11_SRV_DIMENSION_BUFFEREX = D3D_SRV_DIMENSION_BUFFEREX
    End Enum

    'typedef D3D_SRV_DIMENSION D3D11_SRV_DIMENSION;
    <StructLayout(LayoutKind.Sequential)>
    Public Structure D3D11_SHADER_RESOURCE_VIEW_DESC
        Public Format As DXGI_FORMAT
        'public D3D11_SRV_DIMENSION ViewDimension;
        Public ViewDimension As D3D_SRV_DIMENSION

        '    union 
        '    {
        '    D3D11_BUFFER_SRV Buffer;
        '    D3D11_TEX1D_SRV Texture1D;
        '    D3D11_TEX1D_ARRAY_SRV Texture1DArray;
        '    D3D11_TEX2D_SRV Texture2D;
        '    D3D11_TEX2D_ARRAY_SRV Texture2DArray;
        '    D3D11_TEX2DMS_SRV Texture2DMS;
        '    D3D11_TEX2DMS_ARRAY_SRV Texture2DMSArray;
        '    D3D11_TEX3D_SRV Texture3D;
        '    D3D11_TEXCUBE_SRV TextureCube;
        '    D3D11_TEXCUBE_ARRAY_SRV TextureCubeArray;
        '    D3D11_BUFFEREX_SRV BufferEx;
        '};
    End Structure

    <ComImport> <Guid("ea82e40d-51dc-4f33-93d4-db7c9125ae8c")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID3D11PixelShader
        Inherits ID3D11DeviceChild
#Region "ID3D11DeviceChild"
        <PreserveSig>
        Overloads Sub GetDevice(<Out> ByRef ppDevice As IntPtr)
        <PreserveSig>
        Overloads Function GetPrivateData(ByRef guid As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateData(ByRef guid As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateDataInterface(ByRef guid As Guid, pData As IntPtr) As HRESULT
#End Region
    End Interface

    <ComImport> <Guid("a6cd7faa-b0b7-4a2f-9436-8662a65797cb")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID3D11ClassInstance
        Inherits ID3D11DeviceChild
#Region "ID3D11DeviceChild"
        <PreserveSig>
        Overloads Sub GetDevice(<Out> ByRef ppDevice As IntPtr)
        <PreserveSig>
        Overloads Function GetPrivateData(ByRef guid As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateData(ByRef guid As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateDataInterface(ByRef guid As Guid, pData As IntPtr) As HRESULT
#End Region

        <PreserveSig>
        Sub GetClassLinkage(<Out> ByRef ppLinkage As ID3D11ClassLinkage)
        <PreserveSig>
        Sub GetDesc(<Out> ByRef pDesc As D3D11_CLASS_INSTANCE_DESC)
        <PreserveSig>
        Sub GetInstanceName(<Out> ByRef pInstanceName As Text.StringBuilder, ByRef pBufferLength As UInteger)
        <PreserveSig>
        Sub GetTypeName(<Out> ByRef pTypeName As Text.StringBuilder, ByRef pBufferLength As UInteger)
    End Interface

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D3D11_CLASS_INSTANCE_DESC
        Public InstanceId As UInteger
        Public InstanceIndex As UInteger
        Public TypeId As UInteger
        Public ConstantBuffer As UInteger
        Public BaseConstantBufferOffset As UInteger
        Public BaseTexture As UInteger
        Public BaseSampler As UInteger
        Public Created As Boolean
    End Structure

    <ComImport> <Guid("ddf57cba-9543-46e4-a12b-f207a0fe7fed")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID3D11ClassLinkage
        Inherits ID3D11DeviceChild
#Region "ID3D11DeviceChild"
        <PreserveSig>
        Overloads Sub GetDevice(<Out> ByRef ppDevice As IntPtr)
        <PreserveSig>
        Overloads Function GetPrivateData(ByRef guid As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateData(ByRef guid As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateDataInterface(ByRef guid As Guid, pData As IntPtr) As HRESULT
#End Region

        <PreserveSig>
        Function GetClassInstance(pClassInstanceName As String, InstanceIndex As UInteger, <Out> ByRef ppInstance As ID3D11ClassInstance) As HRESULT
        <PreserveSig>
        Function CreateClassInstance(pClassTypeName As String, ConstantBufferOffset As UInteger, ConstantVectorOffset As UInteger, TextureOffset As UInteger, SamplerOffset As UInteger, <Out> ByRef ppInstance As ID3D11ClassInstance) As HRESULT
    End Interface


    <ComImport> <Guid("da6fea51-564c-4487-9810-f0d0f9b4e3a5")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID3D11SamplerState
        Inherits ID3D11DeviceChild
#Region "ID3D11DeviceChild"
        <PreserveSig>
        Overloads Sub GetDevice(<Out> ByRef ppDevice As IntPtr)
        <PreserveSig>
        Overloads Function GetPrivateData(ByRef guid As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateData(ByRef guid As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateDataInterface(ByRef guid As Guid, pData As IntPtr) As HRESULT
#End Region

        <PreserveSig>
        Sub GetDesc(<Out> ByRef pDesc As D3D11_SAMPLER_DESC)
    End Interface

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D3D11_SAMPLER_DESC
        Private Filter As D3D11_FILTER
        Private AddressU As D3D11_TEXTURE_ADDRESS_MODE
        Private AddressV As D3D11_TEXTURE_ADDRESS_MODE
        Private AddressW As D3D11_TEXTURE_ADDRESS_MODE
        Public MipLODBias As Single
        Public MaxAnisotropy As UInteger
        Private ComparisonFunc As D3D11_COMPARISON_FUNC
        <MarshalAs(UnmanagedType.R4, SizeConst:=4)>
        Public BorderColor As Single
        Public MinLOD As Single
        Public MaxLOD As Single
    End Structure
    Public Enum D3D11_FILTER
        D3D11_FILTER_MIN_MAG_MIP_POINT = 0
        D3D11_FILTER_MIN_MAG_POINT_MIP_LINEAR = &H1
        D3D11_FILTER_MIN_POINT_MAG_LINEAR_MIP_POINT = &H4
        D3D11_FILTER_MIN_POINT_MAG_MIP_LINEAR = &H5
        D3D11_FILTER_MIN_LINEAR_MAG_MIP_POINT = &H10
        D3D11_FILTER_MIN_LINEAR_MAG_POINT_MIP_LINEAR = &H11
        D3D11_FILTER_MIN_MAG_LINEAR_MIP_POINT = &H14
        D3D11_FILTER_MIN_MAG_MIP_LINEAR = &H15
        D3D11_FILTER_ANISOTROPIC = &H55
        D3D11_FILTER_COMPARISON_MIN_MAG_MIP_POINT = &H80
        D3D11_FILTER_COMPARISON_MIN_MAG_POINT_MIP_LINEAR = &H81
        D3D11_FILTER_COMPARISON_MIN_POINT_MAG_LINEAR_MIP_POINT = &H84
        D3D11_FILTER_COMPARISON_MIN_POINT_MAG_MIP_LINEAR = &H85
        D3D11_FILTER_COMPARISON_MIN_LINEAR_MAG_MIP_POINT = &H90
        D3D11_FILTER_COMPARISON_MIN_LINEAR_MAG_POINT_MIP_LINEAR = &H91
        D3D11_FILTER_COMPARISON_MIN_MAG_LINEAR_MIP_POINT = &H94
        D3D11_FILTER_COMPARISON_MIN_MAG_MIP_LINEAR = &H95
        D3D11_FILTER_COMPARISON_ANISOTROPIC = &HD5
        D3D11_FILTER_MINIMUM_MIN_MAG_MIP_POINT = &H100
        D3D11_FILTER_MINIMUM_MIN_MAG_POINT_MIP_LINEAR = &H101
        D3D11_FILTER_MINIMUM_MIN_POINT_MAG_LINEAR_MIP_POINT = &H104
        D3D11_FILTER_MINIMUM_MIN_POINT_MAG_MIP_LINEAR = &H105
        D3D11_FILTER_MINIMUM_MIN_LINEAR_MAG_MIP_POINT = &H110
        D3D11_FILTER_MINIMUM_MIN_LINEAR_MAG_POINT_MIP_LINEAR = &H111
        D3D11_FILTER_MINIMUM_MIN_MAG_LINEAR_MIP_POINT = &H114
        D3D11_FILTER_MINIMUM_MIN_MAG_MIP_LINEAR = &H115
        D3D11_FILTER_MINIMUM_ANISOTROPIC = &H155
        D3D11_FILTER_MAXIMUM_MIN_MAG_MIP_POINT = &H180
        D3D11_FILTER_MAXIMUM_MIN_MAG_POINT_MIP_LINEAR = &H181
        D3D11_FILTER_MAXIMUM_MIN_POINT_MAG_LINEAR_MIP_POINT = &H184
        D3D11_FILTER_MAXIMUM_MIN_POINT_MAG_MIP_LINEAR = &H185
        D3D11_FILTER_MAXIMUM_MIN_LINEAR_MAG_MIP_POINT = &H190
        D3D11_FILTER_MAXIMUM_MIN_LINEAR_MAG_POINT_MIP_LINEAR = &H191
        D3D11_FILTER_MAXIMUM_MIN_MAG_LINEAR_MIP_POINT = &H194
        D3D11_FILTER_MAXIMUM_MIN_MAG_MIP_LINEAR = &H195
        D3D11_FILTER_MAXIMUM_ANISOTROPIC = &H1D5
    End Enum

    Public Enum D3D11_TEXTURE_ADDRESS_MODE
        D3D11_TEXTURE_ADDRESS_WRAP = 1
        D3D11_TEXTURE_ADDRESS_MIRROR = 2
        D3D11_TEXTURE_ADDRESS_CLAMP = 3
        D3D11_TEXTURE_ADDRESS_BORDER = 4
        D3D11_TEXTURE_ADDRESS_MIRROR_ONCE = 5
    End Enum

    Public Enum D3D11_COMPARISON_FUNC
        D3D11_COMPARISON_NEVER = 1
        D3D11_COMPARISON_LESS = 2
        D3D11_COMPARISON_EQUAL = 3
        D3D11_COMPARISON_LESS_EQUAL = 4
        D3D11_COMPARISON_GREATER = 5
        D3D11_COMPARISON_NOT_EQUAL = 6
        D3D11_COMPARISON_GREATER_EQUAL = 7
        D3D11_COMPARISON_ALWAYS = 8
    End Enum

    <ComImport> <Guid("3b301d64-d678-4289-8897-22f8928b72f3")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID3D11VertexShader
        Inherits ID3D11DeviceChild
#Region "ID3D11DeviceChild"
        <PreserveSig>
        Overloads Sub GetDevice(<Out> ByRef ppDevice As IntPtr)
        <PreserveSig>
        Overloads Function GetPrivateData(ByRef guid As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateData(ByRef guid As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateDataInterface(ByRef guid As Guid, pData As IntPtr) As HRESULT
#End Region
    End Interface

    Public Enum D3D11_MAP
        D3D11_MAP_READ = 1
        D3D11_MAP_WRITE = 2
        D3D11_MAP_READ_WRITE = 3
        D3D11_MAP_WRITE_DISCARD = 4
        D3D11_MAP_WRITE_NO_OVERWRITE = 5
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D3D11_MAPPED_SUBRESOURCE
        Public pData As IntPtr
        Public RowPitch As UInteger
        Public DepthPitch As UInteger
    End Structure

    <ComImport> <Guid("e4819ddc-4cf0-4025-bd26-5de82a3e07b7")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID3D11InputLayout
        Inherits ID3D11DeviceChild
#Region "ID3D11DeviceChild"
        <PreserveSig>
        Overloads Sub GetDevice(<Out> ByRef ppDevice As IntPtr)
        <PreserveSig>
        Overloads Function GetPrivateData(ByRef guid As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateData(ByRef guid As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateDataInterface(ByRef guid As Guid, pData As IntPtr) As HRESULT
#End Region
    End Interface

    <ComImport> <Guid("38325b96-effb-4022-ba02-2e795b70275c")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID3D11GeometryShader
        Inherits ID3D11DeviceChild
#Region "ID3D11DeviceChild"
        <PreserveSig>
        Overloads Sub GetDevice(<Out> ByRef ppDevice As IntPtr)
        <PreserveSig>
        Overloads Function GetPrivateData(ByRef guid As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateData(ByRef guid As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateDataInterface(ByRef guid As Guid, pData As IntPtr) As HRESULT
#End Region
    End Interface

    Public Enum D3D_PRIMITIVE_TOPOLOGY
        D3D_PRIMITIVE_TOPOLOGY_UNDEFINED = 0
        D3D_PRIMITIVE_TOPOLOGY_POINTLIST = 1
        D3D_PRIMITIVE_TOPOLOGY_LINELIST = 2
        D3D_PRIMITIVE_TOPOLOGY_LINESTRIP = 3
        D3D_PRIMITIVE_TOPOLOGY_TRIANGLELIST = 4
        D3D_PRIMITIVE_TOPOLOGY_TRIANGLESTRIP = 5
        D3D_PRIMITIVE_TOPOLOGY_LINELIST_ADJ = 10
        D3D_PRIMITIVE_TOPOLOGY_LINESTRIP_ADJ = 11
        D3D_PRIMITIVE_TOPOLOGY_TRIANGLELIST_ADJ = 12
        D3D_PRIMITIVE_TOPOLOGY_TRIANGLESTRIP_ADJ = 13
        D3D_PRIMITIVE_TOPOLOGY_1_CONTROL_POINT_PATCHLIST = 33
        D3D_PRIMITIVE_TOPOLOGY_2_CONTROL_POINT_PATCHLIST = 34
        D3D_PRIMITIVE_TOPOLOGY_3_CONTROL_POINT_PATCHLIST = 35
        D3D_PRIMITIVE_TOPOLOGY_4_CONTROL_POINT_PATCHLIST = 36
        D3D_PRIMITIVE_TOPOLOGY_5_CONTROL_POINT_PATCHLIST = 37
        D3D_PRIMITIVE_TOPOLOGY_6_CONTROL_POINT_PATCHLIST = 38
        D3D_PRIMITIVE_TOPOLOGY_7_CONTROL_POINT_PATCHLIST = 39
        D3D_PRIMITIVE_TOPOLOGY_8_CONTROL_POINT_PATCHLIST = 40
        D3D_PRIMITIVE_TOPOLOGY_9_CONTROL_POINT_PATCHLIST = 41
        D3D_PRIMITIVE_TOPOLOGY_10_CONTROL_POINT_PATCHLIST = 42
        D3D_PRIMITIVE_TOPOLOGY_11_CONTROL_POINT_PATCHLIST = 43
        D3D_PRIMITIVE_TOPOLOGY_12_CONTROL_POINT_PATCHLIST = 44
        D3D_PRIMITIVE_TOPOLOGY_13_CONTROL_POINT_PATCHLIST = 45
        D3D_PRIMITIVE_TOPOLOGY_14_CONTROL_POINT_PATCHLIST = 46
        D3D_PRIMITIVE_TOPOLOGY_15_CONTROL_POINT_PATCHLIST = 47
        D3D_PRIMITIVE_TOPOLOGY_16_CONTROL_POINT_PATCHLIST = 48
        D3D_PRIMITIVE_TOPOLOGY_17_CONTROL_POINT_PATCHLIST = 49
        D3D_PRIMITIVE_TOPOLOGY_18_CONTROL_POINT_PATCHLIST = 50
        D3D_PRIMITIVE_TOPOLOGY_19_CONTROL_POINT_PATCHLIST = 51
        D3D_PRIMITIVE_TOPOLOGY_20_CONTROL_POINT_PATCHLIST = 52
        D3D_PRIMITIVE_TOPOLOGY_21_CONTROL_POINT_PATCHLIST = 53
        D3D_PRIMITIVE_TOPOLOGY_22_CONTROL_POINT_PATCHLIST = 54
        D3D_PRIMITIVE_TOPOLOGY_23_CONTROL_POINT_PATCHLIST = 55
        D3D_PRIMITIVE_TOPOLOGY_24_CONTROL_POINT_PATCHLIST = 56
        D3D_PRIMITIVE_TOPOLOGY_25_CONTROL_POINT_PATCHLIST = 57
        D3D_PRIMITIVE_TOPOLOGY_26_CONTROL_POINT_PATCHLIST = 58
        D3D_PRIMITIVE_TOPOLOGY_27_CONTROL_POINT_PATCHLIST = 59
        D3D_PRIMITIVE_TOPOLOGY_28_CONTROL_POINT_PATCHLIST = 60
        D3D_PRIMITIVE_TOPOLOGY_29_CONTROL_POINT_PATCHLIST = 61
        D3D_PRIMITIVE_TOPOLOGY_30_CONTROL_POINT_PATCHLIST = 62
        D3D_PRIMITIVE_TOPOLOGY_31_CONTROL_POINT_PATCHLIST = 63
        D3D_PRIMITIVE_TOPOLOGY_32_CONTROL_POINT_PATCHLIST = 64
        D3D10_PRIMITIVE_TOPOLOGY_UNDEFINED = D3D_PRIMITIVE_TOPOLOGY_UNDEFINED
        D3D10_PRIMITIVE_TOPOLOGY_POINTLIST = D3D_PRIMITIVE_TOPOLOGY_POINTLIST
        D3D10_PRIMITIVE_TOPOLOGY_LINELIST = D3D_PRIMITIVE_TOPOLOGY_LINELIST
        D3D10_PRIMITIVE_TOPOLOGY_LINESTRIP = D3D_PRIMITIVE_TOPOLOGY_LINESTRIP
        D3D10_PRIMITIVE_TOPOLOGY_TRIANGLELIST = D3D_PRIMITIVE_TOPOLOGY_TRIANGLELIST
        D3D10_PRIMITIVE_TOPOLOGY_TRIANGLESTRIP = D3D_PRIMITIVE_TOPOLOGY_TRIANGLESTRIP
        D3D10_PRIMITIVE_TOPOLOGY_LINELIST_ADJ = D3D_PRIMITIVE_TOPOLOGY_LINELIST_ADJ
        D3D10_PRIMITIVE_TOPOLOGY_LINESTRIP_ADJ = D3D_PRIMITIVE_TOPOLOGY_LINESTRIP_ADJ
        D3D10_PRIMITIVE_TOPOLOGY_TRIANGLELIST_ADJ = D3D_PRIMITIVE_TOPOLOGY_TRIANGLELIST_ADJ
        D3D10_PRIMITIVE_TOPOLOGY_TRIANGLESTRIP_ADJ = D3D_PRIMITIVE_TOPOLOGY_TRIANGLESTRIP_ADJ
        D3D11_PRIMITIVE_TOPOLOGY_UNDEFINED = D3D_PRIMITIVE_TOPOLOGY_UNDEFINED
        D3D11_PRIMITIVE_TOPOLOGY_POINTLIST = D3D_PRIMITIVE_TOPOLOGY_POINTLIST
        D3D11_PRIMITIVE_TOPOLOGY_LINELIST = D3D_PRIMITIVE_TOPOLOGY_LINELIST
        D3D11_PRIMITIVE_TOPOLOGY_LINESTRIP = D3D_PRIMITIVE_TOPOLOGY_LINESTRIP
        D3D11_PRIMITIVE_TOPOLOGY_TRIANGLELIST = D3D_PRIMITIVE_TOPOLOGY_TRIANGLELIST
        D3D11_PRIMITIVE_TOPOLOGY_TRIANGLESTRIP = D3D_PRIMITIVE_TOPOLOGY_TRIANGLESTRIP
        D3D11_PRIMITIVE_TOPOLOGY_LINELIST_ADJ = D3D_PRIMITIVE_TOPOLOGY_LINELIST_ADJ
        D3D11_PRIMITIVE_TOPOLOGY_LINESTRIP_ADJ = D3D_PRIMITIVE_TOPOLOGY_LINESTRIP_ADJ
        D3D11_PRIMITIVE_TOPOLOGY_TRIANGLELIST_ADJ = D3D_PRIMITIVE_TOPOLOGY_TRIANGLELIST_ADJ
        D3D11_PRIMITIVE_TOPOLOGY_TRIANGLESTRIP_ADJ = D3D_PRIMITIVE_TOPOLOGY_TRIANGLESTRIP_ADJ
        D3D11_PRIMITIVE_TOPOLOGY_1_CONTROL_POINT_PATCHLIST = D3D_PRIMITIVE_TOPOLOGY_1_CONTROL_POINT_PATCHLIST
        D3D11_PRIMITIVE_TOPOLOGY_2_CONTROL_POINT_PATCHLIST = D3D_PRIMITIVE_TOPOLOGY_2_CONTROL_POINT_PATCHLIST
        D3D11_PRIMITIVE_TOPOLOGY_3_CONTROL_POINT_PATCHLIST = D3D_PRIMITIVE_TOPOLOGY_3_CONTROL_POINT_PATCHLIST
        D3D11_PRIMITIVE_TOPOLOGY_4_CONTROL_POINT_PATCHLIST = D3D_PRIMITIVE_TOPOLOGY_4_CONTROL_POINT_PATCHLIST
        D3D11_PRIMITIVE_TOPOLOGY_5_CONTROL_POINT_PATCHLIST = D3D_PRIMITIVE_TOPOLOGY_5_CONTROL_POINT_PATCHLIST
        D3D11_PRIMITIVE_TOPOLOGY_6_CONTROL_POINT_PATCHLIST = D3D_PRIMITIVE_TOPOLOGY_6_CONTROL_POINT_PATCHLIST
        D3D11_PRIMITIVE_TOPOLOGY_7_CONTROL_POINT_PATCHLIST = D3D_PRIMITIVE_TOPOLOGY_7_CONTROL_POINT_PATCHLIST
        D3D11_PRIMITIVE_TOPOLOGY_8_CONTROL_POINT_PATCHLIST = D3D_PRIMITIVE_TOPOLOGY_8_CONTROL_POINT_PATCHLIST
        D3D11_PRIMITIVE_TOPOLOGY_9_CONTROL_POINT_PATCHLIST = D3D_PRIMITIVE_TOPOLOGY_9_CONTROL_POINT_PATCHLIST
        D3D11_PRIMITIVE_TOPOLOGY_10_CONTROL_POINT_PATCHLIST = D3D_PRIMITIVE_TOPOLOGY_10_CONTROL_POINT_PATCHLIST
        D3D11_PRIMITIVE_TOPOLOGY_11_CONTROL_POINT_PATCHLIST = D3D_PRIMITIVE_TOPOLOGY_11_CONTROL_POINT_PATCHLIST
        D3D11_PRIMITIVE_TOPOLOGY_12_CONTROL_POINT_PATCHLIST = D3D_PRIMITIVE_TOPOLOGY_12_CONTROL_POINT_PATCHLIST
        D3D11_PRIMITIVE_TOPOLOGY_13_CONTROL_POINT_PATCHLIST = D3D_PRIMITIVE_TOPOLOGY_13_CONTROL_POINT_PATCHLIST
        D3D11_PRIMITIVE_TOPOLOGY_14_CONTROL_POINT_PATCHLIST = D3D_PRIMITIVE_TOPOLOGY_14_CONTROL_POINT_PATCHLIST
        D3D11_PRIMITIVE_TOPOLOGY_15_CONTROL_POINT_PATCHLIST = D3D_PRIMITIVE_TOPOLOGY_15_CONTROL_POINT_PATCHLIST
        D3D11_PRIMITIVE_TOPOLOGY_16_CONTROL_POINT_PATCHLIST = D3D_PRIMITIVE_TOPOLOGY_16_CONTROL_POINT_PATCHLIST
        D3D11_PRIMITIVE_TOPOLOGY_17_CONTROL_POINT_PATCHLIST = D3D_PRIMITIVE_TOPOLOGY_17_CONTROL_POINT_PATCHLIST
        D3D11_PRIMITIVE_TOPOLOGY_18_CONTROL_POINT_PATCHLIST = D3D_PRIMITIVE_TOPOLOGY_18_CONTROL_POINT_PATCHLIST
        D3D11_PRIMITIVE_TOPOLOGY_19_CONTROL_POINT_PATCHLIST = D3D_PRIMITIVE_TOPOLOGY_19_CONTROL_POINT_PATCHLIST
        D3D11_PRIMITIVE_TOPOLOGY_20_CONTROL_POINT_PATCHLIST = D3D_PRIMITIVE_TOPOLOGY_20_CONTROL_POINT_PATCHLIST
        D3D11_PRIMITIVE_TOPOLOGY_21_CONTROL_POINT_PATCHLIST = D3D_PRIMITIVE_TOPOLOGY_21_CONTROL_POINT_PATCHLIST
        D3D11_PRIMITIVE_TOPOLOGY_22_CONTROL_POINT_PATCHLIST = D3D_PRIMITIVE_TOPOLOGY_22_CONTROL_POINT_PATCHLIST
        D3D11_PRIMITIVE_TOPOLOGY_23_CONTROL_POINT_PATCHLIST = D3D_PRIMITIVE_TOPOLOGY_23_CONTROL_POINT_PATCHLIST
        D3D11_PRIMITIVE_TOPOLOGY_24_CONTROL_POINT_PATCHLIST = D3D_PRIMITIVE_TOPOLOGY_24_CONTROL_POINT_PATCHLIST
        D3D11_PRIMITIVE_TOPOLOGY_25_CONTROL_POINT_PATCHLIST = D3D_PRIMITIVE_TOPOLOGY_25_CONTROL_POINT_PATCHLIST
        D3D11_PRIMITIVE_TOPOLOGY_26_CONTROL_POINT_PATCHLIST = D3D_PRIMITIVE_TOPOLOGY_26_CONTROL_POINT_PATCHLIST
        D3D11_PRIMITIVE_TOPOLOGY_27_CONTROL_POINT_PATCHLIST = D3D_PRIMITIVE_TOPOLOGY_27_CONTROL_POINT_PATCHLIST
        D3D11_PRIMITIVE_TOPOLOGY_28_CONTROL_POINT_PATCHLIST = D3D_PRIMITIVE_TOPOLOGY_28_CONTROL_POINT_PATCHLIST
        D3D11_PRIMITIVE_TOPOLOGY_29_CONTROL_POINT_PATCHLIST = D3D_PRIMITIVE_TOPOLOGY_29_CONTROL_POINT_PATCHLIST
        D3D11_PRIMITIVE_TOPOLOGY_30_CONTROL_POINT_PATCHLIST = D3D_PRIMITIVE_TOPOLOGY_30_CONTROL_POINT_PATCHLIST
        D3D11_PRIMITIVE_TOPOLOGY_31_CONTROL_POINT_PATCHLIST = D3D_PRIMITIVE_TOPOLOGY_31_CONTROL_POINT_PATCHLIST
        D3D11_PRIMITIVE_TOPOLOGY_32_CONTROL_POINT_PATCHLIST = D3D_PRIMITIVE_TOPOLOGY_32_CONTROL_POINT_PATCHLIST
    End Enum

    <ComImport> <Guid("4b35d0cd-1e15-4258-9c98-1b1333f6dd3b")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID3D11Asynchronous
        Inherits ID3D11DeviceChild
#Region "ID3D11DeviceChild"
        <PreserveSig>
        Overloads Sub GetDevice(<Out> ByRef ppDevice As IntPtr)
        <PreserveSig>
        Overloads Function GetPrivateData(ByRef guid As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateData(ByRef guid As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateDataInterface(ByRef guid As Guid, pData As IntPtr) As HRESULT
#End Region

        <PreserveSig>
        Function GetDataSize() As UInteger
    End Interface

    <ComImport> <Guid("d6c00747-87b7-425e-b84d-44d108560afd")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID3D11Query
        Inherits ID3D11Asynchronous
#Region "ID3D11Asynchronous"
#Region "ID3D11DeviceChild"
        <PreserveSig>
        Overloads Sub GetDevice(<Out> ByRef ppDevice As IntPtr)
        <PreserveSig>
        Overloads Function GetPrivateData(ByRef guid As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateData(ByRef guid As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateDataInterface(ByRef guid As Guid, pData As IntPtr) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function GetDataSize() As UInteger
#End Region

        <PreserveSig>
        Sub GetDesc(<Out> ByRef pDesc As D3D11_QUERY_DESC)
    End Interface


    <StructLayout(LayoutKind.Sequential)>
    Public Structure D3D11_QUERY_DESC
        Public Query As D3D11_QUERY
        Public MiscFlags As UInteger
    End Structure

    Public Enum D3D11_QUERY
        D3D11_QUERY_EVENT = 0
        D3D11_QUERY_OCCLUSION = D3D11_QUERY_EVENT + 1
        D3D11_QUERY_TIMESTAMP = D3D11_QUERY_OCCLUSION + 1
        D3D11_QUERY_TIMESTAMP_DISJOINT = D3D11_QUERY_TIMESTAMP + 1
        D3D11_QUERY_PIPELINE_STATISTICS = D3D11_QUERY_TIMESTAMP_DISJOINT + 1
        D3D11_QUERY_OCCLUSION_PREDICATE = D3D11_QUERY_PIPELINE_STATISTICS + 1
        D3D11_QUERY_SO_STATISTICS = D3D11_QUERY_OCCLUSION_PREDICATE + 1
        D3D11_QUERY_SO_OVERFLOW_PREDICATE = D3D11_QUERY_SO_STATISTICS + 1
        D3D11_QUERY_SO_STATISTICS_STREAM0 = D3D11_QUERY_SO_OVERFLOW_PREDICATE + 1
        D3D11_QUERY_SO_OVERFLOW_PREDICATE_STREAM0 = D3D11_QUERY_SO_STATISTICS_STREAM0 + 1
        D3D11_QUERY_SO_STATISTICS_STREAM1 = D3D11_QUERY_SO_OVERFLOW_PREDICATE_STREAM0 + 1
        D3D11_QUERY_SO_OVERFLOW_PREDICATE_STREAM1 = D3D11_QUERY_SO_STATISTICS_STREAM1 + 1
        D3D11_QUERY_SO_STATISTICS_STREAM2 = D3D11_QUERY_SO_OVERFLOW_PREDICATE_STREAM1 + 1
        D3D11_QUERY_SO_OVERFLOW_PREDICATE_STREAM2 = D3D11_QUERY_SO_STATISTICS_STREAM2 + 1
        D3D11_QUERY_SO_STATISTICS_STREAM3 = D3D11_QUERY_SO_OVERFLOW_PREDICATE_STREAM2 + 1
        D3D11_QUERY_SO_OVERFLOW_PREDICATE_STREAM3 = D3D11_QUERY_SO_STATISTICS_STREAM3 + 1
    End Enum

    <ComImport> <Guid("9eb576dd-9f77-4d86-81aa-8bab5fe490e2")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID3D11Predicate
        Inherits ID3D11Query
#Region "ID3D11Query"
#Region "ID3D11Asynchronous"
#Region "ID3D11DeviceChild"
        'void GetDevice(out ID3D11Device ppDevice);
        Overloads Sub GetDevice(<Out> ByRef ppDevice As IntPtr)
        Overloads Function GetPrivateData(ByRef guid As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        Overloads Function SetPrivateData(ByRef guid As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        Overloads Function SetPrivateDataInterface(ByRef guid As Guid, pData As IntPtr) As HRESULT
#End Region
        <PreserveSig>
        Overloads Function GetDataSize() As UInteger
#End Region

        <PreserveSig>
        Overloads Sub GetDesc(<Out> ByRef pDesc As D3D11_QUERY_DESC)
#End Region
    End Interface

    <ComImport> <Guid("dfdba067-0b8d-4865-875b-d7b4516cc164")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID3D11RenderTargetView
        Inherits ID3D11View
#Region "ID3D11View"
#Region "ID3D11DeviceChild"
        <PreserveSig>
        Overloads Sub GetDevice(<Out> ByRef ppDevice As IntPtr)
        <PreserveSig>
        Overloads Function GetPrivateData(ByRef guid As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateData(ByRef guid As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateDataInterface(ByRef guid As Guid, pData As IntPtr) As HRESULT
#End Region
        <PreserveSig>
        Overloads Sub GetResource(<Out> ByRef ppResource As ID3D11Resource)
#End Region

        <PreserveSig>
        Sub GetDesc(<Out> ByRef pDesc As D3D11_RENDER_TARGET_VIEW_DESC)
    End Interface

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D3D11_RENDER_TARGET_VIEW_DESC
        Public Format As DXGI_FORMAT
        Public ViewDimension As D3D11_RTV_DIMENSION
        '    union 
        '    {
        '    D3D11_BUFFER_RTV Buffer;
        '    D3D11_TEX1D_RTV Texture1D;
        '    D3D11_TEX1D_ARRAY_RTV Texture1DArray;
        '    D3D11_TEX2D_RTV Texture2D;
        '    D3D11_TEX2D_ARRAY_RTV Texture2DArray;
        '    D3D11_TEX2DMS_RTV Texture2DMS;
        '    D3D11_TEX2DMS_ARRAY_RTV Texture2DMSArray;
        '    D3D11_TEX3D_RTV Texture3D;
        '};

    End Structure

    Public Enum D3D11_RTV_DIMENSION
        D3D11_RTV_DIMENSION_UNKNOWN = 0
        D3D11_RTV_DIMENSION_BUFFER = 1
        D3D11_RTV_DIMENSION_TEXTURE1D = 2
        D3D11_RTV_DIMENSION_TEXTURE1DARRAY = 3
        D3D11_RTV_DIMENSION_TEXTURE2D = 4
        D3D11_RTV_DIMENSION_TEXTURE2DARRAY = 5
        D3D11_RTV_DIMENSION_TEXTURE2DMS = 6
        D3D11_RTV_DIMENSION_TEXTURE2DMSARRAY = 7
        D3D11_RTV_DIMENSION_TEXTURE3D = 8
    End Enum


    <ComImport> <Guid("9fdac92a-1876-48c3-afad-25b94f84a9b6")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID3D11DepthStencilView
        Inherits ID3D11View
#Region "ID3D11View"
#Region "ID3D11DeviceChild"
        <PreserveSig>
        Overloads Sub GetDevice(<Out> ByRef ppDevice As IntPtr)
        <PreserveSig>
        Overloads Function GetPrivateData(ByRef guid As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateData(ByRef guid As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateDataInterface(ByRef guid As Guid, pData As IntPtr) As HRESULT
#End Region
        <PreserveSig>
        Overloads Sub GetResource(<Out> ByRef ppResource As ID3D11Resource)
#End Region

        <PreserveSig>
        Sub GetDesc(<Out> ByRef pDesc As D3D11_DEPTH_STENCIL_VIEW_DESC)
    End Interface

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D3D11_DEPTH_STENCIL_VIEW_DESC
        Public Format As DXGI_FORMAT
        Public ViewDimension As D3D11_DSV_DIMENSION
        Public Flags As UInteger
        '    union 
        '    {
        '    D3D11_TEX1D_DSV Texture1D;
        '    D3D11_TEX1D_ARRAY_DSV Texture1DArray;
        '    D3D11_TEX2D_DSV Texture2D;
        '    D3D11_TEX2D_ARRAY_DSV Texture2DArray;
        '    D3D11_TEX2DMS_DSV Texture2DMS;
        '    D3D11_TEX2DMS_ARRAY_DSV Texture2DMSArray;
        '};

    End Structure

    Public Enum D3D11_DSV_DIMENSION
        D3D11_DSV_DIMENSION_UNKNOWN = 0
        D3D11_DSV_DIMENSION_TEXTURE1D = 1
        D3D11_DSV_DIMENSION_TEXTURE1DARRAY = 2
        D3D11_DSV_DIMENSION_TEXTURE2D = 3
        D3D11_DSV_DIMENSION_TEXTURE2DARRAY = 4
        D3D11_DSV_DIMENSION_TEXTURE2DMS = 5
        D3D11_DSV_DIMENSION_TEXTURE2DMSARRAY = 6
    End Enum

    <ComImport> <Guid("28acf509-7f5c-48f6-8611-f316010a6380")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID3D11UnorderedAccessView
        Inherits ID3D11View
#Region "ID3D11View"
#Region "ID3D11DeviceChild"
        <PreserveSig>
        Overloads Sub GetDevice(<Out> ByRef ppDevice As IntPtr)
        <PreserveSig>
        Overloads Function GetPrivateData(ByRef guid As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateData(ByRef guid As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateDataInterface(ByRef guid As Guid, pData As IntPtr) As HRESULT
#End Region

        <PreserveSig>
        Overloads Sub GetResource(<Out> ByRef ppResource As ID3D11Resource)
#End Region

        <PreserveSig>
        Sub GetDesc(<Out> ByRef pDesc As D3D11_UNORDERED_ACCESS_VIEW_DESC)
    End Interface

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D3D11_UNORDERED_ACCESS_VIEW_DESC
        Public Format As DXGI_FORMAT
        Public ViewDimension As D3D11_UAV_DIMENSION
        '    union 
        '    {
        '    D3D11_BUFFER_UAV Buffer;
        '    D3D11_TEX1D_UAV Texture1D;
        '    D3D11_TEX1D_ARRAY_UAV Texture1DArray;
        '    D3D11_TEX2D_UAV Texture2D;
        '    D3D11_TEX2D_ARRAY_UAV Texture2DArray;
        '    D3D11_TEX3D_UAV Texture3D;
        '};

    End Structure

    Public Enum D3D11_UAV_DIMENSION
        D3D11_UAV_DIMENSION_UNKNOWN = 0
        D3D11_UAV_DIMENSION_BUFFER = 1
        D3D11_UAV_DIMENSION_TEXTURE1D = 2
        D3D11_UAV_DIMENSION_TEXTURE1DARRAY = 3
        D3D11_UAV_DIMENSION_TEXTURE2D = 4
        D3D11_UAV_DIMENSION_TEXTURE2DARRAY = 5
        D3D11_UAV_DIMENSION_TEXTURE3D = 8
    End Enum

    <ComImport> <Guid("75b68faa-347d-4159-8f45-a0640f01cd9a")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID3D11BlendState
        Inherits ID3D11DeviceChild
#Region "ID3D11DeviceChild"
        <PreserveSig>
        Overloads Sub GetDevice(<Out> ByRef ppDevice As IntPtr)
        <PreserveSig>
        Overloads Function GetPrivateData(ByRef guid As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateData(ByRef guid As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateDataInterface(ByRef guid As Guid, pData As IntPtr) As HRESULT
#End Region

        <PreserveSig>
        Sub GetDesc(<Out> ByRef pDesc As D3D11_BLEND_DESC)
    End Interface

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D3D11_BLEND_DESC
        Public AlphaToCoverageEnable As Boolean
        Public IndependentBlendEnable As Boolean
        <MarshalAs(UnmanagedType.LPStruct, SizeConst:=8)>
        Public RenderTarget As D3D11_RENDER_TARGET_BLEND_DESC
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D3D11_RENDER_TARGET_BLEND_DESC
        Public BlendEnable As Boolean
        Public SrcBlend As D3D11_BLEND
        Public DestBlend As D3D11_BLEND
        Public BlendOp As D3D11_BLEND_OP
        Public SrcBlendAlpha As D3D11_BLEND
        Public DestBlendAlpha As D3D11_BLEND
        Public BlendOpAlpha As D3D11_BLEND_OP
        Public RenderTargetWriteMask As Byte
    End Structure

    Public Enum D3D11_BLEND
        D3D11_BLEND_ZERO = 1
        D3D11_BLEND_ONE = 2
        D3D11_BLEND_SRC_COLOR = 3
        D3D11_BLEND_INV_SRC_COLOR = 4
        D3D11_BLEND_SRC_ALPHA = 5
        D3D11_BLEND_INV_SRC_ALPHA = 6
        D3D11_BLEND_DEST_ALPHA = 7
        D3D11_BLEND_INV_DEST_ALPHA = 8
        D3D11_BLEND_DEST_COLOR = 9
        D3D11_BLEND_INV_DEST_COLOR = 10
        D3D11_BLEND_SRC_ALPHA_SAT = 11
        D3D11_BLEND_BLEND_FACTOR = 14
        D3D11_BLEND_INV_BLEND_FACTOR = 15
        D3D11_BLEND_SRC1_COLOR = 16
        D3D11_BLEND_INV_SRC1_COLOR = 17
        D3D11_BLEND_SRC1_ALPHA = 18
        D3D11_BLEND_INV_SRC1_ALPHA = 19
    End Enum

    Public Enum D3D11_BLEND_OP
        D3D11_BLEND_OP_ADD = 1
        D3D11_BLEND_OP_SUBTRACT = 2
        D3D11_BLEND_OP_REV_SUBTRACT = 3
        D3D11_BLEND_OP_MIN = 4
        D3D11_BLEND_OP_MAX = 5
    End Enum

    <ComImport> <Guid("03823efb-8d8f-4e1c-9aa2-f64bb2cbfdf1")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID3D11DepthStencilState
        Inherits ID3D11DeviceChild
#Region "ID3D11DeviceChild"
        <PreserveSig>
        Overloads Sub GetDevice(<Out> ByRef ppDevice As IntPtr)
        <PreserveSig>
        Overloads Function GetPrivateData(ByRef guid As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateData(ByRef guid As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateDataInterface(ByRef guid As Guid, pData As IntPtr) As HRESULT
#End Region

        <PreserveSig>
        Sub GetDesc(<Out> ByRef pDesc As D3D11_DEPTH_STENCIL_DESC)
    End Interface

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D3D11_DEPTH_STENCIL_DESC
        Public DepthEnable As Boolean
        Public DepthWriteMask As D3D11_DEPTH_WRITE_MASK
        Public DepthFunc As D3D11_COMPARISON_FUNC
        Public StencilEnable As Boolean
        Public StencilReadMask As Byte
        Public StencilWriteMask As Byte
        Public FrontFace As D3D11_DEPTH_STENCILOP_DESC
        Public BackFace As D3D11_DEPTH_STENCILOP_DESC
    End Structure

    Public Enum D3D11_DEPTH_WRITE_MASK
        D3D11_DEPTH_WRITE_MASK_ZERO = 0
        D3D11_DEPTH_WRITE_MASK_ALL = 1
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D3D11_DEPTH_STENCILOP_DESC
        Public StencilFailOp As D3D11_STENCIL_OP
        Public StencilDepthFailOp As D3D11_STENCIL_OP
        Public StencilPassOp As D3D11_STENCIL_OP
        Public StencilFunc As D3D11_COMPARISON_FUNC
    End Structure
    Public Enum D3D11_STENCIL_OP
        D3D11_STENCIL_OP_KEEP = 1
        D3D11_STENCIL_OP_ZERO = 2
        D3D11_STENCIL_OP_REPLACE = 3
        D3D11_STENCIL_OP_INCR_SAT = 4
        D3D11_STENCIL_OP_DECR_SAT = 5
        D3D11_STENCIL_OP_INVERT = 6
        D3D11_STENCIL_OP_INCR = 7
        D3D11_STENCIL_OP_DECR = 8
    End Enum

    <ComImport> <Guid("9bb4ab81-ab1a-4d8f-b506-fc04200b6ee7")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID3D11RasterizerState
        Inherits ID3D11DeviceChild
#Region "ID3D11DeviceChild"
        <PreserveSig>
        Overloads Sub GetDevice(<Out> ByRef ppDevice As IntPtr)
        <PreserveSig>
        Overloads Function GetPrivateData(ByRef guid As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateData(ByRef guid As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateDataInterface(ByRef guid As Guid, pData As IntPtr) As HRESULT
#End Region

        <PreserveSig>
        Sub GetDesc(<Out> ByRef pDesc As D3D11_RASTERIZER_DESC)
    End Interface

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D3D11_RASTERIZER_DESC
        Public FillMode As D3D11_FILL_MODE
        Public CullMode As D3D11_CULL_MODE
        Public FrontCounterClockwise As Boolean
        Public DepthBias As Integer
        Public DepthBiasClamp As Single
        Public SlopeScaledDepthBias As Single
        Public DepthClipEnable As Boolean
        Public ScissorEnable As Boolean
        Public MultisampleEnable As Boolean
        Public AntialiasedLineEnable As Boolean
    End Structure

    Public Enum D3D11_FILL_MODE
        D3D11_FILL_WIREFRAME = 2
        D3D11_FILL_SOLID = 3
    End Enum

    Public Enum D3D11_CULL_MODE
        D3D11_CULL_NONE = 1
        D3D11_CULL_FRONT = 2
        D3D11_CULL_BACK = 3
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D3D11_VIEWPORT
        Public TopLeftX As Single
        Public TopLeftY As Single
        Public Width As Single
        Public Height As Single
        Public MinDepth As Single
        Public MaxDepth As Single
    End Structure

    <ComImport> <Guid("a24bc4d1-769e-43f7-8013-98ff566c18e2")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID3D11CommandList
        Inherits ID3D11DeviceChild
#Region "ID3D11DeviceChild"
        <PreserveSig>
        Overloads Sub GetDevice(<Out> ByRef ppDevice As IntPtr)
        <PreserveSig>
        Overloads Function GetPrivateData(ByRef guid As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateData(ByRef guid As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateDataInterface(ByRef guid As Guid, pData As IntPtr) As HRESULT
#End Region

        <PreserveSig>
        Function GetContextFlags() As UInteger
    End Interface

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D3D11_BOX
        Public left As UInteger
        Public top As UInteger
        Public front As UInteger
        Public right As UInteger
        Public bottom As UInteger
        Public back As UInteger
    End Structure

    <ComImport> <Guid("8e5c6061-628a-4c8e-8264-bbe45cb3d5dd")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID3D11HullShader
        Inherits ID3D11DeviceChild
#Region "ID3D11DeviceChild"
        <PreserveSig>
        Overloads Sub GetDevice(<Out> ByRef ppDevice As IntPtr)
        <PreserveSig>
        Overloads Function GetPrivateData(ByRef guid As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateData(ByRef guid As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateDataInterface(ByRef guid As Guid, pData As IntPtr) As HRESULT
#End Region
    End Interface

    <ComImport> <Guid("f582c508-0f36-490c-9977-31eece268cfa")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID3D11DomainShader
        Inherits ID3D11DeviceChild
#Region "ID3D11DeviceChild"
        <PreserveSig>
        Overloads Sub GetDevice(<Out> ByRef ppDevice As IntPtr)
        <PreserveSig>
        Overloads Function GetPrivateData(ByRef guid As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateData(ByRef guid As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateDataInterface(ByRef guid As Guid, pData As IntPtr) As HRESULT
#End Region
    End Interface

    <ComImport> <Guid("4f5b196e-c2bd-495e-bd01-1fded38e4969")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID3D11ComputeShader
        Inherits ID3D11DeviceChild
#Region "ID3D11DeviceChild"
        <PreserveSig>
        Overloads Sub GetDevice(<Out> ByRef ppDevice As IntPtr)
        <PreserveSig>
        Overloads Function GetPrivateData(ByRef guid As Guid, ByRef pDataSize As UInteger, <Out> ByRef pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateData(ByRef guid As Guid, DataSize As UInteger, pData As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function SetPrivateDataInterface(ByRef guid As Guid, pData As IntPtr) As HRESULT
#End Region
    End Interface

    Public Enum D3D11_DEVICE_CONTEXT_TYPE
        D3D11_DEVICE_CONTEXT_IMMEDIATE = 0
        D3D11_DEVICE_CONTEXT_DEFERRED = D3D11_DEVICE_CONTEXT_IMMEDIATE + 1
    End Enum



    <StructLayout(LayoutKind.Sequential)>
    Public Structure DXGI_MATRIX_3X2_F
        Public _11 As Single
        Public _12 As Single
        Public _21 As Single
        Public _22 As Single
        Public _31 As Single
        Public _32 As Single
    End Structure

    '[StructLayout(LayoutKind.Explicit, Size = 24)]
    'public class DXGI_MATRIX_3X2_F
    '{
    '    [FieldOffset(0)]
    '    public float _11;
    '    [FieldOffset(4)]
    '    public float _12;
    '    [FieldOffset(8)]
    '    public float _21;
    '    [FieldOffset(12)]
    '    public float _22;
    '    [FieldOffset(16)]
    '    public float _31;
    '    [FieldOffset(20)]
    '    public float _32;
    '}

    Public Enum D2D1_RENDERING_PRIORITY As UInteger
        D2D1_RENDERING_PRIORITY_NORMAL = 0
        D2D1_RENDERING_PRIORITY_LOW = 1
        D2D1_RENDERING_PRIORITY_FORCE_DWORD = &HFFFFFFFFUI
    End Enum


    <ComImport> <Guid("9eb767fd-4269-4467-b8c2-eb30cb305743")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1CommandSink1
        Inherits ID2D1CommandSink
#Region "ID2D1CommandSink"
        <PreserveSig>
        Overloads Function BeginDraw() As HRESULT
        <PreserveSig>
        Overloads Function EndDraw() As HRESULT
        <PreserveSig>
        Overloads Function SetAntialiasMode(antialiasMode As D2D1_ANTIALIAS_MODE) As HRESULT
        <PreserveSig>
        Overloads Function SetTags(tag1 As ULong, tag2 As ULong) As HRESULT
        <PreserveSig>
        Overloads Function SetTextAntialiasMode(textAntialiasMode As D2D1_TEXT_ANTIALIAS_MODE) As HRESULT
        Overloads Function SetTextRenderingParams(textRenderingParams As IDWriteRenderingParams) As HRESULT
        Overloads Function SetTransform(transform As D2D1_MATRIX_3X2_F) As HRESULT
        Overloads Function SetPrimitiveBlend(primitiveBlend As D2D1_PRIMITIVE_BLEND) As HRESULT
        Overloads Function SetUnitMode(unitMode As D2D1_UNIT_MODE) As HRESULT
        Overloads Function Clear(color As D2D1_COLOR_F) As HRESULT
        <PreserveSig>
        Overloads Function DrawGlyphRun(baselineOrigin As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, glyphRunDescription As IntPtr, foregroundBrush As ID2D1Brush, measuringMode As DWRITE_MEASURING_MODE) As HRESULT
        <PreserveSig>
        Overloads Function DrawLine(point0 As D2D1_POINT_2F, point1 As D2D1_POINT_2F, brush As ID2D1Brush, strokeWidth As Single, strokeStyle As ID2D1StrokeStyle) As HRESULT
        <PreserveSig>
        Overloads Function DrawGeometry(geometry As ID2D1Geometry, brush As ID2D1Brush, strokeWidth As Single, strokeStyle As ID2D1StrokeStyle) As HRESULT
        <PreserveSig>
        Overloads Function DrawRectangle(rect As D2D1_RECT_F, brush As ID2D1Brush, strokeWidth As Single, strokeStyle As ID2D1StrokeStyle) As HRESULT
        <PreserveSig>
        Overloads Function DrawBitmap(bitmap As ID2D1Bitmap, destinationRectangle As D2D1_RECT_F, opacity As Single, interpolationMode As D2D1_INTERPOLATION_MODE, sourceRectangle As D2D1_RECT_F, perspectiveTransform As D2D1_MATRIX_4X4_F) As HRESULT
        <PreserveSig>
        Overloads Function DrawImage(image As ID2D1Image, ByRef targetOffset As D2D1_POINT_2F, imageRectangle As D2D1_RECT_F, interpolationMode As D2D1_INTERPOLATION_MODE, compositeMode As D2D1_COMPOSITE_MODE) As HRESULT
        <PreserveSig>
        Overloads Function DrawGdiMetafile(gdiMetafile As ID2D1GdiMetafile, ByRef targetOffset As D2D1_POINT_2F) As HRESULT
        Overloads Function FillMesh(mesh As ID2D1Mesh, brush As ID2D1Brush) As HRESULT
        Overloads Function FillOpacityMask(opacityMask As ID2D1Bitmap, brush As ID2D1Brush, destinationRectangle As D2D1_RECT_F, sourceRectangle As D2D1_RECT_F) As HRESULT
        Overloads Function FillGeometry(geometry As ID2D1Geometry, brush As ID2D1Brush, opacityBrush As ID2D1Brush) As HRESULT
        Overloads Function FillRectangle(rect As D2D1_RECT_F, brush As ID2D1Brush) As HRESULT
        Overloads Function PushAxisAlignedClip(clipRect As D2D1_RECT_F, antialiasMode As D2D1_ANTIALIAS_MODE) As HRESULT
        Overloads Function PushLayer(ByRef layerParameters1 As D2D1_LAYER_PARAMETERS1, layer As ID2D1Layer) As HRESULT
        Overloads Function PopAxisAlignedClip() As HRESULT
        Overloads Function PopLayer() As HRESULT
#End Region

        Function SetPrimitiveBlend1(primitiveBlend As D2D1_PRIMITIVE_BLEND) As HRESULT
    End Interface

    <ComImport> <Guid("d21768e1-23a4-4823-a14b-7c3eba85d658")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1Device1
        Inherits ID2D1Device
#Region "<ID2D1Device>"
#Region "<ID2D1Resource>"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region

        <PreserveSig>
        Overloads Function CreateDeviceContext(options As D2D1_DEVICE_CONTEXT_OPTIONS, <Out> ByRef deviceContext As ID2D1DeviceContext) As HRESULT
        <PreserveSig>
        Overloads Function CreatePrintControl(wicFactory As IWICImagingFactory, documentTarget As IntPtr, ByRef printControlProperties As D2D1_PRINT_CONTROL_PROPERTIES, <Out> ByRef printControl As ID2D1PrintControl) As HRESULT
        <PreserveSig>
        Overloads Sub SetMaximumTextureMemory(maximumInBytes As ULong)
        <PreserveSig>
        Overloads Function GetMaximumTextureMemory() As ULong
        <PreserveSig>
        Overloads Sub ClearResources(Optional millisecondsSinceUse As UInteger = 0)
#End Region

        <PreserveSig>
        Sub GetRenderingPriority(<Out> ByRef renderingPriority As D2D1_RENDERING_PRIORITY)
        <PreserveSig>
        Sub SetRenderingPriority(renderingPriority As D2D1_RENDERING_PRIORITY)
        <PreserveSig>
        Overloads Function CreateDeviceContext(options As D2D1_DEVICE_CONTEXT_OPTIONS, <Out> ByRef deviceContext1 As ID2D1DeviceContext1) As HRESULT
    End Interface

    <ComImport> <Guid("94f81a73-9212-4376-9c58-b16a3a0d3992")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1Factory2
        Inherits ID2D1Factory1
#Region "<ID2D1Factory1>"
#Region "<ID2D1Factory>"
        <PreserveSig>
        Overloads Function ReloadSystemMetrics() As HRESULT
        <PreserveSig>
        Overloads Function GetDesktopDpi(<Out> ByRef dpiX As Single, <Out> ByRef dpiY As Single) As HRESULT
        <PreserveSig>
        Overloads Function CreateRectangleGeometry(ByRef rectangle As D2D1_RECT_F, <Out> ByRef rectangleGeometry As ID2D1RectangleGeometry) As HRESULT
        <PreserveSig>
        Overloads Function CreateRoundedRectangleGeometry(ByRef roundedRectangle As D2D1_ROUNDED_RECT, <Out> ByRef roundedRectangleGeometry As ID2D1RoundedRectangleGeometry) As HRESULT
        <PreserveSig>
        Overloads Function CreateEllipseGeometry(ByRef ellipse As D2D1_ELLIPSE, <Out> ByRef ellipseGeometry As ID2D1EllipseGeometry) As HRESULT
        <PreserveSig>
        Overloads Function CreateGeometryGroup(fillMode As D2D1_FILL_MODE, geometries As ID2D1Geometry, geometriesCount As UInteger, <Out> ByRef geometryGroup As ID2D1GeometryGroup) As HRESULT
        <PreserveSig>
        Overloads Function CreateTransformedGeometry(sourceGeometry As ID2D1Geometry, transform As D2D1_MATRIX_3X2_F, <Out> ByRef transformedGeometry As ID2D1TransformedGeometry) As HRESULT
        <PreserveSig>
        Overloads Function CreatePathGeometry(<Out> ByRef pathGeometry As ID2D1PathGeometry) As HRESULT
        <PreserveSig>
        Overloads Function CreateStrokeStyle(ByRef strokeStyleProperties As D2D1_STROKE_STYLE_PROPERTIES,
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=2)> dashes As Single(), Optional dashesCount As UInteger = 0) As ID2D1StrokeStyle
        <PreserveSig>
        Overloads Function CreateDrawingStateBlock(ByRef drawingStateDescription As D2D1_DRAWING_STATE_DESCRIPTION, textRenderingParams As IDWriteRenderingParams, <Out> ByRef drawingStateBlock As ID2D1DrawingStateBlock) As HRESULT
        <PreserveSig>
        Overloads Function CreateWicBitmapRenderTarget(target As IWICBitmap, renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES, <Out> ByRef renderTarget As ID2D1RenderTarget) As HRESULT
        <PreserveSig>
        Overloads Function CreateHwndRenderTarget(ByRef renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES, ByRef hwndRenderTargetProperties As D2D1_HWND_RENDER_TARGET_PROPERTIES, <Out> ByRef hwndRenderTarget As ID2D1HwndRenderTarget) As HRESULT
        <PreserveSig>
        Overloads Function CreateDxgiSurfaceRenderTarget(dxgiSurface As IntPtr, ByRef renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES, ByRef renderTarget As ID2D1RenderTarget) As HRESULT
        <PreserveSig>
        Overloads Function CreateDCRenderTarget(ByRef renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES, ByRef dcRenderTarget As ID2D1DCRenderTarget) As HRESULT
#End Region

        Overloads Function CreateDevice(dxgiDevice As IDXGIDevice, <Out> ByRef d2dDevice As ID2D1Device) As HRESULT
        <PreserveSig>
        Overloads Function CreateStrokeStyle(ByRef strokeStyleProperties As D2D1_STROKE_STYLE_PROPERTIES1,
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=2)> dashes As Single(), dashesCount As UInteger, <Out> ByRef strokeStyle As ID2D1StrokeStyle1) As HRESULT
        <PreserveSig>
        Overloads Function CreatePathGeometry(<Out> ByRef pathGeometry As ID2D1PathGeometry1) As HRESULT
        <PreserveSig>
        Overloads Function CreateDrawingStateBlock(ByRef drawingStateDescription As D2D1_DRAWING_STATE_DESCRIPTION1, ByRef textRenderingParams As IDWriteRenderingParams, <Out> ByRef drawingStateBlock As ID2D1DrawingStateBlock1) As HRESULT
        <PreserveSig>
        Overloads Function CreateGdiMetafile(metafileStream As ComTypes.IStream, <Out> ByRef metafile As ID2D1GdiMetafile) As HRESULT
        <PreserveSig>
        Overloads Function RegisterEffectFromStream(
<MarshalAs(UnmanagedType.LPStruct)> classId As Guid, propertyXml As ComTypes.IStream,
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=3)> bindings As D2D1_PROPERTY_BINDING(), bindingsCount As Integer, effectFactory As PD2D1_EFFECT_FACTORY) As HRESULT
        <PreserveSig>
        Overloads Function RegisterEffectFromString(
<MarshalAs(UnmanagedType.LPStruct)> classId As Guid,
<MarshalAs(UnmanagedType.LPWStr)> propertyXml As String,
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=3)> bindings As D2D1_PROPERTY_BINDING(), bindingsCount As Integer, effectFactory As PD2D1_EFFECT_FACTORY) As HRESULT
        <PreserveSig>
        Overloads Function UnregisterEffect(
<MarshalAs(UnmanagedType.LPStruct)> classId As Guid) As HRESULT
        <PreserveSig>
        Overloads Function GetRegisteredEffects(
<Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> effects As Guid(), effectsCount As Integer, effectsReturned As IntPtr, effectsRegistered As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function GetEffectProperties(
<MarshalAs(UnmanagedType.LPStruct)> effectId As Guid, <Out> ByRef properties As ID2D1Properties) As HRESULT

#End Region

        <PreserveSig>
        Overloads Function CreateDevice(dxgiDevice As IDXGIDevice, <Out> ByRef d2dDevice1 As ID2D1Device1) As HRESULT
    End Interface

    <ComImport> <Guid("0869759f-4f00-413f-b03e-2bda45404d0f")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1Factory3
        Inherits ID2D1Factory2
#Region "<ID2D1Factory2>"
#Region "<ID2D1Factory1>"
#Region "<ID2D1Factory>"
        <PreserveSig>
        Overloads Function ReloadSystemMetrics() As HRESULT
        <PreserveSig>
        Overloads Function GetDesktopDpi(<Out> ByRef dpiX As Single, <Out> ByRef dpiY As Single) As HRESULT
        <PreserveSig>
        Overloads Function CreateRectangleGeometry(ByRef rectangle As D2D1_RECT_F, <Out> ByRef rectangleGeometry As ID2D1RectangleGeometry) As HRESULT
        <PreserveSig>
        Overloads Function CreateRoundedRectangleGeometry(ByRef roundedRectangle As D2D1_ROUNDED_RECT, <Out> ByRef roundedRectangleGeometry As ID2D1RoundedRectangleGeometry) As HRESULT
        <PreserveSig>
        Overloads Function CreateEllipseGeometry(ByRef ellipse As D2D1_ELLIPSE, <Out> ByRef ellipseGeometry As ID2D1EllipseGeometry) As HRESULT
        <PreserveSig>
        Overloads Function CreateGeometryGroup(fillMode As D2D1_FILL_MODE, geometries As ID2D1Geometry, geometriesCount As UInteger, <Out> ByRef geometryGroup As ID2D1GeometryGroup) As HRESULT
        <PreserveSig>
        Overloads Function CreateTransformedGeometry(sourceGeometry As ID2D1Geometry, transform As D2D1_MATRIX_3X2_F, <Out> ByRef transformedGeometry As ID2D1TransformedGeometry) As HRESULT
        <PreserveSig>
        Overloads Function CreatePathGeometry(<Out> ByRef pathGeometry As ID2D1PathGeometry) As HRESULT
        <PreserveSig>
        Overloads Function CreateStrokeStyle(ByRef strokeStyleProperties As D2D1_STROKE_STYLE_PROPERTIES,
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=2)> dashes As Single(), Optional dashesCount As UInteger = 0) As ID2D1StrokeStyle
        <PreserveSig>
        Overloads Function CreateDrawingStateBlock(ByRef drawingStateDescription As D2D1_DRAWING_STATE_DESCRIPTION, textRenderingParams As IDWriteRenderingParams, <Out> ByRef drawingStateBlock As ID2D1DrawingStateBlock) As HRESULT
        <PreserveSig>
        Overloads Function CreateWicBitmapRenderTarget(target As IWICBitmap, renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES, <Out> ByRef renderTarget As ID2D1RenderTarget) As HRESULT
        <PreserveSig>
        Overloads Function CreateHwndRenderTarget(ByRef renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES, ByRef hwndRenderTargetProperties As D2D1_HWND_RENDER_TARGET_PROPERTIES, <Out> ByRef hwndRenderTarget As ID2D1HwndRenderTarget) As HRESULT
        <PreserveSig>
        Overloads Function CreateDxgiSurfaceRenderTarget(dxgiSurface As IntPtr, ByRef renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES, ByRef renderTarget As ID2D1RenderTarget) As HRESULT
        <PreserveSig>
        Overloads Function CreateDCRenderTarget(ByRef renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES, ByRef dcRenderTarget As ID2D1DCRenderTarget) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateDevice(dxgiDevice As IDXGIDevice, <Out> ByRef d2dDevice As ID2D1Device) As HRESULT
        <PreserveSig>
        Overloads Function CreateStrokeStyle(ByRef strokeStyleProperties As D2D1_STROKE_STYLE_PROPERTIES1,
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=2)> dashes As Single(), dashesCount As UInteger, <Out> ByRef strokeStyle As ID2D1StrokeStyle1) As HRESULT
        <PreserveSig>
        Overloads Function CreatePathGeometry(<Out> ByRef pathGeometry As ID2D1PathGeometry1) As HRESULT
        <PreserveSig>
        Overloads Function CreateDrawingStateBlock(ByRef drawingStateDescription As D2D1_DRAWING_STATE_DESCRIPTION1, ByRef textRenderingParams As IDWriteRenderingParams, <Out> ByRef drawingStateBlock As ID2D1DrawingStateBlock1) As HRESULT
        <PreserveSig>
        Overloads Function CreateGdiMetafile(metafileStream As ComTypes.IStream, <Out> ByRef metafile As ID2D1GdiMetafile) As HRESULT
        <PreserveSig>
        Overloads Function RegisterEffectFromStream(
<MarshalAs(UnmanagedType.LPStruct)> classId As Guid, propertyXml As ComTypes.IStream,
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=3)> bindings As D2D1_PROPERTY_BINDING(), bindingsCount As Integer, effectFactory As PD2D1_EFFECT_FACTORY) As HRESULT
        <PreserveSig>
        Overloads Function RegisterEffectFromString(
<MarshalAs(UnmanagedType.LPStruct)> classId As Guid,
<MarshalAs(UnmanagedType.LPWStr)> propertyXml As String,
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=3)> bindings As D2D1_PROPERTY_BINDING(), bindingsCount As Integer, effectFactory As PD2D1_EFFECT_FACTORY) As HRESULT
        <PreserveSig>
        Overloads Function UnregisterEffect(
<MarshalAs(UnmanagedType.LPStruct)> classId As Guid) As HRESULT
        <PreserveSig>
        Overloads Function GetRegisteredEffects(
<Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> effects As Guid(), effectsCount As Integer, effectsReturned As IntPtr, effectsRegistered As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function GetEffectProperties(
<MarshalAs(UnmanagedType.LPStruct)> effectId As Guid, <Out> ByRef properties As ID2D1Properties) As HRESULT

#End Region

        <PreserveSig>
        Overloads Function CreateDevice(dxgiDevice As IDXGIDevice, <Out> ByRef d2dDevice1 As ID2D1Device1) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateDevice(dxgiDevice As IDXGIDevice, <Out> ByRef d2dDevice2 As ID2D1Device2) As HRESULT
    End Interface

    <ComImport> <Guid("bd4ec2d2-0662-4bee-ba8e-6f29f032e096")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1Factory4
        Inherits ID2D1Factory3
#Region "<ID2D1Factory3>"
#Region "<ID2D1Factory2>"
#Region "<ID2D1Factory1>"
#Region "<ID2D1Factory>"
        <PreserveSig>
        Overloads Function ReloadSystemMetrics() As HRESULT
        <PreserveSig>
        Overloads Function GetDesktopDpi(<Out> ByRef dpiX As Single, <Out> ByRef dpiY As Single) As HRESULT
        <PreserveSig>
        Overloads Function CreateRectangleGeometry(ByRef rectangle As D2D1_RECT_F, <Out> ByRef rectangleGeometry As ID2D1RectangleGeometry) As HRESULT
        <PreserveSig>
        Overloads Function CreateRoundedRectangleGeometry(ByRef roundedRectangle As D2D1_ROUNDED_RECT, <Out> ByRef roundedRectangleGeometry As ID2D1RoundedRectangleGeometry) As HRESULT
        <PreserveSig>
        Overloads Function CreateEllipseGeometry(ByRef ellipse As D2D1_ELLIPSE, <Out> ByRef ellipseGeometry As ID2D1EllipseGeometry) As HRESULT
        <PreserveSig>
        Overloads Function CreateGeometryGroup(fillMode As D2D1_FILL_MODE, geometries As ID2D1Geometry, geometriesCount As UInteger, <Out> ByRef geometryGroup As ID2D1GeometryGroup) As HRESULT
        <PreserveSig>
        Overloads Function CreateTransformedGeometry(sourceGeometry As ID2D1Geometry, transform As D2D1_MATRIX_3X2_F, <Out> ByRef transformedGeometry As ID2D1TransformedGeometry) As HRESULT
        <PreserveSig>
        Overloads Function CreatePathGeometry(<Out> ByRef pathGeometry As ID2D1PathGeometry) As HRESULT
        <PreserveSig>
        Overloads Function CreateStrokeStyle(ByRef strokeStyleProperties As D2D1_STROKE_STYLE_PROPERTIES,
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=2)> dashes As Single(), Optional dashesCount As UInteger = 0) As ID2D1StrokeStyle
        <PreserveSig>
        Overloads Function CreateDrawingStateBlock(ByRef drawingStateDescription As D2D1_DRAWING_STATE_DESCRIPTION, textRenderingParams As IDWriteRenderingParams, <Out> ByRef drawingStateBlock As ID2D1DrawingStateBlock) As HRESULT
        <PreserveSig>
        Overloads Function CreateWicBitmapRenderTarget(target As IWICBitmap, renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES, <Out> ByRef renderTarget As ID2D1RenderTarget) As HRESULT
        <PreserveSig>
        Overloads Function CreateHwndRenderTarget(ByRef renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES, ByRef hwndRenderTargetProperties As D2D1_HWND_RENDER_TARGET_PROPERTIES, <Out> ByRef hwndRenderTarget As ID2D1HwndRenderTarget) As HRESULT
        <PreserveSig>
        Overloads Function CreateDxgiSurfaceRenderTarget(dxgiSurface As IntPtr, ByRef renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES, ByRef renderTarget As ID2D1RenderTarget) As HRESULT
        <PreserveSig>
        Overloads Function CreateDCRenderTarget(ByRef renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES, ByRef dcRenderTarget As ID2D1DCRenderTarget) As HRESULT
#End Region

        Overloads Function CreateDevice(dxgiDevice As IDXGIDevice, <Out> ByRef d2dDevice As ID2D1Device) As HRESULT
        <PreserveSig>
        Overloads Function CreateStrokeStyle(ByRef strokeStyleProperties As D2D1_STROKE_STYLE_PROPERTIES1,
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=2)> dashes As Single(), dashesCount As UInteger, <Out> ByRef strokeStyle As ID2D1StrokeStyle1) As HRESULT
        <PreserveSig>
        Overloads Function CreatePathGeometry(<Out> ByRef pathGeometry As ID2D1PathGeometry1) As HRESULT
        <PreserveSig>
        Overloads Function CreateDrawingStateBlock(ByRef drawingStateDescription As D2D1_DRAWING_STATE_DESCRIPTION1, ByRef textRenderingParams As IDWriteRenderingParams, <Out> ByRef drawingStateBlock As ID2D1DrawingStateBlock1) As HRESULT
        <PreserveSig>
        Overloads Function CreateGdiMetafile(metafileStream As ComTypes.IStream, <Out> ByRef metafile As ID2D1GdiMetafile) As HRESULT
        <PreserveSig>
        Overloads Function RegisterEffectFromStream(
<MarshalAs(UnmanagedType.LPStruct)> classId As Guid, propertyXml As ComTypes.IStream,
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=3)> bindings As D2D1_PROPERTY_BINDING(), bindingsCount As Integer, effectFactory As PD2D1_EFFECT_FACTORY) As HRESULT
        <PreserveSig>
        Overloads Function RegisterEffectFromString(
<MarshalAs(UnmanagedType.LPStruct)> classId As Guid,
<MarshalAs(UnmanagedType.LPWStr)> propertyXml As String,
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=3)> bindings As D2D1_PROPERTY_BINDING(), bindingsCount As Integer, effectFactory As PD2D1_EFFECT_FACTORY) As HRESULT
        <PreserveSig>
        Overloads Function UnregisterEffect(
<MarshalAs(UnmanagedType.LPStruct)> classId As Guid) As HRESULT
        <PreserveSig>
        Overloads Function GetRegisteredEffects(
<Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> effects As Guid(), effectsCount As Integer, effectsReturned As IntPtr, effectsRegistered As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function GetEffectProperties(
<MarshalAs(UnmanagedType.LPStruct)> effectId As Guid, <Out> ByRef properties As ID2D1Properties) As HRESULT

#End Region

        <PreserveSig>
        Overloads Function CreateDevice(dxgiDevice As IDXGIDevice, <Out> ByRef d2dDevice1 As ID2D1Device1) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateDevice(dxgiDevice As IDXGIDevice, <Out> ByRef d2dDevice2 As ID2D1Device2) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateDevice(dxgiDevice As IDXGIDevice, <Out> ByRef d2dDevice3 As ID2D1Device3) As HRESULT
    End Interface

    <ComImport> <Guid("c4349994-838e-4b0f-8cab-44997d9eeacc")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1Factory5
        Inherits ID2D1Factory4
#Region "<ID2D1Factory4>"
#Region "<ID2D1Factory3>"
#Region "<ID2D1Factory2>"
#Region "<ID2D1Factory1>"
#Region "<ID2D1Factory>"
        <PreserveSig>
        Overloads Function ReloadSystemMetrics() As HRESULT
        <PreserveSig>
        Overloads Function GetDesktopDpi(<Out> ByRef dpiX As Single, <Out> ByRef dpiY As Single) As HRESULT
        <PreserveSig>
        Overloads Function CreateRectangleGeometry(ByRef rectangle As D2D1_RECT_F, <Out> ByRef rectangleGeometry As ID2D1RectangleGeometry) As HRESULT
        <PreserveSig>
        Overloads Function CreateRoundedRectangleGeometry(ByRef roundedRectangle As D2D1_ROUNDED_RECT, <Out> ByRef roundedRectangleGeometry As ID2D1RoundedRectangleGeometry) As HRESULT
        <PreserveSig>
        Overloads Function CreateEllipseGeometry(ByRef ellipse As D2D1_ELLIPSE, <Out> ByRef ellipseGeometry As ID2D1EllipseGeometry) As HRESULT
        <PreserveSig>
        Overloads Function CreateGeometryGroup(fillMode As D2D1_FILL_MODE, geometries As ID2D1Geometry, geometriesCount As UInteger, <Out> ByRef geometryGroup As ID2D1GeometryGroup) As HRESULT
        <PreserveSig>
        Overloads Function CreateTransformedGeometry(sourceGeometry As ID2D1Geometry, transform As D2D1_MATRIX_3X2_F, <Out> ByRef transformedGeometry As ID2D1TransformedGeometry) As HRESULT
        <PreserveSig>
        Overloads Function CreatePathGeometry(<Out> ByRef pathGeometry As ID2D1PathGeometry) As HRESULT
        <PreserveSig>
        Overloads Function CreateStrokeStyle(ByRef strokeStyleProperties As D2D1_STROKE_STYLE_PROPERTIES,
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=2)> dashes As Single(), Optional dashesCount As UInteger = 0) As ID2D1StrokeStyle
        <PreserveSig>
        Overloads Function CreateDrawingStateBlock(ByRef drawingStateDescription As D2D1_DRAWING_STATE_DESCRIPTION, textRenderingParams As IDWriteRenderingParams, <Out> ByRef drawingStateBlock As ID2D1DrawingStateBlock) As HRESULT
        <PreserveSig>
        Overloads Function CreateWicBitmapRenderTarget(target As IWICBitmap, renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES, <Out> ByRef renderTarget As ID2D1RenderTarget) As HRESULT
        <PreserveSig>
        Overloads Function CreateHwndRenderTarget(ByRef renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES, ByRef hwndRenderTargetProperties As D2D1_HWND_RENDER_TARGET_PROPERTIES, <Out> ByRef hwndRenderTarget As ID2D1HwndRenderTarget) As HRESULT
        <PreserveSig>
        Overloads Function CreateDxgiSurfaceRenderTarget(dxgiSurface As IntPtr, ByRef renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES, ByRef renderTarget As ID2D1RenderTarget) As HRESULT
        <PreserveSig>
        Overloads Function CreateDCRenderTarget(ByRef renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES, ByRef dcRenderTarget As ID2D1DCRenderTarget) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateDevice(dxgiDevice As IDXGIDevice, <Out> ByRef d2dDevice As ID2D1Device) As HRESULT
        <PreserveSig>
        Overloads Function CreateStrokeStyle(ByRef strokeStyleProperties As D2D1_STROKE_STYLE_PROPERTIES1,
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=2)> dashes As Single(), dashesCount As UInteger, <Out> ByRef strokeStyle As ID2D1StrokeStyle1) As HRESULT
        <PreserveSig>
        Overloads Function CreatePathGeometry(<Out> ByRef pathGeometry As ID2D1PathGeometry1) As HRESULT
        <PreserveSig>
        Overloads Function CreateDrawingStateBlock(ByRef drawingStateDescription As D2D1_DRAWING_STATE_DESCRIPTION1, ByRef textRenderingParams As IDWriteRenderingParams, <Out> ByRef drawingStateBlock As ID2D1DrawingStateBlock1) As HRESULT
        <PreserveSig>
        Overloads Function CreateGdiMetafile(metafileStream As ComTypes.IStream, <Out> ByRef metafile As ID2D1GdiMetafile) As HRESULT
        <PreserveSig>
        Overloads Function RegisterEffectFromStream(
<MarshalAs(UnmanagedType.LPStruct)> classId As Guid, propertyXml As ComTypes.IStream,
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=3)> bindings As D2D1_PROPERTY_BINDING(), bindingsCount As Integer, effectFactory As PD2D1_EFFECT_FACTORY) As HRESULT
        <PreserveSig>
        Overloads Function RegisterEffectFromString(
<MarshalAs(UnmanagedType.LPStruct)> classId As Guid,
<MarshalAs(UnmanagedType.LPWStr)> propertyXml As String,
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=3)> bindings As D2D1_PROPERTY_BINDING(), bindingsCount As Integer, effectFactory As PD2D1_EFFECT_FACTORY) As HRESULT
        <PreserveSig>
        Overloads Function UnregisterEffect(
<MarshalAs(UnmanagedType.LPStruct)> classId As Guid) As HRESULT
        <PreserveSig>
        Overloads Function GetRegisteredEffects(
<Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> effects As Guid(), effectsCount As Integer, effectsReturned As IntPtr, effectsRegistered As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function GetEffectProperties(
<MarshalAs(UnmanagedType.LPStruct)> effectId As Guid, <Out> ByRef properties As ID2D1Properties) As HRESULT

#End Region

        <PreserveSig>
        Overloads Function CreateDevice(dxgiDevice As IDXGIDevice, <Out> ByRef d2dDevice1 As ID2D1Device1) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateDevice(dxgiDevice As IDXGIDevice, <Out> ByRef d2dDevice2 As ID2D1Device2) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateDevice(dxgiDevice As IDXGIDevice, <Out> ByRef d2dDevice3 As ID2D1Device3) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateDevice(dxgiDevice As IDXGIDevice, <Out> ByRef d2dDevice4 As ID2D1Device4) As HRESULT
    End Interface

    <ComImport> <Guid("f9976f46-f642-44c1-97ca-da32ea2a2635")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1Factory6
        Inherits ID2D1Factory5
#Region "<ID2D1Factory5>"
#Region "<ID2D1Factory4>"
#Region "<ID2D1Factory3>"
#Region "<ID2D1Factory2>"
#Region "<ID2D1Factory1>"
#Region "<ID2D1Factory>"
        <PreserveSig>
        Overloads Function ReloadSystemMetrics() As HRESULT
        <PreserveSig>
        Overloads Function GetDesktopDpi(<Out> ByRef dpiX As Single, <Out> ByRef dpiY As Single) As HRESULT
        <PreserveSig>
        Overloads Function CreateRectangleGeometry(ByRef rectangle As D2D1_RECT_F, <Out> ByRef rectangleGeometry As ID2D1RectangleGeometry) As HRESULT
        <PreserveSig>
        Overloads Function CreateRoundedRectangleGeometry(ByRef roundedRectangle As D2D1_ROUNDED_RECT, <Out> ByRef roundedRectangleGeometry As ID2D1RoundedRectangleGeometry) As HRESULT
        <PreserveSig>
        Overloads Function CreateEllipseGeometry(ByRef ellipse As D2D1_ELLIPSE, <Out> ByRef ellipseGeometry As ID2D1EllipseGeometry) As HRESULT
        <PreserveSig>
        Overloads Function CreateGeometryGroup(fillMode As D2D1_FILL_MODE, geometries As ID2D1Geometry, geometriesCount As UInteger, <Out> ByRef geometryGroup As ID2D1GeometryGroup) As HRESULT
        <PreserveSig>
        Overloads Function CreateTransformedGeometry(sourceGeometry As ID2D1Geometry, transform As D2D1_MATRIX_3X2_F, <Out> ByRef transformedGeometry As ID2D1TransformedGeometry) As HRESULT
        <PreserveSig>
        Overloads Function CreatePathGeometry(<Out> ByRef pathGeometry As ID2D1PathGeometry) As HRESULT
        <PreserveSig>
        Overloads Function CreateStrokeStyle(ByRef strokeStyleProperties As D2D1_STROKE_STYLE_PROPERTIES,
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=2)> dashes As Single(), Optional dashesCount As UInteger = 0) As ID2D1StrokeStyle
        <PreserveSig>
        Overloads Function CreateDrawingStateBlock(ByRef drawingStateDescription As D2D1_DRAWING_STATE_DESCRIPTION, textRenderingParams As IDWriteRenderingParams, <Out> ByRef drawingStateBlock As ID2D1DrawingStateBlock) As HRESULT
        <PreserveSig>
        Overloads Function CreateWicBitmapRenderTarget(target As IWICBitmap, renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES, <Out> ByRef renderTarget As ID2D1RenderTarget) As HRESULT
        <PreserveSig>
        Overloads Function CreateHwndRenderTarget(ByRef renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES, ByRef hwndRenderTargetProperties As D2D1_HWND_RENDER_TARGET_PROPERTIES, <Out> ByRef hwndRenderTarget As ID2D1HwndRenderTarget) As HRESULT
        <PreserveSig>
        Overloads Function CreateDxgiSurfaceRenderTarget(dxgiSurface As IntPtr, ByRef renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES, ByRef renderTarget As ID2D1RenderTarget) As HRESULT
        <PreserveSig>
        Overloads Function CreateDCRenderTarget(ByRef renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES, ByRef dcRenderTarget As ID2D1DCRenderTarget) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateDevice(dxgiDevice As IDXGIDevice, <Out> ByRef d2dDevice As ID2D1Device) As HRESULT
        <PreserveSig>
        Overloads Function CreateStrokeStyle(ByRef strokeStyleProperties As D2D1_STROKE_STYLE_PROPERTIES1,
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=2)> dashes As Single(), dashesCount As UInteger, <Out> ByRef strokeStyle As ID2D1StrokeStyle1) As HRESULT
        <PreserveSig>
        Overloads Function CreatePathGeometry(<Out> ByRef pathGeometry As ID2D1PathGeometry1) As HRESULT
        <PreserveSig>
        Overloads Function CreateDrawingStateBlock(ByRef drawingStateDescription As D2D1_DRAWING_STATE_DESCRIPTION1, ByRef textRenderingParams As IDWriteRenderingParams, <Out> ByRef drawingStateBlock As ID2D1DrawingStateBlock1) As HRESULT
        <PreserveSig>
        Overloads Function CreateGdiMetafile(metafileStream As ComTypes.IStream, <Out> ByRef metafile As ID2D1GdiMetafile) As HRESULT
        <PreserveSig>
        Overloads Function RegisterEffectFromStream(
<MarshalAs(UnmanagedType.LPStruct)> classId As Guid, propertyXml As ComTypes.IStream,
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=3)> bindings As D2D1_PROPERTY_BINDING(), bindingsCount As Integer, effectFactory As PD2D1_EFFECT_FACTORY) As HRESULT
        <PreserveSig>
        Overloads Function RegisterEffectFromString(
<MarshalAs(UnmanagedType.LPStruct)> classId As Guid,
<MarshalAs(UnmanagedType.LPWStr)> propertyXml As String,
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=3)> bindings As D2D1_PROPERTY_BINDING(), bindingsCount As Integer, effectFactory As PD2D1_EFFECT_FACTORY) As HRESULT
        <PreserveSig>
        Overloads Function UnregisterEffect(
<MarshalAs(UnmanagedType.LPStruct)> classId As Guid) As HRESULT
        <PreserveSig>
        Overloads Function GetRegisteredEffects(
<Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> effects As Guid(), effectsCount As Integer, effectsReturned As IntPtr, effectsRegistered As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function GetEffectProperties(
<MarshalAs(UnmanagedType.LPStruct)> effectId As Guid, <Out> ByRef properties As ID2D1Properties) As HRESULT

#End Region

        <PreserveSig>
        Overloads Function CreateDevice(dxgiDevice As IDXGIDevice, <Out> ByRef d2dDevice1 As ID2D1Device1) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateDevice(dxgiDevice As IDXGIDevice, <Out> ByRef d2dDevice2 As ID2D1Device2) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateDevice(dxgiDevice As IDXGIDevice, <Out> ByRef d2dDevice3 As ID2D1Device3) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateDevice(dxgiDevice As IDXGIDevice, <Out> ByRef d2dDevice4 As ID2D1Device4) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateDevice(dxgiDevice As IDXGIDevice, <Out> ByRef d2dDevice5 As ID2D1Device5) As HRESULT
    End Interface

    <ComImport> <Guid("bdc2bdd3-b96c-4de6-bdf7-99d4745454de")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1Factory7
        Inherits ID2D1Factory6
#Region "<ID2D1Factory6>"
#Region "<ID2D1Factory5>"
#Region "<ID2D1Factory4>"
#Region "<ID2D1Factory3>"
#Region "<ID2D1Factory2>"
#Region "<ID2D1Factory1>"
#Region "<ID2D1Factory>"
        <PreserveSig>
        Overloads Function ReloadSystemMetrics() As HRESULT
        <PreserveSig>
        Overloads Function GetDesktopDpi(<Out> ByRef dpiX As Single, <Out> ByRef dpiY As Single) As HRESULT
        <PreserveSig>
        Overloads Function CreateRectangleGeometry(ByRef rectangle As D2D1_RECT_F, <Out> ByRef rectangleGeometry As ID2D1RectangleGeometry) As HRESULT
        <PreserveSig>
        Overloads Function CreateRoundedRectangleGeometry(ByRef roundedRectangle As D2D1_ROUNDED_RECT, <Out> ByRef roundedRectangleGeometry As ID2D1RoundedRectangleGeometry) As HRESULT
        <PreserveSig>
        Overloads Function CreateEllipseGeometry(ByRef ellipse As D2D1_ELLIPSE, <Out> ByRef ellipseGeometry As ID2D1EllipseGeometry) As HRESULT
        <PreserveSig>
        Overloads Function CreateGeometryGroup(fillMode As D2D1_FILL_MODE, geometries As ID2D1Geometry, geometriesCount As UInteger, <Out> ByRef geometryGroup As ID2D1GeometryGroup) As HRESULT
        <PreserveSig>
        Overloads Function CreateTransformedGeometry(sourceGeometry As ID2D1Geometry, transform As D2D1_MATRIX_3X2_F, <Out> ByRef transformedGeometry As ID2D1TransformedGeometry) As HRESULT
        <PreserveSig>
        Overloads Function CreatePathGeometry(<Out> ByRef pathGeometry As ID2D1PathGeometry) As HRESULT
        <PreserveSig>
        Overloads Function CreateStrokeStyle(ByRef strokeStyleProperties As D2D1_STROKE_STYLE_PROPERTIES,
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=2)> dashes As Single(), Optional dashesCount As UInteger = 0) As ID2D1StrokeStyle
        <PreserveSig>
        Overloads Function CreateDrawingStateBlock(ByRef drawingStateDescription As D2D1_DRAWING_STATE_DESCRIPTION, textRenderingParams As IDWriteRenderingParams, <Out> ByRef drawingStateBlock As ID2D1DrawingStateBlock) As HRESULT
        <PreserveSig>
        Overloads Function CreateWicBitmapRenderTarget(target As IWICBitmap, renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES, <Out> ByRef renderTarget As ID2D1RenderTarget) As HRESULT
        <PreserveSig>
        Overloads Function CreateHwndRenderTarget(ByRef renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES, ByRef hwndRenderTargetProperties As D2D1_HWND_RENDER_TARGET_PROPERTIES, <Out> ByRef hwndRenderTarget As ID2D1HwndRenderTarget) As HRESULT
        <PreserveSig>
        Overloads Function CreateDxgiSurfaceRenderTarget(dxgiSurface As IntPtr, ByRef renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES, ByRef renderTarget As ID2D1RenderTarget) As HRESULT
        <PreserveSig>
        Overloads Function CreateDCRenderTarget(ByRef renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES, ByRef dcRenderTarget As ID2D1DCRenderTarget) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateDevice(dxgiDevice As IDXGIDevice, <Out> ByRef d2dDevice As ID2D1Device) As HRESULT
        <PreserveSig>
        Overloads Function CreateStrokeStyle(ByRef strokeStyleProperties As D2D1_STROKE_STYLE_PROPERTIES1,
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=2)> dashes As Single(), dashesCount As UInteger, <Out> ByRef strokeStyle As ID2D1StrokeStyle1) As HRESULT
        <PreserveSig>
        Overloads Function CreatePathGeometry(<Out> ByRef pathGeometry As ID2D1PathGeometry1) As HRESULT
        <PreserveSig>
        Overloads Function CreateDrawingStateBlock(ByRef drawingStateDescription As D2D1_DRAWING_STATE_DESCRIPTION1, ByRef textRenderingParams As IDWriteRenderingParams, <Out> ByRef drawingStateBlock As ID2D1DrawingStateBlock1) As HRESULT
        <PreserveSig>
        Overloads Function CreateGdiMetafile(metafileStream As ComTypes.IStream, <Out> ByRef metafile As ID2D1GdiMetafile) As HRESULT
        <PreserveSig>
        Overloads Function RegisterEffectFromStream(
<MarshalAs(UnmanagedType.LPStruct)> classId As Guid, propertyXml As ComTypes.IStream,
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=3)> bindings As D2D1_PROPERTY_BINDING(), bindingsCount As Integer, effectFactory As PD2D1_EFFECT_FACTORY) As HRESULT
        <PreserveSig>
        Overloads Function RegisterEffectFromString(
<MarshalAs(UnmanagedType.LPStruct)> classId As Guid,
<MarshalAs(UnmanagedType.LPWStr)> propertyXml As String,
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=3)> bindings As D2D1_PROPERTY_BINDING(), bindingsCount As Integer, effectFactory As PD2D1_EFFECT_FACTORY) As HRESULT
        <PreserveSig>
        Overloads Function UnregisterEffect(
<MarshalAs(UnmanagedType.LPStruct)> classId As Guid) As HRESULT
        <PreserveSig>
        Overloads Function GetRegisteredEffects(
<Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> effects As Guid(), effectsCount As Integer, effectsReturned As IntPtr, effectsRegistered As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function GetEffectProperties(
<MarshalAs(UnmanagedType.LPStruct)> effectId As Guid, <Out> ByRef properties As ID2D1Properties) As HRESULT

#End Region

        <PreserveSig>
        Overloads Function CreateDevice(dxgiDevice As IDXGIDevice, <Out> ByRef d2dDevice1 As ID2D1Device1) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateDevice(dxgiDevice As IDXGIDevice, <Out> ByRef d2dDevice2 As ID2D1Device2) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateDevice(dxgiDevice As IDXGIDevice, <Out> ByRef d2dDevice3 As ID2D1Device3) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateDevice(dxgiDevice As IDXGIDevice, <Out> ByRef d2dDevice4 As ID2D1Device4) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateDevice(dxgiDevice As IDXGIDevice, <Out> ByRef d2dDevice5 As ID2D1Device5) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateDevice(dxgiDevice As IDXGIDevice, <Out> ByRef d2dDevice6 As ID2D1Device6) As HRESULT
    End Interface

    <ComImport> <Guid("677c9311-f36d-4b1f-ae86-86d1223ffd3a")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1Factory8
        Inherits ID2D1Factory7
#Region "<ID2D1Factory6>"
#Region "<ID2D1Factory5>"
#Region "<ID2D1Factory4>"
#Region "<ID2D1Factory3>"
#Region "<ID2D1Factory2>"
#Region "<ID2D1Factory1>"
#Region "<ID2D1Factory>"
        <PreserveSig>
        Overloads Function ReloadSystemMetrics() As HRESULT
        <PreserveSig>
        Overloads Function GetDesktopDpi(<Out> ByRef dpiX As Single, <Out> ByRef dpiY As Single) As HRESULT
        <PreserveSig>
        Overloads Function CreateRectangleGeometry(ByRef rectangle As D2D1_RECT_F, <Out> ByRef rectangleGeometry As ID2D1RectangleGeometry) As HRESULT
        <PreserveSig>
        Overloads Function CreateRoundedRectangleGeometry(ByRef roundedRectangle As D2D1_ROUNDED_RECT, <Out> ByRef roundedRectangleGeometry As ID2D1RoundedRectangleGeometry) As HRESULT
        <PreserveSig>
        Overloads Function CreateEllipseGeometry(ByRef ellipse As D2D1_ELLIPSE, <Out> ByRef ellipseGeometry As ID2D1EllipseGeometry) As HRESULT
        <PreserveSig>
        Overloads Function CreateGeometryGroup(fillMode As D2D1_FILL_MODE, geometries As ID2D1Geometry, geometriesCount As UInteger, <Out> ByRef geometryGroup As ID2D1GeometryGroup) As HRESULT
        <PreserveSig>
        Overloads Function CreateTransformedGeometry(sourceGeometry As ID2D1Geometry, transform As D2D1_MATRIX_3X2_F, <Out> ByRef transformedGeometry As ID2D1TransformedGeometry) As HRESULT
        <PreserveSig>
        Overloads Function CreatePathGeometry(<Out> ByRef pathGeometry As ID2D1PathGeometry) As HRESULT
        <PreserveSig>
        Overloads Function CreateStrokeStyle(ByRef strokeStyleProperties As D2D1_STROKE_STYLE_PROPERTIES,
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=2)> dashes As Single(), Optional dashesCount As UInteger = 0) As ID2D1StrokeStyle
        <PreserveSig>
        Overloads Function CreateDrawingStateBlock(ByRef drawingStateDescription As D2D1_DRAWING_STATE_DESCRIPTION, textRenderingParams As IDWriteRenderingParams, <Out> ByRef drawingStateBlock As ID2D1DrawingStateBlock) As HRESULT
        <PreserveSig>
        Overloads Function CreateWicBitmapRenderTarget(target As IWICBitmap, renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES, <Out> ByRef renderTarget As ID2D1RenderTarget) As HRESULT
        <PreserveSig>
        Overloads Function CreateHwndRenderTarget(ByRef renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES, ByRef hwndRenderTargetProperties As D2D1_HWND_RENDER_TARGET_PROPERTIES, <Out> ByRef hwndRenderTarget As ID2D1HwndRenderTarget) As HRESULT
        <PreserveSig>
        Overloads Function CreateDxgiSurfaceRenderTarget(dxgiSurface As IntPtr, ByRef renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES, ByRef renderTarget As ID2D1RenderTarget) As HRESULT
        <PreserveSig>
        Overloads Function CreateDCRenderTarget(ByRef renderTargetProperties As D2D1_RENDER_TARGET_PROPERTIES, ByRef dcRenderTarget As ID2D1DCRenderTarget) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateDevice(dxgiDevice As IDXGIDevice, <Out> ByRef d2dDevice As ID2D1Device) As HRESULT
        <PreserveSig>
        Overloads Function CreateStrokeStyle(ByRef strokeStyleProperties As D2D1_STROKE_STYLE_PROPERTIES1,
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=2)> dashes As Single(), dashesCount As UInteger, <Out> ByRef strokeStyle As ID2D1StrokeStyle1) As HRESULT
        <PreserveSig>
        Overloads Function CreatePathGeometry(<Out> ByRef pathGeometry As ID2D1PathGeometry1) As HRESULT
        <PreserveSig>
        Overloads Function CreateDrawingStateBlock(ByRef drawingStateDescription As D2D1_DRAWING_STATE_DESCRIPTION1, ByRef textRenderingParams As IDWriteRenderingParams, <Out> ByRef drawingStateBlock As ID2D1DrawingStateBlock1) As HRESULT
        <PreserveSig>
        Overloads Function CreateGdiMetafile(metafileStream As ComTypes.IStream, <Out> ByRef metafile As ID2D1GdiMetafile) As HRESULT
        <PreserveSig>
        Overloads Function RegisterEffectFromStream(
<MarshalAs(UnmanagedType.LPStruct)> classId As Guid, propertyXml As ComTypes.IStream,
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=3)> bindings As D2D1_PROPERTY_BINDING(), bindingsCount As Integer, effectFactory As PD2D1_EFFECT_FACTORY) As HRESULT
        <PreserveSig>
        Overloads Function RegisterEffectFromString(
<MarshalAs(UnmanagedType.LPStruct)> classId As Guid,
<MarshalAs(UnmanagedType.LPWStr)> propertyXml As String,
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=3)> bindings As D2D1_PROPERTY_BINDING(), bindingsCount As Integer, effectFactory As PD2D1_EFFECT_FACTORY) As HRESULT
        <PreserveSig>
        Overloads Function UnregisterEffect(
<MarshalAs(UnmanagedType.LPStruct)> classId As Guid) As HRESULT
        <PreserveSig>
        Overloads Function GetRegisteredEffects(
<Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> effects As Guid(), effectsCount As Integer, effectsReturned As IntPtr, effectsRegistered As IntPtr) As HRESULT
        <PreserveSig>
        Overloads Function GetEffectProperties(
<MarshalAs(UnmanagedType.LPStruct)> effectId As Guid, <Out> ByRef properties As ID2D1Properties) As HRESULT

#End Region

        <PreserveSig>
        Overloads Function CreateDevice(dxgiDevice As IDXGIDevice, <Out> ByRef d2dDevice1 As ID2D1Device1) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateDevice(dxgiDevice As IDXGIDevice, <Out> ByRef d2dDevice2 As ID2D1Device2) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateDevice(dxgiDevice As IDXGIDevice, <Out> ByRef d2dDevice3 As ID2D1Device3) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateDevice(dxgiDevice As IDXGIDevice, <Out> ByRef d2dDevice4 As ID2D1Device4) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateDevice(dxgiDevice As IDXGIDevice, <Out> ByRef d2dDevice5 As ID2D1Device5) As HRESULT
#End Region

#Region "<ID2D1Factory7>"
        <PreserveSig>
        Overloads Function CreateDevice(dxgiDevice As IDXGIDevice, <Out> ByRef d2dDevice6 As ID2D1Device6) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateDevice(dxgiDevice As IDXGIDevice, <Out> ByRef d2dDevice7 As ID2D1Device7) As HRESULT
    End Interface

    <ComImport> <Guid("3bab440e-417e-47df-a2e2-bc0be6a00916")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1CommandSink2
        Inherits ID2D1CommandSink1
#Region "ID2D1CommandSink1"
#Region "ID2D1CommandSink"
        <PreserveSig>
        Overloads Function BeginDraw() As HRESULT
        <PreserveSig>
        Overloads Function EndDraw() As HRESULT
        <PreserveSig>
        Overloads Function SetAntialiasMode(antialiasMode As D2D1_ANTIALIAS_MODE) As HRESULT
        <PreserveSig>
        Overloads Function SetTags(tag1 As ULong, tag2 As ULong) As HRESULT
        <PreserveSig>
        Overloads Function SetTextAntialiasMode(textAntialiasMode As D2D1_TEXT_ANTIALIAS_MODE) As HRESULT
        Overloads Function SetTextRenderingParams(textRenderingParams As IDWriteRenderingParams) As HRESULT
        Overloads Function SetTransform(transform As D2D1_MATRIX_3X2_F) As HRESULT
        Overloads Function SetPrimitiveBlend(primitiveBlend As D2D1_PRIMITIVE_BLEND) As HRESULT
        Overloads Function SetUnitMode(unitMode As D2D1_UNIT_MODE) As HRESULT
        Overloads Function Clear(color As D2D1_COLOR_F) As HRESULT
        <PreserveSig>
        Overloads Function DrawGlyphRun(baselineOrigin As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, glyphRunDescription As IntPtr, foregroundBrush As ID2D1Brush, measuringMode As DWRITE_MEASURING_MODE) As HRESULT
        <PreserveSig>
        Overloads Function DrawLine(point0 As D2D1_POINT_2F, point1 As D2D1_POINT_2F, brush As ID2D1Brush, strokeWidth As Single, strokeStyle As ID2D1StrokeStyle) As HRESULT
        <PreserveSig>
        Overloads Function DrawGeometry(geometry As ID2D1Geometry, brush As ID2D1Brush, strokeWidth As Single, strokeStyle As ID2D1StrokeStyle) As HRESULT
        <PreserveSig>
        Overloads Function DrawRectangle(rect As D2D1_RECT_F, brush As ID2D1Brush, strokeWidth As Single, strokeStyle As ID2D1StrokeStyle) As HRESULT
        <PreserveSig>
        Overloads Function DrawBitmap(bitmap As ID2D1Bitmap, destinationRectangle As D2D1_RECT_F, opacity As Single, interpolationMode As D2D1_INTERPOLATION_MODE, sourceRectangle As D2D1_RECT_F, perspectiveTransform As D2D1_MATRIX_4X4_F) As HRESULT
        <PreserveSig>
        Overloads Function DrawImage(image As ID2D1Image, ByRef targetOffset As D2D1_POINT_2F, imageRectangle As D2D1_RECT_F, interpolationMode As D2D1_INTERPOLATION_MODE, compositeMode As D2D1_COMPOSITE_MODE) As HRESULT
        <PreserveSig>
        Overloads Function DrawGdiMetafile(gdiMetafile As ID2D1GdiMetafile, ByRef targetOffset As D2D1_POINT_2F) As HRESULT
        Overloads Function FillMesh(mesh As ID2D1Mesh, brush As ID2D1Brush) As HRESULT
        Overloads Function FillOpacityMask(opacityMask As ID2D1Bitmap, brush As ID2D1Brush, destinationRectangle As D2D1_RECT_F, sourceRectangle As D2D1_RECT_F) As HRESULT
        Overloads Function FillGeometry(geometry As ID2D1Geometry, brush As ID2D1Brush, opacityBrush As ID2D1Brush) As HRESULT
        Overloads Function FillRectangle(rect As D2D1_RECT_F, brush As ID2D1Brush) As HRESULT
        Overloads Function PushAxisAlignedClip(clipRect As D2D1_RECT_F, antialiasMode As D2D1_ANTIALIAS_MODE) As HRESULT
        Overloads Function PushLayer(ByRef layerParameters1 As D2D1_LAYER_PARAMETERS1, layer As ID2D1Layer) As HRESULT
        Overloads Function PopAxisAlignedClip() As HRESULT
        Overloads Function PopLayer() As HRESULT
#End Region

        <PreserveSig>
        Overloads Function SetPrimitiveBlend1(primitiveBlend As D2D1_PRIMITIVE_BLEND) As HRESULT
#End Region

        <PreserveSig>
        Function DrawInk(ink As ID2D1Ink, brush As ID2D1Brush, inkStyle As ID2D1InkStyle) As HRESULT
        <PreserveSig>
        Function DrawGradientMesh(gradientMesh As ID2D1GradientMesh) As HRESULT
        <PreserveSig>
        Overloads Function DrawGdiMetafile(gdiMetafile As ID2D1GdiMetafile, ByRef destinationRectangle As D2D1_RECT_F, ByRef sourceRectangle As D2D1_RECT_F) As HRESULT
    End Interface

    <ComImport> <Guid("18079135-4cf3-4868-bc8e-06067e6d242d")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1CommandSink3
        Inherits ID2D1CommandSink2
#Region "ID2D1CommandSink2"
#Region "ID2D1CommandSink1"
#Region "ID2D1CommandSink"
        <PreserveSig>
        Overloads Function BeginDraw() As HRESULT
        <PreserveSig>
        Overloads Function EndDraw() As HRESULT
        <PreserveSig>
        Overloads Function SetAntialiasMode(antialiasMode As D2D1_ANTIALIAS_MODE) As HRESULT
        <PreserveSig>
        Overloads Function SetTags(tag1 As ULong, tag2 As ULong) As HRESULT
        <PreserveSig>
        Overloads Function SetTextAntialiasMode(textAntialiasMode As D2D1_TEXT_ANTIALIAS_MODE) As HRESULT
        Overloads Function SetTextRenderingParams(textRenderingParams As IDWriteRenderingParams) As HRESULT
        Overloads Function SetTransform(transform As D2D1_MATRIX_3X2_F) As HRESULT
        Overloads Function SetPrimitiveBlend(primitiveBlend As D2D1_PRIMITIVE_BLEND) As HRESULT
        Overloads Function SetUnitMode(unitMode As D2D1_UNIT_MODE) As HRESULT
        Overloads Function Clear(color As D2D1_COLOR_F) As HRESULT
        <PreserveSig>
        Overloads Function DrawGlyphRun(baselineOrigin As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, glyphRunDescription As IntPtr, foregroundBrush As ID2D1Brush, measuringMode As DWRITE_MEASURING_MODE) As HRESULT
        <PreserveSig>
        Overloads Function DrawLine(point0 As D2D1_POINT_2F, point1 As D2D1_POINT_2F, brush As ID2D1Brush, strokeWidth As Single, strokeStyle As ID2D1StrokeStyle) As HRESULT
        <PreserveSig>
        Overloads Function DrawGeometry(geometry As ID2D1Geometry, brush As ID2D1Brush, strokeWidth As Single, strokeStyle As ID2D1StrokeStyle) As HRESULT
        <PreserveSig>
        Overloads Function DrawRectangle(rect As D2D1_RECT_F, brush As ID2D1Brush, strokeWidth As Single, strokeStyle As ID2D1StrokeStyle) As HRESULT
        <PreserveSig>
        Overloads Function DrawBitmap(bitmap As ID2D1Bitmap, destinationRectangle As D2D1_RECT_F, opacity As Single, interpolationMode As D2D1_INTERPOLATION_MODE, sourceRectangle As D2D1_RECT_F, perspectiveTransform As D2D1_MATRIX_4X4_F) As HRESULT
        <PreserveSig>
        Overloads Function DrawImage(image As ID2D1Image, ByRef targetOffset As D2D1_POINT_2F, imageRectangle As D2D1_RECT_F, interpolationMode As D2D1_INTERPOLATION_MODE, compositeMode As D2D1_COMPOSITE_MODE) As HRESULT
        <PreserveSig>
        Overloads Function DrawGdiMetafile(gdiMetafile As ID2D1GdiMetafile, ByRef targetOffset As D2D1_POINT_2F) As HRESULT
        Overloads Function FillMesh(mesh As ID2D1Mesh, brush As ID2D1Brush) As HRESULT
        Overloads Function FillOpacityMask(opacityMask As ID2D1Bitmap, brush As ID2D1Brush, destinationRectangle As D2D1_RECT_F, sourceRectangle As D2D1_RECT_F) As HRESULT
        Overloads Function FillGeometry(geometry As ID2D1Geometry, brush As ID2D1Brush, opacityBrush As ID2D1Brush) As HRESULT
        Overloads Function FillRectangle(rect As D2D1_RECT_F, brush As ID2D1Brush) As HRESULT
        Overloads Function PushAxisAlignedClip(clipRect As D2D1_RECT_F, antialiasMode As D2D1_ANTIALIAS_MODE) As HRESULT
        Overloads Function PushLayer(ByRef layerParameters1 As D2D1_LAYER_PARAMETERS1, layer As ID2D1Layer) As HRESULT
        Overloads Function PopAxisAlignedClip() As HRESULT
        Overloads Function PopLayer() As HRESULT
#End Region

        <PreserveSig>
        Overloads Function SetPrimitiveBlend1(primitiveBlend As D2D1_PRIMITIVE_BLEND) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function DrawInk(ink As ID2D1Ink, brush As ID2D1Brush, inkStyle As ID2D1InkStyle) As HRESULT
        <PreserveSig>
        Overloads Function DrawGradientMesh(gradientMesh As ID2D1GradientMesh) As HRESULT
        <PreserveSig>
        Overloads Function DrawGdiMetafile(gdiMetafile As ID2D1GdiMetafile, ByRef destinationRectangle As D2D1_RECT_F, ByRef sourceRectangle As D2D1_RECT_F) As HRESULT
#End Region

        <PreserveSig>
        Function DrawSpriteBatch(spriteBatch As ID2D1SpriteBatch, startIndex As UInteger, spriteCount As UInteger, bitmap As ID2D1Bitmap, interpolationMode As D2D1_BITMAP_INTERPOLATION_MODE, spriteOptions As D2D1_SPRITE_OPTIONS) As HRESULT
    End Interface

    <ComImport> <Guid("c78a6519-40d6-4218-b2de-beeeb744bb3e")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1CommandSink4
        Inherits ID2D1CommandSink3
#Region "ID2D1CommandSink3"
#Region "ID2D1CommandSink2"
#Region "ID2D1CommandSink1"
#Region "ID2D1CommandSink"
        <PreserveSig>
        Overloads Function BeginDraw() As HRESULT
        <PreserveSig>
        Overloads Function EndDraw() As HRESULT
        <PreserveSig>
        Overloads Function SetAntialiasMode(antialiasMode As D2D1_ANTIALIAS_MODE) As HRESULT
        <PreserveSig>
        Overloads Function SetTags(tag1 As ULong, tag2 As ULong) As HRESULT
        <PreserveSig>
        Overloads Function SetTextAntialiasMode(textAntialiasMode As D2D1_TEXT_ANTIALIAS_MODE) As HRESULT
        Overloads Function SetTextRenderingParams(textRenderingParams As IDWriteRenderingParams) As HRESULT
        Overloads Function SetTransform(transform As D2D1_MATRIX_3X2_F) As HRESULT
        Overloads Function SetPrimitiveBlend(primitiveBlend As D2D1_PRIMITIVE_BLEND) As HRESULT
        Overloads Function SetUnitMode(unitMode As D2D1_UNIT_MODE) As HRESULT
        Overloads Function Clear(color As D2D1_COLOR_F) As HRESULT
        <PreserveSig>
        Overloads Function DrawGlyphRun(baselineOrigin As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, glyphRunDescription As IntPtr, foregroundBrush As ID2D1Brush, measuringMode As DWRITE_MEASURING_MODE) As HRESULT
        <PreserveSig>
        Overloads Function DrawLine(point0 As D2D1_POINT_2F, point1 As D2D1_POINT_2F, brush As ID2D1Brush, strokeWidth As Single, strokeStyle As ID2D1StrokeStyle) As HRESULT
        <PreserveSig>
        Overloads Function DrawGeometry(geometry As ID2D1Geometry, brush As ID2D1Brush, strokeWidth As Single, strokeStyle As ID2D1StrokeStyle) As HRESULT
        <PreserveSig>
        Overloads Function DrawRectangle(rect As D2D1_RECT_F, brush As ID2D1Brush, strokeWidth As Single, strokeStyle As ID2D1StrokeStyle) As HRESULT
        <PreserveSig>
        Overloads Function DrawBitmap(bitmap As ID2D1Bitmap, destinationRectangle As D2D1_RECT_F, opacity As Single, interpolationMode As D2D1_INTERPOLATION_MODE, sourceRectangle As D2D1_RECT_F, perspectiveTransform As D2D1_MATRIX_4X4_F) As HRESULT
        <PreserveSig>
        Overloads Function DrawImage(image As ID2D1Image, ByRef targetOffset As D2D1_POINT_2F, imageRectangle As D2D1_RECT_F, interpolationMode As D2D1_INTERPOLATION_MODE, compositeMode As D2D1_COMPOSITE_MODE) As HRESULT
        <PreserveSig>
        Overloads Function DrawGdiMetafile(gdiMetafile As ID2D1GdiMetafile, ByRef targetOffset As D2D1_POINT_2F) As HRESULT
        Overloads Function FillMesh(mesh As ID2D1Mesh, brush As ID2D1Brush) As HRESULT
        Overloads Function FillOpacityMask(opacityMask As ID2D1Bitmap, brush As ID2D1Brush, destinationRectangle As D2D1_RECT_F, sourceRectangle As D2D1_RECT_F) As HRESULT
        Overloads Function FillGeometry(geometry As ID2D1Geometry, brush As ID2D1Brush, opacityBrush As ID2D1Brush) As HRESULT
        Overloads Function FillRectangle(rect As D2D1_RECT_F, brush As ID2D1Brush) As HRESULT
        Overloads Function PushAxisAlignedClip(clipRect As D2D1_RECT_F, antialiasMode As D2D1_ANTIALIAS_MODE) As HRESULT
        Overloads Function PushLayer(ByRef layerParameters1 As D2D1_LAYER_PARAMETERS1, layer As ID2D1Layer) As HRESULT
        Overloads Function PopAxisAlignedClip() As HRESULT
        Overloads Function PopLayer() As HRESULT
#End Region

        <PreserveSig>
        Overloads Function SetPrimitiveBlend1(primitiveBlend As D2D1_PRIMITIVE_BLEND) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function DrawInk(ink As ID2D1Ink, brush As ID2D1Brush, inkStyle As ID2D1InkStyle) As HRESULT
        <PreserveSig>
        Overloads Function DrawGradientMesh(gradientMesh As ID2D1GradientMesh) As HRESULT
        <PreserveSig>
        Overloads Function DrawGdiMetafile(gdiMetafile As ID2D1GdiMetafile, ByRef destinationRectangle As D2D1_RECT_F, ByRef sourceRectangle As D2D1_RECT_F) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function DrawSpriteBatch(spriteBatch As ID2D1SpriteBatch, startIndex As UInteger, spriteCount As UInteger, bitmap As ID2D1Bitmap, interpolationMode As D2D1_BITMAP_INTERPOLATION_MODE, spriteOptions As D2D1_SPRITE_OPTIONS) As HRESULT
#End Region

        <PreserveSig>
        Function SetPrimitiveBlend2(primitiveBlend As D2D1_PRIMITIVE_BLEND) As HRESULT
    End Interface

    <ComImport> <Guid("7047dd26-b1e7-44a7-959a-8349e2144fa8")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1CommandSink5
        Inherits ID2D1CommandSink4
#Region "ID2D1CommandSink4"
#Region "ID2D1CommandSink3"
#Region "ID2D1CommandSink2"
#Region "ID2D1CommandSink1"
#Region "ID2D1CommandSink"
        <PreserveSig>
        Overloads Function BeginDraw() As HRESULT
        <PreserveSig>
        Overloads Function EndDraw() As HRESULT
        <PreserveSig>
        Overloads Function SetAntialiasMode(antialiasMode As D2D1_ANTIALIAS_MODE) As HRESULT
        <PreserveSig>
        Overloads Function SetTags(tag1 As ULong, tag2 As ULong) As HRESULT
        <PreserveSig>
        Overloads Function SetTextAntialiasMode(textAntialiasMode As D2D1_TEXT_ANTIALIAS_MODE) As HRESULT
        Overloads Function SetTextRenderingParams(textRenderingParams As IDWriteRenderingParams) As HRESULT
        Overloads Function SetTransform(transform As D2D1_MATRIX_3X2_F) As HRESULT
        Overloads Function SetPrimitiveBlend(primitiveBlend As D2D1_PRIMITIVE_BLEND) As HRESULT
        Overloads Function SetUnitMode(unitMode As D2D1_UNIT_MODE) As HRESULT
        Overloads Function Clear(color As D2D1_COLOR_F) As HRESULT
        <PreserveSig>
        Overloads Function DrawGlyphRun(baselineOrigin As D2D1_POINT_2F, ByRef glyphRun As DWRITE_GLYPH_RUN, glyphRunDescription As IntPtr, foregroundBrush As ID2D1Brush, measuringMode As DWRITE_MEASURING_MODE) As HRESULT
        <PreserveSig>
        Overloads Function DrawLine(point0 As D2D1_POINT_2F, point1 As D2D1_POINT_2F, brush As ID2D1Brush, strokeWidth As Single, strokeStyle As ID2D1StrokeStyle) As HRESULT
        <PreserveSig>
        Overloads Function DrawGeometry(geometry As ID2D1Geometry, brush As ID2D1Brush, strokeWidth As Single, strokeStyle As ID2D1StrokeStyle) As HRESULT
        <PreserveSig>
        Overloads Function DrawRectangle(rect As D2D1_RECT_F, brush As ID2D1Brush, strokeWidth As Single, strokeStyle As ID2D1StrokeStyle) As HRESULT
        <PreserveSig>
        Overloads Function DrawBitmap(bitmap As ID2D1Bitmap, destinationRectangle As D2D1_RECT_F, opacity As Single, interpolationMode As D2D1_INTERPOLATION_MODE, sourceRectangle As D2D1_RECT_F, perspectiveTransform As D2D1_MATRIX_4X4_F) As HRESULT
        <PreserveSig>
        Overloads Function DrawImage(image As ID2D1Image, ByRef targetOffset As D2D1_POINT_2F, imageRectangle As D2D1_RECT_F, interpolationMode As D2D1_INTERPOLATION_MODE, compositeMode As D2D1_COMPOSITE_MODE) As HRESULT
        <PreserveSig>
        Overloads Function DrawGdiMetafile(gdiMetafile As ID2D1GdiMetafile, ByRef targetOffset As D2D1_POINT_2F) As HRESULT
        Overloads Function FillMesh(mesh As ID2D1Mesh, brush As ID2D1Brush) As HRESULT
        Overloads Function FillOpacityMask(opacityMask As ID2D1Bitmap, brush As ID2D1Brush, destinationRectangle As D2D1_RECT_F, sourceRectangle As D2D1_RECT_F) As HRESULT
        Overloads Function FillGeometry(geometry As ID2D1Geometry, brush As ID2D1Brush, opacityBrush As ID2D1Brush) As HRESULT
        Overloads Function FillRectangle(rect As D2D1_RECT_F, brush As ID2D1Brush) As HRESULT
        Overloads Function PushAxisAlignedClip(clipRect As D2D1_RECT_F, antialiasMode As D2D1_ANTIALIAS_MODE) As HRESULT
        Overloads Function PushLayer(ByRef layerParameters1 As D2D1_LAYER_PARAMETERS1, layer As ID2D1Layer) As HRESULT
        Overloads Function PopAxisAlignedClip() As HRESULT
        Overloads Function PopLayer() As HRESULT
#End Region

        Overloads Function SetPrimitiveBlend1(primitiveBlend As D2D1_PRIMITIVE_BLEND) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function DrawInk(ink As ID2D1Ink, brush As ID2D1Brush, inkStyle As ID2D1InkStyle) As HRESULT
        <PreserveSig>
        Overloads Function DrawGradientMesh(gradientMesh As ID2D1GradientMesh) As HRESULT
        <PreserveSig>
        Overloads Function DrawGdiMetafile(gdiMetafile As ID2D1GdiMetafile, ByRef destinationRectangle As D2D1_RECT_F, ByRef sourceRectangle As D2D1_RECT_F) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function DrawSpriteBatch(spriteBatch As ID2D1SpriteBatch, startIndex As UInteger, spriteCount As UInteger, bitmap As ID2D1Bitmap, interpolationMode As D2D1_BITMAP_INTERPOLATION_MODE, spriteOptions As D2D1_SPRITE_OPTIONS) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function SetPrimitiveBlend2(primitiveBlend As D2D1_PRIMITIVE_BLEND) As HRESULT
#End Region

        <PreserveSig>
        Function BlendImage(image As ID2D1Image, blendMode As D2D1_BLEND_MODE, ByRef targetOffset As D2D1_POINT_2F, ByRef imageRectangle As D2D1_RECT_F, interpolationMode As D2D1_INTERPOLATION_MODE) As HRESULT
    End Interface

    <ComImport> <Guid("a44472e1-8dfb-4e60-8492-6e2861c9ca8b")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1Device2
        Inherits ID2D1Device1
#Region "<ID2D1Device1>"
#Region "<ID2D1Device>"
#Region "<ID2D1Resource>"
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region

        <PreserveSig>
        Overloads Function CreateDeviceContext(options As D2D1_DEVICE_CONTEXT_OPTIONS, <Out> ByRef deviceContext As ID2D1DeviceContext) As HRESULT
        <PreserveSig>
        Overloads Function CreatePrintControl(wicFactory As IWICImagingFactory, documentTarget As IntPtr, ByRef printControlProperties As D2D1_PRINT_CONTROL_PROPERTIES, <Out> ByRef printControl As ID2D1PrintControl) As HRESULT
        <PreserveSig>
        Overloads Sub SetMaximumTextureMemory(maximumInBytes As ULong)
        <PreserveSig>
        Overloads Function GetMaximumTextureMemory() As ULong
        <PreserveSig>
        Overloads Sub ClearResources(Optional millisecondsSinceUse As UInteger = 0)
#End Region

        <PreserveSig>
        Overloads Sub GetRenderingPriority(<Out> ByRef renderingPriority As D2D1_RENDERING_PRIORITY)
        <PreserveSig>
        Overloads Sub SetRenderingPriority(renderingPriority As D2D1_RENDERING_PRIORITY)
        <PreserveSig>
        Overloads Function CreateDeviceContext(options As D2D1_DEVICE_CONTEXT_OPTIONS, <Out> ByRef deviceContext1 As ID2D1DeviceContext1) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateDeviceContext(options As D2D1_DEVICE_CONTEXT_OPTIONS, <Out> ByRef deviceContext2 As ID2D1DeviceContext2) As HRESULT
        <PreserveSig>
        Sub FlushDeviceContexts(bitmap As ID2D1Bitmap)
        <PreserveSig>
        Function GetDxgiDevice(<Out> ByRef dxgiDevice As IDXGIDevice) As HRESULT
    End Interface

    <ComImport> <Guid("852f2087-802c-4037-ab60-ff2e7ee6fc01")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1Device3
        Inherits ID2D1Device2
#Region "<ID2D1Device2>"
#Region "<ID2D1Device1>"
#Region "<ID2D1Device>"
#Region "<ID2D1Resource>"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region

        <PreserveSig>
        Overloads Function CreateDeviceContext(options As D2D1_DEVICE_CONTEXT_OPTIONS, <Out> ByRef deviceContext As ID2D1DeviceContext) As HRESULT
        <PreserveSig>
        Overloads Function CreatePrintControl(wicFactory As IWICImagingFactory, documentTarget As IntPtr, ByRef printControlProperties As D2D1_PRINT_CONTROL_PROPERTIES, <Out> ByRef printControl As ID2D1PrintControl) As HRESULT
        <PreserveSig>
        Overloads Sub SetMaximumTextureMemory(maximumInBytes As ULong)
        <PreserveSig>
        Overloads Function GetMaximumTextureMemory() As ULong
        <PreserveSig>
        Overloads Sub ClearResources(Optional millisecondsSinceUse As UInteger = 0)
#End Region

        <PreserveSig>
        Overloads Sub GetRenderingPriority(<Out> ByRef renderingPriority As D2D1_RENDERING_PRIORITY)
        <PreserveSig>
        Overloads Sub SetRenderingPriority(renderingPriority As D2D1_RENDERING_PRIORITY)
        <PreserveSig>
        Overloads Function CreateDeviceContext(options As D2D1_DEVICE_CONTEXT_OPTIONS, <Out> ByRef deviceContext1 As ID2D1DeviceContext1) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateDeviceContext(options As D2D1_DEVICE_CONTEXT_OPTIONS, <Out> ByRef deviceContext2 As ID2D1DeviceContext2) As HRESULT
        <PreserveSig>
        Overloads Sub FlushDeviceContexts(bitmap As ID2D1Bitmap)
        <PreserveSig>
        Overloads Function GetDxgiDevice(<Out> ByRef dxgiDevice As IDXGIDevice) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateDeviceContext(options As D2D1_DEVICE_CONTEXT_OPTIONS, <Out> ByRef deviceContext3 As ID2D1DeviceContext3) As HRESULT
    End Interface

    <ComImport> <Guid("d7bdb159-5683-4a46-bc9c-72dc720b858b")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1Device4
        Inherits ID2D1Device3
#Region "<ID2D1Device3>"
#Region "<ID2D1Device2>"
#Region "<ID2D1Device1>"
#Region "<ID2D1Device>"
#Region "<ID2D1Resource>"
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region

        <PreserveSig>
        Overloads Function CreateDeviceContext(options As D2D1_DEVICE_CONTEXT_OPTIONS, <Out> ByRef deviceContext As ID2D1DeviceContext) As HRESULT
        <PreserveSig>
        Overloads Function CreatePrintControl(wicFactory As IWICImagingFactory, documentTarget As IntPtr, ByRef printControlProperties As D2D1_PRINT_CONTROL_PROPERTIES, <Out> ByRef printControl As ID2D1PrintControl) As HRESULT
        <PreserveSig>
        Overloads Sub SetMaximumTextureMemory(maximumInBytes As ULong)
        <PreserveSig>
        Overloads Function GetMaximumTextureMemory() As ULong
        <PreserveSig>
        Overloads Sub ClearResources(Optional millisecondsSinceUse As UInteger = 0)
#End Region

        <PreserveSig>
        Overloads Sub GetRenderingPriority(<Out> ByRef renderingPriority As D2D1_RENDERING_PRIORITY)
        <PreserveSig>
        Overloads Sub SetRenderingPriority(renderingPriority As D2D1_RENDERING_PRIORITY)
        <PreserveSig>
        Overloads Function CreateDeviceContext(options As D2D1_DEVICE_CONTEXT_OPTIONS, <Out> ByRef deviceContext1 As ID2D1DeviceContext1) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateDeviceContext(options As D2D1_DEVICE_CONTEXT_OPTIONS, <Out> ByRef deviceContext2 As ID2D1DeviceContext2) As HRESULT
        <PreserveSig>
        Overloads Sub FlushDeviceContexts(bitmap As ID2D1Bitmap)
        <PreserveSig>
        Overloads Function GetDxgiDevice(<Out> ByRef dxgiDevice As IDXGIDevice) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateDeviceContext(options As D2D1_DEVICE_CONTEXT_OPTIONS, <Out> ByRef deviceContext3 As ID2D1DeviceContext3) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateDeviceContext(options As D2D1_DEVICE_CONTEXT_OPTIONS, <Out> ByRef deviceContext4 As ID2D1DeviceContext4) As HRESULT
        <PreserveSig>
        Sub SetMaximumColorGlyphCacheMemory(maximumInBytes As ULong)
        <PreserveSig>
        Function GetMaximumColorGlyphCacheMemory() As ULong
    End Interface

    <ComImport> <Guid("d55ba0a4-6405-4694-aef5-08ee1a4358b4")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1Device5
        Inherits ID2D1Device4
#Region "<ID2D1Device4>"
#Region "<ID2D1Device3>"
#Region "<ID2D1Device2>"
#Region "<ID2D1Device1>"
#Region "<ID2D1Device>"
#Region "<ID2D1Resource>"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region

        <PreserveSig>
        Overloads Function CreateDeviceContext(options As D2D1_DEVICE_CONTEXT_OPTIONS, <Out> ByRef deviceContext As ID2D1DeviceContext) As HRESULT
        <PreserveSig>
        Overloads Function CreatePrintControl(wicFactory As IWICImagingFactory, documentTarget As IntPtr, ByRef printControlProperties As D2D1_PRINT_CONTROL_PROPERTIES, <Out> ByRef printControl As ID2D1PrintControl) As HRESULT
        <PreserveSig>
        Overloads Sub SetMaximumTextureMemory(maximumInBytes As ULong)
        <PreserveSig>
        Overloads Function GetMaximumTextureMemory() As ULong
        <PreserveSig>
        Overloads Sub ClearResources(Optional millisecondsSinceUse As UInteger = 0)
#End Region

        <PreserveSig>
        Overloads Sub GetRenderingPriority(<Out> ByRef renderingPriority As D2D1_RENDERING_PRIORITY)
        <PreserveSig>
        Overloads Sub SetRenderingPriority(renderingPriority As D2D1_RENDERING_PRIORITY)
        <PreserveSig>
        Overloads Function CreateDeviceContext(options As D2D1_DEVICE_CONTEXT_OPTIONS, <Out> ByRef deviceContext1 As ID2D1DeviceContext1) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateDeviceContext(options As D2D1_DEVICE_CONTEXT_OPTIONS, <Out> ByRef deviceContext2 As ID2D1DeviceContext2) As HRESULT
        <PreserveSig>
        Overloads Sub FlushDeviceContexts(bitmap As ID2D1Bitmap)
        <PreserveSig>
        Overloads Function GetDxgiDevice(<Out> ByRef dxgiDevice As IDXGIDevice) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateDeviceContext(options As D2D1_DEVICE_CONTEXT_OPTIONS, <Out> ByRef deviceContext3 As ID2D1DeviceContext3) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateDeviceContext(options As D2D1_DEVICE_CONTEXT_OPTIONS, <Out> ByRef deviceContext4 As ID2D1DeviceContext4) As HRESULT
        <PreserveSig>
        Overloads Sub SetMaximumColorGlyphCacheMemory(maximumInBytes As ULong)
        <PreserveSig>
        Overloads Function GetMaximumColorGlyphCacheMemory() As ULong
#End Region

        <PreserveSig>
        Overloads Function CreateDeviceContext(options As D2D1_DEVICE_CONTEXT_OPTIONS, <Out> ByRef deviceContext5 As ID2D1DeviceContext5) As HRESULT
    End Interface

    <ComImport> <Guid("7bfef914-2d75-4bad-be87-e18ddb077b6d")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1Device6
        Inherits ID2D1Device5
#Region "<ID2D1Device5>"
#Region "<ID2D1Device4>"
#Region "<ID2D1Device3>"
#Region "<ID2D1Device2>"
#Region "<ID2D1Device1>"
#Region "<ID2D1Device>"
#Region "<ID2D1Resource>"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region

        <PreserveSig>
        Overloads Function CreateDeviceContext(options As D2D1_DEVICE_CONTEXT_OPTIONS, <Out> ByRef deviceContext As ID2D1DeviceContext) As HRESULT
        <PreserveSig>
        Overloads Function CreatePrintControl(wicFactory As IWICImagingFactory, documentTarget As IntPtr, ByRef printControlProperties As D2D1_PRINT_CONTROL_PROPERTIES, <Out> ByRef printControl As ID2D1PrintControl) As HRESULT
        <PreserveSig>
        Overloads Sub SetMaximumTextureMemory(maximumInBytes As ULong)
        <PreserveSig>
        Overloads Function GetMaximumTextureMemory() As ULong
        <PreserveSig>
        Overloads Sub ClearResources(Optional millisecondsSinceUse As UInteger = 0)
#End Region

        <PreserveSig>
        Overloads Sub GetRenderingPriority(<Out> ByRef renderingPriority As D2D1_RENDERING_PRIORITY)
        <PreserveSig>
        Overloads Sub SetRenderingPriority(renderingPriority As D2D1_RENDERING_PRIORITY)
        <PreserveSig>
        Overloads Function CreateDeviceContext(options As D2D1_DEVICE_CONTEXT_OPTIONS, <Out> ByRef deviceContext1 As ID2D1DeviceContext1) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateDeviceContext(options As D2D1_DEVICE_CONTEXT_OPTIONS, <Out> ByRef deviceContext2 As ID2D1DeviceContext2) As HRESULT
        <PreserveSig>
        Overloads Sub FlushDeviceContexts(bitmap As ID2D1Bitmap)
        <PreserveSig>
        Overloads Function GetDxgiDevice(<Out> ByRef dxgiDevice As IDXGIDevice) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateDeviceContext(options As D2D1_DEVICE_CONTEXT_OPTIONS, <Out> ByRef deviceContext3 As ID2D1DeviceContext3) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateDeviceContext(options As D2D1_DEVICE_CONTEXT_OPTIONS, <Out> ByRef deviceContext4 As ID2D1DeviceContext4) As HRESULT
        <PreserveSig>
        Overloads Sub SetMaximumColorGlyphCacheMemory(maximumInBytes As ULong)
        <PreserveSig>
        Overloads Function GetMaximumColorGlyphCacheMemory() As ULong
#End Region

        <PreserveSig>
        Overloads Function CreateDeviceContext(options As D2D1_DEVICE_CONTEXT_OPTIONS, <Out> ByRef deviceContext5 As ID2D1DeviceContext5) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateDeviceContext(options As D2D1_DEVICE_CONTEXT_OPTIONS, <Out> ByRef deviceContext6 As ID2D1DeviceContext6) As HRESULT
    End Interface

    <ComImport> <Guid("f07c8968-dd4e-4ba6-9cbd-eb6d3752dcbb")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1Device7
        Inherits ID2D1Device6
#Region "<ID2D1Device5>"
#Region "<ID2D1Device4>"
#Region "<ID2D1Device3>"
#Region "<ID2D1Device2>"
#Region "<ID2D1Device1>"
#Region "<ID2D1Device>"
#Region "<ID2D1Resource>"
        <PreserveSig>
        Overloads Sub GetFactory(<Out> ByRef factory As ID2D1Factory)
#End Region

        <PreserveSig>
        Overloads Function CreateDeviceContext(options As D2D1_DEVICE_CONTEXT_OPTIONS, <Out> ByRef deviceContext As ID2D1DeviceContext) As HRESULT
        <PreserveSig>
        Overloads Function CreatePrintControl(wicFactory As IWICImagingFactory, documentTarget As IntPtr, ByRef printControlProperties As D2D1_PRINT_CONTROL_PROPERTIES, <Out> ByRef printControl As ID2D1PrintControl) As HRESULT
        <PreserveSig>
        Overloads Sub SetMaximumTextureMemory(maximumInBytes As ULong)
        <PreserveSig>
        Overloads Function GetMaximumTextureMemory() As ULong
        <PreserveSig>
        Overloads Sub ClearResources(Optional millisecondsSinceUse As UInteger = 0)
#End Region

        <PreserveSig>
        Overloads Sub GetRenderingPriority(<Out> ByRef renderingPriority As D2D1_RENDERING_PRIORITY)
        <PreserveSig>
        Overloads Sub SetRenderingPriority(renderingPriority As D2D1_RENDERING_PRIORITY)
        <PreserveSig>
        Overloads Function CreateDeviceContext(options As D2D1_DEVICE_CONTEXT_OPTIONS, <Out> ByRef deviceContext1 As ID2D1DeviceContext1) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateDeviceContext(options As D2D1_DEVICE_CONTEXT_OPTIONS, <Out> ByRef deviceContext2 As ID2D1DeviceContext2) As HRESULT
        <PreserveSig>
        Overloads Sub FlushDeviceContexts(bitmap As ID2D1Bitmap)
        <PreserveSig>
        Overloads Function GetDxgiDevice(<Out> ByRef dxgiDevice As IDXGIDevice) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateDeviceContext(options As D2D1_DEVICE_CONTEXT_OPTIONS, <Out> ByRef deviceContext3 As ID2D1DeviceContext3) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateDeviceContext(options As D2D1_DEVICE_CONTEXT_OPTIONS, <Out> ByRef deviceContext4 As ID2D1DeviceContext4) As HRESULT
        <PreserveSig>
        Overloads Sub SetMaximumColorGlyphCacheMemory(maximumInBytes As ULong)
        <PreserveSig>
        Overloads Function GetMaximumColorGlyphCacheMemory() As ULong
#End Region

        <PreserveSig>
        Overloads Function CreateDeviceContext(options As D2D1_DEVICE_CONTEXT_OPTIONS, <Out> ByRef deviceContext5 As ID2D1DeviceContext5) As HRESULT
#End Region

#Region "<ID2D1Device6>"
        <PreserveSig>
        Overloads Function CreateDeviceContext(options As D2D1_DEVICE_CONTEXT_OPTIONS, <Out> ByRef deviceContext6 As ID2D1DeviceContext6) As HRESULT
#End Region

        <PreserveSig>
        Overloads Function CreateDeviceContext(options As D2D1_DEVICE_CONTEXT_OPTIONS, <Out> ByRef deviceContext As ID2D1DeviceContext7) As HRESULT
    End Interface

    Public Enum D2D1_YCBCR_CHROMA_SUBSAMPLING As UInteger
        D2D1_YCBCR_CHROMA_SUBSAMPLING_AUTO = 0
        D2D1_YCBCR_CHROMA_SUBSAMPLING_420 = 1
        D2D1_YCBCR_CHROMA_SUBSAMPLING_422 = 2
        D2D1_YCBCR_CHROMA_SUBSAMPLING_444 = 3
        D2D1_YCBCR_CHROMA_SUBSAMPLING_440 = 4
        D2D1_YCBCR_CHROMA_SUBSAMPLING_FORCE_DWORD = &HFFFFFFFFUI
    End Enum

    Public Enum D2D1_YCBCR_INTERPOLATION_MODE As UInteger
        D2D1_YCBCR_INTERPOLATION_MODE_NEAREST_NEIGHBOR = 0
        D2D1_YCBCR_INTERPOLATION_MODE_LINEAR = 1
        D2D1_YCBCR_INTERPOLATION_MODE_CUBIC = 2
        D2D1_YCBCR_INTERPOLATION_MODE_MULTI_SAMPLE_LINEAR = 3
        D2D1_YCBCR_INTERPOLATION_MODE_ANISOTROPIC = 4
        D2D1_YCBCR_INTERPOLATION_MODE_HIGH_QUALITY_CUBIC = 5
        D2D1_YCBCR_INTERPOLATION_MODE_FORCE_DWORD = &HFFFFFFFFUI
    End Enum

    Public Enum D2D1_YCBCR_PROP As UInteger
        D2D1_YCBCR_PROP_CHROMA_SUBSAMPLING = 0
        D2D1_YCBCR_PROP_TRANSFORM_MATRIX = 1
        D2D1_YCBCR_PROP_INTERPOLATION_MODE = 2
        D2D1_YCBCR_PROP_FORCE_DWORD = &HFFFFFFFFUI
    End Enum

    Public Enum D2D1_SVG_PAINT_TYPE As UInteger
        ''' <summary>
        ''' The fill or stroke is not rendered.
        ''' </summary>
        D2D1_SVG_PAINT_TYPE_NONE = 0

        ''' <summary>
        ''' A solid color is rendered.
        ''' </summary>
        D2D1_SVG_PAINT_TYPE_COLOR = 1

        ''' <summary>
        ''' The current color is rendered.
        ''' </summary>
        D2D1_SVG_PAINT_TYPE_CURRENT_COLOR = 2

        ''' <summary>
        ''' A paint server, defined by another element in the SVG document, is used.
        ''' </summary>
        D2D1_SVG_PAINT_TYPE_URI = 3

        ''' <summary>
        ''' A paint server, defined by another element in the SVG document, is used. If the
        ''' paint server reference is invalid, fall back to D2D1_SVG_PAINT_TYPE_NONE.
        ''' </summary>
        D2D1_SVG_PAINT_TYPE_URI_NONE = 4

        ''' <summary>
        ''' A paint server, defined by another element in the SVG document, is used. If the
        ''' paint server reference is invalid, fall back to D2D1_SVG_PAINT_TYPE_COLOR.
        ''' </summary>
        D2D1_SVG_PAINT_TYPE_URI_COLOR = 5

        ''' <summary>
        ''' A paint server, defined by another element in the SVG document, is used. If the
        ''' paint server reference is invalid, fall back to
        ''' D2D1_SVG_PAINT_TYPE_CURRENT_COLOR.
        ''' </summary>
        D2D1_SVG_PAINT_TYPE_URI_CURRENT_COLOR = 6
        D2D1_SVG_PAINT_TYPE_FORCE_DWORD = &HFFFFFFFFUI
    End Enum

    ''' <summary>
    ''' Specifies the units for an SVG length.
    ''' </summary>
    Public Enum D2D1_SVG_LENGTH_UNITS As UInteger
        ''' <summary>
        ''' The length is unitless.
        ''' </summary>
        D2D1_SVG_LENGTH_UNITS_NUMBER = 0

        ''' <summary>
        ''' The length is a percentage value.
        ''' </summary>
        D2D1_SVG_LENGTH_UNITS_PERCENTAGE = 1
        D2D1_SVG_LENGTH_UNITS_FORCE_DWORD = &HFFFFFFFFUI
    End Enum

    ''' <summary>
    ''' Specifies a value for the SVG display property.
    ''' </summary>
    Public Enum D2D1_SVG_DISPLAY As UInteger
        ''' <summary>
        ''' The element uses the default display behavior.
        ''' </summary>
        D2D1_SVG_DISPLAY_INLINE = 0

        ''' <summary>
        ''' The element and all children are not rendered directly.
        ''' </summary>
        D2D1_SVG_DISPLAY_NONE = 1
        D2D1_SVG_DISPLAY_FORCE_DWORD = &HFFFFFFFFUI
    End Enum

    ''' <summary>
    ''' Specifies a value for the SVG visibility property.
    ''' </summary>
    Public Enum D2D1_SVG_VISIBILITY As UInteger
        ''' <summary>
        ''' The element is visible.
        ''' </summary>
        D2D1_SVG_VISIBILITY_VISIBLE = 0

        ''' <summary>
        ''' The element is invisible.
        ''' </summary>
        D2D1_SVG_VISIBILITY_HIDDEN = 1
        D2D1_SVG_VISIBILITY_FORCE_DWORD = &HFFFFFFFFUI
    End Enum

    ''' <summary>
    ''' Specifies a value for the SVG overflow property.
    ''' </summary>
    Public Enum D2D1_SVG_OVERFLOW As UInteger
        ''' <summary>
        ''' The element is not clipped to its viewport.
        ''' </summary>
        D2D1_SVG_OVERFLOW_VISIBLE = 0

        ''' <summary>
        ''' The element is clipped to its viewport.
        ''' </summary>
        D2D1_SVG_OVERFLOW_HIDDEN = 1
        D2D1_SVG_OVERFLOW_FORCE_DWORD = &HFFFFFFFFUI
    End Enum

    ''' <summary>
    ''' Specifies a value for the SVG stroke-linecap property.
    ''' </summary>
    Public Enum D2D1_SVG_LINE_CAP As UInteger
        ''' <summary>
        ''' The property is set to SVG's 'butt' value.
        ''' </summary>
        D2D1_SVG_LINE_CAP_BUTT = D2D1_CAP_STYLE.D2D1_CAP_STYLE_FLAT

        ''' <summary>
        ''' The property is set to SVG's 'square' value.
        ''' </summary>
        D2D1_SVG_LINE_CAP_SQUARE = D2D1_CAP_STYLE.D2D1_CAP_STYLE_SQUARE

        ''' <summary>
        ''' The property is set to SVG's 'round' value.
        ''' </summary>
        D2D1_SVG_LINE_CAP_ROUND = D2D1_CAP_STYLE.D2D1_CAP_STYLE_ROUND
        D2D1_SVG_LINE_CAP_FORCE_DWORD = &HFFFFFFFFUI
    End Enum

    Public Enum D2D1_CAP_STYLE As UInteger
        ''' <summary>
        ''' Flat line cap.
        ''' </summary>
        D2D1_CAP_STYLE_FLAT = 0

        ''' <summary>
        ''' Square line cap.
        ''' </summary>
        D2D1_CAP_STYLE_SQUARE = 1

        ''' <summary>
        ''' Round line cap.
        ''' </summary>
        D2D1_CAP_STYLE_ROUND = 2

        ''' <summary>
        ''' Triangle line cap.
        ''' </summary>
        D2D1_CAP_STYLE_TRIANGLE = 3
        D2D1_CAP_STYLE_FORCE_DWORD = &HFFFFFFFFUI

    End Enum

    ''' <summary>
    ''' Specifies a value for the SVG stroke-linejoin property.
    ''' </summary>
    Public Enum D2D1_SVG_LINE_JOIN As UInteger
        ''' <summary>
        ''' The property is set to SVG's 'bevel' value.
        ''' </summary>
        D2D1_SVG_LINE_JOIN_BEVEL = D2D1_LINE_JOIN.D2D1_LINE_JOIN_BEVEL

        ''' <summary>
        ''' The property is set to SVG's 'miter' value. Note that this is equivalent to
        ''' D2D1_LINE_JOIN_MITER_OR_BEVEL, not D2D1_LINE_JOIN_MITER.
        ''' </summary>
        D2D1_SVG_LINE_JOIN_MITER = D2D1_LINE_JOIN.D2D1_LINE_JOIN_MITER_OR_BEVEL

        ''' <summary>
        ''' \ The property is set to SVG's 'round' value.
        ''' </summary>
        D2D1_SVG_LINE_JOIN_ROUND = D2D1_LINE_JOIN.D2D1_LINE_JOIN_ROUND
        D2D1_SVG_LINE_JOIN_FORCE_DWORD = &HFFFFFFFFUI
    End Enum

    ''' <summary>
    ''' The alignment portion of the SVG preserveAspectRatio attribute.
    ''' </summary>
    Public Enum D2D1_SVG_ASPECT_ALIGN As UInteger
        ''' <summary>
        ''' The alignment is set to SVG's 'none' value.
        ''' </summary>
        D2D1_SVG_ASPECT_ALIGN_NONE = 0

        ''' <summary>
        ''' The alignment is set to SVG's 'xMinYMin' value.
        ''' </summary>
        D2D1_SVG_ASPECT_ALIGN_X_MIN_Y_MIN = 1

        ''' <summary>
        ''' The alignment is set to SVG's 'xMidYMin' value.
        ''' </summary>
        D2D1_SVG_ASPECT_ALIGN_X_MID_Y_MIN = 2

        ''' <summary>
        ''' The alignment is set to SVG's 'xMaxYMin' value.
        ''' </summary>
        D2D1_SVG_ASPECT_ALIGN_X_MAX_Y_MIN = 3

        ''' <summary>
        ''' The alignment is set to SVG's 'xMinYMid' value.
        ''' </summary>
        D2D1_SVG_ASPECT_ALIGN_X_MIN_Y_MID = 4

        ''' <summary>
        ''' The alignment is set to SVG's 'xMidYMid' value.
        ''' </summary>
        D2D1_SVG_ASPECT_ALIGN_X_MID_Y_MID = 5

        ''' <summary>
        ''' The alignment is set to SVG's 'xMaxYMid' value.
        ''' </summary>
        D2D1_SVG_ASPECT_ALIGN_X_MAX_Y_MID = 6

        ''' <summary>
        ''' The alignment is set to SVG's 'xMinYMax' value.
        ''' </summary>
        D2D1_SVG_ASPECT_ALIGN_X_MIN_Y_MAX = 7

        ''' <summary>
        ''' The alignment is set to SVG's 'xMidYMax' value.
        ''' </summary>
        D2D1_SVG_ASPECT_ALIGN_X_MID_Y_MAX = 8

        ''' <summary>
        ''' The alignment is set to SVG's 'xMaxYMax' value.
        ''' </summary>
        D2D1_SVG_ASPECT_ALIGN_X_MAX_Y_MAX = 9
        D2D1_SVG_ASPECT_ALIGN_FORCE_DWORD = &HFFFFFFFFUI
    End Enum

    ''' <summary>
    ''' The meetOrSlice portion of the SVG preserveAspectRatio attribute.
    ''' </summary>
    Public Enum D2D1_SVG_ASPECT_SCALING As UInteger
        ''' <summary>
        ''' Scale the viewBox up as much as possible such that the entire viewBox is visible
        ''' within the viewport.
        ''' </summary>
        D2D1_SVG_ASPECT_SCALING_MEET = 0

        ''' <summary>
        ''' Scale the viewBox down as much as possible such that the entire viewport is
        ''' covered by the viewBox.
        ''' </summary>
        D2D1_SVG_ASPECT_SCALING_SLICE = 1
        D2D1_SVG_ASPECT_SCALING_FORCE_DWORD = &HFFFFFFFFUI
    End Enum

    ''' <summary>
    ''' Represents a path commmand. Each command may reference floats from the segment
    ''' data. Commands ending in _ABSOLUTE interpret data as absolute coordinate.
    ''' Commands ending in _RELATIVE interpret data as being relative to the previous
    ''' point.
    ''' </summary>
    Public Enum D2D1_SVG_PATH_COMMAND As UInteger
        ''' <summary>
        ''' Closes the current subpath. Uses no segment data.
        ''' </summary>
        D2D1_SVG_PATH_COMMAND_CLOSE_PATH = 0

        ''' <summary>
        ''' Starts a new subpath at the coordinate (x y). Uses 2 floats of segment data.
        ''' </summary>
        D2D1_SVG_PATH_COMMAND_MOVE_ABSOLUTE = 1

        ''' <summary>
        ''' Starts a new subpath at the coordinate (x y). Uses 2 floats of segment data.
        ''' </summary>
        D2D1_SVG_PATH_COMMAND_MOVE_RELATIVE = 2

        ''' <summary>
        ''' Draws a line to the coordinate (x y). Uses 2 floats of segment data.
        ''' </summary>
        D2D1_SVG_PATH_COMMAND_LINE_ABSOLUTE = 3

        ''' <summary>
        ''' Draws a line to the coordinate (x y). Uses 2 floats of segment data.
        ''' </summary>
        D2D1_SVG_PATH_COMMAND_LINE_RELATIVE = 4

        ''' <summary>
        ''' Draws a cubic Bezier curve (x1 y1 x2 y2 x y). The curve ends at (x, y) and is
        ''' defined by the two control points (x1, y1) and (x2, y2). Uses 6 floats of
        ''' segment data.
        ''' </summary>
        D2D1_SVG_PATH_COMMAND_CUBIC_ABSOLUTE = 5

        ''' <summary>
        ''' Draws a cubic Bezier curve (x1 y1 x2 y2 x y). The curve ends at (x, y) and is
        ''' defined by the two control points (x1, y1) and (x2, y2). Uses 6 floats of
        ''' segment data.
        ''' </summary>
        D2D1_SVG_PATH_COMMAND_CUBIC_RELATIVE = 6

        ''' <summary>
        ''' Draws a quadratic Bezier curve (x1 y1 x y). The curve ends at (x, y) and is
        ''' defined by the control point (x1 y1). Uses 4 floats of segment data.
        ''' </summary>
        D2D1_SVG_PATH_COMMAND_QUADRADIC_ABSOLUTE = 7

        ''' <summary>
        ''' Draws a quadratic Bezier curve (x1 y1 x y). The curve ends at (x, y) and is
        ''' defined by the control point (x1 y1). Uses 4 floats of segment data.
        ''' </summary>
        D2D1_SVG_PATH_COMMAND_QUADRADIC_RELATIVE = 8

        ''' <summary>
        ''' Draws an elliptical arc (rx ry x-axis-rotation large-arc-flag sweep-flag x y).
        ''' The curve ends at (x, y) and is defined by the arc parameters. The two flags are
        ''' considered set if their values are non-zero. Uses 7 floats of segment data.
        ''' </summary>
        D2D1_SVG_PATH_COMMAND_ARC_ABSOLUTE = 9

        ''' <summary>
        ''' Draws an elliptical arc (rx ry x-axis-rotation large-arc-flag sweep-flag x y).
        ''' The curve ends at (x, y) and is defined by the arc parameters. The two flags are
        ''' considered set if their values are non-zero. Uses 7 floats of segment data.
        ''' </summary>
        D2D1_SVG_PATH_COMMAND_ARC_RELATIVE = 10

        ''' <summary>
        ''' Draws a horizontal line to the coordinate (x). Uses 1 float of segment data.
        ''' </summary>
        D2D1_SVG_PATH_COMMAND_HORIZONTAL_ABSOLUTE = 11

        ''' <summary>
        ''' Draws a horizontal line to the coordinate (x). Uses 1 float of segment data.
        ''' </summary>
        D2D1_SVG_PATH_COMMAND_HORIZONTAL_RELATIVE = 12

        ''' <summary>
        ''' Draws a vertical line to the coordinate (y). Uses 1 float of segment data.
        ''' </summary>
        D2D1_SVG_PATH_COMMAND_VERTICAL_ABSOLUTE = 13

        ''' <summary>
        ''' Draws a vertical line to the coordinate (y). Uses 1 float of segment data.
        ''' </summary>
        D2D1_SVG_PATH_COMMAND_VERTICAL_RELATIVE = 14

        ''' <summary>
        ''' Draws a smooth cubic Bezier curve (x2 y2 x y). The curve ends at (x, y) and is
        ''' defined by the control point (x2, y2). Uses 4 floats of segment data.
        ''' </summary>
        D2D1_SVG_PATH_COMMAND_CUBIC_SMOOTH_ABSOLUTE = 15

        ''' <summary>
        ''' Draws a smooth cubic Bezier curve (x2 y2 x y). The curve ends at (x, y) and is
        ''' defined by the control point (x2, y2). Uses 4 floats of segment data.
        ''' </summary>
        D2D1_SVG_PATH_COMMAND_CUBIC_SMOOTH_RELATIVE = 16

        ''' <summary>
        ''' Draws a smooth quadratic Bezier curve ending at (x, y). Uses 2 floats of segment
        ''' data.
        ''' </summary>
        D2D1_SVG_PATH_COMMAND_QUADRADIC_SMOOTH_ABSOLUTE = 17

        ''' <summary>
        ''' Draws a smooth quadratic Bezier curve ending at (x, y). Uses 2 floats of segment
        ''' data.
        ''' </summary>
        D2D1_SVG_PATH_COMMAND_QUADRADIC_SMOOTH_RELATIVE = 18
        D2D1_SVG_PATH_COMMAND_FORCE_DWORD = &HFFFFFFFFUI
    End Enum

    ''' <summary>
    ''' Defines the coordinate system used for SVG gradient or clipPath elements.
    ''' </summary>
    Public Enum D2D1_SVG_UNIT_TYPE As UInteger
        ''' <summary>
        ''' The property is set to SVG's 'userSpaceOnUse' value.
        ''' </summary>
        D2D1_SVG_UNIT_TYPE_USER_SPACE_ON_USE = 0

        ''' <summary>
        ''' The property is set to SVG's 'objectBoundingBox' value.
        ''' </summary>
        D2D1_SVG_UNIT_TYPE_OBJECT_BOUNDING_BOX = 1
        D2D1_SVG_UNIT_TYPE_FORCE_DWORD = &HFFFFFFFFUI
    End Enum

    ''' <summary>
    ''' Defines the type of SVG string attribute to set or get.
    ''' </summary>
    Public Enum D2D1_SVG_ATTRIBUTE_STRING_TYPE As UInteger
        ''' <summary>
        ''' The attribute is a string in the same form as it would appear in the SVG XML.
        ''' 
        ''' Note that when getting values of this type, the value returned may not exactly
        ''' match the value that was set. Instead, the output value is a normalized version
        ''' of the value. For example, an input color of 'red' may be output as '#FF0000'.
        ''' </summary>
        D2D1_SVG_ATTRIBUTE_STRING_TYPE_SVG = 0

        ''' <summary>
        ''' The attribute is an element ID.
        ''' </summary>
        D2D1_SVG_ATTRIBUTE_STRING_TYPE_ID = 1
        D2D1_SVG_ATTRIBUTE_STRING_TYPE_FORCE_DWORD = &HFFFFFFFFUI
    End Enum

    ''' <summary>
    ''' Defines the type of SVG POD attribute to set or get.
    ''' </summary>
    Public Enum D2D1_SVG_ATTRIBUTE_POD_TYPE As UInteger
        ''' <summary>
        ''' The attribute is a FLOAT.
        ''' </summary>
        D2D1_SVG_ATTRIBUTE_POD_TYPE_FLOAT = 0

        ''' <summary>
        ''' The attribute is a D2D1_COLOR_F.
        ''' </summary>
        D2D1_SVG_ATTRIBUTE_POD_TYPE_COLOR = 1

        ''' <summary>
        ''' The attribute is a D2D1_FILL_MODE.
        ''' </summary>
        D2D1_SVG_ATTRIBUTE_POD_TYPE_FILL_MODE = 2

        ''' <summary>
        ''' The attribute is a D2D1_SVG_DISPLAY.
        ''' </summary>
        D2D1_SVG_ATTRIBUTE_POD_TYPE_DISPLAY = 3

        ''' <summary>
        ''' The attribute is a D2D1_SVG_OVERFLOW.
        ''' </summary>
        D2D1_SVG_ATTRIBUTE_POD_TYPE_OVERFLOW = 4

        ''' <summary>
        ''' The attribute is a D2D1_SVG_LINE_CAP.
        ''' </summary>
        D2D1_SVG_ATTRIBUTE_POD_TYPE_LINE_CAP = 5

        ''' <summary>
        ''' The attribute is a D2D1_SVG_LINE_JOIN.
        ''' </summary>
        D2D1_SVG_ATTRIBUTE_POD_TYPE_LINE_JOIN = 6

        ''' <summary>
        ''' The attribute is a D2D1_SVG_VISIBILITY.
        ''' </summary>
        D2D1_SVG_ATTRIBUTE_POD_TYPE_VISIBILITY = 7

        ''' <summary>
        ''' The attribute is a D2D1_MATRIX_3X2_F.
        ''' </summary>
        D2D1_SVG_ATTRIBUTE_POD_TYPE_MATRIX = 8

        ''' <summary>
        ''' The attribute is a D2D1_SVG_UNIT_TYPE.
        ''' </summary>
        D2D1_SVG_ATTRIBUTE_POD_TYPE_UNIT_TYPE = 9

        ''' <summary>
        ''' The attribute is a D2D1_EXTEND_MODE.
        ''' </summary>
        D2D1_SVG_ATTRIBUTE_POD_TYPE_EXTEND_MODE = 10

        ''' <summary>
        ''' The attribute is a D2D1_SVG_PRESERVE_ASPECT_RATIO.
        ''' </summary>
        D2D1_SVG_ATTRIBUTE_POD_TYPE_PRESERVE_ASPECT_RATIO = 11

        ''' <summary>
        ''' The attribute is a D2D1_SVG_VIEWBOX.
        ''' </summary>
        D2D1_SVG_ATTRIBUTE_POD_TYPE_VIEWBOX = 12

        ''' <summary>
        ''' The attribute is a D2D1_SVG_LENGTH.
        ''' </summary>
        D2D1_SVG_ATTRIBUTE_POD_TYPE_LENGTH = 13
        D2D1_SVG_ATTRIBUTE_POD_TYPE_FORCE_DWORD = &HFFFFFFFFUI
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_SVG_LENGTH
        Public value As Single
        Public units As D2D1_SVG_LENGTH_UNITS
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_SVG_PRESERVE_ASPECT_RATIO
        ''' <summary>
        ''' Sets the 'defer' portion of the preserveAspectRatio settings. This field only
        ''' has an effect on an 'image' element that references another SVG document. As
        ''' this is not currently supported, the field has no impact on rendering.
        ''' </summary>
        Public defer As Boolean

        ''' <summary>
        ''' Sets the align portion of the preserveAspectRatio settings.
        ''' </summary>
        Public align As D2D1_SVG_ASPECT_ALIGN

        ''' <summary>
        ''' Sets the meetOrSlice portion of the preserveAspectRatio settings.
        ''' </summary>
        Public meetOrSlice As D2D1_SVG_ASPECT_SCALING
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_SVG_VIEWBOX
        Public x As Single
        Public y As Single
        Public width As Single
        Public height As Single
    End Structure

    ''' <summary>
    ''' Indicates what has changed since the last time the effect was asked to prepare
    ''' to render.
    ''' </summary>
    Public Enum D2D1_CHANGE_TYPE As UInteger
        ''' <summary>
        ''' Nothing has changed.
        ''' </summary>
        D2D1_CHANGE_TYPE_NONE = 0

        ''' <summary>
        ''' The effect's properties have changed.
        ''' </summary>
        D2D1_CHANGE_TYPE_PROPERTIES = 1

        ''' <summary>
        ''' The internal context has changed and should be inspected.
        ''' </summary>
        D2D1_CHANGE_TYPE_CONTEXT = 2

        ''' <summary>
        ''' A new graph has been set due to a change in the input count.
        ''' </summary>
        D2D1_CHANGE_TYPE_GRAPH = 3
        D2D1_CHANGE_TYPE_FORCE_DWORD = &HFFFFFFFFUI
    End Enum

    ''' <summary>
    ''' Indicates options for drawing using a pixel shader.
    ''' </summary>
    Public Enum D2D1_PIXEL_OPTIONS As UInteger
        ''' <summary>
        ''' Default pixel processing.
        ''' </summary>
        D2D1_PIXEL_OPTIONS_NONE = 0

        ''' <summary>
        ''' Indicates that the shader samples its inputs only at exactly the same scene
        ''' coordinate as the output pixel, and that it returns transparent black whenever
        ''' the input pixels are also transparent black.
        ''' </summary>
        D2D1_PIXEL_OPTIONS_TRIVIAL_SAMPLING = 1
        D2D1_PIXEL_OPTIONS_FORCE_DWORD = &HFFFFFFFFUI
    End Enum

    ''' <summary>
    ''' Indicates options for drawing custom vertices set by transforms.
    ''' </summary>
    Public Enum D2D1_VERTEX_OPTIONS As UInteger
        ''' <summary>
        ''' Default vertex processing.
        ''' </summary>
        D2D1_VERTEX_OPTIONS_NONE = 0

        ''' <summary>
        ''' Indicates that the output rectangle does not need to be cleared before drawing
        ''' custom vertices. This must only be used by transforms whose custom vertices
        ''' completely cover their output rectangle.
        ''' </summary>
        D2D1_VERTEX_OPTIONS_DO_NOT_CLEAR = 1

        ''' <summary>
        ''' Causes a depth buffer to be used while drawing custom vertices. This impacts
        ''' drawing behavior when primitives overlap one another.
        ''' </summary>
        D2D1_VERTEX_OPTIONS_USE_DEPTH_BUFFER = 2

        ''' <summary>
        ''' Indicates that custom vertices do not form primitives which overlap one another.
        ''' </summary>
        D2D1_VERTEX_OPTIONS_ASSUME_NO_OVERLAP = 4
        D2D1_VERTEX_OPTIONS_FORCE_DWORD = &HFFFFFFFFUI
    End Enum

    ''' <summary>
    ''' Describes how a vertex buffer is to be managed.
    ''' </summary>
    Public Enum D2D1_VERTEX_USAGE As UInteger

        ''' <summary>
        ''' The vertex buffer content do not change frequently from frame to frame.
        ''' </summary>
        D2D1_VERTEX_USAGE_STATIC = 0

        ''' <summary>
        ''' The vertex buffer is intended to be updated frequently.
        ''' </summary>
        D2D1_VERTEX_USAGE_DYNAMIC = 1
        D2D1_VERTEX_USAGE_FORCE_DWORD = &HFFFFFFFFUI
    End Enum

    ''' <summary>
    ''' Describes a particular blend in the D2D1_BLEND_DESCRIPTION structure.
    ''' </summary>
    Public Enum D2D1_BLEND_OPERATION As UInteger
        D2D1_BLEND_OPERATION_ADD = 1
        D2D1_BLEND_OPERATION_SUBTRACT = 2
        D2D1_BLEND_OPERATION_REV_SUBTRACT = 3
        D2D1_BLEND_OPERATION_MIN = 4
        D2D1_BLEND_OPERATION_MAX = 5
        D2D1_BLEND_OPERATION_FORCE_DWORD = &HFFFFFFFFUI
    End Enum

    ''' <summary>
    ''' Describes a particular blend in the D2D1_BLEND_DESCRIPTION structure.
    ''' </summary>
    Public Enum D2D1_BLEND As UInteger
        D2D1_BLEND_ZERO = 1
        D2D1_BLEND_ONE = 2
        D2D1_BLEND_SRC_COLOR = 3
        D2D1_BLEND_INV_SRC_COLOR = 4
        D2D1_BLEND_SRC_ALPHA = 5
        D2D1_BLEND_INV_SRC_ALPHA = 6
        D2D1_BLEND_DEST_ALPHA = 7
        D2D1_BLEND_INV_DEST_ALPHA = 8
        D2D1_BLEND_DEST_COLOR = 9
        D2D1_BLEND_INV_DEST_COLOR = 10
        D2D1_BLEND_SRC_ALPHA_SAT = 11
        D2D1_BLEND_BLEND_FACTOR = 14
        D2D1_BLEND_INV_BLEND_FACTOR = 15
        D2D1_BLEND_FORCE_DWORD = &HFFFFFFFFUI
    End Enum

    ''' <summary>
    ''' Allows a caller to control the channel depth of a stage in the rendering
    ''' pipeline.
    ''' </summary>
    Public Enum D2D1_CHANNEL_DEPTH As UInteger
        D2D1_CHANNEL_DEPTH_DEFAULT = 0
        D2D1_CHANNEL_DEPTH_1 = 1
        D2D1_CHANNEL_DEPTH_4 = 4
        D2D1_CHANNEL_DEPTH_FORCE_DWORD = &HFFFFFFFFUI
    End Enum

    ''' <summary>
    ''' Represents filtering modes transforms may select to use on their input textures.
    ''' </summary>
    Public Enum D2D1_FILTER As UInteger
        D2D1_FILTER_MIN_MAG_MIP_POINT = &H0
        D2D1_FILTER_MIN_MAG_POINT_MIP_LINEAR = &H1
        D2D1_FILTER_MIN_POINT_MAG_LINEAR_MIP_POINT = &H4
        D2D1_FILTER_MIN_POINT_MAG_MIP_LINEAR = &H5
        D2D1_FILTER_MIN_LINEAR_MAG_MIP_POINT = &H10
        D2D1_FILTER_MIN_LINEAR_MAG_POINT_MIP_LINEAR = &H11
        D2D1_FILTER_MIN_MAG_LINEAR_MIP_POINT = &H14
        D2D1_FILTER_MIN_MAG_MIP_LINEAR = &H15
        D2D1_FILTER_ANISOTROPIC = &H55
        D2D1_FILTER_FORCE_DWORD = &HFFFFFFFFUI
    End Enum

    ''' <summary>
    ''' Defines capabilities of the underlying D3D device which may be queried using
    ''' CheckFeatureSupport.
    ''' </summary>
    Public Enum D2D1_FEATURE As UInteger
        D2D1_FEATURE_DOUBLES = 0
        D2D1_FEATURE_D3D10_X_HARDWARE_OPTIONS = 1
        D2D1_FEATURE_FORCE_DWORD = &HFFFFFFFFUI
    End Enum

    ''' <summary>
    ''' Defines a property binding to a function. The name must match the property
    ''' defined in the registration schema.
    ''' </summary>
    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_PROPERTY_BINDING
        ''' <summary>
        ''' The name of the property.
        ''' </summary>
        Public propertyName As String

        ''' <summary>
        ''' The function that will receive the data to set.
        ''' </summary>
        Public setFunction As PD2D1_PROPERTY_SET_FUNCTION

        ''' <summary>
        ''' The function that will be asked to write the output data.
        ''' </summary>
        Public getFunction As PD2D1_PROPERTY_GET_FUNCTION
    End Structure

    Public Delegate Function PD2D1_PROPERTY_GET_FUNCTION(effect As IntPtr, <Out> ByRef data As IntPtr, dataSize As UInteger, <Out> ByRef actualSize As UInteger) As HRESULT

    Public Delegate Function PD2D1_PROPERTY_SET_FUNCTION(effect As IntPtr, data As IntPtr, dataSize As UInteger) As HRESULT

    ''' <summary>
    ''' This is used to define a resource texture when that resource texture is created.
    ''' </summary>
    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_RESOURCE_TEXTURE_PROPERTIES
        '_Field_size_(dimensions) CONST uint* extents;
        Public extents As IntPtr
        Public dimensions As UInteger
        Public bufferPrecision As D2D1_BUFFER_PRECISION
        Public channelDepth As D2D1_CHANNEL_DEPTH
        Public filter As D2D1_FILTER
        '_Field_size_(dimensions) CONST D2D1_EXTEND_MODE *extendModes;
        Public extendModes As IntPtr
    End Structure

    ''' <summary>
    ''' This defines a single element of the vertex layout.
    ''' </summary>
    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_INPUT_ELEMENT_DESC
        Public semanticName As String
        Public semanticIndex As UInteger
        Public format As DXGI_FORMAT
        Public inputSlot As UInteger
        Public alignedByteOffset As UInteger
    End Structure

    ''' <summary>
    ''' This defines the properties of a vertex buffer which uses the default vertex
    ''' layout.
    ''' </summary>
    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_VERTEX_BUFFER_PROPERTIES
        Public inputCount As UInteger
        Public usage As D2D1_VERTEX_USAGE
        Public data As IntPtr
        Public byteWidth As UInteger
    End Structure

    ''' <summary>
    ''' This defines the input layout of vertices and the vertex shader which processes
    ''' them.
    ''' </summary>
    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_CUSTOM_VERTEX_BUFFER_PROPERTIES
        Public shaderBufferWithInputSignature As IntPtr
        Public shaderBufferSize As UInteger
        'public D2D1_INPUT_ELEMENT_DESC *inputElements;
        Public inputElements As IntPtr
        Public elementCount As UInteger
        Public stride As UInteger
    End Structure

    ''' <summary>
    ''' This defines the range of vertices from a vertex buffer to draw.
    ''' </summary>
    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_VERTEX_RANGE
        Public startVertex As UInteger
        Public vertexCount As UInteger
    End Structure

    ''' <summary>
    ''' Blend description which configures a blend transform object.
    ''' </summary>
    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_BLEND_DESCRIPTION
        Public sourceBlend As D2D1_BLEND
        Public destinationBlend As D2D1_BLEND
        Public blendOperation As D2D1_BLEND_OPERATION
        Public sourceBlendAlpha As D2D1_BLEND
        Public destinationBlendAlpha As D2D1_BLEND
        Public blendOperationAlpha As D2D1_BLEND_OPERATION
        <MarshalAs(UnmanagedType.R4, SizeConst:=4)>
        Public blendFactor As Single
    End Structure

    ''' <summary>
    ''' Describes options transforms may select to use on their input textures.
    ''' </summary>
    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_INPUT_DESCRIPTION
        Public filter As D2D1_FILTER
        Public levelOfDetailCount As UInteger
    End Structure

    ''' <summary>
    ''' Indicates whether shader support for doubles is present on the underlying
    ''' hardware.  This may be populated using CheckFeatureSupport.
    ''' </summary>
    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_FEATURE_DATA_DOUBLES
        Public doublePrecisionFloatShaderOps As Boolean
    End Structure

    ''' <summary>
    ''' Indicates support for features which are optional on D3D10 feature levels.  This
    ''' may be populated using CheckFeatureSupport.
    ''' </summary>
    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_FEATURE_DATA_D3D10_X_HARDWARE_OPTIONS
        Public computeShaders_Plus_RawAndStructuredBuffers_Via_Shader_4_x As Boolean
    End Structure

    <ComImport> <Guid("0359dc30-95e6-4568-9055-27720d130e93")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1AnalysisTransform
        Function ProcessAnalysisResults(analysisData As IntPtr, analysisDataCount As UInteger) As HRESULT
    End Interface

    <ComImport> <Guid("63ac0b32-ba44-450f-8806-7f4ca1ff2f1b")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1BlendTransform
        Inherits ID2D1ConcreteTransform
#Region "<ID2D1ConcreteTransform>"
#Region "<ID2D1TransformNode>"
        <PreserveSig>
        Overloads Function GetInputCount() As UInteger
#End Region

        <PreserveSig>
        Overloads Function SetOutputBuffer(bufferPrecision As D2D1_BUFFER_PRECISION, channelDepth As D2D1_CHANNEL_DEPTH) As HRESULT
        <PreserveSig>
        Overloads Sub SetCached(isCached As Boolean)
#End Region

        <PreserveSig>
        Sub SetDescription(description As D2D1_BLEND_DESCRIPTION)
        <PreserveSig>
        Sub GetDescription(<Out> ByRef description As D2D1_BLEND_DESCRIPTION)
    End Interface

    <ComImport> <Guid("b2efe1e7-729f-4102-949f-505fa21bf666")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1TransformNode
        <PreserveSig>
        Function GetInputCount() As UInteger
    End Interface

    <ComImport> <Guid("1a799d8a-69f7-4e4c-9fed-437ccc6684cc")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1ConcreteTransform
        Inherits ID2D1TransformNode
#Region "<ID2D1TransformNode>"
        <PreserveSig>
        Overloads Function GetInputCount() As UInteger
#End Region

        <PreserveSig>
        Function SetOutputBuffer(bufferPrecision As D2D1_BUFFER_PRECISION, channelDepth As D2D1_CHANNEL_DEPTH) As HRESULT
        <PreserveSig>
        Sub SetCached(isCached As Boolean)
    End Interface

    <ComImport> <Guid("4998735c-3a19-473c-9781-656847e3a347")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1BorderTransform
        Inherits ID2D1ConcreteTransform
#Region "<ID2D1ConcreteTransform>"
#Region "<ID2D1TransformNode>"
        <PreserveSig>
        Overloads Function GetInputCount() As UInteger
#End Region

        <PreserveSig>
        Overloads Function SetOutputBuffer(bufferPrecision As D2D1_BUFFER_PRECISION, channelDepth As D2D1_CHANNEL_DEPTH) As HRESULT
        <PreserveSig>
        Overloads Sub SetCached(isCached As Boolean)
#End Region

        <PreserveSig>
        Sub SetExtendModeX(extendMode As D2D1_EXTEND_MODE)
        <PreserveSig>
        Sub SetExtendModeY(extendMode As D2D1_EXTEND_MODE)
        <PreserveSig>
        Sub GetExtendModeX(<Out> ByRef extendedModeX As D2D1_EXTEND_MODE)
        <PreserveSig>
        Sub GetExtendModeY(<Out> ByRef extendedModeY As D2D1_EXTEND_MODE)
    End Interface

    <ComImport> <Guid("90f732e2-5092-4606-a819-8651970baccd")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1BoundsAdjustmentTransform
        Inherits ID2D1TransformNode
#Region "<ID2D1TransformNode>"
        <PreserveSig>
        Overloads Function GetInputCount() As UInteger
#End Region

        <PreserveSig>
        Sub SetOutputBounds(ByRef outputBounds As D2D1_RECT_L)
        <PreserveSig>
        Sub GetOutputBounds(<Out> ByRef outputBounds As D2D1_RECT_L)
    End Interface

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_RECT_L
        Public left As Long
        Public top As Long
        Public right As Long
        Public bottom As Long
        Public Sub New(left As Long, top As Long, right As Long, bottom As Long)
            Me.left = left
            Me.top = top
            Me.right = right
            Me.bottom = bottom
        End Sub
    End Structure

    <ComImport> <Guid("519ae1bd-d19a-420d-b849-364f594776b7")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1RenderInfo
        <PreserveSig>
        Function SetInputDescription(inputIndex As UInteger, inputDescription As D2D1_INPUT_DESCRIPTION) As HRESULT
        <PreserveSig>
        Function SetOutputBuffer(bufferPrecision As D2D1_BUFFER_PRECISION, channelDepth As D2D1_CHANNEL_DEPTH) As HRESULT
        <PreserveSig>
        Sub SetCached(isCached As Boolean)
        <PreserveSig>
        Sub SetInstructionCountHint(instructionCount As UInteger)
    End Interface

    <ComImport> <Guid("5598b14b-9fd7-48b7-9bdb-8f0964eb38bc")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1ComputeInfo
        Inherits ID2D1RenderInfo
#Region "<ID2D1RenderInfo>"
        <PreserveSig>
        Overloads Function SetInputDescription(inputIndex As UInteger, inputDescription As D2D1_INPUT_DESCRIPTION) As HRESULT
        <PreserveSig>
        Overloads Function SetOutputBuffer(bufferPrecision As D2D1_BUFFER_PRECISION, channelDepth As D2D1_CHANNEL_DEPTH) As HRESULT
        <PreserveSig>
        Overloads Sub SetCached(isCached As Boolean)
        <PreserveSig>
        Overloads Sub SetInstructionCountHint(instructionCount As UInteger)
#End Region

        <PreserveSig>
        Function SetComputeShaderConstantBuffer(buffer As IntPtr, bufferCount As UInteger) As HRESULT
        <PreserveSig>
        Function SetComputeShader(ByRef shaderId As Guid) As HRESULT
        <PreserveSig>
        Function SetResourceTexture(textureIndex As UInteger, resourceTexture As ID2D1ResourceTexture) As HRESULT
    End Interface

    <ComImport> <Guid("688d15c3-02b0-438d-b13a-d1b44c32c39a")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1ResourceTexture
        <PreserveSig>
        Function Update(minimumExtents As UInteger, maximimumExtents As UInteger, strides As UInteger, dimensions As UInteger, data As IntPtr, dataCount As UInteger) As HRESULT
    End Interface

    <ComImport> <Guid("db1800dd-0c34-4cf9-be90-31cc0a5653e1")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1SourceTransform
        Inherits ID2D1Transform
#Region "<ID2D1Transform>"
#Region "<ID2D1TransformNode>"
        <PreserveSig>
        Overloads Function GetInputCount() As UInteger
#End Region

        <PreserveSig>
        Overloads Function MapOutputRectToInputRects(outputRect As D2D1_RECT_L, <Out> ByRef inputRects As D2D1_RECT_L, inputRectsCount As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function MapInputRectsToOutputRect(inputRects As IntPtr, inputOpaqueSubRects As IntPtr, inputRectCount As UInteger, <Out> ByRef outputRect As D2D1_RECT_L, <Out> ByRef outputOpaqueSubRect As D2D1_RECT_L) As HRESULT
        <PreserveSig>
        Overloads Function MapInvalidRect(inputIndex As UInteger, invalidInputRect As D2D1_RECT_L, <Out> ByRef invalidOutputRect As D2D1_RECT_L) As HRESULT
#End Region

        <PreserveSig>
        Function SetRenderInfo(renderInfo As ID2D1RenderInfo) As HRESULT
        <PreserveSig>
        Function Draw(target As ID2D1Bitmap1, ByRef drawRect As D2D1_RECT_L, targetOrigin As D2D1_POINT_2U) As HRESULT
    End Interface

    <ComImport> <Guid("ef1a287d-342a-4f76-8fdb-da0d6ea9f92b")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1Transform
        Inherits ID2D1TransformNode
#Region "<ID2D1TransformNode>"
        <PreserveSig>
        Overloads Function GetInputCount() As UInteger
#End Region

        <PreserveSig>
        Function MapOutputRectToInputRects(outputRect As D2D1_RECT_L, <Out> ByRef inputRects As D2D1_RECT_L, inputRectsCount As UInteger) As HRESULT
        <PreserveSig>
        Function MapInputRectsToOutputRect(inputRects As IntPtr, inputOpaqueSubRects As IntPtr, inputRectCount As UInteger, <Out> ByRef outputRect As D2D1_RECT_L, <Out> ByRef outputOpaqueSubRect As D2D1_RECT_L) As HRESULT
        <PreserveSig>
        Function MapInvalidRect(inputIndex As UInteger, invalidInputRect As D2D1_RECT_L, <Out> ByRef invalidOutputRect As D2D1_RECT_L) As HRESULT
    End Interface

    <ComImport> <Guid("13d29038-c3e6-4034-9081-13b53a417992")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1TransformGraph
        <PreserveSig>
        Function GetInputCount() As UInteger
        <PreserveSig>
        Function SetSingleTransformNode(node As ID2D1TransformNode) As HRESULT
        <PreserveSig>
        Function AddNode(node As ID2D1TransformNode) As HRESULT
        <PreserveSig>
        Function RemoveNode(node As ID2D1TransformNode) As HRESULT
        <PreserveSig>
        Function SetOutputNode(node As ID2D1TransformNode) As HRESULT
        <PreserveSig>
        Function ConnectNode(fromNode As ID2D1TransformNode, toNode As ID2D1TransformNode, toNodeInputIndex As UInteger) As HRESULT
        <PreserveSig>
        Function ConnectToEffectInput(toEffectInputIndex As UInteger, node As ID2D1TransformNode, toNodeInputIndex As UInteger) As HRESULT
        <PreserveSig>
        Sub Clear()
        <PreserveSig>
        Function SetPassthroughGraph(effectInputIndex As UInteger) As HRESULT
    End Interface

    <ComImport> <Guid("693ce632-7f2f-45de-93fe-18d88b37aa21")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1DrawInfo
        Inherits ID2D1RenderInfo
#Region "<ID2D1RenderInfo>"
        <PreserveSig>
        Overloads Function SetInputDescription(inputIndex As UInteger, inputDescription As D2D1_INPUT_DESCRIPTION) As HRESULT
        <PreserveSig>
        Overloads Function SetOutputBuffer(bufferPrecision As D2D1_BUFFER_PRECISION, channelDepth As D2D1_CHANNEL_DEPTH) As HRESULT
        <PreserveSig>
        Overloads Sub SetCached(isCached As Boolean)
        <PreserveSig>
        Overloads Sub SetInstructionCountHint(instructionCount As UInteger)
#End Region

        <PreserveSig>
        Function SetPixelShaderantBuffer(buffer As IntPtr, bufferCount As UInteger) As HRESULT
        <PreserveSig>
        Function SetResourceTexture(textureIndex As UInteger, resourceTexture As ID2D1ResourceTexture) As HRESULT
        <PreserveSig>
        Function SetVertexShaderantBuffer(buffer As IntPtr, bufferCount As UInteger) As HRESULT
        <PreserveSig>
        Function SetPixelShader(ByRef shaderId As Guid, Optional pixelOptions As D2D1_PIXEL_OPTIONS = D2D1_PIXEL_OPTIONS.D2D1_PIXEL_OPTIONS_NONE) As HRESULT
        <PreserveSig>
        Function SetVertexProcessing(vertexBuffer As ID2D1VertexBuffer, vertexOptions As D2D1_VERTEX_OPTIONS, blendDescription As D2D1_BLEND_DESCRIPTION, vertexRange As D2D1_VERTEX_RANGE, ByRef vertexShader As Guid) As HRESULT
    End Interface

    <ComImport> <Guid("9b8b1336-00a5-4668-92b7-ced5d8bf9b7b")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1VertexBuffer
        <PreserveSig>
        Function Map(<Out> ByRef data As IntPtr, bufferSize As UInteger) As HRESULT
        <PreserveSig>
        Function Unmap() As HRESULT
    End Interface

    <ComImport> <Guid("3d9f916b-27dc-4ad7-b4f1-64945340f563")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1EffectContext
        <PreserveSig>
        Sub GetDpi(<Out> ByRef dpiX As Single, <Out> ByRef dpiY As Single)
        <PreserveSig>
        Function CreateEffect(ByRef effectId As Guid, <Out> ByRef effect As ID2D1Effect) As HRESULT
        <PreserveSig>
        Function GetMaximumSupportedFeatureLevel(
         <MarshalAs(UnmanagedType.LPArray)> featureLevels As Integer(), featureLevelsCount As UInteger, <Out> ByRef maximumSupportedFeatureLevel As D3D_FEATURE_LEVEL) As HRESULT
        <PreserveSig>
        Function CreateTransformNodeFromEffect(effect As ID2D1Effect, <Out> ByRef transformNode As ID2D1TransformNode) As HRESULT
        <PreserveSig>
        Function CreateBlendTransform(numInputs As UInteger, blendDescription As D2D1_BLEND_DESCRIPTION, <Out> ByRef transform As ID2D1BlendTransform) As HRESULT
        <PreserveSig>
        Function CreateBorderTransform(extendModeX As D2D1_EXTEND_MODE, extendModeY As D2D1_EXTEND_MODE, <Out> ByRef transform As ID2D1BorderTransform) As HRESULT
        <PreserveSig>
        Function CreateOffsetTransform(ByRef offset As D2D1_POINT_2L, <Out> ByRef transform As ID2D1OffsetTransform) As HRESULT
        <PreserveSig>
        Function CreateBoundsAdjustmentTransform(ByRef outputRectangle As D2D1_RECT_L, <Out> ByRef transform As ID2D1BoundsAdjustmentTransform) As HRESULT
        <PreserveSig>
        Function LoadPixelShader(ByRef shaderId As Guid, shaderBuffer As IntPtr, shaderBufferCount As UInteger) As HRESULT
        <PreserveSig>
        Function LoadVertexShader(ByRef resourceId As Guid, shaderBuffer As IntPtr, shaderBufferCount As UInteger) As HRESULT
        <PreserveSig>
        Function LoadComputeShader(ByRef resourceId As Guid, shaderBuffer As IntPtr, shaderBufferCount As UInteger) As HRESULT
        <PreserveSig>
        Function IsShaderLoaded(ByRef shaderId As Guid) As Boolean
        Function CreateResourceTexture(ByRef resourceId As Guid, ByRef resourceTextureProperties As D2D1_RESOURCE_TEXTURE_PROPERTIES, data As IntPtr, strides As UInteger, dataSize As UInteger, <Out> ByRef resourceTexture As ID2D1ResourceTexture) As HRESULT
        <PreserveSig>
        Function FindResourceTexture(ByRef resourceId As Guid, <Out> ByRef resourceTexture As ID2D1ResourceTexture) As HRESULT
        <PreserveSig>
        Function CreateVertexBuffer(ByRef vertexBufferProperties As D2D1_VERTEX_BUFFER_PROPERTIES, ByRef resourceId As Guid, ByRef customVertexBufferProperties As D2D1_CUSTOM_VERTEX_BUFFER_PROPERTIES, <Out> ByRef buffer As ID2D1VertexBuffer) As HRESULT
        <PreserveSig>
        Function FindVertexBuffer(ByRef resourceId As Guid, <Out> ByRef buffer As ID2D1VertexBuffer) As HRESULT
        <PreserveSig>
        Function CreateColorContext(space As D2D1_COLOR_SPACE, profile As IntPtr, profileSize As UInteger, <Out> ByRef colorContext As ID2D1ColorContext) As HRESULT
        <PreserveSig>
        Function CreateColorContextFromFilename(filename As String, <Out> ByRef colorContext As ID2D1ColorContext) As HRESULT
        <PreserveSig>
        Function CreateColorContextFromWicColorContext(wicColorContext As IWICColorContext, <Out> ByRef colorContext As ID2D1ColorContext) As HRESULT
        <PreserveSig>
        Function CheckFeatureSupport(feature As D2D1_FEATURE, <Out> ByRef featureSupportData As IntPtr, featureSupportDataSize As UInteger) As HRESULT
        <PreserveSig>
        Function IsBufferPrecisionSupported(bufferPrecision As D2D1_BUFFER_PRECISION) As Boolean
    End Interface

    <Guid("84ab595a-fc81-4546-bacd-e8ef4d8abe7a")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1EffectContext1
        Inherits ID2D1EffectContext
#Region "<ID2D1EffectContext>"
        <PreserveSig>
        Overloads Sub GetDpi(<Out> ByRef dpiX As Single, <Out> ByRef dpiY As Single)
        <PreserveSig>
        Overloads Function CreateEffect(ByRef effectId As Guid, <Out> ByRef effect As ID2D1Effect) As HRESULT
        <PreserveSig>
        Overloads Function GetMaximumSupportedFeatureLevel(
         <MarshalAs(UnmanagedType.LPArray)> featureLevels As Integer(), featureLevelsCount As UInteger, <Out> ByRef maximumSupportedFeatureLevel As D3D_FEATURE_LEVEL) As HRESULT
        <PreserveSig>
        Overloads Function CreateTransformNodeFromEffect(effect As ID2D1Effect, <Out> ByRef transformNode As ID2D1TransformNode) As HRESULT
        <PreserveSig>
        Overloads Function CreateBlendTransform(numInputs As UInteger, blendDescription As D2D1_BLEND_DESCRIPTION, <Out> ByRef transform As ID2D1BlendTransform) As HRESULT
        <PreserveSig>
        Overloads Function CreateBorderTransform(extendModeX As D2D1_EXTEND_MODE, extendModeY As D2D1_EXTEND_MODE, <Out> ByRef transform As ID2D1BorderTransform) As HRESULT
        <PreserveSig>
        Overloads Function CreateOffsetTransform(ByRef offset As D2D1_POINT_2L, <Out> ByRef transform As ID2D1OffsetTransform) As HRESULT
        <PreserveSig>
        Overloads Function CreateBoundsAdjustmentTransform(ByRef outputRectangle As D2D1_RECT_L, <Out> ByRef transform As ID2D1BoundsAdjustmentTransform) As HRESULT
        <PreserveSig>
        Overloads Function LoadPixelShader(ByRef shaderId As Guid, shaderBuffer As IntPtr, shaderBufferCount As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function LoadVertexShader(ByRef resourceId As Guid, shaderBuffer As IntPtr, shaderBufferCount As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function LoadComputeShader(ByRef resourceId As Guid, shaderBuffer As IntPtr, shaderBufferCount As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function IsShaderLoaded(ByRef shaderId As Guid) As Boolean
        <PreserveSig>
        Overloads Function CreateResourceTexture(ByRef resourceId As Guid, ByRef resourceTextureProperties As D2D1_RESOURCE_TEXTURE_PROPERTIES, data As IntPtr, strides As UInteger, dataSize As UInteger, <Out> ByRef resourceTexture As ID2D1ResourceTexture) As HRESULT
        <PreserveSig>
        Overloads Function FindResourceTexture(ByRef resourceId As Guid, <Out> ByRef resourceTexture As ID2D1ResourceTexture) As HRESULT
        <PreserveSig>
        Overloads Function CreateVertexBuffer(ByRef vertexBufferProperties As D2D1_VERTEX_BUFFER_PROPERTIES, ByRef resourceId As Guid, ByRef customVertexBufferProperties As D2D1_CUSTOM_VERTEX_BUFFER_PROPERTIES, <Out> ByRef buffer As ID2D1VertexBuffer) As HRESULT
        <PreserveSig>
        Overloads Function FindVertexBuffer(ByRef resourceId As Guid, <Out> ByRef buffer As ID2D1VertexBuffer) As HRESULT
        <PreserveSig>
        Overloads Function CreateColorContext(space As D2D1_COLOR_SPACE, profile As IntPtr, profileSize As UInteger, <Out> ByRef colorContext As ID2D1ColorContext) As HRESULT
        <PreserveSig>
        Overloads Function CreateColorContextFromFilename(filename As String, <Out> ByRef colorContext As ID2D1ColorContext) As HRESULT
        <PreserveSig>
        Overloads Function CreateColorContextFromWicColorContext(wicColorContext As IWICColorContext, <Out> ByRef colorContext As ID2D1ColorContext) As HRESULT
        <PreserveSig>
        Overloads Function CheckFeatureSupport(feature As D2D1_FEATURE, <Out> ByRef featureSupportData As IntPtr, featureSupportDataSize As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function IsBufferPrecisionSupported(bufferPrecision As D2D1_BUFFER_PRECISION) As Boolean
#End Region

        <PreserveSig>
        Function CreateLookupTable3D(precision As D2D1_BUFFER_PRECISION,
<MarshalAs(UnmanagedType.LPArray, SizeConst:=3)> extents As UInteger(),
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=3)> data As Byte(), dataCount As UInteger,
        <MarshalAs(UnmanagedType.LPArray, SizeConst:=2)> strides As UInteger(), <Out> ByRef lookupTable As ID2D1LookupTable3D) As HRESULT
    End Interface

    <Guid("577ad2a0-9fc7-4dda-8b18-dab810140052")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1EffectContext2
        Inherits ID2D1EffectContext1
#Region "<ID2D1EffectContext>"
        <PreserveSig>
        Overloads Sub GetDpi(<Out> ByRef dpiX As Single, <Out> ByRef dpiY As Single)
        <PreserveSig>
        Overloads Function CreateEffect(ByRef effectId As Guid, <Out> ByRef effect As ID2D1Effect) As HRESULT
        <PreserveSig>
        Overloads Function GetMaximumSupportedFeatureLevel(
         <MarshalAs(UnmanagedType.LPArray)> featureLevels As Integer(), featureLevelsCount As UInteger, <Out> ByRef maximumSupportedFeatureLevel As D3D_FEATURE_LEVEL) As HRESULT
        <PreserveSig>
        Overloads Function CreateTransformNodeFromEffect(effect As ID2D1Effect, <Out> ByRef transformNode As ID2D1TransformNode) As HRESULT
        <PreserveSig>
        Overloads Function CreateBlendTransform(numInputs As UInteger, blendDescription As D2D1_BLEND_DESCRIPTION, <Out> ByRef transform As ID2D1BlendTransform) As HRESULT
        <PreserveSig>
        Overloads Function CreateBorderTransform(extendModeX As D2D1_EXTEND_MODE, extendModeY As D2D1_EXTEND_MODE, <Out> ByRef transform As ID2D1BorderTransform) As HRESULT
        <PreserveSig>
        Overloads Function CreateOffsetTransform(ByRef offset As D2D1_POINT_2L, <Out> ByRef transform As ID2D1OffsetTransform) As HRESULT
        <PreserveSig>
        Overloads Function CreateBoundsAdjustmentTransform(ByRef outputRectangle As D2D1_RECT_L, <Out> ByRef transform As ID2D1BoundsAdjustmentTransform) As HRESULT
        <PreserveSig>
        Overloads Function LoadPixelShader(ByRef shaderId As Guid, shaderBuffer As IntPtr, shaderBufferCount As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function LoadVertexShader(ByRef resourceId As Guid, shaderBuffer As IntPtr, shaderBufferCount As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function LoadComputeShader(ByRef resourceId As Guid, shaderBuffer As IntPtr, shaderBufferCount As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function IsShaderLoaded(ByRef shaderId As Guid) As Boolean
        <PreserveSig>
        Overloads Function CreateResourceTexture(ByRef resourceId As Guid, ByRef resourceTextureProperties As D2D1_RESOURCE_TEXTURE_PROPERTIES, data As IntPtr, strides As UInteger, dataSize As UInteger, <Out> ByRef resourceTexture As ID2D1ResourceTexture) As HRESULT
        <PreserveSig>
        Overloads Function FindResourceTexture(ByRef resourceId As Guid, <Out> ByRef resourceTexture As ID2D1ResourceTexture) As HRESULT
        <PreserveSig>
        Overloads Function CreateVertexBuffer(ByRef vertexBufferProperties As D2D1_VERTEX_BUFFER_PROPERTIES, ByRef resourceId As Guid, ByRef customVertexBufferProperties As D2D1_CUSTOM_VERTEX_BUFFER_PROPERTIES, <Out> ByRef buffer As ID2D1VertexBuffer) As HRESULT
        <PreserveSig>
        Overloads Function FindVertexBuffer(ByRef resourceId As Guid, <Out> ByRef buffer As ID2D1VertexBuffer) As HRESULT
        <PreserveSig>
        Overloads Function CreateColorContext(space As D2D1_COLOR_SPACE, profile As IntPtr, profileSize As UInteger, <Out> ByRef colorContext As ID2D1ColorContext) As HRESULT
        <PreserveSig>
        Overloads Function CreateColorContextFromFilename(filename As String, <Out> ByRef colorContext As ID2D1ColorContext) As HRESULT
        <PreserveSig>
        Overloads Function CreateColorContextFromWicColorContext(wicColorContext As IWICColorContext, <Out> ByRef colorContext As ID2D1ColorContext) As HRESULT
        <PreserveSig>
        Overloads Function CheckFeatureSupport(feature As D2D1_FEATURE, <Out> ByRef featureSupportData As IntPtr, featureSupportDataSize As UInteger) As HRESULT
        <PreserveSig>
        Overloads Function IsBufferPrecisionSupported(bufferPrecision As D2D1_BUFFER_PRECISION) As Boolean
#End Region

#Region "<ID2D1EffectContext1>"
        <PreserveSig>
        Overloads Function CreateLookupTable3D(precision As D2D1_BUFFER_PRECISION,
<MarshalAs(UnmanagedType.LPArray, SizeConst:=3)> extents As UInteger(),
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=3)> data As Byte(), dataCount As UInteger,
        <MarshalAs(UnmanagedType.LPArray, SizeConst:=2)> strides As UInteger(), <Out> ByRef lookupTable As ID2D1LookupTable3D) As HRESULT
#End Region

        <PreserveSig>
        Function CreateColorContextFromDxgiColorSpace(colorSpace As DXGI_COLOR_SPACE_TYPE, <Out> ByRef colorContext As ID2D1ColorContext1) As HRESULT
        <PreserveSig>
        Function CreateColorContextFromSimpleColorProfile(ByRef simpleProfile As D2D1_SIMPLE_COLOR_PROFILE, <Out> ByRef colorContext As ID2D1ColorContext1) As HRESULT
    End Interface

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D1_POINT_2L
        Public x As Long
        Public y As Long

        Public Sub New(x As Long, y As Long)
            Me.x = x
            Me.y = y
        End Sub
    End Structure

    <ComImport> <Guid("3fe6adea-7643-4f53-bd14-a0ce63f24042")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID2D1OffsetTransform
        Inherits ID2D1TransformNode
#Region "<ID2D1TransformNode>"
        <PreserveSig>
        Overloads Function GetInputCount() As UInteger
#End Region

        <PreserveSig>
        Sub SetOffset(offset As D2D1_POINT_2L)
        <PreserveSig>
        Sub GetOffset(<Out> ByRef offset As D2D1_POINT_2L)
    End Interface

#If Not DWRITE Then
    <ComImport> <Guid("2f0da53a-2add-47cd-82ee-d9ec34688e75")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteRenderingParams
        <PreserveSig>
        Function GetGamma() As Single
        <PreserveSig>
        Function GetEnhancedContrast() As Single
        <PreserveSig>
        Function GetClearTypeLevel() As Single
        <PreserveSig>
        Sub GetPixelGeometry(<Out> ByRef pixelGeometry As DWRITE_PIXEL_GEOMETRY)
        <PreserveSig>
        Sub GetRenderingMode(<Out> ByRef renderingMode As DWRITE_RENDERING_MODE)
    End Interface

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

    <ComImport> <Guid("9c906818-31d7-4fd3-a151-7c5e225db55a")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteTextFormat
        <PreserveSig>
        Function SetTextAlignment(textAlignment As DWRITE_TEXT_ALIGNMENT) As HRESULT
        <PreserveSig>
        Function SetParagraphAlignment(paragraphAlignment As DWRITE_PARAGRAPH_ALIGNMENT) As HRESULT
        <PreserveSig>
        Function SetWordWrapping(wordWrapping As DWRITE_WORD_WRAPPING) As HRESULT
        <PreserveSig>
        Function SetReadingDirection(readingDirection As DWRITE_READING_DIRECTION) As HRESULT
        <PreserveSig>
        Function SetFlowDirection(flowDirection As DWRITE_FLOW_DIRECTION) As HRESULT
        <PreserveSig>
        Function SetIncrementalTabStop(incrementalTabStop As Single) As HRESULT
        <PreserveSig>
        Function SetTrimming(trimmingOptions As DWRITE_TRIMMING, trimmingSign As IDWriteInlineObject) As HRESULT
        <PreserveSig>
        Function SetLineSpacing(lineSpacingMethod As DWRITE_LINE_SPACING_METHOD, lineSpacing As Single, baseline As Single) As HRESULT
        <PreserveSig>
        Sub GetTextAlignment(<Out> ByRef textAlignment As DWRITE_TEXT_ALIGNMENT)
        <PreserveSig>
        Sub GetParagraphAlignment(<Out> ByRef paragraphAlignment As DWRITE_PARAGRAPH_ALIGNMENT)
        <PreserveSig>
        Sub GetWordWrapping(<Out> ByRef wordWrapping As DWRITE_WORD_WRAPPING)
        <PreserveSig>
        Sub GetReadingDirection(<Out> ByRef readingDirection As DWRITE_READING_DIRECTION)
        <PreserveSig>
        Sub GetFlowDirection(<Out> ByRef flowDirection As DWRITE_FLOW_DIRECTION)
        <PreserveSig>
        Function GetIncrementalTabStop() As Single
        <PreserveSig>
        Function GetTrimming(<Out> ByRef trimmingOptions As DWRITE_TRIMMING, <Out> ByRef trimmingSign As IDWriteInlineObject) As HRESULT
        <PreserveSig>
        Function GetLineSpacing(<Out> ByRef lineSpacingMethod As DWRITE_LINE_SPACING_METHOD, <Out> ByRef lineSpacing As Single, <Out> ByRef baseline As Single) As HRESULT
        <PreserveSig>
        Function GetFontCollection(<Out> ByRef fontCollection As IDWriteFontCollection) As HRESULT
        <PreserveSig>
        Function GetFontFamilyNameLength() As UInteger
        <PreserveSig>
        Function GetFontFamilyName(<Out> ByRef fontFamilyName As String, nameSize As UInteger) As HRESULT
        <PreserveSig>
        Sub GetFontWeight(<Out> ByRef fontWeight As DWRITE_FONT_WEIGHT)
        <PreserveSig>
        Sub GetFontStyle(<Out> ByRef fontStyle As DWRITE_FONT_STYLE)
        <PreserveSig>
        Sub GetFontStretch(<Out> ByRef fontStretch As DWRITE_FONT_STRETCH)
        <PreserveSig>
        Function GetFontSize() As Single
        <PreserveSig>
        Function GetLocaleNameLength() As UInteger
        <PreserveSig>
        Function GetLocaleName(<Out> ByRef localeName As String, nameSize As UInteger) As HRESULT
    End Interface

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

    <ComImport> <Guid("8339FDE3-106F-47ab-8373-1C6295EB10B3")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteInlineObject
        <PreserveSig>
        Function Draw(clientDrawingContext As IntPtr, renderer As IDWriteTextRenderer, originX As Single, originY As Single, isSideways As Boolean, isRightToLeft As Boolean, clientDrawingEffect As IntPtr) As HRESULT
        <PreserveSig>
        Function GetMetrics(<Out> ByRef metrics As DWRITE_INLINE_OBJECT_METRICS) As HRESULT
        <PreserveSig>
        Function GetOverhangMetrics(<Out> ByRef overhangs As DWRITE_OVERHANG_METRICS) As HRESULT
        <PreserveSig>
        Function GetBreakConditions(<Out> ByRef breakConditionBefore As DWRITE_BREAK_CONDITION, <Out> ByRef breakConditionAfter As DWRITE_BREAK_CONDITION) As HRESULT
    End Interface

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

    <ComImport> <Guid("eaf3a2da-ecf4-4d24-b644-b34f6842024b")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWritePixelSnapping
        <PreserveSig>
        Function IsPixelSnappingDisabled(clientDrawingContext As IntPtr, <Out> ByRef isDisabled As Boolean) As HRESULT
        <PreserveSig>
        Function GetCurrentTransform(clientDrawingContext As IntPtr, <Out> ByRef transform As DWRITE_MATRIX) As HRESULT
        <PreserveSig>
        Function GetPixelsPerDip(clientDrawingContext As IntPtr, <Out>
<MarshalAs(UnmanagedType.R4)> ByRef pixelsPerDip As Single) As HRESULT
    End Interface

    <ComImport> <Guid("ef8a8135-5cc6-45fe-8825-c5a0724eb819")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
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

    <ComImport> <Guid("a84cee02-3eea-4eee-a827-87c1a02a0fcc")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontCollection
        '[return: MarshalAs(UnmanagedType.U4)]
        <PreserveSig>
        Function GetFontFamilyCount() As UInteger
        <PreserveSig>
        Function GetFontFamily(index As UInteger, <Out> ByRef fontFamily As IDWriteFontFamily) As HRESULT
        <PreserveSig>
        Function FindFamilyName(familyName As String, <Out> ByRef index As UInteger, <Out> ByRef exists As Boolean) As HRESULT
        <PreserveSig>
        Function GetFontFromFontFace(fontFace As IDWriteFontFace, <Out> ByRef font As IDWriteFont) As HRESULT
    End Interface

    <ComImport> <Guid("1a0d8438-1d97-4ec1-aef9-a2fb86ed6acb")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontList
        <PreserveSig>
        Function GetFontCollection(<Out> ByRef fontCollection As IDWriteFontCollection) As HRESULT
        <PreserveSig>
        Function GetFontCount() As Integer
        <PreserveSig>
        Function GetFont(index As Integer, <Out> ByRef font As IDWriteFont) As HRESULT
    End Interface

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

    <ComImport> <Guid("da20d8ef-812a-4c43-9802-62ec4abd7add")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontFamily
        Inherits IDWriteFontList
#Region "IDWriteFontList"
        <PreserveSig>
        Overloads Function GetFontCollection(<Out> ByRef fontCollection As IDWriteFontCollection) As HRESULT
        <PreserveSig>
        Overloads Function GetFontCount() As Integer
        <PreserveSig>
        Overloads Function GetFont(index As Integer, <Out> ByRef font As IDWriteFont) As HRESULT
#End Region

        <PreserveSig>
        Function GetFamilyNames(<Out> ByRef names As IDWriteLocalizedStrings) As HRESULT
        <PreserveSig>
        Function GetFirstMatchingFont(weight As DWRITE_FONT_WEIGHT, stretch As DWRITE_FONT_STRETCH, style As DWRITE_FONT_STYLE, <Out> ByRef matchingFont As IDWriteFont) As HRESULT
        <PreserveSig>
        Function GetMatchingFonts(weight As DWRITE_FONT_WEIGHT, stretch As DWRITE_FONT_STRETCH, style As DWRITE_FONT_STYLE, <Out> ByRef matchingFonts As IDWriteFontList) As HRESULT
    End Interface

    <ComImport> <Guid("acd16696-8c14-4f5d-877e-fe3fc1d32737")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFont
        <PreserveSig>
        Function GetFontFamily(<Out> ByRef fontFamily As IDWriteFontFamily) As HRESULT
        <PreserveSig>
        Sub GetWeight(<Out> ByRef fontWeight As DWRITE_FONT_WEIGHT)
        <PreserveSig>
        Sub GetStretch(<Out> ByRef fontStretch As DWRITE_FONT_STRETCH)
        <PreserveSig>
        Sub GetStyle(<Out> ByRef fontStyle As DWRITE_FONT_STYLE)
        <PreserveSig>
        Function IsSymbolFont() As Boolean
        <PreserveSig>
        Function GetFaceNames(<Out> ByRef names As IDWriteLocalizedStrings) As HRESULT
        <PreserveSig>
        Function GetInformationalStrings(informationalStringID As DWRITE_INFORMATIONAL_STRING_ID, <Out> ByRef informationalStrings As IDWriteLocalizedStrings, <Out> ByRef exists As Boolean) As HRESULT
        <PreserveSig>
        Sub GetSimulations(<Out> ByRef fontSimulations As DWRITE_FONT_SIMULATIONS)
        <PreserveSig>
        Sub GetMetrics(<Out> ByRef fontMetrics As DWRITE_FONT_METRICS)
        <PreserveSig>
        Function HasCharacter(unicodeValue As Integer, <Out> ByRef exists As Boolean) As HRESULT
        <PreserveSig>
        Function CreateFontFace(<Out> ByRef fontFace As IDWriteFontFace) As HRESULT
    End Interface

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

    <ComImport> <Guid("08256209-099a-4b34-b86d-c22b110e7771")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteLocalizedStrings
        <PreserveSig>
        Function GetCount() As Integer
        <PreserveSig>
        Function FindLocaleName(localeName As String, <Out> ByRef index As UInteger, <Out> ByRef exists As Boolean) As HRESULT
        <PreserveSig>
        Function GetLocaleNameLength(index As UInteger, <Out> ByRef length As UInteger) As HRESULT
        <PreserveSig>
        Function GetLocaleName(index As UInteger, <Out> ByRef localeName As String, size As UInteger) As HRESULT
        <PreserveSig>
        Function GetStringLength(index As UInteger, <Out> ByRef length As UInteger) As HRESULT
        <PreserveSig>
        Function GetString(index As UInteger, stringBuffer As Text.StringBuilder, size As UInteger) As HRESULT
    End Interface

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

    <ComImport> <Guid("5f49804d-7024-4d43-bfa9-d25984f53849")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontFace
        <PreserveSig>
        Sub [GetType](<Out> ByRef fontFaceType As DWRITE_FONT_FACE_TYPE)
        <PreserveSig>
        Function GetFiles(
<[In], Out> ByRef numberOfFiles As UInteger,
<Out, MarshalAs(UnmanagedType.LPArray)> fontFiles As IDWriteFontFile()) As HRESULT
        <PreserveSig>
        Function GetIndex() As Integer
        <PreserveSig>
        Sub GetSimulations(<Out> ByRef fontSimulations As DWRITE_FONT_SIMULATIONS)
        <PreserveSig>
        Function IsSymbolFont() As Boolean
        <PreserveSig>
        Sub GetMetrics(<Out> ByRef fontFaceMetrics As DWRITE_FONT_METRICS)
        <PreserveSig>
        Function GetGlyphCount() As UShort
        <PreserveSig>
        Function GetDesignGlyphMetrics(
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> glyphIndices As UShort(), glyphCount As Integer,
<Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> glyphMetrics As DWRITE_GLYPH_METRICS(), Optional isSideways As Boolean = False) As HRESULT
        <PreserveSig>
        Function GetGlyphIndices(
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> codePoints As UInteger(), codePointCount As Integer,
<Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> glyphIndices As UShort()) As HRESULT
        <PreserveSig>
        Function TryGetFontTable(openTypeTableTag As Integer, <Out> ByRef tableData As IntPtr, <Out> ByRef tableSize As Integer, <Out> ByRef tableContext As IntPtr, <Out> ByRef exists As Boolean) As HRESULT
        <PreserveSig>
        Sub ReleaseFontTable(tableContext As IntPtr)
        <PreserveSig>
        Function GetGlyphRunOutline(emSize As Single,
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=4)> glyphIndices As UShort(),
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=4)> glyphAdvances As Single(),
<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=4)> glyphOffsets As DWRITE_GLYPH_OFFSET(), glyphCount As Integer, isSideways As Boolean, isRightToLeft As Boolean, geometrySink As ID2D1SimplifiedGeometrySink) As HRESULT
        <PreserveSig>
        Function GetRecommendedRenderingMode(emSize As Single, pixelsPerDip As Single, measuringMode As DWRITE_MEASURING_MODE, renderingParams As IDWriteRenderingParams, <Out> ByRef renderingMode As DWRITE_RENDERING_MODE) As HRESULT
        <PreserveSig>
        Function GetGdiCompatibleMetrics(emSize As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, <Out> ByRef fontFaceMetrics As DWRITE_FONT_METRICS) As HRESULT
        <PreserveSig>
        Function GetGdiCompatibleGlyphMetrics(emSize As Single, pixelsPerDip As Single, transform As DWRITE_MATRIX, useGdiNatural As Boolean, glyphIndices As UShort, glyphCount As Integer, <Out> ByRef glyphMetrics As DWRITE_GLYPH_METRICS, Optional isSideways As Boolean = False) As HRESULT
    End Interface

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
        DWRITE_FONT_SIMULATIONS_NONE = &H0
        ''' <summary>
        ''' Algorithmic emboldening is performed.
        ''' </summary>
        DWRITE_FONT_SIMULATIONS_BOLD = &H1
        ''' <summary>
        ''' Algorithmic italicization is performed.
        ''' </summary>
        DWRITE_FONT_SIMULATIONS_OBLIQUE = &H2
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

    <ComImport> <Guid("53737037-6d14-410b-9bfe-0b182bb70961")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteTextLayout
        Inherits IDWriteTextFormat
#Region "IDWriteTextFormat"
        <PreserveSig>
        Overloads Function SetTextAlignment(textAlignment As DWRITE_TEXT_ALIGNMENT) As HRESULT
        <PreserveSig>
        Overloads Function SetParagraphAlignment(paragraphAlignment As DWRITE_PARAGRAPH_ALIGNMENT) As HRESULT
        <PreserveSig>
        Overloads Function SetWordWrapping(wordWrapping As DWRITE_WORD_WRAPPING) As HRESULT
        <PreserveSig>
        Overloads Function SetReadingDirection(readingDirection As DWRITE_READING_DIRECTION) As HRESULT
        <PreserveSig>
        Overloads Function SetFlowDirection(flowDirection As DWRITE_FLOW_DIRECTION) As HRESULT
        <PreserveSig>
        Overloads Function SetIncrementalTabStop(incrementalTabStop As Single) As HRESULT
        <PreserveSig>
        Overloads Function SetTrimming(trimmingOptions As DWRITE_TRIMMING, trimmingSign As IDWriteInlineObject) As HRESULT
        <PreserveSig>
        Overloads Function SetLineSpacing(lineSpacingMethod As DWRITE_LINE_SPACING_METHOD, lineSpacing As Single, baseline As Single) As HRESULT
        <PreserveSig>
        Overloads Sub GetTextAlignment(<Out> ByRef textAlignment As DWRITE_TEXT_ALIGNMENT)
        <PreserveSig>
        Overloads Sub GetParagraphAlignment(<Out> ByRef paragraphAlignment As DWRITE_PARAGRAPH_ALIGNMENT)
        <PreserveSig>
        Overloads Sub GetWordWrapping(<Out> ByRef wordWrapping As DWRITE_WORD_WRAPPING)
        <PreserveSig>
        Overloads Sub GetReadingDirection(<Out> ByRef readingDirection As DWRITE_READING_DIRECTION)
        <PreserveSig>
        Overloads Sub GetFlowDirection(<Out> ByRef flowDirection As DWRITE_FLOW_DIRECTION)
        <PreserveSig>
        Overloads Function GetIncrementalTabStop() As Single
        <PreserveSig>
        Overloads Function GetTrimming(<Out> ByRef trimmingOptions As DWRITE_TRIMMING, <Out> ByRef trimmingSign As IDWriteInlineObject) As HRESULT
        <PreserveSig>
        Overloads Function GetLineSpacing(<Out> ByRef lineSpacingMethod As DWRITE_LINE_SPACING_METHOD, <Out> ByRef lineSpacing As Single, <Out> ByRef baseline As Single) As HRESULT
        <PreserveSig>
        Overloads Function GetFontCollection(<Out> ByRef fontCollection As IDWriteFontCollection) As HRESULT
        <PreserveSig>
        Overloads Function GetFontFamilyNameLength() As UInteger
        <PreserveSig>
        Overloads Function GetFontFamilyName(<Out> ByRef fontFamilyName As String, nameSize As UInteger) As HRESULT
        <PreserveSig>
        Overloads Sub GetFontWeight(<Out> ByRef fontWeight As DWRITE_FONT_WEIGHT)
        <PreserveSig>
        Overloads Sub GetFontStyle(<Out> ByRef fontStyle As DWRITE_FONT_STYLE)
        <PreserveSig>
        Overloads Sub GetFontStretch(<Out> ByRef fontStretch As DWRITE_FONT_STRETCH)
        <PreserveSig>
        Overloads Function GetFontSize() As Single
        <PreserveSig>
        Overloads Function GetLocaleNameLength() As UInteger
        <PreserveSig>
        Overloads Function GetLocaleName(<Out> ByRef localeName As String, nameSize As UInteger) As HRESULT
#End Region

        <PreserveSig>
        Function SetMaxWidth(maxWidth As Single) As HRESULT
        <PreserveSig>
        Function SetMaxHeight(maxHeight As Single) As HRESULT
        <PreserveSig>
        Function SetFontCollection(fontCollection As IDWriteFontCollection, textRange As DWRITE_TEXT_RANGE) As HRESULT
        <PreserveSig>
        Function SetFontFamilyName(fontFamilyName As String, textRange As DWRITE_TEXT_RANGE) As HRESULT
        <PreserveSig>
        Function SetFontWeight(fontWeight As DWRITE_FONT_WEIGHT, textRange As DWRITE_TEXT_RANGE) As HRESULT
        <PreserveSig>
        Function SetFontStyle(fontStyle As DWRITE_FONT_STYLE, textRange As DWRITE_TEXT_RANGE) As HRESULT
        <PreserveSig>
        Function SetFontStretch(fontStretch As DWRITE_FONT_STRETCH, textRange As DWRITE_TEXT_RANGE) As HRESULT
        <PreserveSig>
        Function SetFontSize(fontSize As Single, textRange As DWRITE_TEXT_RANGE) As HRESULT
        <PreserveSig>
        Function SetUnderline(hasUnderline As Boolean, textRange As DWRITE_TEXT_RANGE) As HRESULT
        <PreserveSig>
        Function SetStrikethrough(hasStrikethrough As Boolean, textRange As DWRITE_TEXT_RANGE) As HRESULT
        <PreserveSig>
        Function SetDrawingEffect(drawingEffect As IntPtr, textRange As DWRITE_TEXT_RANGE) As HRESULT
        <PreserveSig>
        Function SetInlineObject(inlineObject As IDWriteInlineObject, textRange As DWRITE_TEXT_RANGE) As HRESULT
        <PreserveSig>
        Function SetTypography(typography As IDWriteTypography, textRange As DWRITE_TEXT_RANGE) As HRESULT
        <PreserveSig>
        Function SetLocaleName(localeName As String, textRange As DWRITE_TEXT_RANGE) As HRESULT
        <PreserveSig>
        Function GetMaxWidth() As Single
        <PreserveSig>
        Function GetMaxHeight() As Single
        <PreserveSig>
        Overloads Function GetFontCollection(currentPosition As UInteger, <Out> ByRef fontCollection As IDWriteFontCollection, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        <PreserveSig>
        Overloads Function GetFontFamilyNameLength(currentPosition As UInteger, <Out> ByRef nameLength As UInteger, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        <PreserveSig>
        Overloads Function GetFontFamilyName(currentPosition As UInteger, <Out> ByRef fontFamilyName As String, nameSize As UInteger, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        <PreserveSig>
        Overloads Function GetFontWeight(currentPosition As UInteger, <Out> ByRef fontWeight As DWRITE_FONT_WEIGHT, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        <PreserveSig>
        Overloads Function GetFontStyle(currentPosition As UInteger, <Out> ByRef fontStyle As DWRITE_FONT_STYLE, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        <PreserveSig>
        Overloads Function GetFontStretch(currentPosition As UInteger, <Out> ByRef fontStretch As DWRITE_FONT_STRETCH, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        <PreserveSig>
        Overloads Function GetFontSize(currentPosition As UInteger, <Out> ByRef fontSize As Single, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        <PreserveSig>
        Function GetUnderline(currentPosition As UInteger, <Out> ByRef hasUnderline As Boolean, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        <PreserveSig>
        Function GetStrikethrough(currentPosition As UInteger, <Out> ByRef hasStrikethrough As Boolean, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        <PreserveSig>
        Function GetDrawingEffect(currentPosition As UInteger, <Out> ByRef drawingEffect As IntPtr, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        <PreserveSig>
        Function GetInlineObject(currentPosition As UInteger, <Out> ByRef inlineObject As IDWriteInlineObject, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        <PreserveSig>
        Function GetTypography(currentPosition As UInteger, <Out> ByRef typography As IDWriteTypography, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        <PreserveSig>
        Overloads Function GetLocaleNameLength(currentPosition As UInteger, <Out> ByRef nameLength As UInteger, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        <PreserveSig>
        Overloads Function GetLocaleName(currentPosition As UInteger, <Out> ByRef localeName As String, nameSize As UInteger, <Out> ByRef textRange As DWRITE_TEXT_RANGE) As HRESULT
        <PreserveSig>
        Function Draw(clientDrawingContext As IntPtr, renderer As IDWriteTextRenderer, originX As Single, originY As Single) As HRESULT
        <PreserveSig>
        Function GetLineMetrics(<Out> ByRef lineMetrics As DWRITE_LINE_METRICS, maxLineCount As UInteger, <Out> ByRef actualLineCount As UInteger) As HRESULT
        <PreserveSig>
        Function GetMetrics(<Out> ByRef textMetrics As DWRITE_TEXT_METRICS) As HRESULT
        <PreserveSig>
        Function GetOverhangMetrics(<Out> ByRef overhangs As DWRITE_OVERHANG_METRICS) As HRESULT
        <PreserveSig>
        Function GetClusterMetrics(<Out> ByRef clusterMetrics As DWRITE_CLUSTER_METRICS, maxClusterCount As UInteger, <Out> ByRef actualClusterCount As UInteger) As HRESULT
        <PreserveSig>
        Function DetermineMinWidth(<Out> ByRef minWidth As Single) As HRESULT
        <PreserveSig>
        Function HitTestPoint(pointX As Single, pointY As Single, <Out> ByRef isTrailingHit As Boolean, <Out> ByRef isInside As Boolean, <Out> ByRef hitTestMetrics As DWRITE_HIT_TEST_METRICS) As HRESULT
        <PreserveSig>
        Function HitTestTextPosition(textPosition As UInteger, isTrailingHit As Boolean, <Out> ByRef pointX As Single, <Out> ByRef pointY As Single, <Out> ByRef hitTestMetrics As DWRITE_HIT_TEST_METRICS) As HRESULT
        <PreserveSig>
        Function HitTestTextRange(textPosition As UInteger, textLength As UInteger, originX As Single, originY As Single, <Out> ByRef hitTestMetrics As DWRITE_HIT_TEST_METRICS, maxHitTestMetricsCount As UInteger, <Out> ByRef actualHitTestMetricsCount As UInteger) As HRESULT
    End Interface

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

    <ComImport> <Guid("55f1112b-1dc2-4b3c-9541-f46894ed85b6")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteTypography
        <PreserveSig>
        Function AddFontFeature(fontFeature As DWRITE_FONT_FEATURE) As HRESULT
        <PreserveSig>
        Function GetFontFeatureCount() As UInteger
        <PreserveSig>
        Function GetFontFeature(fontFeatureIndex As UInteger, <Out> ByRef fontFeature As DWRITE_FONT_FEATURE) As HRESULT
    End Interface

    Public Enum DWRITE_FONT_FEATURE_TAG
        DWRITE_FONT_FEATURE_TAG_ALTERNATIVE_FRACTIONS = &H63726661 ' 'afrc'
        DWRITE_FONT_FEATURE_TAG_PETITE_CAPITALS_FROM_CAPITALS = &H63703263 ' 'c2pc'
        DWRITE_FONT_FEATURE_TAG_SMALL_CAPITALS_FROM_CAPITALS = &H63733263 ' 'c2sc'
        DWRITE_FONT_FEATURE_TAG_CONTEXTUAL_ALTERNATES = &H746C6163 ' 'calt'
        DWRITE_FONT_FEATURE_TAG_CASE_SENSITIVE_FORMS = &H65736163 ' 'case'
        DWRITE_FONT_FEATURE_TAG_GLYPH_COMPOSITION_DECOMPOSITION = &H706D6363 ' 'ccmp'
        DWRITE_FONT_FEATURE_TAG_CONTEXTUAL_LIGATURES = &H67696C63 ' 'clig'
        DWRITE_FONT_FEATURE_TAG_CAPITAL_SPACING = &H70737063 ' 'cpsp'
        DWRITE_FONT_FEATURE_TAG_CONTEXTUAL_SWASH = &H68777363 ' 'cswh'
        DWRITE_FONT_FEATURE_TAG_CURSIVE_POSITIONING = &H73727563 ' 'curs'
        DWRITE_FONT_FEATURE_TAG_DEFAULT = &H746C6664 ' 'dflt'
        DWRITE_FONT_FEATURE_TAG_DISCRETIONARY_LIGATURES = &H67696C64 ' 'dlig'
        DWRITE_FONT_FEATURE_TAG_EXPERT_FORMS = &H74707865 ' 'expt'
        DWRITE_FONT_FEATURE_TAG_FRACTIONS = &H63617266 ' 'frac'
        DWRITE_FONT_FEATURE_TAG_FULL_WIDTH = &H64697766 ' 'fwid'
        DWRITE_FONT_FEATURE_TAG_HALF_FORMS = &H666C6168 ' 'half'
        DWRITE_FONT_FEATURE_TAG_HALANT_FORMS = &H6E6C6168 ' 'haln'
        DWRITE_FONT_FEATURE_TAG_ALTERNATE_HALF_WIDTH = &H746C6168 ' 'halt'
        DWRITE_FONT_FEATURE_TAG_HISTORICAL_FORMS = &H74736968 ' 'hist'
        DWRITE_FONT_FEATURE_TAG_HORIZONTAL_KANA_ALTERNATES = &H616E6B68 ' 'hkna'
        DWRITE_FONT_FEATURE_TAG_HISTORICAL_LIGATURES = &H67696C68 ' 'hlig'
        DWRITE_FONT_FEATURE_TAG_HALF_WIDTH = &H64697768 ' 'hwid'
        DWRITE_FONT_FEATURE_TAG_HOJO_KANJI_FORMS = &H6F6A6F68 ' 'hojo'
        DWRITE_FONT_FEATURE_TAG_JIS04_FORMS = &H3430706A ' 'jp04'
        DWRITE_FONT_FEATURE_TAG_JIS78_FORMS = &H3837706A ' 'jp78'
        DWRITE_FONT_FEATURE_TAG_JIS83_FORMS = &H3338706A ' 'jp83'
        DWRITE_FONT_FEATURE_TAG_JIS90_FORMS = &H3039706A ' 'jp90'
        DWRITE_FONT_FEATURE_TAG_KERNING = &H6E72656B ' 'kern'
        DWRITE_FONT_FEATURE_TAG_STANDARD_LIGATURES = &H6167696C ' 'liga'
        DWRITE_FONT_FEATURE_TAG_LINING_FIGURES = &H6D756E6C ' 'lnum'
        DWRITE_FONT_FEATURE_TAG_LOCALIZED_FORMS = &H6C636F6C ' 'locl'
        DWRITE_FONT_FEATURE_TAG_MARK_POSITIONING = &H6B72616D ' 'mark'
        DWRITE_FONT_FEATURE_TAG_MATHEMATICAL_GREEK = &H6B72676D ' 'mgrk'
        DWRITE_FONT_FEATURE_TAG_MARK_TO_MARK_POSITIONING = &H6B6D6B6D ' 'mkmk'
        DWRITE_FONT_FEATURE_TAG_ALTERNATE_ANNOTATION_FORMS = &H746C616E ' 'nalt'
        DWRITE_FONT_FEATURE_TAG_NLC_KANJI_FORMS = &H6B636C6E ' 'nlck'
        DWRITE_FONT_FEATURE_TAG_OLD_STYLE_FIGURES = &H6D756E6F ' 'onum'
        DWRITE_FONT_FEATURE_TAG_ORDINALS = &H6E64726F ' 'ordn'
        DWRITE_FONT_FEATURE_TAG_PROPORTIONAL_ALTERNATE_WIDTH = &H746C6170 ' 'palt'
        DWRITE_FONT_FEATURE_TAG_PETITE_CAPITALS = &H70616370 ' 'pcap'
        DWRITE_FONT_FEATURE_TAG_PROPORTIONAL_FIGURES = &H6D756E70 ' 'pnum'
        DWRITE_FONT_FEATURE_TAG_PROPORTIONAL_WIDTHS = &H64697770 ' 'pwid'
        DWRITE_FONT_FEATURE_TAG_QUARTER_WIDTHS = &H64697771 ' 'qwid'
        DWRITE_FONT_FEATURE_TAG_REQUIRED_LIGATURES = &H67696C72 ' 'rlig'
        DWRITE_FONT_FEATURE_TAG_RUBY_NOTATION_FORMS = &H79627572 ' 'ruby'
        DWRITE_FONT_FEATURE_TAG_STYLISTIC_ALTERNATES = &H746C6173 ' 'salt'
        DWRITE_FONT_FEATURE_TAG_SCIENTIFIC_INFERIORS = &H666E6973 ' 'sinf'
        DWRITE_FONT_FEATURE_TAG_SMALL_CAPITALS = &H70636D73 ' 'smcp'
        DWRITE_FONT_FEATURE_TAG_SIMPLIFIED_FORMS = &H6C706D73 ' 'smpl'
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
        DWRITE_FONT_FEATURE_TAG_TITLING = &H6C746974 ' 'titl'
        DWRITE_FONT_FEATURE_TAG_TRADITIONAL_NAME_FORMS = &H6D616E74 ' 'tnam'
        DWRITE_FONT_FEATURE_TAG_TABULAR_FIGURES = &H6D756E74 ' 'tnum'
        DWRITE_FONT_FEATURE_TAG_TRADITIONAL_FORMS = &H64617274 ' 'trad'
        DWRITE_FONT_FEATURE_TAG_THIRD_WIDTHS = &H64697774 ' 'twid'
        DWRITE_FONT_FEATURE_TAG_UNICASE = &H63696E75 ' 'unic'
        DWRITE_FONT_FEATURE_TAG_VERTICAL_WRITING = &H74726576 ' 'vert'
        DWRITE_FONT_FEATURE_TAG_VERTICAL_ALTERNATES_AND_ROTATION = &H32747276 ' 'vrt2'
        DWRITE_FONT_FEATURE_TAG_SLASHED_ZERO = &H6F72657A ' 'zero'
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

    <ComImport> <Guid("727cad4e-d6af-4c9e-8a08-d695b11caa49")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontFileLoader
        <PreserveSig>
        Function CreateStreamFromKey(fontFileReferenceKey As IntPtr, fontFileReferenceKeySize As UInteger, <Out> ByRef fontFileStream As IDWriteFontFileStream) As HRESULT
    End Interface

    <ComImport> <Guid("739d886a-cef5-47dc-8769-1a8b41bebbb0")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontFile
        <PreserveSig>
        Function GetReferenceKey(<Out> ByRef fontFileReferenceKey As IntPtr, <Out> ByRef fontFileReferenceKeySize As Integer) As HRESULT
        <PreserveSig>
        Function GetLoader(<Out> ByRef fontFileLoader As IDWriteFontFileLoader) As HRESULT
        <PreserveSig>
        Function Analyze(<Out> ByRef isSupportedFontType As Boolean, <Out> ByRef fontFileType As DWRITE_FONT_FILE_TYPE, <Out> ByRef fontFaceType As DWRITE_FONT_FACE_TYPE, <Out> ByRef numberOfFaces As Integer) As HRESULT
    End Interface

    <ComImport> <Guid("6d4865fe-0ab8-4d91-8f62-5dd6be34a3e0")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDWriteFontFileStream
        <PreserveSig>
        Function ReadFileFragment(<Out> ByRef fragmentStart As IntPtr, fileOffset As ULong, fragmentSize As ULong, <Out> ByRef fragmentContext As IntPtr) As HRESULT
        <PreserveSig>
        Sub ReleaseFileFragment(fragmentContext As IntPtr)
        <PreserveSig>
        Function GetFileSize(<Out> ByRef fileSize As ULong) As HRESULT
        <PreserveSig>
        Function GetLastWriteTime(<Out> ByRef lastWriteTime As ULong) As HRESULT
    End Interface

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

#End If

    ' incomplete : d2d1effectauthor.h...
End Namespace
