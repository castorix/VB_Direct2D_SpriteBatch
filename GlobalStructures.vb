Imports System

Imports System.Runtime.InteropServices

Namespace Global.GlobalStructures
    Public Enum HRESULT As Integer
        S_OK = 0
        S_FALSE = 1
        E_NOTIMPL = &H80004001
        E_NOINTERFACE = &H80004002
        E_POINTER = &H80004003
        E_FAIL = &H80004005
        E_UNEXPECTED = &H8000FFFF
        E_OUTOFMEMORY = &H8007000E
        E_INVALIDARG = &H80070057
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure RECT
        Public left As Integer
        Public top As Integer
        Public right As Integer
        Public bottom As Integer
        Public Sub New(Left As Integer, Top As Integer, Right As Integer, Bottom As Integer)
            Me.left = Left
            Me.top = Top
            Me.right = Right
            Me.bottom = Bottom
        End Sub
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure POINT
        Public x As Integer
        Public y As Integer

        Public Sub New(x As Integer, y As Integer)
            Me.x = x
            Me.y = y
        End Sub
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure SIZE
        Public cx As Integer
        Public cy As Integer
        Public Sub New(cx As Integer, cy As Integer)
            Me.cx = cx
            Me.cy = cy
        End Sub
    End Structure

    <StructLayout(LayoutKind.Explicit)>
    Public Structure LARGE_INTEGER
        <FieldOffset(0)>
        Public LowPart As Integer
        <FieldOffset(4)>
        Public HighPart As Integer
        <FieldOffset(0)>
        Public QuadPart As Long
    End Structure

    Public Enum VARENUM
        VT_EMPTY = 0
        VT_NULL = 1
        VT_I2 = 2
        VT_I4 = 3
        VT_R4 = 4
        VT_R8 = 5
        VT_CY = 6
        VT_DATE = 7
        VT_BSTR = 8
        VT_DISPATCH = 9
        VT_ERROR = 10
        VT_BOOL = 11
        VT_VARIANT = 12
        VT_UNKNOWN = 13
        VT_DECIMAL = 14
        VT_I1 = 16
        VT_UI1 = 17
        VT_UI2 = 18
        VT_UI4 = 19
        VT_I8 = 20
        VT_UI8 = 21
        VT_INT = 22
        VT_UINT = 23
        VT_VOID = 24
        VT_HRESULT = 25
        VT_PTR = 26
        VT_SAFEARRAY = 27
        VT_CARRAY = 28
        VT_USERDEFINED = 29
        VT_LPSTR = 30
        VT_LPWSTR = 31
        VT_RECORD = 36
        VT_INT_PTR = 37
        VT_UINT_PTR = 38
        VT_FILETIME = 64
        VT_BLOB = 65
        VT_STREAM = 66
        VT_STORAGE = 67
        VT_STREAMED_OBJECT = 68
        VT_STORED_OBJECT = 69
        VT_BLOB_OBJECT = 70
        VT_CF = 71
        VT_CLSID = 72
        VT_VERSIONED_STREAM = 73
        VT_BSTR_BLOB = &HFFF
        VT_VECTOR = &H1000
        VT_ARRAY = &H2000
        VT_BYREF = &H4000
        VT_RESERVED = &H8000
        VT_ILLEGAL = &HFFFF
        VT_ILLEGALMASKED = &HFFF
        VT_TYPEMASK = &HFFF
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure PROPARRAY
        Public cElems As UInteger
        Public pElems As IntPtr
    End Structure

    <StructLayout(LayoutKind.Explicit, Pack:=1)>
    Public Structure PROPVARIANT
        <FieldOffset(0)>
        Public varType As UShort
        <FieldOffset(2)>
        Public wReserved1 As UShort
        <FieldOffset(4)>
        Public wReserved2 As UShort
        <FieldOffset(6)>
        Public wReserved3 As UShort

        <FieldOffset(8)>
        Public bVal As Byte
        <FieldOffset(8)>
        Public cVal As SByte
        <FieldOffset(8)>
        Public uiVal As UShort
        <FieldOffset(8)>
        Public iVal As Short
        <FieldOffset(8)>
        Public uintVal As UInteger
        <FieldOffset(8)>
        Public intVal As Integer
        <FieldOffset(8)>
        Public ulVal As ULong
        <FieldOffset(8)>
        Public lVal As Long
        <FieldOffset(8)>
        Public fltVal As Single
        <FieldOffset(8)>
        Public dblVal As Double
        <FieldOffset(8)>
        Public boolVal As Short
        <FieldOffset(8)>
        Public pclsidVal As IntPtr ' GUID ID pointer
        <FieldOffset(8)>
        Public pszVal As IntPtr ' Ansi string pointer
        <FieldOffset(8)>
        Public pwszVal As IntPtr ' Unicode string pointer
        <FieldOffset(8)>
        Public punkVal As IntPtr ' punkVal (interface pointer)
        <FieldOffset(8)>
        Public ca As PROPARRAY
        <FieldOffset(8)>
        Public filetime As ComTypes.FILETIME
    End Structure

    <StructLayout(LayoutKind.Sequential, Pack:=4)>
    Public Structure PROPERTYKEY
        Private ReadOnly _fmtid As Guid
        Private ReadOnly _pid As UInteger

        Public Sub New(fmtid As Guid, pid As UInteger)
            _fmtid = fmtid
            _pid = pid
        End Sub

        Public Shared ReadOnly PKEY_ItemNameDisplay As PROPERTYKEY = New PROPERTYKEY(New Guid("B725F130-47EF-101A-A5F1-02608C9EEBAC"), 10)
        Public Shared ReadOnly PKEY_FileVersion As PROPERTYKEY = New PROPERTYKEY(New Guid("0CEF7D53-FA64-11D1-A203-0000F81FEDEE"), 4)
    End Structure

    Friend Class GlobalTools
        Public Shared Sub SafeRelease(Of T As Class)(ByRef comObject As T)
            Dim tObject = comObject
            comObject = Nothing
            If tObject IsNot Nothing Then
                If Marshal.IsComObject(tObject) Then Marshal.ReleaseComObject(tObject)
            End If
        End Sub

        Public Const DELETE As Long = &H10000L
        Public Const READ_CONTROL As Long = &H20000L
        Public Const WRITE_DAC As Long = &H40000L
        Public Const WRITE_OWNER As Long = &H80000L
        Public Const SYNCHRONIZE As Long = &H100000L

        Public Const GENERIC_READ As Long = &H80000000L
        Public Const GENERIC_WRITE As Long = &H40000000L
        Public Const GENERIC_EXECUTE As Long = &H20000000L
        Public Const GENERIC_ALL As Long = &H10000000L

        Public Enum STGC As Integer
            STGC_DEFAULT = 0
            STGC_OVERWRITE = 1
            STGC_ONLYIFCURRENT = 2
            STGC_DANGEROUSLYCOMMITMERELYTODISKCACHE = 4
            STGC_CONSOLIDATE = 8
        End Enum

        <DllImport("User32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
        Public Shared Function GetDpiForWindow(hwnd As IntPtr) As UInteger
        End Function

        Public Shared Function HIWORD(n As Integer) As Integer
            Return (n >> 16) And &HFFFF
        End Function

        Public Shared Function LOWORD(n As Integer) As Integer
            Return n And &HFFFF
        End Function
    End Class
End Namespace
