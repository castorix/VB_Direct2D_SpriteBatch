Imports System

Imports Direct2D
Imports System.Runtime.InteropServices
Imports GlobalStructures

Namespace Global.Sprite
    Friend Class CSprite
        Implements IDisposable
        Private m_pBitmap As ID2D1Bitmap = Nothing
        Private m_pRectSource As D2D1_RECT_U()
        Private m_pRectDest As D2D1_RECT_F()
        Private m_pColors As D2D1_COLOR_F_STRUCT()
        Private m_pTransforms As D2D1_MATRIX_3X2_F_STRUCT()
        Private m_pSpriteBatch As ID2D1SpriteBatch = Nothing
        Private m_nNbImagesX As UInteger = 1
        Private m_nNbImagesY As UInteger = 1
        Private m_nNbImages As UInteger = 1

        Public Enum HORIZONTALFLIP As Integer
            LEFT = 0
            RIGHT
            NONE
        End Enum

        Public Enum BOUNCE As Integer
            HORIZONTAL = 0
            VERTICAL
            BOTH
            NONE
        End Enum

        Public Width As Integer = 0
        Public Height As Integer = 0
        Public StartTime As Long = 0
        Public Tag As String = ""

        Private m_nCurrentIndex As UInteger = 0
        Public Property CurrentIndex As UInteger
            Get
                Return m_nCurrentIndex
            End Get
            Set(value As UInteger)
                m_nCurrentIndex = value
                If m_nCurrentIndex >= m_nNbImages Then m_nCurrentIndex = 0
            End Set
        End Property

        Private m_nStepX As Single = 1
        Public Property StepX As Single
            Get
                Return m_nStepX
            End Get
            Set(value As Single)
                m_nStepX = value
            End Set
        End Property

        Private m_nStepY As Single = 1
        Public Property StepY As Single
            Get
                Return m_nStepY
            End Get
            Set(value As Single)
                m_nStepY = value
            End Set
        End Property

        Public Property X As Single
            Get
                Return m_pRectDest(0).left
            End Get
            Set(value As Single)
                m_pRectDest(0).left = value
            End Set
        End Property

        Public Property Y As Single
            Get
                Return m_pRectDest(0).top
            End Get
            Set(value As Single)
                m_pRectDest(0).top = value
            End Set
        End Property

        Public Sub New(pDC As ID2D1DeviceContext3, pBitmap As ID2D1Bitmap, nNbImagesX As UInteger, nNbImagesY As UInteger, Optional nNbImages As UInteger = 0, Optional nStepX As Single = 1, Optional nStepY As Single = 1, Optional color As D2D1_COLOR_F = Nothing, Optional matrix As D2D1_MATRIX_3X2_F = Nothing)
            Dim hr As HRESULT = HRESULT.S_OK
            m_pBitmap = pBitmap
            m_nNbImagesX = nNbImagesX
            m_nNbImagesY = nNbImagesY
            m_nNbImages = If(nNbImages = 0, m_nNbImagesX * m_nNbImagesY, nNbImages)
            hr = pDC.CreateSpriteBatch(m_pSpriteBatch)
            Dim bmpSize As D2D1_SIZE_F = Nothing
            pBitmap.GetSize(bmpSize)

            ' Only first pRectDest used in the test...
            m_pRectSource = New D2D1_RECT_U(m_nNbImagesX * m_nNbImagesY - 1) {}
            m_pRectDest = New D2D1_RECT_F(m_nNbImagesX * m_nNbImagesY - 1) {}
            m_pTransforms = New D2D1_MATRIX_3X2_F_STRUCT(m_nNbImagesX * m_nNbImagesY - 1) {}
            For m As UInteger = 0 To m_nNbImagesX * m_nNbImagesY - 1
                If matrix IsNot Nothing Then
                    m_pTransforms(m)._11 = matrix._11
                    m_pTransforms(m)._12 = matrix._12
                    m_pTransforms(m)._21 = matrix._21
                    m_pTransforms(m)._22 = matrix._22
                    m_pTransforms(m)._31 = matrix._31
                    m_pTransforms(m)._32 = matrix._32
                End If
            Next

            StepX = nStepX
            StepY = nStepY
            Dim nWidth As Single = bmpSize.width / m_nNbImagesX
            Dim nHeight As Single = bmpSize.height / m_nNbImagesY
            Width = CInt(nWidth)
            Height = CInt(nHeight)
            Dim n = 0
            For j As UInteger = 0 To m_nNbImagesY - 1
                For i As UInteger = 0 To m_nNbImagesX - 1
                    m_pRectSource(n) = New D2D1_RECT_U(CUInt(i * bmpSize.width / m_nNbImagesX), CUInt(j * bmpSize.height / m_nNbImagesY), CUInt(i * bmpSize.width / m_nNbImagesX + CUInt(bmpSize.width) / m_nNbImagesX), CUInt(j * bmpSize.height / m_nNbImagesY + CUInt(bmpSize.height) / m_nNbImagesY))
                    m_pRectDest(n) = New D2D1_RECT_F(i * nWidth, j * nHeight, i * nWidth + nWidth, j * nHeight + nHeight)
                    n += 1
                Next
            Next
            m_pColors = New D2D1_COLOR_F_STRUCT(m_nNbImagesX * m_nNbImagesY - 1) {}
            For c As UInteger = 0 To m_nNbImagesX * m_nNbImagesY - 1
                If color IsNot Nothing Then
                    m_pColors(c).a = color.a
                    m_pColors(c).r = color.r
                    m_pColors(c).g = color.g
                    m_pColors(c).b = color.b
                End If
            Next
            hr = m_pSpriteBatch.AddSprites(m_nNbImagesX * m_nNbImagesY, m_pRectDest, m_pRectSource, If(color Is Nothing, Nothing, m_pColors), If(matrix Is Nothing, Nothing, m_pTransforms), 0, Marshal.SizeOf(GetType(D2D1_RECT_U)), 0, Marshal.SizeOf(GetType(D2D1_MATRIX_3X2_F)))
        End Sub

        Public Sub Draw(pDC As ID2D1DeviceContext3, nSpriteIndex As UInteger, nSpriteCount As UInteger, bIncrement As Boolean)
            pDC.DrawSpriteBatch(m_pSpriteBatch, nSpriteIndex, nSpriteCount, m_pBitmap)
        End Sub

        Public Sub Move(clientSize As D2D1_SIZE_F, pDC As ID2D1DeviceContext3, nHorizontalFlip As HORIZONTALFLIP, nBounce As BOUNCE)
            Dim hr As HRESULT = HRESULT.S_OK
            Dim size As D2D1_SIZE_F
            If Not clientSize.Equals(New D2D1_SIZE_F()) Then
                size = clientSize
            Else
                pDC.GetSize(size)
            End If
            Dim bmpSize As D2D1_SIZE_F = Nothing
            m_pBitmap.GetSize(bmpSize)

            Dim nWidth As Single = bmpSize.width / m_nNbImagesX
            Dim nHeight As Single = bmpSize.height / m_nNbImagesY
            If m_pTransforms(0)._11 <> 0 Then
                'nWidth *= 1/m_pTransforms[0]._11;
                size.width *= 1 / m_pTransforms(0)._11
            End If
            If m_pTransforms(0)._22 <> 0 Then
                'nHeight *= 1/m_pTransforms[0]._22;
                size.height *= 1 / m_pTransforms(0)._22
            End If

            m_pRectDest(0).right = m_pRectDest(0).left + nWidth
            m_pRectDest(0).bottom = m_pRectDest(0).top + nHeight

            ' Tests to bounce the sprite
            If nBounce = BOUNCE.BOTH Then
                If m_pRectDest(0).left >= size.width - nWidth Then
                    m_nStepX = -Math.Abs(m_nStepX)
                    m_pRectDest(0).left = size.width - nWidth
                End If
                If m_pRectDest(0).top >= size.height - nHeight Then
                    m_nStepY = -Math.Abs(m_nStepY)
                    m_pRectDest(0).top = size.height - nHeight
                End If
                If m_pRectDest(0).left <= 0 Then
                    m_nStepX = Math.Abs(m_nStepX)
                    m_pRectDest(0).left = 0
                End If
                If m_pRectDest(0).top <= 0 Then
                    m_nStepY = Math.Abs(m_nStepY)
                    m_pRectDest(0).top = 0
                End If
            ElseIf nBounce = BOUNCE.HORIZONTAL Then
                If m_pRectDest(0).left >= size.width - nWidth Then
                    m_nStepX = -Math.Abs(m_nStepX)
                    m_pRectDest(0).left = size.width - nWidth
                End If
                If m_pRectDest(0).left <= 0 Then
                    m_nStepX = Math.Abs(m_nStepX)
                    m_pRectDest(0).left = 0
                End If

                If m_pRectDest(0).top >= size.height AndAlso m_nStepY >= 0 Then
                    m_pRectDest(0).top = 0 - bmpSize.height / m_nNbImagesY
                    m_pRectDest(0).bottom = m_pRectDest(0).top + bmpSize.height / m_nNbImagesY
                End If
                If m_pRectDest(0).bottom <= 0 AndAlso m_nStepY < 0 Then
                    m_pRectDest(0).top = size.height
                    m_pRectDest(0).bottom = m_pRectDest(0).top + bmpSize.height / m_nNbImagesY
                End If
            ElseIf nBounce = BOUNCE.VERTICAL Then
                If m_pRectDest(0).top >= size.height - nHeight Then
                    m_nStepY = -Math.Abs(m_nStepY)
                    m_pRectDest(0).top = size.height - nHeight
                End If
                If m_pRectDest(0).top <= 0 Then
                    m_nStepY = Math.Abs(m_nStepY)
                    m_pRectDest(0).top = 0
                End If

                If m_pRectDest(0).left >= size.width AndAlso m_nStepX >= 0 Then
                    m_pRectDest(0).left = 0 - bmpSize.width / m_nNbImagesX
                    m_pRectDest(0).right = m_pRectDest(0).left + bmpSize.width / m_nNbImagesX
                End If
                If m_pRectDest(0).right <= 0 AndAlso m_nStepX < 0 Then
                    m_pRectDest(0).left = size.width
                    m_pRectDest(0).right = m_pRectDest(0).left + bmpSize.width / m_nNbImagesX
                End If
            End If

            If nHorizontalFlip = HORIZONTALFLIP.LEFT OrElse nHorizontalFlip = HORIZONTALFLIP.RIGHT Then
                If m_nStepX >= 0 AndAlso nHorizontalFlip = HORIZONTALFLIP.RIGHT OrElse m_nStepX < 0 AndAlso nHorizontalFlip = HORIZONTALFLIP.LEFT Then
                    Dim n = 0
                    For j As UInteger = 0 To m_nNbImagesY - 1
                        For i As UInteger = 0 To m_nNbImagesX - 1
                            m_pRectSource(n) = New D2D1_RECT_U(CUInt(i * bmpSize.width / m_nNbImagesX + CUInt(bmpSize.width) / m_nNbImagesX), CUInt(j * bmpSize.height / m_nNbImagesY), CUInt(i * bmpSize.width / m_nNbImagesX), CUInt(j * bmpSize.height / m_nNbImagesY + CUInt(bmpSize.height) / m_nNbImagesY))
                            n += 1
                        Next
                    Next
                Else
                    Dim n = 0
                    For j As UInteger = 0 To m_nNbImagesY - 1
                        For i As UInteger = 0 To m_nNbImagesX - 1
                            m_pRectSource(n) = New D2D1_RECT_U(CUInt(i * bmpSize.width / m_nNbImagesX), CUInt(j * bmpSize.height / m_nNbImagesY), CUInt(i * bmpSize.width / m_nNbImagesX + CUInt(bmpSize.width) / m_nNbImagesX), CUInt(j * bmpSize.height / m_nNbImagesY + CUInt(bmpSize.height) / m_nNbImagesY))
                            n += 1
                        Next
                    Next
                End If
            End If
            hr = m_pSpriteBatch.SetSprites(0, m_nNbImagesX * m_nNbImagesY, m_pRectDest, m_pRectSource, Nothing, Nothing, 0, Marshal.SizeOf(GetType(D2D1_RECT_U)), 0, Marshal.SizeOf(GetType(D2D1_MATRIX_3X2_F)))
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            m_pSpriteBatch.Clear()
            Marshal.ReleaseComObject(m_pSpriteBatch)
            m_pSpriteBatch = Nothing
        End Sub
    End Class
End Namespace
