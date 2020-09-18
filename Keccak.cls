VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Keccak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
            (MyDest As Any, MySource As Any, ByVal MySize As Long)
#Else
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (Destination As Any, Source As Any, ByVal Length As LongPtr)
#End If

Private mlPadding As Long
Private mlRate As Long
Private mlMode As Long
Private maP2(), maRClo(), maRChi()

Property Let Padding(NewValue As Long)
    Select Case NewValue
        Case 1, 2: mlPadding = NewValue
        Case Else
            Err.Raise vbObjectError + 1111, , "The value not aceptable." & vbNewLine & _
                                              "Valid values for Padding:" & vbNewLine & _
                                              "  SHA-3  - 1" & vbNewLine & _
                                              "  Keccak - 2"
    End Select
End Property

Property Get Padding() As Long

    Padding = mlPadding
End Property

Property Let Mode(NewValue As Long)

    Select Case NewValue
        Case 224: mlMode = NewValue / 8: mlRate = 1152 / 8
        Case 256: mlMode = NewValue / 8: mlRate = 1088 / 8
        Case 384: mlMode = NewValue / 8: mlRate = 832 / 8
        Case 512: mlMode = NewValue / 8: mlRate = 576 / 8
        Case Else
            Err.Raise vbObjectError + 1112, , "The value not aceptable." & vbNewLine & _
                                              "Valid values for Mode:" & vbNewLine & _
                                              "  SHA3-224  - 224" & vbNewLine & _
                                              "  SHA3-256  - 256" & vbNewLine & _
                                              "  SHA3-384  - 384" & vbNewLine & _
                                              "  SHA3-512  - 512"
    End Select
End Property

Property Get Mode() As Long

    Mode = mlMode * 8
End Property

Function ComputeHash(Bytes() As Byte) As Byte()
    
    Dim lPad As Long
    Dim aState(0 To 4, 0 To 4, 0 To 1) As Long, aC(0 To 4, 0 To 1) As Long
    Dim lI As Long, lJ As Long, lK As Long
    Dim lHi As Long, lLo As Long, lRlo As Long, lRhi As Long
    Dim lX As Long, lY As Long, lXa As Long, lXb As Long, lYa As Long, lYb As Long
    ReDim aResult(0 To mlMode - 1) As Byte
    Dim lOut As Long

    lPad = mlRate - (UBound(Bytes) + 1) Mod mlRate
    If lPad = 1 Then
        ReDim Preserve Bytes(0, UBound(Bytes) + 1) As Byte
        Bytes(UBound(Bytes)) = Choose(mlPadding, 134, 129)
    Else
        ReDim Preserve Bytes(0 To UBound(Bytes) + lPad) As Byte
        Bytes(UBound(Bytes) + 1 - lPad) = Choose(mlPadding, 6, 1)
        Bytes(UBound(Bytes)) = 128
    End If
    For lI = 0 To UBound(Bytes) Step mlRate
        For lJ = 0 To mlRate / 8 - 1
            CopyMemory lLo, Bytes(lI + lJ * 8), 4
            CopyMemory lHi, Bytes(lI + lJ * 8 + 4), 4
            lX = lJ Mod 5
            lY = lJ \ 5
            aState(lX, lY, 0) = aState(lX, lY, 0) Xor lLo
            aState(lX, lY, 1) = aState(lX, lY, 1) Xor lHi
        Next lJ
        For lJ = 0 To 23
            For lX = 0 To 4
                aC(lX, 0) = aState(lX, 0, 0) Xor aState(lX, 1, 0) Xor aState(lX, 2, 0) Xor _
                            aState(lX, 3, 0) Xor aState(lX, 4, 0)
                aC(lX, 1) = aState(lX, 0, 1) Xor aState(lX, 1, 1) Xor aState(lX, 2, 1) Xor _
                            aState(lX, 3, 1) Xor aState(lX, 4, 1)
            Next lX
            For lX = 0 To 4
                lXa = (lX + 1) Mod 5 '1,2,3,4,0
                lXb = (lX + 4) Mod 5 '4,0,1,2,3
                RotateLeft64 aC(lXa, 0), aC(lXa, 1), 1, lRlo, lRhi
                For lY = 0 To 4
                    aState(lX, lY, 0) = aState(lX, lY, 0) Xor aC(lXb, 0) Xor lRlo
                    aState(lX, lY, 1) = aState(lX, lY, 1) Xor aC(lXb, 1) Xor lRhi
                Next lY
            Next lX
            lXa = 1
            lYa = 0
            lLo = aState(lXa, lYa, 0)
            lHi = aState(lXa, lYa, 1)
            For lK = 0 To 23
                RotateLeft64 lLo, lHi, ((lK + 1) * (lK + 2) / 2) Mod 64, lRlo, lRhi
                lXb = lYa
                lYb = (2 * lXa + 3 * lYa) Mod 5
                lLo = aState(lXb, lYb, 0)
                lHi = aState(lXb, lYb, 1)
                aState(lXb, lYb, 0) = lRlo
                aState(lXb, lYb, 1) = lRhi
                lXa = lXb
                lYa = lYb
            Next lK
            For lY = 0 To 4
                Erase aC
                For lX = 0 To 4
                    aC(lX, 0) = aState(lX, lY, 0)
                    aC(lX, 1) = aState(lX, lY, 1)
                Next lX
                For lX = 0 To 4
                    aState(lX, lY, 0) = aC(lX, 0) Xor (Not (aC((lX + 1) Mod 5, 0)) And aC((lX + 2) Mod 5, 0))
                    aState(lX, lY, 1) = aC(lX, 1) Xor (Not (aC((lX + 1) Mod 5, 1)) And aC((lX + 2) Mod 5, 1))
                Next lX
            Next lY
            aState(0, 0, 0) = aState(0, 0, 0) Xor maRClo(lJ)
            aState(0, 0, 1) = aState(0, 0, 1) Xor maRChi(lJ)
        Next lJ
    Next lI
    Do
        lX = (lOut / 8) Mod 5
        lY = (lOut / 8) \ 5
        CopyMemory aResult(lOut), aState(lX, lY, 0), 4
        lOut = lOut + 4
        If lOut = mlMode Then Exit Do
        CopyMemory aResult(lOut), aState(lX, lY, 1), 4
        lOut = lOut + 4
        If lOut = mlMode Then Exit Do
    Loop
    ComputeHash = aResult
End Function

Private Sub RotateLeft64(ByVal Lo As Long, ByVal Hi As Long, ByVal Shifts As Byte, Rlo As Long, Rhi As Long)

    If Shifts > 32 Then
        Shifts = Shifts - 32
        Rlo = ShiftLeft(Hi, Shifts) Or ShiftRight(Lo, 32 - Shifts)
        Rhi = ShiftLeft(Lo, Shifts) Or ShiftRight(Hi, 32 - Shifts)
    Else
        Rlo = ShiftLeft(Lo, Shifts) Or ShiftRight(Hi, 32 - Shifts)
        Rhi = ShiftLeft(Hi, Shifts) Or ShiftRight(Lo, 32 - Shifts)
    End If
End Sub

Private Function ShiftLeft(ByVal N As Long, ByVal Shifts As Long) As Long

    Select Case True
        Case Shifts = 0: ShiftLeft = N
        Case Shifts = 31
            If N And &H1& Then ShiftLeft = &H80000000 Else ShiftLeft = &H0&
        Case Shifts = 32: ShiftLeft = 0
        Case N And maP2(31 - Shifts)
            ShiftLeft = (N And maP2(31 - Shifts) - 1) * maP2(Shifts) Or &H80000000
        Case Else
            ShiftLeft = (N And maP2(31 - Shifts) - 1) * maP2(Shifts)
    End Select
End Function

Private Function ShiftRight(ByVal N As Long, ByVal Shifts As Long) As Long

    Select Case True
        Case Shifts = 0: ShiftRight = N
        Case Shifts = 31: If N And &H80000000 Then ShiftRight = 1 Else ShiftRight = 0
        Case Shifts = 32: ShiftRight = 0
        Case N And &H80000000
            ShiftRight = (N And -maP2(Shifts) And &H7FFFFFFF) \ maP2(Shifts) Or maP2(31 - Shifts)
        Case Else
            ShiftRight = (N And -maP2(Shifts)) \ maP2(Shifts)
    End Select
End Function

Private Sub Class_Initialize()

    maP2() = Array(&H1&, &H2&, &H4&, &H8&, &H10&, &H20&, &H40&, &H80&, _
        &H100&, &H200&, &H400&, &H800&, &H1000&, &H2000&, &H4000&, &H8000&, _
        &H10000, &H20000, &H40000, &H80000, &H100000, &H200000, &H400000, &H800000, _
        &H1000000, &H2000000, &H4000000, &H8000000, &H10000000, &H20000000, &H40000000)
    maRClo = Array(&H1&, &H8082&, &H808A&, &H80008000, &H808B&, &H80000001, _
        &H80008081, &H8009&, &H8A&, &H88&, &H80008009, &H8000000A, _
        &H8000808B, &H8B&, &H8089&, &H8003&, &H8002&, &H80&, _
        &H800A&, &H8000000A, &H80008081, &H8080&, &H80000001, &H80008008)
    maRChi = Array(&H0&, &H0&, &H80000000, &H80000000, &H0&, &H0&, _
        &H80000000, &H80000000, &H0&, &H0&, &H0&, &H0&, _
        &H0&, &H80000000, &H80000000, &H80000000, &H80000000, &H80000000, _
        &H0&, &H80000000, &H80000000, &H80000000, &H0&, &H8000000)
    Padding = 1
    Mode = 512
End Sub
