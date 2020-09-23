VERSION 5.00
Begin VB.UserControl Encrypt 
   CanGetFocus     =   0   'False
   ClientHeight    =   960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1155
   InvisibleAtRuntime=   -1  'True
   Picture         =   "Encrypt.ctx":0000
   ScaleHeight     =   960
   ScaleWidth      =   1155
End
Attribute VB_Name = "Encrypt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim Matrix(9, 9) As String * 1
Private Sub UserControl_Initialize()
    'Set the chr. table values
    FillMatrix
End Sub

Public Function Encode(Txt As String) As String

    Dim I As Integer, TLen As Integer, tmpX As String
    Dim Part1, Part2
    TLen = Len(Txt)

    'Level 1
    For I = 1 To TLen
        tmpX = tmpX & InMatrix(Mid(Txt, I, 1))
    Next I


    'Level 2
    For I = 1 To Len(tmpX) Step 2
        Part1 = Part1 + Mid(tmpX, I, 1)
    Next I
    For I = 2 To Len(tmpX) Step 2
        Part2 = Part2 + Mid(tmpX, I, 1)
    Next I

    Encode = Part1 & Part2
End Function

Public Function Decode(Txt As String) As String
    Dim I As Integer, TLen As Integer, tmpX As String, Jx, Jy
    Dim Part1, Part2
    TLen = Len(Txt)
    
    Part1 = Left(Txt, TLen / 2)
    Part2 = Right(Txt, TLen / 2)
    
    Txt = ""
    For I = 1 To TLen / 2
        Txt = Txt & Mid(Part1, I, 1) & Mid(Part2, I, 1)
    Next I
    
    For I = 1 To TLen Step 2
        Jx = Mid(Txt, I, 1)
        Jy = Mid(Txt, I + 1, 1)
        tmpX = tmpX & Matrix(Jx, Jy)
    Next I
    
    Decode = tmpX

End Function
Private Function FillMatrix()
    Dim I1, I2, J
    'Chrz
    Matrix(0, 0) = "-"
    Matrix(0, 1) = " "
    Matrix(0, 2) = "_"
    Matrix(0, 3) = "."
    Matrix(0, 4) = ","
    Matrix(0, 5) = ";"
    Matrix(0, 6) = "" '!!!!!!
    Matrix(0, 7) = "'"
    Matrix(0, 8) = "("
    Matrix(0, 9) = ")"
    '*************************************
    'Small Caps
    J = 97
    For I1 = 1 To 9
        For I2 = 0 To 9
            Matrix(I1, I2) = Chr(J)
            'Debug.Print I1 & ", " & I2 & ": " & Chr(J) & J
            J = J + 1
            If J > 122 Then Exit For
        Next I2
        If J > 122 Then Exit For
    Next I1
    '*************************************
    'Cap. Caps
    Matrix(3, 6) = "A"
    Matrix(3, 7) = "B"
    Matrix(3, 8) = "C"
    Matrix(3, 9) = "D"
    J = 69
    For I1 = 4 To 9
        For I2 = 0 To 9
            Matrix(I1, I2) = Chr(J)
            'Debug.Print I1 & ", " & I2 & ": " & Chr(J) & J
            J = J + 1
            If J > 90 Then Exit For
        Next I2
        If J > 90 Then Exit For
    Next I1
    '*************************************
    'Numbers
    Matrix(6, 2) = "0"
    Matrix(6, 3) = "1"
    Matrix(6, 4) = "2"
    Matrix(6, 5) = "3"
    Matrix(6, 6) = "4"
    Matrix(6, 7) = "5"
    Matrix(6, 8) = "6"
    Matrix(6, 9) = "7"
    Matrix(7, 0) = "8"
    Matrix(7, 1) = "9"
    '*************************************
    Matrix(7, 2) = Chr(13)
    Matrix(7, 3) = Chr(10)
End Function
Private Function InMatrix(ChrX As String) As String
    For I1 = 0 To 9
        For I2 = 0 To 9
            If Matrix(I1, I2) = Left(ChrX, 1) Then InMatrix = Format(Val(I1 & I2), "00"): Exit Function
        Next I2
    Next I1
    
    
    'if not found = 99
    InMatrix = "01"
End Function



Private Sub UserControl_Resize()
    UserControl.Width = 1155
    UserControl.Height = 960
End Sub
