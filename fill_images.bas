Attribute VB_Name = "Module1"
Option Explicit
 
'BMPファイルの仕様
'18-21byteに横の画素数
'22-25byteに縦の画素数
Const WIDTH_POS As Long = 18
Const HEIGHT_POS As Long = 22
 
Sub Main()
    Dim IMG_CNT As Long '画像枚数
    IMG_CNT = 1383
    Dim IMG_H As Long '画像の高さ
    IMG_H = 180
     
    Dim myCnt As Long
    Dim next_row As Long
    next_row = 0
     
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
     
    For myCnt = 0 To IMG_CNT
        Dim str As String
        str = CStr(myCnt)
        Call ReadBMP(str, next_row)
        next_row = next_row + IMG_H
        Application.StatusBar = "Processing " & myCnt & " row"
        If myCnt Mod 5 = 0 Then DoEvents
         
    Next myCnt
     
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
 
 
Function ReadBMP(filenum As String, next_row As Long)
    Dim openFileName As String      '開くファイル名
    Dim a() As Byte                 'Byte列読み込み
    Dim File_Size As Long           '読み込むファイルのサイズ
    Dim Image_Width_Pixel As Long   '画像の横のPixel数
    Dim Image_Height_Pixel As Long  '画像の縦のPixel数
    Dim Line_Width_Size As Long     '横ラインのByte数
    Dim Line_Last_Size As Long      '横ラインの最後につけられたByte数
    Dim Image_Data_Pos As Long      'イメージデータの開始位置
    Dim image() As Byte             '画像の配列
     
    Application.ErrorCheckingOptions.BackgroundChecking = False
    ChDir ThisWorkbook.Path & "\"
    openFileName = "D:\excel_anime\video\" + filenum + ".bmp"
    If openFileName = "False" Then
        MsgBox "BMPファイルを選択してください"
        Exit Function
    End If
    Open openFileName For Binary As #1
    File_Size = LOF(1)
    ReDim a(File_Size)
    Get #1, , a()
    Close #1
     
    Image_Width_Pixel = myHex2Dec(a(), WIDTH_POS, WIDTH_POS + 3)
    Image_Height_Pixel = myHex2Dec(a(), HEIGHT_POS, HEIGHT_POS + 3)
    Line_Width_Size = myCalcLineSize(Image_Width_Pixel)
    Line_Last_Size = Line_Width_Size - Image_Width_Pixel * 3
    Image_Data_Pos = myHex2Dec(a(), 10, 13)
    ReDim image(Image_Width_Pixel - 1, Image_Height_Pixel - 1, 3 - 1)
     
    Call WriteImage2Array(image(), a(), File_Size)
     
    Call ChangeColumnWidth(image(), next_row)
    Call WriteArray2Cells(image(), next_row)
     
End Function
 
 
'16To10変換 (Byte配列, 最初のByte位置, 終わりのByte位置)
'   複数バイトの16進数を10進数に変換する
Function myHex2Dec(a() As Byte, First As Long, Last As Long) As Long
    Dim i As Long
    Dim str As String
    str = ""
     
    For i = Last To First Step -1
        str = str & Right("00" & Hex(a(i)), 2)
    Next
     
    myHex2Dec = CInt("&H" & str)
End Function
 
'BMP画像の1行分のbyte数を計算
'BMP画像は1行のbyte数が4の倍数になるよう、行末に架空のbyteを置いてある
Function myCalcLineSize(Width_Pixel As Long) As Long
    Dim Width_Byte As Long
    Width_Byte = Width_Pixel * 3
     
    If Width_Byte Mod 4 <> 0 Then
        Width_Byte = Width_Byte + (4 - (Width_Byte Mod 4))
    End If
     
    myCalcLineSize = Width_Byte
End Function
 
'画像の配列をセルに塗る
Function WriteArray2Cells(a() As Byte, next_row As Long)
    Dim r As Long
    Dim c As Long
    Dim color As Long
    Dim rMax As Long
    Dim cMax As Long
    rMax = UBound(a, 2)
    cMax = UBound(a, 1)
    For r = 0 To rMax
        For c = 0 To cMax
            Cells(r + 1 + next_row, c + 1).Interior.color = RGB(a(c, r, 0), a(c, r, 1), a(c, r, 2))
        Next
    Next
End Function
 
'byte列(a())を3次元配列(image(x,y,color))にいれる
Function WriteImage2Array(image() As Byte, a() As Byte, fileSize As Long)
    Dim r As Long
    Dim c As Long
    Dim color As Long
    Dim rMax As Long
    Dim cMax As Long
    Dim i As Long
    rMax = UBound(image, 2)
    cMax = UBound(image, 1)
    i = fileSize - 1
 
    For r = 0 To rMax
        For c = cMax To 0 Step -1
            For color = 0 To 2
                image(c, r, color) = a(i)
                i = i - 1
            Next
        Next
    Next
         
End Function
 
'配列の長さ分の行と列の幅を2pixelにする
'height 1pix = 0.75, width 1pix = 0.118
Function ChangeColumnWidth(image() As Byte, next_row As Long)
    Dim r As Long
    Dim c As Long
    r = UBound(image, 2) + 1
    c = UBound(image, 1) + 1
 
    Range(Columns(1), Columns(c)).ColumnWidth = 0.1
    Range(Rows(1), Rows(r + next_row)).RowHeight = 0.75
End Function
