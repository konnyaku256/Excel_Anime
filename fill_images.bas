Attribute VB_Name = "Module11"
Option Explicit

'BMP�t�@�C���̎d�l
'18-21byte�ɉ��̉�f��
'22-25byte�ɏc�̉�f��
Const WIDTH_POS As Long = 18
Const HEIGHT_POS As Long = 22

Sub Main()
    Dim IMG_CNT As Long '�摜����
    IMG_CNT = 1383
    Dim IMG_H As Long '�摜�̍���
    IMG_H = 180
    
    Dim myCnt As Long
    Dim next_row As Long
    next_row = 0
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    For myCnt = 0 To IMG_CNT
        Dim str As String
        str = CStr(myCnt)
        'Application.ScreenUpdating = False
        Call ReadBMP(str, next_row)
        next_row = next_row + IMG_H
        'If myCnt Mod 10 = 0 Then DoEvents
        'Application.ScreenUpdating = True
        Application.StatusBar = "Processing " & myCnt & " row"
        If myCnt Mod 5 = 0 Then DoEvents
        
    Next myCnt
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub


Function ReadBMP(filenum As String, next_row As Long)
    Dim openFileName As String      '�J���t�@�C����
    Dim a() As Byte                 'Byte��ǂݍ���
    Dim File_Size As Long           '�ǂݍ��ރt�@�C���̃T�C�Y
    Dim Image_Width_Pixel As Long   '�摜�̉���Pixel��
    Dim Image_Height_Pixel As Long  '�摜�̏c��Pixel��
    Dim Line_Width_Size As Long     '�����C����Byte��
    Dim Line_Last_Size As Long      '�����C���̍Ō�ɂ���ꂽByte��
    Dim Image_Data_Pos As Long      '�C���[�W�f�[�^�̊J�n�ʒu
    Dim image() As Byte             '�摜�̔z��
    
    Application.ErrorCheckingOptions.BackgroundChecking = False
    ChDir ThisWorkbook.Path & "\"
    openFileName = "D:\excel_anime\video\" + filenum + ".bmp"
    If openFileName = "False" Then
        MsgBox "BMP�t�@�C����I�����Ă�������"
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
    
    'Call WriteImage2Jpeg(image())
End Function


'16To10�ϊ� (Byte�z��, �ŏ���Byte�ʒu, �I����Byte�ʒu)
'   �����o�C�g��16�i����10�i���ɕϊ�����
Function myHex2Dec(a() As Byte, First As Long, Last As Long) As Long
    Dim i As Long
    Dim str As String
    str = ""
    
    For i = Last To First Step -1
        str = str & Right("00" & Hex(a(i)), 2)
    Next
    
    myHex2Dec = CInt("&H" & str)
End Function

'BMP�摜��1�s����byte�����v�Z
'BMP�摜��1�s��byte����4�̔{���ɂȂ�悤�A�s���ɉˋ��byte��u���Ă���
Function myCalcLineSize(Width_Pixel As Long) As Long
    Dim Width_Byte As Long
    Width_Byte = Width_Pixel * 3
    
    If Width_Byte Mod 4 <> 0 Then
        Width_Byte = Width_Byte + (4 - (Width_Byte Mod 4))
    End If
    
    myCalcLineSize = Width_Byte
End Function

'�摜�̔z����Z���ɓh��
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

'byte��(a())��3�����z��(image(x,y,color))�ɂ����
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

'�z��̒������̍s�Ɨ�̕���2pixel�ɂ���
'height 1pix = 0.75, width 1pix = 0.118
Function ChangeColumnWidth(image() As Byte, next_row As Long)
    Dim r As Long
    Dim c As Long
    r = UBound(image, 2) + 1
    c = UBound(image, 1) + 1

    Range(Columns(1), Columns(c)).ColumnWidth = 0.1
    Range(Rows(1), Rows(r + next_row)).RowHeight = 0.75
End Function

Function WriteImage2Jpeg(image() As Byte)
    Dim r As Long
    Dim c As Long
    r = UBound(image, 2) + 1
    c = UBound(image, 1) + 1
    Dim rc As Long
    
    Range(Cells(1, 1), Cells(r, c)).CopyPicture
        
    rc = Shell("mspaint", vbNormalFocus)
    Application.Wait Now + TimeValue("00:00:01")
    SendKeys "^v", True
    SendKeys "^s", True
    
End Function
