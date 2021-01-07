Attribute VB_Name = "Module1"
Option Explicit

Public Const MODBUS_16_CRC_POLYNOMIAL = &HA001


Public Function PadLeft(sData As Variant, iLen As Integer, Optional sPadChar As String = " ") As String
    If Len(sData) < iLen Then
        sData = String(iLen - Len(sData), sPadChar) & sData
    Else
        sData = Mid(sData, 1, iLen)
    End If

    PadLeft = sData
End Function

Public Function PadRight(sData As Variant, iLen As Integer, Optional sPadChar As String = " ") As String
    If Len(sData) < iLen Then
        sData = sData & String(iLen - Len(sData), sPadChar)
    Else
        sData = Mid(sData, 1, iLen)
    End If
    
    PadRight = sData
End Function


Public Function MODBUS_CRC_16(sData As String, sLen As Integer) As Integer
    
    Dim l_iCRC As Integer
    Dim l_cCount As Integer, l_cSubCount As Integer, l_iRotate As Integer
    
    l_iCRC = &HFFFF
    
    For l_cCount = 0 To sLen - 1 Step 1
        l_iCRC = l_iCRC Xor Asc(Mid(sData, l_cCount + 1, 1))
        
        For l_cSubCount = 0 To 7 Step 1
            l_iRotate = l_iCRC And 1
            l_iCRC = ((l_iCRC And &HFFFE) / 2) And &H7FFF
            
            If l_iRotate > 0 Then
                l_iCRC = l_iCRC Xor MODBUS_16_CRC_POLYNOMIAL
            End If
        Next l_cSubCount
    Next l_cCount

    MODBUS_CRC_16 = l_iCRC
End Function

Public Function MODBUS_CRC_16_Received(sData() As Byte, sLen As Integer) As Integer
    
    Dim l_iCRC As Integer
    Dim l_cCount As Integer, l_cSubCount As Integer, l_iRotate As Integer
    
    l_iCRC = &HFFFF
    
    For l_cCount = 0 To sLen - 1 Step 1
        l_iCRC = l_iCRC Xor sData(l_cCount)
        
        For l_cSubCount = 0 To 7 Step 1
            l_iRotate = l_iCRC And 1
            l_iCRC = ((l_iCRC And &HFFFE) / 2) And &H7FFF
            
            If l_iRotate > 0 Then
                l_iCRC = l_iCRC Xor MODBUS_16_CRC_POLYNOMIAL
            End If
        Next l_cSubCount
    Next l_cCount
    
    MODBUS_CRC_16_Received = l_iCRC
    
    End Function
