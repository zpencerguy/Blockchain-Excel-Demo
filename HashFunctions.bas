Attribute VB_Name = "HashFunctions"
Option Explicit

Function hash1(x As String, y As Boolean)

    Dim oT As Object, oSHA1 As Object
    Dim TextToHash() As Byte
    Dim bytes() As Byte
'    'Dim sIn As String
'    'Dim bB64 As Boolean
'
'
    Set oT = CreateObject("System.Text.UTF8Encoding")
    Set oSHA1 = CreateObject("System.Security.Cryptography.SHA1Managed")


    TextToHash = oT.GetBytes_4(x)
    bytes = oSHA1.ComputeHash_2((TextToHash))

    If y = True Then
       hash1 = ConvToBase64String(bytes)
    Else
       hash1 = ConvToHexString(bytes)
    End If

    Set oT = Nothing
    Set oSHA1 = Nothing


End Function

Function ConvToBase64String(vIn As Variant) As Variant

    Dim oD As Object
      
    Set oD = CreateObject("MSXML2.DOMDocument")
      With oD
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.base64"
        .DocumentElement.nodeTypedValue = vIn
      End With
    ConvToBase64String = Replace(oD.DocumentElement.Text, vbLf, "")
    
    Set oD = Nothing

End Function

Function ConvToHexString(vIn As Variant) As Variant

    Dim oD As Object
      
    Set oD = CreateObject("MSXML2.DOMDocument")
      
      With oD
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.Hex"
        .DocumentElement.nodeTypedValue = vIn
      End With
    ConvToHexString = Replace(oD.DocumentElement.Text, vbLf, "")
    
    Set oD = Nothing

End Function



