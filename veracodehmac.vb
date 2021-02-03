Imports System.Security.Cryptography
Imports System.Text

Public Class HmacAuthHeader
    Private ReadOnly RngRandom = New RNGCryptoServiceProvider() ' was static
    Private ReadOnly GetHashAlgorithm$ = "HmacSHA256"
    Private ReadOnly GetAuthorizationScheme$ = "VERACODE-HMAC-SHA-256"
    Private ReadOnly GetRequestVersion$ = "vcode_request_version_1"
    Private ReadOnly GetTextEncoding$ = "UTF-8"
    Private ReadOnly GetNonceSize As Integer = 16

    Protected Function CurrentDateStamp() As String
        Dim numMS As Long = (DateTime.UtcNow - New DateTime(1970, 1, 1)).TotalMilliseconds
        Return numMS.ToString
    End Function

    Protected Function NewNonce(size As Integer) As Byte()
        Dim nonceBytes As Byte() = New Byte(size) {}
        RngRandom.GetBytes(nonceBytes)
        Return nonceBytes
    End Function

    Protected Function ComputeHash(data As Byte(), key As Byte()) As Byte()
        Dim mac As HMAC = HMAC.Create(GetHashAlgorithm)
        mac.Key = key
        Return mac.ComputeHash(data)
    End Function

    Protected Function CalculateDataSignature(apiKeyBytes As Byte(), nonceBytes As Byte(), dateStamp$, data$) As Byte()
        Dim kNonce As Byte() = ComputeHash(nonceBytes, apiKeyBytes)
        Dim kDate As Byte() = ComputeHash(Encoding.GetEncoding(GetTextEncoding).GetBytes(dateStamp), kNonce)
        Dim kSignature As Byte() = ComputeHash(Encoding.GetEncoding(GetTextEncoding).GetBytes(GetRequestVersion), kDate)
        Return ComputeHash(Encoding.GetEncoding(GetTextEncoding).GetBytes(data), kSignature)
    End Function

    Public Function CalculateAuthorizationHeader(apiId$, apiKey$, hostName$, uriString$, urlQueryParams$, httpMethod$) As String
        uriString += (urlQueryParams)
        Dim Data$ = $"id={apiId}&host={hostName}&url={uriString}&method={httpMethod}"
        Dim dateStamp$ = CurrentDateStamp()
        Dim nonceBytes As Byte() = NewNonce(GetNonceSize)
        Dim dataSignature As Byte() = CalculateDataSignature(FromHexBinary(apiKey), nonceBytes, dateStamp, Data)
        Dim authorizationParam$ = $"id={apiId},ts={dateStamp},nonce={ToHexBinary(nonceBytes)},sig={ToHexBinary(dataSignature)}"

        Return GetAuthorizationScheme + " " + authorizationParam
    End Function

    Public Function ToHexBinary(bytes As Byte()) As String
        Return BitConverter.ToString(bytes).Replace("-", "")
    End Function

    Public Function FromHexBinary(hexBinaryString$) As Byte()
        Dim bytes As Byte() = New Byte(hexBinaryString.Length \ 2 - 1) {}
        For i As Integer = 0 To hexBinaryString.Length - 1 Step 2
            bytes(i \ 2) = Byte.Parse(hexBinaryString(i).ToString() & hexBinaryString(i + 1).ToString(), System.Globalization.NumberStyles.HexNumber)
        Next
        Return bytes
    End Function

End Class
