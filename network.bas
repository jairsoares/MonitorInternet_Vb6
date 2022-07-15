Attribute VB_Name = "network"
'----------------------------------------------------------
'Constantes do PING
'----------------------------------------------------------

'Icmp constants converted from
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wmisdk/wmi/win32_pingstatus.asp

Private Const ICMP_SUCCESS As Long = 0
Private Const ICMP_STATUS_BUFFER_TO_SMALL = 11001                   'Buffer Too Small
Private Const ICMP_STATUS_DESTINATION_NET_UNREACH = 11002           'Destination Net Unreachable
Private Const ICMP_STATUS_DESTINATION_HOST_UNREACH = 11003          'Destination Host Unreachable
Private Const ICMP_STATUS_DESTINATION_PROTOCOL_UNREACH = 11004      'Destination Protocol Unreachable
Private Const ICMP_STATUS_DESTINATION_PORT_UNREACH = 11005          'Destination Port Unreachable
Private Const ICMP_STATUS_NO_RESOURCE = 11006                       'No Resources
Private Const ICMP_STATUS_BAD_OPTION = 11007                        'Bad Option
Private Const ICMP_STATUS_HARDWARE_ERROR = 11008                    'Hardware Error
Private Const ICMP_STATUS_LARGE_PACKET = 11009                      'Packet Too Big
Private Const ICMP_STATUS_REQUEST_TIMED_OUT = 11010                 'Request Timed Out
Private Const ICMP_STATUS_BAD_REQUEST = 11011                       'Bad Request
Private Const ICMP_STATUS_BAD_ROUTE = 11012                         'Bad Route
Private Const ICMP_STATUS_TTL_EXPIRED_TRANSIT = 11013               'TimeToLive Expired Transit
Private Const ICMP_STATUS_TTL_EXPIRED_REASSEMBLY = 11014            'TimeToLive Expired Reassembly
Private Const ICMP_STATUS_PARAMETER = 11015                         'Parameter Problem
Private Const ICMP_STATUS_SOURCE_QUENCH = 11016                     'Source Quench
Private Const ICMP_STATUS_OPTION_TOO_BIG = 11017                    'Option Too Big
Private Const ICMP_STATUS_BAD_DESTINATION = 11018                   'Bad Destination
Private Const ICMP_STATUS_NEGOTIATING_IPSEC = 11032                 'Negotiating IPSEC
Private Const ICMP_STATUS_GENERAL_FAILURE = 11050                   'General Failure

Public Const WINSOCK_ERROR = "Windows Sockets not responding correctly."
Public Const INADDR_NONE As Long = &HFFFFFFFF
Public Const WSA_SUCCESS = 0
Public Const WS_VERSION_REQD As Long = &H101

Global tictac As Integer

'Clean up sockets.
'http://support.microsoft.com/default.aspx?scid=kb;EN-US;q154512

Private Declare Function WSACleanup Lib "wsock32.dll" () As Long

'Open the socket connection.
'http://support.microsoft.com/default.aspx?scid=kb;EN-US;q154512
Private Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequired As Long, lpWSAData As WSAData) As Long

'Create a handle on which Internet Control Message Protocol (ICMP) requests can be issued.
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wcesdkr/htm/_wcesdk_icmpcreatefile.asp
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long

'Convert a string that contains an (Ipv4) Internet Protocol dotted address into a correct address.
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winsock/wsapiref_4esy.asp
Private Declare Function inet_addr Lib "wsock32.dll" (ByVal CP As String) As Long

'Close an Internet Control Message Protocol (ICMP) handle that IcmpCreateFile opens.
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wcesdkr/htm/_wcesdk_icmpclosehandle.asp

Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long

Declare Function gethostbyname Lib "wsock32" (ByVal HostName As String) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (xDest As Any, xSource As Any, ByVal nbytes As Long)

'Information about the Windows Sockets implementation
'http://support.microsoft.com/default.aspx?scid=kb;EN-US;q154512
Private Type WSAData
   wVersion As Integer
   wHighVersion As Integer
   szDescription(0 To 256) As Byte
   szSystemStatus(0 To 128) As Byte
   iMaxSockets As Long
   iMaxUdpDg As Long
   lpVendorInfo As Long
End Type

'Send an Internet Control Message Protocol (ICMP) echo request, and then return one or more replies.
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wcetcpip/htm/cerefIcmpSendEcho.asp
Private Declare Function IcmpSendEcho Lib "icmp.dll" _
   (ByVal IcmpHandle As Long, _
    ByVal DestinationAddress As Long, _
    ByVal RequestData As String, _
    ByVal RequestSize As Long, _
    ByVal RequestOptions As Long, _
    ReplyBuffer As ICMP_ECHO_REPLY, _
    ByVal ReplySize As Long, _
    ByVal Timeout As Long) As Long
 
'This structure describes the options that will be included in the header of an IP packet.
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wcetcpip/htm/cerefIP_OPTION_INFORMATION.asp
Private Type IP_OPTION_INFORMATION
   Ttl             As Byte
   Tos             As Byte
   Flags           As Byte
   OptionsSize     As Byte
   OptionsData     As Long
End Type

'This structure describes the data that is returned in response to an echo request.
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wcesdkr/htm/_wcesdk_icmp_echo_reply.asp
Public Type ICMP_ECHO_REPLY
   Address         As Long
   STATUS          As Long
   RoundTripTime   As Long
   DataSize        As Long
   Reserved        As Integer
   ptrData         As Long
   Options        As IP_OPTION_INFORMATION
   DATA            As String * 250
End Type


' ---------------------------------------------------------
Private Const TH32CS_SNAPPROCESS As Long = 2
Private Const MAX_PATH As Long = 260
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, _
                                                                  ByVal lProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, _
                                                        typProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, _
                                                       typProcess As PROCESSENTRY32) As Long
Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)

Private Declare Function GetIpAddrTable_API Lib "IpHlpApi" Alias "GetIpAddrTable" (pIPAddrTable As Any, pdwSize As Long, ByVal bOrder As Long) As Long

Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Function InternetGetConnectedState Lib "wininet" _
(ByRef dwFlags As Long, _
ByVal dwReserved As Long) As Long

Private Declare Function apiSerialNumber Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Private Declare Function URLDownloadToFile Lib "urlmon" _
   Alias "URLDownloadToFileA" _
  (ByVal pCaller As Long, _
   ByVal szURL As String, _
   ByVal szFileName As String, _
   ByVal dwReserved As Long, _
   ByVal lpfnCB As Long) As Long
   
   
Private Declare Function DeleteUrlCacheEntry Lib "wininet.dll" _
   Alias "DeleteUrlCacheEntryA" _
  (ByVal lpszUrlName As String) As Long
  
Private Const ERROR_SUCCESS As Long = 0
Private Const BINDF_GETNEWESTVERSION As Long = &H10
Private Const INTERNET_FLAG_RELOAD As Long = &H80000000

Private Const TOKEN_QUERY As Long = &H8

Private Enum TOKEN_INFORMATION_CLASS 'Stripped to essentials.
    TokenElevationType = 18
    TokenElevation = 20
End Enum

Private Enum TOKEN_ELEVATION_TYPE
    TokenElevationTypeDefault = 1
    TokenElevationTypeFull
    TokenElevationTypeLimited
End Enum

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Private Declare Function OpenProcessToken Lib "advapi32" ( _
    ByVal ProcessHandle As Long, _
    ByVal DesiredAccess As Long, _
    ByRef TokenHandle As Long) As Long

Private Declare Function GetTokenInformation Lib "advapi32" ( _
    ByVal TokenHandle As Long, _
    ByVal TokenInformationClass As Long, _
    ByRef TokenInformation As Any, _
    ByVal TokenInformationLength As Long, _
    ByRef ReturnLength As Long) As Long

Public Const ENCRYPT = 1, DECRYPT = 2




Public Function VerificaInternet() As Long
'By JPaulo ® Maximo Access
Dim strResultado As Long
VerificaInternet = InternetGetConnectedState(strResultado, 0)
End Function



' Rotina de PING pro Servidor do HiZAP
'-------------------------------------------------------------------------
Public Function pingWS() As Boolean
Dim ret As Boolean
Dim Reply As ICMP_ECHO_REPLY
Dim lngSuccess As Long
Dim strIpAddress As String

   
    Screen.MousePointer = vbHourglass
    DoEvents
    Debug.Print "****************************************"

    'Get the sockets ready.
    If SocketsInitialize() Then
        'Address to ping
        
        Select Case tictac
               Case 0
                    strIpAddress = "187.17.111.101"  ' Autotecweb.com
                    tictac = 1
               Case 1
                    strIpAddress = "177.12.168.51"  ' HiZap
                    tictac = 0
        End Select
        
        Debug.Print "TicTAc", "Pingando para ->", strIpAddress
        
        'Ping the IP that is passing the address and get a reply.
        lngSuccess = Ping(strIpAddress, 5000, Reply)
        If Ping(strIpAddress, 5000, Reply) = 0 Then
           ret = True
        Else
          ' MsgBox "Atenção, informe a CSS-Sistemas sobre o seguinte status: " & vbNewLine & _
          '        "Status do Servidor: " & EvaluatePingResponse(lngSuccess), vbCritical
           registraLogErros 0, "Retorno do Ping: " & EvaluatePingResponse(lngSuccess) & ".", "Retorno do Ping"
           ret = False
        End If
        'Display the results.
        Debug.Print "Address to Ping: " & strIpAddress
        Debug.Print "Raw ICMP code: " & lngSuccess
        Debug.Print "Ping Response Message : " & EvaluatePingResponse(lngSuccess)
        Debug.Print "Time : " & Reply.RoundTripTime & " ms"

        'Clean up the sockets.
        SocketsCleanup
    Else
        'Winsock error failure, initializing the sockets.
        Debug.Print WINSOCK_ERROR
    End If

    Screen.MousePointer = vbDefault


    pingWS = ret
End Function





'-- Ping a string representation of an IP address.
' -- Return a reply.
' -- Return long code.
Public Function Ping(ByVal sAddress As String, ByVal time_out As Long, Reply As ICMP_ECHO_REPLY) As Long

Dim hIcmp As Long
Dim lAddress As Long
Dim lTimeOut As Long
Dim StringToSend As String

'Short string of data to send
StringToSend = "hello"

'ICMP (ping) timeout
lTimeOut = time_out 'ms

'Convert string address to a long representation.
lAddress = inet_addr(sAddress)

If (lAddress <> -1) And (lAddress <> 0) Then
        
    'Create the handle for ICMP requests.
    hIcmp = IcmpCreateFile()
    
    If hIcmp Then
        'Ping the destination IP address.
        Call IcmpSendEcho(hIcmp, lAddress, StringToSend, Len(StringToSend), 0, Reply, Len(Reply), lTimeOut)

        'Reply status
        Ping = Reply.STATUS
        
        'Close the Icmp handle.
        IcmpCloseHandle hIcmp
    Else
        Debug.Print "failure opening icmp handle."
        Ping = -1
    End If
Else
    Ping = -1
End If

End Function

'Clean up the sockets.
'http://support.microsoft.com/default.aspx?scid=kb;EN-US;q154512
Public Sub SocketsCleanup()
   
   WSACleanup
    
End Sub

'Get the sockets ready.
'http://support.microsoft.com/default.aspx?scid=kb;EN-US;q154512
Public Function SocketsInitialize() As Boolean

   Dim WSAD As WSAData

   SocketsInitialize = WSAStartup(WS_VERSION_REQD, WSAD) = ICMP_SUCCESS

End Function

'Convert the ping response to a message that you can read easily from constants.
'For more information about these constants, visit the following Microsoft Web site:
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wmisdk/wmi/win32_pingstatus.asp

Public Function EvaluatePingResponse(PingResponse As Long) As String

  Select Case PingResponse
    
  'Success
  Case ICMP_SUCCESS: EvaluatePingResponse = "Success!"
            
  'Some error occurred
  Case ICMP_STATUS_BUFFER_TO_SMALL:    EvaluatePingResponse = "Buffer Too Small"
  Case ICMP_STATUS_DESTINATION_NET_UNREACH: EvaluatePingResponse = "Destination Net Unreachable"
  Case ICMP_STATUS_DESTINATION_HOST_UNREACH: EvaluatePingResponse = "Destination Host Unreachable"
  Case ICMP_STATUS_DESTINATION_PROTOCOL_UNREACH: EvaluatePingResponse = "Destination Protocol Unreachable"
  Case ICMP_STATUS_DESTINATION_PORT_UNREACH: EvaluatePingResponse = "Destination Port Unreachable"
  Case ICMP_STATUS_NO_RESOURCE: EvaluatePingResponse = "No Resources"
  Case ICMP_STATUS_BAD_OPTION: EvaluatePingResponse = "Bad Option"
  Case ICMP_STATUS_HARDWARE_ERROR: EvaluatePingResponse = "Hardware Error"
  Case ICMP_STATUS_LARGE_PACKET: EvaluatePingResponse = "Packet Too Big"
  Case ICMP_STATUS_REQUEST_TIMED_OUT: EvaluatePingResponse = "Request Timed Out"
  Case ICMP_STATUS_BAD_REQUEST: EvaluatePingResponse = "Bad Request"
  Case ICMP_STATUS_BAD_ROUTE: EvaluatePingResponse = "Bad Route"
  Case ICMP_STATUS_TTL_EXPIRED_TRANSIT: EvaluatePingResponse = "TimeToLive Expired Transit"
  Case ICMP_STATUS_TTL_EXPIRED_REASSEMBLY: EvaluatePingResponse = "TimeToLive Expired Reassembly"
  Case ICMP_STATUS_PARAMETER: EvaluatePingResponse = "Parameter Problem"
  Case ICMP_STATUS_SOURCE_QUENCH: EvaluatePingResponse = "Source Quench"
  Case ICMP_STATUS_OPTION_TOO_BIG: EvaluatePingResponse = "Option Too Big"
  Case ICMP_STATUS_BAD_DESTINATION: EvaluatePingResponse = "Bad Destination"
  Case ICMP_STATUS_NEGOTIATING_IPSEC: EvaluatePingResponse = "Negotiating IPSEC"
  Case ICMP_STATUS_GENERAL_FAILURE: EvaluatePingResponse = "General Failure"
            
  'Unknown error occurred
  Case Else: EvaluatePingResponse = "Unknown Response"
        
  End Select

End Function

