Attribute VB_Name = "LocalIPAddressesFunction"

'LB> I haven't been able to get ahold of Randy Birch (no reply to my emails),
'but I believe I am allowed to provide this code, according to this note on his site:
'
'This means you, as an application developer or hobbyist, can :
'
'» use any code here in any application you develop, either personally or corporately.
'
'» provide any code from this site as part of your application's original source code,
'  regardless of whether that code is provided free or at a cost.
'
'» use any code here to create a complete demo application for distribution on the web,
'  either commercially, privately or as an open source project,
'  so long as the purpose of that demo application is to illustrate or
'  create something other than that code found here. In other words the VBnet code,
'  while enhancing functionality of the demo application, forms only a portion of a
'  larger application whose primary goal is not to demonstrate the code taken from
'  this site.
'
'http://www.mvps.org/vbnet/index.html

'This code is taken from the following page:
'http://www.mvps.org/vbnet/code/network/getadaptersinfo-localipaddresses.htm


Option Explicit

' (LB)
Public IPAddressToUse As String  ' With multiple NICs, this stores the IP address of the NIC to use


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2003 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const MAX_ADAPTER_NAME_LENGTH         As Long = 256
Private Const MAX_ADAPTER_DESCRIPTION_LENGTH  As Long = 128
Private Const MAX_ADAPTER_ADDRESS_LENGTH      As Long = 8
Private Const ERROR_SUCCESS  As Long = 0

Private Type IP_ADDRESS_STRING
    IpAddr(0 To 15)  As Byte
End Type

Private Type IP_MASK_STRING
    IpMask(0 To 15)  As Byte
End Type

Private Type IP_ADDR_STRING
    dwNext     As Long
    IpAddress  As IP_ADDRESS_STRING
    IpMask     As IP_MASK_STRING
    dwContext  As Long
End Type

Private Type IP_ADAPTER_INFO
  dwNext                As Long
  ComboIndex            As Long  'reserved
  sAdapterName(0 To (MAX_ADAPTER_NAME_LENGTH + 3))        As Byte
  sDescription(0 To (MAX_ADAPTER_DESCRIPTION_LENGTH + 3)) As Byte
  dwAddressLength       As Long
  sIPAddress(0 To (MAX_ADAPTER_ADDRESS_LENGTH - 1))       As Byte
  dwIndex               As Long
  uType                 As Long
  uDhcpEnabled          As Long
  CurrentIpAddress      As Long
  IPAddressList         As IP_ADDR_STRING
  GatewayList           As IP_ADDR_STRING
  DhcpServer            As IP_ADDR_STRING
  bHaveWins             As Long
  PrimaryWinsServer     As IP_ADDR_STRING
  SecondaryWinsServer   As IP_ADDR_STRING
  LeaseObtained         As Long
  LeaseExpires          As Long
End Type

Private Declare Function GetAdaptersInfo Lib "iphlpapi.dll" _
  (pTcpTable As Any, _
   pdwSize As Long) As Long
   
Private Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (dst As Any, _
   src As Any, _
   ByVal bcount As Long)
   

'pass a character to be used as the
'delimiter in the list of returned addresses.
   
Public Function LocalIPAddresses(ByVal sDelim As String) As String
   
  'api vars
   Dim cbRequired  As Long
   Dim buff()      As Byte
   Dim Adapter     As IP_ADAPTER_INFO
   Dim AdapterStr  As IP_ADDR_STRING
    
  'working vars
   Dim ptr1        As Long
   Dim sIPAddr     As String
   Dim sAllAddr    As String
   
   Call GetAdaptersInfo(ByVal 0&, cbRequired)

   If cbRequired > 0 Then
    
      ReDim buff(0 To cbRequired - 1) As Byte
      
      If GetAdaptersInfo(buff(0), cbRequired) = ERROR_SUCCESS Then
      
        'get a pointer to the data stored in buff()
         ptr1 = VarPtr(buff(0))
                  
        'ptr1 is 0 when no more adapters
         Do While (ptr1 <> 0)
         
           'copy the data from the pointer to the
           'first adapter into the IP_ADAPTER_INFO type
            CopyMemory Adapter, ByVal ptr1, LenB(Adapter)
         
            With Adapter
         
              'the DHCP IP address is in the
              'IpAddress.IpAddr member
               sIPAddr = TrimNull(StrConv(.IPAddressList.IpAddress.IpAddr, vbUnicode))
               sAllAddr = sAllAddr & sIPAddr & sDelim  'Bugfix here by LB
               
              'more?
               ptr1 = .dwNext
               
            End With  'With Adapter
            
         Loop  'Do While (ptr1 <> 0)

      End If  'If GetAdaptersInfo
   End If  'If cbRequired > 0

  'remove the last comma
   If Len(sAllAddr) > 0 Then
      sAllAddr = Left$(sAllAddr, Len(sAllAddr) - 1)
   End If
         
  'return any string found
   LocalIPAddresses = sAllAddr
   
   
End Function


Private Function TrimNull(item As String)

    Dim pos As Integer
   
   'double check that there is a chr$(0) in the string
    pos = InStr(item, Chr$(0))
    If pos Then
          TrimNull = Left$(item, pos - 1)
    Else: TrimNull = item
    End If
  
End Function
'--end block--'


' Get the list of IP addresses on the computer, and add them to the drop-down list.
' VV  Written by Leif Bloomquist VV

Public Sub DetermineIPs()

On Error GoTo DetermineIPsError:

    Dim IPAddressesTemp As String
    Dim AllIPAddresses() As String
    Dim T As Integer
    
    'Determine IP addresses of all NICs
    IPAddressesTemp = LocalIPAddresses(",")
    
    If (Advanced.DetailedDiagnostics.value) Then AddMessage "IPs Detected: " & IPAddressesTemp
    AllIPAddresses = Split(IPAddressesTemp, ",")
    
    For T = 0 To UBound(AllIPAddresses)
        TelnetBBS.IPAddressList.AddItem AllIPAddresses(T)
    Next T
    
    'If there's only one IP address - use it.
    If (UBound(AllIPAddresses) = 0) Then
        IPAddressToUse = AllIPAddresses(0)
    End If

    Exit Sub
    
DetermineIPsError:
    AddMessage "DetermineIPs(): " & Err.Description & " (" & Err.Number & ")"
    Exit Sub
End Sub
