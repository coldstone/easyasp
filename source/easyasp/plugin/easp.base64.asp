<%
class EasyASP_Base64
	private sBASE_64_CHARACTERS
	Private Sub Class_Initialize()
		sBASE_64_CHARACTERS = String2Bytes("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/")
	end sub
	
	'****************************************************
	'@DESCRIPTION:	encode string
	'@PARAM:	str [String] : source string need to be encoded.
	'@RETURN:	[String] Base64 string
	'****************************************************
	function Encode(byval str)
		Encode=EncodeBytes(String2Bytes(str))
	end function
	
	'****************************************************
	'@DESCRIPTION:	encode bytes
	'@PARAM:	str [Byte[]] : Byte Array need to be encoded.
	'@RETURN:	[String] Base64 string
	'****************************************************
	function EncodeBytes(byval str)
		EncodeBytes=Bytes2String(Base64encode(str),"gb2312")
	end function
	
	'****************************************************
	'@DESCRIPTION:	decode base64 string as default charset(charset value is from response.charset)
	'@PARAM:	str [String] : base64 string need to be decoded.
	'@RETURN:	[String] source string
	'****************************************************
	function Decode(byval str)
		dim charset:charset = response.Charset
		if charset="" then charset = "gb2312"
		Decode=decodeAny(str,charset)
	end function
	
	'****************************************************
	'@DESCRIPTION:	decode base64 string as utf-8 string
	'@PARAM:	str [String] :  base64 string need to be decoded.
	'@RETURN:	[String] source string
	'****************************************************
	function DecodeUTF8(byval str)
		DecodeUTF8=decodeAny(str,"utf-8")
	end function
	
	'****************************************************
	'@DESCRIPTION:	decode base64 string as any charset
	'@PARAM:	str [String] : base64 string need to be decoded.
	'@PARAM:	charset [String] : charset of source string.
	'@RETURN:	[String] source string.
	'****************************************************
	function DecodeAny(byval str,byval charset)
		DecodeAny=Bytes2String(DecodeBytes(str),charset)
	end function
	
	'****************************************************
	'@DESCRIPTION:	decode base64 string as bytes data
	'@PARAM:	str [String] : base64 string need to be decoded.
	'@RETURN:	[Byte[]] decode result
	'****************************************************
	function DecodeBytes(byval str)
		DecodeBytes=Base64decode(String2Bytes(str))
	end function
	
	'****************************************************
	'@DESCRIPTION:	decode base64 as binary data(you can write the data to stream)
	'@PARAM:	data [String] : base64 string need to be decoded.
	'@RETURN:	[Binary] Binary data.
	'****************************************************
	function DecodeBinary(byval data)
		dim xmlstr:xmlstr="<?xml version=""1.0"" encoding=""gb2312""?><root xmlns:dt=""urn:schemas-microsoft-com:datatypes""><data dt:dt=""bin.base64"">" & data & "</data></root>"
		dim xmldom:set xmldom = server.CreateObject("Microsoft.XMLDOM")
		xmldom.loadxml xmlstr
		DecodeBinary = xmldom.selectSingleNode("//root/data").nodeTypedValue
		set xmldom = nothing
	end function
	
	Function Base64encode(asContents)
	 Dim lnPosition
	 Dim lsResult
	 Dim Char1
	 Dim Char2
	 Dim Char3
	 Dim Char4
	 Dim Byte1
	 Dim Byte2
	 Dim Byte3
	 Dim SaveBits1
	 Dim SaveBits2
	 Dim lsGroupBinary
	 Dim lsGroup64
	 Dim m4,len1,len2
		
	 len1=Lenb(asContents)
	 if len1<1 then
	  Base64encode=""
	  exit Function
	 end if
	 asContents = midb(asContents,1)
	 m3=Len1 Mod 3
	 If M3 > 0 Then asContents = asContents & String(3-M3, chrb(0))
	 IF m3 > 0 THEN
	  len1=len1+(3-m3)
	  len2=len1-3
	 else
	  len2=len1
	 end if
	 lsResult = ""
	 For lnPosition = 1 To len2 Step 3
	  lsGroup64 = ""
	  lsGroupBinary = Midb(asContents, lnPosition, 3)
	  
	  Byte1 = Ascb(Midb(lsGroupBinary, 1, 1)): SaveBits1 = Byte1 And 3
	  Byte2 = Ascb(Midb(lsGroupBinary, 2, 1)): SaveBits2 = Byte2 And 15
	  Byte3 = Ascb(Midb(lsGroupBinary, 3, 1))
	  
	  Char1 = Midb(sBASE_64_CHARACTERS, ((Byte1 And 252) \ 4) + 1, 1)
	  Char2 = Midb(sBASE_64_CHARACTERS, (((Byte2 And 240) \ 16) Or (SaveBits1 * 16) And &HFF) + 1, 1)
	  Char3 = Midb(sBASE_64_CHARACTERS, (((Byte3 And 192) \ 64) Or (SaveBits2 * 4) And &HFF) + 1, 1)
	  Char4 = Midb(sBASE_64_CHARACTERS, (Byte3 And 63) + 1, 1)
	  lsGroup64 = Char1 & Char2 & Char3 & Char4
	   
	  lsResult = lsResult & lsGroup64
	 Next
	  
	 if M3 > 0 then
	  lsGroup64 = ""
	  lsGroupBinary = Midb(asContents, len2+1, 3)
	  
	  Byte1 = Ascb(Midb(lsGroupBinary, 1, 1)): SaveBits1 = Byte1 And 3
	  Byte2 = Ascb(Midb(lsGroupBinary, 2, 1)): SaveBits2 = Byte2 And 15
	  Byte3 = Ascb(Midb(lsGroupBinary, 3, 1))
	  
	  Char1 = Midb(sBASE_64_CHARACTERS, ((Byte1 And 252) \ 4) + 1, 1)
	  Char2 = Midb(sBASE_64_CHARACTERS, (((Byte2 And 240) \ 16) Or (SaveBits1 * 16) And &HFF) + 1, 1)
	  Char3 = Midb(sBASE_64_CHARACTERS, (((Byte3 And 192) \ 64) Or (SaveBits2 * 4) And &HFF) + 1, 1)
	   
	  if M3=1 then
	   lsGroup64 = Char1 & Char2 & ChrB(61) & ChrB(61)
	  else
	   lsGroup64 = Char1 & Char2 & Char3 & ChrB(61)
	  end if
	   
	  lsResult = lsResult & lsGroup64
	 end if
	  
	 Base64encode = lsResult  
	End Function
	
	Function Base64decode(asContents)
	 Dim lsResult
	 Dim lnPosition
	 Dim lsGroup64, lsGroupBinary
	 Dim Char1, Char2, Char3, Char4
	 Dim Byte1, Byte2, Byte3
	 Dim M4,len1,len2
	 
	 len1= Lenb(asContents)
	 M4 = len1 Mod 4
	 
	 if len1 < 1 or M4 > 0 then
	  Base64decode = ""
	  exit Function
	 end if
	
	 if midb(asContents, len1, 1) = chrb(61) then m4=3
	 if midb(asContents, len1-1, 1) = chrb(61) then m4=2
	 
	 if m4 = 0 then
	  len2=len1
	 else
	  len2=len1-4
	 end if
	
	 For lnPosition = 1 To Len2 Step 4
	  lsGroupBinary = ""
	  lsGroup64 = Midb(asContents, lnPosition, 4)
	  Char1 = InStrb(sBASE_64_CHARACTERS, Midb(lsGroup64, 1, 1)) - 1
	  Char2 = InStrb(sBASE_64_CHARACTERS, Midb(lsGroup64, 2, 1)) - 1
	  Char3 = InStrb(sBASE_64_CHARACTERS, Midb(lsGroup64, 3, 1)) - 1
	  Char4 = InStrb(sBASE_64_CHARACTERS, Midb(lsGroup64, 4, 1)) - 1
	  Byte1 = Chrb(((Char2 And 48) \ 16) Or (Char1 * 4) And &HFF)
	  Byte2 = lsGroupBinary & Chrb(((Char3 And 60) \ 4) Or (Char2 * 16) And &HFF)
	  Byte3 = Chrb((((Char3 And 3) * 64) And &HFF) Or (Char4 And 63))
	  lsGroupBinary = Byte1 & Byte2 & Byte3
	
	  lsResult = lsResult & lsGroupBinary
	 Next
	
	 if M4 > 0 then
	  lsGroupBinary = ""
	  lsGroup64 = Midb(asContents, len2+1, m4) & chrB(65)
	  if M4=2 then
	   lsGroup64 = lsGroup64 & chrB(65)
	  end if
	  Char1 = InStrb(sBASE_64_CHARACTERS, Midb(lsGroup64, 1, 1)) - 1
	  Char2 = InStrb(sBASE_64_CHARACTERS, Midb(lsGroup64, 2, 1)) - 1
	  Char3 = InStrb(sBASE_64_CHARACTERS, Midb(lsGroup64, 3, 1)) - 1
	  Char4 = InStrb(sBASE_64_CHARACTERS, Midb(lsGroup64, 4, 1)) - 1
	  Byte1 = Chrb(((Char2 And 48) \ 16) Or (Char1 * 4) And &HFF)
	  Byte2 = lsGroupBinary & Chrb(((Char3 And 60) \ 4) Or (Char2 * 16) And &HFF)
	  Byte3 = Chrb((((Char3 And 3) * 64) And &HFF) Or (Char4 And 63))
	
	  if M4=2 then
	   lsGroupBinary = Byte1
	  elseif M4=3 then
	   lsGroupBinary = Byte1 & Byte2
	  end if
	
	  lsResult = lsResult & lsGroupBinary
	 end if
	
	 Base64decode = lsResult
	
	End Function
	function String2Bytes(byval content)
		content = Server.URLEncode(content)
		content = replace(content,"+"," ")
		dim ret,i,c
		i=1
		do while i<=len(content)
			c = mid(content,i,1)
			if c="%" then
				ret = ret & chrb(cbyte("&H" & mid(content,i+1,2)))
				i=i+3
			else
				ret = ret & chrb(asc(c))
				i=i+1
			end if
		loop
		String2Bytes = ret
	end function
	public Function Bytes2String(ByVal byt,byval charset)
		If LenB(byt) = 0 Then
			Bytes2String = ""
			Exit Function
		End If
		Dim mystream, bstr
		Set mystream =Server.CreateObject("ADODB.Stream")
		mystream.Type = 2
		mystream.Mode = 3
		mystream.Open
		mystream.WriteText byt
		mystream.Position = 0
		mystream.CharSet = charset
		mystream.Position = 2
		bstr = mystream.ReadText()
		mystream.Close
		Set mystream = Nothing
		Bytes2String = bstr
	End Function
end class
%>