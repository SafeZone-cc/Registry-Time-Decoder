'Registry time decoder by Alex Dragokas ver.2.2

const QT = """"

Dim UnixTime, UTCTime

Set oFSO   = CreateObject("Scripting.FileSystemObject")
Set oShell = CreateObject("WScript.Shell")

cHost64 = oShell.ExpandEnvironmentStrings("%SystemRoot%") & "\sysnative\cscript.exe" 'защита от запуска из-под Wow64
cHost32 = oShell.ExpandEnvironmentStrings("%SystemRoot%") & "\system32\cscript.exe"

' Запущен ли из консоли
vbHost = oFSO.GetBaseName(Wscript.FullName)
if strcomp(vbHost, "cscript", 1) <> 0 then
	On Error resume next
    oShell.Run QT & cHost64 & QT & " //nologo " & QT & WScript.ScriptFullName & QT, 1, false
	if err.number <> 0 then
	    oShell.Run QT & cHost32 & QT & " //nologo " & QT & WScript.ScriptFullName & QT, 1, false		
	end if
    WScript.Quit
end if

'sKey = inputBox("Введите путь к параметру реестра в формате Улей\Ключ\Параметр, либо введите сырое 16-ричное значение")

WScript.Echo ("--------------------------------------------")
WScript.Echo ("   Registry time decoder by Alex Dragokas   ")
WScript.Echo ("--------------------------------------------" & vbcrlf)

WScript.Echo("Введите путь к параметру реестра в формате Улей\Ключ\Параметр," & vbcrlf & "набор байт с запятой или без, либо 16-ричное значение с префиксом 0x:")
WScript.Echo("")
WScript.Echo("Примеры:")
WScript.Echo("")
WScript.Echo("1. Unix-Time (4 byte)")
WScript.Echo("58b69de3 - [Hex]")
WScript.Echo("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\InstallDate - [Reg]")
WScript.Echo("")
WScript.Echo("2. FILETIME (8 byte)")
WScript.Echo("0x01c694c5a38c8000 - [Hex]")
WScript.Echo("00,80,8c,a3,c5,94,c6,01 - [Binary Hex]")
WScript.Echo("")
WScript.Echo("3. SYSTEMTIME (16 bytes)")
WScript.Echo("e0,07,04,00,03,00,14,00,06,00,2b,00,3a,00,33,01 - [Binary Hex]")

if WScript.Arguments.Count <> 0 then sKeyArg = UnQuote(WScript.Arguments(0))

do

  b16byte = false
  b8byte = false
  b4byte = false
  bRegBased = false
  bHexValue = false

  if sKeyArg <> "" then ' если преобразование запрошено через аргумент запуска скрипта
    sKey = sKeyArg
    sKeyArg = ""
	WScript.Echo vbcrlf & ">>> " & sKey
  else
    WScript.StdOut.Write vbcrlf & ">>> "
    sKey = WScript.StdIn.ReadLine()
  end if

  sKey = NormalizeRegPath(Trim(sKey))

  WScript.Echo("")

  if instr(sKey, "\") = 0 then 'value based
    sKey = replace(sKey, ",", "")
    sKey = replace(sKey, " ", "")

	if StrComp(Left(sKey,2),"0x",1) = 0 then 'just 16h number (not a byte data)
		bHexValue = true
		sKey = mid(sKey, 3)
	end if

	if Len(sKey) = 0 then WScript.Quit

	if Len(sKey) > 32 or not IsHex(sKey) then
		WScript.Echo ("Неверный формат входящих данных! Требуется HEX-значение в виде числа или бинарной строки.")
	elseif len(sKey) > 8 and len(sKey) <= 16 then '8 bytes values (FILETIME)
        b8byte = true
		if bHexValue then
			UTCTime = sKey
		else
			UTCTime = ReverseBytesLine(sKey)
		end if
	elseif len(sKey) <= 8 then '4 bytes value (Unix-time)
		b4byte = true
		UnixTime =  CLng("&H" & sKey)
    else '16 byte values (SYSTEMTIME)
		b16byte = true
		redim Bytes(15)
		For i = 1 to 31 step 2
			if Len(sKey) >= i then
				Bytes((i+1)\2-1) = CLng("&H" & mid(sKey,i,2))
			end if
		Next
	end if
  else 'registry based
    bRegBased = true
    set oShell = CreateObject("WScript.Shell")
    On Error resume next
    Bytes = oShell.RegRead(sKey)
    If Err.Number <> 0 then
		if msgbox ("Указанного параметра нет, неверный формат параметра, либо нужно запустить этот скрипт от имени администратора!" & vbcrlf & "Запустить от имени админа сейчас?", vbYesNo or vbExclamation, "Ошибка") = vbYes then
			Set oShellApp = CreateObject("Shell.Application")
			oShellApp.ShellExecute cHost32, "//nologo " & QT & WScript.ScriptFullName & QT & " " & QT & sKey & QT, "", "runas", 1
		end if
	end if
	On Error Goto 0
	if isArray(Bytes) then
		if UBound(Bytes) = 7 then '8-byte value
			b8byte = true
			UTCTime = ""
			for i = 7 to 0 step -1
				UTCTime = UTCTime & right("0" & Hex(Bytes(i)),2)
			next
		elseif UBound(Bytes) = 15 then '16-byte value
			b16byte = true
		else
			WScript.Echo("Неподдерживаемый тип параметра.")
		end if
	else
		if isNumeric(Bytes) then '4-byte value
			b4byte = true
			UnixTime = Bytes
		else
			WScript.Echo("Неподдерживаемый тип параметра.")
		end if
	end if
  end if

  if b4byte then call Decode4byte()
  if b8byte then call Decode8byte()
  if b16byte then call Decode16byte()

loop


'Unix-Time
sub Decode4byte()
	if bRegBased then
		WScript.Echo "REG = " & UnixTime & vbcrlf
	end if
	WScript.Echo (DateAdd("s", UnixTime, #1/1/1970#))
end sub

function IsHexNumber(str)
	if strcomp(left(str, 2), "0x", 1) = 0 then
		IsHexNumber = true
	else
		dim i, codeA, codeF, code
		codeA = asc("A")
		codeF = asc("F")
		for i = 1 to len(str)
			code = asc(ucase(mid(str, i, 1)))
			if code >= codeA and code <= codeF then
				IsHexNumber = true
				exit for
			end if
		next
	end if
end function

'FILETIME
sub Decode8byte()
    'typedef struct _FILETIME {
    '  DWORD dwLowDateTime;
    '  DWORD dwHighDateTime;
    '} FILETIME, *PFILETIME;

    '100 ns. after #1/1/1601#

	if bRegBased then
		WScript.Echo "REG = " & UTCTime & vbcrlf
	else
		WScript.Echo "==> " & UTCTime & vbcrlf
	end if

	Set dateTime = CreateObject("WbemScripting.SWbemDateTime")
	On Error Resume Next
	dateTime.SetFileTime HexToDec(UTCTime), false
	If Err.Number <> 0 then
		WScript.Echo ("Слишком большое число! Возможно, Вы указали 16-ричное число вместо байтов." & vbcrlf & "Пробую в обратном порядке ...")
		UTCTime = ReverseBytesLine(UTCTime)
		WScript.Echo ("==> " & UTCTime)
		Err.Clear
		dateTime.SetFileTime HexToDec(UTCTime), false
		If Err.Number <> 0 then
			WScript.Echo ("Неудача!")
			if not IsHexNumber(UTCTime) then
				Err.Clear
				WScript.Echo ("Проверяю число как 10-ричное:")
				UnixTime = Clng(UTCTime)
				WScript.Echo ("==> " & Cstr(UnixTime))
				WScript.Echo ""
				Decode4byte
				If Err.Number <> 0 then
					WScript.Echo ("Неудача!")
				end if
			end if
		else
			WScript.Echo (vbcrlf & dateTime.GetVarDate)
		end if
	else
		WScript.Echo (dateTime.GetVarDate)
	end if
end sub

'SYSTEMTIME
sub Decode16byte()
	'typedef struct _SYSTEMTIME {
	'  WORD wYear;
	'  WORD wMonth;
	'  WORD wDayOfWeek;
	'  WORD wDay;
	'  WORD wHour;
	'  WORD wMinute;
	'  WORD wSecond;
	'  WORD wMilliseconds;
	'} SYSTEMTIME, *PSYSTEMTIME;

	if bRegBased then 'Print reg. info
		strReg = ""
		for i = 0 to Ubound(Bytes)
			strReg = strReg & right("0" & Hex(Bytes(i)),2) & ","
		next
		WScript.Echo "REG = " & left(strReg, len(strReg)-1) & vbcrlf
	end if

	if Ubound(Bytes) >= 1 then wYear = Byte2Word(Bytes(0), Bytes(1))
	if Ubound(Bytes) >= 3 then wMonth = Byte2Word(Bytes(2), Bytes(3))
	if Ubound(Bytes) >= 7 then wDay = Byte2Word(Bytes(6), Bytes(7))
	if Ubound(Bytes) >= 9 then wHour = Byte2Word(Bytes(8), Bytes(9))
	if Ubound(Bytes) >= 11 then wMinute = Byte2Word(Bytes(10), Bytes(11))
	if Ubound(Bytes) >= 13 then wSecond = Byte2Word(Bytes(12), Bytes(13))
	if Ubound(Bytes) >= 15 then wMilliseconds = Byte2Word(Bytes(14), Bytes(15))

	WScript.Echo ( Right("0" & wDay, 2) & "." & _
			   Right("0" & wMonth, 2) & "." & _
			   Right("000" & wYear, 4) & " " & _
			   Right("0" & wHour, 2) & ":" & _
			   Right("0" & wMinute, 2) & ":" & _
			   Right("0" & wSecond, 2))
end sub

function Byte2Word(LoByte, HiByte)
	Byte2Word = HiByte * 256 + LoByte
end function

function IsHex(num)
  tmp = num
  for i = 1 to 16
    tmp = replace(tmp, mid("0123456789ABCDEF",i,1), "", 1, -1, 1)
  next
  if len(tmp) = 0 and len(num) > 0 then IsHex = true
end function

Function HexToDec(strHex)
    size = Len(strHex) - 1
    ret = CDbl(0)
    For i = 0 To size
        ret = ret + CDbl("&H" & Mid(strHex, size - i + 1, 1)) * (CDbl(16) ^ CDbl(i))
    Next
	ret = CStr(ret)
	ret = replace(ret,",","")
	pos = instr(1,ret,"e+",1)
	e = mid(ret,pos+2)
	ret = left(ret, pos -1)
	ret = ret & string(e - len(ret) +1, "0")
    HexToDec = ret
End Function

function UnQuote(byval str)
	if left(str,1) = QT then str = mid(str, 2)
	if right(str,1) = QT then str = left(str, len(str)-1)
	UnQuote = str
end function

function ReverseBytesLine(byval sLine) '00808ca3c594c601 => 01c694c5a38c8000
	strRet = ""
	if len(sLine) mod 2 = 1 then sLine = "0" & sLine
	for i = len(sLine) - 1 to 1 step -2
		strRet = strRet & mid(sLine, i, 2)
	next
	ReverseBytesLine = strRet
end function

function NormalizeRegPath(path)
	if StrComp(Left(path, 2), "HK", 1) = 0 then
		NormalizeRegPath = path
	else
		Dim pos
		pos = instr(1, path, "HK", 1)
		if pos <> 0 then NormalizeRegPath = Mid(path, pos)
	end if
end function

'Examples:
'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\InstallDate (Unix-Time)
'HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Control\Class \...\ => DriverDateData (FILETIME)
'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\NetworkList\Profiles \...\ => DateCreated (SYSTEMTIME)

'58b69de3
'00,80,8c,a3,c5,94,c6,01
'e0,07,04,00,03,00,14,00,06,00,2b,00,3a,00,33,01
