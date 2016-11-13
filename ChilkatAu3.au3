;~ https://www.chilkatsoft.com/refdoc/xChilkatSshTunnelRef.html
;~ https://www.youtube.com/watch?v=5_iSdS05CKA

Global $acceptLog
Global $SshIdleTimeoutMs = 15000, $SshConnectTimeoutMs = 15000
Global $SshHostname = "64.132.38.245", $SshUsername = "admin", $SshPassword = "default"
Global $SshPort = 22, $SshOutboundBindIpAddress = "0.0.0.0", $SshListenBindIpAddress = "127.0.0.1", $SshListenPort = 8889

Global $oChilkat = ObjCreate("Chilkat_9_5_0.SshTunnel")
If Not IsObj($oChilkat) Then
	Run("regsvr32 /s ChilkatAx-9.5.0-win32.dll")
	Exit
EndIf

With $oChilkat
	.TcpNoDelay = 1
	.VerboseLogging = True
	.IdleTimeoutMs = $SshIdleTimeoutMs
	.ConnectTimeoutMs = $SshConnectTimeoutMs
	.InboundSocksVersion = 5
	.OutboundBindIpAddress = $SshOutboundBindIpAddress
	.ListenBindIpAddress = $SshListenBindIpAddress
	.KeepAcceptLog = 1
	.AcceptLogPath = @ScriptDir & "\AcceptLog.txt"

	ConsoleWrite("-->.UnlockComponent" & @CRLF)
	If Not .UnlockComponent('ĐÉO CÓ KEY') Then
		ConsoleWrite(.LastErrorText & @CRLF)
		Exit
	EndIf

	ConsoleWrite("-->.Connect" & @CRLF)
	If Not .Connect($SshHostname, $SshPort) Then
		ConsoleWrite(.LastErrorText & @CRLF)
		Exit
	EndIf

	ConsoleWrite("-->.AuthenticatePw" & @CRLF)
	If Not .AuthenticatePw($SshUsername, $SshPassword) Then
		ConsoleWrite(.LastErrorText & @CRLF)
		Exit
	EndIf

	ConsoleWrite("-->.DynamicPortForwarding" & @CRLF)
	.DynamicPortForwarding = 1
	If Not .LastMethodSuccess Then
		ConsoleWrite(.LastErrorText & @CRLF)
		Exit
	EndIf

	AdlibRegister("_ErrorLog", 100)

	ConsoleWrite("-->.BeginAccepting" & @CRLF)
	If Not .BeginAccepting($SshListenPort) Then
		ConsoleWrite(.LastErrorText & @CRLF)
		Exit
	EndIf
EndWith

Func _Exit()
	$oChilkat.StopAccepting()
	$oChilkat.CloseTunnel()
EndFunc   ;==>_Exit

Func _ErrorLog()
	Local $acceptLog = $oChilkat.AcceptLog
	If StringInStr($acceptLog, 'SocketError') Or StringInStr($acceptLog, 'Socket bind failed') Or StringInStr($acceptLog, 'bind-and-listen failed') Then
		AdlibUnRegister("_ErrorLog")
		MsgBox(0, "Error", $acceptLog)
		Exit
	EndIf
EndFunc   ;==>_ErrorLog