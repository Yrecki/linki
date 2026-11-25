strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
 & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colChassis = objWMIService.ExecQuery _
 ("SELECT * FROM Win32_SystemEnclosure")
For Each objChassis in colChassis
 For Each intType in objChassis.ChassisTypes
	Select Case intType
		Case 1	Wscript.Echo "This MainBoard is a Other. (" & intType & ")"
		Case 2	Wscript.Echo "This MainBoard is a Unknown. (" & intType & ")"
		Case 3	Wscript.Echo "This MainBoard is a Desktop. (" & intType & ")"
		Case 4	Wscript.Echo "This MainBoard is a Low Profile Desktop. (" & intType & ")"
		Case 5	Wscript.Echo "This MainBoard is a Pizza Box. (" & intType & ")"
		Case 6	Wscript.Echo "This MainBoard is a Mini Tower. (" & intType & ")"
		Case 7	Wscript.Echo "This MainBoard is a Tower. (" & intType & ")"
		Case 8	Wscript.Echo "This MainBoard is a Portable. (" & intType & ")"
		Case 9	Wscript.Echo "This MainBoard is a Laptop. (" & intType & ")"
		Case 10	Wscript.Echo "This MainBoard is a Notebook. (" & intType & ")"
		Case 11	Wscript.Echo "This MainBoard is a Hand Held. (" & intType & ")"
		Case 12	Wscript.Echo "This MainBoard is a Docking Station. (" & intType & ")"
		Case 13	Wscript.Echo "This MainBoard is a All in One. (" & intType & ")"
		Case 14	Wscript.Echo "This MainBoard is a Sub Notebook. (" & intType & ")"
		Case 15	Wscript.Echo "This MainBoard is a Space-Saving. (" & intType & ")"
		Case 16	Wscript.Echo "This MainBoard is a Lunch Box. (" & intType & ")"
		Case 17	Wscript.Echo "This MainBoard is a Main System Chassis. (" & intType & ")"
		Case 18	Wscript.Echo "This MainBoard is a Expansion Chassis. (" & intType & ")"
		Case 19	Wscript.Echo "This MainBoard is a Sub Chassis. (" & intType & ")"
		Case 20	Wscript.Echo "This MainBoard is a Bus Expansion Chassis. (" & intType & ")"
		Case 21	Wscript.Echo "This MainBoard is a Peripheral Chassis. (" & intType & ")"
		Case 22	Wscript.Echo "This MainBoard is a Storage Chassis. (" & intType & ")"
		Case 23	Wscript.Echo "This MainBoard is a Rack Mount Chassis. (" & intType & ")"
		Case 24	Wscript.Echo "This MainBoard is a Sealed-Case PC. (" & intType & ")"
	Case Else
		Wscript.Echo "This MainBoard is a " & intType
	End Select
 Next
Next
