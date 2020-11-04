Option Explicit

Function hs_cond_ToBoolean(ByVal tarValue)
	'*** History ***********************************************************************************
	' 2020/11/04, BBS:	- First release
	'
	'***********************************************************************************************

	'*** Documentation *****************************************************************************
	' Helper function for Boolean value conditioning
	'  
	'***********************************************************************************************

	On Error Resume Next
	hs_cond_ToBoolean = False

	'*** Pre-Validation ****************************************************************************
	' Nothing to be validated

	'*** Initialization ****************************************************************************
	' Nothing to be initialized

	'*** Operations ********************************************************************************
	If VarType(tarValue) <> 11 Then
		If IsNumeric(tarValue) Then
			tarValue = CInt(tarValue)

			If tarValue <> 0 Then
				hs_cond_ToBoolean = True
			Else
				hs_cond_ToBoolean = False
			End If
		Else
			tarValue = CStr(tarValue)

			If LCase(tarValue) = "false" or LCase(tarValue) = "no" or LCase(tarValue) = "off" Then
				hs_cond_ToBoolean = False
			Else
				hs_cond_ToBoolean = True
			End If
		End If
	Else
		hs_cond_ToBoolean = tarValue
	End If

	'--- Release -----------------------------------------------------------------------------------
	If Err.Number <> 0 Then
		Err.Clear
	End If
End Function
