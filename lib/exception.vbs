'Option Explicit
' vbsProcaine:Exception v0.01
'    Copyright (C) 2011  Anton Vattay
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program.  If not, see <http://www.gnu.org/licenses/>.

'===============================================================================
'===============================================================================
' Error Handling ---------------------------------------------------------------
' This section supports try-catch&throw functionality in vbscript.
' You should only surround one exception throwing command with this
' construct, otherwise you might lose the error.
' The usage idiom is:
' On Error Resume Next 'try
' 	... 'code that could throw exception
' Set Ex = New ErrWrap.catch() 'catch
' On Error GoTo 0 'catch part two
' If (Ex = <some_err_num>) Then
' 	... 'Handle error
' End If

' Note that code called within an error handler that re-throws (using Err.raise)
' must be "exception raise safe" all the way up the call chain.
' If your called function has an "On Error..." statement in it, that will reset
' The global Err object, thereby losing the exception the code was handling. When
' The raise is called at the end of the handling to re-throw, it will throw an
' "non-error" Err object with code 0, which will then slip by any upstream
' error handlers. A nightmare to debug if it happens.

Class ErrWrap
	Private pNumber
	Private pSource
	Private pDescription
	Private pHelpContext
	Private pHelpFile
	Private objReasonEx
	
	Public Function catch()
		init()		
		objReasonEx = NULL
		Set catch = Me
	End Function
	
	Public Function init()
		pNumber = Err.Number
		pSource = Err.Source
		pDescription = Err.Description
		pHelpContext = Err.HelpContext
		pHelpFile = Err.HelpFile		
		Set init = Me
	End Function
	
	Public Function initM(intCode, strSource, strDescription)
		pNumber = intCode
		pSource = strSource
		pDescription = strDescription
		pHelpContext = ""
		pHelpFile = ""		
		Set initM = Me
	End Function
	
	Public Function initExM(intCode, strSource, strDescription, objEx)
		pNumber = intCode
		pSource = strSource
		pDescription = strDescription
		pHelpContext = ""
		pHelpFile = ""		
		Set objReasonEx = objEx
		Set initExM = Me
	End Function
	
	Public Function initEx(objEx)
		pNumber = Err.Number
		pSource = Err.Source
		pDescription = Err.Description
		pHelpContext = Err.HelpContext
		pHelpFile = Err.HelpFile
		Set objReasonEx = objEx
		Set initEx = Me
	End Function
	
	Public Function getReason() 'returns objEx
		If NOT isObject(objReasonEx) Then
			getReason = NULL
		Else
			Set getReason = objReasonEx
		End If
	End Function
	
	Public Default Property Get Number
		Number = pNumber
	End Property
	
	Public Property Get Source
		Source = pSource
	End Property
	
	Public Property Get Description
		Description = pDescription
	End Property
	
	Public Property Get HelpContext
		HelpContext = pHelpContextl
	End Property
	
	Public Property Get HelpFile
		HelpFile = HelpFile
	End Property
	
End Class

'===============================================================================
'===============================================================================
Class ExceptionManager
	Dim currentEx
	
	Function init()
		currentEx = NULL
		Set init = Me
	End Function
	
	Function catch()
		If isNull(currentEx) Then
			Set currentEx = New ErrWrap.catch()
		Else
			If (Err <> currentEx) Then
				'Exception mismatch, when the current exception
				'does not match the last recorded currentEx.
				'Happens when an exception is thrown in an 
				'exception handler
				Set currentEx = New ErrWrap.initEx(currentEx)
			End IF
		End If
		
		Set catch = currentEx
	End Function
	
	Function throw(objEx)
		Set currentEx = objEx
		Err.Raise currentEx.number, currentEx.Source, currentEx.Description
	End Function
	
	Function dump(objEx)
		dump = ""
		If NOT (isObject(objEx.getReason)) Then
			dump = "[" & objEx & ", " & objEx.Source & ", " & objEx.Description & ")" & VbCrLf
		Else
			dump = dump(objEx.getReason) & "[" & objEx & ", " & objEx.Source & ", " & objEx.Description & ")" & VbCrLf
		End If
	End Function
End Class