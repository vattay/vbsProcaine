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
' Exception Handling -----------------------------------------------------------
' This section supports try-catch&throw functionality in vbscript.
' You should only surround one exception throwing command with this
' construct, otherwise you might lose the error.
' You must use this format in your code to simulate a try catch
' On Error Resume Next 'try
' 	... 'code that could throw exception
' Set caughtErr = New ErrWrap.catch() 'catch
' On Error GoTo 0 'catch part two
' If (caughtErr = <some_err_num>) Then
' 	... 'Handle error
' End If

Class ErrWrap
	Private pNumber
	Private pSource
	Private pDescription
	Private pHelpContext
	Private pHelpFile
	
	Public Function catch()
		pNumber = Err.Number
		pSource = Err.Source
		pDescription = Err.Description
		pHelpContext = Err.HelpContext
		pHelpFile = Err.HelpFile		
		Set catch = Me
	End Function
	
	Public Function Newk(strSource, ErrWrap)
		pNumber = ErrWrap.Number
		pSource = strSource & "->" & ErrWrap.Source
		pDescription = ErrWrap.Description
		pHelpContext = ErrWrap.HelpContext
		pHelpFile = ErrWrap.HelpFile
		Set Newk = Me
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
		HelpContext = pHelpContext
	End Property
	
	Public Property Get HelpFile
		HelpFile = HelpFile
	End Property
	
End Class