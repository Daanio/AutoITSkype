#cs ----------------------------------------------------------------------------

 AutoIt Version: 1.0
 Author: Ralph van Roosmalen (ralphvanroosmalen@gmail.com), based on a script of Andy Flesner (Airwolf123)

 Script Function:
	Read a line from a text file and send this line as chat message to a groupchat.
	
http://www.autoitscript.com/forum/topic/72869-skype-com-examples-skype4comlib/

 Licensed under BSD License:
	Copyright (c) 2004-2006, Skype Limited.
	All rights reserved. 

	Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met: 
		- Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer. 
		- Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer 
		  in the documentation and/or other materials provided with the distribution. 
		- Neither the name of the Skype Limited nor the names of its contributors may be used to endorse or promote products derived from 
		  this software without specific prior written permission. 
	
	THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT 
	LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT 
	OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT 
	LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON 
	ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF 
	THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE. 

#ce ----------------------------------------------------------------------------

#include <file.au3>

Dim $aRecords

;// Create a Skype4COM object:
$oSkypeObjectCreatedToUseInAutoITScript = ObjCreate("Skype4COM.Skype")
$oSkypeObjectCreatedToUseInAutoITScriptEvent = ObjEvent($oSkypeObjectCreatedToUseInAutoITScript,"Skype_")

;// Start the Skype client:
If Not $oSkypeObjectCreatedToUseInAutoITScript.Client.IsRunning Then
	$oSkypeObjectCreatedToUseInAutoITScript.Client.Start()
EndIf

;// Verify that a user is signed in and online before continuing
While 1
	If $oSkypeObjectCreatedToUseInAutoITScript.CurrentUserStatus = $oSkypeObjectCreatedToUseInAutoITScript.Convert.TextToUserStatus("ONLINE") Then
		ExitLoop
	Else
		$oSkypeObjectCreatedToUseInAutoITScript.ChangeUserStatus($oSkypeObjectCreatedToUseInAutoITScript.Convert.TextToUserStatus("ONLINE"))
	EndIf
	Sleep(2000)
WEnd

;//Retrieve the chat id with the for example Tracer, http://developer.skype.com/resources/Tracer.exe
$oChat = $oSkypeObjectCreatedToUseInAutoITScript.Chat("#ralphvanroosmalen/1450d4bf63dfbd86")

;On some special days we inform the testers about the day that is coming.
if (@MON = 4) and (@MDAY = 29) Then
 $Message = "Tomorrow is a public holiday in the Netherlands, http://en.wikipedia.org/wiki/Koninginnedag"
Elseif (@MON = 1) and (@MDAY = 2) Then
  $Message = "Happy New Year to you all!"
Elseif (@MON = 12) and (@MDAY = 24) Then
 $Message = "Enjoy X-Mas!"
Elseif (@MON = 8) and (@MDAY = 14) Then
$Message = "Tomorrow is a public holiday in India, Independence Day, http://en.wikipedia.org/wiki/Independence_Day_(India)"
Elseif (@MON = 1) and (@MDAY = 25) Then
                     $Message = "Tomorrow is a public holiday in India, Republic Day, http://en.wikipedia.org/wiki/Republic_Day_(India)"
ElseIf (@MON = 10) and (@MDAY = 1) Then 
   $Message = "Tomorrow is a public holiday in India, Mahatma Gandhi's Birthday, http://en.wikipedia.org/wiki/Gandhi_Jayanti"
Else
	If Not _FileReadToArray("tipofthedayfortesters.txt",$aRecords) Then
		MsgBox(4096,"Error", " Error reading log to Array error:" & @error)
		Exit
	 EndIf
	 $MDday = Random($aRecords[0] + 1)
	$TipOfTheDayNr = mod(@MDAY,$aRecords[0])
	if $TipOfTheDayNr = 0 Then
		$TipOfTheDayNr = $aRecords[0]
	EndIf
	
	$Message = "Tip of the day: " & $aRecords[$TipOfTheDayNr]
EndIf

;Send the message to the groupchat.
$oChat.SendMessage($Message)
