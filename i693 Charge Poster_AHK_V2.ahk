#Requires AutoHotkey v2.0
#SingleInstance Force
;TODO need to make this conditional might need to turn off other script. Why?
;TODO create a picture that shows what are the hotkeys
^XButton2::EraPostingKeys_CourtesyWriteOff
^RButton::PaymentPaste ;when selecting the insurance balance it will paste it into the payment section

GuiWidth := 250


i693ChargePoster:=Gui('AlwaysOnTop','i693 Charge Poster')
i693ChargePoster.SetFont('s9')
i693ChargePoster.Add('Button', 'xp+4 yp+20 w220 h30 vv1','Exam').OnEvent('Click',ExamFee)

i693ChargePoster.Add('Button', 'xp yp+35 w100 hp vv2 Section','Tuberculosis').OnEvent('Click',ChargePoster)
i693ChargePoster.Add('Button', 'xp yp+35 wp hp vv3','Syphilis').OnEvent('Click',ChargePoster)
i693ChargePoster.Add('Button', 'xp yp+35 wp hp vv4','Gonorrhea').OnEvent('Click',ChargePoster)

i693ChargePoster.Add('Button', 'x+20 yp+35 wp hp vv5','Tetanus Vaccine').OnEvent('Click',ChargePoster)

i693ChargePoster.Add('Button', 'xS yp+35 wp hp vv6 Section','MMR Titer').OnEvent('Click',ChargePoster)
i693ChargePoster.Add('Button', 'x+20 yp wp hp vv7','MMR Vaccine').OnEvent('Click',ChargePoster)

i693ChargePoster.Add('Button', 'xS yp+35 wp hp vv8 Section','Varicella Titer').OnEvent('Click',ChargePoster)
i693ChargePoster.Add('Button', 'x+20 yp wp hp vv9','Varicella Vaccine').OnEvent('Click',ChargePoster)

i693ChargePoster.Add('Button', 'xS yp+35 wp hp vv10 Section','Hep B Titer').OnEvent('Click',ChargePoster)
i693ChargePoster.Add('Button', 'x+20 yp wp hp vv11','Hep B Vaccine').OnEvent('Click',ChargePoster)

i693ChargePoster.Add('Button', 'xp yp+35 wp hp vv12','Polio Vaccine').OnEvent('Click',ChargePoster)

i693ChargePoster.Add('Button', 'xp yp+35 wp hp vv13','Flu').OnEvent('Click',ChargePoster)

i693ChargePoster.Add('Button', 'xS yp+35 w220 hp vv14','Mailing Fee').OnEvent('Click',ChargePoster)
i693ChargePoster.Add('Button', 'xp yp+35 wp hp vv15','5 Day Turn').OnEvent('Click',ChargePoster)
i693ChargePoster.Add('Button', 'xp yp+35 wp hp vv16','Misc').OnEvent('Click',ChargePoster)
i693ChargePoster.Add('Text', 'xp yp+40 wp hp Center','Total')
i693ChargePoster.Add('Edit', 'xp yp+20 wp  vtotal')
i693ChargePoster.Add('Text', 'xp yp+30 wp hp Center','CC')
i693ChargePoster.Add('Edit', 'xp yp+20 wp  vCC').OnEvent('Change',DebitAdjustCalculation)
i693ChargePoster.Add('Edit', 'xp yp+25 wp  vDebitAdjust')
i693ChargePoster.BackColor := 'ed830a'
i693ChargePoster.Show('w' GuiWidth)
i693ChargePoster.OnEvent('Close',i693ChargePoster_Close)





DebitAdjustCalculation(*){
	Try
		i693ChargePoster['DebitAdjust'].value := i693ChargePoster['cc'].value - i693ChargePoster['Total'].value
	return
}

ExamFee(*){
	try {
		WinActivate('Charge Entry --')
		loop 15
			i693ChargePoster['v' A_Index].visible := False
		Send('i6new')
		Sleep(200)
		Send('{Tab}')
		Sleep(200)
		Send('{alt down}a{alt up}')
		Sleep(200)
		Send('{Enter}')
		Sleep(200)
		send("+{tab 2}")
		Sleep(200)
		Send('z02.89')
		Sleep(500)
		Send('{Tab}')
		Send('{alt down}a{alt up}')
		loop 15
			i693ChargePoster['v' A_Index].visible := True
		return
	}
	Catch as e {
		MsgBox('An error was thrown!`nSpecifically: ' e.Message,'Error',4096)
		return
	}
}
	

ChargePoster(ButtonText,NotUsed){
	;Mapping the name of the button to CPT
	ButtonMap := Map('Exam','i6new',
						'Syphilis','i6syp',
						'Gonorrhea','i6gono',
						'Tuberculosis','i6tbgol',
						'MMR Titer','i6mt',
						'MMR Vaccine','i6mv',
						'Hep B Titer','i6ht',
						'Hep B Vaccine','i6hv',
						'Varicella Titer','i6vt',
						'Varicella Vaccine','i6vv',
						'Tetanus Vaccine','i6tv',
						'Flu','i6flu',
						'Mailing Fee','i6mail',
						'5 Day Turn','i6fast',
						'Misc','i6misc',
						'Polio Vaccine', 'i6pv')
	Try {	
		WinActivate('Charge Entry --')
		loop 16
			i693ChargePoster['v' A_Index].visible := False
		Send(ButtonMap.Get(ButtonText.text))
		Sleep(500)
		Send('{Tab}')
		Sleep(500)
		Send('{alt down}a{alt up}')
		Sleep(200)
		loop 16
			i693ChargePoster['v' A_Index].visible := True
	}
	Catch as e {
		MsgBox('An error was thrown!`nSpecifically: ' e.Message,'Error',4096)
		return
	}
	return
}



PaymentPaste(*){
	OldClipboard := A_Clipboard
	A_Clipboard := ""
	Send("{Ctrl Down}c{Ctrl Up}")
	if !ClipWait(1){
		MsgBox "The attempt to copy text onto the clipboard failed."
		return
	}
	InsuranceBalance := A_Clipboard 
	Send("{Tab}")
	Send("{Ctrl Down}v{Ctrl Up}")
	return
}

EraPostingKeys_CourtesyWriteOff(){
	OldClipboard := A_Clipboard
	A_Clipboard := ""
	Send("{Ctrl Down}c{Ctrl Up}")
	if !ClipWait(1){
		MsgBox "The attempt to copy text onto the clipboard failed."
		return
	}
	InsuranceBalance := A_Clipboard 
	Send("{Tab}")
	A_Clipboard := ""
	Send("{Ctrl Down}c{Ctrl Up}")
	if !ClipWait(1) {
		MsgBox 'The attempt to copy text onto the clipboard failed.'
		return
	}
	sleep(200)
	Send("{Tab 2}WOCOUR{Tab}")
	If !Send(Round(InsuranceBalance-A_Clipboard,2)){
		;  msgbox 'here'
	}
	A_Clipboard := OldClipboard
	Return
}


;TODO test this
i693ChargePoster_Close(*) {
	;Run("C:\Users\" A_UserName "\Google Drive\Great City Medical\Scripts\Main Script AHK_V2\Main Script.ahk")
	ExitApp
    ; Do not call ExitApp -- that would prevent other callbacks from being called.
}