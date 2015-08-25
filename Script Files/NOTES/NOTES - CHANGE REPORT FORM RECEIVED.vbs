'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - CHANGE REPORTED"
start_time = timer 'manual time= ? 


'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF default_directory = "C:\DHS-MAXIS-Scripts\Script Files\" THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		ELSEIF beta_agency = "" or beta_agency = True then							'If you're a beta agency, you should probably use the beta branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else																		'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
			MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
					vbCr & _
					"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
					vbCr & _
					"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
					vbTab & "- The name of the script you are running." & vbCr &_
					vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
					vbTab & "- The name and email for an employee from your IT department," & vbCr & _
					vbTab & vbTab & "responsible for network issues." & vbCr &_
					vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
					vbCr & _
					"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
					vbCr &_
					"URL: " & FuncLib_URL
					script_end_procedure("Script ended due to error connecting to GitHub.")
		END IF
	ELSE
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'THE DIALOG--------------------------------------------------------------------------------------------------------------

BeginDialog change_reported_dialog, 0, 0, 411, 335, "Change Reported"
  EditBox 60, 5, 55, 15, case_number
  EditBox 270, 5, 60, 15, date_reported
  ComboBox 70, 30, 260, 15, " "+chr(9)+"Change Report Form "+chr(9)+"Phone contact from client"+chr(9)+"Written Correspondence "+chr(9)+"Email Received", How_change_reported_combobox
  EditBox 50, 55, 330, 15, address_notes
  CheckBox 390, 55, 10, 10, "", address_verif_checkbox
  EditBox 50, 75, 330, 15, household_notes
  CheckBox 390, 75, 10, 10, "", household_verif_checkbox
  EditBox 50, 95, 330, 15, savings_notes
  CheckBox 390, 95, 10, 10, "", savings_verif_checkbox
  EditBox 50, 115, 330, 15, property_notes
  CheckBox 390, 115, 10, 10, "", property_verif_checkbox
  EditBox 50, 135, 330, 15, vehicles_notes
  CheckBox 390, 135, 10, 10, "", vehicle_verif_checkbox
  EditBox 50, 155, 330, 15, income_notes
  CheckBox 390, 155, 10, 10, "", Income_verif_checkbox
  EditBox 50, 175, 330, 15, shelter_notes
  CheckBox 390, 175, 10, 10, "", Shelter_verif_checkbox
  EditBox 50, 195, 330, 15, other
  CheckBox 390, 195, 10, 10, "", Other_verif_checkbox
  EditBox 50, 225, 340, 15, actions_taken
  EditBox 50, 245, 340, 15, other_notes
  EditBox 65, 265, 325, 15, verifs_requested
  CheckBox 10, 290, 140, 10, "Check here to navigate to DAIL/WRIT", tikl_nav_check
  DropListBox 270, 290, 95, 15, "Select One..."+chr(9)+"will continue next month"+chr(9)+"will not continue next month", changes_continue
  EditBox 80, 315, 85, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 290, 315, 50, 15
    CancelButton 350, 315, 50, 15
  Text 5, 10, 50, 10, "Case Number:"
  Text 165, 10, 90, 10, "Change Reported Date"
  Text 5, 25, 60, 20, "How was this change reported?"
  Text 385, 30, 25, 20, "Verif Recvd "
  Text 5, 60, 30, 15, "Address:"
  Text 5, 80, 45, 10, "HHLD Comp:"
  Text 5, 100, 30, 15, "Savings:"
  Text 5, 120, 35, 15, "Property:"
  Text 5, 140, 35, 15, "Vehicles:"
  Text 5, 160, 30, 10, "Income:"
  Text 5, 180, 30, 10, "Shelter:"
  Text 5, 200, 25, 10, "Other:"
  Text 5, 230, 45, 15, "Action Taken:"
  Text 5, 250, 40, 15, "Other notes:"
  Text 5, 270, 60, 10, "Verifs Requested:"
  Text 180, 290, 90, 10, "The changes client reports:"
  Text 10, 315, 60, 15, "Worker Signature"
EndDialog








'THE SCRIPT--------------------------------------------------------------------------------------------------------------

EMConnect "" 'Connect to Bluezone

CALL MAXIS_case_number_finder(case_number) 'Grabs Maxis Case number
Call Check_for_MAXIS(True)

DO
	err_msg = ""
	Dialog change_reported_dialog
	cancel_confirmation
	IF case_number = "" OR (case_number <> "" AND IsNumeric(case_number) = False) THEN err_msg = err_msg & vbNewLine & "*Please enter a valid case number"
	If date_reported = "" THEN err_msg = err_msg & vbNewLine & "*You must enter the date the change was reported"
	IF How_change_reported_combobox = " " OR How_change_reported_combobox = "" THEN err_msg = err_msg & vbNewLine & "*You must enter how the change was reported"
	IF actions_taken = "" THEN err_msg = err_msg & vbNewLine & "*You must enter an action taken"
	IF worker_signature = "" THEN err_msg = err_msg & vbNewLine & "*You must sign your case note"
	IF err_msg <> "" THEN Msgbox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine & vbNewLine & "Please resolve for the script to continue"
LOOP UNTIL err_msg = ""


Dim Changes_reported
If address_notes <> "" THEN Changes_reported = changes_reported & "Addr, "
If household_notes <> "" Then changes_reported = changes_reported & "HHLD Comp, "
If savings_notes <> "" OR property_notes <> "" OR vehicle_notes <> "" THEN changes_reported = changes_reported & "Assets, "
If income_notes <> "" Then changes_reported = changes_reported & "Income, "
If other <> "" Then changes_reported = changes_reported & "Other, "
changes_reported = left(changes_reported, len(changes_reported) -2)



'Checks Maxis for password prompt
CALL check_for_MAXIS(FALSE)

'THE CASE NOTE----------------------------------------------------------------------------------------------------
'Navigates to case note
Call start_a_blank_CASE_NOTE

CALL write_variable_in_case_note("Changes Reported: " & changes_reported)   
CALL write_bullet_and_variable_in_case_note("Date change reported",  date_reported)
Call write_bullet_and_variable_in_case_note("Change was reported by",  how_change_reported_combobox)

IF address_verif_checkbox = checked THEN 
	CALL write_bullet_and_variable_in_case_note("Address", address_notes & "- Verified" )
Else 
	CALL write_bullet_and_variable_in_case_note("Address", address_notes)
End If 


If household_verif_checkbox = checked THEN 
	CALL write_bullet_and_variable_in_case_note("HHLD Comp", household_notes & "- Verified")
Else 
	CALL write_bullet_and_variable_in_case_note("HHLD Comp", household_notes)
End If


If savings_verif_checkbox = checked THEN 
	CALL write_bullet_and_variable_in_case_note("Savings", savings_notes & "- Verified")
Else
	CALL write_bullet_and_variable_in_case_note("Savings", savings_notes)
End If


If property_verif_checkbox = checked THEN 
	CALL write_bullet_and_variable_in_case_note("Property", property_notes & "- Verified")
Else 
	CALL write_bullet_and_variable_in_case_note("Property", property_notes)
End If


If vehicle_verif_checkbox = checked THEN
	CALL write_bullet_and_variable_in_case_note("Vehicles", vehicles_notes & "- Verified")
Else 
	CALL write_bullet_and_variable_in_case_note("Vehicles", vehicles_notes)
End If


If Income_verif_checkbox = checked THEN 
	CALL write_bullet_and_variable_in_case_note("Income", income_notes & "- Verified")
Else 
	CALL write_bullet_and_variable_in_case_note("Income", income_notes)
End If


If Shelter_verif_checkbox = checked THEN 
	CALL write_bullet_and_variable_in_case_note("Shelter", shelter_notes & "- Verified")
Else 
	CALL write_bullet_and_variable_in_case_note("Shelter", shelter_notes)
End If

If Other_verif_checkbox = checked THEN 
	CALL write_bullet_and_variable_in_case_note("Other", other & "- Verified")
Else
	CALL write_bullet_and_variable_in_case_note("Other", other)
End If


CALL write_bullet_and_variable_in_case_note("Action Taken", actions_taken)
CALL write_bullet_and_variable_in_case_note("Other Notes", other_notes)
CALL write_bullet_and_variable_in_case_note("Verifs Requested", verifs_requested)
IF changes_continue <> "Select One..." THEN CALL write_bullet_and_variable_in_case_note("The changes reported", changes_continue)
CALL write_variable_in_case_note("---")
CALL write_variable_in_case_note(worker_signature)

'If we checked to TIKL out, it goes to TIKL and sends a TIKL
IF tikl_nav_check = 1 THEN
	CALL navigate_to_MAXIS_screen("DAIL", "WRIT")
	CALL create_MAXIS_friendly_date(date, 10, 5, 18)
	EMSetCursor 9, 3
END IF

script_end_procedure("")
