'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - APPLICATION RECEIVED.vbs"
start_time = timer
msgbox "arbitrary"
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

'DIALOGS-------------------------------------------------------------
BeginDialog case_appld_dialog, 0, 0, 161, 65, "Application Received"
  EditBox 95, 5, 60, 15, case_number
  EditBox 95, 25, 60, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 45, 45, 50, 15
    CancelButton 105, 45, 50, 15
  Text 5, 10, 85, 10, "Enter your case number:"
  Text 5, 30, 85, 10, "Worker Signature"
EndDialog

BeginDialog app_detail_dialog, 0, 0, 221, 205, "Detail of application"
  DropListBox 80, 5, 135, 45, "Select One"+chr(9)+"In Person"+chr(9)+"Droped Off"+chr(9)+"Mail"+chr(9)+"Online"+chr(9)+"Fax"+chr(9)+"Email", how_app_recvd
  DropListBox 80, 25, 135, 20, "Select One"+chr(9)+"CAF"+chr(9)+"ApplyMN"+chr(9)+"HC - Certain Populations"+chr(9)+"HCAPP"+chr(9)+"Addendum", app_type
  EditBox 80, 45, 135, 15, confirmation_number
  EditBox 60, 65, 75, 15, time_of_app
  DropListBox 145, 65, 70, 15, "AM"+chr(9)+"PM", AM_PM
  EditBox 50, 85, 165, 15, worker_name
  EditBox 120, 105, 95, 15, worker_number
  EditBox 150, 125, 65, 15, pended_date
  EditBox 5, 145, 210, 15, entered_notes
  CheckBox 5, 165, 205, 15, "Check here to have script transfer case to assigned worker", transfer_case
  ButtonGroup ButtonPressed
    OkButton 110, 185, 50, 15
    CancelButton 165, 185, 50, 15
  Text 5, 10, 70, 10, "Application received"
  Text 5, 30, 65, 10, "Type of application"
  Text 5, 50, 60, 10, "Confirmation #"
  Text 5, 70, 50, 10, "Time received"
  Text 5, 90, 40, 10, "Assigned to:"
  Text 5, 110, 110, 10, "Worker Number (X###### format)"
  Text 5, 130, 25, 10, "Notes:"
  Text 105, 130, 40, 10, "Pended on:"
EndDialog

'Grabs the case number
EMConnect ""

CALL MAXIS_case_number_finder (case_number)

'Runs the first dialog - which confirms the case number and gathers worker signature
Do 
	Dialog case_appld_dialog
	If buttonpressed = cancel then stopscript
	If case_number = "" then MsgBox "You must have a case number to continue!"
	If worker_signature = "" then Msgbox "Please sign your case note"
Loop until case_number <> "" AND worker_signature <> ""

call check_for_MAXIS(true)

'Gathers Date of application and creates MAXIS friendly dates to be sure to navigate to the correct time frame

call navigate_to_MAXIS_screen("REPT","PND2")
dateofapp_row = 1 
dateofapp_col = 1 
EMSearch case_number, dateofapp_row, dateofapp_col
EMReadScreen footer_month, 2, dateofapp_row, 38
EMReadScreen app_day, 2, dateofapp_row, 41
EMReadScreen footer_year, 2, dateofapp_row, 44
date_of_app = footer_month & "/" & app_day & "/" & footer_year

'Determines which programs are currently pending in the month of application
call navigate_to_MAXIS_screen("STAT","PROG")
EMReadScreen cash1_pend, 4, 6, 74
EMReadScreen cash2_pend, 4, 7, 74
EMReadScreen emer_pend, 4, 8, 74
EMReadScreen grh_pend, 4, 9, 74
EMReadScreen fs_pend, 4, 10, 74
EMReadScreen ive_pend, 4, 11, 74
EMReadScreen hc_pend, 4, 12, 74

IF cash1_pend = "PEND" THEN
	cash1_pend = 1 
	Else 
	cash1_pend = 0 
End If

If cash2_pend = "PEND" THEN 
	cash2_pend = 1
	Else 
	cash2_pend = 0
End if
	
If emer_pend = "PEND" THEN 
	emer_pend = 1
	Else 
	emer_pend = 0
End if

If grh_pend = "PEND" THEN 
	ghr__pend = 1
	Else 
	grh_pend = 0
End if

If fs_pend = "PEND" THEN 
	fs_pend = 1
	Else 
	fs_pend = 0
End if

If ive_pend = "PEND" THEN 
	ive_pend = 1
	Else 
	ive_pend = 0
End if

If hc_pend = "PEND" THEN 
	hc_pend = 1
	Else 
	hc_pend = 0
End if
	
'Defaults the date pended to today
pended_date = date & ""

'Runs the second dialog - which gathers information about the application
Do
	Do
		Do	
			Dialog app_detail_dialog
			cancel_confirmation
			If app_type = "Select One" then MsgBox "Please enter the type of application received."
			If how_app_recvd = "Select One" then MsgBox "Please enter how the application was received to the agency."
			If worker_name = "" then MsgBox "Please enter who this case was assigned to."
		Loop until (app_type <> "Select One" AND how_app_recvd <> "Select One" AND worker_name <> "")
		If transfer_case = 1 AND (worker_number = "" OR len(worker_number) <> 7) then MsgBox "You must enter the MAXIS number of the worker if you would like the case to be transfered by the script, be sure that it is in X###### format."
	Loop until (worker_number <> "" AND len(worker_number) = 7 OR transfer_case = 0)
	If app_type = "ApplyMN" AND isnumeric(confirmation_number) = false AND time_of_app = "" = true then MsgBox "If an ApplyMN was received, you must enter the confirmation number and time received"
Loop until (app_type = "ApplyMN" and isnumeric(confirmation_number) = true) AND time_of_app <> "" OR app_type <> "ApplyMN"

'Creates a variable that lists all the programs pending.
If cash1_pend = 1 THEN programs_applied_for = programs_applied_for & "Cash, "
If cash2_pend = 1 THEN programs_applied_for = programs_applied_for & "Cash, "
If emer_pend = 1 THEN programs_applied_for = programs_applied_for & "Emergency, "
If grh_pend = 1 THEN programs_applied_for = programs_applied_for & "GRH, "
If fs_pend = 1 THEN programs_applied_for = programs_applied_for & "SNAP, "
If ive_pend = 1 THEN programs_applied_for = programs_applied_for & "IV-E, "
If hc_pend = 1 THEN programs_applied_for = programs_applied_for & "HC"

'Transfers the case to the assigned worker if this was selected in the second dialog box
IF transfer_case = 1 THEN
	CALL navigate_to_MAXIS_screen ("SPEC", "XFER")
	EMWriteScreen "x", 7, 16
	transmit
	PF9
	EMWriteScreen worker_number, 18, 61
	transmit
	EMReadScreen worker_check, 9, 24, 2 
	IF worker_check = "SERVICING" THEN 
		MsgBox "The correct worker number was not entered, this X-Number is not a valid worker in MAXIS. You will need to transfer the case manually"
		PF10
		transfer_case = unchecked
	End If
End If


IF time_of_app <> "" Then
	time_stamp = " at " & time_of_app & " " & AM_PM
ELSE 
	time_stamp = " "
End If

'Writes the case note 	
CALL start_a_blank_case_note
CALL write_variable_in_CASE_NOTE ("APP PENDED - " & app_type & " rec'vd via " & how_app_recvd & " on " & date_of_app & time_stamp)
IF isnumeric(confirmation_number) = true THEN CALL write_bullet_and_variable_in_CASE_NOTE ("Confirmation # ", confirmation_number)
CALL write_bullet_and_variable_in_CASE_NOTE ("Requesting", programs_applied_for)
CALL write_bullet_and_variable_in_CASE_NOTE ("Pended on", pended_date)
CALL write_bullet_and_variable_in_CASE_NOTE ("Application assigned to", worker_name)
IF transfer_case = checked THEN CALL write_variable_in_CASE_NOTE ("* Case transfered to " & worker_name & " in MAXIS")
IF entered_notes <> "" THEN CALL write_bullet_and_variable_in_CASE_NOTE ("Notes", entered_notes)
CALL write_variable_in_CASE_NOTE ("---")
CALL write_variable_in_CASE_NOTE (worker_signature)

'Reminder to screen for XFS if SNAP is pending.
IF fs_pend = 1 THEN MsgBox ("SNAP is pending, be sure to run the NOTES-Expedited Screening script as well to note potential XFS eligibility")

script_end_procedure ("")
