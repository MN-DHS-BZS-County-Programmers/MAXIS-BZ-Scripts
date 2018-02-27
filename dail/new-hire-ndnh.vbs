'------------------THIS SCRIPT IS DESIGNED TO BE RUN FROM THE DAIL SCRUBBER.
'------------------As such, it does NOT include protections to be ran independently.

'Required for statistical purposes===============================================================================
name_of_script = "DAIL - NEW HIRE NDNH.vbs"
start_time = timer
STATS_counter = 1              'sets the stats counter at one
STATS_manualtime = 345         'manual run time in seconds
STATS_denomination = "C"       'C is for each MEMBER
'END OF stats block==============================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message
			critical_error_msgbox = MsgBox ("Something has gone wrong. The Functions Library code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
                                            "FuncLib URL: " & FuncLib_URL & vbNewLine & vbNewLine &_
                                            "The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
                                            vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
            StopScript
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

'DIALOGS----------------------------------------------------------------------------------------------
'This is a dialog asking if the job is known to the agency.
BeginDialog new_HIRE_dialog, 0, 0, 291, 195, "New HIRE dialog"
  EditBox 80, 10, 25, 15, HH_memb
  CheckBox 5, 30, 160, 10, "Check here if this job is known to the agency.", job_known_checkbox
  EditBox 95, 45, 190, 15, employer
  CheckBox 5, 65, 190, 10, "Check here to have the script make a new JOBS panel.", create_JOBS_checkbox
  CheckBox 5, 80, 190, 10, "Check here if you sent a status update to CCA.", CCA_checkbox
  CheckBox 5, 95, 160, 10, "Check here is you sent a status update to ES. ", ES_checkbox
  CheckBox 5, 110, 165, 10, "Check here if you send a Work Number request. ", work_number_checkbox
  CheckBox 5, 125, 165, 10, "Check here if you are requesting CEI/OHI docs.", requested_CEI_OHI_docs_checkbox
  CheckBox 5, 140, 235, 10, "Check here to have the script send a TIKL to return proofs in 10 days.", TIKL_checkbox
  EditBox 50, 155, 235, 15, other_notes
  EditBox 65, 175, 110, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 180, 175, 50, 15
    CancelButton 235, 175, 50, 15
    PushButton 175, 15, 45, 10, "prev. panel", prev_panel_button
    PushButton 175, 25, 45, 10, "next panel", next_panel_button
    PushButton 235, 15, 45, 10, "prev. memb", prev_memb_button
    PushButton 235, 25, 45, 10, "next memb", next_memb_button
  Text 5, 180, 60, 10, "Worker signature:"
  GroupBox 170, 5, 115, 35, "STAT-based navigation"
  Text 5, 50, 85, 10, "Job on DAIL is listed as:"
  Text 5, 160, 40, 10, "Other notes:"
  Text 5, 15, 70, 10, "HH member number:"
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------------

'Connecting to BlueZone
EMConnect ""

'The script needs to determine what the day is in a MAXIS friendly format. The following does that.
current_month = datepart("m", date)
If len(current_month) = 1 then current_month = "0" & current_month
current_day = datepart("d", date)
If len(current_day) = 1 then current_day = "0" & current_day
current_year = datepart("yyyy", date)
current_year = current_year - 2000

	DAIL_check = MsgBox("Please resolve all previous months DAILs prior to running the script." & vbcr & new_hire_third_line & vbcr & new_hire_fourth_line & vbcr & "Please review and click OK if you wish to continue and CANCEL if there are DAILs from previous months for this case.", vbOKCancel)
	If DAIL_check = vbCancel then script_end_procedure("The script has ended. Please resolve any previous months DAILs and rerun the script.")

'Brings the highlighted message to the top and finds the case number
EMSendKey "t"
transmit
EMReadScreen case_number, 7, 5, 73


'SELECTS THE DAIL MESSAGE AND READS THE RESPONSE
EMSendKey "x"
transmit
row = 1
col = 1
EMSearch "DATE HIRED", row, col 					'Has to search, because every once in a while the rows and columns can slide one or two positions.
If row = 0 then script_end_procedure("MAXIS may be busy: the script appears to have errored out. This should be temporary. Try again in a moment. If it happens repeatedly contact the alpha user for your agency.")
EMReadScreen new_hire_first_line, 61, row - 1, col - 2 'Reads each line for the case note. COL needs to be subtracted from because of NDNH message format differs from original new hire format. 
EMReadScreen new_hire_second_line, 61, row , col				
EMReadScreen new_hire_third_line, 61, row + 1, col
EMReadScreen new_hire_fourth_line, 61, row + 2, col
EMReadScreen MEMB_name, 61, 11, 22
	MEMB_name = replace(MEMB_name, "  ", "")
IF right(new_hire_third_line, 46) <> right(new_hire_fourth_line, 46) then 				'script was being run on cases where the names did not match but SSN did. This will allow users to review.
	warning_box = MsgBox("The names found on the NEW HIRE message do not match exactly." & vbcr & new_hire_third_line & vbcr & new_hire_fourth_line & vbcr & "Please review and click OK if you wish to continue and CANCEL if the name is incorrect.", vbOKCancel)
	If warning_box = vbCancel then script_end_procedure("The script has ended. Please review the new hire as you indicated that the name read from the NEW HIRE and the MAXIS name did not match.")
END IF
row = 1 									'Now it's searching for info on the hire date as well as employer
col = 1
EMSearch "DATE HIRED   :", row, col
EMReadScreen date_hired, 10, row, col + 15  '+ 15 because of the offset where the search finds it. 
If date_hired = "  -  -  EM" then date_hired = current_month & "-" & current_day & "-" & current_year
date_hired = CDate(date_hired)
month_hired = Datepart("m", date_hired)
If len(month_hired) = 1 then month_hired = "0" & month_hired
day_hired = Datepart("d", date_hired)
If len(day_hired) = 1 then day_hired = "0" & day_hired
year_hired = Datepart("yyyy", date_hired)
year_hired = year_hired - 2000
EMSearch "EMPLOYER:", row, col
EMReadScreen employer, 25, row, col + 10
row = 1 									'Now it's searching for the SSN
col = 1
'EMSearch "SSN #", row, col           No longer has SSN # in the DAIL message.
EMReadScreen new_HIRE_SSN, 9, 9, 5

'removing any extra spaces after the employer name
employer = replace(employer, "  ", "")

PF3

'CHECKING CASE CURR. MFIP AND SNAP HAVE DIFFERENT RULES. 
EMWriteScreen "h", 6, 3
transmit
row = 1
col = 1
EMSearch "FS: ", row, col
If row <> 0 then FS_case = True
If row = 0 then FS_case = False
row = 1
col = 1
EMSearch "MFIP: ", row, col
If row <> 0 then MFIP_case = True
If row = 0 then MFIP_case = False

PF3

'GOING TO STAT
EMSendKey "s" 
transmit
EMReadScreen stat_check, 4, 20, 21
If stat_check <> "STAT" then script_end_procedure("Unable to get to stat due to an error screen. Clear the error screen and return to the DAIL. Then try the script again.")


'GOING TO MEMB, NEED TO CHECK THE HH MEMBER
EMWriteScreen "memb", 20, 71
transmit
Do
	EMReadScreen MEMB_current, 1, 2, 73
	EMReadScreen MEMB_total, 1, 2, 78
	EMReadScreen MEMB_SSN, 11, 7, 42
	If new_HIRE_SSN = replace(MEMB_SSN, " ", "") then
		EMReadScreen HH_memb, 2, 4, 33
		EMReadScreen memb_age, 2, 8, 76
		If cint(memb_age) < 19 then MsgBox "This client is under 19, so make sure to check that school verification is on file."
	End if
	transmit
Loop until (MEMB_current = MEMB_total) or (new_HIRE_SSN = replace(MEMB_SSN, " ", "-"))


'GOING TO JOBS
EMWriteScreen "jobs", 20, 71
EMWriteScreen HH_memb, 20, 76
transmit


'MFIP cases need to manually add the JOBS panel for ES purposes.
If MFIP_case = False then create_JOBS_checkbox = checked 

'Defaulting the "set TIKL" variable to checked
TIKL_checkbox = checked

'Setting the variable for the following do...loop
HH_memb_row = 5 

'Show dialog
Do	
	Do
		Dialog new_HIRE_dialog
		cancel_confirmation
		MAXIS_dialog_navigation
	Loop until ButtonPressed = -1
call check_for_password(are_we_passworded_out)  			'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

'Checking to see if 5 jobs already exist. If so worker will need to manually delete one first. 
EMReadScreen jobs_total_panel_count, 1, 2, 78
IF create_JOBS_checkbox = checked AND jobs_total_panel_count = "5" THEN script_end_procedure("This client has 5 jobs panels already. Please review and delete and unneeded panels if you want the script to add a new one.")

'If new job is known, script ends.
If job_known_checkbox = checked then 

		'Navigates back to DAIL
		Do
			EMReadScreen DAIL_check, 4, 2, 48
			If DAIL_check = "DAIL" then exit do
			PF3
		Loop until DAIL_check = "DAIL"
	
		EMWriteScreen "I", 6, 3
		transmit
		EMWriteScreen new_HIRE_SSN, 3, 63
		EMWriteScreen "HIRE", 20, 71
		transmit
	

		HIRE_selected = 7

		DO
			MsgBox ("Please enter a U next to the HIRE match you wish to update for case number " & case_number & chr(13) & "There will be a 5 second delay in the script to allow time to make your selection" & chr(13) & chr(13) & "Please do NOT hit transmit!" & chr(13) & "The script will transmit for you.")
			If HIRE_selected = 1 then exit DO
			'puts in a 5 second delay in the script to allow time for the worker to select the HIRE match to update.
			Dim dteWait
			dteWait = DateAdd("s", 5, Now())
			Do Until (Now() > dteWait)
			Loop
		HIRE_selected = MsgBox ("Have you selected the HIRE match you wish to update?", 51, "Hire selected?")
			If HIRE_selected = 2 then
				cancel_confirm = MsgBox("Are you sure you want to cancel the script? Press YES to cancel. Press NO to return to the script.", vbYesNo)
				If cancel_confirm = vbYes then stopscript
			End if
		Loop until HIRE_selected = 6

		transmit
		EMWriteScreen "Y", 16, 54
		transmit
		transmit
		PF3
		PF3

	script_end_procedure("The script will stop as this job is known.")
End IF

'Now it will create a new JOBS panel for this case.
If create_JOBS_checkbox = checked then
	EMWriteScreen "nn", 20, 79						'Creates new panel
	transmit	'Transmits
	EMReadScreen MAXIS_footer_month, 2, 20, 55			'Reads footer month for updating the panel
	EMReadScreen MAXIS_footer_year, 2, 20, 58				'Reads footer year
	EMWriteScreen "w", 5, 34						'Wage income is the type
	EMWriteScreen "n", 6, 34						'No proof has been provided	
	EMWriteScreen employer, 7, 42						'Adds employer info
	EMWriteScreen month_hired, 9, 35					'Adds month hired to start date (this is actually the day income was received)
	EMWriteScreen day_hired, 9, 38					'Adds day hired
	EMWriteScreen year_hired, 9, 41					'Adds year hired
	EMWriteScreen MAXIS_footer_month, 12, 54				'Puts footer month in as the month on prospective side of panel
		IF month_hired = MAXIS_footer_month THEN     			'This accounts for rare cases when new hire footer month is the same as the hire date. 
			EMWriteScreen day_hired, 12, 57				'Puts date hired if message is from same month as hire ex 01/16 new hire for 1/17/16 start date.
		ELSE
			EMWriteScreen current_day, 12, 57				'Puts today in as the day on prospective side, because that's the day we edited the panel		
		END IF
	EMWriteScreen MAXIS_footer_year, 12, 60				'Puts footer year in on prospective side
	EMWriteScreen "0", 12, 67						'Puts $0 in as the received income amt
	EMWriteScreen "0", 18, 72						'Puts 0 hours in as the worked hours
		If FS_case = True then 							'If case is SNAP, it creates a PIC
			EMWriteScreen "x", 19, 38			
			transmit	
			IF month_hired = MAXIS_footer_month THEN     		'This accounts for rare cases when new hire footer month is the same as the hire date. 
				EMWriteScreen month_hired, 5, 34
				EMWriteScreen day_hired, 5, 37
				EMWriteScreen year_hired, 5, 40
			ELSE
				EMWriteScreen current_month, 5, 34		
				EMWriteScreen current_day, 5, 37
				EMWriteScreen current_year, 5, 40
			END IF
			EMWriteScreen "1", 5, 64
			EMWriteScreen "0", 8, 64
			EMWriteScreen "0", 9, 66
			transmit
			transmit
			transmit
		End if
	transmit

								'Transmits to submit the panel
	EMReadScreen expired_check, 6, 24, 17 				'Checks to see if the jobs panel will carry over by looking for the "This information will expire" at the bottom of the page and adds the JOBS panel to the next month
		If expired_check = "EXPIRE" THEN 
			PF3
			EMWriteScreen "y", 16, 54
			transmit

			'GOING TO JOBS
			EMWriteScreen "jobs", 20, 71
			EMWriteScreen HH_memb, 20, 76
			EMWriteScreen "nn", 20, 79
			transmit
			EMWriteScreen "w", 5, 34						'Wage income is the type
			EMWriteScreen "n", 6, 34						'No proof has been provided	
			EMWriteScreen employer, 7, 42						'Adds employer info
			EMWriteScreen month_hired, 9, 35					'Adds month hired to start date (this is actually the day income was received)
			EMWriteScreen day_hired, 9, 38					'Adds day hired
			EMWriteScreen year_hired, 9, 41					'Adds year hired
				EMReadScreen Current_month, 2, 20, 55
			EMWriteScreen Current_month, 12, 54					'Puts footer month in as the month on prospective side of panel
				IF month_hired = MAXIS_footer_month THEN     			'This accounts for rare cases when new hire footer month is the same as the hire date. 
					EMWriteScreen day_hired, 12, 57				'Puts date hired if message is from same month as hire ex 01/16 new hire for 1/17/16 start date.
				ELSE
					EMWriteScreen current_day, 12, 57				'Puts today in as the day on prospective side, because that's the day we edited the panel
				END IF
			EMWriteScreen MAXIS_footer_year, 12, 60				'Puts footer year in on prospective side
			EMWriteScreen "0", 12, 67						'Puts $0 in as the received income amt
			EMWriteScreen "0", 18, 72						'Puts 0 hours in as the worked hours
			
		If FS_case = True then 							'If case is SNAP, it creates a PIC
			EMWriteScreen "x", 19, 38			
			transmit	
			IF month_hired = MAXIS_footer_month THEN     		'This accounts for rare cases when new hire footer month is the same as the hire date. 
				EMWriteScreen month_hired, 5, 34
				EMWriteScreen day_hired, 5, 37
				EMWriteScreen year_hired, 5, 40
			ELSE
				EMWriteScreen MAXIS_footer_month, 5, 34		
				EMWriteScreen current_day, 5, 37
				EMWriteScreen current_year, 5, 40
			END IF
			EMWriteScreen "1", 5, 64
			EMWriteScreen "0", 8, 64
			EMWriteScreen "0", 9, 66

					transmit
					transmit
					transmit


			End IF	

		End if

					transmit
					transmit
					transmit


	'Navigates to case note
	PF4

	'Creates blank case note
	PF9
	transmit

	'Writes new hire message but removes the SSN. 
	EMSendKey replace(new_hire_first_line, new_HIRE_SSN, "XXX-XX-XXXX") & "<newline>" & new_hire_second_line & "<newline>" & new_hire_third_line + "<newline>" & new_hire_fourth_line & "<newline>" & "---" & "<newline>"

	'Writes that the message is unreported, and that the proofs are being sent/TIKLed for.
	call write_variable_in_case_note("* Job unreported to the agency.")
	call write_variable_in_case_note("* Sent employment verification and DHS-2919B (Verification Request Form - B).")
	If create_JOBS_checkbox = checked then call write_variable_in_case_note("* JOBS updated with new hire info from DAIL.")
	if CCA_checkbox = 1 then call write_variable_in_case_note("* Sent status update to CCA.")
	if ES_checkbox = 1 then call write_variable_in_case_note("* Sent status update to ES.")
	if work_number_checkbox = 1 then call write_variable_in_case_note("* Sent Work Number request.")
	If requested_CEI_OHI_docs_checkbox = checked then call write_variable_in_case_note("* Requested CEI/OHI docs.")
	If TIKL_checkbox = checked then call write_variable_in_case_note("* TIKLed for 10-day return and INFC - HIRE match update.")
	call write_bullet_and_variable_in_case_note("Other notes", other_notes)
	call write_variable_in_case_note("---")
	call write_variable_in_case_note(worker_signature & ", using automated script.")
	PF3
	PF3
End if



'Navigates back to DAIL
Do
	EMReadScreen DAIL_check, 4, 2, 48
	If DAIL_check = "DAIL" then exit do
	PF3
Loop until DAIL_check = "DAIL"



'If TIKL_checkbox is checked it enters a TIKL.
IF TIKL_checkbox = checked then

	'Navigates to TIKL
	EMSendKey "w"
	transmit
	'The following will generate a TIKL formatted date for 10 days from now, and add it to the TIKL
	call create_MAXIS_friendly_date(date, 10, 5, 18)

	'Setting cursor on 9, 3, because the message goes beyond a single line and EMWriteScreen does not word wrap.
	EMSetCursor 9, 3

	'Sending TIKL text.
	call write_variable_in_TIKL("INFC - HIRE Match update needed. Verification of " & employer & " job via NEW HIRE should have been returned by now. If not received and processed, take appropriate action. (TIKL auto-generated from script).")


	'Submits TIKL
	transmit
	'Exits TIKL
	PF3
end IF


Do
	EMReadScreen DAIL_check, 4, 2, 48
	If DAIL_check = "DAIL" then exit do
	PF3
Loop until DAIL_check = "DAIL"


'If TIKL_checkbox is unchecked, it needs to end here.
If TIKL_checkbox = unchecked then script_end_procedure("Success! MAXIS updated for new HIRE message, and a case note made. An Employment Verification and Verif Req Form B should now be sent. The job is at " & employer & ".")


script_end_procedure("Success! MAXIS updated for new HIRE message, a case note made, and a TIKL has been sent for 10 days from now. An Employment Verification and Verif Req Form B should now be sent." & chr(13) & chr(13) & "The job is at " & employer & " for " & MEMB_name & " (SSN - " & new_HIRE_SSN & ")" & " on case number " & case_number & ".")

