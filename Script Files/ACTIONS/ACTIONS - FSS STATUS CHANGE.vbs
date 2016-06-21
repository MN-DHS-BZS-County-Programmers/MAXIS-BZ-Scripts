'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - FSS STATUS CHANGE.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 49                	'manual run time in seconds
STATS_denomination = "M"       		'M is for each MEMBER
'END OF stats block=========================================================================================================

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

'DIALOGS ===================================================================================================================
BeginDialog fss_status_dialog, 0, 0, 221, 335, "FSS Status Update"
  EditBox 60, 5, 65, 15, MAXIS_case_number
  ButtonGroup ButtonPressed
    PushButton 135, 10, 75, 10, "Reload Client Name", get_client_name_button
  EditBox 30, 25, 25, 15, ref_number
  EditBox 60, 25, 155, 15, client_name
  CheckBox 10, 45, 55, 10, "Pre-60 MFIP", pre_sixty_checkbox
  CheckBox 75, 45, 55, 10, "Post-60 MFIP", post_sixty_checkbox
  Text 5, 65, 185, 10, "Select all the ES Status Codes that applies to this client."
  CheckBox 5, 80, 155, 10, "Age 60 or Over - 21", age_sixty_checkbox
  CheckBox 5, 95, 155, 10, "Pregnant/Incapacitated - 22", preg_checkbox
  CheckBox 5, 110, 155, 10, "Ill/Incapacitated for more than 60 Days - 23", ill_incap_checkbox
  CheckBox 5, 125, 155, 10, "Care of Ill/Incap Family Member - 24", care_of_ill_Incap_checkbox
  CheckBox 5, 140, 155, 10, "Care of Child Under 12 Months - 25", child_under_one_checkbox
  CheckBox 5, 155, 155, 10, "Family Violence Waiver - 26", fam_violence_checkbox
  CheckBox 5, 170, 155, 10, "Special Medical Criteria - 27", Special_medical_checkbox
  CheckBox 5, 185, 155, 10, "IQ Tested - 28", iq_test_checkbox
  CheckBox 5, 200, 155, 10, "Learning Disabled - 29", learning_disabled_checkbox
  CheckBox 5, 215, 155, 10, "Mentally Ill - 30", mentally_ill_checkbox
  CheckBox 5, 230, 155, 10, "Developmentally Delayed - 31", dev_delayed_checkbox
  CheckBox 5, 245, 155, 10, "Unemployable - 32", unemployable_checkbox
  CheckBox 5, 260, 155, 10, "SSI/RSDI Pending - 33", ssi_pending_checkbox
  CheckBox 5, 275, 155, 10, "Newly Arrived Immigrant - 34", new_imig_checkbox
  CheckBox 5, 300, 155, 10, "Universal Participant - 20", universal_partipant_checkbox
  ButtonGroup ButtonPressed
    OkButton 90, 315, 50, 15
    CancelButton 145, 315, 50, 15
  Text 5, 30, 25, 10, "Client"
  Text 5, 10, 45, 10, "Case Number"
EndDialog

BeginDialog FSS_final_dialog, 0, 0, 420, 210, "FSS Case Note Information"
  EditBox 80, 5, 335, 15, fss_category_list
  CheckBox 230, 30, 85, 10, "MFIP Results approved", results_approved_checkbox
  CheckBox 230, 50, 105, 10, "MFIP Results NOT approved", not_approved_checkbox
  EditBox 95, 65, 320, 15, notes_not_approved
  EditBox 65, 85, 350, 15, other_notes
  EditBox 10, 115, 395, 15, MFIP_results
  ButtonGroup ButtonPressed
    PushButton 10, 135, 75, 15, "Send case to BGTX", CASE_BGTX_button
    PushButton 85, 170, 25, 10, "BUSI", BUSI_button
    PushButton 110, 170, 25, 10, "JOBS", JOBS_button
    PushButton 135, 170, 25, 10, "UNEA", UNEA_button
    PushButton 185, 170, 25, 10, "MEMB", MEMB_button
    PushButton 210, 170, 25, 10, "MEMI", MEMI_button
    PushButton 235, 170, 25, 10, "EMPS", EMPS_button
    PushButton 260, 170, 25, 10, "REVW", REVW_button
    PushButton 285, 170, 25, 10, "MONT", MONT_button
    PushButton 310, 170, 25, 10, "PBEN", PBEN_button
    PushButton 335, 170, 25, 10, "DISA", DISA_button
    PushButton 360, 170, 25, 10, "IMIG", IMIG_button
    PushButton 385, 170, 25, 10, "TIME", TIME_button
  EditBox 230, 190, 70, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 310, 190, 50, 15
    CancelButton 365, 190, 50, 15
  Text 15, 25, 205, 15, "If the case is ready for approval with the results shown below. APP the results before pressing 'OK' and check this box:"
  Text 15, 50, 205, 10, "Else, check here if MFIP is not ready for approval"
  Text 10, 90, 45, 10, "Other Notes:"
  GroupBox 5, 105, 410, 50, "MFIP Results"
  Text 165, 195, 60, 10, "Worker Signature:"
  Text 30, 10, 50, 10, "FSS Category:"
  GroupBox 80, 160, 85, 25, "Income panels"
  Text 15, 70, 75, 10, "Reason not approved:"
  GroupBox 180, 160, 235, 25, "other STAT panels:"
  Text 275, 140, 85, 10, "Initial Footer Month/Year"
  EditBox 360, 135, 20, 15, month_to_start
  EditBox 385, 135, 20, 15, year_to_start
EndDialog

'===========================================================================================================================

'FUNCTIONS==================================================================================================================
FUNCTION Read_MFIP_Results(month_to_start, year_to_start, MFIP_results)
	Call date_array_generator(month_to_start, year_to_start, date_array)

	For Each version in date_array
		MAXIS_footer_month = right("00" & datepart("m", version), 2)
		MAXIS_footer_year = right(datepart("yyyy", version), 2)
		Back_to_SELF
		Call Navigate_to_MAXIS_screen ("ELIG", "MFIP")
		EMReadScreen elig_check, 4, 3, 47
		If elig_check = "MFPR" Then 
			EMReadScreen process_date, 8, 2, 73
			If CDate(process_date) = date Then 
				EMWriteScreen "MFSM", 20, 71
				transmit
				Do 
					EMReadScreen benefit_status, 13, 10, 31
					benefit_status = trim(benefit_status)
					If benefit_status = "NO CHANGE" Then 
						no_change = TRUE
						EMReadScreen total_grant, 8, 13, 73
						If trim(total_grant) = "0.00" Then 
							EMReadScreen version, 1, 2, 12
							version = abs(version)
							prev_version = version - 1
							EMWriteScreen "0" & prev_version, 20, 79
							transmit
						Else Exit Do
						End If 
					End If 
				Loop Until benefit_status <> "NO CHANGE"
				EMReadScreen total_grant, 8, 13, 73
				EMReadScreen cash_amt, 8, 14, 73
				EMReadScreen food_amt, 8, 15, 73
				EMReadScreen housing_grant, 8, 16, 73
				MFIP_results = MFIP_results & MAXIS_footer_month & "/" & MAXIS_footer_year & " Total Grant: " & total_grant & "; Cash Portion: " & cash_amt & "; Food Portion: " & food_amt & "; Housing Grant: " & housing_grant & "; "
			Else 
				CALL Navigate_to_MAXIS_screen ("STAT", "SUMM")
				summ_row = 2
				Do 
					EMReadScreen edit_msg, 23, summ_row, 20
					If edit_msg = "CASH HAS BEEN INHIBITED" Then 
						inhibiting_error = TRUE
						Exit do 
					End If 
					If trim(edit_msg) = "" Then 
						EMReadScreen next_page, 7, summ_row, 71
						If next_page = "MORE: +" Then 
							PF8
							summ_row = 1
						End If 
					End iF 
					summ_row = summ_row + 1
				Loop until summ_row = 23
				If inhibiting_error = TRUE then 
					MFIP_results = MFIP_results & MAXIS_footer_month & "/" & MAXIS_footer_year & " has an Inhibiting EDIT in STAT - resolve and rerun to generate results."
					inhibiting_error = FALSE
				End If 
			End IF 
		Else 
			CALL Navigate_to_MAXIS_screen ("STAT", "SUMM")
			summ_row = 2
			Do 
				EMReadScreen edit_msg, 23, summ_row, 20
				If edit_msg = "CASH HAS BEEN INHIBITED" Then 
					inhibiting_error = TRUE
					Exit do 
				End If 
				If trim(edit_msg) = "" Then 
					EMReadScreen next_page, 7, summ_row, 71
					If next_page = "MORE: +" Then 
						PF8
						summ_row = 1
					End If 
				End iF 
				summ_row = summ_row + 1
			Loop until summ_row = 23
			If inhibiting_error = TRUE then 
				MFIP_results = MFIP_results & MAXIS_footer_month & "/" & MAXIS_footer_year & " has an Inhibiting EDIT in STAT - resolve and rerun to generate results."
				inhibiting_error = FALSE
			End If 
		End If 
	Next
End Function 

FUNCTION date_array_generator(initial_month, initial_year, date_array)
	'defines an intial date from the initial_month and initial_year parameters
	initial_date = initial_month & "/1/" & initial_year
	'defines a date_list, which starts with just the initial date
	date_list = initial_date
	'This loop creates a list of dates
	Do
		If datediff("m", date, initial_date) = 1 then exit do		'if initial date is the current month plus one then it exits the do as to not loop for eternity'
		working_date = dateadd("m", 1, right(date_list, len(date_list) - InStrRev(date_list,"|")))	'the working_date is the last-added date + 1 month. We use dateadd, then grab the rightmost characters after the "|" delimiter, which we determine the location of using InStrRev
		date_list = date_list & "|" & working_date	'Adds the working_date to the date_list
	Loop until datediff("m", date, working_date) = 1	'Loops until we're at current month plus one

	'Splits this into an array
	date_array = split(date_list, "|")
End function

'===============================================================================================================================

EMConnect ""

developer_mode = FALSE 

Call MAXIS_case_number_finder(MAXIS_case_number)

If MAXIS_case_number <> "" Then
	Call Navigate_to_MAXIS_screen("STAT", "MEMB")

	EMReadScreen ref_number, 2, 4, 33
	EMReadScreen first_name, 12, 6, 63
	EMReadScreen last_name, 25, 6, 30

	first_name = Replace(first_name, "_", "")
	last_name = Replace(last_name, "_", "")
	client_name = first_name & " " & last_name & ""
	ref_number = ref_number & ""
	
	Call Navigate_to_MAXIS_screen ("STAT", "TIME")
	EMWriteScreen ref_number, 20, 76
	transmit
	EMReadScreen tanf_ext, 2, 17, 69
	tanf_ext = abs(tanf_ext)
	If tanf_ext < 60 Then 
		Extension_case = FALSE
		pre_sixty_checkbox = checked
		post_sixty_checkbox = unchecked
	Else 
		Extension_case = TRUE 
		post_sixty_checkbox = checked 
		pre_sixty_checkbox = checked 
	End IF 
End IF 

If client_name = "" then client_name = "Enter Ref Numb and press 'Reload Client Name'"

Do
	err_msg = ""
	Dialog fss_status_dialog
	Cancel_confirmation
	If ButtonPressed = get_client_name_button Then 
		ref_number = right("00" & ref_number, 2)
		Call Navigate_to_MAXIS_screen("STAT", "MEMB")
		EMWriteScreen ref_number, 20, 76
		transmit
		EMReadScreen first_name, 12, 6, 63
		EMReadScreen last_name, 25, 6, 30
		
		first_name = Replace(first_name, "_", "")
		last_name = Replace(last_name, "_", "")
		client_name = first_name & " " & last_name & ""
		
		Call Navigate_to_MAXIS_screen ("STAT", "TIME")
		EMWriteScreen ref_number, 20, 76
		transmit
		EMReadScreen tanf_ext, 2, 17, 69
		tanf_ext = abs(tanf_ext)
		If tanf_ext < 60 Then 
			Extension_case = FALSE
			pre_sixty_checkbox = checked
			post_sixty_checkbox = unchecked
		Else 
			Extension_case = TRUE 
			post_sixty_checkbox = checked 
			pre_sixty_checkbox = checked 
		End IF 
	End If 
	If universal_partipant_checkbox = unchecked AND new_imig_checkbox = unchecked AND age_sixty_checkbox = unchecked AND preg_checkbox = unchecked AND ill_incap_checkbox = unchecked AND care_of_ill_Incap_checkbox = unchecked AND child_under_one_checkbox = unchecked AND fam_violence_checkbox = unchecked AND Special_medical_checkbox = unchecked AND iq_test_checkbox = unchecked AND learning_disabled_checkbox = unchecked AND mentally_ill_checkbox = unchecked AND ssi_pending_checkbox = unchecked AND unemployable_checkbox = unchecked AND dev_delayed_checkbox = unchecked Then err_msg = err_msg & vbNewLine & "You must select a code to update."
	If MAXIS_case_number = "" Then err_msg = err_msg & vbNewLine & "You must enter a case number."
	If ref_number = "" Then err_msg = err_msg & vbNewLine & "Please enter the reference number of the person the SU is for."
	If pre_sixty_checkbox = checked AND post_sixty_checkbox = checked Then err_msg = err_msg & vbNewLine & "Case cannot be Pre-60 and Post-60, please select only one."
	If pre_sixty_checkbox = unchecked AND post_sixty_checkbox = unchecked Then err_msg = err_msg & vbNewLine & "Case must be Pre-60 or Post-60, please select one."
	If err_msg <> "" AND ButtonPressed <> get_client_name_button Then MsgBox "Please resolve to continue." & vbNewLine & err_msg
Loop until err_msg = ""

If age_sixty_checkbox = checked AND preg_checkbox = checked AND child_under_one_checkbox = checked AND universal_partipant_checkbox = checked then 
	developer_mode = TRUE
	MsgBox "You have enabled Developer Mode, if you did not intend to do this, stop the script and start it again."
End IF 
	
ref_number = right("00" & ref_number, 2)
Call Navigate_to_MAXIS_screen("STAT", "EMPS")
EMWriteScreen ref_number, 20, 76
transmit
EMReadScreen current_emps_status, 38, 15, 40
current_emps_status = trim(current_emps_status)

If current_emps_status = "20 (UP) Universal Participation" Then 
	ill_incap_new_checkbox = checked 
	rel_care_new_checkbox = checked
	unemployable_new_checkbox = checked 
	fvw_new_checkbox = checked 
	ssa_app_new_checkbox = checked
	child_under_1_new_checkbox = checked 
	imig_new_checkbox = checked 
	smc_new_checkbox = checked 
End If 

If pre_sixty_checkbox = checked then Extension_case = FALSE
If post_sixty_checkbox = checked then Extension_case = TRUE 

If child_under_one_checkbox = checked then 
	baby_on_case = FALSE
	Call Navigate_to_MAXIS_screen ("STAT", "PNLP")
	maxis_row = 3
	Do 
		EMReadScreen panel_name, 4, maxis_row, 5
		If panel_name = "MEMB" Then 
			EMReadScreen client_age, 2, maxis_row, 71
			If client_age = " 0" Then baby_on_case = TRUE
		End If 
		If panel_name = "MEMI" Then Exit Do
		maxis_row = maxis_row + 1
		If maxis_row = 20 Then 
			transmit
			maxis_row = 3
		End If 
	Loop until panel_name = "REVW"
	If baby_on_case = FALSE Then 
		no_baby_message = MsgBox("You are reporting an FSS status with a child under 12 months but there is no child under 1 listed in the household. You must add the baby fisrt." & vbNewLine & "Press Cancel to stop the script." & vbNewLine & "Press OK to continue the script if you have selected other FSS reasons.", vbOKCancel + VBAlert, "Review Child Under 12 Months Selection")
		If no_baby_message = VBCancel then cancel_confirmation
		child_under_one_checkbox = unchecked
	End If 
End If 

If child_under_one_checkbox = checked Then 
	Call Navigate_to_MAXIS_screen ("STAT", "EMPS")
	EMWriteScreen "X", 12, 39
	transmit
	emps_row = 7
	emps_col = 22
	Do
		EMReadScreen month_used, 2, emps_row, emps_col
		If month_used = "__" Then Exit Do
		EMReadScreen year_used, 4, emps_row, emps_col + 5
		emps_exemption_month_used = emps_exemption_month_used & "~" & month_used & "/" & year_used
		emps_col = emps_col + 11
		If emps_col = 66 Then 
			emps_col = 22
			emps_row = emps_row + 1
		End If 
	Loop Until emps_row = 10
	emps_exemption_month_used = right(emps_exemption_month_used, len(emps_exemption_month_used)-1)
	used_expemption_months_array = split(emps_exemption_month_used, "~")
	months_for_use = Join(used_expemption_months_array, ", ")
	number_of_months_available = 12 - (ubound(used_expemption_months_array) + 1) & ""
End If 

months_to_fill = "Enter the date of request and click 'Calculate' to fill this field."
detail_dialog_length = 45
If ill_incap_checkbox = checked Then detail_dialog_length = detail_dialog_length + 40
If care_of_ill_Incap_checkbox = checked Then detail_dialog_length = detail_dialog_length + 60
If iq_test_checkbox = checked OR learning_disabled_checkbox = checked OR mentally_ill_checkbox = checked OR dev_delayed_checkbox = checked OR unemployable_checkbox = checked Then detail_dialog_length = detail_dialog_length + 40
If fam_violence_checkbox = checked Then detail_dialog_length = detail_dialog_length + 40
If ssi_pending_checkbox = checked Then detail_dialog_length = detail_dialog_length + 40
If child_under_one_checkbox = checked Then detail_dialog_length = detail_dialog_length + 60
If new_imig_checkbox = checked Then detail_dialog_length = detail_dialog_length + 35
If Special_medical_checkbox = checked Then detail_dialog_length = detail_dialog_length + 50
y_pos_counter = 25

BeginDialog fss_code_detail, 0, 0, 440, detail_dialog_length, "Update FSS Information from the Status Update"
  Text 5, 10, 40, 10, "Date of SU"
  EditBox 50, 5, 50, 15, SU_date
  Text 120, 10, 40, 10, "ES Agency"
  EditBox 165, 5, 65, 15, es_agency
  Text 240, 10, 40, 10, "ES Worker"
  EditBox 285, 5, 110, 15, es_worker
  
  If ill_incap_checkbox = checked Then 
	  GroupBox 5, y_pos_counter, 430, 35, "Client Illness/Incapacity"
	  CheckBox 100, y_pos_counter, 30, 10, "New", ill_incap_new_checkbox
	  CheckBox 145, y_pos_counter, 35, 10, "Renew", ill_incap_renew_checkbox
	  CheckBox 195, y_pos_counter, 25, 10, "End", ill_incap_end_checkbox
	  Text 15, y_pos_counter + 20, 40, 10, "Start Date"
	  EditBox 75, y_pos_counter + 15, 50, 15, ill_incap_start_date
	  Text 135, y_pos_counter + 20, 35, 10, "End Date"
	  EditBox 180, y_pos_counter + 15, 50, 15, ill_incap_end_date
	  Text 260, y_pos_counter + 20, 70, 10, "Documentation with:"
	  CheckBox 335, y_pos_counter + 20, 25, 10, "ES", ill_incap_docs_with_es
	  CheckBox 370, y_pos_counter + 20, 50, 10, "Financial", ill_incap_docs_with_fas
	  
	  y_pos_counter = y_pos_counter + 40
  End If 
  
  If care_of_ill_Incap_checkbox = checked Then 
	  GroupBox 5, y_pos_counter, 430, 55, "Needed in Home to care for Family Member"
	  CheckBox 160, y_pos_counter, 30, 10, "New", rel_care_new_checkbox
	  CheckBox 205, y_pos_counter, 35, 10, "Renew", rel_care_renew_checkbox
	  CheckBox 255, y_pos_counter, 25, 10, "End", rel_care_end_checkbox
	  Text 15, y_pos_counter + 20, 95, 10, "Person in HH requiring care"
	  EditBox 115, y_pos_counter + 15, 25, 15, disa_HH_memb
	  Text 15, y_pos_counter + 40, 55, 10, "DISA Start Date"
	  EditBox 75, y_pos_counter + 35, 50, 15, rel_care_start_date
	  Text 135, y_pos_counter + 40, 35, 10, "End Date"
	  EditBox 180, y_pos_counter + 35, 50, 15, rel_care_end_date
	  Text 260, y_pos_counter + 40, 70, 10, "Documentation with:"
	  CheckBox 335, y_pos_counter + 40, 25, 10, "ES", rel_care_docs_with_es
	  CheckBox 370, y_pos_counter + 40, 50, 10, "Financial", rel_care_docs_with_fas
	  
	  y_pos_counter = y_pos_counter + 60
  End If
  
  If iq_test_checkbox = checked OR learning_disabled_checkbox = checked OR mentally_ill_checkbox = checked OR dev_delayed_checkbox = checked OR unemployable_checkbox = checked Then 
	  GroupBox 5, y_pos_counter, 430, 35, "Unemployable"
	  CheckBox 75, y_pos_counter, 30, 10, "New", unemployable_new_checkbox
	  CheckBox 120, y_pos_counter, 35, 10, "Renew", unemployable_renew_checkbox
	  CheckBox 170, y_pos_counter, 25, 10, "End", unemployable_end_checkbox
	  Text 15, y_pos_counter + 20, 55, 10, "Start Date on SU"
	  EditBox 75, y_pos_counter + 15, 50, 15, unemployable_start_date
	  Text 135, y_pos_counter + 20, 35, 10, "End Date"
	  EditBox 180, y_pos_counter + 15, 50, 15, unemployable_end_date
	  Text 260, y_pos_counter + 20, 70, 10, "Documentation with:"
	  CheckBox 335, y_pos_counter + 20, 25, 10, "ES", unemployable_docs_with_es
	  CheckBox 370, y_pos_counter + 20, 50, 10, "Financial", unemployable_docs_with_fas
	  
	  y_pos_counter = y_pos_counter + 40
  End If 
  
  If fam_violence_checkbox = checked Then 
	  GroupBox 5, y_pos_counter, 430, 35, "Family Violence Waiver"
	  CheckBox 100, y_pos_counter, 30, 10, "New", fvw_new_checkbox
	  CheckBox 145, y_pos_counter, 35, 10, "Renew", fvw_renew_checkbox
	  CheckBox 195, y_pos_counter, 25, 10, "End", fvw_end_checkbox
	  Text 15, y_pos_counter + 20, 55, 10, "Start Date "
	  EditBox 75, y_pos_counter + 15, 50, 15, fvw_start_date
	  Text 135, y_pos_counter + 20, 35, 10, "End Date"
	  EditBox 180, y_pos_counter + 15, 50, 15, fvw_end_date
	  
	  y_pos_counter = y_pos_counter + 40
  End If
  
  If ssi_pending_checkbox = checked Then 
	  GroupBox 5, y_pos_counter, 430, 35, "SSI/RSDI Pending"
	  CheckBox 85, y_pos_counter, 30, 10, "New", ssa_app_new_checkbox
	  CheckBox 130, y_pos_counter, 35, 10, "Renew", ssa_app_renew_checkbox
	  CheckBox 180, y_pos_counter, 25, 10, "End", ssa_app_end_checkbox
	  Text 15, y_pos_counter + 20, 55, 10, "Application Date"
	  EditBox 75, y_pos_counter + 10, 50, 15, ssa_app_date
	  Text 260, y_pos_counter + 20, 70, 10, "Documentation with:"
	  CheckBox 335, y_pos_counter + 20, 25, 10, "ES", ssa_app_docs_with_es
	  CheckBox 370, y_pos_counter + 20, 50, 10, "Financial", ssa_app_docs_with_fas
	  
	  y_pos_counter = y_pos_counter + 40
  End If 
  
  If child_under_one_checkbox = checked Then 
	  GroupBox 5, y_pos_counter, 430, 55, "Child Under 12 Months"
	  CheckBox 100, y_pos_counter, 30, 10, "New", child_under_1_new_checkbox
	  CheckBox 145, y_pos_counter, 35, 10, "Renew", child_under_1_renew_checkbox
	  CheckBox 195, y_pos_counter, 25, 10, "End", child_under_1_end_checkbox
	  Text 15, y_pos_counter + 20, 55, 10, "Request Date"
	  EditBox 75, y_pos_counter + 15, 50, 15, child_under_1_request_date
	  Text 275, y_pos_counter + 20, 85, 10, "Request made to:"
	  CheckBox 335, y_pos_counter + 20, 25, 10, "ES", child_under_1_at_es
	  CheckBox 370, y_pos_counter + 20, 50, 10, "Financial", child_under_1_at_fas
	  Text 130, y_pos_counter + 10, 65, 10, "Months used:"
	  Text 175, y_pos_counter + 10, 100, 30, months_for_use
	  Text 15, y_pos_counter + 40, 75, 10, "Months of exemption"
	  EditBox 90, y_pos_counter + 35, 280, 15, months_to_fill
	  ButtonGroup ButtonPressed
	    PushButton 380, y_pos_counter + 40, 35, 10, "Calculate", child_under_1_months_calculate

	  y_pos_counter = y_pos_counter + 60
  End If 
  
  If new_imig_checkbox = checked Then 
	  GroupBox 5, y_pos_counter, 430, 30, "Newly Arrived Immigrant"
	  CheckBox 100, y_pos_counter, 30, 10, "New", imig_new_checkbox
	  CheckBox 145, y_pos_counter, 35, 10, "Renew", imig_renew_checkbox
	  CheckBox 195, y_pos_counter, 25, 10, "End", imig_end_checkbox
	  Text 15, y_pos_counter + 15, 110, 10, "Spoken Language (SPL) from SU"
	  EditBox 130, y_pos_counter + 10, 25, 15, spl_listed
	  CheckBox 170, y_pos_counter + 15, 260, 10, "Check here to confirm that the SU indicates clt is enrolled in ELL/ESL classes", ell_confirm_checkbox
	  
	  y_pos_counter = y_pos_counter + 35
  End If 
  
  If Special_medical_checkbox = checked Then 
	  GroupBox 5, y_pos_counter, 430, 45, "Special Medical Criteria"
	  CheckBox 100, y_pos_counter, 30, 10, "New", smc_new_checkbox
	  CheckBox 145, y_pos_counter, 35, 10, "Renew", smc_renew_checkbox
	  CheckBox 195, y_pos_counter, 25, 10, "End", smc_end_checkbox
	  Text 20, y_pos_counter + 15, 100, 10, "Person in HH meeting Criteria"
	  EditBox 125, y_pos_counter + 10, 20, 15, smc_hh_memb
	  Text 155, y_pos_counter + 15, 60, 10, "Medical Criteria"
	  DropListBox 210, y_pos_counter + 10, 60, 40, "Select One ..."+chr(9)+"1 - Home-Health/Waiver Services"+chr(9)+"2 - Child who meets SED Criteria"+chr(9)+"3 - other Adult who meets SPMI", medical_criteria
	  Text 280, y_pos_counter + 15, 70, 10, "Date of Diagnosis"
	  EditBox 345, y_pos_counter + 10, 50, 15, smc_diagnosis_date
	  Text 265, y_pos_counter + 30, 70, 10, "Documentation with:"
	  CheckBox 340, y_pos_counter + 30, 25, 10, "ES", smp_docs_with_es
	  CheckBox 375, y_pos_counter + 30, 50, 10, "Financial", smc_docs_with_fas
	  
	  y_pos_counter = y_pos_counter + 50
  End If 
  
  Text 15, y_pos_counter, 85, 15, "Caregiver SU received for:"
  Text 100, y_pos_counter, 150, 15, client_name
  ButtonGroup ButtonPressed
	OkButton 330, y_pos_counter, 50, 15
	CancelButton 385, y_pos_counter, 50, 15
EndDialog



Do
	Do
		err_msg = ""
		dialog fss_code_detail
		cancel_confirmation
		If ButtonPressed = child_under_1_months_calculate Then 
			If IsDate(child_under_1_request_date) = TRUE Then 
				For add_month = 1 to number_of_months_available
					this_month = DatePart("m", DateAdd ("m", add_month, child_under_1_request_date))
					If len(this_month) = 1 Then this_month = "0" & this_month
					this_year = DatePart("yyyy", DateAdd("m", add_month, child_under_1_request_date))
					new_exemption_months = new_exemption_months & "~" & this_month & " / " & this_year
				Next
				new_exemption_months = right(new_exemption_months, len(new_exemption_months) - 1)
				new_exemption_months_array = split(new_exemption_months, "~")
				months_to_fill = Join(new_exemption_months_array, ", ")
			Else 
				MsgBox "You must enter a valid date to calculate the which months will have an exemption."
			End If 
		End If 
		If es_worker = "" Then err_msg = err_msg & vbNewLine & "** You must enter the name of the ES worker that completed the SU."
		If es_agency = "" Then err_msg = err_msg & vbNewLine & "** You must enter the ES Agency that provided the SU."
		If IsDate(SU_date) = FALSE Then err_msg = err_msg & vbNewLine & "** Enter the date of the Status Update." 
		If ill_incap_checkbox = checked Then 
			If IsDate(ill_incap_start_date) = False Then err_msg = err_msg & vbNewLine &"- You must enter a valid date for the start of client Ill/Incap. If one was not provided on the SU, an new SU is required."
			If ill_incap_docs_with_es = unchecked AND ill_incap_docs_with_fas = unchecked Then err_msg = err_msg & vbNewLine & "- Please indicate if verification of client's ill/incap are held in ES file or Financial File."
			If IsDate(ill_incap_end_date) = FALSE AND ill_incap_renew_checkbox = checked Then MsgBox "YOU HAVE NOT ENTERED AN END DATE FOR ILL/INCAP - since this is a renewal of this status the script will defaut the end date to six months from now."
			If ISDate(ill_incap_end_date) = FALSE AND ill_incap_end_checkbox = checked Then err_msg = err_msg & vbNewLine & "- If you are ending the ill/incap category, you must enter an end date."
		End If 
		If care_of_ill_Incap_checkbox = checked Then 
			If IsNumeric(disa_HH_memb) = False Then err_msg = err_msg & vbNewLine & "- List the reference number of the household member the client is needed in the home to care for. The person must be listed on the case, if the person has not yet been added to the case, cancel the script and do that first."
			If IsDate(rel_care_start_date) = False Then err_msg = err_msg & vbNewLine &"- You must enter a valid date for the start need to be at home. If one was not provided on the SU, an new SU is required."
			If rel_care_docs_with_es = unchecked AND rel_care_docs_with_fas = unchecked Then err_msg = err_msg & vbNewLine & "- Please indicate if verification of need to be at home for care of a family member is held in ES file or Financial File."
			If IsDate(rel_care_end_date) = FALSE AND rel_care_renew_checkbox = checked Then MsgBox "YOU HAVE NOT ENTERED AN END DATE FOR CARE OF ILL FAMILY MEMEBER - since this is a renewal of this status the script will defaut the end date to six months from now."
			If ISDate(rel_care_end_date) = FALSE AND rel_care_end_checkbox = checked Then err_msg = err_msg & vbNewLine & "- If you are ending the care of ill/incap family member category, you must enter an end date."
		End If 
		If iq_test_checkbox = checked OR learning_disabled_checkbox = checked OR mentally_ill_checkbox = checked OR dev_delayed_checkbox = checked OR unemployable_checkbox = checked Then 
			If IsDate(unemployable_start_date) = False Then err_msg = err_msg & vbNewLine &"- You must enter a valid date for the start of client determined to be unemployable. If one was not provided on the SU, an new SU is required."
			If ill_incap_docs_with_es = unchecked AND ill_incap_docs_with_fas = unchecked Then err_msg = err_msg & vbNewLine & "- Please indicate if verification of client's unemployability is held in ES file or Financial File."
		End If 
		If fam_violence_checkbox = checked Then 
			If IsDate(fvw_start_date) = False Then err_msg = err_msg & vbNewLine & "- Start date of Family Violence Waiver must be listed. If one was not provided on the SU, an new SU is required."
			If IsDate(fvw_end_date) = False Then err_msg = err_msg & vbNewLine & "- End date of Family Violence Waiver must be listed. If one was not provided on the SU, an new SU is required."
		End If 
		If ssi_pending_checkbox = checked Then 
			If IsDate(ssa_app_date) = False Then err_msg = err_msg & vbNewLine &"- You must enter a valid date of applicaiton for SSI/RSDI."
			If ssa_app_docs_with_es = unchecked AND ssa_app_docs_with_fas = unchecked Then err_msg = err_msg & vbNewLine & "- Please indicate if verification of client's SSI/RSDI Application is held in ES file or Financial File"
		End If 
		If child_under_one_checkbox = checked Then 
			If IsDate(child_under_1_request_date) = False Then err_msg = err_msg & vbNewLine &"- You must enter a the date the Child Under 12 Months Exemption was requested."
			If child_under_1_at_es = unchecked AND child_under_1_at_fas = unchecked Then err_msg = err_msg & vbNewLine & "- Please indicate if the request for Child Under 12 Months Exemption was requested to ES or Financial."
		End If 
		If new_imig_checkbox = checked Then 
			If ell_confirm_checkbox = unchecked Then err_msg = err_msg & vbNewLine & "- The SU must confirm that clt is enrolled in ELL Classes, if it does not a new SU is required."
			spl_listed = abs(spl_listed)
			If spl_listed >= 6 Then err_msg = err_msg & vbNewLine & "- Spoken Language (SPL) must be less than 6 to qualify for this FSS Coding. Connect with ES worker to clarify."
		End If
		If Special_medical_checkbox = checked Then 
			If IsNumeric(smc_hh_memb) = False Then err_msg = err_msg& vbNewLine & "- List the reference number of the household member who qualifies for Special Medical Criteria. The person must be listed on the case, if the person has not yet been added to the case, cancel the script and do that first."
			If IsDate(smc_diagnosis_date) = False Then MsgBox "No Diagnosis Date was listed, it is not required, but TANF Banked Months cannot be determined without it."
			If smp_docs_with_es = unchecked AND smc_docs_with_fas = unchecked Then err_msg = err_msg & vbNewLine & "- Please indicate if verification of need to be at home for care of a family member is held in ES file or Financial File."
			If medical_criteria = "Select One ..." Then err_msg = err_msg & "- Select a Medical Criteria from what is indicated on the SU."
		End If 
		If err_msg <> "" AND ButtonPressed <> child_under_1_months_calculate Then MsgBox "You must resolve to continue:" & vbNewLine & vbNewLine & err_msg
	Loop until err_msg = ""
	call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = false

Back_to_SELF
MAXIS_footer_month = right("00" & DatePart("m", SU_date), 2)
MAXIS_footer_year = right(DatePart ("yyyy", SU_date), 2)

If ill_incap_checkbox = checked Then 
	fvw_only = FALSE 
	fss_category_list = fss_category_list & "; Ill/Incap >60 Days"
	If ill_incap_end_date = "" Then 
		If ill_incap_new_checkbox = checked Then 
			ill_incap_end_date = DateAdd("m", 6, ill_incap_start_date)
		ElseIf ill_incap_renew_checkbox = checked Then 
			ill_incap_end_date = DateAdd("m", 6, date)
		End IF 
	End If 
	Call Navigate_to_MAXIS_screen ("STAT", "DISA")
	EMWriteScreen ref_number, 20, 76
	transmit
	start_month = right("00" & DatePart("m", ill_incap_start_date), 2)
	start_day = right("00" & DatePart("d", ill_incap_start_date), 2)
	start_year = DatePart("yyyy", ill_incap_start_date)
	
	end_month = right("00" & DatePart("m", ill_incap_end_date), 2)
	end_day = right("00" & DatePart("d", ill_incap_end_date), 2)
	end_year = DatePart("yyyy", ill_incap_end_date)
	EMReadScreen disa_exist, 4, 6, 53
	If ill_incap_new_checkbox = checked Then 
		fss_category_list = fss_category_list & " - NEW"
		If disa_exist <> "____" Then 
			EMReadScreen listed_end_month, 2, 6, 69
			EMReadScreen listed_end_day, 2, 6, 72
			EMReadScreen listed_end_year, 4, 6, 75
			If listed_end_year = "____" Then disa_info = "It appears there is an open ended DISA for this person." 
			listed_end_date = listed_end_month & "/" & listed_end_day & "/" & listed_end_year
			listed_end_date = cDate(listed_end_date)
			If listed_end_date > date Then disa_info = "It appears there is DISA with a future end date for this person." 
			If listed_end_date <= date Then disa_info = "It appears there is a DISA for this person that has already ended."
			change_disa_message = MsgBox(disa_info & vbNewLine & "Do you want the script to replace the dates on the panel with these?" & vbNewLine & vbNewLine & "Disability & Certification Begin: " & start_month & "/" & start_day & "/" & start_year & vbNewLine & "Disability & Certification End: " & end_month & "/" & end_day & "/" & end_year, vbYesNo + vbQuestion, "Update DISA?")
			If change_disa_message = VBNo Then panels_reviewed = panels_reviewed & "DISA for Memb " & ref_number & " & "
		End If 
		If disa_exist = "____" or change_disa_message = VBYes Then
			EMReadScreen numb_of_panels, 1, 2, 78
			IF numb_of_panels = "0" Then 
				EMWriteScreen "NN", 20, 79
				transmit
			Else
				PF9
			End IF 
			start_month = right("00" & DatePart("m", ill_incap_start_date), 2)
			start_day = right("00" & DatePart("d", ill_incap_start_date), 2)
			start_year = DatePart("yyyy", ill_incap_start_date)
			'Writing the Disability Begin Date'
			EMWriteScreen start_month, 6, 47
			EMWriteScreen start_day, 6, 50
			EMWriteScreen start_year, 6, 53
			'Writing the Certification Begin Date'
			EMWriteScreen start_month, 7, 47
			EMWriteScreen start_day, 7, 50
			EMWriteScreen start_year, 7, 53
			'Writing the Disability End Date'
			EMWriteScreen end_month, 6, 69
			EMWriteScreen end_day, 6, 72
			EMWriteScreen end_year, 6, 75
			'Writing the Certification End Date'
			EMWriteScreen end_month, 7, 69
			EMWriteScreen end_day, 7, 72
			EMWriteScreen end_year, 7, 75
			'Writing the verif code'
			EMWriteScreen "09", 11, 59
			EMWriteScreen "6", 11, 69
			transmit
			
			panels_updated = panels_updated & "DISA for Memb " & ref_number & " & "
		End If 
	Else	'If the category is being ended or renewed the action is the same - update the end date
		If ill_incap_renew_checkbox = checked Then fss_category_list = fss_category_list & " - RENEW"
		IF ill_incap_end_checkbox = checked Then fss_category_list = fss_category_list & " - ENDED"
		EMReadScreen numb_of_panels, 1, 2, 78
		IF numb_of_panels = "0" Then 
			MsgBox "The script is attempting to renew or end the Ill/Incap FSS Categort and a DISA panel does not exist. The script will case note the information but you must check DISA manually to generate correct results."
			panels_reviewed = panels_reviewed & "DISA for Memb " & ref_number & " & "
		Else
			PF9
		End IF 
		'Writing the Disability End Date'
		EMWriteScreen end_month, 6, 69
		EMWriteScreen end_day, 6, 72
		EMWriteScreen end_year, 6, 75
		'Writing the Certification End Date'
		EMWriteScreen end_month, 7, 69
		EMWriteScreen end_day, 7, 72
		EMWriteScreen end_year, 7, 75
		'Writing the verif code'
		EMWriteScreen "09", 11, 59
		EMWriteScreen "6", 11, 69
		transmit
		
		panels_updated = panels_updated & "DISA for Memb " & ref_number & " & "
	End If 
End If 

If care_of_ill_Incap_checkbox = checked Then 
	fvw_only = FALSE 
	fss_category_list = fss_category_list & "; Care of Ill/Incap Family Member"
	If rel_care_end_date = "" Then 
		If rel_care_new_checkbox = checked Then 
			rel_care_end_date = DateAdd("m", 6, rel_care_start_date)
		ElseIf rel_care_renew_checkbox = checked Then 
			rel_care_end_date = DateAdd("m", 6, date)
		End If 
	End IF 
	Call Navigate_to_MAXIS_screen ("STAT", "DISA")
	EMWriteScreen disa_HH_memb, 20, 76
	transmit
	start_month = right("00" & DatePart("m", rel_care_start_date), 2)
	start_day = right("00" & DatePart("d", rel_care_start_date), 2)
	start_year = DatePart("yyyy", rel_care_start_date)

	end_month = right("00" & DatePart("m", rel_care_end_date), 2)
	end_day = right("00" & DatePart("d", rel_care_end_date), 2)
	end_year = DatePart("yyyy", rel_care_end_date)
	EMReadScreen disa_exist, 4, 6, 53
	If rel_care_new_checkbox = checked Then 
		fss_category_list = fss_category_list & " - NEW"
		If disa_exist <> "____" Then 
			EMReadScreen listed_end_month, 2, 6, 69
			EMReadScreen listed_end_day, 2, 6, 72
			EMReadScreen listed_end_year, 4, 6, 75
			If listed_end_year = "____" Then disa_info = "It appears there is an open ended DISA for this person." 
			listed_end_date = listed_end_month & "/" & listed_end_day & "/" & listed_end_year
			'listed_end_date = cDate(listed_end_date)
			If listed_end_date > date Then disa_info = "It appears there is DISA with a future end date for this person." 
			If listed_end_date <= date Then disa_info = "It appears there is a DISA for this person that has already ended."
			change_disa_message = MsgBox(disa_info & vbNewLine & "Do you want the script to replace the dates on the panel with these?" & vbNewLine & vbNewLine & "Disability & Certification Begin: " & start_month & "/" & start_day & "/" & start_year & vbNewLine & "Disability & Certification End: " & end_month & "/" & end_day & "/" & end_year, vbYesNo + vbQuestion, "Update DISA?")
			If change_disa_message = vbNo Then panels_reviewed = panels_reviewed & "DISA for Memb " & disa_HH_memb & " & "
		End If 
		If disa_exist = "____" or change_disa_message = VBYes Then
			EMReadScreen numb_of_panels, 1, 2, 78
			IF numb_of_panels = "0" Then 
				EMWriteScreen "NN", 20, 79
				transmit
			Else
				PF9
			End IF 
			start_month = right("00" & DatePart("m", ill_incap_start_date), 2)
			start_day = right("00" & DatePart("d", ill_incap_start_date), 2)
			start_year = DatePart("yyyy", ill_incap_start_date)
			'Writing the Disability Begin Date'
			EMWriteScreen start_month, 6, 47
			EMWriteScreen start_day, 6, 50
			EMWriteScreen start_year, 6, 53
			'Writing the Certification Begin Date'
			EMWriteScreen start_month, 7, 47
			EMWriteScreen start_day, 7, 50
			EMWriteScreen start_year, 7, 53
			'Writing the Disability End Date'
			EMWriteScreen end_month, 6, 69
			EMWriteScreen end_day, 6, 72
			EMWriteScreen end_year, 6, 75
			'Writing the Certification End Date'
			EMWriteScreen end_month, 7, 69
			EMWriteScreen end_day, 7, 72
			EMWriteScreen end_year, 7, 75
			'Writing the verif code'
			EMWriteScreen "09", 11, 59
			EMWriteScreen "6", 11, 69
			transmit
			
			panels_updated = panels_updated & "DISA for Memb " & disa_HH_memb & " & "
		End If 
	Else 			'If the category is being ended or renewed the action is the same - update the end date
		IF rel_care_renew_checkbox = checked Then fss_category_list = fss_category_list & " - RENEW"
		IF rel_care_end_checkbox = checked Then fss_category_list = fss_category_list & " - ENDED"
		EMReadScreen numb_of_panels, 1, 2, 78
		IF numb_of_panels = "0" Then 
			MsgBox "The script is attempting to renew or end the Care of Ill/Incap Family Member FSS Category but a DISA panel does not exist for this person. The script will continue and case note, but you must check DISA manually to generate the correct approval."
			panels_reviewed = panels_reviewed & "DISA for Memb " & disa_HH_memb & " & "
		Else
			PF9
		End IF 
		'Writing the Disability End Date'
		EMWriteScreen end_month, 6, 69
		EMWriteScreen end_day, 6, 72
		EMWriteScreen end_year, 6, 75
		'Writing the Certification End Date'
		EMWriteScreen end_month, 7, 69
		EMWriteScreen end_day, 7, 72
		EMWriteScreen end_year, 7, 75
		'Writing the verif code'
		EMWriteScreen "09", 11, 59
		EMWriteScreen "6", 11, 69
		transmit
		
		panels_updated = panels_updated & "DISA for Memb " & disa_HH_memb & " & "
	End If 
End If 

If iq_test_checkbox = checked OR learning_disabled_checkbox = checked OR mentally_ill_checkbox = checked OR dev_delayed_checkbox = checked OR unemployable_checkbox = checked Then 
	fvw_only = FALSE 
	fss_category_list = fss_category_list & "; Unemployable"
	If iq_test_checkbox = checked Then fss_category_list = fss_category_list & " - IQ Tested < 80"
	If learning_disabled_checkbox = checked Then fss_category_list = fss_category_list & " - Learning Diabled"
	If mentally_ill_checkbox = checked Then fss_category_list = fss_category_list & " - Mentally Ill"
	If dev_delayed_checkbox = checked Then fss_category_list = fss_category_list & " - Developmentally Delayed"
	If unemployable_new_checkbox = checked Then fss_category_list = fss_category_list & " - NEW"
	if unemployable_renew_checkbox = checked Then fss_category_list = fss_category_list & " - RENEW"
	If unemployable_end_checkbox = checked Then fss_category_list = fss_category_list & " - ENDED"
	Call Navigate_to_MAXIS_screen ("STAT", "EMPS")
	PF9 
	If unemployable_checkbox = checked Then EMWriteScreen "UN", 11, 76
	If dev_delayed_checkbox = checked Then EMWriteScreen "DD", 11, 76
	If mentally_ill_checkbox = checked Then EMWriteScreen "MI", 11, 76
	If learning_disabled_checkbox = checked Then EMWriteScreen "LD", 11, 76
	IF iq_test_checkbox = checked Then EMWriteScreen "IQ", 11, 76
	transmit 
	panels_updated = panels_updated & "EMPS for Memb " & ref_number & " & "
End If 

If fam_violence_checkbox = checked Then 
	fss_category_list = fss_category_list & "; Family Violence Waiver"
	If fvw_new_checkbox = checked Then 
		fss_category_list = fss_category_list & " - NEW"
		MAXIS_footer_month = right("00" & DatePart("m", fvw_start_date), 2)
		MAXIS_footer_year = right(DatePart("yyyy", fvw_start_date), 2)
		Back_to_SELF
		Call Navigate_to_MAXIS_screen ("STAT", "MEMI")
		EMWriteScreen ref_number, 20, 76
		transmit
		PF9 
		EMWriteScreen "02", 17, 78
		EMWriteScreen MAXIS_footer_month, 18, 49
		EMWriteScreen MAXIS_footer_year, 18, 55
		transmit
		transmit
		panels_updated = panels_updated & "MEMI for Memb " & ref_number & " & "
		
		next_month = DateAdd("m", 1, date)
		next_mo = right("00" & DatePart("m", next_month) , 2)
		next_yr = right(DatePart("yyyy", next_month), 2)
		next_MAXIS_month = next_mo & "/" & next_yr
		Call Navigate_to_MAXIS_screen ("STAT", "TIME")
		EMWriteScreen ref_number, 20, 76
		transmit
		Do 
			If MAXIS_footer_month = "01" Then fvw_month_col = 15
			If MAXIS_footer_month = "02" Then fvw_month_col = 20
			If MAXIS_footer_month = "03" Then fvw_month_col = 25
			If MAXIS_footer_month = "04" Then fvw_month_col = 30
			If MAXIS_footer_month = "05" Then fvw_month_col = 35
			If MAXIS_footer_month = "06" Then fvw_month_col = 40
			If MAXIS_footer_month = "07" Then fvw_month_col = 45
			If MAXIS_footer_month = "08" Then fvw_month_col = 50
			If MAXIS_footer_month = "09" Then fvw_month_col = 55
			If MAXIS_footer_month = "10" Then fvw_month_col = 60
			If MAXIS_footer_month = "11" Then fvw_month_col = 65
			If MAXIS_footer_month = "12" Then fvw_month_col = 70
			For row = 5 to 16
				EMReadScreen find_year, 2, row, 11
				If MAXIS_footer_year = find_year Then 
					fvw_month_row = row
					Exit For 
				End If 
			Next
			EMReadScreen is_counted, 2, fvw_month_row, first_fvw_month_col
			If is_counted = "SS" OR is_counted = "SF" OR is_counted = "WS" OR is_counted = "WF" Then 
				If Extension_case = TRUE THEN 
					PF9 
					EMWriteScreen "Y0", fvw_month_row, fvw_month_col
				ElseIf Extension_case = FALSE Then 
					PF9 
					EMWriteScreen "WD", fvw_month_row, fvw_month_col
				End IF 
				counted_months_changed = counted_months_changed & " & " & MAXIS_footer_month & "/" & MAXIS_footer_year
			End If 
			Call month_change(1, MAXIS_footer_month, MAXIS_footer_year, MAXIS_footer_month, MAXIS_footer_year)
		Loop until MAXIS_footer_month & "/" & MAXIS_footer_year = next_MAXIS_month
		transmit
		panels_updated = panels_updated & "TIME for Memb " & ref_number & " & "
		EMReadScreen tanf_used, 3, 17, 69
		EMReadScreen ext_tanf_used, 19, 69
	ElseIf fvw_renew_checkbox = checked Then 
		fss_category_list = fss_category_list & " - RENEW"
		Call Navigate_to_MAXIS_screen ("STAT", "MEMI")
		panels_reviewed = panels_reviewed & "MEMI for Memb " & disa_HH_memb & " & "
		EMWriteScreen ref_number, 20, 76
		transmit
		EMReadScreen fvw_code, 2, 17, 78
		If fvw_code <> "02" Then 
			MsgBox "The script has attempted to confirm the Family Violence Waiver coding was correct in MEMI for the renewal of the Waiver." & vbNewLine & vbNewLine & "The TANF Exemption is NOT coded as '02'. The script will continue and will case note the renewal but you must review the MEMI and TIME panel manually"
			fvw_memi_error = TRUE
		End If 
	ElseIf fvw_end_date = checked Then 
		fss_category_list = fss_category_list & " - ENDED"
		MAXIS_footer_month = right("00" & DatePart("m", DateAdd("m", 1, fvw_end_date)), 2)
		MAXIS_footer_year = right(DatePart("yyyy", DateAdd("m", 1, fvw_end_date)), 2)
		Call Navigate_to_MAXIS_screen ("STAT", "MEMI")
		EMWriteScreen ref_number, 20, 76
		transmit
		PF9 
		EMWriteScreen "  ", 17, 78
		EMWriteScreen "  ", 18, 49
		EMWriteScreen "  ", 18, 55
		transmit
		transmit 
		CALL Navigate_to_MAXIS_screen ("STAT", "TIME")
		EMReadScreen tanf_used, 3, 17, 69
		EMReadScreen ext_tanf_used, 19, 69
		panels_updated = panels_updated & "TIME for Memb " & ref_number & " & "
	End If 		
End If 

tanf_used = trim(tanf_used)
ext_tanf_used = trim(ext_tanf_used)
If counted_months_changed <> "" Then counted_months_changed = right (counted_months_changed, len(counted_months_changed)-3)

Back_to_SELF
MAXIS_footer_month = right("00" & DatePart("m", SU_date), 2)
MAXIS_footer_year = right(DatePart ("yyyy", SU_date), 2)

If ssi_pending_checkbox = checked Then 
	fvw_only = FALSE 
	fss_category_list = fss_category_list & "SSI/RSDI Pending"
	IF ssa_app_new_checkbox = checked Then fss_category_list = fss_category_list & " - NEW"
	IF ssa_app_renew_checkbox = checked Then fss_category_list = fss_category_list &  " - RENEW"
	IF ssa_app_end_checkbox = checked  Then fss_category_list = fss_category_list & " - ENDED"
	MAXIS_footer_month = right("00" & DatePart("m", fvw_start_date), 2)
	MAXIS_footer_year = right(DatePart("yyyy", fvw_start_date), 2)
	Back_to_SELF
	Call Navigate_to_MAXIS_screen ("STAT", "PBEN")
	ssa_app_month = right("00" & DatePart("m", ssa_app_date), 2)
	ssa_app_day = right("00" & DatePart("d", ssa_app_date), 2)
	ssa_app_year = right(DatePart("yyyy", ssa_app_date), 2)
	pben_row = 8 
	Do 
		EMReadScreen pben_exist, 2, pben_row, 24
		If pben_exist - "__" Then 
			EMReadScreen numb_of_panels, 1, 2, 78
			IF numb_of_panels = "0" Then 
				EMWriteScreen "NN", 20, 79
				transmit
			Else
				PF9
			End IF 
			EMWriteScreen "01", pben_row, 24
			EMWriteScreen ssa_app_month, pben_row, 51
			EMWriteScreen ssa_app_day, pben_row, 54
			EMWriteScreen ssa_app_year, pben_row, 57
			EMWriteScreen "5", pben_row, 62
			EMWriteScreen "P", pben_row, 77
			
			EMWriteScreen "02", pben_row + 1, 24
			EMWriteScreen ssa_app_month, pben_row + 1, 51
			EMWriteScreen ssa_app_day, pben_row + 1, 54
			EMWriteScreen ssa_app_year, pben_row + 1, 57
			EMWriteScreen "5", pben_row + 1, 62
			EMWriteScreen "P", pben_row + 1, 77
			
			panels_updated = panels_updated & "PBEN for Memb " & ref_number & " & "
			Exit Do 
		Else 
			pben_row = pben_row + 1
		End If 
	Loop until pben_row = 12
	If pben_row = 12 Then replace_pben_message = MSGBox("It appears the PBEN Panel is full." & vbNewLine & vbNewLine & "The script can overwrite the first 2 lines with the pending SSI/RSDI application." & vbNewLine & vbNewLine & "If you agree to application information being entered on the first 2 lines, press 'Yes'", vbYesNo + vbAlert, "Update PBEN?")
	If replace_pben_message = vbYes Then 
		PF9
		EMWriteScreen "01", 8, 24
		EMWriteScreen ssa_app_month, 8, 51
		EMWriteScreen ssa_app_day, 8, 54
		EMWriteScreen ssa_app_year, 8, 57
		EMWriteScreen "5", 8, 62
		EMWriteScreen "P", 8, 77
		
		EMWriteScreen "02", 9, 24
		EMWriteScreen ssa_app_month, 9, 51
		EMWriteScreen ssa_app_day, 9, 54
		EMWriteScreen ssa_app_year, 9, 57
		EMWriteScreen "5", 9, 62
		EMWriteScreen "P", 9, 77
		
		panels_updated = panels_updated & "PBEN for Memb " & ref_number & " & "
	Else 
		panels_reviewed = panels_reviewed & "PBEN for Memb " & disa_HH_memb & " & "
	End If 
	transmit
	
	Call Navigate_to_MAXIS_screen ("STAT", "DISA")
	
	ssa_end_date = DateAdd(6, "m", ssa_app_date)
	ssa_end_month = right("00" & DatePart("m", ssa_end_date), 2)
	ssa_end_day = right("00" & DatePart("d", ssa_end_date), 2)
	ssa_end_year = DatePart("yyyy", ssa_end_date)
	EMWriteScreen ref_number, 20, 76
	transmit
	EMReadScreen numb_of_panels, 1, 2, 78
	IF numb_of_panels = "0" Then 
		EMWriteScreen "NN", 20, 79
		transmit
	Else
		PF9
	End IF 
	EMWriteScreen ssa_app_month, 6, 47
	EMWriteScreen ssa_app_day, 6, 50
	EMWriteScreen "20" & ssa_app_year, 6, 53
	
	EMWriteScreen ssa_app_month, 7, 47
	EMWriteScreen ssa_app_day, 7, 50
	EMWriteScreen "20" & ssa_app_year, 7, 53
	
	EMWriteScreen ssa_end_month, 6, 69
	EMWriteScreen ssa_end_day, 6, 72
	EMWriteScreen ssa_app_year, 6, 75
	
	EMWriteScreen ssa_end_month, 7, 69
	EMWriteScreen ssa_end_day, 7, 72
	EMWriteScreen ssa_app_year, 7, 75
	
	EMWriteScreen "06", 11, 59
	EMWriteScreen "6", 11, 69
	
	panels_updated = panels_updated & "DISA for Memb " & ref_number & " & "
	
	Transmit
End If 

Back_to_SELF
MAXIS_footer_month = right("00" & DatePart("m", SU_date), 2)
MAXIS_footer_year = right(DatePart ("yyyy", SU_date), 2)

If child_under_one_checkbox = checked Then 
	fvw_only = FALSE 
	fss_category_list = fss_category_list & "; Care of Child < 12 Months"
	If child_under_1_new_checkbox = checked Then 
		fss_category_list = fss_category_list & " - NEW"
		MAXIS_footer_month = left(used_expemption_months_array(0), 2)
		MAXIS_footer_year = right(used_expemption_months_array(0), 2)
		Call Navigate_to_MAXIS_screen ("STAT", "EMPS")
		EMWriteScreen ref_number, 20, 76
		transmit
		PF9
		EMWriteScreen "Y", 12, 76
		EMWriteScreen "X", 12, 39 
		transmit
		
		emps_row = 7
		emps_col = 22
		Do
			EMReadScreen month_used, 2, emps_row, emps_col
			If month_used = "__" Then Exit Do
			emps_col = emps_col + 11
			If emps_col = 66 Then 
				emps_col = 22
				emps_row = emps_row + 1
			End If 
		Loop Until emps_row = 10
		IF emps_row = 10 Then 
			MsgBox "It appears the client has used all of their Exempt Months. EMPS will need to be updated manually."
			PF3
			PF10
		Else 
			For each exempt_month in used_expemption_months_array
				EMWriteScreen left(exempt_month, 2), emps_row, emps_col
				EMWriteScreen right(exempt_month, 4), emps_row, emps_col + 5
				emps_col = emps_col + 11
				If emps_col = 66 Then 
					emps_col = 22
					emps_row = emps_row + 1
				End If 
			Next
			PF3
			transmit
			panels_updated = panels_updated & "EMPS for Memb " & ref_number & " & "
		End IF 
	ElseIF child_under_1_end_checkbox = checked Then 
		fss_category_list = fss_category_list & " - ENDED"
		MAXIS_footer_month = right("00" & DatePart("m", DateAdd("M", 1, SU_date)), 2)
		MAXIS_footer_year = right(DatePart ("yyyy", DateAdd("M", 1, SU_date)), 2)
		Call Navigate_to_MAXIS_screen ("STAT", "EMPS")
		EMWriteScreen ref_number, 20, 76
		transmit
		PF9
		EMWriteScreen "N", 12, 76
		EMWriteScreen "X", 12, 39 
		transmit
		
		emps_row = 7
		emps_col = 22
		Do while emps_row < 10
			EMReadScreen month_used, 2, emps_row, emps_col
			If month_used = MAXIS_footer_month Then
				EMReadScreen year_used, 2, emps_row, emps_col + 7
				If year_used = MAXIS_footer_year Then 
					Do 
						EMWriteScreen "__", emps_row, emps_col
						EMWriteScreen "____", emps_row, emps_col + 5
						exemption_months_for_future_use = exemption_months_for_future_use + 1
						emps_col = emps_col + 11
						If emps_col = 66 Then 
							emps_col = 22
							emps_row = emps_row + 1
						End If 
					Loop Until emps_row = 10
					panels_updated = panels_updated & "EMPS for Memb " & ref_number & " & "
				End IF 
			End If 
		Loop
		PF3 
		transmit
	End If 
End If

Back_to_SELF
MAXIS_footer_month = right("00" & DatePart("m", SU_date), 2)
MAXIS_footer_year = right(DatePart ("yyyy", SU_date), 2)

If new_imig_checkbox = checked Then
	fvw_only = FALSE 
 	fss_category_list = fss_category_list & "; Newly Arrived Immigrant"
	If imig_new_checkbox = checked OR imig_renew_checkbox = checked Then 
		fss_category_list = fss_category_list & " - NEW"
		Call Navigate_to_MAXIS_screen ("STAT", "IMIG")
		EMWriteScreen ref_number, 20, 76
		transmit
		EMReadScreen numb_of_panels, 1, 2, 78
		IF numb_of_panels = "0" Then 
			MsgBox "No IMIG Panel exists for this person. This coding cannot be completed for someone without an IMIG panel. The script will now end."
			script_end_procedure("")
		Else
			PF9
		End IF 
		EMWriteScreen "Y", 18, 56
		transmit
		panels_updated = panels_updated & "IMIG for Memb " & ref_number & " & "
	ElseIF imig_end_checkbox = checked Then 
		fss_category_list = fss_category_list & " - ENDED"
		Call Navigate_to_MAXIS_screen ("STAT", "IMIG")
		EMWriteScreen ref_number, 20, 76
		transmit
		EMReadScreen numb_of_panels, 1, 2, 78
		IF numb_of_panels = "0" Then 
			MsgBox "No IMIG Panel exists for this person. This coding cannot be completed for someone without an IMIG panel. The script will now end."
			script_end_procedure("")
		Else
			PF9
		End IF 
		EMWriteScreen "N", 18, 56
		transmit
		panels_updated = panels_updated & "IMIG for Memb " & ref_number & " & "
	End If 
End If

If IsDate(smc_diagnosis_date) = TRUE Then 
	MAXIS_footer_month = right ("00" & DatePart ("m",smc_diagnosis_date), 2)
	MAXIS_footer_year = right (DatePart("yyyy", smc_diagnosis_date), 2)
End IF 

If Special_medical_checkbox = checked Then 
	fvw_only = FALSE 
	fss_category_list = fss_category_list & "; Special Medical Criteria"
	If smc_new_checkbox = checked Then 
		fss_category_list = fss_category_list & " - NEW"
		'Find Correct footer month
		Call Navigate_to_MAXIS_screen ("STAT", "EMPS")
		EMWriteScreen ref_number, 20, 76
		transmit
		PF9
		Select Case medical_criteria
		Case "1 - Home-Health/Waiver Services"
			EMWriteScreen "1", 8, 76
		Case "2 - Child who meets SED Criteria"
			EMWriteScreen "2", 8, 76
		Case "3 - other Adult who meets SPMI"
			EMWriteScreen "3", 8, 76
		End Select 
		transmit
		panels_updated = panels_updated & "EMPS for Memb " & ref_number & " & "
		
		IF Extension_case = FALSE Then 
			next_month = DateAdd(1, "m", date)
			next_mo = right("00" & DatePart("m", next_month) , 2)
			next_yr = right(DatePart("yyyy", next_month), 2)
			next_MAXIS_month = next_mo & "/" & next_yr
			TANF_banked_month = MAXIS_footer_month
			TANF_banked_year = MAXIS_footer_year
			Call Navigate_to_MAXIS_screen ("STAT", "TIME")
			EMWriteScreen ref_number, 20, 76
			transmit
			Do 
				If TANF_banked_month = "01" Then smc_month_col = 15
				If TANF_banked_month = "02" Then smc_month_col = 20
				If TANF_banked_month = "03" Then smc_month_col = 25
				If TANF_banked_month = "04" Then smc_month_col = 30
				If TANF_banked_month = "05" Then smc_month_col = 35
				If TANF_banked_month = "06" Then smc_month_col = 40
				If TANF_banked_month = "07" Then smc_month_col = 45
				If TANF_banked_month = "08" Then smc_month_col = 50
				If TANF_banked_month = "09" Then smc_month_col = 55
				If TANF_banked_month = "10" Then smc_month_col = 60
				If TANF_banked_month = "11" Then smc_month_col = 65
				If TANF_banked_month = "12" Then smc_month_col = 70
				For row = 5 to 16
					EMReadScreen find_year, 2, row, 11
					If TANF_banked_year = find_year Then 
						smc_month_row = row
						Exit For 
					End If 
				Next
				EMReadScreen is_counted, 2, fvw_month_row, first_fvw_month_col
				If is_counted = "SF" OR is_counted = "WF" Then 
					EMWriteScreen "FM", smc_month_row, smc_month_col
					tanf_banked_months_coded = tanf_banked_months_coded + 1
					banked_months_changed = banked_months_changed & " & " & TANF_banked_month & "/" & TANF_banked_year
				ElseIF is_counted = "SS" OR is_counted = "WS" Then 
					EMWriteScreen "SM", smc_month_row, smc_month_col
					tanf_banked_months_coded = tanf_banked_months_coded + 1
					banked_months_changed = banked_months_changed & " & " & TANF_banked_month & "/" & TANF_banked_year
				End If 
				Call month_change(1, TANF_banked_month, TANF_banked_year, TANF_banked_month, TANF_banked_year)
			Loop until TANF_banked_month & "/" & TANF_banked_year = next_MAXIS_month
			transmit
			panels_updated = panels_updated & "TIME for Memb " & ref_number & " & "
			If banked_months_changed <> "" Then banked_months_changed = right(banked_months_changed, len(banked_months_changed)-3)
		End IF 
	ElseIf smc_renew_checkbox = checked Then 
		fss_category_list = fss_category_list & " - RENEW"
		Call Navigate_to_MAXIS_screen ("STAT", "EMPS")
		panels_reviewed = panels_reviewed & "EMPS for Memb " & disa_HH_memb & " & "
		EMWriteScreen ref_number, 20, 76
		transmit
		EMReadScreen smc_code, 1, 8, 76
		Select Case medical_criteria
		Case "1 - Home-Health/Waiver Services"
			IF smc_code <> "1" Then 
				PF9
				EMWriteScreen "1", 8, 76
				other_notes = "Special Medical Criteria changed from " & smc_code & " to 1 - Home-Health/Waiver Services; "
				transmit
			End If 
		Case "2 - Child who meets SED Criteria"
			If smc_code <> "2" Then 
				PF9
				EMWriteScreen "2", 8, 76
				other_notes = "Special Medical Criteria changed from " & smc_code & " to 2 - Child who meets SED Criteria; "
				transmit
			End If 
		Case "3 - other Adult who meets SPMI"
			If smc_code <> "3" Then 
				PF9
				EMWriteScreen "3", 8, 76
				other_notes = "Special Medical Criteria changed from " & smc_code & " to 3 - other Adult who meets SPMI; "
				transmit
			End If 
		End Select 
	ElseIf smc_end_checkbox = checked Then 
		fss_category_list = fss_category_list & " - ENDED"
		Call Navigate_to_MAXIS_screen ("STAT", "EMPS")
		EMWriteScreen ref_number, 20, 76
		transmit
		PF9
		EMWriteScreen "N", 8, 76
		transmit
		panels_updated = panels_updated & "EMPS for Memb " & ref_number & " & "
	End If 
End If 

inhibiting_error = FALSE 
month_to_start = right("00" & DatePart("m", date), 2) & ""
year_to_start = right(DatePart ("yyyy", date), 2) & ""

Call Navigate_to_MAXIS_screen ("STAT", "SUMM")
EMWriteScreen "BGTX", 20, 71
transmit
Call date_array_generator (month_to_start, year_to_start, date_array)
For Each version in date_array
	MAXIS_footer_month = right("00" & datepart("m", version), 2)
	MAXIS_footer_year = right(datepart("yyyy", version), 2)
	Do 
		Call Navigate_to_MAXIS_screen ("STAT", "REVW")
		EMReadScreen revw_panel_check, 4, 2, 46
	Loop until revw_panel_check = "REVW"
	If er_due <> TRUE Then 
		EMReadScreen er_code, 1, 7, 40
		Select Case er_code
		Case "_", "A"
			er_due = FALSE
		Case "I", "N"
			er_due = TRUE
			er_due_month = MAXIS_footer_month & "/" & MAXIS_footer_year
		End Select
	End If 
	If mont_due <> TRUE Then 
		Call Navigate_to_MAXIS_screen ("STAT", "MONT")
		EMReadScreen mont_code, 1, 11, 43
		Select Case mont_code
		Case "_", "A"
			mont_due = FALSE
		Case "I", "N"
			mont_due = TRUE
			mont_due_month = MAXIS_footer_month & "/" & MAXIS_footer_year
		End Select
	End If 
Next 

If er_due = TRUE Then notes_not_approved = notes_not_approved & "ER due for " & er_due_month & "; "
IF mont_due = TRUE Then notes_not_approved = notes_not_approved & "HRF due for " & mont_due_month & "; "
 

fss_category_list = right(fss_category_list, len(fss_category_list) - 1) & ""
If panels_updated <> "" Then panels_updated = left(panels_updated, len(panels_updated)-3)
If panels_reviewed <> "" Then panels_reviewed = left(panels_reviewed, len(panels_reviewed)-3)
If other_notes <> "" THEN other_notes = left(other_notes, len(other_notes)-1) & ""
If notes_not_approved <> "" Then notes_not_approved = left (notes_not_approved, len(notes_not_approved)-2)
Call Read_MFIP_Results(month_to_start, year_to_start, MFIP_results)

Do 
	err_msg = ""
	Dialog FSS_final_dialog
	Cancel_confirmation
	MAXIS_dialog_navigation
	If worker_signature = "" Then err_msg = err_msg & vbNewLine & "Sign your case note!"
	If results_approved_checkbox = unchecked AND not_approved_checkbox = unchecked Then err_msg = err_msg & vbNewLine & "You must indicate if you approved the new MFIP results or not."
	IF results_approved_checkbox = checked AND not_approved_checkbox = checked Then err_msg = err_msg & vbNewLine & "You must pick if you have approved the new MFIP results or not - it cannot be both."
	IF not_approved_checkbox = checked AND notes_not_approved = "" Then err_msg = err_msg & vbNewLine & "If you did not approve the new MFIP results, you must explain why the approval is not being done."
	If ButtonPressed = CASE_BGTX_button Then 
		err_msg = err_msg & "new results needed"
		Call Read_MFIP_Results(month_to_start, year_to_start, MFIP_results)
	End IF 
	If err_msg <> "" AND ButtonPressed <> CASE_BGTX_button Then MsgBox "** Resolve to continue **" & vbNewLine & vbNewLine & err_msg
Loop until err_msg = ""

IF child_under_one_checkbox = checked AND child_under_1_end_checkbox = unchecked Then 
	Call navigate_to_MAXIS_screen ("DAIL", "WRIT")
	
	last_exemption_month = left(used_expemption_months_array(ubound(used_expemption_months_array)), 2)
	last_exemption_year = right(used_expemption_months_array(ubound(used_expemption_months_array)), 2)
	EMWriteScreen last_exemption_month, 5, 18
	EMWriteScreen "01", 5, 21
	EMWriteScreen last_exemption_year, 5, 24
	Call Write_variable_in_TIKL ("Child under one year exemption to end this month. Case needs to be sent through background ")
	transmit
	EMReadScreen TIKL_verified, 4, 24, 2
	IF TIKL_verified = "    " Then 
		TIKL_verified = TRUE 
	ELSE
		TIKL_verified = FALSE
		MsgBox "Script could not write a TIKL for the end of Child Under 12 Months exemption. You will need to set the TIKL manually."
	End If
	PF3	
End If 

IF fam_violence_checkbox = checked Then 
	CALL start_a_blank_CASE_NOTE
	IF fvw_new_checkbox = checked Then 
		CALL write_variable_in_CASE_NOTE ("***** DOMESTIC VIOLENCE WAIVER *****")
		CALL write_bullet_and_variable_in_CASE_NOTE ("Effective Date", fvw_start_date)
	ElseIF fvw_renew_checkbox = checked Then 
		CALL write_variable_in_CASE_NOTE ("***** DVW RENEWED *****")
		CALL write_bullet_and_variable_in_CASE_NOTE ("Effective Date", fvw_start_date)
	ElseIF fvw_end_checkbox = checked Then 
		CALL write_variable_in_CASE_NOTE ("***** DVW ENDS *****")
		CALL write_bullet_and_variable_in_CASE_NOTE ("End Date", fvw_end_date)
	End IF 
	CALL write_bullet_and_variable_in_CASE_NOTE ("ES Worker", es_worker)
	CALL write_bullet_and_variable_in_CASE_NOTE ("ES Agency", es_agency)
	CALL write_variable_in_CASE_NOTE ("* All documentation needed for the waiver is with Employment Services, including advocate information and review details.")
	CALL write_bullet_and_variable_in_CASE_NOTE ("Months Changed due to Waiver", counted_months_changed)
	CALL write_bullet_and_variable_in_CASE_NOTE ("TANF Months Used", tanf_used)
	CALL write_bullet_and_variable_in_CASE_NOTE ("Extension Months Used", ext_tanf_used)
	CALL write_variable_in_CASE_NOTE ("---")
	IF results_approved_checkbox = checked Then CALL write_bullet_and_variable_in_CASE_NOTE ("MFIP Results Approved", MFIP_results)
	IF not_approved_checkbox = checked Then Call write_bullet_and_variable_in_CASE_NOTE ("New MFIP NOT Approved Due To", notes_not_approved)
	CALL write_variable_in_CASE_NOTE ("---")
	CALL write_variable_in_CASE_NOTE (worker_signature)
End IF 

IF ill_incap_new_checkbox = checked OR rel_care_new_checkbox = checked OR unemployable_new_checkbox = checked OR ssa_app_new_checkbox = checked OR child_under_1_new_checkbox = checked OR imig_new_checkbox = checked OR smc_new_checkbox = checked Then
	new_category = TRUE 
Else 
	new_category = FALSE
End If 

IF ill_incap_renew_checkbox = checked OR rel_care_renew_checkbox = checked OR unemployable_renew_checkbox = checked OR ssa_app_renew_checkbox = checked OR child_under_1_renew_checkbox = checked OR imig_renew_checkbox = checked OR smc_renew_checkbox = checked Then
	renew_category = TRUE 
Else 
	renew_category = FALSE
End If 

IF ill_incap_end_checkbox = checked OR rel_care_end_checkbox = checked OR unemployable_end_checkbox = checked OR ssa_app_end_checkbox = checked OR child_under_1_end_checkbox = checked OR imig_end_checkbox = checked OR smc_end_checkbox = checked Then
	end_category = TRUE 
Else 
	end_category = FALSE
End If 

If fvw_only = FALSE Then 
	CALL start_a_blank_CASE_NOTE
	IF new_category = TRUE AND renew_category = FALSE AND end_category = FALSE Then 
		CALL write_variable_in_CASE_NOTE ("**** FSS ELIGIBLE ****")
		CALL write_bullet_and_variable_in_CASE_NOTE ("Approved change to state funding effective", month_to_start & "/" & year_to_start)
	ElseIF end_category = TRUE AND new_category = FALSE AND renew_category = FALSE Then 
		CALL write_variable_in_CASE_NOTE ("**** FSS ENDED ****")
		CALL write_bullet_and_variable_in_CASE_NOTE ("Approved change to federal funding effective", month_to_start & "/" & year_to_start)
	Else 
		CALL write_variable_in_CASE_NOTE ("**** FSS EXTENDED ****")
		CALL write_bullet_and_variable_in_CASE_NOTE ("Approved continued state funding effective", month_to_start & "/" & year_to_start)
	End IF 
	CALL write_bullet_and_variable_in_CASE_NOTE ("Eligibility of Category", fss_category_list)
	CALL write_bullet_and_variable_in_CASE_NOTE ("ES Worker", es_worker)
	CALL write_bullet_and_variable_in_CASE_NOTE ("ES Agency", es_agency)
	If ill_incap_checkbox = checked Then 
		IF ill_incap_docs_with_es = checked Then CALL write_variable_in_CASE_NOTE ("* Documentation of clt Ill/Incap is with Employment Services.")
		IF ill_incap_docs_with_fas = checked Then CALL write_variable_in_CASE_NOTE ("* Documentation of clt Ill/Incap is with Financial Case File.")
	End If 
	If care_of_ill_Incap_checkbox = checked Then
		IF rel_care_docs_with_es = checked Then CALL write_variable_in_CASE_NOTE ("* Documentation of Ill/Incap HH Member is with Employment Services.")
		IF rel_care_docs_with_fas = checked Then CALL write_variable_in_CASE_NOTE ("* Documentation of Ill/Incap HH Member is with Financial Case File.")
		CALL write_variable_in_CASE_NOTE ("* Caregiver is required in the home to care for " & disa_HH_memb)
	End If 
	If iq_test_checkbox = checked OR learning_disabled_checkbox = checked OR mentally_ill_checkbox = checked OR dev_delayed_checkbox = checked OR unemployable_checkbox = checked Then
		IF unemployable_docs_with_es = checked Then CALL write_variable_in_CASE_NOTE ("* Documentation of Unemployability is with Employment Services.")
		IF unemployable_docs_with_fas = checked Then CALL write_variable_in_CASE_NOTE ("* Documentation of Unemployability is with Financial Case File.")
	End IF 
	If ssi_pending_checkbox = checked Then
		IF ssa_app_docs_with_es = checked Then CALL write_variable_in_CASE_NOTE ("* Documentation of Application for SSI/RSDI is with Employment Services.")
		IF ssa_app_docs_with_fas = checked Then CALL write_variable_in_CASE_NOTE ("* Documentation of Application for SSI/RSDI is with Financial Case File.")
	End If 
	If child_under_one_checkbox = checked Then
		IF child_under_1_at_es = checked Then CALL write_variable_in_CASE_NOTE ("* Request to take the Child Under 12 Months exemption was made to ES Worker.")
		IF child_under_1_at_fas = checked Then CALL write_variable_in_CASE_NOTE ("* Request to take the Child Under 12 Months exemption was made to FW.")
		IF TIKL_verified = TRUE Then CALL write_variable_in_CASE_NOTE ("* TIKL set to end the exemption and do a new MFIP approval when months are all used.")
		CALL write_bullet_and_variable_in_CASE_NOTE ("Months coded for Child < 12 Months Exemption", Join(used_expemption_months_array, ", "))
	End IF 
	If new_imig_checkbox = checked Then 
		CALL write_variable_in_CASE_NOTE ("* Documentation of particilation with ELL Classes and SPL is with Employment Services.")
	End If 
	If Special_medical_checkbox = checked Then 
		IF smc_docs_with_es = checked Then CALL write_variable_in_CASE_NOTE ("* Documentation of Special Medical Criteria is with Employment Services.")
		IF smc_docs_with_fas = checked Then CALL write_variable_in_CASE_NOTE ("* Documentation of Special Medical Criteria is with Financial Case File.")
		CALL write_variable_in_CASE_NOTE ("* Special Medical Criteria for Memb " & smc_hh_memb & " for " & medical_criteria)
		CALL write_bullet_and_variable_in_CASE_NOTE ("Date of Diagnosis", smc_diagnosis_date)
		CALL write_bullet_and_variable_in_CASE_NOTE ("Banked Months Changed on TIME", banked_months_changed)
	End IF 		
	CALL write_bullet_and_variable_in_CASE_NOTE ("STAT Panels Updated", panels_updated)
	CALL write_bullet_and_variable_in_CASE_NOTE ("STAT Panels Reviewed", panels_reviewed)
	CALL write_variable_in_CASE_NOTE ("---")
	IF results_approved_checkbox = checked Then CALL write_bullet_and_variable_in_CASE_NOTE ("MFIP Results Approved", MFIP_results)
	IF not_approved_checkbox = checked Then Call write_bullet_and_variable_in_CASE_NOTE ("New MFIP NOT Approved Due To", notes_not_approved)
	CALL write_variable_in_CASE_NOTE ("---")
	CALL write_variable_in_CASE_NOTE (worker_signature)
End If 

script_end_procedure("Success!")