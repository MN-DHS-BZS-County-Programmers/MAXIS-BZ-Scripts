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

'===========================================================================================================================

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
	  Text 15, y_pos_counter + 20, 55, 10, "Start Date "
	  EditBox 75, y_pos_counter + 15, 50, 15, fvw_start_date
	  Text 135, y_pos_counter + 20, 35, 10, "End Date"
	  EditBox 180, y_pos_counter + 15, 50, 15, fvw_end_date
	  
	  y_pos_counter = y_pos_counter + 40
  End If
  
  If ssi_pending_checkbox = checked Then 
	  GroupBox 5, y_pos_counter, 430, 35, "SSI/RSDI Pending"
	  Text 15, y_pos_counter + 20, 55, 10, "Application Date"
	  EditBox 75, y_pos_counter + 10, 50, 15, ssa_app_date
	  Text 260, y_pos_counter + 20, 70, 10, "Documentation with:"
	  CheckBox 335, y_pos_counter + 20, 25, 10, "ES", ssa_app_docs_with_es
	  CheckBox 370, y_pos_counter + 20, 50, 10, "Financial", ssa_app_docs_with_fas
	  
	  y_pos_counter = y_pos_counter + 40
  End If 
  
  If child_under_one_checkbox = checked Then 
	  GroupBox 5, y_pos_counter, 430, 55, "Child Under 12 Months"
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
	  Text 15, y_pos_counter + 15, 110, 10, "Spoken Language (SPL) from SU"
	  EditBox 130, y_pos_counter + 10, 25, 15, spl_listed
	  CheckBox 170, y_pos_counter + 15, 260, 10, "Check here to confirm that the SU indicates clt is enrolled in ELL/ESL classes", ell_confirm_checkbox
	  
	  y_pos_counter = y_pos_counter + 35
  End If 
  
  If Special_medical_checkbox = checked Then 
	  GroupBox 5, y_pos_counter, 430, 45, "Special Medical Criteria"
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
		End If 
		If care_of_ill_Incap_checkbox = checked Then 
			If IsNumeric(disa_HH_memb) = False Then err_msg = err_msg& vbNewLine & "- List the reference number of the household member the client is needed in the home to care for. The person must be listed on the case, if the person has not yet been added to the case, cancel the script and do that first."
			If IsDate(rel_care_start_date) = False Then err_msg = err_msg & vbNewLine &"- You must enter a valid date for the start need to be at home. If one was not provided on the SU, an new SU is required."
			If rel_care_docs_with_es = unchecked AND rel_care_docs_with_fas = unchecked Then err_msg = err_msg & vbNewLine & "- Please indicate if verification of need to be at home for care of a family member is held in ES file or Financial File."
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
	If ill_incap_end_date = "" Then ill_incap_end_date = DateAdd("m", 6, ill_incap_start_date)
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
	End If 
	If disa_exist = "____" or change_disa_message = VBYes Then
		PF9
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
	End If 
End If 

If care_of_ill_Incap_checkbox = checked Then 
	If rel_care_end_date = "" Then rel_care_end_date = DateAdd("m", 6, rel_care_start_date)
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
	End If 
	If disa_exist = "____" or change_disa_message = VBYes Then
		PF9
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
	End If 
End If 

If iq_test_checkbox = checked OR learning_disabled_checkbox = checked OR mentally_ill_checkbox = checked OR dev_delayed_checkbox = checked OR unemployable_checkbox = checked Then 
	Call Navigate_to_MAXIS_screen ("STAT", "EMPS")
	PF9 
	If unemployable_checkbox = checked Then EMWriteScreen "UN", 11, 76
	If dev_delayed_checkbox = checked Then EMWriteScreen "DD", 11, 76
	If mentally_ill_checkbox = checked Then EMWriteScreen "MI", 11, 76
	If learning_disabled_checkbox = checked Then EMWriteScreen "LD", 11, 76
	IF iq_test_checkbox = checked Then EMWriteScreen "IQ", 11, 76
	transmit 
End If 

If fam_violence_checkbox = checked Then 
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
	next_month = DateAdd(1, "m", date)
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
End If 

Back_to_SELF
MAXIS_footer_month = right("00" & DatePart("m", SU_date), 2)
MAXIS_footer_year = right(DatePart ("yyyy", SU_date), 2)

If ssi_pending_checkbox = checked Then 
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
			PF9
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
	End If 
	transmit
	
	Call Navigate_to_MAXIS_screen ("STAT", "DISA")
	
	ssa_end_date = DateAdd(6, "m", ssa_app_date)
	ssa_end_month = right("00" & DatePart("m", ssa_end_date), 2)
	ssa_end_day = right("00" & DatePart("d", ssa_end_date), 2)
	ssa_end_year = DatePart("yyyy", ssa_end_date)
	EMWriteScreen ref_number, 20, 76
	transmit
	PF9
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
	
	Transmit
End If 

Back_to_SELF
MAXIS_footer_month = right("00" & DatePart("m", SU_date), 2)
MAXIS_footer_year = right(DatePart ("yyyy", SU_date), 2)

If child_under_one_checkbox = checked Then 
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
	End IF 
End If

Back_to_SELF
MAXIS_footer_month = right("00" & DatePart("m", SU_date), 2)
MAXIS_footer_year = right(DatePart ("yyyy", SU_date), 2)

If new_imig_checkbox = checked Then 
	Call Navigate_to_MAXIS_screen ("STAT", "IMIG")
	EMWriteScreen ref_number, 20, 76
	transmit
	PF9
	EMWriteScreen "Y", 18, 56
	transmit
End If

If Special_medical_checkbox = checked Then 
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
End If 

'use current emps status to determine if this is an es status change, add or ENd
'crate a dynamic dialog that asks for information based on the status selected in the first dialog 

script_end_proceedure("")
