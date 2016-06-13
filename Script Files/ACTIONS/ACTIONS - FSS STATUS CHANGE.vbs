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
BeginDialog fss_status_dialog, 0, 0, 201, 335, "FSS Status Update"
  EditBox 60, 5, 65, 15, MAXIS_case_number
  EditBox 30, 25, 25, 15, ref_number
  EditBox 60, 25, 135, 15, client_name
  Text 5, 55, 185, 10, "Select all the ES Status Codes that applies to this client."
  CheckBox 5, 70, 155, 10, "Age 60 or Over - 21", age_sixty_checkbox
  CheckBox 5, 85, 155, 10, "Pregnant/Incapacitated - 22", preg_checkbox
  CheckBox 5, 100, 155, 10, "Ill/Incapacitated for more than 60 Days - 23", ill_incap_checkbox
  CheckBox 5, 115, 155, 10, "Care of Ill/Incap Family Member - 24", care_of_ill-Incap_checkbox
  CheckBox 5, 130, 155, 10, "Care of Child Under 12 Months - 25", child_under_one_checkbox
  CheckBox 5, 145, 155, 10, "Family Violence Waiver - 26", fam_violence_checkbox
  CheckBox 5, 160, 155, 10, "Special Medical Criteria - 27", Special_medical_checkbox
  CheckBox 5, 175, 155, 10, "IQ Tested - 28", iq_test_checkbox
  CheckBox 5, 190, 155, 10, "Learning Disabled - 29", learning_disabled_checkbox
  CheckBox 5, 205, 155, 10, "Mentally Ill - 30", mentally_ill_checkbox
  CheckBox 5, 220, 155, 10, "Developmentally Delayed - 31", dev_delayed_checkbox
  CheckBox 5, 235, 155, 10, "Unemployable - 32", unemployable_checkbox
  CheckBox 5, 250, 155, 10, "SSI/RSDI Pending - 33", ssi_pending_checkbox
  CheckBox 5, 265, 155, 10, "Newly Arrived Immigrant - 34", new_imig_checkbox
  CheckBox 5, 290, 155, 10, "Universal Participant - 20", universal_partipant_checkbox
  ButtonGroup ButtonPressed
    OkButton 90, 315, 50, 15
    CancelButton 145, 315, 50, 15
  Text 5, 10, 45, 10, "Case Number"
  Text 5, 30, 25, 10, "Client"
EndDialog

'===========================================================================================================================

EMConnect ""

Call MAXIS_case_number_finder (MAXIS_case_number)

Call Navigate_to_MAXIS_screen ("STAT", "MEMB")

EMReadScreen ref_number, 2, 4, 33
EMReadScreen first_name, 12, 6, 63
EMReadScreen last_name, 25, 6, 30

Replace(first_name, "_", "")
Replace(last_name, "_", "")
client_name = first_name & last_name & ""
ref_number = ref_number & ""

Dialog fss_status_dialog

ref_number = right("00" & ref_number, 2)
Navigate_to_MAXIS_screen ("STAT", "EMPS")
EMWriteScreen ref_number, 20, 76
transmit
EMReadScreen current_emps_status, 2, 15, 40

'use current emps status to determine if this is an es status change, add or ENd
'crate a dynamic dialog that asks for information based on the status selected in the first dialog 

script_end_proceedure ("")
