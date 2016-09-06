'PLEASE NOTE: this script was designed to run off of the BULK - pull data into Excel script.
'As such, it might not work if ran separately from that.

'Required for statistical purposes==========================================================================================
name_of_script = "BULK - SPENDDOWN ERROR REPORT.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 13                      'manual run time in seconds
STATS_denomination = "C"       							'C is for each CASE
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

'This function is used to grab all active X numbers according to the supervisor X number(s) inputted 
FUNCTION create_array_of_all_active_x_numbers_by_supervisor(array_name, supervisor_array) 
	'Getting to REPT/USER 
	CALL navigate_to_MAXIS_screen("REPT", "USER") 


	'Sorting by supervisor 
	PF5 
	PF5 


	'Reseting array_name 
	array_name = "" 


	'Splitting the list of inputted supervisors... 
	supervisor_array = replace(supervisor_array, " ", "") 
	supervisor_array = split(supervisor_array, ",") 
	FOR EACH unit_supervisor IN supervisor_array 
		IF unit_supervisor <> "" THEN 
			'Entering the supervisor number and sending a transmit 
			CALL write_value_and_transmit(unit_supervisor, 21, 12) 


			MAXIS_row = 7 
			DO 
				EMReadScreen worker_ID, 8, MAXIS_row, 5 
				worker_ID = trim(worker_ID) 
				IF worker_ID = "" THEN EXIT DO 
				array_name = trim(array_name & " " & worker_ID) 
				MAXIS_row = MAXIS_row + 1 
				IF MAXIS_row = 19 THEN 
					PF8 
					EMReadScreen end_check, 9, 24,14
					If end_check = "LAST PAGE" Then Exit Do
					MAXIS_row = 7 
				END IF 
			LOOP 
		END IF 
	NEXT 
	'Preparing array_name for use... 
	array_name = split(array_name) 
END FUNCTION 



'DIALOGS----------------------------------------------------------------------
BeginDialog find_spenddowns_dialog, 0, 0, 221, 150, "Pull REPT data into Excel dialog"
  EditBox 85, 20, 130, 15, worker_number
  CheckBox 5, 95, 150, 10, "Check here to run this query county-wide.", all_workers_check
  ButtonGroup ButtonPressed
    OkButton 110, 130, 50, 15
    CancelButton 165, 130, 50, 15
  Text 5, 25, 65, 10, "Worker(s) to check:"
  Text 5, 110, 210, 20, "NOTE: running queries county-wide can take a significant amount of time and resources. This should be done after hours."
  Text 50, 5, 125, 10, "*** REPT ON MAXIS SPENDDOW ***"
  Text 5, 40, 210, 20, "Enter 7 digits of your workers' x1 numbers (ex: x######), separated by a comma."
  Text 5, 65, 210, 25, "** If a supervisor 'x1 number' is entered, the script will add the 'x1 numbers' of all workers listed in MAXIS under that supervisor number."
EndDialog

BeginDialog find_spenddowns_month_spec_dialog, 0, 0, 221, 170, "Pull REPT data into Excel dialog"
  EditBox 85, 20, 130, 15, worker_number
  DropListBox 130, 130, 80, 45, "ALL"+chr(9)+"January"+chr(9)+"February"+chr(9)+"March"+chr(9)+"April"+chr(9)+"May"+chr(9)+"June"+chr(9)+"July"+chr(9)+"August"+chr(9)+"September"+chr(9)+"October"+chr(9)+"November"+chr(9)+"December", revw_month_list
  ButtonGroup ButtonPressed
    OkButton 110, 150, 50, 15
    CancelButton 165, 150, 50, 15
  Text 50, 5, 125, 10, "*** REPT ON MAXIS SPENDDOW ***"
  Text 5, 25, 65, 10, "Worker(s) to check:"
  Text 5, 40, 210, 20, "Enter 7 digits of your workers' x1 numbers (ex: x######), separated by a comma."
  Text 5, 65, 210, 25, "** If a supervisor 'x1 number' is entered, the script will add the 'x1 numbers' of all workers listed in MAXIS under that supervisor number."
  CheckBox 5, 95, 150, 10, "Check here to run this query county-wide.", all_workers_check
  Text 5, 110, 210, 20, "NOTE: running queries county-wide can take a significant amount of time and resources. This should be done after hours."
  Text 5, 135, 120, 10, "Only pull cases with next review in:"
EndDialog

'THE SCRIPT-------------------------------------------------------------------------
'Determining specific county for multicounty agencies...
get_county_code

'Connects to BlueZone
EMConnect ""

one_month_only = FALSE 
MAXIS_footer_month = right("00" & datepart("m", date), 2)
MAXIS_footer_year = right("00" & datepart("yyyy", date), 2)

MsgBox MAXIS_footer_month & "/" & MAXIS_footer_year

'Shows dialog
'Dialog find_spenddowns_dialog
'If buttonpressed = cancel then stopscript

'Shows dialog
Dialog find_spenddowns_month_spec_dialog
If buttonpressed = cancel then stopscript

If revw_month_list <> "ALL" AND revw_month_list <> "" Then 
	one_month_only = TRUE 
	Select Case revw_month_list
		Case "January"
			month_selected = 1
		Case "February"
			month_selected = 2
		Case "March"
			month_selected = 3
		Case "April"
			month_selected = 4
		Case "May"
			month_selected = 5
		Case "June"
			month_selected = 6
		Case "July"
			month_selected = 7
		Case "August"
			month_selected = 8
		Case "September"
			month_selected = 9
		Case "October"
			month_selected = 10
		Case "November"
			month_selected = 11
		Case "December"
			month_selected = 12					
	End Select 
End If

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

'Checking for MAXIS
Call check_for_MAXIS(True)

'If all workers are selected, the script will go to REPT/USER, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
If all_workers_check = checked then
	call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)
Else
	x1s_from_dialog = split(worker_number, ",")	'Splits the worker array based on commas

	'formatting array
	For each x1_number in x1s_from_dialog
		x1_number = trim(ucase(x1_number))
		Call navigate_to_MAXIS_screen ("REPT", "USER")
		PF5
		PF5
		EMWriteScreen x1_number, 21, 12
		transmit
		EMReadScreen sup_id_check, 7, 7, 5
		IF sup_id_check <> "       " Then 
			supervisor_array = trim(supervisor_array & " " & x1_number)
		Else			
			If worker_array = "" then
				worker_array = trim(x1_number)		'replaces worker_county_code if found in the typed x1 number
			Else
				worker_array = worker_array & ", " & trim(ucase(x1_number)) 'replaces worker_county_code if found in the typed x1 number
			End if
		End If 
		PF3
	Next

	If supervisor_array <> "" Then 
		Call create_array_of_all_active_x_numbers_by_supervisor (more_workers_array, supervisor_array)
		workers_to_add = join(more_workers_array, ", ")
		If worker_array = "" then
			worker_array = workers_to_add		'replaces worker_county_code if found in the typed x1 number
		Else
			worker_array = worker_array & ", " & trim(ucase(workers_to_add)) 'replaces worker_county_code if found in the typed x1 number
		End if
	End If 
	
	'Split worker_array
	worker_array = split(worker_array, ", ")
End if

Const wrk_num   = 0
Const case_num  = 1
Const next_revw = 2
Const clt_name  = 3
Const ref_numb  = 4
Const clt_pmi   = 5
Const mobl_spdn = 6
Const spd_pd    = 7
Const hc_excess = 8
Const mmis_spdn = 9
Const add_xcl   = 10


Dim clts_with_spdwn_array()
ReDim clts_with_spdwn_array (3, 0)

Dim spenddown_error_array ()
ReDim spenddown_error_array (11, 0)

'Setting the variable for what's to come
excel_row = 2
hc_clt = 0

For each worker in worker_array
	back_to_self	'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason
	Call navigate_to_MAXIS_screen("rept", "actv")
	EMWriteScreen worker, 21, 13
	transmit
	EMReadScreen user_worker, 7, 21, 71		'
	EMReadScreen p_worker, 7, 21, 13
	IF user_worker = p_worker THEN PF7		'If the user is checking their own REPT/ACTV, the script will back up to page 1 of the REPT/ACTV

	'Skips workers with no info
	EMReadScreen has_content_check, 1, 7, 8
	If has_content_check <> " " then

		'Grabbing each case number on screen
		Do
			'Set variable for next do...loop
			MAXIS_row = 7

			'Checking for the last page of cases.
			EMReadScreen last_page_check, 21, 24, 2	'because on REPT/ACTV it displays right away, instead of when the second F8 is sent
			Do
				EMReadScreen MAXIS_case_number, 8, MAXIS_row, 12	'Reading case number
				EMReadScreen client_name, 21, MAXIS_row, 21			'Reading client name
				EMReadScreen next_revw_date, 8, MAXIS_row, 42		'Reading application date
				EMReadScreen HC_status, 1, MAXIS_row, 64			'Reading HC status

				'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and stops if we've seen this one before.
				If trim(MAXIS_case_number) <> "" and instr(all_case_numbers_array, MAXIS_case_number) <> 0 then exit do
				all_case_numbers_array = trim(all_case_numbers_array & " " & MAXIS_case_number)

				If MAXIS_case_number = "        " then exit do			'Exits do if we reach the end

				'Using if...thens to decide if a case should be added (status isn't blank or inactive and respective box is checked)
				If HC_status = "A" then 
					If one_month_only = TRUE Then 
						revw_month = abs(left(next_revw_date, 2))
						If revw_month = month_selected Then 
							ReDim Preserve clts_with_spdwn_array (3, hc_clt)
							clts_with_spdwn_array(wrk_num, hc_clt)   = worker
							clts_with_spdwn_array(case_num, hc_clt)  = MAXIS_case_number
							clts_with_spdwn_array(next_revw, hc_clt) = next_revw_date
							hc_clt = hc_clt + 1
						End If 
					Else 
						ReDim Preserve clts_with_spdwn_array (3, hc_clt)
						clts_with_spdwn_array(wrk_num, hc_clt)   = worker
						clts_with_spdwn_array(case_num, hc_clt)  = MAXIS_case_number
						clts_with_spdwn_array(next_revw, hc_clt) = next_revw_date
						hc_clt = hc_clt + 1
					End If 
				End If 


				MAXIS_row = MAXIS_row + 1
				STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
			Loop until MAXIS_row = 19
			PF8
		Loop until last_page_check = "THIS IS THE LAST PAGE"
	End if
next

spd_case = 0

For hc_case = 0 to UBound(clts_with_spdwn_array, 2)
	MAXIS_case_number = clts_with_spdwn_array(case_num, hc_case)
	Call navigate_to_MAXIS_screen ("ELIG", "HC")
	row = 8
	Do 
		EMReadScreen prog, 2, row, 28
		If prog = "MA" Then 
			EMWriteScreen "X", row, 26
			transmit
			Exit Do
		End if 
		row = row + 1
	Loop until row = 20  
	If row <> 20 Then 
		EMWriteScreen "X", 18, 3
		transmit
		
		row = 6
		Do
			EMReadScreen spd_type, 20, row, 39
			spd_type = trim(spd_type)
			If spd_type = "" Then Exit Do
			If spd_type <> "NO SPENDDOWN" Then 
				EMReadScreen reference, 2, row, 6
				EMReadScreen period, 13, row, 61
				EMReadScreen cname, 21, row, 10
				cname = trim(cname)
				If cname = "" Then EMReadScreen cname, 21, row - 1, 10
				cname = trim(cname)
				
				ReDim Preserve spenddown_error_array (11, spd_case)
				
				spenddown_error_array (wrk_num,   spd_case) = clts_with_spdwn_array(wrk_num, hc_case)
				spenddown_error_array (case_num,  spd_case) = MAXIS_case_number
				spenddown_error_array (next_revw, spd_case) = replace(clts_with_spdwn_array(next_revw, hc_case), " ", "/")
				spenddown_error_array (clt_name,  spd_case) = cname
				spenddown_error_array (ref_numb,  spd_case) = reference
				spenddown_error_array (mobl_spdn, spd_case) = spd_type
				spenddown_error_array (spd_pd,    spd_case) = period 
				
				spd_case = spd_case + 1
				
			End If 
			row = row + 1
		Loop until row = 19
	End If 	
Next

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Setting the first 4 col as worker, case number, name, and APPL date
ObjExcel.Cells(1, 1).Value = "WORKER"
objExcel.Cells(1, 1).Font.Bold = TRUE
ObjExcel.Cells(1, 2).Value = "CASE NUMBER"
objExcel.Cells(1, 2).Font.Bold = TRUE
ObjExcel.Cells(1, 3).Value = "REF NO"
objExcel.Cells(1, 3).Font.Bold = TRUE
ObjExcel.Cells(1, 4).Value = "NAME"
objExcel.Cells(1, 4).Font.Bold = TRUE
ObjExcel.Cells(1, 5).Value = "NEXT REVW DATE"
objExcel.Cells(1, 5).Font.Bold = TRUE
ObjExcel.Cells(1, 6).Value = "SPDWN ON MOBL"
objExcel.Cells(1, 6).Font.Bold = TRUE
ObjExcel.Cells(1, 7).Value = "HC OVERAGE"
objExcel.Cells(1, 7).Font.Bold = TRUE
ObjExcel.Cells(1, 8).Value = "MMIS SPDWN"
objExcel.Cells(1, 8).Font.Bold = TRUE

For spd_case = 0 to UBound(spenddown_error_array, 2)
	MAXIS_case_number = spenddown_error_array(case_num, spd_case)				'Setting the case number for global functions
	spenddown_error_array(add_xcl, spd_case) = TRUE 
	Call navigate_to_MAXIS_screen ("CASE", "PERS")
	row = 9
	Do 
		EMReadScreen person, 2, row, 3
		If person = spenddown_error_array(ref_numb, spd_case) Then 
			EMReadScreen hc_stat, 1, row, 61
			If hc_stat = "I" Then spenddown_error_array(add_xcl, spd_case) = FALSE 
			Exit Do 
		Else 
			row = row + 1
			If row = 18 Then 
				EMReadScreen next_page, 7, row, 3
				If next_page = "More: +" Then 
					PF8
					row = 9
				End If 
			End If 
		End If 
	Loop until row = 18
	IF spenddown_error_array(add_xcl, spd_case) = TRUE Then 
		Call navigate_to_MAXIS_screen ("ELIG", "HC")								'Need a closer look at HC
		EMWriteScreen "BSUM", 20, 71												'Navigating to the HC ELIG for the right clt
		EMWriteScreen spenddown_error_array (ref_numb, spd_case), 20, 76
		transmit
		EMReadScreen bsum_check, 4, 3, 57
		If bsum_check <> "BSUM" Then	
			EMWriteScreen "    ", 20, 71
			EMWriteScreen "  ", 20, 76
			transmit
			row = 8
			Do 
				EMReadScreen person, 2, row, 3
				If person = spenddown_error_array(ref_numb, spd_case) Then 
					Do 
						EMReadScreen prog, 2, row, 28
						If prog = "MA" Then 
							Call write_value_and_transmit("x", row, 26)
							Exit Do 
						ELSE
							row = row + 1
							EMReadScreen person, 2, row, 3
							If person <> "  "  Then Exit Do
						End If 
					Loop Until row = 20
					Exit Do 
				Else 
					row = row + 1
				End If 
			Loop until row = 20 
		End If 
		spd_amt = 0
		col = 18
		Do 
			EMReadScreen month_net_inc, 8, 15, col 
			EMReadScreen month_std_inc, 8, 16, col
			month_net_inc = trim(month_net_inc)
			If month_net_inc = "" Then month_net_inc = 0
			month_std_inc = trim(month_std_inc)
			If month_std_inc = "" Then month_std_inc = 0
			tot_net_inc = tot_net_inc + abs(month_net_inc)
			tot_std_inc = tot_std_inc + abs(trim(month_std_inc))
			col = col + 11
		Loop until col = 84
		spd_amt =  tot_net_inc - tot_std_inc
		If spd_amt < 0 Then spd_amt = 0
		spenddown_error_array(hc_excess, spd_case) = spd_amt
		If spd_amt = 0 Then spenddown_error_array(add_xcl,   spd_case) = TRUE 

		'If spenddown_error_array(add_xcl, spd_case) = TRUE Then 
			ObjExcel.Cells(excel_row, 1).Value = spenddown_error_array(wrk_num, spd_case)
			ObjExcel.Cells(excel_row, 2).Value = spenddown_error_array(case_num, spd_case)
			ObjExcel.Cells(excel_row, 3).Value = "Memb " & spenddown_error_array(ref_numb, spd_case)
			ObjExcel.Cells(excel_row, 4).Value = spenddown_error_array(clt_name, spd_case)
			ObjExcel.Cells(excel_row, 5).Value = spenddown_error_array(next_revw, spd_case)
			ObjExcel.Cells(excel_row, 6).Value = spenddown_error_array(mobl_spdn, spd_case) & " for " & spenddown_error_array(spd_pd, spd_case)
			ObjExcel.Cells(excel_row, 7).Value = spenddown_error_array(hc_excess, spd_case)

			excel_row = excel_row + 1
		'End If 
	End If 
	back_to_self
Next


'Query date/time/runtime info
objExcel.Cells(1, 9).Font.Bold = TRUE
objExcel.Cells(2, 9).Font.Bold = TRUE
ObjExcel.Cells(1, 9).Value = "Query date and time:"	'Goes back one, as this is on the next row
ObjExcel.Cells(1, 10).Value = now
ObjExcel.Cells(2, 9).Value = "Query runtime (in seconds):"	'Goes back one, as this is on the next row
ObjExcel.Cells(2, 10).Value = timer - query_start_time


'Autofitting columns
For col_to_autofit = 1 to 10
	ObjExcel.columns(col_to_autofit).AutoFit()
Next

'Logging usage stats
STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("")
