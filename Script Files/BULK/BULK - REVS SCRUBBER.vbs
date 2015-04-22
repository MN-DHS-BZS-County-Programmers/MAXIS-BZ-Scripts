time_array_30_min = "7:00"+chr(9)+"7:30"+chr(9)+"8:00"+chr(9)+"8:30"+chr(9)+"9:00"+chr(9)+"9:30"+chr(9)+"10:00"+chr(9)+"10:30"+chr(9)+"11:00"+chr(9)+"11:30"+chr(9)+"12:00"+chr(9)+"12:30"+chr(9)+"1:00"+chr(9)+"1:30"+chr(9)+"2:00"+chr(9)+"2:30"+chr(9)+"3:00"+chr(9)+"3:30"+chr(9)+"4:00"+chr(9)+"4:30"+chr(9)+"5:00"+chr(9)+"5:30"+chr(9)+"6:00"
appt_time_list = "15 mins"+chr(9)+"30 mins"+chr(9)+"45 mins"+chr(9)+"60 mins"

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - pull cases into Excel-revised"
start_time = timer

'LOADING ROUTINE FUNCTIONS FROM GITHUB REPOSITORY---------------------------------------------------------------------------
url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"

SET req = CreateObject("Msxml2.XMLHttp.6.0") 'Creates an object to get a URL
req.open "GET", url, FALSE	'Attempts to open the URL
req.send 'Sends request

IF req.Status = 200 THEN	'200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject") 'Creates an FSO
	Execute req.responseText 'Executes the script code
ELSE	'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
	MsgBox "Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_
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
	"URL: " & url
	script_end_procedure("Script ended due to error connecting to GitHub.")
END IF

FUNCTION create_calendar(month_to_use, month_array)
	'Generating a calendar
	'Determining the number of days in the calendar month.
	next_month = DateAdd("M", 1, month_to_use)
	next_month = DatePart("M", next_month) & "/01/" & DatePart("YYYY", next_month)
	num_of_days = DatePart("D", (DateAdd("D", -1, next_month)))

	ReDim month_array(num_of_days, 0)
	
	'=====Another dialog=====
	BeginDialog calendar_dlg, 0, 0, 280, 160, month_to_use
		Text 5, 10, 265, 10, "Please select days for the script to ignore."
		Text 5, 25, 265, 10, ("Month: " & DatePart("M", month_to_use) & "/" & DatePart("YYYY", month_to_use))
		y = 45
		FOR i = 1 TO num_of_days
			use_date = (DatePart("M", month_to_use) & "/" & i & "/" & DatePart("YYYY", month_to_use))
			x = 15 + (40 * (WeekDay(use_date) - 1))
			IF WeekDay(use_date) = 1 AND i <> 1 THEN y = y + 15
			IF WeekDay(use_date) = 1 OR WeekDay(use_date) = 7 THEN month_array(i, 0) = 1
			CheckBox x, y, 30, 10, i, month_array(i, 0)
		NEXT
		ButtonGroup ButtonPressed
		OkButton 175, 140, 50, 15
		CancelButton 225, 140, 50, 15
	EndDialog
	
	Dialog calendar_dlg
		IF ButtonPressed = 0 THEN stopscript
END FUNCTION


BeginDialog REVS_scrubber_initial_dialog, 0, 0, 136, 65, "REVS scrubber initial dialog"
  EditBox 65, 5, 60, 15, worker_number
  EditBox 65, 25, 60, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 25, 45, 50, 15
    CancelButton 80, 45, 50, 15
  Text 5, 10, 55, 10, "Worker number:"
  Text 5, 30, 60, 10, "Worker signature:"
EndDialog

BeginDialog REVS_scrubber_time_dialog, 0, 0, 141, 130, "REVS scrubber time dialog"
  DropListBox 70, 5, 60, 15, "Select one..."+chr(9)+time_array_30_min, first_appointment_listbox
  DropListBox 70, 25, 60, 15, "Select one..."+chr(9)+time_array_30_min, last_appointment_listbox
  DropListBox 80, 45, 50, 15, "Select one..."+chr(9)+appt_time_list, List3
  CheckBox 5, 70, 135, 10, "Duplicate appointments per time slot?", duplicate_appt_times
  EditBox 100, 85, 35, 15, appointments_per_time_slot
  ButtonGroup ButtonPressed
    OkButton 25, 105, 50, 15
    CancelButton 80, 105, 50, 15
  Text 5, 90, 90, 10, "Appointments per time slot:"
  Text 5, 10, 60, 10, "First appointment:"
  Text 5, 50, 65, 10, "Appointment Time:"
  Text 5, 30, 60, 10, "Last appointment:"
EndDialog


'-----THE SCRIPT, dawg
EMConnect ""

'Stopping the script is the user is running it before the 16th of the month.
day_of_month = DatePart("D", date)
IF day_of_month < 16 THEN script_end_procedure("You cannot run this script before the 16th of the month.")

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
Set objWorkbook = objExcel.Workbooks.Add() 
objExcel.DisplayAlerts = True

'formatting excel file
objExcel.cells(1, 1).Value = "CASE NUMBER"
objExcel.Cells(1, 1).Font.Bold = TRUE
objExcel.Cells(1, 2).Value = "Interview Date"
objExcel.cells(1, 2).Font.Bold = TRUE
objExcel.Cells(1, 3).Value = "Interview Time"
objExcel.cells(1, 3).Font.Bold = TRUE

'creating month plus 1 and plus 2
cm_plus_1 = dateadd("M", 1, date)
cm_plus_2 = dateadd("M", 2, date)
'creating a last day of recert variable
last_day_of_recert = Left(cm_plus_2, 2) & "/01/" & Right(cm_plus_2, 2)
last_day_of_recert = dateadd("D", -1, last_day_of_recert)

'Grabbing the worker's X number.
CALL find_variable("User: ", worker_number, 7)

DIALOG REVS_scrubber_initial_dialog
	IF ButtonPressed = 0 THEN stopscript

revs_month = DateAdd("M", 2, date)
next_month = DateAdd("M", 1, revs_month)
next_month = DatePart("M", next_month) & "/01/" & DatePart("YYYY", next_month)
num_of_days = DatePart("D", (DateAdd("D", -1, next_month)))

'Generating the calendar
ReDim month_array(num_of_days, 0)
CALL create_calendar(revs_month, month_array)

'Determining the appropriate times to set appointments.
DIALOG REVS_scrubber_time_dialog
	IF ButtonPressed = 0 THEN stopscript

CALL check_for_MAXIS(false)
back_to_SELF
current_month = DatePart("M", date)
	IF len(current_month) = 1 THEN current_month = "0" & current_month
current_year = DatePart("YYYY", date)
	current_year = right(current_year, 2)
	
revs_month = DateAdd("M", 2, date)
revs_year = DatePart("YYYY", revs_month)
	revs_year = right(revs_year, 2)
revs_month = DatePart("M", revs_month)
	IF len(revs_month) = 1 THEN revs_month = "0" & revs_month
	
EMWriteScreen current_month, 20, 43
EMWriteScreen current_year, 20, 46
transmit

CALL navigate_to_MAXIS_screen("REPT", "REVS")
EMWriteScreen revs_month, 20, 55
EMWriteScreen revs_year, 20, 58
transmit

EMReadScreen current_worker, 7, 21, 6
IF UCASE(current_worker) <> UCASE(worker_number) THEN
	EMWriteScreen worker_number, 21, 6
	transmit
END IF


Excel_row = 2
DO
	MAXIS_row = 7
	DO
		EMReadScreen case_number, 8, MAXIS_row, 6
		EMReadScreen SNAP_status, 1, MAXIS_row, 45
		
		'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and stops if we've seen this one before.
		IF trim(case_number) <> "" and instr(all_case_numbers_array, case_number) <> 0 then exit do 
		all_case_numbers_array = trim(all_case_numbers_array & " " & case_number)
		
		IF case_number = "        " then exit do
		
		'For some goofy reason the dash key shows up instead of the space key. No clue why. This will turn them into null variables.
		If cash_status = "-" then cash_status = ""
		If SNAP_status = "-" then SNAP_status = ""
		If HC_status = "-" then HC_status = ""
		
				'Using if...thens to decide if a case should be added (status isn't blank and respective box is checked)
		If trim(SNAP_status) = "N" or trim(SNAP_status) = "I" or trim(SNAP_status) = "U" then add_case_info_to_Excel = True
		'Adding the case to Excel
		If add_case_info_to_Excel = True then 
			ObjExcel.Cells(excel_row, 1).Value = case_number
			excel_row = excel_row + 1
		End if
		MAXIS_row = MAXIS_row + 1
		add_case_info_to_Excel = ""	'Blanking out variable
		case_number = ""			'Blanking out variable
	Loop until MAXIS_row = 19
	PF8
	EMReadScreen last_page_check, 21, 24, 2	'checking to see if we're at the end
Loop until last_page_check = "THIS IS THE LAST PAGE"


'Now the script needs to go back to the start of the Excel file and start assigning appointments.
'FOR EACH day that is not checked, start assigning appointments according to DatePart("N", appointment) because DatePart"N" is minutes. Once datepart("N") = last_appointment_time THEN the script needs to jump to the next day.
'

FOR i = 1 to num_of_days
	IF month_array(i, 0) = 0 THEN		'These are the dates that the user has determined the agency/unit/worker
		appointment_time = revs_month & "/" & i & "/" & revs_year 
	
	'do stuff
	
	END IF
NEXT


'This is the bit that Charles wrote that 

all_case_numbers_array = ""   'resetting array
excel_row = 2					'resetting excel row to start reading at the top 
DO 								'looping until it meets a blank excel cell without a case number
	recert_status = ""			'resetting recert status for each run through the loop/case number
	case_number = objExcel.cells(excel_row, 1).value
	IF case_number = "" THEN EXIT DO      'exiting do if it finds a blank cell on the case number column
	
	back_to_self
	
	EMwritescreen left(date, 2), 20, 43			'writing current month
	EMwritescreen right(date, 2), 20, 46		'writing current year
	transmit
	call navigate_to_screen("STAT", "REVW")
	
	ERRR_screen_check

	EMwritescreen "x", 5, 58
	Transmit
	
	DO											'looping to check if the SNAP REVW popup is on the screen
		EMReadScreen SNAP_popup_check, 7, 5, 43
	LOOP until SNAP_popup_check = "Reports"

	'The script will now read the CSR MO/YR and the Recert MO/YR
	EMReadScreen CSR_mo, 2, 9, 26
	EMReadScreen CSR_yr, 2, 9, 32
	EMReadScreen recert_mo, 2, 9, 64
	EMReadScreen recert_yr, 2, 9, 70
	
	'It then compares what it read to the previously established current month plus 2 and determine if it is a recert or not. If it is a recert we need an interview
	IF CSR_mo = left(cm_plus_2, 2) and CSR_yr = right(cm_plus_2, 2) THEN RECERT_STATUS = "NO"
	IF recert_mo = left(cm_plus_2, 2) and recert_yr = right(cm_plus_2, 2) THEN RECERT_STATUS = "YES"
	CALL navigate_to_screen("STAT", "ADDR")
	EMReadScreen area_code, 3, 17, 45
	EMReadScreen remaining_digits, 9, 17, 50
	phone_number = area_code & remaining_digits
	
	'If the case is up for a CSR in CM+2 then it will delete the row from the excel file.
	IF RECERT_STATUS = "NO" THEN
		SET objRange = objExcel.Cells(excel_row, 1).EntireRow
		objRange.Delete
		excel_row = excel_row - 1
	END If
	
	IF RECERT_STATUS = "YES" Then
		back_to_self
		CALL navigate_to_screen("SPEC", "MEMO")
		PF5
		EMReadScreen memo_display_check, 12, 2, 33
		If memo_display_check = "Memo Display" then script_end_procedure("You are not able to go into update mode. Did you enter in inquiry by mistake? Please try again in production.")
		'Checking for AREP
		row = 4
		col = 1
		EMSearch "ALTREP", row, col
		IF row > 4 THEN
			arep_row = row
			CALL navigate_to_screen("STAT", "AREP")
			EMReadscreen forms_to_arep, 1, 10, 45
			call navigate_to_screen("SPEC", "MEMO")
			PF5
		END IF
		'Checking for SWKR
		row = 4
		col = 1
		EMSearch "SOCWKR", row, col
		IF row > 4 THEN
			swkr_row = row
			call navigate_to_screen("STAT", "SWKR")
			EMReadscreen forms_to_swkr, 1, 15, 63
			call navigate_to_screen("SPEC", "MEMO")
			PF5
		END IF
		EMWriteScreen "x", 5, 10
		IF forms_to_arep = "Y" THEN EMWriteScreen "x", arep_row, 10
		IF forms_to_swkr = "Y" THEN EMWriteScreen "x", swkr_row, 10
		transmit
		
		EMSendKey("************************************************************")
		CALL write_new_line_in_SPEC_MEMO("Your SNAP case is set to recertify on " & Left(cm_plus_2, 2) & "/" & Right(cm_plus_2, 2) & ". An interview is required to process your application.")
		CALL write_new_line_in_SPEC_MEMO("")
		CALL write_new_line_in_SPEC_MEMO("Your phone interview is scheduled on " & interview_date & " at " & interview_time & ".")
		CALL write_new_line_in_SPEC_MEMO("We will be calling you at this number " & phone_number & ".")
		CALL write_new_line_in_SPEC_MEMO("If this date and/or time does not work, or if you would prefer an in-person interview, please call our office.")
		CALL write_new_line_in_SPEC_MEMO("")
		CALL write_new_line_in_SPEC_MEMO("If we do not hear from you by " & last_day_of_recert & " your SNAP case will close.")
		CALL write_new_line_in_SPEC_MEMO("")
		CALL write_new_line_in_SPEC_MEMO("A recertification packet has been sent to you, containing an application form. Please complete, sign, and date the form, and return it along with any required verifications by the date of your interview.")
		CALL write_new_line_in_SPEC_MEMO("")
		CALL write_new_line_in_SPEC_MEMO("Common items to be verified include income, housing costs, and medical costs. Some ways to verify items area included below.")
		CALL write_new_line_in_SPEC_MEMO("")
		CALL write_new_line_in_SPEC_MEMO("Income examples: paystubs, pension, unemployment, sponsor income etc.")
		CALL write_new_line_in_SPEC_MEMO("     Note: the agency will verify social security income.")
		CALL write_new_line_in_SPEC_MEMO("* Housing cost examples (if changed): rent/house payment receipt, mortgage, lease, etc.")
		CALL write_new_line_in_SPEC_MEMO("* Medical cost examples (if changed): prescription and medical bills, etc.")
		CALL write_new_line_in_SPEC_MEMO("")
		CALL write_new_line_in_SPEC_MEMO("Please contact the agency with any questions. Thank you.")
		PF4
		back_to_self
		
		CALL navigate_to_screen("CASE", "NOTE")
		PF9
		
		EMSendKey "***SNAP Recertification Interview Scheduled***"
		CALL write_variable_in_case_note("* A phone interview has been scheduled for " & interview_date & " at " & interview_time & ".")
		CALL write_variable_in_case_note("* Client phone: " & phone_number)
		If forms_to_arep = "Y" then call write_variable_in_case_note("* Copy of notice sent to AREP.")
		If forms_to_swkr = "Y" then call write_variable_in_case_note("* Copy of notice sent to Social Worker.")
		call write_variable_in_case_note("---")
		call write_variable_in_case_note(worker_signature)
	END IF
	
	
	excel_row = excel_row + 1
		
LOOP until objExcel.cells(excel_row, 1).Value = ""

	
'Formatting the columns to autofit after they are all finished being created. 
objExcel.Columns(1).autofit()
objExcel.Columns(2).autofit()
objExcel.Columns(3).autofit()
objExcel.Columns(4).autofit()

script_end_procedure("Success, the excel file now has all of the cases that have had interviews scheduled.")


