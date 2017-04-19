' PURPOSE: This script allows users to FIAT the HC determination for clients that are active on GRH or MSA...policy change means that these clients are no
'			longer automatically eligible for MA eligibility.
' DESIGN...
'		1. The script must collect the MAXIS Case Number and select the individual to FIAT.
'		2. The script must collect income information for the individual...
'			2a. GROSS UNEARNED INCOME
'			2b. GROSS DEEMED UNEARNED INCOME
'			2c. EXCLUDED UNEARNED INCOME
'		3. The script must collect asset information for the individual, determine which is COUNTED, EXCLUDED, and UNAVAILABLE.
'   4. More stuff tbd



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

'these variables are needed to input the values of each individual amount to the ELIG/HC FIAT
DIM ttl_CASH_counted, ttl_CASH_excluded, ttl_CASH_unavail
DIM ttl_ACCT_counted, ttl_ACCT_excluded, ttl_ACCT_unavail
DIM ttl_SECU_counted, ttl_SECU_excluded, ttl_SECU_unavail
DIM ttl_CARS_counted, ttl_CARS_excluded, ttl_CARS_unavail
DIM ttl_REST_counted, ttl_REST_excluded, ttl_REST_unavail
DIM ttl_OTHR_counted, ttl_OTHR_excluded, ttl_OTHR_unavail
DIM ttl_BURY_counted, ttl_BURY_excluded, ttl_BURY_unavail
DIM ttl_SPON_counted, ttl_SPON_excluded, ttl_SPON_unavail

'these variables are needed for to input values to the FIAT of income

'this class is needed for to keep track of data for individual assets
class asset_object
	'variables...going to keep them public for to cut down on the work needed to manipulate
	public asset_panel
	public asset_amount
	public asset_type
	public asset_amount_dialog  ' used for display in the dialog
	public asset_type_dialog 	' used for the dialog...you'll see

	' function to read the amount for a specific asset
	public function read_asset_amount(len, row, col)
		EMReadScreen asset_amt, len, row, col
		asset_amt = replace(asset_amt, "_", "")
		asset_amt = trim(asset_amt)
		IF asset_amt = "" THEN asset_amt = 0
		IF asset_amt < 0 THEN 																' }
			MsgBox "ERROR: Asset found with negative balance. The script will now stop."  	' } 
			stopscript																		' } should probably just have object function reject negative balance
		END IF																				' }
		asset_amount = asset_amt
	end function	
	
	' function to read whether the asset is counted
	public function read_asset_counted(row, col)
		EMReadScreen asset_counted, 1, row, col
		IF asset_counted = "Y" THEN asset_type = "COUNTED"
		IF asset_counted = "N" OR asset_counted = "_" THEN asset_type = "EXCLUDED"
	end function
	
	' function to assign value to panel name
	public function set_asset_panel(panel_name)
		asset_panel = panel_name
	end function
	
	' function to re-set amount of the asset
	public function set_asset_amount(specified_amount)
		asset_amount = specified_amount
	end function	
	
	' function to re-set whether or not the asset is counted
	public function set_asset_type(user_selection)
		asset_type = user_selection
	end function
end class

' this class is going to be used for grabbing information from UNEA, JOBS, and BUSI
class income_object
	' variables for income objects
	public income_amt
	public hc_income_amt
	public retro_income_amt
	public prosp_income_amt
	public snap_pic_income_amt
	public grh_pic_income_amt
	public income_category			' "UNEARNED", "EARNED", "DEEMED UNEARNED", "DEEMED EARNED"
	public income_type				' type read from panel
	public income_frequency
	public income_start_date
	public income_end_date
	
	' === member functions for income object ===

	' private member function for identifying UNEA
	private sub are_we_at_unea 
		row = 1																																		' }
		col = 1																																		' }		
		EMSearch "(UNEA)", row, col																													' }		
		IF row = 0 THEN 																															' } safeguarding that the script				
			MsgBox "Invalid function application. The script cannot confirm you are on UNEA. The script will now stop.", vbCritical 				' }	finds UNEA					
			stopscript																																' }		
		END IF																																		' }
	end sub

	' private member function for identifying JOBS	
	private sub are_we_at_jobs
		row = 1
		col = 1
		EMSearch "(JOBS)", row, col
		IF row = 0 THEN 
			MsgBox "Invalid function application. The script cannot confirm you are on JOBS. The script will now stop.", vbCritical		' }	finds JOBS			
			stopscript																													' }		
		END IF																															' }
	end sub			

	' private member function for identifying BUSI
	private sub are_we_at_busi
		row = 1
		col = 1
		EMSearch "(BUSI)", row, col
		IF row = 0 THEN 
			MsgBox "Invalid function application. The script cannot confirm you are on BUSI. The script will now stop.", vbCritical		' }	finds BUSI											' }					
			stopscript																													' }		
		END IF																															' }
	end sub		
	
	public function set_income_category(specific_income_category)
		income_category = specific_income_category
	end function
	
	public sub read_income_type
		IF income_category = "" THEN MsgBox "No category has been set for this income type. The script will not perform optimally."
		row = 1
		col = 1
		EMSearch "Inc Type: ", row, col
		IF row <> 0 THEN 			' } Then we are on the JOBS panel...
			EMReadScreen specific_income_type, 1, row, col + 10
			IF specific_income_type = "J" THEN 
				specific_income_type = "WIOA"
			ELSEIF specific_income_type = "W" THEN 
				specific_income_type = "Wages"
			ELSEIF specific_income_type = "E" THEN 
				specific_income_type = "EITC"
			ELSEIF specific_income_type = "G" THEN 
				specific_income_type = "Experience Works"
			ELSEIF specific_income_type = "F" THEN 
				specific_income_type = "Federal Work Study" 
			ELSEIF specific_income_type = "S" THEN 
				specific_income_type = "State Work Study"
			ELSEIF specific_income_type = "O" THEN 
				specific_income_type = "Other"
			ELSEIF specific_income_type = "C" THEN 
				specific_income_type = "Contract Income"
			ELSEIF specific_income_type = "T" THEN 
				specific_income_type = "Training Program"
			ELSEIF specific_income_type = "P" THEN 
				specific_income_type = "Service Program"
			ELSEIF specific_income_type = "R" THEN 
				specific_income_type = "Rehab Program"
			END IF
		ELSE						' } THEN we are on either BUSI or UNEA
			row = 1
			col = 1
			EMSearch "Income Type: ", row, col
			IF income_category = "EARNED" OR income_category = "DEEMED EARNED" THEN 			' } THEN WE ARE ON BUSI
				EMReadScreen specific_income_type, 2, row, col + 13
				IF specific_income_type = "01" THEN 
					specific_income_type = "01 Farming"
				ELSEIF specific_income_type = "02" THEN 
					specified_income_type = "02 Real Estate"
				ELSEIF specific_income_type = "03" THEN 
					specific_income_type = "03 Home Product Sales"
				ELSEIF specific_income_type = "04" THEN 
					specific_income_type = "04 Other Sales"
				ELSEIF specific_income_type = "05" THEN 
					specific_income_type = "05 Personal Services"
				ELSEIF specific_income_type = "06" THEN 
					specific_income_type = "06 Paper Route"
				ELSEIF specific_income_type = "07" THEN 
					specific_income_type = "07 In-Home Daycare"
				ELSEIF specific_income_type = "08" THEN 
					specific_income_type = "08 Rental Income"
				ELSEIF specific_income_type = "09" THEN 
					specific_income_type = "09 Other"
				END IF
			ELSEIF income_category = "UNEARNED"	or income_category = "DEEMED UNEARNED" THEN 		' } THEN WE ARE ON UNEA
				EMReadScreen specific_income_type, 20, row, col + 13
				specific_income_type = trim(specific_income_type)
			END IF
		END IF
		income_type = specific_income_type
	end sub
	
	' member functions for reading from JOBS
	public sub read_jobs_for_hc
		are_we_at_jobs
		row = 1
		col = 1
		EMSearch "HC Income Estimate", row, col
		IF row = 0 THEN 
			row = 1
			col = 1
			EMSearch "_ HC Est", row, col
			CALL write_value_and_transmit("X", row, col)
		END IF
		EMReadScreen hc_jobs_amount, 8, 11, 63
		hc_jobs_amount = replace(hc_jobs_amount, "_", "")
		hc_jobs_amount = trim(hc_jobs_amount)
		IF hc_jobs_amount = "" THEN hc_jobs_amount = 0.00
		transmit
		hc_income_amt = hc_jobs_amount
	end sub
	
	' member functions for reading from UNEA	
	public sub read_unea_for_hc
		are_we_at_unea
		row = 1
		col = 1
		EMSearch "_ HC Income Estimate", row, col
		IF row <> 0 THEN CALL write_value_and_transmit("X", row, col)

		EMReadScreen hc_income_info, 8, 9, 65
		EMReadScreen hc_inc_est_pay_freq, 1, 10, 63
		hc_income_info = replace(hc_income_info, "_", "")
		hc_income_info = trim(hc_income_info)
		IF hc_income_info = "" THEN hc_income_info = 0.00

		transmit							' } to close the pop-up
		hc_income_amt = hc_income_info		' } assigning value
	end sub
	
	public sub read_unea_for_snap
		are_we_at_unea
		row = 1
		col = 1 
		EMSearch "SNAP Prospective Income", row, col
		IF row = 0 THEN 
			row = 1
			col = 1
			EMSearch "SNAP Prosp Inc", row, col
			CALL write_value_and_transmit("X", row, col - 2)
		END IF
			
		EMReadScreen snap_income_info, 8, 18, 56
		EMReadScreen snap_income_pay_frequency, 1, 5, 64
		snap_income_info = trim(snap_income_info)
		transmit
		snap_pic_income_amt = snap_income_info		
		income_frequency = snap_income_pay_frequency
	end sub
end class


'FUNCTION ======================================
FUNCTION calculate_assets(input_array)
	number_of_assets = ubound(input_array)
	
	'parralel array for user input
	redim parallel_array(number_of_assets, 1)	
	
	'determining height of dialog
	dialog_height = 115 + (20 * number_of_assets)
	
	DO		
		asset_counted_total = 0
		asset_excluded_total = 0
		asset_unavailable_total = 0
		'calculating the values of the totals...
		FOR i = 0 TO number_of_assets
			parallel_array(i, 0) = input_array(i).asset_amount
			parallel_array(i, 1) = input_array(i).asset_type
		
			IF input_array(i).asset_type = "COUNTED" THEN asset_counted_total = asset_counted_total + (input_array(i).asset_amount * 1)
			IF input_array(i).asset_type = "EXCLUDED" THEN asset_excluded_total = asset_excluded_total + (input_array(i).asset_amount * 1)
			IF input_array(i).asset_type = "UNAVAILABLE" THEN asset_unavailable_total = asset_unavailable_total + (input_array(i).asset_amount * 1)
		NEXT
	
     BeginDialog Dialog1, 0, 0, 385, dialog_height, "Asset Dialog"
       FOR i = 0 TO number_of_assets
     	Text 10, 15 + (i * 20), 55, 10, "Asset Panel:"
     	Text 75, 15 + (i * 20), 40, 10, input_array(i).asset_panel
     	Text 130, 15 + (i * 20), 35, 10, "Amount:"
     	EditBox 170, 10 + (i * 20), 65, 15, parallel_array(i, 0)
     	Text 250, 15 + (i * 20), 45, 10, "Counted?"
     	DropListBox 305, 10 + (i * 20), 65, 15, "COUNTED"+chr(9)+"EXCLUDED"+chr(9)+"UNAVAILABLE", parallel_array(i, 1)
       NEXT
       Text 10, dialog_height - 40, 60, 10, "COUNTED Total:"
       EditBox 70, dialog_height - 45, 50, 15, asset_counted_total & ""
       Text 130, dialog_height - 40, 60, 10, "EXCLUDED Total:"
       EditBox 195, dialog_height - 45, 50, 15, asset_excluded_total & ""
       Text 250, dialog_height - 40, 70, 10, "UNAVAILABLE Total:"
       EditBox 325, dialog_height - 45, 50, 15, asset_unavailable_total & ""
       ButtonGroup ButtonPressed
         OkButton 10, dialog_height - 20, 50, 15
         CancelButton 60, dialog_height - 20, 50, 15
         PushButton 320, dialog_height - 20, 55, 15, "CALCULATE", calculator_button	
     EndDialog

		DIALOG Dialog1
			cancel_confirmation
			IF ButtonPressed = calculator_button THEN
				'Changing the values of the 
				FOR i = 0 TO number_of_assets	
					CALL input_array(i).set_asset_amount(parallel_array(i, 0))
					CALL input_array(i).set_asset_type(parallel_array(i, 1))
				NEXT
			END IF
		
	LOOP UNTIL ButtonPressed = -1

END FUNCTION

' DIALOGS
BeginDialog case_number_dialog, 0, 0, 171, 65, "Enter Case Number"
  ButtonGroup ButtonPressed
    OkButton 65, 45, 50, 15
    CancelButton 115, 45, 50, 15
  Text 10, 15, 75, 10, "MAXIS Case Number"
  EditBox 95, 10, 70, 15, maxis_case_number
EndDialog

' ================ the script ====================
EMConnect ""

CALL check_for_MAXIS(true)		' checking for MAXIS

row = 1																' }
col = 1																' }
EMSearch "Case Nbr: ", row, col										' }
IF row <> 0 THEN 													' }
	EMReadScreen maxis_case_number, 8, row, col + 10				' }
	maxis_case_number = trim(maxis_case_number)						' }
	maxis_case_number = replace(maxis_case_number, "_", "")			' }
ELSEIF row = 0 THEN 												' }	looking for the MAXIS case number
	EMReadScreen at_self, 4, 2, 50									' }
	IF at_self = "SELF" THEN 										' }
		EMReadScreen maxis_case_number, 8, 18, 43					' }
		maxis_case_number = trim(maxis_case_number)					' }
		maxis_case_number = replace(maxis_case_number, "_", "")		' }
	END IF															' }
END IF																' }

DO																				' }
	DIALOG case_number_dialog													' }
		cancel_confirmation														' } initial dialog
		IF maxis_case_number = "" THEN MsgBox "Enter a MAXIS Case Number."		' }
LOOP UNTIL maxis_case_number <> "" 												' }

CALL check_for_MAXIS(false)	'checking for MAXIS again

DO
	' Getting the individual on the case
	CALL HH_member_custom_dialog(HH_member_array)
	IF ubound(HH_member_array) <> 0 THEN MsgBox "Please pick one and only one person for this."
LOOP UNTIL ubound(HH_member_array) = 0

FOR EACH person in HH_member_array
	hh_memb = left(person, 2)
	EXIT FOR
NEXT

' ==============
' ... ASSETS ...
' ==============
' VARIABLES WE NEED FOR THIS BIT...
'		asset_acct_amt X
'		asset_cash_amt X
'		asset_secu_amt X
'		asset_cars_amt X
'		asset_rest_amt X
'		asset_othr_amt X
'		asset_bury_amt 
'		asset_spon_amt 

' ==================
' ... ACCT PANEL ...
' ==================
num_assets = -1
redim asset_array(0)

'asset_acct_amt = 0													' }
CALL navigate_to_MAXIS_screen("STAT", "ACCT")						' }
EMWriteScreen hh_memb, 20, 76										' }
CALL write_value_and_transmit("01", 20, 79)							' }
EMReadScreen num_acct, 1, 2, 78										' }
IF num_acct <> "0" THEN 											' }
	Do																' }
		num_assets = num_assets + 1									' }
		redim preserve asset_array(num_assets)						' } STAT/ACCT
		set asset_array(num_assets) = new asset_object				' }
		asset_array(num_assets).set_asset_panel "ACCT"				' }			
		asset_array(num_assets).read_asset_amount 8, 10, 46			' }
		asset_array(num_assets).read_asset_counted 14, 64			' }
		transmit													' }
		EMReadScreen enter_a_valid, 21, 24, 2						' }
		IF enter_a_valid = "ENTER A VALID COMMAND" THEN EXIT DO		' }
	LOOP															' }
END IF

' ==================
' ... CASH PANEL ...
' ==================
CALL navigate_to_MAXIS_screen("STAT", "CASH")						' }	
CALL write_value_and_transmit(hh_memb, 20, 76)						' }
EMReadScreen number_of_cash, 1, 2, 78								' }
IF number_of_cash <> "0" THEN 										' }
	num_assets = num_assets + 1										' }
	redim preserve asset_array(num_assets)							' }
	set asset_array(num_assets) = new asset_object					' } STAT/CASH	
	asset_array(num_assets).set_asset_panel "CASH"					' }			
	asset_array(num_assets).read_asset_amount 8, 8, 39				' } 
	asset_array(num_assets).set_asset_type "COUNTED"				' }
END IF																' }

' ==================
' ... OTHR PANEL ...
' ==================	
CALL navigate_to_MAXIS_screen("STAT", "OTHR")							' }				
EMWriteScreen hh_memb, 20, 76											' }		
CALL write_value_and_transmit("01", 20, 79)								' }		
EMReadScreen number_of_other, 1, 2, 78									' }
IF number_of_other <> "0" THEN 											' }
	DO																	' }
		num_assets = num_assets + 1										' }
		redim preserve asset_array(num_assets)							' }
		set asset_array(num_assets) = new asset_object					' } STAT/OTHR
		asset_array(num_assets).set_asset_panel "OTHR"					' }
		asset_array(num_assets).read_asset_amount 10, 8, 40				' }
		asset_array(num_assets).read_asset_counted 12, 64				' }
		transmit														' }
		EMReadScreen enter_a_valid, 21, 24, 2							' }
		IF enter_a_valid = "ENTER A VALID COMMAND" THEN EXIT DO			' }
	LOOP																' }
END IF																	' }

' ==================
' ... SECU PANEL ...
' ==================
CALL navigate_to_MAXIS_screen("STAT", "SECU")							' } 
EMWriteScreen hh_memb, 20, 76											' }	
CALL write_value_and_transmit("01", 20, 79)								' }						
EMReadScreen number_of_secu, 1, 2, 78									' }							
IF number_of_secu <> "0" THEN 											' }			
	DO																	' }		
		num_assets = num_assets + 1										' }				
		redim preserve asset_array(num_assets)							' }	STAT/SECU	
		set asset_array(num_assets) = new asset_object					' }		
		CALL asset_array(num_assets).set_asset_panel("SECU")			' }	
		CALL asset_array(num_assets).read_asset_amount(8, 10, 52)		' }		
		CALL asset_array(num_assets).read_asset_counted(15, 64)			' }		
		transmit														' }			
		EMReadScreen enter_a_valid, 21, 24, 2							' }		
		IF enter_a_valid = "ENTER A VALID COMMAND" THEN EXIT DO			' }		
	LOOP																' }	
END IF																	' }

' ==================
' ... CARS PANEL ...
' ==================
CALL navigate_to_MAXIS_screen("STAT", "CARS")							' }
EMWriteScreen hh_memb, 20, 76											' }			
CALL write_value_and_transmit("01", 20, 79)								' }						
EMReadScreen number_of_cars, 1, 2, 78									' }			
IF number_of_cars <> "0" THEN 											' }				
	DO																	' }		
		num_assets = num_assets + 1										' }				
		redim preserve asset_array(num_assets)							' }					
		set asset_array(num_assets) = new asset_object					' } STAT/CARS					
		CALL asset_array(num_assets).set_asset_amount("CARS")			' }					
		CALL asset_array(num_assets).read_asset_amount(8, 9, 45)		' }						
		CALL asset_array(num_assets).read_asset_counted(15, 76)			' }					
		transmit														' }							
		EMReadScreen enter_a_valid, 21, 24, 2							' }										
		IF enter_a_valid = "ENTER A VALID COMMAND" THEN EXIT DO			' }									
	LOOP																' }							
END IF																	' }			

CALL calculate_assets(asset_array)

CALL check_for_MAXIS(false) 	' checking for MAXIS again again


' ==============
' ... Income ...
' ==============
num_income = -1
redim income_array(0)
' ====================
' ...earned income ...
' ====================

' ==================
' ... JOBS PANEL ...
' ==================
CALL navigate_to_MAXIS_screen("STAT", "JOBS")
EMWriteScreen hh_memb, 20, 76
CALL write_value_and_transmit("01", 20, 79)
EMReadScreen number_of_jobs, 1, 2, 78
IF number_of_jobs <> "0" THEN 
	DO
		num_income = num_income + 1
		redim preserve income_array(num_income)
		set income_array(num_income) = new income_object
		CALL income_array(num_income).set_income_category("EARNED")
		income_array(num_income).read_jobs_for_hc
		income_array(num_income).read_income_type
		transmit
		EMReadScreen enter_a_valid, 21, 24, 2
		IF enter_a_valid = "ENTER A VALID COMMAND" THEN EXIT DO		
	LOOP
END IF

' ==================
' ... BUSI PANEL ...
' ==================
CALL navigate_to_MAXIS_screen("STAT", "BUSI")
EMWriteScreen hh_memb, 20, 76
CALL write_value_and_transmit("01", 20, 79)

' =====================
' ...unearned income...
' =====================

' ==================
' ... UNEA PANEL ...
' ==================
CALL navigate_to_MAXIS_screen("STAT", "UNEA")
EMWriteScreen hh_memb, 20, 76
CALL write_value_and_transmit("01", 20, 79)
EMReadScreen number_of_unea, 1, 2, 78
IF number_of_unea <> "0" THEN 
	DO
		num_income = num_income + 1
		redim preserve income_array(num_income)
		set income_array(num_income) = new income_object
		CALL income_array(num_income).set_income_category("UNEARNED")
		income_array(num_income).read_unea_for_hc
		income_array(num_income).read_income_type
		transmit													' }
		EMReadScreen enter_a_valid, 21, 24, 2						' } navigating to the next UNEA
		IF enter_a_valid = "ENTER A VALID COMMAND" THEN EXIT DO		' }
	LOOP
END IF

FOR i = 0 to ubound(income_array)
	MsgBox "Income Category: " & income_array(i).income_category & vbNewLine & "Income Type: " & income_array(i).income_type & vbNewLine & "Income Budgetted for HC: " & income_array(i).hc_income_amt
NEXT

stopscript

CALL check_for_MAXIS(false) 	' checking for MAXIS again again

'The business of FIATing
CALL navigate_to_MAXIS_screen("ELIG", "HC")

'finding the correct household member
FOR hhmm_row = 8 to 19
	EMReadScreen hhmm_pers, 2, hhmm_row, 3
	IF hhmm_pers = hh_memb THEN EXIT FOR
NEXT

EMReadScreen ma_case, 4, hhmm_row, 26				' }
IF ma_case <> "_ MA" THEN msgbox "error"			' } looking to see that the client has MA

CALL write_value_and_transmit("X", hhmm_row, 26)	' navigating to BSUM for that client's MA

PF9													' } 
EMSendKey "04"										' } FIAT 500 for POLICY CHANGE
transmit											' } 

' updating budget method away from B
FOR i = 0 to 5
	EMWriteScreen "B", 13, (21 + (i * 11))
NEXT

' going through and updating the budget with income and assets
FOR i = 0 TO 5
	CALL write_value_and_transmit("X", 9, (21 + (i * 11)))			' pooting the X on the BUDGET field for that month in the benefit period

NEXT
