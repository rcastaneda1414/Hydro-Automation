#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <constants.au3>
#include <WinAPI.au3>
#include <Array.au3>
#include "CUIAutomation2.au3"
; #INDEX# =======================================================================================================================
; Title .........: UI automation helper functions
; AutoIt Version : 3.3.8.1
; Language ......: English (language independent)
; Description ...: Brings UI automation to AutoIt.
; Author(s) .....: junkew
; Copyright .....: Copyright (C) 2013. All rights reserved.
; License .......: GPL
;
; This program is distributed in the hope that it will be useful,
; but WITHOUT ANY WARRANTY; without even the implied warranty of
; MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
; See the Artistic License for more details.
;
; ===============================================================================================================================

;~ Some core default variables frequently to use in wrappers/helpers/global objects and pointers
; ===============================================================================================================================
Global $objUIAutomation          ;The main library core CUI automation reference
Global $oDesktop, $pDesktop		 ;Desktop will be frequently the starting point
Global $oUIElement, $pUIElement  ;Used frequently to get an element
Global $oTW, $pTW                ;Used frequently for treewalking
Local Const $UIA_tryMax=3		 ;Retry
Global $UIA_Vars                 ;Hold global UIA data in a dictionary object
Global $UIA_DefaultWaitTime=150  ;Frequently it makes sense to have a small waiting time to have windows rebuild, could be set to 0 if good synch is happening

;~ Loglevels that can be used in scripting following log4j defined standard
const $_UIA_Log_trace=10, $_UIA_Log_debug=20, $_UIA_Log_info=30, $_UIA_Log_warn =40, $_UIA_Log_error=50, $_UIA_Log_fatal=60

_UIA_Init()

; #FUNCTION# ====================================================================================================================
; Name...........: _UIA_Init
; Description ...: Initializes the basic stuff for the UI Automation library of MS
; Syntax.........: _UIA_Init()
; Parameters ....: none
; Return values .: Success      - Returns 1
;                  Failure		- Returns 0 and sets @error on errors:
;                  |@error=1     - UI automation failed
;                  |@error=2     - UI automation desktop failed
; Author ........:
; Modified.......:
; Remarks .......: None
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
func _UIA_Init()
;~ 	consolewrite("initializing")
	;~ The main object with acces to the windows automation api 3.0
	$objUIAutomation = ObjCreateInterface($sCLSID_CUIAutomation, $sIID_IUIAutomation, $dtagIUIAutomation)
	If IsObj($objUIAutomation) = 0 Then
;~ 		msgbox(1,"UI automation failed", "UI Automation failed",10)
		Return SetError(1, 0, 0)
	EndIf

	;~ Try to get the desktop as a generic reference/global for all samples
	$objUIAutomation.GetRootElement($pDesktop)
	$oDesktop = ObjCreateInterface($pDesktop, $sIID_IUIAutomationElement,$dtagIUIAutomationElement)
	If IsObj($oDesktop) = 0 Then
;~ 		msgbox(1,"UI automation desktop failed", "UI Automation desktop failed",10)
		Return SetError(2, 0, 0)
	EndIf
;~ 	consolewrite("At least it seems I have the desktop as a frequently used starting point" 	& "[" &_UIA_getPropertyValue($oDesktop, $UIA_NamePropertyId) & "][" &_UIA_getPropertyValue($oDesktop, $UIA_ClassNamePropertyId) & "]" & @CRLF)

;~ 	Dictionary object to store a lot of handy global data
	$UIA_vars=ObjCreate("Scripting.Dictionary")
	$UIA_Vars.comparemode=2 ; Text comparison case insensitive

;~ 	By default turn debugging on
	_UIA_setVar("Global.Debug",true)
	_UIA_setVar("Global.Highlight",true)

	_UIA_setVar("RTI.ACTIONCOUNT",0)

;~ Check if We can find configuration file(s)
	_UIA_LoadConfiguration()

	return 1
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _UIA_LoadConfiguration
; Description ...: Load all settings from a CFG file
; Syntax.........: _UIA_LoadConfiguration()
; Parameters ....: none
; Return values .: Success      - Returns 1
;                  Failure		- Returns 0 and sets @error on errors:
;                  |@error=1     - UI automation failed
; Author ........:
; Modified.......:
; Remarks .......: None
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================

func _UIA_LoadConfiguration()
	if fileexists("UIA.CFG") Then
		_UIA_LoadCFGFile("UIA.CFG")
	EndIf

;~ 		_UIA_Debug("Script name " & stringinstr(@scriptname)
EndFunc

func _UIA_loadCFGFile($strFname )
	Local $var
	$sections=IniReadSectionNames($strFName)

	If @error Then
		_UIA_DEBUG("Error occurred on reading " & $strFName & @CRLF)
	Else
;~ 		Load all settings into the dictionary
		For $i = 1 To $sections[0]
			$values=IniReadSection($strFName, $sections[$i])
			If @error Then
				_UIA_DEBUG("Error occurred on reading " & $strFName & @CRLF)
			Else
			;~ 		Load all settings into the dictionary
				For $j = 1 To $values[0][0]
					$strKey=$sections[$i] & "." & $values[$j][0]
					$strVal=$values[$j][1]

					if stringlower($strVal)="true" then $strVal=True
					if stringlower($strVal)="false" then $strVal=False
					if stringlower($strVal)="on" then $strVal=true
					if stringlower($strVal)="off" then $strVal=False

					if stringlower($strVal)="minimized" then $strVal=@SW_minimize
					if stringlower($strVal)="maximized" then $strVal=@SW_maximize
					if stringlower($strVal)="normal" then $strVal=@SW_restore

					$strval=stringreplace($strval,"%windowsdir%", @windowsdir)
					$strval=stringreplace($strval,"%programfilesdir%", @programfilesdir)

;~ 					_UIA_DEBUG("Key: [" & $strKey & "] Value: [" &  $strVal & "]" & @CRLF)
					_UIA_setvar($strKey,$strVal)
				Next
			EndIf
		Next
	endif
EndFunc

;~ Propertynames to match to numeric values
local $propertiesSupportedArray[10][3]=[ _
	["name",$UIA_NamePropertyId], _
	["title",$UIA_NamePropertyId], _
	["automationid",$UIA_AutomationIdPropertyId], _
	["classname", $UIA_ClassNamePropertyId], _
	["class", $UIA_ClassNamePropertyId], _
	["iaccessiblevalue",$UIA_LegacyIAccessibleValuePropertyId], _
	["iaccessiblechildId", $UIA_LegacyIAccessibleChildIdPropertyId], _
	["controltype", $UIA_ControlTypePropertyId,1], _
	["processid", $UIA_ProcessIdPropertyId], _
	["acceleratorkey", $UIA_AcceleratorKeyPropertyId] _
]

local $controlArray[41][3]= [ _
["UIA_AppBarControlTypeId",50040 ,"Identifies the AppBar control type. Supported starting with Windows 8.1."], _
["UIA_ButtonControlTypeId",50000 ,"Identifies the Button control type."], _
["UIA_CalendarControlTypeId",50001 ,"Identifies the Calendar control type."], _
["UIA_CheckBoxControlTypeId",50002 ,"Identifies the CheckBox control type."], _
["UIA_ComboBoxControlTypeId",50003 ,"Identifies the ComboBox control type."], _
["UIA_CustomControlTypeId",50025 ,"Identifies the Custom control type. For more information, see Custom Properties, Events, and Control Patterns."], _
["UIA_DataGridControlTypeId",50028 ,"Identifies the DataGrid control type."], _
["UIA_DataItemControlTypeId",50029 ,"Identifies the DataItem control type."], _
["UIA_DocumentControlTypeId",50030 ,"Identifies the Document control type."], _
["UIA_EditControlTypeId",50004 ,"Identifies the Edit control type."], _
["UIA_GroupControlTypeId",50026 ,"Identifies the Group control type."], _
["UIA_HeaderControlTypeId",50034 ,"Identifies the Header control type."], _
["UIA_HeaderItemControlTypeId",50035 ,"Identifies the HeaderItem control type."], _
["UIA_HyperlinkControlTypeId",50005 ,"Identifies the Hyperlink control type."], _
["UIA_ImageControlTypeId",50006 ,"Identifies the Image control type."], _
["UIA_ListControlTypeId",50008 ,"Identifies the List control type."], _
["UIA_ListItemControlTypeId",50007 ,"Identifies the ListItem control type."], _
["UIA_MenuBarControlTypeId",50010 ,"Identifies the MenuBar control type."], _
["UIA_MenuControlTypeId",50009 ,"Identifies the Menu control type."], _
["UIA_MenuItemControlTypeId",50011 ,"Identifies the MenuItem control type."], _
["UIA_PaneControlTypeId",50033 ,"Identifies the Pane control type."], _
["UIA_ProgressBarControlTypeId",50012 ,"Identifies the ProgressBar control type."], _
["UIA_RadioButtonControlTypeId",50013 ,"Identifies the RadioButton control type."], _
["UIA_ScrollBarControlTypeId",50014 ,"Identifies the ScrollBar control type."], _
["UIA_SemanticZoomControlTypeId",50039 ,"Identifies the SemanticZoom control type. Supported starting with Windows 8."], _
["UIA_SeparatorControlTypeId",50038 ,"Identifies the Separator control type."], _
["UIA_SliderControlTypeId",50015 ,"Identifies the Slider control type."], _
["UIA_SpinnerControlTypeId",50016 ,"Identifies the Spinner control type."], _
["UIA_SplitButtonControlTypeId",50031 ,"Identifies the SplitButton control type."], _
["UIA_StatusBarControlTypeId",50017 ,"Identifies the StatusBar control type."], _
["UIA_TabControlTypeId",50018 ,"Identifies the Tab control type."], _
["UIA_TabItemControlTypeId",50019 ,"Identifies the TabItem control type."], _
["UIA_TableControlTypeId",50036 ,"Identifies the Table control type."], _
["UIA_TextControlTypeId",50020 ,"Identifies the Text control type."], _
["UIA_ThumbControlTypeId",50027 ,"Identifies the Thumb control type."], _
["UIA_TitleBarControlTypeId",50037 ,"Identifies the TitleBar control type."], _
["UIA_ToolBarControlTypeId",50021 ,"Identifies the ToolBar control type."], _
["UIA_ToolTipControlTypeId",50022 ,"Identifies the ToolTip control type."], _
["UIA_TreeControlTypeId",50023 ,"Identifies the Tree control type."], _
["UIA_TreeItemControlTypeId",50024 ,"Identifies the TreeItem control type."], _
["UIA_WindowControlTypeId",50032 ,"Identifies the Window control type."] _
]

; #FUNCTION# ====================================================================================================================
; Name...........: _UIA_getControlName
; Description ...: Transforms the number of a control to a readable name
; Syntax.........: _UIA_getControlName($controlID)
; Parameters ....: $controlID
; Return values .: Success      - Returns string
;                  Failure		- Returns 0 and sets @error on errors:
;                  |@error=1     - UI automation failed
;                  |@error=2     - UI automation desktop failed
; Author ........:
; Modified.......:
; Remarks .......: None
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
func _UIA_getControlName($controlID)
	seterror(1,0,0)
	for $i=0 to ubound($controlArray)-1
		if ($controlArray[$i][1]=$controlID) then
			return $controlArray[$i][0]
		endIf
	Next
EndFunc

func _UIA_getControlID($controlName)
	local $tName
	$tName=stringupper($controlName)
	if stringleft($tname,3)<>"UIA" Then
		$tName="UIA_" & $controlname & "CONTROLTYPEID"
	endif
	seterror(1,0,0)
	for $i=0 to ubound($controlArray)-1
		if (stringupper($controlArray[$i][0])=$tName) then
			return $controlArray[$i][1]
		endIf
	Next
EndFunc

func _UIA_getPropertyIndex($propName)
	for $i=0 to ubound($propertiesSupportedArray,1)-1
		if stringlower($propertiesSupportedArray[$i][0])=stringlower($propName) Then
			return $i
		endif
	next
	_UIA_Debug("[FATAL] : property not found ")
EndFunc



; #FUNCTION# ====================================================================================================================
; Name...........: _UIA_setVar($varName, $varValue)
; Description ...: Just sets a variable to a certain value
; Syntax.........: _UIA_setVar("Global.UIADebug",True)
; Parameters ....: $varName  - A name for a variable
;				   $varValue - A value to assign to the variable
; Return values .: Success      - Returns 1
;                  Failure		- Returns 0 and sets @error on errors:
;                  |@error=1     - UI automation failed
;                  |@error=2     - UI automation desktop failed
; Author ........:
; Modified.......:
; Remarks .......: None
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
;~ Just set a value in a dictionary object
func _UIA_setVar($varName, $varValue)
	if $UIA_VARS.exists($varName) Then
		$UIA_Vars($varName)=$varvalue
	Else
		$UIA_Vars.add($varName, $VarValue)
	endif
EndFunc

Func _UIA_setVarsFromArray(ByRef $_array, $prefix="")
    If Not IsArray($_array) Then Return 0
    For $x = 0 To ubound($_array,1)-1
        _UIA_setVar($prefix & $_array[$x][0], $_array[$x][1])
    Next
EndFunc

Func _UIA_launchScript(ByRef $_scriptArray)
    If Not IsArray($_scriptArray) Then
		Return SetError(1,0,0)
	EndIf

    For $x = 0 To ubound($_scriptArray,1)-1
		if ($_scriptArray[$x][0]<>"") Then
			_UIA_action($_scriptArray[$x][0],$_scriptArray[$x][1],$_scriptArray[$x][2],$_scriptArray[$x][3],$_scriptArray[$x][4],$_scriptArray[$x][5])
		endif
	Next
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _UIA_getVar($varName)
; Description ...: Just returns a value as set before
; Syntax.........: _UIA_getVar("Global.UIADebug")
; Parameters ....: $varName  - A name for a variable
; Return values .: Success      - Returns the value of the variable
;                  Failure		- Returns "*** ERROR ***" and sets error to 1
; Author ........:
; Modified.......:
; Remarks .......: None
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
;~ Just get a value in a dictionary object
func _UIA_getVar($varName)
	if $UIA_VARS.exists($varName) Then
		return $UIA_Vars($varName)
	Else
		SetError(1) ;~ Not defined in repository
		return "*** ERROR ***" & $varname
	endif
EndFunc

Func _UIA_getVars2Array($prefix="")

;~ 	_UIA_debug($uia_vars.count-1 & @CRLF)
	$keys=$uia_vars.keys
	$it=$uia_vars.items
	for $i=0 to $uia_vars.count-1
;~ 		_UIA_debug("[" & $keys[$i] & "]:=[" & $it[$i] & "]"  & @CRLF)
	Next

EndFunc



; #FUNCTION# ====================================================================================================================
; Name...........: _UIA_getPropertyValue($obj, $id)
; Description ...: Just return a single property or if its an array string them together
; Syntax.........: _UIA_getPropertyValue
; Parameters ....: $obj - An UI Element object
;				   $id - A reference to the property id
; Return values .: Success      - Returns 1
;                  Failure		- Returns 0 and sets @error on errors:
;                  |@error=1     - UI automation failed
;                  |@error=2     - UI automation desktop failed
; Author ........:
; Modified.......:
; Remarks .......: None
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
;~ Just return a single property or if its an array string them together
func _UIA_getPropertyValue($obj, $id)
	local $tval
	local $tStr

	$obj.GetCurrentPropertyValue($Id,$tVal)
  	$tStr=""
	if isarray($tVal) Then
		for $i=0 to ubound($tval)-1
			$tStr=$tStr & $tVal[$i]
			if $i <> ubound($tVal)-1 Then
				$tStr=$tStr & ";"
			endif
		Next
		return $tStr
	endIf
	return $tVal
EndFunc



; ~ Just get all available properties for desktop/should work on all IUIAutomationElements depending on ControlTypePropertyID they work yes/no
 ; ~ Just make it a very long string name:= value pairs
func getAllPropertyValues($oUIElement)
	local $tStr, $tSeparator
	$tStr=""
	$tSeparator = @crLF  ; To make sure its not a value you normally will get back for values
	$tStr=$tStr & "UIA_AcceleratorKeyPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_AcceleratorKeyPropertyId) & $tSeparator ; Shortcut key for the element's default action.
	$tStr=$tStr & "UIA_AccessKeyPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_AccessKeyPropertyId) & $tSeparator ; Keys used to move focus to a control.
	$tStr=$tStr & "UIA_AriaPropertiesPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_AriaPropertiesPropertyId) & $tSeparator ; A collection of Accessible Rich Internet Application (ARIA) properties, each consisting of a name/value pair delimited by ‘-’ and ‘ ; ’ (for example, ("checked=true ; disabled=false").
	$tStr=$tStr & "UIA_AriaRolePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_AriaRolePropertyId) & $tSeparator ; ARIA role information.
	$tStr=$tStr & "UIA_AutomationIdPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_AutomationIdPropertyId) & $tSeparator ; UI Automation identifier.
	$tStr=$tStr & "UIA_BoundingRectanglePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_BoundingRectanglePropertyId) & $tSeparator ; Coordinates of the rectangle that completely encloses the element.
	$tStr=$tStr & "UIA_ClassNamePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_ClassNamePropertyId) & $tSeparator ; Class name of the element as assigned by the control developer.
	$tStr=$tStr & "UIA_ClickablePointPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_ClickablePointPropertyId) & $tSeparator ; Screen coordinates of any clickable point within the control.
	$tStr=$tStr & "UIA_ControllerForPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_ControllerForPropertyId) & $tSeparator ; Array of elements controlled by the automation element that supports this property.
	$tStr=$tStr & "UIA_ControlTypePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_ControlTypePropertyId) & $tSeparator ; Control Type of the element.
	$tStr=$tStr & "UIA_CulturePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_CulturePropertyId) & $tSeparator ; Locale identifier of the element.
	$tStr=$tStr & "UIA_DescribedByPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_DescribedByPropertyId) & $tSeparator ; Array of elements that provide more information about the element.
	$tStr=$tStr & "UIA_DockDockPositionPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_DockDockPositionPropertyId) & $tSeparator ; Docking position.
	$tStr=$tStr & "UIA_ExpandCollapseExpandCollapseStatePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_ExpandCollapseExpandCollapseStatePropertyId) & $tSeparator ; The expand/collapse state.
	$tStr=$tStr & "UIA_FlowsToPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_FlowsToPropertyId) & $tSeparator ; Array of elements that suggest the reading order after the corresponding element.
	$tStr=$tStr & "UIA_FrameworkIdPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_FrameworkIdPropertyId) & $tSeparator ; Underlying UI framework that the element is part of.
	$tStr=$tStr & "UIA_GridColumnCountPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_GridColumnCountPropertyId) & $tSeparator ; Number of columns.
	$tStr=$tStr & "UIA_GridItemColumnPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_GridItemColumnPropertyId) & $tSeparator ; Column the item is in.
	$tStr=$tStr & "UIA_GridItemColumnSpanPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_GridItemColumnSpanPropertyId) & $tSeparator ; number of columns that the item spans.
	$tStr=$tStr & "UIA_GridItemContainingGridPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_GridItemContainingGridPropertyId) & $tSeparator ; UI Automation provider that implements IGridProvider and represents the container of the cell or item.
	$tStr=$tStr & "UIA_GridItemRowPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_GridItemRowPropertyId) & $tSeparator ; Row the item is in.
	$tStr=$tStr & "UIA_GridItemRowSpanPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_GridItemRowSpanPropertyId) & $tSeparator ; Number of rows that the item spzns.
	$tStr=$tStr & "UIA_GridRowCountPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_GridRowCountPropertyId) & $tSeparator ; Number of rows.
	$tStr=$tStr & "UIA_HasKeyboardFocusPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_HasKeyboardFocusPropertyId) & $tSeparator ; Whether the element has the keyboard focus.
	$tStr=$tStr & "UIA_HelpTextPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_HelpTextPropertyId) & $tSeparator ; Additional information about how to use the element.
	$tStr=$tStr & "UIA_IsContentElementPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_IsContentElementPropertyId) & $tSeparator ; Whether the element appears in the content view of the automation element tree.
	$tStr=$tStr & "UIA_IsControlElementPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_IsControlElementPropertyId) & $tSeparator ; Whether the element appears in the control view of the automation element tree.
	$tStr=$tStr & "UIA_IsDataValidForFormPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_IsDataValidForFormPropertyId) & $tSeparator ; Whether the data in a form is valid.
	$tStr=$tStr & "UIA_IsDockPatternAvailablePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_IsDockPatternAvailablePropertyId) & $tSeparator ; Whether the Dock control pattern is available on the element.
	$tStr=$tStr & "UIA_IsEnabledPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_IsEnabledPropertyId) & $tSeparator ; Whether the control is enabled.
	$tStr=$tStr & "UIA_IsExpandCollapsePatternAvailablePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_IsExpandCollapsePatternAvailablePropertyId) & $tSeparator ; Whether the ExpandCollapse control pattern is available on the element.
	$tStr=$tStr & "UIA_IsGridItemPatternAvailablePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_IsGridItemPatternAvailablePropertyId) & $tSeparator ; Whether the GridItem control pattern is available on the element.
	$tStr=$tStr & "UIA_IsGridPatternAvailablePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_IsGridPatternAvailablePropertyId) & $tSeparator ; Whether the Grid control pattern is available on the element.
	$tStr=$tStr & "UIA_IsInvokePatternAvailablePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_IsInvokePatternAvailablePropertyId) & $tSeparator ; Whether the Invoke control pattern is available on the element.
	$tStr=$tStr & "UIA_IsItemContainerPatternAvailablePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_IsItemContainerPatternAvailablePropertyId) & $tSeparator ; Whether the ItemContainer control pattern is available on the element.
	$tStr=$tStr & "UIA_IsKeyboardFocusablePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_IsKeyboardFocusablePropertyId) & $tSeparator ; Whether the element can accept the keyboard focus.
	$tStr=$tStr & "UIA_IsLegacyIAccessiblePatternAvailablePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_IsLegacyIAccessiblePatternAvailablePropertyId) & $tSeparator ; Whether the LegacyIAccessible control pattern is available on the control.
	$tStr=$tStr & "UIA_IsMultipleViewPatternAvailablePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_IsMultipleViewPatternAvailablePropertyId) & $tSeparator ; Whether the pattern is available on the control.
	$tStr=$tStr & "UIA_IsOffscreenPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_IsOffscreenPropertyId) & $tSeparator ; Whether the element is scrolled or collapsed out of view.
	$tStr=$tStr & "UIA_IsPasswordPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_IsPasswordPropertyId) & $tSeparator ; Whether the element contains protected content or a password.
	$tStr=$tStr & "UIA_IsRangeValuePatternAvailablePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_IsRangeValuePatternAvailablePropertyId) & $tSeparator ; Whether the RangeValue pattern is available on the control.
	$tStr=$tStr & "UIA_IsRequiredForFormPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_IsRequiredForFormPropertyId) & $tSeparator ; Whether the element is a required field on a form.
	$tStr=$tStr & "UIA_IsScrollItemPatternAvailablePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_IsScrollItemPatternAvailablePropertyId) & $tSeparator ; Whether the ScrollItem control pattern is available on the element.
	$tStr=$tStr & "UIA_IsScrollPatternAvailablePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_IsScrollPatternAvailablePropertyId) & $tSeparator ; Whether the Scroll control pattern is available on the element.
	$tStr=$tStr & "UIA_IsSelectionItemPatternAvailablePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_IsSelectionItemPatternAvailablePropertyId) & $tSeparator ; Whether the SelectionItem control pattern is available on the element.
	$tStr=$tStr & "UIA_IsSelectionPatternAvailablePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_IsSelectionPatternAvailablePropertyId) & $tSeparator ; Whether the pattern is available on the element.
	$tStr=$tStr & "UIA_IsSynchronizedInputPatternAvailablePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_IsSynchronizedInputPatternAvailablePropertyId) & $tSeparator ; Whether the SynchronizedInput control pattern is available on the element.
	$tStr=$tStr & "UIA_IsTableItemPatternAvailablePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_IsTableItemPatternAvailablePropertyId) & $tSeparator ; Whether the TableItem control pattern is available on the element.
	$tStr=$tStr & "UIA_IsTablePatternAvailablePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_IsTablePatternAvailablePropertyId) & $tSeparator ; Whether the Table conntrol pattern is available on the element.
	$tStr=$tStr & "UIA_IsTextPatternAvailablePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_IsTextPatternAvailablePropertyId) & $tSeparator ; Whether the Text control pattern is available on the element.
	$tStr=$tStr & "UIA_IsTogglePatternAvailablePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_IsTogglePatternAvailablePropertyId) & $tSeparator ; Whether the Toggle control pattern is available on the element.
	$tStr=$tStr & "UIA_IsTransformPatternAvailablePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_IsTransformPatternAvailablePropertyId) & $tSeparator ; Whether the Transform control pattern is available on the element.
	$tStr=$tStr & "UIA_IsValuePatternAvailablePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_IsValuePatternAvailablePropertyId) & $tSeparator ; Whether the Value control pattern is available on the element.
	$tStr=$tStr & "UIA_IsVirtualizedItemPatternAvailablePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_IsVirtualizedItemPatternAvailablePropertyId) & $tSeparator ; Whether the VirtualizedItem control pattern is available on the element.
	$tStr=$tStr & "UIA_IsWindowPatternAvailablePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_IsWindowPatternAvailablePropertyId) & $tSeparator ; Whether the Window control pattern is available on the element.
	$tStr=$tStr & "UIA_ItemStatusPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_ItemStatusPropertyId) & $tSeparator ; Control-specific status.
	$tStr=$tStr & "UIA_ItemTypePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_ItemTypePropertyId) & $tSeparator ; Description of the item type, such as "Document File" or "Folder".
	$tStr=$tStr & "UIA_LabeledByPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_LabeledByPropertyId) & $tSeparator ; Element that contains the text label for this element.
	$tStr=$tStr & "UIA_LegacyIAccessibleChildIdPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_LegacyIAccessibleChildIdPropertyId) & $tSeparator ; MSAA child ID of the element.
	$tStr=$tStr & "UIA_LegacyIAccessibleDefaultActionPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_LegacyIAccessibleDefaultActionPropertyId) & $tSeparator ; MSAA default action.
	$tStr=$tStr & "UIA_LegacyIAccessibleDescriptionPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_LegacyIAccessibleDescriptionPropertyId) & $tSeparator ; MSAA description.
	$tStr=$tStr & "UIA_LegacyIAccessibleHelpPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_LegacyIAccessibleHelpPropertyId) & $tSeparator ; MSAA help string.
	$tStr=$tStr & "UIA_LegacyIAccessibleKeyboardShortcutPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_LegacyIAccessibleKeyboardShortcutPropertyId) & $tSeparator ; MSAA shortcut key.
	$tStr=$tStr & "UIA_LegacyIAccessibleNamePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_LegacyIAccessibleNamePropertyId) & $tSeparator ; MSAA name.
	$tStr=$tStr & "UIA_LegacyIAccessibleRolePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_LegacyIAccessibleRolePropertyId) & $tSeparator ; MSAA role.
	$tStr=$tStr & "UIA_LegacyIAccessibleSelectionPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_LegacyIAccessibleSelectionPropertyId) & $tSeparator ; MSAA selection.
	$tStr=$tStr & "UIA_LegacyIAccessibleStatePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_LegacyIAccessibleStatePropertyId) & $tSeparator ; MSAA state.
	$tStr=$tStr & "UIA_LegacyIAccessibleValuePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_LegacyIAccessibleValuePropertyId) & $tSeparator ; MSAA value.
	$tStr=$tStr & "UIA_LocalizedControlTypePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_LocalizedControlTypePropertyId) & $tSeparator ; Localized string describing the control type of element.
	$tStr=$tStr & "UIA_MultipleViewCurrentViewPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_MultipleViewCurrentViewPropertyId) & $tSeparator ; Current view state of the control.
	$tStr=$tStr & "UIA_MultipleViewSupportedViewsPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_MultipleViewSupportedViewsPropertyId) & $tSeparator ; Supported control-specific views.
	$tStr=$tStr & "UIA_NamePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_NamePropertyId) & $tSeparator ; Name of the control.
	$tStr=$tStr & "UIA_NativeWindowHandlePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_NativeWindowHandlePropertyId) & $tSeparator ; Underlying HWND of the element, if one exists.
	$tStr=$tStr & "UIA_OrientationPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_OrientationPropertyId) & $tSeparator ; Orientation of the element.
	$tStr=$tStr & "UIA_ProcessIdPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_ProcessIdPropertyId) & $tSeparator ; Identifier of the process that the element resides in.
	$tStr=$tStr & "UIA_ProviderDescriptionPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_ProviderDescriptionPropertyId) & $tSeparator ; Description of the UI Automation provider.
	$tStr=$tStr & "UIA_RangeValueIsReadOnlyPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_RangeValueIsReadOnlyPropertyId) & $tSeparator ; Whether the value is read-only.
	$tStr=$tStr & "UIA_RangeValueLargeChangePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_RangeValueLargeChangePropertyId) & $tSeparator ; Amount by which the value is adjusted by input such as PgDn.
	$tStr=$tStr & "UIA_RangeValueMaximumPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_RangeValueMaximumPropertyId) & $tSeparator ; Maximum value in the range.
	$tStr=$tStr & "UIA_RangeValueMinimumPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_RangeValueMinimumPropertyId) & $tSeparator ; Minimum value in the range.
	$tStr=$tStr & "UIA_RangeValueSmallChangePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_RangeValueSmallChangePropertyId) & $tSeparator ; Amount by which the value is adjusted by input such as an arrow key.
	$tStr=$tStr & "UIA_RangeValueValuePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_RangeValueValuePropertyId) & $tSeparator ; Current value.
	$tStr=$tStr & "UIA_RuntimeIdPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_RuntimeIdPropertyId) & $tSeparator ; Run time identifier of the element.
	$tStr=$tStr & "UIA_ScrollHorizontallyScrollablePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_ScrollHorizontallyScrollablePropertyId) & $tSeparator ; Whether the control can be scrolled horizontally.
	$tStr=$tStr & "UIA_ScrollHorizontalScrollPercentPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_ScrollHorizontalScrollPercentPropertyId) & $tSeparator ; How far the element is currently scrolled.
	$tStr=$tStr & "UIA_ScrollHorizontalViewSizePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_ScrollHorizontalViewSizePropertyId) & $tSeparator ; The viewable width of the control.
	$tStr=$tStr & "UIA_ScrollVerticallyScrollablePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_ScrollVerticallyScrollablePropertyId) & $tSeparator ; Whether the control can be scrolled vertically.
	$tStr=$tStr & "UIA_ScrollVerticalScrollPercentPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_ScrollVerticalScrollPercentPropertyId) & $tSeparator ; How far the element is currently scrolled.
	$tStr=$tStr & "UIA_ScrollVerticalViewSizePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_ScrollVerticalViewSizePropertyId) & $tSeparator ; The viewable height of the control.
	$tStr=$tStr & "UIA_SelectionCanSelectMultiplePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_SelectionCanSelectMultiplePropertyId) & $tSeparator ; Whether multiple items can be in the selection.
	$tStr=$tStr & "UIA_SelectionIsSelectionRequiredPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_SelectionIsSelectionRequiredPropertyId) & $tSeparator ; Whether at least one item must be in the selection at all times.
	$tStr=$tStr & "UIA_SelectionselectionPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_SelectionselectionPropertyId) & $tSeparator ; The items in the selection.
	$tStr=$tStr & "UIA_SelectionItemIsSelectedPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_SelectionItemIsSelectedPropertyId) & $tSeparator ; Whether the item can be selected.
	$tStr=$tStr & "UIA_SelectionItemSelectionContainerPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_SelectionItemSelectionContainerPropertyId) & $tSeparator ; The control that contains the item.
	$tStr=$tStr & "UIA_TableColumnHeadersPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_TableColumnHeadersPropertyId) & $tSeparator ; Collection of column header providers.
	$tStr=$tStr & "UIA_TableItemColumnHeaderItemsPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_TableItemColumnHeaderItemsPropertyId) & $tSeparator ; Column headers.
	$tStr=$tStr & "UIA_TableRowHeadersPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_TableRowHeadersPropertyId) & $tSeparator ; Collection of row header providers.
	$tStr=$tStr & "UIA_TableRowOrColumnMajorPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_TableRowOrColumnMajorPropertyId) & $tSeparator ; Whether the table is primarily organized by row or column.
	$tStr=$tStr & "UIA_TableItemRowHeaderItemsPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_TableItemRowHeaderItemsPropertyId) & $tSeparator ; Row headers.
	$tStr=$tStr & "UIA_ToggleToggleStatePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_ToggleToggleStatePropertyId) & $tSeparator ; The toggle state of the control.
	$tStr=$tStr & "UIA_TransformCanMovePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_TransformCanMovePropertyId) & $tSeparator ; Whether the element can be moved.
	$tStr=$tStr & "UIA_TransformCanResizePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_TransformCanResizePropertyId) & $tSeparator ; Whether the element can be resized.
	$tStr=$tStr & "UIA_TransformCanRotatePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_TransformCanRotatePropertyId) & $tSeparator ; Whether the element can be rotated.
	$tStr=$tStr & "UIA_ValueIsReadOnlyPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_ValueIsReadOnlyPropertyId) & $tSeparator ; Whether the value is read-only.
	$tStr=$tStr & "UIA_ValueValuePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_ValueValuePropertyId) & $tSeparator ; Current value.
	$tStr=$tStr & "UIA_WindowCanMaximizePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_WindowCanMaximizePropertyId) & $tSeparator ; Whether the window can be maximized.
	$tStr=$tStr & "UIA_WindowCanMinimizePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_WindowCanMinimizePropertyId) & $tSeparator ; Whether the window can be minimized.
	$tStr=$tStr & "UIA_WindowIsModalPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_WindowIsModalPropertyId) & $tSeparator ; Whether the window is modal.
	$tStr=$tStr & "UIA_WindowIsTopmostPropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_WindowIsTopmostPropertyId) & $tSeparator ; Whether the window is on top of other windows.
	$tStr=$tStr & "UIA_WindowWindowInteractionStatePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_WindowWindowInteractionStatePropertyId) & $tSeparator ; Whether the window can receive input.
	$tStr=$tStr & "UIA_WindowWindowVisualStatePropertyId :=" &_UIA_getPropertyValue($oUIElement, $UIA_WindowWindowVisualStatePropertyId) & $tSeparator ; Whether the window is maximized, minimized, or restored (normal).
	return $tStr
endFunc

; Draw rectangle on screen.
Func _DrawRect($tLeft, $tRight, $tTop, $tBottom, $color = 0xFF, $PenWidth = 4)
    Local $hDC, $hPen, $obj_orig, $x1, $x2, $y1, $y2
    $x1 = $tLeft
    $x2 = $tRight
    $y1 = $tTop
    $y2 = $tBottom
    $hDC = _WinAPI_GetWindowDC(0) ; DC of entire screen (desktop)
    $hPen = _WinAPI_CreatePen($PS_SOLID, $PenWidth, $color)
    $obj_orig = _WinAPI_SelectObject($hDC, $hPen)

    _WinAPI_DrawLine($hDC, $x1, $y1, $x2, $y1) ; horizontal to right
    _WinAPI_DrawLine($hDC, $x2, $y1, $x2, $y2) ; vertical down on right
    _WinAPI_DrawLine($hDC, $x2, $y2, $x1, $y2) ; horizontal to left right
    _WinAPI_DrawLine($hDC, $x1, $y2, $x1, $y1) ; vertical up on left

    ; clear resources
    _WinAPI_SelectObject($hDC, $obj_orig)
    _WinAPI_DeleteObject($hPen)
    _WinAPI_ReleaseDC(0, $hDC)
EndFunc   ;==>_DrawtRect

;~ Small helper function to get an object out of a treeSearch based on the name / title
func _UIA_getFirstObjectOfElement($obj,$str,$treeScope)
	local $tResult
	local $pCondition
	local $propertyID

;~ 	Split a description into multiple subdescription/properties
	$tResult=stringsplit($str,":=",1)
	Local $tVal
	if $tResult[0]=1 Then
		$propertyID=$UIA_NamePropertyId
		$tVal=$str
	Else
		for $i=0 to ubound($propertiesSupportedArray)-1
			if $propertiesSupportedArray[$i][0]=stringlower($tResult[1]) Then
				_UIA_Debug("matched: " & $propertiesSupportedArray[$i][0] & $propertiesSupportedArray[$i][1] & @crlf)
				$propertyID=$propertiesSupportedArray[$i][1]

;~ 				Some properties expect a number (otherwise system will break)
				switch $propertiesSupportedArray[$i][1]
					case $UIA_ControlTypePropertyId
						$tVal=number($tResult[2])
					case else
						$tVal=$tResult[2]
				endswitch
			endif
		next
	EndIf

	_UIA_Debug("Matching: " & $PropertyId & " for " & $tVal & @CRLF)

;~ Tricky when numeric values to pass
	$objUIAutomation.createPropertyCondition($PropertyId, $tVal, $pCondition)
	Local $oCondition
	$oCondition=ObjCreateInterface($pCondition,$sIID_IUIAutomationPropertyCondition,$dtagIUIAutomationPropertyCondition)

	Local $iTry=1
	$oUIElement=""
	while not isobj($oUIElement) and $iTry<= $UIA_tryMax
		$t=$obj.Findfirst($TreeScope,$oCondition,$pUIElement)
		$oUIElement=ObjCreateInterface($pUIElement, $sIID_IUIAutomationElement, $dtagIUIAutomationElement)
		if not isobj($oUIElement) Then
			sleep(100)
			$iTry=$iTry+1
		endif
	WEnd

	if isobj($oUIElement) Then
;~ 		_UIA_Debug("UIA found the element" & @CRLF)
		if _UIA_getVar("Global.Highlight")= true Then
			_UIA_Highlight($oUIElement)
		EndIf

		return $oUIElement
	Else
;~ 		_UIA_Debug("UIA failing ** NOT ** found the element" & @CRLF)
		if _UIA_getVar("Global.Debug")= true Then
			_UIA_DumpThemAll($obj, $treescope)
		EndIf

		return ""
	endif

EndFunc


;~ Find it by using a findall array of the UIA framework
func _UIA_getObjectByFindAll($obj, $str, $treescope,$p1=0)
	dim $pCondition, $pTrueCondition
	dim $pElements, $iLength

	local $tResult
	local $propertyID
	local $tPos
	local $relPos
	local $relIndex=0
	local $tMatch
	local $tStr
	local $properties2Match[1][2]   ;~ All properties of the expression to match in a normalized form

;~ 	Split it first into multiple sections representing each property
	$allProperties=stringsplit($str,";",1)

;~ Redefine the array to have all properties that are used to identify
	$propertyCount=$allProperties[0]
	redim $properties2Match[$propertyCount][2]
;~ 	_UIA_Debug("_UIA_getObjectByFindAll " &  $str & $propertyCount & @crlf)
	for $i=1 to $allProperties[0]
;~ 		_UIA_Debug("  _UIA_getObjectByFindAll " &  $allProperties[$i] & @crlf)
		$tResult=stringsplit($allProperties[$i],":=",1)
		$tResult[1]=stringstripws($tresult[1],3)
		$tResult[2]=stringstripws($tresult[2],3)

		;~ Handle syntax without a property to have default name property:  Ok as Name:=Ok
		if $tResult[0]=1 Then
			$propName=$UIA_NamePropertyId
			$propValue=$allProperties[$i]

			$properties2Match[$i-1][0]=$propName
			$properties2Match[$i-1][1]=$propValue
		Else
			$propName=$tResult[1]
			$propValue=$tResult[2]

			$bAdd=True
			if $propName="indexrelative" Then
				$relPos=$propValue
				$bAdd=False
			EndIf
			if ($propName="index") or ($propName="instance") Then
				$relIndex=$propValue
				$bAdd=False
			EndIf

			if $bAdd=true Then
				$index=_UIA_getPropertyIndex($propName)

;~ 				Some properties expect a number (otherwise system will break)
				switch $propertiesSupportedArray[$index][1]
					case $UIA_ControlTypePropertyId
						$propValue=number(_UIA_getControlID($propValue))
				endswitch
;~ 				_UIA_Debug("  adding " &  $propname & $propvalue & @crlf)

;~ Add it to the normalized array
				$properties2Match[$i-1][0]= $propertiesSupportedArray[$index][1]  ;~ store the propertyID (numeric value)
				$properties2Match[$i-1][1]= $propvalue

			endif
		endif
	Next

;~ Now walk thru the tree
;~ 	_UIA_Debug("_UIA_getObjectByFindAll walk thru the tree" &  $allProperties[0] )
    $objUIAutomation.CreateTrueCondition($pTrueCondition)
    $oCondition=ObjCreateInterface($pTrueCondition, $sIID_IUIAutomationCondition,$dtagIUIAutomationCondition)
	$obj.FindAll($treescope, $oCondition, $pElements)

    $oAutomationElementArray = ObjCreateInterFace($pElements, $sIID_IUIAutomationElementArray, $dtagIUIAutomationElementArray)
	$matchCount=0
;~ All elements to inspect are in this array
    $oAutomationElementArray.Length($iLength)
;~ 		_UIA_Debug("_UIA_getObjectByFindAll walk thru the tree" &  $iLength )
    For $i = 0 To $iLength - 1; it's zero based
		$oAutomationElementArray.GetElement($i, $pUIElement)
        $oUIElement = ObjCreateInterface($pUIElement, $sIID_IUIAutomationElement, $dtagIUIAutomationElement)

;~ 		_UIA_Debug( $tval & "searching the Name is: <" &  _UIA_getPropertyValue($oUIElement,$UIA_NamePropertyId) &  ">" & @TAB _
;~ 			& "Class   := <" & _UIA_getPropertyValue($oUIElement,$uia_classnamepropertyid) &  ">" & @TAB _
;~ 			& "controltype:= <" &  _UIA_getPropertyValue($oUIElement,$UIA_ControlTypePropertyId) &  ">" & @TAB & @CRLF)

;~		Walk thru all properties in the properties2Match array to match
		for $j=0 to ubound($properties2Match,1)-1
			$propertyID=$properties2Match[$j][0]
			$propertyVal=$properties2Match[$j][1]

;~ 			Some properties expect a number (otherwise system will break)
			switch $propertyId
				case $UIA_ControlTypePropertyId
					$propertyVal=number($propertyVal)
			endswitch

			$propertyActualValue=_UIA_getPropertyValue($oUIElement,$PropertyId)
;~ 			_UIA_Debug("j:" & $j & "[" & $propertyID & "][" & $propertyVal & "][" & $propertyActualValue & "]" & @CRLF)

			$tMatch=stringregexp($propertyActualValue, $propertyVal,0)
			if $tMatch=0 Then ExitLoop
		Next

		if $tMatch=1 Then
				if $relPos <> 0 Then
;~ 					_UIA_Debug("Relative position used")
					$oAutomationElementArray.GetElement($i+$relPos, $pUIElement)
					$oUIElement = ObjCreateInterface($pUIElement, $sIID_IUIAutomationElement, $dtagIUIAutomationElement)
				EndIf
				if $relIndex <> 0 Then
;~ 					_UIA_Debug("Index position used")
					$matchCount=$matchCount+1
					if $matchCount <> $relIndex then $tMatch=0
				EndIf

				if $tMatch=1 Then
;~ 					_UIA_Debug( " Found the Name is: <" &  _UIA_getPropertyValue($oUIElement,$UIA_NamePropertyId) &  ">" & @TAB _
;~ 				    & "Class   := <" & _UIA_getPropertyValue($oUIElement,$uia_classnamepropertyid) &  ">" & @TAB _
;~ 					& "controltype:= <" &  _UIA_getPropertyValue($oUIElement,$UIA_ControlTypePropertyId) &  ">" & @TAB  _
;~ 					& " (" &  hex(_UIA_getPropertyValue($oUIElement,$UIA_ControlTypePropertyId)) &  ")" & @TAB & @CRLF)

					if _UIA_getVar("Global.Highlight")= true Then
						_UIA_Highlight($oUIElement)
					EndIf

;~ 					Add element to runtime information object reference
					if isstring($p1) Then
;~ 						_UIA_DEBUG("Storing in RTI as RTI." & $p1 & @CRLF)
						_UIA_setVar("RTI." & $p1, $oUIElement)
					EndIf
					return $oUIElement
				endif
		EndIf
	Next

	if _UIA_getVar("Global.Debug")= true Then
		_UIA_DumpThemAll($obj, $treescope)
	EndIf
	return ""
endfunc

func _UIA_getPattern($obj,$patternID)
local $patternArray[21][3]=[ _
	[$UIA_ValuePatternId    , 			$sIID_IUIAutomationValuePattern    , 		$dtagIUIAutomationValuePattern], _
	[$UIA_InvokePatternId   , 			$sIID_IUIAutomationInvokePattern   , 		$dtagIUIAutomationInvokePattern], _
	[$UIA_SelectionPatternId, 			$sIID_IUIAutomationSelectionPattern, 		$dtagIUIAutomationSelectionPattern], _
    [$UIA_LegacyIAccessiblePatternId, 	$sIID_IUIAutomationLegacyIAccessiblePattern,$dtagIUIAutomationLegacyIAccessiblePattern], _
    [$UIA_SelectionItemPatternId, 		$sIID_IUIAutomationSelectionItemPattern,	$dtagIUIAutomationSelectionItemPattern], _
    [$UIA_RangeValuePatternId, 			$sIID_IUIAutomationRangeValuePattern,		$dtagIUIAutomationRangeValuePattern], _
	[$UIA_ScrollPatternId, 				$sIID_IUIAutomationScrollPattern,			$dtagIUIAutomationScrollPattern], _
	[$UIA_GridPatternId, 				$sIID_IUIAutomationGridPattern,				$dtagIUIAutomationGridPattern], _
	[$UIA_GridItemPatternId, 			$sIID_IUIAutomationGridItemPattern,			$dtagIUIAutomationGridItemPattern], _
	[$UIA_MultipleViewPatternId, 		$sIID_IUIAutomationMultipleViewPattern,		$dtagIUIAutomationMultipleViewPattern], _
	[$UIA_WindowPatternId, 				$sIID_IUIAutomationWindowPattern,			$dtagIUIAutomationWindowPattern], _
	[$UIA_DockPatternId, 				$sIID_IUIAutomationDockPattern,				$dtagIUIAutomationDockPattern], _
	[$UIA_TablePatternId, 				$sIID_IUIAutomationTablePattern,			$dtagIUIAutomationTablePattern], _
	[$UIA_TextPatternId, 				$sIID_IUIAutomationTextPattern,				$dtagIUIAutomationTextPattern], _
	[$UIA_TogglePatternId, 				$sIID_IUIAutomationTogglePattern,			$dtagIUIAutomationTogglePattern], _
	[$UIA_TransformPatternId, 			$sIID_IUIAutomationTransformPattern,		$dtagIUIAutomationTransformPattern], _
	[$UIA_ScrollItemPatternId, 			$sIID_IUIAutomationScrollItemPattern,		$dtagIUIAutomationScrollItemPattern], _
	[$UIA_ItemContainerPatternId, 		$sIID_IUIAutomationItemContainerPattern,	$dtagIUIAutomationItemContainerPattern], _
	[$UIA_VirtualizedItemPatternId, 	$sIID_IUIAutomationVirtualizedItemPattern,	$dtagIUIAutomationVirtualizedItemPattern], _
	[$UIA_SynchronizedInputPatternId, 	$sIID_IUIAutomationSynchronizedInputPattern,$dtagIUIAutomationSynchronizedInputPattern], _
	[$UIA_ExpandCollapsePatternId, 		$sIID_IUIAutomationExpandCollapsePattern, 	$dtagIUIAutomationExpandCollapsePattern] _
		]

    local $pPattern
    local $sIID_Pattern
	local $sdTagPattern
	local $i

	for $i=0 to ubound($patternArray)-1
		if $patternArray[$i][0]=$patternId Then
;~ 			consolewrite("Pattern identified " & @crlf)
			$sIID_Pattern=$patternArray[$i][1]
			$sdTagPattern=$patternArray[$i][2]
		EndIf
	next
;~ 	consolewrite($patternid & $sIID_Pattern & $sdTagPattern & @CRLF)

	$obj.getCurrentPattern($PatternId, $pPattern)
	$oPattern=objCreateInterface($pPattern, $sIID_Pattern, $sdtagPattern)
	if isobj($oPattern) Then
;~ 		consolewrite("UIA found the pattern" & @CRLF)
		return $oPattern
	Else
		_UIA_Debug("UIA WARNING ** NOT ** found the pattern" & @CRLF)
	endif
EndFunc

func _UIA_getTaskBar()
	return _UIA_getFirstObjectOfElement($oDesktop,"classname:=Shell_TrayWnd",$TreeScope_Children)
EndFunc

;~ func _UIA_action($obj, $strAction, $p1=0, $p2=0, $p3=0)
func _UIA_action($obj_or_string, $strAction, $p1=0, $p2=0, $p3=0, $p4=0)

	local $tPattern
	local $x, $y
;~ 	local $objElement
	local $oElement

;~ If we are giving a description then try to make an object first by looking from repository
;~ Otherwise assume an advanced description we should search under one of the previously referenced elements at runtime

	if isobj($obj_or_string) Then
		$oElement=$obj_or_string
		$obj=$obj_or_string
	else
;~ 		_UIA_DEBUG("Finding object " & $obj_or_string & @CRLF)
		$tPhysical=_UIA_getVar($obj_or_string)
;~ If not found in repository assume its a physical description
		if @error=1 Then
;~ 			_UIA_DEBUG("Finding object (bypassing repository) with physical description " & $tPhysical & @CRLF)
			$tPhysical=$obj_or_string
		EndIf

;~ 			TODO: If its a physical description the searching should start at one of the last active elements or parent of that last active element
;~ 		else
;~ 			We found a reference try to make it an object
;~ 		_UIA_DEBUG("Finding object with physical description " & $tPhysical & @CRLF)
;~ if its a mainwindow reference find it under the desktop
			if stringinstr($obj_or_string,".mainwindow") Then

					$startElement="Desktop"
					$oStart=$oDesktop
;~ 					_UIA_DEBUG("Finding object under " & $startElement & @CRLF)
					$oElement=_UIA_getObjectByFindAll($oStart, $tPhysical, $treescope_subtree,$obj_or_string)
					_UIA_setVar("RTI.MAINWINDOW",$oElement)
			else
			;~ 	Find the object under the last referenced mainwindow / parent window
;~ 				_UIA_DEBUG("TPhysical is now " & $tPhysical & @CRLF)
;~ 				_UIA_DEBUG("$obj_or_string is now " & $obj_or_string & @CRLF)

				$startElement="RTI." & stringleft($obj_or_string,stringinstr($obj_or_string,".")) & "mainwindow"
				$oStart=_UIA_getVar($startElement)
;~ 				if isobj($oStart) Then
;~ 					_UIA_DEBUG("Its an object ")
;~ 				endif
;~ 				_UIA_DEBUG("Finding object under " & $startElement & @CRLF)
;~ 				_UIA_getVars2Array()

				$oElement=_UIA_getObjectByFindAll($oStart, $tPhysical, $treescope_subtree)
			endif
;~ 		endif

;~ And just continue the action by setting the $obj value to an UIA element

		if isobj($oElement) Then
			$obj=$oElement
		Else
			seterror(1)
			Return
		endif
	EndIf

	_UIA_setVar("RTI.ACTIONCOUNT",_UIA_getVar("RTI.ACTIONCOUNT")+1)
	_UIA_DEBUG("Action " & _UIA_getVar("RTI.ACTIONCOUNT") & " " & $strAction & " on " & $obj_or_string & @CRLF & "   ")

	$controlType=_UIA_getPropertyValue($oUIElement,$UIA_ControlTypePropertyId)


;~ Execute the given action
	switch $strAction
;~ 		All mouse click actions
		case "leftclick", "left", "click", "leftdoubleclick", "leftdouble", "doubleclick", _
			 "rightclick", "right",        "rightdoubleclick", "rightdouble", _
             "middleclick", "middle",      "middledoubleclick", "middledouble"

			local $clickAction="left"  ;~ Default action is the left mouse button
			local $clickCount=1      ;~ Default action is the single click

			if stringinstr($strAction, "right") then $clickAction="right"
			if stringinstr($strAction, "middle") then $clickAction="middle"
			if stringinstr($strAction, "double") then $clickCount=2

			;~ consolewrite("So you saw it selected but did not click" & @crlf)
			;~ still you can click as you now know the dimensions where to click
			dim $t
			$t=stringsplit(_UIA_getPropertyValue($obj, $UIA_BoundingRectanglePropertyId),";")
;~ 			consolewrite($t[1] & ";" & $t[2] & ";" & $t[3] & ";" & $t[4] & @crlf)
;~ 			_winapi_mouse_event($MOUSEEVENTF_ABSOLUTE + $MOUSEEVENTF_MOVE,$t[1],$t[2])
			$x=int($t[1]+($t[3]/2))
			$y=int($t[2]+$t[4]/2)

;~ Split into 2 actions for screenrecording purposes intentionally a slowdown
;~ Arguable that this delay should be configurable or removed on synchronizing differently in future
;~ 			First try to set the focus to make sure right window is active

;~ 	_UIA_DEBUG("Title is: <" &  _UIA_getPropertyValue($oUIElement,$UIA_NamePropertyId) &  ">" & $clickcount & ":" & $clickaction & ":" & $x & ":" & $y & ":" & @CRLF)

;~ TODO: Check if setting focus should happen as it influences behavior before clicking
;~ Tricky when using setfocus on menuitems, seems to do the click already
;~ 			$obj.setfocus()

;~ 			Mouse should move to keep it as userlike as possible
			mousemove($x,$y,0)
;~ 			mouseclick($clickAction,Default,Default,$clickCount,0)
			mouseclick($clickAction,$x,$y,$clickCount,0)
			sleep($UIA_DefaultWaitTime)

		case "setvalue"

;~ TODO: Find out how to set title for a window with UIA commands
;~ winsettitle(hwnd(_UIA_getVar("RTI.calculator.HWND")),"","nicer")
;~ winsettitle("Naamloos - Kladblok","","This works better")

			if ($controltype=$UIA_WindowControlTypeId) then
				$hwnd=0
				$obj.CurrentNativeWindowHandle($hwnd)
				consolewrite($hwnd)
				winsettitle(hwnd($hwnd),"",$p1)
			Else
				$obj.setfocus()
				sleep($UIA_DefaultWaitTime)
				$tPattern=_UIA_getPattern($obj,$UIA_ValuePatternId)
				$tPattern.setvalue($p1)
			EndIf

		case "setvalue using keys"
			$obj.setfocus()
			send("^a")
			send($p1)
			sleep($UIA_DefaultWaitTime)
		case "sendkeys", "enterstring"
			$obj.setfocus()
			send($p1)
		case "invoke"
			$obj.setfocus()
			sleep($UIA_DefaultWaitTime)
			$tPattern=_UIA_getPattern($obj,$UIA_InvokePatternId)
			$tPattern.invoke()
		case "focus","setfocus"
			$obj.setfocus()
			sleep($UIA_DefaultWaitTime)
		case Else
	endswitch
EndFunc

;~ Just dumps all information under a certain object
func _UIA_DumpThemAll($oElementStart, $TreeScope)
;~  Get result with findall function alternative could be the treewalker
    dim $pCondition, $pTrueCondition
	dim $pElements, $iLength

	_UIA_Debug("***** Dumping tree *****" & @CRLF)

    $objUIAutomation.CreateTrueCondition($pTrueCondition)
    $oCondition=ObjCreateInterface($pTrueCondition, $sIID_IUIAutomationCondition,$dtagIUIAutomationCondition)

	$oElementStart.FindAll($TreeScope, $oCondition, $pElements)

    $oAutomationElementArray = ObjCreateInterFace($pElements, $sIID_IUIAutomationElementArray, $dtagIUIAutomationElementArray)

    $oAutomationElementArray.Length($iLength)
    For $i = 0 To $iLength - 1; it's zero based
		$oAutomationElementArray.GetElement($i, $pUIElement)
        $oUIElement = ObjCreateInterface($pUIElement, $sIID_IUIAutomationElement, $dtagIUIAutomationElement)
        _UIA_Debug( "Title is: <" &  _UIA_getPropertyValue($oUIElement,$UIA_NamePropertyId) &  ">" & @TAB _
				    & "Class   := <" & _UIA_getPropertyValue($oUIElement,$uia_classnamepropertyid) &  ">" & @TAB _
					& "controltype:= " _
					& "<" &  _UIA_getControlName(_UIA_getPropertyValue($oUIElement,$UIA_ControlTypePropertyId)) &  ">" & @TAB  _
					& ",<" &  _UIA_getPropertyValue($oUIElement,$UIA_ControlTypePropertyId) &  ">" & @TAB  _
					& ", (" &  hex(_UIA_getPropertyValue($oUIElement,$UIA_ControlTypePropertyId)) &  ")" & @TAB _
					& ", acceleratorkey:= <" &  _UIA_getPropertyValue($oUIElement,$UIA_AcceleratorKeyPropertyId) &  ">" & @TAB _
					& ", automationid:= <" &  _UIA_getPropertyValue($oUIElement,$UIA_AutomationIdPropertyId) &  ">" & @TAB & @CRLF)

	Next

EndFunc
;~ For the moment just dump to the consolewindow
;~ TODO: Differentiate between debug, error, warning, informational
func _UIA_Debug($s)
;~ 	filewrite("log.txt",$s)

	consolewrite($s)
EndFunc

func _UIA_StartSUT($SUT_VAR)
	local $fullName=_UIA_getVar( $SUT_VAR & ".Fullname")
	local $processName=_UIA_getVar($SUT_VAR & ".processname")
	local $app2Start=$fullName & " " & _UIA_getVar($SUT_VAR & ".Parameters")
	local $workingDir= _UIA_getVar($SUT_VAR & ".Workingdir")
	local $windowState=_UIA_getVar($SUT_VAR & ".Windowstate")
    local $result, $result2   ; Holds the process id's

;~ 	_UIA_Debug("SUT 1 Starting : " & $fullName & @CRLF)
	if fileexists($fullName) Then
;~ 		_UIA_Debug("SUT 2 Starting : " & $fullName & @CRLF)
;~ 		Only start new instance when not found
		$result2=processexists($processName)
		if $result2=0 Then
			_UIA_Debug("Starting : " & $app2Start & " from " & $workingDir)
			$result=run($app2Start,$workingDir, $windowState )
			$result2=ProcessWait($processName,60)
;~ 			sleep(500) ;~ Just to give the system some time to show everything
		EndIf

;~ Wait for the window to be there
		$oSUT=_UIA_getObjectByFindAll($oDesktop, "processid:=" & $result2, $treescope_children)
		if not isobj($oSUT) Then
			_UIA_Debug("No window found in SUT : " & $app2Start & " from " & $workingDir)
		Else
		;~ Add it to the Runtime Type Information
			_UIA_setVar("RTI." & $SUT_VAR & ".PID", $result2)
			_UIA_setVar("RTI." & $SUT_VAR & ".HWND", hex(_UIA_getPropertyValue($oSUT, $UIA_NativeWindowHandlePropertyId)))
;~ 			_UIA_DumpThemAll($oSUT,$treescope_subtree)
		EndIf
	Else
		_UIA_Debug("No clue where to find the system under test (SUT) on your system, please start manually:" & @CRLF )
		_UIA_Debug($app2Start & @CRLF)
	EndIf
EndFunc

func _UIA_Highlight($oElement)
	$t=stringsplit(_UIA_getPropertyValue($oElement, $UIA_BoundingRectanglePropertyId),";")
	_DrawRect($t[1],$t[3]+$t[1],$t[2],$t[4]+$t[2])
EndFunc


;~ ***** Experimental catching the events that are flowing around *****
;~ ;===============================================================================
;~ #interface "IUnknown"
;~ Global Const $sIID_IUnknown = "{00000000-0000-0000-C000-000000000046}"
;~ ; Definition
;~ Global $dtagIUnknown = "QueryInterface hresult(ptr;ptr*);" & _
;~ 		"AddRef dword();" & _
;~ 		"Release dword();"
;~ ; List
;~ Global $ltagIUnknown = "QueryInterface;" & _
;~ 		"AddRef;" & _
;~ 		"Release;"
;~ ;===============================================================================
;~ ;===============================================================================
;~ #interface "IDispatch"
;~ Global Const $sIID_IDispatch = "{00020400-0000-0000-C000-000000000046}"
;~ ; Definition
;~ Global $dtagIDispatch = $dtagIUnknown & _
;~ 		"GetTypeInfoCount hresult(dword*);" & _
;~ 		"GetTypeInfo hresult(dword;dword;ptr*);" & _
;~ 		"GetIDsOfNames hresult(ptr;ptr;dword;dword;ptr);" & _
;~ 		"Invoke hresult(dword;ptr;dword;word;ptr;ptr;ptr;ptr);"
;~ ; List
;~ Global $ltagIDispatch = $ltagIUnknown & _
;~ 		"GetTypeInfoCount;" & _
;~ 		"GetTypeInfo;" & _
;~ 		"GetIDsOfNames;" & _
;~ 		"Invoke;"
;~ ;===============================================================================
;~ ; #FUNCTION# ====================================================================================================================
;~ ; Name...........: UIA_ObjectFromTag($obj, $id)
;~ ; Description ...: Get an object from a DTAG
;~ ; Syntax.........:
;~ ; Parameters ....:
;~ ;
;~ ; Return values .: Success      - Returns 1
;~ ;                  Failure		- Returns 0 and sets @error on errors:
;~ ;                  |@error=1     - UI automation failed
;~ ;                  |@error=2     - UI automation desktop failed
;~ ; Author ........: TRANCEXX
;~ ; Modified.......:
;~ ; Remarks .......: None
;~ ; Related .......:
;~ ; Link ..........:
;~ ; Example .......: Yes
;~ ; ===============================================================================================================================
;~ http://www.autoitscript.com/forum/topic/153859-objevent-possible-with-addfocuschangedeventhandler/
;~ Func UIA_ObjectFromTag($sFunctionPrefix, $tagInterface, ByRef $tInterface)
;~     Local Const $tagIUnknown = "QueryInterface hresult(ptr;ptr*);" & _
;~             "AddRef dword();" & _
;~             "Release dword();"
;~     ; Adding IUnknown methods
;~     $tagInterface = $tagIUnknown & $tagInterface
;~     Local Const $PTR_SIZE = DllStructGetSize(DllStructCreate("ptr"))
;~     ; Below line really simple even though it looks super complex. It's just written weird to fit one line, not to steal your eyes
;~     Local $aMethods = StringSplit(StringReplace(StringReplace(StringReplace(StringReplace(StringTrimRight(StringReplace(StringRegExpReplace($tagInterface, "\h*(\w+)\h*(\w+\*?)\h*(\((.*?)\))\h*(;|;*\z)", "$1\|$2;$4" & @LF), ";" & @LF, @LF), 1), "object", "idispatch"), "variant*", "ptr"), "hresult", "long"), "bstr", "ptr"), @LF, 3)
;~     Local $iUbound = UBound($aMethods)
;~     Local $sMethod, $aSplit, $sNamePart, $aTagPart, $sTagPart, $sRet, $sParams
;~     ; Allocation. Read http://msdn.microsoft.com/en-us/library/ms810466.aspx to see why like this (object + methods):
;~     $tInterface = DllStructCreate("ptr[" & $iUbound + 1 & "]")
;~     If @error Then Return SetError(1, 0, 0)
;~     For $i = 0 To $iUbound - 1
;~         $aSplit = StringSplit($aMethods[$i], "|", 2)
;~         If UBound($aSplit) <> 2 Then ReDim $aSplit[2]
;~         $sNamePart = $aSplit[0]
;~         $sTagPart = $aSplit[1]
;~         $sMethod = $sFunctionPrefix & $sNamePart
;~         $aTagPart = StringSplit($sTagPart, ";", 2)
;~         $sRet = $aTagPart[0]
;~         $sParams = StringReplace($sTagPart, $sRet, "", 1)
;~         $sParams = "ptr" & $sParams
;~         DllStructSetData($tInterface, 1, DllCallbackGetPtr(DllCallbackRegister($sMethod, $sRet, $sParams)), $i + 2) ; Freeing is left to AutoIt.
;~     Next
;~     DllStructSetData($tInterface, 1, DllStructGetPtr($tInterface) + $PTR_SIZE) ; Interface method pointers are actually pointer size away
;~     Return ObjCreateInterface(DllStructGetPtr($tInterface), "", $tagInterface, False) ; and first pointer is object pointer that's wrapped
;~ EndFunc