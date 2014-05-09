#include "CUIAutomation2.au3"
#include "UIAWrappers.au3"

;Opt( "MustDeclareVars", 1 )

MainFunc()


Func MainFunc()

  Local $hWindow = WinGetHandle( "HOBOnode Viewer Utility" )
  If Not $hWindow Then Return ConsoleWrite( "Window handle ERR" & @CRLF )
  ConsoleWrite( "Window handle OK" & @CRLF )

  ; Create UI Automation object
  Local $oUIAutomation = ObjCreateInterface( $sCLSID_CUIAutomation, $sIID_IUIAutomation, $dtagIUIAutomation )
  If Not IsObj( $oUIAutomation ) Then Return ConsoleWrite( "UI Automation object ERR" & @CRLF )
  ConsoleWrite( "UI Automation object OK" & @CRLF )

  ; Get UI Automation element from window handle
  Local $pWindow, $oWindow
  $oUIAutomation.ElementFromHandle( $hWindow, $pWindow )
  $oWindow = ObjCreateInterface( $pWindow, $sIID_IUIAutomationElement, $dtagIUIAutomationElement )
  If Not IsObj( $oWindow ) Then Return ConsoleWrite( "Automation element from window ERR" & @CRLF )
  ConsoleWrite( "Automation element from window OK" & @CRLF )





  ; Condition to find "Save Data..." button
  Local $pCondition, $pCondition1, $pCondition2
  $oUIAutomation.CreatePropertyCondition( $UIA_ControlTypePropertyId, $UIA_ButtonControlTypeId, $pCondition1 )
  $oUIAutomation.CreatePropertyCondition( $UIA_NamePropertyId, "Save Data...", $pCondition2 )
  $oUIAutomation.CreateAndCondition( $pCondition1, $pCondition2, $pCondition )
  If Not $pCondition Then Return ConsoleWrite( "Property condition ERR Save" & @CRLF )
  ConsoleWrite( "Property condition OK Save" & @CRLF )

; Condition to find "Get/Refresh Time Range" button
  Local $pConditionTime, $pCondition1Time, $pCondition2Time
  $oUIAutomation.CreatePropertyCondition( $UIA_ControlTypePropertyId, $UIA_ButtonControlTypeId, $pCondition1Time )
  $oUIAutomation.CreatePropertyCondition( $UIA_NamePropertyId, "Get/Refresh Time Range", $pCondition2Time )
  $oUIAutomation.CreateAndCondition( $pCondition1Time, $pCondition2Time, $pConditionTime )
  If Not $pConditionTime Then Return ConsoleWrite( "Property condition ERR Time" & @CRLF )
  ConsoleWrite( "Property condition OK Time" & @CRLF )

; Condition to find "Generate Data Streams" button
  Local $pConditionGenerate, $pCondition1Generate, $pCondition2Generate
  $oUIAutomation.CreatePropertyCondition( $UIA_ControlTypePropertyId, $UIA_ButtonControlTypeId, $pCondition1Generate )
  $oUIAutomation.CreatePropertyCondition( $UIA_NamePropertyId, "Generate Data Streams", $pCondition2Generate )
  $oUIAutomation.CreateAndCondition( $pCondition1Generate, $pCondition2Generate, $pConditionGenerate )
  If Not $pConditionGenerate Then Return ConsoleWrite( "Property condition ERR Generate" & @CRLF )
  ConsoleWrite( "Property condition OK Generate" & @CRLF )

 ;Condition to find "416 SN: 10485416" sensor
  Local $pConditionSensor1, $pCondition1Sensor1, $pCondition2Sensor1
  $oUIAutomation.CreatePropertyCondition( $UIA_ControlTypePropertyId, $UIA_ListItemControlTypeId, $pCondition1Sensor1 )
  $oUIAutomation.CreatePropertyCondition( $UIA_NamePropertyId, "416 SN: 10485416; controltype:=ListItem", $pCondition2Sensor1 )
  $oUIAutomation.CreateAndCondition( $pCondition1Sensor1, $pCondition2Sensor1, $pConditionSensor1 )
  If Not $pConditionSensor1 Then Return ConsoleWrite( "Property condition ERR Sensor1" & @CRLF )
  ConsoleWrite( "Property condition OK Sensor1" & @CRLF )

  ;Local $oSensor1
  ;_UIA_setVar("HOBO.sensor1","name:=416 SN:10485416; controltype:=ListItem" )
  ;$oSensor1=_UIA_getFirstObjectOfElement($hWindow,_UIA_getVar("HOBO.sensor2"),$treescope_subtree)
  ;_UIA_action("HOBO.sensor1","leftclick")


  ; Find "Save Data..." button
  Local $pButton, $oButton
  $oWindow.FindFirst( $TreeScope_Descendants, $pCondition, $pButton )
  $oButton = ObjCreateInterface( $pButton, $sIID_IUIAutomationElement, $dtagIUIAutomationElement )
 ; If Not IsObj( $oButton ) Then Return ConsoleWrite( "Find button ERR" & @CRLF )
 ; ConsoleWrite( "Find button OK" & @CRLF )

  ; Find "Get/Refresh Time Range" button
  Local $pButtonTime, $oButtonTime
  $oWindow.FindFirst( $TreeScope_Descendants, $pConditionTime, $pButtonTime )
  $oButtonTime = ObjCreateInterface( $pButtonTime, $sIID_IUIAutomationElement, $dtagIUIAutomationElement )
 ; If Not IsObj( $oButtonTime ) Then Return ConsoleWrite( "Find button ERR Time" & @CRLF )
 ; ConsoleWrite( "Find button OK Time" & @CRLF )

 ; Find "Generate Data Streams" button
  Local $pButtonGenerate, $oButtonGenerate
  $oWindow.FindFirst( $TreeScope_Descendants, $pConditionGenerate, $pButtonGenerate )
  $oButtonGenerate = ObjCreateInterface( $pButtonGenerate, $sIID_IUIAutomationElement, $dtagIUIAutomationElement )
 ; If Not IsObj( $oButtonGenerate ) Then Return ConsoleWrite( "Find button ERR Generate" & @CRLF )
 ; ConsoleWrite( "Find button OK Generate" & @CRLF )

; Find "416 SN: 10485416" sensor
  Local $pListSensor1, $oListSensor1
  $oWindow.FindFirst( $TreeScope_Descendants, $pConditionSensor1, $pListSensor1 )
  $oListSensor1 = ObjCreateInterface( $pListSensor1, $sIID_IUIAutomationElement, $dtagIUIAutomationElement )
  If Not IsObj( $oListSensor1 ) Then Return ConsoleWrite( "Find List ERR Sensor 1" & @CRLF )
  ConsoleWrite( "Find List OK Sensor 1" & @CRLF )





  ; Click (invoke) "416 SN: 10485416"
  Local $pInvokeSensor1, $oInvokeSensor1
  $oListSensor1.GetCurrentPattern( $UIA_InvokePatternId, $pInvokeSensor1 )
  $oInvokeSensor1 = ObjCreateInterface( $pInvokeSensor1, $sIID_IUIAutomationInvokePattern, $dtagIUIAutomationInvokePattern )
  If Not IsObj( $oInvokeSensor1) Then Return ConsoleWrite( "Invoke pattern ERR Sensor 1" & @CRLF )
  ConsoleWrite( "Invoke pattern OK Sensor 1" & @CRLF )
  $oInvokeSensor1.Invoke()

  ; Click (invoke) "Get/Refresh Time Range" button
  Local $pInvokeTime, $oInvokeTime
  $oButtonTime.GetCurrentPattern( $UIA_InvokePatternId, $pInvokeTime )
  $oInvokeTime = ObjCreateInterface( $pInvokeTime, $sIID_IUIAutomationInvokePattern, $dtagIUIAutomationInvokePattern )
  If Not IsObj( $oInvokeTime) Then Return ConsoleWrite( "Invoke pattern ERR Time" & @CRLF )
  ConsoleWrite( "Invoke pattern OK Time" & @CRLF )
  $oInvokeTime.Invoke()

 ; Click (invoke) "Generate Data Streams" button
  Local $pInvokeGenerate, $oInvokeGenerate
  $oButtonGenerate.GetCurrentPattern( $UIA_InvokePatternId, $pInvokeGenerate )
  $oInvokeGenerate = ObjCreateInterface( $pInvokeGenerate, $sIID_IUIAutomationInvokePattern, $dtagIUIAutomationInvokePattern )
  If Not IsObj( $oInvokeGenerate) Then Return ConsoleWrite( "Invoke pattern ERR Generate" & @CRLF )
  ConsoleWrite( "Invoke pattern OK Generate" & @CRLF )
  $oInvokeGenerate.Invoke()

; Click (invoke) "Save Data..." button
  Local $pInvoke, $oInvoke
  $oButton.GetCurrentPattern( $UIA_InvokePatternId, $pInvoke )
  $oInvoke = ObjCreateInterface( $pInvoke, $sIID_IUIAutomationInvokePattern, $dtagIUIAutomationInvokePattern )
  If Not IsObj( $oInvoke ) Then Return ConsoleWrite( "Invoke pattern ERR" & @CRLF )
  ConsoleWrite( "Invoke pattern OK" & @CRLF )
  $oInvoke.Invoke()

EndFunc