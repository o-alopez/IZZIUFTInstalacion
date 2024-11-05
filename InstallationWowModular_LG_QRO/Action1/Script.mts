'************************************************************************************************* @@ script infofile_;_ZIP::ssf240.xml_;_
'Script Description     :Main script which calls different relative modules to complete the flow
'Test Tool/Version		: Unified Functional Testing 14.50
'Application Automated	: Siebel CRM
'Author				    : Jose Luis Santiago Williams/Mariana Aldana Sanchez
'Date Created			: 10/10/2022 (mm/dd/yyyy)
'Date last Modified     : 10/10/2022
'**************************************************************************************************
'ADD20200805 JIEC CAMBIA RUTA DE ARCHIVO DE ENTRADA
 'strNameFile="..\datatables\InstallationDisney.xlsx"
 strNameFile="W:\flowsN\Instalacion\dtConComplementoN\InstallationWowModular_LG_QRO.xlsx"
 
If Datatable.GetSheet("Installation").GetCurrentRow = 1 Then

'* Load the Environment Variables includes Siebel URL, Username and Password
    fnLoadEnvironmentVariables()

    '*Creation of sheets
    DataTable.AddSheet("ProfileCreation")
    DataTable.AddSheet("ProductAvailability")
    DataTable.AddSheet("ProductSelection")
    DataTable.AddSheet("Output")
    
    
    '*Import of data
    Datatable.ImportSheet strNameFile,"Installation", "Installation"
    Datatable.ImportSheet strNameFile,"ProfileCreation", "ProfileCreation"
    Datatable.ImportSheet strNameFile,"ProductAvailability", "ProductAvailability"
    Datatable.ImportSheet strNameFile,"ProductSelection", "ProductSelection"
    Datatable.ImportSheet strNameFile,"Output", "Output"
    
    Systemutil.CloseProcessByName("SiebelAX_Test_Automation_21233.exe")
    
    'copia de seguridad ADDJLSW 19/01/20222
    Call CopiaSeguridad("W:\flowsN\RespaldoData\Instalacion\","InstallationWowModular_LG_QRO")
        
    call fnDisplay("p_Execute", "Installation")
    print  "                                     import      "
'    If i>1 Then
'      PRINT "saLIR "
'      Reporter.ReportEvent micFail, "Siebel Error", "sALIR POR QUE SON MUCHOS INTENTOS"
'    ExitAction
'    End If    
End  If

' ADD20200621 JIEC mostrar los status ejecutados

If Datatable.GetSheet("Installation").GetRowCount= Datatable.GetSheet("Installation").GetCurrentRow Then
	call fnDisplay("p_Execute", "Installation")
	
	Call fnExportToExcel(strNameFile,"Installation", "Installation")
	Call fnExportToExcel(strNameFile,"Output", "Output")
	
         print "               -----Export Extra ---                    .."		
End If


'* To check Whether the Script to be exuected or not with the flag value
'...................:::::::OPCIONES     ::::::::::.................................
'  Y--> DE INCIO A FIN LA CUENTA                                   ::::::::::::::::
'  PF--> CREA LA CUENTAS HASTA SELECCIONAR PRODUCTOS               :::::::::::::::: 
'  PE ---> PEGAR EQUIPOS, PROGRAMAR Y ENVIAR                       ::::::::::::::::
'  SEP --> TEST PROBAR SI SELECCIONAR PRODUCTOS(SOLAMENTE DEV)     ::::::::::::::::
'  CP --> PROFILE, DIRECCION Y CONTRATO (ANTEES DE NUEWVO ORDEN)
' CPP  --> PROFILE HASTA AGREGAR PORTAFOLIO
'...................................................................................

If DataTable.Value ("p_Execute", "Installation") = uCase("MSG") or  DataTable.Value ("p_Execute", "Installation") = uCase("CPP") or DataTable.Value ("p_Execute", "Installation") = uCase("Y") or DataTable.Value ("p_Execute", "Installation") = uCase("YA") or DataTable.Value ("p_Execute", "Installation") = uCase("PF") or DataTable.Value ("p_Execute", "Installation") = uCase("PE") or DataTable.Value ("p_Execute", "Installation") = uCase("PEA") or DataTable.Value ("p_Execute", "Installation") = uCase("PEA2") or DataTable.Value ("p_Execute", "Installation") = uCase("SEP") or DataTable.Value ("p_Execute", "Installation") = uCase("CP") Then    
'If DataTable.Value ("p_Execute", "Installation") = uCase("SEP") Then    

		 '*Create the data table objects
		    Set dtInstall = Datatable.GetSheet("Installation")
		    Set dtProfileCreation = Datatable.GetSheet("ProfileCreation")
		    Set dtProductAvailability = Datatable.GetSheet("ProductAvailability")
		    Set dtProductSelection = Datatable.GetSheet("ProductSelection")
		    Set dtOutput = Datatable.GetSheet("Output")
		    
		    
		    '*Set the current rows for each Tab based on the current row of the Casos tab
		    dtProfileCreation.SetCurrentRow(dtInstall.GetCurrentRow)
		    dtProductAvailability.SetCurrentRow(dtInstall.GetCurrentRow)
		    dtProductSelection.SetCurrentRow(dtInstall.GetCurrentRow)
		    dtOutput.SetCurrentRow(dtInstall.GetCurrentRow)
		     
		     print dtInstall.GetCurrentRow&"  ......::::::::  "&DataTable.Value ("TestCase", "Installation") 
'              ?Escribir en archivo LOg
             strLog=dtInstall.GetCurrentRow&"  ......::::::::  "&DataTable.Value ("TestCase", "Installation")&":::: ..."&DataTable.Value ("p_Execute", "Installation") 
             fnUtilWriteLog strLog   
             
'			Systemutil.CloseProcessByName("SiebelAX_Test_Automation_21233.exe")
				sUserName = Environment.Value ("gUserChromeDEV")
				sPassword = Environment.Value ("gPasswordChromeDEV")
			'Call LoginToSiebel(sUserName, sPassword)
			If DataTable.Value ("p_Execute", "Installation") = "PEA2" Then
				sUserName = "TSTCUENTASESPECIALES"'Environment.Value ("gUserProfileiZZi")
				sPassword = "TSTCUENTASESPECIALES"'Environment.Value ("gPasswordProfileiZZi")
			End If
			 Call fnLoginToSiebelNCDesarrollo(sUserName, sPassword) 'LoginToSiebelNC
          '			Call LoginToSiebelSmartCHROME(sUserName, sPassword)
			
			DataTable.Value ("o_TestCase", "Output") =DataTable.Value ("TestCase", "Installation") 
			

				Select Case DataTable.Value ("p_Execute", "Installation") 
			
				Case "Y","YA"
					print "                  "&DataTable.Value ("p_Execute", "Installation")
								Call fnProfilecreationNewC()
								Call fnAddressCreateWowC()
								Call fnBillingTypeC()
								Call fnProductSelectionC()
								Call fnPersonalizerWowC("ProductSelection")
								'Call fnConsultarOrdenServicio()
								
								finalstatus=Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebEdit("acc_name:=Estado","name:=s_1_1_.*_0","role:=combobox").GetROProperty("value")
								Datatable.Value("o_Status","Output") = finalstatus
								DataTable.Value ("p_Execute", "Installation") = "Y;PartePFCompleto"
								
								              	Call PegadodeTelefono()
                                   
                                    Call subPegarTag(DataTable.Value ("p_Tag", "ProductSelection")) 
                                   
                                    Call subPegarNombreContrato(DataTable.Value ("o_ContractoNo", "Output"))
                              
                                    Call fhPersonalizerDatosC("ProductSelection")
                                    
                                    Call Programar()
 									
 									Call PegadoDeEquipo(DataTable.Value ("p_SerialVideo", "ProductSelection"), DataTable.Value ("p_SerialEMTA", "ProductSelection"), DataTable.Value ("p_SerialEMTA", "ProductSelection"))
 									
 									wait 10
 									'Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebButton("acc_name:=Detalles:Enviar", "html id:=s_4_1_1_0_Ctrl").Click
								finalstatus=Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebEdit("acc_name:=Estado","name:=s_1_1_.*_0","role:=combobox").GetROProperty("value")
								Datatable.Value("o_Status","Output") = finalstatus
								DataTable.Value ("p_Execute", "Installation") = "Y;Completo"
		'						valorpro = DataTable.Value ("p_Portafolio", "ProductSelection") 
		'						If valorpro <> "Portafolio Productos Negocios 2_0" Then
		'						Call fnPersonalizerWow("ProductSelection")
		'						else
		'						Call fnPersonalizerNegociosDOS()
        '                        End If
                                                                '20201228  JIEC se mueve aqui flow principal
		'					        Call fhPersonalizerDatos("ProductSelection")
		'						
		'						If DataTable.Value ("p_Execute", "Installation") = uCase("YA") Then
		'								valorpro = DataTable.Value ("p_Portafolio", "ProductSelection")
		'										If valorpro <> "Portafolio Productos Negocios 2_0" Then
		'							                     	If valorpro = "Portafolio Productos wizzplus wow" Then
								'	                           Call fnFinalSubmissionWizzAA()
		'							                        End If
		'							                     Call fnFinalSubmissionWowAAF("ProductSelection")
		'					                         else
		'					                             Call fnFinalSubmissionNegocioAA()
		'					                         End If
										'Call fnFinalSubmissionWowAAF("ProductSelection")
		'								If fnFindOS(DataTable.Value ("o_OrderService", "Output")) Then
		'								     Call fhPersonalizerDatos("ProductSelection")									
		'								End If
		'						End If
		'						valorpro = DataTable.Value ("p_Portafolio", "ProductSelection") 
		'						If valorpro <> "Portafolio Productos Negocios 2_0" Then
		'							If valorpro = "Portafolio Productos wizzplus wow" Then
		'							   Call fnFinalSubmissionWizz()	
		'							End If
		'							Call fnFinalSubmissionWowI("ProductSelection")
		'					    else
		'					        Call fnFinalSubmissionNegocio()
		'					    End If
								'Call fnInventoryValidationResidencial()
			
				Case "PF"
					print "                  "&DataTable.Value ("p_Execute", "Installation")
					If true Then
						        Call fnProfilecreationNewC()
								Call fnAddressCreateWowC()
								Call fnBillingTypeC()
								Call fnProductSelectionC()
								Call fnPersonalizerWowC("ProductSelection")
								'Call fnConsultarOrdenServicio()
								finalstatus=Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebEdit("acc_name:=Estado","name:=s_1_1_.*_0","role:=combobox").GetROProperty("value")
								Datatable.Value("o_Status","Output") = finalstatus
								DataTable.Value ("p_Execute", "Installation") = "PF;Pendiente"
								'Call fnProfilecreationNew()
								'Call fnAddressCreateSIPRE()  'fnAddressCreateWow() antes tenia este pero no funciona bien desde 08092021 20078959 20081740
								'Call fnBillingType()
								'Call fnProductSelection()
					End If
								'valorpro = DataTable.Value ("p_Portafolio", "ProductSelection") 
								'If valorpro <> "Portafolio Productos Negocios 2_0" Then
								'Call fnPersonalizerWow("ProductSelection")
								'else
								'Call fnPersonalizerNegociosDOS()
                                'End If '20201228  JIEC se mueve aqui flow principal
								'Call fhPersonalizerDatos("ProductSelection")
						
				Case "SEP"
					print "                  "&DataTable.Value ("p_Execute", "Installation")
'														
					             If DataTable.Value ("o_Numeroid", "Output")<>"" Then
					             	SearchAccountSBL DataTable.Value ("o_Numeroid", "Output")
					                HandleSiebelDialogErrorAndContinue()
					                'Profundizar en la cuenta
					                
					                SiebApplication("Siebel Communications").SiebScreen("Cuentas").SiebView("Resumen de la compaÃ±Ã­a").SiebApplet("Ordenes de Servicio").SiebList("List").DrillDownColumn "Order Number (eService)",0
					                ' AGREGAR PORTAFOLIO 
					                
					                Call fnAgregarPortafolio()  ' Revisa si tiene motivo de la orden, agrega portafolio y personaliza
					                
					                Call fnPersonalizerWow("ProductSelection")
					                '20201228  JIEC
									Call fhPersonalizerDatos("ProductSelection")
					             End If
					             
					              
					
					
				Case "PE"
					print "                  "&DataTable.Value ("p_Execute", "Installation")
								If fnFindOS(DataTable.Value ("o_OrderService", "Output")) Then
									
                                    'ADD20191101 JIEC personalizar Complemento
                                    finalstatus=SiebApplication("Siebel Communications").SiebScreen("Ordenes de Servicio").SiebView("Detalles").SiebApplet("Orden de servicio").SiebPicklist("Estado").GetROProperty("activeitem")
                                    If finalstatus<>"Enviada" and finalstatus<>"Completa" Then
                                    			Call fhPersonalizerDatos("ProductSelection")									
												'Call fnFinalSubmissionWowI("ProductSelection")
												'Call fnFinalSubmissionNegocio()
												valorpro = DataTable.Value ("p_Portafolio", "ProductSelection")
												If valorpro <> "Portafolio Productos Negocios 2_0" Then
									                	If valorpro = "Portafolio Productos wizzplus wow" Then
									                       Call fnFinalSubmissionWizz()	
									                    End If
									                Call fnFinalSubmissionWowI("ProductSelection")
							                    else
							                        Call fnFinalSubmissionNegocio()
							                    End If
									ElseIf finalstatus="Completa" Then
									    osFecha=SiebApplication("Siebel Communications").SiebScreen("Ordenes de Servicio").SiebView("Detalles").SiebApplet("Orden de servicio").SiebText("Ultima Modificacion").GetROProperty("text")
										print 	osFecha
										Datatable.Value("o_OSFecha","Output") = osFecha
										'Store the Order Status in Output Sheet in Data Table
										Datatable.Value("o_Status","Output") = finalstatus
									Else 
									     call fnProgramarOS()
                                         Call fmSendSOWow()									
                                    End If

					'				Call fnInventoryValidationResidencial()
								End If
				Case "PEA"
					print "                  "&DataTable.Value ("p_Execute", "Installation")
								'If fnFindOS(DataTable.Value ("o_OrderService", "Output")) Then
								If fnFindOSG(DataTable.Value ("o_OrderService", "Output")) Then
								
 @@ script infofile_;_ZIP::ssf235.xml_;_
                                    'ADD20191101 JIEC personalizar Complemento
 @@ script infofile_;_ZIP::ssf198.xml_;_
                                  	Call PegadodeTelefono() @@ script infofile_;_ZIP::ssf239.xml_;_
                                   
                                    Call subPegarTag(DataTable.Value ("p_Tag", "ProductSelection")) 
                                   
                                    Call subPegarNombreContrato(DataTable.Value ("o_ContractoNo", "Output"))
                              
                                    Call fhPersonalizerDatosC("ProductSelection")
                                    
                                    '..........................................
                                    if DynamicWaitBotonProgramar_2( 30 ) then
                                       print "Entro a programar"
                                       Call Programar()
									wait 1
                                    else
                                       strMensaje = "No encontro el boton programar"
                                       print mensaje
                                       Datatable.Value("o_error","Output") = strMensaje
			                           DataTable.Value ("p_Execute", "Installation") = DataTable.Value ("p_Execute", "Installation")&";"&strMensaje
									   wait 1
									   ExitActioniteration
                                    End If
                                    '..........................................

                                    
 									
 									Call PegadoDeEquipo(DataTable.Value ("p_SerialVideo", "ProductSelection"), DataTable.Value ("p_SerialEMTA", "ProductSelection"), DataTable.Value ("p_SerialEMTA", "ProductSelection"))
 									
 									'wait 10
 									'Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebButton("acc_name:=Detalles:Enviar", "html id:=s_4_1_1_0_Ctrl").Click
 									wait 3
 									wait 3
 									wait 3
 									wait 1 									
								    'Call FuncionesFinales("FIN")
								    Call FuncionesFinales_2("FIN")
 									wait 3
 									wait 3
 									wait 3
 									wait 1
                                    'wait 30 									
								    
                                    finalstatus=Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebEdit("acc_name:=Estado","name:=s_1_1_.*_0","role:=combobox").GetROProperty("value")
                                    Datatable.Value("o_Status","Output") = finalstatus
                                    DataTable.Value ("p_Execute", "Installation") = DataTable.Value ("p_Execute", "Installation")&";"&finalstatus

 @@ script infofile_;_ZIP::ssf241.xml_;_
'                                    finalstatus=SiebApplication("Siebel Communications").SiebScreen("Ordenes de Servicio").SiebView("Detalles").SiebApplet("Orden de servicio").SiebPicklist("Estado").GetROProperty("activeitem")
'                                    If finalstatus<>"Enviada" and finalstatus<>"Completa" Then
'                                    			Call fhPersonalizerDatosC("ProductSelection")									
'												valorpro = DataTable.Value ("p_Portafolio", "ProductSelection")
'												If valorpro <> "Portafolio Productos Negocios 2_0" Then
'									                     	If valorpro = "Portafolio Productos wizzplus wow" Then
'									                           Call fnFinalSubmissionWizzAA()
'									                        End If
'									                     Call fnFinalSubmissionWowAAF("ProductSelection")
'							                         else
'							                             Call fnFinalSubmissionNegocioAA()
'							                         End If
'												If fnFindOS(DataTable.Value ("o_OrderService", "Output")) Then
'												     Call fhPersonalizerDatos("ProductSelection")									
'												     'Call fnFinalSubmissionWowI("ProductSelection")
'												     'Call fnFinalSubmissionNegocio()
'												     valorpro = DataTable.Value ("p_Portafolio", "ProductSelection")
'												     If valorpro <> "Portafolio Productos Negocios 2_0" Then
'									                     	If valorpro = "Portafolio Productos wizzplus wow" Then
'									                           Call fnFinalSubmissionWizz()	
'									                        End If
'									                     Call fnFinalSubmissionWowI("ProductSelection")
'							                         else
'							                             Call fnFinalSubmissionNegocio()
'							                         End If
'												End If
'									ElseIf finalstatus="Completa" Then
'									    osFecha=SiebApplication("Siebel Communications").SiebScreen("Ordenes de Servicio").SiebView("Detalles").SiebApplet("Orden de servicio").SiebText("Ultima Modificacion").GetROProperty("text")
'										print 	osFecha
'										Datatable.Value("o_OSFecha","Output") = osFecha
'										'Store the Order Status in Output Sheet in Data Table
'										Datatable.Value("o_Status","Output") = finalstatus
'									Else 
'									     call fnProgramarOS()
'                                         Call fmSendSOWow()									
'                                    End If
'
'					'				Call fnInventoryValidationResidencial()
								End If
				Case "PEA2"
					print "                  "&DataTable.Value ("p_Execute", "Installation")
								'If fnFindOS(DataTable.Value ("o_OrderService", "Output")) Then
								If fnFindOSG(DataTable.Value ("o_OrderService", "Output")) Then
								
 @@ script infofile_;_ZIP::ssf235.xml_;_
                                    'ADD20191101 JIEC personalizar Complemento
                                     wait 2
 Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("acc_name:=CIC Potencial").Set Datatable.Value("p_CIC_Potencial","ProductSelection")
 wait 3
 If not Browser("Todas las cuentas").Page("Todas las cuentas").WebButton("Detalles:Agregar Items").Exist(2) Then
Set oSave = CreateObject("Wscript.Shell")
oSave.SendKeys "{ENTER}"
Set oSave = Nothing
 End If
 wait 1
 @@ script infofile_;_ZIP::ssf198.xml_;_
                                  	Call PegadodeTelefono() @@ script infofile_;_ZIP::ssf239.xml_;_
                                   
                                    Call subPegarTag(DataTable.Value ("p_Tag", "ProductSelection")) 
                                   
                                    Call subPegarNombreContrato(DataTable.Value ("o_ContractoNo", "Output"))
                              
                                    Call fhPersonalizerDatosC("ProductSelection")
                                    
                                    Call Programar()
 									
 									Call PegadoDeEquipo(DataTable.Value ("p_SerialVideo", "ProductSelection"), DataTable.Value ("p_SerialEMTA", "ProductSelection"), DataTable.Value ("p_SerialEMTA", "ProductSelection"))
 									
 									wait 10
 									'Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebButton("acc_name:=Detalles:Enviar", "html id:=s_4_1_1_0_Ctrl").Click
								    Call FuncionesFinales("FIN")
                                    finalstatus=Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebEdit("acc_name:=Estado","name:=s_1_1_.*_0","role:=combobox").GetROProperty("value")
                                    Datatable.Value("o_Status","Output") = finalstatus
                                    DataTable.Value ("p_Execute", "Installation") = DataTable.Value ("p_Execute", "Installation")&";"&finalstatus

 @@ script infofile_;_ZIP::ssf241.xml_;_
'                                    finalstatus=SiebApplication("Siebel Communications").SiebScreen("Ordenes de Servicio").SiebView("Detalles").SiebApplet("Orden de servicio").SiebPicklist("Estado").GetROProperty("activeitem")
'                                    If finalstatus<>"Enviada" and finalstatus<>"Completa" Then
'                                    			Call fhPersonalizerDatosC("ProductSelection")									
'												valorpro = DataTable.Value ("p_Portafolio", "ProductSelection")
'												If valorpro <> "Portafolio Productos Negocios 2_0" Then
'									                     	If valorpro = "Portafolio Productos wizzplus wow" Then
'									                           Call fnFinalSubmissionWizzAA()
'									                        End If
'									                     Call fnFinalSubmissionWowAAF("ProductSelection")
'							                         else
'							                             Call fnFinalSubmissionNegocioAA()
'							                         End If
'												If fnFindOS(DataTable.Value ("o_OrderService", "Output")) Then
'												     Call fhPersonalizerDatos("ProductSelection")									
'												     'Call fnFinalSubmissionWowI("ProductSelection")
'												     'Call fnFinalSubmissionNegocio()
'												     valorpro = DataTable.Value ("p_Portafolio", "ProductSelection")
'												     If valorpro <> "Portafolio Productos Negocios 2_0" Then
'									                     	If valorpro = "Portafolio Productos wizzplus wow" Then
'									                           Call fnFinalSubmissionWizz()	
'									                        End If
'									                     Call fnFinalSubmissionWowI("ProductSelection")
'							                         else
'							                             Call fnFinalSubmissionNegocio()
'							                         End If
'												End If
'									ElseIf finalstatus="Completa" Then
'									    osFecha=SiebApplication("Siebel Communications").SiebScreen("Ordenes de Servicio").SiebView("Detalles").SiebApplet("Orden de servicio").SiebText("Ultima Modificacion").GetROProperty("text")
'										print 	osFecha
'										Datatable.Value("o_OSFecha","Output") = osFecha
'										'Store the Order Status in Output Sheet in Data Table
'										Datatable.Value("o_Status","Output") = finalstatus
'									Else 
'									     call fnProgramarOS()
'                                         Call fmSendSOWow()									
'                                    End If
'
'					'				Call fnInventoryValidationResidencial()
								End If				
				Case "CP"
					print "                  "&DataTable.Value ("p_Execute", "Installation")
					strMsg="-"
		            Call fnUtilDTWriterMsg("Installation","p_Execute",strMsg)
		        
								Call fnProfilecreationNew()
								Call fnAddressCreateWow()
								Call fnBillingType()
'								Call fnProductSelection() No porque se necesita hasta antes de portafolio
										
				Case "CPP"
					print "                  "&DataTable.Value ("p_Execute", "Installation")
					strMsg="-"
		            Call fnUtilDTWriterMsg("Installation","p_Execute",strMsg)
		        
								Call fnProfilecreationNew()
								Call fnAddressCreateWow()
								Call fnBillingType()
								Call fnProductSelection()
										
				Case Else
					print "         No existe opcion:          "&DataTable.Value ("p_Execute", "Installation")
					ExitActioniteration	
					
				End Select
		
	    print  "        -----  "& DataTable.Value ("o_Status", "Output")   	
		strMsg=" "&DataTable.Value ("o_Status", "Output")
		'Call fnUtilDTWriterMsg("Installation","p_Execute",strMsg)
		print "                   .."
		'         Escirbir archivo log
'			o_TestCase	o_Numeroid	o_OrderService	o_Status	o_ContractoNo

         strLog=DataTable.Value ("o_TestCase", "Output")&" "&DataTable.Value ("o_Numeroid", "Output")&" "&DataTable.Value ("o_OrderService", "Output")&" "&DataTable.Value ("o_Status", "Output")&" "&DataTable.Value ("o_ContractoNo", "Output")
         fnUtilWriteLog strLog
		 
	    Call fnExportToExcel(strNameFile,"Installation", "Installation")
	    Call fnExportToExcel(strNameFile,"Output", "Output")
         print "                                   .."		
		
Else
	'If the Flag is not "Y" then exit the Test
	ExitActionIteration

End if
'----------------------------------------------------------------
'----------------------------------------------------------------
'----------------------------------------------------------------
Call Final("Installation","W:\flowsN\RespaldoDataFinal\Instalacion\","InstallationWow_ConComplementoN","","")

''* To check Whether the Script to be exuected or not with the flag value
'
'If DataTable.Value ("p_Execute", "Installation") = uCase("Y") or DataTable.Value ("p_Execute", "Installation") = uCase("PF") Then    
'    
'    '*Create the data table objects
'    Set dtInstall = Datatable.GetSheet("Installation")
'    Set dtProfileCreation = Datatable.GetSheet("ProfileCreation")
'    Set dtProductAvailability = Datatable.GetSheet("ProductAvailability")
'    Set dtProductSelection = Datatable.GetSheet("ProductSelection")
'    Set dtOutput = Datatable.GetSheet("Output")
'    
'    
'    '*Set the current rows for each Tab based on the current row of the Casos tab
'    dtProfileCreation.SetCurrentRow(dtInstall.GetCurrentRow)
'    dtProductAvailability.SetCurrentRow(dtInstall.GetCurrentRow)
'    dtProductSelection.SetCurrentRow(dtInstall.GetCurrentRow)
'    dtOutput.SetCurrentRow(dtInstall.GetCurrentRow)
'     
'     print dtInstall.GetCurrentRow&"  ......::::::::  "&DataTable.Value ("TestCase", "Installation") 
'
''** Calling the Functions in Sequence to execute the End to End flow of Installation Scenario
'															'If DataTable.Value("p_Category","ProductSelection") = "izzi WOW" Then
'															'	sUserName = Environment.Value ("gUserProfileWow")
'															'	sPassword = Environment.Value ("gPasswordProfileWow")
'															   
'															'else
'   				sUserName = Environment.Value ("gUserProfileiZZi")
'				sPassword = Environment.Value ("gPasswordProfileiZZi")
'															 '   sUserName = Environment.Value ("gUserProfile1")
'															'	sPassword = Environment.Value ("gPasswordProfile1")
'											'End If
'	Dim blnDebug
'	blnDebug = false				
'	If blnDebug = False Then
'								
'									Systemutil.CloseProcessByName("SiebelAX_Test_Automation_21233.exe")
'									Call LoginToSiebel(sUserName, sPassword)
'									
'									DataTable.Value ("o_TestCase", "Output") =DataTable.Value ("TestCase", "Installation") 
'										
'									Call fnProfilecreationNew()
'									Call fnAddressCreateWow()
'									Call fnBillingType()
'									Call fnProductSelection()
'								
'								
'									Call fnPersonalizerWow("ProductSelection")
'						   
'	End If   'Solo para desarrollar		
'                  'Call fhPersonalizerDatos("ProductSelection")
'         	       If not DataTable.Value ("p_Execute", "Installation") = uCase("PF") Then ' Cuando es PF se queda en seleccionar el Producto y no se pega equipos
'						 	 
'							    Call fnFinalSubmissionWow("ProductSelection")
'								' Call fnInventoryValidationResidencial()
'	      			 End If	
'
'				       	'ADD20190718 GUARDAR ESTATUS EN hoja Installation 
'					strMsg=" "&DataTable.Value ("o_Status", "Output")
'					Call fnUtilDTWriterMsg("Installation","p_Execute",strMsg)
'					
'				    Call fnExportToExcel("..\datatables\InstallationWow.xlsx","Installation", "Installation")
'				    Call fnExportToExcel("..\datatables\InstallationWow.xlsx","Output", "Output")
'			    
'
'ElseIf DataTable.Value ("p_Execute", "Installation") = uCase("PE") Then   
'					''  *****************************************************************
'					'
'					'   Este parte de codigo es para un reproceso 
'					'  Continua desde pegado de equipo ...
'					' *****************************************************************
'					
'					      '*Create the data table objects
'					    Set dtInstall = Datatable.GetSheet("Installation")
'					    Set dtProfileCreation = Datatable.GetSheet("ProfileCreation")
'					    Set dtProductAvailability = Datatable.GetSheet("ProductAvailability")
'					    Set dtProductSelection = Datatable.GetSheet("ProductSelection")
'					    Set dtOutput = Datatable.GetSheet("Output")
'					    
'					    
'					    '*Set the current rows for each Tab based on the current row of the Casos tab
'					    dtProfileCreation.SetCurrentRow(dtInstall.GetCurrentRow)
'					    dtProductAvailability.SetCurrentRow(dtInstall.GetCurrentRow)
'					    dtProductSelection.SetCurrentRow(dtInstall.GetCurrentRow)
'					    dtOutput.SetCurrentRow(dtInstall.GetCurrentRow)
'					    
'					'----------------------------------------------------------------------------------------------------------------------    
'					     print dtInstall.GetCurrentRow&"  ......::::REPROCESO::::::: ::::  "&DataTable.Value ("TestCase", "Installation") 
'					     
'
'					        sUserName = Environment.Value ("gUserProfileiZZi")
'						    sPassword = Environment.Value ("gPasswordProfileiZZi")
'					
'								Systemutil.CloseProcessByName("SiebelAX_Test_Automation_21233.exe")
'								Call LoginToSiebel(sUserName, sPassword)
'								If fnFindOS(DataTable.Value ("o_OrderService", "Output")) Then
'									
'                                    'ADD20191101 JIEC personalizar Complemento
'									Call fhPersonalizerDatos("ProductSelection")									
'									Call fnFinalSubmissionWow("ProductSelection")
'
'					'				Call fnInventoryValidationResidencial()
'								End If
'							
'					
'					       	'ADD20190718 GUARDAR ESTATUS EN hoja Installation 
'						strMsg=" "&DataTable.Value ("o_Status", "Output")
'						Call fnUtilDTWriterMsg("Installation","p_Execute",strMsg)
'						
'					    Call fnExportToExcel("..\datatables\InstallationWow.xlsx","Installation", "Installation")
'					    Call fnExportToExcel("..\datatables\InstallationWow.xlsx","Output", "Output")
'					     
'     
'Else
'	'If the Flag is not "Y" then exit the Test
'	ExitActionIteration
'
'End if

Function fnAddressCreateSIPRE()
'*************************************************************************************************
'Function Description   : Address Selection for Billing Address and Facturation
'Test Tool/Version		: Unified Functional Testing 14.50
'Application Automated	: Siebel CRM
'Author				    : Jose Luis Santiago Williams
'Date Created			: 08/09/2021 (mm/dd/yyyy)
'Date last Modified     : 08/09/2021
'**************************************************************************************************
On error resume next
Dim numero 'Numero id value retrieved from Previous Action

'* Navigate back to Cuentas Tab and Search the Numero id again
SiebApplication("Siebel Communications").SiebPageTabs("PageTabs").Dynamicwait(20)
SiebApplication("Siebel Communications").SiebPageTabs("PageTabs").GotoScreen "Accounts Screen"
SiebApplication("Siebel Communications").SiebScreen("Cuentas").SiebView("Todas las compañías").SiebApplet("Cuentas").SiebButton("Consulta").Click
numero = Datatable.value("o_Numeroid","Output")
wait 1 
print "Buscando cuenta para registrar direccion ... " + numero 
'* Enter the numero id and Click enter
SiebApplication("Siebel Communications").SiebScreen("Cuentas").SiebView("Todas las compañías").SiebApplet("Cuentas").SiebList("List").SiebText("Nro. Cuenta").SetText numero
SiebApplication("Siebel Communications").SiebScreen("Cuentas").SiebView("Todas las compañías").SiebApplet("Cuentas").SiebList("List").SiebText("Nro. Cuenta").ProcessKey "EnterKey"
wait 1

SiebApplication("Siebel Communications").SiebScreen("Cuentas").SiebScreenViews("ScreenViews").Dynamicwait(20)

'----------------------------------------------

'ADDD20191113 JIEC 
subPageUPx2()
SiebApplication("Siebel Communications").SiebScreen("Cuentas").SiebScreenViews("ScreenViews").Goto "CV All Account Capture View","L2"

'* Activate the numer id from Search Result
SiebApplication("Siebel Communications").SiebScreen("Cuentas").SiebView("Todas las cuentas").SiebApplet("Cuenta").SiebPicklist("Giro Negocio").ActiveItem

'* Navigate with pagedown
Set oPagedn = Createobject("Wscript.Shell")
oPagedn.SendKeys "{PgDn}"
oPagedn.SendKeys "{PgDn}"
Set oPagedn = Nothing

'From Direcciones de la Cuenta Applet, Click on Neuova Button
If SiebApplication("Siebel Communications").SiebScreen("Cuentas").SiebView("Todas las cuentas").SiebApplet("Direcciones de la Cuenta").Exist(5) Then
  SiebApplication("Siebel Communications").SiebScreen("Cuentas").SiebView("Todas las cuentas").SiebApplet("Direcciones de la Cuenta").SiebButton("Nuevo").Click
  SiebApplication("Siebel Communications").SiebScreen("Cuentas").SiebView("Todas las cuentas").SiebApplet("Direcciones de la Cuenta").SiebPicklist("Tipo Dirección").Select "Servicio"
End If

Set oPagedn = Createobject("Wscript.Shell")
oPagedn.SendKeys "{PgDn}"
Set oPagedn = Nothing

'Click on Calle and Enter the Postcode and RPT Details from Datasheet
SiebApplication("Siebel Communications").SiebScreen("Cuentas").SiebView("Todas las cuentas").SiebApplet("Direcciones de la Cuenta").SiebText("Calle").OpenPopup
wait 3

If Datatable.Value("p_Sipre_id","ProfileCreation")<>"" Then
'   SiebApplication("Siebel Communications").SiebScreen("Cuentas").SiebView("Todas las cuentas").SiebApplet("Nombres Calles").SiebText("CV SIPRE Id").SetText Datatable.Value("p_Sipre_id","ProfileCreation")
SiebApplication("Siebel Communications").SiebScreen("Cuentas").SiebView("Todas las cuentas").SiebApplet("Nombres Calles").SiebPicklist("SiebPicklist").Select "CV SIPRE Id"
SiebApplication("Siebel Communications").SiebScreen("Cuentas").SiebView("Todas las cuentas").SiebApplet("Nombres Calles").SiebText("SiebText").SetText Datatable.Value("p_Sipre_id","ProfileCreation")
SiebApplication("Siebel Communications").SiebScreen("Cuentas").SiebView("Todas las cuentas").SiebApplet("Nombres Calles").SiebButton("Ir_2").Click
else
SiebApplication("Siebel Communications").SiebScreen("Cuentas").SiebView("Todas las cuentas").SiebApplet("Nombres Calles").SiebButton("Consulta").Dynamicwait(10)
SiebApplication("Siebel Communications").SiebScreen("Cuentas").SiebView("Todas las cuentas").SiebApplet("Nombres Calles").SiebButton("Consulta").Click

End If

If Datatable.Value("p_PostalCode","ProfileCreation") <> "" and Datatable.Value("p_Sipre_id","ProfileCreation") = "" Then
	'Browser("SiebWebPopupWindow").Page("SiebWebPopupWindow").Frame("_swepopcontent Frame").WinEdit("RPT").Set Datatable.value("p_RPTCode","ProfileCreation")
	'SiebApplication("Siebel Communications").SiebScreen("Cuentas").SiebView("Todas las cuentas").SiebApplet("Nombres Calles").SiebPicklist("Codigo RPT").Select Datatable.value("p_RPTCode","ProfileCreation")
    Browser("SiebWebPopupWindow").Page("SiebWebPopupWindow").Frame("_swepopcontent Frame").WinEdit("CodigoPostal").Set Datatable.Value("p_PostalCode","ProfileCreation")
End If
'print "       p_RPTCode                          "&Datatable.value("p_RPTCode","ProfileCreation")
If Datatable.value("p_RPTCode","ProfileCreation") <> "" and Datatable.Value("p_Sipre_id","ProfileCreation") = "" Then
	'Browser("SiebWebPopupWindow").Page("SiebWebPopupWindow").Frame("_swepopcontent Frame").WinEdit("RPT").Set Datatable.value("p_RPTCode","ProfileCreation")
	SiebApplication("Siebel Communications").SiebScreen("Cuentas").SiebView("Todas las cuentas").SiebApplet("Nombres Calles").SiebPicklist("Codigo RPT").Select Datatable.value("p_RPTCode","ProfileCreation")

End If


'* From the Search result , Select the first Address displayed
SiebApplication("Siebel Communications").SiebScreen("Cuentas").SiebView("Todas las cuentas").SiebApplet("Nombres Calles").SiebButton("Ir").Click
SiebApplication("Siebel Communications").SiebScreen("Cuentas").SiebView("Todas las cuentas").SiebApplet("Nombres Calles").SiebList("List").ActivateRow 1
SiebApplication("Siebel Communications").SiebScreen("Cuentas").SiebView("Todas las cuentas").SiebApplet("Nombres Calles").SiebButton("Aceptar").Dynamicwait(10)
SiebApplication("Siebel Communications").SiebScreen("Cuentas").SiebView("Todas las cuentas").SiebApplet("Nombres Calles").SiebButton("Aceptar").Click

'Press Ctrl+S to save the record
Set oSave = Createobject("Wscript.Shell")
oSave.SendKeys "^s"
Set oSave = Nothing

wait 10
'Wait for Application to Sync/Response after ctrl+s
SiebApplication("Siebel Communications").SiebScreen("Cuentas").SiebView("Todas las cuentas").SiebApplet("Cuentas").SiebList("List").Dynamicwait(180)

'Activate the Applet to Focus
SiebApplication("Siebel Communications").SiebScreen("Cuentas").SiebView("Todas las cuentas").SiebApplet("Cuentas").SiebList("List").DoubleClick

'Navigate below with pagedown action
Set oSave1 = Createobject("Wscript.Shell")
oSave1.SendKeys "{PgDn}"
Set oSave1 = Nothing

'Navigate below with pagedown action
Set oPageDown = Createobject("Wscript.Shell")
oPageDown.SendKeys "{PgDn}"
Set oPageDown = Nothing

'Activate the applet to set focus ,before passing ctrl+B in the next step
SiebApplication("Siebel Communications").SiebScreen("Cuentas").SiebView("Todas las cuentas").SiebApplet("Direcciones de la Cuenta").SiebList("List").ActivateRow 0
SiebApplication("Siebel Communications").SiebScreen("Cuentas").SiebView("Todas las cuentas").SiebApplet("Direcciones de la Cuenta").SiebList("List").Dynamicwait(10)
SiebApplication("Siebel Communications").SiebScreen("Cuentas").SiebView("Todas las cuentas").SiebApplet("Direcciones de la Cuenta").SiebList("List").DoubleClick 0
wait 2
'Press ctrl+b to copy the address for Facturation
Set Octrlb = Createobject("Wscript.Shell")
Octrlb.SendKeys "^b"
Set Octrlb= Nothing

wait 5
'Allow the application to Sync, until ctrl+b action is saved
SiebApplication("Siebel Communications").SiebScreen("Cuentas").SiebView("Todas las cuentas").SiebApplet("Direcciones de la Cuenta").SiebList("List").Dynamicwait(180)

'Press Ctrl+S to save the record
Set oSave = Createobject("Wscript.Shell")
oSave.SendKeys "^s"
Set oSave = Nothing

wait 10
'** This event to ensure the Siebelframe for Order de Servicio is visible**
Set oPageEnd = Createobject("Wscript.Shell")
oPageEnd.SendKeys "{End}"
oPageEnd.SendKeys "{PgUp}"
Set oPageEnd = Nothing

End Function


'-----------------------------------
Function LoginToSiebelNC (strUserName, strPass)
	   
	   SystemUtil.CloseProcessByName "SiebelAx_Test_Automation_21233.exe"
	   print "                       Iniciando Login to Siebel ....  " 
 	   	      SystemUtil.CloseProcessByName "chrome.exe"'"iexplore.exe"
			'https://crmqa.izzi.mx/siebel/app/ecommunications/esn?
			'https://172.19.138.222:9011/siebel/app/ecommunications/esn?
			'SystemUtil.Run "chrome.exe", Environment.value("gURLSiebelQACHROME"), , ,3
			SystemUtil.Run "chrome.exe", "https://172.19.138.74:9011/siebel/app/eCommunications/esn", , ,3
		    wait 10
			If not Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("SWEUserName").Exist(5) Then
				reporter.ReportEvent micFail, "Access to Siebel", "Siebel system is down"
				ExitActionIteration
			End If
'				Browser("title:=Siebel Communications.*").Page("title:=Siebel Communications.*").WebEdit("html id:=s_swepi_1").Set strUserName
'				Browser("title:=Siebel Communications.*").Page("title:=Siebel Communications.*").WebEdit("html id:=s_swepi_2").Set strPass
'				Browser("title:=Siebel Communications.*").Page("title:=Siebel Communications.*").WebEdit("html id:=s_swepi_2").highlight
'				Browser("title:=Siebel Communications.*").Page("title:=Siebel Communications.*").WebEdit("html id:=s_swepi_2").Click
			wait 3
            Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("SWEUserName").Set "ROBOTDEV"'"CAPACITACION" @@ script infofile_;_ZIP::ssf21.xml_;_
            Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("SWEPassword").SetSecure "Password_#2022_QAS"'"6264cf4feef1289779e0af080336bcef2eee24f10aa50d68925f3c4af788d42e63f2eec40830324c909a"  'Password_#2022_QAS @@ script infofile_;_ZIP::ssf22.xml_;_
            'Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("SWEUserName").Set "ROBOTDEV"'"CAPACITACION"
            'Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("SWEPassword").Set "Password_#2022_QAS"'Secure "6264cf4feef1289779e0af080336bcef2eee24f10aa50d68925f3c4af788d42e63f2eec40830324c909a"  'Password_#2022_QAS
            Browser("Todas las cuentas").Page("Todas las cuentas").Link("Inicio de sesión").Click @@ script infofile_;_ZIP::ssf23.xml_;_
            wait 30
		
End Function

'wait 3
'Browser("Siebel Communications_3").Page("Siebel Communications").WebEdit("SWEUserName").Set "CAPACITACION" @@ script infofile_;_ZIP::ssf14.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebEdit("SWEPassword").SetSecure "6264b9c052948759a8c7750fe37065b1dd0f33018067a170a8dc4db999a2e0ad5a866a442e821060712b" @@ script infofile_;_ZIP::ssf15.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").Link("Inicio de sesión").Click @@ script infofile_;_ZIP::ssf16.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebTabStrip("Página inicialPantalla").Select "Cuentas" @@ script infofile_;_ZIP::ssf17.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebList("s_vis_div").Select "Captura de Cuentas" @@ script infofile_;_ZIP::ssf18.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebButton("Cuentas:Nuevo").Click @@ script infofile_;_ZIP::ssf19.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebEdit("Primer Nombre").Set "JOSE" @@ script infofile_;_ZIP::ssf20.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebEdit("Segundo Nombre").Set "luis" @@ script infofile_;_ZIP::ssf21.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebEdit("Apellido Paterno").Set "Santiago" @@ script infofile_;_ZIP::ssf22.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebEdit("Apellido Materno").Set "williams" @@ script infofile_;_ZIP::ssf23.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebElement("Sr./Sra.:").Click @@ script infofile_;_ZIP::ssf24.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebEdit("Licencia Conducir").Set "345234234214" @@ script infofile_;_ZIP::ssf25.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebEdit("Credencial Elector").Set "23452345324" @@ script infofile_;_ZIP::ssf26.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebElement("Seleccionar varios campos").Click @@ script infofile_;_ZIP::ssf27.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebButton("Teléfonos:Nuevo").Click @@ script infofile_;_ZIP::ssf28.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebEdit("Numero Telefonico").Set "5512963466" @@ script infofile_;_ZIP::ssf29.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebElement("s_9_2_165_0_icon").Click @@ script infofile_;_ZIP::ssf30.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebList("ui-id-302").Click @@ script infofile_;_ZIP::ssf31.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebEdit("Tipo de Telefono").Set "Celular" @@ script infofile_;_ZIP::ssf32.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebElement("s_9_2_166_0_icon").Click @@ script infofile_;_ZIP::ssf33.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebList("ui-id-303").Click @@ script infofile_;_ZIP::ssf34.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebEdit("Compañia Telefonica").Set "Telcel" @@ script infofile_;_ZIP::ssf35.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebButton("Teléfonos:Guardar").Click @@ script infofile_;_ZIP::ssf36.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebButton("Teléfonos:Aceptar").Click @@ script infofile_;_ZIP::ssf37.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebElement("Seleccionar varios campos_2").Click @@ script infofile_;_ZIP::ssf38.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebButton("Cuenta de Email:Nuevo").Click @@ script infofile_;_ZIP::ssf39.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebEdit("Correo").Set "mas1" @@ script infofile_;_ZIP::ssf40.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebElement("s_10_2_187_0_icon").Click @@ script infofile_;_ZIP::ssf41.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebList("ui-id-302").Click @@ script infofile_;_ZIP::ssf42.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebEdit("Dominio").Set "gmail.com" @@ script infofile_;_ZIP::ssf43.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebEdit("Confirma Correo").Set "mas1" @@ script infofile_;_ZIP::ssf44.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebElement("s_10_2_189_0_icon").Click @@ script infofile_;_ZIP::ssf45.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebList("ui-id-303").Click @@ script infofile_;_ZIP::ssf46.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebEdit("Confirma Dominio").Set "gmail.com" @@ script infofile_;_ZIP::ssf47.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebCheckBox("Promociones").Set "ON" @@ script infofile_;_ZIP::ssf48.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebCheckBox("Newsletter").Set "ON" @@ script infofile_;_ZIP::ssf49.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebButton("Cuenta de Email:Guardar").Click @@ script infofile_;_ZIP::ssf50.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebButton("Cuenta de Email:Validar").Click @@ script infofile_;_ZIP::ssf51.xml_;_
'Browser("Siebel Communications_3").HandleDialog micOK @@ hightlight id_;_5573814_;_script infofile_;_ZIP::ssf52.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebButton("Cuenta de Email:Validar").Click @@ script infofile_;_ZIP::ssf53.xml_;_
'Browser("Siebel Communications_3").HandleDialog micOK @@ hightlight id_;_5573814_;_script infofile_;_ZIP::ssf54.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebElement("1_s_10_l_CV_Correo").Click @@ script infofile_;_ZIP::ssf55.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebButton("Cuenta de Email:Aceptar").Click @@ script infofile_;_ZIP::ssf56.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebElement("Centro de costos:").Click @@ script infofile_;_ZIP::ssf57.xml_;_
'Browser("Siebel Communications_3").HandleDialog micOK @@ hightlight id_;_5573814_;_script infofile_;_ZIP::ssf58.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebElement("s_1_1_14_0_icon").Click @@ script infofile_;_ZIP::ssf59.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebList("ui-id-302").Click @@ script infofile_;_ZIP::ssf60.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebEdit("Pago Anticipado Vendedor").Set "NO" @@ script infofile_;_ZIP::ssf61.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebElement("Volumen potencial:").Click @@ script infofile_;_ZIP::ssf62.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebButton("Direcciones de la Cuenta:Nuevo").Click @@ script infofile_;_ZIP::ssf63.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebElement("Seleccionar campo").Click @@ script infofile_;_ZIP::ssf64.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebElement("s_11_1_115_0_icon").Click @@ script infofile_;_ZIP::ssf65.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebList("ui-id-302").Click @@ script infofile_;_ZIP::ssf66.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebEdit("s_11_1_115_0").Set "Codigo Postal" @@ script infofile_;_ZIP::ssf67.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebEdit("s_11_1_116_0").Set "03100" @@ script infofile_;_ZIP::ssf68.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebButton("Nombres Calles:Ir").Click @@ script infofile_;_ZIP::ssf69.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebButton("Nombres Calles:Aceptar").Click @@ script infofile_;_ZIP::ssf70.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebEdit("Calle").Set "CAPULIN" @@ script infofile_;_ZIP::ssf71.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebEdit("Calle").Set "CAPULIN" @@ script infofile_;_ZIP::ssf72.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebButton("Perfil de facturación:Nuevo").Click @@ script infofile_;_ZIP::ssf73.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebButton("Efectivo:Nuevo").Click @@ script infofile_;_ZIP::ssf74.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebButton("Vista de la lista de contratos").Click @@ script infofile_;_ZIP::ssf75.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebElement("s_4_2_20_0_icon").Click @@ script infofile_;_ZIP::ssf76.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebList("ui-id-302").Click @@ script infofile_;_ZIP::ssf77.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebEdit("1-124893392622").Set "Oro" @@ script infofile_;_ZIP::ssf78.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebElement("WebElement").Click @@ script infofile_;_ZIP::ssf79.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebElement("s_4_2_20_0_icon").Click @@ script infofile_;_ZIP::ssf80.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebList("ui-id-302").Click @@ script infofile_;_ZIP::ssf81.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebEdit("Tipo").Set "Mensual" @@ script infofile_;_ZIP::ssf82.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebElement("WebElement").Click @@ script infofile_;_ZIP::ssf83.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").Link("33534302").Click @@ script infofile_;_ZIP::ssf84.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").WebButton("Ítems de facturación:Nuevo").Click @@ script infofile_;_ZIP::ssf85.xml_;_
'Browser("Siebel Communications_3").Page("Siebel Communications").Link("1-124893392622").Click @@ script infofile_;_ZIP::ssf86.xml_;_
'Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("WebEdit").Set
'Browser("Todas las cuentas").Page("Todas las cuentas").WebTable("Numero Cuenta:").GetCellData
' @@ script infofile_;_ZIP::ssf14.xml_;_


'Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("Primer Nombre").Set "jose" @@ script infofile_;_ZIP::ssf15.xml_;_
'Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("Nombre Comercial:").Click @@ script infofile_;_ZIP::ssf16.xml_;_
'Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("Segundo Nombre").Set "luis" @@ script infofile_;_ZIP::ssf17.xml_;_
'Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("Apellido Paterno").Set "santiago" @@ script infofile_;_ZIP::ssf18.xml_;_
'Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("Apellido Materno").Set "williams" @@ script infofile_;_ZIP::ssf19.xml_;_
'Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("WebElement").Click @@ script infofile_;_ZIP::ssf20.xml_;_

If false Then
	
'------------------------NUEVO CODIGO
wait 3
Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("SWEUserName").Set "CAPACITACION" @@ script infofile_;_ZIP::ssf21.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("SWEPassword").SetSecure "6264cf4feef1289779e0af080336bcef2eee24f10aa50d68925f3c4af788d42e63f2eec40830324c909a" @@ script infofile_;_ZIP::ssf22.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").Link("Inicio de sesión").Click @@ script infofile_;_ZIP::ssf23.xml_;_
wait 3
Browser("Todas las cuentas").Page("Todas las cuentas").WebTabStrip("Página inicialPantalla").Select "Cuentas" @@ script infofile_;_ZIP::ssf24.xml_;_
wait 3
Browser("Todas las cuentas").Page("Todas las cuentas").WebList("s_vis_div").Select "Captura de Cuentas" @@ script infofile_;_ZIP::ssf25.xml_;_
wait 6
Browser("Todas las cuentas").Page("Todas las cuentas").WebButton("Cuentas:Nuevo").Click @@ script infofile_;_ZIP::ssf26.xml_;_
wait 5
Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("Primer Nombre").Set "jose" @@ script infofile_;_ZIP::ssf27.xml_;_

Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("Segundo Nombre").Set "luis" @@ script infofile_;_ZIP::ssf29.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("Apellido Paterno").Set "santiago" @@ script infofile_;_ZIP::ssf30.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("Apellido Materno").Set "williams" @@ script infofile_;_ZIP::ssf31.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("Licencia Conducir").Set "4563453245324" @@ script infofile_;_ZIP::ssf32.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("Credencial Elector").Set "234532453245324" @@ script infofile_;_ZIP::ssf33.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("s_1_1_14_0_icon").Click @@ script infofile_;_ZIP::ssf34.xml_;_
'Browser("Todas las cuentas").Page("Todas las cuentas").WebList("ui-id-586").Select "NO" @@ script infofile_;_ZIP::ssf35.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("Pago Anticipado Vendedor").Set "NO" @@ script infofile_;_ZIP::ssf36.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("Seleccionar varios campos").Click @@ script infofile_;_ZIP::ssf37.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebButton("Teléfonos:Nuevo").Click @@ script infofile_;_ZIP::ssf38.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("Numero Telefonico").Set "5512963466" @@ script infofile_;_ZIP::ssf39.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("s_9_2_165_0_icon").Click @@ script infofile_;_ZIP::ssf40.xml_;_
'Browser("Todas las cuentas").Page("Todas las cuentas").WebList("ui-id-675").Click @@ script infofile_;_ZIP::ssf41.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("Tipo de Telefono").Set "Celular" @@ script infofile_;_ZIP::ssf42.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("s_9_2_166_0_icon").Click @@ script infofile_;_ZIP::ssf43.xml_;_
'Browser("Todas las cuentas").Page("Todas las cuentas").WebList("ui-id-676").Click @@ script infofile_;_ZIP::ssf44.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("Compañia Telefonica").Set "Telcel" @@ script infofile_;_ZIP::ssf45.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebButton("Teléfonos:Guardar").Click @@ script infofile_;_ZIP::ssf46.xml_;_
wait 15
Browser("Todas las cuentas").Page("Todas las cuentas").WebButton("Teléfonos:Aceptar").Click @@ script infofile_;_ZIP::ssf47.xml_;_
wait 5
Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("Seleccionar varios campos_2").Click @@ script infofile_;_ZIP::ssf48.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebButton("Cuenta de Email:Nuevo").Click @@ script infofile_;_ZIP::ssf49.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("Correo").Set "aura" @@ script infofile_;_ZIP::ssf50.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("s_10_2_187_0_icon").Click @@ script infofile_;_ZIP::ssf51.xml_;_
'Browser("Todas las cuentas").Page("Todas las cuentas").WebList("ui-id-675").Click @@ script infofile_;_ZIP::ssf52.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("Dominio").Set "gmail.com" @@ script infofile_;_ZIP::ssf53.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("Confirma Correo").Set "aura" @@ script infofile_;_ZIP::ssf54.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("s_10_2_189_0_icon").Click @@ script infofile_;_ZIP::ssf55.xml_;_
'Browser("Todas las cuentas").Page("Todas las cuentas").WebList("ui-id-676").Click @@ script infofile_;_ZIP::ssf56.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("Confirma Dominio").Set "gmail.com" @@ script infofile_;_ZIP::ssf57.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebCheckBox("Promociones").Set "ON" @@ script infofile_;_ZIP::ssf58.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebCheckBox("Newsletter").Set "ON" @@ script infofile_;_ZIP::ssf59.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebCheckBox("Paperless").Set "ON" @@ script infofile_;_ZIP::ssf60.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebButton("Cuenta de Email:Guardar").Click @@ script infofile_;_ZIP::ssf61.xml_;_
wait 3
Browser("Todas las cuentas").Page("Todas las cuentas").WebButton("Cuenta de Email:Validar").Click @@ script infofile_;_ZIP::ssf62.xml_;_
'Browser("Todas las cuentas").HandleDialog micOK @@ hightlight id_;_24775614_;_script infofile_;_ZIP::ssf63.xml_;_
wait 2
mensaje = Browser("Todas las cuentas").GetDialogText
wait 2
If trim(mensaje) = "Error, la direccion de correo no existe(SBL-EXL-00151)" Then
	Set oSave = Createobject("Wscript.Shell")
oSave.SendKeys "{ENTER}"
Set oSave = Nothing
End If
wait 2
print mensaje
Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("Exención validación").Click @@ script infofile_;_ZIP::ssf64.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebButton("Cuenta de Email:Aceptar").Click @@ script infofile_;_ZIP::ssf65.xml_;_
wait 5
Set oSave = Createobject("Wscript.Shell")
oSave.SendKeys "^s"
Set oSave = Nothing
wait 5
numerocuenta = Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("acc_name:=Numero Cuenta").GetROProperty("value")
print numerocuenta


wait 3
Browser("Todas las cuentas").Page("Todas las cuentas").WebButton("Direcciones de la Cuenta:Nuevo").Click @@ script infofile_;_ZIP::ssf66.xml_;_
wait 3
Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("Seleccionar campo").Click @@ script infofile_;_ZIP::ssf67.xml_;_
wait 3
Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("s_11_1_115_0_icon").Click @@ script infofile_;_ZIP::ssf68.xml_;_
wait 3
'Browser("Todas las cuentas").Page("Todas las cuentas").WebList("ui-id-586").Click @@ script infofile_;_ZIP::ssf69.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("s_11_1_115_0").Set "Codigo Postal" @@ script infofile_;_ZIP::ssf70.xml_;_
wait 3
Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("s_11_1_116_0").Set "03100" @@ script infofile_;_ZIP::ssf71.xml_;_
wait 3
Browser("Todas las cuentas").Page("Todas las cuentas").WebButton("Nombres Calles:Ir").Click @@ script infofile_;_ZIP::ssf72.xml_;_
wait 3
Browser("Todas las cuentas").Page("Todas las cuentas").WebButton("Nombres Calles:Aceptar").Click @@ script infofile_;_ZIP::ssf73.xml_;_
wait 3
Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("Calle").Set "CAPULIN" @@ script infofile_;_ZIP::ssf74.xml_;_
wait 3
Set oSave = Createobject("Wscript.Shell")
oSave.SendKeys "^s"
Set oSave = Nothing
wait 10
Set oSave = Createobject("Wscript.Shell")
oSave.SendKeys "^b"
Set oSave = Nothing
wait 10
'Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("Calle").Set "CAPULIN" @@ script infofile_;_ZIP::ssf75.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebButton("Perfil de facturación:Nuevo").Click @@ script infofile_;_ZIP::ssf76.xml_;_
wait 3
Browser("Todas las cuentas").Page("Todas las cuentas").WebButton("Efectivo:Nuevo").Click @@ script infofile_;_ZIP::ssf77.xml_;_
wait 3
Browser("Todas las cuentas").Page("Todas las cuentas").WebButton("Vista de la lista de contratos").Click @@ script infofile_;_ZIP::ssf78.xml_;_
wait 3
Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("s_4_2_20_0_icon").Click @@ script infofile_;_ZIP::ssf79.xml_;_
'Browser("Todas las cuentas").Page("Todas las cuentas").WebList("ui-id-586").Click @@ script infofile_;_ZIP::ssf80.xml_;_
wait 3
Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("html id:=1_Service_Type").Set "Oro" @@ script infofile_;_ZIP::ssf81.xml_;_
wait 3
Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("WebElement_3").Click @@ script infofile_;_ZIP::ssf82.xml_;_
wait 3
Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("s_4_2_20_0_icon").Click @@ script infofile_;_ZIP::ssf83.xml_;_
'Browser("Todas las cuentas").Page("Todas las cuentas").WebList("ui-id-586").Click @@ script infofile_;_ZIP::ssf84.xml_;_
wait 3
Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("Tipo").Set "Mensual" @@ script infofile_;_ZIP::ssf85.xml_;_
wait 3
Set oSave = Createobject("Wscript.Shell")
oSave.SendKeys "^s"
Set oSave = Nothing
wait 10
wait 3
contrato = Browser("Todas las cuentas").Page("Todas las cuentas").Link("html tag:=A","abs_x:=205","color:=rgb\(20, 116, 191\)").GetROProperty("name")
print contrato
Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("jqgh_s_4_l_Name").Click @@ script infofile_;_ZIP::ssf86.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("jqgh_s_3_l_Name").Click @@ script infofile_;_ZIP::ssf87.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("1_s_3_l_CV_Calculated_Alias_Na").Click @@ script infofile_;_ZIP::ssf88.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("WebElement_4").Click @@ script infofile_;_ZIP::ssf89.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("s_4_2_20_0_icon").Click
wait 3
'Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("WebElement_4").Click @@ script infofile_;_ZIP::ssf89.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").Link("html tag:=A","abs_x:=214","color:=rgb\(20, 116, 191\)").Click
'Browser("Todas las cuentas").Page("Todas las cuentas").Link("html tag:=A","color:=rgb\(20, 116, 191\)").Click @@ script infofile_;_ZIP::ssf28.xml_;_
 '-----------
'---------------------------FIN NUEVO CODIGO

'Browser("Siebel Communications_3").Page("Siebel Communications").WebEdit("SWEUserName").Set

Browser("Todas las cuentas").Page("Todas las cuentas").WebButton("Detalles:Expandir Ítems").Click @@ script infofile_;_ZIP::ssf93.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("s_9_2_166_0_icon").Click @@ script infofile_;_ZIP::ssf94.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebList("ui-id-586").Click @@ script infofile_;_ZIP::ssf95.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("Tipo de orden:").Click @@ script infofile_;_ZIP::ssf96.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("s_1_1_88_0_icon").Click @@ script infofile_;_ZIP::ssf97.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("Horario de Atención:").Click @@ script infofile_;_ZIP::ssf98.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebButton("Detalles:Agregar Items").Click @@ script infofile_;_ZIP::ssf99.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("s_10_2_189_0_icon").Click @@ script infofile_;_ZIP::ssf100.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebList("ui-id-676").Click @@ script infofile_;_ZIP::ssf101.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("s_10_2_189_0_icon").Click @@ script infofile_;_ZIP::ssf102.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebList("ui-id-676").Click @@ script infofile_;_ZIP::ssf103.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("Buscar").Set "Producto" @@ script infofile_;_ZIP::ssf104.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("Que empiece por").Set "Port*" @@ script infofile_;_ZIP::ssf105.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebButton("Seleccionar producto:Ir").Click @@ script infofile_;_ZIP::ssf106.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("gbox_s_11_l").Click @@ script infofile_;_ZIP::ssf107.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("Que empiece por").Set "Portafolio Productos wow" @@ script infofile_;_ZIP::ssf108.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebButton("Seleccionar producto:Ir").Click @@ script infofile_;_ZIP::ssf109.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebButton("Seleccionar producto:Agregar").Click @@ script infofile_;_ZIP::ssf110.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebButton("Seleccionar producto:Agregar_2").Click @@ script infofile_;_ZIP::ssf111.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("gbox_s_4_l").Click @@ script infofile_;_ZIP::ssf112.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("WebElement_8").Click @@ script infofile_;_ZIP::ssf113.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("WebElement_8").Click @@ script infofile_;_ZIP::ssf114.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebButton("Detalles:Desagrupar").Click @@ script infofile_;_ZIP::ssf115.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebButton("Detalles Menú List").Click @@ script infofile_;_ZIP::ssf116.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebMenu("s_at_m_4-menu").Select "Eliminar registro [Ctrl+D]" @@ script infofile_;_ZIP::ssf117.xml_;_
Browser("Todas las cuentas").HandleDialog micOK @@ hightlight id_;_17434780_;_script infofile_;_ZIP::ssf118.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebButton("Detalles:Personalizar").Click @@ script infofile_;_ZIP::ssf119.xml_;_
wait 3
Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("WebElement_9").highlight
Set objMultiplay_Table = Browser("Siebel Communications_Numero").Page("Siebel Communications").Frame("CfgMainFrame Frame").WebTable("Multiplay_Table")
			print "                      "&sp_var(i)
			fnSelectWebTableValue objMultiplay_Table, sp_var(i)

wait 3

Browser("Todas las cuentas").Page("Todas las cuentas").Link("Internet").Click
wait 2
Set desc2 =Description.Create
    desc2("micclass").value="WebCheckBox"
    'desc2("column names").value=";Ítem;Precio de lista;Descripcion"

Set childItems =   Browser("Todas las cuentas").Page("Todas las cuentas").ChildObjects(desc2)
  print childItems.count
  
   'Set childTable= childItems(0).ChildObjects(desc2)
   'print " "&childTable.count
   
   
   For i = 0 To childItems.count-1
        valor = childItems(i).GetROProperty("value")
        If valor = "Hazlo 30 NxtGen" Then '"izzi_20_" Then
        	childItems(i).highlight()
        	childItems(i).set "ON"
   	        'inner = childItems(i).GetROProperty("innertext")
   	        print valor
   	        Exit for
        End If
   	    'childItems(i).highlight()
   	    'inner = childItems(i).GetROProperty("innertext")
   	    'print inner
   Next			


'Browser("Todas las cuentas").Page("Todas las cuentas").Link("1-124895732632").Click @@ script infofile_;_ZIP::ssf120.xml_;_
'Browser("Todas las cuentas").Page("Todas las cuentas").Link("Internet").Click @@ script infofile_;_ZIP::ssf121.xml_;_
'Browser("Todas las cuentas").Page("Todas las cuentas").Link("Telefonia").Click @@ script infofile_;_ZIP::ssf122.xml_;_
'Browser("Todas las cuentas").Page("Todas las cuentas").Link("OTT").Click @@ script infofile_;_ZIP::ssf123.xml_;_
'Browser("Todas las cuentas").Page("Todas las cuentas").Link("FTTH").Click @@ script infofile_;_ZIP::ssf124.xml_;_
'Browser("Todas las cuentas").Page("Todas las cuentas").Link("Otros Servicios").Click @@ script infofile_;_ZIP::ssf125.xml_;_
'Browser("Todas las cuentas").Page("Todas las cuentas").Link("Movil").Click @@ script infofile_;_ZIP::ssf126.xml_;_
'Browser("Todas las cuentas").Page("Todas las cuentas").Link("Internet").Click @@ script infofile_;_ZIP::ssf127.xml_;_

'Browser("Todas las cuentas").Page("Todas las cuentas").WebButton("Terminado").Click @@ script infofile_;_ZIP::ssf128.xml_;_
'Browser("Todas las cuentas").Page("Todas las cuentas").WebButton("Detalles:Expandir Ítems").Click @@ script infofile_;_ZIP::ssf129.xml_;_
'Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("WebElement_10").Click @@ script infofile_;_ZIP::ssf130.xml_;_
'Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("WebElement_11").Click @@ script infofile_;_ZIP::ssf131.xml_;_
'Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("s_4_2_78_0_icon").Click @@ script infofile_;_ZIP::ssf132.xml_;_
'Browser("Todas las cuentas").Page("Todas las cuentas").WebButton("Close").Click @@ script infofile_;_ZIP::ssf133.xml_;_
'Browser("Todas las cuentas").HandleDialog micOK @@ hightlight id_;_17434780_;_script infofile_;_ZIP::ssf134.xml_;_

wait 2
Set desc2 =Description.Create
    desc2("micclass").value="WebTable"
    'desc2("column names").value=";Ítem;Precio de lista;Descripcion"

Set childItems =   Browser("Todas las cuentas").Page("Todas las cuentas").ChildObjects(desc2)
  print childItems.count
  
   'Set childTable= childItems(0).ChildObjects(desc2)
   'print " "&childTable.count
   
   
   For i = 0 To childItems.count-1
        valor = childItems(i).GetROProperty("name")
        If valor = "s_4_lSelectAll" Then '"izzi_20_" Then
        	childItems(i).highlight()
'        	childItems(i).set "ON"
'   	        'inner = childItems(i).GetROProperty("innertext")
   	        print valor
   	        Exit for
        End If
   	    'childItems(i).highlight()
   	    'inner = childItems(i).GetROProperty("innertext")
   	    'print inner
   Next		

wait 4

set table = Browser("Todas las cuentas").Page("Todas las cuentas").WebTable("name:=s_4_lSelectAll")
fila = table.GetRowWithCellText("Hazlo 30 NxtGen")
print fila
dato = table.GetCellData(fila,4)
print dato
'Browser("Todas las cuentas").Page("Todas las cuentas").WebTable("name:=s_4_lSelectAll").SetTOProperty
'ctchildproduct = SiebApplication("Siebel Communications").SiebScreen("Ordenes de Servicio").SiebView("Detalles").SiebApplet("Detalles").SiebList("List").GetCellText("Product",q)
'Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("WebElement_12").Click
'Browser("Todas las cuentas").Page("Todas las cuentas").WebTable("1").GetCellData
'Lineitem= SiebApplication("Siebel Communications").SiebScreen("Ordenes de Servicio").SiebView("Detalles").SiebApplet("Detalles").SiebList("List").SiebText("No. de Serie Equipo").GetROProperty("isenabled")
'Browser("Siebel Communications_3").Page("Siebel Communications").WebEdit("SWEUserName").Set

Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("No. de Serie Equipo").Click @@ script infofile_;_ZIP::ssf135.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("Hazlo 30 NxtGen").Set "qa"',True @@ script infofile_;_ZIP::ssf136.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("Motivo de Cancelacion").Set
Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("WebEdit_3").Set
wait 3
Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("acc_name:=Bloqueos izzi","html id:=6_s_4_l_CV_TN").Click
Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("acc_name:=Bloqueos izzi","name:=CV_TN").Set "1342"
Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("s_4_2_78_0_icon").Click

Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("s_4_2_78_0_icon").Click @@ script infofile_;_ZIP::ssf137.xml_;_
Browser("Todas las cuentas").Page("Todas las cuentas").WebButton("Close").Click @@ script infofile_;_ZIP::ssf138.xml_;_
Browser("Todas las cuentas").HandleDialog micOK @@ hightlight id_;_23397228_;_script infofile_;_ZIP::ssf139.xml_;_

wait 3
'Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("acc_name:=Bloqueos izzi","html id:=.*_s_4_l_Outline_Number").Click
'Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("acc_name:=Bloqueos izzi","html id:=6_s_4_l_CV_TN").Click
outlinenumber = "html id:="&"1"&"_s_4_l_Outline_Number"
Browser("Todas las cuentas").Page("Todas las cuentas").WebElement(outlinenumber).Click
wait 3
outlinenumber = "html id:="&"6"&"_s_4_l_Outline_Number"
Browser("Todas las cuentas").Page("Todas las cuentas").WebElement(outlinenumber).Click
'funcion agregarproductos
wait 3
Browser("Todas las cuentas").Page("Todas las cuentas").WebButton("title:=Ítems de facturación:Nuevo").highlight
Browser("Todas las cuentas").Page("Todas las cuentas").WebButton("title:=Ítems de facturación:Nuevo").Click
wait 10
sOrderService = Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("acc_name:=Número").GetROProperty("value")
Datatable.Value("o_OrderService","Output") = sOrderService
print "sOrderService "&sOrderService
wait 3
Set sibOrderMot=Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("acc_name:=Motivo de la orden")
    'sibOrderMot.highlight
    strOrdenMotivo=sibOrderMot.GetROProperty("value")
 If strOrdenMotivo="" and (DataTable.value("p_MotiveOrder","ProductSelection")="AXTEL" or DataTable.value("p_MotiveOrder","ProductSelection")="FTTH") Then
 	sibOrderMot.Set  DataTable.value("p_MotiveOrder","ProductSelection") '"Triple Play Video Internet Tel"
 	'Reporter.ReportEvent mic,"Application Issue","Motivo de la Orden: No se llena en automatico,se tuvo que asignar manualemente"
	strMsg="Motivo de la Orden: No se llena en automatico,se tuvo que asignar manualemente"
	Call fnUtilDTWriterMsg("Output","o_error",strMsg)  
 End If
 wait 2
 Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("acc_name:=CIC Potencial").Set Datatable.Value("p_CIC_Potencial","ProductSelection")
 wait 3
Browser("Todas las cuentas").Page("Todas las cuentas").WebButton("Detalles:Agregar Items").Click @@ script infofile_;_ZIP::ssf99.xml_;_
wait 2
       If Datatable.Value("p_Portafolio","ProductSelection") = "Paquete Cablevision" Then
		    Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("Buscar").Set "Nº de pieza"
            Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("Que empiece por").Set "1-3T4QKX"
		else
		    Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("Que empiece por").Set Datatable.Value("p_Portafolio","ProductSelection")

		End If

Browser("Todas las cuentas").Page("Todas las cuentas").WebButton("Seleccionar producto:Ir").Click @@ script infofile_;_ZIP::ssf109.xml_;_
wait 2
Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("jqgh_s_13_l_Name").Click @@ script infofile_;_ZIP::ssf140.xml_;_
wait 1
Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("class:=siebui-icon-arrowsm-down","role:=presentation","visible:=True").Click
wait 2
Browser("Todas las cuentas").Page("Todas las cuentas").WebButton("acc_name:=Seleccionar producto:Agregar y cerrar").Click
wait 10
		If NOT DataTable.Value ("p_Execute", "Admision") = uCase("CP") and NOT DataTable.Value ("p_Execute", "Admision") = uCase("CPP")  Then 'UNICA  JIEC	crea hasta portafolio						
						
				'Check whether the product added successful by checking the property of Personlizer button
				If Browser("Todas las cuentas").Page("Todas las cuentas").WebButton("Detalles:Personalizar").Exist(3) Then
					reporter.ReportEvent micPass,"Product Portfolio","Product Added successfully"
				Else
					Reporter.ReportEvent micFail,"Product Portfolio","Failed to add the Product",ReporterCapture()
					ExitActioniteration
				End If
				Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("Estado").GetROProperty("value")

				finalstatus = Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("Estado").GetROProperty("value")
                print "Estado de la orden:"&finalstatus
				Datatable.Value("o_Status","Output") = finalstatus
				'CLick on Personalier Button
				'SiebApplication("Siebel Communications").SiebScreen("Ordenes de Servicio").SiebView("Detalles").SiebApplet("Detalles").SiebButton("Personalizar").Click
				Browser("Todas las cuentas").Page("Todas las cuentas").WebButton("Detalles:Personalizar").Click
                wait 5
		End If
wait 2
Browser("Todas las cuentas").Page("Todas las cuentas").WebList("s_vis_div").Select "Extension de Video B wow"
wait 2
Browser("Todas las cuentas").Page("Todas las cuentas").WebButton("class:=siebui-ctrl-btn ","html tag:=BUTTON").Click
wait 2
Browser("Todas las cuentas").Page("Todas las cuentas").WebTable("column names:=;;Extension Video.*").WebCheckBox("value:=Extension HD","checked:=0").Set "ON"
'reviso consecutivo y reviso nombre del producto
'al revisar si el consecutivo mide mas de 3 y coincide con internet tel o video, le pongo el equipo correspondiente
'reviso si el producto se llama "internet" o "telefonia" o "video"
'en caso de ser internet asocio al consecutivo que es equipo internet, igual para telefonia y video


Browser("Todas las cuentas").Page("Todas las cuentas").WebTable("OrdenServicioDetalles").GetCellData

Browser("Todas las cuentas").Page("Todas las cuentas").WebTable("OrdenServicioDetalles").GetCellData

wait 3
filastotal = Browser("Todas las cuentas").Page("Todas las cuentas").WebTable("OrdenServicioDetalles").RowCount
papaVideo = ""
papaInternet = ""
papaTelefonia = ""
For I = 1 To filastotal-1
    idhtml = "html id:="&I&"_s_4_l_Outline_Number"
	consecutivo = Browser("Todas las cuentas").Page("Todas las cuentas").WebElement(idhtml).GetROProperty("title")
	print consecutivo
	If Len(consecutivo) = 3 Then
		print "producto principal"
		idactual =  "html id:="&I&"_Outline_Number"
		Browser("Todas las cuentas").Page("Todas las cuentas").WebElement(idhtml).Click
		Producto = Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit(idactual).GetROProperty("acc_name")
        If InStr(Producto,"Video") > 0 Then
        	papaVideo = consecutivo
        End If
        If InStr(Producto,"Internet") > 0 Then
        	papaInternet = consecutivo
        End If
        If InStr(Producto,"Telefonia") > 0 Then
        	papaTelefonia = consecutivo
        End If
	End If
	If Len(consecutivo) > 3 Then
	    equipo = "html id:="&I&"_s_4_l_Service_Id"
	    equipoInside = "html id:="&I&"_Service_Id"
	    Browser("Todas las cuentas").Page("Todas las cuentas").WebElement(idhtml).Click
	    Browser("Todas las cuentas").Page("Todas las cuentas").WebElement(equipo).Click
	    enable = Browser("Todas las cuentas").Page("Todas las cuentas").WebElement(equipoInside).GetROProperty("readonly")
	    print "enable:"&enable
	    If enable = "0" Then
	    	hijo = Left(consecutivo,3)
		    If hijo = papaVideo Then
			   print "le pego equipo de Video"
		    End If
		    If hijo = papaInternet Then
			   print "le pego equipo de Internet"
		    End If
		    If hijo = papaTelefonia Then
			   print "le pego equipo de Telefonia"
		    End If
	    End If		
	End If
Next

Print papaVideo
Print papaInternet
print papaTelefonia

wait 3
filastotal = Browser("Todas las cuentas").Page("Todas las cuentas").WebTable("OrdenServicioDetalles").RowCount
'Browser("Todas las cuentas").Page("Todas las cuentas").WebTable("OrdenServicioDetalles").ChildItem(1,2,"WebEdit",0).click
papaTelefonia = ""
ITelefonia = ""
telefono1 = ""
telefono2 = ""
varB = 0
varC = 0
varS = 0
For I = 1 To filastotal-1
    idhtml = "html id:="&I&"_s_4_l_Outline_Number"
	consecutivo = Browser("Todas las cuentas").Page("Todas las cuentas").WebElement(idhtml).GetROProperty("title")
	print consecutivo
	If Len(consecutivo) = 3 Then
		print "producto principal"
		idactual =  "html id:="&I&"_Outline_Number"
		Browser("Todas las cuentas").Page("Todas las cuentas").WebElement(idhtml).Click
		Producto = Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit(idactual).GetROProperty("acc_name")
        If InStr(Producto,"Telefonia") > 0 Then
        	papaTelefonia = consecutivo
        	ITelefonia = I
        End If
	End If
	If Len(consecutivo) > 3 Then
	    equipo = "html id:="&I&"_s_4_l_CV_TN"
	    equipoInside = "html id:="&I&"_CV_TN"
	    Npro = "html id:="&I&"_s_4_l_Product"
	    NproInside = "html id:="&I&"_Product"
	    Browser("Todas las cuentas").Page("Todas las cuentas").WebElement(idhtml).Click
	    Browser("Todas las cuentas").Page("Todas las cuentas").WebElement(Npro).Click
	    Producto = Browser("Todas las cuentas").Page("Todas las cuentas").WebElement(NproInside).GetROProperty("value")
	    Browser("Todas las cuentas").Page("Todas las cuentas").WebElement(equipo).Click
	    enable = Browser("Todas las cuentas").Page("Todas las cuentas").WebElement(equipoInside).GetROProperty("readonly")
	    print "enable:"&enable
	    If Producto = "izzi Telefonia" and enable = "0" Then
	    	print "telefono principal"
	    	print "pongo telefono1"
	    	telefono1 = ""
	    	Exit for
	    End If
	End If
Next
For I = ITelefonia+1 To filastotal-1
	idhtml = "html id:="&I&"_s_4_l_Outline_Number"
	consecutivo = Browser("Todas las cuentas").Page("Todas las cuentas").WebElement(idhtml).GetROProperty("title")
	print consecutivo
	If Len(consecutivo) = 3 Then
		print "producto principal"
		Exit for
	End If
	If Len(consecutivo) > 3 Then
	    equipo = "html id:="&I&"_s_4_l_CV_TN"
	    equipoInside = "html id:="&I&"_CV_TN"
	    Npro = "html id:="&I&"_s_4_l_Product"
	    NproInside = "html id:="&I&"_Product"
	    Browser("Todas las cuentas").Page("Todas las cuentas").WebElement(idhtml).Click
	    Browser("Todas las cuentas").Page("Todas las cuentas").WebElement(Npro).Click
	    Producto = Browser("Todas las cuentas").Page("Todas las cuentas").WebElement(NproInside).GetROProperty("value")
	    Browser("Todas las cuentas").Page("Todas las cuentas").WebElement(equipo).Click
	    enable = Browser("Todas las cuentas").Page("Todas las cuentas").WebElement(equipoInside).GetROProperty("readonly")
	    print "enable:"&enable
	    If Producto = "Bloqueos izzi" and enable = "0" Then
	    	   print "Bloqueos"
	    	If varB = 0 Then
	    	   print "pongo telefono1"
               varB = 1	
	    	else
	    	   print "pongo telefono2"
	    	End If
	    End If
	    If Producto = "Complementos telefonia" and enable = "0" Then
	    	   print "Complementos telefonia"
	    	If varC = 0 Then
	    	   print "pongo telefono1"
               varC = 1	
	    	else
	    	   print "pongo telefono2"
	    	End If
	    End If
	    If Producto = "Soluciones Digitales izzi" and enable = "0" Then
	    	   print "Soluciones Digitales izzi"
	    	If varS = 0 Then
	    	   print "pongo telefono1"
               varS = 1	
	    	else
	    	   print "pongo telefono2"
	    	End If
	    End If
	End If
Next

wait 3
Browser("Todas las cuentas").Page("Todas las cuentas").WebElement("html id:=s_4_2_78_0_icon","visible:=True").Click
wait 5
Browser("Todas las cuentas").Page("Todas las cuentas").WebButton("value:=Seleccionar").Click

'print papaTelefonia

Browser("Todas las cuentas").Page("Resumen de la cuenta:").Link("Video").Click @@ script infofile_;_ZIP::ssf141.xml_;_
Browser("Todas las cuentas").Page("Resumen de la cuenta:").Link("Internet").Click @@ script infofile_;_ZIP::ssf142.xml_;_
Browser("Todas las cuentas").Page("Resumen de la cuenta:").Link("Telefonia").Click @@ script infofile_;_ZIP::ssf143.xml_;_
Browser("Todas las cuentas").Page("Resumen de la cuenta:").Link("OTT").Click @@ script infofile_;_ZIP::ssf144.xml_;_
Browser("Todas las cuentas").Page("Resumen de la cuenta:").Link("FTTH").Click @@ script infofile_;_ZIP::ssf145.xml_;_
Browser("Todas las cuentas").Page("Resumen de la cuenta:").Link("Otros Servicios").Click @@ script infofile_;_ZIP::ssf146.xml_;_
Browser("Todas las cuentas").Page("Resumen de la cuenta:").Link("Multiplay").Click @@ script infofile_;_ZIP::ssf147.xml_;_
Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebCheckBox("GRPITEM[~^^1-1NDFHFTR^^~[PORT[").Set "ON" @@ script infofile_;_ZIP::ssf148.xml_;_
Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebButton("Terminado").Click @@ script infofile_;_ZIP::ssf149.xml_;_
Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebButton("Detalles:Expandir Ítems").Click @@ script infofile_;_ZIP::ssf150.xml_;_

wait 3
Call ProductoActivaALL(sp_subvar(1),"ON") 
Call ProductoActiva("Telefonia wow","ON") 
wait 2

Browser("Siebel Communications_3").Page("Detalles: 1-129242227357").WebElement("WebElement").Click @@ script infofile_;_ZIP::ssf155.xml_;_
End If
'Password_#2022_QAS
   

   
'Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebList("select").Select "Extension de Video A wow" @@ script infofile_;_ZIP::ssf151.xml_;_
'Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebButton("Agregar ítem").Click @@ script infofile_;_ZIP::ssf152.xml_;_
'Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebCheckBox("GRPITEM[~^^1-1NDFHG0A^^~[PORT[").Set "ON" @@ script infofile_;_ZIP::ssf153.xml_;_
'Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebCheckBox("GRPITEM[~^^1-1NDFHG0A^^~[PORT[").Set "OFF" @@ script infofile_;_ZIP::ssf154.xml_;_


'
'wait 4
'Browser("Todas las cuentas").Page("Todas las cuentas").WebButton("Detalles:Agregar Items").Click
'	wait 5
'	Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebCheckBox("s_17_lSelectAll").Set "OFF"
'	
'	Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("Que empiece por").Click
'	
'
'	wait 2
'       
'		    Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("Que empiece por").Set "Portafolio Productos wizzplus wow"
'
'	Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebButton("Seleccionar producto:Ir").Click
'	wait 2
'	Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebButton("Seleccionar producto:Agregar").Click
'
'	wait 5
'	If not Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebButton("Detalles:Expandir Ítems").Exist(2) Then
'	Set oSave = CreateObject("Wscript.Shell")
'	oSave.SendKeys "{ENTER}"
'	Set oSave = Nothing
'	ExitActioniteration
' 	End If
' 	
' 	Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebButton("Detalles:Personalizar").Click
'	wait 2
'	If Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebButton("Detalles:Expandir Ítems").Exist(3) Then
'					finalstatus = "Error en el portafolio"
'               		print "Detalle:" &finalstatus
'					'Datatable.Value("o_error","Output") = finalstatus
'					ExitActioniteration
'				Else
'					reporter.ReportEvent micPass,"Product Portfolio","Product Added successfully"
'				End If
'wait 3
'contrato = Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebElement("html id:=1_s_4_l_Name").GetROProperty("title")
'print contrato
 @@ script infofile_;_ZIP::ssf184.xml_;_
 @@ script infofile_;_ZIP::ssf192.xml_;_

'wait 3
'filas = Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebElement("html id:=s_4_rc").GetROProperty("outertext")
'print filas
'vArr=split (filas,"-")
'iniPage=CInt(trim(vArr(0)))
'vArrSp=split (vArr(1),"de")
'intLastRowList=(trim(vArrSp(0)))
'intTOTALRowList=(trim(vArrSp(1)))
'print "total ="&iniPage&" "&intLastRowList&" "&intTOTALRowList
'If intLastRowList <> intTOTALRowList Then	
'      Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebElement("WebElement_3").Click @@ script infofile_;_ZIP::ssf199.xml_;_
'End If
'Call PegadodeTelefono()
''"No. Telefonico"
'numeroColumna = GetNumeroColumna("No. Telefonico")
'print numeroColumna
'
'Function GetNumeroColumna(ColumnaN)
'    Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebMenu("s_0").Select "Ver;Columnas mostradas...[Ctrl+Mayús+K]"
'For I = 1 To 10
'	columna = Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebList("acc_name:=Columnas seleccionadas:").GetItem(I)
'	print columna
'	If columna = ColumnaN Then
'		print "lo encontre en"&I
'		GetNumeroColumna = I
'		exit for
'	End If
'Next
'Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebButton("acc_name:=Columnas mostradas:Cancelar").Click	
'End Function
'
'
'Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebList("WebList").GetItem(0)
'I = 9
'equipoInside = "html id:="&I&"_CV_TN"
'enable = Browser("Todas las cuentas").Page("Todas las cuentas").WebElement(equipoInside).GetROProperty(ENTITY_REFERENCE_NODE)
'print "numero:"&enable
'call fnPickAndChoiceTelefonoPrincipal(9,5)
'

'Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebElement("jqgh_s_4_l_CV_TN").Click @@ script infofile_;_ZIP::ssf193.xml_;_
'Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebMenu("s_S_A4_headerMenu").Select "" @@ script infofile_;_ZIP::ssf194.xml_;_
'
'
'wait 2
'Browser("Todas las cuentas").Page("Resumen de la cuenta:").Link("Página inicial").Click @@ script infofile_;_ZIP::ssf201.xml_;_
'wait 2
'Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebList("Barra de vista de primer").Select "Ordenes de Servicio" @@ script infofile_;_ZIP::ssf202.xml_;_
'wait 2
'Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebButton("Ordenes de servicio:Consulta").Click @@ script infofile_;_ZIP::ssf203.xml_;_
'wait 2
'Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebElement("WebElement_4").Click @@ script infofile_;_ZIP::ssf204.xml_;_
'wait 2
'Browser("Todas las cuentas").Page("Todas las cuentas").WebEdit("Nro. Cuenta").highlight
'wait 2
'Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebEdit("Nº de orden").Set "1-129242338707" @@ script infofile_;_ZIP::ssf205.xml_;_
'wait 2
'Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebButton("Ordenes de servicio:Ir").Click @@ script infofile_;_ZIP::ssf206.xml_;_
'wait 2
'Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebElement("WebElement_4").Click @@ script infofile_;_ZIP::ssf207.xml_;_
'wait 2
''Browser("Todas las cuentas").Page("Resumen de la cuenta:").Link("1-129242338707").Click @@ script infofile_;_ZIP::ssf208.xml_;_
'Browser("Todas las cuentas").Page("Resumen de la cuenta:").Link("Class Name:=Link").Click
'wait 2
'Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebButton("Detalles:Expandir Ítems").Click @@ script infofile_;_ZIP::ssf209.xml_;_


'
'Function fnConsultarOrdenServicio()
'
'	ordenDeServicio = Datatable.Value("o_OrderService","Output")
'	print "Portafolio "&ordenDeServicio
'	
'    blnExisteOS = false
'	wait 2
'	Browser("Todas las cuentas").Page("Resumen de la cuenta:").Link("Página inicial").Click
'	Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebList("Barra de vista de primer").Select "Ordenes de Servicio" @@ script infofile_;_ZIP::ssf202.xml_;_
'	wait 2
'	Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebButton("Ordenes de servicio:Consulta").Click
'	wait 2
'	Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebElement("WebElement_4").Click
'	wait 1
'	Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebEdit("Nº de orden").Set ordenDeServicio
'	wait 1
'	Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebButton("Ordenes de servicio:Ir").Click
'	wait 3
'	Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebElement("WebElement_4").Click
'	wait 1
'	strLink = "name:="&ordenDeServicio
'	print strLink
'	Browser("Todas las cuentas").Page("Resumen de la cuenta:").Link(strLink).Click
' @@ script infofile_;_ZIP::ssf208.xml_;_
' 	wait 1
'	Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebButton("Detalles:Expandir Ítems").Click
'
'	wait 2
'	
'	strOSLink = Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebEdit("Nº de orden").GetROProperty("value")
'	strOSLink=fnISEmpty(strOSLink)
'	print strOSLink
'	
'	wait 3
'	
'	If strOSLink <> "campo vacio" Then
'	   wait 3
'	   blnExisteOS = true
'	elseBrowser("Todas las cuentas").Page("Resumen de la cuenta:").WebElement("WebElement_5").Click

'		print "No Existe NroCuenta"
'	   blnExisteOS = false
'	End If
'	wait 2
'	fnConsultarOrdenServicio = blnExisteOS
'End Function
'wait 4
'Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebEdit("acc_name:=No\. VTS").Set  "U001"
'Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebButton("title:=Orden de servicio:Programar").highlight

Function FuncionesFinales_2(valor1)
	Call subPegarTag(DataTable.Value ("p_Tag", "ProductSelection")) 
	wait 5
	Call BotonEnviar_2()
	
	  if esperaBotonE(80) then
            Ultima=Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebEdit("acc_name:=Ultima Modificacion","name:=s_1_1_.*_0","disabled:=0").GetROProperty("value")
            Datatable.Value("o_OSFecha","Output") = Ultima
      else
			Set oSave = CreateObject("Wscript.Shell")
			oSave.SendKeys "{ENTER}"
			Set oSave = Nothing
   		    print "Enviar tardo mas de 80 segundos" &""
			Datatable.Value("o_error","Output") = Datatable.Value("o_error","Output")&"Mas de 80 segundos tardo el boton Enviar"
      end if
	
End Function

Function esperaBotonE(max1)
    esperaBotonE = false
	For Iterator = 1 To max1
	    finalstatus=Browser("Todas las cuentas").Page("Resumen de la cuenta:").WebEdit("acc_name:=Estado","name:=s_1_1_.*_0","role:=combobox").GetROProperty("value")
    	If finalstatus = "Abierta" Then
		    print "(BotonEnviar) En el ciclo..."& Iterator
              wait 3
			  'wait 3
			  'wait 3
			  'wait 1
			  intCantidadCiclos = Iterator
		else
		    print "Salio ciclo"
		    esperaBotonE = true
		    Exit for
   		end if 
	Next

End Function

Function BotonEnviar_2()
  print "inicio BotonEnviar_2  ....................................	"
  banderaBotonEnviar = false
'  max1 = 40
   max1 = 80
  intCantidadCiclos =0
	For Iterator = 1 To max1
    	If not Browser("Todas las cuentas").Page("Página inicial de Siebel").WebButton("acc_name:=Enviar").Exist Then
		    print "(BotonEnviar) En el ciclo..."& Iterator
              wait 3
			  wait 3
			  wait 3
			  wait 1
			  intCantidadCiclos = Iterator
		else
		    print "Salio ciclo"
		    banderaBotonEnviar = true
		    Exit for
   		end if 
	Next

  if banderaBotonEnviar then
            Browser("Todas las cuentas").Page("Página inicial de Siebel").WebButton("acc_name:=Enviar").Click
  else
			Set oSave = CreateObject("Wscript.Shell")
			oSave.SendKeys "{ENTER}"
			Set oSave = Nothing
   		    print "Orden no ENVIADA:" &""
			Datatable.Value("o_error","Output") = Datatable.Value("o_error","Output")&"Orden no Enviada salio Mensaje  (Ciclos Error) :"&intCantidadCiclos    
  end if 
    print "Fin BotonEnviar_2 ....................................	"
'....................................	
End Function

 @@ script infofile_;_ZIP::ssf237.xml_;_
