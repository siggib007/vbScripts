REM ### Script: W2Kupdate
REM ### ScriptVersion: 2.2
REM ### Created: 2/14/2000
REM ### Modified: 9/25/2000
REM ### Group: Global Infrastructure Services
REM ### Contact: Doug Young
REM ### Description: Used to upgrade software on Windows 2000 servers at Microsoft.



REM ##################################################################################################
REM ######             Declare and set variables and Definitions for the script                 ######
REM #############d####################################################################################
REM ###### Declare variables ######
option explicit
dim arg
dim Wshshell, Wshnetwork, Wshmapdrives, Wshfile
dim computername, systemdrive, systemroot, backuproot, procarc, numproc, path, userdomain, username, tempdir
dim begdate, begtime, enddate, endtime, regarray(100,4), regarrayloc, netfactor, cdrunvar
dim bypass, debug, forceupdate, locationtestvar, argtext, okargtext, srvdebug

dim alldrives, harddrives, netdrives, cddrives, allservices, extraservices

dim scriptnamevar, scriptversionvar, scriptipakvar, scriptosvervar, scriptarcvar, scriptbuildvar
dim scriptbinvar, scripttempvar, scriptlogvar, scriptrebootvar, scripttsvar
dim scriptsitevar, scriptconswivar, scriptnetworkvar
dim scriptpathlocvar, scriptcmdlocvar, scriptinilocvar, scriptfileslocvar
dim scriptsharelocvar, scriptdrivelocvar, scriptserverlocvar, scriptbinlocvar, scriptloglocvar

dim srvvervar, srvbuildvar, srvcsdvar, srvarcvar, srvrolevar, srvsuitevar, srvencvar, srvdomainvar, srvdomtypevar
dim srvhardwarevar, srvhardtypevar, srvbootvar, srvmemvar, srvstrwrksvar

dim spvar, spchkvar, spfilevar, spdirvar, splocvar, splogvar
dim fixvar, fixlocvar, fixlogvar
dim symvar, symlocvar, symfulllocvar, symlogvar
dim encvar, enclocvar, enclogvar

dim wwwvar, wwwlogvar, wwwlocvar, wwwsymlocvar
dim termvar, sqlvar, exchangevar, sapvar

dim cpqbiosdate, cpqssdvar, cpqssdchkvar, cpqssdfilevar, cpqssdlogvar, cpqssdlocvar
dim cpqcimvar, cpqcimchkvar, cpqcimfilevar, cpqcimlogvar, cpqcimlocvar
dim cpqstrdrvvar, cpqstrdrvchkvar, cpqstrdrvfilevar, cpqstrdrvlogvar, cpqstrdrvlocvar
dim delbiosdate, delmannvar, delmannchkvar, delmannfilevar, delmannlogvar, delmannlocvar
dim delfastvar, delfastchkvar, delfastfilevar, delfastlogvar, delfastlocvar
dim deldrvvar, deldrvchkvar, deldrvfilevar, deldrvlogvar, deldrvlocvar

dim bldvar, bldlocvar
dim recconsvar, recconschkvar, recconsfilevar, recconslogvar, recconslocvar
dim opavar, opachkvar, opafilevar, opalogvar, opalocvar
dim senvar, senchkvar, senfilevar, senlogvar, senlocvar
dim niqvar, niqchkvar, niqfilevar, niqlogvar, niqlocvar
dim inocvar, inocchkvar, inocfilevar, inoclogvar, inoclocvar
dim oicvar, oicchkvar, oicfilevar, oiclogvar, oiclocvar
dim bacvar, bacchkvar, bacfilevar, baclogvar, baclocvar



REM ###### Set definitions ######
Set Wshnetwork = Wscript.CreateObject("Wscript.Network")
Set Wshshell = Wscript.CreateObject("Wscript.shell")
Set Wshfile = Wscript.CreateObject("Scripting.FileSystemObject")
computername = WshShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
systemdrive = WshShell.ExpandEnvironmentStrings("%SYSTEMDRIVE%")
systemroot = WshShell.ExpandEnvironmentStrings("%SYSTEMROOT%")
procarc = WshShell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")
numproc = WshShell.ExpandEnvironmentStrings("%NUMBER_OF_PROCESSORS%")
path = WshShell.ExpandEnvironmentStrings("%PATH%")
username = WshShell.ExpandEnvironmentStrings("%USERNAME%")
userdomain = WshShell.ExpandEnvironmentStrings("%USERDOMAIN%")
tempdir = WshShell.ExpandEnvironmentStrings("%TEMP%")

REM ###### Set script variables ######
scriptnamevar = left(wscript.ScriptName, Instr(1, wscript.ScriptName, ".vbs", 1) -1)
scriptversionvar = "2.2"
scriptipakvar = "NT5.02"
scriptosvervar = "5.0"
scriptarcvar = "x86"
scriptbuildvar = 2195
scriptbinvar = systemdrive & "\localbin"
scripttempvar = systemdrive & "\temp\" & scriptnamevar
scriptlogvar = systemroot & "\ipaklogs"
scriptrebootvar = "ask"
scripttsvar = "no"
scriptloglocvar = "\\cpitgfsa02\scriptlogs"
srvdebug = 0
debug = 0
regarrayloc = 1
forceupdate = 0
locationtestvar = 0
netfactor = 1
cdrunvar = "off"
bypass = "\\cpitgddsa01\gmas\bronze\ipak\nt5.02"
okargtext = ",/com,/del,/oth,/nostrwrk,/allfixes,/fullsym,/clrsym,/noupd,"
okargtext = okargtext & "/reb,/noreb,/debug,/forceup,/nobitmap,/slownet,"
okargtext = okargtext & "/inocon,/buacon,/nocim,/noopa,/nosen,/noniq,/noinoc,/nooic,/nobac,/notools,/new,"
okargtext = okargtext & "/wwwreg,/sqlreg,/excreg,/sapreg,/wwwipak,-?,/?,"
okargtext = okargtext & "head,paths,auth,config,locs,check,path,defpath,bug,scriptbug,noscriptbug,loctest,"
okargtext = okargtext & "notas,noots,noall,build,sp,nontbuilds,strtest,forcecd,128bit,"

REM ###### IPAK variables ######
bldvar = "on"
recconsvar = "1.0"
recconschkvar = systemdrive & "\cmdcons\autochk.exe"
recconsfilevar = ""
recconslogvar = ""
spvar = "1"
spdirvar = "SP1"
spchkvar = systemroot & "\system32\cmd.exe"
spfilevar = "7/21/2000"
splogvar = ""
fixvar = "on"
fixlogvar = ""
symvar = "small"
symlogvar = ""
encvar = "on"
enclogvar = ""

REM ###### Set hardware software variables ######
cpqbiosdate = "8/17/1998"
cpqssdvar = "5.03a"
cpqssdchkvar = systemroot & "\system32\drivers\sysmgmt.sys"
cpqssdfilevar = "6/1/2000"
cpqssdlogvar = ""
cpqcimvar = "4.80"
cpqcimchkvar = systemroot & "\system32\cpqmgmt\CqMgServ\CqMgServ.exe"
cpqcimfilevar = "5/30/2000"
cpqcimlogvar = ""
cpqstrdrvvar = "5-4.41A5"
cpqstrdrvchkvar = systemroot & "\system32\drivers\cpqkgpsa.sys"
cpqstrdrvfilevar = "1/31/2000"
cpqstrdrvlogvar = ""
delbiosdate = "10/27/1999"
delmannvar = "1.52"
delmannchkvar = systemdrive & "\program files\dell\OpenManage\nnm-HIP\bin\dcbase.exe"
delmannfilevar = "12/29/1999"
delmannlogvar = ""
delfastvar = "2.1-2972"
delfastchkvar = systemdrive & "\program files\perc2\afa\afaagent.exe"
delfastfilevar = "7/7/2000"
delfastlogvar = ""
deldrvvar = "2.1-2963"
deldrvchkvar = systemroot & "\system32\drivers\perc2.sys"
deldrvfilevar = "4/7/2000"
deldrvlogvar = ""

REM ###### Set tool variables ######
opavar = "3.0"
opachkvar = systemdrive & "\program files\opassist\opassist.exe"
opafilevar = "2/8/2000"
opalogvar = ""
senvar = "2.5.2.7"
senchkvar = systemdrive & "\sentry\progs\sass.exe"
senfilevar = "9/2/1998"
senlogvar = ""
niqvar = "3.0.361.8"
niqchkvar = systemdrive & "\netiq\bin\netiqmc.exe"
niqfilevar = "2/3/1999"
niqlogvar = ""
inocvar = "375"
inocchkvar = systemroot & "\system32\drivers\ino_fltr.sys"
inocfilevar = "3/7/2000"
inoclogvar = ""
oicvar = "2"
oicchkvar = systemdrive & "\oic\ObjectInactivityService.exe"
oicfilevar = "7/27/1999"
oiclogvar = ""
bacvar = "8.x"
bacchkvar = systemdrive & "\bentaa\beremote.exe"
bacfilevar = "4/20/2000"
baclogvar = ""

REM ###### Additional services variables ######
termvar = ""
wwwvar = ""
wwwlogvar = ""
sqlvar = ""
exchangevar = ""
sapvar = ""


REM #################################################################################################
REM ###                                                                                           ###
REM ###                                         MAIN                                              ###
REM ###                                                                                           ###
REM #################################################################################################
dim temp, reply, control, check
on error resume next

REM ###### Do pre-config tasks ######
if argument("scriptbug") <> 0 then debug = 1
if argument("noscriptbug") <> 0 then debug = 0
control = config_preconfig


REM ###### Print header ######
if control <> 1 then screenout "          _________________________________________________"
if control <> 1 then screenout "         |                                                 |"
if control <> 1 then screenout "         |                                                 |"
if control <> 1 then screenout "         | Windows 2000 Custom Configuration/Update Script |"
if control <> 1 then screenout "         |             " & scriptnamevar & " (Version " & scriptversionvar & ")             |"  
if control <> 1 then screenout "         |                                                 |"
if control <> 1 then screenout "         |            (Written for IPAK " & scriptipakvar & ")            |"
if control <> 1 then screenout "         |_________________________________________________|"
if control <> 1 then screenout "         \\" & Computername & " - " & userdomain & "\" & username
if argument("head") then control = 1


REM ###### Check for Question arguments ######
if control <> 1 then if argument("-?") <> 0 then control = syntax
if control <> 1 then if argument("/?") <> 0 then control = syntax

REM ###### Script paths and variables ######
if control <> 1 then screenout ""
if control <> 1 then screenout ""
if control <> 1 then screenout "SCRIPT PATHS AND VARIABLES"
if control <> 1 then control = config_scriptpath
if argument("paths") then control=1

REM ###### Script config and authen ######
if control <> 1 then screenout ""
if control <> 1 then screenout "SCRIPT AUTHENTICATION AND CONFIGURATION"
if control <> 1 then control = config_timestart
if control <> 1 then control = config_authen
if argument("auth") then genevent "I", "3", "An undocumented script debugging command was used to only run a portion of the script!" : control=1
if control <> 1 then control = config_getinfo
if control <> 1 then control = config_getswitches
if argument("config") then genevent "I", "3", "An undocumented script debugging command was used to only run a portion of the script!" : control=1
if control <> 1 then control = config_setlocvars
if control <> 1 and argument("bug") then bug : control=1
if argument("locs") then genevent "I", "3", "An undocumented script debugging command was used to only run a portion of the script!" : control=1

REM ###### Server Config Pre-check ######
if control <> 1 then screenout ""
if control <> 1 then screenout "SERVER REBOOT PRE-CHECKS (5)"
if control <> 1 then check = check + check_delldriver
if control <> 1 then check = check + hardware_perc2fast("true")
if control <> 1 then check = check + hardware_managednode("true")
if control <> 1 then check = check + update_sp("true")
if control <> 1 then check = check + check_sysbios
if check > 0 then
	control = 1
	screenout ""
	screenout ""
	screenout "************************************************************"
	screenout "Please perform or correct the above tasks or issues!"
	scrennout "Make sure you do the tasks in the above order!"
	screenout "When every thing is completed reboot the server!"
	screenout "When the server comes back up just re-run the update script."
	screenout "************************************************************"
end if
if argument("check") then genevent "I", "3", "An undocumented script debugging command was used to only run a portion of the script" : control=1

REM ###### Server information ######
if control <> 1 then screenout ""
if control <> 1 then screenout "ADMINISTRATION"
if control <> 1 then control = admin_asset
if control <> 1 then control = admin_prework
if control <> 1 then control = admin_groups
if control <> 1 then control = admin_audit

REM ###### NT services ######
if control <> 1 then screenout ""
if control <> 1 then screenout "NT SERVICES"
if control <> 1 then control = services_schedule
if control <> 1 then control = services_time
if control <> 1 then control = services_snmp
if control <> 1 then control = services_msmq

REM ###### Hardware Software ######
if control <> 1 then screenout ""
if control <> 1 then screenout "HARDWARE SOFTWARE"
if control <> 1 then control = hardware_cpqnic
if control <> 1 then control = hardware_ssd
if control <> 1 then control = hardware_cim
if control <> 1 then control = hardware_storageworksdriver
if control <> 1 then control = hardware_managednode("false")
if control <> 1 then control = hardware_perc2fast("false")

REM ###### Upgrade Software ######
if control <> 1 then screenout ""
if control <> 1 then screenout "UPDATE SOFTWARE"
if control <> 1 then getbuildcd
if control <> 1 then control = update_reccons
if control <> 1 then getupdatecd
if control <> 1 then control = update_sp("false")
if control <> 1 then control = update_sym
if control <> 1 then control = update_fix
if control <> 1 then control = update_128

REM ###### File ######
if control <> 1 then screenout ""
if control <> 1 then screenout "FILE CHANGES"
if control <> 1 then control = file_localbin
if control <> 1 then control = file_exchange
if control <> 1 then control = file_boot
if control <> 1 then control = file_delete
if control <> 1 then control = file_bitmap

REM ###### Registry ######
if control <> 1 then screenout ""
if control <> 1 then screenout "REGISTRY CHANGES"
if control <> 1 then control = registry_backup
if control <> 1 then control = registry_main
if control <> 1 then control = registry_filters
if control <> 1 then control = registry_pagefile
if control <> 1 then control = registry_path
if control <> 1 then control = registry_source
if control <> 1 then control = registry_diskperf
if control <> 1 then control = registry_site
if control <> 1 then control = registry_debug
if control <> 1 then control = registry_services

REM ###### Additional Tools ######
if control <> 1 then screenout ""
if control <> 1 then screenout "ADDITIONAL TOOLS"
if control <> 1 then control = tools_perfcol
if control <> 1 then control = tools_oic
if control <> 1 then gettoolscd
if control <> 1 then control = tools_opassist
if control <> 1 then control = tools_sentry
if control <> 1 then control = tools_netiq
if control <> 1 then control = tools_inoculan
if control <> 1 then control = tools_backupexec
if control <> 1 then getupdatecd
if control <> 1 then control = tools_other

REM ###### Additional Services ######
if control <> 1 then screenout ""
if control <> 1 then screenout "ADDITIONAL SERVICES"
if control <> 1 then control = services_www

REM ###### Script Completion ######
if control <> 1 then screenout ""
if control <> 1 then screenout "SCRIPT COMPLETION"
temp = wshshell.Run("change user /execute", 0, true)
if err.number <> 0 then err.clear
if instr(1, netdrives, scriptdrivelocvar, 1) = 0 then
	rem screenout "Removing network drive connection (" & scriptdrivelocvar & ")..."
	WshNetwork.RemoveNetworkDrive scriptdrivelocvar, TRUE, TRUE
end if
if control <> 1 then control = completion_other
if control <> 1 then control = completion_locallogging
if control <> 1 then control = completion_destlogging
if control <> 1 then control = completion_reboot
if control <> 1 and (Wshfile.fileexists(scriptlogvar & "\" & scriptnamevar & "-run.log")) then wshfile.deletefile scriptlogvar & "\" & scriptnamevar & "-run.log", TRUE
if control <> 1 and (Wshfile.fileexists(systemdrive & "\" & scriptnamevar & "-run.log")) then wshfile.movefile systemdrive & "\" & scriptnamevar & "-run.log", scriptlogvar & "\"

Wscript.DisconnectObject Wshnetwork
Wscript.DisconnectObject Wshshell
Wscript.DisconnectObject Wshfile
Set Wshnetwork=nothing
Set Wshshell=nothing
Set Wshfile=nothing





REM #################################################################################################
REM ###                                                                                           ###
REM ###                               One Time Run Functions                                      ###
REM ###                                                                                           ###
REM #################################################################################################
REM #################################################################################################
REM ###                                   Pre-Config                                              ###
REM #################################################################################################
Function config_preconfig()
	dim arg, temp
	on error resume next

	REM ###### Sets up Vbscript to always run in command window ######
	if instr(1, wscript.fullname, "cscript.exe", 1) = 0 then
		if err.number <> 0 then err.clear
		temp = wshshell.Run("cmd /c ""cscript //h:cscript //nologo //s 1>nul 2>nul""", 0, true)
		temp = MsgBox ("The script has changed the default output of Windows Scripting Host to the command prompt." & vbCrLf & "This is pop up is normal, just re-run the script!", 0, "WSH default changed to cscript.")
		config_preconfig = 1
		exit function
	else
		temp = wshshell.Run("change user /install", 0, true)
		if err.number <> 0 then err.clear
	end if

	if err.number <> 0 then logerror "End of config_preconfig Function", err.number : err.clear
	config_preconfig = 0
End Function




REM #################################################################################################
REM ###                                  Core Script Paths                                        ###
REM #################################################################################################
Function config_scriptpath()
	dim arg, text, a, drivecoll, dc, wshdrive
	on error resume next

	REM ######  Set script path variable ######
	screenout "Getting script path information..."
	scriptpathlocvar = Replace(wscript.ScriptFullName, "\" & wscript.ScriptName, "", 1, -1, 1)
	if argument("defpath") <> 0 then
		scriptpathlocvar = bypass
	end if
	if argument("path") <> 0 then
		arg = argument("path")
		scriptpathlocvar = nextargument(arg)
		if scriptpathlocvar = "" then
			screenout ""
			screenout "You must specify the Script path after the 'path' switch!"
			screenout "Example: path \\servername\gmas\gold\ipak\nt5.01"
			config_scriptpath = 1
			exit function
		end if
	end if

	REM ### Check script path source location ###
	if not (Wshfile.FolderExists(scriptpathlocvar)) then
		screenout ""
		screenout "The Script is unable to find the directory in which it was ran from!"
		screenout "  (" & scriptpathlocvar & ")"
		screenout "Make sure you are running the script as a UNC or from a net use drive letter."
		config_scriptpath = 1
		exit function
	end if

	REM ###### Set core path variables  ######
	scriptcmdlocvar = scriptpathlocvar
	scriptinilocvar = scriptpathlocvar & "\ini"
	scriptfileslocvar = scriptpathlocvar & "\bin\" & procarc

	REM ### Check ini files source location ###
	if not (Wshfile.FolderExists(scriptinilocvar)) then
		screenout ""
		screenout "Could not locate the ini directory needed for the script!"
		screenout "Make sure you are running the script as a UNC path or from a net use drive letter."
		screenout "Make sure the below location exists from where you are running the script."
		screenout "  (" & scriptinilocvar & ")"
		config_scriptpath = 1
		exit function
	end if

	REM ### Check script files location path ###
	if not (Wshfile.FolderExists(scriptfileslocvar)) then
		screenout ""
		screenout "Could not locate the executable files needed for the script!"
		screenout "Make sure you are running the script as a UNC path or from a net use drive letter."
		screenout "Make sure the below location exists from where you are running the script."
		screenout "  (" & scriptfileslocvar & ")"
		config_scriptpath = 1
		exit function
	end if

	REM ### Check script files ###
	if not (Wshfile.FileExists(scriptfileslocvar & "\reg.exe")) then screenout "Could not find Reg.exe for the script!" : config_scriptpath = 1: exit function
	if not (Wshfile.FileExists(scriptfileslocvar & "\regsecadd.exe")) then screenout "Could not find RegSecAdd.exe for the script!" : config_scriptpath = 1 : exit function
	if not (Wshfile.FileExists(scriptfileslocvar & "\auditpol.exe")) then screenout "Could not find AuditPol.exe for the script!" : config_scriptpath = 1 : exit function
	if not (Wshfile.FileExists(scriptfileslocvar & "\bmpedit.exe")) then screenout "Could not find BMPEdit.exe for the script!" : config_scriptpath = 1 : exit function
	if not (Wshfile.FileExists(scriptfileslocvar & "\ntmem.exe")) then screenout "Could not find NTmem.exe for the script!" : config_scriptpath = 1 : exit function
	if not (Wshfile.FileExists(scriptfileslocvar & "\tlist.exe")) then screenout "Could not find TList.exe for the script!" : config_scriptpath = 1 : exit function
	if not (Wshfile.FileExists(scriptfileslocvar & "\kill.exe")) then screenout "Could not find Kill.exe for the script!" : config_scriptpath = 1 : exit function
	if not (Wshfile.FileExists(scriptfileslocvar & "\reboot.exe")) then screenout "Could not find Reboot.exe for the script!" : config_scriptpath = 1 : exit function
	if not (Wshfile.FileExists(scriptfileslocvar & "\ipakevnt.exe")) then screenout "Could not find IpakEvnt.exe for the script!" : config_scriptpath = 1 : exit function
	if not (Wshfile.FileExists(scriptfileslocvar & "\assettag.exe")) then screenout "Could not find AssetTag.exe for the script!" : config_scriptpath = 1 : exit function
	if not (Wshfile.FileExists(scriptfileslocvar & "\evaltest.exe")) then screenout "Could not find EvalTest.exe for the script!" : config_scriptpath = 1 : exit function
	screenout "  Script Path and name - " & scriptpathlocvar & "\" & scriptnamevar

	REM ###### Get drives ######
	screenout "Getting drive information..."
	set drivecoll=wshfile.drives
	for each dc in drivecoll
		alldrives = alldrives & dc.driveletter & ": "
		Set wshdrive = wshfile.GetDrive(dc.driveletter)
		Select Case wshdrive.DriveType
			Case 0: 
			Case 1: 
			Case 2: harddrives = harddrives & dc.driveletter & ": "
			Case 3: netdrives = netdrives & dc.driveletter & ": "
			Case 4: cddrives = cddrives & dc.driveletter & ": "
			Case 5: 
		End Select
	next
	screenout "  All drives found on machine - "& alldrives
	screenout "  Logical drive(s) on machine - "& harddrives
	screenout "  CD-ROM drive(s) on machine - "& cddrives
	screenout "  Network drive(s) on machine - "& netdrives

	screenout "Getting script share and drive path information..."
	REM ###### Check for UNC or drive letter connection ######
	if Instr(1, scriptpathlocvar, ":", 1) = 2 then
		scriptdrivelocvar = left(scriptpathlocvar, Instr(1, scriptpathlocvar, ":", 1))
		if scriptdrivelocvar = "" then
			screenout ""
			text = "Unable to get drive letter from script location source directory!"
			screenout text
			genevent "E", "3", text
			config_scriptpath = 1
			exit function
		end if
		Set wshmapDrives = WshNetwork.EnumNetworkDrives
		for a = 0 to wshmapdrives.Count -1 step 2
			screenout "  Mapped Drive - " & wshmapdrives.Item(a) & " - " & wshmapdrives.Item(a+1)
			if not wshmapdrives.Item(a) = "" and Instr(1, scriptpathlocvar, wshmapdrives.Item(a), 1) = 1 and a < wshmapdrives.Count then
				scriptsharelocvar = wshmapdrives.Item(a+1)
			end if
		next
		if scriptsharelocvar = "" then scriptsharelocvar = scriptdrivelocvar
		if Instr(Instr(3, scriptsharelocvar, "\", 1)+1, scriptsharelocvar, "\", 1) <> 0 then 
			screenout ""
			text = "The directory structure that is used in the net use command is too deep!"
			screenout text
			screenout "Remove the existing drive connection and re-create it using ONLY the server name and share name."
			screenout "  Example: Net use * \\server\share   NOT: Net use * \\server\share\dir"
			screenout "Make sure only ONE drive letter is mapped to the source server."
			genevent "E", "3", text
			config_scriptpath = 1
			exit function
		end if
		
	else
		scriptsharelocvar = left(scriptpathlocvar, Instr(instr(3,scriptpathlocvar, "\", 1)+1, scriptpathlocvar, "\", 1)-1)
		if scriptsharelocvar = "" then
			screenout ""
			text = "Unable to get server and share from script location source directory!"
			screenout text
			genevent "E", "3", text
			config_scriptpath = 1
			exit function
		end if
		Set wshmapDrives = WshNetwork.EnumNetworkDrives
		for a = 0 to wshmapdrives.Count -1 step 2
			screenout "  Mapped Drive - " & wshmapdrives.Item(a) & " - " & wshmapdrives.Item(a+1)
			if not wshmapdrives.Item(a) = "" and wshmapdrives.Item(a+1) = scriptsharelocvar and a < wshmapdrives.Count then
				scriptdrivelocvar = wshmapdrives.Item(a)
			end if
		next
		Set wshmapDrives=nothing
		if scriptdrivelocvar = "" then
			if instr(1, alldrives, "Z:", 1) = 0 then scriptdrivelocvar = "Z:"
			if instr(1, alldrives, "Y:", 1) = 0 then scriptdrivelocvar = "Y:"
			if instr(1, alldrives, "X:", 1) = 0 then scriptdrivelocvar = "X:"
			if instr(1, alldrives, "W:", 1) = 0 then scriptdrivelocvar = "W:"
			if instr(1, alldrives, "V:", 1) = 0 then scriptdrivelocvar = "V:"
			if instr(1, alldrives, "U:", 1) = 0 then scriptdrivelocvar = "U:"
			if instr(1, alldrives, "T:", 1) = 0 then scriptdrivelocvar = "T:"
			if instr(1, alldrives, "S:", 1) = 0 then scriptdrivelocvar = "S:"
			if instr(1, alldrives, "R:", 1) = 0 then scriptdrivelocvar = "R:"
			if instr(1, alldrives, "Q:", 1) = 0 then scriptdrivelocvar = "Q:"
			if instr(1, alldrives, "P:", 1) = 0 then scriptdrivelocvar = "P:"
			if instr(1, alldrives, "O:", 1) = 0 then scriptdrivelocvar = "O:"
			if instr(1, alldrives, "N:", 1) = 0 then scriptdrivelocvar = "N:"
			if instr(1, alldrives, "M:", 1) = 0 then scriptdrivelocvar = "M:"
			if instr(1, alldrives, "L:", 1) = 0 then scriptdrivelocvar = "L:"
			if instr(1, alldrives, "K:", 1) = 0 then scriptdrivelocvar = "K:"
			if scriptdrivelocvar = "" then
				screenout ""
				text = "No drive letters are available to map to the source server!"
				screenout text
				genevent "E", "3", text
				config_scriptpath = 1
				exit function
			end if
			if debug = 1 then screenout "  Mapping Drive - " & scriptdrivelocvar & " - " & scriptsharelocvar
			if err.number <> 0 then err.clear
			WshNetwork.MapNetworkDrive scriptdrivelocvar, scriptsharelocvar
			if err.number <> 0 then
				err.clear
				screenout ""
				text = "Could not map a drive letter to the source server share!"
				screenout text
				genevent "E", "3", text
				config_scriptpath = 1
				exit function
			end if
		end if			
	end if

	REM ###### Get script source server name ######
	if instr(3, scriptsharelocvar, "\", 1) <> 0 then
		scriptserverlocvar = mid(scriptsharelocvar, 3, instr(3, scriptsharelocvar, "\", 1) - 3)
	else
		scriptserverlocvar = ""
	end if

	REM ###### Display variables ######
	screenout "  Script Server variable - " & scriptserverlocvar
	screenout "  Script Share variable - " & scriptsharelocvar
	screenout "  Script Drive variable - " & scriptdrivelocvar

	if err.number <> 0 then logerror "End of config_scriptpath Function", err.number : err.clear
	config_scriptpath = 0
End Function





REM #################################################################################################
REM ###                                 Set Time and Start entry                                  ###
REM #################################################################################################
Function config_timestart()
	dim strcmdline, temp, text
	on error resume next

	strcmdLine = "cmd /c ""net time \\"& scriptserverlocvar & " /set /y"""
	temp = wshshell.Run(strcmdline, 0, true)
	begdate=date
	begtime=time
	text = scriptnamevar & " for IPAK " & scriptipakvar & " started on " & begdate & " at " & begtime & "."
	screenout text
	genevent "I", "1", text

	if err.number <> 0 then logerror "End of config_timestart Function", err.number : err.clear
	config_timestart = 0
End Function





REM #################################################################################################
REM ###                                     Authentication                                        ###
REM #################################################################################################
Function config_authen()
	dim strcmdline, temp, text, a, wshtempfile, alltext, suite, upgrade
	on error resume next

	REM ### Create temp directory ###
	if not (Wshfile.FolderExists(systemdrive & "\temp")) then Wshfile.CreateFolder(systemdrive & "\temp")
	if not (Wshfile.FolderExists(scripttempvar)) then Wshfile.CreateFolder(scripttempvar)

	REM ### check NT version ###
	srvvervar = WshShell.Regread("HKLM\Software\microsoft\windows nt\currentversion\currentversion")
	if err.number <> 0 then err.clear
	if srvvervar = scriptosvervar then
		screenout "Authentication for Windows Version " & srvvervar & " OK."
	else
		screenout ""
		text = "This is the wrong script for this version of Windows!"
		screenout text
		screenout "Please run a script that is designed to run on Windows version " & srvversion
		genevent "E", "3", text
		config_authen = 1
		exit function
	end if

	REM ### Check build number ###
	srvbuildvar = WshShell.RegRead("HKLM\Software\microsoft\windows nt\currentversion\currentbuildnumber")
	if err.number <> 0 then err.clear
	srvcsdvar = WshShell.Regread("HKLM\Software\microsoft\windows nt\currentversion\csdversion")
	if err.number <> 0 then err.clear
	if cint(srvbuildvar) = scriptbuildvar then
		screenout "Authentication for build " & srvbuildvar & " OK."
	else
		screenout ""
		text = "This script does not support running on this Build of Windows 2000!"
		screenout text
		screenout "Since most of the software now being installed on Windows 2000 is dependent upon 2195"
		screenout "older builds can no longer be supported."
		screenout "Please upgrade the server to 2195 RTM."
		genevent "E", "3", text
		config_authen = 1
		exit Function
	end if

	REM ###### Get server Suite type ######
	suite =  WshShell.RegRead("HKLM\system\currentcontrolset\control\productoptions\productsuite")
	if err.number <> 0 then err.clear
	for a = 0 to 5
		if isempty(suite(a)) = True then Exit for
		if suite(a) = "Enterprise" then srvsuitevar = "AS"
		if suite(a) = "DataCenter" then srvsuitevar = "DTC"
	Next
	if srvsuitevar = "AS" then
		screenout "Authentication for " & srvsuitevar & " product suite OK."
	elseif srvsuitevar = "DTC" then
		screenout "Authentication for " & srvsuitevar & " product suite OK."
		spvar = "off"
	else
		screenout ""
		text = "This script does not support the Professional or Server product suites!"
		screenout text
		screenout "Please rebuild the server to the Advanced Server or DataCenter suite."
		screenout "Upgrades to Advanced Server or DataCenter are not supported."
		genevent "E", "3", text
		config_authen = 1
		exit Function
	end if

	REM ###### Check architecture ######
	if instr(1, scriptarcvar, procarc, 1) <> 0 then
		screenout "Authentication for " & procarc & " Processor Architecture is OK."
		srvarcvar = lcase(procarc)
		if srvarcvar = "x86" then srvarcvar = "i386"
	else
		screenout ""
		text = "This script does not support this architecture of Windows 2000!"
		screenout text
		screenout "You have to run it on an x86 machine or a project needs started to support this type of architecture."
		genevent "E", "3", text
		config_authen = 1
		exit function
	end if	

	REM ###### Check for Full version ######
	strcmdLine = "cmd /c """& scriptfileslocvar & "\evaltest -s" & computername & " -v >" & scripttempvar & "\evaltemp.txt"""
	temp = wshshell.Run(strcmdline, 0, true)
	if (Wshfile.fileexists(scripttempvar & "\evaltemp.txt")) then 
		Set wshtempfile = wshfile.OpenTextFile(scripttempvar & "\evaltemp.txt", 1)
		alltext = wshtempfile.readall
		wshtempfile.Close
		wscript.DisconnectObject wshtempfile
		Set wshtempfile=nothing
		if alltext = "" then
			screenout ""
			text = "Unable to create a text file in the temporary directory for the script!"
			screenout text
			screenout "Make sure that " & scripttempvar & " exists and is writable."
			screenout "Make sure that evaltest.exe is available and working properly."
			genevent "E", "3", text
			config_authen = 1
			exit function
		elseif instr(1, alltext, "Evaluation expires", 1) and srvsuitevar = "AS" then
			text = "This server is currently running on the evaluation copy of Windows 2000!"
			screenout text
			screenout "The server will need to be upgraded before the script will continue."
			screenout "Please upgrade the server to 2195 RTM."
			screenout "Thank you!"
			genevent "E", "3", text
			config_authen = 1
			exit function
		elseif instr(1, alltext, "Evaluation expires", 1) and srvsuitevar = "DTC" then
			screenout "Authentication for evaluation version is acceptable for " & srvsuitevar & "."
		else
			screenout "Authentication for RTM " & srvsuitevar & " OK."
		end if
	else
		screenout ""
		text = "Unable to create a text file in the temporary directory for the script!"
		screenout text
		screenout "Make sure that " & scripttempvar & " exists and is writable."
		screenout "Make sure that evaltest.exe is available and working properly."
		genevent "E", "3", text
		config_authen = 1
		exit function
	end if

	REM ###### Check for admin access ######
	strcmdline = "cmd /c ""net localgroup administrators bogus\account /del 2>" & scripttempvar & "\admintemp.txt"""	
	temp = wshshell.Run(strcmdline, 0, true)
	if (Wshfile.fileexists(scripttempvar & "\admintemp.txt")) then 
		Set wshtempfile = wshfile.OpenTextFile(scripttempvar & "\admintemp.txt", 1)
		alltext = wshtempfile.readall
		wshtempfile.Close
		wscript.DisconnectObject wshtempfile
		Set wshtempfile=nothing
		if instr(1, alltext, "Access is Denied", 1) then
			screenout ""
			text = "This script cannot run without the user having administrative access to the server!"
			screenout text
			screenout "Please log off and log back on using an account with administrative access."
			genevent "E", "3", text
			config_authen = 1
			exit function
		else
			screenout "Authentication for Administrative access is OK."
		end if
	else
		screenout ""
		text = "Unable to create a text file in the temporary directory for the script!"
		screenout text
		screenout "Make sure that " & scripttempvar & " exists and is writable."
		screenout "Make sure that net.exe is available and working properly (net localgroup)."
		genevent "E", "3", text
		config_authen = 1
		exit function
	end if

	REM ###### Check for TS Session ######
	strcmdline = "cmd /c ""qwinsta >" & scripttempvar & "\tssesstemp.txt"""	
	temp = wshshell.Run(strcmdline, 0, true)
	if (Wshfile.fileexists(scripttempvar & "\tssesstemp.txt")) then 
		Set wshtempfile = wshfile.OpenTextFile(scripttempvar & "\tssesstemp.txt", 1)
		alltext = wshtempfile.readall
		wshtempfile.Close
		wscript.DisconnectObject wshtempfile
		Set wshtempfile=nothing
		if instr(1, alltext, ">rdp", 1) then
			screenout "Authentication for Terminal Service session is OK."
			scripttsvar = "yes"
		else
			screenout "Authentication for Terminal Service session is OK."
		end if
	else
		screenout ""
		text = "Unable to create a text file in the temporary directory for the script!"
		screenout text
		screenout "Make sure that " & scripttempvar & " exists and is writable."
		screenout "Make sure that qwinsta.exe is available and working properly (qwinsta)."
		genevent "E", "3", text
		config_authen = 1
		exit function
	end if

	REM ###### Check Drive free space  ######
	if checkfreespace(150) = 0 then
		screenout "Authentication for 150 MB minimum boot drive space (" & systemdrive & ") is OK."
	else
		screenout ""
		text = "There is not enough drive space on the boot drive (" & systemdrive & ") to run the update script!"
		screenout text
		screenout "Please free up a minimum of 150 MB for the script to run."
		genevent "E", "3", text
		config_authen = 1
		exit function
	end if

	if err.number <> 0 then logerror "End of config_authen Function", err.number : err.clear
	config_authen = 0
End Function



REM #################################################################################################
REM ###                                Get config information                                     ###
REM #################################################################################################
Function config_getinfo()
	dim temp, strcmdline, text, a, line, pos, pos2, pos3, size1, size2, description
	dim wshbootdrive, wshtempfile, srvrole, srvenc, alltext, terminst, termchk, sapchk
	dim system, systemset, computerob, object, objLocator, objService, objEnumerator, objInstance
	dim oemname, oemmodel, oemtype
	dim santype, securepath, hubaddress, validaddress
	on error resume next

	REM ### Get server role ###
	srvrole = WshShell.RegRead("HKLM\system\currentcontrolset\control\productoptions\producttype")
	if err.number <> 0 then err.clear
	if srvrole = "ServerNT" then
		srvrolevar = "Server"
	elseif srvrole = "LanmanNT" then
		srvrolevar = "Domain Controller"
		symvar = "full"
	else
		srvrolevar = "Unknown"
	end if
	screenout "  Detected " & srvrolevar & " role."

	REM ###### Get encryption type ######
	srvenc =  WshShell.RegRead("HKLM\software\microsoft\windows nt\currentversion\hotfix\128bit\file 1\new file")
	for a = 0 to 5
		if isempty(srvenc) = True then srvencvar = "40" : encvar = "off"
		if srvenc = "Encpack" then srvencvar = "128" : encvar = "128"
	Next
	if argument("128bit") <> 0 then encvar = "128"
	screenout "  Detected " & srvencvar & " Bit Encryption."

	REM ###### Get Domain name ######
	Set SystemSet = GetObject("winmgmts:").InstancesOf ("Win32_ComputerSystem")
	for each System in SystemSet
		srvdomainvar = System.Domain
	next
	Set System=nothing
	Wscript.DisconnectObject SystemSet
	Set SystemSet=nothing
	screenout "  Detected " & srvdomainvar & " Domain." 	

	REM ###### Get domain type ######
	if (Wshfile.fileexists(scriptinilocvar & "\muddomainlistw2k.ini")) then
		Set wshtempfile = wshfile.OpenTextFile(scriptinilocvar & "\muddomainlistw2k.ini", 1)
		alltext = wshtempfile.readall
		wshtempfile.Close
		wscript.DisconnectObject wshtempfile
		Set wshtempfile=nothing
		if Instr(1, alltext, srvdomainvar, 1) <> 0 then
			srvdomtypevar = "MUD"
		else
			srvdomtypevar = "RESOURCE"
		end if
	else
		screenout ""
		text = "The script was unable to find the MudDomainListW2K.ini file!"
		screenout text
		screenout "Make sure that " & scriptinilocvar & " exists and contains the necessary INI file."
		genevent "E", "3", text
		config_getinfo = 1
		exit function
	end if
	screenout "  Detected " & srvdomtypevar & " Domain." 	


	REM ###### Get hardware type ######
	Set objLocator = CreateObject("WbemScripting.SWbemLocator")
	Set objService = objLocator.ConnectServer ("", "root\cimv2", "", "")
	ObjService.Security_.impersonationlevel = 3
    	Set objEnumerator = objService.ExecQuery("Select vendor, version From Win32_ComputerSystemProduct",,0)
  	For Each objInstance in objEnumerator
 		oemname = objInstance.properties_("vendor")
 		oemmodel = trim(objInstance.properties_("version"))
 		oemtype = trim(objInstance.properties_("name"))
	Next
	Wscript.DisconnectObject objEnumerator
	set objEnumerator = nothing
	Wscript.DisconnectObject objService
	set objService = nothing
	Wscript.DisconnectObject objLocator
	set objLocator = nothing
	if Instr(1, oemname, "COMPAQ", 1) <> 0 then
		if (Wshfile.FileExists(systemroot & "\system32\drivers\cpqarray.sys")) then srvhardwarevar = "COMPAQ"
		if (Wshfile.FileExists(systemroot & "\system32\drivers\cpqarry2.sys")) then srvhardwarevar = "COMPAQ"
		if (Wshfile.FileExists(systemroot & "\system32\drivers\cpqcissm.SYS")) then srvhardwarevar = "COMPAQ"
		if (Wshfile.FileExists(systemroot & "\system32\drivers\perc2.sys")) then srvhardwarevar = "OTHER"
		if (Wshfile.FileExists(systemroot & "\system32\drivers\afascsi.sys")) then srvhardwarevar = "OTHER"
	elseif Instr(1, oemname, "Dell", 1) <> 0 then
		if (Wshfile.FileExists(systemroot & "\system32\drivers\cpqarray.sys")) then srvhardwarevar = "OTHER"
		if (Wshfile.FileExists(systemroot & "\system32\drivers\cpqarry2.sys")) then srvhardwarevar = "OTHER"
		if (Wshfile.FileExists(systemroot & "\system32\drivers\cpqcissm.SYS")) then srvhardwarevar = "OTHER"
		if (Wshfile.FileExists(systemroot & "\system32\drivers\perc2.sys")) then srvhardwarevar = "DELL"
		if (Wshfile.FileExists(systemroot & "\system32\drivers\afascsi.sys")) then srvhardwarevar = "DELL"
	elseif oemname = "" then
		if (Wshfile.FileExists(systemroot & "\system32\drivers\cpqarray.sys")) then srvhardwarevar = "COMPAQ"
		if (Wshfile.FileExists(systemroot & "\system32\drivers\cpqarry2.sys")) then srvhardwarevar = "COMPAQ"
		if (Wshfile.FileExists(systemroot & "\system32\drivers\cpqcissm.SYS")) then srvhardwarevar = "COMPAQ"
		if (Wshfile.FileExists(systemroot & "\system32\drivers\perc2.sys")) then srvhardwarevar = "DELL"
		if (Wshfile.FileExists(systemroot & "\system32\drivers\afascsi.sys")) then srvhardwarevar = "DELL"
	else
		srvhardwarevar = "OTHER"
	end if
	if not oemmodel = "" then
		srvhardtypevar = oemmodel
	elseif oemmodel = "" and instr(1, oemtype, "ProLiant ", 1) then
		srvhardtypevar = right(oemtype, len(oemtype) - 9)
	elseif oemmodel = "" and instr(1, oemtype, "PowerEdge ", 1) then
		srvhardtypevar = mid(oemtype, 11, 4)
	else
		srvhardtypevar = "UNKNOWN"
	end if
	screenout "  Detected " & srvhardwarevar & " hardware."
	screenout "  Detected " & srvhardtypevar & " server model."


	REM ###### Get storage works information ######
	if (argument("strtest") <> 0 or (Wshfile.FileExists(systemroot & "\system32\drivers\lp6nds35.sys")) or (Wshfile.FileExists(systemroot & "\system32\drivers\cpqkgpsa.sys"))) and argument("/nostrwrk") = 0 then 
		description = "Is this server a 'hub' or a 'switch'?" & vbCrLf & vbCrLf & "(Enter 'cancel' to exit the script.)" & vbCrLf & "(You can use the first letter of the word.)"
		santype = inputbox(description, scriptnamevar & " - Question!")
		description = "Is this server running (or going to be running) SecurePath software?"
		securepath = msgbox(description, 4, scriptnamevar & " - Question!")
		rem screenout "  (Secure Path):" & securepath
		if santype = "h" or santype = "hub" then
			validaddress = "0x01, 0x02, 0x04, 0x08, 0x0F, 0x10, 0x17, 0x18, 0x1B, 0x1D, 0x1E, 0x1F, 0x23, 0x25, 0x26, 0x27,"
			validaddress = validaddress & vbCrLf & "0x29, 0x2A, 0x2B, 0x2C, 0x2D, 0x2E, 0x31, 0x32, 0x33, 0x34, 0x35, 0x36, 0x39, 0x3A, 0x3C, 0x43,"
			validaddress = validaddress & vbCrLf & "0x45, 0x46, 0x47, 0x49, 0x4A, 0x4B, 0x4C, 0x4D, 0x4E, 0x51, 0x52, 0x53, 0x54, 0x55, 0x56, 0x59,"
			validaddress = validaddress & vbCrLf & "0x5A, 0x5C, 0x63, 0x65, 0x66, 0x67, 0x69, 0x6A, 0x6B, 0x6C, 0x6D, 0x6E, 0x71, 0x72, 0x73, 0x74,"
			validaddress = validaddress & vbCrLf & "0x75, 0x76, 0x79, 0x7A, 0x7C, 0x80, 0x81, 0x82, 0x84, 0x88, 0x8F, 0x90, 0x97, 0x98, 0x9B, 0x9D,"
			validaddress = validaddress & vbCrLf & "0x9E, 0x9F, 0xA3, 0xA5, 0xA6, 0xA7, 0xA9, 0xAA, 0xAB, 0xAC, 0xAD, 0xAE, 0xB1, 0xB2, 0xB3, 0xB4,"
			validaddress = validaddress & vbCrLf & "0xB5, 0xB6, 0xB9, 0xBA, 0xBC, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCC, 0xCD, 0xCE,"
			validaddress = validaddress & vbCrLf & "0xD1, 0xD2, 0xD3, 0xD4, 0xD5, 0xD6, 0xD9, 0xDA, 0xDC, 0xE0, 0xE1, 0xE2, 0xE4, 0xE8, 0xEF."
			description = "Please enter the Hub's ALPA address..." & vbCrLf & vbCrLf & "Default address should be: 0x01" & vbCrLf & "You must use a valid ALPA address."
			hubaddress = inputbox(description, scriptnamevar & " - ALPA Address!", "0x01")
			if hubaddress = "" or instr(1, validaddress, hubaddress, 1) = 0 then
				screenout ""
				text = "An invalid entry was entered in the input box for the ALPA address!"
				screenout text
				screenout "The ALPA address should be a hexadecimal number of this form '0x01'."
				screenout "If you are not sure of the address of this machine, contact the server owner or a "
				screenout "SAN administrator. Do NOT attempt to update the server unless you are sure of the address!"
				screenout "Valid ALPA addresses are:"
				screenout validaddress
				genevent "E", "3", text
				config_getinfo = 1
				exit function
			end if
			if securepath = 6 then
				srvstrwrksvar = "HUB-YES-" & hubaddress
			elseif securepath = 7 then
				srvstrwrksvar = "HUB-NO-" & hubaddress
			end if
		elseif (santype = "s" or santype = "switch") and securepath = 6 then
			srvstrwrksvar = "SWITCH-YES"
		elseif (santype = "s" or santype = "switch") and securepath = 7 then
			srvstrwrksvar = "SWITCH-NO"
		elseif santype = "c" or santype = "cancel" then
			screenout ""
			text = "The script has been cancelled by the user from SAN type information!"
			screenout text
			screenout "If you are not sure of the configuration of this machine, contact the server owner or a "
			screenout "SAN administrator. Do NOT attempt to update the server unless you are sure of the configuration!"
			genevent "I", "3", text
			config_getinfo = 1
			exit function
		else
			screenout ""
			text = "An invalid entry was entered in the input box for the SAN type!"
			screenout text
			screenout "When inputting SAN type information use 'switch' ('s') or 'hub' ('h')."
			screenout "Always use lower case letters."
			genevent "E", "3", text
			config_getinfo = 1
			exit function
		end if
	elseif argument("/nostrwrk") <> 0 then
		srvstrwrksvar = "DISABLED"
	else
		srvstrwrksvar = "ABSENT"
	end if
	screenout "  Detected " & srvstrwrksvar & " Storage Works."

	REM ###### Get File System ######
	Set wshbootdrive = wshfile.GetDrive(systemdrive)
	srvbootvar = wshbootdrive.FileSystem
	Wscript.DisconnectObject Wshbootdrive
	Set Wshbootdrive=nothing
	screenout "  Detected " & srvbootvar & " boot partition."

	REM ###### Get Current memory size ######
	strCmdLine = "cmd /c """& scriptfileslocvar & "\ntmem >" & scripttempvar & "\pagefile1temp.txt"""
	temp = wshshell.Run(strcmdline, 0, true)
	if err.number <> 0 then err.clear
	if (Wshfile.fileexists(scripttempvar & "\pagefile1temp.txt")) then 
		Set wshtempfile = wshfile.OpenTextFile(scripttempvar & "\pagefile1temp.txt", 1)
		if err.number <> 0 then err.clear
		Do While wshtempfile.AtEndOfStream <> true
			line = wshtempFile.ReadLine
			if instr(1, line, "Total Physical Memory", 1) then
				pos = instr(1, line, ":", 1)
				pos2 = instr(pos+2, line, " ", 1)
				size1 = cint(mid(line, pos+2, pos2-(pos+2)))
				size2 = cint(size1\32) + 1
				srvmemvar = size2 * 32
			end if
		Loop
		wshtempfile.Close
		wscript.DisconnectObject wshtempfile
		Set wshtempfile=nothing
	else
		screenout ""
		text = "Unable to create a text file in the temporary directory for the script!"
		screenout text
		screenout "Make sure that " & scripttempvar & " exists and is writable."
		screenout "Make sure that ntmem.exe in " & scriptfileslocvar & " is working properly."
		genevent "E", "3", text
		config_getinfo = 1
		exit function
	end if
	screenout "  Detected " & srvmemvar & " MB of memory."
	

	REM ###### Get Network factor ######
	if not scriptserverlocvar = "" then
		strCmdLine = "cmd /c ""ping "& scriptserverlocvar & " -w 3000 -n 1 >" & scripttempvar & "\pingtemp.txt"""
		temp = wshshell.Run(strcmdline, 0, true)
		if err.number <> 0 then err.clear
		if (Wshfile.fileexists(scripttempvar & "\pingtemp.txt")) then 
			Set wshtempfile = wshfile.OpenTextFile(scripttempvar & "\pingtemp.txt", 1)
			if err.number <> 0 then err.clear
			Do While wshtempfile.AtEndOfStream <> true
				line = wshtempFile.ReadLine
				if instr(1, line, "time=", 1) then
					pos = instr(1, line, "time=", 1)
					pos2 = instr(pos, line, "ms ", 1)
					rem screenout mid(line, pos + 5, pos2 - (pos + 5))
					netfactor = round(cint(mid(line, pos + 5, pos2 - (pos + 5)))/10)
					if netfactor <= 0 then netfactor = 1
					if netfactor >200 then netfactor = 200
				end if
			Loop
			wshtempfile.Close
			wscript.DisconnectObject wshtempfile
			Set wshtempfile=nothing
		else
			screenout ""
			text = "Unable to create a text file in the temporary directory for the script!"
			screenout text
			screenout "Make sure that " & scripttempvar & " exists and is writable."
			screenout "Make sure that ntmem.exe in " & scriptfileslocvar & " is working properly."
			genevent "E", "3", text
			config_getinfo = 1
			exit function
		end if
	end if
	screenout "  Detected network factor to be " & netfactor & "."

	REM ###### Get installed and extra services ######
	if (Wshfile.fileexists(scriptinilocvar & "\coreservicesw2k.ini")) then 
		Set wshtempfile = wshfile.OpenTextFile(scriptinilocvar & "\coreservicesw2k.ini", 1)
		alltext = wshtempFile.Readall
		wshtempfile.Close
		wscript.DisconnectObject wshtempfile
		Set wshtempfile=nothing
	else
		screenout ""
		text = "The script was unable to find the CoreServicesW2K.ini file!"
		screenout text
		screenout "Make sure that " & scriptinilocvar & " exists and contains the necessary INI file."
		genevent "E", "3", text
		config_getinfo = 1
		exit function
	end if
	allservices = ""
	extraservices = ""
	Set computerob = GetObject("WinNT://" & computername & ",computer" )
	computerob.filter = Array("Service")
	if err.number <> 0 then logerror "Test!", err.number : err.clear
	for each object in computerob
		rem screenout object.name & " (" & object.class & ")"
		if object.class = "Service" then
			allservices = allservices & object.Name & vbCrLf
			if Instr(1, alltext, object.name, 1) = 0 then
				extraservices = extraservices & object.Name & vbCrLf
			end if
		end if
	next
	if not extraservices = "" then screenout "  Detected extra services installed."

	REM ###### Get special services ######
	Terminst = WshShell.RegRead("HKLM\system\currentcontrolset\services\termservice\Start")
	Termchk = WshShell.RegRead("HKLM\system\currentcontrolset\control\Terminal Server\TSAppCompat")
	if err.number <> 0 then logerror "Reading UseLicenseServer Registry Key", err.number : err.clear
	if terminst = "2" and termchk = "0" then 
		screenout "  Detected Terminal Services service (RA)."
		termvar = "TermRa"
	elseif terminst = "2" and termchk = "1" then
		screenout "  Detected Terminal Services service (APP)."
		termvar = "TermApp"
	end if
	if Instr(1, allservices, "W3SVC", 1) <> 0 then
		screenout "  Detected WWW service."
		wwwvar = "Running"
	end if
	if Instr(1, allservices, "SQLServer", 1) <> 0 then
		screenout "  Detected SQL service."
		sqlvar = "Running"
	end if
	if Instr(1, allservices, "MSExchange", 1) <> 0 then
		screenout "  Detected Exchange service."
		exchangevar = "Running"
	end if
	sapchk = WshShell.RegRead("HKLM\software\microsoft\windows nt\currentversion\ServerType")
	if err.number <> 0 then err.clear
	if sapchk = "SAP" then 
		screenout "  Detected SAP service."
		sapvar = "Running"
	end if

	if err.number <> 0 then logerror "End of config_getinfo Function", err.number : err.clear
	config_getinfo = 0
End Function



REM #################################################################################################
REM ###                                Get switch information                                     ###
REM #################################################################################################
Function config_getswitches()
	dim arg, text, a
	on error resume next

	REM ###### Get text for all arguments ######
	argtext = wscript.ScriptFullName & " "
	for a = 0 to wscript.arguments.count - 1
		argtext = argtext & wscript.arguments(a) & " "
		if instr(1, wscript.arguments(a), "-", 1) <> 0 then
			scriptconswivar = wscript.arguments(a)
		elseif instr(1, okargtext, wscript.arguments(a), 1) = 0 then
			screenout ""
			text = "The switch '" & wscript.arguments(a) & "' is not a supported optional switch!"
			screenout text
			screenout "Please check the list of available switches by running the script with a '/?'."
			genevent "E", "3", text
			config_getswitches = 1
			exit function
		end if
		if wscript.arguments(a) = "path" or wscript.arguments(a) = "build" or wscript.arguments(a) = "bug" or wscript.arguments(a) = "sp" or wscript.arguments(a) = "/slownet" then a=a+1
	next

	REM ###### Check for switches and set variables ######
	if scriptconswivar = "" then 
		screenout ""
		text = "The configuration switch is either missing or has been entered improperly!"
		screenout text
		screenout "The configuration switch must contain a site description and a network description seperated by a dash."
		screenout "Run the script with a /? for more information."
		genevent "E", "3", text
		config_getswitches = 1
		exit function
	else
		screenout "  Switch for script configuration is: " & scriptconswivar
	end if

	if instr(1, scriptconswivar, "-corp", 1) then
		scriptnetworkvar = "CORPORATE"
	elseif instr(1, scriptconswivar, "-int", 1) then
		scriptnetworkvar = "INTERNET"
	elseif instr(1, scriptconswivar, "-ext", 1) then
		scriptnetworkvar = "EXTRANET"
	elseif instr(1, scriptconswivar, "-pri", 1) then
		scriptnetworkvar = "PRIVATE"
	else
		screenout ""
		text = "The network portion of the configuration switch has been entered improperly!"
		screenout text
		screenout "The networking portion must contain one of the following extensions: 'int', 'corp', 'ext', 'pri'"
		screenout "Run the script with a /? for more information."
		genevent "E", "3", text
		config_getswitches = 1
		exit function
	end if

	if instr(1, scriptconswivar, "/noam", 1) then
		scriptsitevar = "North America"
	elseif instr(1, scriptconswivar, "/soam", 1) then
		scriptsitevar = "South America"
	elseif instr(1, scriptconswivar, "/euro", 1) then
		scriptsitevar = "Europe"
	elseif instr(1, scriptconswivar, "/sopa", 1) then
		scriptsitevar = "South Pacific"
	elseif instr(1, scriptconswivar, "/faea", 1) then
		scriptsitevar = "Far East"
	elseif instr(1, scriptconswivar, "/miea", 1) then
		scriptsitevar = "Middle East"
	elseif instr(1, scriptconswivar, "/afca", 1) then
		scriptsitevar = "Africa"
	elseif instr(1, scriptconswivar, "/b11", 1) then
		scriptsitevar = "Building 11 Data Center"
	elseif instr(1, scriptconswivar, "/cp", 1) then
		scriptsitevar = "Canyon Park Data Center"
	elseif instr(1, scriptconswivar, "/tuk", 1) then
		scriptsitevar = "Tukwila Data Center"
	elseif instr(1, scriptconswivar, "/sat", 1) then
		scriptsitevar = "Saturn Lab"
	elseif instr(1, scriptconswivar, "/jup", 1) then
		scriptsitevar = "Jupiter Lab"
	elseif instr(1, scriptconswivar, "/soc", 1) then
		scriptsitevar = "MSN/SOC Servers"
	elseif instr(1, scriptconswivar, "/dsk", 1) then
		scriptsitevar = "DESK"
	else
		screenout ""
		text = "The site description portion of the configuration switch has been entered improperly!"
		screenout text
		screenout "The site description portion must contain a specific or a special site description."
		screenout "Run the script with a /? for more information."
		genevent "E", "3", text
		config_getswitches = 1
		exit function
	end if
	screenout "    Network type: " & scriptnetworkvar
	screenout "    Site location: " & scriptsitevar

	if argument("/com") <> 0 and srvhardwarevar = "COMPAQ" then
		screenout ""
		text = "The script has all ready detected the server to be Compaq hardware!"
		screenout text
		screenout "The /com switch should NOT be used unless the script fails to detect Compaq hardware."
		screenout "Re-Run the script without the '/com' switch."
		genevent "E", "3", text
		config_getswitches = 1
		exit function
	elseif argument("/com") <> 0 then
		screenout "  Switch for COMPAQ Hardware."
		srvhardwarevar = "COMPAQ"
	end if
	if argument("/del") <> 0 and srvhardwarevar = "DELL" then
		screenout ""
		text = "The script has all ready detected the server to be Dell hardware!"
		screenout text
		screenout "The /del switch should NOT be used unless the script fails to detect Dell hardware."
		screenout "Re-Run the script without the '/del' switch."
		genevent "E", "3", text
		config_getswitches = 1
		exit function
	elseif argument("/del") <> 0 then
		screenout "  Switch for DELL Hardware."
		srvhardwarevar = "DELL"
	end if
	if argument("/oth") <> 0 then
		screenout "  Switch for OTHER Hardware (No Compaq or Dell Components)."
		srvhardwarevar = "OTHER"
	end if
	if argument("/nocim") <> 0 then
		screenout "  Switch for NO CIM."
		cpqcimvar = "switch"
	end if


	if argument("/fullsym") <> 0 then
		screenout "  Switch for Full Symbol set install."
		symvar = "full"
	end if
	if argument("/clrsym") <> 0 then
		screenout "  Switch for clearing Symbols."
		symvar = "clear"
	end if
	if argument("/noupd") <> 0 then
		screenout "  Switch for No Update (Service Pack, Hotfixes, and Symbols disabled)."
		recconsvar = "switch"
		spvar = "switch"
		fixvar = "switch"
		symvar = "switch"
		encvar = "switch"
	end if


	if argument("/reb") <> 0 then
		screenout "  Switch for auto reboot at the end of the script."
		scriptrebootvar = "reboot"
	end if
	if argument("/noreb") <> 0 then
		screenout "  Switch for no reboot at the end of the script."
		scriptrebootvar = "noreboot"
	end if
	if argument("/debug") <> 0 then
		screenout "  Switch for full debugger settings."
		symvar = "full"
		srvdebug = 1
	end if
	if argument("/forceup") <> 0 then
		screenout "  Switch to force update of All up-to-date components."
		forceupdate = 1
	end if
	if argument("/slownet") <> 0 then
		arg = argument("/slownet")
		netfactor = nextargument(arg)
		if isnumeric(netfactor) = false or netfactor <= 0 then
			screenout ""
			text = "The '/slownet' switch requires a number to be used after the switch!"
			screenout text
			screenout "Re-run the script and include a number from 1 to 200 after the switch."
			screenout "  Example: /slownet 2"
			genevent "E", "3", text
			config_getswitches = 1
			exit function
		elseif netfactor > 200 then
			netfactor = 200
			screenout "  Switch to change network factor to " & netfactor & " (Limit of 200)."
		else
			screenout "  Switch to change network factor to " & netfactor & "."
		end if	
	end if


	if argument("/noopa") <> 0 then
		screenout "  Switch for NO OpAssist install."
		opavar = "switch"
	end if
	if argument("/nosen") <> 0 then
		screenout "  Switch for NO Sentry install."
		senvar = "switch"
	end if
	if argument("/noniq") <> 0 then
		screenout "  Switch for NO NetIQ install."
		niqvar = "switch"
	end if
	if argument("/noinoc") <> 0 then
		screenout "  Switch for NO Inoculan install."
		inocvar = "switch"
	end if
	if argument("/nooic") <> 0 then
		screenout "  Switch for NO Object Inactivity Checker install."
		oicvar = "switch"
	end if
	if argument("/nobac") <> 0 then
		screenout "  Switch for NO Backup Accelerator install."
		bacvar = "switch"
	end if
	if argument("/notools") <> 0 then
		screenout "  Switch for NO additional service changes."
		opavar = "switch"
		senvar = "switch"
		niqvar = "switch"
		inocvar = "switch"
		oicvar = "switch"
		bacvar = "switch"
	end if

	if argument("/wwwreg") <> 0 then
		screenout "  Switch for WWW related NT registry settings."
		if wwwvar = "" then wwwvar = "Switch"
	end if
	if argument("/sqlreg") <> 0 then
		screenout "  Switch for SQL related NT registry settings."
		if sqlvar = "" then sqlvar = "Switch"
	end if
	if argument("/excreg") <> 0 then
		screenout "  Switch for Exchange related NT registry settings."
		if exchangevar = "" then exchangevar = "Switch"
	end if
	if argument("/sapreg") <> 0 then
		screenout "  Switch for SAP related NT registry settings."
		if sapvar = "" then sapvar = "Switch"
	end if
	if argument("/wwwipak") <> 0 then
		screenout "  Switch for WWW IPAK execution."
		wwwvar = "IPAK"
	end if
	if argument("loctest") <> 0 then
		screenout "  Switch for WWW IPAK execution."
		locationtestvar = 1
	end if
	if argument("nontbuilds") <> 0 then
		screenout "  Switch for no Recovery Console or NTbuilds location."
		recconsvar = "switch"
		bldvar = "switch"
	end if
	if argument("notas") <> 0 then
		screenout "  Switch for no Tools and Services debug."
		opavar = "switch"
		senvar = "switch"
		niqvar = "switch"
		inocvar = "switch"
		oicvar = "switch"
		bacvar = "switch"
		wwwvar = ""
		sqlvar = ""
		exchangevar = ""
		sapvar = ""
	end if
	if argument("noots") <> 0 then
		screenout "  Switch for no OEM, Tools, or Services debug."
		cpqcimvar = "switch"
		cpqssdvar = "switch"
		delmannvar = "switch"
		delfastvar = "switch"
		deldrvvar = "switch"
		opavar = "switch"
		senvar = "switch"
		niqvar = "switch"
		inocvar = "switch"
		oicvar = "switch"
		bacvar = "switch"
		wwwvar = ""
		sqlvar = ""
		exchangevar = ""
		sapvar = ""
	end if
	if argument("noall") <> 0 then
		screenout "  Switch for NOTHING debug."
		spvar = "switch"
		fixvar = "switch"
		symvar = "switch"
		encvar = "switch"
		cpqcimvar = "switch"
		cpqssdvar = "switch"
		delmannvar = "switch"
		delfastvar = "switch"
		deldrvvar = "switch"
		opavar = "switch"
		senvar = "switch"
		niqvar = "switch"
		inocvar = "switch"
		oicvar = "switch"
		bacvar = "switch"
		wwwvar = ""
		sqlvar = ""
		exchangevar = ""
		sapvar = ""
	end if
	if argument("build") <> 0 then
		arg = argument("build")
		srvbuildvar = nextargument(arg)
		if srvbuildvar = "" then
			screenout ""
			text = "The build number needs to be included with the 'build' switch!"
			screenout text
			screenout "Please re-run the script and include a build number with the 'build' switch."
			screenout "  Example: build 2195"
			genevent "E", "3", text
			config_getswitches = 1
			exit function
		else
			screenout "  Switch to change source build location to " & srvbuildvar & "."
		end if	
	end if

	if argument("sp") <> 0 then
		arg = argument("sp")
		spdirvar = nextargument(arg)
		if spdirvar = "" then
			screenout ""
			text = "The SP number needs to be included with the 'sp' switch!"
			screenout text
			screenout "Please re-run the script and include a build number with the 'build' switch."
			screenout "  Example: sp SP1.059"
			genevent "E", "3", text
			config_getswitches = 1
			exit function
		else
			screenout "  Switch to change source SP location to " & spdirvar & "."
		end if	
	end if

	REM ###### Check for script being ran from CD ######
	if Instr(1, cddrives, scriptdrivelocvar, 1) <> 0 then 
		text = "This script is running from CD-ROM so some features will be disabled!"
		screenout text
		screenout "  Due to lack of space on the CD, the following tools and features have been disabled:"
		screenout "    - Installing the FULL symbol set"
		screenout "    - Running IPAK IIS5.01"
		screenout "  Since the CD is used only for servers on isolated networks this should not cause any problems."
		genevent "W", "4", text
		temp = MsgBox (text, 0, scriptnamevar & " - Tools and Features will be disabled!")
		cdrunvar = "on"
		symvar = "small"
		if wwwvar = "IPAK" then wwwvar = "Running"
	end if

	REM ###### Adjust settings based on Network Factor #######
	if netfactor >= 75 and netfactor < 150 then
		if symvar = "full" then
			screenouut "  Symbol install set to 'subset' due to net factor."
			symvar = "on"
		end if
	elseif netfactor >=150 then
		if symvar = "full" then
			screenouut "  Symbol install set to 'clear' due to net factor."
			symvar = "clear"
		end if
	end if

	if err.number <> 0 then logerror "End of config_getswitches Function", err.number : err.clear
	config_getswitches = 0
End Function



REM #################################################################################################
REM ###                                Set location variables                                     ###
REM #################################################################################################
Function config_setlocvars()
	dim text, color, oemcolor, toolcolor
	on error resume next

	REM ###### Get directory name ######
	if Instr(1, scriptpathlocvar, "bronze", 1) <> 0 then 
		color = "bronze"
		oemcolor = "silver"
		toolcolor = "silver"
	elseif Instr(1, scriptpathlocvar, "silver", 1) <> 0 then
		color = "silver"
		oemcolor = "silver"
		toolcolor = "silver"
	elseif Instr(1, scriptpathlocvar, "gold", 1) <> 0 then
		color = "gold"
		oemcolor = "gold"
		toolcolor = "gold"
	elseif Instr(1, scriptpathlocvar, "rust", 1) <> 0 then
		color = "rust"
		oemcolor = "rust"
		toolcolor = "rust"
	else 
		screenout ""
		text = "Unable to find the 'Color' directory in the source path to the script!"
		screenout text
		genevent "E", "3", text
		screenout "   (Gold, Silver, Bronze, or Rust)"
		screenout "Make sure that the source servers directory structure is created properly."
		screenout "Make sure that you are running the script properly."
		config_setlocvars = 1
		exit function
	end if


	REM #####################################################################################################
	REM #####################################################################################################
	REM ######                              Set location variables                                     ######
	REM #####################################################################################################
	REM #####################################################################################################
	scriptbinlocvar = scriptpathlocvar & "\" & srvbuildvar & "\" & srvsuitevar & "\" & srvarcvar & "\localbin"
	splocvar = scriptpathlocvar & "\" & srvbuildvar & "\" & srvsuitevar & "\" & srvarcvar & "\" & spdirvar
	fixlocvar = scriptpathlocvar & "\" & srvbuildvar & "\" & srvsuitevar & "\" & srvarcvar & "\" & spdirvar & "\hftrack"
	symlocvar = scriptpathlocvar & "\" & srvbuildvar & "\" & srvsuitevar & "\" & srvarcvar & "\" & spdirvar & "\symbolsubset"
	symfulllocvar = scriptpathlocvar & "\" & srvbuildvar & "\" & srvsuitevar & "\" & srvarcvar & "\" & spdirvar & "\symbols"
	wwwsymlocvar = scriptpathlocvar & "\" & srvbuildvar & "\" & srvsuitevar & "\" & srvarcvar & "\" & spdirvar & "\iissubset"
	bldlocvar = scriptsharelocvar & "\" & color & "\ntbuilds\win2k\" & srvbuildvar & "." & spdirvar & "\"& srvsuitevar
	recconslocvar = scriptsharelocvar & "\" & color & "\ntbuilds\win2k\" & srvbuildvar & "." & spdirvar & "\"& srvsuitevar& "\" & srvarcvar & "\winnt32.exe"
	enclocvar = scriptsharelocvar & "\" & color & "_dom\ipak\" & left(scriptipakvar, 6) & "\w2kencrypt.vbs"
	wwwlocvar = scriptpathlocvar & "\..\iis5.02"

	cpqssdlocvar = scriptdrivelocvar & "\" & oemcolor & "\compaq\win2000\503a"
	cpqcimlocvar = scriptdrivelocvar & "\" & oemcolor & "\compaq\win2000\503a"
	cpqstrdrvlocvar = scriptdrivelocvar & "\" & oemcolor & "\compaq\storageworks\win2000\hsg80_v85b_x86nt\kgpsa\win2k\"
	delmannlocvar = scriptdrivelocvar & "\" & oemcolor & "\dell\win2000\management\hpov1.52\hip3.5.2"
	delfastlocvar = scriptdrivelocvar & "\" & oemcolor & "\dell\win2000\management\perc2\fast_2972"
	deldrvlocvar = scriptdrivelocvar & "\" & oemcolor & "\dell\win2000\drivers\array_cntrl\perc2\rev_2963"

	oiclocvar = scriptfileslocvar & "\oic2\oicinstall.bat"
	opalocvar = scriptdrivelocvar & "\" & toolcolor & "\tools\opassist\opinst-v4.cmd"
	senlocvar = scriptdrivelocvar & "\" & toolcolor & "\tools\sentry\install5.02.bat"
	niqlocvar = scriptdrivelocvar & "\" & toolcolor & "\tools\netiq\install.bat"
	inoclocvar = scriptdrivelocvar & "\" & toolcolor & "\inoculan\inocinst.cmd"
	baclocvar = scriptdrivelocvar & "\" & toolcolor & "\veritas\be8.0\winnt\install\eng\" & srvarcvar & "\ntaa\setupaa.cmd"

	REM ### Check for localbin source location ###
	if (Wshfile.FolderExists(scriptbinlocvar)) then
		screenout "Localbin files location variable is OK."
		if debug = 1 then screenout "  (" & scriptbinlocvar & ")"
	else
		screenout ""
		text = "Could not locate the reskit files directory for the script!"
		screenout text
		screenout "  (" & scriptbinlocvar & ")"
		genevent "E", "3", text
		config_setlocvars = 1
		exit function
	end if

	REM ###### Check for needed executables ######
	if not (Wshfile.FileExists(scriptbinlocvar & "\regback.exe")) then
		text = "Could not find Regback.exe in Localbin files location!"
		screenout text
		genevent "E", "3", text
		config_setlocvars = 1
		exit function
	elseif not (Wshfile.FileExists(scriptbinlocvar & "\logevent.exe")) then
		text =  "Could not find logevent.exe in Localbin files location!"
		screenout text
		genevent "E", "3", text
		config_setlocvars = 1
		exit function
	end if

	REM ### Check for service pack location ###
	if spvar = "off" or spvar = "switch" or spdirvar = "SP0" and locationtestvar = 0 then
	elseif (Wshfile.FolderExists(splocvar)) then
		screenout "Service Pack location variable is OK."
		if debug = 1 then screenout "  (" & splocvar & ")"
	else
		screenout ""
		text = "Could not locate the Service Pack directory for the script!"
		screenout text
		screenout "  (" & splocvar & ")"
		genevent "E", "3", text
		config_setlocvars = 1
		exit function
	end if
	
	REM ### Check for hotfix location ###
	if fixvar = "off" or fixvar = "switch" and locationtestvar = 0 then
	elseif (Wshfile.FolderExists(fixlocvar)) then
		screenout "Hotfix location variable is OK."
		if debug = 1 then screenout "  (" & fixlocvar & ")"
	else
		screenout ""
		text = "Could not locate the hotfix directory for the script!"
		screenout text
		screenout "  (" & fixlocvar & ")"
		genevent "E", "3", text
		config_setlocvars = 1
		exit function
	end if
	
	REM ### Check for symbols location ###
	if symvar = "off"  or symvar = "switch" and locationtestvar = 0 then
	elseif (Wshfile.FolderExists(symlocvar)) then
		screenout "Symbol subset location variable is OK."
		if debug = 1 then screenout "  (" & symlocvar & ")"
	else
		screenout ""
		text = "Could not locate the symbol subset directory for the script!"
		screenout text
		screenout "  (" & symlocvar & ")"
		genevent "E", "3", text
		config_setlocvars = 1
		exit function
	end if

	REM ### Check for iis symbols location ###
	if symvar = "off"  or symvar = "switch" or wwwvar = "" and locationtestvar = 0 then
	elseif (Wshfile.FolderExists(wwwsymlocvar)) then
		screenout "IIS Symbols location variable is OK."
		if debug = 1 then screenout "  (" & wwwsymlocvar & ")"
	else
		screenout ""
		text = "Could not locate the IIS symbols directory for the script!"
		screenout text
		screenout "  (" & wwwsymlocvar & ")"
		genevent "E", "3", text
		config_setlocvars = 1
		exit function
	end if

	REM ### Check for full symbols location ###
	if not symvar = "full" and locationtestvar = 0 then
	elseif (Wshfile.FolderExists(symfulllocvar)) then
		screenout "Full Symbols location variable is OK."
		if debug = 1 then screenout "  (" & symfulllocvar & ")"
	else
		screenout ""
		screenout "Could not locate the Full symbols source...switching symbol install to small symbol set."
		screenout "  (" & symfulllocvar & ")"
		symvar = srvbuildvar
	end if

	REM ### Check for build tree location ###
	if (bldvar = "off" or bldvar = "switch" or cdrunvar = "on") and locationtestvar = 0 then
	elseif (Wshfile.FolderExists(bldlocvar)) then
		screenout "Build tree location variable is OK."
		if debug = 1 then screenout "  (" & bldlocvar & ")"
	else
		screenout ""
		text = "Could not locate the Build tree directory for the script!"
		screenout text
		screenout "  (" & bldlocvar & ")"
		genevent "E", "3", text
		config_setlocvars = 1
		exit function
	end if

	REM ### Check for Recovery Console location ###
	if (recconsvar = "off" or not srvbootvar = "NTFS" or recconsvar = "switch" or cdrunvar = "on") and locationtestvar = 0 then
	elseif (Wshfile.FileExists(recconslocvar)) then
		screenout "Recovery Console location variable is OK."
		if debug = 1 then screenout "  (" & recconslocvar & ")"
	else
		screenout ""
		text = "Could not locate the Recovery Console directory for the script!"
		screenout text
		screenout "  (" & recconslocvar & ")"
		genevent "E", "3", text
		config_setlocvars = 1
		exit function
	end if

	REM ### Check for 128 Bit encryption location ###
	if encvar = "off" or symvar = "switch" and locationtestvar = 0 then
	elseif (Wshfile.FileExists(enclocvar)) then
		screenout "128 Bit encyption location variable is OK."
		if debug = 1 then screenout "  (" & enclocvar & ")"
	else
		screenout ""
		text = "Could not locate the 128 Bit encryption directory for the script!"
		screenout text
		screenout "  (" & enclocvar & ")"
		genevent "E", "3", text
		config_setlocvars = 1
		exit function
	end if
	
	REM ### Check for ssd location ###
	if (not srvhardwarevar = "COMPAQ" or cpqssdvar = "off" or cpqssdvar = "switch") and locationtestvar = 0 then
	elseif (Wshfile.FolderExists(cpqssdlocvar)) then
		screenout "SSD location variable is OK."
		if debug = 1 then screenout "  (" & cpqssdlocvar & ")"
	else
		screenout ""
		text = "Could not locate the SSD directory for the script!"
		screenout text
		screenout "  (" & cpqssdlocvar & ")"
		genevent "E", "3", text
		config_setlocvars = 1
		exit function
	end if
	
	REM ### Check for CIM location ###
	if (not srvhardwarevar = "COMPAQ" or cpqcimvar = "off" or cpqcimvar = "switch") and locationtestvar = 0 then
	elseif (Wshfile.FolderExists(cpqcimlocvar)) then
		screenout "CIM location variable is OK."
		if debug = 1 then screenout "  (" & cpqcimlocvar & ")"
	else
		screenout ""
		text = "Could not locate the CIM directory for the script!"
		screenout text
		screenout "  (" & cpqcimlocvar & ")"
		genevent "E", "3", text
		config_setlocvars = 1
		exit function
	end if
	
	REM ### Check for Storage Works Driver location ###
	if (srvstrwrksvar = "DISABLED" or srvstrwrksvar = "ABSENT") and locationtestvar = 0 then
	elseif (Wshfile.FolderExists(cpqstrdrvlocvar)) then
		screenout "Storage Works Driver location variable is OK."
		if debug = 1 then screenout "  (" & cpqstrdrvlocvar & ")"
	else
		screenout ""
		text = "Could not locate the Storage Works Driver directory for the script!"
		screenout text
		screenout "  (" & cpqstrdrvlocvar & ")"
		genevent "E", "3", text
		config_setlocvars = 1
		exit function
	end if

	REM ### Check for DMN location ###
	if (not srvhardwarevar = "DELL" or delmannvar = "off" or delmannvar = "switch") and locationtestvar = 0  then
	elseif (Wshfile.FolderExists(delmannlocvar)) then
		screenout "DMN location variable is OK."
		if debug = 1 then screenout "  (" & delmannlocvar & ")"
	else
		screenout ""
		text = "Could not locate the DMN directory for the script!"
		screenout text
		screenout "  (" & delmannlocvar & ")"
		genevent "E", "3", text
		config_setlocvars = 1
		exit function
	end if

	REM ### Check for DFAST location ###
	if (not srvhardwarevar = "DELL" or delfastvar = "off" or delfastvar = "switch") and locationtestvar = 0  then
	elseif (Wshfile.FolderExists(delfastlocvar)) then
		screenout "DFAST location variable is OK."
		if debug = 1 then screenout "  (" & delfastlocvar & ")"
	else
		screenout ""
		text = "Could not locate the DFAST directory for the script!"
		screenout text
		screenout "  (" & delfastlocvar & ")"
		genevent "E", "3", text
		config_setlocvars = 1
		exit function
	end if

	REM ### Check for Perc2 driver location ###
	if (not srvhardwarevar = "DELL" or deldrvvar = "off" or deldrvvar = "switch") and locationtestvar = 0  then
	elseif (Wshfile.FolderExists(deldrvlocvar)) then
		screenout "Perc2 driver location variable is OK."
		if debug = 1 then screenout "  (" & deldrvlocvar & ")"
	else
		screenout ""
		text = "Could not locate the Perc2 driver directory for the script!"
		screenout text
		screenout "  (" & deldrvlocvar & ")"
		genevent "E", "3", text
		config_setlocvars = 1
		exit function
	end if

	if debug = 1 then screenout "Opassist location."
	if debug = 1 then screenout "  (" & opalocvar & ")"
	if debug = 1 then screenout "Sentry location."
	if debug = 1 then screenout "  (" & senlocvar & ")"
	if debug = 1 then screenout "NetIQ location."
	if debug = 1 then screenout "  (" & niqlocvar & ")"
	if debug = 1 then screenout "Inoculan location."
	if debug = 1 then screenout "  (" & inoclocvar & ")"
	if debug = 1 then screenout "OIC location."
	if debug = 1 then screenout "  (" & oiclocvar & ")"
	if debug = 1 then screenout "Backup Accelerator location."
	if debug = 1 then screenout "  (" & baclocvar & ")"

	REM ### Check for www location ###
	if (wwwvar = "" or wwwvar = "Running") and locationtestvar = 0 then
	elseif (Wshfile.FolderExists(wwwlocvar)) then
		screenout "WWW IPAK script location variable is OK."
		if debug = 1 then screenout "  (" & wwwlocvar & ")"
	else
		screenout ""
		text = "Could not locate the WWW IPAK script directory for the script!"
		screenout text
		screenout "  (" & wwwlocvar & ")"
		genevent "E", "3", text
		config_setlocvars = 1
		exit function
	end if

	if locationtestvar = 1 then config_setlocvars = 1: exit function
	if err.number <> 0 then logerror "End of config_setlocvars Function", err.number : err.clear
	config_setlocvars = 0
End Function





REM #################################################################################################
REM ###                                     Bug Function                                          ###
REM #################################################################################################
Function bug()
	dim a, location, text
	on error resume next

	screenout ""
	screenout "SCRIPT DEBUGGER"
	if argument("hardware_cpqnic") then hardware_cpqnic
	if argument("check_sysbios") then check_sysbios
	if argument("check_delldriver") then check_delldriver
	if argument("admin_asset") then admin_asset
	if argument("admin_prework") then admin_prework
	if argument("admin_groups") then admin_groups
	if argument("admin_audit") then admin_audit
	if argument("services_schedule") then services_schedule
	if argument("services_time") then services_time
	if argument("services_snmp") then services_snmp
	if argument("services_msmq") then services_msmq
	if argument("hardware_ssd") then hardware_ssd
	if argument("hardware_cim") then hardware_cim
	if argument("hardware_storageworksdriver") then hardware_storageworksdriver
	if argument("hardware_managednode") then hardware_managednode("false")
	if argument("hardware_perc2fast") then hardware_perc2fast("false")
	if argument("update_reccons") then getbuildcd: update_reccons: getupdatecd
	if argument("update_sp") then update_sp("false")
	if argument("update_sym") then update_sym
	if argument("update_fix") then update_fix
	if argument("update_128") then update_128
	if argument("file_localbin") then file_localbin
	if argument("file_exchange") then file_exchange
	if argument("file_boot") then file_boot
	if argument("file_delete") then file_delete
	if argument("file_bitmap") then file_bitmap
	if argument("registry_main") then registry_main
	if argument("registry_filters") then registry_filters
	if argument("registry_pagefile") then registry_pagefile
	if argument("registry_path") then registry_path
	if argument("registry_source") then registry_source
	if argument("registry_backup") then registry_backup
	if argument("registry_diskperf") then registry_diskperf
	if argument("registry_site") then registry_site
	if argument("registry_debug") then registry_debug
	if argument("registry_services") then registry_services
	if argument("tools_perfcol") then tools_perfcol
	if argument("tools_opassist") then tools_opassist
	if argument("tools_sentry") then tools_sentry
	if argument("tools_netiq") then tools_netiq
	if argument("tools_inoculan") then tools_inoculan
	if argument("tools_oic") then tools_oic
	if argument("tools_backupexec") then tools_backupexec
	if argument("tools_other") then tools_other
	if argument("services_www") then services_www
	if argument("completion_other") then completion_other
	if argument("completion_locallogging") then completion_locallogging
	if argument("completion_destlogging") then completion_destlogging
	if argument("completion_reboot") then completion_reboot
	genevent "I", "3", "An undocumented script debugging command was used to only run a portion of the script"

End Function





REM #################################################################################################
REM ###                                  Check System Bios                                        ###
REM #################################################################################################
Function check_sysbios()
	dim temp, strcmdline, text, biosdate, endscript
	on error resume next

	REM ### Get firmware date ###
	screenout "System BIOS Check:"
	temp = WshShell.Regread("HKLM\hardware\description\system\systembiosdate")
	if err.number = 0 then
		biosdate = datevalue(temp)
	else
		err.clear
		biosdate = "no"
	end if
	screenout "  Current System BIOS date - " & biosdate
	cpqbiosdate = datevalue(cpqbiosdate)
	delbiosdate = datevalue(delbiosdate)

	REM ### Check for biosdate location ###
	if srvhardwarevar = "COMPAQ" and srvhardtypevar = "UNKNOWN" and datecomp(cpqbiosdate, biosdate) then
		screenout "  The Compaq System BIOS on this unknown server model may be OK."
	elseif srvhardwarevar = "COMPAQ" and (srvhardtypevar = "DL380" or srvhardtypevar = "DL360") and datecomp("6/2/2000", biosdate) then
		screenout "  The Compaq System BIOS on this " & srvhardtypevar & " is OK."
	elseif srvhardwarevar = "COMPAQ" and srvhardtypevar = "850R" and datecomp("12/9/1999", biosdate) then
		screenout "  The Compaq System BIOS on this " & srvhardtypevar & " is OK."
	elseif srvhardwarevar = "COMPAQ" and (srvhardtypevar = "5500" or srvhardtypevar = "6400R" or srvhardtypevar = "6500") and datecomp("12/8/1999", biosdate) then
		screenout "  The Compaq System BIOS on this " & srvhardtypevar & " is OK."
	elseif srvhardwarevar = "COMPAQ" and (srvhardtypevar = "1850R" or srvhardtypevar = "1600" or srvhardtypevar = "3000" or srvhardtypevar = "5500" or srvhardtypevar = "6000" or srvhardtypevar = "7000" or srvhardtypevar = "8000"or srvhardtypevar = "8500") and datecomp("12/7/1999", biosdate) then
		screenout "  The Compaq System BIOS on this " & srvhardtypevar & " is OK."
	elseif srvhardwarevar = "COMPAQ" and (srvhardtypevar = "1600" or srvhardtypevar = "2500") and datecomp("6/28/1999", biosdate) then
		screenout "  The Compaq System BIOS on this " & srvhardtypevar & " is OK."
	elseif srvhardwarevar = "COMPAQ" and (srvhardtypevar = "5000" or srvhardtypevar = "6000" or srvhardtypevar = "6500" or srvhardtypevar = "7000") and datecomp("4/30/1999", biosdate) then
		screenout "  The Compaq System BIOS on this " & srvhardtypevar & " is OK."
	elseif srvhardwarevar = "COMPAQ" and srvhardtypevar = "4500" and datecomp("8/17/1998", biosdate) then
		screenout "  The Compaq System BIOS on this " & srvhardtypevar & " is OK."
	elseif srvhardwarevar = "COMPAQ" and datecomp(cpqbiosdate, biosdate) then
		screenout "  The Compaq System BIOS on this " & srvhardtypevar & " is OK."
	elseif srvhardwarevar = "COMPAQ" then
		screenout ""
		text = "The Compaq System BIOS on this server needs to be updated before the update script can continue!"
		screenout text
		screenout "The known minimum System BIOS dates that are required for this IPAK are:"
		screenout "  Proliant DL380, DL360 - 6/2/2000"
		screenout "  Proliant 850R (P04) - 12/9/1999"
		screenout "  Proliant 5500 Xeon (P12), 6400R (E25), and 6500R Xeon (P11) - 12/8/1999"
		screenout "  Proliant 1850R (P07), 1600 (P08), 3000 (P09/E39), 5500 (E39) - 12/7/1999"
		screenout "  Proliant 6000 Xeon (P40/P43), 7000 Xeon (P40/P43), 8000 (P41), and 8500 (P42) - 12/7/1999"
		screenout "  Proliant 1600 (E34) and 2500 (E25/E39) - 6/28/1999"
		screenout "  Proliant 5000 (E16), 6000 (E20), 6500 (E25), and 7000 (E40) - 4/30/1999"
		screenout "  Proliant 4500 (E14) - 8/17/1998"
		screenout "The Rompaq's for the system bios can be found at:"
		screenout  "  " & scriptsharelocvar & "\gold\compaq\roms\system"
		screenout ""
		genevent "E", "3", text
		check_sysbios = 1
		exit function
	elseif srvhardwarevar = "DELL" and srvhardtypevar = "UNKNOWN" and datecomp(delbiosdate, biosdate) then
		screenout "  The Dell System BIOS on this unknown server model may be OK."
	elseif srvhardwarevar = "DELL" and (srvhardtypevar = "6300" or srvhardtypevar = "6350") and datecomp("10/27/99", biosdate) then
		screenout "  The Dell System BIOS on this " & srvhardtypevar & " is OK."
	elseif srvhardwarevar = "DELL" and srvhardtypevar = "2450" and datecomp("12/21/99", biosdate) then
		screenout "  The Dell System BIOS on this " & srvhardtypevar & " is OK."
	elseif srvhardwarevar = "DELL" and datecomp(delbiosdate, biosdate) then
		screenout "  The Dell System BIOS on this " & srvhardtypevar & " is OK."
	elseif srvhardwarevar = "DELL" then
		screenout ""
		text = "The Dell System BIOS on this server needs to be updated before the update script can continue!"
		screenout text
		screenout "The known minimum System BIOS dates that are required for this IPAK are:"
		screenout "   Dell Poweredge 6300 and 6350 - 10/27/99 (A11)"
		screenout "   Dell Poweredge 2450 - 12/21/99 (A02)"
		screenout "The updates for the System BIOS can be found at:"
		screenout  "  " & scriptsharelocvar & "\gold\dell\roms\system\6300\bios_A11"
		screenout ""
		genevent "E", "3", text
		check_sysbios = 1
		exit function
	else
		screenout "  The System BIOS on this platform does not need checked."
	end if

	if err.number <> 0 then logerror "End of check_sysbios Function", err.number : err.clear
	check_sysbios = 0
End Function





REM #################################################################################################
REM ###                                     Dell Perc2 Driver                                     ###
REM #################################################################################################
Function check_delldriver()
	dim temp, strcmdline, text, wrongfile, filethere, servfile
	on error resume next


	REM ### Check for driver dates ###
	screenout "Dell PERC2 Driver (2.1 - 2963):"
	if not srvhardwarevar = "DELL" then
		screenout "  The Dell Perc2 driver does not exist on this hardware platform."
		deldrvlogvar = "Different Hardware Platform."
	else
		REM ### Get perc2 driver date ###
		if (Wshfile.fileexists(deldrvchkvar)) then
			set servfile = Wshfile.getfile(deldrvchkvar)
			filethere = datevalue(left(servfile.datelastmodified, Instr(1, servfile.datelastmodified, " ", 1)-1))
			wscript.DisconnectObject servfile
			set servfile=nothing
		else
			filethere = "no"
		end if
		if (Wshfile.fileexists(systemroot & "\system32\drivers\afascsi.sys")) then filethere = "Upgrade"
		deldrvfilevar = datevalue(deldrvfilevar)

		if debug = 1 then screenout "  (Perc2 File: " & filethere & ")"

		if datecomp(deldrvfilevar, filethere) then
			screenout "  Dell Perc2 driver (" & deldrvvar & ") OK."
			deldrvlogvar = "OK"
		else
			screenout ""
			text = "The Dell Perc2 Driver needs to be updated before the update script can continue!"
			screenout text
			screenout "The Dell Perc2 Driver needs to be running version " & deldrvvar & "." 
			screenout "This requires user intervention! Please follow these steps:"
			screenout "  In 'Device Manager' of 'Computer Manager'"
			screenout "    - Open the 'SCSI and RAID Controllers' Devices"
			screenout "  On the 'Dell PERC 2 RAID Controller' device"
			screenout "    - Right click and select 'Properties'"
			screenout "  In the 'Dell PERC 2 RAID Controller Properties' window"
			screenout "    - Select the 'Driver' tab and click on the 'Update Driver' button"
			screenout "  In the 'Upgrade Device Driver Wizard' window"
			screenout "    - Press the 'Next' button"
			screenout "    - Make sure the 'Display a list of known drivers for this device"
			screenout "      so that I can choose a specific driver' radial button is checked"
			screenout "    - Press the 'Next' button"
			screenout "    - Press the 'Have Disk' button"
			screenout "  In the 'Install from Disk' window"
			screenout "    - Press the 'Browse' button"
			screenout "  In the 'Locate File' window"
			screenout "    - Open up the 'Look in' box and select the " & scriptdrivelocvar & " drive"
			screenout "    - Move into the " & deldrvlocvar & " directory"
			screenout "    - Highlight the 'oemsetup.inf' file"
			screenout "    - Press the 'Open' button"
			screenout "  In the 'Install from Disk' window"
			screenout "    - Press the 'OK' button"
			screenout "  In the 'Upgrade Device Driver Wizard' window"
			screenout "    - Highlight the 'Dell PERC 2 RAID Controller' Model and press the 'Next' button"
			screenout "    - Press the 'Next' button"
			screenout "  In the 'Digital Signature Not Found' window"
			screenout "    - Press the 'Yes' button "
			screenout " In the 'Upgrade Device Driver Wizard' window"
			screenout "    - Press the 'Finish' button"
			screenout "    - Close out of the properties window"
			screenout "  In the 'System Settings Change' window"
			screenout "    - Press the 'No' button"
			screenout "Repeat the above steps for ALL PERC2 devices!"
			screenout "Check 'Other Devices' in 'Device Manager' for any unknown Dell devices."
			screenout "  (Such as 'Dell 1x8 U2W SCSI BP SCSI Processor Device')"
			screenout "  If a device exists, follow the above procedure on that device except"
			screenout "  use the 'OEMSETUP2.inf' file. There will not be a Digital Signature pop up."
			screenout "  If the above procedure fails, use the 'Search for a suitable driver...' option."
			screenout "When the server is down, check the PERC2 Firmware and upgrade it if needed"
			screenout "  The PERC2 firmware should be running Version 2.1 Build 2939 or newer!"
			screenout "  Make sure the PERC2 driver and firmware are not mis-matched versions!"
			screenout "  It is critical that both PERC2 driver and firmware are Version 2.1!"
			screenout ""
			genevent "E", "3", text
			check_delldriver = 1
			strcmdline = "cmd /c ""compmgmt.msc /s"""	
			temp = wshshell.Run(strcmdline, 0, true)
			exit function
		end if
	end if

	if err.number <> 0 then logerror "End of check_delldriver Function", err.number : err.clear
	check_delldriver = 0
End Function




REM #################################################################################################
REM ###                                           Asset Tag                                       ###
REM #################################################################################################
Function admin_asset()
	dim temp, strcmdline, text, entriesok, count, asset, serial, location
	on error resume next

	REM ### Check Asset Tag Stuff ###
	screenout "Asset Tagger:"
	if scriptsitevar = "DESK" then
		screenout "  Asset Tag Information is not needed on a Desktop configurations."
	else
		entriesok = 0
		do
			asset = WshShell.RegRead("HKLM\system\currentcontrolset\services\snmp\parameters\rfc1156agent\syscontact")
			if err.number <> 0 then err.clear
			serial = WshShell.RegRead("HKLM\system\currentcontrolset\services\snmp\parameters\rfc1156agent\sysserialnumber")
			if err.number <> 0 then err.clear
			location = WshShell.RegRead("HKLM\system\currentcontrolset\services\snmp\parameters\rfc1156agent\syslocation")
			if err.number <> 0 then err.clear
			if count > 20 then
				screenout ""
				text = "  Entries for the Asset Tagger Utility have failed more then 20 times!"
				screenout text
				screenout "    Please check the script documentation for the appropriate entries and re-run the script."
				genevent "E", "3", text
				entriesok = 1
				admin_asset = 1
				exit function
			elseif asset = "" and serial = "" and location = "" then
				screenout "  All entries in the Asset Tagger Utility need to be filled out..."
				screenout "    Example:"
				screenout "            321561"
				screenout "            D740BPV10025"
				screenout "            11/S25-W12 (3)"
				strcmdline = scriptfileslocvar & "\assettag.exe"
				temp = wshshell.Run(strcmdline, 1, true)
			elseif asset = "" then
				screenout "  The 'Asset Tag' entry in the Asset Tagger Utility needs to be filled out..."
				screenout "    The Asset Tag number should contain the number on the Microsoft Asset sticker on the computer."
				screenout "    Examples:"
				screenout "             321561"
				screenout "             MS0011111 "
				screenout "             L321561"
				screenout "    If the server is not owned by Microsoft you can enter '000000' into the Asset Tag # field."
				strcmdline = scriptfileslocvar & "\assettag.exe"
				temp = wshshell.Run(strcmdline, 1, true)
			elseif serial = "" then
				screenout "  The 'Serial Number' entry in the Asset Tagger Utility needs to be filled out..."
				screenout "    The Serial number should contain the serial number marked on the computer."
				screenout "    Example:"
				screenout "            D740BPV10025"
				strcmdline = scriptfileslocvar & "\assettag.exe"
				temp = wshshell.Run(strcmdline, 1, true)
			elseif location = "" then
				screenout "  The 'Location' entry in the Asset Tagger Utility needs to be filled out..."
				screenout "    The Location entry should contain a building abreviation of no less then 2 characters,"
				screenout "    followed by a forward slash, followed by a rack location of no less then 4 characters."
				screenout "    Additional information should be enclosed in parenthesis (such as concentrator position)."
				screenout "    Examples:"
				screenout "             11/S25-W12 (6,1)"
				screenout "             CP-E/W01N36 (3)"
				screenout "             CPL/2162"
				screenout "    If this server is a test machine you can enter 'Test' for the Location field."
				strcmdline = scriptfileslocvar & "\assettag.exe"
				temp = wshshell.Run(strcmdline, 1, true)
			elseif not location = "Test" and Instr(1, location, "/", 1) = 0  then
				screenout "  The 'Location' entry in the Asset Tagger Utility needs to contain a forward slash seperator..."
				screenout "    The Location entry should contain a building abreviation of no less then 2 characters,"
				screenout "    followed by a forward slash, followed by a rack location of no less then 4 characters."
				screenout "    Additional information should be enclosed in parenthesis (such as concentrator position)."
				screenout "    Examples:"
				screenout "             11/S25-W12 (6,1)"
				screenout "             CP-E/W01N36 (3)"
				screenout "             CPL/2162"
				screenout "    If this server is a test machine you can enter 'Test' for the Location field."
				strcmdline = scriptfileslocvar & "\assettag.exe"
				temp = wshshell.Run(strcmdline, 1, true)
			elseif count = 0 then
				screenout "  Check Asset Tagger Utility entries for changes..."
				strcmdline = scriptfileslocvar & "\assettag.exe"
				temp = wshshell.Run(strcmdline, 1, false)
				entriesok = 1
				exit do
			else
				screenout "  Asset Tagger Utility entries are OK."
				entriesok = 1
				exit do
			end if
			count = count + 1
		loop while entriesok = 0
	end if

	if err.number <> 0 then logerror "End of admin_asset Function", err.number : err.clear
	admin_asset = 0
End Function





REM #################################################################################################
REM ###                                         Pre Work                                          ###
REM #################################################################################################
Function admin_prework()
	dim strcmdline, temp, text
	on error resume next

	REM ###### Stop File replication Services ######
	screenout "Stopping Windows 2000 File Replication Service..."
	strcmdline = "cmd /c ""net stop ntfrs"""	
	temp = wshshell.Run(strcmdline, 0, true)
	if err.number <> 0 then err.clear

	if err.number <> 0 then logerror "End of admin_prework Function", err.number : err.clear
	admin_prework = 0
End Function





REM #################################################################################################
REM ###                                        Groups                                             ###
REM #################################################################################################
Function admin_groups()
	dim strcmdline, temp, text
	on error resume next

	REM ###### Check for admin access ######
	if scriptsitevar = "DESK" then
		screenout "Administration Group manipulation not needed on DESK top configurations."
	elseif srvrolevar = "Server" then 
		screenout "Manipulating Groups for administration..."
		strcmdline = "cmd /c ""net localgroup administrators Redmond\itg-admin /add"""	
		temp = wshshell.Run(strcmdline, 0, true)
		strcmdline = "cmd /c ""net localgroup administrators Houston\itg-admin /add"""	
		temp = wshshell.Run(strcmdline, 0, true)
		strcmdline = "cmd /c ""net localgroup ""Backup Operators"" ""REDMOND\itg-backup operators"" /add"""	
		temp = wshshell.Run(strcmdline, 0, true)
		strcmdline = "cmd /c ""net localgroup ""Backup Operators"" ""HOUSTON\BUAdmin"" /add"""	
		temp = wshshell.Run(strcmdline, 0, true)
	else
		screenout "Administration Group manipulation not supported on Domain Controllers."
	end if

	if err.number <> 0 then logerror "End of admin_groups Function", err.number : err.clear
	admin_groups = 0
End Function





REM #################################################################################################
REM ###                                      Auditing                                             ###
REM #################################################################################################
Function admin_audit()
	dim strcmdline, temp, text
	on error resume next

	REM ###### Check for admin access ######
	if scriptsitevar = "DESK" then
		screenout "Administration Audit manipulation not needed on DESK top configurations."
	elseif srvrolevar = "Server" then 
		screenout "Manipulating Auditing for administration..."
		strcmdLine = "cmd /c """& scriptfileslocvar & "\auditpol.exe /logon:all /sam:all /policy:all /account:all """
		temp = wshshell.Run(strcmdline, 0, true)
	else
		screenout "Administration Audit manipulation not supported on Domain Controllers."
	end if

	if err.number <> 0 then logerror "End of admin_audit Function", err.number : err.clear
	admin_audit = 0
End Function



REM #################################################################################################
REM ###                                  Schedule service                                         ###
REM #################################################################################################
Function services_schedule()
	dim strcmdline, temp, alltext, wshtempfile, service, text
	on error resume next

	REM ###### Scheduler service ######
	strcmdline = "cmd /c ""net start > " & scripttempvar & "\servicestemp.txt"""	
	temp = wshshell.Run(strcmdline, 0, true)
	if (Wshfile.fileexists(scripttempvar & "\servicestemp.txt")) then 
		Set wshtempfile = wshfile.OpenTextFile(scripttempvar & "\servicestemp.txt", 1)
		alltext = wshtempfile.readall
		wshtempfile.Close
		wscript.DisconnectObject wshtempfile
		Set wshtempfile=nothing
		if instr(1, alltext, "Task Scheduler", 1) then
			screenout "Task Scheduler OK."
		else
			screenout "Task Scheduler service being configured and started..."
			wshshell.regwrite "HKLM\system\currentcontrolset\services\schedule\start", 2, "REG_DWORD"
			logregistryentry "Scheduler Service Automatic Start", "Corrected", "HKLM\system\currentcontrolset\services\schedule\start", "2"
			if err.number <> 0 then err.clear
			wshshell.regwrite "HKLM\system\currentcontrolset\services\schedule\objectname", "LocalSystem", "REG_SZ"
			logregistryentry "Scheduler Service Account", "Corrected", "HKLM\system\currentcontrolset\services\schedule\objectname", "LocalSystem"
			if err.number <> 0 then err.clear
			temp = wshshell.Run("cmd /c ""net start schedule""", 0, true)
		end if
	else
		screenout "Unable to open services text file!"
	end if

	if err.number <> 0 then logerror "End of services_schedule Function", err.number : err.clear
	services_schedule = 0
End Function



REM #################################################################################################
REM ###                                  Time service                                             ###
REM #################################################################################################
Function services_time()
	dim strcmdline, temp, text, pos1, pos2, line, wshtempfile, scriptinfravar, ipconfigtext, nettimetext, sntplist
	dim attext, atlist, number, hour, minute
	on error resume next

	REM ###### Find what operating system infrastructure the server is in ######
	screenout "Time service configuration..."
	strcmdline = "cmd /c ""ipconfig /all > " & scripttempvar & "\ipconfigtemp.txt"""	
	temp = wshshell.Run(strcmdline, 0, true)
	if (Wshfile.fileexists(scripttempvar & "\ipconfigtemp.txt")) then
		Set wshtempfile = wshfile.OpenTextFile(scripttempvar & "\ipconfigtemp.txt", 1)
		if err.number <> 0 then err.clear
		ipconfigtext = wshtempfile.readall
		wshtempfile.Close
		wscript.DisconnectObject wshtempfile
		Set wshtempfile=nothing

		pos1 = Instr(1, ipconfigtext, "Primary DNS Suffix", 1)
		pos2 = Instr(pos1, ipconfigtext, chr(10), 1)
		line = mid(ipconfigtext, pos1+36, pos2-pos1-37)
		if line = "redmond.corp.microsoft.com" and Instr(1, ipconfigtext, "PRXY", 1) <> 0 then
			scriptinfravar = "W2Kprxy"
			screenout "  Server is a Redmond Proxy time server (" & line & ")."
		elseif line = "corp.microsoft.com" then
			scriptinfravar = "W2Kroot"
			screenout "  Server is in a Windows 2000 infrastructure root domain (" & line & ")."
		elseif line = "microsoft.com" then
			scriptinfravar = "W2Kroot"
			screenout "  Server is in a Windows 2000 infrastructure root domain (" & line & ")."
		elseif line = "" then
			scriptinfravar = "NT"
			screenout "  Server is in an NT 4.0 infrastructure (" & line & ")."
		else
			scriptinfravar = "W2K"
			screenout "  Server is in a Windows 2000 infrastructure (" & line & ")."
		end if
	else
		screenout ""
		text = "Unable to create a text file in the temporary directory for the script!"
		screenout text
		screenout "Make sure that " & scripttempvar & " exists and is writable."
		screenout "Make sure that ipconfig.exe is available and working properly (ipconfig /all)."
		genevent "E", "3", text
		services_time = 1
		exit function
	end if

	REM ###### Find if SNTP values are set or not ######
	strcmdline = "cmd /c ""net time /querysntp > " & scripttempvar & "\timetemp.txt"""	
	temp = wshshell.Run(strcmdline, 0, true)
	if (Wshfile.fileexists(scripttempvar & "\timetemp.txt")) then 
		Set wshtempfile = wshfile.OpenTextFile(scripttempvar & "\timetemp.txt", 1)
		nettimetext = wshtempfile.readall
		wshtempfile.Close
		wscript.DisconnectObject wshtempfile
		Set wshtempfile=nothing
		if Instr(1, nettimetext, "131.107.1.10 192.5.41.40 192.43.244.18", 1) <> 0 then 
			sntplist = "internet"
			screenout "  SNTP values set to Internet Time servers (131.107.1.10 192.5.41.40 192.43.244.18)."
		elseif Instr(1, nettimetext, "itgproxy.redmond.corp.microsoft.com", 1) <> 0 then 
			sntplist = "proxy"
			screenout "  SNTP values set to proxy servers (itgproxy.redmond.corp.microsoft.com)."
		else  
			sntplist = "empty"
			screenout "  SNTP values NOT currently configured on this server. ()"
		end if
	else
		screenout ""
		text = "Unable to create a text file in the temporary directory for the script!"
		screenout text
		screenout "Make sure that " & scripttempvar & " exists and is writable."
		screenout "Make sure that net.exe is available and working properly (net time)."
		genevent "E", "3", text
		services_time = 1
		exit function
	end if


	REM ###### Find if net time at jobs exist ######
	strcmdline = "cmd /c ""at \\" & computername & " > " & scripttempvar & "\atjobstemp.txt"""	
	temp = wshshell.Run(strcmdline, 0, true)
	if (Wshfile.fileexists(scripttempvar & "\atjobstemp.txt")) then 
		Set wshtempfile = wshfile.OpenTextFile(scripttempvar & "\atjobstemp.txt", 1)
		attext = wshtempfile.readall
		wshtempfile.Close
		wscript.DisconnectObject wshtempfile
		Set wshtempfile=nothing
		if Instr(1, attext, "timesync.cmd", 1) <> 0 then 
			atlist = "full"
			screenout "  TimeSync At jobs exist. (Timesync.cmd)"
		else 
			atlist = "empty"
			screenout "  TimeSync At jobs do not exist. ()"
		end if
	else
		screenout ""
		text = "Unable to create a text file in the temporary directory for the script!"
		screenout text
		screenout "Make sure that " & scripttempvar & " exists and is writable."
		screenout "Make sure that at.exe is available and working properly (at \\servername)."
		genevent "E", "3", text
		services_time = 1
		exit function
	end if


	REM ###### Set Time service accordingly ######
	if scriptinfravar = "W2Kprxy" and not sntplist = "internet" then
		screenout "  Time service being set with NTP Internet servers list..."
		strcmdline = "cmd /c ""net time /setsntp:""131.107.1.10 192.5.41.40 192.43.244.18"""""	
		temp = wshshell.Run(strcmdline, 0, true)
	elseif scriptinfravar = "W2Kroot" and not sntplist = "proxy" then
		screenout "  Time service being set with NTP itgproxy value..."
		strcmdline = "cmd /c ""net time /setsntp:""itgproxy.redmond.corp.microsoft.com"""""	
		temp = wshshell.Run(strcmdline, 0, true)
	elseif scriptinfravar = "W2K" and not sntplist = "empty" then
		screenout "  Time service being cleared of NTP server list..."
		strcmdline = "cmd /c ""net time /setsntp:"""	
		temp = wshshell.Run(strcmdline, 0, true)
	elseif scriptinfravar = "NT" and not sntplist = "empty" then
		screenout "  Time service being cleared of NTP server list..."
		strcmdline = "cmd /c ""net time /setsntp:"""	
		temp = wshshell.Run(strcmdline, 0, true)
	else
		screenout "  Time service configured OK."
	end if


	REM ###### Set AT jobs correctly ######
	if scriptinfravar = "NT" and atlist = "empty" then
		screenout "  Adding Timesync AT jobs..."
		randomize
		hour = Int((6 - 1 + 1) * Rnd(0) + 1)
		randomize
		minute = Int((59 - 0 + 1) * Rnd(0) + 0)
		strcmdline = "cmd /c ""at \\" & computername & " " & hour &":" & minute & " /EVERY:m,t,w,th,f,s,su " & scriptbinvar & "\timesync.cmd"""	
		temp = wshshell.Run(strcmdline, 0, true)
		strcmdline = "cmd /c ""at \\" & computername & " " & hour + 6 &":" & minute & " /EVERY:m,t,w,th,f,s,su " & scriptbinvar & "\timesync.cmd"""	
		temp = wshshell.Run(strcmdline, 0, true)
		strcmdline = "cmd /c ""at \\" & computername & " " & hour + 12 &":" & minute & " /EVERY:m,t,w,th,f,s,su " & scriptbinvar & "\timesync.cmd"""	
		temp = wshshell.Run(strcmdline, 0, true)
		strcmdline = "cmd /c ""at \\" & computername & " " & hour + 18 &":" & minute & " /EVERY:m,t,w,th,f,s,su " & scriptbinvar & "\timesync.cmd"""	
		temp = wshshell.Run(strcmdline, 0, true)
	elseif not scriptinfravar = "NT" and atlist = "full" then
		screenout "  Removing AT jobs..."
		if (Wshfile.fileexists(scripttempvar & "\atjobstemp.txt")) then 
			Set wshtempfile = wshfile.OpenTextFile(scripttempvar & "\atjobstemp.txt", 1)
			Do While wshtempfile.AtEndOfStream <> true
				line = wshtempFile.ReadLine
				if Instr(1, line, "timesync", 1) then
					number = int(mid(line, 8, 2))
					screenout "    ID " & number & "..."
					strcmdline = "cmd /c ""at " & number & " /d"""	
					temp = wshshell.Run(strcmdline, 0, true)
				end if
			Loop
			wshtempfile.Close
			wscript.DisconnectObject wshtempfile
			Set wshtempfile=nothing
		else
			screenout "  Unable to open At file!"
		end if
	else
		screenout "  Timesync AT jobs OK."
	end if


	if err.number <> 0 then logerror "End of services_time Function", err.number : err.clear
	services_time = 0
End Function



REM #################################################################################################
REM ###                                      SNMP service                                         ###
REM #################################################################################################
Function services_snmp()
	dim srvinstalled, strcmdline, temp, text, wshaddonfile
	on error resume next

	REM ###### Snmp service ######
	srvinstalled = WshShell.RegRead("HKLM\system\currentcontrolset\services\snmp\imagepath")
	if err.number <> 0 then err.clear
        if instr(1, srvinstalled, "snmp.exe", 1) <> 0 then
		screenout "SNMP service OK."
	else
		screenout "SNMP service being installed..."
		if (Wshfile.FileExists(systemdrive & "\addon.inf")) then wshfile.deletefile systemdrive & "\addon.inf", TRUE
		set wshaddonfile = wshfile.createtextfile(systemdrive & "\addon.inf",1)
 		wshaddonfile.writeline("[NetOptionalComponents]")
		wshaddonfile.writeline("SNMP = 1")
		wshaddonfile.close
		wscript.DisconnectObject wshaddonfile
		set wshaddonfile=nothing
		strcmdline = "cmd /c ""sysocmgr /i:sysoc.inf /u:" & systemdrive & "\addon.inf /r"""	
		temp = wshshell.Run(strcmdline, 0, true)
		if (Wshfile.FileExists(systemdrive & "\addon.inf")) then wshfile.deletefile systemdrive & "\addon.inf", TRUE
	end if

	if err.number <> 0 then logerror "End of services_snmp Function", err.number : err.clear
	services_snmp = 0
End Function



REM #################################################################################################
REM ###                                      MSMQ service                                         ###
REM #################################################################################################
Function services_msmq()
	dim srvinstalled, strcmdline, temp, text, wshaddonfile
	on error resume next

	REM ###### MSMQ service ######
	srvinstalled = WshShell.RegRead("HKLM\system\currentcontrolset\services\msmq\imagepath")
	if err.number <> 0 then err.clear
	if srvrolevar = "Domain Controller" and srvdomtypevar = "MUD" and instr(1, srvinstalled, "mqsvc.exe", 1) <> 0 then
		screenout "MSMQ service OK."
	elseif srvrolevar = "Domain Controller" and srvdomtypevar = "MUD" then
		screenout "MSMQ service being installed..."
		if (Wshfile.FileExists(systemdrive & "\addon.inf")) then wshfile.deletefile systemdrive & "\addon.inf", TRUE
		set wshaddonfile = wshfile.createtextfile(systemdrive & "\addon.inf",1)
 		wshaddonfile.writeline("[Components]")
		wshaddonfile.writeline("msmq = on")
		wshaddonfile.close
		wscript.DisconnectObject wshaddonfile
		set wshaddonfile=nothing
		strcmdline = "cmd /c ""sysocmgr /i:sysoc.inf /u:" & systemdrive & "\addon.inf /r"""	
		temp = wshshell.Run(strcmdline, 0, true)
		if (Wshfile.FileExists(systemdrive & "\addon.inf")) then wshfile.deletefile systemdrive & "\addon.inf", TRUE
	else
		screenout "MSMQ service needed only on MUDC's."
	end if

	if err.number <> 0 then logerror "End of services_msmq Function", err.number : err.clear
	services_msmq = 0
End Function




REM #################################################################################################
REM ###                                    Check Compaq NIC                                       ###
REM #################################################################################################
Function hardware_cpqnic()
	dim text, temp, strcmdline, wshtempfile, ipconfigtext, wshinifile, nic
	on error resume next

	REM ###### Find what operating system infrastructure the server is in ######
	screenout "COMPAQ DTC NIC Check:"
	if srvsuitevar = "DTC" and srvhardwarevar = "COMPAQ" then
		strcmdline = "cmd /c ""ipconfig /all > " & scripttempvar & "\ipconfigtemp.txt"""	
		temp = wshshell.Run(strcmdline, 0, true)
		if (Wshfile.fileexists(scripttempvar & "\ipconfigtemp.txt")) then
			Set wshtempfile = wshfile.OpenTextFile(scripttempvar & "\ipconfigtemp.txt", 1)
			if err.number <> 0 then err.clear
			ipconfigtext = wshtempfile.readall
			wshtempfile.Close
			wscript.DisconnectObject wshtempfile
			Set wshtempfile=nothing
			rem screenout ipconfigtext
			if (Wshfile.fileexists(scriptinilocvar & "\DisableNicsW2K.ini")) then 
				Set wshinifile = wshfile.OpenTextFile(scriptinilocvar & "\DisableNicsW2K.ini", 1)
				Do While wshinifile.AtEndOfStream <> true
					nic = wshiniFile.ReadLine
					rem screenout nic
					if not instr(1, ipconfigtext, nic, 1) = 0 then
						screenout ""
						text = "The script detected a Compaq NIC that needs to be disabled when used with the Datacenter suite!" 
						screenout text
						text = text & " (" & nic & ")" 
						screenout "  (" & nic & ")"
						screenout "The device can be disabled using 'Device Manager' in 'Computer Manager'."
						screenout "Find the above listed device under 'Network Adapters', right click on it, and select disable."
						screenout "If this is the only NIC in the machine then it MUST be replaced with a supported network adapter."
						genevent "E", "3", text
						hardware_cpqnic = 1
						exit function
					end if
				loop
				wshinifile.Close
				screenout "  Check for Compaq Netflex3 NIC on DTC OK."
			else
				screenout ""
				text = "The script was unable to find the DisableNics.ini file!"
				screenout text
				screenout "Make sure that " & scriptinilocvar & " exists and contains the necessary INI file."
				genevent "E", "3", text
				hardware_cpqnic = 1
				exit function
			end if

		else
			screenout ""
			text = "Unable to create a text file in the temporary directory for the script!"
			screenout text
			screenout "Make sure that " & scripttempvar & " exists and is writable."
			screenout "Make sure that ipconfig.exe is available and working properly (ipconfig /all)."
			genevent "E", "3", text
			hardware_cpqnic = 1
			exit function
		end if
	else
		screenout "  Check for Compaq Netflex3 NIC on DTC not needed."
	end if


	if err.number <> 0 then logerror "End of admin_asset Function", err.number : err.clear
	hardware_cpqnic = 0
End Function






REM #################################################################################################
REM ###                                           SSD                                             ###
REM #################################################################################################
Function hardware_ssd()
	dim temp, strcmdline, text, regthere, filethere, servfile, lettervar, locationvar, ended
	dim acufile
	on error resume next

	REM ### SSD ###
	screenout "Compaq SSD (5.03a):"
	if not srvhardwarevar = "COMPAQ"  then
		screenout "  SSD is not needed on this hardware platform."
		cpqssdlogvar = "Different Hardware Platform."
	elseif cpqssdvar = "off"  then
		screenout "  SSD install is disabled."
		cpqssdlogvar = "Disabled"
	elseif cpqssdvar = "switch"  then
		screenout "  SSD install has been disabled by the user."
		cpqssdlogvar = "Disabled - By User"
	else
		temp = WshShell.Regread("HKLM\system\currentcontrolset\services\sysmgmt\imagepath")
		if err.number = 0 then
			regthere = "yes"
		else
			err.clear
			regthere = "no"
		end if
		if (Wshfile.fileexists(cpqssdchkvar)) then
			set servfile = Wshfile.getfile(cpqssdchkvar)
			filethere = datevalue(left(servfile.datelastmodified, Instr(1, servfile.datelastmodified, " ", 1)-1))
			wscript.DisconnectObject servfile
			set servfile=nothing
		else
			filethere = "no"
		end if
		if (Wshfile.fileexists(systemdrive & "\program files\compaq\cpqacu\cpqacu.exe")) then
			set servfile = Wshfile.getfile(systemdrive & "\program files\compaq\cpqacu\cpqacu.exe")
			acufile = datevalue(left(servfile.datelastmodified, Instr(1, servfile.datelastmodified, " ", 1)-1))
			wscript.DisconnectObject servfile
			set servfile=nothing
		else
			acufile = "no"
		end if
		cpqssdfilevar = datevalue(cpqssdfilevar)

		if debug = 1 then screenout "  (SSD Registry: " & regthere & ")"
		if debug = 1 then screenout "  (SSD File: " & filethere & ")"
		if debug = 1 then screenout "  (ACU File: " & acufile & ")"

		if not filethere = "no" and not datecomp(datevalue("1/13/2000"), filethere) then 
			screenout "  Running Primer to remove old utilities and services..."
			screenout "    SNMP may AV during this process so please clear the pop up message if it does."
			screenout "    When the server comes up after the reboot chkdsk may run as well so please let it continue."
			screenout "    Both of these are normal when Primer runs."
			strcmdline = cpqssdlocvar & "\primer\primer s"
			if debug = 1 then screenout "  (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, true)
			cpqssdlogvar = "Removed"
			wscript.sleep 5000 * netfactor
			screenout ""
			text = "The Primer utility has been ran to remove services and the server must now be rebooted!"
			screenout text
			screenout "CIM will not function properly on a re-install until the server is rebooted."
			screenout "The script will reboot the server after the pop up message is cleared."
			screenout "After the server comes back up, just re-run the script."
		 	genevent "I", "4", text
			genevent "I", "5", "Server was rebooted by " & scriptnamevar & "."
		 	strCmdLine = "cmd /c """& scriptfileslocvar &"\reboot /L /R /T:5"""
		 	temp = wshshell.Run(strcmdline, 0, true)
			hardware_ssd = 1
			exit function
		elseif regthere = "yes" and datecomp(cpqssdfilevar, filethere) and datecomp("5/25/2000", acufile) and (forceupdate = 1 or argument("newcpq") <> 0) then 
			screenout "  Updating Compaq SSD components..."
			cpqssdlogvar = "Updated"
		elseif regthere = "yes" and datecomp(cpqssdfilevar, filethere) and datecomp("5/25/2000", acufile) then 
			screenout "  Compaq SSD components OK."
			screenout "    Running Compaq Advanced System Management Controller Driver(184)..."
			strcmdline = cpqssdlocvar & "\cp000184.exe /s /f"
			if debug = 1 then screenout "      (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, true)
			screenout "    Running Compaq Integrated System Management Controller Driver(223)..."
			strcmdline = cpqssdlocvar & "\cp000223.exe /s /f"
			if debug = 1 then screenout "      (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, true)
			screenout "    Running Compaq System Management Controller Driver(224)..."
			strcmdline = cpqssdlocvar & "\cp000224.exe /s /f"
			if debug = 1 then screenout "      (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, true)
			cpqssdlogvar = "OK"
			hardware_ssd = 0
			exit function
		elseif regthere = "no" then 
			screenout "  Installing Compaq SSD components..."
			cpqssdlogvar = "Installed"
		else
			screenout "  Upgrading Compaq SSD components..."
			cpqssdlogvar = "Upgraded"
		end if
		if (Wshfile.FileExists(cpqssdlocvar & "\CP000224.exe")) and checkfreespace(25) = 0 then
			screenout "    Compaq Advanced System Management Controller Driver(184)..."
			strcmdline = cpqssdlocvar & "\cp000184.exe /s /f"
			if debug = 1 then screenout "      (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, true)
			screenout "    Compaq Integrated System Management Controller Driver(223)..."
			strcmdline = cpqssdlocvar & "\cp000223.exe /s /f"
			if debug = 1 then screenout "      (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, true)
			screenout "    Compaq System Management Controller Driver(224)..."
			strcmdline = cpqssdlocvar & "\cp000224.exe /s /f"
			if debug = 1 then screenout "      (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, true)

			screenout "    Compaq Drive Array Driver(175)..."
			strcmdline = cpqssdlocvar & "\cp000175.exe /s"
			if debug = 1 then screenout "      (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, false)
			ended = waituntil("cp000175.exe", "")
			ended = waituntil("Upgrade Device Driver Wizard", "cp000175.exe")
			if ended = 0 then 
				screenout "      Clearing Signature pop up..."
				winmanipulate "Upgrade Device Driver Wizard", "Digital Signature Not Found", "%y"
			end if
			ended = waituntil("", "cp000175.exe")

			screenout "    Compaq Smart Array-2 Controllers Driver(176)..."
			strcmdline = cpqssdlocvar & "\cp000176.exe /s"
			if debug = 1 then screenout "      (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, true)
			screenout "    Compaq 32-Bit SCSI Controller(177)..."
			strcmdline = cpqssdlocvar & "\cp000177.exe /s"
			if debug = 1 then screenout "      (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, true)
			screenout "    Compaq 64-Bit/66-Mhz Wide Ultra3 Controller Driver(193)..."
			strcmdline = cpqssdlocvar & "\cp000193.exe /s"
			if debug = 1 then screenout "      (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, true)
			screenout "    Compaq StorageWorks Fibre Channel Host Bus Adapter Driver(178)..."
			strcmdline = cpqssdlocvar & "\cp000178.exe /s"
			if debug = 1 then screenout "       (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, true)

			screenout "    Compaq NetFlex/Netelligent Adapter Driver(181)..."
			strcmdline = cpqssdlocvar & "\cp000181.exe /s"
			if debug = 1 then screenout "      (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, false)
			ended = waituntil("cp000181.exe", "")
			ended = waituntil("Upgrade Device Driver Wizard", "cp000181.exe")
			if ended = 0 then 
				screenout "      Clearing Signature pop up..."
				winmanipulate "Upgrade Device Driver Wizard", "Digital Signature Not Found", "%y"
			end if
			ended = waituntil("", "cp000181.exe")
			screenout "      Pausing to wait for network to re-establish..."
			wscript.sleep 12000 * netfactor

			screenout "    Compaq Ethernet or Fast Ethernet NIC Driver(182)..."
			strcmdline = cpqssdlocvar & "\cp000182.exe /s /f"
			if debug = 1 then screenout "      (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, true)
			screenout "      Pausing to wait for network to re-establish..."
			wscript.sleep 12000 * netfactor

			screenout "    Compaq Gigabit Ethernet NIC Drive(183)..."
			strcmdline = cpqssdlocvar & "\cp000183.exe /s"
			if debug = 1 then screenout "      (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, true)
			screenout "      Pausing to wait for network to re-establish..."
			wscript.sleep 12000 * netfactor

			screenout "    Compaq Network Teaming and Configuration(195)..."
			strcmdline = cpqssdlocvar & "\cp000195.exe /s"
			if debug = 1 then screenout "      (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, true)
			screenout "    Compaq ProLiant Storage System Support(179)..."
			strcmdline = cpqssdlocvar & "\cp000179.exe /s"
			if debug = 1 then screenout "      (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, true)
			screenout "    Compaq Drive Array Notification(180)..."
			strcmdline = cpqssdlocvar & "\cp000180.exe /s"
			if debug = 1 then screenout "      (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, true)
			screenout "    Compaq Remote Monitor Service(191)..."
			strcmdline = cpqssdlocvar & "\cp000191.exe /s"
			if debug = 1 then screenout "      (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, true)
			screenout "    Compaq PCI Hot Plug Controller Driver(86)..."
			strcmdline = cpqssdlocvar & "\cp000086.exe /s"
			if debug = 1 then screenout "      (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, true)
			screenout "    Compaq Integrated Management Display Utility(187)..."
			strcmdline = cpqssdlocvar & "\cp000187.exe /s"
			if debug = 1 then screenout "      (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, true)
			screenout "    Compaq Integrated Management Log Viewer(186)..."
			strcmdline = cpqssdlocvar & "\cp000186.exe /s"
			if debug = 1 then screenout "      (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, true)
			screenout "    Compaq Power Supply Viewer(188)..."
			strcmdline = cpqssdlocvar & "\cp000188.exe /s"
			if debug = 1 then screenout "      (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, true)
			screenout "    Compaq Power Down Manager(189)..."
			strcmdline = cpqssdlocvar & "\cp000184.exe /s"
			if debug = 1 then screenout "      (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, true)

			REM ###### This module is causing lock ups on DTC servers! ######
			REM screenout "    Compaq Enhanced Integrated Management Display Service(190)..."
			REM strcmdline = cpqssdlocvar & "\cp000190.exe /s"
			REM if debug = 1 then screenout "      (" & strcmdline & ")"
			REM temp = wshshell.Run(strcmdline, 1, true)

			screenout "    Compaq Array Configuration Utility(192)..."
			strcmdline = cpqssdlocvar & "\cp000192.exe /s"
			if debug = 1 then screenout "      (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, true)
			screenout "    Compaq ATI RAGE IIC Video Controller Support(95)..."
			strcmdline = cpqssdlocvar & "\cp000095.exe /s"
			if debug = 1 then screenout "      (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, true)
			screenout "    Compaq Remote Insight Board Driver(196)..."
			strcmdline = cpqssdlocvar & "\cp000196.exe /s"
			if debug = 1 then screenout "      (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, true)
		elseif checkfreespace(25) = 1 then
			text = "There is not enough free space on the boot drive to install SSD!"
			screenout text
			screenout "Please free up 25 MB of disk space on the boot drive and re-run the script."
			genevent "E", "3", text
			hardware_ssd = 1
			exit function
		else
			text = "The script could not locate the SSD setup file!"
			screenout text
			screenout "Make sure the below file exists on the source that you are running the script from."
			screenout "  (" & cpqssdlocvar & "\CP000224.exe)"
			genevent "E", "3", text
			hardware_ssd = 1
			exit Function
		end if
	end if

	if err.number <> 0 then logerror "End of Function hardware_ssd", err.number : err.clear
	hardware_ssd = 0
End Function




REM #################################################################################################
REM ###                                           CIM                                             ###
REM #################################################################################################
Function hardware_cim()
	dim temp, strcmdline, regthere, filethere, servfile, lettervar, locationvar, ended
	on error resume next

	REM ### CIM ###
	screenout "Compaq Insight Manager (4.80):"
	if not srvhardwarevar = "COMPAQ"  then
		screenout "  CIM is not needed on this hardware platform."
		cpqcimlogvar = "Different Hardware Platform."
	elseif cpqcimvar = "off"  then
		screenout "  CIM install is disabled."
		cpqcimlogvar = "Disabled"
	elseif cpqcimvar = "switch"  then
		screenout "  CIM install has been disabled by the user."
		cpqcimlogvar = "Disabled - By User."
	else
		temp = WshShell.Regread("HKLM\system\currentcontrolset\services\cqmgserv\imagepath")
		if err.number = 0 then
			regthere = "yes"
		else
			err.clear
			regthere = "no"
		end if
		if (Wshfile.fileexists(cpqcimchkvar)) then
			set servfile = Wshfile.getfile(systemroot & "\system32\cpqmgmt\CqMgServ\CqMgServ.exe")
			filethere = datevalue(left(servfile.datelastmodified, Instr(1, servfile.datelastmodified, " ", 1)-1))
			wscript.DisconnectObject servfile
			set servfile=nothing
		else
			filethere = "no"
		end if
		cpqcimfilevar = datevalue(cpqcimfilevar)

		if debug = 1 then screenout "  (CIM Registry: " & regthere & ")"
		if debug = 1 then screenout "  (CIM File: " & filethere & ")"

		if regthere = "yes"  and datecomp(cpqcimfilevar, filethere) and (forceupdate = 1 or argument("newcpq") <> 0) then 
			screenout "  Updating Compaq Insight Manager (CIM) Agents..."
			cpqcimlogvar = "Updated"
		elseif regthere = "yes"  and datecomp(cpqcimfilevar, filethere) then 
			screenout "  Compaq Insight Manager (CIM) Agents OK."
			cpqcimlogvar = "OK"
			hardware_cim = 0
			exit function
		elseif regthere = "no" then 
			screenout "  Installing Compaq Insight Manager (CIM) Agents..."
			cpqcimlogvar = "Installed"
		else
			screenout "  Upgrading Compaq Insight Manager (CIM) Agents..."
			cpqcimlogvar = "Upgraded"
		end if

		if (Wshfile.FileExists(cpqcimlocvar & "\cp000150.exe")) and checkfreespace(25) = 0 then
			screenout "    Compaq Foundation Agent(150)..."
			strcmdline = cpqssdlocvar & "\cp000150.exe /s"
			if debug = 1 then screenout "      (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, true)
			screenout "    Compaq Server Agent(151)..."
			strcmdline = cpqssdlocvar & "\cp000151.exe /s"
			if debug = 1 then screenout "      (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, true)
			screenout "    Compaq Storage Agent(153)..."
			strcmdline = cpqssdlocvar & "\cp000153.exe /s"
			if debug = 1 then screenout "      (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, true)
			screenout "    Compaq NIC Agent(152)..."	
			strcmdline = cpqssdlocvar & "\cp000152.exe /s"
			if debug = 1 then screenout "      (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, true)
			screenout "    Compaq Survey Utility(154)..."
			strcmdline = cpqssdlocvar & "\cp000154.exe /s"
			if debug = 1 then screenout "      (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, true)

			REM ###### Set CIM Registry entries ######	
			screenout "  Checking CIM registry entries:"
			registry(scriptinilocvar & "\CIMRegistryW2K.ini")
		elseif checkfreespace(25) = 1 then
			screenout ""
			text = "There is not enough free space on the boot drive to install CIM!"
			screenout text
			screenout "Please free up 25 MB of disk space on the boot drive and re-run the script."
			genevent "E", "3", text
			hardware_cim = 1
			exit function
		else
			screenout ""
			text = "The script could not locate the CIM setup file!"
			screenout text
			screenout "Make sure the below file exists on the source that you are running the script from."
			screenout "  (" & cpqcimlocvar & "\cp000150.exe)"
			genevent "E", "3", text
			hardware_cim = 1
			exit Function
		end if
	end if

	if err.number <> 0 then logerror "End of Function", err.number : err.clear
	hardware_cim = 0
End Function





REM #################################################################################################
REM ###                         Compaq Storage Works LP7000 & LP8000 Driver                       ###
REM #################################################################################################
Function hardware_storageworksdriver()
	dim temp, strcmdline, text, wrongfile, filethere, servfile
	dim hubaddress, data
	on error resume next


	REM ### Check for driver dates ###
	screenout "Compaq Storage Works LP Driver (5-4.41A5):"
	if srvstrwrksvar = "ABSENT" then
		screenout "  The Compaq Storage Works LP driver does not exist on this server."
		cpqstrdrvlogvar = "Not installed."
	elseif srvstrwrksvar = "DISABLED" then
		screenout "  The Compaq Storage Works LP driver update has been disabled by the user."
		cpqstrdrvlogvar = "Disabled - By User."
	elseif cpqstrdrvvar = "off" then
		screenout "  The Compaq Storage Works LP driver update is disabled."
		cpqstrdrvlogvar = "Disabled."
	else
		REM ### Get LP driver date ###
		if (Wshfile.fileexists(cpqstrdrvchkvar)) then
			set servfile = Wshfile.getfile(cpqstrdrvchkvar)
			filethere = datevalue(left(servfile.datelastmodified, Instr(1, servfile.datelastmodified, " ", 1)-1))
			wscript.DisconnectObject servfile
			set servfile=nothing
		else
			filethere = "no"
		end if
		cpqstrdrvfilevar = datevalue(cpqstrdrvfilevar)

		if debug = 1 then screenout "  (Storage Works LP File: " & filethere & ")"

		if datecomp(cpqstrdrvfilevar, filethere) and 0 = 1 then
			screenout "  Compaq Storage Works LP driver (" & cpqstrdrvvar & ") OK."
			cpqstrdrvlogvar = "OK"
		else
			screenout "  The Compaq Storage Works LP Driver needs to be updated before the update script can continue!"
			screenout "  The Compaq Storage Works LP Driver needs to be running version " & cpqstrdrvvar & "." 
			screenout "  This requires user intervention! Please follow these steps:"
			screenout "    In 'Device Manager' of 'Computer Manager'"
			screenout "      - Open the 'SCSI and RAID Controllers' Devices"
			screenout "    On the 'Emulex LP6000/7000/8000/9000/850/950, PCI-Fibre Channel Adapter' device"
			screenout "      - Right click and select 'Properties'"
			screenout "    In the 'Emulex LP6000/7000/8000/9000/850/950, PCI-Fibre Channel Adapter' window"
			screenout "      - Select the 'Driver' tab and click on the 'Update Driver' button"
			screenout "    In the 'Upgrade Device Driver Wizard' window"
			screenout "      - Press the 'Next' button"
			screenout "     - Make sure the 'Display a list of known drivers for this device"
			screenout "        so that I can choose a specific driver' radial button is checked"
			screenout "      - Press the 'Next' button"
			screenout "      - Press the 'Have Disk' button"
			screenout "    In the 'Install from Disk' window"
			screenout "      - Press the 'Browse' button"
			screenout "    In the 'Locate File' window"
			screenout "      - Open up the 'Look in' box and select the " & scriptdrivelocvar & " drive"
			screenout "      - Move into the " & cpqstrdrvlocvar & " directory"
			screenout "      - Highlight the 'oemsetup.inf' file"
			screenout "      - Press the 'Open' button"
			screenout "    In the 'Install from Disk' window"
			screenout "      - Press the 'OK' button"
			screenout "    In the 'Upgrade Device Driver Wizard' window"
			screenout "      - Highlight the 'Compaq KGPSA' Model and press the 'Next' button"
			screenout "      - Press the 'Next' button"
			screenout "    In the 'Digital Signature Not Found' window"
			screenout "      - Press the 'Yes' button "
			screenout "   In the 'Upgrade Device Driver Wizard' window"
			screenout "      - Press the 'Finish' button"
			screenout "      - Close out of the properties window"
			screenout "    In the 'System Settings Change' window"
			screenout "      - Press the 'No' button"
			screenout "  Repeat the above steps for ALL Storage Works LP devices!"
			screenout "  Close out of Computer Manager when finished."
			strcmdline = "cmd /c ""compmgmt.msc /s"""	
			temp = wshshell.Run(strcmdline, 0, true)
			cpqstrdrvlogvar = "Upgraded"
		end if
		if instr(1, srvstrwrksvar, "HUB-YES-", 1) <> 0 then
			hubaddress = right(srvstrwrksvar, 4)
			data = "RetryIoTimeOut=1;RetryInterval=52;enabledpc=1;queuetarget=1;queuedepth=25;Topology=0;ScanDown=1;ElsRetryCount=6;SimulateDevice=0;HardALPA=" & hubaddress & ";NodeTimeOut=10;LinkTimeOut=40;HlinkTimeOut=5"
			screenout "  Setting Storage Works SecurePath Hub driver parameter..."
			wshshell.regwrite "HKLM\system\currentcontrolset\services\cqpkgpsa\parameters\device\driverparameter", data, "REG_SZ"
			if err.number <> 0 then logerror "Writing Storage Works SecurePath Hub driver parameter registry location", err.number : err.clear
			logregistryentry "Storage Works SecurePath Hub driver parameter", "Updated", "HKLM\system\currentcontrolset\services\cqpkgpsa\parameters\device\driverparameter", data
		elseif instr(1, srvstrwrksvar, "HUB-NO-", 1) <> 0  then
			hubaddress = right(srvstrwrksvar, 4)
			data = "RetryIoTimeOut=1;RetryInterval=52;enabledpc=1;queuetarget=1;queuedepth=25;NodeTimeout=60;Topology=0;ScanDown=1;ElsRetryCount=6;SimulateDevice=0;HardALPA=" & hubaddress
			screenout "  Setting Storage Works Hub driver parameter..."
			wshshell.regwrite "HKLM\system\currentcontrolset\services\cqpkgpsa\parameters\device\driverparameter", data, "REG_SZ"
			if err.number <> 0 then logerror "Writing Storage Works Hub driver parameter registry location", err.number : err.clear
			logregistryentry "Storage Works Hub driver parameter", "Updated", "HKLM\system\currentcontrolset\services\cqpkgpsa\parameters\device\driverparameter", data
		elseif srvstrwrksvar = "SWITCH-YES" then
			data = "RetryIoTimeOut=1;RetryInterval=52;enabledpc=1;queuetarget=1;queuedepth=25;Topology=1;ElsRetryCount=6;SimulateDevice=0;NodeTimeOut=10;LinkTimeOut=40;HlinkTimeOut=5"
			screenout "  Setting Storage Works SecurePath Switch driver parameter..."
			wshshell.regwrite "HKLM\system\currentcontrolset\services\cqpkgpsa\parameters\device\driverparameter", data, "REG_SZ"
			if err.number <> 0 then logerror "Writing Storage Works SecurePath Switch driver parameter registry location", err.number : err.clear
			logregistryentry "Storage Works SecurePath Switch driver parameter", "Updated", "HKLM\system\currentcontrolset\services\cqpkgpsa\parameters\device\driverparameter", data
		elseif srvstrwrksvar = "SWITCH-NO" then
			data = "RetryIoTimeOut=1;RetryInterval=52;enabledpc=1;queuetarget=1;queuedepth=25;NodeTimeout=60;Topology=1;ElsRetryCount=6;SimulateDevice=0"
			screenout "  Setting Storage Works Switch driver parameter..."
			wshshell.regwrite "HKLM\system\currentcontrolset\services\cqpkgpsa\parameters\device\driverparameter", data, "REG_SZ"
			if err.number <> 0 then logerror "Writing Storage Works Switch driver parameter registry location", err.number : err.clear
			logregistryentry "Storage Works Switch driver parameter", "Updated", "HKLM\system\currentcontrolset\services\cqpkgpsa\parameters\device\driverparameter", data
		end if

	end if

	if err.number <> 0 then logerror "End of hardware_storageworksdriver Function", err.number : err.clear
	hardware_storageworksdriver = 0
End Function




REM #################################################################################################
REM ###                                       Managed Node                                        ###
REM #################################################################################################
Function hardware_managednode(rebootcheck)
	dim temp, strcmdline, text, regthere, filethere, servfile, lettervar, locationvar
	on error resume next

	REM ### Check for  location ###
	screenout "Dell Managed Node (1.52):"
	if not srvhardwarevar = "DELL"  then
		screenout "  Managed Node is not needed on this hardware platform."
		delmannlogvar = "Different Hardware Platform."
	elseif delmannvar = "off"  then
		screenout "  Managed Node install is disabled."
		delmannlogvar = "Disabled"
	elseif delmannvar = "switch"  then
		screenout "  Managed Node install has been turned off by the user."
		delmannlogvar = "Disabled - By User"
	else
		if (Wshfile.fileexists(systemroot & "\system32\msvcrt.old")) then wshfile.deletefile systemroot & "\system32\msvcrt.old", TRUE
		if err.number <> 0 then err.clear
		temp = WshShell.Regread("HKLM\system\currentcontrolset\services\dell baseboard agent\imagepath")
		if err.number = 0 then
			regthere = "yes"
		else
			err.clear
			regthere = "no"
		end if
		if (Wshfile.fileexists(delmannchkvar)) then
			set servfile = Wshfile.getfile(delmannchkvar)
			filethere = datevalue(left(servfile.datelastmodified, Instr(1, servfile.datelastmodified, " ", 1)-1))
			wscript.DisconnectObject servfile
			set servfile=nothing
		else
			filethere = "no"
		end if
		delmannfilevar = datevalue(delmannfilevar)

		if debug = 1 then screenout "  (DMN Registry: " & regthere & ")"
		if debug = 1 then screenout "  (DMN File: " & filethere & ")"

		if regthere = "yes" and not datecomp(delmannfilevar, filethere) then
			screenout ""
			text = "Dell Managed Node needs to be removed before the update script can continue!"
			screenout text
			screenout "Removing Dell Managed Node requires user intervention."
			screenout "To remove Version 1.50, follow these steps:"
			screenout "  - Go to Start - Programs - 'Dell OpenManage Applications'"
			screenout "  - Select 'Uninstall NNM Se Components'."
			screenout "  - Press 'OK'."
			screenout "  - Press 'ALL'."
			screenout "  - Press 'OK'."
			screenout "  - Press 'OK', when the removal is complete."
			screenout "  - Press 'No' to reboot the server."
			screenout "Other versions may be named different but the procedure is still the same."
			screenout ""
			genevent "W", "3", text
			hardware_managednode = 1
			exit function
		elseif rebootcheck then 
			screenout "  The reboot check for Dell Managed Node is OK."
			hardware_managednode = 0
			exit function
		elseif regthere = "yes" and datecomp(delmannfilevar, filethere) then 
			screenout "  Dell Managed Node OK."
			delmannlogvar = "OK"
			hardware_managednode = 0
			exit function
		elseif regthere = "no" then
			screenout "  Installing Dell Managed Node..."
			delmannlogvar = "Installed"
		else 
			screenout "  Upgrading Dell Managed Node..."
			delmannlogvar = "Upgraded"
		end if
		if (Wshfile.FileExists(delmannlocvar & "\setup.exe")) and checkfreespace(25) = 0 then
			screenout "  This requires user intervention! Please follow these steps:"
			screenout "  In the 'Dell OpenManage Master Setup' screen:"
			screenout "    - Use all defaults"
			screenout "    - Press the 'Next' Button"
			screenout "  In the 'License' screen:"
			screenout "    - Press the 'I agree' button"
			screenout "  In the 'Dell OpenManage Managed Node Setup' screen:"
			screenout "    - Select 'Dell Managed Node for Windows NT'"
			screenout "    - Press the 'Install' button"
			screenout "  In the 'Dell OpenManage Managed Node Manager - Windows NT' screen:"
			screenout "    - Press the 'OK' button"
			screenout "  In the 'Dell OpenManage Connection Setup' screen:"
			screenout "    - Press the 'No' button"
			lettervar = left(delmannlocvar, Instr(1, delmannlocvar, ":", 1))
			locationvar = right(delmannlocvar, len(delmannlocvar)-2)
			strcmdline = "cmd /c """& lettervar & "& cd " & locationvar & "&" & " setup.exe & " & systemdrive & """"
			if debug = 1 then screenout "  (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 0, true)
			if (Wshfile.fileexists(systemroot & "\system32\msvcrt.dlltcb")) then 
				screenout "  Removing file msvcrt.dlltcb..."
				wshfile.deletefile systemroot & "\system32\msvcrt.dlltcb", TRUE
			end if
			if err.number <> 0 then err.clear
		elseif checkfreespace(25) = 1 then
			screenout ""
			text = "There is not enough free space on the boot drive to install Managed node!"
			screenout text
			screenout "Please free up 25 MB of disk space on the boot drive and re-run the script."
			genevent "E", "3", text
			hardware_managednode = 1
			exit function
		else
			screenout ""
			text = "The script could not locate the Managed Node setup file!"
			screenout text
			screenout "Make sure the below file exists on the source that you are running the script from."
			screenout "  (" & delmannlocvar & "\setup.exe)"
			genevent "E", "3", text
			hardware_managednode = 1
			exit Function
		end if
	end if

	if err.number <> 0 then logerror "End of hardware_managednode Function", err.number : err.clear
	hardware_managednode = 0
End Function




REM #################################################################################################
REM ###                                       Perc2 Fast                                          ###
REM #################################################################################################
Function hardware_perc2fast(rebootcheck)
	dim temp, strcmdline, text, regthere, filethere, servfile, lettervar, locationvar
	on error resume next

	REM ### Check for  location ###
	screenout "Dell FAST Utility (2.1 - 2972):"
	if not srvhardwarevar = "DELL"  then
		screenout "  The FAST utility is not needed on this hardware platform."
		delfastlogvar = "Different Hardware Platform."
	elseif delfastvar = "off"  then
		screenout "  The Fast Utility install is disabled."
		delfastlogvar = "Disabled"
	elseif delfastvar = "switch"  then
		screenout "  The Fast Utility install has been turned off by the user."
		delfastlogvar = "Disabled - By User"
	else
		temp = WshShell.Regread("HKLM\system\currentcontrolset\services\afa_agent\imagepath")
		if err.number = 0 then
			regthere = "yes"
		else
			err.clear
			regthere = "no"
		end if
		if (Wshfile.fileexists(delfastchkvar)) then
			set servfile = Wshfile.getfile(delfastchkvar)
			filethere = datevalue(left(servfile.datelastmodified, Instr(1, servfile.datelastmodified, " ", 1)-1))
			wscript.DisconnectObject servfile
			set servfile=nothing
		else
			filethere = "no"
		end if
		delfastfilevar = datevalue(delfastfilevar)

		if debug = 1 then screenout "  (FAST Registry: " & regthere & ")"
		if debug = 1 then screenout "  (FAST File: " & filethere & ")"

		if regthere = "yes" and not datecomp(delfastfilevar, filethere) then
			screenout ""
			text = "Dell FAST needs to be removed before the update script can continue!"
			screenout text
			screenout "The old version must be removed before the new version can be installed."
			screenout "Removing Dell FAST requires user intervention."
			screenout "To remove Version 2941, follow these steps:"
			screenout "  - Go to Start - Settings - 'Control Panel'"
			screenout "  - Select 'Add/Remove Programs'."
			screenout "  - Highligh 'Dell Perc 2'."
			screenout "  - Press the 'Change/Remove' button."
			screenout "  - Press 'Yes'."
			screenout "  - Press 'Yes'."
			screenout "      (If the uninstall script hangs, close out of the .cmd window and re-start it.)"
			screenout "  - Press 'OK'."
			screenout "  - Close out of 'Add/Remove Programs' and 'Control Panel'."
			screenout "Other versions may be named different but the procedure is still the same."
			screenout ""
			genevent "W", "3", text
			hardware_perc2fast = 1
			exit function
		elseif rebootcheck then 
			screenout "  The reboot check for Dell FAST is OK."
			hardware_perc2fast = 0
			exit function
		elseif regthere = "yes" and datecomp(delfastfilevar, filethere) then 
			screenout "  Dell FAST OK."
			delfastlogvar = "OK"
			hardware_perc2fast = 0
			exit function
		elseif regthere = "no" then
			screenout "  Installing Dell FAST..."
			if debug = 1 then screenout "    (FAST Registry: " & regthere & ")"
			if debug = 1 then screenout "    (FAST File: " & filethere & ")"
			delfastlogvar = "Installed"
		else
			screenout "  Upgrading Dell FAST..."
			if debug = 1 then screenout "    (FAST Registry: " & regthere & ")"
			if debug = 1 then screenout "    (FAST File: " & filethere & ")"
			delfastlogvar = "Upgraded"
		end if
		if (Wshfile.FileExists(delfastlocvar & "\setup.exe")) and checkfreespace(25) = 0 then
			screenout "  This requires user intervention! You will need to:"
			screenout "  In the 'Welcome to the setup of Dell Perc 2' screen:"
			screenout "    - Press the 'Next' Button"
			screenout "  In the 'User Information' screen:"
			screenout "    - Press the 'Next' button"
			screenout "  In the 'Select Program Folder' screen:"
			screenout "    - Press the 'Next' button"
			screenout "  In the 'Dell Perc2; 2/Si; 3/Di; 3/Si Setup: Destination Base Directory Selection' screen:"
			screenout "    - Press the 'Next' button"
			screenout "  In the 'Installation Type Selection' screen:"
			screenout "    - Press the 'Next' button"
			screenout "  In the 'Start Copying Files' screen:"
			screenout "    - Press the 'Next' button"
			screenout "  In the 'Question' screen:"
			screenout "    - Press the 'No' button"
			screenout "  In the 'Dell Perc2; 2/Si; 3/Di; 3/Si Setup Complete' screen:"
			screenout "    - Select the 'No, I will restart my computer later' radial button"
			screenout "    - Press the 'Finish' button"
			lettervar = left(delfastlocvar, Instr(1, delfastlocvar, ":", 1))
			locationvar = right(delfastlocvar, len(delfastlocvar)-2)
			strcmdline = "cmd /c """& lettervar & "& cd " & locationvar & "&" & " setup.exe & " & systemdrive & """"
			if debug = 1 then screenout "    (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 0, true)
			waituntil "", "Dell PERC 2; 2/Si; 3/Di; 3/Si Setup"
			screenout "  Setting all Perc2 array controllers write cache to Always on..."
			strcmdline = "cmd /c """""& systemdrive & "\program files\perc2\afa\afacli"" @" & delfastlocvar & "\cacheon.afa"""
			if debug = 1 then screenout "    (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, true)
		elseif checkfreespace(25) = 1 then
			screenout ""
			text = "There is not enough free space on the boot drive to install the FAST Utility!"
			screenout text
			screenout "Please free up 25 MB of disk space on the boot drive and re-run the script."
			genevent "E", "3", text
			hardware_perc2fast = 1
			exit function
		else
			screenout ""
			text = "The script could not locate the Fast Utility setup file!"
			screenout text
			screenout "Make sure the below file exists on the source that you are running the script from."
			screenout "  (" & delfastlocvar & "\setup.exe)"
			genevent "E", "3", text
			hardware_perc2fast = 1
			exit Function
		end if
	end if

	if err.number <> 0 then logerror "End of hardware_perc2fast Function", err.number : err.clear
	hardware_perc2fast = 0
End Function




REM #################################################################################################
REM ###                                    Recovery Console                                        ###
REM #################################################################################################
Function update_reccons()
	dim temp, strcmdline, text, servfile, filethere
	on error resume next

	REM ### Check for Recovery Console location ###
	screenout "Recovery Console:"
	if recconsvar = "off"  then
		screenout "  Recovery Console install is disabled."
		recconslogvar = "Disabled"
	elseif recconsvar = "switch"  then
		screenout "  Recovery Console install has been disabled by the user."
		recconslogvar = "Disabled - By User"
	elseif not srvbootvar = "NTFS" then
		screenout "  Recovery Console does not need installed on Non-NTFS partitions."
		recconslogvar = "Non-NTFS Boot Partition"
	else
		if not recconsfilevar = "" then
		elseif recconsfilevar = "" and (Wshfile.fileexists(bldlocvar & "\" & srvarcvar & "\autochk.exe")) then
			set servfile = Wshfile.getfile(bldlocvar & "\" & srvarcvar & "\autochk.exe")
			recconsfilevar = left(servfile.datelastmodified, Instr(1, servfile.datelastmodified, " ", 1)-1)
			wscript.DisconnectObject servfile
			set servfile=nothing
			if err.number <> 0 then err.clear
		else
			screenout ""
			text = "The script could not locate the Recovery Console check file!"
			screenout text
			screenout "  (" & bldlocvar & "\" & srvarcvar & "\autochk.exe" & ")"
			recconslogvar = "No Check File"
			genevent "E", "3", text
			update_reccons = 1
			exit function
		end if
		if (Wshfile.fileexists(recconschkvar)) then
			set servfile = Wshfile.getfile(recconschkvar)
			filethere = datevalue(left(servfile.datelastmodified, Instr(1, servfile.datelastmodified, " ", 1)-1))
			wscript.DisconnectObject servfile
			set servfile=nothing
			if err.number <> 0 then err.clear
		else
			filethere = "no"
		end if
		recconsfilevar = datevalue(recconsfilevar)

		if debug = 1 then screenout "  (Source File: " & recconsfilevar & ")"
		if debug = 1 then screenout "  (Check File: " & filethere & ")"

		if datecomp(recconsfilevar, filethere) and forceupdate = 1  then 
			screenout "  Updating Recovery Console..."
			recconslogvar = "Updated"
		elseif datecomp(recconsfilevar, filethere) then 
			screenout "  Recovery Console OK."
			recconslogvar = "OK"
			update_reccons = 0
			exit function
		elseif filethere = "no" then 
			screenout "  Installing Recovery Console..."
			recconslogvar = "Installed"
		else
			screenout "  Upgrading Recovery Console..."
			recconslogvar = "Upgraded"
		end if
		if (Wshfile.FileExists(recconslocvar)) then
			strcmdline = recconslocvar & " /cmdcons"
			if debug = 1 then screenout "  (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, false)
			screenout "      Waiting for Recovery Console to start..."
			waituntil "Windows 2000 Setup", ""
			screenout "      Starting Recovery Console file copy..."
			winmanipulate "Windows 2000 Setup", "Windows 2000 Setup", "Y"
			screenout "      Waiting for copy to complete..."
			waituntil "Microsoft Windows 2000 Advanced Server Setup", "winnt32.exe"
			wscript.sleep 25000 * netfactor
			screenout "      Closing out of Recovery Console..."
			winmanipulate "Microsoft Windows 2000 Advanced Server Setup", "Microsoft Windows 2000 Advanced Server Setup", "{ENTER}"
			if err.number <> 0 then err.clear

			REM ###### Set Reccons Registry entries ######	
			screenout "Setting Recovery Console registry entries..."
			registry(scriptinilocvar & "\RecconsRegistryW2K.ini")
		else
			screenout ""
			text = "The script could not locate the Recovery Console setup file!"
			screenout text
			screenout "  (" & recconslocvar & ")"
			recconslogvar = "No Install File"
			genevent "E", "3", text
			update_reccons = 1
			exit function
		end if
	end if

	if err.number <> 0 then logerror "End of update_reccons Function", err.number : err.clear
	update_reccons = 0
End Function




REM #################################################################################################
REM ###                                  Service Pack                                             ###
REM #################################################################################################
Function update_sp(rebootcheck)
	dim temp, strcmdline, text, srvsp, srvrc, scriptsp, scriptrc, backup, servfile, filethere
	on error resume next

	REM ### Install service Pack ###
	screenout "Service Pack 1:"
	if spvar = "off" then
		screenout "  Service Pack install is disabled at this time."
		splogvar = "Disabled"
	elseif spvar = "switch" then 
		screenout "  Service Pack install has been turned off by the user."
		splogvar = "Disabled - By User"
	elseif spdirvar = "SP0" then 
		screenout "  No Service Pack install at this time."
		splogvar = "None"
	else
		if Instr(1, spdirvar, "SP", 1) <> 0 and Instr(1, spdirvar, ".", 1) <> 0 then
			scriptsp = int(mid(spdirvar, 3, 1))
			scriptrc = int(right(spdirvar, 3))
		elseif Instr(1, spdirvar, "SP", 1) = 0 and Instr(1, spdirvar, ".", 1) <> 0 then
			scriptsp = 0
			scriptrc = int(right(spdirvar, 3))
		elseif Instr(1, spdirvar, "SP", 1) <> 0 and Instr(1, spdirvar, ".", 1) = 0 then
			scriptsp = int(mid(spdirvar, 3, 1))
			scriptrc = 0
		else
			scriptsp = 0
			scriptrc = 0
		end if
		if Instr(1, srvcsdvar, "Service Pack", 1) <> 0 and Instr(1, srvcsdvar, "RC", 1) <> 0 then
			srvsp = int(mid(srvcsdvar, 14, 1))
			srvrc = int(right(srvcsdvar, len(srvcsdvar) - Instr(1, srvcsdvar, ".", 1)))
		elseif Instr(1, srvcsdvar, "Service Pack", 1) = 0 and Instr(1, srvcsdvar, "RC", 1) <> 0 then
			srvsp = 0
			srvrc = int(right(srvcsdvar, len(srvcsdvar) - Instr(1, srvcsdvar, ".", 1)))
		elseif Instr(1, srvcsdvar, "Service Pack", 1) <> 0 and Instr(1, srvcsdvar, "RC", 1) = 0 then
			srvsp = int(mid(srvcsdvar, 14, 1))
			srvrc = 0
		else
			srvsp = 0
			srvrc = 0
		end if
		if (Wshfile.fileexists(spchkvar)) then
			set servfile = Wshfile.getfile(spchkvar)
			filethere = datevalue(left(servfile.datelastmodified, Instr(1, servfile.datelastmodified, " ", 1)-1))
			wscript.DisconnectObject servfile
			set servfile=nothing
			if err.number <> 0 then err.clear
		else
			filethere = "no"
		end if
		if (Wshfile.folderexists(systemroot & "\$NtServicePackUninstall$")) then 
			backup = "yes"
		else
			backup = "no"
		end if
		if debug = 1 then screenout "  (Script SP: " & scriptsp & ")"
		if debug = 1 then screenout "  (Server SP: " & srvsp & ")"
		if debug = 1 then screenout "  (SP Backup: " & backup & ")"
		if debug = 1 then screenout "  (SP File: " & filethere & ")"

 
		if srvsp > scriptsp then 
			screenout ""
			text = "This server is all ready running Service Pack " & srvsp & " which is newer then what this script supports!"
			screenout text
			screenout "You must run an update script that supports this Service Pack or remove the current SP."
			screenout ""
			genevent "E", "3", text
			update_sp = 1
			exit function
		elseif srvsp = scriptsp and not backup = "no" then 
			screenout ""
			text = "This server is all ready running Service Pack " & srvsp & " with a backup directory!"
			screenout text
			screenout "You must remove the current SP before running the update script again."
			screenout "You can remove the current SP by using Control Panel - Add/remove programs."
			screenout ""
			genevent "E", "3", text
			update_sp = 1
			exit function
		elseif rebootcheck then 
			screenout "  The reboot check for Service Pack " & scriptsp & " is OK."
			update_sp = 0
			exit function
		elseif srvsp = scriptsp and datecomp(spfilevar, filethere) and forceupdate = 1 then 
			screenout "  Updating Service Pack " & scriptsp & "..."
			splogvar = "Updated"
			update_sp = 0
			exit function
		elseif srvsp = scriptsp and datecomp(spfilevar, filethere) then 
			screenout "  Service Pack " & scriptsp & " OK."
			splogvar = "OK"
			update_sp = 0
			exit function
		elseif srvsp = 0 then 
			screenout "  Installing Service Pack " & scriptsp & "..."
			splogvar = "Installed"
		else
			screenout "  Upgrading to Service Pack " & scriptsp & "..."
			splogvar = "Upgraded"
		end if

		if (Wshfile.FileExists(splocvar & "\" & srvarcvar & "\update\update.exe")) and checkfreespace(60) = 0 then
			strcmdline = splocvar & "\" & srvarcvar & "\update\update.exe -u -n -o -z"
			rem strcmdline = splocvar & "\" & srvarcvar & "\update\update.exe -u -o -z"
			if debug = 1 then screenout "  (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, true)
		elseif checkfreespace(60) = 1 then
			screenout ""
			text = "There is not enough free space on the boot drive to install the Service Pack!"
			screenout text
			screenout "Please free up 60 MB of disk space on the boot drive and re-run the script."
			genevent "E", "3", text
			update_sp = 1
			exit function
		else
			screenout ""
			text = "The script could not locate the Service Pack setup file!"
			screenout text
			screenout "Make sure the below file exists on the source that you are running the script from."
			screenout "  (" & splocvar & "\" & srvarcvar & "\update\update.exe)"
			genevent "E", "3", text
			update_sp = 1
			exit Function
		end if
	end if

	if err.number <> 0 then logerror "End of update_sp Function", err.number : err.clear
	update_sp = 0
End Function



REM #################################################################################################
REM ###                                  Symbols                                                  ###
REM #################################################################################################
Function update_sym()
	dim text, wshinifile, file, strcmdline, temp
	on error resume next

	REM ### Install Symbols ###
	screenout "Symbols:"
	if (Wshfile.folderexists(systemroot & "\symbolsold")) then wshfile.deletefolder systemroot & "\symbolsold", TRUE
	if err.number <> 0 then err.clear
	if symvar = "off" then
		screenout "  Symbol install is disabled."
		symlogvar="Disabled"
	elseif symvar = "switch" then
		screenout "  Symbol install has been disabled by the user."
		symlogvar = "Disabled - By User"
	elseif symvar = "no" then
		screenout "  Symbol install not needed at this site."
		symlogvar = "Disabled - By Site"
	else
		screenout "  Removing Windows 2000 Symbol directories..."
		if (Wshfile.folderexists(systemroot & "\symbols\acm")) then wshfile.deletefolder systemroot & "\symbols\acm", TRUE
		if (Wshfile.folderexists(systemroot & "\symbols\ax")) then wshfile.deletefolder systemroot & "\symbols\ax", TRUE
		if (Wshfile.folderexists(systemroot & "\symbols\cfm")) then wshfile.deletefolder systemroot & "\symbols\cfm", TRUE
		if (Wshfile.folderexists(systemroot & "\symbols\cnv")) then wshfile.deletefolder systemroot & "\symbols\cnv", TRUE
		if (Wshfile.folderexists(systemroot & "\symbols\cpl")) then wshfile.deletefolder systemroot & "\symbols\cpl", TRUE
		if (Wshfile.folderexists(systemroot & "\symbols\ds")) then wshfile.deletefolder systemroot & "\symbols\ds", TRUE
		if (Wshfile.folderexists(systemroot & "\symbols\flg")) then wshfile.deletefolder systemroot & "\symbols\flg", TRUE
		if (Wshfile.folderexists(systemroot & "\symbols\io")) then wshfile.deletefolder systemroot & "\symbols\io", TRUE
		if (Wshfile.folderexists(systemroot & "\symbols\scr")) then wshfile.deletefolder systemroot & "\symbols\scr", TRUE
		if (Wshfile.folderexists(systemroot & "\symbols\tsp")) then wshfile.deletefolder systemroot & "\symbols\tsp", TRUE
		if (Wshfile.folderexists(systemroot & "\symbols\usa")) then wshfile.deletefolder systemroot & "\symbols\usa", TRUE
		if (Wshfile.folderexists(systemroot & "\symbols\wpc")) then wshfile.deletefolder systemroot & "\symbols\wpc", TRUE
		if (Wshfile.folderexists(systemroot & "\symbols\16bit")) then wshfile.deletefolder systemroot & "\symbols\16bit", TRUE
		if (Wshfile.folderexists(systemroot & "\symbols\winnt32")) then wshfile.deletefolder systemroot & "\symbols\winnt32", TRUE
		if (Wshfile.folderexists(systemroot & "\symbols\noexport")) then wshfile.deletefolder systemroot & "\symbols\noexport", TRUE
		if err.number <> 0 then logerror "Deleting Symbols directories", err.number : err.clear
		if (Wshfile.fileexists(scriptinilocvar & "\SymbolFilesW2K.ini")) then 
			Set wshinifile = wshfile.OpenTextFile(scriptinilocvar & "\SymbolFilesW2K.ini", 1)
			screenout "  Removing Windows 2000 Symbol files..."
			Do While wshinifile.AtEndOfStream <> true
				file = wshiniFile.ReadLine
				if (Wshfile.FileExists(systemroot & file)) then
					rem screenout "    " & file
					wshfile.deletefile systemroot & file, TRUE
					if err.number <> 0 then err.clear
				end if
			loop
			wshinifile.Close
		else
			screenout ""
			text = "The script was unable to find the SymbolFilesW2K.ini file!"
			screenout text
			screenout "Make sure that " & scriptinilocvar & " exists and contains the necessary INI file."
			genevent "E", "3", text
			update_sym = 1
			exit function
		end if
		if symvar = "clear" then
			screenout "  Symbols have been cleared by the user."
			symlogvar = "Cleared - By User"
		elseif scriptsitevar = "DESK" then
			screenout "  Symbols do not need installed on Desk Top machines."
			symlogvar = "Desk Top Machine"
		elseif symvar = "full" and checkfreespace(750) = 0 then
			if not (Wshfile.folderexists(systemroot & "\symbols")) then wshfile.createfolder systemroot & "\symbols", TRUE
			screenout "  Starting FULL Symbol set copy..."
			symlogvar="Installed FULL set"
			strcmdline = "cmd /c ""xcopy /cdefhkrvyz " & symfulllocvar & " " & systemroot & "\symbols\"""
	       		if debug = 1 then screenout "    (" & strcmdline & ")"
               		temp = wshshell.Run(strcmdline, 1, true)
			if temp <> 0 then 
				screenout ""
				text = "The script was unable to complete the symbol files copy to the " & systemroot & "\symbols directory!"
				screenout text
				screenout "Make sure the directory or a file in the directory is not in use by an application."
				screenout "Do not close the Command Prompt window that is copying files."
				screenout "Try removing the symbols directory from the server and run the script again."
				genevent "E", "3", text
				update_sym = 1
				exit function
			end if
		elseif checkfreespace(80) = 0  then
			if not (Wshfile.folderexists(systemroot & "\symbols")) then wshfile.createfolder systemroot & "\symbols", TRUE
			screenout "  Starting SMALL Symbol set copy..."
			symlogvar="Installed SMALL set"
			strcmdline = "cmd /c ""xcopy /cdefhkrvyz " & symlocvar & " " & systemroot & "\symbols\"""
	       		if debug = 1 then screenout "    (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, true)
			if temp <> 0 then 
				screenout ""
				text = "The script was unable to complete the symbol files copy to the " & systemroot & "\symbols directory!"
				screenout text
				screenout "Make sure the directory or a file in the directory is not in use by an application."
				screenout "Do not close the Command Prompt window that is copying files."
				screenout "Try removing the symbols directory from the server and run the script again."
				genevent "E", "3", text
				update_sym = 1
				exit function
			end if
			if not wwwvar = "" then screenout "  Starting IIS Symbol copy..."
			if not wwwvar = "" then symlogvar="Installed SMALL set and IIS small set"
			if not wwwvar = "" then strcmdline = "cmd /c ""xcopy /cdefhkrvyz " & wwwsymlocvar & " " & systemroot & "\symbols\"""
	       		if not wwwvar = "" and debug = 1 then screenout "    (" & strcmdline & ")"
			if not wwwvar = "" then temp = wshshell.Run(strcmdline, 1, true)
			if not wwwvar = "" and err.number <> 0 then err.clear
		else
			screenout ""
			text = "There is not enough free space on the boot drive to install symbols!"
			screenout text
			screenout "Please free up 80 MB on the boot drive for the small symbol set (750 MB for the Full symbols set)."
			genevent "E", "3", text
			update_sym = 1
			exit function
		end if
	end if

	if err.number <> 0 then logerror "End of update_sym Function", err.number : err.clear
	update_sym = 0
End Function



REM #################################################################################################
REM ###                                  Hotfixes                                                 ###
REM #################################################################################################
Function update_fix()
	dim temp, strcmdline, text, wshinifile, rkey, dchotfixes
	on error resume next

	REM ### install Hotfixes ###
	screenout "Hotfixes:"
	if fixvar = "off" then
		screenout "  Hotfix install is disabled."
		fixlogvar = "Disabled"
	elseif fixvar = "switch" then
		screenout "  Hotfix install has been turned off by the user."
		fixlogvar = "Disabled - By User"
	else
		screenout "  Clearing IPAK tracking registry entries (For Tools and Update script)..."
		if (Wshfile.fileexists(scriptinilocvar & "\DeleteRegKeysW2K.ini")) then 
			Set wshinifile = wshfile.OpenTextFile(scriptinilocvar & "\DeleteRegKeysW2K.ini", 1)
			Do While wshinifile.AtEndOfStream <> true
				rkey = wshiniFile.ReadLine
				wshshell.regdelete rkey
				if err.number <> 0 then err.clear
			loop
			wshinifile.Close
		else
			screenout ""
			text = "The script was unable to find the DeleteRegKeysW2K.ini file!"
			screenout text
			screenout "Make sure that " & scriptinilocvar & " exists and contains the necessary INI file."
			genevent "E", "3", text
			update_fix = 1
			exit function
		end if
		screenout "  Clearing Hotfix tracking registry entries (For srvinfo)..."
		if (Wshfile.fileexists(scriptinilocvar & "\DelhotfixRegKeysW2K.ini")) then 
			Set wshinifile = wshfile.OpenTextFile(scriptinilocvar & "\DelhotfixRegKeysW2K.ini", 1)
			Do While wshinifile.AtEndOfStream <> true
				rkey = wshiniFile.ReadLine
				rem if debug = 1 then screenout "    (" & rkey & ")"
				strCmdLine = "cmd /c """& scriptfileslocvar & "\reg delete """ & rkey & """ /f"""
				rem if debug = 1 then screenout "    (" & strcmdline & ")"
				temp = wshshell.Run(strcmdline, 0, true)
				if err.number <> 0 then err.clear
			loop
			wshinifile.Close
		else
			screenout ""
			text = "The script was unable to find the DelhotfixRegKeysW2K.ini file!"
			screenout text
			screenout "Make sure that " & scriptinilocvar & " exists and contains the necessary INI file."
			genevent "E", "3", text
			update_fix = 1
			exit function
		end if
		if (Wshfile.FileExists(fixlocvar & "\hotfix.exe")) then
			screenout "  Installing Hotfixes..."
			fixlogvar = "Installed"
			strcmdline = fixlocvar & "\hotfix.exe -n -z -m"
	       		if debug = 1 then screenout "    (" & strcmdline & ")"
               		temp = wshshell.Run(strcmdline, 1, true)
		else
			screenout "  No Hotfixes at this time."
			fixlogvar = "None"
		end if
		if (Wshfile.FileExists(fixlocvar & "\hftrack.exe")) then
			screenout "  Installing Hotfix and IPAK tracking registry entries..."
			strcmdline = fixlocvar & "\hftrack.exe"
	       		if debug = 1 then screenout "    (" & strcmdline & ")"
               		temp = wshshell.Run(strcmdline, 1, true)
		else
			screenout ""
			text = "The script could not locate the HfTrack setup file!"
			screenout text
			screenout "Make sure that the HfTrack.exe file is located at the below location."
			screenout "    (" & fixlocvar & "\hftrack.exe)"
			genevent "E", "3", text
			update_fix = 1
			exit function
		end if
		dchotfixes = scriptpathlocvar & "\" & srvbuildvar & "\" & srvsuitevar & "\" & srvarcvar & "\" & spdirvar & "\hotfix\dcfixes"
		if argument("/allfixes") <> 0 and (Wshfile.FileExists(scriptcmdlocvar & "\allfixes.vbs")) then
			screenout "  Installing All additional Hotfixes..."
			strcmdline = "cmd /c """ & scriptcmdlocvar & "\allfixes.vbs"""
	       		if debug = 1 then screenout "    (" & strcmdline & ")"
               		temp = wshshell.Run(strcmdline, 1, true)
		elseif srvrolevar = "Domain Controller" and (Wshfile.FileExists(dchotfixes & "\setup.cmd")) then
			screenout "  Installing additional DC Hotfixes..."
			strcmdline = "cmd /c """ & dchotfixes & "\setup.cmd"""
	       		if debug = 1 then screenout "    (" & strcmdline & ")"
               		temp = wshshell.Run(strcmdline, 1, true)
		elseif argument("/allfixes") <> 0 then
			screenout "  No additional Hotfixes."
		end if
	end if

	if err.number <> 0 then logerror "End of update_fix Function", err.number : err.clear
	update_fix = 0
End Function



REM #################################################################################################
REM ###                                  128 Bit encryption                                       ###
REM #################################################################################################
Function update_128()
	dim temp, strcmdline, text
	on error resume next

	REM ### Run 128 bit encryption script ###
	screenout "128 bit encryption:"
	if encvar="off" then
		screenout "  128 Bit encryption is disabled."
		enclogvar = "Disabled"
	elseif encvar= "switch" then
		screenout "  128 Bit encryption disabled by the user."
		enclogvar = "Disabled - By User"
	else
		if (Wshfile.FileExists(enclocvar)) then
			screenout "  Starting 128 Bit Encryption IPAK script..."
			enclogvar = "Executed"
			strcmdline = "cmd /c """ & enclocvar & " /noreb"""
			if debug = 1 then screenout "  (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, true)
		else
			screenout ""
			text = "The script could not locate the 128 Bit encryption script!"
			screenout text
			screenout "Make sure that this site supports 128 bit encryption."
			screenout "Make sure that the 128 bit encryption script has been replicated to this server."
			screenout "  (" & enclocvar & ")"
			enclogvar = "No Script"
			genevent "E", "3", text
			update_128 = 1
			exit function
		end if
	end if

	if err.number <> 0 then logerror "End of update_128 Function", err.number : err.clear
	update_128 = 0
End Function





REM #################################################################################################
REM ###                                       Localbin Files                                      ###
REM #################################################################################################
Function file_localbin()
	dim strcmdline, temp, text, servfile, filethere, bincheckfile
	on error resume next

	screenout "Localbin Files:"
	if (Wshfile.folderexists(systemdrive & "\localold")) then wshfile.deletefolder systemdrive & "\localold", TRUE
	if err.number <> 0 then err.clear
	if (Wshfile.fileexists(scriptbinlocvar & "\binupdate.log")) then
		set servfile = Wshfile.getfile(scriptbinlocvar & "\binupdate.log")
		bincheckfile = left(servfile.datelastmodified, Instr(1, servfile.datelastmodified, " ", 1)-1)
		wscript.DisconnectObject servfile
		set servfile=nothing
		if err.number <> 0 then err.clear
	else
		bincheckfile = "no"
	end if
	if (Wshfile.fileexists(scriptbinvar & "\binupdate.log")) then
		set servfile = Wshfile.getfile(scriptbinvar & "\binupdate.log")
		filethere = datevalue(left(servfile.datelastmodified, Instr(1, servfile.datelastmodified, " ", 1)-1))
		wscript.DisconnectObject servfile
		set servfile=nothing
		if err.number <> 0 then err.clear
	else
		filethere = "no"
	end if

	if debug = 1 then screenout "  (Source File: " & bincheckfile & ")"
	if debug = 1 then screenout "  (Check File: " & filethere & ")"

	if datecomp(bincheckfile, filethere) and forceupdate = 0 then 
		screenout "  Localbin files OK."
	else
		REM #### Purge and replace localbin files ######
		screenout "  Removing the " & scriptbinvar & " directory..."
		if (Wshfile.folderexists(scriptbinvar)) then wshfile.deletefolder scriptbinvar, TRUE
		if err.number <> 0 then
			err.clear
			screenout "    Folder could not be removed."
		end if
		if checkfreespace(15) = 1 then
			screenout ""
			text = "There is not enough free space on the boot drive to install the localbin files!"
			screenout text
			screenout "Please free up at least 15 MB of disk space on the boot drive for the localbin files."
			genevent "E", "3", text
			file_localbin = 1
			exit function
		end if
		if not (Wshfile.folderexists(scriptbinvar)) then wshfile.createfolder(scriptbinvar)
		screenout "  Starting copy of files to " & scriptbinvar & "..."
		strcmdline = "cmd /c ""xcopy /cdefhkrvyz " & scriptbinlocvar & " " & scriptbinvar & "\"""
		if debug = 1 then screenout "    (" & strcmdline & ")"
		temp = wshshell.Run(strcmdline, 1, true)
		if temp <> 0 then 
			screenout ""
			text = "The script was unable to complete the localbin files copy to the " & scriptbinvar & " directory!"
			screenout text
			screenout "Make sure the directory or a file in the directory is not in use by an application."
			screenout "Do not close the Command Prompt window that is copying files."
			screenout "Try removing the localbin directory from the server and run the script again."
			genevent "E", "3", text
			file_localbin = 1
			exit function
		end if
	end if

	REM ###### Timesync.cmd to bin directory ######
	screenout "  Updating Timesync.cmd in " & scriptbinvar & "..."
	if (Wshfile.fileexists(scriptbinvar & "\timesync.cmd")) then wshfile.deletefile scriptbinvar & "\timesync.cmd", TRUE
	if err.number <> 0 then err.clear
	if not (Wshfile.fileexists(scriptbinvar & "\timesync.cmd")) then wshfile.copyfile scriptcmdlocvar & "\timesync.cmd", scriptbinvar & "\", TRUE
	if err.number <> 0 then logerror "Copying TimeSync.cmd", err.number : err.clear

	REM ###### ITGBackup.vbs to bin directory ######
	screenout "  Updating ITGBackup.vbs in " & scriptbinvar & "..."
	if (Wshfile.fileexists(scriptbinvar & "\ITGBackup.vbs")) then wshfile.deletefile scriptbinvar & "\ITGBackup.vbs", TRUE
	if err.number <> 0 then err.clear
	if not (Wshfile.fileexists(scriptbinvar & "\ITGBackup.vbs")) then wshfile.copyfile scriptcmdlocvar & "\ITGBackup.vbs", scriptbinvar & "\", TRUE
	if err.number <> 0 then logerror "Copying ITGBackup.vbs", err.number : err.clear

	REM ###### BackGround.cmd to bin directory ######
	screenout "  Updating BackGround.cmd in " & scriptbinvar & "..."
	if (Wshfile.fileexists(scriptbinvar & "\BackGround.cmd")) then wshfile.deletefile scriptbinvar & "\BackGround.cmd", TRUE
	if err.number <> 0 then err.clear
	if not (Wshfile.fileexists(scriptbinvar & "\BackGround.cmd")) then wshfile.copyfile scriptcmdlocvar & "\BackGround.cmd", scriptbinvar & "\", TRUE
	if err.number <> 0 then logerror "Copying BackGround.cmd", err.number : err.clear

	if err.number <> 0 then logerror "End of file_localbin Function", err.number : err.clear
	file_localbin = 0
End Function





REM #################################################################################################
REM ###                                       Exchange Files                                      ###
REM #################################################################################################
Function file_exchange()
	dim text
	on error resume next


	REM ###### create d:\ntdumps on exchange servers ######
	if not exchangevar = "" then
		screenout "Checking file changes for Exchange services..."
		if Instr(1, harddrives, "D:", 1) <> 0 then
			if not (Wshfile.folderexists("d:\ntdumps")) then
				screenout "  Creating D:\ntdumps directory..."
				wshfile.createfolder("d:\ntdumps")
				if err.number <> 0 then err.clear
			else
				screenout "  D:\ntdumps is all created."
			end if
			if not (Wshfile.folderexists("d:\exchsrvr")) then
				screenout "  Creating D:\exchsrvr directory..."
				wshfile.createfolder("d:\exchsrvr")
				if err.number <> 0 then err.clear
			else
				screenout "  D:\exchsrvr is all created."
			end if
		else
			screenout ""
			text = "Unable to create Exchange dump file directory as D: is not currently a valid hard drive!"
			screenout text
			screenout "Make sure that the D: drive is partitioned and formatted NTFS."
			genevent "E", "3", text
			file_localbin = 1
			exit function
		end if
	else
		screenout "File changes for Exchange services not needed."
	end if

	if err.number <> 0 then logerror "End of file_exchange Function", err.number : err.clear
	file_exchange = 0
End Function





REM #################################################################################################
REM ###                                       Boot.ini file                                       ###
REM #################################################################################################
Function file_boot()
	dim strcmdline, temp, text
	dim wshgetbootfile, attribute
	dim wshreadbootfile, wshnewbootfile, change, line, addon, newline, first, last
	dim wshgetnewbootfile
	on error resume next

	REM ###### Change Boot.ini file for the right architecture ######
	if procarc = "x86" then
		screenout "Checking i386 Boot.ini..."
		REM ###### Removes oldboot.ini file and set attributes so file can be read ######
		if (wshfile.fileexists(systemdrive & "\boot.ini")) then 
			set wshgetbootfile = wshfile.getfile(systemdrive & "\boot.ini")
			attribute = wshgetbootfile.attributes
			wshgetbootfile.attributes = 0
		else
			screenout "  Cannot find Boot.ini file!"
		end if

		REM ###### Open files, read, change if needed, close files #######
		if (wshfile.fileexists(systemdrive & "\boot.ini")) then 
			set wshreadbootfile = wshfile.OpenTextFile(systemdrive & "\boot.ini", 1)
			set wshnewbootfile = Wshfile.createtextfile(systemdrive & "\newboot.ini",1)
			change = 0
			addon = ""
			backuproot = ""
			Do While wshreadbootfile.AtEndOfStream <> true
				line = wshreadbootfile.ReadLine
				if Instr(1, line, "Microsoft Windows 2000 Advanced Server", 1) <> 0 then
					if Instr(1, line, "/sos", 1) = 0 then
						screenout "  Adding /sos..."
						change = 1
						addon = " /sos"
					end if
					if Instr(1, line, "/debug", 1) = 0 and not scriptsitevar = "DESK" then
						screenout "  Adding /debug..."
						change = 1
						addon = addon & " /debug"
					end if
					if Instr(1, line, "/baudrate=57600", 1) = 0 and not scriptsitevar = "DESK" then
						screenout "  Adding /baudrate=57600..."
						change = 1
						addon = addon & " /baudrate=57600"
					end if
					if Instr(1, line, "/debugport=com1", 1) = 0 and srvhardtypevar = "6400R" then
						screenout "  Adding /debugport=com1..."
						change = 1
						addon = addon & " /debugport=com1"
					end if
					wshnewbootfile.writeline(line & addon)
				elseif Instr(1, line, "timeout", 1) <> 0 then
					if Instr(1, line, "=10", 1) = 0 then
						screenout "  Changing timeout to 10 seconds..."
						change=1
						line = left(line,8) & "10"
					end if
					wshnewbootfile.writeline(line)
				elseif Instr(1, line, "backup", 1) <> 0 then
					screenout "  Removing Backup build entry..."
					change = 1
					first = Instr(1, line, "\", 1)
					last = Instr(1, line, "=", 1)
					backuproot = systemdrive & mid(line, first, last-first)
				elseif Instr(1, line, systemdrive & "\=", 1) <> 0 and srvbootvar = "NTFS" then
					screenout "  Removing MS-DOS or Microsoft Windows option..."
					change = 1
				else
					wshnewbootfile.writeline(line)
				end if
			Loop
			wshreadbootfile.close
			wshnewbootfile.close
			wscript.DisconnectObject wshreadbootfile
			wscript.DisconnectObject wshnewbootfile
			set wshreadbootfile=nothing
			set wshnewbootfile=nothing
		end if

		REM ###### if new file was changed then rename files, otherwise delete new file ######
		if change = 1 then
			screenout "  Changing to new boot.ini file..."
			if (wshfile.fileexists(systemdrive & "\oldboot.ini")) then wshfile.deletefile systemdrive & "\oldboot.ini", TRUE
			wshgetbootfile.name = "oldboot.ini"
			set wshgetnewbootfile = wshfile.getfile(systemdrive & "\newboot.ini")
			wshgetnewbootfile.name = "boot.ini"
			wshgetnewbootfile.attributes = attribute
			if err.number <> 0 then err.clear
		else
			screenout "  Boot.ini file OK."
			wshgetbootfile.attributes = attribute
			wshfile.deletefile systemdrive & "\newboot.ini", TRUE
			if err.number <> 0 then err.clear
		end if
		if not backuproot = "" then
			screenout "  (Backup build directory is " & backuproot & ".)"
		end if
	elseif srvarcvar = "alpha" then
		screenout "Changing Boot options for Alpha servers..."
		strCmdLine = "cmd /c """& scriptfileslocvar & "\nvram /set OSLOADOPTIONS = ""/sos /debug"""""
		temp = wshshell.Run(strcmdline, 0, true)
		strCmdLine = "cmd /c """& scriptfileslocvar & "\nvram /set COUNTDOWN = ""10"""""
		temp = wshshell.Run(strcmdline, 0, true)
	else
		screenout "Changes to the boot.ini for this architecture are not supported."
	end if

	if err.number <> 0 then logerror "End of file_boot Function", err.number : err.clear
	file_boot = 0
End Function




REM #################################################################################################
REM ###                                       Delete files                                        ###
REM #################################################################################################
Function file_delete()
	dim wshinifile, file, wshgettempfile
	on error resume next

	REM ###### Remove attributes of autoexec.bat and config.sys file ######
	if (wshfile.fileexists(systemdrive & "\autoexec.bat")) then 
		set wshgettempfile = wshfile.getfile(systemdrive & "\autoexec.bat")
		attribute = wshgettempfile.attributes
		wshgettempfile.attributes = 0
	end if
	if (wshfile.fileexists(systemdrive & "\config.sys")) then 
		set wshgettempfile = wshfile.getfile(systemdrive & "\config.sys")
		attribute = wshgettempfile.attributes
		wshgettempfile.attributes = 0
	end if
	if (Wshfile.FileExists(systemdrive & "\autoexec.bat")) then wshfile.deletefile systemdrive & "\autoexec.bat", TRUE
	if (Wshfile.FileExists(systemdrive & "\config.sys")) then wshfile.deletefile systemdrive & "\config.sys", TRUE

	REM ###### Remove install, promotion, and addon files ######
	if (Wshfile.FileExists(systemdrive & "\answer.txt")) then wshfile.deletefile systemdrive & "\answer.txt", TRUE
	if (Wshfile.FileExists(systemdrive & "\winnt.sif")) then wshfile.deletefile systemdrive & "\winnt.sif", TRUE
	if (Wshfile.FileExists(systemdrive & "\addon.inf")) then wshfile.deletefile systemdrive & "\addon.inf", TRUE
	if (Wshfile.FileExists(systemdrive & "\dcpromo-ntupg.inf")) then wshfile.deletefile systemdrive & "\dcpromo-ntupg.inf", TRUE
	if (Wshfile.FolderExists(systemdrive & "\ESP")) then wshfile.deletefolder systemdrive & "\ESP", TRUE
	if (Wshfile.FolderExists(systemroot & "\itg")) then wshfile.deletefolder systemroot & "\itg", TRUE
	if err.number <> 0 then err.clear


	REM ###### Delete Files ######	
	if (Wshfile.fileexists(scriptinilocvar & "\DeleteFilesW2K.ini")) then 
		Set wshinifile = wshfile.OpenTextFile(scriptinilocvar & "\DeleteFilesW2K.ini", 1)
		screenout "Deleting unwanted file(s)..."
		Do While wshinifile.AtEndOfStream <> true
			file = wshiniFile.ReadLine
			if (Wshfile.FileExists(file)) then wshfile.deletefile file, TRUE
			if err.number <> 0 then err.clear
		loop
		wshinifile.Close
	else
		screenout ""
		text = "The script was unable to find the DeleteFilesW2K.ini file!"
		screenout text
		screenout "Make sure that " & scriptinilocvar & " exists and contains the necessary INI file."
		genevent "E", "3", text
		file_delete = 1
		exit function
	end if

	wscript.DisconnectObject wshinifile
	set wshinifile=nothing

	if err.number <> 0 then logerror "End of file_delete Function", err.number : err.clear
	file_delete = 0
End Function




REM #################################################################################################
REM ###                                  Backgroup Bitmap File                                    ###
REM #################################################################################################
Function file_bitmap()
	dim strcmdline, temp, background
	on error resume next

	REM ### Check Background Bitmap stuff ###
	if argument("/nobitmap") <> 0 then
		screenout "Background bitmap changes have been disabled by the user."
	else
		screenout "Checking background bitmap settings..."
		background = WshShell.RegRead("HKEY_USERS\.default\Control Panel\desktop\wallpaper")
		if err.number <> 0 then logerror "Reading ITG background bitmap registry location", err.number : err.clear
		if cstr(background) = systemroot & "\ITGServerName.bmp" then
			screenout "  ITG background bitmap registry location OK."
		else
			screenout "  Setting ITG background bitmap registry location..."
			wshshell.regwrite "HKEY_USERS\.default\Control Panel\desktop\wallpaper", systemroot & "\ITGServerName.bmp", "REG_SZ"
			wshshell.regwrite "HKEY_USERS\.default\Control Panel\desktop\tilewallpaper", "1", "REG_SZ"
			if err.number <> 0 then logerror "Writing ITG background bitmap registry location", err.number : err.clear
			wshshell.regwrite "HKCU\Control Panel\desktop\wallpaper", "(None)", "REG_SZ"
			wshshell.regwrite "HKCU\Control Panel\desktop\tilewallpaper", "0", "REG_SZ"
			if err.number <> 0 then logerror "Clearing users background registry location", err.number : err.clear
			logregistryentry "ITG background bitmap registry location", "Updated", "HKEY_USERS\.default\Control Panel\desktop\wallpaper", systemroot & "\ITGServerName.bmp"
		end if
		if termvar = "TermRa" then
			screenout "  Setting ITG background bitmap registry location for TS-RA..."
			wshshell.regwrite "HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp\UserOverride\Control Panel\desktop\wallpaper", systemroot & "\ITGServerName.bmp", "REG_SZ"
			wshshell.regwrite "HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp\UserOverride\Control Panel\desktop\tilewallpaper", "0", "REG_SZ"
			if err.number <> 0 then logerror "Writing TS Background bitmap registry location", err.number : err.clear
		elseif termvar = "TermApp" then
			screenout "  Clearing ITG background bitmap registry location for TS-APP..."
			wshshell.regwrite "HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp\UserOverride\Control Panel\desktop\wallpaper", "(None)", "REG_SZ"
			if err.number <> 0 then logerror "Clearing TS Background bitmap registry location", err.number : err.clear
		end if
		if (Wshfile.FileExists(scriptpathlocvar & "\bin\ITGSN.bmp")) then
			screenout "  Updating ITG background bitmap on server..."
			if (Wshfile.FileExists(systemroot & "\ITGSN.bmp")) then wshfile.deletefile systemroot & "\ITGSN.bmp"
			wshfile.copyfile scriptpathlocvar & "\bin\ITGSN.bmp", systemroot & "\ITGSN.bmp", TRUE
		else
			screenout "  Cannot find the source background bitmap at " & scriptpathlocvar & "\bin\ITGSN.bmp!"
		end if
		if (Wshfile.FileExists(systemroot & "\ITGSN.bmp")) then
			screenout "  Updating ITG background bitmap with the server name..."
			if (Wshfile.FileExists(systemroot & "\ITGServerName.bmp")) then wshfile.deletefile systemroot & "\ITGServerName.bmp"
			strcmdLine = "cmd /c """& scriptfileslocvar & "\bmpedit -m " & computername & " -v ""IPAK " & scriptipakvar & """ -r 255 -g 255 -b 255 -f1 28 -f2 12 -i " & systemroot & "\ITGSN.bmp -o " & systemroot & "\ITGServerName.bmp"""
			temp = wshshell.Run(strcmdline, 0, true)
		else
			screenout "  The source background bitmap did not get copied to the server!"
		end if
		if not (Wshfile.FileExists(systemroot & "\ITGServerName.bmp")) then
			screenout "  The source background bitmap did not get created on the server!"
		end if
	end if

	if err.number <> 0 then logerror "End of file_bitmap Function", err.number : err.clear
	file_bitmap = 0
End Function




REM #################################################################################################
REM ###                                      Registry                                             ###
REM #################################################################################################
Function registry(filetoprocess)
	dim text, wshinifile, space, rtext, rkey, rwhat, rvalue, rtype, temp, exists, correct
	on error resume next

	REM ###### Set Registry values ######	
	if (Wshfile.FileExists(filetoprocess)) then
		Set wshinifile = wshfile.OpenTextFile(filetoprocess, 1)
		if err.number <> 0 then logerror "Opening Registry INI file", err.number : err.clear
		Do While wshinifile.AtEndOfStream <> true
			space = wshiniFile.ReadLine
			rtext = wshinifile.readline
			rwhat = wshinifile.ReadLine
			rkey = wshiniFile.ReadLine
			rvalue = wshiniFile.ReadLine
			rtype = wshiniFile.ReadLine
			temp = WshShell.RegRead(rkey)
			if err.number = 0 then
				exists = 1
			else
				exists = 0
				if err.number <> 0 then err.clear
			end if
			if cstr(temp) = cstr(rvalue) then
				correct = 1
			else
				correct = 0
			end if
			if rwhat = "create" and exists = 0 then
				screenout "  Registry entry being created...  (" & rtext & ")"
				wshshell.regwrite rkey, rvalue, rtype
				logregistryentry rtext, "Created", rkey, rvalue
				if err.number <> 0 then logerror "Writing to registry", err.number : err.clear
			elseif rwhat = "update" and exists = 1 and correct = 0 then
				screenout "  Registry entry being corrected...  (" & rtext & ")"
				wshshell.regwrite rkey, rvalue, rtype
				logregistryentry rtext, "Corrected", rkey, rvalue
				if err.number <> 0 then logerror "Writing to registry", err.number : err.clear
			elseif rwhat = "set" and exists = 0 then
				screenout "  Registry entry being created...  (" & rtext & ")"
				wshshell.regwrite rkey, rvalue, rtype
				logregistryentry rtext, "Created", rkey, rvalue
				if err.number <> 0 then logerror "Writing to registry", err.number : err.clear
			elseif rwhat = "set" and exists = 1 and correct = 0 then
				screenout "  Registry entry being corrected...  (" & rtext & ")"
				wshshell.regwrite rkey, rvalue, rtype
				logregistryentry rtext, "Corrected", rkey, rvalue
				if err.number <> 0 then logerror "Writing to registry", err.number : err.clear
			elseif rwhat = "remove" and exists = 1 then
				screenout "  Registry entry being removed...  (" & rtext & ")"
				wshshell.regdelete rkey
				logregistryentry rtext, "Removed", rkey, rvalue
				if err.number <> 0 then logerror "Writing to registry", err.number : err.clear
			elseif rwhat = "" or not (rwhat = "create" or rwhat = "update" or rwhat = "set" or rwhat = "remove") then 
				screenout "  Invalid Command!  (" & rwhat & ")" 
			else
				screenout "  Registry entry OK.  (" & rtext & ")"
			end if
		Loop
		wshinifile.Close
		wscript.DisconnectObject wshinifile
		set wshinifile=nothing
	else
		screenout ""
		text = "The script was unable to find the " & filetoprocess & " file!"
		screenout text
		screenout "Make sure that the file is located in the above location."
		genevent "E", "3", text
		registry = 1
		exit function
	end if

	if err.number <> 0 then logerror "End of registry Function", err.number : err.clear
	registry = 0
End Function




REM #################################################################################################
REM ###                                  Registry backup                                          ###
REM #################################################################################################
Function registry_backup()
	dim strcmdline, temp, text, wshtempfile, line, number, timesync, backup
	on error resume next

	REM ###### Checks for registry backup AT jobs ######
	screenout "Checking registry backup at jobs:"
	strcmdline = "cmd /c ""at \\" & computername & " > " & scripttempvar & "\atjobstemp.txt"""	
	temp = wshshell.Run(strcmdline, 0, true)
	if (Wshfile.fileexists(scripttempvar & "\atjobstemp.txt")) then 
		Set wshtempfile = wshfile.OpenTextFile(scripttempvar & "\atjobstemp.txt", 1)
		backup=0
		Do While wshtempfile.AtEndOfStream <> true
			line = wshtempFile.ReadLine
			if Instr(1, line, "rdisk", 1) then
				number = mid(line, 9, 1)
				screenout "  Removing Rdisk AT job ID " & number & "..."
				strcmdline = "cmd /c ""at " & number & " /d"""	
				temp = wshshell.Run(strcmdline, 0, true)
			elseif Instr(1, line, "regbackup", 1) then
				number = mid(line, 9, 1)
				screenout "  Removing RegBackup AT job ID " & number & "..."
				strcmdline = "cmd /c ""at " & number & " /d"""	
				temp = wshshell.Run(strcmdline, 0, true)
			elseif Instr(1, line, "itgbackup", 1) then
				screenout "  AT job for ItgBackup OK."
				backup=1
			end if
		Loop
		wshtempfile.Close
		wscript.DisconnectObject wshtempfile
		Set wshtempfile=nothing
		if backup <> 1 then
			screenout "  Adding AT job for ITGBackup..."
			strcmdline = "cmd /c ""at \\" & computername & " 3:00 /EVERY:m,t,w,th,f,s,su " & scriptbinvar & "\ITGBackup.vbs"""	
			temp = wshshell.Run(strcmdline, 0, true)
		end if
	else
		screenout ""
		text = "Unable to create a text file in the temporary directory for the script!"
		screenout text
		screenout "Make sure that " & scripttempvar & " exists and is writable."
		screenout "Make sure that at.exe is available and working properly (at \\servername)."
		genevent "E", "3", text
		registry_backup = 1
		exit function
	end if

	if err.number <> 0 then logerror "End of registry_backup Function", err.number : err.clear
	registry_backup = 0
End Function



REM #################################################################################################
REM ###                                      Registry Main                                        ###
REM #################################################################################################
Function registry_main()
	dim text
	on error resume next

	REM ###### Set Registry values ######	
	screenout "Checking Main registry entries:"
	registry_main = registry(scriptinilocvar & "\MainRegistryW2K.ini")

	if err.number <> 0 then logerror "End of registry_main Function", err.number : err.clear
End Function


REM #################################################################################################
REM ###                                    Registry Filters                                       ###
REM #################################################################################################
Function registry_filters()
	dim temp, strcmdline, text, wshtempfile, alltext, line, pos, value, value1
	on error resume next

	REM ###### check for PASSFILT security filter ######
	screenout "Checking security filter list:"

	REM ###### Get Current Filters ######
	strCmdLine = "cmd /c """& scriptfileslocvar & "\reg query ""HKLM\system\currentcontrolset\Control\LSA"" /v ""Notification Packages"" >" & scripttempvar & "\passfilttemp.txt"""
	temp = wshshell.Run(strcmdline, 0, true)
	if err.number <> 0 then err.clear
	if (Wshfile.fileexists(scripttempvar & "\passfilttemp.txt")) then 
		Set wshtempfile = wshfile.OpenTextFile(scripttempvar & "\passfilttemp.txt", 1)
		if wshtempfile.AtEndOfStream = true then screenout "Reg.exe cannot get filter information!"
		Do While wshtempfile.AtEndOfStream <> true
			line = wshtempFile.ReadLine
			if Instr(1, line, "Notification Packages", 1) <> 0 and Instr(1, line, "PASSFILT", 1) <> 0 then
				screenout "  Removing PASSFILT from filter list..."
				Value1 = replace(line, "PASSFILT\0", "", 40, 1, 1)
				Value = replace(value1, "\0\0", "", 1, 1)
				strCmdLine = "cmd /c """& scriptfileslocvar & "\reg add ""HKLM\system\currentcontrolset\Control\LSA"" /v ""Notification Packages"" /t REG_MULTI_SZ /d " & value & " /f"""
				temp = wshshell.Run(strcmdline, 0, true)
				if err.number <> 0 then err.clear
				exit do
			elseif Instr(1, line, "Notification Packages", 1) <> 0 and Instr(1, line, "PASSFILT", 1) = 0 then
				screenout "  PASSFILT does not exist as a filter."
				exit do
			else
			end if	
		Loop
		wshtempfile.Close
		wscript.DisconnectObject wshtempfile
		Set wshtempfile=nothing
	else
		screenout ""
		text = "Unable to create a text file in the temporary directory for the script!"
		screenout text
		screenout "Make sure that " & scripttempvar & " exists and is writable."
		screenout "Make sure that reg.exe in " & scriptfileslocvar & " is working properly."
		genevent "E", "3", text
		registry_filters = 1
		exit function
	end if

	if err.number <> 0 then logerror "End of registry_filters Function", err.number : err.clear
	registry_filters = 0
End Function



REM #################################################################################################
REM ###                                   Registry PageFile                                       ###
REM #################################################################################################
Function registry_pagefile()
	dim text, temp, pfarray, strcmdline, wshtempfile, line, pos, pos2, pos3, size1, size2
	dim currentmemsize, currentpagedrive, currentpagemin, currentpagemax, correctpagedrive, correctpagemin, correctpagemax, wshpagefile, pagefilesize
	on error resume next

	REM ###### Check pagefile setting ######
	screenout "Checking servers pagefile setting:"

	REM ### Don't run on SAP servers ###
	if not sapvar = "" then
		screenout "  The Pagefile does not get modified on SAP servers."
		registry_pagefile = 0
		exit function
	end if

	REM ###### Get Current memory size ######
	currentmemsize = srvmemvar
	screenout "  Current Physical memory is " & currentmemsize & " MB."

	REM ###### Get Current pagefile size ######
	pfarray =  WshShell.RegRead("HKLM\system\currentcontrolset\Control\session manager\memory management\pagingfiles")
	pos = instr(1, pfarray(0), "\", 1)
	currentpagedrive = left(pfarray(0), pos-1)
	pos2 = instr(1, pfarray(0), " ", 1)
	pos3 = instr(pos2+1, pfarray(0), " ", 1)
	currentpagemin = cint(mid(pfarray(0), pos2, pos3-pos2))
	currentpagemax = cint(mid(pfarray(0), pos3, len(pfarray(0))+1-pos3))
	screenout "  Current Pagefile is " & currentpagedrive & "\pagefile.sys " & currentpagemin & " " & currentpagemax & " MB."
	
	REM ###### get correct pagefile settings ######
	if currentmemsize >=2048 then
		correctpagedrive = systemdrive
		correctpagemin = currentmemsize
		correctpagemax  = currentmemsize
	else
		correctpagedrive = systemdrive
		correctpagemin = currentmemsize * 1.5
		correctpagemax  = currentmemsize * 1.5
	end if
	if currentmemsize >4095 then
		correctpagemin = 4095
		correctpagemax  = 4095
	end if

	screenout "  Correct Pagefile is " & correctpagedrive & "\pagefile.sys " & correctpagemin & " " & correctpagemax & " MB."
	REM ###### Check for correct size ######
	if currentpagedrive = "" or currentpagemin = "" or currentpagemax = "" then
		screenout "  Pagefile setting not being set due to problem getting current pagefile settings!"
	elseif currentpagedrive = correctpagedrive and currentpagemin = correctpagemin and currentpagemax = correctpagemax then
		screenout "  Pagefile setting OK."
	elseif checkfreespace((correctpagemin-currentpagemin) + 50) = 1 then
		screenout "  Pagefile setting not being set due to lack of drive space!"
	elseif not currentpagedrive = correctpagedrive then
		screenout "  Pagefile setting not being set due to being on a different drive!"
	elseif checkfreespace((correctpagemin-currentpagemin) + 50) = 0 then
		pfarray(0) = correctpagedrive & "\pagefile.sys " & correctpagemin & " " & correctpagemax
		screenout "  Pagefile setting being set to " & pfarray(0) & "..."
		strCmdLine = "cmd /c """& scriptfileslocvar & "\reg add ""HKLM\system\currentcontrolset\Control\session manager\memory management"" /v ""pagingfiles"" /t REG_MULTI_SZ /d """ & pfarray(0) & """ /f"""
		temp = wshshell.Run(strcmdline, 0, true)
		if err.number <> 0 then err.clear
	else
		screenout "  Pagefile cannot be set with existing conditions!"
	end if

	if err.number <> 0 then logerror "End of registry_pagefile Function", err.number : err.clear
	registry_pagefile = 0
End Function




REM #################################################################################################
REM ###                                     Registry Path                                         ###
REM #################################################################################################
Function registry_path()
	dim check, oldpath, newpath, firstpos, lastpos, onepart
	on error resume next

	REM ###### Get current path ######
	check = 0
	screenout "Checking machines Path:"
	oldpath = WshShell.Regread("HKLM\system\currentcontrolset\control\session manager\environment\path")
	if err.number <> 0 then err.clear
	screenout "  Old Path - " & oldpath
	newpath = "%SystemRoot%\system32;%SystemRoot%;%SystemRoot%\System32\Wbem;%SystemDrive%\localbin"

	REM ###### Get each part of the current path and append it to the new path if needed ######
	screenout "  Creating new path..."
	firstpos = 1
	do
		lastpos = Instr(firstpos, oldpath, ";", 1)
		if lastpos = 0 then lastpos = len(oldpath) + 1
		onepart = mid(oldpath, firstpos, lastpos-firstpos)
		if check = 1 then screenout "    One part - " & onepart
		if check = 1 then screenout "      First Position - " & firstpos
		if check = 1 then screenout "      Last Position - " & lastpos
		if onepart = "" then
			if check = 1 then screenout "      Skipping 1... "
		elseif onepart = "c:\winnt" then
			if check = 1 then screenout "      Skipping 2... "
		elseif onepart = systemroot then
			if check = 1 then screenout "      Skipping 3... "
		elseif onepart = "%systemroot%" then
			if check = 1 then screenout "      Skipping 4... "
		elseif onepart = "%SystemRoot%" then
			if check = 1 then screenout "      Skipping 5... "
		elseif onepart = "c:\winnt\system32" then
			if check = 1 then screenout "      Skipping 6... "
		elseif onepart = systemroot & "\system32" then
			if check = 1 then screenout "      Skipping 7... "
		elseif onepart = "%systemroot%\system32" then
			if check = 1 then screenout "      Skipping 8... "
		elseif onepart = "%SystemRoot%\system32" then
			if check = 1 then screenout "      Skipping 9... "
		elseif onepart = "c:\winnt\system32\wbem" then
			if check = 1 then screenout "      Skipping 10... "
		elseif onepart = systemroot & "\System32\Wbem" then
			if check = 1 then screenout "      Skipping 11... "
		elseif onepart = "%systemroot%\System32\Wbem" then
			if check = 1 then screenout "      Skipping 12... "
		elseif onepart = "%SystemRoot%\System32\Wbem" then
			if check = 1 then screenout "      Skipping 13... "
		elseif onepart = "%SystemRoot%\system32\Wbem" then
			if check = 1 then screenout "      Skipping 14... "
		elseif onepart = "%SystemRoot%\system32\WBEM" then
			if check = 1 then screenout "      Skipping 25... "
		elseif onepart = "c:\localbin" then
			if check = 1 then screenout "      Skipping 15... "
		elseif onepart = "%systemdrive%\localbin" then
			if check = 1 then screenout "      Skipping 16... "
		elseif onepart = "%SystemDrive%\localbin" then
			if check = 1 then screenout "      Skipping 17... "
		elseif onepart = systemdrive & "\localbin" then
			if check = 1 then screenout "      Skipping 18... "
		elseif onepart = "C:\Program" then
			if check = 1 then screenout "      Skipping 19... "
		elseif onepart = "C:\PROGRA~1\Dell\bin" then
			if check = 1 then screenout "      Skipping 20... "
		elseif onepart = "C:\PROGRA~1\Dell\ihv\bin" then
			if check = 1 then screenout "      Skipping 21... "
		elseif onepart = "C:\PROGRA~1\Dell\dmi\bin" then
			if check = 1 then screenout "      Skipping 22... "
		elseif onepart = "C:\Program Files\PERC2\System" then
			if check = 1 then screenout "      Skipping 23... "
		elseif onepart = "c:\localbin\dos" then
			if check = 1 then screenout "      Skipping 24... "
		elseif onepart = "c:\" then
		else
			screenout "      Appending " & onepart & " to the new path..."
			newpath = newpath & ";" & onepart
		end if
		firstpos = lastpos + 1
	loop while not lastpos = len(oldpath) + 1
	screenout "  New Path - " & newpath

	REM ### Add new path to machine if it has changed ###
	if newpath = oldpath or check = 1 then
		screenout "  Machine path OK."
	else
		screenout "  Correcting machines path..."
		wshshell.regwrite "HKLM\system\currentcontrolset\control\session manager\environment\path", newpath, "REG_EXPAND_SZ"
		if err.number <> 0 then logerror "Setting Machine Path", err.number : err.clear
		logregistryentry "System Path", "Corrected", "HKLM\system\currentcontrolset\control\session manager\environment\path", newpath
	end if

	if err.number <> 0 then logerror "End of registry_path Function", err.number : err.clear
	registry_path = 0
End Function




REM #################################################################################################
REM ###                                  Registry Source location                                 ###
REM #################################################################################################
Function registry_source()
	dim source1, source2, source3, text
	on error resume next

	if bldvar = "off" or bldvar = "switch" then
	else
		REM ###### Checking default source path ######
		screenout "Checking machines default source path(s):"

		REM ###### Get path and check ######
		source1 = WshShell.RegRead("HKLM\software\microsoft\windows nt\currentversion\SourcePath")
		source2 = WshShell.RegRead("HKLM\software\microsoft\windows\currentversion\setup\SourcePath")
		source3 = WshShell.RegRead("HKLM\software\microsoft\windows\currentversion\setup\ServicePackSourcePath")
		if err.number <> 0 then logerror "Reading first sourcepath", err.number : err.clear
		if cstr(source1) = cstr(bldlocvar) and cstr(source2) = cstr(bldlocvar) and cstr(source3) = cstr(splocvar) then
			screenout "  Default source path registry locations OK."
		else
			screenout "  Setting default source path registry locations..."
			wshshell.regwrite "HKLM\software\microsoft\windows nt\currentversion\SourcePath", bldlocvar, "REG_SZ"
			wshshell.regwrite "HKLM\software\microsoft\windows\currentversion\setup\SourcePath", bldlocvar, "REG_SZ"
			wshshell.regwrite "HKLM\software\microsoft\windows\currentversion\setup\ServicePackSourcePath", splocvar, "REG_SZ"
			if err.number <> 0 then logerror "Default source path registry locations", err.number : err.clear
			logregistryentry "Source Files", "Updated", "HKLM\software\microsoft\windows nt\currentversion\SourcePath", bldlocvar
			logregistryentry "Source Files 2", "Updated", "HKLM\software\microsoft\windows\currentversion\setup\SourcePath", bldlocvar
			logregistryentry "Source Files 3", "Updated", "HKLM\software\microsoft\windows\currentversion\setup\ServicePackSourcePath", splocvar
		end if
	end if

	if err.number <> 0 then logerror "End of registry_source Function", err.number : err.clear
	registry_source = 0
End Function




REM #################################################################################################
REM ###                                  Registry DiskPerf settings                               ###
REM #################################################################################################
Function registry_diskperf()
	dim strcmdline, temp, text
	on error resume next

	REM ###### Registering DiskPerf counters ######
	screenout "Checking disk perfmon counters:"

	REM ###### Registering DiskPerf counters ######
	screenout "  Disabling disk perfmon counters..."
	strcmdline = "cmd /c ""diskperf -N"""
        temp = wshshell.Run(strcmdline, 0, TRUE)
        if err.number <> 0 then err.clear
	screenout "  Enabling disk perfmon counters..."
	strcmdline = "cmd /c ""diskperf -Y"""
        temp = wshshell.Run(strcmdline, 0, TRUE)
        if err.number <> 0 then err.clear

	if err.number <> 0 then logerror "End of registry_diskperf Function", err.number : err.clear
	registry_diskperf = 0
End Function




REM #################################################################################################
REM ###                                      Registry Site                                        ###
REM #################################################################################################
Function registry_site()
	dim text, siteregistryfile
	on error resume next

	Select Case scriptconswivar
		Case "/b11-corp": siteregistryfile = "b11corp"
		Case "/b11-int": siteregistryfile = "cpint"
		Case "/b11-ext": siteregistryfile = "cpint"
		Case "/cp-corp": siteregistryfile = "b11corp"
		Case "/cp-int": siteregistryfile = "cpint"
		Case "/cp-ext": siteregistryfile = "cpint"
		Case "/tuk-corp": siteregistryfile = "b11corp"
		Case "/tuk-int": siteregistryfile = "cpint"
		Case "/tuk-ext": siteregistryfile = "cpint"
		Case "/sat-corp": siteregistryfile = "satcorp"
		Case "/sat-int": siteregistryfile = "satint"
		Case "/sat-ext": siteregistryfile = "satint"
		Case "/jup-corp": siteregistryfile = "jupcorp"
		Case "/jup-int": siteregistryfile = "jupcorp"
		Case "/jup-ext": siteregistryfile = "jupcorp"
		Case "/soc-corp": siteregistryfile = "socint"
		Case "/soc-int": siteregistryfile = "socint"
		Case "/soc-ext": siteregistryfile = "socint"
		Case "/dsk-corp": siteregistryfile = "deskcorp"
		Case "/dsk-int": siteregistryfile = "deskcorp"
		Case "/dsk-ext": siteregistryfile = "deskcorp"
		Case "/noam-corp": siteregistryfile = "noamcorp"
		Case "/noam-int": siteregistryfile = "noamint"
		Case "/noam-ext": siteregistryfile = "noamint"
		Case "/soam-corp": siteregistryfile = "noamcorp"
		Case "/soam-int": siteregistryfile = "noamint"
		Case "/soam-ext": siteregistryfile = "noamint"
		Case "/sopa-corp": siteregistryfile = "faeacorp"
		Case "/sopa-int": siteregistryfile = "faeaint"
		Case "/sopa-ext": siteregistryfile = "faeaint"
		Case "/faea-corp": siteregistryfile = "faeacorp"
		Case "/faea-int": siteregistryfile = "faeaint"
		Case "/faea-ext": siteregistryfile = "faeaint"
		Case "/euro-corp": siteregistryfile = "eurocorp"
		Case "/euro-int": siteregistryfile = "euroint"
		Case "/euro-ext": siteregistryfile = "euroint"
		Case "/miea-corp": siteregistryfile = "eurocorp"
		Case "/miea-int": siteregistryfile = "euroint"
		Case "/miea-ext": siteregistryfile = "euroint"
		Case "/afca-corp": siteregistryfile = "eurocorp"
		Case "/afca-int": siteregistryfile = "euroint"
		Case "/afca-ext": siteregistryfile = "euroint"
		Case else
			screenout "No registry configuration exists for '" & scriptconswivar & "'!"
			registry_site = 0
			exit function
	End Select
	REM ###### Set Site Registry values ######
	screenout "Checking registry entries for " & siteregistryfile & " configuration:"
	registry_site = registry(scriptinilocvar & "\config\" & siteregistryfile & "RegistryW2K.ini")
	if registry_site = 1 then exit function

	if err.number <> 0 then logerror "End of registry_site Function", err.number : err.clear
	registry_site = 0
End Function



REM #################################################################################################
REM ###                                      Registry Debug                                        ###
REM #################################################################################################
Function registry_debug()
	dim text
	on error resume next

	REM ###### Set memory dump settings ######	
	screenout "Checking memory dump setting:"
	if srvmemvar <= 2048 then
		registry_debug = registry(scriptinilocvar & "\DumpComRegistryW2K.ini")
		if registry_debug = 1 then exit function
	else
		registry_debug = registry(scriptinilocvar & "\DumpKernRegistryW2K.ini")
		if registry_debug = 1 then exit function
	end if

	REM ###### Set full debug settings ######
	if srvdebug = 1 then
		screenout "Checking full debugger settings:"
		registry_debug = registry(scriptinilocvar & "\FullDebugRegistryW2K.ini")
		if registry_debug = 1 then exit function
	end if

	REM ###### Set DC Debug/registry Settings ######
	if srvrolevar = "Domain Controller" then
		screenout "Checking DC Registry Settings:"
		registry_debug = registry(scriptinilocvar & "\DCRegistryW2K.ini")
		if registry_debug = 1 then exit function
	end if

	if err.number <> 0 then logerror "End of registry_debug Function", err.number : err.clear
	registry_debug = 0
End Function



REM #################################################################################################
REM ###                                  Registry Services                                        ###
REM #################################################################################################
Function registry_services()
	dim text
	on error resume next

	REM #### Set optimization to Network Application for given services ######
	if not wwwvar = "" or not sqlvar = "" or not exchangevar = "" or not sapvar = "" then
		screenout "Checking registry entries for Optimization:"
		registry_services = registry(scriptinilocvar & "\OPTORegistryW2K.ini")
		if registry_services = 1 then exit function
	end if

	REM ###### Set TSRA Registry values ######
	if termvar = "TermRa" and (Wshfile.FileExists(scriptinilocvar & "\TermRaRegistryW2K.ini")) then
		screenout "Checking registry entries for Terminal Service (Remote Administration):"
		registry_services = registry(scriptinilocvar & "\TermRaRegistryW2K.ini")
		if registry_services = 1 then exit function
	elseif termvar = "TermApp" and (Wshfile.FileExists(scriptinilocvar & "\TermAppRegistryW2K.ini")) then
		screenout "Checking registry entries for Terminal Service (Application):"
		registry_services = registry(scriptinilocvar & "\TermAppRegistryW2K.ini")
		if registry_services = 1 then exit function
	end if

	REM #### Set registry settings on WWW servers ######
	if not wwwvar = "" and (Wshfile.FileExists(scriptinilocvar & "\WWWRegistryW2K.ini")) then
		screenout "Checking registry entries for WWW services:"
		registry_services = registry(scriptinilocvar & "\WWWRegistryW2K.ini")
		if registry_services = 1 then exit function
	end if

	REM #### Set registry settings on SQL servers ######
	if not sqlvar = "" and (Wshfile.FileExists(scriptinilocvar & "\SQLRegistryW2K.ini")) then
		screenout "Checking registry entries for SQL services:"
		registry_services = registry(scriptinilocvar & "\SQLRegistryW2K.ini")
		if registry_services = 1 then exit function
	end if

	REM #### Set registry settings on exchange servers ######
	if not exchangevar = "" and (Wshfile.FileExists(scriptinilocvar & "\EXCRegistryW2K.ini"))then
		screenout "Checking registry entries for Exchange services:"
		registry_services = registry(scriptinilocvar & "\EXCRegistryW2K.ini")
		if registry_services = 1 then exit function
	end if

	REM #### Set registry settings on sap servers ######
	if not sapvar = "" and (Wshfile.FileExists(scriptinilocvar & "\SAPRegistryW2K.ini")) then
		screenout "Checking registry entries for SAP services:"
		registry_services = registry(scriptinilocvar & "\SAPRegistryW2K.ini")
		if registry_services = 1 then exit function
	end if

	if err.number <> 0 then logerror "End of registry_services Function", err.number : err.clear
	registry_services = 0
End Function




REM #################################################################################################
REM ###                                    Perfcol                                                ###
REM #################################################################################################
Function tools_perfcol()
	dim temp, strcmdline, text
	on error resume next

	REM ### Add Perfcol group for perfmon access ###
	screenout "Adding Perfcol group to registry permissions for perfcol counter access..."
	strCmdLine = "cmd /c """& scriptfileslocvar & "\regsecadd -l -a ""software\microsoft\windows nt\currentversion\perflib"" Redmond\Perfcol"""
	temp = wshshell.Run(strcmdline, 0, true)
	if err.number <> 0 then err.clear
	strCmdLine = "cmd /c """& scriptfileslocvar & "\regsecadd -l -a ""system\currentcontrolset\control\securepipeservers\winreg"" Redmond\Perfcol"""
	temp = wshshell.Run(strcmdline, 0, true)
	if err.number <> 0 then err.clear
	strCmdLine = "cmd /c """& scriptfileslocvar & "\regsecadd -l -a ""system\currentcontrolset\control\securepipeservers\winreg\allowedpaths"" Redmond\Perfcol"""
 	temp = wshshell.Run(strcmdline, 0, true)
	if err.number <> 0 then err.clear
	strCmdLine = "cmd /c """& scriptfileslocvar & "\regsecadd -l -a ""software\microsoft\windows nt\currentversion\perflib"" Houston\Perfcol"""
	temp = wshshell.Run(strcmdline, 0, true)
	if err.number <> 0 then err.clear
	strCmdLine = "cmd /c """& scriptfileslocvar & "\regsecadd -l -a ""system\currentcontrolset\control\securepipeservers\winreg"" Houston\Perfcol"""
	temp = wshshell.Run(strcmdline, 0, true)
	if err.number <> 0 then err.clear
	strCmdLine = "cmd /c """& scriptfileslocvar & "\regsecadd -l -a ""system\currentcontrolset\control\securepipeservers\winreg\allowedpaths"" Houston\Perfcol"""
 	temp = wshshell.Run(strcmdline, 0, true)
	if err.number <> 0 then err.clear
	if temp <> 0 then
		screenout "  Failed to properly add Perfcol group(s) to registy!" 
	end if

	if err.number <> 0 then logerror "End of tools_perfcol Function", err.number : err.clear
	tools_perfcol = 0
End Function




REM #################################################################################################
REM ###                                            OIC                                            ###
REM #################################################################################################
Function tools_oic()
	dim temp, strcmdline, text, regthere, filethere, servfile, lettervar, locationvar, reboot, inocsitevar
	on error resume next

	REM ### Check for Object Inactivity Checker location ###
	screenout "Object Inactivity Checker:"
	if oicvar = "off"  then
		screenout "  Object Inactivity Checker install is disabled."
		oiclogvar = "Disabled"
		oiclogvar = "Did Not Run Correctly"
	elseif oicvar = "switch"  then
		screenout "  Object Inactivity Checker install has been turned off by the user."
		oiclogvar = "Disabled - By User"
	elseif not srvrolevar = "Domain Controller" then
		screenout "  Object Inactivity Checker install is not needed on Servers (DC's Only!)."
		oiclogvar = "Not a Domain Controller"
	else	
		if (Wshfile.fileexists(oicchkvar)) then
			set servfile = Wshfile.getfile(oicchkvar)
			filethere = datevalue(left(servfile.datelastmodified, Instr(1, servfile.datelastmodified, " ", 1)-1))
			wscript.DisconnectObject servfile
			set servfile=nothing
			if err.number <> 0 then err.clear
		else
			filethere = "no"
		end if
		oicfilevar = datevalue(oicfilevar)

		if debug = 1 then screenout "  (Object Inactivity Checker File - " & filethere & ")"

		if datecomp(oicfilevar, filethere) and forceupdate = 1 then 
			screenout "  Updating Object Inactivity Checker..."
			oiclogvar = "Updated"
		elseif datecomp(oicfilevar, filethere) then 
			screenout "  Object Inactivity Checker OK."
			oiclogvar = "OK"
			tools_oic = 0
			exit function
		elseif filethere = "no" then 
			screenout "  Installing Object Inactivity Checker..."
			oiclogvar = "Installed"
		else
			screenout "  Upgrading Object Inactivity Checker..."
			oiclogvar = "Upgraded"
		end if
		if (Wshfile.FileExists(oiclocvar)) and checkfreespace(5) = 0 then
			strcmdline = "cmd /c """ & oiclocvar & """"
			if debug = 1 then screenout "  (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, true)
		elseif checkfreespace(5) = 1 then
			text = "  There is not enough free space on the boot drive to install OIC!"
			screenout text
			genevent "W", "4", text
			oiclogvar = "Not Enough Disk Space"
		else
			text = "  The script could not locate the OIC install script!"
			screenout text
			screenout "  (" & oiclocvar & ")"
			genevent "W", "4", text
			oiclogvar = "No Install File"
		end if
	end if


	if err.number <> 0 then logerror "End of tools_oic Function", err.number : err.clear
	tools_oic = 0
End Function




REM #################################################################################################
REM ###                                           Opassist                                        ###
REM #################################################################################################
Function tools_opassist()
	dim temp, strcmdline, servfile, filethere, text
	on error resume next

	REM ### Check for opassist location ###
	screenout "OpAssist:"
	if opavar = "off"  then
		screenout "  OpAssist install is disabled."
		opalogvar = "Disabled"
	elseif scriptsitevar = "DESK" then
		screenout "  Opassist is not needed on a Desktop configuration."
		opalogvar = "Desktop configuration."
	elseif opavar = "switch"  then
		screenout "  OpAssist install has been turned off by the user."
		opalogvar = "Disabled - By User"
	else
		if (Wshfile.fileexists(opachkvar)) then
			set servfile = Wshfile.getfile(opachkvar)
			filethere = datevalue(left(servfile.datelastmodified, Instr(1, servfile.datelastmodified, " ", 1)-1))
			wscript.DisconnectObject servfile
			set servfile=nothing
			if err.number <> 0 then err.clear
		else
			filethere = "no"
		end if
		opafilevar = datevalue(opafilevar)

		if debug = 1 then screenout "  (OpAssist File: " & filethere & ")"

		if datecomp(opafilevar, filethere) and forceupdate = 1  then 
			screenout "  Updating OpAssist..."
			opalogvar = "Updated"
		elseif datecomp(opafilevar, filethere) then 
			screenout "  OpAssist OK."
			opalogvar = "OK"
			tools_opassist = 0
			exit function
		elseif filethere = "no" then 
			screenout "  Installing OpAssist..."
			opalogvar = "Installed"
		else
			screenout "  Upgrading OpAssist..."
			opalogvar = "Upgraded"
		end if
		if (Wshfile.FileExists(opalocvar)) and checkfreespace(10) = 0 then
			strcmdline = "cmd /c """ & opalocvar & """"
			if debug = 1 then screenout "  (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, true)
		elseif checkfreespace(10) = 1 then
			text = "  There is not enough free space on the boot drive to install OpAssist!"
			screenout text
			genevent "W", "4", text
			opalogvar = "Not Enough Disk Space"
		else
			text = "  The script could not locate the OpAssist install script!"
			screenout text
			screenout "  (" & opalocvar & ")"
			genevent "W", "4", text
			opalogvar = "No Install File"
		end if
	end if


	if err.number <> 0 then logerror "End of tools_opassist Function", err.number : err.clear
	tools_opassist = 0
End Function




REM #################################################################################################
REM ###                                           Sentry                                          ###
REM #################################################################################################
Function tools_sentry()
	dim temp, strcmdline, text, regthere, filethere, servfile
	on error resume next

	REM ### Check for Sentry location ###
	screenout "Sentry:"
	if senvar = "off"  then
		screenout "  Sentry install is disabled."
		senlogvar = "Disabled"
	elseif scriptsitevar = "DESK" then
		screenout "  Sentry is not needed on a Desktop configuration."
		senlogvar = "Desktop configuration."
	elseif senvar = "switch"  then
		screenout "  Sentry install has been turned off by the user."
		senlogvar = "Disabled - By User"
	else
		if (Wshfile.fileexists(senchkvar)) then
			set servfile = Wshfile.getfile(senchkvar)
			filethere = datevalue(left(servfile.datelastmodified, Instr(1, servfile.datelastmodified, " ", 1)-1))
			wscript.DisconnectObject servfile
			set servfile=nothing
			if err.number <> 0 then err.clear
		else
			filethere = "no"
		end if
		senfilevar = datevalue(senfilevar)

		if debug = 1 then screenout "  (Current Sentry File: " & filethere & ")"

		if datecomp(senfilevar, filethere) and forceupdate = 1  then 
			screenout "  Updating Sentry..."
			senlogvar = "Updated"
		elseif datecomp(senfilevar, filethere) then 
			screenout "  Sentry OK."
			senlogvar = "OK"
			tools_sentry = 0
			exit function
		elseif filethere = "no" then 
			screenout "  Installing Sentry..."
			senlogvar = "Installed"
		else
			screenout "  Upgrading Sentry..."
			senlogvar = "Upgraded"
		end if
		if (Wshfile.FileExists(senlocvar)) and checkfreespace(25) = 0 then
			strcmdline = "cmd /c """ & senlocvar & " " & scriptconswivar & """"
			if debug = 1 then screenout "  (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, true)
		elseif checkfreespace(25) = 1 then
			text = "  There is not enough free space on the boot drive to install Sentry!"
			screenout text
			genevent "W", "4", text
			senlogvar = "Not Enough Disk Space"
		else
			text = "  The script could not locate the Sentry install script!"
			screenout text
			screenout "  (" & senlocvar & ")"
			genevent "W", "4", text
			senlogvar = "No Install File"
		end if
	end if


	if err.number <> 0 then logerror "End of tools_Sentry Function", err.number : err.clear
	tools_Sentry = 0
End Function




REM #################################################################################################
REM ###                                        NetIQ                                           ###
REM #################################################################################################
Function tools_netiq()
	dim temp, strcmdline, text, regthere, filethere, servfile
	on error resume next

	REM ### Check for NetIQ location ###
	screenout "NetIQ:"
	if niqvar = "off"  then
		screenout "  NetIQ install is disabled."
		niqlogvar = "Disabled"
	elseif scriptsitevar = "DESK" then
		screenout "  NetIQ is not needed on a Desktop configuration."
		niqlogvar = "Desktop configuration."
	elseif niqvar = "switch"  then
		screenout "  NetIQ install has been turned off by the user."
		niqlogvar = "Disabled - By User"
	else	
		if (Wshfile.fileexists(niqchkvar)) then
			set servfile = Wshfile.getfile(niqchkvar)
			filethere = datevalue(left(servfile.datelastmodified, Instr(1, servfile.datelastmodified, " ", 1)-1))
			wscript.DisconnectObject servfile
			set servfile=nothing
			if err.number <> 0 then err.clear
		else
			filethere = "no"
		end if
		niqfilevar = datevalue(niqfilevar)

		if debug = 1 then screenout "  (NetIQ File: " & filethere & ")"

		if datecomp(niqfilevar, filethere) and forceupdate = 1  then 
			screenout "  Updating NetIQ..."
			niqlogvar = "Updated"
		elseif datecomp(niqfilevar, filethere) then 
			screenout "  NetIQ OK."
			niqlogvar = "OK"
			tools_netiq = 0
			exit function
		elseif filethere = "no" then 
			screenout "  Installing NetIQ..."
			niqlogvar = "Installed"
		else
			screenout "  Upgrading NetIQ..."
			niqlogvar = "Upgraded"
		end if
		if (Wshfile.FileExists(niqlocvar)) and checkfreespace(25) = 0 then
			strcmdline = "cmd /c """ & niqlocvar & " " & scriptconswivar & """"
			if debug = 1 then screenout "  (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, true)
		elseif checkfreespace(25) = 1 then
			text = "  There is not enough free space on the boot drive to install NetIQ!"
			screenout text
			genevent "W", "4", text
			niqlogvar = "Not Enough Disk Space"
		else
			text = "  The script could not locate the NetIQ install script!"
			screenout text
			screenout "  (" & niqlocvar & ")"
			genevent "W", "4", text
			niqlogvar = "No Install File"
		end if
	end if


	if err.number <> 0 then logerror "End of tools_netiq Function", err.number : err.clear
	tools_netiq = 0
End Function




REM #################################################################################################
REM ###                                        Inoculan                                           ###
REM #################################################################################################
Function tools_inoculan()
	dim temp, strcmdline, text, regthere, filethere, servfile
	on error resume next

	REM ### Check for Inoculan location ###
	screenout "Inoculan:"
	if inocvar = "off"  then
		screenout "  Inoculan install is disabled."
		inoclogvar = "Disabled"
		inoclogvar = "Did Not Run Correctly"
	elseif inocvar = "switch"  then
		screenout "  Inoculan install has been turned off by the user."
		inoclogvar = "Disabled - By User"
	else	
		if (Wshfile.fileexists(inocchkvar)) then
			set servfile = Wshfile.getfile(inocchkvar)
			filethere = datevalue(left(servfile.datelastmodified, Instr(1, servfile.datelastmodified, " ", 1)-1))
			wscript.DisconnectObject servfile
			set servfile=nothing
			if err.number <> 0 then err.clear
		else
			filethere = "no"
		end if
		inocchkvar = datevalue(inocchkvar)

		if debug = 1 then screenout "  (Current Inoculan File - " & filethere & ")"

		if filethere = "no" and scriptsitevar = "DESK" then 
			screenout "  Installing Inoculan on this Desk top machine..."
			inoclogvar = "Installed - Desk Top"
		elseif filethere = "no" and not argument("/inocon") <> 0 then 
			screenout "  Inoculan is not enabled to install (Add the '/inocon' switch to enable)."
			inoclogvar = "Not Enabled"
			tools_inoculan = 0
			exit function
		elseif datecomp(inocfilevar, filethere) and forceupdate = 1 then 
			screenout "  Updating Inoculan..."
			inoclogvar = "Updated"
		elseif datecomp(inocfilevar, filethere) then 
			screenout "  Inoculan OK."
			inoclogvar = "OK"
			tools_inoculan = 0
			exit function
		elseif filethere = "no" and argument("/inocon") <> 0 then 
			screenout "  Installing Inoculan on this server..."
			inoclogvar = "Installed"
		else
			screenout "  Upgrading Inoculan..."
			inoclogvar = "Upgraded"
		end if
		if (Wshfile.FileExists(inoclocvar)) and checkfreespace(25) = 0 then
			strcmdline = "cmd /c """ & inoclocvar & " " & scriptconswivar & """"
			if debug = 1 then screenout "  (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, true)
			screenout "    Clearing Inoculan for Windows NT window..."
			winmanipulate "", "Inoculan for Windows NT", "%fc"
			screenout "    Clearing Startup window..."
			winmanipulate "", "Startup", "%fc"
		elseif checkfreespace(25) = 1 then
			text = "  There is not enough free space on the boot drive to install Inoculan!"
			screenout text
			genevent "W", "4", text
			inoclogvar = "Not Enough Disk Space"
		else
			text = "  The script could not locate the Inoculan install script!"
			screenout text
			screenout "  (" & inoclocvar & ")"
			genevent "W", "4", text
			inoclogvar = "No Install File"
		end if
	end if

	if err.number <> 0 then logerror "End of tools_inoculan Function", err.number : err.clear
	tools_inoculan = 0
End Function




REM #################################################################################################
REM ###                                   Backup Accelerator                                      ###
REM #################################################################################################
Function tools_backupexec()
	dim temp, strcmdline, text, regthere, filethere, servfile
	on error resume next

	REM ### Check for Backup Accelerator location ###
	screenout "Backup Accelerator:"
	if bacvar = "off"  then
		screenout "  Backup Accelerator install is disabled."
		baclogvar = "Disabled"
		baclogvar = "Did Not Run Correctly"
	elseif bacvar = "switch"  then
		screenout "  Backup Accelerator install has been turned off by the user."
		baclogvar = "Disabled - By User"
	else	
		if (Wshfile.fileexists(bacchkvar)) then
			set servfile = Wshfile.getfile(bacchkvar)
			filethere = datevalue(left(servfile.datelastmodified, Instr(1, servfile.datelastmodified, " ", 1)-1))
			wscript.DisconnectObject servfile
			set servfile=nothing
			if err.number <> 0 then err.clear
		else
			filethere = "no"
		end if
		bacchkvar = datevalue(bacchkvar)

		if debug = 1 then screenout "  (Current Backup Accelerator File - " & filethere & ")"

		if filethere = "no" and not argument("/buacon") <> 0 then 
			screenout "  Backup Accelerator is not enabled to install (Add the '/buacon' switch to enable)."
			baclogvar = "Not Enabled"
			tools_backupexec = 0
			exit function
		elseif datecomp(bacfilevar, filethere) and forceupdate = 1 then 
			screenout "  Updating Backup Accelerator..."
			baclogvar = "Updated"
		elseif datecomp(bacfilevar, filethere) then 
			screenout "  Backup Accelerator OK."
			baclogvar = "OK"
			tools_backupexec = 0
			exit function
		elseif filethere = "no" and argument("/buacon") <> 0 then 
			screenout "  Installing Backup Accelerator on this server..."
			baclogvar = "Installed"
		else
			screenout "  Upgrading Backup Accelerator..."
			baclogvar = "Upgraded"
		end if
		if (Wshfile.FileExists(baclocvar)) and checkfreespace(25) = 0 then
			strcmdline = "cmd /c """ & baclocvar & """"
			if debug = 1 then screenout "  (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, true)
		elseif checkfreespace(25) = 1 then
			text = "  There is not enough free space on the boot drive to install Backup Accelerator!"
			screenout text
			genevent "W", "4", text
			baclogvar = "Not Enough Disk Space"
		else
			text = "  The script could not locate the Backup Accelerator install script!"
			screenout text
			screenout "  (" & baclocvar & ")"
			genevent "W", "4", text
			baclogvar = "No Install File"
		end if
	end if

	if err.number <> 0 then logerror "End of tools_backupexec Function", err.number : err.clear
	tools_backupexec = 0
End Function



REM #################################################################################################
REM ###                                     Other tools stuff                                     ###
REM #################################################################################################
Function tools_other()
	dim text
	on error resume next

	REM ### Check for new switch and set some services to manual start ###
	if argument("/new") <> 0 then
		screenout "Switch to disable monitoring services on New builds detected."
		wshshell.regwrite "HKLM\system\currentcontrolset\services\Sentry Alert Sender (SASS)\start", "3", "REG_DWORD"
		logregistryentry "Sentry Service to manual start", "Changed", "HKLM\system\currentcontrolset\services\Sentry Alert Sender (SASS)\start", "3"
		if err.number <> 0 then err.clear
	end if


	if err.number <> 0 then logerror "End of tools_other Function", err.number : err.clear
	tools_other = 0
End Function



REM #################################################################################################
REM ###                                        WWW Services                                       ###
REM #################################################################################################
Function services_www()
	dim temp, strcmdline, text
	on error resume next

	REM #### Set optimization to Network Application for given services ######
	if wwwvar = "IPAK" then
		if (Wshfile.FileExists(wwwlocvar & "\wwwsetup.vbs"))  then
			wwwlogvar = "Executed"
			screenout "Executing the WWW IPAK script..."
			strcmdline = "cmd /c """ & wwwlocvar & "\wwwsetup " & scriptconswivar & """"
			if debug = 1 then screenout "  (" & strcmdline & ")"
			temp = wshshell.Run(strcmdline, 1, true)
			if err.number <> 0 then err.clear
		else 
			text = "The script could not locate the WWW IPAK install script!"
			screenout text
			screenout "  (" & wwwlocvar & "\wwwsetup.vbs)"
			genevent "W", "4", text
		end if
	else
		wwwlogvar = "Not enabled"
		screenout "WWW IPAK not enabled."
	end if

	if err.number <> 0 then logerror "End of services_www Function", err.number : err.clear
	services_www = 0
End Function




REM #################################################################################################
REM ###                                         Other                                           ###
REM #################################################################################################
Function completion_other()
	dim  text, wshinifile, rkey
	on error resume next

	REM ###### register IPAK event log DLL ######
	screenout "Registering IPAK Event log DLL..."
	genevent "0", "0", "dll"

	REM ###### Create last ran registry entries ######
	enddate=date
	endtime=time
	if fixvar = "off" or fixvar = "switch" then 
		screenout "Removing Last Ran registry identifiers for the IPAK script..."
		if (Wshfile.fileexists(scriptinilocvar & "\DeleteRegKeysW2K.ini")) then 
			Set wshinifile = wshfile.OpenTextFile(scriptinilocvar & "\DeleteRegKeysW2K.ini", 1)
			Do While wshinifile.AtEndOfStream <> true
				rkey = wshiniFile.ReadLine
				wshshell.regdelete rkey
				if err.number <> 0 then err.clear
			loop
			wshinifile.Close
		else
			screenout ""
			text = "The script was unable to find the DeleteRegKeysW2K.ini file!"
			screenout text
			screenout "Make sure that " & scriptinilocvar & " exists and contains the necessary INI file."
			genevent "E", "3", text
			completion_other = 1
			exit function
		end if
		if err.number <> 0 then err.clear
	else
		screenout "Creating Last Ran registry identifiers..."
		wshshell.regwrite "HKLM\software\microsoft\windows nt\currentversion\LastUpdateScript", enddate & " " & endtime, "REG_SZ"
		wshshell.regwrite "HKLM\software\microsoft\windows nt\currentversion\IpakVersion", scriptipakvar , "REG_SZ"
		wshshell.regwrite "HKLM\software\microsoft\windows nt\currentversion\LastUpdateScriptName", scriptnamevar & " - Version " & scriptversionvar, "REG_SZ"
		wshshell.regwrite "HKLM\software\microsoft\windows nt\currentversion\ipak\ipaknt\Version", scriptipakvar , "REG_SZ"
		wshshell.regwrite "HKLM\software\microsoft\windows nt\currentversion\ipak\ipaknt\Scriptinfo", scriptnamevar & " - Version " & scriptversionvar, "REG_SZ"
		wshshell.regwrite "HKLM\software\microsoft\windows nt\currentversion\ipak\ipaknt\installedfrom", scriptpathlocvar, "REG_SZ"
		wshshell.regwrite "HKLM\software\microsoft\windows nt\currentversion\ipak\ipaknt\installedon", enddate & " " & endtime, "REG_SZ"
		wshshell.regwrite "HKLM\software\microsoft\windows nt\currentversion\ipak\ipaknt\installedby", userdomain & "\" & username, "REG_SZ"
		if scriptnetworkvar = "CORPORATE" then
			wshshell.regwrite "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\SNMP\Parameters\RFC1156Agent\sysObjectID", "1.3.6.1.4.1.311.24.1", "REG_SZ"
		elseif scriptnetworkvar = "INTERNET" or scriptnetworkvar = "EXTRANET" then
			wshshell.regwrite "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\SNMP\Parameters\RFC1156Agent\sysObjectID", "1.3.6.1.4.1.311.24.2", "REG_SZ"
		end if
		if err.number <> 0 then logerror "Creating Last Ran Registry entries", err.number : err.clear
	end if

	if err.number <> 0 then logerror "End of completion_other Function", err.number : err.clear
	completion_other = 0
End Function




REM #################################################################################################
REM ###                                         Logging                                           ###
REM #################################################################################################
Function completion_locallogging()
	dim  wshscriptlogfile, a, text
	on error resume next

	REM ### create logs directory and remove existing logs ###
	if not (Wshfile.folderexists(scriptlogvar)) then Wshfile.createFolder(scriptlogvar)
	if (Wshfile.fileexists(scriptlogvar & "\" & scriptnamevar & ".log")) then wshfile.deletefile scriptlogvar & "\" & scriptnamevar & ".log", TRUE
	if (Wshfile.fileexists(scriptlogvar & "\symbols.log")) then wshfile.deletefile scriptlogvar & "\symbols.log", TRUE
	if (Wshfile.fileexists(tempdir & "\" & computername & ".log")) then wshfile.deletefile tempdir & "\" & computername & ".log", TRUE
	if (Wshfile.fileexists("d:\exchsrvr" & scriptnamevar & ".log")) then wshfile.deletefile "d:\exchsrvr" & scriptnamevar & ".log", TRUE


	REM ###### Create server log ######
	screenout "Creating server log file on local machine..."
	set wshscriptlogfile = wshfile.createtextfile(scriptlogvar & "\" & scriptnamevar & ".log",1)
 	wshscriptlogfile.writeline("")
	wshscriptlogfile.writeline("Script Name: " & scriptnamevar)
	wshscriptlogfile.writeline("Script Version: " & scriptversionvar)
	wshscriptlogfile.writeline("IPAK Version: " & scriptipakvar)
 	wshscriptlogfile.writeline("Date and Time (start): " & begdate & " " & begtime)
 	wshscriptlogfile.writeline("Date and Time (end): " & enddate & " " & endtime)
	wshscriptlogfile.writeline("User Name: " & username)
	wshscriptlogfile.writeline("Arguments: " & argtext)
	wshscriptlogfile.writeline("Configuration Switch: " & scriptconswivar)
	wshscriptlogfile.writeline("Site Location: " & scriptsitevar)
	wshscriptlogfile.writeline("Network Type: " & scriptnetworkvar)
	wshscriptlogfile.writeline("Source Location path: " & scriptpathlocvar)
	wshscriptlogfile.writeline("################ Server Information ####################")
	wshscriptlogfile.writeline("Server Name: " & computername)
	wshscriptlogfile.writeline("Windows Version: " & srvvervar)
	wshscriptlogfile.writeline("Windows Build: " & srvbuildvar)
	wshscriptlogfile.writeline("Product Suite: " & srvsuitevar)
	wshscriptlogfile.writeline("Windows Architecture: " & srvarcvar)
	wshscriptlogfile.writeline("Computer Role: " & srvrolevar)
	wshscriptlogfile.writeline("Encryption Bits: " & srvencvar)
	wshscriptlogfile.writeline("Hardware Platform: " & srvhardwarevar)
	wshscriptlogfile.writeline("Boot Drive Format: " & srvbootvar)
	wshscriptlogfile.writeline("Physical Memory: " & srvmemvar)
	wshscriptlogfile.writeline("Logical Drives: " & harddrives)
	wshscriptlogfile.writeline("CD-ROM Drives: " & cddrives)
	wshscriptlogfile.writeline("################# Hardware Software ####################")
	wshscriptlogfile.writeline("Compaq SSD  " & cpqssdvar & ": " & cpqssdlogvar)
	wshscriptlogfile.writeline("Compaq CIM  " & cpqcimvar & ": " & cpqcimlogvar)
	wshscriptlogfile.writeline("Compaq Storage Works Driver  " & cpqstrdrvvar & ": " & cpqstrdrvlogvar)
	wshscriptlogfile.writeline("Dell Managed Node " & delmannvar & ": " & delmannlogvar)
	wshscriptlogfile.writeline("Dell FAST Utility " & delfastvar & ": " & delfastlogvar)
	wshscriptlogfile.writeline("Dell Perc2 Driver " & deldrvvar & ": " & deldrvlogvar)
	wshscriptlogfile.writeline("################# Update Software ######################")
	wshscriptlogfile.writeline("Recovery Console: " & recconslogvar)
	wshscriptlogfile.writeline("Service Pack: " & splogvar)
	wshscriptlogfile.writeline("Symbols: " & symlogvar)
	wshscriptlogfile.writeline("Hotfixes: " & fixlogvar)
	wshscriptlogfile.writeline("128 Bit Encryption: " & enclogvar)
	wshscriptlogfile.writeline("####################### Tools ##########################")
	wshscriptlogfile.writeline("OpAssist: " & opalogvar)
	wshscriptlogfile.writeline("Sentry: " & senlogvar)
	wshscriptlogfile.writeline("NetIQ: " & niqlogvar)
	wshscriptlogfile.writeline("Inoculan: " & inoclogvar)
	wshscriptlogfile.writeline("Object Inactivity Checker: " & oiclogvar)
	wshscriptlogfile.writeline("Backup Accelerator: " & baclogvar)
	wshscriptlogfile.writeline("################## Windows Services ####################")
	wshscriptlogfile.writeline("TS Changes: " & termvar)
	wshscriptlogfile.writeline("################## Special Services ####################")
	wshscriptlogfile.writeline("WWW Changes: " & wwwvar)
	wshscriptlogfile.writeline("SQL Changes: " & sqlvar)
	wshscriptlogfile.writeline("Exchange Changes: " & exchangevar)
	wshscriptlogfile.writeline("SAP Changes: " & sapvar)
	wshscriptlogfile.writeline("WWW IPAK: " & wwwlogvar)
	wshscriptlogfile.writeline("################### Extra Services #####################")
	wshscriptlogfile.writeline("Installed extra services:")
	wshscriptlogfile.writeline(extraservices)
	wshscriptlogfile.writeline("################## Registry Values #####################")
	for a = 1 to regarrayloc - 1 step 1
		wshscriptlogfile.writeline(regarray(a,1) & " - " & regarray(a,2))
		wshscriptlogfile.writeline("  " & regarray(a,3) & " - " & regarray(a,4))
	next
	wshscriptlogfile.close
	wscript.DisconnectObject wshscriptlogfile
	set wshscriptlogfile=nothing


	REM ###### Create exchange log ######
	if not exchangevar = "" then 
		screenout "Creating Exchange server log file on local machine..."
		set wshscriptlogfile = wshfile.createtextfile("d:\exchsrvr\" & scriptnamevar & ".log",1)
 		wshscriptlogfile.writeline("")
		wshscriptlogfile.writeline("Script Name: " & scriptnamevar)
		wshscriptlogfile.writeline("Script Version: " & scriptversionvar)
 		wshscriptlogfile.writeline("Date and Time (start): " & begdate & " " & begtime)
 		wshscriptlogfile.writeline("Date and Time (end): " & enddate & " " & endtime)
		wshscriptlogfile.writeline("User Name: " & username)
		wshscriptlogfile.writeline("Configuraiton Switch: " & scriptconswivar)
		wshscriptlogfile.writeline("Site Location: " & scriptsitevar)
		wshscriptlogfile.writeline("Network Type: " & scriptnetworkvar)
		wshscriptlogfile.writeline("Source Location path: " & scriptpathlocvar)
		wshscriptlogfile.writeline("################ Server Information ####################")
		wshscriptlogfile.writeline("Server Name: " & computername)
		wshscriptlogfile.writeline("Windows Version: " & srvvervar)
		wshscriptlogfile.writeline("Windows Architecture: " & srvarcvar)
		wshscriptlogfile.writeline("Windows Build: " & srvbuildvar)
		wshscriptlogfile.writeline("Computer Role: " & srvrolevar)
		wshscriptlogfile.writeline("Product Suite: " & srvsuitevar)
		wshscriptlogfile.writeline("Encryption Bits: " & srvencvar)
		wshscriptlogfile.writeline("Hardware Platform: " & srvhardwarevar)
		wshscriptlogfile.writeline("Boot Drive Format: " & srvbootvar)
		wshscriptlogfile.writeline("Physical Memory: " & srvmemvar)
		wshscriptlogfile.writeline("Logical Drives: " & harddrives)
		wshscriptlogfile.writeline("CD-ROM Drives: " & cddrives)
		wshscriptlogfile.writeline("################## Special Services ####################")
		wshscriptlogfile.writeline("Exchange Changes: " & exchangevar)
		wshscriptlogfile.close
		wscript.DisconnectObject wshscriptlogfile
		set wshscriptlogfile=nothing
	end if

	if err.number <> 0 then logerror "End of completion_locallogging Function", err.number : err.clear
	completion_locallogging = 0
End Function




REM #################################################################################################
REM ###                                  Destination Logging                                      ###
REM #################################################################################################
Function completion_destlogging()
	dim  wshscriptlogfile, text, checkaccess
	on error resume next

	REM ##### Exception out test servers for logging ######
	if Instr(1, computername, "mustangs", 1) <> 0 then
		screenout "Destination logs not getting updated as this machine is a test machine."
		completion_destlogging = 0
		exit function
	elseif not (Wshfile.folderexists(scriptloglocvar)) then
		screenout "Logging destination directory cannot be found. Logs will be only be available on the server."
		completion_destlogging = 0
		exit function
	else
		set wshscriptlogfile = wshfile.createtextfile(scriptloglocvar & "\Access.txt")	
		if err.number <> 0 then checkaccess = 1 : err.clear
		wshscriptlogfile.close
		wscript.DisconnectObject wshscriptlogfile
		set wshscriptlogfile=nothing
		if (Wshfile.fileexists(scriptloglocvar & "\access.txt")) then wshfile.deletefile scriptloglocvar & "\access.txt", TRUE
		if checkaccess = 1 then
			screenout "Logging destination directory access problems. Logs will be only be available on the server."
			completion_destlogging = 0
			exit function
		else
			screenout "Manipulating destination logging server logs..."
			REM ###### Create directories if they do not exist ######
			if not (Wshfile.folderexists(scriptloglocvar & "\servers")) then wshfile.createfolder(scriptloglocvar & "\servers")
			if not (Wshfile.folderexists(scriptloglocvar & "\scripts")) then wshfile.createfolder(scriptloglocvar & "\scripts")

			REM ###### Copying server log file to logging destination ######
			screenout "  Copying server log file to destination logging server..."
			if (Wshfile.fileexists(scriptloglocvar & "\servers\" & computername & ".bak.log")) then wshfile.deletefile scriptloglocvar & "\servers\" & computername & ".bak.log", TRUE
			if (Wshfile.fileexists(scriptloglocvar & "\servers\" & computername & "-3.log")) then wshfile.deletefile scriptloglocvar & "\servers\" & computername & "-3.log", TRUE
			if (Wshfile.fileexists(scriptloglocvar & "\servers\" & computername & "-2.log")) then wshfile.copyfile scriptloglocvar & "\servers\" & computername & "-2.log", scriptloglocvar & "\servers\" & computername & "-3.log"
			if (Wshfile.fileexists(scriptloglocvar & "\servers\" & computername & "-1.log")) then wshfile.copyfile scriptloglocvar & "\servers\" & computername & "-1.log", scriptloglocvar & "\servers\" & computername & "-2.log"
			if (Wshfile.fileexists(scriptloglocvar & "\servers\" & computername & ".log")) then wshfile.copyfile scriptloglocvar & "\servers\" & computername & ".log", scriptloglocvar & "\servers\" & computername & "-1.log"
			if (Wshfile.fileexists(scriptlogvar & "\" & scriptnamevar & ".log")) then wshfile.copyfile scriptlogvar & "\" & scriptnamevar & ".log", scriptloglocvar & "\servers\" & computername & ".log"


			REM ###### Append to Script log on logging destination ######
			screenout "  Appending information to script log on destination logging server..."
			if (Wshfile.fileexists(scriptloglocvar & "\scripts\_" & scriptnamevar & scriptversionvar & ".log")) then 
				set wshscriptlogfile = wshfile.opentextfile(scriptloglocvar & "\scripts\_" & scriptnamevar & scriptversionvar & ".log",8)
			else
				set wshscriptlogfile = wshfile.createtextfile(scriptloglocvar & "\scripts\_" & scriptnamevar & scriptversionvar & ".log",1)
			end if
			wshscriptlogfile.writeline computername& " - " & begdate & " " & begtime   & " - " & enddate & " " & endtime  & " - " & srvbuildvar & " - " & scriptipakvar & " - " & username
			wshscriptlogfile.close
			wscript.DisconnectObject wshscriptlogfile
			set wshscriptlogfile=nothing
		end if
	end if

	if err.number <> 0 then logerror "End of completion_destlogging Function", err.number : err.clear
	completion_destlogging = 0
End Function




REM #################################################################################################
REM ###                                         Reboot                                            ###
REM #################################################################################################
Function completion_reboot()
	dim  temp, strcmdline, text, reboot
	on error resume next

	REM #### Delete temporary directory ######
	if argument("config") <> 0 or argument("setup") <> 0 or argument("bug") <> 0 then
	else
		wshfile.deletefolder scripttempvar, TRUE
		if err.number <> 0 then err.clear
	end if

	genevent "S", "2", scriptnamevar & " for IPAK " & scriptipakvar & " Completed Successfully."
	if scriptrebootvar  = "reboot" then
		screenout "The server is now being rebooted by " & scriptnamevar & "..."
		genevent "I", "5", "Server was rebooted by " & scriptnamevar & "."
		strCmdLine = "cmd /c """& scriptfileslocvar & "\reboot /L /R /T:10"""
		temp = wshshell.Run(strcmdline, 0, true)
	elseif scriptrebootvar = "ask" then
		screenout "The server should now be rebooted."
		reboot = msgbox("Do you wish to reboot the server now?", 20, scriptnamevar & " - Reboot Server?")
		if reboot = 6 then
			screenout "The server is now being rebooted by " & scriptnamevar & "..."
			genevent "I", "5", "Server was rebooted by " & scriptnamevar & "."
			strCmdLine = "cmd /c """& scriptfileslocvar & "\reboot /L /R /T:5"""
			temp = wshshell.Run(strcmdline, 0, true)
		else
			text = "The server must be rebooted for changes to take effect!"
			screenout text
			genevent "W", "4", text
		end if	 
	else
		text = "The server must be rebooted for changes to take effect!"
		screenout text
		genevent "W", "4", text
	end if

	if err.number <> 0 then logerror "End of completion_reboot Function", err.number : err.clear
	completion_reboot = 0		
End Function






REM #################################################################################################
REM ###                                                                                           ###
REM ###                                   Global Functions                                        ###
REM ###                                                                                           ###
REM #################################################################################################
REM #################################################################################################
REM ###                                       Screen Out                                          ###
REM #################################################################################################
Function screenout(text)
	dim wsherrorlogfile
	on error resume next

	wscript.echo text
	if (Wshfile.fileexists(systemdrive & "\" & scriptnamevar & "-run.log")) then 
		set wsherrorlogfile = wshfile.openTextfile(systemdrive & "\" & scriptnamevar & "-run.log",8)
 		wsherrorlogfile.writeline(text)
		wsherrorlogfile.close
		wscript.DisconnectObject Wsherrorlogfile
		set Wsherrorlogfile=nothing
		if err.number <> 0 then err.clear
	else
		set wsherrorlogfile = wshfile.createtextfile(systemdrive & "\" & scriptnamevar & "-run.log",1)
 		wsherrorlogfile.writeline(text)
		wsherrorlogfile.close
		wscript.DisconnectObject Wsherrorlogfile
		set Wsherrorlogfile=nothing
		if err.number <> 0 then err.clear
	end if

End Function




REM #################################################################################################
REM ###                                   Registry writes                                         ###
REM #################################################################################################
Function logregistryentry(text, status, rkey, rvalue)
	dim wshbootdrive, freespace
	on error resume next

		regarray(regarrayloc,1) = text
		regarray(regarrayloc,2) = status
		regarray(regarrayloc,3) = rkey
		regarray(regarrayloc,4) = rvalue
		regarrayloc = regarrayloc + 1

End Function




REM #################################################################################################
REM ###                                Argument Function                                          ###
REM #################################################################################################
Function argument(sttext)
	dim a
	on error resume next

	for a = 0 to wscript.arguments.count - 1
		if wscript.arguments(a) = sttext then argument = a + 1
	next

End Function



REM #################################################################################################
REM ###                                NextArgument Function                                      ###
REM #################################################################################################
Function nextargument(location)
	dim sttext
	on error resume next

	if wscript.arguments.count - 1 < location - 1 then
		sttext = ""
	else
		sttext = wscript.arguments(location)
	end if
	nextargument = sttext

End Function



REM #################################################################################################
REM ###                                GetArgument Function                                       ###
REM #################################################################################################
Function getargument(location)
	dim sttext
	on error resume next

	if wscript.arguments.count - 1 < location - 1 then
		sttext = ""
	else
		sttext = wscript.arguments(location - 1)
	end if
	getargument = sttext

End Function



REM #################################################################################################
REM ###                                InArgument Function                                          ###
REM #################################################################################################
Function inargument(sttext)
	dim a
	on error resume next

	for a = 0 to wscript.arguments.count - 1
		if instr(1, wscript.arguments(a), sttext, 1) <> 0 then inargument = a + 1
	next

End Function



REM #################################################################################################
REM ###                                    Check Freespace                                        ###
REM #################################################################################################
Function checkfreespace(size)
	dim wshbootdrive, freespace
	on error resume next

	Set wshbootdrive = wshfile.GetDrive(systemdrive)
	freespace = int(formatnumber(wshbootdrive.freespace/1048576,0))
	Wscript.DisconnectObject Wshbootdrive
	if freespace < size then
		checkfreespace = 1
	else
		checkfreespace = 0
	end if
	Set Wshbootdrive=nothing

End Function




REM #################################################################################################
REM ###                             Wait until text is found or not found                         ###
REM #################################################################################################
Function waituntil(there, notthere)
	dim alltext, checkpoint, strcmdline, temp, wshtempfile
	on error resume next

	REM ###### Loop until first arg is found or second arg is not found ######
	checkpoint = 0
	Do While checkpoint = 0
		strCmdLine = "cmd /c """& scriptfileslocvar & "\tlist >" & scripttempvar & "\tlisttemp.txt"""
		temp = wshshell.Run(strcmdline, 0, true)
		if err.number <> 0 then err.clear
		Set wshtempfile = wshfile.OpenTextFile(scripttempvar & "\tlisttemp.txt", 1)
		alltext = wshtempFile.Readall
		wshtempfile.Close
		wscript.DisconnectObject wshtempfile
		Set wshtempfile=nothing
		rem screenout "First Arg - " & there & " - " & instr(1, alltext, there, 1)
		rem screenout "Second Arg - " & notthere & " - " & instr(1, alltext, notthere, 1)
		if not there = "" and not instr(1, alltext, there, 1) = 0 then
			waituntil = 0
			checkpoint = 1
			wscript.sleep 5000 * netfactor
		end if
		if not notthere = "" and instr(1, alltext, notthere, 1) = 0 then
			waituntil = 1
			checkpoint = 1
		end if
	Loop

	if err.number <> 0 then logerror "End of WaitUntil Function", err.number : err.clear
end function




REM #################################################################################################
REM ###                                   Manipulate a window                                     ###
REM #################################################################################################
Function winmanipulate(check, popup, key)
	dim strcmdline, temp, wshtempfile, alltext
	on error resume next

	REM ###### Check for open window and manipulate that window ######
	strCmdLine = "cmd /c """& scriptfileslocvar & "\tlist >" & scripttempvar & "\tlisttemp.txt"""
	temp = wshshell.Run(strcmdline, 0, true)
	if err.number <> 0 then err.clear
	Set wshtempfile = wshfile.OpenTextFile(scripttempvar & "\tlisttemp.txt", 1)
	alltext = wshtempFile.Readall
	wshtempfile.Close
	wscript.DisconnectObject wshtempfile
	Set wshtempfile=nothing
	if check = "" or not instr(1, alltext, check, 1) = 0 then
		rem screenout popup
		rem screenout key
		wshshell.appactivate popup
		wscript.sleep 2000 * netfactor
		wshshell.sendkeys key
		winmanipulate = 0
	else
		winmanipulate = 1
	end if

	if err.number <> 0 then logerror "End of WinManipulate Function", err.number : err.clear
end function




REM #################################################################################################
REM ###                                       Error logging                                       ###
REM #################################################################################################
Function logerror(location, error)
	on error resume next
	screenout "******************************************************"
	screenout "******** Description: Script Error"
	screenout "******** Script Location: " & location
	screenout "******** Error Code: " & error
	screenout "******** User Action: Check the IPAK " & scriptipakvar & " TSG on Http://gnsweb/ipak"
	screenout "********              If the TSG does not contain any information, then fill out the"
	screenout "********              escalation template on Http://gnsweb/ipak and submit."
	screenout "******************************************************"
	genevent "E", "4", "Error with internal scripting component - (Description: " & location & "Error: " & error & ")"
	logerror = 0
End Function




REM #################################################################################################
REM ###                                   Generate Eventlog Entry                                 ###
REM #################################################################################################
Function genevent(severity, eventid, description)
	dim strcmdline, temp
	on error resume next

	if severity = eventid and description = "dll" then 
		strCmdLine = "cmd /c """& scriptfileslocvar & "\ipakevnt"""""
		temp = wshshell.Run(strcmdline, 0, true)
	else
		strCmdLine = "cmd /c """& scriptfileslocvar & "\ipakevnt -s " & severity & " -c " & eventid & " -r ""IPAKNT"" -e 200" & eventid & " """ & description & """"""
		rem if debug = 1 then screenout strcmdline
		temp = wshshell.Run(strcmdline, 0, true)
		if severity = "E" and eventid = 3 then temp = msgbox(description, 0, scriptnamevar & " - ERROR!")

	end if

	if err.number <> 0 then logerror "End of GenEvent Function", err.number : err.clear
	genevent = 0
End Function




REM #################################################################################################
REM ###                                        Data Compare                                       ###
REM #################################################################################################
Function datecomp(date1, date2)
	dim daydiff
	on error resume next

		if isdate(date1) and isdate(date2) then
			daydiff = datediff("d", date1, date2)
		else
			daydiff = -1
		end if
		if daydiff >= 0 then
			datecomp = True
		else
			datecomp = false
		end if

	if err.number <> 0 then logerror "End of Datecomp Function", err.number : err.clear
End Function





REM #################################################################################################
REM ###                                     Get Build CD                                          ###
REM #################################################################################################
Function getbuildcd()
	dim description, temp
	on error resume next

	if recconsvar = "off" or not srvbootvar = "NTFS" or recconsvar = "switch" then
	elseif cdrunvar = "on" or argument("forcecd") <> 0 then
		do while (Wshfile.folderexists(bldlocvar)) = false
			description = "Please insert the 'ITG Build CD (SP1 Series)', Version 7.40 or newer into the CD-ROM!"
			temp = msgbox(description, 0, scriptnamevar & " - Need Build CD!")
			if debug = 1 then screenout "  " & bldlocvar
		loop
		screenout "Found Build CD OK."
	end if

End Function





REM #################################################################################################
REM ###                                     Get Tools CD                                          ###
REM #################################################################################################
Function gettoolscd()
	dim description, temp
	on error resume next

	if argument("/notools") <> 0 then
	elseif cdrunvar = "on" or argument("forcecd") <> 0 then
		do while (Wshfile.fileexists(opalocvar)) = false
			description = "Please insert the 'ITG Tools CD', Version 1.00 or newer into the CD-ROM!"
			temp = msgbox(description, 0, scriptnamevar & " - Need Tools CD!")
			if debug = 1 then screenout "  " & opalocvar
		loop
		screenout "Found Tools CD OK."
	end if

End Function





REM #################################################################################################
REM ###                                    Get Update CD                                          ###
REM #################################################################################################
Function getupdatecd()
	dim description, temp
	on error resume next

	if cdrunvar = "on" or argument("forcecd") <> 0 then
		do while (Wshfile.fileexists(scriptpathlocvar & "\" & scriptnamevar & ".vbs")) = false
			description = "Please insert the 'ITG Update CD' for IPAK NT5.02 into the CD-ROM!"
			temp = msgbox(description, 0, scriptnamevar & " - Need Update CD!")
			if debug = 1 then screenout "  " & scriptpathlocvar & "\" & scriptnamevar & ".vbs"
		loop
		screenout "Found Update CD OK."
	end if

End Function





REM #################################################################################################
REM ###                                           Syntax                                          ###
REM #################################################################################################
Function syntax()
	on error resume next

	screenout ""
	screenout ""
	screenout "Name: " & scriptnamevar
	screenout "Usage: " & scriptnamevar & " [Configuration Switch] [Other Switches]"
	screenout "Description: " & scriptnamevar & " is used to update/config software on production W2K servers at Microsoft."
	screenout "Switches:"
	screenout "    Configuration Switch Syntax (One Required):"
	screenout "      /'location'-'network'"
	screenout "        'location' description:"
	screenout "           b11 - Building 11 configuration"
	screenout "           cp -  Canyon Park Data Center configuration"
	screenout "           tuk - Tukwila Data Center configuration"
	screenout "           sat - Saturn Lab configuration"
	screenout "           jup - Jupiter Lab configuration"
	screenout "           soc - MSN/SOC configuration"
	screenout "           dsk - Desk Top Machine configuration"
	screenout "           noam - North America configuration"
	screenout "           soam - South America configuration"
	screenout "           euro - Europe configuration"
	screenout "           sopa - South Pacific configuration"
	screenout "           faea - Far East configuration"
	screenout "           miea - Middle East configuration"
	screenout "           afca - Africa configuration"
	screenout "        'network' description:"
	screenout "           corp - Server on the Corporate network"
	screenout "           int - Server on the Internet network"
	screenout "           pri - Server on a Private network"
	screenout "           ext - Server on the Extranet network"
	screenout "    Examples:"
	screenout "      " & scriptnamevar & " /b11-corp"
	screenout "      " & scriptnamevar & " /noam-corp"
	screenout "      " & scriptnamevar & " /euro-corp"
	screenout "    Other Switches (optional):"
	screenout "        /com      - Forces the script to set hardware platform to Compaq."
	screenout "        /del      - Forces the script to set hardware platform to Dell."
	screenout "        /oth      - Forces the script to set hardware platform to other."
	screenout "                      No Compaq or Dell Components."
	screenout "        /nocim    - Forces the script to NOT install CIM."
	screenout "        /nostrwrk - Forces the script to NOT make any changes for Storage Works."
	screenout "        /allfixes - Enables install of ALL additional Hotfixes."
	screenout "        /fullsym  - Forces the script to copy files for the Full Symbol set."
	screenout "        /clrsym   - Forces the script to clear existing symbols and not copy in new ones."
	screenout "        /noupd    - Forces the script to NOT update the Windows components."
	screenout "                      No Service Pack, Hotfixes, or Symbols."
	screenout "        /reb      - Forces the script to automatically reboot when completed."
	screenout "        /noreb    - Forces the script to NOT reboot and end when completed."
	screenout "        /debug    - Forces the script to set full debugger settings."
	screenout "        /forceup  - Enables the script to update ALL up-to-date components."
	screenout "                      Except: Dell Management Software - Software will hang."
	screenout "        /slownet '#' - Gives the script a multipler to slow down the screen manipulations"
	screenout "                        for slow network links. A 2 causes the script to wait twice as"
	screenout "                        long before manipulating the screens."
	screenout "        /inocon   - Enables the Inoculan install on File Servers."
	screenout "        /buacon   - Enables the Backup Accelerator install."
	screenout "        /noopa    - Forces the script to NOT install OpAssist."
	screenout "        /nosen    - Forces the script to NOT install Sentry."
	screenout "        /noniq    - Forces the script to NOT install NetIQ."
	screenout "        /noinoc   - Forces the script to NOT update or install Inoculan."
	screenout "        /nooic    - Forces the script to NOT install Object Inactivity Checker. (DC's ONLY!)"
	screenout "        /nobac    - Forces the script to NOT install Backup Accelerator."
	screenout "        /notools  - Forces the script to NOT install any of the Tools."
	screenout "                      No Opassist, Sentry, NetIQ, Inoculan, OIC, or Backup Accelerator."
	screenout "        /nobitmap - Forces the script to NOT change the background bitmap."
	screenout "        /new      - Forces the script to set Sentry services to Manual start for new builds."
	screenout "        /wwwreg   - Forces the script to make NT registry changes for WWW services."
	screenout "        /sqlreg   - Forces the script to make NT registry changes for SQL services."
	screenout "        /excreg   - Forces the script to make NT registry changes for Exchange services."
	screenout "        /sapreg   - Forces the script to make NT registry changes for SAP services."
	screenout "        /wwwipak  - Enables execution of the WWW IPAK script."

	syntax=1
End Function

