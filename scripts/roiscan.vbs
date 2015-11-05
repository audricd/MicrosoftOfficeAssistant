' Name: Robust Office Inventory Scan - Version 1.8.1
' Author: Microsoft Customer Support Services
' Copyright (c) Microsoft Corporation. All rights reserved.
' Script to create an inventory scan of installed Office applications
' Supported Office Families: 2000, 2002, 2003, 2007
'                            2010, 2013, 2016, O365

Option Explicit
On Error Resume Next
Const SCRIPTBUILD = "1.8.1"
Dim sPathOutputFolder : sPathOutputFolder = ""
Dim fQuiet : fQuiet = False
Dim fLogFeatures : fLogFeatures = False
Dim fLogFull : fLogFull = False
dim fBasicMode : fBasicMode = False
Dim fLogChainedDetails : fLogChainedDetails = False
Dim fLogVerbose : fLogVerbose = False
Dim fListNonOfficeProducts : fListNonOfficeProducts = False
Dim fFileInventory : fFileInventory = False
Dim fFeatureTree : fFeatureTree = False
Dim fDisallowCScript : fDisallowCScript = False

'=======================================================================================================
'[INI] Section for script behavior customizations

'Directory for Log output.
'Example: "\\<server>\<share>\"
'Default: sPathOutputFolder = vbNullString -> %temp% directory is used
sPathOutputFolder = ""

'Quiet switch.
'Default: False -> Open inventory log when done
fQuiet = False


'Basic Mode
'Generates a basic list of installed Office products with licensing information only
'Disables all other extended analysis options
'Default: False -> Allow extended analysis options
fBasicMode = False

'Log full (verbose) details. This enables all possible scans for Office products.
'Default: False -> Only list standard details
fLogFull = False

'Enables additional logging details
'like the "FeatureTree" display and
'the additional individual listing of the chained Office SKU's
'Default: False -> Do not log additional details
fLogVerbose = False

'Starting with Office 2007 a SKU can contain several .msi packages
'The default does not list the details for each chained package
'This option allows to show the full details for each chained package
'Default: Comprehensive view - do not show details for chained packages
fLogChainedDetails = False

'The script filters for products of the Microsoft Office family.
'Set this option to 'True' to get a list of all Windows Installer products in the inventory log
'Default: False -> Don't list other products in the log
fListNonOfficeProducts = False

'File level inventory of installed Office products
'Depending on the number of installed products this can be an extremely time consuming task!
'Default: False -> Don't create a file level inventory
fFileInventory = False

'Detect all features of a product and include a feature tree in the log
'Default: False -> Don't include the feature detection
fFeatureTree = False

'DO NOT CUSTOMIZE BELOW THIS LINE!
'=======================================================================================================

'Measure total scan runtime
Dim tStart,tEnd
tStart = Time()
'Call the command line parser
ParseCmdLine
'Definition of non customizable settings
'Strings 
Dim sComputerName, sTemp, sCurUserSid, sDebugErr, sError, sErrBpa, sProductCodes_C2R
Dim sStack, sCacheLog, sLogFile, sInstalledProducts, sSystemType, sLogFormat
Dim sPackageGuid, sUserName, sDomain
'Arrays
Dim arrAllProducts(), arrMProducts(), arrUUProducts(), arrUMProducts(), arrMaster(), arrArpProducts()
Dim arrVirtProducts(), arrVirt2Products(), arrVirt3Products(), arrPatch(), arrAipPatch, arrMspFiles
Dim arrLog(4), arrLogFormat(), arrUUSids(), arrUMSids(), arrMVProducts(), arrIS(), arrFeature()
Dim arrProdVer09(), arrProdVer10(), arrProdVer11(), arrProdVer12(), arrProdVer14(), arrProdVer15()
Dim arrProdVer16(), arrFiles()
'Booleans
Dim fIsAdmin, fIsElevated, fIsCriticalError, fGuidCaseWarningOnly, f64, fPatchesOk, fPatchesExOk
Dim fCScript, bOsppInit, fZipError, fInitArrProdVer
'Integers
Dim iWiVersionMajor, iVersionNt, iPCount, iPatchesExError
'Dictionaries
Dim dicFolders, dicAssembly, dicMspIndex, dicProducts, dicArp, dicMissingChild
Dim dicPatchLevel, dicScenario, dicKeyComponents
Dim dicPolHKCU, dicPolHKLM
Dim dicProductCodeC2R, dicActiveC2Rv2Versions
Dim dicKeyComponentsV2, dicScenarioV2, dicC2RPropV2, dicVirt2Cultures
Dim dicKeyComponentsV3, dicScenarioV3, dicC2RPropV3, dicVirt3Cultures
'Other
Dim oMsi, oShell, oFso, oReg, oWsh
Dim TextStream, ShellApp, AppFolder, Ospp, Spp

'Identifier for product family
Const OFFICE_ALL                      = "78E1-11D2-B60F-006097C998E7}.0001-11D2-92F2-00104BC947F0}.6000-11D3-8CFE-0050048383C9}.6000-11D3-8CFE-0150048383C9}.7000-11D3-8CFE-0150048383C9}.BE5F-4ED1-A0F7-759D40C7622E}.BDCA-11D1-B7AE-00C04FB92F3D}.6D54-11D4-BEE3-00C04F990354}.CFDA-404E-8992-6AF153ED1719}.{9AC08E99-230B-47e8-9721-4577B7F124EA}"
'Office 2000 -> KB230848; Office XP -> KB302663; Office 2003 -> KB832672
Const OFFICE_2000                     = "78E1-11D2-B60F-006097C998E7}"
Const ORK_2000                        = "0001-11D2-92F2-00104BC947F0}"
Const PRJ_2000                        = "BDCA-11D1-B7AE-00C04FB92F3D}"
Const VIS_2002                        = "6D54-11D4-BEE3-00C04F990354}"
Const OFFICE_2002                     = "6000-11D3-8CFE-0050048383C9}"
Const OFFICE_2003                     = "6000-11D3-8CFE-0150048383C9}"
Const WSS_2                           = "7000-11D3-8CFE-0150048383C9}"
Const SPS_2003                        = "BE5F-4ED1-A0F7-759D40C7622E}"
Const PPS_2007                        = "CFDA-404E-8992-6AF153ED1719}" 'Project Portfolio Server 2007
Const POWERPIVOT_2010                 = "{72F8ECCE-DAB0-4C23-A471-625FEDABE323},{A37E1318-29CA-4A9F-9CCA-D9BFDD61D17B}" 'UpgradeCode!
Const O15_C2R                         = "{9AC08E99-230B-47e8-9721-4577B7F124EA}"
Const OFFICEID                        = "000-0000000FF1CE}" 'cover O12, O14 with 32 & 64 bit
Const OREGREFC2R15                    = "Microsoft Office 15"

Const PRODLEN                         = 13
Const FOR_READING                     = 1
Const FOR_WRITING                     = 2
Const FOR_APPENDING                   = 8
Const TRISTATE_USEDEFAULT             = -2 'Opens the file using the system default. 
Const TRISTATE_TRUE                   = -1 'Opens the file as Unicode. 
Const TRISTATE_FALSE                  = 0  'Opens the file as ASCII. 

Const USERSID_EVERYONE                = "s-1-1-0"
Const MACHINESID                      = ""
Const PRODUCTCODE_EMPTY               = ""
Const MSIOPENDATABASEMODE_READONLY    = 0
Const MSIOPENDATABASEMODE_PATCHFILE   = 32
Const MSICOLUMNINFONAMES              = 0
Const MSICOLUMNINFOTYPES              = 1
'Summary Information fields
Const PID_TITLE                       = 2 'Type of installer package. E.g. "Installation Database" or "Transform" or "Patch"
Const PID_SUBJECT                     = 3 'Displayname
Const PID_TEMPLATE                    = 7 'compatible platform and language versions for .msi / PatchTargets for .msp
Const PID_REVNUMBER                   = 9 'PackageCode
Const PID_WORDCOUNT                   = 15'InstallSource type 
Const MSIPATCHSTATE_UNKNOWN           = -1 'Patch is in an unknown state to this product instance. 
Const MSIPATCHSTATE_APPLIED           = 1 'Patch is applied to this product instance. 
Const MSIPATCHSTATE_SUPERSEDED        = 2 'Patch is applied to this product instance but is superseded.  
Const MSIPATCHSTATE_OBSOLETED         = 4 'Patch is applied in this product instance but obsolete.  
Const MSIPATCHSTATE_REGISTERED        = 8 'The enumeration includes patches that are registered but not yet applied.
Const MSIPATCHSTATE_ALL               = 15
Const MSIINSTALLCONTEXT_USERMANAGED   = 1
Const MSIINSTALLCONTEXT_USERUNMANAGED = 2
Const MSIINSTALLCONTEXT_MACHINE       = 4
Const MSIINSTALLCONTEXT_ALL           = 7
Const MSIINSTALLCONTEXT_C2RV2         = 8 'C2r V2 virtualized context
Const MSIINSTALLCONTEXT_C2RV3         = 15 'C2r V3 virtualized context
Const MSIINSTALLMODE_DEFAULT          = 0    'Provide the component and perform any installation necessary to provide the component. 
Const MSIINSTALLMODE_EXISTING         = -1   'Provide the component only if the feature exists. This option will verify that the assembly exists.
Const MSIINSTALLMODE_NODETECTION      = -2   'Provide the component only if the feature exists. This option does not verify that the assembly exists.
Const MSIINSTALLMODE_NOSOURCERESOLUTION = -3 'Provides the assembly only if the assembly is installed local.
 Const MSIPROVIDEASSEMBLY_NET          = 0    'A .NET assembly.
Const MSIPROVIDEASSMBLY_WIN32         = 1    'A Win32 side-by-side assembly.
Const MSITRANSFORMERROR_ALL           = 319
'Installstates for products, features, components
Const INSTALLSTATE_NOTUSED            = -7  ' component disabled
Const INSTALLSTATE_BADCONFIG          = -6  ' configuration data corrupt
Const INSTALLSTATE_INCOMPLETE         = -5  ' installation suspended or in progress
Const INSTALLSTATE_SOURCEABSENT       = -4  ' run from source, source is unavailable
Const INSTALLSTATE_MOREDATA           = -3  ' return buffer overflow
Const INSTALLSTATE_INVALIDARG         = -2  ' invalid function argument. The product/feature is neither advertised or installed. 
Const INSTALLSTATE_UNKNOWN            = -1  ' unrecognized product or feature
Const INSTALLSTATE_BROKEN             =  0  ' broken
Const INSTALLSTATE_ADVERTISED         = 1 'The product/feature is advertised but not installed. 
Const INSTALLSTATE_REMOVED            =  1 'The component is being removed (action state, not settable)
Const INSTALLSTATE_ABSENT             = 2 'The product/feature is not installed. 
Const INSTALLSTATE_LOCAL              = 3 'The product/feature/component is installed. 
Const INSTALLSTATE_SOURCE             = 4 'The product or feature is installed to run from source, CD, or network. 
Const INSTALLSTATE_DEFAULT            = 5 'The product or feature will be installed to use the default location: local or source.
Const INSTALLSTATE_VIRTUALIZED        = 8 'The product is virtualized (C2R).
Const VERSIONCOMPARE_LOWER            = -1 'Left hand file version is lower than right hand 
Const VERSIONCOMPARE_MATCH            =  0 'File versions are identical
Const VERSIONCOMPARE_HIGHER           =  1 'Left hand file versin is higher than right hand
Const VERSIONCOMPARE_INVALID          =  2 'Cannot compare. Invalid compare attempt.

Const COPY_OVERWRITE                  = &H10&
Const COPY_SUPPRESSERROR              = &H400& 

Const HKEY_CLASSES_ROOT               = &H80000000
Const HKEY_CURRENT_USER               = &H80000001
Const HKEY_LOCAL_MACHINE              = &H80000002
Const HKEY_USERS                      = &H80000003
Const HKCR                            = &H80000000
Const HKCU                            = &H80000001
Const HKLM                            = &H80000002
Const HKU                             = &H80000003
Const KEY_QUERY_VALUE                 = &H0001
Const KEY_SET_VALUE                   = &H0002
Const KEY_CREATE_SUB_KEY              = &H0004
Const DELETE                          = &H00010000
Const REG_SZ                          = 1
Const REG_EXPAND_SZ                   = 2
Const REG_BINARY                      = 3
Const REG_DWORD                       = 4
Const REG_MULTI_SZ                    = 7
Const REG_QWORD                       = 11
Const REG_GLOBALCONFIG                = "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\"
Const REG_CONTEXTMACHINE              = "Installer\"
Const REG_CONTEXTUSER                 = "Software\Microsoft\Installer\"
Const REG_CONTEXTUSERMANAGED          = "Software\Microsoft\Windows\CurrentVersion\Installer\Managed\"
Const REG_ARP                         = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
Const REG_OFFICE                      = "SOFTWARE\Microsoft\Office\"
Const REG_C2RVIRT_HKLM                = "\ClickToRun\REGISTRY\MACHINE\"

Const STR_NOTCONFIGURED             = "Not Configured"
Const STR_PACKAGEGUID               = "PackageGUID"
Const STR_REGPACKAGEGUID            = "RegPackageGUID"
Const STR_VERSION                   = "Version"
Const STR_PLATFORM                  = "Platform"
Const STR_CDNBASEURL                = "CDNBaseUrl"
Const STR_LASTUSEDBASEURL           = "Last used InstallSource"
Const STR_UPDATELOCATION            = "Custom UpdateLocation"
Const STR_USEDUPDATELOCATION        = "Winning UpdateLocation"
Const STR_UPDATESENABLED            = "UpdatesEnabled"
Const STR_UPDATEBRANCH              = "UpdateBranch"
Const STR_UPDATETOVERSION           = "UpdateToVersion"
Const STR_UPDATETHROTTLE            = "UpdatesThrottleValue"
Const STR_POLUPDATESENABLED         = "Policy UpdatesEnabled"
Const STR_POLUPGRADEENABLED         = "Policy EnableAutomaticUpgrade"
Const STR_POLUPDATEBRANCH           = "Policy UpdateBranch"
Const STR_POLUPDATELOCATION         = "Policy UpdateLocation"
Const STR_POLUPDATETOVERSION        = "Policy UpdateToVersion"
Const STR_POLUPDATEDEADLINE         = "Policy UpdateDeadline"
Const STR_POLUPDATENOTIFICATIONS    = "Policy UpdateHideNotifications"
Const STR_POLHIDEUPDATECFGOPT       = "Policy HideUpdateConfigOptions"
Const STR_SCA                       = "Shared Computer Activation"

Const GUID_UNCOMPRESSED               = 0
Const GUID_COMPRESSED                 = 1
Const GUID_SQUISHED                   = 2
Const LOGPOS_COMPUTER                 = 0 '    ArrLogPosition 0: "Computer"
Const LOGPOS_REVITEM                  = 1 '    ArrLogPosition 1: "Review Items"
Const LOGPOS_PRODUCT                  = 2 '    ArrLogPosition 2: "Products Inventory"
Const LOGPOS_RAW                      = 3 '    ArrLogPosition 3: "Raw Data"
Const LOGHEADING_NONE                 = 0 '    Not a heading
Const LOGHEADING_H1                   = 1 '    Heading 1 '='
Const LOGHEADING_H2                   = 2 '    Heading 2 '-'
Const LOGHEADING_H3                   = 3 '    Heading 3 ' '
Const TEXTINDENT                      = "                        "
Const CATEGORY                        = 1
Const TAG                             = 2
Const OSPP_ACTIVATIONTYPE = 0
Const OSPP_ID = 1
Const OSPP_APPLICATIONID = 2
Const OSPP_PARTIALPRODUCTKEY = 3
Const OSPP_DESCRIPTION = 4
Const OSPP_NAME = 5
Const OSPP_LICENSESTATUS = 6
Const OSPP_LICENSESTATUSREASON = 7
Const OSPP_PRODUCTKEYID = 8
Const OSPP_GRACEPERIODREMAINING = 9
Const OSPP_LICENSEFAMILY = 10
Const OSPP_DISCOVEREDKEYMANAGEMENTSERVICEMACHINENAME = 11
Const OSPP_DISCOVEREDKEYMANAGEMENTSERVICEMACHINEPORT = 12
Const OSPP_KEYMANAGEMENTSERVICEPORT = 13
Const OSPP_KEYMANGEMENTSERVICELOOKUPDOMAIN = 14
Const OSPP_VLACTIVATIONINTERVAL = 15
Const OSPP_VLRENEWALINTERVAL = 16
Const OSPP_ADACTIVATIONOBJECTNAME = 17
Const OSPP_ADACTIVATIONOBJECTDN = 18
Const OSPP_ADACTIVATIONCSVLPID = 19
Const OSPP_ADACTIVATIONCSVLSKUID = 20
Const OSPP_PRODUCTKEYID2 = 21
Const OSPP_MACHINEKEY  = 22

'Global_Access_Core - msaccess.exe
Const CID_ACC16_64 = "{27C919A6-3FA5-47F9-A3EC-BC7FF2AAD452}"
Const CID_ACC16_32 = "{E34AA7C4-8845-4BD7-BAC6-26554B60823B}"
Const CID_ACC15_64 = "{3CE2B4B3-DA38-4113-8DB2-965847CDE94F}"
Const CID_ACC15_32 = "{A3E12EF0-7C3B-4493-99A3-F92FCD0AA512}"
Const CID_ACC14_64 = "{02F5CBEC-E7B5-4FC1-BD72-6043152BD1D4}"
Const CID_ACC14_32 = "{AE393348-E564-4894-B8C5-EBBC5E72EFC6}"
Const CID_ACC12    = "{0638C49D-BB8B-4CD1-B191-054E8F325736}"
Const CID_ACC11    = "{F2D782F8-6B14-4FA4-8FBA-565CDDB9B2A8}"
'Global_Excel_Core - excel.exe
Const CID_XL16_64 = "{C4ACE6DB-AA99-401F-8BE6-8784BD09F003}"
Const CID_XL16_32 = "{C845E028-E091-442E-8202-21F596C559A0}"
Const CID_XL15_64 = "{58A9998B-6103-436F-85A1-52720802CA0A}"
Const CID_XL15_32 = "{107E1A9A-03AE-4F2B-ACF7-0CC519E60E7B}"
Const CID_XL14_64 = "{8B1BF0B4-A1CA-4656-AA46-D11C50BC55A4}"
Const CID_XL14_32 = "{538F6C89-2AD5-4006-8154-C6670774E980}"
Const CID_XL12    = "{0638C49D-BB8B-4CD1-B191-052E8F325736}"
Const CID_XL11    = "{A2B280D4-20FB-4720-99F7-40C09FBCE10A}"
'WAC_CoreSPD - spdesign.exe (frontpage.exe)
Const CID_SPD16_64 = "{2FB768AF-8F57-424A-BBDA-81611CFF3ED2}"
Const CID_SPD16_32 = "{C3F352B2-A43B-4948-AE54-12E265647697}"
Const CID_SPD15_64 = "{25B4430E-E7D6-406F-8468-D9B65BC240F3}"
Const CID_SPD15_32 = "{0F0A451D-CB3C-44BE-B8A4-E72C2B89C4A2}"
Const CID_SPD14_64 = "{6E4D3AA2-2AD9-4DD2-8C2D-8C55B656A5C9}"
Const CID_SPD14_32 = "{E5344AC3-915E-4655-AF0D-98BC878805DC}"
Const CID_SPD12    = "{0638C49D-BB8B-4CD1-B191-056E8F325736}"
Const CID_SPD11    = "{81E9830C-5A6B-436A-BEC9-4FB759282DE3}" ' FrontPage
'Groove_Core - groove.exe
Const CID_GRV16_64 = "{EEE31981-E2D9-45AE-B134-FD9276C19588}"
Const CID_GRV16_32 = "{6C26357C-A2D8-4C68-8BC6-A8091BECDA02}"
Const CID_GRV15_64 = "{AD8AD7F2-98CB-4257-BE7A-05CBCA1354B4}"
Const CID_GRV15_32 = "{87E86C36-1368-4841-9152-766F31BC46E8}"
Const CID_GRV14_64 = "{61CD70FF-C6B7-4F6A-8491-5B8B9B0040F8}"
Const CID_GRV14_32 = "{EFE67578-E52B-410E-9178-9911443DBF5A}"
Const CID_GRV12    = "{0A048D77-2DE9-4672-ACF7-12429662397D}"
'Lync_Corelync - lync.exe
Const CID_LYN16_64 = "{3CFF5AB2-9B16-4A31-BC3F-FAD761D92780}"
Const CID_LYN16_32 = "{E1AFBCD9-12F0-4FC0-9177-BFD3148AEC74}"
Const CID_LYN15_64 = "{D5B16A67-9FA6-4B77-AE2A-3B1F49CE9D3B}"
Const CID_LYN15_32 = "{F8D36F1C-6196-4FFA-94AA-736644D458E3}"
'Global_OneNote_Core - onenote.exe
Const CID_ONE16_64 = "{8265A5EF-46C7-4D46-812C-076F2A28F7CB}"
Const CID_ONE16_32 = "{2A8FA8D7-B728-4792-AC02-463FD7A423BD}"
Const CID_ONE15_64 = "{74F233A9-A17A-477C-905F-853F5FCDAD40}"
Const CID_ONE15_32 = "{DACE5A15-C57C-44DE-9AFF-89B4412485AF}"
Const CID_ONE14_64 = "{9542A6E5-2FAF-4191-B525-6ED00F2D0127}"
Const CID_ONE14_32 = "{62F8C897-D359-4D8F-9659-CF1E9E3E6B74}"
Const CID_ONE12    = "{0638C49D-BB8B-4CD1-B191-057E8F325736}"
Const CID_ONE11    = "{D2C0E18B-C463-4E90-92AC-CA94EBEC26CE}"
'Global_Office_Core - mso.dll
Const CID_MSO16_64 = "{625F5772-C1B3-497E-8ABE-7254EDB00506}"
Const CID_MSO16_32 = "{68477CB0-662A-48FB-AF2E-9573C92869F7}"
Const CID_MSO15_64 = "{D01398A1-F26F-4545-A441-567F097A57D7}"
Const CID_MSO15_32 = "{9CC2CF5E-9A2E-41AC-AF95-432890A9659A}"
Const CID_MSO14_64 = "{E6AC97ED-6651-4C00-A8FE-790DB0485859}"
Const CID_MSO14_32 = "{398E906A-826B-48DD-9791-549C649CACE5}"
Const CID_MSO12    = "{0638C49D-BB8B-4CD1-B191-050E8F325736}"
Const CID_MSO11    = "{A2B280D4-20FB-4720-99F7-10C09FBCE10A}"
'Global_Outlook_Core - outlook.exe
Const CID_OL16_64 = "{7C6D92EF-7B45-46E5-8670-819663220E4E}"
Const CID_OL16_32 = "{2C6C511D-4542-4E0C-95D0-05D4406032F2}"
Const CID_OL15_64 = "{3A5F96E7-F51D-4942-98DB-3CD037FB39E5}"
Const CID_OL15_32 = "{E9E5CFFC-AFFE-4F83-A695-7734FA4775B9}"
Const CID_OL14_64 = "{ECCC8A38-7855-46CA-88FB-3BAA7CD95E56}"
Const CID_OL14_32 = "{CFF13DD8-6EF2-49EB-B265-E3BFC6501C1D}"
Const CID_OL12    = "{0638C49D-BB8B-4CD1-B191-055E8F325736}"
Const CID_OL11    = "{3CE26368-6322-4ABF-B11B-458F5C450D0F}"
'Global_PowerPoint_Core - powerpnt.exe
Const CID_PPT16_64 = "{E0A76492-0FD5-4EC2-8570-AE1BAA61DC88}"
Const CID_PPT16_32 = "{9E73CEA4-29D0-4D16-8FB9-5AB17387C960}"
Const CID_PPT15_64 = "{8C1B8825-A280-4657-A7B8-8172C553A4C4}"
Const CID_PPT15_32 = "{258D5292-6DDA-4B39-B301-58405FA16638}"
Const CID_PPT14_64 = "{EE8D8E0A-D905-401D-9BC3-0D20156D5E30}"
Const CID_PPT14_32 = "{E72E0D20-0D63-438B-BC71-92AB9F9E8B54}"
Const CID_PPT12    = "{0638C49D-BB8B-4CD1-B191-053E8F325736}"
Const CID_PPT11    = "{C86C0B92-63C0-4E35-8605-281275C21F97}"
'Global_Project_ClientCore - winproj.exe
Const CID_PRJ16_64 = "{107BCD9A-F1DC-4004-A444-33706FC10058}"
Const CID_PRJ16_32 = "{0B6EDA1D-4A15-4F88-8B20-EA6528978E4E}"
Const CID_PRJ15_64 = "{760CE47D-9512-40D9-8C6D-CF232851B4BB}"
Const CID_PRJ15_32 = "{5296AE31-2F7D-480C-BFDC-CE0797426395}"
Const CID_PRJ14_64 = "{64A809BD-6EE9-475C-B4E8-95B0D7FF3B97}"
Const CID_PRJ14_32 = "{51894540-193D-40AE-83F9-D3FC5DB24D91}"
Const CID_PRJ12    = "{43C3CF66-AA31-476D-B029-6D274E46F86C}"
Const CID_PRJ11    = "{C33FFB81-6E54-4541-AFF4-D84DC60460F7}"
'Global_Publisher_Core - mspub.exe
Const CID_PUB16_64 = "{7ECBF2AA-14AA-4F89-B9A5-C064274CFA83}"
Const CID_PUB16_32 = "{81DD86EC-5F1C-4DDE-9211-98AF184EAD47}"
Const CID_PUB15_64 = "{22299AFF-DC4C-45A8-9A8F-651FB6467057}"
Const CID_PUB15_32 = "{C9C0167D-3FE0-4078-B47E-83272A4B8B04}"
Const CID_PUB14_64 = "{A716400F-5D5D-45CF-94B4-05B17A98B901}"
Const CID_PUB14_32 = "{CD0D7B29-89E7-49C5-8EE1-5D858EFF2593}"
Const CID_PUB12    = "{CD0D7B29-89E7-49C5-8EE1-5D858EFF2593}"
Const CID_PUB11    = "{0638C49D-BB8B-4CD1-B191-05CE8F325736}"
'Global_XDocs_Core - infopath.exe
Const CID_IP16_64 = "{2774AAC0-1433-46BE-993F-8088018C3B09}"
Const CID_IP15_64 = "{19AF7201-09A2-4C73-AB50-FCEF94CB2BA9}"
Const CID_IP15_32 = "{3741355B-72CF-4CEE-948E-CC9FBDBB8E7A}"
Const CID_IP14_64 = "{28B2FBA8-B95F-47CB-8F8F-0885ACDAC69B}"
Const CID_IP14_32 = "{E3898C62-6EC3-4491-8194-9C88AD716468}"
Const CID_IP12    = "{0638C49D-BB8B-4CD1-B191-058E8F325736}"
Const CID_IP11    = "{1A66B512-C4BE-4347-9F0C-8638F8D1E6E4}"
'Global_Visio_visioexe - visio.exe
Const CID_VIS16_64 = "{2D4540EC-2C88-4C28-AE88-2614B5460648}"
Const CID_VIS16_32 = "{A4C55BC1-B94C-4058-B15C-B9D4AE540AD1}"
Const CID_VIS15_64 = "{7069FF90-1D63-4F85-A2AB-6F0D01C78D83}"
Const CID_VIS15_32 = "{5D502092-1543-4D9B-89FE-7B4364417CC6}"
Const CID_VIS14_64 = "{DB2B19E4-F894-47B1-A6F1-9B391A4AE0A8}"
Const CID_VIS14_32 = "{4371C2B1-3F27-41F5-A849-9987AB91D990}"
Const CID_VIS12    = "{0638C49D-BB8B-4CD1-B191-05DE8F325736}"
Const CID_VIS11    = "{7E5F9F34-8EA7-4EA2-ABFB-CA4E742EFFA1}"
'Global_Word_Core - winword.exe
Const CID_WD16_64 = "{DC5CCACD-A7AC-4FD3-9F70-9454B5DE5161}"
Const CID_WD16_32 = "{30CAC893-3CA4-494C-A5E9-A99141352216}"
Const CID_WD15_64 = "{6FF09BDF-B087-4E23-A9B9-272DBFD64099}"
Const CID_WD15_32 = "{09D07EFC-505F-4D9C-BFD5-ACE3217F6654}"
Const CID_WD14_64 = "{C0AC079D-A84B-4CBD-8DBA-F1BB44146899}"
Const CID_WD14_32 = "{019C826E-445A-4649-A5B0-0BF08FCC4EEE}"
Const CID_WD12    = "{0638C49D-BB8B-4CD1-B191-051E8F325736}"
Const CID_WD11    = "{1EBDE4BC-9A51-4630-B541-2561FA45CCC5}"

'Arrays
Const UBOUND_LOGARRAYS                = 12
Const UBOUND_LOGCOLUMNS               = 31 ' Controlled by array with the most columns
Redim arrLogFormat(UBOUND_LOGARRAYS,UBOUND_LOGCOLUMNS)

Const ARRAY_MASTER                    = 0 'Master data array id
    Const UBOUND_MASTER               = 31
' ProductCode
    Const COL_PRODUCTCODE             = 0
    arrLogFormat(ARRAY_MASTER,COL_PRODUCTCODE)     = "ProductCode"
' Msi ProductName
    Const COL_PRODUCTNAME             = 1
    arrLogFormat(ARRAY_MASTER,COL_PRODUCTNAME)     = "Msi ProductName"
' UserSid
    Const COL_USERSID                 = 2
    arrLogFormat(ARRAY_MASTER,COL_USERSID)         = "UserSid"
' ProductContext
    Const COL_CONTEXTSTRING           = 3
    arrLogFormat(ARRAY_MASTER,COL_CONTEXTSTRING)   = "ProductContext"
' ProductState
    Const COL_STATESTRING             = 4
    arrLogFormat(ARRAY_MASTER,COL_STATESTRING)     = "ProductState"
' ProductContext
    Const COL_CONTEXT                 = 5
    arrLogFormat(ARRAY_MASTER,COL_CONTEXT)         = "ProductContext"
' ProductState
    Const COL_STATE                   = 6
    arrLogFormat(ARRAY_MASTER,COL_STATE)           = "ProductState"
' Arp SystemComponent
    Const COL_SYSTEMCOMPONENT         = 7
    arrLogFormat(ARRAY_MASTER,COL_SYSTEMCOMPONENT) = "Arp SystemComponent"
' Arp ParentCount
    Const COL_ARPPARENTCOUNT          = 8
    arrLogFormat(ARRAY_MASTER,COL_ARPPARENTCOUNT)  = "Arp ParentCount"
' Arp Parents
    Const COL_ARPPARENTS              = 9
    arrLogFormat(ARRAY_MASTER,COL_ARPPARENTS)      = "Configuration SKU"
' Arp Productname
    Const COL_ARPPRODUCTNAME          = 10
    arrLogFormat(ARRAY_MASTER,COL_ARPPRODUCTNAME)  = "ARP ProductName"
' ProductVersion
    Const COL_PRODUCTVERSION          = 11
    arrLogFormat(ARRAY_MASTER,COL_PRODUCTVERSION)  = "ProductVersion"
' ServicePack Level
    Const COL_SPLEVEL                 = 12
    arrLogFormat(ARRAY_MASTER,COL_SPLEVEL)         = "ServicePack Level"
' InstallDate
    Const COL_INSTALLDATE             = 13
    arrLogFormat(ARRAY_MASTER,COL_INSTALLDATE)     = "InstallDate"
' Cached .msi 
    Const COL_CACHEDMSI               = 14
    arrLogFormat(ARRAY_MASTER,COL_CACHEDMSI)       = "Cached .msi Package"
' Original .msi name
    Const COL_ORIGINALMSI             = 15
    arrLogFormat(ARRAY_MASTER,COL_ORIGINALMSI)     = "Original .msi Name"
' Build/Origin Property
    Const COL_ORIGIN                  = 16
    arrLogFormat(ARRAY_MASTER,COL_ORIGIN)          = "Build/Origin"
' ProductID Property
    Const COL_PRODUCTID               = 17
    arrLogFormat(ARRAY_MASTER,COL_PRODUCTID)       = "ProductID (MSI)"
' Package Code
    Const COL_PACKAGECODE             = 18
    arrLogFormat(ARRAY_MASTER,COL_PACKAGECODE)     = "Package Code"
' Transform
    Const COL_TRANSFORMS              = 19
    arrLogFormat(ARRAY_MASTER,COL_TRANSFORMS)      = "Transforms"
' Architecture
    Const COL_ARCHITECTURE            = 20
    arrLogFormat(ARRAY_MASTER,COL_ARCHITECTURE)    = "Architecture"
' Error
    Const COL_ERROR                   = 21
    arrLogFormat(ARRAY_MASTER,COL_ERROR)           = "Errors"
' Notes
    Const COL_NOTES                   = 22
    arrLogFormat(ARRAY_MASTER,COL_NOTES)           = "Notes"
' MetadataState
    Const COL_METADATASTATE           = 23
    arrLogFormat(ARRAY_MASTER,COL_METADATASTATE)   = "MetadataState"
' IsOfficeProduct
    Const COL_ISOFFICEPRODUCT         = 24
    arrLogFormat(ARRAY_MASTER,COL_ISOFFICEPRODUCT) = "IsOfficeProduct"
' PatchFamily
    Const COL_PATCHFAMILY             = 25
    arrLogFormat(ARRAY_MASTER,COL_PATCHFAMILY)     = "PatchFamily"
' OSPP License
    Const COL_OSPPLICENSE             = 26
    arrLogFormat(ARRAY_MASTER,COL_OSPPLICENSE)     = "OSPP License"
    Const COL_OSPPLICENSEXML          = 27
    arrLogFormat(ARRAY_MASTER, COL_OSPPLICENSEXML) = "OSPP License XML"
' UpgradeCode
    Const COL_UPGRADECODE             = 28
    arrLogFormat(ARRAY_MASTER,COL_UPGRADECODE)     = "UpgradeCode"
' Virtualized
    Const COL_VIRTUALIZED             = 29
    arrLogFormat(ARRAY_MASTER,COL_VIRTUALIZED)     = "Virtualized"
' InstallType
    Const COL_INSTALLTYPE             = 30
    arrLogFormat(ARRAY_MASTER,COL_INSTALLTYPE)     = "InstallType"
    ' KeyComponents
    Const COL_KEYCOMPONENTS           = 31
    arrLogFormat(ARRAY_MASTER,COL_KEYCOMPONENTS)   = "KeyComponents"

Const ARRAY_PATCH                     = 1 'Patch data array id
    Const PATCH_COLUMNCOUNT           = 14
    Const PATCH_LOGSTART              = 1
    Const PATCH_LOGCHAINEDMAX         = 8
    Const PATCH_LOGMAX                = 11
' Product
    Const PATCH_PRODUCT               = 0
    arrLogFormat(ARRAY_PATCH,PATCH_PRODUCT)        = "Patched Product: "
' KB
    Const PATCH_KB                    = 1
    arrLogFormat(ARRAY_PATCH,PATCH_KB)             = "KB: "
' PackageName
    Const PATCH_PACKAGE               = 3
    arrLogFormat(ARRAY_PATCH,PATCH_PACKAGE)        = "Package: "
' PatchState
    Const PATCH_PATCHSTATE            = 2
    arrLogFormat(ARRAY_PATCH,PATCH_PATCHSTATE)     = "State: "
' Sequence
    Const PATCH_SEQUENCE              = 4
    arrLogFormat(ARRAY_PATCH,PATCH_SEQUENCE)       = "Sequence: "
' Uninstallable
    Const PATCH_UNINSTALLABLE         = 5
    arrLogFormat(ARRAY_PATCH,PATCH_UNINSTALLABLE)  = "Uninstallable: "
' InstallDate
    Const PATCH_INSTALLDATE           = 6
    arrLogFormat(ARRAY_PATCH,PATCH_INSTALLDATE)    = "InstallDate: "
' PatchCode
    Const PATCH_PATCHCODE             = 7
    arrLogFormat(ARRAY_PATCH,PATCH_PATCHCODE)      = "PatchCode: "
' LocalPackage
    Const PATCH_LOCALPACKAGE          = 8
    arrLogFormat(ARRAY_PATCH,PATCH_LOCALPACKAGE)   = "LocalPackage: "
' PatchTransform
    Const PATCH_TRANSFORM             = 9
    arrLogFormat(ARRAY_PATCH,PATCH_TRANSFORM)      = "PatchTransform: "
' DisplayName
    Const PATCH_DISPLAYNAME           = 10
    arrLogFormat(ARRAY_PATCH,PATCH_DISPLAYNAME)    = "DisplayName: "
' MoreInfoUrl
    Const PATCH_MOREINFOURL           = 11
    arrLogFormat(ARRAY_PATCH,PATCH_MOREINFOURL)    = "MoreInfoUrl: "
' Client side patch or patched AIP
    Const PATCH_CSP                   = 12
    arrLogFormat(ARRAY_PATCH,PATCH_CSP)            = "ClientSidePatch: "
' Local .msp package OK/available
    Const PATCH_CPOK                  = 13
    arrLogFormat(ARRAY_PATCH,PATCH_CPOK)           = "CachedMspOK: "

' arrMsp(MSPDEFAULT,MSP_COLUMNCOUNT)
Const ARRAY_MSPFILES                  = 10
    Const MSPFILES_COLUMNCOUNT           = 18
    Const MSPFILES_LOGMAX                = 9
' Product
    Const MSPFILES_TARGETS               = 0
    arrLogFormat(ARRAY_MSPFILES,MSPFILES_TARGETS)        = "Patch Targets: "
' KB
    Const MSPFILES_KB                    = 1
    arrLogFormat(ARRAY_MSPFILES,MSPFILES_KB)             = "KB: "
' PackageName
    Const MSPFILES_PACKAGE               = 2
    arrLogFormat(ARRAY_MSPFILES,MSPFILES_PACKAGE)        = "Package: "
' Family
    Const MSPFILES_FAMILY                = 3
    arrLogFormat(ARRAY_MSPFILES,MSPFILES_FAMILY)         = "Family: "
' Sequence
    Const MSPFILES_SEQUENCE              = 4
    arrLogFormat(ARRAY_MSPFILES,MSPFILES_SEQUENCE)       = "Sequence: "
' PatchState
    Const MSPFILES_PATCHSTATE            = 5
    arrLogFormat(ARRAY_MSPFILES,MSPFILES_PATCHSTATE)     = "PatchState: "
' Uninstallable
    Const MSPFILES_UNINSTALLABLE         = 6
    arrLogFormat(ARRAY_MSPFILES,MSPFILES_UNINSTALLABLE)  = "Uninstallable: "
' InstallDate
    Const MSPFILES_INSTALLDATE           = 7
    arrLogFormat(ARRAY_MSPFILES,MSPFILES_INSTALLDATE)    = "InstallDate: "
' DisplayName
    Const MSPFILES_DISPLAYNAME           = 8
    arrLogFormat(ARRAY_MSPFILES,MSPFILES_DISPLAYNAME)    = "DisplayName: "
' MoreInfoUrl
    Const MSPFILES_MOREINFOURL           = 9
    arrLogFormat(ARRAY_MSPFILES,MSPFILES_MOREINFOURL)    = "MoreInfoUrl: "
' PatchCode
    Const MSPFILES_PATCHCODE             = 10
    arrLogFormat(ARRAY_MSPFILES,MSPFILES_PATCHCODE)      = "PatchCode: "
' LocalPackage
    Const MSPFILES_LOCALPACKAGE          = 11
    arrLogFormat(ARRAY_MSPFILES,MSPFILES_LOCALPACKAGE)   = "LocalPackage: "
' Bucket
    Const MSPFILES_BUCKET                = 12
    arrLogFormat(ARRAY_MSPFILES,MSPFILES_BUCKET)         = "Bucket: "
' Attribute msidbPatchSequenceSupersedeEarlier
    Const MSPFILES_ATTRIBUTE             = 14
    arrLogFormat(ARRAY_MSPFILES,MSPFILES_ATTRIBUTE)      = "msidbPatchSequenceSupersedeEarlier: "
' PatchTransform
    Const MSPFILES_TRANSFORM             = 14
    arrLogFormat(ARRAY_MSPFILES,MSPFILES_TRANSFORM)      = "PatchTransform: "
' PatchXml
    Const MSPFILES_XML                   = 15
    arrLogFormat(ARRAY_MSPFILES,MSPFILES_XML)            = "PatchXml: "
' PatchTables
    Const MSPFILES_TABLES                = 16
    arrLogFormat(ARRAY_MSPFILES,MSPFILES_TABLES)         = "PatchTables: "
' Local .msp package OK/available
    Const MSPFILES_CPOK                  = 17
    arrLogFormat(ARRAY_MSPFILES,MSPFILES_CPOK)           = "CachedMspOK: "

Const ARRAY_AIPPATCH                  = 11
    Const AIPPATCH_COLUMNCOUNT        = 3
' Product
    Const AIPPATCH_PRODUCT            = 0
    arrLogFormat(ARRAY_AIPPATCH,AIPPATCH_PRODUCT)  = "Patched Product: "
' PatchCode
    Const AIPPATCH_PATCHCODE          = 1
    arrLogFormat(ARRAY_AIPPATCH,AIPPATCH_PATCHCODE)= "PatchCode: "
' DisplayName
    Const AIPPATCH_DISPLAYNAME        = 3
    arrLogFormat(ARRAY_AIPPATCH,AIPPATCH_DISPLAYNAME)= "DisplayName: "

Const ARRAY_FEATURE                   = 2 'Feature data array id
    Const FEATURE_COLUMNCOUNT         = 1
    Const FEATURE_PRODUCTCODE         = 0
    Const FEATURE_TREE                = 1

Const ARRAY_ARP                       = 4 'Add/remove products data array id
    Const ARP_CHILDOFFSET             = 6
' Config Productcode
    Const ARP_CONFIGPRODUCTCODE       = 0
    arrLogFormat(ARRAY_ARP,ARP_CONFIGPRODUCTCODE)  = "Config ProductCode"
' Config Productname
    Const COL_CONFIGNAME              = 1
    arrLogFormat(ARRAY_ARP,COL_CONFIGNAME)         = "Config ProductName"
' Config ProductVersion
    Const ARP_PRODUCTVERSION          = 2
    arrLogFormat(ARRAY_ARP,ARP_PRODUCTVERSION)     = "ProductVersion"
' Config InstallType
    Const COL_CONFIGINSTALLTYPE       = 3
    arrLogFormat(ARRAY_ARP,COL_CONFIGINSTALLTYPE)  = "Config InstallType"
' Config PackageID
    Const COL_CONFIGPACKAGEID         = 4
    arrLogFormat(ARRAY_ARP,COL_CONFIGPACKAGEID)    = "Config PackageID"
    Const COL_ARPALLPRODUCTS          = 5
    Const COL_LBOUNDCHAINLIST         = 6

Const ARRAY_IS                        = 5 ' MSI InstallSource data array id
    Const UBOUND_IS                   = 6
    Const IS_LOG_LBOUND               = 2
    Const IS_LOG_UBOUND               = 6
    Const IS_PRODUCTCODE              = 0
    Const IS_SOURCETYPE               = 1
    Const IS_SOURCETYPESTRING         = 2
    arrLogFormat(ARRAY_IS,IS_SOURCETYPESTRING)     = "InstallSource Type"
    Const IS_ORIGINALSOURCE           = 3
    arrLogFormat(ARRAY_IS,IS_ORIGINALSOURCE)       = "Initially Used Source"
    Const IS_LASTUSEDSOURCE           = 4
    arrLogFormat(ARRAY_IS,IS_LASTUSEDSOURCE)       = "Last Used Source"
    Const IS_LISRESILIENCY            = 5
    arrLogFormat(ARRAY_IS,IS_LISRESILIENCY)        = "LIS Resiliency Sources"
    Const IS_ADDITIONALSOURCES        = 6
    arrLogFormat(ARRAY_IS,IS_ADDITIONALSOURCES)    = "Network Sources"

Const ARRAY_VIRTPROD                  = 6 ' Non MSI based virtualized products
    Const UBOUND_VIRTPROD             = 10

' Productcode
    arrLogFormat(ARRAY_VIRTPROD, COL_PRODUCTCODE)       = "ProductCode"
' Productname
    arrLogFormat(ARRAY_VIRTPROD, COL_PRODUCTNAME)       = "ProductName"
' KeyName
    Const VIRTPROD_KEYNAME          = 2
    arrLogFormat(ARRAY_VIRTPROD, VIRTPROD_KEYNAME)      = "KeyName"
' ConfigName
    Const VIRTPROD_CONFIGNAME       = 3
    arrLogFormat(ARRAY_VIRTPROD, VIRTPROD_CONFIGNAME)   = "Config ProductName"
' ProductVersion
    Const VIRTPROD_PRODUCTVERSION   = 4
    arrLogFormat(ARRAY_VIRTPROD, VIRTPROD_PRODUCTVERSION) = "ProductVersion"
' Service Pack Level
    Const VIRTPROD_SPLEVEL          = 5
    arrLogFormat(ARRAY_VIRTPROD, VIRTPROD_SPLEVEL)      = "ServicePack Level"
' (O)SPP License
    Const VIRTPROD_OSPPLICENSE      = 6
    arrLogFormat(ARRAY_VIRTPROD, VIRTPROD_OSPPLICENSE)  = "OSPP License"
' (O)SPP License XML
    Const VIRTPROD_OSPPLICENSEXML   = 7
    arrLogFormat(ARRAY_VIRTPROD, VIRTPROD_OSPPLICENSEXML) = "OSPP License XML"
' Child Packages
    Const VIRTPROD_CHILDPACKAGES    = 8
    arrLogFormat(ARRAY_VIRTPROD, VIRTPROD_CHILDPACKAGES) = "Child Packages"
' KeyComponents
    Const VIRTPROD_KEYCOMPONENTS    = 9
    arrLogFormat(ARRAY_VIRTPROD, VIRTPROD_KEYCOMPONENTS) = "KeyComponents"
' Excluded Applications
    Const VIRTPROD_EXCLUDEAPP       = 10
    arrLogFormat(ARRAY_VIRTPROD, VIRTPROD_EXCLUDEAPP) = "Excluded Applications"

Const CSV                             = ", "
Const DSV                             = " - "
Const DOT                             = ". "
Const ERR_CATEGORYNOTE                = "Note: "
Const ERR_CATEGORYWARN                = "Warning: "
Const ERR_CATEGORYERROR               = "Error: "
Const ERR_NONADMIN                    = "The script appears to run outside administrator context"
Const ERR_NONELEVATED                 = "The script does not appear to run elevated"
Const ERR_DATAINTEGRITY               = "A script internal error occurred. The integrity of the logged data might be affected"
Const ERR_OBJPRODUCTINFO              = "Installer.ProductInfo -> "
Const ERR_INITSUMINFO                 = "Could not connect to summary information stream"
Const ERR_NOARRAY                     = "Array check failed"
Const ERR_UNKNOWNHANDLER              = "Unknown Error Handler: '"
Const ERR_PRODUCTSEXALL               = "ProductsEx for MSIINSTALLCONTEXT_ALL failed"
Const ERR_PATCHESEX                   = "PatchesEx failed to get a list of patches for: "
Const ERR_PATCHES                     = "Installer.Patches failed to get a list of patches"
Const ERR_MISSINGCHILD                = "A chained product is missing which breaks the ability to maintain or uninstall this product. "
Const ERR_ORPHANEDITEM                = "Office application without entry point in Add/Remove Programs"
Const ERR_INVALIDPRODUCTCODE          = "Critical Windows Installer metadata corruption detected 'Invalid ProductCode'"
Const ERR_INVALIDGUID                 = "GUID validation failed"
Const ERR_INVALIDGUIDCHAR             = "Guid contains invalid character(s)"
Const ERR_INVALIDGUIDLENGTH           = "Invalid length for GUID  "
Const ERR_GUIDCASE                    = "Guid contains lower case character(s)"
Const ERR_BADARPMETADATA              = "Crititcal ARP metadata corruption detected in key: "
Const ERR_OFFSCRUB_TERMINATED         = "Bad ARP metadata. This can be caused by an OffScrub run that was terminated before it could complete:"
Const ERR_ARPENTRYMISSING             = "Expected regkey not present for ARP config parent"
Const ERR_REGKEYMISSING               = "Regkey does not exist: "
Const ERR_CUSTOMSTACKCORRUPTION       = "Custom stack list string corrupted"
Const ERR_BADMSPMETADATA              = "Metadata mismatch for patch registration"
Const ERR_BADMSINAMEMETADATA          = "Failed to retrieve value for original .msi name"
Const ERR_BADPACKAGEMETADATA          = "Failed to retrieve value for cached .msi package"
Const ERR_PACKAGEAPIFAILURE           = "API failed to retrieve value for cached .msi package"
Const ERR_BADPACKAGECODEMETADATA      = "Failed to retrieve value for Package Code"
Const ERR_PACKAGECODEMISMATCH         = "PackageCode mismatch between registered value and cached .msi"
Const ERR_LOCALPACKAGEMISSING         = "Local cached .msi appears to be missing"
Const ERR_BADTRANSFORMSMETADATA       = "Failed to retrieve value for Transforms"
Const ERR_SICONNECTFAILED             = "Failed to connect to SummaryInformation stream"
Const ERR_MSPOPENFAILED               = "OpenDatabase failed to open .msp file "
Const ERR_MSIOPENFAILED               = "OpenDatabase failed to open .msi file "
Const ERR_BADFILESTATE                = " has unexpected file state(s). "
Const ERR_FILEVERSIONLOW              = "Review file versions for product "
Const BPA_GUID                        = "For details on 'GUID' see https://msdn.microsoft.com/en-us/library/Aa368767.aspx"
Const BPA_PACKAGECODE                 = "For details on 'Package Codes' see https://msdn.microsoft.com/en-us/library/aa370568.aspx"
Const BPA_PRODUCTCODE                 = "For details on 'Product Codes' see https://msdn.microsoft.com/en-us/library/aa370860.aspx"
Const BPA_PACKAGECODEMISMATCH         = "A mismatch of the PackageCode will force the Windows Installer to recache the local .msi from the InstallSource. For details on 'Package Code' see https://msdn.microsoft.com/en-us/library/aa370568.aspx"
'=======================================================================================================

Main

'=======================================================================================================
'Module Main
'=======================================================================================================
Sub Main
    Dim fCheckPreReq, FsoLogFile, FsoXmlLogFile
    On Error Resume Next
' Check type of scripting host
    fCScript = (UCase(Mid(Wscript.FullName, Len(Wscript.Path) + 2, 1)) = "C")
' Ensure all required objects are available. Prerequisite Checks with inline error handling. 
    fCheckPreReq = CheckPreReq()
    If fCheckPreReq = False Then Exit Sub
' Initializations 
    Initialize 
' Get computer specific properties
    If fCScript AND NOT fQuiet Then wscript.echo "Stage 1 of 11: ComputerProperties"
    ComputerProperties 
' Build an array with a list of all Windows Installer based products
' After this point the master array "arrMaster" is instantiated and has the basic product details
    If fCScript AND NOT fQuiet Then wscript.echo "Stage 2 of 11: Product detection"
    FindAllProducts 
' Get additional product properties and add them to the master array
    If fCScript AND NOT fQuiet Then wscript.echo "Stage 3 of 11: ProductProperties"
    ProductProperties 
' Build an array with data on the InstallSource(s)
    If fCScript AND NOT fQuiet Then wscript.echo "Stage 4 of 11: InstallSources"
    ReadMsiInstallSources
' Build an array with data from Add/Remove Products
' Only Office >= 2007 products that use a multiple .msi structure will be covered here
    If fCScript AND NOT fQuiet Then wscript.echo "Stage 5 of 11: Add/Remove Programs analysis"
    ARPData 
' Add Licensing data
' Only Office >= 2010 products that use OSPP are covered here
    If fCScript AND NOT fQuiet Then wscript.echo "Stage 6 of 11: Licensing (OSPP)"
    OsppCollect 
' Build an array with all patch data.
    If fCScript AND NOT fQuiet Then wscript.echo "Stage 7 of 11: Patch detection"
    FindAllPatches
    If fFeatureTree Then
    ' Build a tree structure for the Features 
        If fCScript AND NOT fQuiet Then wscript.echo "Stage 8 of 11: FeatureStates"
        FindFeatureStates
    Else
        If fCScript AND NOT fQuiet Then wscript.echo "Skipping stage 8 of 11: (FeatureStates)"
    End If 'fFeatureTree
' Create file inventory XML files
    If fFileInventory Then
        If fCScript AND NOT fQuiet Then wscript.echo "Stage 9 of 11: FileInventory"
        FileInventory
    Else
        If fCScript AND NOT fQuiet Then wscript.echo "Skipping stage 9 of 11: (FileInventory)"
    End If 'fFileInventory
' Prepare the collected data for the output file
    If fCScript AND NOT fQuiet Then wscript.echo "Stage 10 of 11: Prepare collected data for output"
    PrepareLog sLogFormat 
' Write the output file
    If fCScript AND NOT fQuiet Then wscript.echo "Stage 11 of 11: Write log"
    WriteLog 
    Set FsoLogFile = oFso.GetFile(sLogFile)
    Set FsoXmlLogFile = oFso.GetFile(sPathOutputFolder & sComputerName & "_ROIScan.xml")
    If fFileInventory Then 
        If (oFso.FileExists(sPathOutputFolder & sComputerName & "_ROIScan.zip") AND NOT fZipError) Then
            CopyToZip ShellApp.NameSpace(sPathOutputFolder & sComputerName & "_ROIScan.zip"), FsoLogFile
            CopyToZip ShellApp.NameSpace(sPathOutputFolder & sComputerName & "_ROIScan.zip"), FsoXmlLogFile
        End If
    End If
    If fCScript AND NOT fQuiet Then wscript.echo "Done!"
' Open the output file
    If Not fQuiet Then
        Set oShell = CreateObject("WScript.Shell")
        If fFileInventory Then 
            If oFso.FileExists(sPathOutputFolder & sComputerName & "_ROIScan.zip") AND NOT fZipError Then
                oShell.Run "explorer /e," & chr(34) & sPathOutputFolder&sComputerName & "_ROIScan.zip" & chr(34) 
            Else 
                oShell.Run "explorer /e," & chr(34) & sPathOutputFolder & "ROIScan" & chr(34)
            End If
        End If 'fFileInventory
        oShell.Run chr(34) & sLogFile & chr(34)
        Set oShell = Nothing
    End If 'fQuiet
' Clear up Objects
    CleanUp
End Sub
'=======================================================================================================

'Initialize defaults, setting and collect current user information
Sub Initialize
    Dim oApp, oWmiLocal, Process, Processes
    Dim sEnvVar, Argument
    Dim iPopup, iInstanceCnt
    Dim fPrompt
    On Error Resume Next
    'Ensure there's only a single instance running of this script
    iInstanceCnt = 0
    Set oWmiLocal = GetObject("winmgmts:\\.\root\cimv2")
    wscript.sleep 500
    Set Processes = oWmiLocal.ExecQuery("Select * From Win32_Process")
    For Each Process in Processes
        If LCase(Mid(Process.Name, 2, 6)) = "script" Then 
            If InStr(LCase(Process.CommandLine), "roiscan") > 0 AND NOT InStr(Process.CommandLine," UAC") > 0 Then iInstanceCnt = iInstanceCnt + 1
        End If
    Next 'Process
    If iInstanceCnt > 1 Then
        If NOT fQuiet Then wscript.echo "Error: Another instance of this script is already running."
        wscript.quit
    End If
   
    'Other defaults
    Set dicPatchLevel = CreateObject("Scripting.Dictionary")
    Set dicScenario = CreateObject("Scripting.Dictionary")
    Set dicC2RPropV2 = CreateObject("Scripting.Dictionary")
    Set dicScenarioV2 = CreateObject("Scripting.Dictionary")
    Set dicKeyComponentsV2 = CreateObject("Scripting.Dictionary")
    Set dicVirt2Cultures = CreateObject("Scripting.Dictionary")
    Set dicC2RPropV3 = CreateObject("Scripting.Dictionary")
    Set dicScenarioV3 = CreateObject("Scripting.Dictionary")
    Set dicKeyComponentsV3 = CreateObject("Scripting.Dictionary")
    Set dicVirt3Cultures = CreateObject("Scripting.Dictionary")
    Set dicProductCodeC2R = CreateObject("Scripting.Dictionary")
    Set dicArp = CreateObject("Scripting.Dictionary")
    Set dicPolHKCU = CreateObject("Scripting.Dictionary")
    Set dicPolHKLM = CreateObject("Scripting.Dictionary")
    fZipError = False
    fInitArrProdVer = False

    ' log output folder
    If sPathOutputFolder = "" Then sPathOutputFolder = "%TEMP%"
    sPathOutputFolder = oShell.ExpandEnvironmentStrings(sPathOutputFolder)
    sPathOutputFolder = Replace(sPathOutputFolder, "'", "")
    If sPathOutputFolder = "." OR Left(sPathOutputFolder, 2) = ".\" Then sPathOutputFolder = GetFullPathFromRelative(sPathOutputFolder)
    sPathOutputFolder = oFso.GetAbsolutePathName(sPathOutputFolder)
    If Trim(UCase(sPathOutputFolder)) = "DESKTOP" Then
        Set oApp = CreateObject ("Shell.Application")
        Const DESKTOP = &H10&
        sPathOutputFolder = oApp.Namespace(DESKTOP).Self.Path
        Set oApp = Nothing
    End If
    If Not oFso.FolderExists(sPathOutputFolder) Then 
        ' custom log folder location does not exist.
        ' try to create the folder before falling back to default
        oFso.CreateFolder sPathOutputFolder
        If NOT Err = 0 Then
            sPathOutputFolder = oShell.ExpandEnvironmentStrings("%TEMP%") & "\" 
            Err.Clear
        End If
    End If
    While Right(sPathOutputFolder, 1) = "\" 
        sPathOutputFolder = Left(sPathOutputFolder, Len(sPathOutputFolder) - 1)
    Wend
    If Not Right(sPathOutputFolder, 1) = "\" Then sPathOutputFolder = sPathOutputFolder & "\"
    sLogFile = sPathOutputFolder & sComputerName & "_ROIScan.log"

    CacheLog LOGPOS_COMPUTER,LOGHEADING_H1,Null,"Computer" 
    CacheLog LOGPOS_REVITEM,LOGHEADING_H1,Null,"Review Items" 
'CacheLog LOGPOS_PRODUCT,LOGHEADING_H1,Null,"Products Inventory" 
    CacheLog LOGPOS_RAW,LOGHEADING_H1,Null,"Raw Data" 
    iPopup = -1
    fPrompt = True
    If Wscript.Arguments.Count > 0 Then
        For Each Argument in Wscript.Arguments
            If Argument = "UAC" Then fPrompt = False
        Next 'Argument
    End If
'Add warning to log if non-admin was detected
    If Not fIsAdmin Then 
        Cachelog LOGPOS_REVITEM,LOGHEADING_NONE,ERR_CATEGORYERROR,ERR_NONADMIN
        If NOT fQuiet AND fPrompt Then RelaunchElevated
    End If
    If fIsAdmin AND (NOT fIsElevated) Then 
        Cachelog LOGPOS_REVITEM,LOGHEADING_NONE,ERR_CATEGORYERROR,ERR_NONELEVATED
        If NOT fQuiet AND fPrompt Then RelaunchElevated
   End If
'Ensure CScript as engine
    If (NOT UCase(Mid(Wscript.FullName, Len(Wscript.Path) + 2, 1)) = "C") AND (NOT fDisallowCScript) Then RelaunchAsCScript
'Check on 64 bit OS -> see CheckPreReq
'Init sCurUserSid
    GetUserSids("Current") 
'Init "arrUUSids"
    Redim arrUUSids(-1)
    GetUserSids ("UserUnmanaged") 
'Init "arrUMSids"
    Redim arrUMSids(-1)
    GetUserSids("UserManaged") 
'Init KeyComponents dictionary
    Set dicKeyComponents = CreateObject("Scripting.Dictionary")
    InitKeyComponents
'Set defaults for ProductList arrays
    InitPLArrays 
    Set dicMspIndex = CreateObject("Scripting.Dictionary")
    bOsppInit = False
End Sub 'Initialize
'=======================================================================================================
'End Of Main Module

'=======================================================================================================
'Module FileInventory
'=======================================================================================================

'File inventory for installed applications
Sub FileInventory
    Dim sProductCode,sQueryFT,sQueryMst,sQueryCompID,sQueryFC,sMst,sFtk,sCmp
    Dim sPath,sName,sXmlLine,sAscCheck,sQueryDir,sQueryAssembly,sPatchCode
    Dim sCurVer,sSqlCreateTable,sTables,sTargetPath
    Dim iPosMaster,iPosPatch,iFoo,iFile,iMst,iCnt,iAsc,iAscCnt,iCmp
    Dim iPosArr,iArrMaxCnt,iColCnt,iBaseRefCnt,iIndex
    Dim bAsc,bMstApplied,bBaseRefFound,bFtkViewComplete,bFtkInScope,bFtkForceOutOfScope,bNeedKeyPathFallback
    Dim bCreate,bDelete,bDrop,bInsert,bFileNameChanged
    Dim MsiDb,SessionDb,MspDb,qViewFT,qViewMst,qViewCompID,qViewFC,qViewDir,qViewAssembly,Record,Record2,Record3
    Dim qViewMspCompId,qViewMspFC,qViewMsiAssembly,AllOfficeFiles,FileStream
    Dim SessionDir,Table,ViewTables,tbl
    Dim dicTransforms,dicKeys
    Dim arrTables
    
    If fBasicMode Then Exit Sub
    On Error Resume Next
    Const FILES_FTK             = 0 'key field for each file
    Const FILES_SOURCE          = 1
    Const FILES_ISPATCHED       = 2
    Const FILES_FILESTATUS      = 3 
    Const FILES_FILE            = 4
    Const FILES_FOLDER          = 5
    Const FILES_FULLNAME        = 6
    Const FILES_LANGUAGE        = 7
    Const FILES_BASEVERSION     = 8
    Const FILES_PATCHVERSION    = 9
    Const FILES_CURRENTVERSION  =10
    Const FILES_VERSIONSTATUS   =11
    Const FILES_PATCHCODE       =12 'key field for patch
    Const FILES_PATCHSTATE      =13
    Const FILES_PATCHKB         =14
    Const FILES_PATCHPACKAGE    =15
    Const FILES_PATCHMOREINFO   =16
    Const FILES_DIRECTORY       =17
    Const FILES_COMPONENTID     =18
    Const FILES_COMPONENTNAME   =19
    Const FILES_COMPONENTSTATE  =20
    Const FILES_KEYPATH         =21
    Const FILES_COMPONENTCLIENTS=22
    Const FILES_FEATURENAMES    =23
    Const FILES_COLUMNCNT       =23
    
    Const SQL_FILETABLE = "SELECT * FROM `_TransformView` WHERE `Table` = 'File' ORDER BY `Row`"
    Const SQL_PATCHTRANSFORMS = "SELECT `Name` FROM `_Storages` ORDER BY `Name`"
    Const INSTALLSTATE_ASSEMBLY = 6
    Const WAITTIME = 500 

' Loop all products
    If fCScript AND NOT fQuiet Then wscript.echo vbTab & "File version scan"
    For iPosMaster = 0 To UBound(arrMaster)
        If (arrMaster(iPosMaster, COL_ISOFFICEPRODUCT) OR fListNonOfficeProducts) AND (arrMaster(iPosMaster, COL_CONTEXT) = MSIINSTALLCONTEXT_MACHINE) Then
        ' Cache ProductCode
            sProductCode = "" : sProductCode = arrMaster(iPosMaster,COL_PRODUCTCODE)
            If fCScript AND NOT fQuiet Then wscript.echo vbTab & sProductCode 
        ' Reset Files array
            iPosArr = -1
            iArrMaxCnt = 5000
            ReDim arrFiles(FILES_COLUMNCNT,iArrMaxCnt)
            iBaseRefCnt = 0
            For iFoo = 1 To 1
            ' Add fields from msi base
            ' ------------------------
            ' Connect to the local .msi file for reading
                Err.Clear
                Set MsiDb = oMsi.OpenDatabase(arrMaster(iPosMaster,COL_CACHEDMSI),MSIOPENDATABASEMODE_READONLY)
                If Not Err = 0 Then 
                    Exit For
                End If 'Err = 0
            ' Check which tables exist in the current .msi
                sTables = ""
                Set ViewTables = MsiDb.OpenView("SELECT `Name` FROM `_Tables` ORDER BY `Name`")
                ViewTables.Execute
                Do
                    Set Table = ViewTables.Fetch
                    If Table Is Nothing then Exit Do
                    sTables = sTables & Table.StringData(1) & ","
                    If Not Err = 0 Then Exit Do
                Loop
                ViewTables.Close
                arrTables = Split(RTrimComma(sTables),",")
            ' Build an assembly reference dictionary
                Set dicAssembly = Nothing
                Set dicAssembly = CreateObject("Scripting.Dictionary")
                If InStr(sTables,"MsiAssembly,")>0 Then
                    sQueryAssembly = "SELECT DISTINCT `Component_` FROM MsiAssembly"
                    Set qViewAssembly = MsiDb.OpenView(sQueryAssembly)
                    qViewAssembly.Execute
                ' If the MsiAssmbly table does not exist it returns an error 
                    If Not Err = 0 Then Err.Clear
                    Set Record = qViewAssembly.Fetch
                ' must not enter the loop in case of an error!
                    If Not Err = 0 Then
                        Err.Clear
                    Else 
                        Do Until Record Is Nothing
                            If Not dicAssembly.Exists(Record.StringData(1)) Then
                                dicAssembly.Add Record.StringData(1),Record.StringData(1)
                            End If
                            Set Record = qViewAssembly.Fetch
                        Loop
                    End If 'Not Err = 0
                    qViewAssembly.Close
                End If 'InStr(sTables,"MsiAssembly")>0
                
                If InStr(sTables,"SxsMsmGenComponents,")>0 Then
                    sQueryAssembly = "SELECT DISTINCT `Component_` FROM SxsMsmGenComponents"
                    Set qViewAssembly = MsiDb.OpenView(sQueryAssembly)
                    qViewAssembly.Execute
                ' If the MsiAssmbly table does not exist it returns an error 
                    If Not Err = 0 Then Err.Clear
                    Set Record = qViewAssembly.Fetch
                ' must not enter the loop in case of an error!
                    If Not Err = 0 Then
                        Err.Clear
                    Else 
                        Do Until Record Is Nothing
                            If Not dicAssembly.Exists(Record.StringData(1)) Then
                                dicAssembly.Add Record.StringData(1),Record.StringData(1)
                            End If
                            Set Record = qViewAssembly.Fetch
                        Loop
                    End If 'Not Err = 0
                    qViewAssembly.Close
                End If 'InStr(sTables,"MsiAssembly")>0
                
            ' Build directory reference
                Set SessionDir = Nothing
                oMsi.UILevel = 2 'None
                Set SessionDir = oMsi.OpenProduct(sProductCode)
                SessionDir.DoAction("CostInitialize")
                SessionDir.DoAction("FileCost")
                SessionDir.DoAction("CostFinalize")
                Set dicFolders = Nothing
                Set dicFolders = CreateObject("Scripting.Dictionary")
                Err.Clear
                Set SessionDb = SessionDir.Database
                sQueryDir = "SELECT DISTINCT `Directory` FROM Directory"
                Set qViewDir = SessionDb.OpenView(sQueryDir)
                qViewDir.Execute
                Set Record = qViewDir.Fetch
            ' must not enter the loop in case of an error!
                If Not Err = 0 Then
                    Err.Clear
                Else 
                    Do Until Record Is Nothing
                        If Not dicFolders.Exists(Record.Stringdata(1)) Then
                            sTargetPath = "" : sTargetPath = SessionDir.TargetPath(Record.Stringdata(1))
                            If NOT sTargetPath = "" Then dicFolders.Add Record.Stringdata(1), sTargetPath
                        End If
                        Set Record = qViewDir.Fetch
                    Loop
                End If 'Not Err = 0
                qViewDir.Close
            
            ' .msi file inventory
            ' -------------------
                sQueryFT = "SELECT * FROM File"
                Set qViewFT = MsiDb.OpenView(sQueryFT)
                qViewFT.Execute
                Set Record = qViewFT.Fetch()
                Do Until Record Is Nothing
                ' Next Row in Array
                ' -----------------
                    iPosArr = iPosArr + 1
                    iBaseRefCnt = iPosArr
                    If iPosArr > iArrMaxCnt Then
                    ' increase array row buffer
                        iArrMaxCnt = iArrMaxCnt + 1000
                        ReDim Preserve arrFiles(FILES_COLUMNCNT,iArrMaxCnt)
                    End If 'iPosArr > iArrMaxCnt
                ' add FTK name
                    arrFiles(FILES_FTK,iPosArr) = Record.StringData(1)
                ' the FilesSource flag allows to filter the data in the report to exclude patch only entries.
                    arrFiles(FILES_SOURCE,iPosArr) = "Msi"
                ' default IsPatched field to 'False'
                    arrFiles(FILES_ISPATCHED,iPosArr) = False
                ' add the LFN (long filename)
                    arrFiles(FILES_FILE,iPosArr) = GetLongFileName(Record.StringData(3))
                ' add ComponentName
                    arrFiles(FILES_COMPONENTNAME,iPosArr) = Record.StringData(2)
                ' add ComponentID and Directory reference from Component table
                    sQueryCompID = "SELECT `Component`,`ComponentId`,`Directory_` FROM Component WHERE `Component` = '" & Record.StringData(2) &"'"
                    Set qViewCompID = MsiDb.OpenView(sQueryCompID)
                    qViewCompID.Execute
                    Set Record2 = qViewCompID.Fetch()
                    arrFiles(FILES_COMPONENTID,iPosArr) = Record2.StringData(2)
                    arrFiles(FILES_DIRECTORY,iPosArr) = Record2.StringData(3)
                    Set Record2 = Nothing
                    qViewCompID.Close
                    Set qViewCompID = Nothing
                ' ComponentState
                    arrFiles(FILES_COMPONENTSTATE,iPosArr) = GetComponentState(sProductCode,arrFiles(FILES_COMPONENTID,iPosArr),iPosMaster)
                ' add ComponentClients
                    arrFiles(FILES_COMPONENTCLIENTS,iPosArr) = GetComponentClients(arrFiles(FILES_COMPONENTID,iPosArr),arrFiles(FILES_COMPONENTSTATE,iPosArr))
                ' add Features that use the component
                    sQueryFC = "SELECT * FROM FeatureComponents WHERE `Component_` = '" & Record.StringData(2) &"'"
                    Set qViewFC = MsiDb.OpenView(sQueryFC)
                    qViewFC.Execute
                    Set Record2 = qViewFC.Fetch()
                    Do Until Record2 Is Nothing
                        arrFiles(FILES_FEATURENAMES,iPosArr) = arrFiles(FILES_FEATURENAMES,iPosArr) & Record2.StringData(1) & _
                                                               "("&TranslateFeatureState(oMsi.FeatureState(sProductCode,Record2.StringData(1)))&")"& ","
                        Set Record2 = qViewFC.Fetch()
                    Loop
                    RTrimComma arrFiles(FILES_FEATURENAMES,iPosArr)
                ' add KeyPath
                    arrFiles(FILES_KEYPATH, iPosArr) = GetComponentPath(sProductCode, arrFiles(FILES_COMPONENTID, iPosArr), arrFiles(FILES_COMPONENTSTATE, iPosArr))
                    sPath = "" : sName = ""
                ' add Componentpath
                    If dicAssembly.Exists(arrFiles(FILES_COMPONENTNAME,iPosArr)) Then
                    ' Assembly
                        If arrFiles(FILES_COMPONENTSTATE,iPosArr) = INSTALLSTATE_LOCAL Then
                            sPath = GetAssemblyPath(arrFiles(FILES_FILE,iPosArr),arrFiles(FILES_KEYPATH,iPosArr),dicFolders.Item(arrFiles(FILES_DIRECTORY,iPosArr)))
                            arrFiles(FILES_FOLDER,iPosArr) = Left(sPath,InStrRev(sPath,"\"))
                        Else
                            arrFiles(FILES_FOLDER,iPosArr) = sPath
                        End If
                    Else
                    ' Regular component
                        arrFiles(FILES_FOLDER, iPosArr) = dicFolders.Item(arrFiles(FILES_DIRECTORY, iPosArr))
                        If arrFiles(FILES_FOLDER, iPosArr) = "" Then
                            ' failed to obtain the directory from the session object
                            ' try again by direct read from session object
                            arrFiles(FILES_FOLDER, iPosArr) = SessionDir.TargetPath(arrFiles(FILES_DIRECTORY, iPosArr))

                            ' if still failed, fall back to the keypath by using the assembly logic to resolve the path
                            If arrFiles(FILES_FOLDER, iPosArr) = "" AND arrFiles(FILES_COMPONENTSTATE, iPosArr) = INSTALLSTATE_LOCAL Then
                                sPath = GetAssemblyPath(arrFiles(FILES_FILE, iPosArr), arrFiles(FILES_KEYPATH,iPosArr), dicFolders.Item(arrFiles(FILES_DIRECTORY,iPosArr)))
                                arrFiles(FILES_FOLDER,iPosArr) = Left(sPath,InStrRev(sPath,"\"))
                            End If
                        End If
                    End If
                ' add file FullName - if sPath contains a string then it's the result of the assembly detection
                    If sPath = "" Then
                        arrFiles(FILES_FULLNAME,iPosArr) = GetFileFullName(arrFiles(FILES_COMPONENTSTATE,iPosArr),arrFiles(FILES_FOLDER,iPosArr),arrFiles(FILES_FILE,iPosArr))
                    Else
                        arrFiles(FILES_FULLNAME,iPosArr) = sPath
                        sName = Right(sPath,Len(sPath)-InStrRev(sPath,"\"))
                        'Update the files field 
                        If Not UCase(sName) = UCase(arrFiles(FILES_FILE,iPosArr)) Then _
                            arrFiles(FILES_FILE,iPosArr) = sName '& " ("&arrFiles(FILES_FILE,iPosArr)&")"
                    End If
                ' add (msi) BaseVersion
                    If Not Err = 0 Then Err.Clear
                    arrFiles(FILES_BASEVERSION,iPosArr) = Record.StringData(5)
                ' add FileState and FileVersion
                    arrFiles(FILES_FILESTATUS,iPosArr) = GetFileState(arrFiles(FILES_COMPONENTSTATE,iPosArr),arrFiles(FILES_FULLNAME,iPosArr))
                    arrFiles(FILES_CURRENTVERSION,iPosArr) = GetFileVersion(arrFiles(FILES_COMPONENTSTATE,iPosArr),arrFiles(FILES_FULLNAME,iPosArr))
                ' add Language
                    arrFiles(FILES_LANGUAGE,iPosArr) = Record.StringData(6)
                ' get next row
                    Set Record = qViewFT.Fetch()
                Loop
                Set Record = Nothing
                qViewFT.Close
                Set qViewFT = Nothing

            ' --------------
            '  Add patches '
            ' --------------
            ' loop through all patches for the current product
                For iPosPatch = 0 to UBound(arrPatch,3)
                    If Not (IsEmpty (arrPatch(iPosMaster,PATCH_PATCHCODE,iPosPatch))) AND (arrPatch(iPosMaster,PATCH_CSP,iPosPatch) = True) Then
                        Err.Clear
                        sPatchCode = arrPatch(iPosMaster,PATCH_PATCHCODE,iPosPatch)
                        Set MspDb = oMsi.OpenDatabase(arrPatch(iPosMaster,PATCH_LOCALPACKAGE,iPosPatch),MSIOPENDATABASEMODE_PATCHFILE)
                        If Err = 0 Then
                        ' create the table structures from .msi schema to allow detailed query of the .msp _TransformView 
                            For Each tbl in arrTables
                                sSqlCreateTable = "CREATE TABLE `" & tbl &"` (" & GetTableColumnDef(MsiDb,tbl) & " PRIMARY KEY " & GetPrimaryTableKeys(MsiDb,tbl) &")"
                                MspDb.OpenView(sSqlCreateTable).Execute
                            Next 'tbl
                        ' check if a PatchTransform is available
                            sMst = "" : bMstApplied = False
                            sMst = arrPatch(iPosMaster,PATCH_TRANSFORM,iPosPatch)
                            Err.Clear
                            If InStr(sMst,";")>0 Then  
                                bMstApplied = True 
                                sMst = Left(sMst,InStr(sMst,";")-1) 
                            ' apply the patch transform  
                            ' msiTransformErrorAll includes msiTransformErrorViewTransform which creates the "_TransformView" table 
                                MspDb.ApplyTransform sMst,MSITRANSFORMERROR_ALL 
                            End If 'InStr(sMst,";")>0 
                        ' if no known .mst or failed to apply the .mst we go into generic patch embedded transform detection loop
                            If (Not bMstApplied) OR (Not Err = 0) Then
                                Err.Clear
                            ' Dictionary object for the patch transforms
                                Set dicTransforms = CreateObject("Scripting.Dictionary")
                            ' create the view to retrieve the patch transforms
                                sQueryMst = SQL_PATCHTRANSFORMS
                                Set qViewMst = MspDb.OpenView(sQueryMst): qViewMst.Execute
                                Set Record = qViewMst.Fetch
                            ' loop all transforms and add them to the dictionary
                                Do Until Record Is Nothing
                                    sMst = Record.StringData(1)
                                        dicTransforms.Add sMst,sMst
                                    Set Record = qViewMst.Fetch
                                Loop
                                qViewMst.Close : Set qViewMst = Nothing
                                Set Record = Nothing
                            ' apply the patch transforms
                                dicKeys = dicTransforms.Keys
                                For iMst = 0 To dicTransforms.Count - 1
                                ' get the transform name
                                    sMst = dicKeys(iMst)
                                ' apply the patch transform / staple them all on the table
                                    MspDb.ApplyTransform ":" & sMst,MSITRANSFORMERROR_ALL
                                    If Not Err = 0 Then Err.Clear
                                Next 'iMst
                            End If '(Not bMstApplied) OR (Not Err = 0)
                            
                        ' _TransformView loop
                        ' -------------------
                        ' update the MsiAssembly reference dictionary
                            Err.Clear
                            If InStr(sTables,"MsiAssembly,")>0 Then
                                Set qViewMsiAssembly = MspDb.OpenView("SELECT * FROM `_TransformView` WHERE `Table`='MsiAssembly' ORDER By `Row`")
                                qViewMsiAssembly.Execute()
                                Set Record = qViewMsiAssembly.Fetch()
                                Do
                                    If Record Is Nothing Then Exit Do
                                    If Not Err = 0 Then
                                        Err.Clear
                                        Exit Do
                                    End If
                                    If Record.StringData(2) = "INSERT" Then
                                        If Not dicAssembly.Exists(Record.StringData(3)) Then dicAssembly.Add Record.StringData(3),Record.StringData(3)
                                    End If
                                    Set Record = qViewMsiAssembly.Fetch()
                                Loop
                                qViewMsiAssembly.Close
                            End If
                            Set Record = Nothing

                            If InStr(sTables,"SxsMsmGenComponents,")>0 Then
                                Set qViewMsiAssembly = MspDb.OpenView("SELECT * FROM `_TransformView` WHERE `Table`='SxsMsmGenComponents' ORDER By `Row`")
                                qViewMsiAssembly.Execute()
                                Set Record = qViewMsiAssembly.Fetch()
                                Do
                                    If Record Is Nothing Then Exit Do
                                    If Not Err = 0 Then
                                        Err.Clear
                                        Exit Do
                                    End If
                                    If Record.StringData(2) = "INSERT" Then
                                        If Not dicAssembly.Exists(Record.StringData(3)) Then dicAssembly.Add Record.StringData(3),Record.StringData(3)
                                    End If
                                    Set Record = qViewMsiAssembly.Fetch()
                                Loop
                                qViewMsiAssembly.Close
                            End If
                            Set Record = Nothing
                        ' get the files being modified from the "_TransformView" 'File' table
                            Set qViewMst = MspDb.OpenView(SQL_FILETABLE) : qViewMst.Execute()
                        ' loop all of the entries in the File table from "_TransformView"
                            Set Record = qViewMst.Fetch()
                        ' initial defaults
                            sFtk = ""
                            bFtkViewComplete = True
                            bFtkInScope = False
                            Do
                            ' is this the next FTK?
                                If (Not sFtk = Record.StringData(3)) OR (Record Is Nothing) Then
                                    If Record Is Nothing Then Err.Clear
                                ' yes this is the next FTK or the last time before exit of the loop
                                ' is previous FTK handling complete?
                                    If Not bFtkViewComplete Then
                                    ' previous FTK handling is not complete
                                    ' is previous FTK in scope?
                                        If bFtkInScope AND NOT bFtkForceOutOfScope Then
                                        'FTK is in scope - reset the scope flag
                                            bFtkInScope = False
                                            If bBaseRefFound Then
                                            ' update base entry fields with patch information
                                            ' check if the filename got updated
                                                If bFileNameChanged Then
                                                    arrFiles(FILES_FILE,iPosArr) = GetLongFileName(arrFiles(FILES_FILE,iPosArr))
                                                ' the filename got changed by a patch
                                                ' if the patch is in the 'Applied' state -> care about this change
                                                    If LCase(arrFiles(FILES_PATCHSTATE,iPosArr)) = "applied" Then
                                                    ' update the filename in the baseref if this is (NOT Assembly) OR (broken)
                                                        If NOT(dicAssembly.Exists(arrFiles(FILES_COMPONENTNAME,iPosArr))) OR (arrFiles(FILES_FILESTATUS,iPosArr) = INSTALLSTATE_BROKEN) Then
                                                        ' correct the baseref filename field
                                                            arrFiles(FILES_FILE,iCnt) = arrFiles(FILES_FILE,iPosArr)
                                                        ' File FullName
                                                            arrFiles(FILES_FULLNAME,iPosArr) = GetFileFullName(arrFiles(FILES_COMPONENTSTATE,iPosArr),arrFiles(FILES_FOLDER,iPosArr),arrFiles(FILES_FILE,iPosArr))
                                                        ' recheck the filestate
                                                            arrFiles(FILES_FILESTATUS,iPosArr) = GetFileState(arrFiles(FILES_COMPONENTSTATE,iPosArr),arrFiles(FILES_FULLNAME,iPosArr))
                                                            arrFiles(FILES_FILESTATUS,iCnt) = arrFiles(FILES_FILESTATUS,iPosArr)
                                                        ' recheck the version
                                                            arrFiles(FILES_CURRENTVERSION,iPosArr) = GetFileVersion(arrFiles(FILES_COMPONENTSTATE,iPosArr),arrFiles(FILES_FULLNAME,iPosArr))
                                                            arrFiles(FILES_CURRENTVERSION,iCnt) = arrFiles(FILES_CURRENTVERSION,iPosArr)
                                                        End If
                                                    End If 'applied
                                                End If 'bFileNameChanged
                                            ' set IsPatched flag
                                                arrFiles(FILES_ISPATCHED,iCnt) = True
                                            ' check if PatchVersion field needs to be updated
                                                iCmp = 2 : sCmp = ""
                                                iCmp = CompareVersion(arrFiles(FILES_PATCHVERSION,iPosArr),arrFiles(FILES_PATCHVERSION,iCnt),True)
                                                Select Case iCmp
                                                Case VERSIONCOMPARE_LOWER   ': sCmp="ERROR_VersionLow"
                                                Case VERSIONCOMPARE_MATCH   ': sCmp="SUCCESS_VersionMatch"
                                                Case VERSIONCOMPARE_HIGHER  ': sCmp="SUCCESS_VersionHigh"
                                                ' update the base field
                                                    arrFiles(FILES_PATCHVERSION,iCnt) = arrFiles(FILES_PATCHVERSION,iPosArr)
                                                Case VERSIONCOMPARE_INVALID ': sCmp=""
                                                End Select
                                            ' update PatchCode field
                                                arrFiles(FILES_PATCHCODE,iCnt) = arrFiles(FILES_PATCHCODE,iCnt) & sPatchCode&","
                                            ' update PatchKB field 
                                                If dicMspIndex.Exists(sPatchCode) Then arrFiles(FILES_PATCHKB,iCnt) = arrFiles(FILES_PATCHKB,iCnt) & arrFiles(FILES_PATCHKB,iPosArr) &","
                                            ' update PatchMoreInfo field
                                                arrFiles(FILES_PATCHMOREINFO,iCnt) = arrFiles(FILES_PATCHMOREINFO,iCnt) & arrPatch(iPosMaster,PATCH_MOREINFOURL,iPosPatch)&","
                                            Else    
                                            ' bBaseRefFound is False. 
                                            ' -----------------------
                                            ' this is a new file introduced with the patch. set FileSource flag
                                                arrFiles(FILES_SOURCE,iPosArr) = "Msp"
                                            ' define IsPatched field
                                                arrFiles(FILES_ISPATCHED,iPosArr) = False
                                            ' add the LFN (long file name)
                                                arrFiles(FILES_FILE,iPosArr) = GetLongFileName(arrFiles(FILES_FILE,iPosArr))
                                            ' locate the ComponentId from Component table
                                                sQueryCompID = "SELECT `Component`,`ComponentId`,´Directory_` FROM Component WHERE `Component` = '" & arrFiles(FILES_COMPONENTNAME,iPosArr) &"'"
                                                Set qViewCompID = MsiDb.OpenView(sQueryCompID)
                                                If Not Err = 0 Then 
                                                    Err.Clear
                                                    Set Record2 = Nothing
                                                Else
                                                    qViewCompID.Execute
                                                    Set Record2 = qViewCompID.Fetch()
                                                End If
                                                If Not Record2 Is Nothing Then 
                                                ' found the ComponentId
                                                ' this is a new file added to an existing component
                                                    arrFiles(FILES_COMPONENTID,iPosArr) = Record2.StringData(2)
                                                ' add the Directory_ reference
                                                    arrFiles(FILES_DIRECTORY,iPosArr) = Record2.StringData(3)
                                                    Set Record2 = Nothing
                                                    qViewCompID.Close
                                                    Set qViewCompID = Nothing
                                                Else
                                                ' did not find the ComponentId in the base .msi
                                                ' this is a new file AND a new component -> need to query the .msp for details
                                                    Set qViewMspCompId = MspDb.OpenView("SELECT * FROM `_TransformView` WHERE `Table`='Component' ORDER BY `Row`")
                                                    qViewMspCompID.Execute
                                                    Do
                                                        Set Record3 = qViewMspCompId.Fetch()
                                                        If Record3 Is Nothing Then Exit Do
                                                        If Record3.StringData(3) = arrFiles(FILES_COMPONENTNAME,iPosArr) Then 
                                                            If Record3.StringData(2) = "ComponentId" Then 
                                                                arrFiles(FILES_COMPONENTID,iPosArr) = Record3.StringData(4)
                                                            ElseIf Record3.StringData(2) = "Directory_" Then
                                                                arrFiles(FILES_DIRECTORY,iPosArr) = Record3.StringData(4)
                                                            End If
                                                        End If
                                                    Loop
                                                    qViewMspCompID.Close
                                                    If arrFiles(FILES_COMPONENTID,iPosArr)="" Then bFtkForceOutOfScope = True
                                                End If 'Not Record2 Is Nothing
                                            ' all other logic is only needed if in scope
                                                If Not bFtkForceOutOfScope Then
                                                ' ensure the directory reference exists
                                                    If Not dicFolders.Exists(arrFiles(FILES_DIRECTORY,iPosArr)) Then
                                                        If SessionDir Is Nothing Then
                                                        ' try to recover lost SessionDir object
                                                            Set SessionDir = oMsi.OpenProduct(sProductCode)
                                                            SessionDir.DoAction("CostInitialize")
                                                            SessionDir.DoAction("FileCost")
                                                            SessionDir.DoAction("CostFinalize")
                                                        End If
                                                        dicFolders.Add arrFiles(FILES_DIRECTORY,iPosArr),SessionDir.TargetPath(arrFiles(FILES_DIRECTORY,iPosArr))
                                                        If Not Err = 0 Then 
                                                            Err.Clear
                                                        ' still failed to identify the path - get rid of this entry
                                                            bNeedKeyPathFallback = True 
                                                        End If
                                                    End If
                                                ' ComponentState
                                                    arrFiles(FILES_COMPONENTSTATE,iPosArr) = GetComponentState(sProductCode,arrFiles(FILES_COMPONENTID,iPosArr),iPosMaster)
                                                ' add ComponentClients
                                                    arrFiles(FILES_COMPONENTCLIENTS,iPosArr) = GetComponentClients(arrFiles(FILES_COMPONENTID,iPosArr),arrFiles(FILES_COMPONENTSTATE,iPosArr))
                                                ' add Features that use the component
                                                    Set qViewMspFC = MspDb.OpenView("SELECT * FROM `_TransformView` WHERE `Table`='FeatureComponents' ORDER BY `Row`")
                                                    qViewMspFC.Execute
                                                    Set Record2 = qViewMspFC.Fetch()
                                                    Do
                                                        Set Record2 = qViewMspFC.Fetch()
                                                        If Record2 Is Nothing Then Exit Do
                                                        If Record2.StringData(4) = arrFiles(FILES_COMPONENTNAME,iPosArr) Then 
                                                            If Record2.StringData(2) = "Feature_" Then 
                                                                arrFiles(FILES_FEATURENAMES,iPosArr) = arrFiles(FILES_FEATURENAMES,iPosArr)&Record2.StringData(3)& _
                                                                "("&TranslateFeatureState(oMsi.FeatureState(sProductCode,Record2.StringData(3)))&")"& ","
                                                            End If
                                                        End If
                                                    Loop
                                                    qViewMspFC.Close
                                                    Set Record2 = Nothing
                                                    RTrimComma arrFiles(FILES_FEATURENAMES,iPosArr)
                                                ' add KeyPath
                                                    arrFiles(FILES_KEYPATH,iPosArr) = GetComponentPath(sProductCode,arrFiles(FILES_COMPONENTID,iPosArr),arrFiles(FILES_COMPONENTSTATE,iPosArr))
                                                ' add Componentpath
                                                    sPath = "" : sName = ""
                                                    If dicAssembly.Exists(arrFiles(FILES_COMPONENTNAME,iPosArr)) Then
                                                    ' Assembly
                                                        If arrFiles(FILES_COMPONENTSTATE,iPosArr) = INSTALLSTATE_LOCAL Then
                                                            sPath = GetAssemblyPath(arrFiles(FILES_FILE,iPosArr),arrFiles(FILES_KEYPATH,iPosArr),dicFolders.Item(arrFiles(FILES_DIRECTORY,iPosArr)))
                                                            arrFiles(FILES_FOLDER,iPosArr) = Left(sPath,InStrRev(sPath,"\"))
                                                            sName = Right(sPath,Len(sPath)-InStrRev(sPath,"\"))
                                                        ' update the files field to ensure the correct value
                                                            If Not UCase(sName) = UCase(arrFiles(FILES_FILE,iPosArr)) Then _
                                                                arrFiles(FILES_FILE,iPosArr) = sName '& " ("&arrFiles(FILES_FILE,iPosArr)&")"
                                                        Else
                                                            arrFiles(FILES_FOLDER,iPosArr) = sPath
                                                        End If
                                                    Else
                                                    ' Regular component
                                                        If bNeedKeyPathFallback Then
                                                            arrFiles(FILES_FOLDER,iPosArr) = Left(arrFiles(FILES_KEYPATH,iPosArr),InStrRev(arrFiles(FILES_KEYPATH,iPosArr),"\"))
                                                        Else
                                                            arrFiles(FILES_FOLDER,iPosArr) = dicFolders.Item(arrFiles(FILES_DIRECTORY,iPosArr))
                                                        End If
                                                    End If
                                                ' add file FullName
                                                ' if sPath contains a string then it's the result of the assembly detection
                                                    If sPath = "" Then
                                                        arrFiles(FILES_FULLNAME,iPosArr) = GetFileFullName(arrFiles(FILES_COMPONENTSTATE,iPosArr),arrFiles(FILES_FOLDER,iPosArr),arrFiles(FILES_FILE,iPosArr))
                                                    Else
                                                        arrFiles(FILES_FULLNAME,iPosArr) = arrFiles(FILES_FOLDER,iPosArr)&arrFiles(FILES_FILE,iPosArr)
                                                    End If
                                                ' add FileState and FileVersion
                                                    arrFiles(FILES_FILESTATUS,iPosArr) = GetFileState(arrFiles(FILES_COMPONENTSTATE,iPosArr),arrFiles(FILES_FULLNAME,iPosArr))
                                                    arrFiles(FILES_CURRENTVERSION,iPosArr) = GetFileVersion(arrFiles(FILES_COMPONENTSTATE,iPosArr),arrFiles(FILES_FULLNAME,iPosArr))
                                                End If
                                            End If 'bBaseRefFound
                                        Else
                                        ' No - FTK not in scope
                                            bFtkForceOutOfScope = False
                                        ' delete all row contents
                                            For iColCnt = 0 To FILES_COLUMNCNT
                                                arrFiles(iColCnt,iPosArr) = ""
                                            Next 'iColCnt
                                        ' decrease array counter
                                            iPosArr = iPosArr - 1
                                        End If 'bFtkInScope
                                        
                                        If bFtkForceOutOfScope Then
                                        ' delete all row contents
                                            For iColCnt = 0 To FILES_COLUMNCNT
                                                arrFiles(iColCnt,iPosArr) = ""
                                            Next 'iColCnt
                                        ' decrease array counter
                                            iPosArr = iPosArr - 1
                                        End If
                                    End If 'bFtkViewComplete
                                    bFtkViewComplete = True
                                    
                                ' Previous FTK handling is now complete
                                ' -------------------------------------
                                    If Record Is Nothing Then Exit Do
                                ' Init new FTK row
                                ' ----------------
                                ' increase array pointer
                                    iPosArr = iPosArr + 1
                                    bFtkViewComplete = False
                                    bFtkInScope = False
                                    bFtkForceOutOfScope = False
                                    bInsert = False
                                    bFileNameChanged = False
                                    If iPosArr > iArrMaxCnt Then
                                    ' add more rows to array
                                        iArrMaxCnt = iArrMaxCnt + 1000
                                        ReDim Preserve arrFiles(FILES_COLUMNCNT,iArrMaxCnt)
                                    End If 'iPosArr > iArrMaxCnt
                                ' update current FTK cache reference
                                    sFtk = Record.StringData(3)
                                ' locate the FTK reference from msi base
                                    bBaseRefFound = False
                                    For iCnt = 0 To iBaseRefCnt
                                        If arrFiles(FILES_FTK,iCnt) = sFtk Then
                                            bBaseRefFound = True
                                        ' copy known fields from Base version if applicable
                                            If bBaseRefFound Then
                                                For iColCnt = 0 To FILES_COLUMNCNT
                                                    arrFiles(iColCnt,iPosArr) = arrFiles(iColCnt,iCnt)
                                                Next 'iColCnt
                                                bFtkInScope = True
                                            End If 'bBaseRefFound
                                            Exit For 'iCnt = 0 To UBound(arrFiles,2)-1
                                        End If 'arrFiles(FILES_FTK,iCnt) = sFtk Then
                                    Next 'iCnt
                                ' add initial available data
                                ' correct/ensure FTK name
                                    arrFiles(FILES_FTK,iPosArr) = Record.StringData(3)
                                ' correct IsPatched field
                                    arrFiles(FILES_ISPATCHED,iPosArr) = False
                                ' correct/ensure FileSource
                                    arrFiles(FILES_SOURCE,iPosArr) = "Msp"
                                ' add fields from patch array
                                    arrFiles(FILES_PATCHSTATE,iPosArr) = arrPatch(iPosMaster,PATCH_PATCHSTATE,iPosPatch)
                                    arrFiles(FILES_PATCHCODE,iPosArr) = arrPatch(iPosMaster,PATCH_PATCHCODE,iPosPatch)
                                    arrFiles(FILES_PATCHMOREINFO,iPosArr) = arrPatch(iPosMaster,PATCH_MOREINFOURL,iPosPatch)
                                ' add KB reference
                                    If dicMspIndex.Exists(sPatchCode) Then
                                        iIndex = dicMspIndex.Item(sPatchCode)
                                        arrFiles(FILES_PATCHKB,iPosArr) = arrMspFiles(iIndex,MSPFILES_KB)
                                        arrFiles(FILES_PATCHPACKAGE,iPosArr) = arrMspFiles(iIndex,MSPFILES_PACKAGE)
                                    End If
                                ' new FTK row init complete
                                End If 'Not sFtk = Record.StringData(3)
                            ' add data from _TransformView
                                Select Case Record.StringData(2)
                                Case "File"
                                Case "FileSize"
                                Case "Component_"
                                    arrFiles(FILES_COMPONENTNAME,iPosArr) = Record.StringData(4)
                                Case "CREATE"
                                Case "DELETE"
                                Case "DROP"
                                Case "FileName"
                                'Add the filename
                                    bFileNameChanged = True
                                    arrFiles(FILES_FILE,iPosArr) = Record.StringData(4)
                                Case "Version"
                                ' don't allow version field to contain alpha characters
                                    bAsc = True : sAscCheck = "" : sAscCheck = Record.StringData(4)
                                    If Len(sAscCheck)>0 Then
                                        For iAscCnt = 1 To Len(sAscCheck)
                                            iAsc = Asc(UCase(Mid(sAscCheck,iAscCnt,1)))
                                            If (iAsc>64) AND (iAsc<91) Then
                                                bAsc = False
                                                Exit For
                                            End If
                                        Next 'iCnt
                                    End If 'Len(sAscCheck)>0
                                    If bAsc Then arrFiles(FILES_PATCHVERSION,iPosArr) = Record.StringData(4)
                                Case "Language"
                                        arrFiles(FILES_LANGUAGE,iPosArr) = Record.StringData(4)
                                Case "Attributes"
                                Case "Sequence"
                                Case "INSERT"
                                ' this is a new file added by the pach
                                    bFtkInScope = True
                                    bInsert = True
                                Case Else
                                End Select
                            ' get the next record (column) from _TransformView
                                Set Record = qViewMst.Fetch()
                            Loop
                        ' _TransformView analysis for this patch complete
                        ' reset views 
                            MspDb.OpenView("ALTER TABLE _TransformView FREE").Execute
                            MspDb.OpenView("DROP TABLE `File`").Execute
                        End If 'Err = 0
                    End If 'IsEmpty
                Next 'iPosPatch
            Next 'iFoo
            
        ' Final field & verb fixups
        ' -------------------------
            For iFile = 0 To iPosArr
            ' File VersionState translation
                iCmp = 2 : sCmp = "" : sCurVer = ""
                If arrFiles(FILES_ISPATCHED,iFile) OR (arrFiles(FILES_SOURCE,iFile)="Msp") Then 
                ' compare actual file version to patch file version 
                    sCurVer = arrFiles(FILES_PATCHVERSION,iFile)
                Else
                ' compare actual file version to base file version 
                    sCurVer = arrFiles(FILES_BASEVERSION,iFile)
                End If 'arrFiles(FILES_ISPATCHED,iFile)
                iCmp = CompareVersion(arrFiles(FILES_CURRENTVERSION,iFile),sCurVer,False)
                Select Case iCmp
                Case VERSIONCOMPARE_LOWER
                    sCmp="ERROR_VersionLow"
                ' log error
                    Cachelog LOGPOS_REVITEM,LOGHEADING_NONE,ERR_CATEGORYERROR,ERR_FILEVERSIONLOW & arrMaster(iPosMaster,COL_PRODUCTCODE) & DSV & _
                    arrMaster(iPosMaster,COL_PRODUCTNAME)
                    If Not InStr(arrMaster(iPosMaster,COL_ERROR),arrFiles(FILES_FILE,iFile)&" expected: ")>0 Then
                        arrMaster(iPosMaster,COL_ERROR) = arrMaster(iPosMaster,COL_ERROR) & ERR_CATEGORYERROR & arrFiles(FILES_FILE,iFile) & " expected: "&sCurVer& " found: " & arrFiles(FILES_CURRENTVERSION,iFile) & CSV
                    End If
                Case VERSIONCOMPARE_MATCH   : sCmp="SUCCESS_VersionMatch"
                Case VERSIONCOMPARE_HIGHER  : sCmp="SUCCESS_VersionHigh"
                Case VERSIONCOMPARE_INVALID : sCmp=""
                End Select
                arrFiles(FILES_VERSIONSTATUS,iFile) = sCmp
            ' FileState translation
                sCmp = ""
                Select Case arrFiles(FILES_FILESTATUS,iFile)
                Case INSTALLSTATE_LOCAL : sCmp = "OK_Local"
                Case INSTALLSTATE_BROKEN
                    sCmp = "ERROR_Broken"
                ' log error
                    Cachelog LOGPOS_REVITEM,LOGHEADING_NONE,ERR_CATEGORYERROR,"Product " & arrMaster(iPosMaster,COL_PRODUCTCODE) & DSV & _
                    arrMaster(iPosMaster,COL_PRODUCTNAME) & ": " & ERR_BADFILESTATE
                    If Not InStr(arrMaster(iPosMaster,COL_ERROR),arrFiles(FILES_FILE,iFile)&" FileState: Broken")>0 Then
                        arrMaster(iPosMaster,COL_ERROR) = arrMaster(iPosMaster,COL_ERROR) & ERR_CATEGORYERROR & arrFiles(FILES_FILE,iFile) & " FileState: Broken" & CSV
                    End If
                Case INSTALLSTATE_UNKNOWN
                    sCmp = "Unknown"
                    If Not arrFiles(FILES_FEATURENAMES,iFile) = "" Then sCmp=Mid(arrFiles(FILES_FEATURENAMES,iFile),InStrRev(arrFiles(FILES_FEATURENAMES,iFile),"(")+1, Len(arrFiles(FILES_FEATURENAMES,iFile))-InStrRev(arrFiles(FILES_FEATURENAMES,iFile),"(")-1)
                Case INSTALLSTATE_NOTUSED : sCmp = "NotUsed"
                Case INSTALLSTATE_ASSEMBLY : sCmp = "Assembly"
                Case Else
                End Select
                arrFiles(FILES_FILESTATUS,iFile) = sCmp
            ' ComponentState translation
                sCmp = ""
                Select Case arrFiles(FILES_COMPONENTSTATE,iFile)
                Case INSTALLSTATE_LOCAL : sCmp = "Local"
                Case INSTALLSTATE_BROKEN : sCmp = "Broken"
                Case INSTALLSTATE_UNKNOWN : sCmp = "Unknown"
                Case INSTALLSTATE_NOTUSED : sCmp = "NotUsed"
                Case Else
                End Select
                arrFiles(FILES_COMPONENTSTATE,iFile) = sCmp
            ' PatchCode field trim
                arrFiles(FILES_PATCHCODE,iFile) = RTrimComma(arrFiles(FILES_PATCHCODE,iFile))
            ' PatchKB field trim
                arrFiles(FILES_PATCHKB,iFile) = RTrimComma(arrFiles(FILES_PATCHKB,iFile))
            ' PatchInfo field trim
                arrFiles(FILES_PATCHMOREINFO,iFile) = RTrimComma(arrFiles(FILES_PATCHMOREINFO,iFile))
            Next 'iFile
        ' dump out the collected data to file
        ' create the AllOffice file
            If IsEmpty(AllOfficeFiles) Then
                If NOT oFso.FolderExists(sPathOutputFolder & "ROIScan") Then oFso.CreateFolder(sPathOutputFolder & "ROIScan")
                Set AllOfficeFiles = oFso.CreateTextFile(sPathOutputFolder&"ROIScan\"&sComputerName&"_OfficeAll_FileList.xml",True,True)
                AllOfficeFiles.WriteLine "<?xml version=""1.0""?>"
                AllOfficeFiles.WriteLine "<FILEDATA>"
            End If
        ' individual products file
            If NOT oFso.FolderExists(sPathOutputFolder & "ROIScan") Then oFso.CreateFolder(sPathOutputFolder & "ROIScan")
            Set FileStream = oFso.CreateTextFile(sPathOutputFolder&"ROIScan\"&sComputerName&"_"&sProductCode&"_FileList.xml",True,True)
            FileStream.WriteLine "<?xml version=""1.0""?>"
            FileStream.WriteLine "<FILEDATA>"
            FileStream.WriteLine vbTab & "<PRODUCT ProductCode="""&sProductCode&""" >"
            If arrMaster(iPosMaster,COL_ISOFFICEPRODUCT) Then AllOfficeFiles.WriteLine vbTab & "<PRODUCT ProductCode="""&sProductCode&""" >"
            For iFile = 0 To iPosArr
                sXmlLine = ""
                sXmlLine = vbTab & vbTab & "<FILE " & _
                                           "FileName="&chr(34)&arrFiles(FILES_FILE,iFile)&chr(34)&" " & _
                                           "FileState="&chr(34)&arrFiles(FILES_FILESTATUS,iFile)&chr(34)&" " & _
                                           "VersionStatus="&chr(34)&arrFiles(FILES_VERSIONSTATUS,iFile)&chr(34)&" " & _
                                           "CurrentVersion="&chr(34)&arrFiles(FILES_CURRENTVERSION,iFile)&chr(34)&" " & _
                                           "InitialVersion="&chr(34)&arrFiles(FILES_BASEVERSION,iFile)&chr(34)&" " & _
                                           "PatchVersion="&chr(34)&arrFiles(FILES_PATCHVERSION,iFile)&chr(34)&" " & _
                                           "FileSource="&chr(34)&arrFiles(FILES_SOURCE,iFile)&chr(34)&" " & _
                                           "IsPatched="&chr(34)&arrFiles(FILES_ISPATCHED,iFile)&chr(34)&" " & _
                                           "KB="&chr(34)&arrFiles(FILES_PATCHKB,iFile)&chr(34)&" " & _
                                           "Package="&chr(34)&arrFiles(FILES_PATCHPACKAGE,iFile)&chr(34)&" " & _
                                           "PatchState="&chr(34)&arrFiles(FILES_PATCHSTATE,iFile)&chr(34)&" " & _
                                           "FolderName="&chr(34)&arrFiles(FILES_FOLDER,iFile)&chr(34)&" " & _
                                           "PatchCode="&chr(34)&arrFiles(FILES_PATCHCODE,iFile)&chr(34)&" " & _
                                           "PatchInfo="&chr(34)&arrFiles(FILES_PATCHMOREINFO,iFile)&chr(34)&" " & _
                                           "FtkName="&chr(34)&arrFiles(FILES_FTK,iFile)&chr(34)&" " & _
                                           "KeyPath="&chr(34)&arrFiles(FILES_KEYPATH,iFile)&chr(34)&" " & _
                                           "MsiDirectory="&chr(34)&arrFiles(FILES_DIRECTORY,iFile)&chr(34)&" " & _
                                           "Language="&chr(34)&arrFiles(FILES_LANGUAGE,iFile)&chr(34)&" " & _
                                           "ComponentState="&chr(34)&arrFiles(FILES_COMPONENTSTATE,iFile)&chr(34)&" " & _
                                           "ComponentID="&chr(34)&arrFiles(FILES_COMPONENTID,iFile)&chr(34)&" " & _
                                           "ComponentName="&chr(34)&arrFiles(FILES_COMPONENTNAME,iFile)&chr(34)&" " & _
                                           "ComponentClients="&chr(34)&arrFiles(FILES_COMPONENTCLIENTS,iFile)&chr(34)&" " & _
                                           "FeatureReference="&chr(34)&arrFiles(FILES_FEATURENAMES,iFile)&chr(34)&" " & _
                                           " />"
                If InStr(sXmlLine,"&")>0 Then sXmlLine = Replace(sXmlLine,"&","&amp;")
                FileStream.WriteLine sXmlLine
                If arrMaster(iPosMaster,COL_ISOFFICEPRODUCT) Then AllOfficeFiles.WriteLine sXmlLine
            Next 'iFile
            FileStream.WriteLine vbTab & "</PRODUCT>"
            If arrMaster(iPosMaster,COL_ISOFFICEPRODUCT) Then AllOfficeFiles.WriteLine vbTab & "</PRODUCT>"
            FileStream.WriteLine "</FILEDATA>"
            FileStream.Close
            Set FileStream = Nothing
        End If 'arrMaster(iPosMaster,COL_ISOFFICEPRODUCT)
    Next 'iPosMaster
' close the AllOffice file
    If Not AllOfficeFiles Is Nothing Then
        AllOfficeFiles.WriteLine "</FILEDATA>"
        AllOfficeFiles.Close
        Set AllOfficeFiles = Nothing
    End if
' compress the files 
    Dim i,iWait
    Dim FileScanFolder,FileVerScanZip,xmlFile,item,zipfile
    Dim sDat,sDatCln
    Dim fCopyComplete
    If oFso.FileExists(sPathOutputFolder&sComputerName&"_ROIScan.zip") Then
    ' rename existing .zip container by appending a timestamp to prevent overwrite.
        Dim oRegExp
        Set oRegExp = CreateObject("Vbscript.RegExp")
        Set zipfile = oFso.GetFile(sPathOutputFolder&sComputerName&"_ROIScan.zip")
        oRegExp.Global = True
        oRegExp.Pattern = "\D"
        Err.Clear
        zipfile.Name = sComputerName&"_ROIScan_" & oRegExp.Replace(zipfile.DateLastModified, "") & ".zip"
        If NOT Err = 0 Then
            zipfile.Delete
            Err.Clear
        End If
    End If
    Set FileVerScanZip = oFso.OpenTextFile(sPathOutputFolder&sComputerName&"_ROIScan.zip",FOR_WRITING,True) 
    FileVerScanZip.write "PK" & chr(5) & chr(6) & String(18,chr(0)) 
    FileVerScanZip.close  
    Set FileScanFolder = oFso.GetFolder(sPathOutputFolder&"ROIScan")
    For Each xmlFile in FileScanFolder.Files 
        If Right(LCase(xmlFile.Name),4)=".xml" Then
            If NOT fZipError Then CopyToZip ShellApp.NameSpace(sPathOutputFolder&sComputerName&"_ROIScan.zip"), xmlFile
        End If 
    Next 'xmlFile 
    If fCScript AND NOT fQuiet Then wscript.echo vbTab & "File version scan complete"

End Sub 'FileInventory
'=======================================================================================================

'Identify the InstallState of a component
Function GetComponentState(sProductCode,sComponentId,iPosMaster)
    On Error Resume Next
    
    Dim Product
    Dim sPath

    GetComponentState = INSTALLSTATE_UNKNOWN
    If iWiVersionMajor > 2 Then
        'WI 3.x or higher
        Set Product = oMsi.Product(sProductCode,arrMaster(iPosMaster,COL_USERSID),arrMaster(iPosMaster,COL_CONTEXT))
        Err.Clear
        GetComponentState = Product.ComponentState(sComponentId)
        If Not Err = 0 Then 
            GetComponentState = INSTALLSTATE_UNKNOWN
            Err.Clear
        End If ' Err = 0
    Else
        'WI 2.x
        If Not Err = 0 Then Err.Clear
        sPath = ""
        sPath = oMsi.ComponentPath(sProductCode,sComponentId)
        If Not Err = 0 Then
            GetComponentState = INSTALLSTATE_UNKNOWN
            Err.Clear
        Else
            If oFso.FileExists(sPath) Then
                GetComponentState = INSTALLSTATE_LOCAL
            Else
                GetComponentState = INSTALLSTATE_NOTUSED
            End If 'oFso.FileExists(sPath)
        End If 'Not Err = 0
    End If 'iWiVersionMajor > 2

End Function 'GetComponentState
'=======================================================================================================

'Get a list of client products that are registered to the component
Function GetComponentClients(sComponentId,iComponentState)
    On Error Resume Next
    
    Dim sClients,prod

    sClients = ""
    GetComponentClients = ""
    If Not(iComponentState = INSTALLSTATE_UNKNOWN) Then
        For Each prod in oMsi.ComponentClients(sComponentId)
            If Not Err = 0 Then 
                Err.Clear
                Exit For
            End If 'Not Err = 0
            sClients = sClients & prod & ","
        Next 'prod
        RTrimComma sClients
        GetComponentClients = sClients
    End If 'Not (arrFiles(FILES_COMPONENTSTATE,...

End Function 'GetComponentClients
'=======================================================================================================

'Get the keypath value for the component
Function GetComponentPath(sProductCode,sComponentId,iComponentState)
    On Error Resume Next
    
    Dim sPath

    sPath = "" 
    If iComponentState = INSTALLSTATE_LOCAL Then
        sPath = oMsi.ComponentPath(sProductCode,sComponentId)
    End If 'iComponentState = INSTALLSTATE_LOCAL
    GetComponentPath = sPath

End Function 'GetComponentPath
'=======================================================================================================


'Use WI ProvideAssembly function to identify the path for an assembly.
'Returns the path to the file if the file exists.
'Returns an empty string if file does not exist

Function GetAssemblyPath(sLfn,sKeyPath,sDir)
    On Error Resume Next
    
    Dim sFile,sFolder,sExt,sRoot,sName
    Dim arrTmp
    
    'Defaults
    GetAssemblyPath=""
    sFile="" : sFolder="" : sExt="" : sRoot="" : sName=""
    

    'The componentpath should already point to the correct folder
    'except for components with a registry keypath element.
    'In that case tweak the directory folder to match
    If Left(sKeyPath,1)="0" Then
        sFolder = sDir
        sFolder = oShell.ExpandEnvironmentStrings("%SYSTEMROOT%") & Mid(sFolder,InStr(LCase(sFolder),"\winsxs\"))
        sFile = sLfn
    End If 'Left(sKeyPath,1)="0"
    
    'Figure out the correct file reference
    If sFolder = "" Then sFolder = Left(sKeyPath,InStrRev(sKeyPath,"\"))
    sRoot = Left(sFolder,InStrRev(sFolder,"\",Len(sFolder)-1))
    arrTmp = Split(sFolder,"\")
    If CheckArray(arrTmp) Then sName = arrTmp(UBound(arrTmp)-1)
    If sFile = "" Then sFile = Right(sKeyPath,Len(sKeyPath)-InStrRev(sKeyPath,"\"))
    If oFso.FileExists(sFolder&sLfn) Then 
        sFile = sLfn
    Else
        'Handle .cat, .manifest and .policy files
        If InStr(sLfn,".")>0 Then
            sExt = Mid(sLfn,InStrRev(sLfn,"."))
            Select Case LCase(sExt)
            Case ".cat"
                sFile = Left(sFile,InStrRev(sFile,"."))&"cat"
                If Not oFso.FileExists(sFolder&sFile) Then
                    'Check Manifest folder
                    If oFso.FileExists(sRoot&"Manifests\"&sName&".cat") Then
                        sFolder = sRoot&"Manifests\"
                        sFile = sName&".cat"
                    Else
                        If oFso.FileExists(sRoot&"Policies\"&sName&".cat") Then
                            sFolder = sRoot&"Policies\"
                            sFile = sName&".cat"
                        End If
                    End If
                End If
            Case ".manifest"
                sFile = Left(sFile,InStrRev(sFile,"."))&"manifest"
                If oFso.FileExists(sRoot&"Manifests\"&sName&".manifest") Then
                    sFolder = sRoot&"Manifests\"
                    sFile = sName&".manifest"
                End If
            Case ".policy"
                If iVersionNT < 600 Then
                    sFile = Left(sFile,InStrRev(sFile,"."))&"policy"
                    If oFso.FileExists(sRoot&"Policies\"&sName&".policy") Then
                        sFolder = sRoot&"Policies\"
                        sFile = sName&".policy"
                    End If
                Else
                    sFile = Left(sFile,InStrRev(sFile,"."))&"manifest"
                    If oFso.FileExists(sRoot&"Manifests\"&sName&".manifest") Then
                        sFolder = sRoot&"Manifests\"
                        sFile = sName&".manifest"
                    End If
                End If
            Case Else
            End Select
            
            'Check if the file exists
            If Not oFso.FileExists(sFolder&sFile) Then
                'Ensure the right folder
            End If
        End If 'InStr(sFile,".")>0
    End If
    
    GetAssemblyPath = sFolder&sFile

End Function 'GetAssemblyPath
'=======================================================================================================


'
Function GetFileFullName(iComponentState,sComponentPath,sFileName)
    On Error Resume Next
    
    Dim sFileFullName

    sFileFullName = ""
    If iComponentState = INSTALLSTATE_LOCAL Then
        If Len(sComponentPath) > 2 Then sFileFullName = sComponentPath & sFileName
    End If 'iComponentState = INSTALLSTATE_LOCAL
    GetFileFullName = sFileFullName

End Function 'GetFileFullName
'=======================================================================================================

'
Function GetLongFileName(sMsiFileName)
    On Error Resume Next
    
    Dim sFileTmp
    
    sFileTmp = ""
    sFileTmp = sMsiFileName
    If InStr(sFileTmp,"|") > 0 Then sFileTmp = Mid(sFileTmp,InStr(sFileTmp,"|")+1,Len(sFileTmp))
    GetLongFileName = sFileTmp

End Function 'GetLongFileName
'=======================================================================================================

'
Function GetFileState(iComponentState,sFileFullName)
    On Error Resume Next
    
    GetFileState = INSTALLSTATE_UNKNOWN
    If iComponentState = INSTALLSTATE_LOCAL Then
        If oFso.FileExists(sFileFullName) Then
            GetFileState = INSTALLSTATE_LOCAL
        Else
            GetFileState = INSTALLSTATE_BROKEN
        End If 'oFso.FileExists(sFileFullName) 
    Else
        If oFso.FileExists(sFileFullName) Then
            'This should not happen!
            GetFileState = INSTALLSTATE_LOCAL
        Else
            GetFileState = iComponentState
        End If 'oFso.FileExists(sFileFullName) 
    End If 'iComponentState = INSTALLSTATE_LOCAL

End Function 'GetFileState
'=======================================================================================================

'
Function GetFileVersion(iComponentState,sFileFullName)
    On Error Resume Next
    
    GetFileVersion = ""
    
    If iComponentState = INSTALLSTATE_LOCAL Then
        If oFso.FileExists(sFileFullName) Then
            GetFileVersion = oFso.GetFileVersion(sFileFullName)
        End If 'oFso.FileExists(sFileFullName) 
    Else
        If oFso.FileExists(sFileFullName) Then
            'This should not happen!
            GetFileVersion = oFso.GetFileVersion(sFileFullName)
        End If 'oFso.FileExists(sFileFullName) 
    End If 'iComponentState = INSTALLSTATE_LOCAL

End Function 'GetFileVersion
'=======================================================================================================


'=======================================================================================================
'Module FeatureStates
'=======================================================================================================

'Builds a FeatureTree indicating the FeatureStates 
Sub FindFeatureStates
    If fBasicMode Then Exit Sub
    On Error Resume Next

    Const ADVARCHAR     = 200
    Const MAXCHARACTERS = 255

    Dim Features,oRecordSet,oDicLevel,oDicParent
    Dim sProductCode,sFeature,sFTree,sFParent,sLeft,sRight
    Dim iFoo,iPosMaster,iMaxNestLevel,iNestLevel,iLevel,iFCnt,iLeft,iStart
    Dim arrFName,arrFLevel,arrFParent
    
    'ReDim the global feature array
    ReDim arrFeature (UBound (arrMaster), FEATURE_COLUMNCOUNT)
    
    'Outer loop to iterate through all products
    For iPosMaster = 0 To UBound (arrFeature)
        iFoo = 0
        sFTree = ""
        'Dummy Loop to allow exit out in case of an error
        Do While iFoo = 0
            iFoo = 1
            'Get the ProductCode for this loop
            sProductCode = arrMaster (iPosMaster, COL_PRODUCTCODE)
            'Get the Features colection for this product
            
            'oMsi.Features is only valid for installed per-machine or current user products.
            'The call will fail for advertised and other user products.
            Set Features = oMsi.Features (sProductCode)
            If Not Err = 0 Then
                Err.Clear
                Exit Do
            End If 'Not Err = 0
            'Create the dictionary objects
            Set oDicLevel  = CreateObject ("Scripting.Dictionary")
            Set oDicParent = CreateObject ("Scripting.Dictionary")
            'Prepare a recordset to allow sorting of the root features
            Set oRecordSet = CreateObject ("ADOR.Recordset")
            oRecordSet.Fields.Append "FeatureName", ADVARCHAR, MAXCHARACTERS
            oRecordSet.Open
            If Not Err=0 Then
                Err.Clear
                Exit Do
            End If 'Not Err = 0
            iMaxNestLevel = 0
            
            'Inner loop # 1 to identify all features
            For Each sFeature in Features
                'Reset the nested level counter
                iNestLevel = -1
                'Check for & cache parent feature
                sFParent = oMsi.FeatureParent (sProductCode, sFeature)
                If sFParent = "" Then 
                    'Found a root feature.
                    iNestLevel = 0
                    'Add to recordset for later sorting
                    oRecordSet.AddNew
                    oRecordSet("FeatureName") = TEXTINDENT & sFeature & " (" & TranslateFeatureState (oMsi.FeatureState (sProductCode, sFeature)) & ")"
                    oRecordSet.Update
                Else
                    'Call the recursive function to get the nest level of the current feature
                    iNestLevel = GetFeatureNestLevel (sProductCode, sFeature, iNestLevel)
                    'Add to dictionary arrays
                    oDicLevel.Add sFeature, iNestLevel
                    oDicParent.Add sFeature, sFParent
                End If 'sFParent=""
                'Max nest level is required for second inner loop
                If iNestLevel > iMaxNestLevel Then iMaxNestLevel = iNestLevel
            Next 'sFeature
            
            'First inner loop complete. Sort the root features
            oRecordSet.Sort = "FeatureName"
            oRecordSet.MoveFirst
            'Write the sorted root features to the 'treeview' string
            Do Until oRecordSet.EOF
                sFTree = sFTree & oRecordSet.Fields.Item ("FeatureName") & vbCrLf
                oRecordSet.MoveNext
            Loop 'oRecordSet.EOF
            
            'Copy dic's to array
            arrFName  = oDicLevel.Keys
            arrFLevel = oDicLevel.Items
            arrFParent= oDicParent.Items
            
            '2nd inner loop to add the features to the 'treeview' string
            For iLevel = 1 To iMaxNestLevel
                For iFCnt = 0 To UBound(arrFName)
                    If arrFLevel (iFCnt) = iLevel Then
                        iStart = InStr (sFTree, arrFParent (iFCnt) & " (") + Len (arrFParent (iFCnt))
                        iLeft  = InStr (iStart, sFTree, ")") + 2
                        sLeft  = Left (sFTree, iLeft)
                        sRight = Right (sFTree, Len (sFTree) - iLeft)
                        sFTree = sLeft & TEXTINDENT & FeatureIndent (iLevel) & arrFName (iFCnt) & " (" & TranslateFeatureState (oMsi.FeatureState (sProductCode, arrFName (iFCnt))) & ")" & vbCrLf & sRight
                    End If 'arrFLevel(iFCnt)=i
                Next 'iFCnt
            Next 'iLevel
            
            'Reset objects for next cycle
            Set oRecordSet = Nothing
            Set oDicLevel  = Nothing
            Set oDicParent = Nothing
        Loop 'iFoo=0
        
        arrFeature (iPosMaster, FEATURE_TREE) = vbCrLf & sFTree
    Next 'iProdMaster

End Sub 'FindFeatureStates
'=======================================================================================================

'Translate the FeatureState value
Function TranslateFeatureState(iFState)

    Select Case iFState
    Case INSTALLSTATE_UNKNOWN       : TranslateFeatureState="Unknown"
    Case INSTALLSTATE_ADVERTISED    : TranslateFeatureState="Advertised"
    Case INSTALLSTATE_ABSENT        : TranslateFeatureState="Absent"
    Case INSTALLSTATE_LOCAL         : TranslateFeatureState="Local"
    Case INSTALLSTATE_SOURCE        : TranslateFeatureState="Source"
    Case INSTALLSTATE_DEFAULT       : TranslateFeatureState="Default"
    Case INSTALLSTATE_VIRTUALIZED   : TranslateFeatureState="Virtualized"
    Case INSTALLSTATE_BADCONFIG     : TranslateFeatureState="BadConfig"
    Case Else                       : TranslateFeatureState="Error"
    End Select

End Function 'GetFeatureStateString
'=======================================================================================================
Function FeatureIndent(iNestLevel)
    Dim iLevel
    Dim sIndent

    For iLevel = 1 To iNestLevel
        sIndent = sIndent & vbTab
    Next 'iLevel
    FeatureIndent = sIndent
End Function 'FeatureIndent
'=======================================================================================================

Function GetFeatureNestLevel(sProductCode,sFeature,iNestLevel)
    Dim sParent : sParent = ""
    iNestLevel=iNestLevel+1
    sParent=oMsi.FeatureParent(sProductCode,sFeature)
    If Not sParent = "" Then iNestLevel=GetFeatureNestLevel(sProductCode,sParent,iNestLevel)
    GetFeatureNestLevel = iNestLevel
End Function 'GetFeatureNestLevel
'=======================================================================================================


'=======================================================================================================
'Module Product InstallSource - 
'=======================================================================================================

Sub ReadMsiInstallSources ()
    If fBasicMode Then Exit Sub
    On Error Resume Next


    Dim oProduct, oSumInfo
    Dim iProdCnt, iSourceCnt
    Dim sSource
    Dim MsiSources

    ReDim arrIS(UBound(arrMaster),UBOUND_IS)

    For iProdCnt = 0 To UBound(arrMaster)
        arrIS(iProdCnt,IS_SOURCETYPESTRING) = "No Data Available"
        arrIS(iProdCnt,IS_ORIGINALSOURCE)   = "No Data Available"
        'Add the ProductCode to the array
        arrIS(iProdCnt,IS_PRODUCTCODE) = arrMaster(iProdCnt,COL_PRODUCTCODE)
        
        If arrMaster(iProdCnt, COL_VIRTUALIZED) = 1 Then
            ' do nothing
        Else
            'SourceType
            If oFso.FileExists(arrMaster(iProdCnt,COL_CACHEDMSI)) Then
                Err.Clear
                Set oSumInfo = oMsi.SummaryInformation(arrMaster(iProdCnt,COL_CACHEDMSI),MSIOPENDATABASEMODE_READONLY)
                If Err = 0 Then
                    arrIS(iProdCnt,IS_SOURCETYPE) = oSumInfo.Property(PID_WORDCOUNT)
                    Select Case arrIS(iProdCnt,IS_SOURCETYPE)
                    Case 0 : arrIS(iProdCnt,IS_SOURCETYPESTRING) = "Original source using long file names"
                    Case 1 : arrIS(iProdCnt,IS_SOURCETYPESTRING) = "Original source using short file names"
                    Case 2 : arrIS(iProdCnt,IS_SOURCETYPESTRING) = "Compressed source files using long file names"
                    Case 3 : arrIS(iProdCnt,IS_SOURCETYPESTRING) = "Compressed source files using short file names"
                    Case 4 : arrIS(iProdCnt,IS_SOURCETYPESTRING) = "Administrative image using long file names"
                    Case 5 : arrIS(iProdCnt,IS_SOURCETYPESTRING) = "Administrative image using short file names"
                    Case Else : arrIS(iProdCnt,IS_SOURCETYPESTRING) = "Unknown InstallSource Type"
                    End Select
                Else
                    'ERR_SICONNECTFAILED
                    Cachelog LOGPOS_REVITEM,LOGHEADING_NONE,ERR_CATEGORYERROR,"Product " & arrMaster(iProdCnt,COL_PRODUCTCODE) & DSV & _
                    arrMaster(iProdCnt,COL_PRODUCTNAME) & ": " & ERR_SICONNECTFAILED
                    arrMaster(iProdCnt,COL_ERROR) = arrMaster(iProdCnt,COL_ERROR) & ERR_CATEGORYERROR & ERR_SICONNECTFAILED & CSV
                End If 'Err
            End If
        
            'Get the original InstallSource
            arrIS(iProdCnt,IS_ORIGINALSOURCE) = oMsi.ProductInfo(arrMaster(iProdCnt,COL_PRODUCTCODE),"InstallSource")
            If Not Len(arrIS(iProdCnt,IS_ORIGINALSOURCE)) > 0 Then arrIS(iProdCnt,IS_ORIGINALSOURCE) = "Not Registered"
        
            'Get Network InstallSource(s)
            'With WI 3.x and later the 'Product' object can be used to gather some data
            If iWiVersionMajor > 2 Then
                Err.Clear
                Set oProduct = oMsi.Product(arrMaster(iProdCnt,COL_PRODUCTCODE),arrMaster(iProdCnt,COL_USERSID),arrMaster(iProdCnt,COL_CONTEXT))
                If Err = 0 Then
                    'Get the last used source
                    arrIS(iProdCnt,IS_LASTUSEDSOURCE) = oProduct.SourceListInfo("LastUsedSource")
            
                    Set MsiSources = oProduct.Sources(1)
                    For Each sSource in MsiSources
                        If IsEmpty(arrIS(iProdCnt,IS_ADDITIONALSOURCES)) Then
                            arrIS(iProdCnt,IS_ADDITIONALSOURCES)=sSource 'MsiSources(iSourceCnt)
                        Else
                            arrIS(iProdCnt,IS_ADDITIONALSOURCES)=arrIS(iProdCnt,IS_ADDITIONALSOURCES)&" || "& sSource'MsiSources(iSourceCnt)
                        End If
                    Next 'MsiSources
                End If 'Err
            End If 'iWiVersionMajor
        
            'Get the LIS resiliency source (if applicable)
            If GetDeliveryResiliencySource(arrMaster(iProdCnt,COL_PRODUCTCODE),iProdCnt,sSource) Then arrIS(iProdCnt,IS_LISRESILIENCY)=sSource
        End If 'Not Virtualized        
    Next 'iProdCnt
End Sub 'ReadMsiInstallSources

'=======================================================================================================

'Return True/False and the LIS source path as sSource
'Empty string for sProductCode forces to identify the DownloadCode from Setup.xml
Function GetDeliveryResiliencySource (sProductCode, iPosMaster, sSource)
    On Error Resume Next
    
    Dim arrSources, arrDownloadCodeKeys
    Dim dicDownloadCode
    Dim sSubKeyName, sValue, key, sku, sSkuName, sText, sDownloadCode, sTmpDownloadCode, source
    Dim arrKeys, arrSku
    Dim iVersionMajor, iSrc
    Dim fFound

    GetDeliveryResiliencySource = False
    sSource = Empty
    iVersionMajor = GetVersionMajor(sProductCode)
    Set dicDownloadCode = CreateObject("Scripting.Dictionary")
    
    If iVersionMajor > 11 Then
        'Note: ProductCode doesn't work consistently for this logic
        '      To locate the Setup.xml requires additional logic so the tweak here is to use the 
        '      original source location to identify the DownloadCode
        sText = arrIS(iPosMaster, IS_ORIGINALSOURCE)
        If InStr(source, "{") > 0 Then sDownloadCode = Mid(sText, InStr(sText, "{"), 40) Else sDownloadCode = sProductCode
        dicDownloadCode.Add sDownloadCode, sProductCode
        
        'Find the additional download locations
        'Check if more than one sources are registered
        If InStr(arrIS(iPosMaster, IS_ADDITIONALSOURCES), "||") > 0 Then
            arrSources = Split(arrIS(iPosMaster, IS_ADDITIONALSOURCES), " || ")
            For Each source in arrSources
                If InStr(source, "{") > 0 Then
                    sTmpDownloadCode = Mid(source, InStr(source, "{"), 40)
                    If Not dicDownloadCode.Exists(sTmpDownloadCode) Then dicDownloadCode.Add sTmpDownloadCode, sProductCode
                End If 'InStr
            Next'
        End If 'InStr
        
        
        arrDownloadCodeKeys = dicDownloadCode.Keys
        For iSrc = 0 To dicDownloadCode.Count-1
            sDownloadCode = UCase(arrDownloadCodeKeys(iSrc))
            'Enum HKLM\SOFTWARE\Microsoft\Office\Delivery\SourceEngine\Downloads
            sSubKeyName="SOFTWARE\Microsoft\Office\Delivery\SourceEngine\Downloads\"
            If RegEnumKey(HKLM,sSubKeyName,arrKeys) Then
                For Each key in arrKeys
                    fFound = False
                    If Len(key) > 37 Then
                        fFound = (UCase(Left(key, 38)) = sDownloadCode) OR (UCase(key) = sDownloadCode)
                    Else
                        fFound = (UCase(key) = sDownloadCode)
                    End If 'Len > 37
                    If fFound Then
                        'Found the Delivery reference
                        'Enum the 'Sources' subkey
                        sSubKeyName = sSubKeyName & key & "\Sources\"
                        If RegEnumKey(HKLM,sSubKeyName,arrSku) Then
                            For Each sku in arrSku
                                If RegReadStringValue(HKLM, sSubKeyName & sku, "Path", sValue) Then
                                    sSkuName = ""
                                    sSkuName = " (" & Left(sku, InStr(sku,"(") - 1) &")"
                                    If IsEmpty(sSource) Then
                                        sSource = sValue & sSkuName
                                    Else
                                        sSource = sSource & " || " & sValue & sSkuName
                                    End If 'IsEmpty
                                End If 'RegReadStringValue
                            Next 'sku
                        End If 'RegEnumKey
                        
                        'GUID is unique no need to continue loop once we found a match
                        Exit For
                    End If
                Next 'key
            End If 'RegEnumKey
        Next 'iSrc
        
    ElseIf iVersionMajor = 11 Then
        'Get the DownloadCode
        sSubKeyName = "SOFTWARE\Microsoft\Office\11.0\Delivery\" & sProductCode&"\"
        If RegReadStringValue(HKLM,sSubKeyName, "DownloadCode", sDownloadCode) Then
            sSubKeyName = "SOFTWARE\Microsoft\Office\Delivery\SourceEngine\Downloads\" & sDownloadCode & "\Sources\" & Mid(sProductCode, 2, 36) & "\"
            If RegReadStringValue(HKLM, sSubKeyName, "Path", sValue) Then sSource = sValue
        End If 
    End If 'iVersionMajor
    
    If Not IsEmpty(sSource) Then
        GetDeliveryResiliencySource = True
        WriteDebug sActiveSub, "Delivery resiliency source for " & sProductCode & " returned 'TRUE': " & sSource
    Else
        WriteDebug sActiveSub, "Delivery resiliency source for " & sProductCode & " returned 'FALSE'"
    End If

End Function 'GetDeliveryResiliencySource

'=======================================================================================================


'=======================================================================================================
'Module Product Properties
'=======================================================================================================

'Gather additional properties for the product
'Add them to master array
Sub ProductProperties
    Dim prod,MsiDb
    Dim iPosMaster,iVersionMajor,iContext,iProd
    Dim sProductCode,sSpLevel,sCachedMsi,sSid,sComp,sComp2,sPath,sRef,sProdId,n
    Dim fVer,fCx
    Dim arrCx,arrCxN,arrCxErr,arrKeys,arrTypes,arrNames
    On Error Resume Next

    If NOT fInitArrProdVer Then InitProdVerArrays 
    For iPosMaster = 0 to UBound (arrMaster)
    ' collect properties only for products in state '5' (Default)
        If arrMaster(iPosMaster,COL_STATE) = INSTALLSTATE_DEFAULT OR arrMaster(iPosMaster,COL_STATE) = INSTALLSTATE_VIRTUALIZED Then
            sProductCode = arrMaster(iPosMaster,COL_PRODUCTCODE)
            sSid = arrMaster(iPosMaster,COL_USERSID)
            iContext = arrMaster(iPosMaster,COL_CONTEXT)
        ' ProductID
            sProdId = GetProductId(sProductCode,iPosMaster) 
            If NOT sProdId = "" Then arrMaster(iPosMaster,COL_PRODUCTID) = sProdId
        ' ProductVersion
            arrMaster(iPosMaster,COL_PRODUCTVERSION) = GetProductVersion(sProductCode,iContext,sSid)
            If NOT fBasicMode Then
            ' cached .msi package 
                Set MsiDb = Nothing
                arrMaster(iPosMaster,COL_CACHEDMSI) = GetCachedMsi(sProductCode,iPosMaster) 
                If (NOT arrMaster(iPosMaster,COL_CACHEDMSI) = "") AND (arrMaster(iPosMaster, COL_VIRTUALIZED) = 0) Then
                    Set MsiDb = oMsi.OpenDatabase(arrMaster(iPosMaster,COL_CACHEDMSI), MSIOPENDATABASEMODE_READONLY)
                    ' PackageCode
                    arrMaster(iPosMaster,COL_PACKAGECODE) = GetPackageCode(sProductCode,iPosMaster,MsiDb)
                    ' UpgradeCode
                    arrMaster(iPosMaster,COL_UPGRADECODE) = GetUpgradeCode(MsiDb)
                    ' Transforms
                    arrMaster(iPosMaster,COL_TRANSFORMS) = GetTransforms(sProductCode, iPosMaster)
                    ' InstallDate
                    arrMaster(iPosMaster,COL_INSTALLDATE) = GetInstallDate(sProductCode,iContext,sSid,arrMaster(iPosMaster,COL_CACHEDMSI))
                    ' original .MSI Name
                    arrMaster(iPosMaster,COL_ORIGINALMSI) = GetOriginalMsiName(sProductCode, iPosMaster)
                End If 'msi not empty
            End If 'fBasicMode
            
            ' some of this is only valid for Office products
            If arrMaster(iPosMaster,COL_ISOFFICEPRODUCT) Then
                ' SP level
                iVersionMajor = GetVersionMajor(sProductCode) 
                sSpLevel = OVersionToSpLevel(sProductCode, iVersionMajor, arrMaster(iPosMaster, COL_PRODUCTVERSION)) 
                arrMaster(iPosMaster, COL_SPLEVEL) = sSpLevel 
                ' Build/Origin
                If NOT fBasicMode Then arrMaster(iPosMaster,COL_ORIGIN) = CheckOrigin(MsiDb)
                ' Architecture (Bitness)
                If Left(arrMaster(iPosMaster,COL_PRODUCTVERSION),2)>11 Then
                    If Mid(sProductCode,21,1) = "1" Then arrMaster(iPosMaster, COL_ARCHITECTURE) = "x64" Else arrMaster(iPosMaster, COL_ARCHITECTURE) = "x86"
                End If
                ' Key ComponentStates
                arrMaster(iPosMaster, COL_KEYCOMPONENTS) = GetKeyComponentStates(sProductCode, False)
                
                If NOT fBasicMode Then
                ' Cx
                    fVer = False : fCx = False
                    Select Case iVersionMajor
                    Case 11
                        sComp = "{1EBDE4BC-9A51-4630-B541-2561FA45CCC5}"
                        sRef  = "11.0.8320.0"
                    Case 12
                        sComp = "{0638C49D-BB8B-4CD1-B191-051E8F325736}"
                        sRef  = "12.0.6514.5001"
                        If Mid(sProductCode,11,4) = "0020" Then fVer = True
                    Case 14
                        If Mid(sProductCode,21,1) = "1" Then
                            sComp = "{C0AC079D-A84B-4CBD-8DBA-F1BB44146899}"
                            sComp2= "{E6AC97ED-6651-4C00-A8FE-790DB0485859}"
                        Else
                            sComp = "{019C826E-445A-4649-A5B0-0BF08FCC4EEE}"
                            sComp2= "{398E906A-826B-48DD-9791-549C649CACE5}"
                        End If
                        sRef  = "14.0.5123.5004"
                    Case Else
                        sComp = "" : sComp2 = "" : sRef = ""
                    End Select
                    ' obtain product handle
                    Err.Clear
                    If oMsi.Product(sProductCode, sSid, iContext).ComponentState(sComp) = INSTALLSTATE_LOCAL Then
                        If Err = 0 Then
                            fCx = True
                            sPath = oMsi.ComponentPath(sProductCode,sComp)
                            If oFso.FileExists(sPath) Then
                                fVer = (CompareVersion(oFso.GetFileVersion(sPath),sRef,True) > -1)
                            End If
                            If iVersionMajor = 14 Then
                                If oMsi.Product(sProductCode, sSid, iContext).ComponentState(sComp2) = INSTALLSTATE_LOCAL Then
                                    sPath = oMsi.ComponentPath(sProductCode,sComp2)
                                    If oFso.FileExists(sPath) Then
                                        fVer = (CompareVersion(oFso.GetFileVersion(sPath),sRef,True) > -1)
                                    End If
                                End If 'INSTALLSTATE_LOCAL
                                Err.Clear
                            End If
                        Else
                            Err.Clear
                        End If 'Err = 0 
                    End If 'INSTALLSTATE_LOCAL
                    If fVer Then
                        arrCxN = Array("44","43","58","46")
                        Select Case iVersionMajor
                        Case 11
                            arrCx = Array("53","4F","46","54","57","41","52","45","5C","4D","69","63","72","6F","73","6F","66","74","5C","4F","66","66","69","63","65","5C","31","31","2E","30","5C","57","6F","72","64")
                            arrCxErr = Array("43","75","73","74","6F","6D","20","58","4D","4C","20","66","65","61","74","75","72","65","20","64","69","73","61","62","6C","65","64","20","62","79","20","4B","42","20","39","37","39","30","34","35")
                            If RegEnumValues(HKLM,hAtS(arrCx),arrNames,arrTypes) Then
                                For Each n in arrNames
                                    If UCase(n) = hAtS(arrCxN) Then arrMaster(iPosMaster,COL_NOTES) = arrMaster(iPosMaster,COL_NOTES) & hAtS(arrCxErr) & CSV
                                Next 'n
                            End If
                        Case 12
                            arrCx = Array("53","4F","46","54","57","41","52","45","5C","4D","69","63","72","6F","73","6F","66","74","5C","4F","66","66","69","63","65","5C","31","32","2E","30","5C","57","6F","72","64")
                            arrCxErr = Array("57","6F","72","64","20","43","75","73","74","6F","6D","20","58","4D","4C","20","66","65","61","74","75","72","65","20","64","69","73","61","62","6C","65","64","20","62","79","20","4B","42","20","39","37","34","36","33","31")
                            If Mid(sProductCode,11,4) = "0020" Then
                                If CompareVersion(sRef,GetMsiProductVersion(arrMaster(iPosMaster,COL_CACHEDMSI)),False) < 1 Then
                                    arrCxErr = Array("57","6F","72","64","20","43","75","73","74","6F","6D","20","58","4D","4C","20","66","65","61","74","75","72","65","20","64","69","73","61","62","6C","65","64","20","62","79","20","43","6F","6D","70","61","74","69","62","69","6C","69","74","79","20","50","61","63","6B")
                                End If
                            End If '"0020"
                            If RegEnumValues(HKLM,hAtS(arrCx),arrNames,arrTypes) Then
                                For Each n in arrNames
                                    If UCase(n) = hAtS(arrCxN) Then arrMaster(iPosMaster,COL_NOTES) = arrMaster(iPosMaster,COL_NOTES) & hAtS(arrCxErr) & CSV
                                Next 'n
                            End If
                        Case 14
                            arrCxErr = Array("43","75","73","74","6F","6D","20","58","4D","4C","20","66","65","61","74","75","72","65","20","65","6E","61","62","6C","65","64","20","62","79","20","57","6F","72","64","20","32","30","31","30","20","4B","42","20","32","34","32","38","36","37","20","61","64","64","2D","69","6E")
                            For Each prod in arrMaster
                                If Mid(prod,11,4)="0126" Then arrMaster(iPosMaster,COL_NOTES) = arrMaster(iPosMaster,COL_NOTES) & hAtS(arrCxErr) & CSV
                            Next 'prod
                        Case Else
                        End Select
                    Else
                        If (iVersionMajor = 14) AND fCx Then
                            arrCxErr = Array("43","75","73","74","6F","6D","20","58","4D","4C","20","66","65","61","74","75","72","65","20","66","6F","72","20","74","68","65","20","62","69","6E","61","72","79","20","2E","64","6F","63","20","66","6F","72","6D","61","74","20","72","65","71","75","69","72","65","73","20","4B","42","20","32","34","31","33","36","35","39")
                            arrMaster(iPosMaster,COL_NOTES) = arrMaster(iPosMaster,COL_NOTES) & hAtS(arrCxErr) & CSV
                        End If '14
                    End If 'fVer
                End If 'fBasicMode
            Else
                arrMaster(iPosMaster,COL_ORIGIN) = "n/a"
            ' checks for known add-ins
            ' POWERPIVOT_2010
                If InStr(POWERPIVOT_2010,arrMaster(iPosMaster,COL_UPGRADECODE)) > 0 Then
                    If CompareVersion(arrMaster(iPosMaster,COL_PRODUCTVERSION),"10.50.1747.0",True) < 1 Then
                        arrMaster(iPosMaster,COL_NOTES) = arrMaster(iPosMaster,COL_NOTES) & ERR_CATEGORYWARN & "This is a preview version. Please obtain version 10.50.1747.0 or higher." & CSV
                    End If
                End If 'POWERPIVOT_2010
            End If 'IsOfficeProduct
        End If 'INSTALLSTATE_DEFAULT
    Next 'iPosMaster
End Sub 'ProductProperties
'=======================================================================================================

'The name of the original installation package 'PackageName' is obtained from HKCR.
'This limits the availability to products which are installed 'per-machine' or for the current user!
'Exception situations WI < 3.x or user profile not available are covered in the 'Error handler'
Function GetOriginalMsiName(sProductCode, iPosMaster)
    Dim iPos
    Dim sCompGuid, sRegName
    Dim fVirtual
    On Error Resume Next

    fVirtual = (arrMaster(iPosMaster, COL_VIRTUALIZED) = 1)
    sRegName = ""
    If NOT fVirtual Then sRegName = oMsi.ProductInfo(sProductCode,"PackageName")

    'Error Handler
    If (Not Err = 0) OR sRegName = "" Then
        'This can happen if WI < 3.x or product is installed for other user
        Err.Clear
        iPos = GetArrayPosition(arrMaster,sProductCode)
        sCompGuid = GetCompressedGuid(sProductCode)
        sRegName = GetRegOriginalMsiName(sCompGuid,arrMaster(iPos,COL_CONTEXT),arrMaster(iPos,COL_USERSID))
        If sRegName = "-" AND NOT fVirtual Then arrMaster(iPos,COL_ERROR) = arrMaster(iPos,COL_ERROR) & ERR_CATEGORYERROR & ERR_BADMSINAMEMETADATA & CSV
    End If
    GetOriginalMsiName = sRegName
    
End Function
'=======================================================================================================

'The 'Transforms' property is obtained from HKCR.
'This limits the availability to products which are installed 'per-machine' or for the current user!
'Exception situations WI < 3.x or user profile not available are covered in the 'Error handler'
Function GetTransforms(sProductCode, iPosMaster)
    Dim sTransforms, sCompGuid, sRegTransforms
    Dim iPos
    On Error Resume Next

    GetTransforms = "-" : sTransforms = ""
    If arrMaster(iPosMaster, COL_VIRTUALIZED) = 0 Then sTransforms = oMsi.ProductInfo(sProductCode,"Transforms")

    'Error Handler
    If NOT Err = 0 OR arrMaster(iPosMaster, COL_VIRTUALIZED) = 1 Then
        Err.Clear
        iPos = GetArrayPosition(arrMaster,sProductCode)
        sCompGuid = GetCompressedGuid(sProductCode)
        sTransforms = GetRegTransforms(sCompGuid,arrMaster(iPos,COL_CONTEXT),arrMaster(iPos,COL_USERSID))
    End If
    If Len(sTransforms) > 0 Then GetTransforms = sTransforms
End Function
'=======================================================================================================

'InstallDate is available as part of the ProductInfo.
'It's stored in the global key. Introduced with WI 3.x
Function GetInstallDate(sProductCode,iContext,sSid,sCachedMsi)
    Dim iPos
    Dim hDefKey
    Dim sSubKeyName, sName, sValue, sDateLocalized, sDateNormalized, sYY, sMM, sDD
    On Error Resume Next

    GetInstallDate = "" 
    hDefKey = HKEY_LOCAL_MACHINE
    sSubKeyName = GetRegConfigKey(sProductCode,iContext,sSid,True) & "InstallProperties"
    sName = "InstallDate"
    GetInstallDate = "-"
    If RegReadValue(hDefKey,sSubKeyName,sName,sValue,"REG_EXPAND_SZ") Then GetInstallDate = sValue

    'The InstallDate is reset with every patch transaction
    'As a workaround the CreateDate of the cached .msi package will be used to obtain the correct date
    If oFso.FileExists(sCachedMsi) Then
        'GetInstallDate = oFso.GetFile(sCachedMsi).DateCreated
        sDateLocalized = oFso.GetFile(sCachedMsi).DateCreated
        sYY = Year(sDateLocalized)
        sMM = Right("0" & Month(sDateLocalized), 2)
        sDD = Right("0" & Day(sDateLocalized), 2)
        sDateNormalized = sYY & " " & sMM & " " & sDD & " (yyyy mm dd)"
        GetInstallDate = sDateNormalized

    End If
    
End Function 'GetInstallDate
'=======================================================================================================

'The package code associates a .msi file with an application or product
'This property is used for source verification
'3 possible checks here:
' a) Installer.ProductInfo     Note: ProductInfo object is limited to per-machine and current user scope
' b) SummaryInformation stream from cached .msi
' c) SummaryInformation stream from .msi in InstallSource
Function GetPackageCode(sProductCode,iPosMaster,MsiDb)
    Dim sValidate, sCompGuid, sPackageCode
    Dim oSumInfo
    On Error Resume Next

    sPackageCode = ""
    If arrMaster(iPosMaster, COL_VIRTUALIZED) = 0 Then sPackageCode = oMsi.ProductInfo(sProductCode, "PackageCode")
    If (Not Err = 0) OR (sPackageCode="") Then
    ' Error Handler
        sCompGuid = GetCompressedGuid(sProductCode)
        sPackageCode = GetRegPackageCode(sCompGuid,arrMaster(iPosMaster,COL_CONTEXT),arrMaster(iPosMaster,COL_USERSID))
        Exit Function
    End If
    If Not sPackageCode = "n/a" Then
        If Not IsValidGuid(sPackageCode,GUID_UNCOMPRESSED) Then
            If fGuidCaseWarningOnly Then
                arrMaster(iPosMaster,COL_NOTES) = arrMaster(iPosMaster,COL_NOTES) & ERR_CATEGORYNOTE & ERR_GUIDCASE & DOT & sErrBpa & CSV
            Else
                Cachelog LOGPOS_REVITEM,LOGHEADING_NONE, ERR_CATEGORYERROR,"Product " & sProductCode & DSV & arrMaster(iPosMaster,COL_PRODUCTNAME) & _
                         ": " & sError & " for PackageCode '" & sPackageCode & "'" & DOT & sErrBpa
                sError = "" : sErrBpa = ""
                arrMaster(iPosMaster,COL_ERROR) = arrMaster(iPosMaster,COL_ERROR) & ERR_CATEGORYERROR & sError & DOT & sErrBpa & CSV
            End If 'fGuidCaseWarningOnly
        End If
    End If
' Scan cached .msi
    If Not arrMaster(iPosMaster,COL_CACHEDMSI) = "" Then
	    Set oSumInfo = MsiDb.SummaryInformation(MSIOPENDATABASEMODE_READONLY)
	    If Not (Err = 0) Then
	        arrMaster(iPosMaster,COL_NOTES) = arrMaster(iPosMaster,COL_NOTES) & ERR_CATEGORYWARN & ERR_INITSUMINFO & CSV
            Exit Function
	    End If 'Not Err
	    If Not sPackageCode = oSumInfo.Property(PID_REVNUMBER) Then 
	        arrMaster(iPosMaster,COL_ERROR) = arrMaster(iPosMaster,COL_ERROR) & ERR_CATEGORYERROR & ERR_PACKAGECODEMISMATCH & CSV
	        Cachelog LOGPOS_REVITEM,LOGHEADING_NONE,ERR_CATEGORYERROR,"Product " & sProductCode & DSV & arrMaster(iPosMaster,COL_PRODUCTNAME) & _
                     ": " & ERR_PACKAGECODEMISMATCH  & DOT & BPA_PACKAGECODEMISMATCH
        End If         
    End If 'arrMaster
' Scan .msi in InstallSource has to be deferred to module InstallSource
    GetPackageCode = sPackageCode
End Function 'GetPackageCode
'=======================================================================================================

Function GetUpgradeCode (MsiDb)
    Dim Record
    Dim qView
    On Error Resume Next
    GetUpgradeCode = ""
    If MsiDb Is Nothing Then Exit Function
    Set qView = MsiDb.OpenView("SELECT `Value` FROM Property WHERE `Property`='UpgradeCode'")
    qView.Execute()
    Set Record = qView.Fetch()
    If NOT Err = 0 Then Exit Function
    GetUpgradeCode = Record.StringData(1)
End Function 'GetPackageCode
'=======================================================================================================

'Read the Build / Origin property from the cached .msi
Function CheckOrigin(MsiDb)
    Dim sQuery, sCachedMsi
    Dim Record
    Dim qView
    On Error Resume Next

' not all products do support this so the return value is not guaranteed.
    CheckOrigin = ""
    If MsiDb Is Nothing Then Exit Function
' read the 'Build' entry first
	sQuery = "SELECT `Value` FROM Property WHERE `Property` = 'BUILD'"
	Set qView = MsiDb.OpenView(sQuery)
	qView.Execute 
	Set Record = qView.Fetch()
	If Not Record Is Nothing Then
	    CheckOrigin = Record.StringData(1)
	End If 'Is Nothing
' add the 'Origin' entry to the same field
	sQuery = "SELECT `Value` FROM Property WHERE `Property` = 'ORIGIN'"
	Set qView = MsiDb.OpenView(sQuery)
	qView.Execute 
	Set Record = qView.Fetch()
	If Not Record Is Nothing Then
	    CheckOrigin = CheckOrigin & " / " & Record.StringData(1)
	End If 'Is Nothing
End Function
'=======================================================================================================

Function GetCachedMsi(sProductCode,iPosMaster)
    Dim sCachedMsi
    Dim oApp
    Dim fVirtual
    On Error Resume Next

    fVirtual = (arrMaster(iPosMaster, COL_VIRTUALIZED) = 1)
    sCachedMsi = ""
    If NOT fVirtual Then
    ' for WI >= 3.x we can use 'InstallProperty' for earlier versions we need to stick to 'ProductInfo'
        If iWiVersionMajor > 2 Then
            Set oApp = oMsi.Product(sProductCode,arrMaster(iPosMaster,COL_USERSID),arrMaster(iPosMaster,COL_CONTEXT))
            sCachedMsi = oApp.InstallProperty("LocalPackage")
        ElseIf (arrMaster(iPosMaster,COL_USERSID) = sCurUserSid) Or (arrMaster(iPosMaster,COL_CONTEXT) = MSIINSTALLCONTEXT_MACHINE) Then
        ' ProductInfo only available for per-machine or current users' products
            sCachedMsi = oMsi.ProductInfo(sProductCode, "LocalPackage")
        Else
        ' rely on error handling to retain the value from direct registry read
        End If 'iWiVersionMajor
    End If 'fVirtual

    If Not Err = 0 Then
    ' ensure direct registry detection is done
        sCachedMsi=""
    ' log Error
        Cachelog LOGPOS_REVITEM,LOGHEADING_NONE,ERR_CATEGORYERROR,"Product " & sProductCode & DSV & arrMaster(iPosMaster,COL_PRODUCTNAME) & _
                 ": " & ERR_PACKAGEAPIFAILURE & "."
    End If
    
    If sCachedMsi="" Then
        sCachedMsi = GetRegCachedMsi(GetCompressedGuid(sProductCode),iPosMaster)
    End If 'sCachedMsi=""
' ensure to not reference a non existent .msi
    If sCachedMsi="" Then
        If NOT fVirtual Then arrMaster(iPosMaster,COL_ERROR) = arrMaster(iPosMaster,COL_ERROR) & ERR_CATEGORYERROR & ERR_BADPACKAGEMETADATA & CSV
    Else
        If Not oFso.FileExists (sCachedMsi) Then 
            Cachelog LOGPOS_REVITEM,LOGHEADING_NONE,ERR_CATEGORYERROR,"Product " & sProductCode & " - " &_
                     arrMaster(iPosMaster,COL_PRODUCTNAME) & ": " & ERR_LOCALPACKAGEMISSING & ": " & sCachedMsi
            arrMaster(iPosMaster,COL_ERROR) = arrMaster(iPosMaster,COL_ERROR) & ERR_CATEGORYERROR & ERR_LOCALPACKAGEMISSING & CSV
            sCachedMsi = ""
        End If 'oFso.FileExists
    End If
    GetCachedMsi = sCachedMsi
End Function 'GetCachedMsi
'=======================================================================================================

Function GetRegCachedMsi (sProductCodeCompressed,iPosMaster)
    Dim hDefKey
    Dim sSubKeyName,sValue,sSid,sName
    Dim iContext
    On Error Resume Next

    GetRegCachedMsi = ""
    'Go global
    hDefKey = HKLM
    iContext = arrMaster(iPosMaster,COL_CONTEXT)
    sSid = arrMaster(iPosMaster,COL_USERSID)
    sName = "LocalPackage"
    'Tweak managed to unmanaged to avoid link to managed global key
    If iContext = MSIINSTALLCONTEXT_USERMANAGED Then 
        iContext = MSIINSTALLCONTEXT_USERUNMANAGED
        sName = "ManagedLocalPackage"
    End If
    
    sSubKeyName = GetRegConfigKey(sProductCodeCompressed,iContext,sSid,True) & "InstallProperties\"
    If RegReadStringValue(hDefKey,sSubKeyName,sName,sValue) Then GetRegCachedMsi = sValue

End Function
'=======================================================================================================

Function GetProductId(sProductCode,iPosMaster)
    Dim sProductId
    Dim oApp
    On Error Resume Next

    GetProductId = ""
    
    If arrMaster(iPosMaster,COL_CONTEXT) = MSIINSTALLCONTEXT_C2RV2 Then
        sProductId = GetRegProductId(GetCompressedGuid(sProductCode),iPosMaster)
        Exit Function
    End If
    If arrMaster(iPosMaster,COL_CONTEXT) = MSIINSTALLCONTEXT_C2RV3 Then
        sProductId = GetRegProductId(GetCompressedGuid(sProductCode),iPosMaster)
        Exit Function
    End If

    'For WI >= 3.x we can use 'InstallProperty' for earlier versions we need to stick to 'ProductInfo'
    If iWiVersionMajor > 2 Then
        Set oApp = oMsi.Product(sProductCode,arrMaster(iPosMaster,COL_USERSID),arrMaster(iPosMaster,COL_CONTEXT))
        sProductId = oApp.InstallProperty("ProductID")
    ElseIf (arrMaster(iPosMaster,COL_USERSID) = sCurUserSid) Or (arrMaster(iPosMaster,COL_CONTEXT) = MSIINSTALLCONTEXT_MACHINE) Then
        'ProductInfo only available for per-machine or current users' products
        sProductId = oMsi.ProductInfo(sProductCode, "ProductID")
    Else
        'Rely on error handling to retain the value from direct registry read
    End If 'iWiVersionMajor

    If Not Err = 0 Then
        'Ensure direct registry detection is done
        sProductId=""
    End If
    
    If sProductId="" Then
        sProductId = GetRegProductId(GetCompressedGuid(sProductCode),iPosMaster)
    End If 'sCachedMsi=""

    GetProductId = sProductId
End Function 'GetProductId
'=======================================================================================================

Function GetRegProductId (sProductCodeCompressed, iPosMaster)
    Dim hDefKey
    Dim sSubKeyName, sValue, sSid, sName
    Dim iContext
    On Error Resume Next

    GetRegProductId = ""
    'Go global
    hDefKey = HKLM
    iContext = arrMaster(iPosMaster,COL_CONTEXT)
    sSid = arrMaster(iPosMaster,COL_USERSID)
    sName = "ProductID"
    'Tweak managed to unmanaged to avoid link to managed global key
    If iContext = MSIINSTALLCONTEXT_USERMANAGED Then 
        iContext = MSIINSTALLCONTEXT_USERUNMANAGED
        sName = "ProductID"
    End If
    
    sSubKeyName = GetRegConfigKey(sProductCodeCompressed,iContext,sSid,True) & "InstallProperties\"
    If RegReadStringValue(hDefKey,sSubKeyName,sName,sValue) Then GetRegProductId = sValue

End Function 'GetRegProductId
'=======================================================================================================

'Get the ProductVersion string from WI ProductInfo
Function GetProductVersion (sProductCode,iContext,sSid)
    Dim sTmp
    On Error Resume Next

    If iContext = MSIINSTALLCONTEXT_C2RV2 Then
         GetProductVersion = GetRegProductVersion(sProductCode,iContext,sSid)
         Exit Function
    End If
    If iContext = MSIINSTALLCONTEXT_C2RV3 Then
         GetProductVersion = GetRegProductVersion(sProductCode,iContext,sSid)
         Exit Function
    End If
    
    sTmp = ""
    sTmp = oMsi.ProductInfo (sProductCode, "VersionString")
    If (sTmp = "") OR (NOT Err = 0) Then
        Err.Clear
        sTmp = GetRegProductVersion(sProductCode,iContext,sSid)
    End If
    GetProductVersion = sTmp

End Function
'=======================================================================================================

'Get the ProductVersion from Registry
Function GetRegProductVersion (sProductCode,iContext,sSid)
    Dim hDefKey
    Dim sSubKeyName,sValue
    Dim iTmpContext
    On Error Resume Next

    GetRegProductVersion = "Error"
    hDefKey = HKEY_LOCAL_MACHINE
    If iContext = MSIINSTALLCONTEXT_USERMANAGED Then 
        iTmpContext = MSIINSTALLCONTEXT_USERUNMANAGED
    Else
        iTmpContext = iContext
    End If 'iContext = MSIINSTALLCONTEXT_USERMANAGED
    sSubKeyName = GetRegConfigKey(GetCompressedGuid(sProductCode),iTmpContext,sSid,True) & "InstallProperties\"
    If RegReadStringValue(hDefKey,sSubKeyName,"DisplayVersion",sValue) Then GetRegProductVersion = sValue
   
End Function
'=======================================================================================================

'Translate the Office ProductVersion to the service pack level
Function OVersionToSpLevel (sProductCode,iVersionMajor,sProductVersion)
    On Error Resume Next
    
    'SKU identifier constants for SP level detection
    Const O16_EXCEPTION = ""
    Const O15_EXCEPTION = ""
    Const O14_EXCEPTION = "007A,007B,007C,007D,007F,2005"
'#Devonly   O12_Server = "1014,1015,104B,104E,1080,1088,10D7,10D8,10EB,10F5,10F6,10F7,10F8,10FB,10FC,10FD,1103,1104,110D,1105,1110,1121,1122"    
    Const O12_EXCEPTION = "001C,001F,0020,003F,0045,00A4,00A7,00B0,00B1,00B2,00B9,011F,CFDA"
    Const O11_EXCEPTION = "14,15,16,17,18,19,1A,1B,1C,24,32,3A,3B,44,51,52,53,5E,A1,A4,A9,E0"
    Const O10_EXCEPTION = "17,1D,25,27,30,36,3A,3B,51,52,53,54"
    Const O09_EXCEPTION = "3A,3B,3C,5F"

    Dim iSpCnt,iExptnCnt,iLevel,iRetry
    Dim sSpLevel,sSku

    iLevel = 0 : iRetry = 0
    Select Case iVersionMajor

    Case 9
        'Sku ProductID is 2 digits starting at pos 4
        sSku = Mid (sProductCode, 4, 2)
        If InStr (O09_EXCEPTION, sSku) > 0 Then
            For iExptnCnt = 1 To UBound (arrProdVer09, 1)
                If InStr (arrProdVer10 (iExptnCnt, 0), sSku) > 0 Then Exit For
            Next 'iExptnCnt
        Else
            iExptnCnt = 0
        End If 'InStr(O09_Exception,sSku)>0
        
        For iRetry = 0 To 1
            For iSpCnt = 1 To UBound (arrProdVer09, 2)
                If sProductVersion = Left (arrProdVer09 (iExptnCnt, iSpCnt), Len (sProductVersion)) Then 
                    'Special release references are noted within same field with a "," separator
                    If InStr (arrProdVer09 (iExptnCnt, iSpCnt), ",") > 0 Then
                        OVersionToSpLevel = Mid (arrProdVer09 (iExptnCnt, iSpCnt), InStr (arrProdVer09 (iExptnCnt, iSpCnt), ",") + 1, Len (arrProdVer09 (iExptnCnt, iSpCnt)))
                        Exit Function
                    Else
                        iLevel = iSpCnt
                    End If
                End If
            Next 'iSpCnt
            If iLevel > 0 Then Exit For
            'Did not find the SP level yet. Retry with core build numbers
            iExptnCnt = 0
        Next 'iRetry

    Case 10
        'Sku ProductID is 2 digits starting at pos 4
        sSku = Mid(sProductCode,4,2)
        If InStr(O10_EXCEPTION,sSku)>0 Then
            For iExptnCnt = 1 To UBound(arrProdVer10,1)
                If InStr(arrProdVer10(iExptnCnt,0),sSku)>0 Then Exit For
            Next 'iExptnCnt
        Else
            iExptnCnt = 0
        End If 'InStr(O10_Exception,sSku)>0
        
        For iRetry = 0 To 1
            For iSpCnt = 1 To UBound(arrProdVer10,2)
                If sProductVersion = Left(arrProdVer10(iExptnCnt,iSpCnt),Len(sProductVersion)) Then 
                    'Special release references are noted within same field with a "," separator
                    If InStr(arrProdVer10(iExptnCnt,iSpCnt),",")>0 Then
                        OVersionToSpLevel = Mid(arrProdVer10(iExptnCnt,iSpCnt),InStr(arrProdVer10(iExptnCnt,iSpCnt),",")+1,Len(arrProdVer10(iExptnCnt,iSpCnt)))
                        Exit Function
                    Else
                        iLevel = iSpCnt
                        Exit For
                    End If
                End If
            Next 'iSpCnt
            If iLevel > 0 Then Exit For
            'Did not find the SP level yet. Retry with core build numbers
            iExptnCnt = 0
        Next 'iRetry

    Case 11
        'Sku ProductID is 2 digits starting at pos 4
        sSku = Mid(sProductCode,4,2)
        If InStr(O11_EXCEPTION,sSku)>0 Then
            For iExptnCnt = 1 To UBound(arrProdVer11,1)
                If InStr(arrProdVer11(iExptnCnt,0),sSku)>0 Then Exit For
            Next 'iExptnCnt
        Else
            iExptnCnt = 0
        End If 'InStr(O11_Exception,sSku)>0
        
        For iRetry = 0 To 1
            For iSpCnt = 1 To UBound(arrProdVer11,2)
                If sProductVersion = Left(arrProdVer11(iExptnCnt,iSpCnt),Len(sProductVersion)) Then 
                    'Special release references are noted within same field with a "," separator
                    If InStr(arrProdVer11(iExptnCnt,iSpCnt),",")>0 Then
                        OVersionToSpLevel = Mid(arrProdVer11(iExptnCnt,iSpCnt),InStr(arrProdVer11(iExptnCnt,iSpCnt),",")+1,Len(arrProdVer11(iExptnCnt,iSpCnt)))
                        Exit Function
                    Else
                        iLevel = iSpCnt
                        Exit For
                    End If
                End If
            Next 'iSpCnt
            If iLevel > 0 Then Exit For
            'Did not find the SP level yet. Retry with core build numbers
            iExptnCnt = 0
        Next 'iRetry

    Case 12
        'Sku ProductID is 4 digits starting at pos 11
        sSku = Mid(sProductCode,11,4)
        If InStr(O12_EXCEPTION,sSku)>0 Then
            For iExptnCnt = 2 To UBound(arrProdVer12,1)
                If InStr(arrProdVer12(iExptnCnt,0),sSku)>0 Then Exit For
            Next 'iExptnCnt
        ElseIf Left(sSku,1)="1" Then 'Server SKU
            iExptnCnt = 1
        Else
            iExptnCnt = 0
        End If 'InStr(O12_Exception,sSku)>0
        
        For iRetry = 0 To 1
            For iSpCnt = 1 To UBound(arrProdVer12,2)
                If Left(sProductVersion,10) = Left(arrProdVer12(iExptnCnt,iSpCnt),10) Then 
                    'Special release references are noted within same field with a "," separator
                    If InStr(arrProdVer12(iExptnCnt,iSpCnt),",")>0 Then
                        OVersionToSpLevel = Mid(arrProdVer12(iExptnCnt,iSpCnt),InStr(arrProdVer12(iExptnCnt,iSpCnt),",")+1,Len(arrProdVer12(iExptnCnt,iSpCnt)))
                        Exit Function
                    Else
                        iLevel = iSpCnt
                        Exit For
                    End If
                End If
            Next 'iSpCnt
            If iLevel > 0 Then Exit For
            'Did not find the SP level yet. Retry with core build numbers
            iExptnCnt = 0
        Next 'iRetry

    Case 14
        'Sku ProductID is 4 digits starting at pos 11
        sSku = Mid(sProductCode,11,4)
        If InStr(O14_EXCEPTION,sSku)>0 Then
            For iExptnCnt = 1 To UBound(arrProdVer14,1)
                If InStr(arrProdVer14(iExptnCnt,0),sSku)>0 Then Exit For
            Next 'iExptnCnt
        Else
            iExptnCnt = 0
        End If 'InStr(O14_Exception,sSku)>0
        
        For iRetry = 0 To 1
            For iSpCnt = 1 To UBound(arrProdVer14,2)
                If Left(sProductVersion,10) = Left(arrProdVer14(iExptnCnt,iSpCnt),10) Then 
                    'Special release references are noted within same field with a "," separator
                    If InStr(arrProdVer14(iExptnCnt,iSpCnt),",")>0 Then
                        OVersionToSpLevel = Mid(arrProdVer14(iExptnCnt,iSpCnt),InStr(arrProdVer14(iExptnCnt,iSpCnt),",")+1,Len(arrProdVer14(iExptnCnt,iSpCnt)))
                        Exit Function
                    Else
                        iLevel = iSpCnt
                        Exit For
                    End If
                End If
            Next 'iSpCnt
            If iLevel > 0 Then Exit For
            'Did not find the SP level yet. Retry with core build numbers
            iExptnCnt = 0
        Next 'iRetry

    Case 15
        'Sku ProductID is 4 digits starting at pos 11
        sSku = Mid(sProductCode, 11, 4)
        If InStr(O15_EXCEPTION, sSku) > 0 Then
            For iExptnCnt = 1 To UBound(arrProdVer15, 1)
                If InStr(arrProdVer15(iExptnCnt, 0), sSku) > 0 Then Exit For
            Next 'iExptnCnt
        Else
            iExptnCnt = 0
        End If 'InStr(O15_Exception, sSku) > 0
        
        For iRetry = 0 To 1
            For iSpCnt = 1 To UBound(arrProdVer15, 2)
                If Left(sProductVersion, 10) = Left(arrProdVer15(iExptnCnt, iSpCnt), 10) Then 
                    'Special release references are noted within same field with a "," separator
                    If InStr(arrProdVer15(iExptnCnt, iSpCnt), ",") > 0 Then
                        OVersionToSpLevel = Mid(arrProdVer15(iExptnCnt, iSpCnt), InStr(arrProdVer15(iExptnCnt, iSpCnt), ",") + 1, Len(arrProdVer15(iExptnCnt, iSpCnt)))
                        Exit Function
                    Else
                        iLevel = iSpCnt
                        Exit For
                    End If
                End If
            Next 'iSpCnt
            If iLevel > 0 Then Exit For
            'Did not find the SP level yet. Retry with core build numbers
            iExptnCnt = 0
        Next 'iRetry
        
        Case 16
        'Sku ProductID is 4 digits starting at pos 11
        sSku = Mid(sProductCode, 11, 4)
        If InStr(O16_EXCEPTION, sSku) > 0 Then
            For iExptnCnt = 1 To UBound(arrProdVer16, 1)
                If InStr(arrProdVer16(iExptnCnt, 0), sSku) > 0 Then Exit For
            Next 'iExptnCnt
        Else
            iExptnCnt = 0
        End If 'InStr(O16_Exception, sSku) > 0
        
        For iRetry = 0 To 1
            For iSpCnt = 1 To UBound(arrProdVer16, 2)
                If Left(sProductVersion, 10) = Left(arrProdVer16(iExptnCnt, iSpCnt), 10) Then 
                    'Special release references are noted within same field with a "," separator
                    If InStr(arrProdVer16(iExptnCnt, iSpCnt), ",") > 0 Then
                        OVersionToSpLevel = Mid(arrProdVer16(iExptnCnt, iSpCnt), InStr(arrProdVer16(iExptnCnt, iSpCnt), ",") + 1, Len(arrProdVer16(iExptnCnt, iSpCnt)))
                        Exit Function
                    Else
                        iLevel = iSpCnt
                        Exit For
                    End If
                End If
            Next 'iSpCnt
            If iLevel > 0 Then Exit For
            'Did not find the SP level yet. Retry with core build numbers
            iExptnCnt = 0
        Next 'iRetry
Case Else
    End Select
        
    Select Case iLevel
    Case 1 : sSpLevel = "RTM"
    Case 2 : sSpLevel = "SP1"
    Case 3 : sSpLevel = "SP2"
    Case 4 : sSpLevel = "SP3"
    Case Else : sSpLevel = ""
    End Select

    OVersionToSpLevel = sSpLevel
    
End Function 'OVersionToSpLevel
'=======================================================================================================

'Initialize arrays for translation ProductVersion -> ServicePackLevel
Sub InitProdVerArrays
    On Error Resume Next

' O16 Products -> KB ?
    ReDim arrProdVer16(0,0) 'n,1=RTM
    arrProdVer16(0,0)="" 

' 2013 Products -> KB 2786054
    ReDim arrProdVer15(0,44) 'n,1=RTM
    arrProdVer15(0,0)="" : arrProdVer15(0,1) = "15.0.4420.1017" : arrProdVer15(0,2) = "15.0.4569.1507"
    arrProdVer15(0,3) = "15.0.4454.1004,2013/01" : arrProdVer15(0,4) = "15.0.4454.1511,2013/02" 
    arrProdVer15(0,5) = "15.0.4481.1005,2013/03" : arrProdVer15(0,6) = "15.0.4481.1510,2013/04" : arrProdVer15(0,7) = "15.0.4505.1006,2013/05" 
    arrProdVer15(0,8) = "15.0.4505.1510,2013/06" : arrProdVer15(0,8) = "15.0.4517.1005,2013/07" : arrProdVer15(0,9) = "15.0.4517.1509,2013/08"
    arrProdVer15(0,10) = "15.0.4535.1004,2013/09" : arrProdVer15(0,11) = "15.0.4535.1511,2013/10" : arrProdVer15(0,12) = "15.0.4551.1005,2013/11"
    arrProdVer15(0,13) = "15.0.4551.1011,2013/12" : arrProdVer15(0,14) = "15.0.4551.1512,2014/01"
    arrProdVer15(0,15) = "15.0.4569.1000,SP1_Preview" : arrProdVer15(0,16) = "15.0.4569.1001,SP1_Preview" : arrProdVer15(0,17) = "15.0.4569.1002,SP1_Preview" : arrProdVer15(0,18) = "15.0.4569.1003,SP1_Preview"
    arrProdVer15(0,19) = "15.0.4569.1508,2014/03" : arrProdVer15(0,20) = "15.0.4605.1003,2014/04" : arrProdVer15(0,21) = "15.0.4615.1002,2014/05" : arrProdVer15(0,22) = "15.0.4623.1003,2014/06"
    arrProdVer15(0,23) = "15.0.4631.1002,2014/07" : arrProdVer15(0,24) = "15.0.4631.1004,2014/07+" : arrProdVer15(0,25) = "15.0.4641.1002,2014/08" : arrProdVer15(0,26) = "15.0.4641.1003,2014/08+"
    arrProdVer15(0,27) = "15.0.4649.1001,2014/09" : arrProdVer15(0,28) = "15.0.4649.1001,2014/09+" : arrProdVer15(0,29) = "15.0.4649.1004,2014/09++" : arrProdVer15(0,30) = "15.0.4659.1001,2014/10"
    arrProdVer15(0,31) = "15.0.4667.1002,2014/11" : arrProdVer15(0,32) = "15.0.4675.1002,2014/12" : arrProdVer15(0,33) = "15.0.4675.1003,2014/12+" : arrProdVer15(0,34) = "15.0.4693.1001,2015/02"
    arrProdVer15(0,35) = "15.0.4693.1002,2015/02+" : arrProdVer15(0,36) = "15.0.4701.1002,2015/03" : arrProdVer15(0,37) = "15.0.4711.1002,2015/04" : arrProdVer15(0,38) = "15.0.4711.1003,2015/04+"
    arrProdVer15(0,39) = "15.0.4719.1002,2015/05" : arrProdVer15(0,40) = "15.0.4727.1002,2015/06" : arrProdVer15(0,41) = "15.0.4727.1003,2015/06+" : arrProdVer15(0,42) = "15.0.4737.1003,2015/07"
    arrProdVer15(0,43) = "15.0.4745.1001,2015/08" : arrProdVer15(0,44) = "15.0.4745.1001,2015/08+"

' 2010 Products -> KB 2186281
    ReDim arrProdVer14(6,3) 'n,1=RTM; n,2=SP1
    arrProdVer14(0,0)="" : arrProdVer14(0,1)="14.0.4763.1000" : arrProdVer14(0,2)="14.0.6029.1000" : arrProdVer14(0,3)="14.0.7015.1000"
    arrProdVer14(1,0)="007A" : arrProdVer14(1,1)="14.0.5118.5000,Web V1" 'Outlook Connector
    arrProdVer14(2,0)="007B" : arrProdVer14(2,1)="14.0.5117.5000,Web V1" 'Outlook Social Connector Provider for Windows Live Messen
    arrProdVer14(3,0)="007C" : arrProdVer14(3,1)="14.0.5117.5000,Web V1" 'Outlook Social Connector Facebook
    arrProdVer14(4,0)="007D" : arrProdVer14(4,1)="14.0.5120.5000,Web V1" 'Outlook Social Connector Windows Live
    arrProdVer14(5,0)="007F" : arrProdVer14(5,1)="14.0.5139.5001,Web V1" 'Outlook Connector
    arrProdVer14(6,0)="2005" : arrProdVer14(6,1)="14.0.5130.5003,Web V1" 'OFV File Validation Add-In
    
' 2007 Products -> KB 928516
    ReDim arrProdVer12(12,4) 'n,1=RTM; n,2=SP1
    arrProdVer12(0,0)="": arrProdVer12(0,1)="12.0.4518.1014": arrProdVer12(0,2)="12.0.6215.1000" : arrProdVer12(0,3)="12.0.6425.1000" : arrProdVer12(0,4)="12.0.6612.1000"
    arrProdVer12(1,0)="" : arrProdVer12(1,1)="12.0.4518.1016": arrProdVer12(1,2)="12.0.6219.1000" ': arrProdVer12(1,3)="12.0.6425.1000" 'Server
    arrProdVer12(2,0)="0045" : arrProdVer12(2,1)="12.0.4518.1084" 'Expression Web 2
    arrProdVer12(3,0)="011F" : arrProdVer12(3,1)="12.0.6407.1000,Web V1 (Windows Live)" 'Outlook Connector
    arrProdVer12(4,0)="001F" : arrProdVer12(4,1)="12.0.4518.1014" : arrProdVer12(4,2)="12.0.6213.1000" 'Office Proof (Container)
    arrProdVer12(5,0)="00B9" : arrProdVer12(5,1)="12.0.6012.5000,KB 932080": arrProdVer12(5,2)="12.0.6015.5000,Web V1 (Windows Live)"'Application Error Reporting
    arrProdVer12(6,0)="0020" : arrProdVer12(6,1)="12.0.4518.1014,Web V1" : arrProdVer12(6,3)="12.0.6021.5000,Web V3" : arrProdVer12(6,4)="12.0.6514.5001,Web V4 (incl. SP2)"  ': arrProdVer12(6,5)="12.0.6425.1000,SP2"'Compatibility Pack
    arrProdVer12(7,0)="00B0,00B1,00B2" : arrProdVer12(7,1)="12.0.4518.1014,Web V1" 'Save As packs
    arrProdVer12(8,0)="004A" : arrProdVer12(8,1)="12.0.4518.1014,Web RTM" : arrProdVer12(8,2)="12.0.6213.1000,Web SP1" 'Web Components
    arrProdVer12(9,0)="00A7" : arrProdVer12(9,1)="12.0.4518.1014,Web RTM" : arrProdVer12(9,2)="12.0.6520.3001,SP2" 'Calendar Printing Assistant
    arrProdVer12(10,0)="CFDA" : arrProdVer12(10,1)="12.0.613.1000,RTM" : arrProdVer12(10,2)="12.1.1313.1000,SP1"  : arrProdVer12(10,2)="12.0.6511.5000,SP2" 'Project Portfolio Server 2007
    arrProdVer12(11,0)="001C" : arrProdVer12(11,1)="12.0.4518.1049,RTM"  : arrProdVer12(11,2)="12.0.6230.1000,SP1" : arrProdVer12(11,3)="12.0.6237.1003,SP1" 'Access Runtime
    arrProdVer12(12,0)="003F" : arrProdVer12(12,1)="12.0.6214.1000,SP1" 'Excel Viewer

' 2003 Products -> KB 832672
    'ProductVersions      -> KB821549
    Redim arrProdVer11(16,4) 'n,1=RTM; n,2=SP1; n,3=SP2; n,4=SP3
    arrProdVer11(0,0)="": arrProdVer11(0,1)="11.0.5614.0": arrProdVer11(0,2)="11.0.6361.0": arrProdVer11(0,3)="11.0.7969.0": arrProdVer11(0,4)="11.0.8173.0" 'Suites & Core
    arrProdVer11(1,0)="14": arrProdVer11(1,1)="11.0.5614.0": arrProdVer11(1,2)="11.0.6361.0": arrProdVer11(1,3)="11.0.7969.0": arrProdVer11(1,4)="11.0.8173.0" 'Sharepoint Services
    arrProdVer11(2,0)="15,1C": arrProdVer11(2,1)="11.0.5614.0": arrProdVer11(2,2)="11.0.6355.0"': arrProdVer11(2,3)="11.0.7969.0": arrProdVer11(2,4)="11.0.8173.0" 'Access
    arrProdVer11(3,0)="16": arrProdVer11(3,1)="11.0.5612.0": arrProdVer11(3,2)="11.0.6355.0"': arrProdVer11(3,3)="11.0.7969.0": arrProdVer11(3,4)="11.0.8173.0" 'Excel
    arrProdVer11(4,0)="17": arrProdVer11(4,1)="11.0.5516.0": arrProdVer11(4,2)="11.0.6356.0"': arrProdVer11(4,3)="11.0.7969.0": arrProdVer11(4,4)="11.0.8173.0" 'FrontPage
    arrProdVer11(5,0)="18": arrProdVer11(5,1)="11.0.5529.0": arrProdVer11(5,2)="11.0.6361.0"': arrProdVer11(5,3)="11.0.7969.0": arrProdVer11(5,4)="11.0.8173.0" 'PowerPoint
    arrProdVer11(6,0)="19": arrProdVer11(6,1)="11.0.5525.0": arrProdVer11(6,2)="11.0.6255.0"': arrProdVer11(6,3)="11.0.7969.0": arrProdVer11(6,4)="11.0.8173.0" 'Publisher
    arrProdVer11(7,0)="1A,E0": arrProdVer11(7,1)="11.0.5510.0": arrProdVer11(7,2)="11.0.6353.0"': arrProdVer11(7,3)="11.0.7969.0": arrProdVer11(7,4)="11.0.8173.0" 'Outlook
    arrProdVer11(8,0)="1B": arrProdVer11(8,1)="11.0.5510.0": arrProdVer11(8,2)="11.0.6353.0"': arrProdVer11(8,3)="11.0.7969.0": arrProdVer11(8,4)="11.0.8173.0" 'Word
    arrProdVer11(9,0)="44": arrProdVer11(9,1)="11.0.5531.0": arrProdVer11(9,2)="11.0.6357.0"': arrProdVer11(9,3)="11.0.7969.0": arrProdVer11(9,4)="11.0.8173.0" 'InfoPath
    arrProdVer11(10,0)="A1": arrProdVer11(10,1)="11.0.5614.0": arrProdVer11(10,2)="11.0.6360.0"': arrProdVer11(10,3)="11.0.7969.0": arrProdVer11(10,4)="11.0.8173.0" 'OneNote
    arrProdVer11(11,0)="3A,3B,32": arrProdVer11(11,1)="11.0.5614.0": arrProdVer11(11,2)="11.0.6707.0"': arrProdVer11(11,3)="11.0.7969.0": arrProdVer11(11,4)="11.0.8173.0" 'Project
    arrProdVer11(12,0)="51,53,5E": arrProdVer11(12,1)="11.0.3216.5614": arrProdVer11(12,2)="11.0.4301.6360"': arrProdVer11(12,3)="11.0.7969.0": arrProdVer11(12,4)="11.0.8173.0" 'Visio
    arrProdVer11(13,0)="24" : arrProdVer11(13,1)="11.0.5614.0,V1 (Web)" : arrProdVer11(13,2)="11.0.6120.0,V2 (Hotfix)": arrProdVer11(13,3)="11.0.6550.0,V2 (Japanese Hotfix)" 'ORK
    arrProdVer11(14,0)="52": arrProdVer11(14,1)="11.0.3709.5614": arrProdVer11(14,2)="11.0.6206.8011,Web V1"': arrProdVer11(14,3)="11.0.7969.0": arrProdVer11(14,4)="11.0.8173.0" 'Visio Viewer
    arrProdVer11(15,0)="A4" : arrProdVer11(15,1)="11.0.5614.0": arrProdVer11(15,2)="11.0.6361.0" : arrProdVer11(15,3)="11.0.6558.0,MSDE"': arrProdVer11(15,4)="11.0.7969.0,SP2": arrProdVer11(15,5)="11.0.8173.0,SP3"'OWC11
    arrProdVer11(16,0)="A9" : arrProdVer11(16,1)="11.0.5614.0,RTM" : arrProdVer11(16,2)="11.0.6553.0,Web V1" 'PIA11

' Office XP
    'Office 10 Numbering Scheme -> KB302663
    'ProuctVersions             -> KB291331
    Redim arrProdVer10(5,4)
    arrProdVer10(0,0)="": arrProdVer10(0,1)="10.0.2627.01": arrProdVer10(0,2)="10.0.3520.0": arrProdVer10(0,3)="10.0.4330.0": arrProdVer10(0,4)="10.0.6626.0"
    'LPK/MUI RTM have inconsisten RTM build versions but do follow the main SP builds.
    '=> Limitation to not cover RTM LPK's
    arrProdVer10(1,0)="17,1D": arrProdVer10(1,1)="10.0.2623.0": arrProdVer10(1,2)="10.0.3506.0": arrProdVer10(1,3)="10.0.4128.0": arrProdVer10(1,4)="10.0.6308.0"'FrontPage
    arrProdVer10(2,0)="27,3A,3B": arrProdVer10(2,1)="10.0.2915.0": arrProdVer10(2,2)="10.0.3416.0": arrProdVer10(2,3)="10.0.4219.0": arrProdVer10(2,4)="10.0.6612.0"'Project
    arrProdVer10(3,0)="30": arrProdVer10(3,1)="10.0.2619.0" 'Media Content (CAG)
    arrProdVer10(4,0)="25,36": arrProdVer10(4,1)="10.0.2701.0,V1": arrProdVer10(4,2)="10.0.6403.0,V2 (Oct 24 2002 - Release)" 'ORK
    arrProdVer10(5,0)="51,52,53,54": arrProdVer10(5,1)="10.0.525": arrProdVer10(5,2)="10.1.2514" : arrProdVer10(5,3)="10.2.5110" 'Visio

' 2000
    Redim arrProdVer09(2,4)
    arrProdVer09(0,0)="": arrProdVer09(0,1)="9.00.2720": arrProdVer09(0,2)="9.00.3821,SR1": arrProdVer09(0,3)="9.00.4527": arrProdVer09(0,4)="9.00.9327"
    arrProdVer09(1,0)="3A,3B,3C": arrProdVer09(1,1)="9.00.2720": arrProdVer09(1,2)="9.00.4527"
    arrProdVer09(2,0)="5F":arrProdVer09(2,1)="9.00.00.2010,Web V1" 'ORK

    fInitArrProdVer = True
End Sub

'=======================================================================================================
'Module Patch 
'=======================================================================================================
Sub FindAllPatches
    If fBasicMode Then Exit Sub
    Dim n
    On Error Resume Next
    
    'Set Default for patch array
    Redim arrPatch(UBound(arrMaster),PATCH_COLUMNCOUNT,0)
    
    'Check which patch API calls are safe for use
    CheckPatchApi
    
    'Iterate all products to add the patches to the array 
    For n = 0 To UBound(arrMaster)
        FindPatches(n)
    Next 'n
    
    AddDetailsFromPackage

End Sub
'=======================================================================================================

'Find all patches registered to a product and add them to the Patch array
Sub FindPatches (iPosMaster)
    On Error Resume Next
    
    Dim Patches,Patch,siSumInfo,MsiDb,Record,qView
    Dim sQuery,sErr
    Dim iMspMax,iMspCnt
    Dim fVirtual

    fVirtual = (arrMaster(iPosMaster,COL_VIRTUALIZED) = 1)
    'Find client side patches (CSP) for the product
    If iWiVersionMajor > 2 Then
        If NOT fVirtual Then Set Patches = oMsi.PatchesEx(arrMaster(iPosMaster,COL_PRODUCTCODE),arrMaster(iPosMaster,COL_USERSID),arrMaster(iPosMaster,COL_CONTEXT),MSIPATCHSTATE_ALL)
        If Not Err = 0 OR fVirtual Then
            'PatchesEx API call failed
            'Log the error
            sErr = GetErrorDescription(Err)
            If fPatchesExOk AND NOT fVirtual Then
                'Only log the error if fPatchesExOk = True.
                'If we're Not fPatchesExOk the error is expected and has already been logged.
                Cachelog LOGPOS_REVITEM,LOGHEADING_NONE,ERR_CATEGORYERROR,"Product " & arrMaster(iPosMaster,COL_PRODUCTCODE) & DSV & arrMaster(iPosMaster,COL_PRODUCTNAME) & _
                     ": " & ERR_PATCHESEX
                arrMaster(iPosMaster,COL_ERROR) = arrMaster(iPosMaster,COL_ERROR) & ERR_CATEGORYERROR & ERR_PATCHESEX & DSV & sErr & CSV
            End If 'fPatchesExOk
            'Fall back to registry detection
            FindRegPatches iPosMaster
        Else
            'PatchesEx API call succeeded
            'Ensure sufficient array size
            iMspMax = Patches.Count
            If iMspMax > UBound(arrPatch,3) Then ReDim Preserve arrPatch(UBound(arrMaster),PATCH_COLUMNCOUNT,iMspMax)
            'Find the patch details
            ReadPatchDetails Patches,iPosMaster
        End If 'Err = 0
        
    Else
        FindRegPatches iPosMaster
    End If 'iWiVersionMajor > 2
    
    'Find patches in InstallSource (AIP)
    'Only valid for Office products
    Err.Clear
    If arrMaster(iPosMaster,COL_ISOFFICEPRODUCT) Then
        Set MsiDb = oMsi.OpenDatabase(arrMaster(iPosMaster,COL_CACHEDMSI), MSIOPENDATABASEMODE_READONLY)
        If Not Err = 0 Then    'Critical error for this sub. Need to terminate it to prevent endless loop.
            Err.Clear
            'Error has already been logged
            Exit Sub
        End If

        sQuery = "SELECT * FROM Property"
        Set qView = MsiDb.OpenView(sQuery)
        qView.Execute 
        Set Record = qView.Fetch()
        ' Loop through the records in the view
        Do Until (Record Is Nothing)
            'Get the properties name. The 'OR' condition covers the  Office 2000 family AIP's
            If (Left(Record.StringData(1),1) = "_" And Mid(Record.StringData(1),15,1) = "_") _
            OR (Left(Record.StringData(1),1) = "{" And Mid(Record.StringData(1),15,1) = "-") Then 
		        'Found AIP patch
		        'ReDim patch array
		        iMspCnt = UBound(arrPatch,3)+1
		        ReDim Preserve arrPatch(UBound(arrMaster),PATCH_COLUMNCOUNT,iMspCnt)
		        'Client side patch flag
		        arrPatch(iPosMaster,PATCH_CSP,iMspCnt) = False
	            'PatchCode
	            arrPatch(iPosMaster,PATCH_PATCHCODE,iMspCnt) = _
	            "{" & Mid(Record.StringData(1),2,8) & _
	            "-" & Mid(Record.StringData(1),11,4) & _
	            "-" & Mid(Record.StringData(1),16,4) & _
	            "-" & Mid(Record.StringData(1),21,4) & _
	            "-" & Mid(Record.StringData(1),26,12) &  "}"
	            'DisplayName
	            arrPatch(iPosMaster,PATCH_DISPLAYNAME,iMspCnt) = Record.StringData(2)
            End If 
            ' Next record
            Set Record = qView.Fetch()
        Loop
    End If 'arrMaster(iPosMaster,COL_ISOFFICEPRODUCT)

End Sub 'FindPatches
'=======================================================================================================

'Find all patches registered to a product by direct registry read and add them to the Patch array
Sub FindRegPatches (iPosMaster)
    Dim sSid,sSubKeyName,sSubKeyNamePatch,sSubKeyNamePatches,sValue,sProductCode
    Dim sPatch,sPatches,sAllPatches,sTmpPatches
    Dim hDefKey,hDefKeyGlobal
    Dim bPatchRegOk,bPatchMultiSzBroken,bPatchOrgMspNameBroken,bPatchCachedMspLinkBroken
    Dim bPatchGlobalOK,bPatchProdGlobalBroken,bPatchGlobalMultiSzBroken
    Dim iContext,iName,iType,iKey,iState,iMspMax
    Dim arrName,arrType,arrKeys,arrRegPatch
    Const REGDIM1 = 1
    On Error Resume Next
    
    iContext = arrMaster(iPosMaster,COL_CONTEXT)
    sSid = arrMaster(iPosMaster,COL_USERSID)
    sProductCode = arrMaster(iPosMaster,COL_PRODUCTCODE)
    sTmpPatches = ""
    bPatchRegOk = False
    bPatchGlobalOk = False
    bPatchMultiSzBroken = False
    bPatchOrgMspNameBroken = False
    ReDim arrRegPatch(REGDIM1,-1)
    
    'Find "Applied" patches
    hDefKeyGlobal = GetRegHive(iContext,sSid,True)
    hDefKey = GetRegHive(iContext,sSid,False)
    sSubKeyName = GetRegConfigKey(sProductCode,iContext,sSid,False) & "Patches\"
    If RegEnumValues(hDefKey,sSubKeyName,arrName,arrType) AND CheckArray(arrName) Then
        bPatchRegOk = True
        'Check Metadata integrity
        sPatches = vbNullString
        If Not RegReadValue(hDefKey,sSubKeyName,"Patches",sPatches,REG_MULTI_SZ) Then bPatchRegOk = False
        sTmpPatches = sPatches
        For iName = 0 To UBound(arrName)
            If arrType(iName) = REG_SZ Then
                ReDim Preserve arrRegPatch(REGDIM1,UBound(arrRegPatch,2)+1)
                sValue = ""
                If RegReadStringValue(hDefKey,sSubKeyName,arrName(iName),sValue) Then sPatch = arrName(iName)
                arrRegPatch(0,UBound(arrRegPatch,2)) = GetExpandedGuid(sPatch)
                arrRegPatch(1,UBound(arrRegPatch,2)) = MSIPATCHSTATE_APPLIED
                If NOT InStr(sTmpPatches,sPatch)>0 Then
                    bPatchRegOk = False
                    bPatchMultiSzBroken = True
                Else
                    'Strip current patch from sTmpPatches
                    sTmpPatches = Replace(sTmpPatches,sPatch,"")
                End If 'InStr(sTmpPatches,sPatch)
                
                'Check Patches key
                sSubKeyNamePatches = GetRegConfigPatchesKey(iContext,sSid,False)
                If Not RegKeyExists(hDefKey,sSubKeyNamePatches & sPatch) Then
                    bPatchOrgMspNameBroken = True
                End If
                
                'Check Global Patches key
                sSubKeyNamePatches = ""
                sSubKeyNamePatches = GetRegConfigPatchesKey(iContext,sSid,True)
                If Not RegKeyExists(hDefKeyGlobal,sSubKeyNamePatches & sPatch) Then
                    bPatchCachedMspLinkBroken = True
                End If
                
                'Check Global Product key 
                sSubKeyNamePatches = ""
                sSubKeyNamePatches = GetRegConfigKey(sProductCode,iContext,sSid,True) & "Patches\" & sPatch
                If Not RegKeyExists(hDefKeyGlobal,sSubKeyNamePatches) Then
                    bPatchProdGlobalBroken = True
                Else
                    If RegReadDWordValue(hDefKeyGlobal,sSubKeyNamePatches,"State",sValue) Then 
                        arrRegPatch(1,UBound(arrRegPatch,2)) = CInt(sValue)
                    End If 'RegReadDWordValue
                End If 'RegKeyExists(hDefKeyGlobal,sSubKeyNamePatches)
            End If 'arrType
        Next 'iName
    End If 'RegEnumValues
    'sTmpPatches should be an empty string now
    sTmpPatches = Replace(sTmpPatches,Chr(34),"")
    If (Not sTmpPatches = "") OR (NOT fPatchesOk) Then arrMaster(iPosMaster,COL_ERROR) = arrMaster(iPosMaster,COL_ERROR) &  ERR_BADMSPMETADATA & CSV

    'Find patches in other states (than 'Applied')
    '---------------------------------------------
    'Get a list from the global patches key
    sSubKeyName = GetRegConfigKey(sProductCode,iContext,sSid,True) & "Patches\"
    bPatchGlobalMultiSzBroken = NOT RegReadValue(hDefKeyGlobal,sSubKeyName,"AllPatches",sAllPatches,REG_MULTI_SZ)
    
    'Sanity Check on metadata integrity
    If RegEnumKey(hDefKeyGlobal,sSubKeyName,arrKeys) AND CheckArray(arrKeys) Then
        bPatchGlobalOk = True
        sTmpPatches = sAllPatches
        For iKey = 0 To UBound(arrKeys)
            If InStr(sTmpPatches,arrKeys(iKey))>0 Then
                sTmpPatches = Replace(sTmpPatches,arrKeys(iKey),"")
            End If 'InStr
        Next 'iKey
        'sTmpPatches should be an empty string now
        sTmpPatches = Replace(sTmpPatches,Chr(34),"")
        
        'Add the patches to the array
        For iKey = 0 To UBound(arrKeys)
            If Not InStr(sPatches,arrKeys(iKey))>0 Then
                ReDim Preserve arrRegPatch(REGDIM1,UBound(arrRegPatch,2)+1)
                sValue = ""
                If RegReadDWordValue(hDefKeyGlobal,sSubKeyName&arrKeys(iKey),"State",sValue) Then iState = CInt(sValue) Else iState = -1
                arrRegPatch(0,UBound(arrRegPatch,2)) = GetExpandedGuid(arrKeys(iKey))
                arrRegPatch(1,UBound(arrRegPatch,2)) = iState
            End If 'Not InStr
        Next 'iKey
    End If 'RegEnumKey
    
    'Ensure sufficient array size
    iMspMax = 0 : If CheckArray(arrRegPatch) Then iMspMax = UBound(arrRegPatch,2)
    If iMspMax > UBound(arrPatch,3) Then ReDim Preserve arrPatch(UBound(arrMaster),PATCH_COLUMNCOUNT,iMspMax)
    'Find the patch details
    ReadRegPatchDetails arrRegPatch,iPosMaster
    
End Sub 'FindRegPatches
'=======================================================================================================

Sub ReadPatchDetails(Patches,iPosMaster)
    Dim sTmpState
    On Error Resume Next
    
    Dim Patch,siSumInfo
    Dim iMspCnt
    
    iMspCnt = 0
    For Each Patch in Patches
        'Add patch data to array
        'ProductCode:
        arrPatch(iPosMaster,PATCH_PRODUCT,iMspCnt) = arrMaster(iPosMaster,COL_PRODUCTCODE)
        'PatchCode:
        arrPatch(iPosMaster,PATCH_PATCHCODE,iMspCnt) = Patch.PatchCode
        'CSP flag:
        arrPatch(iPosMaster,PATCH_CSP,iMspCnt) = True
        'LocalPackage:
        arrPatch(iPosMaster,PATCH_LOCALPACKAGE,iMspCnt) = Patch.PatchProperty("LocalPackage")
        'LocalPackage availability:
        arrPatch(iPosMaster,PATCH_CPOK,iMspCnt) = False
        arrPatch(iPosMaster,PATCH_CPOK,iMspCnt) = oFso.FileExists(arrPatch(iPosMaster,PATCH_LOCALPACKAGE,iMspCnt))
        'DisplayName:
        arrPatch(iPosMaster,PATCH_DISPLAYNAME,iMspCnt) = Patch.PatchProperty("DisplayName")
        'Ensure displayname is not blank
        If IsEmpty(arrPatch(iPosMaster,PATCH_DISPLAYNAME,iMspCnt)) OR (arrPatch(iPosMaster,PATCH_DISPLAYNAME,iMspCnt)=vbNullString) Then
            'Fall back to registry mode detection
            arrPatch(iPosMaster,PATCH_DISPLAYNAME,iMspCnt) = GetRegMspProperty(GetCompressedGuid(Patch.PatchCode),iMspCnt,iPosMaster,arrMaster(iPosMaster,COL_CONTEXT),arrMaster(iPosMaster,COL_USERSID),"DisplayName")
        End If 'IsEmpty(iPosMaster,PATCH_DISPLAYNAME,iMspCnt)
        
        'InstallState:
        Select Case (Patch.State)
	        Case MSIPATCHSTATE_APPLIED,MSIPATCHSTATE_SUPERSEDED
	            If Patch.State = MSIPATCHSTATE_APPLIED Then sTmpState = "Applied" Else sTmpState = "Superseded"
	            arrPatch(iPosMaster,PATCH_PATCHSTATE,iMspCnt) = sTmpState
	            'Check stuff that is only useful/valid for patches in Applied/Superseded state
	            'Uninstallable
	            If Patch.PatchProperty("Uninstallable") = "1" Then
		            arrPatch(iPosMaster,PATCH_UNINSTALLABLE,iMspCnt) = "Yes"
	            Else
		            arrPatch(iPosMaster,PATCH_UNINSTALLABLE,iMspCnt) = "No"
	            End If
	            'Check cached package
	            If Not arrPatch(iPosMaster,PATCH_CPOK,iMspCnt) Then 
	                arrMaster(iPosMaster,COL_ERROR) = arrMaster(iPosMaster,COL_ERROR) & _
                     ERR_CATEGORYERROR & "Local patch package '" & arrPatch(iPosMaster,PATCH_LOCALPACKAGE,iMspCnt) & _
	                 "' missing for patch '." & arrPatch(iPosMaster,PATCH_PATCHCODE,iMspCnt) & "'" & CSV
                    Cachelog LOGPOS_REVITEM,LOGHEADING_NONE,ERR_CATEGORYERROR,"Local patch package '" & arrPatch(iPosMaster,PATCH_LOCALPACKAGE,iMspCnt) & _
	                 "' missing for patch '" & arrPatch(iPosMaster,PATCH_PATCHCODE,iMspCnt) & "'"
                End If
			    'InstallDate
			    arrPatch(iPosMaster,PATCH_INSTALLDATE,iMspCnt) = Patch.PatchProperty("InstallDate")
			    'MoreInfoUrl
			    arrPatch(iPosMaster,PATCH_MOREINFOURL,iMspCnt) = Patch.PatchProperty("MoreInfoURL")

	        Case MSIPATCHSTATE_OBSOLETED
	            arrPatch(iPosMaster,PATCH_PATCHSTATE,iMspCnt) = "Obsoleted"
			    'InstallDate
			    arrPatch(iPosMaster,PATCH_INSTALLDATE,iMspCnt) = Patch.PatchProperty("InstallDate")
			    'MoreInfoUrl
			    arrPatch(iPosMaster,PATCH_MOREINFOURL,iMspCnt) = Patch.PatchProperty("MoreInfoURL")

	        Case MSIPATCHSTATE_REGISTERED
	            arrPatch(iPosMaster,PATCH_PATCHSTATE,iMspCnt) = "Registered"
			    'MoreInfoUrl
			    arrPatch(iPosMaster,PATCH_MOREINFOURL,iMspCnt) = Patch.PatchProperty("MoreInfoURL")

	        Case Else
	            arrPatch(iPosMaster,PATCH_PATCHSTATE,iMspCnt) = "Unknown Patchstate"
        End Select
        'PatchTransform
        arrPatch(iPosMaster,PATCH_TRANSFORM,iMspCnt) = GetRegMspProperty(GetCompressedGuid(Patch.PatchCode),iMspCnt,iPosMaster,arrMaster(iPosMaster,COL_CONTEXT),arrMaster(iPosMaster,COL_USERSID),"PatchTransform")
        iMspCnt = iMspCnt + 1
    Next 'Patch
End Sub 'ReadPatchDetails
'=======================================================================================================

Sub ReadRegPatchDetails(arrRegPatch,iPosMaster)
    Dim hDefKey
    Dim sSid,sSubKeyName,sSubKeyNamePatch,sValue,sProductCode,sPatchCodeCompressed,sTmpState
    Dim iContext
    On Error Resume Next
    
    Dim Patch,siSumInfo
    Dim iMspCnt
    
    iContext = arrMaster(iPosMaster,COL_CONTEXT)
    sSid = arrMaster(iPosMaster,COL_USERSID)
    sProductCode = arrMaster(iPosMaster,COL_PRODUCTCODE)
    hDefKey = GetRegHive(iContext,sSid,True)
    sSubKeyName = GetRegConfigKey(sProductCode,iContext,sSid,True) & "Patches\"
    'iMspCnt = 0
    'For Each Patch in Patches
    For iMspCnt = 0 To UBound(arrRegPatch,2)
        sPatchCodeCompressed = GetCompressedGuid(arrRegPatch(0,iMspCnt))
        'Add patch data to array
        'ProductCode:
        arrPatch(iPosMaster,PATCH_PRODUCT,iMspCnt) = arrMaster(iPosMaster,COL_PRODUCTCODE)
        'PatchCode:
        arrPatch(iPosMaster,PATCH_PATCHCODE,iMspCnt) = arrRegPatch(0,iMspCnt)
        'CSP flag:
        arrPatch(iPosMaster,PATCH_CSP,iMspCnt) = True
        'LocalPackage:
        arrPatch(iPosMaster,PATCH_LOCALPACKAGE,iMspCnt) = GetRegMspProperty(sPatchCodeCompressed,iMspCnt,iPosMaster,iContext,sSid,"LocalPackage")
        arrPatch(iPosMaster,PATCH_CPOK,iMspCnt) = oFso.FileExists(arrPatch(iPosMaster,PATCH_LOCALPACKAGE,iMspCnt))
        'DisplayName:
        arrPatch(iPosMaster,PATCH_DISPLAYNAME,iMspCnt) = GetRegMspProperty(sPatchCodeCompressed,iMspCnt,iPosMaster,iContext,sSid,"DisplayName")
        'InstallState:
        Select Case (arrRegPatch(1,iMspCnt))
	        Case MSIPATCHSTATE_APPLIED,MSIPATCHSTATE_SUPERSEDED
	            If arrRegPatch(1,iMspCnt) = MSIPATCHSTATE_APPLIED Then sTmpState = "Applied" Else sTmpState = "Superseded"
	            arrPatch(iPosMaster,PATCH_PATCHSTATE,iMspCnt) = sTmpState
	            'Check stuff that is only useful/valid for patches in 'Applied/Superseded' state
	            'Uninstallable
	            arrPatch(iPosMaster,PATCH_UNINSTALLABLE,iMspCnt) = GetRegMspProperty(sPatchCodeCompressed,iMspCnt,iPosMaster,iContext,sSid,"Uninstallable")
	            'Check cached package
	            If Not arrPatch(iPosMaster,PATCH_CPOK,iMspCnt) Then 
	                arrMaster(iPosMaster,COL_ERROR) = arrMaster(iPosMaster,COL_ERROR) & _
                     ERR_CATEGORYERROR & "Local patch package '" & arrPatch(iPosMaster,PATCH_LOCALPACKAGE,iMspCnt) & _
	                 "' missing for patch '." & arrPatch(iPosMaster,PATCH_PATCHCODE,iMspCnt) & "'" & CSV
                    Cachelog LOGPOS_REVITEM,LOGHEADING_NONE,ERR_CATEGORYERROR,"Local patch package '" & arrPatch(iPosMaster,PATCH_LOCALPACKAGE,iMspCnt) & _
	                 "' missing for patch '" & arrPatch(iPosMaster,PATCH_PATCHCODE,iMspCnt) & "'" 
                End If
			    'InstallDate
			    arrPatch(iPosMaster,PATCH_INSTALLDATE,iMspCnt) = GetRegMspProperty(sPatchCodeCompressed,iMspCnt,iPosMaster,iContext,sSid,"InstallDate")
			    'MoreInfoUrl
			    arrPatch(iPosMaster,PATCH_MOREINFOURL,iMspCnt) = GetRegMspProperty(sPatchCodeCompressed,iMspCnt,iPosMaster,iContext,sSid,"MoreInfoURL")
	        Case MSIPATCHSTATE_OBSOLETED
	            arrPatch(iPosMaster,PATCH_PATCHSTATE,iMspCnt) = "Obsoleted"
			    'InstallDate
			    arrPatch(iPosMaster,PATCH_INSTALLDATE,iMspCnt) = GetRegMspProperty(sPatchCodeCompressed,iMspCnt,iPosMaster,iContext,sSid,"InstallDate")
			    'MoreInfoUrl
			    arrPatch(iPosMaster,PATCH_MOREINFOURL,iMspCnt) = GetRegMspProperty(sPatchCodeCompressed,iMspCnt,iPosMaster,iContext,sSid,"MoreInfoURL")
	        Case MSIPATCHSTATE_REGISTERED
	            arrPatch(iPosMaster,PATCH_PATCHSTATE,iMspCnt) = "Registered"
			    'MoreInfoUrl
			    arrPatch(iPosMaster,PATCH_MOREINFOURL,iMspCnt) = GetRegMspProperty(sPatchCodeCompressed,iMspCnt,iPosMaster,iContext,sSid,"MoreInfoURL")
	        Case Else
	            arrPatch(iPosMaster,PATCH_PATCHSTATE,iMspCnt) = "Unknown Patchstate"
        End Select
        'PatchTransform
        arrPatch(iPosMaster,PATCH_TRANSFORM,iMspCnt) = GetRegMspProperty(sPatchCodeCompressed,iMspCnt,iPosMaster,iContext,sSid,"PatchTransform")
    Next 'iMspCnt
End Sub 'ReadRegPatchDetails
'=======================================================================================================

'Get a patch property without MSI API
Function GetRegMspProperty(sPatchCodeCompressed,iMspCnt,iPosMaster,iContext,sSid,sProperty)
    Dim hDefKey
    Dim sProductCode,sSubKeyName,sValue
    Dim siSumInfo
    On Error Resume Next
    
    GetRegMspProperty = ""
    sProductCode = arrMaster(iPosMaster,COL_PRODUCTCODE)
    'Default to 'Global' registry patch location
    hDefKey = GetRegHive(iContext,sSid,True)
    sSubKeyName = GetRegConfigKey(sProductCode,iContext,sSid,True) & "Patches\" & sPatchCodeCompressed & "\"
    sValue = ""
    Select Case UCase(sProperty)
    Case "DISPLAYNAME"
        If RegReadStringValue(hDefKey,sSubKeyName,"DisplayName",sValue) Then GetRegMspProperty = sValue
        'Ensure displayname is not blank
        If GetRegMspProperty="" Then
            'Try to get the value from SummaryInformation stream. Use the PatchCode in case of failure
            If arrPatch(iPosMaster,PATCH_CPOK,iMspCnt) Then
                Err.Clear
                Set siSumInfo = oMsi.SummaryInformation(arrPatch(iPosMaster,PATCH_LOCALPACKAGE,iMspCnt),MSIOPENDATABASEMODE_READONLY)
                If Not Err = 0 Then
                    GetRegMspProperty = GetExpandedGuid(sPatchCodeCompressed)
                Else
                    GetRegMspProperty = siSumInfo.Property(PID_TITLE)
                    If GetRegMspProperty = vbNullString Then GetRegMspProperty = GetExpandedGuid(sPatchCodeCompressed)
                End If 'Err = 0
            Else
                GetRegMspProperty = GetExpandedGuid(sPatchCodeCompressed)
            End If 'arrPatch(iPosMaster,PATCH_CPOK,iMspCnt)
        End If 'GetRegMspProperty=""
    Case "INSTALLDATE"
        If RegReadStringValue(hDefKey,sSubKeyName,"Installed",sValue) Then GetRegMspProperty = sValue
    Case "LOCALPACKAGE"
        If iContext = MSIINSTALLCONTEXT_USERMANAGED Then iContext = MSIINSTALLCONTEXT_USERUNMANAGED
        sSubKeyName = GetRegConfigPatchesKey(iContext,sSid,True) & sPatchCodeCompressed & "\"
        If RegReadStringValue(hDefKey,sSubKeyName,"LocalPackage",sValue) Then GetRegMspProperty = sValue
    Case "MOREINFOURL"
        If RegReadStringValue(hDefKey,sSubKeyName,"MoreInfoURL",sValue) Then GetRegMspProperty = sValue
    Case "PATCHTRANSFORM"
        hDefKey = GetRegHive(iContext,sSid,False)
        sSubKeyName = GetRegConfigKey(sProductCode,iContext,sSid,False) & "Patches\" 
        If RegReadStringValue(hDefKey,sSubKeyName,sPatchCodeCompressed,sValue) Then GetRegMspProperty = sValue
    Case "UNINSTALLABLE"
        GetRegMspProperty = "No"
        If (RegReadDWordValue(hDefKey,sSubKeyName,"Uninstallable",sValue) AND sValue = "1") Then GetRegMspProperty = "Yes"
    Case Else
    End Select
End Function 'GetRegMspDisplayName
'=======================================================================================================

'Check if usage of PatchesEx API calls are safe for use.
'PatchesEx requires WI 3.x or higher
'If PatchesEx API calls fail this triggers a fallback to direct registry detection mode
Sub CheckPatchApi
    Dim sActiveSub, sErrHnd 
    sActiveSub = "CheckPatchApi" : sErrHnd = "" 
    On Error Resume Next
    Dim MspCheckEx, MspCheck
        
    'Defaults
    fPatchesExOk = False
    fPatchesOk = False
    iPatchesExError = 0
    'WI 2.x specific check
    Set MspCheck = oMsi.Patches(arrMaster(0,COL_PRODUCTCODE))
    If Err = 0 Then fPatchesOk = True Else Err.Clear
    'Ensure WI 3.x for PatchesEx checks
    If iWiVersionMajor = 2 Then 
        Exit Sub
    End If 'iWiVersionMajor = 2
    'ROIScan error values for PatchesEx to allow to distinguish into unique error messages
    Set MspCheckEx = oMsi.PatchesEx(PRODUCTCODE_EMPTY,USERSID_EVERYONE,MSIINSTALLCONTEXT_ALL,MSIPATCHSTATE_ALL)
    If Err <> 0 Then
        iPatchesExError = 8
        CheckError sActiveSub,sErrHnd 
    Else
        fPatchesExOk = True
        Exit Sub
    End If 'Not Err <> 0
        
    Set MspCheckEx = oMsi.PatchesEx(PRODUCTCODE_EMPTY,USERSID_EVERYONE,MSIINSTALLCONTEXT_USERMANAGED,MSIPATCHSTATE_ALL)
    If Err <> 0 Then
        iPatchesExError = iPatchesExError + MSIINSTALLCONTEXT_USERMANAGED
        CheckError sActiveSub,sErrHnd 
    End If 'Not Err <> 0 - MSIINSTALLCONTEXT_USERMANAGED
    
    Set MspCheckEx = oMsi.PatchesEx(PRODUCTCODE_EMPTY,USERSID_EVERYONE,MSIINSTALLCONTEXT_USERUNMANAGED,MSIPATCHSTATE_ALL)
    If Err <> 0 Then
        iPatchesExError = iPatchesExError + MSIINSTALLCONTEXT_USERUNMANAGED
        CheckError sActiveSub,sErrHnd 
    End If 'Not Err <> 0 - MSIINSTALLCONTEXT_USERUNMANAGED
    
    Set MspCheckEx = oMsi.PatchesEx(PRODUCTCODE_EMPTY,MACHINESID,MSIINSTALLCONTEXT_MACHINE,MSIPATCHSTATE_ALL)
    If Err <> 0 Then
        iPatchesExError = iPatchesExError + MSIINSTALLCONTEXT_MACHINE
        CheckError sActiveSub,sErrHnd
    End If 'Not Err <> 0 - MSIINSTALLCONTEXT_MACHINE
    
    Select Case iPatchesExError
        Case 8  CacheLog LOGPOS_REVITEM,LOGHEADING_NONE,ERR_CATEGORYERROR,ERR_PATCHESEX & "MSIINSTALLCONTEXT_ALL" & DOT
        Case 9  CacheLog LOGPOS_REVITEM,LOGHEADING_NONE,ERR_CATEGORYERROR,ERR_PATCHESEX & "MSIINSTALLCONTEXT_ALL & MSIINSTALLCONTEXT_USERMANAGED" & DOT
        Case 10 CacheLog LOGPOS_REVITEM,LOGHEADING_NONE,ERR_CATEGORYERROR,ERR_PATCHESEX & "MSIINSTALLCONTEXT_ALL & MSIINSTALLCONTEXT_USERUNMANAGED" & DOT
        Case 11 CacheLog LOGPOS_REVITEM,LOGHEADING_NONE,ERR_CATEGORYERROR,ERR_PATCHESEX & "MSIINSTALLCONTEXT_ALL & MSIINSTALLCONTEXT_USERMANAGED & MSIINSTALLCONTEXT_USERUNMANAGED" & DOT
        Case 12 CacheLog LOGPOS_REVITEM,LOGHEADING_NONE,ERR_CATEGORYERROR,ERR_PATCHESEX & "MSIINSTALLCONTEXT_ALL & MSIINSTALLCONTEXT_MACHINE" & DOT
        Case 13 CacheLog LOGPOS_REVITEM,LOGHEADING_NONE,ERR_CATEGORYERROR,ERR_PATCHESEX & "MSIINSTALLCONTEXT_ALL & MSIINSTALLCONTEXT_USERMANAGED & MSIINSTALLCONTEXT_MACHINE" & DOT
        Case 14 CacheLog LOGPOS_REVITEM,LOGHEADING_NONE,ERR_CATEGORYERROR,ERR_PATCHESEX & "MSIINSTALLCONTEXT_ALL & MSIINSTALLCONTEXT_USERUNMANAGED & MSIINSTALLCONTEXT_MACHINE" & DOT
        Case 15 CacheLog LOGPOS_REVITEM,LOGHEADING_NONE,ERR_CATEGORYERROR,ERR_PATCHESEX & "MSIINSTALLCONTEXT_ALL & MSIINSTALLCONTEXT_USERMANAGED & MSIINSTALLCONTEXT_USERUNMANAGED & MSIINSTALLCONTEXT_MACHINE" & DOT
        Case Else 
    End Select
End Sub 'CheckPatchApi
'=======================================================================================================

'Gather details from the cached patch package
Sub AddDetailsFromPackage
    On Error Resume Next

    Dim Element, Elements, Key, XmlDoc, Msp, SumInfo, Record
    Dim dicMspSequence, dicMspTmp
    Dim sFamily, sSeq, sMsp, sBaseSeqShort
    Dim iPosMaster, iPosPatch, iCnt, iIndex, iVersionMajor
    Dim fAdd
    Dim qView
    Dim arrTitle, arrTmpFam, arrTmpSeq, arrVer, arrPatchTargets
        
    'Defaults
    Set XmlDoc = CreateObject("Microsoft.XMLDOM")
    Set dicMspSequence = CreateObject("Scripting.Dictionary")
    Set dicMspTmp = CreateObject("Scripting.Dictionary")
    'Initialize the patch index dictionary to assign each distinct PatchCode an index number
    iCnt = -1
    For iPosMaster = 0 To UBound(arrMaster)
        For iPosPatch = 0 to UBound(arrPatch, 3)
            If Not (IsEmpty (arrPatch(iPosMaster, PATCH_PATCHCODE, iPosPatch))) AND (arrPatch(iPosMaster, PATCH_CSP, iPosPatch)) Then
                If NOT dicMspIndex.Exists(arrPatch(iPosMaster, PATCH_PATCHCODE, iPosPatch)) Then
                    iCnt = iCnt + 1
                    dicMspIndex.Add arrPatch(iPosMaster, PATCH_PATCHCODE, iPosPatch), iCnt
                    dicMspTmp.Add arrPatch(iPosMaster, PATCH_PATCHCODE, iPosPatch), arrPatch(iPosMaster, PATCH_LOCALPACKAGE, iPosPatch)
                End If
            End If
        Next 'iPosPatch
    Next 'iPosMaster
    'Initialize the PatchFiles array
    ReDim arrMspFiles(dicMspIndex.Count -1, MSPFILES_COLUMNCOUNT -1)
    'Open the .msp for reading to add the data to the array
    For Each Key in dicMspTmp.Keys
        iPosPatch = dicMspIndex.Item(Key)
        sMsp = dicMspTmp.Item(Key)
        Err.Clear
        Set Msp = oMsi.OpenDatabase (sMsp, MSIOPENDATABASEMODE_PATCHFILE)
        Set SumInfo = Msp.SummaryInformation
        If Not Err = 0 Then
            'An error at this points indicates a severe issue
            Cachelog LOGPOS_REVITEM, LOGHEADING_NONE, ERR_CATEGORYERROR, ERR_MSPOPENFAILED & sMsp
        Else
            'LocalPackage
            arrMspFiles(iPosPatch, MSPFILES_LOCALPACKAGE) = sMsp
            'PatchCode
            arrMspFiles(iPosPatch, MSPFILES_PATCHCODE) = Key
            'PatchTargets
            arrMspFiles(iPosPatch, MSPFILES_TARGETS) = SumInfo.Property(PID_TEMPLATE)
            'PatchTables
            arrMspFiles(iPosPatch, MSPFILES_TABLES) = GetPatchTables(Msp)
            'KB
            If InStr(arrMspFiles(iPosPatch, MSPFILES_TABLES), "MsiPatchMetadata") > 0 Then
                Set qView = Msp.OpenView("SELECT `Property`,`Value` FROM MsiPatchMetadata WHERE `Property`='KBArticle Number'")
                qView.Execute : Set Record = qView.Fetch()
                If Not Record Is Nothing Then
                    arrMspFiles(iPosPatch, MSPFILES_KB) = UCase(Record.StringData(2))
                    arrMspFiles(iPosPatch, MSPFILES_KB) = Replace(arrMspFiles(iPosPatch, MSPFILES_KB), "KB", "")
                Else
                    arrMspFiles(iPosPatch, MSPFILES_KB) = ""
                End If
                qView.Close
            'StdPackageName
                Set qView = Msp.OpenView("SELECT `Property`,`Value` FROM MsiPatchMetadata WHERE `Property`='StdPackageName'")
                qView.Execute : Set Record = qView.Fetch()
                If Not Record Is Nothing Then
                    arrMspFiles(iPosPatch, MSPFILES_PACKAGE) = Record.StringData(2)
                Else
                    arrMspFiles(iPosPatch, MSPFILES_PACKAGE) = ""
                End If
                qView.Close
            Else
                arrMspFiles(iPosPatch, MSPFILES_KB) = ""
                arrMspFiles(iPosPatch, MSPFILES_PACKAGE) = ""
            End If
            arrTitle = Split(SumInfo.Property(PID_TITLE), ";")
            If arrMspFiles(iPosPatch,MSPFILES_KB) = "" Then
                If UBound(arrTitle) > 0 Then
                    arrMspFiles(iPosPatch, MSPFILES_KB) = arrTitle(1)
                End If
            End If
            If arrMspFiles(iPosPatch, MSPFILES_PACKAGE) = "" Then
                If UBound(arrTitle) > 0 Then
                    arrMspFiles(iPosPatch, MSPFILES_PACKAGE) = arrTitle(1)
                End If
            End If
            'PatchSequence & PatchFamily
            If InStr(arrMspFiles(iPosPatch, MSPFILES_TABLES), "MsiPatchSequence") > 0 Then
                Set qView = Msp.OpenView("SELECT `PatchFamily`,`Sequence`,`Attributes` FROM MsiPatchSequence")
                qView.Execute : Set Record = qView.Fetch()
                If Not Record Is Nothing Then
                    Do Until Record Is Nothing
                        arrMspFiles(iPosPatch, MSPFILES_FAMILY) = arrMspFiles(iPosPatch, MSPFILES_FAMILY) & LCase(Record.StringData(1)) & ","
                        sSeq = Record.StringData(2)
                        If NOT InStr(sSeq, ".") > 0 Then sSeq = GetVersionMajor(Left(arrMspFiles(iPosPatch, MSPFILES_TARGETS), 38)) & ".0." & sSeq & ".0"
                        arrMspFiles(iPosPatch, MSPFILES_SEQUENCE) = arrMspFiles(iPosPatch, MSPFILES_SEQUENCE) & sSeq & ","
                        arrMspFiles(iPosPatch, MSPFILES_ATTRIBUTE) = Record.StringData(3)
                        Set Record = qView.Fetch()
                    Loop
                    arrMspFiles(iPosPatch, MSPFILES_FAMILY) = RTrimComma(arrMspFiles(iPosPatch, MSPFILES_FAMILY))
                    arrMspFiles(iPosPatch, MSPFILES_SEQUENCE) = RTrimComma(arrMspFiles(iPosPatch, MSPFILES_SEQUENCE))
                Else
                    arrMspFiles(iPosPatch, MSPFILES_FAMILY) = ""
                    arrMspFiles(iPosPatch, MSPFILES_SEQUENCE) = "0"
                    arrMspFiles(iPosPatch, MSPFILES_ATTRIBUTE) = "0"
                End If
                qView.Close
            Else
                arrMspFiles(iPosPatch, MSPFILES_FAMILY) = ""
                arrMspFiles(iPosPatch, MSPFILES_SEQUENCE) = "0"
                arrMspFiles(iPosPatch, MSPFILES_ATTRIBUTE) = "0"
            End If
        End If
    Next 'Key

    'Copy the collected data to the arrPatch array
    For iPosMaster = 0 To UBound(arrMaster)
        dicMspSequence.RemoveAll
        sBaseSeqShort = ""
        arrVer = Split(arrMaster(iPosMaster,COL_PRODUCTVERSION),".")
        If UBound(arrVer)>1 Then sBaseSeqShort = arrVer(2)
        For iPosPatch = 0 to UBound(arrPatch,3)
            If Not (IsEmpty (arrPatch(iPosMaster,PATCH_PATCHCODE,iPosPatch))) AND (arrPatch(iPosMaster,PATCH_CSP,iPosPatch)) Then
                iIndex = -1
                iIndex = dicMspIndex.Item(arrPatch(iPosMaster,PATCH_PATCHCODE,iPosPatch))
                If NOT iIndex = -1 Then
                    'KB field
                    arrPatch(iPosMaster,PATCH_KB,iPosPatch) = arrMspFiles(iIndex,MSPFILES_KB)
                    'StdPackageName
                    arrPatch(iPosMaster,PATCH_PACKAGE,iPosPatch) = arrMspFiles(iIndex,MSPFILES_PACKAGE)
                    'PatchSequenece
                    Set arrTmpSeq = Nothing
                    arrTmpSeq = Split(arrMspFiles(iIndex,MSPFILES_SEQUENCE),",")
                    arrPatch(iPosMaster,PATCH_SEQUENCE,iPosPatch) = arrTmpSeq(0)
                    'PatchFamily
                    Set arrTmpFam = Nothing
                    arrTmpFam = Split(arrMspFiles(iIndex,MSPFILES_FAMILY),",")
                    'dicMspSequence
                    'Only add to the family sequence number if it's marked to supersede earlier
                    If arrMspFiles(iIndex,MSPFILES_ATTRIBUTE) = "1" Then
                        iCnt = -1
                        For iCnt = 0 To UBound(arrTmpSeq)
                            sFamily="" : sSeq=""
                            sFamily = arrTmpFam(iCnt)
                            sSeq = arrTmpSeq(iCnt)
                            fAdd = False
                            'Don't care if it's not higher than the current baseline
                            Select Case Len(sSeq)
                            Case 4
                                fAdd = (sSeq > sBaseSeqShort)
                            Case Else
                                fAdd = (sSeq > arrMaster(iPosMaster, COL_PRODUCTVERSION))
                            End Select
                            If fAdd Then 
                                If dicMspSequence.Exists(sFamily) Then
                                    If (sSeq > dicMspSequence.Item(sFamily)) Then dicMspSequence.Item(sFamily)=sSeq
                                Else
                                    dicMspSequence.Add sFamily,sSeq
                                End If
                            End If 'fAdd
                        Next 'iCnt
                    End If 'Attributes = 1
                    ' update the global PatchLevel dic
                    iCnt = -1
                    For iCnt = 0 To UBound(arrTmpSeq)
                        sFamily="" : sSeq=""
                        sFamily = arrTmpFam(iCnt)
                        sSeq = arrTmpSeq(iCnt)
                        If dicPatchLevel.Exists(sFamily) Then
                            If (sSeq > dicPatchLevel.Item(sFamily)) Then dicPatchLevel.Item(sFamily) = sSeq
                        Else
                            dicPatchLevel.Add sFamily, sSeq
                        End If
                    Next 'iCnt

                End If 'iIndex
            End If
        Next 'iPosPatch
        'Add the sequence data to the master array
        For Each Key in dicMspSequence.Keys
            arrMaster(iPosMaster,COL_PATCHFAMILY)=arrMaster(iPosMaster,COL_PATCHFAMILY)&Key&":"&dicMspSequence.Item(Key)&","
        Next 'Key
        RTrimComma arrMaster(iPosMaster,COL_PATCHFAMILY)
    Next 'iPosMaster
End Sub 'AddDetailsFromPackage
'=======================================================================================================

'Return the tables of a given .msp file
Function GetPatchTables(Msp)

Dim ViewTables,Table
Dim sTables

    On Error Resume Next
    sTables = ""
    Set Table = Nothing
    Set ViewTables = Msp.OpenView("SELECT `Name` FROM `_Tables` ORDER BY `Name`")
    ViewTables.Execute
    Do
        Set Table = ViewTables.Fetch
        If Table Is Nothing then Exit Do
        sTables = sTables&Table.StringData(1)&","
    Loop
    ViewTables.Close
    GetPatchTables=RTrimComma(sTables)
End Function 'GetPatchTables

'=======================================================================================================
'Module ARP - Add Remove Programs
'=======================================================================================================
'Build array with data from Add/Remove Programs.
'This is essential for Office 2007 family applications 
Sub ARPData
    On Error Resume Next

    'Filter for >O12 products since the logic is proprietary to >O12
     FindArpParents 
     AddArpParentsToMaster 
     AddSystemComponentFlagToMaster 
     ValidateSystemComponentProducts 
     FindArpConfigMsi 
End Sub 'ARPData
'=======================================================================================================

'Identify the core configuration .msi for multi .msi ARP entries
Sub FindArpConfigMsi
    Dim n, i, k
    Dim hDefKey
    Dim sSubKeyName, sCurKey, sName, sValue, sArpProd, sArpProdWW
    Dim arrKeys, arrValues, arrTmpArp
    Dim ProdId

    On Error Resume Next
    
    'Get Reg_Multi_Sz 'PackageIds' entry from 'uninstall' key
    'Do an InStr compare on each entry against ARP keyname
    'The identified position matches the position of the Reg_Multi_Sz 'ProductCodes'

    hDefKey = HKEY_LOCAL_MACHINE
    If Not CheckArray(arrArpProducts) Then 
        If Not UBound(arrArpProducts) = -1 Then 
            Cachelog LOGPOS_REVITEM,LOGHEADING_NONE,ERR_CATEGORYERROR,ERR_NOARRAY& " Terminated ARP detection"
        End If
        Err.Clear
        Exit Sub
    End If
   
    For n = 0 To UBound(arrArpProducts)
        k = -1
        sArpProd = arrArpProducts(n,COL_CONFIGNAME)
        'Server products do not follow the 'WW' naming logic
        If Not InStr(sArpProd,".") > 0 Then sArpProdWW = sArpProd & "WW"
        sName = "PackageIds"
        If UCase(Left(sArpProd,9))="OFFICE14." OR UCase(Left(sArpProd,9))="OFFICE15." Then 
            sName = "PackageRefs"
            sArpProdWW = Right(sArpProd,Len(sArpProd)-9) & "WW"
        End If
        sSubKeyName = REG_ARP & arrArpProducts(n,COL_CONFIGNAME) & "\"

        If RegReadMultiStringValue(hDefKey,sSubKeyName,sName,arrValues) Then
            i = 0
            'The target product is usually at the last entry of the array
            'Check last entry first. If this fails search the array
            If InStr(UCase(arrValues(UBound(arrValues))),UCase(sArpProdWW)) > 0 Then
                k = UBound(arrValues)
            ElseIf InStr(UCase(arrValues(UBound(arrValues))),UCase(sArpProd)) > 0 Then
                k = UBound(arrValues)
            Else
                For Each ProdId in arrValues
                    If InStr(UCase(ProdID),UCase(sArpProdWW)) > 0 Then 
                        k = i
                        Exit For
                    End If 'InStr
                    i = i + 1
                Next 'ProdId
                'Before failing completely try without the 'WW' extension
                If k = - 1 Then
                    For Each ProdId in arrValues
                        If InStr(UCase(ProdID),UCase(sArpProd)) > 0 Then 
                            k = i
                            Exit For
                        End If 'InStr
                        i = i + 1
                    Next 'ProdId
                End If 'k = -1
            End If 'InStr
        Else
            'Cachelog LOGPOS_REVITEM,LOGHEADING_NONE,ERR_CATEGORYERROR,ERR_NOARRAY
        End If 'RegReadMultiStringValue
        
        If Not k = -1 Then 'found a matching 'Config Guid' for the product
            'Add the Config PackageID to the array
            arrArpProducts(n,COL_CONFIGPACKAGEID) = arrValues(k) 
        End If 'k
    Next 'n
End Sub 'FindArpConfigMsi
'=======================================================================================================

'Check if every Office product flagged as 'SystemComponent = 1' has a parent registered
Sub ValidateSystemComponentProducts
    Dim n
    On Error Resume Next
    
    For n = 0 To UBound (arrMaster)
        If Not Err=0 Then Exit For
        If arrMaster(n,COL_ISOFFICEPRODUCT) Then
            If (NOT arrMaster(n, COL_SYSTEMCOMPONENT) = 0)  AND (arrMaster(n, COL_ARPPARENTCOUNT) = 0) Then
                'Identified orphaned product
                'Log warning except for C2R, AER or Rosebud
                Select Case Mid(arrMaster(n, COL_PRODUCTCODE), 11, 4)
                Case "0010", "00B9", "008C", "008F", "007E"
                    ' don't add warning
                Case Else
                    Cachelog LOGPOS_REVITEM,LOGHEADING_NONE,ERR_CATEGORYNOTE,ERR_ORPHANEDITEM & DSV & arrMaster(n,COL_PRODUCTNAME) & DSV & arrMaster(n,COL_PRODUCTCODE)
                    arrMaster(n,COL_NOTES) = arrMaster(n,COL_NOTES) & ERR_CATEGORYNOTE & ERR_ORPHANEDITEM & CSV
                End Select
            End If 'Not arrMaster
        End If 'IsOfficeProduct
    Next 'n
End Sub
'=======================================================================================================

'Adds the ARP config product information to the individual products
Sub AddArpParentsToMaster
    Dim i, n
    On Error Resume Next

    For n = 0 To UBound(arrMaster)
        If Not Err=0 Then Exit For
        If arrMaster(n,COL_ISOFFICEPRODUCT) Then
            For i = 0 To UBound(arrArpProducts)
                If InStr(arrArpProducts(i,COL_ARPALLPRODUCTS),arrMaster(n,COL_PRODUCTCODE)) > 0 Then
                    'add COL_ARPPARENTCOUNT COL_ARPPARENTS
                    arrMaster(n,COL_ARPPARENTCOUNT) = arrMaster(n,COL_ARPPARENTCOUNT) + 1
                    arrMaster(n,COL_ARPPARENTS) = arrMaster(n,COL_ARPPARENTS) & "," & arrArpProducts(i,COL_CONFIGNAME)
                End If
            Next 'i
        End If
    Next 'n
End Sub
'=======================================================================================================

'Add applications 'SystemComponent' flag from ARP to Master array
Sub AddSystemComponentFlagToMaster
    Dim sValue, sSubKeyName
    Dim n
    On Error Resume Next

    For n = 0 To UBound(arrMaster)
        If Not Err=0 Then Exit For
        sSubKeyName = REG_ARP & arrMaster(n,COL_PRODUCTCODE) & "\"
        If RegReadDWordValue(HKEY_LOCAL_MACHINE,sSubKeyName,"SystemComponent",sValue) Then
            arrMaster(n,COL_SYSTEMCOMPONENT) = sValue
        Else
            arrMaster(n,COL_SYSTEMCOMPONENT) = 0
        End If 'RegReadDWordValue "SystemComponent"
    Next 'n
End Sub
'=======================================================================================================

'Find parent entries from 'Add/Remove Programs' for Office applications which are flagged as 'Systemcomponent'
'ARP entries flagged as Systemcomponent will not display under 'Add/Remove Programs'
'From Office 2007 on ARP groups the multi MSI configuration together.
Sub FindArpParents
    Dim ArpProd, ProductCode, Key, dicKey
    Dim n, i, j, iMaxVal, iArpCnt, iPosMaster, iLoop
    Dim bNoSysComponent, bOConfigEntry, fFoundConfigProductCode, fUninstallString
    Dim hDefKey
    Dim sSubKeyName, sCurKey, sName, sValue, sNames, sArpProductName, sRegArp, sMondo
    Dim arrKeys, arrValues, arrTmpArp, arrNames, arrTypes
    Dim dicArpTmp
    Dim tModulStart, tModulEnd
    On Error Resume Next
    
    iLoop = 0 : iMaxVal = 0: iArpCnt = 0 : Redim arrTmpArp(-1)
    Set dicArpTmp = CreateObject("Scripting.Dictionary")
    
    hDefKey = HKEY_LOCAL_MACHINE
    sRegArp = REG_ARP
    sSubKeyName = sRegArp
    Set dicMissingChild = CreateObject("Scripting.Dictionary")
    'If RegEnumKey (hDefKey,sSubKeyName,arrKeys) Then
    For Each ArpProd in dicArp.Keys
        sCurKey = dicArp.Item(ArpProd) & "\"
        bNoSysComponent = True
        bOConfigEntry = False
        sNames = "" : Redim arrNames(-1) : Redim arrTypes(-1)
        If RegEnumValues(hDefKey,sCurKey,arrNames,arrTypes) Then sNames = Join(arrNames)
        If InStr(sNames,"PackageIds") OR _
           InStr(sNames,"PackageRefs") OR _
           UCase(Left(ArpProd, 9)) = "OFFICE14." OR _
           UCase(Left(ArpProd, 9)) = "OFFICE15." OR _
           UCase(Left(ArpProd, 9)) = "OFFICE16." Then bOConfigEntry = True
            
        If bOConfigEntry Then 
            sName = "SystemComponent"
            'If RegDWORD 'SystemComponent' = 1 then it's not showing up in ARP
            If RegReadDWordValue(HKLM, sCurKey, sName, sValue) Then
                If CInt(sValue) = 1 Then
                    CacheLog LOGPOS_REVITEM, LOGHEADING_NONE, ERR_CATEGORYWARN, "Office configuration entry that has been hidden from ARP: " & ArpProd
                End If
            End If 'RegReadValue "SystemComponent"
            If bOConfigEntry Then
                ' found ARP Parent entry
                ' only care if it's not a 'helper' entry 
                If UBound(arrNames) > 5 OR InStr(UCase(ArpProd), ".MONDO") > 0 Then
                    Redim Preserve arrTmpArp(iArpCnt)
                    arrTmpArp(iArpCnt) = ArpProd
                    If NOT dicArpTmp.Exists(sCurKey) Then dicArpTmp.Add sCurKey, ArpProd
                    iArpCnt = iArpCnt + 1
                    Cachelog LOGPOS_RAW, LOGHEADING_NONE, Null, "ARP Entry: " & ArpProd
                Else
                    Cachelog LOGPOS_RAW, LOGHEADING_NONE, Null, "Bad ARP Entry: " & ArpProd
                    Cachelog LOGPOS_REVITEM, LOGHEADING_NONE, ERR_CATEGORYERROR, ERR_OFFSCRUB_TERMINATED
                    Cachelog LOGPOS_REVITEM, LOGHEADING_NONE, ERR_CATEGORYERROR, ERR_BADARPMETADATA & GetHiveString(hDefKey) & "\" & sCurKey
                End If
            End If
        Else
'           'Handled in Click2Run V2 below
        End If
    Next 'ArpProd

    ' initialize arrArpProducts
    If iArpCnt > 0 Then 
        iArpCnt = iArpCnt - 1
        Redim arrArpProducts(iArpCnt,ARP_CHILDOFFSET)
    Else
        Redim arrArpProducts(-1)
    End If
    ' copy the identified configuration parents to the actual target array and add additional fields
    n = 0
    'For Each ArpProd in arrTmpArp
    For Each dicKey in dicArpTmp.Keys
        ArpProd = dicArpTmp.Item(dicKey)
        sName = "ProductCodes"
        'sSubKeyName = sRegArp & ArpProd & "\"
        sSubKeyName = dicKey
        If RegReadMultiStringValue(hDefKey, sSubKeyName, sName, arrValues) Then
            ' fill the array set
            arrArpProducts(n,COL_CONFIGNAME) = ArpProd
            arrArpProducts(n, COL_CONFIGINSTALLTYPE) = "MSI"
            RegReadStringValue hDefKey, sSubKeyName, "DisplayVersion", sValue
            arrArpProducts(n, ARP_PRODUCTVERSION) = sValue
            If UBound (arrValues) + ARP_CHILDOFFSET > iMaxVal Then 
                iMaxVal = UBound (arrValues)+ARP_CHILDOFFSET
                Redim Preserve arrArpProducts(iArpCnt,iMaxVal)
            End If 'UBound(arrValues)
            j = ARP_CHILDOFFSET
            For Each ProductCode in arrValues
                arrArpProducts(n,COL_ARPALLPRODUCTS) = arrArpProducts(n,COL_ARPALLPRODUCTS) & ProductCode
                arrArpProducts(n,j) = ProductCode
            ' validate that the product is installed
                If NOT dicProducts.Exists(ProductCode) Then
                    Cachelog LOGPOS_REVITEM,LOGHEADING_NONE,ERR_CATEGORYERROR,ERR_MISSINGCHILD&"Config Product: "&ArpProd& CSV & "Missing Product: "&ProductCode
                    If NOT dicMissingChild.Exists(ProductCode) Then dicMissingChild.Add ProductCode,ArpProd
                End If
                j = j + 1
            Next 'ProductCode
        ' identify the config ProductCode
            fFoundConfigProductCode = False
            sArpProductName = ""
            sArpProductName = GetArpProductName(ArpProd)
            For Each ProductCode in arrValues
                iPosMaster = GetArrayPosition(arrMaster,ProductCode)
                If NOT iPosMaster = -1 Then
                    If LCase(Trim(sArpProductName)) = LCase(Trim(arrMaster(GetArrayPosition(arrMaster,ProductCode),COL_PRODUCTNAME))) Then
                        arrArpProducts(n,ARP_CONFIGPRODUCTCODE) = ProductCode
                        fFoundConfigProductCode = True
                        Exit For
                    End If
                End If
            Next 'ProductCode
            If NOT fFoundConfigProductCode Then
                i = 0
                For Each ProductCode in arrValues
                    If (NOT Mid(ProductCode,11,4)="002A") AND (Mid(ProductCode,16,4)="0000") _ 
                        OR (Mid(ProductCode,12,2)="10") _ 
                    Then
                        i = i + 1
                        arrArpProducts(n,ARP_CONFIGPRODUCTCODE) = ProductCode
                    End If
                Next 'ProductCode
            ' ensure to not have ambigious hits
                If i > 1 Then
                    Dim MsiDb,Record
                    Dim qView
                    For Each ProductCode in arrValues
                        If (NOT Mid(ProductCode,11,4)="002A") AND (Mid(ProductCode,16,4)="0000") _ 
                            OR (Mid(ProductCode,12,2)="10") _ 
                        Then
                            Set MsiDb = oMsi.OpenDatabase(arrMaster(GetArrayPosition(arrMaster,ProductCode),COL_CACHEDMSI),MSIOPENDATABASEMODE_READONLY)
	                        Set qView = MsiDb.OpenView("SELECT * FROM Property WHERE `Property` = 'SetupExeArpId'")
	                        qView.Execute 
	                        Set Record = qView.Fetch()
	                        If Not Record Is Nothing Then
	                            arrArpProducts(n,ARP_CONFIGPRODUCTCODE) = ProductCode
	                            qView.Close
	                            Exit For
	                        End If 'Is Nothing
                            'SetupExeArpId' 
                        End If
                    Next 'ProductCode
                End If
            End If 'NOT fFoundConfigProductCode
        Else
        ' handle Click2Run
            arrArpProducts(n, COL_CONFIGNAME) = ArpProd
            arrArpProducts(n, COL_CONFIGINSTALLTYPE) = "C2R"
            If 1 + ARP_CHILDOFFSET > iMaxVal Then 
                iMaxVal = 1 + ARP_CHILDOFFSET
                Redim Preserve arrArpProducts(iArpCnt, iMaxVal)
            End If 
            ' V1
            'If RegEnumKey (HKLM, REG_ARP, arrKeys) Then
            For Each Key in dicArp.Keys
                If Len(Key) = 38 Then
                    If Mid(Key, 11, 4) = "006D" Then
                        arrArpProducts(n, COL_ARPALLPRODUCTS) = Key
                        arrArpProducts(n, ARP_CHILDOFFSET) = Key
                        arrArpProducts(n, ARP_CONFIGPRODUCTCODE) = Key
                    End If
                End If
            Next 'Key
            'End If
            ' V2
            sMondo = ""
            If InStr(UCase(ArpProd), ".MONDO") > 0 Then 
                sMondo = Mid(ArpProd, 7, 2) & "0000-000F-0000-0000-0000000FF1CE}"
            Else
                RegReadValue HKLM, dicArp.Item(ArpProd), "UninstallString", sValue, "REG_SZ"
                sValue = Mid(sValue, InStr(sValue, "Microsoft Office ") + 17, 2)
                sMondo = sValue & "0000-000F-0000-0000-0000000FF1CE}"
            End If
            iPosMaster = GetArrayPositionFromPattern(arrMaster, sMondo)
            If InStr(UCase(ArpProd), ".MONDO") > 0 Then
                arrArpProducts(n, COL_CONFIGINSTALLTYPE) = "VIRTUAL"
                If NOT iPosMaster = -1 Then
                    arrArpProducts(n, ARP_CONFIGPRODUCTCODE) = arrMaster(iPosMaster, COL_PRODUCTCODE)
                End If
                If UBound (arrMVProducts) + ARP_CHILDOFFSET > iMaxVal Then 
                    iMaxVal = UBound (arrMVProducts) + ARP_CHILDOFFSET
                    Redim Preserve arrArpProducts(iArpCnt,iMaxVal)
                End If 'UBound(arrMVProducts)
                ' fill the array set
                j = ARP_CHILDOFFSET
                i = 0
                For i = 0 To UBound(arrMVProducts)
                    'If arrMaster(i, COL_VIRTUALIZED) = 1 Then
                        arrArpProducts(n, COL_ARPALLPRODUCTS) = arrArpProducts(n, COL_ARPALLPRODUCTS) & arrMVProducts(i, COL_PRODUCTCODE)
                        arrArpProducts(n, j) = arrMVProducts(i, COL_PRODUCTCODE)
                        j = j + 1
                    'End If
                Next 'i

            Else
                If NOT iPosMaster = -1 Then
                    arrArpProducts(n,COL_ARPALLPRODUCTS) = arrMaster(iPosMaster, COL_PRODUCTCODE)
                    arrArpProducts(n,ARP_CHILDOFFSET) = arrMaster(iPosMaster, COL_PRODUCTCODE)
                    arrArpProducts(n,ARP_CONFIGPRODUCTCODE) = arrMaster(iPosMaster, COL_PRODUCTCODE)
                End If
            End If
        End If 'RegReadValue "ProductCodes"
        n = n + 1
    Next 'ArpProd

End Sub 'FindArpParents

'=======================================================================================================
'Module OSPP - Licensing Data / LicenseState
'=======================================================================================================

Sub OsppInit
    On Error Resume Next
    Dim oWmiLocal
    Set oWmiLocal = GetObject("winmgmts:\\.\root\cimv2")
    If iVersionNt > 601 Then 
        Set Spp = oWmiLocal.ExecQuery("SELECT ID, ApplicationId, EvaluationEndDate, PartialProductKey, Description, Name, LicenseStatus, LicenseStatusReason, ProductKeyID, GracePeriodRemaining, LicenseFamily, KeyManagementServiceLookupDomain, VLActivationType, ADActivationObjectName, ADActivationObjectDN, ADActivationCsvlkPid, ADActivationCsvlkSkuId, VLActivationTypeEnabled, DiscoveredKeyManagementServiceMachineName, DiscoveredKeyManagementServiceMachinePort, VLActivationInterval, VLRenewalInterval, KeyManagementServiceMachine, KeyManagementServicePort, ProductKeyID2 FROM SoftwareLicensingProduct")
    End If
    Set Ospp = oWmiLocal.ExecQuery("SELECT ID, ApplicationId, EvaluationEndDate, PartialProductKey, Description, Name, LicenseStatus, LicenseStatusReason, ProductKeyID, GracePeriodRemaining, LicenseFamily, DiscoveredKeyManagementServiceMachineName, DiscoveredKeyManagementServiceMachinePort, VLActivationInterval, VLRenewalInterval, KeyManagementServiceMachine, KeyManagementServicePort, ProductKeyID2 FROM OfficeSoftwareProtectionProduct")
    bOsppInit = True
End Sub 'OsppInit
'=======================================================================================================

Function GetLicCnt(iPosMaster, iVersionMajor, ByVal sConfigName, sPossibleSkus, sPossibleSkusFull)
    On Error Resume Next
    Dim iLicCnt, iLeft, iCnt
    Dim sPrefix
    Dim ProdLic, ConfigProd, prop, ProtectionClass
    Dim arrLic
    Dim fMsiMatchesOsppID, fExclude
    
    If NOT bOsppInit Then OsppInit
    If NOT CheckArray(Ospp) AND NOT IsObject(Ospp) Then Exit Function
    GetPrefixAndConfig iVersionMajor, sConfigName, sPrefix
    If sConfigName <> "" Then arrLic = Split(sConfigName,";")
    iLicCnt = 0
    If CheckArray(arrLic) Then
        For Each ConfigProd in arrLic
            iLeft = Len(sPrefix) + Len(ConfigProd)
            If iVersionMajor > 14 and iVersionNt > 601 Then
                Set ProtectionClass = Spp
            Else
                Set ProtectionClass = Ospp
            End If
            Err.Clear
            For Each ProdLic in ProtectionClass
                If NOT Err = 0 Then
                    err.Clear
                    Exit For
                End If
                fMsiMatchesOsppID = False
                fExclude = False
                If NOT iPosMaster = -1 Then
                    If (NOT IsNull (Mid(Prodlic.ProductKeyID, 13, 10))) AND NOT IsNull(arrMaster(iPosMaster, COL_PRODUCTID)) Then _ 
                        fMsiMatchesOsppID = (Mid(arrMaster(iPosMaster, COL_PRODUCTID), 7, 10) = Mid(Prodlic.ProductKeyID, 13, 10) )
                End If
            ' check for exceptions
                If UCase(ConfigProd) = "VISIO" Then
                    fExclude = InStr(Prodlic.Name, "DeltaTrial") > 0
                End If 'visio
            ' add matches to counter and sku-string
                If (LCase(Left(ProdLic.Name, iLeft)) = LCase(sPrefix & ConfigProd) OR fMsiMatchesOsppID) AND NOT fExclude Then 
                    iLicCnt = iLicCnt + 1
                    sPossibleSkus = sPossibleSkus & "; " & ProdLic.Name
                End If
            Next 'ProdLic
        Next 'ConfigProd
    End If
    sPossibleSkusFull = sPossibleSkus
    sPossibleSkus = Replace(sPossibleSkus,sPrefix,"")
    sPossibleSkus = Replace(sPossibleSkus," edition;",",")
    GetLicCnt = iLicCnt
End Function 'GetLicCnt
'=======================================================================================================

Function GetLicenseData(iPosMaster, iVersionMajor, ByVal sConfigName, iLicPos, sPossibleSkusFull)
    On Error Resume Next
    Dim arrLicData (22)
    Dim iPropCnt, iLicCnt, iLeft
    Dim ProdLic, prop, ConfigProd, ProtectionClass
    Dim sPrefix, sAllLic, sName, sTmp
    Dim arrLic
    Dim fMsiMatchesOsppID
    
    If NOT bOsppInit Then OsppInit
    iLeft = Len(sConfigName)
    GetPrefixAndConfig iVersionMajor, sConfigName, sPrefix
    If sConfigName <> "" Then arrLic = Split(sConfigName, ";")
    For Each ConfigProd in arrLic
        sAllLic = sAllLic & sPrefix&ConfigProd & ";"
    Next
    If CheckArray(arrLic) Then
        iLicCnt = 0
        If iVersionMajor > 14 and iVersionNt > 601 Then
            Set ProtectionClass = Spp
        Else
            Set ProtectionClass = Ospp
        End If
        For Each ProdLic in ProtectionClass
            fMsiMatchesOsppID = False
            If NOT iPosMaster = -1 Then
                If (NOT IsNull (Mid(Prodlic.ProductKeyID, 13, 10))) AND NOT IsNull(arrMaster(iPosMaster, COL_PRODUCTID)) Then _ 
                    fMsiMatchesOsppID = (Mid(arrMaster(iPosMaster, COL_PRODUCTID), 7, 10) = Mid(Prodlic.ProductKeyID, 13, 10) )
            End If
            iLicCnt = iLicCnt + 1
            For Each ConfigProd in arrLic
                iLeft = Len(sPrefix) + Len(ConfigProd)
                If Len(ProdLic.Name) > iLeft -1 Then
                    If (LCase(Left(ProdLic.Name, iLeft)) = LCase(sPrefix & ConfigProd) OR fMsiMatchesOsppID) AND (iLicCnt > iLicPos) AND InStr(sPossibleSkusFull, Prodlic.Name) > 0 Then 
                        iLicPos = iLicCnt
                        arrLicData(OSPP_ACTIVATIONTYPE) = ProdLic.VLActivationTypeEnabled
                        arrLicData(OSPP_ID) = ProdLic.ID
                        arrLicData(OSPP_APPLICATIONID) = ProdLic.ApplicationId
                        arrLicData(OSPP_PARTIALPRODUCTKEY) = ProdLic.PartialProductKey
                        arrLicData(OSPP_DESCRIPTION) = ProdLic.Description
                        arrLicData(OSPP_NAME) = ProdLic.Name
                        arrLicData(OSPP_LICENSESTATUS) = ProdLic.LicenseStatus
                        arrLicData(OSPP_LICENSESTATUSREASON) = ProdLic.LicenseStatusReason
                        arrLicData(OSPP_GRACEPERIODREMAINING) = ProdLic.GracePeriodRemaining
                        arrLicData(OSPP_LICENSEFAMILY) = ProdLic.LicenseFamily
                        arrLicData(OSPP_DISCOVEREDKEYMANAGEMENTSERVICEMACHINENAME) = ProdLic.DiscoveredKeyManagementServiceMachineName
                        arrLicData(OSPP_KEYMANAGEMENTSERVICEPORT) = ProdLic.KeyManagementServicePort
                        arrLicData(OSPP_VLACTIVATIONINTERVAL) = ProdLic.VLActivationInterval
                        arrLicData(OSPP_VLRENEWALINTERVAL) = ProdLic.VLRenewalInterval
                        If Not IsNull (ProdLic.ProductKeyID) Then
                            arrLicData(OSPP_PRODUCTKEYID) = ProdLic.ProductKeyID
                            arrLicData(OSPP_PRODUCTKEYID2) = ProdLic.ProductKeyID2
                            sTmp = Right (Replace (arrLicData(OSPP_PRODUCTKEYID2), "-", ""), 19)
                            arrLicData(OSPP_MACHINEKEY) = Mid (sTmp, 1, 5) & "-" & Mid (sTmp, 6, 3) & "-" & Mid (sTmp, 9, 6)
                        End If
                        If iVersionNt > 601 Then
                            arrLicData(OSPP_DISCOVEREDKEYMANAGEMENTSERVICEMACHINEPORT) = ProdLic.DiscoveredKeyManagementServiceMachinePort
                            arrLicData(OSPP_KEYMANGEMENTSERVICELOOKUPDOMAIN) = ProdLic.KeyManagementServiceLookupDomain
                            Select Case ProdLic.VLActivationTypeEnabled
                            Case 0, 1 'AD
                                arrLicData(OSPP_ADACTIVATIONOBJECTNAME) = ProdLic.ADActivationObjectName
                                arrLicData(OSPP_ADACTIVATIONOBJECTDN) = ProdLic.ADActivationObjectDN
                                arrLicData(OSPP_ADACTIVATIONCSVLPID) = ProdLic.ADActivationCsvlkPid
                                arrLicData(OSPP_ADACTIVATIONCSVLSKUID) = ProdLic.ADActivationCsvlkSkuId
                            Case 2 'KMS
                            Case 3 'Token
                            End Select
                        End If
                        GetLicenseData = arrLicData
                        Exit Function
                    End If
                End If
            Next
        Next 'ProdLic
    End If
    GetLicenseData = arrLicData
End Function 'GetLicenseData
'=======================================================================================================

'-------------------------------------------------------------------------------
'   GetLicPrefix
'
'   Get the prefix string used in the Prodlic.Name string
'   used in Prodlic.Name
'   Returns the updated ConfigProd name string that matches the Prodlic.Name
'-------------------------------------------------------------------------------

Sub GetPrefixAndConfig(iVersionMajor, sConfigName, sPrefix)
    On Error Resume Next
    Select Case iVersionMajor
    Case 14
        sPrefix = "Office 14, Office"
        If sConfigName = "SingleImage" Then sConfigName = "SingleImage;Professional;HomeBusiness;HomeStudent;OneNote;Word;HSOneNote;HSWord;OEM"
        If sConfigName = "Click2Run" Then sConfigName = "Click2Run;HomeBusiness;HomeStudent;Starter;OEM"
        If Len(sConfigName) > 2 Then
            If UCase(Left(sConfigName, 3)) = "PRJ" Then sConfigName = "Project"
            If UCase(Right(sConfigName, 1)) = "R" Then sConfigName = Left(sConfigName,(Len(sConfigName) - 1))
        End If
    Case 15
        sPrefix = "Office 15, Office"
        If Len(sConfigName) > 2 Then
            If UCase(Left(sConfigName, 3)) = "PRJ" Then sConfigName = "Project"
            If UCase(Right(sConfigName, 1)) = "R" Then sConfigName = Left(sConfigName, (Len(sConfigName) - 1))
            If UCase(sConfigName) = "VISSTD" Then sConfigName = "VISSTD;VisioStd"
            If UCase(sConfigName) = "VISPRO" Then sConfigName = "VISPRO;VisioPro"
            If UCase(sConfigName) = "PROPLUSR" Then sConfigName = "PROPLUSR;O365ProPlusR"
        End If
    Case 16
        sPrefix = "Office 16, Office16"
        If Len(sConfigName) > 2 Then
            If UCase(Left(sConfigName, 3)) = "PRJ" Then sConfigName = "Project"
            If UCase(Right(sConfigName, 1)) = "R" Then sConfigName = Left(sConfigName, (Len(sConfigName) - 1))
            If UCase(sConfigName) = "VISSTD" Then sConfigName = "VISSTD;VisioStd"
            If UCase(sConfigName) = "VISPRO" Then sConfigName = "VISPRO;VisioPro"
        End If
    Case Else
    End Select
End Sub
'=======================================================================================================

Function GetLicErrDesc(hErr)
On Error Resume Next
Select Case "0x"& hErr
Case "0x0" : GetLicErrDesc = "Success."
Case "0xC004B001" : GetLicErrDesc = "The activation server determined that the license is invalid."
Case "0xC004B002" : GetLicErrDesc = "The activation server determined that the license is invalid."
Case "0xC004B003" : GetLicErrDesc = "The activation server determined that the license is invalid."
Case "0xC004B004" : GetLicErrDesc = "The activation server determined that the license is invalid."
Case "0xC004B005" : GetLicErrDesc = "The activation server determined that the license is invalid."
Case "0xC004B006" : GetLicErrDesc = "The activation server determined that the license is invalid."
Case "0xC004B007" : GetLicErrDesc = "The activation server reported that the computer could not connect to the activation server."
Case "0xC004B008" : GetLicErrDesc = "The activation server determined that the computer could not be activated."
Case "0xC004B009" : GetLicErrDesc = "The activation server determined that the license is invalid."
Case "0xC004B011" : GetLicErrDesc = "The activation server determined that your computer clock time is not correct. You must correct your clock before you can activate."
Case "0xC004B100" : GetLicErrDesc = "The activation server determined that the computer could not be activated."
Case "0xC004C001" : GetLicErrDesc = "The activation server determined the specified product key is invalid."
Case "0xC004C002" : GetLicErrDesc = "The activation server determined there is a problem with the specified product key."
Case "0xC004C003" : GetLicErrDesc = "The activation server determined the specified product key has been blocked."
Case "0xC004C004" : GetLicErrDesc = "The activation server determined the specified product key is invalid."
Case "0xC004C005" : GetLicErrDesc = "The activation server determined the license is invalid."
Case "0xC004C006" : GetLicErrDesc = "The activation server determined the license is invalid."
Case "0xC004C007" : GetLicErrDesc = "The activation server determined the specified product key is invalid."
Case "0xC004C008" : GetLicErrDesc = "The activation server determined that the specified product key could not be used."
Case "0xC004C009" : GetLicErrDesc = "The activation server determined the license is invalid."
Case "0xC004C00A" : GetLicErrDesc = "The activation server determined the license is invalid."
Case "0xC004C00B" : GetLicErrDesc = "The activation server determined the license is invalid."
Case "0xC004C00C" : GetLicErrDesc = "The activation server experienced an error."
Case "0xC004C00D" : GetLicErrDesc = "The activation server determined the license is invalid."
Case "0xC004C00E" : GetLicErrDesc = "The activation server determined the specified product key is invalid."
Case "0xC004C00F" : GetLicErrDesc = "The activation server determined the specified product key is invalid."
Case "0xC004C010" : GetLicErrDesc = "The activation server determined the specified product key is invalid."
Case "0xC004C011" : GetLicErrDesc = "The activation server determined the license is invalid."
Case "0xC004C012" : GetLicErrDesc = "The activation server experienced a network error."
Case "0xC004C013" : GetLicErrDesc = "The activation server experienced an error."
Case "0xC004C014" : GetLicErrDesc = "The activation server experienced an error."
Case "0xC004C020" : GetLicErrDesc = "The activation server reported that the Multiple Activation Key has exceeded its limit."
Case "0xC004C021" : GetLicErrDesc = "The activation server reported that the Multiple Activation Key extension limit has been exceeded."
Case "0xC004C022" : GetLicErrDesc = "The activation server reported that the re-issuance limit was not found."
Case "0xC004C023" : GetLicErrDesc = "The activation server reported that the override request was not found."
Case "0xC004C016" : GetLicErrDesc = "The activation server reported that the specified product key cannot be used for online activation."
Case "0xC004C017" : GetLicErrDesc = "The activation server determined the specified product key has been blocked for this geographic location."
Case "0xC004C015" : GetLicErrDesc = "The activation server experienced an error."
Case "0xC004C050" : GetLicErrDesc = "The activation server experienced a general error."
Case "0xC004C030" : GetLicErrDesc = "The activation server reported that time based activation attempted before start date."
Case "0xC004C031" : GetLicErrDesc = "The activation server reported that time based activation attempted after end date."
Case "0xC004C032" : GetLicErrDesc = "The activation server reported that new time based activation not available."
Case "0xC004C033" : GetLicErrDesc = "The activation server reported that time based product key not configured for activation."
Case "0xC004C04F" : GetLicErrDesc = "The activation server reported that no business rules available to activate specified product key."
Case "0xC004C700" : GetLicErrDesc = "The activation server reported that business rule cound not find required input."
Case "0xC004C750" : GetLicErrDesc = "The activation server reported that NULL value specified for business property name and Id."
Case "0xC004C751" : GetLicErrDesc = "The activation server reported that property name specifies unknown property."
Case "0xC004C752" : GetLicErrDesc = "The activation server reported that property Id specifies unknown property."
Case "0xC004C755" : GetLicErrDesc = "The activation server reported that it failed to update product key binding."
Case "0xC004C756" : GetLicErrDesc = "The activation server reported that it failed to insert product key binding."
Case "0xC004C757" : GetLicErrDesc = "The activation server reported that it failed to delete product key binding."
Case "0xC004C758" : GetLicErrDesc = "The activation server reported that it failed to process input XML for product key bindings."
Case "0xC004C75A" : GetLicErrDesc = "The activation server reported that it failed to insert product key property."
Case "0xC004C75B" : GetLicErrDesc = "The activation server reported that it failed to update product key property."
Case "0xC004C75C" : GetLicErrDesc = "The activation server reported that it failed to delete product key property."
Case "0xC004C764" : GetLicErrDesc = "The activation server reported that the product key type is unknown."
Case "0xC004C770" : GetLicErrDesc = "The activation server reported that the product key type is being used by another user."
Case "0xC004C780" : GetLicErrDesc = "The activation server reported that it failed to insert product key record."
Case "0xC004C781" : GetLicErrDesc = "The activation server reported that it failed to update product key record."
Case "0xC004C401" : GetLicErrDesc = "The Vista Genuine Advantage Service determined that the installation is not genuine."
Case "0xC004C600" : GetLicErrDesc = "The Vista Genuine Advantage Service determined that the installation is not genuine."
Case "0xC004C801" : GetLicErrDesc = "The activation server determined the specified product key is invalid."
Case "0xC004C802" : GetLicErrDesc = "The activation server determined the specified product key is invalid."
Case "0xC004C803" : GetLicErrDesc = "The activation server determined the specified product key has been revoked."
Case "0xC004C804" : GetLicErrDesc = "The activation server determined the specified product key is invalid."
Case "0xC004C805" : GetLicErrDesc = "The activation server determined the specified product key is invalid."
Case "0xC004C810" : GetLicErrDesc = "The activation server determined the specified product key is invalid."
Case "0xC004C811" : GetLicErrDesc = "The activation server determined the license is invalid."
Case "0xC004C812" : GetLicErrDesc = "The activation server determined that the specified product key has exceeded its activation count."
Case "0xC004C813" : GetLicErrDesc = "The activation server determined the license is invalid."
Case "0xC004C814" : GetLicErrDesc = "The activation server determined the specified product key is invalid."
Case "0xC004C815" : GetLicErrDesc = "The activation server determined the license is invalid."
Case "0xC004C816" : GetLicErrDesc = "The activation server reported that the specified product key cannot be used for online activation."
Case "0xC004E001" : GetLicErrDesc = "The Software Licensing Service determined that the specified context is invalid."
Case "0xC004E002" : GetLicErrDesc = "The Software Licensing Service reported that the license store contains inconsistent data."
Case "0xC004E003" : GetLicErrDesc = "The Software Licensing Service reported that license evaluation failed."
Case "0xC004E004" : GetLicErrDesc = "The Software Licensing Service reported that the license has not been evaluated."
Case "0xC004E005" : GetLicErrDesc = "The Software Licensing Service reported that the license is not activated."
Case "0xC004E006" : GetLicErrDesc = "The Software Licensing Service reported that the license contains invalid data."
Case "0xC004E007" : GetLicErrDesc = "The Software Licensing Service reported that the license store does not contain the requested license."
Case "0xC004E008" : GetLicErrDesc = "The Software Licensing Service reported that the license property is invalid."
Case "0xC004E009" : GetLicErrDesc = "The Software Licensing Service reported that the license store is not initialized."
Case "0xC004E00A" : GetLicErrDesc = "The Software Licensing Service reported that the license store is already initialized."
Case "0xC004E00B" : GetLicErrDesc = "The Software Licensing Service reported that the license property is invalid."
Case "0xC004E00C" : GetLicErrDesc = "The Software Licensing Service reported that the license could not be opened or created."
Case "0xC004E00D" : GetLicErrDesc = "The Software Licensing Service reported that the license could not be written."
Case "0xC004E00E" : GetLicErrDesc = "The Software Licensing Service reported that the license store could not read the license file."
Case "0xC004E00F" : GetLicErrDesc = "The Software Licensing Service reported that the license property is corrupted."
Case "0xC004E010" : GetLicErrDesc = "The Software Licensing Service reported that the license property is missing."
Case "0xC004E011" : GetLicErrDesc = "The Software Licensing Service reported that the license store contains an invalid license file."
Case "0xC004E012" : GetLicErrDesc = "The Software Licensing Service reported that the license store failed to start synchronization properly."
Case "0xC004E013" : GetLicErrDesc = "The Software Licensing Service reported that the license store failed to synchronize properly."
Case "0xC004E014" : GetLicErrDesc = "The Software Licensing Service reported that the license property is invalid."
Case "0xC004E015" : GetLicErrDesc = "The Software Licensing Service reported that license consumption failed."
Case "0xC004E016" : GetLicErrDesc = "The Software Licensing Service reported that the product key is invalid."
Case "0xC004E017" : GetLicErrDesc = "The Software Licensing Service reported that the product key is invalid."
Case "0xC004E018" : GetLicErrDesc = "The Software Licensing Service reported that the product key is invalid."
Case "0xC004E019" : GetLicErrDesc = "The Software Licensing Service determined that validation of the specified product key failed."
Case "0xC004E01A" : GetLicErrDesc = "The Software Licensing Service reported that invalid add-on information was found."
Case "0xC004E01B" : GetLicErrDesc = "The Software Licensing Service reported that not all hardware information could be collected."
Case "0xC004E01C" : GetLicErrDesc = "This evaluation product key is no longer valid."
Case "0xC004E01D" : GetLicErrDesc = "The new product key cannot be used on this installation of Windows. Type a different product key. (CD-AB)"
Case "0xC004E01E" : GetLicErrDesc = "The new product key cannot be used on this installation of Windows. Type a different product key. (AB-AB)"
Case "0xC004E01F" : GetLicErrDesc = "The new product key cannot be used on this installation of Windows. Type a different product key. (AB-CD)"
Case "0xC004E020" : GetLicErrDesc = "The Software Licensing Service reported that there is a mismatched between a policy value and information stored in the OtherInfo section."
Case "0xC004E021" : GetLicErrDesc = "The Software Licensing Service reported that the Genuine information contained in the license is not consistent."
Case "0xC004E022" : GetLicErrDesc = "The Software Licensing Service reported that the secure store id value in license does not match with the current value."
Case "0x8004E101" : GetLicErrDesc = "The Software Licensing Service reported that the Token Store file version is invalid."
Case "0x8004E102" : GetLicErrDesc = "The Software Licensing Service reported that the Token Store contains an invalid descriptor table."
Case "0x8004E103" : GetLicErrDesc = "The Software Licensing Service reported that the Token Store contains a token with an invalid header/footer."
Case "0x8004E104" : GetLicErrDesc = "The Software Licensing Service reported that a Token Store token has an invalid name."
Case "0x8004E105" : GetLicErrDesc = "The Software Licensing Service reported that a Token Store token has an invalid extension."
Case "0x8004E106" : GetLicErrDesc = "The Software Licensing Service reported that the Token Store contains a duplicate token."
Case "0x8004E107" : GetLicErrDesc = "The Software Licensing Service reported that a token in the Token Store has a size mismatch."
Case "0x8004E108" : GetLicErrDesc = "The Software Licensing Service reported that a token in the Token Store contains an invalid hash."
Case "0x8004E109" : GetLicErrDesc = "The Software Licensing Service reported that the Token Store was unable to read a token."
Case "0x8004E10A" : GetLicErrDesc = "The Software Licensing Service reported that the Token Store was unable to write a token."
Case "0x8004E10B" : GetLicErrDesc = "The Software Licensing Service reported that the Token Store attempted an invalid file operation."
Case "0x8004E10C" : GetLicErrDesc = "The Software Licensing Service reported that there is no active transaction."
Case "0x8004E10D" : GetLicErrDesc = "The Software Licensing Service reported that the Token Store file header is invalid."
Case "0x8004E10E" : GetLicErrDesc = "The Software Licensing Service reported that a Token Store token descriptor is invalid."
Case "0xC004F001" : GetLicErrDesc = "The Software Licensing Service reported an internal error."
Case "0xC004F002" : GetLicErrDesc = "The Software Licensing Service reported that rights consumption failed."
Case "0xC004F003" : GetLicErrDesc = "The Software Licensing Service reported that the required license could not be found."
Case "0xC004F004" : GetLicErrDesc = "The Software Licensing Service reported that the product key does not match the range defined in the license."
Case "0xC004F005" : GetLicErrDesc = "The Software Licensing Service reported that the product key does not match the product key for the license."
Case "0xC004F006" : GetLicErrDesc = "The Software Licensing Service reported that the signature file for the license is not available."
Case "0xC004F007" : GetLicErrDesc = "The Software Licensing Service reported that the license could not be found."
Case "0xC004F008" : GetLicErrDesc = "The Software Licensing Service reported that the license could not be found."
Case "0xC004F009" : GetLicErrDesc = "The Software Licensing Service reported that the grace period expired."
Case "0xC004F00A" : GetLicErrDesc = "The Software Licensing Service reported that the application ID does not match the application ID for the license."
Case "0xC004F00B" : GetLicErrDesc = "The Software Licensing Service reported that the product identification data is not available."
Case "0x4004F00C" : GetLicErrDesc = "The Software Licensing Service reported that the application is running within the valid grace period."
Case "0x4004F00D" : GetLicErrDesc = "The Software Licensing Service reported that the application is running within the valid out of tolerance grace period."
Case "0xC004F00E" : GetLicErrDesc = "The Software Licensing Service determined that the license could not be used by the current version of the security processor component."
Case "0xC004F00F" : GetLicErrDesc = "The Software Licensing Service reported that the hardware ID binding is beyond the level of tolerance."
Case "0xC004F010" : GetLicErrDesc = "The Software Licensing Service reported that the product key is invalid."
Case "0xC004F011" : GetLicErrDesc = "The Software Licensing Service reported that the license file is not installed."
Case "0xC004F012" : GetLicErrDesc = "The Software Licensing Service reported that the call has failed because the value for the input key was not found."
Case "0xC004F013" : GetLicErrDesc = "The Software Licensing Service determined that there is no permission to run the software."
Case "0xC004F014" : GetLicErrDesc = "The Software Licensing Service reported that the product key is not available."
Case "0xC004F015" : GetLicErrDesc = "The Software Licensing Service reported that the license is not installed."
Case "0xC004F016" : GetLicErrDesc = "The Software Licensing Service determined that the request is not supported."
Case "0xC004F017" : GetLicErrDesc = "The Software Licensing Service reported that the license is not installed."
Case "0xC004F018" : GetLicErrDesc = "The Software Licensing Service reported that the license does not contain valid location data for the activation server."
Case "0xC004F019" : GetLicErrDesc = "The Software Licensing Service determined that the requested event ID is invalid."
Case "0xC004F01A" : GetLicErrDesc = "The Software Licensing Service determined that the requested event is not registered with the service."
Case "0xC004F01B" : GetLicErrDesc = "The Software Licensing Service reported that the event ID is already registered."
Case "0xC004F01C" : GetLicErrDesc = "The Software Licensing Service reported that the license is not installed."
Case "0xC004F01D" : GetLicErrDesc = "The Software Licensing Service reported that the verification of the license failed."
Case "0xC004F01E" : GetLicErrDesc = "The Software Licensing Service determined that the input data type does not match the data type in the license."
Case "0xC004F01F" : GetLicErrDesc = "The Software Licensing Service determined that the license is invalid."
Case "0xC004F020" : GetLicErrDesc = "The Software Licensing Service determined that the license package is invalid."
Case "0xC004F021" : GetLicErrDesc = "The Software Licensing Service reported that the validity period of the license has expired."
Case "0xC004F022" : GetLicErrDesc = "The Software Licensing Service reported that the license authorization failed."
Case "0xC004F023" : GetLicErrDesc = "The Software Licensing Service reported that the license is invalid."
Case "0xC004F024" : GetLicErrDesc = "The Software Licensing Service reported that the license is invalid."
Case "0xC004F025" : GetLicErrDesc = "The Software Licensing Service reported that the action requires administrator privilege."
Case "0xC004F026" : GetLicErrDesc = "The Software Licensing Service reported that the required data is not found."
Case "0xC004F027" : GetLicErrDesc = "The Software Licensing Service reported that the license is tampered."
Case "0xC004F028" : GetLicErrDesc = "The Software Licensing Service reported that the policy cache is invalid."
Case "0xC004F029" : GetLicErrDesc = "The Software Licensing Service cannot be started in the current OS mode."
Case "0xC004F02A" : GetLicErrDesc = "The Software Licensing Service reported that the license is invalid."
Case "0xC004F02C" : GetLicErrDesc = "The Software Licensing Service reported that the format for the offline activation data is incorrect."
Case "0xC004F02D" : GetLicErrDesc = "The Software Licensing Service determined that the version of the offline Confirmation ID (CID) is incorrect."
Case "0xC004F02E" : GetLicErrDesc = "The Software Licensing Service determined that the version of the offline Confirmation ID (CID) is not supported."
Case "0xC004F02F" : GetLicErrDesc = "The Software Licensing Service reported that the length of the offline Confirmation ID (CID) is incorrect."
Case "0xC004F030" : GetLicErrDesc = "The Software Licensing Service determined that the Installation ID (IID) or the Confirmation ID (CID) could not been saved."
Case "0xC004F031" : GetLicErrDesc = "The Installation ID (IID) and the Confirmation ID (CID) do not match. Please confirm the IID and reacquire a new CID if necessary."
Case "0xC004F032" : GetLicErrDesc = "The Software Licensing Service determined that the binding data is invalid."
Case "0xC004F033" : GetLicErrDesc = "The Software Licensing Service reported that the product key is not allowed to be installed. Please see the eventlog for details."
Case "0xC004F034" : GetLicErrDesc = "The Software Licensing Service reported that the license could not be found or was invalid."
Case "0xC004F035" : GetLicErrDesc = "The Software Licensing Service reported that the computer could not be activated with a Volume license product key. Volume-licensed systems require upgrading from a qualifying operating system. Please contact your system administrator or use a different type of key."
Case "0xC004F038" : GetLicErrDesc = "The Software Licensing Service reported that the computer could not be activated. The count reported by your Key Management Service (KMS) is insufficient. Please contact your system administrator."
Case "0xC004F039" : GetLicErrDesc = "The Software Licensing Service reported that the computer could not be activated.  The Key Management Service (KMS) is not enabled."
Case "0x4004F040" : GetLicErrDesc = "The Software Licensing Service reported that the computer was activated but the owner should verify the Product Use Rights."
Case "0xC004F041" : GetLicErrDesc = "The Software Licensing Service determined that the Key Management Service (KMS) is not activated. KMS needs to be activated. Please contact system administrator."
Case "0xC004F042" : GetLicErrDesc = "The Software Licensing Service determined that the specified Key Management Service (KMS) cannot be used."
Case "0xC004F047" : GetLicErrDesc = "The Software Licensing Service reported that the proxy policy has not been updated."
Case "0xC004F04D" : GetLicErrDesc = "The Software Licensing Service determined that the Installation ID (IID) or the Confirmation ID (CID) is invalid."
Case "0xC004F04F" : GetLicErrDesc = "The Software Licensing Service reported that license management information was not found in the licenses."
Case "0xC004F050" : GetLicErrDesc = "The Software Licensing Service reported that the product key is invalid."
Case "0xC004F051" : GetLicErrDesc = "The Software Licensing Service reported that the product key is blocked."
Case "0xC004F052" : GetLicErrDesc = "The Software Licensing Service reported that the licenses contain duplicated properties."
Case "0xC004F053" : GetLicErrDesc = "The Software Licensing Service determined that the license is invalid. The license contains an override policy that is not configured properly."
Case "0xC004F054" : GetLicErrDesc = "The Software Licensing Service reported that license management information has duplicated data."
Case "0xC004F055" : GetLicErrDesc = "The Software Licensing Service reported that the base SKU is not available."
Case "0xC004F056" : GetLicErrDesc = "The Software Licensing Service reported that the computer could not be activated using the Key Management Service (KMS)."
Case "0xC004F057" : GetLicErrDesc = "The Software Licensing Service reported that the computer BIOS is missing a required license."
Case "0xC004F058" : GetLicErrDesc = "The Software Licensing Service reported that the computer BIOS is missing a required license."
Case "0xC004F059" : GetLicErrDesc = "The Software Licensing Service reported that a license in the computer BIOS is invalid."
Case "0xC004F060" : GetLicErrDesc = "The Software Licensing Service determined that the version of the license package is invalid."
Case "0xC004F061" : GetLicErrDesc = "The Software Licensing Service determined that this specified product key can only be used for upgrading, not for clean installations."
Case "0xC004F062" : GetLicErrDesc = "The Software Licensing Service reported that a required license could not be found."
Case "0xC004F063" : GetLicErrDesc = "The Software Licensing Service reported that the computer BIOS is missing a required license."
Case "0xC004F064" : GetLicErrDesc = "The Software Licensing Service reported that the non-genuine grace period expired."
Case "0x4004F065" : GetLicErrDesc = "The Software Licensing Service reported that the application is running within the valid non-genuine grace period."
Case "0xC004F066" : GetLicErrDesc = "The Software Licensing Service reported that the genuine information property can not be set before dependent property been set."
Case "0xC004F067" : GetLicErrDesc = "The Software Licensing Service reported that the non-genuine grace period expired (type 2)."
Case "0x4004F068" : GetLicErrDesc = "The Software Licensing Service reported that the application is running within the valid non-genuine grace period (type 2)."
Case "0xC004F069" : GetLicErrDesc = "The Software Licensing Service reported that the product SKU is not found."
Case "0xC004F06A" : GetLicErrDesc = "The Software Licensing Service reported that the requested operation is not allowed."
Case "0xC004F06B" : GetLicErrDesc = "The Software Licensing Service determined that it is running in a virtual machine. The Key Management Service (KMS) is not supported in this mode."
Case "0xC004F06C" : GetLicErrDesc = "The Software Licensing Service reported that the computer could not be activated. The Key Management Service (KMS) determined that the request timestamp is invalid."
Case "0xC004F071" : GetLicErrDesc = "The Software Licensing Service reported that the plug-in manifest file is incorrect."
Case "0xC004F072" : GetLicErrDesc = "The Software Licensing Service reported that the license policies for fast query could not be found."
Case "0xC004F073" : GetLicErrDesc = "The Software Licensing Service reported that the license policies for fast query have not been loaded."
Case "0xC004F074" : GetLicErrDesc = "The Software Licensing Service reported that the computer could not be activated. No Key Management Service (KMS) could be contacted. Please see the Application Event Log for additional information."
Case "0xC004F075" : GetLicErrDesc = "The Software Licensing Service reported that the operation cannot be completed because the service is stopping."
Case "0xC004F076" : GetLicErrDesc = "The Software Licensing Service reported that the requested plug-in cannot be found."
Case "0xC004F077" : GetLicErrDesc = "The Software Licensing Service determined incompatible version of authentication data."
Case "0xC004F078" : GetLicErrDesc = "The Software Licensing Service reported that the key is mismatched."
Case "0xC004F079" : GetLicErrDesc = "The Software Licensing Service reported that the authentication data is not set."
Case "0xC004F07A" : GetLicErrDesc = "The Software Licensing Service reported that the verification could not be done."
Case "0xC004F07B" : GetLicErrDesc = "The requested operation is unavailable while the Software Licensing Service is running."
Case "0xC004F07C" : GetLicErrDesc = "The Software Licensing Service determined that the version of the computer BIOS is invalid."
Case "0xC004F200" : GetLicErrDesc = "The Software Licensing Service reported that current state is not genuine."
Case "0xC004F301" : GetLicErrDesc = "The Software Licensing Service reported that the computer could not be activated. The token-based activation challenge has expired."
Case "0xC004F302" : GetLicErrDesc = "The Software Licensing Service reported that Silent Activation failed. The Software Licensing Service reported that there are no certificates found in the system that could activate the product without user interaction."
Case "0xC004F303" : GetLicErrDesc = "The Software Licensing Service reported that the certificate chain could not be built or failed validation."
Case "0xC004F304" : GetLicErrDesc = "The Software Licensing Service reported that required license could not be found."
Case "0xC004F305" : GetLicErrDesc = "The Software Licensing Service reported that there are no certificates found in the system that could activate the product."
Case "0xC004F306" : GetLicErrDesc = "The Software Licensing Service reported that this software edition does not support token-based activation."
Case "0xC004F307" : GetLicErrDesc = "The Software Licensing Service reported that the computer could not be activated. Activation data is invalid."
Case "0xC004F308" : GetLicErrDesc = "The Software Licensing Service reported that the computer could not be activated. Activation data is tampered."
Case "0xC004F309" : GetLicErrDesc = "The Software Licensing Service reported that the computer could not be activated. Activation challenge and response do not match."
Case "0xC004F30A" : GetLicErrDesc = "The Software Licensing Service reported that the computer could not be activated. The certificate does not match the conditions in the license."
Case "0xC004F30B" : GetLicErrDesc = "The Software Licensing Service reported that the inserted smartcard could not be used to activate the product."
Case "0xC004F30C" : GetLicErrDesc = "The Software Licensing Service reported that the token-based activation license content is invalid."
Case "0xC004F30D" : GetLicErrDesc = "The Software Licensing Service reported that the computer could not be activated. The thumbprint is invalid."
Case "0xC004F30E" : GetLicErrDesc = "The Software Licensing Service reported that the computer could not be activated. The thumbprint does not match any certificate."
Case "0xC004F30F" : GetLicErrDesc = "The Software Licensing Service reported that the computer could not be activated. The certificate does not match the criteria specified in the issuance license."
Case "0xC004F310" : GetLicErrDesc = "The Software Licensing Service reported that the computer could not be activated. The certificate does not match the trust point identifier (TPID) specified in the issuance license."
Case "0xC004F311" : GetLicErrDesc = "The Software Licensing Service reported that the computer could not be activated. A soft token cannot be used for activation."
Case "0xC004F312" : GetLicErrDesc = "The Software Licensing Service reported that the computer could not be activated. The certificate cannot be used because its private key is exportable."
Case "0xC004F313" : GetLicErrDesc = "The Software Licensing Service reported that the CNG encryption library could not be loaded.  The current certificate may not be available on this version of Windows."
Case "0xC004FC03" : GetLicErrDesc = "A networking problem has occurred while activating your copy of Windows."
Case "0x4004FC04" : GetLicErrDesc = "The Software Licensing Service reported that the application is running within the timebased validity period."
Case "0x4004FC05" : GetLicErrDesc = "The Software Licensing Service reported that the application has a perpetual grace period."
Case "0x4004FC06" : GetLicErrDesc = "The Software Licensing Service reported that the application is running within the valid extended grace period."
Case "0xC004FC07" : GetLicErrDesc = "The Software Licensing Service reported that the validity period expired."
Case "0xC004FE00" : GetLicErrDesc = "The Software Licensing Service reported that activation is required to recover from tampering of SL Service trusted store."
Case "0xC004D101" : GetLicErrDesc = "The security processor reported an initialization error."
Case "0x8004D102" : GetLicErrDesc = "The security processor reported that the machine time is inconsistent with the trusted time."
Case "0xC004D103" : GetLicErrDesc = "The security processor reported that an error has occurred."
Case "0xC004D104" : GetLicErrDesc = "The security processor reported that invalid data was used."
Case "0xC004D105" : GetLicErrDesc = "The security processor reported that the value already exists."
Case "0xC004D107" : GetLicErrDesc = "The security processor reported that an insufficient buffer was used."
Case "0xC004D108" : GetLicErrDesc = "The security processor reported that invalid data was used."
Case "0xC004D109" : GetLicErrDesc = "The security processor reported that an invalid call was made."
Case "0xC004D10A" : GetLicErrDesc = "The security processor reported a version mismatch error."
Case "0x8004D10B" : GetLicErrDesc = "The security processor cannot operate while a debugger is attached."
Case "0xC004D301" : GetLicErrDesc = "The security processor reported that the trusted data store was tampered."
Case "0xC004D302" : GetLicErrDesc = "The security processor reported that the trusted data store was rearmed."
Case "0xC004D303" : GetLicErrDesc = "The security processor reported that the trusted store has been recreated."
Case "0xC004D304" : GetLicErrDesc = "The security processor reported that entry key was not found in the trusted data store."
Case "0xC004D305" : GetLicErrDesc = "The security processor reported that the entry key already exists in the trusted data store."
Case "0xC004D306" : GetLicErrDesc = "The security processor reported that the entry key is too big to fit in the trusted data store."
Case "0xC004D307" : GetLicErrDesc = "The security processor reported that the maximum allowed number of re-arms has been exceeded.  You must re-install the OS before trying to re-arm again."
Case "0xC004D308" : GetLicErrDesc = "The security processor has reported that entry data size is too big to fit in the trusted data store."
Case "0xC004D309" : GetLicErrDesc = "The security processor has reported that the machine has gone out of hardware tolerance."
Case "0xC004D30A" : GetLicErrDesc = "The security processor has reported that the secure timer already exists."
Case "0xC004D30B" : GetLicErrDesc = "The security processor has reported that the secure timer was not found."
Case "0xC004D30C" : GetLicErrDesc = "The security processor has reported that the secure timer has expired."
Case "0xC004D30D" : GetLicErrDesc = "The security processor has reported that the secure timer name is too long."
Case "0xC004D30E" : GetLicErrDesc = "The security processor reported that the trusted data store is full."
Case "0xC004D401" : GetLicErrDesc = "The security processor reported a system file mismatch error."
Case "0xC004D402" : GetLicErrDesc = "The security processor reported a system file mismatch error."
Case "0xC004D501" : GetLicErrDesc = "The security processor reported an error with the kernel data."
Case Else : GetLicErrDesc = ""
End Select
End Function 'GetLicErrDesc



'Use WMI to query the license details for installed Office products
'The license details of the products are mapped to the OSPP data based on the "Configuration Productname"

Sub OsppCollect
    Dim iArpCnt, iPosMaster, iLicCnt, iLicPos, iLic, iVPCnt, iVersionMajor
    Dim sText, sOsppLicenses, sPossibleSkus, sPossibleSkusFull, sTmp, sXmlLogLine
    Dim arrLicData

    If CheckArray(arrArpProducts) Then
        For iArpCnt = 0 To UBound(arrArpProducts)
            sOsppLicenses = ""
            sXmlLogLine = ""
            'Get the link to the master array
            iPosMaster = GetArrayPosition(arrMaster,arrArpProducts(iArpCnt,ARP_CONFIGPRODUCTCODE))
            If iPosMaster > -1 Then
                iLicCnt = 0
                'OSPP is first used by O14
                iVersionMajor = CInt(GetVersionMajor(arrMaster(iPosMaster, COL_PRODUCTCODE)))
                If (arrMaster(iPosMaster, COL_ISOFFICEPRODUCT)) AND iVersionMajor > 13 Then
                    'Get the Config Productname
                    sText = "" : sText = arrArpProducts(iArpCnt,COL_CONFIGNAME)
                    If InStr(sText,".") > 0 Then sText = Mid(sText, InStr(sText, ".") + 1)
                    'A list of all possible licenses this product can use is stored in sPossibleSkus
                    sPossibleSkus = ""
                    sXmlLogLine = ""
                    'Loop all licenses
                    iLicCnt = GetLicCnt(iPosMaster, iVersionMajor, sText, sPossibleSkus, sPossibleSkusFull)
                    If iLicCnt > 0 Then
                        sOsppLicenses = "Possible Licenses;" & Mid(sPossibleSkus, 3)
                        'List installed licenses (ProductKeyID <> "")
                        iLicPos = 0
                        sXmlLogLine = ""
                        For iLic = 1 To iLicCnt
                            arrLicData = GetLicenseData(iPosMaster,CInt(GetVersionMajor(arrMaster(iPosMaster,COL_PRODUCTCODE))),sText,iLicPos,sPossibleSkusFull)
                            If InStr(sPossibleSkusFull, arrLicData(OSPP_NAME)) > 0 Then AddLicXmlString sXmlLogLine, arrLicData
                            If arrLicData(OSPP_PRODUCTKEYID) <> "" Then AddLicTxtString sOsppLicenses, arrLicData
                        Next 'iLic
                        arrMaster(iPosMaster, COL_OSPPLICENSEXML) = sXmlLogLine
                        arrMaster(iPosMaster, COL_OSPPLICENSE) = sOsppLicenses
                    End If 'iLicCnt > 0
                End If 'VersionMajor > 14
            End If 'iPosMaster > -1
        Next 'iArpCnt
    End If 'arrArpProducts

    If UBound(arrVirt2Products) > -1 Then
        For iVPCnt = 0 To UBound(arrVirt2Products, 1)
            sText = "" : sText = arrVirt2Products(iVPCnt, VIRTPROD_CONFIGNAME)
            'sText = Replace(sText, "O365", "")
            sText = Replace(sText, "Volume", "")
            sText = Replace(sText, "Retail", "")
            
            iVersionMajor = 15
            'A list of all possible licenses this product can use is stored in sPossibleSkus
            sPossibleSkus = ""
            sXmlLogLine = ""
            'Loop all licenses
            iLicCnt = GetLicCnt(-1, iVersionMajor, sText, sPossibleSkus, sPossibleSkusFull)
            If iLicCnt > 0 Then
                sOsppLicenses = "Possible Licenses;" & Mid(sPossibleSkus, 3)
                'List installed licenses (ProductKeyID <> "")
                iLicPos = 0
                sXmlLogLine = ""
                For iLic = 1 To iLicCnt
                    arrLicData = GetLicenseData(-1, iVersionMajor, sText, iLicPos, sPossibleSkusFull)
                    If InStr(sPossibleSkusFull, arrLicData(OSPP_NAME)) > 0 Then AddLicXmlString sXmlLogLine, arrLicData
                    If arrLicData(OSPP_PRODUCTKEYID) <> "" Then AddLicTxtString sOsppLicenses, arrLicData
                Next 'iLic
                arrVirt2Products(iVPCnt, VIRTPROD_OSPPLICENSEXML) = sXmlLogLine
                arrVirt2Products(iVPCnt, VIRTPROD_OSPPLICENSE) = sOsppLicenses
            End If 'iLicCnt
        Next 'iVPCnt
    End If 'arrVirt2Products

    If UBound(arrVirt3Products) > -1 Then
        For iVPCnt = 0 To UBound(arrVirt3Products, 1)
            sText = "" : sText = arrVirt3Products(iVPCnt, VIRTPROD_CONFIGNAME)
            'sText = Replace(sText, "O365", "")
            sText = Replace(sText, "Volume", "")
            sText = Replace(sText, "Retail", "")
            
            iVersionMajor = 16
            'A list of all possible licenses this product can use is stored in sPossibleSkus
            sPossibleSkus = ""
            sXmlLogLine = ""
            'Loop all licenses
            iLicCnt = GetLicCnt(-1, iVersionMajor, sText, sPossibleSkus, sPossibleSkusFull)
            If iLicCnt > 0 Then
                sOsppLicenses = "Possible Licenses;" & Mid(sPossibleSkus, 3)
                'List installed licenses (ProductKeyID <> "")
                iLicPos = 0
                sXmlLogLine = ""
                For iLic = 1 To iLicCnt
                    arrLicData = GetLicenseData(-1, iVersionMajor, sText, iLicPos, sPossibleSkusFull)
                    If InStr(sPossibleSkusFull, arrLicData(OSPP_NAME)) > 0 Then AddLicXmlString sXmlLogLine, arrLicData
                    If arrLicData(OSPP_PRODUCTKEYID) <> "" Then AddLicTxtString sOsppLicenses, arrLicData
                Next 'iLic
                arrVirt3Products(iVPCnt, VIRTPROD_OSPPLICENSEXML) = sXmlLogLine
                arrVirt3Products(iVPCnt, VIRTPROD_OSPPLICENSE) = sOsppLicenses
            End If 'iLicCnt
        Next 'iVPCnt
    End If 'arrVirt2Products
End Sub 'OsppCollect

'=======================================================================================================
Function GetLicenseStateString (iState)
    Dim sTmp

    sTmp = ""
    Select Case iState
    Case 0 : sTmp = "UNLICENSED"
    Case 1 : sTmp = "LICENSED"
    Case 2 : sTmp = "OOB GRACE"
    Case 3 : sTmp = "OOT GRACE"
    Case 4 : sTmp = "NON GENUINE GRACE"
    Case 5 : sTmp = "NOTIFICATIONS"
    Case Else : sTmp = "UNKNOWN"
    End Select

    GetLicenseStateString = sTmp
End Function 'GetLicenseStateString

'=======================================================================================================
Sub AddLicXmlString (sXmlLogLine, arrLicData)
    Dim sTmp
    
    If NOT sXmlLogLine = "" Then sXmlLogLine = sXmlLogLine & vbCrLf
    If arrLicData(OSPP_PRODUCTKEYID) <> "" Then sTmp = "TRUE" Else sTmp = "FALSE"
    sXmlLogLine = sXmlLogLine & "<License"
    sXmlLogLine = sXmlLogLine & " IsActive=" & chr(34) & sTmp & chr(34)
    sXmlLogLine = sXmlLogLine & " Name=" & chr(34) & arrLicData(OSPP_NAME) & chr(34)
    sXmlLogLine = sXmlLogLine & " Description=" & chr(34) & arrLicData(OSPP_DESCRIPTION) & chr(34)
    sXmlLogLine = sXmlLogLine & " Family=" & chr(34) & arrLicData(OSPP_LICENSEFAMILY) & chr(34)
    sXmlLogLine = sXmlLogLine & " Status=" & chr(34) & arrLicData(OSPP_LICENSESTATUS) & chr(34)
    sXmlLogLine = sXmlLogLine & " StatusString=" & chr(34) & GetLicenseStateString(arrLicData(OSPP_LICENSESTATUS)) & chr(34)
    sXmlLogLine = sXmlLogLine & " StatusCode=" & chr(34) & "0x" & Hex(arrLicData(OSPP_LICENSESTATUSREASON)) & chr(34)
    sXmlLogLine = sXmlLogLine & " StatusDescription=" & chr(34) & GetLicErrDesc(Hex(arrLicData(OSPP_LICENSESTATUSREASON))) & chr(34)
    sXmlLogLine = sXmlLogLine & " PartialProductkey=" & chr(34) & arrLicData(OSPP_PARTIALPRODUCTKEY) & chr(34)
    sXmlLogLine = sXmlLogLine & " ApplicationID=" & chr(34) & arrLicData(OSPP_APPLICATIONID) & chr(34)
    sXmlLogLine = sXmlLogLine & " ProductKeyID=" & chr(34) & arrLicData(OSPP_PRODUCTKEYID) & chr(34)
    sXmlLogLine = sXmlLogLine & " ProductID=" & chr(34) & arrLicData(OSPP_PRODUCTKEYID2) & chr(34)
    sXmlLogLine = sXmlLogLine & " MachineKey=" & chr(34) & arrLicData(OSPP_MACHINEKEY) & chr(34)
    sXmlLogLine = sXmlLogLine & " ActivationType=" & chr(34) & arrLicData(OSPP_ACTIVATIONTYPE) & chr(34)
    sXmlLogLine = sXmlLogLine & " SkuID=" & chr(34) & arrLicData(OSPP_ID) & chr(34)
    sXmlLogLine = sXmlLogLine & " KmsServer=" & chr(34) & arrLicData(OSPP_DISCOVEREDKEYMANAGEMENTSERVICEMACHINENAME) & chr(34)
    sXmlLogLine = sXmlLogLine & " KmsPort=" & chr(34) & arrLicData(OSPP_DISCOVEREDKEYMANAGEMENTSERVICEMACHINEPORT) & chr(34)
    sXmlLogLine = sXmlLogLine & " ActivationObjectName=" & chr(34) & arrLicData(OSPP_ADACTIVATIONOBJECTNAME) & chr(34)
    sXmlLogLine = sXmlLogLine & " ActivationObjectDN=" & chr(34) & arrLicData(OSPP_ADACTIVATIONOBJECTDN) & chr(34)
    sXmlLogLine = sXmlLogLine & " ActivationObjectExtendedPID=" & chr(34) & arrLicData(OSPP_ADACTIVATIONCSVLPID) & chr(34)
    sXmlLogLine = sXmlLogLine & " ActivationObjectActivationID=" & chr(34) & arrLicData(OSPP_ADACTIVATIONCSVLSKUID) & chr(34)
    sXmlLogLine = sXmlLogLine & " RemainingGracePeriod=" & chr(34) & CInt(arrLicData(OSPP_GRACEPERIODREMAINING) / 1440) & chr(34)
    sXmlLogLine = sXmlLogLine & " ActivationInterval=" & chr(34) & arrLicData(OSPP_VLACTIVATIONINTERVAL) / 60 & chr(34)
    sXmlLogLine = sXmlLogLine & " RenewalInterval=" & chr(34) & arrLicData(OSPP_VLRENEWALINTERVAL) / 1440 & chr(34) & " />"
End Sub 'AddLicXmlString

'=======================================================================================================
Sub AddLicTxtString (sOsppLicenses, arrLicData)
    Dim sTmp
    
    sOsppLicenses = sOsppLicenses & "#;#" & "Active License;" & arrLicData(OSPP_NAME)
    sOsppLicenses = sOsppLicenses & "#;#" & "Description;" & arrLicData(OSPP_DESCRIPTION)
    sOsppLicenses = sOsppLicenses & "#;#" & "License Family;" & arrLicData(OSPP_LICENSEFAMILY)
    sTmp = GetLicenseStateString(arrLicData(OSPP_LICENSESTATUS))
    sTmp = arrLicData(OSPP_LICENSESTATUS) & " - " & sTmp & " (Error Code: 0x" & Hex(arrLicData(OSPP_LICENSESTATUSREASON)) 
    sTmp = sTmp & " - " & GetLicErrDesc(Hex(arrLicData(OSPP_LICENSESTATUSREASON))) & ")"
    sOsppLicenses = sOsppLicenses & "#;#" & "License Status;" & sTmp
    sOsppLicenses = sOsppLicenses & "#;#" & "Partial ProductKey;" & arrLicData(OSPP_PARTIALPRODUCTKEY)
    sOsppLicenses = sOsppLicenses & "#;#" & "ApplicationID;" & arrLicData(OSPP_APPLICATIONID)
    sOsppLicenses = sOsppLicenses & "#;#" & "SKU ID;" & arrLicData(OSPP_ID)
    sOsppLicenses = sOsppLicenses & "#;#" & "ProductKeyID;" & arrLicData(OSPP_PRODUCTKEYID)
    sOsppLicenses = sOsppLicenses & "#;#" & "Product ID;" & arrLicData(OSPP_PRODUCTKEYID2)
    sOsppLicenses = sOsppLicenses & "#;#" & "Machine Key;" & arrLicData(OSPP_MACHINEKEY)
    Select Case arrLicData (OSPP_ACTIVATIONTYPE)
    Case 0 ' ALL
        sOsppLicenses = sOsppLicenses & "#;#" & "Activation Type;" & arrLicData(OSPP_ACTIVATIONTYPE) & " (ALL)"
        sOsppLicenses = sOsppLicenses & "#;#" & "KMS Server;" & arrLicData(OSPP_DISCOVEREDKEYMANAGEMENTSERVICEMACHINENAME)
        sOsppLicenses = sOsppLicenses & "#;#" & "KMS Port;" & arrLicData(OSPP_DISCOVEREDKEYMANAGEMENTSERVICEMACHINEPORT)
        sOsppLicenses = sOsppLicenses & "#;#" & "Activation Object Name;" & arrLicData(OSPP_ADACTIVATIONOBJECTNAME)
        sOsppLicenses = sOsppLicenses & "#;#" & "Activation Object DN;" & arrLicData(OSPP_ADACTIVATIONOBJECTDN)
        sOsppLicenses = sOsppLicenses & "#;#" & "Activation Object Extended PID;" & arrLicData(OSPP_ADACTIVATIONCSVLPID)
        sOsppLicenses = sOsppLicenses & "#;#" & "Activation Object Activation ID;" & arrLicData(OSPP_ADACTIVATIONCSVLSKUID)
        sOsppLicenses = sOsppLicenses & "#;#" & "Licensed Days Remaining;" & CInt(arrLicData(OSPP_GRACEPERIODREMAINING) / 1440)
        sOsppLicenses = sOsppLicenses & "#;#" & "VL Activation Interval;" & arrLicData(OSPP_VLACTIVATIONINTERVAL) / 60 & " hours"
        sOsppLicenses = sOsppLicenses & "#;#" & "VL Renewal Interval;" & arrLicData(OSPP_VLRENEWALINTERVAL) / 1440 & " days"
    Case 1 ' AD
        sOsppLicenses = sOsppLicenses & "#;#" & "Activation Type;" & arrLicData(OSPP_ACTIVATIONTYPE) & " (AD)"
        sOsppLicenses = sOsppLicenses & "#;#" & "Activation Object Name;" & arrLicData(OSPP_ADACTIVATIONOBJECTNAME)
        sOsppLicenses = sOsppLicenses & "#;#" & "Activation Object DN;" & arrLicData(OSPP_ADACTIVATIONOBJECTDN)
        sOsppLicenses = sOsppLicenses & "#;#" & "Activation Object Extended PID;" & arrLicData(OSPP_ADACTIVATIONCSVLPID)
        sOsppLicenses = sOsppLicenses & "#;#" & "Activation Object Activation ID;" & arrLicData(OSPP_ADACTIVATIONCSVLSKUID)
        sOsppLicenses = sOsppLicenses & "#;#" & "Licensed Days Remaining;" & CInt(arrLicData(OSPP_GRACEPERIODREMAINING) / 1440)
        sOsppLicenses = sOsppLicenses & "#;#" & "VL Activation Interval;" & arrLicData(OSPP_VLACTIVATIONINTERVAL) / 60 & " hours"
        sOsppLicenses = sOsppLicenses & "#;#" & "VL Renewal Interval;" & arrLicData(OSPP_VLRENEWALINTERVAL) / 1440 & " days"
    Case 2 ' KMS
        sOsppLicenses = sOsppLicenses & "#;#" & "Activation Type;" & arrLicData(OSPP_ACTIVATIONTYPE) & " (KMS)"
        sOsppLicenses = sOsppLicenses & "#;#" & "KMS Server;" & arrLicData(OSPP_DISCOVEREDKEYMANAGEMENTSERVICEMACHINENAME)
        sOsppLicenses = sOsppLicenses & "#;#" & "KMS Port;" & arrLicData(OSPP_KEYMANAGEMENTSERVICEPORT)
        sOsppLicenses = sOsppLicenses & "#;#" & "Licensed Days Remaining;" & CInt(arrLicData(OSPP_GRACEPERIODREMAINING) / 1440)
        sOsppLicenses = sOsppLicenses & "#;#" & "VL Activation Interval;" & arrLicData(OSPP_VLACTIVATIONINTERVAL) / 60 & " hours"
        sOsppLicenses = sOsppLicenses & "#;#" & "VL Renewal Interval;" & arrLicData(OSPP_VLRENEWALINTERVAL) / 1440 & " days"
    Case 3 ' Token
        sOsppLicenses = sOsppLicenses & "#;#" & "Activation Type;" & arrLicData(OSPP_ACTIVATIONTYPE) & " (Token)"
    Case Else
        sOsppLicenses = sOsppLicenses & "#;#" & "Activation Type;" & arrLicData(OSPP_ACTIVATIONTYPE) & " (?)"
        If arrLicData(OSPP_GRACEPERIODREMAINING) <> 0 Then _
            sOsppLicenses = sOsppLicenses & "#;#" & "Remaining Grace Period;" & CInt(arrLicData(OSPP_GRACEPERIODREMAINING) / 1440) & " days"
    End Select
End Sub 'AddLicTxtString

'=======================================================================================================
'Module Productslist
'=======================================================================================================
Sub FindAllProducts
    Dim sActiveSub, sErrHnd 
    Dim AllProducts, ProdX, key
    Dim arrKeys, Arr, arrTmpSids()
    Dim sSid
    Dim i
    sActiveSub = "FindAllProducts" : sErrHnd = "" 
    On Error Resume Next
    
    'Cache ARP entries
    If RegEnumKey (HKLM, REG_ARP, arrKeys) Then
        For Each key in arrKeys
            If NOT dicArp.Exists(key) Then dicArp.Add key, REG_ARP & key
        Next 'key
    End If
    
    'Build an array of all applications registered to Windows Installer
    'Iterate products depending of available WI version
    Select Case iWiVersionMajor
    Case 3, 4, 5, 6
        sErrHnd = "_ErrorHandler3x" : Err.Clear
        
        Set AllProducts = oMsi.ProductsEx("",USERSID_EVERYONE,MSIINSTALLCONTEXT_ALL) : CheckError sActiveSub,sErrHnd 
        If CheckObject(AllProducts) Then WritePLArrayEx 3, arrAllProducts, AllProducts, Null, Null
    Case 2
        'Only available for backwards compatibility reasons
        'Will use direct registry reading. Reuse logic from FindProducts_ErrorHandler
        For i = 1 to 4
            'Note that 'i=3' is not a valid context which is simply ignored in FindRegProducts.
            FindRegProducts i
        Next 'i

    Case Else
        'Not Handled - Not Supported
    End Select
    
    GetC2Rv2VersionsActive
    FindRegProducts MSIINSTALLCONTEXT_C2RV2
    FindRegProducts MSIINSTALLCONTEXT_C2RV3

    InitMasterArray
    Set dicProducts = CreateObject("Scripting.Dictionary")
    For i = 0 To UBound(arrMaster)
        arrMaster(i,COL_ISOFFICEPRODUCT) = IsOfficeProduct(arrMaster(i,COL_PRODUCTCODE))
        If NOT dicProducts.Exists(arrMaster(i,COL_PRODUCTCODE)) Then dicProducts.Add arrMaster(i,COL_PRODUCTCODE),arrMaster(i,COL_USERSID)
    Next 'i
    'Build an array of all virtualized applications (Click2Run)
    FindV1VirtualizedProducts
    FindV2VirtualizedProducts
    FindV3VirtualizedProducts
End Sub
'=======================================================================================================

Sub FindAllProducts_ErrorHandler3x
    
    Cachelog LOGPOS_REVITEM,LOGHEADING_NONE,ERR_CATEGORYERROR,ERR_PRODUCTSEXALL & DOT & " Error Details: " & _
    Err.Source & " " & Hex( Err ) & ": " & Err.Description
    
    On Error Resume Next
    
    FindProducts "",MSIINSTALLCONTEXT_MACHINE
    
    FindProducts USERSID_EVERYONE,MSIINSTALLCONTEXT_USERUNMANAGED
    
    FindProducts USERSID_EVERYONE,MSIINSTALLCONTEXT_USERMANAGED
End Sub
'=======================================================================================================

Sub FindProducts(sSid, iContext)
    Dim sActiveSub, sErrHnd 
    Dim Arr, Products
    sActiveSub = "FindProducts" : sErrHnd = ""
    On Error Resume Next

    sErrHnd = iContext : Err.Clear
    Set Products = oMsi.ProductsEx("",sSid,iContext) : CheckError sActiveSub,sErrHnd
    Select Case iContext
        Case MSIINSTALLCONTEXT_MACHINE :     If CheckObject(Products) Then WritePLArrayEx 3,arrMProducts,Products,iContext,sSid
        Case MSIINSTALLCONTEXT_USERUNMANAGED : If CheckObject(Products) Then WritePLArrayEx 3,arrUUProducts,Products,iContext,sSid
        Case MSIINSTALLCONTEXT_USERMANAGED : If CheckObject(Products) Then WritePLArrayEx 3,arrUMProducts,Products,iContext,sSid
        Case Else
    End Select
End Sub
'=======================================================================================================

Sub FindProducts_ErrorHandler (iContext)
    Dim sSid, Arr
    Dim n
    
    Cachelog LOGPOS_REVITEM,LOGHEADING_NONE,ERR_CATEGORYERROR,"ProductsEx for " & GetContextString(iContext) & " failed. Error Details: " & _
    Err.Source & " " & Hex( Err ) & ": " & Err.Description
    
    On Error Resume Next

    FindRegProducts iContext
End Sub
'=======================================================================================================

'General entry point for getting a list of products from the registry.
Sub FindRegProducts (iContext)
    Dim arrKeys, arrTmpKeys
    Dim sSid, sSubKeyName, sUMSids, sProd, sTmpProd, sVer
    Dim hDefKey
    Dim n, iProdCnt, iProdFind, iProdTotal
    On Error Resume Next

    Select Case iContext

    Case MSIINSTALLCONTEXT_MACHINE
        sSid = MACHINESID
        hDefKey = GetRegHive(iContext, sSid, False)
        sSubKeyName = GetRegConfigKey("", iContext, sSid, False)
        FindRegProductsEx hDefKey, sSubKeyName, sSid, iContext, arrMProducts
    
    Case MSIINSTALLCONTEXT_C2RV2
        sSid = MACHINESID
        hDefKey = GetRegHive(iContext, sSid, False)
        sSubKeyName = REG_OFFICE & "15.0" & REG_C2RVIRT_HKLM & REG_GLOBALCONFIG & "S-1-5-18\Products\"
        FindRegProductsEx hDefKey, sSubKeyName, sSid, iContext, arrMVProducts
    
    Case MSIINSTALLCONTEXT_C2RV3
        sSid = MACHINESID
        hDefKey = GetRegHive(iContext, sSid, False)
        sSubKeyName = REG_OFFICE & "ClickToRun\REGISTRY\MACHINE\" & REG_GLOBALCONFIG & "S-1-5-18\Products\"
        FindRegProductsEx hDefKey, sSubKeyName, sSid, iContext, arrMVProducts
    
    Case MSIINSTALLCONTEXT_USERUNMANAGED
        sUMSids = ""
        If CheckArray (arrUMSids) Then sUMSids = Join(arrUMSids)
        If CheckArray (arrUUSids) Then
            iProdTotal = -1
            For iProdFind = 0 To 1
                For n = 0 To UBound(arrUUSids)
                    sSid = arrUUSids(n)
                    hDefKey = GetRegHive(iContext, sSid, False)
                    sSubKeyName = GetRegConfigKey("", iContext, sSid, False)
                    If InStr(sUMSids, sSid)>0 Then
                        'Current SID has installed managed per-user products
                        'Create a string list with products
                        sTmpProd = ""
                        If NOT sUMSids="" Then
                            For iProdCnt = 0 To UBound(arrUMProducts)
                                If arrUMProducts(iProdCnt, COL_USERSID)=sSid Then sTmpProd = sTmpProd & arrUMProducts(iProdCnt, COL_PRODUCTCODE)
                            Next 'iProdCnt
                        End If
                        If RegEnumKey(hDefKey, sSubKeyName, arrTmpKeys) Then
                            ReDim arrKeys(-1)
                            For iProdCnt = 0 To UBound(arrTmpKeys)
                                If Not InStr(sTmpProd, GetExpandedGuid(arrTmpKeys(iProdCnt)))>0 Then
                                    ReDim Preserve arrKeys(UBound(arrKeys)+1)
                                    arrKeys(UBound(arrKeys))=arrTmpKeys(iProdCnt)
                                End If 'Not InStr(sTmpProd, arrTmpKeys(iProdCnt))>0
                            Next 'iProdCnt
                            If iProdFind = 0 Then
                                iProdTotal = iProdTotal + UBound(arrKeys) + 1
                            Else
                                WritePLArrayEx 0, arrUUProducts, arrKeys, iContext, sSid
                            End If
                        End If 'RegEnumKey(hDefKey, sSubKeyName, arrTmpKeys)
                    Else
                        'No conflict with managed per user products
                        If RegEnumKey(hDefKey,sSubKeyName,arrKeys) Then
                            If iProdFind = 0 Then
                                iProdTotal = iProdTotal + UBound(arrKeys) + 1
                            Else
                                WritePLArrayEx 0,arrUUProducts,arrKeys,iContext,sSid
                            End If
                         End If 'RegEnumKey(hDefKey,sSubKeyName,arrKeys)
                    End If 'InStr(sUMSid,sSid)>0
                Next 'n
                If iProdFind = 0 Then ReDim arrUUProducts(iProdFind,UBOUND_MASTER)
            Next 'iProdFind
        End If 'CheckArray

    Case MSIINSTALLCONTEXT_USERMANAGED
        iProdCnt = -1
        If CheckArray (arrUMSids) Then
            'Determine number of products
            For n = 0 To UBound(arrUMSids)
                sSid = arrUMSids (n)
                hDefKey = GetRegHive(iContext,sSid,False)
                sSubKeyName = GetRegConfigKey("",iContext,sSid,False)
                If RegEnumKey(hDefKey,sSubKeyName,arrTmpKeys) Then
                    iProdCnt = iProdCnt + UBound(arrTmpKeys) + 1
                End If
            Next 'n
            ReDim arrUMProducts(iProdCnt,UBOUND_MASTER)
                
            'Add the products to array
            For n = 0 To UBound(arrUMSids)
                sSid = arrUMSids (n)
                hDefKey = GetRegHive(iContext,sSid,False)
                sSubKeyName = GetRegConfigKey("",iContext,sSid,False)
                
                If RegEnumKey(hDefKey,sSubKeyName,arrTmpKeys) Then WritePLArrayEx 0,arrUMProducts,arrTmpKeys,iContext,sSid
            Next 'n
        End If 'CheckArray

    Case Else
    End Select
    
End Sub
'=======================================================================================================

'Usually called from FindRegProducts
Sub FindRegProductsEx (hDefKey,sSubKeyName,sSid,iContext,Arr)
    Dim arrKeys
    On Error Resume Next

    If RegEnumKey(hDefKey,sSubKeyName,arrKeys) Then 
        WritePLArrayEx 0,Arr,arrKeys,iContext,sSid
    End If 'RegKeyExists
End Sub
'=======================================================================================================

'Starting with O14 virtualized applications are available as part of "Click2Run"
'These SKU's are not fully covered by native .msi installations

Sub FindV1VirtualizedProducts

Dim Key, sValue, VProd
Dim arrKeys
Dim dicVirtProd
Dim iVCnt

'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\14.0\CVH
'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\<ProductCode> -> InstallLocation = "Virtualized Application"
Set dicVirtProd = CreateObject("Scripting.Dictionary")

For Each Key in dicArp.Keys
    If Len(Key) = 38 Then
        If IsOfficeProduct(Key) Then
            If RegReadValue(HKLM, REG_ARP & Key, "CVH", sValue, "REG_DWORD") Then
                If sValue = "1" Then
                    If NOT dicVirtProd.Exists(Key) Then dicVirtProd.Add Key,Key
                End If
            End If
        End If
    End If
Next 'Key

'Fill the virtual products array 
If dicVirtProd.Count > 0 Then
    ReDim arrVirtProducts(dicVirtProd.Count -1, UBOUND_VIRTPROD)
    iVCnt = 0
    For Each VProd in dicVirtProd.Keys
        'ProductCode
        arrVirtProducts(iVCnt,COL_PRODUCTCODE) = VProd
        
        'ProductName
        If RegReadValue(HKLM, REG_ARP & VProd, "DisplayName", sValue, "REG_SZ") Then
            arrVirtProducts(iVCnt,COL_PRODUCTNAME) = sValue
        End If

        'DisplayVersion
        If RegReadValue(HKLM, REG_ARP & VProd, "DisplayVersion", sValue, "REG_SZ") Then
            arrVirtProducts(iVCnt, VIRTPROD_PRODUCTVERSION) = sValue
        End If

        If IsOfficeProduct(VProd) Then
            'SP level
            If NOT fInitArrProdVer Then InitProdVerArrays
            arrVirtProducts(iVCnt,VIRTPROD_SPLEVEL) = OVersionToSpLevel (VProd,GetVersionMajor(VProd),arrVirtProducts(iVCnt,VIRTPROD_PRODUCTVERSION)) 

            'Architecture (bitness)
            arrVirtProducts(iVCnt,VIRTPROD_BITNESS) = "x86"
            If Mid(VProd,21,1)="1" Then arrVirtProducts(iVCnt,VIRTPROD_BITNESS) = "x64"
        End If 'IsOfficeProduct

        iVCnt = iVCnt + 1
    Next 'VProd
End If 'dicVirtProd > 0

End Sub 'FindV1VirtualizedProducts
'=======================================================================================================

'-------------------------------------------------------------------------------
'   FindV2VirtualizedProducts
'
'   Locate virtualized C2R_v2 products
'-------------------------------------------------------------------------------
Sub FindV2VirtualizedProducts
    Dim ArpItem, VProd, ConfigProd
    Dim component, culture, child, name, key, subKey, prod
    Dim sSubKeyName, sConfigName, sValue, sChild
    Dim sCurKey, sCurKeyL0, sCurKeyL1, sCurKeyL2, sCurKeyL3, sKey
    Dim sVersion, sVersionFallback, sFileName, sKeyComponents
    Dim sActiveConfiguration, sProd, sCult
    Dim iVCnt
    Dim dicVirt2Prod, dicVirt2ConfigID
    Dim arrKeys, arrConfigProducts, arrVersion, arrCultures, arrChildPackages, arrSubKeys
    Dim arrNames, arrTypes, arrScenario, arrKeyComponents, arrComponentData
    Dim fUninstallString, fIsSP1, fC2RPolEnabled

	Const REG_C2R				    = "SOFTWARE\Microsoft\Office\15.0\ClickToRun\"
   	Const REG_C2RCONFIGURATION	    = "SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration\"
   	Const REG_C2RPROPERTYBAG        = "SOFTWARE\Microsoft\Office\15.0\ClickToRun\propertyBag\"
   	Const REG_C2RSCENARIO           = "SOFTWARE\Microsoft\Office\15.0\ClickToRun\Scenario\"
   	Const REG_C2RUPDATES            = "SOFTWARE\Microsoft\Office\15.0\ClickToRun\Updates\"
	Const REG_C2RPRODUCTIDS		    = "SOFTWARE\Microsoft\Office\15.0\ClickToRun\ProductReleaseIDs\"
	Const REG_C2RUPDATEPOL		    = "SOFTWARE\Policies\Microsoft\Office\15.0\Common\OfficeUpdate\"

    On Error Resume Next

    Set dicVirt2Prod = CreateObject ("Scripting.Dictionary")
    Set dicVirt2ConfigID = CreateObject("Scripting.Dictionary")
    fIsSP1 = False
    fC2RPolEnabled = False
    ' extend ARP dic to contain virt2 references
    sKey = REG_C2R & "REGISTRY\MACHINE\" & REG_ARP
    If RegEnumKey (HKLM, sKey, arrKeys) Then
        For Each key in arrKeys
            If NOT dicArp.Exists(key) Then dicArp.Add key, sKey & key
        Next 'key
    End If
    
    'Integration PackageGUID
    If RegReadValue(HKLM, REG_C2R, "PackageGUID", sValue, REG_SZ) Then
    	sPackageGuid = "{" & sValue & "}"
        dicC2RPropV2.Add STR_REGPACKAGEGUID, sValue
        dicC2RPropV2.Add STR_PACKAGEGUID, sPackageGuid
    Else
        If RegEnumKey(HKLM, REG_C2R & "appvMachineRegistryStore\Integration\Packages\", arrKeys) Then
    	    sPackageGuid = sValue
    	    sValue = Replace(sValue, "{", "")
    	    sValue = Replace(sValue, "}", "")
            dicC2RPropV2.Add STR_REGPACKAGEGUID, sValue
            dicC2RPropV2.Add STR_PACKAGEGUID, sPackageGuid
        End If

    End If

    'ActiveConfiguration & ConfigProducts
    'Config IDs
    'Try (but not rely on) the ProductRleaseIds entry in Configuration
    If RegReadValue(HKLM, REG_C2RCONFIGURATION, "ProductReleaseIds", sValue, REG_SZ) Then
        For Each prod in Split(sValue, ",")
            If NOT dicVirt2ConfigID.Exists(prod) Then
                dicVirt2ConfigID.Add prod, prod
            End If
        Next 'prod
    End If

    If RegEnumKey(HKLM, REG_C2RPRODUCTIDS & "Active", arrConfigProducts) Then
    	For Each prod In arrConfigProducts
    		sProd = prod
    		Select Case LCase(sProd)
    		Case "culture", "stream"
    		Case Else
	            'add to ConfigID collection
	            If NOT dicVirt2ConfigID.Exists(sProd) Then
	                dicVirt2ConfigID.Add sProd, prod
	            End If
    		End Select
    	Next 'prod
    End If 'arrConfigProducts

    'Shared ProductVersion
    If RegReadStringValue(HKLM, REG_C2RPRODUCTIDS & "Active\culture\", "x-none", sVersionFallback) Then
        dicC2RPropV2.Add STR_VERSION, sVersionFallback
    End If
    	
    'Cultures
    If RegEnumValues (HKLM, REG_C2RPRODUCTIDS & "Active\culture", arrCultures, arrTypes) Then
    	For Each culture in arrCultures
    		sCult = culture
    		Select Case LCase(sCult)
    		Case "x-none"
    		Case Else
	            'add to ConfigID collection
	            If NOT dicVirt2Cultures.Exists(sCult) Then
	                dicVirt2Cultures.Add sCult, culture
	            End If
    		End Select
    	Next 'culture
    End If 'cultures

    ' enum ARP to identify configuration products
    
    For Each ArpItem in dicArp.Keys
        ' filter on C2Rv2 products
        sCurKey = REG_ARP & ArpItem & "\"
        fUninstallString = RegReadValue(HKLM, sCurKey, "UninstallString", sValue, "REG_SZ")
        If InStr(LCase(sValue), "microsoft office 15") > 0 Then
        	For Each key In dicVirt2ConfigID.Keys
        		If InStr(sValue, key) > 0 Then
		            If NOT dicVirt2Prod.Exists(ArpItem) Then
		            	dicVirt2Prod.Add ArpItem, sCurKey
		            End If
        		End If
        	Next
            prod = Mid(sValue, InStr(sValue, "productsdata ") + 13)
            prod = Trim(prod)
            prod = Trim(Mid(prod, InStrRev(prod, " ")))
            prod = Replace(prod, "productstoremove=", "")
            If InStr(prod, "_") > 0 Then
                prod = Left(prod, InStr(prod, "_") - 1)
            End If
	        If NOT dicVirt2Prod.Exists(ArpItem) Then 
	            dicVirt2Prod.Add ArpItem, sCurKey
	        End If

        End If
    Next 'ArpItem

    'Fill the v2 virtual products array 
    If dicVirt2Prod.Count > 0 Then
        ReDim arrVirt2Products(dicVirt2Prod.Count - 1, UBOUND_VIRTPROD)
        iVCnt = 0
        ReDim arrVirt2Products(dicVirt2Prod.Count - 1, UBOUND_VIRTPROD)
        iVCnt = 0
        fIsSP1 = (CompareVersion(sVersionFallback, "15.0.4569.1506", True) > -1)
        fC2RPolEnabled = (CompareVersion(sVersionFallback, "15.0.4605.1003", True) > -1)
        
        'Global settings - applicable for all v2 products
        '------------------------------------------------
        'Scenario key state(s)
        If RegEnumKey(HKLM, REG_C2RSCENARIO, arrKeys) Then
            For Each key in arrKeys
                If RegEnumKey (HKLM, REG_C2RSCENARIO & key, arrSubKeys) Then
                    For Each subKey in arrSubKeys
                        If RegEnumValues(HKLM, REG_C2RSCENARIO & key & "\" & subKey, arrNames, arrTypes) Then
                            For Each name in arrNames
                                RegReadValue HKLM, REG_C2RSCENARIO & key & "\" & subKey, name, sValue, "REG_SZ"
                                If InStr (name, ":") > 0 Then name = Left (name, InStr(name , ":") - 1)
                                If NOT dicScenarioV2.Exists(key & "\" & name) Then dicScenarioV2.Add key & "\" & name, sValue
                            Next 'name
                        End If
                    Next 'subKey
                End If
            Next 'name
        Else
            If RegEnumValues(HKLM, REG_C2RSCENARIO, arrNames, arrTypes) Then
                For Each name in arrNames
                    RegReadValue HKLM, REG_C2RSCENARIO, name, sValue, "REG_DWORD"
                    If NOT dicScenarioV2.Exists(name) Then dicScenarioV2.Add name, sValue
                Next 'name
            End If
        End If

        'Architecture (bitness)
        If RegReadValue (HKLM, REG_C2RCONFIGURATION, "Platform", sValue, "REG_SZ") Then
            dicC2RPropV2.Add STR_PLATFORM, sValue
        ElseIf RegReadValue (HKLM, REG_C2RPROPERTYBAG, "platform", sValue, "REG_SZ") Then 
            dicC2RPropV2.Add STR_PLATFORM, sValue
        Else
            dicC2RPropV2.Add STR_PLATFORM, "Error"
        End If
            
        'SCA
        If RegReadValue (HKLM, REG_C2RCONFIGURATION, "SharedComputerLicensing", sValue, "REG_SZ") Then
            dicC2RPropV2.Add STR_SCA, sValue
        Else
            dicC2RPropV2.Add STR_SCA, STR_NOTCONFIGURED
        End If

        'CDNBaseUrl
        If RegReadValue(HKLM, REG_C2RCONFIGURATION, "CDNBaseUrl", sValue, "REG_SZ") Then
            dicC2RPropV2.Add STR_CDNBASEURL, sValue
            dicC2RPropV2.Add hAtS(Array("49","6E","73","74","61","6C","6C","53","6F","75","72","63","65","20","42","72","61","6E","63","68")), GetBr(sValue)
            'Last known InstallSource used
            If RegReadValue (HKLM, REG_C2RSCENARIO & "Install\", "BaseUrl", sValue, "REG_SZ") Then
                dicC2RPropV2.Add STR_LASTUSEDBASEURL, sValue
            End If 
        ElseIf RegReadValue (HKLM, REG_C2RPROPERTYBAG, "wwwbaseurl", sValue, "REG_SZ") Then
            dicC2RPropV2.Add STR_CDNBASEURL, sValue
            dicC2RPropV2.Add hAtS(Array("49","6E","73","74","61","6C","6C","53","6F","75","72","63","65","20","42","72","61","6E","63","68")), GetBr(sValue)
        Else
            dicC2RPropV2.Add STR_CDNBASEURL, "Error"
        End If

        'UpdatesEnabled
        If RegReadValue (HKLM, REG_C2RCONFIGURATION, "UpdatesEnabled", sValue, "REG_SZ") Then
            dicC2RPropV2.Add STR_UPDATESENABLED, sValue
        Else
            dicC2RPropV2.Add STR_UPDATESENABLED, "True"
        End If

        'UpdatesEnabled Policy
        If RegReadValue (HKLM, REG_C2RUPDATEPOL, "enableautomaticupdates", sValue, "REG_SZ") Then
            dicC2RPropV2.Add STR_POLUPDATESENABLED, sValue
        Else
            dicC2RPropV2.Add STR_POLUPDATESENABLED, STR_NOTCONFIGURED
        End If
            
        'UpGradeEnabled Policy
        If RegReadValue (HKLM, REG_C2RUPDATEPOL, "enableautomaticupgrade", sValue, "REG_SZ") Then
            dicC2RPropV2.Add STR_POLUPGRADEENABLED, sValue
        Else
            dicC2RPropV2.Add STR_POLUPGRADEENABLED, STR_NOTCONFIGURED
        End If
            
        'UpdateUrl / Path
        If RegReadValue (HKLM, REG_C2RCONFIGURATION, "UpdateUrl", sValue, "REG_SZ") Then
            dicC2RPropV2.Add STR_UPDATELOCATION, sValue
        ElseIf RegReadValue (HKLM, REG_C2RPROPERTYBAG, "UpdateUrl", sValue, "REG_SZ") Then
                dicC2RPropV2.Add STR_UPDATELOCATION, sValue
        Else
            dicC2RPropV2.Add STR_UPDATELOCATION, STR_NOTCONFIGURED
        End If
        
        'UpdateUrl / Path Policy
        If RegReadValue (HKLM, REG_C2RUPDATEPOL, "updatepath", sValue, "REG_SZ") AND fC2RPolEnabled Then
            dicC2RPropV2.Add STR_POLUPDATELOCATION, sValue
        Else
            dicC2RPropV2.Add STR_POLUPDATELOCATION, STR_NOTCONFIGURED
        End If

        'Winning Update Location
        'Default to CDNBaseUrl
        sValue = dicC2RPropV2.Item (STR_CDNBASEURL)
        If NOT dicC2RPropV2.Item (STR_UPDATELOCATION) = STR_NOTCONFIGURED Then sValue = dicC2RPropV2.Item (STR_UPDATELOCATION)
        If NOT dicC2RPropV2.Item (STR_POLUPDATELOCATION) = STR_NOTCONFIGURED Then sValue = dicC2RPropV2.Item (STR_POLUPDATELOCATION)
        If UCASE (dicC2RPropV2.Item (STR_UPDATESENABLED)) = "FALSE" Then sValue = "-"
        If UCASE (dicC2RPropV2.Item (STR_POLUPDATESENABLED)) = "FALSE" Then sValue = "-"
        If NOT GetBr (sValue) = "Custom" Then sValue = sValue & " (" & GetBr (sValue) & ")"
        dicC2RPropV2.Add STR_USEDUPDATELOCATION, sValue

         'UpdateVersion
        If RegReadValue (HKLM, REG_C2R & "Updates\", "UpdateToVersion", sValue, "REG_SZ") Then
            If sValue = "" Then sValue = STR_NOTCONFIGURED
            dicC2RPropV2.Add STR_UPDATETOVERSION, sValue
        ElseIf RegReadValue (HKLM, REG_C2RPROPERTYBAG, "UpdateToVersion", sValue, "REG_SZ") Then
                If sValue = "" Then sValue = STR_NOTCONFIGURED
                dicC2RPropV2.Add STR_UPDATETOVERSION, sValue
        Else
            dicC2RPropV2.Add STR_UPDATETOVERSION, STR_NOTCONFIGURED
        End If
        
        'UpdateVersion Policy
        If RegReadValue (HKLM, REG_C2RUPDATEPOL, "updatetargetversion", sValue, "REG_SZ") AND fC2RPolEnabled Then
            dicC2RPropV2.Add STR_POLUPDATETOVERSION, sValue
        Else
            dicC2RPropV2.Add STR_POLUPDATETOVERSION, STR_NOTCONFIGURED
        End If

        'UpdateDeadline Policy
        If RegReadValue (HKLM, REG_C2RUPDATEPOL, "updatedeadline", sValue, "REG_SZ") Then
            dicC2RPropV2.Add STR_POLUPDATEDEADLINE, sValue
        Else
            dicC2RPropV2.Add STR_POLUPDATEDEADLINE, STR_NOTCONFIGURED
        End If
            
        'UpdateHideEnableDisableUpdates Policy
        If RegReadValue (HKLM, REG_C2RUPDATEPOL, "hideenabledisableupdates ", sValue, "REG_SZ") Then
            dicC2RPropV2.Add STR_POLHIDEUPDATECFGOPT, sValue
        Else
            dicC2RPropV2.Add STR_POLHIDEUPDATECFGOPT, STR_NOTCONFIGURED
        End If
            
        'UpdateNotifications Policy
        If RegReadValue (HKLM, REG_C2RUPDATEPOL, "hideupdatenotifications ", sValue, "REG_SZ") Then
            dicC2RPropV2.Add STR_POLUPDATENOTIFICATIONS, sValue
        Else
            dicC2RPropV2.Add STR_POLUPDATENOTIFICATIONS, STR_NOTCONFIGURED
        End If
            
        'UpdateThrottle
        If RegReadValue (HKLM, REG_C2RUPDATES, "UpdatesThrottleValue", sValue, "REG_SZ") Then
            dicC2RPropV2.Add STR_UPDATETHROTTLE, sValue
        Else
            dicC2RPropV2.Add STR_UPDATETHROTTLE, ""
        End If
            
        'KeyComponentStates
        sKeyComponents = GetKeyComponentStates(sPackageGuid, False)
        arrKeyComponents = Split(sKeyComponents, ";")
        For Each component in arrKeyComponents
            arrComponentData = Split(component, ",")
            If CheckArray(arrComponentData) Then
                dicKeyComponentsV2.Add arrComponentData(0), component
            End If
        Next 'component

        ' Loop for product specific details
        For Each VProd in dicVirt2Prod.Keys
            
            'KeyName
            arrVirt2Products(iVCnt, VIRTPROD_KEYNAME) = VProd

            'ProductName
            If RegReadValue(HKLM, REG_ARP & VProd, "DisplayName", sValue, "REG_SZ") Then arrVirt2Products(iVCnt, COL_PRODUCTNAME) = sValue
            If NOT Len(sValue) > 0 Then arrVirt2Products(iVCnt, COL_PRODUCTNAME) = VProd

            'ConfigName
            fUninstallString = RegReadValue(HKLM, dicVirt2Prod.item(VProd), "UninstallString", sValue, "REG_SZ")
	        If InStr(LCase(sValue), "productstoremove=") > 0 Then
	        	prod = ""
	        	For Each key In dicVirt2ConfigID.Keys
	        		If InStr(sValue, key) > 0 Then
			            arrVirt2Products (iVCnt, VIRTPROD_CONFIGNAME) = key
			            prod = key
	        		End If
	        	Next
	        	If prod = "" Then
                    prod = Mid(sValue, InStr(sValue, "productsdata ") + 13)
                    prod = Trim(prod)
                    prod = Trim(Mid(prod, InStrRev(prod, " ")))
                    prod = Replace(prod, "productstoremove=", "")
                    If InStr(prod, "_") > 0 Then
                        prod = Left(prod, InStr(prod, "_") - 1)
                    End If
                    arrVirt2Products (iVCnt, VIRTPROD_CONFIGNAME) = prod
                End If 'prod = ""
            End If 'productstoremove

            'DisplayVersion
            If RegReadValue (HKLM, REG_ARP & VProd, "DisplayVersion", sValue, "REG_SZ") Then 
                arrVirt2Products(iVCnt, VIRTPROD_PRODUCTVERSION) = sValue
            Else
                arrVirt2Products(iVCnt, VIRTPROD_PRODUCTVERSION) = sVersionFallback
            End If

            'SP level
            If Len(arrVirt2Products(iVCnt, VIRTPROD_PRODUCTVERSION)) > 0 Then
                arrVersion = Split(arrVirt2Products(iVCnt, VIRTPROD_PRODUCTVERSION),".")
                If NOT fInitArrProdVer Then InitProdVerArrays
                arrVirt2Products (iVCnt, VIRTPROD_SPLEVEL) = OVersionToSpLevel ("{90150000-000F-0000-0000-0000000FF1CE}", arrVersion(0), arrVirt2Products(iVCnt, VIRTPROD_PRODUCTVERSION)) 
            End If

            'Child Packages
            sCurKeyL0 = REG_C2RPRODUCTIDS & "Active\"
            If RegEnumKey (HKLM, sCurKeyL0, arrConfigProducts) Then
                For Each ConfigProd in arrConfigProducts
                    sCurKeyL1 = sCurKeyL0 & ConfigProd
                    Select Case ConfigProd
                    Case "culture", "stream"
                        ' ignore
                    Case arrVirt2Products (iVCnt, VIRTPROD_CONFIGNAME)
                        If RegEnumKey(HKLM, sCurKeyL1, arrCultures) Then
                            For Each culture in arrCultures
                                sCurKeyL2 = sCurKeyL1 & "\" & culture
                                RegReadValue HKLM, sCurKeyL2, "Version", sValue, "REG_SZ"
                                arrVirt2Products(iVCnt,VIRTPROD_CHILDPACKAGES) = arrVirt2Products(iVCnt, VIRTPROD_CHILDPACKAGES) & culture & " - " & sValue & ";"
                                If RegEnumKey(HKLM, sCurKeyL2, arrChildPackages) Then
                                    For Each child in arrChildPackages
                                        sCurKeyL3 = sCurKeyL2 & "\" & child
                                        sChild = "" : sPackageGuid = "" : sVersion = "" : sFileName = ""
                                        RegReadValue HKLM, sCurKeyL3, "PackageGuid", sPackageGuid, "REG_SZ"
                                        If Len(sPackageGuid) > 0 Then
                                            sChild = "{" & sPackageGuid & "}" & " - "
                                            ' call component detection to allow correct mapping for C2R
                                            sKeyComponents = GetKeyComponentStates("{" & sPackageGuid & "}", True)
                                            If NOT sKeyComponents = "" Then arrVirt2Products(iVCnt, VIRTPROD_KEYCOMPONENTS) = arrVirt2Products(iVCnt, VIRTPROD_KEYCOMPONENTS) & sKeyComponents & ";"
                                        End If
                                        RegReadValue HKLM, sCurKeyL3, "Version", sVersion, "REG_SZ"
                                        If Len(sVersion) > 0 Then sChild = sChild & sVersion & " - "
                                        RegReadValue HKLM, sCurKeyL3, "FileName", sFileName, "REG_SZ"
                                        If Len(sFileName) > 0 Then
                                            sFileName = Replace(sFileName, ".zip", "")
                                            sChild = sChild & sFileName 
                                        End If
                                        If NOT sChild = "" Then arrVirt2Products(iVCnt, VIRTPROD_CHILDPACKAGES) = arrVirt2Products(iVCnt, VIRTPROD_CHILDPACKAGES) & sChild & ";"
                                    Next 'child
                                End If
                            Next 'culture
                        End If
                    Case Else
                        ' not the targeted product
                    End Select
                Next 'ConfigProd
            End If

            iVCnt = iVCnt + 1
        Next 'VProd
    End If 'dicVirtProd > 0


End Sub 'FindV2VirtualizedProducts

'-------------------------------------------------------------------------------
'   FindV3VirtualizedProducts
'
'   Locate virtualized C2R_v3 products
'-------------------------------------------------------------------------------
Sub FindV3VirtualizedProducts
    Dim ArpItem, VProd, culture, name, key, subKey, prod, component
    Dim sValue, sCurKey, sKey, sVersionFallback, sKeyComponents
    Dim sActiveConfiguration, sProd, sCult
    Dim iVCnt
    Dim dicVirt3Prod, dicVirt3ConfigID
    Dim arrKeys, arrConfigProducts, arrVersion, arrCultures
    Dim arrNames, arrTypes, arrSubKeys, arrKeyComponents, arrComponentData
    Dim fUninstallString
    
   	Const REG_C2RCONFIGURATION	    = "SOFTWARE\Microsoft\Office\ClickToRun\Configuration\"
   	Const REG_C2RSCENARIO           = "SOFTWARE\Microsoft\Office\ClickToRun\Scenario\"
   	Const REG_C2RUPDATES            = "SOFTWARE\Microsoft\Office\ClickToRun\Updates\"
	Const REG_C2RPRODUCTIDS		    = "SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs\"
	Const REG_C2R				    = "SOFTWARE\Microsoft\Office\ClickToRun\"
	Const REG_C2RUPDATEPOL		    = "SOFTWARE\Policies\Microsoft\Office\16.0\Common\OfficeUpdate\"

    On Error Resume Next

    Set dicVirt3Prod = CreateObject ("Scripting.Dictionary")
    Set dicVirt3ConfigID = CreateObject("Scripting.Dictionary")
    ' extend ARP dic to contain virt3 references
    sKey = REG_C2R & "REGISTRY\MACHINE\" & REG_ARP
    If RegEnumKey (HKLM, sKey, arrKeys) Then
        For Each key in arrKeys
            If NOT dicArp.Exists(key) Then dicArp.Add key, sKey & key
        Next 'key
    End If
    
    'Integration PackageGUID
    If RegReadValue(HKLM, REG_C2R, "PackageGUID", sValue, REG_SZ) Then
    	sPackageGuid = "{" & sValue & "}"
        dicC2RPropV3.Add STR_REGPACKAGEGUID, sValue
        dicC2RPropV3.Add STR_PACKAGEGUID, sPackageGuid
    End If
    
    'ActiveConfiguration & ConfigProducts
    If RegReadValue(HKLM, REG_C2RPRODUCTIDS, "ActiveConfiguration", sActiveConfiguration, REG_SZ) Then
        'Config IDs
        'Try (but not rely on) the ProductRleaseIds entry in Configuration
        If RegReadValue(HKLM, REG_C2RCONFIGURATION, "ProductReleaseIds", sValue, REG_SZ) Then
            For Each prod in Split(sValue, ",")
                If NOT dicVirt3ConfigID.Exists(prod) Then
                    dicVirt3ConfigID.Add prod, prod
                End If
            Next 'prod
        End If

        If RegEnumKey(HKLM, REG_C2RPRODUCTIDS & sActiveConfiguration, arrConfigProducts) Then
    	    For Each prod In arrConfigProducts
    		    sProd = prod
			    If InStr(sProd, ".16") > 0 Then sProd = Left(sProd, InStr(sProd, ".16") - 1)
    		    Select Case LCase(sProd)
    		    Case "culture", "stream"
    		    Case Else
	                'add to ConfigID collection
	                If NOT dicVirt3ConfigID.Exists(sProd) Then
	                    dicVirt3ConfigID.Add sProd, prod
	                End If
    		    End Select
    	    Next 'prod
        End If 'arrConfigProducts

        'Shared ProductVersion
        If RegReadStringValue(HKLM, REG_C2RPRODUCTIDS & sActiveConfiguration & "\culture\x-none.16\", "Version", sVersionFallback) Then
            dicC2RPropV3.Add STR_VERSION, sVersionFallback
        End If
    	
    	'Cultures
        If RegEnumKey(HKLM, REG_C2RPRODUCTIDS & sActiveConfiguration & "\culture", arrCultures) Then
    		For Each culture in arrCultures
    			sCult = culture
				If InStr(sCult, ".16") > 0 Then sCult = Left(sCult, InStr(sCult, ".16") - 1)
    			Select Case LCase(sCult)
    			Case "x-none"
    			Case Else
	                'add to ConfigID collection
	                If NOT dicVirt3Cultures.Exists(sCult) Then
	                	dicVirt3Cultures.Add sCult, culture
	                End If
    			End Select
    		Next 'culture
        End If 'cultures
    End If 'ActiveConfiguration
        
    ' enum ARP to identify configuration products
    For Each ArpItem in dicArp.Keys
        ' filter on C2Rv3 products
        sCurKey = REG_ARP & ArpItem & "\"
        fUninstallString = RegReadValue(HKLM, sCurKey, "UninstallString", sValue, "REG_SZ")
        If InStr(LCase(sValue), "productstoremove=") > 0 Then
        	For Each key In dicVirt3ConfigID.Keys
        		If InStr(sValue, key) > 0 Then
		            If NOT dicVirt3Prod.Exists(ArpItem) Then
		            	dicVirt3Prod.Add ArpItem, sCurKey
		            End If
        		End If
        	Next
        	prod = Mid(sValue, InStr(sValue, "productstoremove="))
	        prod = Replace(prod, "productstoremove=", "")
	        If InStr(prod, "_") > 0 Then
	            prod = Left(prod, InStr(prod, "_") - 1)
	        End If
	        If InStr(prod, ".16") > 0 Then
	            prod = Left(prod, InStr(prod, ".16") - 1)
	            If NOT dicVirt3Prod.Exists(ArpItem) Then 
	            	dicVirt3Prod.Add ArpItem, sCurKey
	            End If
	        End If
        End If
    Next 'ArpItem

    'Fill the v3 virtual products array 
    If dicVirt3Prod.Count > 0 Then
        ReDim arrVirt3Products(dicVirt3Prod.Count - 1, UBOUND_VIRTPROD)
        iVCnt = 0
        
        'Global settings - applicable for all v3 products
        '------------------------------------------------
        'Scenario key state(s)
        If RegEnumKey(HKLM, REG_C2RSCENARIO, arrKeys) Then
            For Each key in arrKeys
                If RegEnumKey (HKLM, REG_C2RSCENARIO & key, arrSubKeys) Then
                    For Each subKey in arrSubKeys
                        If RegEnumValues(HKLM, REG_C2RSCENARIO & key & "\" & subKey, arrNames, arrTypes) Then
                            For Each name in arrNames
                                RegReadValue HKLM, REG_C2RSCENARIO & key & "\" & subKey, name, sValue, "REG_SZ"
                                If InStr (name, ":") > 0 Then name = Left (name, InStr(name , ":") - 1)
                                If NOT dicScenarioV3.Exists(key & "\" & name) Then dicScenarioV3.Add key & "\" & name, sValue
                            Next 'name
                        End If
                    Next 'subKey
                End If
            Next 'name
        End If
        
        'Architecture (bitness)
        If RegReadValue (HKLM, REG_C2RCONFIGURATION, "Platform", sValue, "REG_SZ") Then
            dicC2RPropV3.Add STR_PLATFORM, sValue
        Else 
            dicC2RPropV3.Add STR_PLATFORM, "Error"
        End If

        'SCA
        If RegReadValue (HKLM, REG_C2RCONFIGURATION, "SharedComputerLicensing", sValue, "REG_SZ") Then
            dicC2RPropV3.Add STR_SCA, sValue
        Else
            dicC2RPropV3.Add STR_SCA, STR_NOTCONFIGURED
        End If

        'CDNBaseUrl
        If RegReadValue(HKLM, REG_C2RCONFIGURATION, "CDNBaseUrl", sValue, "REG_SZ") Then
            dicC2RPropV3.Add STR_CDNBASEURL, sValue
            dicC2RPropV3.Add hAtS(Array("49","6E","73","74","61","6C","6C","53","6F","75","72","63","65","20","42","72","61","6E","63","68")), GetBr(sValue)
            'Last known InstallSource used
            If RegReadValue(HKLM, REG_C2RSCENARIO & "Install", "BaseUrl", sValue, "REG_SZ") Then
                dicC2RPropV3.Add STR_LASTUSEDBASEURL, sValue
            End If 
        Else
            dicC2RPropV3.Add STR_CDNBASEURL, "Error"
        End If
            
        'UpdatesEnabled
        If RegReadValue (HKLM, REG_C2RCONFIGURATION, "UpdatesEnabled", sValue, "REG_SZ") Then
            dicC2RPropV3.Add STR_UPDATESENABLED, sValue
        Else
            dicC2RPropV3.Add STR_UPDATESENABLED, "True"
        End If
            
        'UpdatesEnabled Policy
        If RegReadValue (HKLM, REG_C2RUPDATEPOL, "enableautomaticupdates", sValue, "REG_SZ") Then
            dicC2RPropV3.Add STR_POLUPDATESENABLED, sValue
        Else
            dicC2RPropV3.Add STR_POLUPDATESENABLED, STR_NOTCONFIGURED
        End If
            
        'UpdateBranch 
        If RegReadValue (HKLM, REG_C2RCONFIGURATION, "UpdateBranch", sValue, "REG_SZ") Then
            dicC2RPropV3.Add STR_UPDATEBRANCH, sValue
        Else
            dicC2RPropV3.Add STR_UPDATEBRANCH, STR_NOTCONFIGURED
        End If
        
        'UpdateBranch Policy
        If RegReadValue (HKLM, REG_C2RUPDATEPOL, "updatebranch", sValue, "REG_SZ") Then
            dicC2RPropV3.Add STR_POLUPDATEBRANCH, sValue
        Else
            dicC2RPropV3.Add STR_POLUPDATEBRANCH, STR_NOTCONFIGURED
        End If

        'UpdateUrl / Path
        If RegReadValue (HKLM, REG_C2RCONFIGURATION, "UpdateUrl", sValue, "REG_SZ") Then
            dicC2RPropV3.Add STR_UPDATELOCATION, sValue
        Else
            dicC2RPropV3.Add STR_UPDATELOCATION, STR_NOTCONFIGURED
        End If
        
        'UpdateUrl / Path Policy
        If RegReadValue (HKLM, REG_C2RUPDATEPOL, "updatepath", sValue, "REG_SZ") Then
            dicC2RPropV3.Add STR_POLUPDATELOCATION, sValue
        Else
            dicC2RPropV3.Add STR_POLUPDATELOCATION, STR_NOTCONFIGURED
        End If

        'Winning Update Location
        'Default to CDNBaseUrl
        sValue = dicC2RPropV3.Item (STR_CDNBASEURL)
        If NOT dicC2RPropV3.Item (STR_UPDATEBRANCH) = STR_NOTCONFIGURED Then sValue = dicC2RPropV3.Item (STR_UPDATEBRANCH)
        If NOT dicC2RPropV3.Item (STR_UPDATELOCATION) = STR_NOTCONFIGURED Then sValue = dicC2RPropV3.Item (STR_UPDATELOCATION)
        If NOT dicC2RPropV3.Item (STR_POLUPDATEBRANCH) = STR_NOTCONFIGURED Then sValue = dicC2RPropV3.Item (STR_POLUPDATEBRANCH)
        If NOT dicC2RPropV3.Item (STR_POLUPDATELOCATION) = STR_NOTCONFIGURED Then sValue = dicC2RPropV3.Item (STR_POLUPDATELOCATION)
        If UCASE (dicC2RPropV3.Item (STR_UPDATESENABLED)) = "FALSE" Then sValue = "-"
        If UCASE (dicC2RPropV3.Item (STR_POLUPDATESENABLED)) = "FALSE" Then sValue = "-"
        If NOT GetBr (sValue) = "Custom" Then sValue = sValue & " (" & GetBr (sValue) & ")"
        dicC2RPropV3.Add STR_USEDUPDATELOCATION, sValue

        'UpdateVersion
        If RegReadValue (HKLM, REG_C2RUPDATES, "UpdateToVersion", sValue, "REG_SZ") Then
            If sValue = "" Then sValue = STR_NOTCONFIGURED
            dicC2RPropV3.Add STR_UPDATETOVERSION, sValue
        Else
            dicC2RPropV3.Add STR_UPDATETOVERSION, STR_NOTCONFIGURED
        End If
        
        'UpdateVersion Policy
        If RegReadValue (HKLM, REG_C2RUPDATEPOL, "updatetargetversion", sValue, "REG_SZ") Then
            dicC2RPropV3.Add STR_POLUPDATETOVERSION, sValue
        Else
            dicC2RPropV3.Add STR_POLUPDATETOVERSION, STR_NOTCONFIGURED
        End If
            
        'UpdateDeadline Policy
        If RegReadValue (HKLM, REG_C2RUPDATEPOL, "updatedeadline", sValue, "REG_SZ") Then
            dicC2RPropV3.Add STR_POLUPDATEDEADLINE, sValue
        Else
            dicC2RPropV3.Add STR_POLUPDATEDEADLINE, STR_NOTCONFIGURED
        End If
            
        'UpdateHideEnableDisableUpdates Policy
        If RegReadValue (HKLM, REG_C2RUPDATEPOL, "hideenabledisableupdates ", sValue, "REG_SZ") Then
            dicC2RPropV3.Add STR_POLHIDEUPDATECFGOPT, sValue
        Else
            dicC2RPropV3.Add STR_POLHIDEUPDATECFGOPT, STR_NOTCONFIGURED
        End If
            
        'UpdateNotifications Policy
        If RegReadValue (HKLM, REG_C2RUPDATEPOL, "hideupdatenotifications ", sValue, "REG_SZ") Then
            dicC2RPropV3.Add STR_POLUPDATENOTIFICATIONS, sValue
        Else
            dicC2RPropV3.Add STR_POLUPDATENOTIFICATIONS, STR_NOTCONFIGURED
        End If
            
        'UpdateThrottle
        If RegReadValue (HKLM, REG_C2RUPDATES, "UpdatesThrottleValue", sValue, "REG_SZ") Then
            dicC2RPropV3.Add STR_UPDATETHROTTLE, sValue
        Else
            dicC2RPropV3.Add STR_UPDATETHROTTLE, ""
        End If

        'KeyComponentStates
        sKeyComponents = GetKeyComponentStates(sPackageGuid, False)
        arrKeyComponents = Split(sKeyComponents, ";")
        For Each component in arrKeyComponents
            arrComponentData = Split(component, ",")
            If CheckArray(arrComponentData) Then
                dicKeyComponentsV3.Add arrComponentData(0), component

            End If
        Next 'component
            
        ' Loop for product specific details
        For Each VProd in dicVirt3Prod.Keys

            'KeyName
            arrVirt3Products(iVCnt, VIRTPROD_KEYNAME) = VProd

            'ProductName
            If RegReadValue(HKLM, REG_ARP & VProd, "DisplayName", sValue, "REG_SZ") Then arrVirt3Products(iVCnt, COL_PRODUCTNAME) = sValue
            If NOT Len(sValue) > 0 Then arrVirt3Products(iVCnt, COL_PRODUCTNAME) = VProd

            'ConfigName
            fUninstallString = RegReadValue(HKLM, dicVirt3Prod.item(VProd), "UninstallString", sValue, "REG_SZ")
	        If InStr(LCase(sValue), "productstoremove=") > 0 Then
	        	prod = ""
	        	For Each key In dicVirt3ConfigID.Keys
	        		If InStr(sValue, key) > 0 Then
			            arrVirt3Products (iVCnt, VIRTPROD_CONFIGNAME) = key
			            prod = key
	        		End If
	        	Next
	        	If prod = "" Then
		        	prod = Mid(sValue, InStr(sValue, "productstoremove="))
			        prod = Replace(prod, "productstoremove=", "")
			        If InStr(prod, "_") > 0 Then
			            prod = Left(prod, InStr(prod, "_") - 1)
			        End If
			        If InStr(prod, ".16") > 0 Then
			            prod = Left(prod, InStr(prod, ".16") - 1)
			            arrVirt3Products (iVCnt, VIRTPROD_CONFIGNAME) = prod
			        End If
	        	End If
	        End If
	        
            'DisplayVersion
            If RegReadValue(HKLM, REG_ARP & VProd, "DisplayVersion", sValue, "REG_SZ") Then 
                arrVirt3Products(iVCnt, VIRTPROD_PRODUCTVERSION) = sValue
            Else
                RegReadStringValue HKLM, REG_C2RPRODUCTIDS & sActiveConfiguration & "\culture\x-none.16\", "Version", sVersionFallback
                arrVirt3Products(iVCnt, VIRTPROD_PRODUCTVERSION) = sVersionFallback
            End If

            'SP level
            If Len(arrVirt3Products(iVCnt, VIRTPROD_PRODUCTVERSION)) > 0 Then
                arrVersion = Split(arrVirt3Products(iVCnt, VIRTPROD_PRODUCTVERSION),".")
                If NOT fInitArrProdVer Then InitProdVerArrays
                arrVirt3Products (iVCnt, VIRTPROD_SPLEVEL) = OVersionToSpLevel ("{90160000-000F-0000-0000-0000000FF1CE}", arrVersion(0), arrVirt3Products(iVCnt, VIRTPROD_PRODUCTVERSION)) 
            End If

            iVCnt = iVCnt + 1
        Next 'VProd
    End If 'dicVirtProd > 0


End Sub 'FindV3VirtualizedProducts

'-------------------------------------------------------------------------------
'   MapAppsToConfigProduct
'
'   Obtain the mapping of which apps are contained in a SKU based on the
'   products ConfigID
'   Pass in a dictionary object which and return the filled dic 
'-------------------------------------------------------------------------------
Sub MapAppsToConfigProduct(dicConfigIdApps, sConfigId, iVM)
    Dim sMondo

    dicConfigIdApps.RemoveAll
    'sMondo = "MondoVolume"
    sMondo = hAtS(Array("4D","6F","6E","64","6F","56","6F","6C","75","6D","65"))
    Select Case sConfigId
    Case sMondo
        dicConfigIdApps.Add "Access", ""
        dicConfigIdApps.Add "Excel", ""
        dicConfigIdApps.Add "Groove", ""
        If iVM = 15 Then dicConfigIdApps.Add "InfoPath", ""
        dicConfigIdApps.Add "Skype for Business", ""
        dicConfigIdApps.Add "OneNote", ""
        dicConfigIdApps.Add "Outlook", ""
        dicConfigIdApps.Add "PowerPoint", ""
        dicConfigIdApps.Add "Publisher", ""
        dicConfigIdApps.Add "Word", ""
        dicConfigIdApps.Add "Project", ""
        dicConfigIdApps.Add "Visio", ""
    Case "O365ProPlusRetail"
        dicConfigIdApps.Add "Access", ""
        dicConfigIdApps.Add "Excel", ""
        dicConfigIdApps.Add "Groove", ""
        If iVM = 15 Then dicConfigIdApps.Add "InfoPath", ""
        dicConfigIdApps.Add "Skype for Business", ""
        dicConfigIdApps.Add "OneNote", ""
        dicConfigIdApps.Add "Outlook", ""
        dicConfigIdApps.Add "PowerPoint", ""
        dicConfigIdApps.Add "Publisher", ""
        dicConfigIdApps.Add "Word", ""
    Case "O365BusinessRetail"
        dicConfigIdApps.Add "Excel", ""
        dicConfigIdApps.Add "Groove", ""
        dicConfigIdApps.Add "Skype for Business", ""
        dicConfigIdApps.Add "OneNote", ""
        dicConfigIdApps.Add "Outlook", ""
        dicConfigIdApps.Add "PowerPoint", ""
        dicConfigIdApps.Add "Publisher", ""
        dicConfigIdApps.Add "Word", ""
    Case "O365SmallBusPremRetail"
        dicConfigIdApps.Add "Excel", ""
        dicConfigIdApps.Add "Groove", ""
        dicConfigIdApps.Add "Skype for Business", ""
        dicConfigIdApps.Add "OneNote", ""
        dicConfigIdApps.Add "Outlook", ""
        dicConfigIdApps.Add "PowerPoint", ""
        dicConfigIdApps.Add "Publisher", ""
        dicConfigIdApps.Add "Word", ""
    Case "VisioProRetail"
        dicConfigIdApps.Add "Visio", ""
    Case "ProjectProRetail"
        dicConfigIdApps.Add "Project", ""
    Case "AccessRetail"
        dicConfigIdApps.Add "Access", ""
    Case "ExcelRetail"
        dicConfigIdApps.Add "Excel", ""
    Case "GrooveRetail"
        dicConfigIdApps.Add "Groove", ""
    Case "HomeBusinessRetail"
        dicConfigIdApps.Add "Excel", ""
        dicConfigIdApps.Add "OneNote", ""
        dicConfigIdApps.Add "Outlook", ""
        dicConfigIdApps.Add "PowerPoint", ""
        dicConfigIdApps.Add "Word", ""
    Case "HomeStudentRetail"
        dicConfigIdApps.Add "Excel", ""
        dicConfigIdApps.Add "OneNote", ""
        dicConfigIdApps.Add "PowerPoint", ""
        dicConfigIdApps.Add "Word", ""
    Case "InfoPathRetail"
        dicConfigIdApps.Add "InfoPath", ""
    Case "LyncEntryRetail"
        dicConfigIdApps.Add "Skype for Business", ""
    Case "LyncRetail"
        dicConfigIdApps.Add "Skype for Business", ""
    Case "SkypeforBusinessEntryRetail"
        dicConfigIdApps.Add "Skype for Business", ""
    Case "SkypeforBusinessRetail"
        dicConfigIdApps.Add "Skype for Business", ""
    Case "ProfessionalRetail"
        dicConfigIdApps.Add "Access", ""
        dicConfigIdApps.Add "Excel", ""
        dicConfigIdApps.Add "OneNote", ""
        dicConfigIdApps.Add "Outlook", ""
        dicConfigIdApps.Add "PowerPoint", ""
        dicConfigIdApps.Add "Publisher", ""
        dicConfigIdApps.Add "Word", ""
    Case "O365HomePremRetail"
        dicConfigIdApps.Add "Access", ""
        dicConfigIdApps.Add "Excel", ""
        dicConfigIdApps.Add "OneNote", ""
        dicConfigIdApps.Add "Outlook", ""
        dicConfigIdApps.Add "PowerPoint", ""
        dicConfigIdApps.Add "Publisher", ""
        dicConfigIdApps.Add "Word", ""
    Case "OneNoteRetail"
        dicConfigIdApps.Add "OneNote", ""
    Case "OutlookRetail"
        dicConfigIdApps.Add "Outlook", ""
    Case "PowerPointRetail"
        dicConfigIdApps.Add "PowerPoint", ""
    Case "ProjectStdRetail"
        dicConfigIdApps.Add "Project", ""
    Case "PublisherRetail"
        dicConfigIdApps.Add "Publisher", ""
    Case "VisioStdRetail"
        dicConfigIdApps.Add "Visio", ""
    Case "WordRetail"
        dicConfigIdApps.Add "Word", ""
    End Select

End Sub

'-------------------------------------------------------------------------------
'   GetC2Rv2VersionsActive
'
'   Obtain the version major for active C2Rv2 versions
'   Returns an array with version numbers that have a C2Rv2 product
'   An active version is detected if the registry entry 'x-none' is found at 
'   HKLM\SOFTWARE\Microsoft\Office\XX.y\ClickToRun\ProductReleaseIDs\Active\culture
'-------------------------------------------------------------------------------
Sub GetC2Rv2VersionsActive()
    Dim key, sValue, sActiveConfiguration
    Dim arrKeys

    On Error Resume Next
    
    If RegEnumKey (HKLM, "Software\Microsoft\Office", arrKeys) Then
        For Each key in arrKeys
            Select Case LCase(key)
            Case "15.0"
                If RegReadValue (HKLM, "Software\Microsoft\Office\" & key & "\ClickToRun\ProductReleaseIDs\Active\culture", "x-none", sValue, "REG_SZ") Then
                    If Not dicActiveC2Rv2Versions.Exists(sValue) Then dicActiveC2Rv2Versions.Add sValue, Left(sValue, 2)
                End If
            Case "clicktorun"
		        If RegReadValue(HKLM, "Software\Microsoft\Office\" & key & "\ProductReleaseIDs", "ActiveConfiguration", sActiveConfiguration, "REG_SZ") Then
		        	If RegReadValue(HKLM, "Software\Microsoft\Office\" & key & "\ProductReleaseIDs\" & sActiveConfiguration & "\culture\x-none.16", "Version", sValue, "REG_SZ") Then
	                    If Not dicActiveC2Rv2Versions.Exists(sValue) Then dicActiveC2Rv2Versions.Add sValue, Left(sValue, 2)
                    End If
                End If
            End Select
        Next
    End If

End Sub 'GetC2Rv2VersionsActive

'-------------------------------------------------------------------------------
'   GetBr
'
'-------------------------------------------------------------------------------
Function GetBr(sValue)
    Dim sCh0, sCh1, sCh2, sCh3, sCh4, sBr

    sBr = hAtS(Array("43","75","73","74","6F","6D"))
    sCh0 = hAtS(Array("33","39","31","36","38","44","37","45","2D","30","37","37","42","2D","34","38","45","37","2D","38","37","32","43","2D","42","32","33","32","43","33","45","37","32","36","37","35"))
    sCh1 = hAtS(Array("34","39","32","33","35","30","66","36","2D","33","61","30","31","2D","34","66","39","37","2D","62","39","63","30","2D","63","37","63","36","64","64","66","36","37","64","36","30"))
    sCh2 = hAtS(Array("36","34","32","35","36","61","66","65","2D","66","35","64","39","2D","34","66","38","36","2D","38","39","33","36","2D","38","38","34","30","61","36","61","34","66","35","62","65"))
    sCh3 = hAtS(Array("62","38","66","39","62","38","35","30","2D","33","32","38","64","2D","34","33","35","35","2D","39","31","34","35","2D","63","35","39","34","33","39","61","30","63","34","63","66"))
    sCh4 = hAtS(Array("37","66","66","62","63","36","62","66","2D","62","63","33","32","2D","34","66","39","32","2D","38","39","38","32","2D","66","39","64","64","31","37","66","64","33","31","31","34"))
    If InStr(sValue, sCh0) > 0 Then sBr = hAtS(Array("4F","31","35","20","50","55"))
    If InStr(sValue, sCh1) > 0 Then sBr = hAtS(Array("43","75","72","72","65","6E","74","20","28","43","42","29"))
    If InStr(sValue, sCh2) > 0 Then sBr = hAtS(Array("46","69","72","73","74","52","65","6C","65","61","73","65","43","75","72","72","65","6E","74","20","28","49","6E","73","69","64","65","72","29"))
    If InStr(sValue, sCh3) > 0 Then sBr = hAtS(Array("46","69","72","73","74","52","65","6C","65","61","73","65","42","75","73","69","6E","65","73","73","20","28","46","52","20","43","42","42","29"))
    If InStr(sValue, sCh4) > 0 Then sBr = hAtS(Array("42","75","73","69","6E","65","73","73","20","28","43","42","42","29"))
    GetBr = sBr
End Function

'-------------------------------------------------------------------------------
'   GetKeyComponentStates
'
'   Obtain the key component states for a product
'   Returns a string with the applicable states for the key components
'   Application: Name, ExeName, VersionMajor, Version, InstallState, 
'                InstallStateString, ComponentId, FeatureName, Path 
'-------------------------------------------------------------------------------
Function GetKeyComponentStates(sProductCode, fVirtualized)
    Dim sSkuId, sReturn, sComponents, sProd, sName, sExeName
    Dim sVersion, sInstallState, sInstallStateString, sPath
    Dim component
    Dim sFeatureName
    Dim iVM
    Dim fIsWW
    Dim arrComponents

    On Error Resume Next
    sReturn = ""
    GetKeyComponentStates = sReturn
	If IsOfficeProduct(sProductCode) Then
        iVM = GetVersionMajor(sProductCode)
        If sProductCode = sPackageGuid Then

        End If
        If iVM < 12 Then 
            sSkuId = Mid(sProductCode, 4, 2)
            fIsWW = True
        Else
            sSkuId = Mid(sProductCode, 11, 4)
            fIsWW = (Mid(sProductCode, 16, 4) = "0000")
        End If
        If sProductCode = sPackageGuid Then
            fIsWW = True
            If dicC2RPropV2.Count > 0 Then iVM = 15
            If dicC2RPropV3.Count > 0 Then iVM = 16
        End If
    End If
    If NOT fIsWW Then Exit Function 
	
    sComponents = GetCoreComponentCode(sSkuId, iVM, fVirtualized)
    arrComponents = Split(sComponents, ";")
    sProd = sProductCode
    If fVirtualized Then sProd = UCase(dicProductCodeC2R.Item(iVM))

    ' get the component data registered to the product
    For Each component in arrComponents
        If dicKeyComponents.Exists (component) Then
            If InStr (UCase(dicKeyComponents.Item (component)), UCase(sProd)) > 0 Then
                sFeatureName = GetMsiFeatureName (component)
                sName = GetApplicationName (sFeatureName)
                sExeName = GetApplicationExeName (component)
                sPath = oMsi.ComponentPath (sProd, component)
                If oFso.FileExists (sPath) Then sVersion = oFso.GetFileVersion (sPath) Else sVersion = ""
                sInstallState = oMsi.FeatureState (sProd, sFeatureName)
                sInstallStateString = TranslateFeatureState (sInstallState)
                sReturn = sReturn & sName & "," & sExeName & "," & iVM & "," & sVersion & "," & sInstallState & "," & sInstallStateString & "," & component & "," & sFeatureName & "," & sPath & ";"
            End If
        End If
    Next 'component
    If NOT sReturn = "" Then sReturn = Left(sReturn, Len(sReturn) - 1)
    GetKeyComponentStates = sReturn

End Function 'GetKeyComponentStates

'-------------------------------------------------------------------------------
'   GetCoreComponentCode
'
'   Returns the component code(s) for core application .exe files
'   from the SKU element of the productcode
'-------------------------------------------------------------------------------
Function GetCoreComponentCode (sSkuId, iVM, fVirtualized)
    Dim sReturn
    Dim arrTmp

    On Error Resume Next
    sReturn = ""
    GetCoreComponentCode = sReturn
    Select Case iVM
    Case 16
        sReturn = Join(Array(CID_ACC16_64,CID_ACC16_32,CID_XL16_64,CID_XL16_32,CID_GRV16_64,CID_GRV16_32,CID_LYN16_64,CID_LYN16_32,CID_ONE16_64,CID_ONE16_32,CID_OL16_64,CID_OL16_32,CID_PPT16_64,CID_PPT16_32,CID_PRJ16_64,CID_PRJ16_32,CID_PUB16_64,CID_PUB16_32,CID_IP16_64,CID_VIS16_64,CID_VIS16_32,CID_WD16_64,CID_WD16_32,CID_SPD16_64,CID_SPD16_32,CID_MSO16_64,CID_MSO16_32), ";")
        If fVirtualized Then
            Select Case sSkuId
            Case "0015" : sReturn = CID_ACC16_64 & ";" & CID_ACC16_32 & ";" & CID_MSO16_64 & ";" & CID_MSO16_32 'Access
            Case "0016","0029" : sReturn = CID_XL16_64 & ";" & CID_XL16_32 & ";" & CID_MSO16_64 & ";" & CID_MSO16_32 'Excel
            Case "0018" : sReturn = CID_PPT16_64 & ";" & CID_PPT16_32 & ";" & CID_MSO16_64 & ";" & CID_MSO16_32 'PowerPoint
            Case "0019" : sReturn = CID_PUB16_64 & ";" & CID_PUB16_32 & ";" & CID_MSO16_64 & ";" & CID_MSO16_32 'Publisher
            Case "001A","00E0" : sReturn = CID_OL16_64 & ";" & CID_OL16_32 & ";" & CID_MSO16_64 & ";" & CID_MSO16_32 'Outlook
            Case "001B","002B" : sReturn = CID_WD16_64 & ";" & CID_WD16_32 & ";" & CID_MSO16_64 & ";" & CID_MSO16_32 'Word
            Case "0027" : sReturn = CID_PRJ16_64 & ";" & CID_PRJ16_32 & ";" & CID_MSO16_64 & ";" & CID_MSO16_32 'Project
            Case "0044" : sReturn = CID_IP16_64 & ";"  & CID_MSO16_64 & ";" & CID_MSO16_32 'InfoPath
            Case "0017" : sReturn = CID_SPD16_64 & ";" & CID_SPD16_32 & ";" & CID_MSO16_64 & ";" & CID_MSO16_32 'SharePointDesigner
            Case "0051","0053","0057" : sReturn = CID_VIS16_64 & ";" & CID_VIS16_32 & ";" & CID_MSO16_64 & ";" & CID_MSO16_32 'Visio
            Case "00A1","00A3" : sReturn = CID_ONE16_64 & ";" & CID_ONE16_32 & ";" & CID_MSO16_64 & ";" & CID_MSO16_32 'OneNote
            Case "00BA" : sReturn = CID_GRV16_64 & ";" & CID_GRV16_32 & ";" & CID_MSO16_64 & ";" & CID_MSO16_32 'Groove
            Case "012B","012C" : sReturn = CID_LYN16_64 & ";" & CID_LYN16_32 & ";" & CID_MSO16_64 & ";" & CID_MSO16_32 'Lync
            Case Else : sReturn = ""
            End Select
        End If
    Case 15
        sReturn = Join(Array(CID_ACC15_64,CID_ACC15_32,CID_XL15_64,CID_XL15_32,CID_GRV15_64,CID_GRV15_32,CID_LYN15_64,CID_LYN15_32,CID_ONE15_64,CID_ONE15_32,CID_OL15_64,CID_OL15_32,CID_PPT15_64,CID_PPT15_32,CID_PRJ15_64,CID_PRJ15_32,CID_PUB15_64,CID_PUB15_32,CID_IP15_64,CID_IP15_32,CID_VIS15_64,CID_VIS15_32,CID_WD15_64,CID_WD15_32,CID_SPD15_64,CID_SPD15_32,CID_MSO15_64,CID_MSO15_32), ";")
        If fVirtualized Then
            Select Case sSkuId
            Case "0015" : sReturn = CID_ACC15_64 & ";" & CID_ACC15_32 & ";" & CID_MSO15_64 & ";" & CID_MSO15_32 'Access
            Case "0016","0029" : sReturn = CID_XL15_64 & ";" & CID_XL15_32 & ";" & CID_MSO15_64 & ";" & CID_MSO15_32 'Excel
            Case "0018" : sReturn = CID_PPT15_64 & ";" & CID_PPT15_32 & ";" & CID_MSO15_64 & ";" & CID_MSO15_32 'PowerPoint
            Case "0019" : sReturn = CID_PUB15_64 & ";" & CID_PUB15_32 & ";" & CID_MSO15_64 & ";" & CID_MSO15_32 'Publisher
            Case "001A","00E0" : sReturn = CID_OL15_64 & ";" & CID_OL15_32 & ";" & CID_MSO15_64 & ";" & CID_MSO15_32 'Outlook
            Case "001B","002B" : sReturn = CID_WD15_64 & ";" & CID_WD15_32 & ";" & CID_MSO15_64 & ";" & CID_MSO15_32 'Word
            Case "0027" : sReturn = CID_PRJ15_64 & ";" & CID_PRJ15_32 & ";" & CID_MSO15_64 & ";" & CID_MSO15_32 'Project
            Case "0044" : sReturn = CID_IP15_64 & ";" & CID_IP15_32 & ";" & CID_MSO15_64 & ";" & CID_MSO15_32 'InfoPath
            Case "0017" : sReturn = CID_SPD15_64 & ";" & CID_SPD15_32 & ";" & CID_MSO15_64 & ";" & CID_MSO15_32 'SharePointDesigner
            Case "0051","0053","0057" : sReturn = CID_VIS15_64 & ";" & CID_VIS15_32 & ";" & CID_MSO15_64 & ";" & CID_MSO15_32 'Visio
            Case "00A1","00A3" : sReturn = CID_ONE15_64 & ";" & CID_ONE15_32 & ";" & CID_MSO15_64 & ";" & CID_MSO15_32 'OneNote
            Case "00BA" : sReturn = CID_GRV15_64 & ";" & CID_GRV15_32 & ";" & CID_MSO15_64 & ";" & CID_MSO15_32 'Groove
            Case "012B","012C" : sReturn = CID_LYN15_64 & ";" & CID_LYN15_32 & ";" & CID_MSO15_64 & ";" & CID_MSO15_32 'Lync
            Case Else : sReturn = ""
            End Select
        End If
    Case 14
        sReturn = Join(Array(CID_ACC14_64,CID_ACC14_32,CID_XL14_64,CID_XL14_32,CID_GRV14_64,CID_GRV14_32,CID_ONE14_64,CID_ONE14_32,CID_OL14_64,CID_OL14_32,CID_PPT14_64,CID_PPT14_32,CID_PRJ14_64,CID_PRJ14_32,CID_PUB14_64,CID_PUB14_32,CID_IP14_64,CID_IP14_32,CID_VIS14_64,CID_VIS14_32,CID_WD14_64,CID_WD14_32,CID_SPD14_64,CID_SPD14_32,CID_MSO14_64,CID_MSO14_32), ";")
    Case 12
        sReturn = Join(Array(CID_ACC12,CID_XL12,CID_GRV12,CID_ONE12,CID_OL12,CID_PPT12,CID_PRJ12,CID_PUB12,CID_IP12,CID_VIS12,CID_WD12,CID_SPD12,CID_MSO12), ";")
    Case 11
        sReturn = Join(Array(CID_ACC11,CID_XL11,CID_ONE11,CID_OL11,CID_PPT11,CID_PRJ11,CID_PUB11,CID_IP11,CID_VIS11,CID_WD11,CID_SPD11,CID_MSO11), ";")
    End Select
    GetCoreComponentCode = sReturn
    'CID_MSO15_64,CID_MSO15_32,CID_MSO14_64,CID_MSO14_32,CID_MSO12,CID_MSO11,
End Function 'GetCoreComponentCode

'-------------------------------------------------------------------------------
'   GetMsiFeatureName
'
'   Get the known Feature name for a core component
'-------------------------------------------------------------------------------
Function GetMsiFeatureName (sComponentId)
    Dim sReturn

    On Error Resume Next
    sReturn = ""
    Select Case sComponentId
    Case CID_ACC16_64,CID_ACC16_32,CID_ACC15_64,CID_ACC15_32,CID_ACC14_64,CID_ACC14_32,CID_ACC12,CID_ACC11
        sReturn = "ACCESSFiles"
    Case CID_XL16_64,CID_XL16_32,CID_XL15_64,CID_XL15_32,CID_XL14_64,CID_XL14_32,CID_XL12,CID_XL11
        sReturn = "EXCELFiles"
    Case CID_GRV16_64,CID_GRV16_32,CID_GRV15_64,CID_GRV15_32
        sReturn = "GrooveFiles2"
    Case CID_GRV14_64,CID_GRV14_32,CID_GRV12
        sReturn = "GrooveFiles"
    Case CID_LYN16_64,CID_LYN16_32,CID_LYN15_64,CID_LYN15_32
        sReturn = "Lync_CoreFiles"
    Case CID_MSO16_64,CID_MSO16_32,CID_MSO15_64,CID_MSO15_32,CID_MSO14_64,CID_MSO14_32,CID_MSO12,CID_MSO11
        sReturn = "ProductFiles"
    Case CID_ONE16_64,CID_ONE16_32,CID_ONE15_64,CID_ONE15_32,CID_ONE14_64,CID_ONE14_32,CID_ONE12,CID_ONE11
        sReturn = "OneNoteFiles"
    Case CID_OL16_64,CID_OL16_32,CID_OL15_64,CID_OL15_32,CID_OL14_64,CID_OL14_32,CID_OL12,CID_OL11
        sReturn = "OUTLOOKFiles"
    Case CID_PPT16_64,CID_PPT16_32,CID_PPT15_64,CID_PPT15_32,CID_PPT14_64,CID_PPT14_32,CID_PPT12,CID_PPT11
        sReturn = "PPTFiles"
    Case CID_PRJ16_64,CID_PRJ16_32,CID_PRJ15_64,CID_PRJ15_32,CID_PRJ14_64,CID_PRJ14_32,CID_PRJ12,CID_PRJ11
        sReturn = "PROJECTFiles"
    Case CID_PUB16_64,CID_PUB16_32,CID_PUB15_64,CID_PUB15_32,CID_PUB14_64,CID_PUB14_32,CID_PUB12,CID_PUB11
        sReturn = "PubPrimary"
    Case CID_IP16_64,CID_IP15_64,CID_IP15_32,CID_IP14_64,CID_IP14_32,CID_IP12,CID_IP11
        sReturn = "XDOCSFiles"
    Case CID_VIS16_64,CID_VIS16_32,CID_VIS15_64,CID_VIS15_32,CID_VIS14_64,CID_VIS14_32,CID_VIS12,CID_VIS11
        sReturn = "VisioCore"
    Case CID_WD16_64,CID_WD16_32,CID_WD15_64,CID_WD15_32,CID_WD14_64,CID_WD14_32,CID_WD12,CID_WD11
        sReturn = "WORDFiles"
    Case CID_SPD16_64,CID_SPD16_32,CID_SPD15_64,CID_SPD15_32,CID_SPD14_64,CID_SPD14_32,CID_SPD12
        sReturn = "WAC_CoreSPD"
    Case CID_SPD11
        sReturn = "FPClientFiles"
    End Select
    GetMsiFeatureName = sReturn
End Function 'GetMsiFeatureName

'-------------------------------------------------------------------------------
'   GetApplicationName
'
'   Get the friendly name from the FeatureName
'-------------------------------------------------------------------------------
Function GetApplicationName (sFeatureName)
    Dim sReturn

    On Error Resume Next
    sReturn = ""
    Select Case sFeatureName
    Case "ACCESSFiles"                  : sReturn = "Access"
    Case "EXCELFiles"                   : sReturn = "Excel"
    Case "GrooveFiles2", "GrooveFiles"  : sReturn = "Groove"
    Case "Lync_CoreFiles"               : sReturn = "Skype for Business"
    Case "OneNoteFiles"                 : sReturn = "OneNote"
    Case "OUTLOOKFiles"                 : sReturn = "Outlook"
    Case "PPTFiles"                     : sReturn = "PowerPoint"
    Case "ProductFiles"                 : sReturn = "Mso"
    Case "PROJECTFiles"                 : sReturn = "Project"
    Case "PubPrimary"                   : sReturn = "Publisher"
    Case "XDOCSFiles"                   : sReturn = "InfoPath"
    Case "VisioCore"                    : sReturn = "Visio"
    Case "WORDFiles"                    : sReturn = "Word"
    Case "WAC_CoreSPD", "FPClientFiles" : sReturn = "SharePoint Designer"
    End Select
    GetApplicationName = sReturn
End Function 'GetApplicationName

'-------------------------------------------------------------------------------
'   GetApplicationExeName
'
'   Get the friendly name from the FeatureName
'-------------------------------------------------------------------------------
Function GetApplicationExeName (sComponentId)
    Dim sReturn

    On Error Resume Next
    sReturn = ""
    Select Case sComponentId
    Case CID_ACC16_64,CID_ACC16_32,CID_ACC15_64,CID_ACC15_32,CID_ACC14_64,CID_ACC14_32,CID_ACC12,CID_ACC11
        sReturn = "MSACCESS.EXE"
    Case CID_XL16_64,CID_XL16_32,CID_XL15_64,CID_XL15_32,CID_XL14_64,CID_XL14_32,CID_XL12,CID_XL11
        sReturn = "EXCEL.EXE"
    Case CID_GRV16_64,CID_GRV16_32,CID_GRV15_64,CID_GRV15_32,CID_GRV14_64,CID_GRV14_32,CID_GRV12
        sReturn = "GROOVE.EXE"
    Case CID_LYN16_64,CID_LYN16_32,CID_LYN15_64,CID_LYN15_32
        sReturn = "LYNC.EXE"
    Case CID_MSO16_64,CID_MSO16_32,CID_MSO15_64,CID_MSO15_32,CID_MSO14_64,CID_MSO14_32,CID_MSO12,CID_MSO11
        sReturn = "MSO.DLL"
    Case CID_ONE16_64,CID_ONE16_32,CID_ONE15_64,CID_ONE15_32,CID_ONE14_64,CID_ONE14_32,CID_ONE12,CID_ONE11
        sReturn = "ONENOTE.EXE"
    Case CID_OL16_64,CID_OL16_32,CID_OL15_64,CID_OL15_32,CID_OL14_64,CID_OL14_32,CID_OL12,CID_OL11
        sReturn = "OUTLOOK.EXE"
    Case CID_PPT16_64,CID_PPT16_32,CID_PPT15_64,CID_PPT15_32,CID_PPT14_64,CID_PPT14_32,CID_PPT12,CID_PPT11
        sReturn = "POWERPNT.EXE"
    Case CID_PRJ16_64,CID_PRJ16_32,CID_PRJ15_64,CID_PRJ15_32,CID_PRJ14_64,CID_PRJ14_32,CID_PRJ12,CID_PRJ11
        sReturn = "WINPROJ.EXE"
    Case CID_PUB16_64,CID_PUB16_32,CID_PUB15_64,CID_PUB15_32,CID_PUB14_64,CID_PUB14_32,CID_PUB12,CID_PUB11
        sReturn = "MSPUB.EXE"
    Case CID_IP16_64,CID_IP15_64,CID_IP15_32,CID_IP14_64,CID_IP14_32,CID_IP12,CID_IP11
        sReturn = "INFOPATH.EXE"
    Case CID_VIS16_64,CID_VIS16_32,CID_VIS15_64,CID_VIS15_32,CID_VIS14_64,CID_VIS14_32,CID_VIS12,CID_VIS11
        sReturn = "VISIO.EXE"
    Case CID_WD16_64,CID_WD16_32,CID_WD15_64,CID_WD15_32,CID_WD14_64,CID_WD14_32,CID_WD12,CID_WD11
        sReturn = "WINWORD.EXE"
    Case CID_SPD16_64,CID_SPD16_32,CID_SPD15_64,CID_SPD15_32,CID_SPD14_64,CID_SPD14_32,CID_SPD12
        sReturn = "SPDESIGN.EXE"
    Case CID_SPD11
        sReturn = "FRONTPG.EXE"
    End Select
    GetApplicationExeName = sReturn
End Function 'GetApplicationExeName

'-------------------------------------------------------------------------------
'   GetConfigName
'
'   Get the configuration name from the ARP key name
'-------------------------------------------------------------------------------
Function GetConfigName(ArpItem)
    Dim sCurKey, sValue, sDisplayVersion, sUninstallString
    dim sCulture, sConfigName
    Dim iLeft, iRight
    Dim fSystemComponent0, fDisplayVersion, fUninstallString

    sCurKey = REG_ARP & ArpItem & "\"
    sValue = ""
    sDisplayVersion = ""
    fSystemComponent0 = NOT (RegReadValue(HKLM, sCurKey, "SystemComponent", sValue, "REG_DWORD") AND (sValue = "1"))
    fDisplayVersion = RegReadValue(HKLM, sCurKey, "DisplayVersion", sValue, "REG_SZ")
    If fDisplayVersion Then
        sDisplayVersion = sValue
        If Len(sValue) > 1 Then
            fDisplayVersion = (Left(sValue, 2) = "15")
        Else
            fDisplayVersion = False
        End If
    End If
    fUninstallString = RegReadValue(HKLM, sCurKey, "UninstallString", sUninstallString, "REG_SZ")

    'C2R
    If (fSystemComponent0 AND fDisplayVersion AND InStr(UCase(sUninstallString), UCase(OREGREFC2R15)) > 0) Then
        iLeft = InStr(ArpItem, " - ") + 2
        iRight = InStr(iLeft, ArpItem, " - ") - 1
        If iRight > 0 Then
            sConfigName = Trim(Mid(ArpItem, iLeft, (iRight - iLeft)))
            sCulture = Mid(ArpItem, iRight + 3)
        Else
            sConfigName = Trim(Left(ArpItem, iLeft - 3))
            sCulture = Mid(ArpItem, iLeft)
        End If
        sConfigName = Replace(sConfigName, "Microsoft", "")
        sConfigName = Replace(sConfigName, "Office", "")
        sConfigName = Replace(sConfigName, "Professional", "Pro")
        sConfigName = Replace(sConfigName, "Standard", "Std")
        sConfigName = Replace(sConfigName, "(Technical Preview)", "")
        sConfigName = Replace(sConfigName, "15", "")
        sConfigName = Replace(sConfigName, "2013", "")
        sConfigName = Replace(sConfigName, " ", "")
        'sConfigName = Replace(sConfigName, "Project", "Prj")
        'sConfigName = Replace(sConfigName, "Visio", "Vis")
        GetConfigName = sConfigName
        Exit Function
    End If

    'Standalone helper MSI products
    Select Case Mid(ArpItem,11,4)
    Case "007E", "008F", "008C"
        GetConfigName = "Habanero"
    Case "24E1", "237A"
        GetConfigName = "MSOIDLOGIN"
    Case Else
        GetConfigName = ""
    End Select

End Function 'GetConfigName

'=======================================================================================================

Sub CopyToMaster (Arr)
    Dim i,j,n
    On Error Resume Next
    
    For n = 0 To UBound(arrMaster)
        If (IsEmpty(arrMaster(n,COL_PRODUCTCODE)) OR Not (Len(arrMaster(n,COL_PRODUCTCODE))>0)) Then Exit For
    Next 'n
    
    For i = 0 To UBound(Arr,1)
        For j = 0 To UBound(Arr,2)
            arrMaster(i+n,j) = arr(i,j)
        Next 'j
    Next 'i
End Sub
'=======================================================================================================

Function GetObjProductState(ProdX,iContext,sSid,iSource)
    Dim iState
    On Error Resume Next

    Select Case iSource

    Case 0 'Registry
        iState = GetRegProductState(ProdX,iContext,sSid)
    Case 2 'WI 2.x

    Case 3 'WI >=3.x
        iState = ProdX.State

    Case Else
        iState = -1

    End Select
    GetObjProductState = iState
End Function
'=======================================================================================================

Function TranslateObjProductState(iState)
    Dim sState
    On Error Resume Next

    Select Case iState

    Case INSTALLSTATE_ADVERTISED '1
        sState = "Advertised"
    Case INSTALLSTATE_ABSENT '2
        sState = "Absent"
    Case INSTALLSTATE_DEFAULT '5
        sState = "Installed"    
    Case INSTALLSTATE_VIRTUALIZED '8
        sState = "Virtualized"  
    Case Else '-1
        sState = "Unknown"
    End Select
    TranslateObjProductState = sState
End Function
'=======================================================================================================

'Obtain the ProductName for product
Function GetObjProductName(ProdX,sProductCode,iContext,sSid,iSource)
    Dim sName
    On Error Resume Next

    Select Case iSource

    Case 0 'Registry
        sName = GetRegProductName(ProdX,iContext,sSid)
    Case 2 'WI 2.x

    Case 3 'WI >=3.x
        Err.Clear
        sName = ProdX.InstallProperty ("ProductName")
        If Not Err = 0 Then
            sName = GetRegProductName(GetCompressedGuid(sProductCode),iContext,sSid)
        End If 'Err

    Case Else

    End Select
    GetObjProductName = sName
End Function
'=======================================================================================================

'Get the current products UserSID.
Function GetObjUserSid(ProdX,sSid,iSource)
    On Error Resume Next

    Select Case iSource

    Case 0 'Registry
        sSid = sSid
    Case 2 'WI 2.x

    Case 3 'WI >=3.x
        sSid = ProdX.UserSid
    Case Else

    End Select
    
    GetObjUserSid = sSid
End Function
'=======================================================================================================

Function GetObjContext(ProdX,iContext,iSource)
    On Error Resume Next

    If iSource = 3 Then iContext = ProdX.Context
    GetObjContext = iContext
End Function
'=======================================================================================================

Function TranslateObjContext(iContext)
    Dim sContext
    On Error Resume Next

    Select Case iContext

    Case MSIINSTALLCONTEXT_USERMANAGED '1
        sContext = "User Managed"
    Case MSIINSTALLCONTEXT_USERUNMANAGED '2
        sContext = "User Unmanaged"
    Case MSIINSTALLCONTEXT_MACHINE '4
        sContext = "Machine"
    Case MSIINSTALLCONTEXT_ALL '7
        sContext = "All"
    Case MSIINSTALLCONTEXT_C2RV2 '8
        sContext = "Machine C2Rv2"
    Case MSIINSTALLCONTEXT_C2RV3 '15
        sContext = "Machine C2Rv3"
    Case Else
        sContext = "Unknown"

    End Select
    TranslateObjContext = sContext
End Function
'=======================================================================================================

Function GetObjGuid(ProdX,iSource)
    Dim sGuid
    On Error Resume Next

    Select Case iSource

    Case 0 'Registry
        If (IsValidGuid(ProdX,GUID_COMPRESSED) OR fGuidCaseWarningOnly) Then 
            sGuid = GetExpandedGuid(ProdX)
        Else
            sGuid = ProdX
        End If

    Case 2 'WI 2.x
        sGuid = ProdX

    Case 3 'WI >=3.x
        sGuid = ProdX.ProductCode

    Case Else

    End Select
    GetObjGuid = sGuid
End Function
'=======================================================================================================

Sub WritePLArrayEx(iSource,Arr,Obj,iContext,sSid)
    Dim ProdX,Product,sProductCode,sProductName
    Dim i, n, iDimCnt,iPosUMP
    On Error Resume Next
    
    If CheckObject(Obj) Or CheckArray(Obj) Then
        i = 0
        If CheckObject(Obj) Then 
            ReDim Arr(Obj.Count-1, UBOUND_MASTER)
        Else
            If UBound(Obj)>UBound(Arr) Then ReDim Arr(UBound(Obj), UBOUND_MASTER)
            Do While Not IsEmpty(Arr(i,0))
                i=i+1
                If i >= UBound(Obj) Then Exit Do
            Loop
        End If 'CheckObject
        For Each ProdX in Obj
        ' preset Recordset with Default Error
            For n = 0 to 4
                Arr(i,n) = "Preset Error String"
            Next 'n
        ' ProductCode
            Arr(i,COL_PRODUCTCODE) = GetObjGUID(ProdX,iSource)
        ' ProductContext
            Arr(i,COL_CONTEXT) = GetObjContext(ProdX,iContext,iSource)
            Arr(i,COL_CONTEXTSTRING) = TranslateObjContext(Arr(i,COL_CONTEXT))
        ' SID    
            Arr(i,COL_USERSID) = GetObjUserSid(ProdX,sSid,iSource)
        ' ProductName
            Arr(i,COL_PRODUCTNAME) = GetObjProductName(ProdX,Arr(i,COL_PRODUCTCODE),Arr(i,COL_CONTEXT),Arr(i,COL_USERSID),iSource)
        ' ProductState    
            Arr(i,COL_STATE) = GetObjProductState(ProdX,Arr(i,COL_CONTEXT),Arr(i,COL_USERSID),iSource)
            Arr(i,COL_STATESTRING) = TranslateObjProductState(Arr(i,COL_STATE))
        ' write to cache
            CacheLog LOGPOS_RAW,LOGHEADING_NONE,Null,Arr(i,COL_PRODUCTCODE) & "," & Arr(i,COL_CONTEXT) & "," & Arr(i,COL_USERSID) & "," & _
                Arr(i,COL_PRODUCTNAME) & "," & Arr(i,COL_STATE) 
        ' ARP ProductName    
            Arr(i,COL_ARPPRODUCTNAME) = GetArpProductname(Arr(i,COL_PRODUCTCODE))
        ' Guid validation
            If Not IsValidGuid(Arr(i,COL_PRODUCTCODE),GUID_UNCOMPRESSED) Then 
                If fGuidCaseWarningOnly Then 
                    Arr(i,COL_NOTES) = Arr(i,COL_NOTES) & ERR_CATEGORYNOTE & ERR_GUIDCASE & CSV
                Else
                    Arr(i,COL_ERROR) = Arr(i,COL_ERROR) & ERR_CATEGORYERROR & sError & CSV
                    Cachelog LOGPOS_REVITEM,LOGHEADING_NONE,ERR_CATEGORYERROR,sError & DSV & Arr(i,COL_PRODUCTCODE) & DSV & _
                    Arr(i,COL_PRODUCTNAME)& DSV & BPA_GUID
                End If
            End If
        ' Virtual flag
            Arr(i, COL_VIRTUALIZED) = 0
            If iContext = MSIINSTALLCONTEXT_C2RV2 Then Arr(i, COL_VIRTUALIZED) = 1
            If iContext = MSIINSTALLCONTEXT_C2RV3 Then Arr(i, COL_VIRTUALIZED) = 1
        ' InstallType flag
            Arr(i, COL_INSTALLTYPE) = "MSI"
            If iContext = MSIINSTALLCONTEXT_C2RV2 Then Arr(i, COL_INSTALLTYPE) = "C2R"
            If iContext = MSIINSTALLCONTEXT_C2RV3 Then Arr(i, COL_INSTALLTYPE) = "C2R"

            i = i + 1
        Next 'ProdX
    End If 'ObjectList.Count > 0
End Sub 'WritePLArrayEx
'=======================================================================================================

Sub InitPLArrays
    On Error Resume Next
    
    ReDim arrAllProducts(-1)
    ReDim arrVirtProducts(-1)
    ReDim arrVirt2Products(-1)
    ReDim arrVirt3Products(-1)
    ReDim arrMProducts(-1)
    ReDim arrMVProducts(-1)
    ReDim arrUUProducts(-1)
    ReDim arrUMProducts(-1)
End Sub

'=======================================================================================================
'Module Prerequisites
'=======================================================================================================
Function CheckPreReq ()
    On Error Resume Next
    Dim sActiveSub, sErrHnd
    Dim sPreReqError, sDebugLogName, sSubKeyName
    Dim i, iFoo
    Dim hDefKey, lAccPermLevel
    Dim oCompItem, oWmiLocal, oItem
    sActiveSub = "CheckPreReq" : sErrHnd = "_ErrorHandler"

    fIsCriticalError = False
    CheckPreReq = True : fIsAdmin = True : fIsElevated = True
    sPreReqError = vbNullString
    Err.Clear

    'Create the WScript Shell Object
    Set oShell = CreateObject("WScript.Shell"): CheckError sActiveSub,sErrHnd
    sComputerName = oShell.ExpandEnvironmentStrings("%COMPUTERNAME%"): CheckError sActiveSub,sErrHnd
    sTemp = oShell.ExpandEnvironmentStrings("%TEMP%"): CheckError sActiveSub,sErrHnd

    'Create the Windows Installer Object
    Set oMsi = CreateObject("WindowsInstaller.Installer"): CheckError sActiveSub,sErrHnd
    iWiVersionMajor = Left(oMsi.Version,Instr(oMsi.Version,".")-1)
    If (CheckPreReq = True And iWiVersionMajor < 2) Then CheckPreReq = False

    'Create the FileSystemObject
    Set oFso = CreateObject("Scripting.FileSystemObject"): CheckError sActiveSub,sErrHnd

    'Connect to WMI Registry Provider
    Set oReg = GetObject("winmgmts:\\.\root\default:StdRegProv"): CheckError sActiveSub,sErrHnd
    
    'Needs to be done here already as registry access calls depend on it
    Set oWmiLocal = GetObject("winmgmts:\\.\root\cimv2")
    Set oCompItem = oWmiLocal.ExecQuery("Select * from Win32_ComputerSystem")
    For Each oItem In oCompItem
        sSystemType = oItem.SystemType
        f64 = Instr(Left(oItem.SystemType,3),"64") > 0
    Next
    
    'Check registry access permissions
    'Failure will not terminate the scipt but noted in the log
    Set dicActiveC2Rv2Versions = CreateObject("Scripting.Dictionary")
    hDefKey = HKEY_LOCAL_MACHINE
    sSubKeyName = "SOFTWARE\Microsoft\Windows"
    For i = 1 To 4
        Select Case i
        Case 1 : lAccPermLevel = KEY_QUERY_VALUE
        Case 2 : lAccPermLevel = KEY_SET_VALUE
        Case 3 : lAccPermLevel = KEY_CREATE_SUB_KEY
        Case 4 : lAccPermLevel = DELETE
        End Select
        
        If Not RegCheckAccess(hDefKey,sSubKeyName,lAccPermLevel) Then
            fIsAdmin = False
            Exit for
        End If
    Next 'i
    
    If fIsAdmin Then
        sSubKeyName = "Software\Microsoft\Windows\"
        For i = 1 To 4
            Select Case i
            Case 1 : lAccPermLevel = KEY_QUERY_VALUE
            Case 2 : lAccPermLevel = KEY_SET_VALUE
            Case 3 : lAccPermLevel = KEY_CREATE_SUB_KEY
            Case 4 : lAccPermLevel = DELETE
            End Select
            
            If Not RegCheckAccess(hDefKey,sSubKeyName,lAccPermLevel) Then
                fIsElevated = False
                Exit for
            End If
        Next 'i
    End If 'fIsAdmin

    
    Set ShellApp = CreateObject("Shell.Application") 

    If Not sPreReqError = vbNullString Then
        If fQuiet = False Then
            Msgbox "Script execution needs to terminate" & vbCrLf & vbCrLf & sPreReqError, vbOkOnly, _
            "Critical Error in Script Prerequisite Check"
        End If
        CheckPreReq = False
    End If ' sPreReqError = vbNullString
End Function 'CheckPreReq()
'=======================================================================================================

Sub CheckPreReq_ErrorHandler

    sPreReqError = sPreReqError & _
        sDebugErr & " returned:" & vbCrLf &_
        "Error Details: " & Err.Source & " " & Hex( Err ) & ": " & Err.Description & vbCrLf & vbCrLf
End Sub

'=======================================================================================================
'Module ComputerProperties
'=======================================================================================================

Sub ComputerProperties
    Dim oOS, oWmi, oOsItem
    Dim sOSinfo, sOSVersion, sUserInfo, sSubKeyName, sName, sValue, sOsMui, sOsLcid, sCulture
    Dim arrKeys, arrNames, arrTypes, arrVersion
    Dim qOS, OsLang, ValueType
    Dim iOSVersion, iValueName
    Dim hDefKey
    Const REG_SZ = 1
    On Error Resume Next
    
    sComputerName = oShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
    
    'Note 64 bit OS check was already done in 'Sub Initialize'
    
    'OS info from WMI Win32_OperatingSystem
    Set oWMI = GetObject("winmgmts:\\.\root\cimv2") 
    Set qOS = oWmi.ExecQuery("Select * from Win32_OperatingSystem") 
    For Each oOsItem in qOS 
        sOSinfo = sOSinfo & oOsItem.Caption 
        sOSinfo = sOSinfo & oOsItem.OtherTypeDescription
        sOSinfo = sOSinfo & CSV & "SP " & oOsItem.ServicePackMajorVersion
        sOSinfo = sOSinfo & CSV & "Version: " & oOsItem.Version
        sOsVersion = oOsItem.Version
        sOSinfo = sOSinfo & CSV & "Codepage: " & oOsItem.CodeSet
        sOSinfo = sOSinfo & CSV & "Country Code: " & oOsItem.CountryCode
        sOSinfo = sOSinfo & CSV & "Language: " & oOsItem.OSLanguage
    Next
    sOSinfo = sOSinfo & CSV & "System Type: " & sSystemType 
    
    'Check for OS MUI languages
    'The MUI registry location has been changed with Windows Vista
    'Win 2000, XP and Server 2003: HKLM\System\CurrentControlSet\Control\Nls\MUILanguages
    'From Vista on: HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\MUI\UILanguages
    
    'Build the VersionNT number
    arrVersion = Split(sOsVersion,Delimiter(sOsVersion))
    iVersionNt = CInt(arrVersion(0))*100 + CInt(arrVersion(1))
    
    hDefKey = HKEY_LOCAL_MACHINE 
    If iVersionNt < 600 Then 
        ' "old" reg location
        sSubKeyName = "System\CurrentControlSet\Control\Nls\MUILanguages\" 
        If RegEnumValues (hDefKey,sSubKeyName,arrNames,arrTypes) Then
            For iValueName = 0 To UBound(arrNames) 
                If arrTypes(iValueName) = REG_SZ Then
                    sName = "": sName = arrNames(iValueName)
                    sOsMui = sOsMui & GetCultureInfo(CInt("&h"&sName)) & " (" & CInt("&h"&sName) & ")" & CSV
                End If
            Next 'ValueName
        sOsMui = RTrimComma(sOsMui) 
        End If 'IsArray
    Else
        ' "new" reg location
        sSubKeyName = "SYSTEM\CurrentControlSet\Control\MUI\UILanguages\"
        sName = "LCID"
        If RegEnumKey(hDefKey,sSubKeyName,arrKeys) Then
            For Each OsLang in arrKeys 
                If Len(sOsMui) > 1 Then sOsMui = sOsMui & CSV
                sOsMui = sOsMui & OsLang 
                If RegReadDWordValue(hDefKey,sSubKeyName&OsLang,sName,sValue) Then sOsMui = sOsMui & " (" & sValue & ")"
            Next 'OsLang
        End If 'RegKeyExists
    End If
    
    'User info from WMI Win32_ComputerSystem
    Set qOS = oWmi.ExecQuery("Select * from Win32_ComputerSystem") 
    For Each oOsItem in qOS
        sUserInfo = sUserInfo & "Username: " & oOsItem.UserName 
    Next
    sUserInfo = sUserInfo & CSV & "IsAdmin: " & fIsAdmin
    sUserInfo = sUserInfo & CSV & "SID: " & sCurUserSid
    
    If NOT fBasicMode Then CacheLog LOGPOS_COMPUTER,LOGHEADING_NONE,"Windows Installer Version",oMsi.Version
    CacheLog LOGPOS_COMPUTER,LOGHEADING_NONE,"ComputerName",sComputerName
    CacheLog LOGPOS_COMPUTER,LOGHEADING_NONE,"OS Details",sOSinfo 
    If NOT fBasicMode Then CacheLog LOGPOS_COMPUTER,LOGHEADING_NONE,"OS MUI Languages",sOsMui 
    If NOT fBasicMode Then CacheLog LOGPOS_COMPUTER,LOGHEADING_NONE,"Current User",sUserInfo 
    If NOT fBasicMode Then CacheLog LOGPOS_COMPUTER,LOGHEADING_NONE,"Logfile Name",sLogFile
    CacheLog LOGPOS_COMPUTER,LOGHEADING_NONE,"ROI Script Build",SCRIPTBUILD
End Sub 'ComputerProperties

'=======================================================================================================
'Module Logging
'=======================================================================================================

'=======================================================================================================

Sub WriteLog
    Dim LogOutput
    Dim i,n
    Dim sScriptSettings
    On Error Resume Next
    
    'Add actual values for customizable settings
    sScriptSettings = ""
    If fListNonOfficeProducts Then sScriptSettings = sScriptSettings & "/All "
    If fLogFull Then sScriptSettings = sScriptSettings & "/Full "
    If fLogVerbose Then sScriptSettings = sScriptSettings & "/LogVerbose "
    If fLogChainedDetails Then sScriptSettings = sScriptSettings & "/LogChainedDetails "
    If fFileInventory Then sScriptSettings = sScriptSettings & "/FileInventory "
    If fFeatureTree Then sScriptSettings = sScriptSettings & "/FeatureTree "
    If fBasicMode Then sScriptSettings = "/Basic "
    If fQuiet Then sScriptSettings = sScriptSettings & "/Quiet "
    If Not sPathOutputFolder = "" Then sScriptSettings = sScriptSettings & "/Logfolder: "&sPathOutputFolder
    CacheLog LOGPOS_COMPUTER,LOGHEADING_NONE,"Script Settings",sScriptSettings 
    
    'Determine total scan time
    tEnd = Time()
    If NOT fBasicMode Then CacheLog LOGPOS_COMPUTER,LOGHEADING_NONE,"Total scan time",Int((tEnd - tStart)*1000000 + 0.5)/10 & " s"

    Set LogOutput = oFso.OpenTextFile(sLogFile, FOR_WRITING, True, True)
    LogOutput.WriteLine "Microsoft Customer Support Services - Robust Office Inventory - " & Now
    LogOutput.Write vbCrLf & String(160,"*") & vbCrLf

    'Flush the prepared output arrays to the log
    n = 2
    'Enable raw output if an error was encountered
    If NOT Len(arrLog(1)) = 32 AND NOT fBasicMode Then n = 3
    For i = 0 to n
        If i = 1 Then
            If NOT Len(arrLog(i)) = 32 Then 
                LogOutput.Write arrLog(i)
                LogOutput.Write vbCrLf & String(160,"*") & vbCrLf
            End If
        Else
            LogOutput.Write arrLog(i)
            If NOT (i = 2) Then LogOutput.Write vbCrLf & String(160,"*") & vbCrLf
        End If
    Next 'i
    LogOutput.Close
    Set LogOutput = Nothing
    
    'Copy the log to the fileinventory folder if needed
    If fFileInventory Then oFso.CopyFile sLogFile,sPathOutputFolder&"ROIScan\", TRUE

    ' write the xml log
    WriteXmlLog

End Sub
'=======================================================================================================

Sub WriteXmlLog
    Dim XmlLogStream, mspSeq, key, lic, component, item
    Dim sXmlLine, sText, sProductCode, sSPLevel, sFamily, sSeq
    Dim i, iVPCnt, iItem, iArpCnt, iDummy, iPos, iPosMaster, iChainProd, iPosPatch, iColPatch
    Dim iColISource, iLogCnt
    Dim arrC2RPackages, arrC2RItems, arrTmp, arrTmpInner, arrLicData, arrKeyComponents
    Dim arrComponentData, arrVAppState
    Dim dicXmlTmp, dicApps
    Dim fLogProduct

    On Error Resume Next
    
    Set dicXmlTmp = CreateObject("Scripting.Dictionary")
    Set dicApps = CreateObject("Scripting.Dictionary")

    Set XmlLogStream = oFso.CreateTextFile(sPathOutputFolder & sComputerName & "_ROIScan.xml", True, True)
    XmlLogStream.WriteLine "<?xml version=""1.0""?>"
    XmlLogStream.WriteLine "<OFFICEINVENTORY>"

    'c2r v3
    If UBound(arrVirt3Products) > -1 Then
        sXmlLine = "<C2RShared "
        For Each key in dicC2RPropV3.Keys
            sXmlLine = sXmlLine & " " & Replace(key, " ", "") & "=" & chr(34) & dicC2RPropV3.Item(key) & chr(34)
        Next
        ' end line
        sXmlLine = sXmlLine & " />"
        'flush
        XmlLogStream.WriteLine sXmlLine

        For iVPCnt = 0 To UBound(arrVirt3Products, 1)
            sXmlLine = ""
            sXmlLine = "<SKU "
            ' ProductName (heading)
            sXmlLine = sXmlLine & "ProductName=" & chr(34) & arrVirt3Products (iVPCnt, COL_PRODUCTNAME) & chr(34)
            ' KeyName
            sXmlLine = sXmlLine & " KeyName=" & chr(34) & arrVirt3Products (iVPCnt, VIRTPROD_KEYNAME) & chr(34)
            ' ConfigName
            sXmlLine = sXmlLine & " ConfigName=" & chr(34) & arrVirt3Products (iVPCnt, VIRTPROD_CONFIGNAME) & chr(34)
            sXmlLine = sXmlLine & " IsChainedChild=" & chr(34) & "FALSE" & chr(34)
            ' ProductCode
            sXmlLine = sXmlLine & " ProductCode=" & chr(34) & "" & chr(34)
            ' ProductVersion
            sXmlLine = sXmlLine & " ProductVersion=" & chr(34) & dicC2RPropV3.Item(STR_VERSION) & chr(34)
            ' SP Level
            sXmlLine = sXmlLine & " ServicePackLevel=" & chr(34) & arrVirt3Products (iVPCnt, VIRTPROD_SPLEVEL) & chr(34)
            ' Architecture
            sXmlLine = sXmlLine & " Architecture=" & chr(34) & dicC2RPropV3.Item(STR_PLATFORM) & chr(34)
            ' InstallType
            sXmlLine = sXmlLine & " InstallType=" & chr(34) & "C2R" & chr(34)
            ' C2R integration ProductCode
            sXmlLine = sXmlLine & " C2rIntegrationProductCode=" & chr(34) & dicC2RPropV3.Item(STR_PACKAGEGUID) & chr(34)
            ' end line
            sXmlLine = sXmlLine & " >"
            'flush
            XmlLogStream.WriteLine sXmlLine

            ' Child Packages
            XmlLogStream.WriteLine "<ChildPackages>"
            arrC2RPackages = Split (arrVirt3Products(iVPCnt, VIRTPROD_CHILDPACKAGES), ";")
            If UBound(arrC2RPackages) > 0 Then
                For i = 0 To UBound(arrC2RPackages) - 1 'strip off last delimiter
                    arrC2RItems = Split(arrC2RPackages(i), " - ")
                    If UBound(arrC2RItems) = 1 Then
                        If NOT i = 0 Then XmlLogStream.WriteLine "</SkuCulture>"
                        XmlLogStream.WriteLine "<SkuCulture culture=" & chr(34) & arrC2RItems(0) & chr(34) & " ProductVersion=" & chr(34) & arrC2RItems(1) & chr(34) & " >"
                    Else
                        XmlLogStream.WriteLine "<ChildPackage ProductCode=" & chr(34) & arrC2RItems(0) & chr(34) & " ProductVersion=" & chr(34) & arrC2RItems(1) & chr(34) & " ProductName=" & chr(34) & arrC2RItems(2) & chr(34) & " />"
                    End If
                Next 'i
                XmlLogStream.WriteLine "</SkuCulture>"
            End If 'UBound(arrC2RPackages) > 0
            XmlLogStream.WriteLine "</ChildPackages>"
            
            'KeyComponents
            MapAppsToConfigProduct dicApps, arrVirt3Products(iVPCnt, VIRTPROD_CONFIGNAME), 16
            XmlLogStream.WriteLine "<KeyComponents>"
            For Each key in dicApps
                Set arrVAppState = Nothing
                'sText = "Installed - "
                arrVAppState = Split(dicKeyComponentsV3.Item(key), ",")
                sXmlLine = "<Application "
                sXmlLine = sXmlLine & " Name=" & chr(34) & key & chr(34)
                sXmlLine = sXmlLine & " ExeName=" & chr(34) & arrVAppState (1) & chr(34)
                sXmlLine = sXmlLine & " VersionMajor=" & chr(34) & arrVAppState (2) & chr(34)
                sXmlLine = sXmlLine & " Version=" & chr(34) & arrVAppState (3) & chr(34)
                sXmlLine = sXmlLine & " InstallState=" & chr(34) & arrVAppState (4) & chr(34)
                sXmlLine = sXmlLine & " InstallStateString=" & chr(34) & arrVAppState (5) & chr(34)
                sXmlLine = sXmlLine & " ComponentId=" & chr(34) & arrVAppState (6) & chr(34)
                sXmlLine = sXmlLine & " FeatureName=" & chr(34) & arrVAppState (7) & chr(34)
                sXmlLine = sXmlLine & " Path=" & chr(34) & arrVAppState (8) & chr(34)
                sXmlLine = sXmlLine & " />"
                XmlLogStream.WriteLine sXmlLine
            Next 'key
            XmlLogStream.WriteLine "</KeyComponents>"

            'LicenseData
            XmlLogStream.WriteLine "<LicenseData>"
            XmlLogStream.WriteLine arrVirt3Products(iVPCnt, VIRTPROD_OSPPLICENSEXML)
            XmlLogStream.WriteLine "</LicenseData>"
            
            XmlLogStream.WriteLine "</SKU>"
        Next 'iVPCnt
    End If 'c2r v2


    'c2r v2
    If UBound(arrVirt2Products) > -1 Then
        sXmlLine = "<C2RShared "
        For Each key in dicC2RPropV2.Keys
            sXmlLine = sXmlLine & " " & Replace(key, " ", "") & "=" & chr(34) & dicC2RPropV2.Item(key) & chr(34)
        Next
        ' end line
        sXmlLine = sXmlLine & " />"
        'flush
        XmlLogStream.WriteLine sXmlLine

        For iVPCnt = 0 To UBound(arrVirt2Products, 1)
            sXmlLine = ""
            sXmlLine = "<SKU "
            ' ProductName (heading)
            sXmlLine = sXmlLine & "ProductName=" & chr(34) & arrVirt2Products (iVPCnt, COL_PRODUCTNAME) & chr(34)
            ' KeyName
            sXmlLine = sXmlLine & " KeyName=" & chr(34) & arrVirt2Products (iVPCnt, VIRTPROD_KEYNAME) & chr(34)
            ' ConfigName
            sXmlLine = sXmlLine & " ConfigName=" & chr(34) & arrVirt2Products (iVPCnt, VIRTPROD_CONFIGNAME) & chr(34)
            sXmlLine = sXmlLine & " IsChainedChild=" & chr(34) & "FALSE" & chr(34)
            ' ProductCode
            sXmlLine = sXmlLine & " ProductCode=" & chr(34) & "" & chr(34)
            ' ProductVersion
            sXmlLine = sXmlLine & " ProductVersion=" & chr(34) & arrVirt2Products (iVPCnt, VIRTPROD_PRODUCTVERSION) & chr(34)
            ' SP Level
            sXmlLine = sXmlLine & " ServicePackLevel=" & chr(34) & arrVirt2Products (iVPCnt, VIRTPROD_SPLEVEL) & chr(34)
            ' Architecture
            sXmlLine = sXmlLine & " Architecture=" & chr(34) & dicC2RPropV2.Item(STR_PLATFORM) & chr(34)
            ' InstallType
            sXmlLine = sXmlLine & " InstallType=" & chr(34) & "C2R" & chr(34)
            ' C2R integration ProductCode
            sXmlLine = sXmlLine & " C2rIntegrationProductCode=" & chr(34) & dicC2RPropV2.Item(STR_PACKAGEGUID) & chr(34)
            ' end line
            sXmlLine = sXmlLine & " >"
            'flush
            XmlLogStream.WriteLine sXmlLine

            ' Child Packages
            XmlLogStream.WriteLine "<ChildPackages>"
            arrC2RPackages = Split (arrVirt2Products(iVPCnt, VIRTPROD_CHILDPACKAGES), ";")
            If UBound(arrC2RPackages) > 0 Then
                For i = 0 To UBound(arrC2RPackages) - 1 'strip off last delimiter
                    arrC2RItems = Split(arrC2RPackages(i), " - ")
                    If UBound(arrC2RItems) = 1 Then
                        If NOT i = 0 Then XmlLogStream.WriteLine "</SkuCulture>"
                        XmlLogStream.WriteLine "<SkuCulture culture=" & chr(34) & arrC2RItems(0) & chr(34) & " ProductVersion=" & chr(34) & arrC2RItems(1) & chr(34) & " >"
                    Else
                        XmlLogStream.WriteLine "<ChildPackage ProductCode=" & chr(34) & arrC2RItems(0) & chr(34) & " ProductVersion=" & chr(34) & arrC2RItems(1) & chr(34) & " ProductName=" & chr(34) & arrC2RItems(2) & chr(34) & " />"
                    End If
                Next 'i
                XmlLogStream.WriteLine "</SkuCulture>"
            End If 'UBound(arrC2RPackages) > 0
            XmlLogStream.WriteLine "</ChildPackages>"
            
            'KeyComponents
            MapAppsToConfigProduct dicApps, arrVirt2Products(iVPCnt, VIRTPROD_CONFIGNAME), 15
            XmlLogStream.WriteLine "<KeyComponents>"
            For Each key in dicApps
                Set arrVAppState = Nothing
                'sText = "Installed - "
                arrVAppState = Split(dicKeyComponentsV2.Item(key), ",")
                sXmlLine = "<Application "
                sXmlLine = sXmlLine & " Name=" & chr(34) & key & chr(34)
                sXmlLine = sXmlLine & " ExeName=" & chr(34) & arrVAppState (1) & chr(34)
                sXmlLine = sXmlLine & " VersionMajor=" & chr(34) & arrVAppState (2) & chr(34)
                sXmlLine = sXmlLine & " Version=" & chr(34) & arrVAppState (3) & chr(34)
                sXmlLine = sXmlLine & " InstallState=" & chr(34) & arrVAppState (4) & chr(34)
                sXmlLine = sXmlLine & " InstallStateString=" & chr(34) & arrVAppState (5) & chr(34)
                sXmlLine = sXmlLine & " ComponentId=" & chr(34) & arrVAppState (6) & chr(34)
                sXmlLine = sXmlLine & " FeatureName=" & chr(34) & arrVAppState (7) & chr(34)
                sXmlLine = sXmlLine & " Path=" & chr(34) & arrVAppState (8) & chr(34)
                sXmlLine = sXmlLine & " />"
                XmlLogStream.WriteLine sXmlLine
            Next 'key
            XmlLogStream.WriteLine "</KeyComponents>"
            
            'LicenseData
            XmlLogStream.WriteLine "<LicenseData>"
            XmlLogStream.WriteLine arrVirt2Products(iVPCnt, VIRTPROD_OSPPLICENSEXML)
            XmlLogStream.WriteLine "</LicenseData>"
            
            XmlLogStream.WriteLine "</SKU>"
        Next 'iVPCnt
    End If 'c2r v2

    'c2r v1
    If UBound(arrVirtProducts) > -1 Then
        For iVPCnt = 0 To UBound(arrVirtProducts, 1)
            sXmlLine = ""
            sXmlLine = "<SKU "
            ' ProductName (heading)
            sXmlLine = sXmlLine & "ProductName=" & chr(34) & arrVirtProducts(iVPCnt, COL_PRODUCTNAME) & chr(34)
            ' ProductVersion
            sXmlLine = sXmlLine & " ProductVersion=" & chr(34) & arrVirtProducts(iVPCnt, VIRTPROD_PRODUCTVERSION) & chr(34)
            ' SP Level
            sXmlLine = sXmlLine & " ServicePackLevel=" & chr(34) & arrVirtProducts(iVPCnt, VIRTPROD_SPLEVEL) & chr(34)
            ' ProductCode
            sXmlLine = sXmlLine & " ProductCode=" & chr(34) & arrVirtProducts(iVPCnt, COL_PRODUCTCODE) & chr(34)
            ' end line
            sXmlLine = sXmlLine & " >"
            'flush
            XmlLogStream.WriteLine sXmlLine
            XmlLogStream.WriteLine "</SKU>"
        Next 'iVPCnt
    End If 'c2r v1

    'MSI Office > v12 Products
    For iArpCnt = 0 To UBound(arrArpProducts)
'        If arrArpProducts(iArpCnt, COL_CONFIGINSTALLTYPE) = "VIRTUAL" Then
'            'do nothing
'        Else
        ' extra loop to allow exit out of a bad product
        iPos = GetArrayPosition(arrMaster, arrArpProducts(iArpCnt, ARP_CONFIGPRODUCTCODE))
        sProductCode = "" 
        If NOT iPos = -1 Then sProductCode = arrMaster(iPos, COL_PRODUCTCODE)
        sXmlLine = ""
        sXmlLine = "<SKU "
        ' ProductName (heading)
        sText = "" : sText = GetArpProductName(arrArpProducts(iArpCnt, COL_CONFIGNAME)) 
        If (sText = "" AND iPos = -1) Then sText = arrMaster(iPos, COL_ARPPRODUCTNAME) 
        sXmlLine = sXmlLine & "ProductName=" & chr(34) & sText & chr(34)
        ' Configuration ProductName
        sText = "" : sText = arrArpProducts(iArpCnt, COL_CONFIGNAME)
        If InStr(sText,".") > 0 Then sText = Mid(sText, InStr(sText, ".") + 1)
        sXmlLine = sXmlLine & " ConfigName=" & chr(34) & sText & chr(34)
        sXmlLine = sXmlLine & " IsChainedChild=" & chr(34) & "FALSE" & chr(34)
        ' ProductCode
        sXmlLine = sXmlLine & " ProductCode=" & chr(34) & "" & chr(34)
        ' ProductVersion
        sXmlLine = sXmlLine & " ProductVersion=" & chr(34) & arrArpProducts(iArpCnt, ARP_PRODUCTVERSION) & chr(34)
        ' ServicePack
        sSPLevel = ""
        If NOT iPos = -1 AND NOT sProductCode = "" AND Len(arrArpProducts(iArpCnt, ARP_PRODUCTVERSION)) > 2 Then
            sSPLevel = OVersionToSpLevel(sProductCode, GetVersionMajor(sProductCode), arrArpProducts(iArpCnt, ARP_PRODUCTVERSION))
        End If
        sXmlLine = sXmlLine & " ServicePackLevel=" & chr(34) & sSPLevel  & chr(34)
        ' Architecture
        sXmlLine = sXmlLine & " Architecture=" & chr(34) & arrMaster(iPos, COL_ARCHITECTURE) & chr(34)
        ' InstallType
        sXmlLine = sXmlLine & " InstallType=" & chr(34) & arrMaster(iPos, COL_INSTALLTYPE) & chr(34)
        ' InstallDate
        sXmlLine = sXmlLine & " InstallDate=" & chr(34) & Left(Replace(arrMaster(iPos, COL_INSTALLDATE), " ", ""), 8) & chr(34)
        ' ProductState
        sXmlLine = sXmlLine & " ProductState=" & chr(34) & arrMaster(iPos, COL_STATESTRING) & chr(34)
        ' ConfigMsi
        sXmlLine = sXmlLine & " ConfigMsi=" & chr(34) & arrMaster(iPos, COL_ORIGINALMSI) & chr(34)
        ' BuildOrigin
        sXmlLine = sXmlLine & " BuildOrigin=" & chr(34) & arrMaster(iPos, COL_ORIGIN) & chr(34)
        ' end line
        sXmlLine = sXmlLine & " >"
        'flush
        XmlLogStream.WriteLine sXmlLine

        ' Child Packages
        XmlLogStream.WriteLine "<ChildPackages>"
            For iChainProd = COL_LBOUNDCHAINLIST To UBound(arrArpProducts, 2)
                If IsEmpty(arrArpProducts(iArpCnt, iChainProd)) Then
                    'log the master entry only to ensure a consistent xml data structure
                    sXmlLine =  "<ChildPackage"
                    'ProductCode
                    sXmlLine = sXmlLine & " ProductCode=" & chr(34) & sProductCode & chr(34)
                    'ProductVersion
                    sXmlLine = sXmlLine & " ProductVersion=" & chr(34) & arrMaster(iPosMaster, COL_PRODUCTVERSION) & chr(34)
                    'ProductName
                    sText = "" : sText = arrMaster(iPosMaster,COL_ARPPRODUCTNAME)
                    If sText = "" Then sText = arrMaster(iPosMaster,COL_PRODUCTNAME)
                    sXmlLine = sXmlLine & " ProductName=" & chr(34) & sText & chr(34)
                    ' end line
                    sXmlLine = sXmlLine & " />"
                    'flush
                    XmlLogStream.WriteLine sXmlLine
                    Exit For
                Else
                    iPosMaster = GetArrayPosition(arrMaster,arrArpProducts(iArpCnt,iChainProd))
                    'Only run if iPosMaster has a valid index #
                    sXmlLine =  "<ChildPackage"
                    'ProductCode
                    sText = "" : If NOT iPosMaster = -1 Then sText = arrMaster(iPosMaster, COL_PRODUCTCODE)
                    sXmlLine = sXmlLine & " ProductCode=" & chr(34) & sText & chr(34)
                    'ProductVersion
                    sText = "" : If NOT iPosMaster = -1 Then sText = arrMaster(iPosMaster, COL_PRODUCTVERSION)
                    sXmlLine = sXmlLine & " ProductVersion=" & chr(34) & sText & chr(34)
                    'ProductName
                    sText = "" : If NOT iPosMaster = -1 Then sText = arrMaster(iPosMaster, COL_PRODUCTNAME)
                    sXmlLine = sXmlLine & " ProductName=" & chr(34) & sText & chr(34)
                    ' end line
                    sXmlLine = sXmlLine & " />"
                    'flush
                    XmlLogStream.WriteLine sXmlLine
                End If
            Next 'iChainProd
        XmlLogStream.WriteLine "</ChildPackages>"

        'KeyComponents
        If Len(arrMaster(iPos, COL_KEYCOMPONENTS)) > 0 Then
            If Right(arrMaster(iPos, COL_KEYCOMPONENTS), 1) = ";" Then arrMaster(iPos, COL_KEYCOMPONENTS) = Left(arrMaster(iPos, COL_KEYCOMPONENTS), Len(arrMaster(iPos, COL_KEYCOMPONENTS)) - 1 )
        End If
        XmlLogStream.WriteLine "<KeyComponents>"
        arrKeyComponents = Split(arrMaster(iPos, COL_KEYCOMPONENTS), ";")
        For Each component in arrKeyComponents
            arrComponentData = Split(component, ",")
            If CheckArray(arrComponentData) Then
            sXmlLine = "<Application "
            sXmlLine = sXmlLine & " Name=" & chr(34) & arrComponentData(0) & chr(34)
            sXmlLine = sXmlLine & " ExeName=" & chr(34) & arrComponentData(1) & chr(34)
            sXmlLine = sXmlLine & " VersionMajor=" & chr(34) & arrComponentData(2) & chr(34)
            sXmlLine = sXmlLine & " Version=" & chr(34) & arrComponentData(3) & chr(34)
            sXmlLine = sXmlLine & " InstallState=" & chr(34) & arrComponentData(4) & chr(34)
            sXmlLine = sXmlLine & " InstallStateString=" & chr(34) & arrComponentData(5) & chr(34)
            sXmlLine = sXmlLine & " ComponentId=" & chr(34) & arrComponentData(6) & chr(34)
            sXmlLine = sXmlLine & " FeatureName=" & chr(34) & arrComponentData(7) & chr(34)
            sXmlLine = sXmlLine & " Path=" & chr(34) & arrComponentData(8) & chr(34)
            sXmlLine = sXmlLine & " />"
            XmlLogStream.WriteLine sXmlLine
            End If
        Next 'component
        XmlLogStream.WriteLine "</KeyComponents>"

        'LicenseData
        '-----------
        XmlLogStream.WriteLine "<LicenseData>"
        If NOT iPos = -1 Then
            If arrMaster(iPos, COL_OSPPLICENSE) <> "" Then
                XmlLogStream.WriteLine arrMaster(iPos, COL_OSPPLICENSEXML)
            End If
        End If
        XmlLogStream.WriteLine "</LicenseData>"

        ' Patches
        XmlLogStream.WriteLine "<PatchData>"
        ' PatchBaseLines
        XmlLogStream.WriteLine "<PatchBaseline Sequence=" & chr(34) & arrArpProducts(iArpCnt, ARP_PRODUCTVERSION) & chr(34) & " >"
        dicXmlTmp.RemoveAll
        For iChainProd = COL_LBOUNDCHAINLIST To UBound(arrArpProducts, 2)
            If IsEmpty(arrArpProducts(iArpCnt, iChainProd)) Then Exit For
            iPosMaster = GetArrayPosition(arrMaster, arrArpProducts(iArpCnt, iChainProd))
            'Only run if iPosMaster has a valid index #
            If Not iPosMaster = -1 Then
                Set arrTmp = Nothing
                arrTmp = Split(arrMaster(iPosMaster, COL_PATCHFAMILY), ",")
                For Each MspSeq in arrTmp
                    arrTmpInner = Split(MspSeq, ":")
                    sFamily = "" : sSeq = ""
                    sFamily = arrTmpInner(0)
                    sSeq  = arrTmpInner(1)
                    If (sSeq>arrMaster(iPosMaster,COL_PRODUCTVERSION)) Then 
                        If dicXmlTmp.Exists(sFamily) Then
                            If (sSeq > dicXmlTmp.Item(sFamily)) Then dicXmlTmp.Item(sFamily)=sSeq
                        Else
                            dicXmlTmp.Add sFamily,sSeq
                        End If
                    End If
                Next 'MspSeq
            End If 'Not iPosMaster = -1
        Next 'iChainProd
        For Each key in dicXmlTmp.Keys
            XmlLogStream.WriteLine "<PostBaseline PatchFamily=" & chr(34) & key & chr(34) & " Sequence=" & chr(34) & dicXmlTmp.Item(key) & chr(34) & " />"
        Next 'key
        XmlLogStream.WriteLine "</PatchBaseline>"

        ' PatchList
        For iChainProd = COL_LBOUNDCHAINLIST To UBound(arrArpProducts, 2)
            If IsEmpty(arrArpProducts(iArpCnt,iChainProd)) Then Exit For
            iPosMaster = GetArrayPosition(arrMaster, arrArpProducts(iArpCnt, iChainProd))
            'Only run if iPosMaster has a valid index #
            If Not iPosMaster = -1 Then
                For iPosPatch = 0 to UBound(arrPatch, 3)
                    If Not IsEmpty (arrPatch(iPosMaster, PATCH_PATCHCODE, iPosPatch)) Then
                        sXmlLine = "<Patch PatchedProduct=" & chr(34) & arrMaster(iPosMaster, COL_PRODUCTCODE) & chr(34)
                        For iColPatch = PATCH_LOGSTART to PATCH_LOGCHAINEDMAX
                            If Not IsEmpty(arrPatch(iPosMaster, iColPatch, iPosPatch)) Then
                                sXmlLine =  sXmlLine & " " & Replace(arrLogFormat(ARRAY_PATCH,iColPatch), ": ", "") & "=" & chr(34) & arrPatch(iPosMaster,iColPatch,iPosPatch) & chr(34)
                            End If
                        Next 'iColPatch
                        sXmlLine =  sXmlLine & " />"
                        XmlLogStream.WriteLine sXmlLine
'                            Set arrTmp = Nothing
'                            arrTmp = Split(arrMaster(iPosMaster, COL_PATCHFAMILY), ",")
'                            For Each MspSeq in arrTmp
'                                arrTmpInner = Split(MspSeq, ":")
'                                XmlLogStream.WriteLine "<MsiPatchSequence PatchFamily=" & chr(34) & arrTmpInner(0) & chr(34) & " Sequence=" & chr(34) & arrTmpInner(1) & chr(34) & " />"
'                            Next 'MspSeq
'                            XmlLogStream.WriteLine "</Patch>"
                    End If 'IsEmpty
                Next 'iPosPatch
            End If ' Not iPosMaster = -1
        Next 'iChainProd

        XmlLogStream.WriteLine "</PatchData>"

        'InstallSource
        XmlLogStream.WriteLine "<InstallSource>"
        For iColISource = IS_LOG_LBOUND To IS_LOG_UBOUND
            If NOT iPos = -1 Then
                If Not IsEmpty(arrIS(iPos, iColISource)) Then _ 
                XmlLogStream.WriteLine "<Source " & Replace(arrLogFormat(ARRAY_IS,iColISource)," ", "") & "=" & chr(34) & arrIS(iPos, iColISource) & chr(34) & " />"
            End If
        Next 'iColISource
        XmlLogStream.WriteLine "</InstallSource>"

        XmlLogStream.WriteLine "</SKU>"
'        End If
    Next 'iArpCnt

    'Other Products
    Err.Clear
    For iLogCnt = 0 To 12
        For iPosMaster = 0 To UBound(arrMaster)
            fLogProduct = CheckLogProduct(iLogCnt, iPosMaster)
            If fLogProduct Then
                'arrMaster contents
                sXmlLine = ""
                sXmlLine = "<SKU "
                ' ProductName (heading)
                sText = "" : sText = arrMaster(iPosMaster,COL_ARPPRODUCTNAME)
                If sText = "" Then sText = arrMaster(iPosMaster,COL_PRODUCTNAME)
                sXmlLine = sXmlLine & "ProductName=" & chr(34) & sText & chr(34)
                sXmlLine = sXmlLine & " ConfigName=" & chr(34) & "" & chr(34)
                Select Case iLogCnt
                Case 1, 3, 5
                    sXmlLine = sXmlLine & " IsChainedChild=" & chr(34) & "TRUE" & chr(34)
                Case Else
                    sXmlLine = sXmlLine & " IsChainedChild=" & chr(34) & "FALSE" & chr(34)
                End Select
                ' ProductCode
                sXmlLine = sXmlLine & " ProductCode=" & chr(34) & arrMaster(iPosMaster, COL_PRODUCTCODE) & chr(34)
                ' ProductVersion
                sXmlLine = sXmlLine & " ProductVersion=" & chr(34) & arrMaster(iPosMaster, COL_PRODUCTVERSION) & chr(34)
                ' ServicePack
                sXmlLine = sXmlLine & " ServicePackLevel=" & chr(34) & arrMaster(iPosMaster, COL_SPLEVEL) & chr(34)
                ' Architecture
                sXmlLine = sXmlLine & " Architecture=" & chr(34) & arrMaster(iPosMaster, COL_ARCHITECTURE) & chr(34)
                ' InstallType
                sXmlLine = sXmlLine & " InstallType=" & chr(34) & arrMaster(iPosMaster, COL_INSTALLTYPE) & chr(34)
                ' InstallDate
                sXmlLine = sXmlLine & " InstallDate=" & chr(34) & Left(Replace(arrMaster(iPosMaster, COL_INSTALLDATE), " ", ""), 8) & chr(34)
                ' ProductState
                sXmlLine = sXmlLine & " ProductState=" & chr(34) & arrMaster(iPosMaster, COL_STATESTRING) & chr(34)
                ' ConfigMsi
                sXmlLine = sXmlLine & " ConfigMsi=" & chr(34) & arrMaster(iPosMaster, COL_ORIGINALMSI) & chr(34)
                ' BuildOrigin
                sXmlLine = sXmlLine & " BuildOrigin=" & chr(34) & arrMaster(iPosMaster, COL_ORIGIN) & chr(34)
                ' end line
                sXmlLine = sXmlLine & " >"
                'flush
                XmlLogStream.WriteLine sXmlLine

                ' Child Packages
                XmlLogStream.WriteLine "<ChildPackages>"
                    'log the master entry only to ensure a consistent xml data structure
                    sXmlLine =  "<ChildPackage"
                    'ProductCode
                    sXmlLine = sXmlLine & " ProductCode=" & chr(34) & arrMaster(iPosMaster, COL_PRODUCTCODE) & chr(34)
                    'ProductVersion
                    sXmlLine = sXmlLine & " ProductVersion=" & chr(34) & arrMaster(iPosMaster, COL_PRODUCTVERSION) & chr(34)
                    'ProductName
                    sText = "" : sText = arrMaster(iPosMaster,COL_ARPPRODUCTNAME)
                    If sText = "" Then sText = arrMaster(iPosMaster,COL_PRODUCTNAME)
                    sXmlLine = sXmlLine & " ProductName=" & chr(34) & sText & chr(34)
                    ' end line
                    sXmlLine = sXmlLine & " />"
                    'flush
                    XmlLogStream.WriteLine sXmlLine
                XmlLogStream.WriteLine "</ChildPackages>"

                'KeyComponents
                If Len(arrMaster(iPosMaster, COL_KEYCOMPONENTS)) > 0 Then
                    If Right(arrMaster(iPosMaster, COL_KEYCOMPONENTS), 1) = ";" Then arrMaster(iPosMaster, COL_KEYCOMPONENTS) = Left(arrMaster(iPosMaster, COL_KEYCOMPONENTS), Len(arrMaster(iPosMaster, COL_KEYCOMPONENTS)) - 1 )
                End If
                XmlLogStream.WriteLine "<KeyComponents>"
                arrKeyComponents = Split(arrMaster(iPosMaster, COL_KEYCOMPONENTS), ";")
                For Each component in arrKeyComponents
                    arrComponentData = Split(component, ",")
                    If CheckArray(arrComponentData) Then
                    sXmlLine = "<Application "
                    sXmlLine = sXmlLine & " Name=" & chr(34) & arrComponentData(0) & chr(34)
                    sXmlLine = sXmlLine & " ExeName=" & chr(34) & arrComponentData(1) & chr(34)
                    sXmlLine = sXmlLine & " VersionMajor=" & chr(34) & arrComponentData(2) & chr(34)
                    sXmlLine = sXmlLine & " Version=" & chr(34) & arrComponentData(3) & chr(34)
                    sXmlLine = sXmlLine & " InstallState=" & chr(34) & arrComponentData(4) & chr(34)
                    sXmlLine = sXmlLine & " InstallStateString=" & chr(34) & arrComponentData(5) & chr(34)
                    sXmlLine = sXmlLine & " ComponentId=" & chr(34) & arrComponentData(6) & chr(34)
                    sXmlLine = sXmlLine & " FeatureName=" & chr(34) & arrComponentData(7) & chr(34)
                    sXmlLine = sXmlLine & " Path=" & chr(34) & arrComponentData(8) & chr(34)
                    sXmlLine = sXmlLine & " />"
                    XmlLogStream.WriteLine sXmlLine
                    End If
                Next 'component
                XmlLogStream.WriteLine "</KeyComponents>"

                ' Patches
                XmlLogStream.WriteLine "<PatchData>"

                ' PatchBaseLines
                XmlLogStream.WriteLine "<PatchBaseline Sequence=" & chr(34) & arrMaster(iPosMaster, COL_PRODUCTVERSION) & chr(34) & " >"
                dicXmlTmp.RemoveAll
                Set arrTmp = Nothing
                arrTmp = Split(arrMaster(iPosMaster, COL_PATCHFAMILY), ",")
                For Each MspSeq in arrTmp
                    arrTmpInner = Split(MspSeq, ":")
                    sFamily = "" : sSeq = ""
                    sFamily = arrTmpInner(0)
                    sSeq  = arrTmpInner(1)
                    If (sSeq>arrMaster(iPosMaster, COL_PRODUCTVERSION)) Then 
                        If dicXmlTmp.Exists(sFamily) Then
                            If (sSeq > dicXmlTmp.Item(sFamily)) Then dicXmlTmp.Item(sFamily) = sSeq
                        Else
                            dicXmlTmp.Add sFamily, sSeq
                        End If
                    End If
                Next 'MspSeq
                For Each key in dicXmlTmp.Keys
                    XmlLogStream.WriteLine "<PostBaseline PatchFamily=" & chr(34) & key & chr(34) & " Sequence=" & chr(34) & dicXmlTmp.Item(key) & chr(34) & " />"
                Next 'key
                XmlLogStream.WriteLine "</PatchBaseline>"

                For iPosPatch = 0 to UBound(arrPatch, 3)
                    If Not IsEmpty (arrPatch(iPosMaster, PATCH_PATCHCODE, iPosPatch)) Then
                        sXmlLine = "<Patch PatchedProduct=" & chr(34) & arrMaster(iPosMaster, COL_PRODUCTCODE) & chr(34)
                        For iColPatch = PATCH_LOGSTART to PATCH_LOGCHAINEDMAX
                            If Not IsEmpty(arrPatch(iPosMaster, iColPatch, iPosPatch)) Then
                                sXmlLine =  sXmlLine & " " & Replace(arrLogFormat(ARRAY_PATCH, iColPatch), ": ", "") & "=" & chr(34) & arrPatch(iPosMaster, iColPatch, iPosPatch) & chr(34)
                            End If
                        Next 'iColPatch
                        sXmlLine =  sXmlLine & " />"
                        XmlLogStream.WriteLine sXmlLine
                    End If 'IsEmpty
                Next 'iPosPatch
                
                XmlLogStream.WriteLine "</PatchData>"

                'InstallSource
                XmlLogStream.WriteLine "<InstallSource>"
                For iColISource = IS_LOG_LBOUND To IS_LOG_UBOUND
                    If Not IsEmpty(arrIS(iPos, iColISource)) Then _ 
                    XmlLogStream.WriteLine "<Source " & Replace(arrLogFormat(ARRAY_IS,iColISource)," ", "") & "=" & chr(34) & arrIS(iPosMaster, iColISource) & chr(34) & " />"
                Next 'iColISource
                XmlLogStream.WriteLine "</InstallSource>"

                XmlLogStream.WriteLine "</SKU>"
            End If 'fLogProduct
        Next 'iPosMaster
    Next 'iLogCnt

    XmlLogStream.WriteLine "</OFFICEINVENTORY>"

End Sub 'WriteXmlLog

'=======================================================================================================

Sub PrepareLog (sLogFormat)
    Dim Key, MspSeq, Lic, ScenarioItem
    Dim i, j, k, m, iLBound, iUBound
    Dim iAip, iPos, iPosMaster, iPosArp, iArpCnt, iChainProd, iPosOrder, iDummy
    Dim iPosPatch, iColPatch, iColISource, iPosISource, iLogCnt, iVPCnt
    Dim sTmp, sText, sTextErr, sTextNotes, sCategory, sSeq, sFamily, sFamilyMain, sFamilyLang
    Dim sProdCache, sLcid, sVer
    Dim bCspCondition, fLogProduct, fDataIntegrity
    Dim fLoggedVirt, fLoggedMulti, fLoggedSingle
    Dim arrOrder(), arrTmp, arrTmpInner, arrLicData, arrC2RPackages, arrVAppState
    Dim dicFamilyLang, dicApps, dicTmp, dicMspFamVer
    On Error Resume Next

    fDataIntegrity = True
    fLoggedMulti = False
    fLoggedSingle = False
    fLoggedVirt = False

    'Filter contents for the log
    i = -1
    Redim arrOrder(18)
    i = i + 1 : arrOrder(i)  = COL_PRODUCTVERSION
    i = i + 1 : arrOrder(i)  = COL_SPLEVEL
    i = i + 1 : arrOrder(i)  = COL_ARCHITECTURE
    i = i + 1 : arrOrder(i)  = COL_INSTALLTYPE
    i = i + 1 : arrOrder(i)  = COL_PRODUCTCODE
    i = i + 1 : arrOrder(i)  = COL_PRODUCTNAME
    i = i + 1 : arrOrder(i)  = COL_ARPPARENTS
    i = i + 1 : arrOrder(i)  = COL_INSTALLDATE
    i = i + 1 : arrOrder(i)  = COL_USERSID
    i = i + 1 : arrOrder(i)  = COL_CONTEXTSTRING
    i = i + 1 : arrOrder(i)  = COL_STATESTRING
    i = i + 1 : arrOrder(i) = COL_TRANSFORMS
    i = i + 1 : arrOrder(i) = COL_ORIGINALMSI
    i = i + 1 : arrOrder(i) = COL_CACHEDMSI
    i = i + 1 : arrOrder(i) = COL_PRODUCTID
    i = i + 1 : arrOrder(i) = COL_ORIGIN
    i = i + 1 : arrOrder(i) = COL_PACKAGECODE
    i = i + 1 : arrOrder(i) = COL_NOTES
    i = i + 1 : arrOrder(i) = COL_ERROR

' trim content fields
    For iPosMaster = 0 To UBound(arrMaster)
        arrMaster(iPosMaster,COL_NOTES) = Trim(arrMaster(iPosMaster,COL_NOTES)) 
        arrMaster(iPosMaster,COL_ERROR) = Trim(arrMaster(iPosMaster,COL_ERROR))
        If Right(arrMaster(iPosMaster,COL_NOTES),1) = "," Then arrMaster(iPosMaster,COL_NOTES) = Left(arrMaster(iPosMaster,COL_NOTES),Len(arrMaster(iPosMaster,COL_NOTES))-1)
        If Right(arrMaster(iPosMaster,COL_ERROR),1) = "," Then arrMaster(iPosMaster,COL_ERROR) = Left(arrMaster(iPosMaster,COL_ERROR),Len(arrMaster(iPosMaster,COL_ERROR))-1)
        If arrMaster(iPosMaster,COL_NOTES) = "" Then arrMaster(iPosMaster,COL_NOTES) = "-"
        If arrMaster(iPosMaster,COL_ERROR) = "" Then arrMaster(iPosMaster,COL_ERROR) = "-"
        If Left(arrMaster(iPosMaster,COL_ARPPARENTS),1)="," Then arrMaster(iPosMaster,COL_ARPPARENTS)=Replace(arrMaster(iPosMaster,COL_ARPPARENTS),",","",1,1,1)
    Next 'iPosMaster
    
    Set dicTmp = CreateObject("Scripting.Dictionary")
    Set dicFamilyLang = CreateObject("Scripting.Dictionary")
    Set dicMspFamVer = CreateObject("Scripting.Dictionary")
    Set dicApps = CreateObject("Scripting.Dictionary")
    
' prepare Virtualized C2R Products
' --------------------------------
    If UBound(arrVirtProducts) > -1 OR UBound(arrVirt2Products) > -1 OR UBound(arrVirt3Products) > -1 Then
        CacheLog LOGPOS_PRODUCT, LOGHEADING_H1, Null, "C2R Products" 
        fLoggedVirt = True
        
        'O16 C2R
        If UBound(arrVirt3Products) > -1 Then
            CacheLog LOGPOS_PRODUCT, LOGHEADING_H2, Null, "C2R Shared Properties" 
                
            'Shared Properties
            For Each key in dicC2RPropV3.Keys
                Select Case key
                Case STR_REGPACKAGEGUID
                Case STR_CDNBASEURL
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "", ""
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "Install", ""
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, key, dicC2RPropV3.Item(key)
                Case STR_UPDATESENABLED
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "", ""
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "Updates", ""
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, key, dicC2RPropV3.Item(key)
                Case STR_UPDATETOVERSION
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "", ""
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, key, dicC2RPropV3.Item(key)
                Case Else
                    sTmp = dicC2RPropV3.Item(key)
                    If NOT sTmp = "" Then CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, key, dicC2RPropV3.Item(key)
                End Select
            Next


            ' Scenario key state
            CacheLog LOGPOS_PRODUCT, LOGHEADING_H2, Null, "Scenario Key State" 
            For Each ScenarioItem in dicScenarioV3.Keys
                CacheLog LOGPOS_PRODUCT,LOGHEADING_NONE, ScenarioItem, dicScenarioV3.Item(ScenarioItem)
            Next 'ScenarioItem

            sProdCache = ""
            For iVPCnt = 0 To UBound(arrVirt3Products, 1)
                If NOT InStr(sProdCache, arrVirt3Products(iVPCnt, VIRTPROD_CONFIGNAME)) > 0 Then
                    sProdCache = sProdCache & arrVirt3Products(iVPCnt, VIRTPROD_CONFIGNAME) & ","

                    ' ProductName (heading)
                    sText = "" : sText = arrVirt3Products(iVPCnt, COL_PRODUCTNAME)
                    If InStr(sText, " - ") > 0 AND (Len(sText) = InStr(sText, " - ") + 7) Then sText = Left(sText, InStr(sText, " - ") - 1)
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_H2, Null, sText 
                    ' ProductVersion
                    sText = "" : sText = arrVirt3Products(iVPCnt, VIRTPROD_PRODUCTVERSION)
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, arrLogFormat(ARRAY_VIRTPROD, VIRTPROD_PRODUCTVERSION), sText
                    ' SP Level
                    sText = "" : sText = arrVirt3Products(iVPCnt, VIRTPROD_SPLEVEL)
                    ' ConfigName
                    sText = "" : sText = arrVirt3Products(iVPCnt, VIRTPROD_CONFIGNAME)
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, arrLogFormat(ARRAY_VIRTPROD, VIRTPROD_CONFIGNAME), sText
                    ' Languages
                    i = 0
                    For Each key in dicVirt3Cultures.Keys
                        If i = 0 Then sText = "Language(s)" Else sText = ""
                        CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, sText, key
                        i = i + 1
                    Next 'key
                    
                    'Application States
                    '------------------
                    MapAppsToConfigProduct dicApps, arrVirt3Products(iVPCnt, VIRTPROD_CONFIGNAME), 16
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "", ""
                    For Each key in dicApps
                        Set arrVAppState = Nothing
                        sText = "Installed - "
                        arrVAppState = Split(dicKeyComponentsV3.Item(key), ",")
                        If NOT arrVAppState(4) = "3" Then
                            sText = "Absent/Excluded - "
                        Else
                            sText = sText & arrVAppState(3) & " - " & arrVAppState(6) & " - " & arrVAppState(7) & " - " & arrVAppState(8)
                        End If
                        CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, key, sText
                    Next 'key
                    
                    'LicenseData
                    '-----------
                    If arrVirt3Products(iVPCnt, VIRTPROD_OSPPLICENSE) <> "" Then
                        arrLicData = Split(arrVirt3Products(iVPCnt, VIRTPROD_OSPPLICENSE), "#;#")
                        If CheckArray(arrLicData) Then
                            If NOT fBasicMode Then CacheLog LOGPOS_PRODUCT,LOGHEADING_NONE, "", ""
                            i = 0
                            For Each Lic in arrLicData
                                arrTmp = Split(Lic, ";")
                                If LCase(arrTmp(0)) = "active license" Then i = 1
                                If i < 2 Then
                                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, arrTmp(0), arrTmp(1)
                                Else
                                    If NOT (fBasicMode AND i > 5) Then
                                        CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "", arrTmp(0) & ":" & Space(33 - Len(arrTmp(0))) & arrTmp(1)
                                    End If
                                End If
                                i = i + 1
                            Next 'Lic
                        End If 'arrLicData
                    End If 'arrMaster
                End If 'sProdCache
            Next 'iVPCnt
        'end C2R v3

        'O15 C2R
        ElseIf UBound(arrVirt2Products) > -1 Then
            CacheLog LOGPOS_PRODUCT, LOGHEADING_H2, Null, "C2R Shared Properties" 

            'Shared Properties
            For Each key in dicC2RPropV2.Keys
                Select Case key
                Case STR_REGPACKAGEGUID
                Case STR_CDNBASEURL
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "", ""
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "Install", ""
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, key, dicC2RPropV2.Item(key)
                Case STR_UPDATESENABLED
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "", ""
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "Updates", ""
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, key, dicC2RPropV2.Item(key)
                Case STR_UPDATETOVERSION
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "", ""
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, key, dicC2RPropV2.Item(key)
                Case Else
                    sTmp = dicC2RPropV2.Item(key)
                    If NOT sTmp = "" Then CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, key, dicC2RPropV2.Item(key)
                End Select
            Next


            ' Scenario key state
            CacheLog LOGPOS_PRODUCT, LOGHEADING_H2, Null, "Scenario Key State" 
            For Each ScenarioItem in dicScenarioV2.Keys
                CacheLog LOGPOS_PRODUCT,LOGHEADING_NONE, ScenarioItem, dicScenarioV2.Item(ScenarioItem)
            Next 'ScenarioItem

            sProdCache = ""
            For iVPCnt = 0 To UBound(arrVirt2Products, 1)
                If NOT InStr(sProdCache, arrVirt2Products(iVPCnt, VIRTPROD_CONFIGNAME)) > 0 Then
                    sProdCache = sProdCache & arrVirt2Products(iVPCnt, VIRTPROD_CONFIGNAME) & ","

                    ' ProductName (heading)
                    sText = "" : sText = arrVirt2Products(iVPCnt, COL_PRODUCTNAME)
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_H2, Null, sText 
                    ' ProductVersion
                    sText = "" : sText = arrVirt2Products(iVPCnt, VIRTPROD_PRODUCTVERSION)
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, arrLogFormat(ARRAY_VIRTPROD, VIRTPROD_PRODUCTVERSION), sText
                    ' SP Level
                    sText = "" : sText = arrVirt2Products(iVPCnt, VIRTPROD_SPLEVEL)
                    ' ConfigName
                    sText = "" : sText = arrVirt2Products(iVPCnt, VIRTPROD_CONFIGNAME)
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, arrLogFormat(ARRAY_VIRTPROD, VIRTPROD_CONFIGNAME), sText
                    ' Languages
                    i = 0
                    For Each key in dicVirt2Cultures.Keys
                        If i = 0 Then sText = "Language(s)" Else sText = ""
                        CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, sText, key
                        i = i + 1
                    Next 'key
                    ' Child Packages
                    arrC2RPackages = Split(arrVirt2Products(iVPCnt, VIRTPROD_CHILDPACKAGES), ";")
                    For i = 0 To UBound(arrC2RPackages)
                        If i = 0 Then CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "Chained Packages", arrC2RPackages(i) _
                        Else CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "", arrC2RPackages(i)
                    Next 'i
                    'Application States
                    '------------------
                    MapAppsToConfigProduct dicApps, arrVirt2Products(iVPCnt, VIRTPROD_CONFIGNAME), 15
                    For Each key in dicApps
                        Set arrVAppState = Nothing
                        sText = "Installed - "
                        arrVAppState = Split(dicKeyComponentsV2.Item(key), ",")
                        If NOT arrVAppState(4) = "3" Then
                            sText = "Absent/Excluded - "
                        Else
                            sText = sText & arrVAppState(3) & " - " & arrVAppState(6) & " - " & arrVAppState(7) & " - " & arrVAppState(8)
                        End If
                        CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, key, sText
                    Next 'key

                    'LicenseData
                    '-----------
                    If arrVirt2Products(iVPCnt, VIRTPROD_OSPPLICENSE) <> "" Then
                        arrLicData = Split(arrVirt2Products(iVPCnt,VIRTPROD_OSPPLICENSE), "#;#")
                        If CheckArray(arrLicData) Then
                            If NOT fBasicMode Then CacheLog LOGPOS_PRODUCT,LOGHEADING_NONE,"",""
                            i = 0
                            For Each Lic in arrLicData
                                arrTmp = Split(Lic, ";")
                                If LCase(arrTmp(0)) = "active license" Then i = 1
                                If i < 2 Then
                                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, arrTmp(0), arrTmp(1)
                                Else
                                    If NOT (fBasicMode AND i > 5) Then
                                        CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "", arrTmp(0) & ":" & Space(25 - Len(arrTmp(0))) & arrTmp(1)
                                    End If
                                End If
                                i = i + 1
                            Next 'Lic
                        End If 'arrLicData
                    End If 'arrMaster
                End If 'sProdCache
            Next 'iVPCnt

        ElseIf UBound(arrVirtProducts) > -1 Then
            For iVPCnt = 0 To UBound(arrVirtProducts, 1)
                sText = "" : sText = arrVirtProducts(iVPCnt,COL_PRODUCTNAME)
                CacheLog LOGPOS_PRODUCT,LOGHEADING_H2,Null,sText 
                sText = "" : sText = arrVirtProducts(iVPCnt,VIRTPROD_PRODUCTVERSION)
                CacheLog LOGPOS_PRODUCT,LOGHEADING_NONE,arrLogFormat(ARRAY_VIRTPROD,VIRTPROD_PRODUCTVERSION),sText
                sText = "" : sText = arrVirtProducts(iVPCnt,VIRTPROD_SPLEVEL)
                CacheLog LOGPOS_PRODUCT,LOGHEADING_NONE,arrLogFormat(ARRAY_VIRTPROD,VIRTPROD_SPLEVEL),sText
                sText = "" : sText = arrVirtProducts(iVPCnt,COL_PRODUCTCODE)
                CacheLog LOGPOS_PRODUCT,LOGHEADING_NONE,arrLogFormat(ARRAY_VIRTPROD,COL_PRODUCTCODE),sText
            Next 'iVPCnt
        End If
        CacheLog LOGPOS_PRODUCT,LOGHEADING_NONE,"",vbCrLf & String(160,"*") & vbCrLf
    Else
        'No virtualized products found
    End If
    
    'Prepare Office > 12 Products
    '----------------------------
    If CheckArray(arrArpProducts) Then
        CacheLog LOGPOS_PRODUCT, LOGHEADING_H1, Null, "Chained Products View" 
        fLoggedMulti = True

    For iArpCnt = 0 To UBound(arrArpProducts)
    ' extra loop to allow exit out of a bad product
        For iDummy = 1 To 1
        ' get the link to the configuration .MSI
            iPos = GetArrayPosition(arrMaster,arrArpProducts(iArpCnt,ARP_CONFIGPRODUCTCODE))
            If iPos = -1 Then
                If NOT dicMissingChild.Exists(arrArpProducts(iArpCnt,ARP_CONFIGPRODUCTCODE)) Then Cachelog LOGPOS_REVITEM,LOGHEADING_NONE,ERR_CATEGORYERROR,ERR_DATAINTEGRITY
            End If
            'If possible use the entry as shown under ARP as heading for the log
            sText = "" : sText = GetArpProductName(arrArpProducts(iArpCnt,COL_CONFIGNAME)) 
            If (sText = "" AND iPos=-1) Then sText = arrMaster(iPos,COL_ARPPRODUCTNAME) 
            CacheLog LOGPOS_PRODUCT,LOGHEADING_H2,Null,sText 

            'Contents from master array
            '--------------------------
            For iPosOrder = 0 To UBound(arrOrder)-2 '-2 to exclude Notes and Error fields
                If fBasicMode AND iPosOrder > 1 Then Exit For
                sText = "" : If NOT iPos=-1 Then sText = arrMaster(iPos,arrOrder(iPosOrder)) 
                'Suppress Configuration SKU for ARP config products. Log all others
                If NOT arrOrder(iPosOrder) = COL_ARPPARENTS Then CacheLog LOGPOS_PRODUCT,LOGHEADING_NONE,arrLogFormat(ARRAY_MASTER,arrOrder(iPosOrder)),sText
            Next 'iPosOrder

            If NOT fBasicMode Then
                'Add Notes and Errors from chained products to ArpProd column
                sTextNotes = "" : sTextErr = ""
                For iChainProd = COL_LBOUNDCHAINLIST To UBound(arrArpProducts,2)
                    If IsEmpty(arrArpProducts(iArpCnt,iChainProd)) Then Exit For
                    iPosMaster = GetArrayPosition(arrMaster,arrArpProducts(iArpCnt,iChainProd))
                    If iPosMaster = -1 AND NOT dicMissingChild.Exists(arrArpProducts(iArpCnt,iChainProd)) Then Cachelog LOGPOS_REVITEM,LOGHEADING_NONE,ERR_CATEGORYERROR,ERR_DATAINTEGRITY
                    'Only run if iPosMaster has a valid index #
                    If Not iPosMaster = -1 Then
                        If Not arrMaster(iPosMaster,COL_NOTES) = "-" Then
                            sTextNotes = sTextNotes & arrMaster(iPosMaster,COL_PRODUCTCODE) & DSV & arrMaster(iPosMaster,COL_NOTES)
                        End If 'arrMaster(iPosMaster,COL_NOTES) = "-"
                        If Not arrMaster(iPosMaster,COL_ERROR) = "-" Then
                            sTextErr = sTextErr & arrMaster(iPosMaster,COL_PRODUCTCODE) & DSV & arrMaster(iPosMaster,COL_ERROR)
                        End If 'arrMaster(iPosMaster,COL_NOTES) = "-"
                    End If 'Not iPosMaster = -1
                Next 'iChainProd
                If sTextNotes = "" Then sTextNotes = "-"
                If sTextErr = "" Then sTextErr = "-"
                CacheLog LOGPOS_PRODUCT,LOGHEADING_NONE,arrLogFormat(ARRAY_MASTER,COL_NOTES),sTextNotes
                CacheLog LOGPOS_PRODUCT,LOGHEADING_NONE,arrLogFormat(ARRAY_MASTER,COL_ERROR),sTextErr
            
                'Configuration ProductName
                '-------------------------
                sText = "" : sText = arrArpProducts(iArpCnt,COL_CONFIGNAME)
                If InStr(sText,".")>0 Then sText = Mid(sText,InStr(sText,".")+1)
                CacheLog LOGPOS_PRODUCT,LOGHEADING_NONE,arrLogFormat(ARRAY_ARP,COL_CONFIGNAME),sText
            
                'Configuration PackageID
                sText = "" : sText = arrArpProducts(iArpCnt,COL_CONFIGPACKAGEID)
                CacheLog LOGPOS_PRODUCT,LOGHEADING_NONE,arrLogFormat(ARRAY_ARP,COL_CONFIGPACKAGEID),sText
            
                'Chained packages
                '----------------
                CacheLog LOGPOS_PRODUCT,LOGHEADING_NONE,"",""
                sCategory = "Chained Packages"
                For iChainProd = COL_LBOUNDCHAINLIST To UBound(arrArpProducts,2)
                    If IsEmpty(arrArpProducts(iArpCnt,iChainProd)) Then Exit For
                    iPosMaster = GetArrayPosition(arrMaster,arrArpProducts(iArpCnt,iChainProd))
                    If iPosMaster = -1  AND NOT dicMissingChild.Exists(arrArpProducts(iArpCnt,iChainProd)) Then Cachelog LOGPOS_REVITEM,LOGHEADING_NONE,ERR_CATEGORYERROR,ERR_DATAINTEGRITY
                    'Only run if iPosMaster has a valid index #
                    If Not iPosMaster = -1 Then
                        sText = "" : sText = arrMaster(iPosMaster,COL_PRODUCTCODE) & DSV & arrMaster(iPosMaster,COL_PRODUCTVERSION) & DSV & arrMaster(iPosMaster,COL_PRODUCTNAME)
                        CacheLog LOGPOS_PRODUCT,LOGHEADING_NONE,sCategory,sText
                        sCategory = ""
                    Else
                        sText = "" : sText = arrArpProducts(iArpCnt,iChainProd) & DSV & "Error - missing chained product!"
                        CacheLog LOGPOS_PRODUCT,LOGHEADING_NONE,sCategory,sText
                        sCategory = ""
                    End If 'Not iPosMaster = -1
                Next 'iChainProd
            End If 'fBasicMode

            'LicenseData
            '-----------
            If NOT iPos = -1 Then
                If arrMaster(iPos,COL_OSPPLICENSE) <> "" Then
                    arrLicData = Split(arrMaster(iPos,COL_OSPPLICENSE),"#;#")
                    If CheckArray(arrLicData) Then
                        If NOT fBasicMode Then CacheLog LOGPOS_PRODUCT,LOGHEADING_NONE,"",""
                        i = 0
                        For Each Lic in arrLicData
                            arrTmp = Split(Lic,";")
                            If LCase(arrTmp(0)) = "active license" Then i = 1
                            If i < 2 Then
                                CacheLog LOGPOS_PRODUCT,LOGHEADING_NONE,arrTmp(0),arrTmp(1)
                            Else
                                If NOT (fBasicMode AND i > 5) Then
                                    CacheLog LOGPOS_PRODUCT,LOGHEADING_NONE,"",arrTmp(0)&":"& Space(25-Len(arrTmp(0)))&arrTmp(1)
                                End If
                            End If
                            i = i + 1
                        Next 'Lic
                    End If 'arrLicData
                End If 'arrMaster
            End If 'iPos -1
            
            If NOT fBasicMode Then
                'Patches
                '-------
                CacheLog LOGPOS_PRODUCT,LOGHEADING_NONE, "", ""
                sCategory = "Patch Baseline"
                sText = "" : If NOT iPos = -1 Then sText = arrMaster(iPos, COL_PRODUCTVERSION)
                CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, sCategory, sText
            
                sCategory = "Post Baseline Sequences"
                dicTmp.RemoveAll
                dicMspFamVer.RemoveAll
                For iChainProd = COL_LBOUNDCHAINLIST To UBound (arrArpProducts, 2)
                    If IsEmpty (arrArpProducts(iArpCnt, iChainProd)) Then Exit For
                    iPosMaster = GetArrayPosition (arrMaster, arrArpProducts(iArpCnt, iChainProd))
                    If iPosMaster = -1  AND NOT dicMissingChild.Exists (arrArpProducts(iArpCnt, iChainProd)) Then Cachelog LOGPOS_REVITEM, LOGHEADING_NONE, ERR_CATEGORYERROR, ERR_DATAINTEGRITY
                    
                    If Not iPosMaster = -1 Then
                        Set arrTmp = Nothing
                        arrTmp = Split (arrMaster(iPosMaster, COL_PATCHFAMILY), ",")
                        For Each MspSeq in arrTmp
                            arrTmpInner = Split (MspSeq, ":")
                            sFamily = "" : sSeq = "" : sFamilyMain = "" : sFamilyLang = ""
                            sFamily = arrTmpInner(0)
                            sSeq  = arrTmpInner(1)
                            If (sSeq > arrMaster(iPosMaster, COL_PRODUCTVERSION)) Then 
                                If dicTmp.Exists(sFamily) Then
                                    If (sSeq > dicTmp.Item(sFamily)) Then dicTmp.Item(sFamily) = sSeq
                                Else
                                    dicTmp.Add sFamily, sSeq
                                End If
                            End If
                        Next 'MspSeq
                    End If 'Not iPosMaster = -1
                Next 'iChainProd
                
                For Each key in dicTmp.Keys
                    sTmp = ""
                    If InStr (key, "_") > 0 Then
                       If Len (key) = InStr (key, "_") + 4 Then
                            sTmp = Left (key, InStr (key, "_")) & dicTmp.Item (key)
                            sLcid = Right (key, 4)
                            If NOT dicMspFamVer.Exists (sTmp) Then
                                dicMspFamVer.Add sTmp, sLcid
                            Else
                                If NOT InStr (dicMspFamVer.Item (sTmp), sLcid) > 0 Then dicMspFamVer.Item (sTmp) = dicMspFamVer.Item (sTmp) & "," & sLcid
                            End If
                        End If
                    Else
                        If NOT dicMspFamVer.Exists (key) Then dicMspFamVer.Add key, dicTmp.Item (key)
                    End If
                Next

                For Each key in dicMspFamVer.Keys
                    If InStr (key, "_") > 0 Then
                        arrTmp = Split (key, "_")
                        CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, sCategory, arrTmp (1) & " - " & arrTmp (0) & " (" & dicMspFamVer.Item (key) & ")"
                    Else
                        CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, sCategory, dicTmp.Item(Key) & " - " & key
                    End If
                    sCategory = ""
                Next 'key
            
                sCategory = "Patchlist by product"
                For iChainProd = COL_LBOUNDCHAINLIST To UBound(arrArpProducts,2)
                    If IsEmpty(arrArpProducts(iArpCnt,iChainProd)) Then Exit For
                    iPosMaster = GetArrayPosition(arrMaster,arrArpProducts(iArpCnt,iChainProd))
                    If iPosMaster = -1  AND NOT dicMissingChild.Exists(arrArpProducts(iArpCnt,iChainProd)) Then Cachelog LOGPOS_REVITEM,LOGHEADING_NONE,ERR_CATEGORYERROR,ERR_DATAINTEGRITY
                    'Only run if iPosMaster has a valid index #
                    If Not iPosMaster = -1 Then
                        For iPosPatch = 0 to UBound(arrPatch,3)
                            If Not IsEmpty (arrPatch(iPosMaster,PATCH_PATCHCODE,iPosPatch)) Then
                                If iPosPatch = 0 Then CacheLog LOGPOS_PRODUCT,LOGHEADING_NONE,"",""
                                If iPosPatch = 0 Then CacheLog LOGPOS_PRODUCT,LOGHEADING_NONE,sCategory,arrMaster(iPosMaster,COL_PRODUCTNAME)&" - "&arrMaster(iPosMaster,COL_PRODUCTCODE)
                                sCategory = ""
                                sText = "  "
                                For iColPatch = PATCH_LOGSTART to PATCH_LOGCHAINEDMAX
                                    If Not IsEmpty(arrPatch(iPosMaster,iColPatch,iPosPatch)) Then sText = sText & arrLogFormat(ARRAY_PATCH,iColPatch) & arrPatch(iPosMaster,iColPatch,iPosPatch) & CSV
                                Next 'iColPatch
                                CacheLog LOGPOS_PRODUCT,LOGHEADING_NONE,sCategory,RTrimComma(sText)
                            End If 'IsEmpty
                        Next 'iPosPatch
                    End If ' Not iPosMaster = -1
                Next 'iChainProd
            
                'InstallSource
                '-------------
                CacheLog LOGPOS_PRODUCT,LOGHEADING_NONE,"",""
                For iColISource = IS_LOG_LBOUND To IS_LOG_UBOUND
                    If NOT iPos = -1 Then
                        If Not IsEmpty(arrIS(iPos,iColISource)) Then _ 
                        CacheLog LOGPOS_PRODUCT,LOGHEADING_NONE,arrLogFormat(ARRAY_IS,iColISource),arrIS(iPos,iColISource)
                    End If
                Next 'iColISource
            End If 'fBasicMode

        Next 'iDummy
    Next 'iArpCnt
    If fLoggedMulti Then CacheLog LOGPOS_PRODUCT,LOGHEADING_NONE,"",vbCrLf & String(160,"*") & vbCrLf
    End If 'CheckArray(arrArpProducts)

    'Prepare Other Products
    '======================
    Err.Clear
    For iLogCnt = 0 To 12
        For iPosMaster = 0 To UBound(arrMaster)
            fLogProduct = CheckLogProduct(iLogCnt, iPosMaster)
            If fLogProduct Then
                If NOT fLoggedSingle Then CacheLog LOGPOS_PRODUCT, LOGHEADING_H1, Null, "Single .msi Products View" 
                fLoggedSingle = True
                For iDummy = 1 To 1
                'arrMaster contents
                '------------------
                sText = "" : sText = arrMaster(iPosMaster,COL_ARPPRODUCTNAME)
                If sText = "" Then sText = arrMaster(iPosMaster,COL_PRODUCTNAME)
                CacheLog LOGPOS_PRODUCT,LOGHEADING_H2,Null,sText
                If arrMaster(iPosMaster,COL_STATE) = INSTALLSTATE_UNKNOWN Then
                    sText = arrMaster(iPosMaster,COL_PRODUCTCODE)
                    CacheLog LOGPOS_PRODUCT,LOGHEADING_NONE,arrLogFormat(ARRAY_MASTER,COL_PRODUCTCODE),sText
                    sText = arrMaster(iPosMaster,COL_USERSID)
                    CacheLog LOGPOS_PRODUCT,LOGHEADING_NONE,arrLogFormat(ARRAY_MASTER,COL_USERSID),sText
                    sText = arrMaster(iPosMaster,COL_CONTEXTSTRING)
                    CacheLog LOGPOS_PRODUCT,LOGHEADING_NONE,arrLogFormat(ARRAY_MASTER,COL_CONTEXTSTRING),sText
                    sText = arrMaster(iPosMaster,COL_STATESTRING)
                    CacheLog LOGPOS_PRODUCT,LOGHEADING_NONE,arrLogFormat(ARRAY_MASTER,COL_STATESTRING),sText
                    Exit For 'Dummy
                End If 'arrMaster(iPosMster,COL_STATE) = INSTALLSTATE_UNKNOWN
                For iPosOrder = 0 To UBound(arrOrder)
                    If fBasicMode AND iPosOrder > 1 Then Exit For
                    sText = "" : sText = arrMaster(iPosMaster,arrOrder(iPosOrder))
                    If Not IsEmpty(arrMaster(iPosMaster,arrOrder(iPosOrder))) Then 
                        CacheLog LOGPOS_PRODUCT,LOGHEADING_NONE,arrLogFormat(ARRAY_MASTER,arrOrder(iPosOrder)),sText
                        If arrMaster(iPosMaster,arrOrder(iPosOrder)) = "Unknown" Then Exit For
                    End If
                Next 'iPosOrder
                
                If NOT fBasicMode Then
                    'Patches
                    '-------
                    CacheLog LOGPOS_PRODUCT,LOGHEADING_NONE, "", ""
                    'First loop will take care of client side patches
                    'Second loop will log patches in the InstallSource

                    sCategory = "Patch Baseline"
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, sCategory, arrMaster(iPosMaster, COL_PRODUCTVERSION)
                
                    sCategory = "Post Baseline Sequences"
                    dicTmp.RemoveAll
                    Set arrTmp = Nothing
                    arrTmp = Split(arrMaster(iPosMaster, COL_PATCHFAMILY), ",")
                    For Each MspSeq in arrTmp
                        arrTmpInner = Split(MspSeq, ":")
                        dicTmp.Add arrTmpInner(0), arrTmpInner(1)
                    Next 'MspSeq
                    For Each Key in dicTmp.Keys
                        CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, sCategory, dicTmp.Item(Key) & " - " & Key
                        sCategory = ""
                    Next
                
                    sCategory = "Client side patches"  
                    bCspCondition = True
                    For iAip = 0 To 1
                        For iPosPatch = 0 to UBound(arrPatch,3)
                            If Not (IsEmpty (arrPatch(iPosMaster,PATCH_PATCHCODE,iPosPatch))) AND (arrPatch(iPosMaster,PATCH_CSP,iPosPatch) = bCspCondition) Then
                                sText = ""
                                For iColPatch = PATCH_LOGSTART to PATCH_LOGMAX
                                    If Not IsEmpty(arrPatch(iPosMaster,iColPatch,iPosPatch)) Then sText = sText & arrLogFormat(ARRAY_PATCH,iColPatch) & arrPatch(iPosMaster,iColPatch,iPosPatch) & CSV
                                Next 'iColPatch
                                CacheLog LOGPOS_PRODUCT,LOGHEADING_NONE,sCategory,sText
                                sCategory = ""
                            End If 'IsEmpty
                        Next 'iPosPatch
                        sCategory = "Patches in InstallSource"
                        bCspCondition = False
                    Next 'iAip

                    'InstallSource
                    '-------------
                    CacheLog LOGPOS_PRODUCT,LOGHEADING_NONE,"",""
                    For iColISource = IS_LOG_LBOUND To IS_LOG_UBOUND
                        If Not IsEmpty(arrIS(iPosMaster,iColISource)) Then _ 
                        CacheLog LOGPOS_PRODUCT,LOGHEADING_NONE,arrLogFormat(ARRAY_IS,iColISource),arrIS(iPosMaster,iColISource)
                    Next 'iColISource
                
                    'FeatureStates
                    '-------------
                    CacheLog LOGPOS_PRODUCT,LOGHEADING_NONE,"",""
                    If fFeatureTree Then
                        CacheLog LOGPOS_PRODUCT,LOGHEADING_NONE,"FeatureStates",arrFeature(iPosMaster,FEATURE_TREE)
                    End If 'fFeatureTree
                End If 'fBasicMode
            Next 'iDummy
            End If 'fLogProduct
        Next 'iPosMaster
    Next 'iLogCnt
    If fLoggedSingle Then CacheLog LOGPOS_PRODUCT,LOGHEADING_NONE,"",vbCrLf & String(160,"*") & vbCrLf


End Sub 'PrepareLog
'=======================================================================================================

Sub Cachelog (iArrLogPosition,iHeading,sCategory,sText)
    On Error Resume Next
    
    If Not iHeading = 0 Then sText = FormatHeading(iHeading,sText) 
    If Not IsNull(sCategory) Then sCategory = FormatCategory(sCategory) 
    If iArrLogPosition = LOGPOS_REVITEM Then 
        If Not InStr(arrLog(iArrLogPosition),sText)>0 Then _
         arrLog(iArrLogPosition) = arrLog(iArrLogPosition) & sCategory & sText &vbCrLf 
    Else
        arrLog(iArrLogPosition) = arrLog(iArrLogPosition) & sCategory & sText &vbCrLf 
    End If
End Sub
'=======================================================================================================

Function CheckLogProduct (iLogCnt, iPosMaster)
    Dim fLogProduct

    Select Case iLogCnt
            
    Case 0 'Add-Ins
        fLogProduct = False
        fLogProduct = fLogProduct OR (InStr(POWERPIVOT_2010, arrMaster(iPosMaster, COL_UPGRADECODE)) > 0 AND Len(arrMaster(iPosMaster, COL_UPGRADECODE)) = 38)
    Case 1 'Office 16
        fLogProduct = fLogChainedDetails AND arrMaster(iPosMaster, COL_ISOFFICEPRODUCT) AND GetVersionMajor(arrMaster(iPosMaster, COL_PRODUCTCODE)) = 16
    Case 2 'Office 16 Single Msi Products
        fLogProduct = NOT fLogChainedDetails AND _ 
                        arrMaster(iPosMaster, COL_ISOFFICEPRODUCT) AND _ 
                        ( (NOT arrMaster(iPosMaster, COL_SYSTEMCOMPONENT) = 1) OR (NOT arrMaster(iPosMaster, COL_SYSTEMCOMPONENT) = 0 AND arrMaster(iPosMaster, COL_ARPPARENTCOUNT) = 0) ) AND _ 
                        IsEmpty(arrMaster(iPosMaster, COL_ARPPARENTCOUNT)) AND _ 
                        GetVersionMajor(arrMaster(iPosMaster, COL_PRODUCTCODE)) = 16
        If arrMaster(iPosMaster, COL_INSTALLTYPE) = "C2R" AND NOT fLogVerbose Then fLogProduct = False
    Case 3 'Office 15 
        fLogProduct = fLogChainedDetails AND arrMaster(iPosMaster, COL_ISOFFICEPRODUCT) AND GetVersionMajor(arrMaster(iPosMaster, COL_PRODUCTCODE)) = 15
    Case 4 'Office 15 Single Msi Products
        fLogProduct = NOT fLogChainedDetails AND _ 
                        arrMaster(iPosMaster, COL_ISOFFICEPRODUCT) AND _ 
                        ( (NOT arrMaster(iPosMaster, COL_SYSTEMCOMPONENT) = 1) OR (NOT arrMaster(iPosMaster, COL_SYSTEMCOMPONENT) = 0 AND arrMaster(iPosMaster, COL_ARPPARENTCOUNT) = 0) ) AND _ 
                        IsEmpty(arrMaster(iPosMaster, COL_ARPPARENTCOUNT)) AND _ 
                        GetVersionMajor(arrMaster(iPosMaster, COL_PRODUCTCODE)) = 15
        If arrMaster(iPosMaster, COL_INSTALLTYPE) = "C2R" AND NOT fLogVerbose Then fLogProduct = False
    Case 5 'Office 14
        fLogProduct = fLogChainedDetails AND arrMaster(iPosMaster,COL_ISOFFICEPRODUCT) AND GetVersionMajor(arrMaster(iPosMaster,COL_PRODUCTCODE))=14
    Case 6 'Office 14 Single Msi Products
        fLogProduct = NOT fLogChainedDetails AND _ 
                        arrMaster(iPosMaster,COL_ISOFFICEPRODUCT) AND _ 
                        ( (NOT arrMaster(iPosMaster,COL_SYSTEMCOMPONENT)=1) OR (NOT arrMaster(iPosMaster,COL_SYSTEMCOMPONENT)=0 AND arrMaster(iPosMaster,COL_ARPPARENTCOUNT)=0) ) AND _ 
                        IsEmpty(arrMaster(iPosMaster,COL_ARPPARENTCOUNT)) AND _ 
                        GetVersionMajor(arrMaster(iPosMaster,COL_PRODUCTCODE))=14
    Case 7 'Office 12
        fLogProduct = fLogChainedDetails AND arrMaster(iPosMaster,COL_ISOFFICEPRODUCT) AND GetVersionMajor(arrMaster(iPosMaster,COL_PRODUCTCODE))=12
    Case 8 'Office 12 Single Msi Products
        fLogProduct = NOT fLogChainedDetails AND _ 
                        arrMaster(iPosMaster,COL_ISOFFICEPRODUCT) AND _ 
                        ( (NOT arrMaster(iPosMaster,COL_SYSTEMCOMPONENT)=1) OR (NOT arrMaster(iPosMaster,COL_SYSTEMCOMPONENT)=0 AND arrMaster(iPosMaster,COL_ARPPARENTCOUNT)=0) ) AND _ 
                        IsEmpty(arrMaster(iPosMaster,COL_ARPPARENTCOUNT)) AND _ 
                        GetVersionMajor(arrMaster(iPosMaster,COL_PRODUCTCODE))=12
    Case 9 'Office 11
        fLogProduct = arrMaster(iPosMaster,COL_ISOFFICEPRODUCT) AND GetVersionMajor(arrMaster(iPosMaster,COL_PRODUCTCODE))=11
    Case 10 'Office 10
        fLogProduct = arrMaster(iPosMaster,COL_ISOFFICEPRODUCT) AND GetVersionMajor(arrMaster(iPosMaster,COL_PRODUCTCODE))=10
    Case 11 'Office  9
        fLogProduct = arrMaster(iPosMaster,COL_ISOFFICEPRODUCT) AND GetVersionMajor(arrMaster(iPosMaster,COL_PRODUCTCODE))=9
    Case 12 'Non Office Products
        fLogProduct = fListNonOfficeProducts AND NOT arrMaster(iPosMaster, COL_ISOFFICEPRODUCT)
    Case Else
    End Select

    CheckLogProduct = fLogProduct
End Function 'CheckLogProduct

'=======================================================================================================

Function FormatCategory(sCategory)
    Dim sTmp : Dim i
    On Error Resume Next

    Const iCATLEN = 33
    If Len(sCategory) > iCATLEN - 1 Then sTmp = sTmp & vbTab Else _
        sTmp = sTmp & Space(iCATLEN - Len(sCategory) - 1)

    FormatCategory = sCategory & sTmp 
End Function
'=======================================================================================================

Function FormatHeading(iHeading,sText)
    Dim sTmp, sStyle 
    Dim i
    On Error Resume Next

    Select Case iHeading
    Case LOGHEADING_H1: sStyle = "="
    Case LOGHEADING_H2: sStyle = "-"
    Case Else: sStyle =" "
    End Select
    
    sTmp = sTmp & String(Len(sText),sStyle)
    FormatHeading = vbCrLf & vbCrLf & sText & vbCrlf & sTmp 

End Function

'=======================================================================================================
'Module Global Helper Functions
'=======================================================================================================

'=======================================================================================================

Sub RelaunchElevated
    Dim Argument
    Dim sCmdLine
    Dim oShell,oWShell

    Set oShell = CreateObject("Shell.Application")
    Set oWShell = CreateObject("Wscript.Shell")

    sCmdLine = Chr(34) & WScript.scriptFullName & Chr(34)
    If Wscript.Arguments.Count > 0 Then
        For Each Argument in Wscript.Arguments
            If Argument = "UAC" Then Exit Sub
            sCmdLine = sCmdLine  &  " " & chr(34) & Argument & chr(34)
        Next 'Argument
    End If
    oShell.ShellExecute oWShell.ExpandEnvironmentStrings("%windir%") & "\system32\cscript.exe",sCmdLine & " UAC", "", "runas", 1
    Wscript.Quit
End Sub 'RelaunchElevated
'=======================================================================================================

Sub RelaunchAsCScript
    Dim Argument
    Dim sCmdLine
    
    sCmdLine = oShell.ExpandEnvironmentStrings("%windir%") & "\system32\cscript.exe " & Chr(34) & WScript.scriptFullName & Chr(34)
    If Wscript.Arguments.Count > 0 Then
        For Each Argument in Wscript.Arguments
            sCmdLine = sCmdLine  &  " " & chr(34) & Argument & chr(34)
        Next 'Argument
    End If
    oShell.Run sCmdLine,1,False
    Wscript.Quit
End Sub 'RelaunchAsCScript
'=======================================================================================================

'Launch a Shell command, wait for the task to complete and return the result
Function Command(sCommand)
    Dim oExec
    Dim sCmdOut

    Set oExec = oShell.Exec(sCommand)
    sCmdOut = oExec.StdOut.ReadAll()
    Do While oExec.Status = 0
         WScript.Sleep 100
    Loop
    Command = oExec.ExitCode & " - " & sCmdOut
End Function
'=======================================================================================================

'-------------------------------------------------------------------------------
'   GetFullPathFromRelative
'
'   Expands a relative path syntax to the full path
'-------------------------------------------------------------------------------
Function GetFullPathFromRelative (sRelativePath)
    Dim sScriptDir

    sScriptDir = Left (wscript.ScriptFullName, InStrRev (wscript.ScriptFullName, "\"))
    ' ensure sRelativePath has no leading "\"
    If Left (sRelativePath, 1) = "\" Then sRelativePath = Mid (sRelativePath, 2)
    GetFullPathFromRelative = oFso.GetAbsolutePathName (sScriptDir & sRelativePath)

End Function 'GetFullPathFromRelative

'=======================================================================================================

Sub CopyToZip (ShellNameSpace, fsoFile)
    Dim fCopyComplete
    Dim item,ShellFolder
    Dim i

    Set ShellFolder = ShellNameSpace

    fCopyComplete = False
    ShellNameSpace.CopyHere fsoFile.Path,COPY_OVERWRITE
    For Each item in ShellFolder.Items
        If item.Name = fsoFile.Name Then fCopyComplete = True
    Next 'item
    i = 0
    While NOT fCopyComplete
        WScript.Sleep 500
        i = i + 1
        For Each item in ShellFolder.Items
            If item.Name = fsoFile.Name Then fCopyComplete = True
        Next 'item
    ' hang protection
        If i > 12 Then
            fCopyComplete = True
            fZipError = True
        End If
    Wend
End Sub
'=======================================================================================================

'Function to compare to numbers of unspecified format
Function CompareVersion (sFile1, sFile2, bAllowBlanks)
'Return values:
'Left file version is lower than right file version     -1
'Left file version is identical to right file version    0
'Left file version is higher than right file version     1
'Invalid comparison                                      2

    Dim file1, file2
    Dim sDelimiter
    Dim iCnt, iAsc, iMax, iF1, iF2
    Dim bLEmpty, bREmpty

    CompareVersion = 0
    bLEmpty = False
    bREmpty = False
    
    'Ensure valid inputs values
    On Error Resume Next
    If IsEmpty(sFile1) Then bLEmpty = True
    If IsEmpty(sFile2) Then bREmpty = True
    If sFile1 = "" Then bLEmpty = True
    If sFile2 = "" Then bREmpty = True
' don't allow alpha characters
    If Not bLEmpty Then
        For iCnt = 1 To Len(sFile1)
            iAsc = Asc(UCase(Mid(sFile1,iCnt,1)))
            If (iAsc>64) AND (iAsc<91) Then
                CompareVersion = 2
                Exit Function
            End If
        Next 'iCnt
    End If
    If Not bREmpty Then
        For iCnt = 1 To Len(sFile2)
            iAsc = Asc(UCase(Mid(sFile2,iCnt,1)))
            If (iAsc>64) AND (iAsc<91) Then
                CompareVersion = 2
                Exit Function
            End If
        Next 'iCnt
    End If
    
    If bLEmpty AND (NOT bREmpty) Then
        If bAllowBlanks Then CompareVersion = -1 Else CompareVersion = 2
        Exit Function
    End If
    
    If (NOT bLEmpty) AND bREmpty Then
        If bAllowBlanks Then CompareVersion = 1 Else CompareVersion = 2
        Exit Function
    End If
    
    If bLEmpty AND bREmpty Then
        CompareVersion = 2
        Exit Function
    End If
    
' if Files are identical we're already done
    If sFile1 = sFile2 Then Exit Function
' split the VersionString
    file1 = Split(sFile1,Delimiter(sFile1))
    file2 = Split(sFile2,Delimiter(sFile2))
' ensure we get the lower count
    iMax = UBound(file1)
    CompareVersion = -1
    If iMax > UBound(file2) Then 
        iMax = UBound(file2)
        CompareVersion = 1
    End If
' compare the file versions
    For iCnt = 0 To iMax
        iF1 = CLng(file1(iCnt))
        iF2 = CLng(file2(iCnt))
        If iF1 > iF2 Then
            CompareVersion = 1
            Exit For
        ElseIf iF1 < iF2 Then
            CompareVersion = -1
            Exit For
        End If
    Next 'iCnt
End Function

'=======================================================================================================

Function hAtS(arrHex)
    Dim c,s

    On Error Resume Next
    hAtS = "" : s = ""
    If NOT IsArray(arrHex) Then Exit Function
    For Each c in arrHex
        s = s & Chr(CInt("&h" & c))
    Next 'c
    hAts = s
End Function
'=======================================================================================================

Function Delimiter (sVersion)
    Dim iCnt, iAsc

    Delimiter = " "
    For iCnt = 1 To Len(sVersion)
        iAsc = Asc(Mid(sVersion, iCnt, 1))
        If Not (iASC >= 48 And iASC <= 57) Then 
            Delimiter = Mid(sVersion, iCnt, 1)
            Exit Function
        End If
    Next 'iCnt
End Function
'=======================================================================================================

'Get the culture info tag from LCID
Function GetCultureInfo (sLcid)
    Dim sLang

    Select Case UCase(Hex(CInt(sLcid)))
        Case "7F" : sLang = ""        'Invariant culture
        Case "36" : sLang = "af"	 ' Afrikaans
        Case "436" : sLang = "af-ZA"	 ' Afrikaans (South Africa)
        Case "1C" : sLang = "sq"	 ' Albanian
        Case "41C" : sLang = "sq-AL"	 ' Albanian (Albania)
        Case "1" : sLang = "ar"	 ' Arabic
        Case "1401" : sLang = "ar-DZ"	 ' Arabic (Algeria)
        Case "3C01" : sLang = "ar-BH"	 ' Arabic (Bahrain)
        Case "C01" : sLang = "ar-EG"	 ' Arabic (Egypt)
        Case "801" : sLang = "ar-IQ"	 ' Arabic (Iraq)
        Case "2C01" : sLang = "ar-JO"	 ' Arabic (Jordan)
        Case "3401" : sLang = "ar-KW"	 ' Arabic (Kuwait)
        Case "3001" : sLang = "ar-LB"	 ' Arabic (Lebanon)
        Case "1001" : sLang = "ar-LY"	 ' Arabic (Libya)
        Case "1801" : sLang = "ar-MA"	 ' Arabic (Morocco)
        Case "2001" : sLang = "ar-OM"	 ' Arabic (Oman)
        Case "4001" : sLang = "ar-QA"	 ' Arabic (Qatar)
        Case "401" : sLang = "ar-SA"	 ' Arabic (Saudi Arabia)
        Case "2801" : sLang = "ar-SY"	 ' Arabic (Syria)
        Case "1C01" : sLang = "ar-TN"	 ' Arabic (Tunisia)
        Case "3801" : sLang = "ar-AE"	 ' Arabic (U.A.E.)
        Case "2401" : sLang = "ar-YE"	 ' Arabic (Yemen)
        Case "2B" : sLang = "hy"	 ' Armenian
        Case "42B" : sLang = "hy-AM"	 ' Armenian (Armenia)
        Case "2C" : sLang = "az"	 ' Azeri
        Case "82C" : sLang = "az-Cyrl-AZ"	 ' Azeri (Azerbaijan, Cyrillic)
        Case "42C" : sLang = "az-Latn-AZ"	 ' Azeri (Azerbaijan, Latin)
        Case "2D" : sLang = "eu"	 ' Basque
        Case "42D" : sLang = "eu-ES"	 ' Basque (Basque)
        Case "23" : sLang = "be"	 ' Belarusian
        Case "423" : sLang = "be-BY"	 ' Belarusian (Belarus)
        Case "2" : sLang = "bg"	 ' Bulgarian
        Case "402" : sLang = "bg-BG"	 ' Bulgarian (Bulgaria)
        Case "3" : sLang = "ca"	 ' Catalan
        Case "403" : sLang = "ca-ES"	 ' Catalan (Catalan)
        Case "C04" : sLang = "zh-HK"	 ' Chinese (Hong Kong SAR, PRC)
        Case "1404" : sLang = "zh-MO"	 ' Chinese (Macao SAR)
        Case "804" : sLang = "zh-CN"	 ' Chinese (PRC)
        Case "4" : sLang = "zh-Hans"	 ' Chinese (Simplified)
        Case "1004" : sLang = "zh-SG"	 ' Chinese (Singapore)
        Case "404" : sLang = "zh-TW"	 ' Chinese (Taiwan)
        Case "7C04" : sLang = "zh-Hant"	 ' Chinese (Traditional)
        Case "1A" : sLang = "hr"	 ' Croatian
        Case "41A" : sLang = "hr-HR"	 ' Croatian (Croatia)
        Case "5" : sLang = "cs"	 ' Czech
        Case "405" : sLang = "cs-CZ"	 ' Czech (Czech Republic)
        Case "6" : sLang = "da"	 ' Danish
        Case "406" : sLang = "da-DK"	 ' Danish (Denmark)
        Case "65" : sLang = "dv"	 ' Divehi
        Case "465" : sLang = "dv-MV"	 ' Divehi (Maldives)
        Case "13" : sLang = "nl"	 ' Dutch
        Case "813" : sLang = "nl-BE"	 ' Dutch (Belgium)
        Case "413" : sLang = "nl-NL"	 ' Dutch (Netherlands)
        Case "9" : sLang = "en"	 ' English
        Case "C09" : sLang = "en-AU"	 ' English (Australia)
        Case "2809" : sLang = "en-BZ"	 ' English (Belize)
        Case "1009" : sLang = "en-CA"	 ' English (Canada)
        Case "2409" : sLang = "en-029"	 ' English (Caribbean)
        Case "1809" : sLang = "en-IE"	 ' English (Ireland)
        Case "2009" : sLang = "en-JM"	 ' English (Jamaica)
        Case "1409" : sLang = "en-NZ"	 ' English (New Zealand)
        Case "3409" : sLang = "en-PH"	 ' English (Philippines)
        Case "1C09" : sLang = "en-ZA"	 ' English (South Africa
        Case "2C09" : sLang = "en-TT"	 ' English (Trinidad and Tobago)
        Case "809" : sLang = "en-GB"	 ' English (United Kingdom)
        Case "409" : sLang = "en-US"	 ' English (United States)
        Case "3009" : sLang = "en-ZW"	 ' English (Zimbabwe)
        Case "25" : sLang = "et"	 ' Estonian
        Case "425" : sLang = "et-EE"	 ' Estonian (Estonia)
        Case "38" : sLang = "fo"	 ' Faroese
        Case "438" : sLang = "fo-FO"	 ' Faroese (Faroe Islands)
        Case "29" : sLang = "fa"	 ' Farsi
        Case "429" : sLang = "fa-IR"	 ' Farsi (Iran)
        Case "B" : sLang = "fi"	 ' Finnish
        Case "40B" : sLang = "fi-FI"	 ' Finnish (Finland)
        Case "C" : sLang = "fr"	 ' French
        Case "80C" : sLang = "fr-BE"	 ' French (Belgium)
        Case "C0C" : sLang = "fr-CA"	 ' French (Canada)
        Case "40C" : sLang = "fr-FR"	 ' French (France)
        Case "140C" : sLang = "fr-LU"	 ' French (Luxembourg)
        Case "180C" : sLang = "fr-MC"	 ' French (Monaco)
        Case "100C" : sLang = "fr-CH"	 ' French (Switzerland)
        Case "56" : sLang = "gl"	 ' Galician
        Case "456" : sLang = "gl-ES"	 ' Galician (Spain)
        Case "37" : sLang = "ka"	 ' Georgian
        Case "437" : sLang = "ka-GE"	 ' Georgian (Georgia)
        Case "7" : sLang = "de"	 ' German
        Case "C07" : sLang = "de-AT"	 ' German (Austria)
        Case "407" : sLang = "de-DE"	 ' German (Germany)
        Case "1407" : sLang = "de-LI"	 ' German (Liechtenstein)
        Case "1007" : sLang = "de-LU"	 ' German (Luxembourg)
        Case "807" : sLang = "de-CH"	 ' German (Switzerland)
        Case "8" : sLang = "el"	 ' Greek
        Case "408" : sLang = "el-GR"	 ' Greek (Greece)
        Case "47" : sLang = "gu"	 ' Gujarati
        Case "447" : sLang = "gu-IN"	 ' Gujarati (India)
        Case "D" : sLang = "he"	 ' Hebrew
        Case "40D" : sLang = "he-IL"	 ' Hebrew (Israel)
        Case "39" : sLang = "hi"	 ' Hindi
        Case "439" : sLang = "hi-IN"	 ' Hindi (India)
        Case "E" : sLang = "hu"	 ' Hungarian
        Case "40E" : sLang = "hu-HU"	 ' Hungarian (Hungary)
        Case "F" : sLang = "is"	 ' Icelandic
        Case "40F" : sLang = "is-IS"	 ' Icelandic (Iceland)
        Case "21" : sLang = "id"	 ' Indonesian
        Case "421" : sLang = "id-ID"	 ' Indonesian (Indonesia)
        Case "10" : sLang = "it"	 ' Italian
        Case "410" : sLang = "it-IT"	 ' Italian (Italy)
        Case "810" : sLang = "it-CH"	 ' Italian (Switzerland)
        Case "11" : sLang = "ja"	 ' Japanese
        Case "411" : sLang = "ja-JP"	 ' Japanese (Japan)
        Case "4B" : sLang = "kn"	 ' Kannada
        Case "44B" : sLang = "kn-IN"	 ' Kannada (India)
        Case "3F" : sLang = "kk"	 ' Kazakh
        Case "43F" : sLang = "kk-KZ"	 ' Kazakh (Kazakhstan)
        Case "57" : sLang = "kok"	 ' Konkani
        Case "457" : sLang = "kok-IN"	 ' Konkani (India)
        Case "12" : sLang = "ko"	 ' Korean
        Case "412" : sLang = "ko-KR"	 ' Korean (Korea)
        Case "40" : sLang = "ky"	 ' Kyrgyz
        Case "440" : sLang = "ky-KG"	 ' Kyrgyz (Kyrgyzstan)
        Case "26" : sLang = "lv"	 ' Latvian
        Case "426" : sLang = "lv-LV"	 ' Latvian (Latvia)
        Case "27" : sLang = "lt"	 ' Lithuanian
        Case "427" : sLang = "lt-LT"	 ' Lithuanian (Lithuania)
        Case "2F" : sLang = "mk"	 ' Macedonian
        Case "42F" : sLang = "mk-MK"	 ' Macedonian (Macedonia, FYROM)
        Case "3E" : sLang = "ms"	 ' Malay
        Case "83E" : sLang = "ms-BN"	 ' Malay (Brunei Darussalam)
        Case "43E" : sLang = "ms-MY"	 ' Malay (Malaysia)
        Case "4E" : sLang = "mr"	 ' Marathi
        Case "44E" : sLang = "mr-IN"	 ' Marathi (India)
        Case "50" : sLang = "mn"	 ' Mongolian
        Case "450" : sLang = "mn-MN"	 ' Mongolian (Mongolia)
        Case "14" : sLang = "no"	 ' Norwegian
        Case "414" : sLang = "nb-NO"	 ' Norwegian (Bokmål, Norway)
        Case "814" : sLang = "nn-NO"	 ' Norwegian (Nynorsk, Norway)
        Case "15" : sLang = "pl"	 ' Polish
        Case "415" : sLang = "pl-PL"	 ' Polish (Poland)
        Case "16" : sLang = "pt"	 ' Portuguese
        Case "416" : sLang = "pt-BR"	 ' Portuguese (Brazil)
        Case "816" : sLang = "pt-PT"	 ' Portuguese (Portugal)
        Case "46" : sLang = "pa"	 ' Punjabi
        Case "446" : sLang = "pa-IN"	 ' Punjabi (India)
        Case "18" : sLang = "ro"	 ' Romanian
        Case "418" : sLang = "ro-RO"	 ' Romanian (Romania)
        Case "19" : sLang = "ru"	 ' Russian
        Case "419" : sLang = "ru-RU"	 ' Russian (Russia)
        Case "4F" : sLang = "sa"	 ' Sanskrit
        Case "44F" : sLang = "sa-IN"	 ' Sanskrit (India)
        Case "C1A" : sLang = "sr-Cyrl-CS"	 ' Serbian (Serbia, Cyrillic)
        Case "81A" : sLang = "sr-Latn-CS"	 ' Serbian (Serbia, Latin)
        Case "1B" : sLang = "sk"	 ' Slovak
        Case "41B" : sLang = "sk-SK"	 ' Slovak (Slovakia)
        Case "24" : sLang = "sl"	 ' Slovenian
        Case "424" : sLang = "sl-SI"	 ' Slovenian (Slovenia)
        Case "A" : sLang = "es"	 ' Spanish
        Case "2C0A" : sLang = "es-AR"	 ' Spanish (Argentina)
        Case "400A" : sLang = "es-BO"	 ' Spanish (Bolivia)
        Case "340A" : sLang = "es-CL"	 ' Spanish (Chile)
        Case "240A" : sLang = "es-CO"	 ' Spanish (Colombia)
        Case "140A" : sLang = "es-CR"	 ' Spanish (Costa Rica)
        Case "1C0A" : sLang = "es-DO"	 ' Spanish (Dominican Republic)
        Case "300A" : sLang = "es-EC"	 ' Spanish (Ecuador)
        Case "440A" : sLang = "es-SV"	 ' Spanish (El Salvador)
        Case "100A" : sLang = "es-GT"	 ' Spanish (Guatemala)
        Case "480A" : sLang = "es-HN"	 ' Spanish (Honduras)
        Case "80A" : sLang = "es-MX"	 ' Spanish (Mexico)
        Case "4C0A" : sLang = "es-NI"	 ' Spanish (Nicaragua)
        Case "180A" : sLang = "es-PA"	 ' Spanish (Panama)
        Case "3C0A" : sLang = "es-PY"	 ' Spanish (Paraguay)
        Case "280A" : sLang = "es-PE"	 ' Spanish (Peru)
        Case "500A" : sLang = "es-PR"	 ' Spanish (Puerto Rico)
        Case "C0A" : sLang = "es-ES"	 ' Spanish (Spain)
        Case "380A" : sLang = "es-UY"	 ' Spanish (Uruguay)
        Case "200A" : sLang = "es-VE"	 ' Spanish (Venezuela)
        Case "41" : sLang = "sw"	 ' Swahili
        Case "441" : sLang = "sw-KE"	 ' Swahili (Kenya)
        Case "1D" : sLang = "sv"	 ' Swedish
        Case "81D" : sLang = "sv-FI"	 ' Swedish (Finland)
        Case "41D" : sLang = "sv-SE"	 ' Swedish (Sweden)
        Case "5A" : sLang = "syr"	 ' Syriac
        Case "45A" : sLang = "syr-SY"	 ' Syriac (Syria)
        Case "49" : sLang = "ta"	 ' Tamil
        Case "449" : sLang = "ta-IN"	 ' Tamil (India)
        Case "44" : sLang = "tt"	 ' Tatar
        Case "444" : sLang = "tt-RU"	 ' Tatar (Russia)
        Case "4A" : sLang = "te"	 ' Telugu
        Case "44A" : sLang = "te-IN"	 ' Telugu (India)
        Case "1E" : sLang = "th"	 ' Thai
        Case "41E" : sLang = "th-TH"	 ' Thai (Thailand)
        Case "1F" : sLang = "tr"	 ' Turkish
        Case "41F" : sLang = "tr-TR"	 ' Turkish (Turkey)
        Case "22" : sLang = "uk"	 ' Ukrainian
        Case "422" : sLang = "uk-UA"	 ' Ukrainian (Ukraine)
        Case "20" : sLang = "ur"	 ' Urdu
        Case "420" : sLang = "ur-PK"	 ' Urdu (Pakistan)
        Case "43" : sLang = "uz"	 ' Uzbek
        Case "843" : sLang = "uz-Cyrl-UZ"	 ' Uzbek (Uzbekistan, Cyrillic)
        Case "443" : sLang = "uz-Latn-UZ"	 ' Uzbek (Uzbekistan, Latin)
        Case "2A" : sLang = "vi"	 ' Vietnamese
        Case "42A" : sLang = "vi-VN"	 ' Vietnamese (Vietnam)
    Case Else : sLang = ""
    End Select
    GetCultureInfo = sLang
End Function
'=======================================================================================================

'Trim away trailing comma from string
Function RTrimComma (sString)
    sString = RTrim (sString)
    If Right(sString,1) = "," Then sString = Left(sString,Len(sString)-1)
    RTrimComma = sString
End Function
'=======================================================================================================

'Return the primary keys of a table by using the PrimaryKeys property of the database object
'in SQL ready syntax 
Function GetPrimaryTableKeys(MsiDb,sTable)
    Dim iKeyCnt
    Dim sPrimaryTmp
    Dim PrimaryKeys
    On Error Resume Next

    sPrimaryTmp = ""
    Set PrimaryKeys = MsiDb.PrimaryKeys(sTable)
    For iKeyCnt = 1 To PrimaryKeys.FieldCount
        sPrimaryTmp = sPrimaryTmp & "`"&PrimaryKeys.StringData(iKeyCnt)&"`, "
    Next 'iKeyCnt
    GetPrimaryTableKeys = Left(sPrimaryTmp,Len(sPrimaryTmp)-2)
End Function 'GetPrimaryTableKeys
'=======================================================================================================

'Return the Column schema definition of a table in SQL ready syntax
Function GetTableColumnDef(MsiDb,sTable)
    On Error Resume Next
    Dim sQuery,sColDefTmp
    Dim View,ColumnNames,ColumnTypes
    Dim iColCnt
    
    'Get the ColumnInfo details
    sColDefTmp = ""
    sQuery = "SELECT * FROM " & sTable
    Set View = MsiDb.OpenView(sQuery)
    View.Execute
    Set ColumnNames = View.ColumnInfo(MSICOLUMNINFONAMES)
    Set ColumnTypes = View.ColumnInfo(MSICOLUMNINFOTYPES)
    For iColCnt = 1 To ColumnNames.FieldCount
        sColDefTmp = sColDefTmp & ColDefToSql(ColumnNames.StringData(iColCnt),ColumnTypes.StringData(iColCnt)) & ", "
    Next 'iColCnt
    View.Close
    
    GetTableColumnDef = Left(sColDefTmp,Len(sColDefTmp)-2)
    
End Function 'GetTableColumnDef
'=======================================================================================================

'Translate the column definition fields into SQL syntax
Function ColDefToSql(sColName,sColType)
    On Error Resume Next
    
    Dim iLen
    Dim sRight,sLeft, sSqlTmp

    iLen = Len(sColType)
    sRight = Right(sColType,iLen-1)
    sLeft = Left(sColType,1)
    sSqlTmp = "`"&sColName&"`"
    Select Case sLeft
    Case "s","S"
        's? String, variable length (?=1-255) -> CHAR(#) or CHARACTER(#)
        's0 String, variable length -> LONGCHAR
        If sRight="0" Then sSqlTmp = sSqlTmp & " LONGCHAR" Else sSqlTmp = sSqlTmp & " CHAR("&sRight&")"
        If sLeft = "s" Then sSqlTmp = sSqlTmp & " NOT NULL"
    Case "l","L"
        'CHAR(#) LOCALIZABLE or CHARACTER(#) LOCALIZABLE
        If sRight="0" Then sSqlTmp = sSqlTmp & " LONGCHAR" Else sSqlTmp = sSqlTmp & " CHAR("&sRight&")"
        If sLeft = "l" Then sSqlTmp = sSqlTmp & " NOT NULL"
        If sRight="0" Then sSqlTmp = sSqlTmp & "  LOCALIZABLE" Else sSqlTmp = sSqlTmp & " LOCALIZABLE"
    Case "i","I"
        'i2 Short integer 
        'i4 Long integer 
        If sRight="2" Then sSqlTmp = sSqlTmp & " SHORT" Else sSqlTmp = sSqlTmp & " LONG"
        If sLeft = "i" Then sSqlTmp = sSqlTmp & " NOT NULL"
    Case "v","V"
        'v0 Binary Stream 
        sSqlTmp = sSqlTmp & " OBJECT"
        If sLeft = "v" Then sSqlTmp = sSqlTmp & " NOT NULL"
    Case "g","G"
        'g? Temporary string (?=0-255)
    Case "j","J"
        'j? Temporary integer (?=0,1,2,4)) 
    Case "o","O"
        'O0 Temporary object 
    Case Else
    End Select

    ColDefToSql = sSqlTmp

End Function 'ColDefToSql
'=======================================================================================================

'Initialize the dicKeyComponents dictionary to link the key components to the ComponentClients
Sub InitKeyComponents
    Dim CompId, CompClients, client
    Dim arrKeyComponents

    On Error Resume Next
    arrKeyComponents = Array (CID_ACC16_64,CID_ACC16_32,CID_ACC15_64,CID_ACC15_32,CID_ACC14_64,CID_ACC14_32,CID_ACC12,CID_ACC11,CID_XL16_64,CID_XL16_32,CID_XL15_64,CID_XL15_32,CID_XL14_64,CID_XL14_32,CID_XL12,CID_XL11,CID_GRV16_64,CID_GRV16_32,CID_GRV15_64,CID_GRV15_32,CID_GRV14_64,CID_GRV14_32,CID_GRV12,CID_LYN16_64,CID_LYN16_32,CID_LYN15_64,CID_LYN15_32,CID_ONE16_64,CID_ONE16_32,CID_ONE15_64,CID_ONE15_32,CID_ONE14_64,CID_ONE14_32,CID_ONE12,CID_ONE11,CID_MSO16_64,CID_MSO16_32,CID_MSO15_64,CID_MSO15_32,CID_MSO14_64,CID_MSO14_32,CID_MSO12,CID_MSO11,CID_OL16_64,CID_OL16_32,CID_OL15_64,CID_OL15_32,CID_OL14_64,CID_OL14_32,CID_OL12,CID_OL11,CID_PPT16_64,CID_PPT16_32,CID_PPT15_64,CID_PPT15_32,CID_PPT14_64,CID_PPT14_32,CID_PPT12,CID_PPT11,CID_PRJ16_64,CID_PRJ16_32,CID_PRJ15_64,CID_PRJ15_32,CID_PRJ14_64,CID_PRJ14_32,CID_PRJ12,CID_PRJ11,CID_PUB16_64,CID_PUB16_32,CID_PUB15_64,CID_PUB15_32,CID_PUB14_64,CID_PUB14_32,CID_PUB12,CID_PUB11,CID_IP16_64,CID_IP15_64,CID_IP15_32,CID_IP14_64,CID_IP14_32,CID_IP12,CID_IP11,CID_VIS16_64,CID_VIS16_32,CID_VIS15_64,CID_VIS15_32,CID_VIS14_64,CID_VIS14_32,CID_VIS12,CID_VIS11,CID_WD16_64,CID_WD16_32,CID_WD15_64,CID_WD15_32,CID_WD14_64,CID_WD14_32,CID_WD12,CID_WD11,CID_SPD16_64,CID_SPD16_32,CID_SPD15_64,CID_SPD15_32,CID_SPD14_64,CID_SPD14_32,CID_SPD12,CID_SPD11)
    On Error Resume Next
    For Each CompId in arrKeyComponents
        Set CompClients = oMsi.ComponentClients (CompId)
        For Each client in CompClients
            If NOT CompClients.Count > 0 Then Exit For
            If NOT client = "" Then
                If NOT dicKeyComponents.Exists (CompId) Then
                    dicKeyComponents.Add CompId, client
                Else
                    dicKeyComponents.Item (CompId) = dicKeyComponents.Item (CompId) & ";" & client
                End If
            End If
        Next
    Next
End Sub 'InitKeyComponents

'=======================================================================================================

'Checks if a productcode belongs to a Office family
Function IsOfficeProduct (sProductCode)
    On Error Resume Next
    
    IsOfficeProduct = False
    If InStr(OFFICE_ALL, UCase(Right(sProductCode, 28))) > 0 OR _
       InStr(sProductCodes_C2R, UCase(sProductCode)) > 0 OR _
       InStr(OFFICEID, UCase(Right(sProductCode, 17))) > 0 OR _
       sProductCode = sPackageGuid Then
           IsOfficeProduct = True
    End If

End Function
'=======================================================================================================

Function GetExpandedGuid (sGuid)
    Dim sExpandGuid
    Dim i
    On Error Resume Next

    sExpandGuid = "{" & StrReverse(Mid(sGuid,1,8)) & "-" & _
                        StrReverse(Mid(sGuid,9,4)) & "-" & _
                        StrReverse(Mid(sGuid,13,4))& "-"
    For i = 17 To 20
	    If i Mod 2 Then
		    sExpandGuid = sExpandGuid & mid(sGuid,(i + 1),1)
	    Else
		    sExpandGuid = sExpandGuid & mid(sGuid,(i - 1),1)
	    End If
    Next
    sExpandGuid = sExpandGuid & "-"
    For i = 21 To 32
	    If i Mod 2 Then
		    sExpandGuid = sExpandGuid & mid(sGuid,(i + 1),1)
	    Else
		    sExpandGuid = sExpandGuid & mid(sGuid,(i - 1),1)
	    End If
    Next
    sExpandGuid = sExpandGuid & "}"
    GetExpandedGuid = sExpandGuid
    
End Function
'=======================================================================================================

Function GetCompressedGuid (sGuid)
'Converts the GUID / ProductCode into the compressed format
    Dim sCompGUID
    Dim i
    On Error Resume Next

    sCompGUID = StrReverse(Mid(sGuid,2,8))  & _
                StrReverse(Mid(sGuid,11,4)) & _
                StrReverse(Mid(sGuid,16,4)) 
    For i = 21 To 24
	    If i Mod 2 Then
		    sCompGUID = sCompGUID & Mid(sGuid, (i + 1), 1)
	    Else
		    sCompGUID = sCompGUID & Mid(sGuid, (i - 1), 1)
	    End If
    Next
    For i = 26 To 37
	    If i Mod 2 Then
		    sCompGUID = sCompGUID & Mid(sGuid, (i - 1), 1)
	    Else
		    sCompGUID = sCompGUID & Mid(sGuid, (i + 1), 1)
	    End If
    Next
    GetCompressedGuid = sCompGUID
End Function
'=======================================================================================================

'Get Version Major from GUID
Function GetVersionMajor(sProductCode)
    Dim iVersionMajor
    On Error Resume Next

    iVersionMajor = 0
    If InStr(OFFICE_2000, UCase(Right(sProductCode,28))) > 0 Then iVersionMajor = 9
    If InStr(ORK_2000,    UCase(Right(sProductCode,28))) > 0 Then iVersionMajor = 9
    If InStr(PRJ_2000,    UCase(Right(sProductCode,28))) > 0 Then iVersionMajor = 9
    If InStr(VIS_2002,    UCase(Right(sProductCode,28))) > 0 Then iVersionMajor = 10
    If InStr(OFFICE_2002, UCase(Right(sProductCode,28))) > 0 Then iVersionMajor = 10
    If InStr(OFFICE_2003, UCase(Right(sProductCode,28))) > 0 Then iVersionMajor = 11
    If InStr(WSS_2,       UCase(Right(sProductCode,28))) > 0 Then iVersionMajor = 11
    If InStr(SPS_2003,    UCase(Right(sProductCode,28))) > 0 Then iVersionMajor = 11
    If InStr(PPS_2007,    UCase(Right(sProductCode,28))) > 0 Then iVersionMajor = 12
    If InStr(OFFICEID,    UCase(Right(sProductCode,17))) > 0 Then iVersionMajor = Mid(sProductCode,4,2)
    
    If iVersionMajor = 0 Then iVersionMajor = oMsi.ProductInfo(sProductCode, "VersionMajor")  

    GetVersionMajor = iVersionMajor
End Function
'=======================================================================================================

'Obtain the ProductVersion from a .msi package
Function GetMsiProductVersion(sMsiFile)
    Dim MsiDb,Record
    Dim qView
    
    On Error Resume Next
    GetMsiProductVersion = ""
    Set Record = Nothing
    Set MsiDb = oMsi.OpenDatabase(sMsiFile,MSIOPENDATABASEMODE_READONLY)
    Set qView = MsiDb.OpenView("SELECT `Value` FROM Property WHERE `Property` = 'ProductVersion'")
    qView.Execute
    Set Record = qView.Fetch
    If NOT Record Is Nothing Then GetMsiProductVersion = Record.StringData(1)
    qView.Close

End Function 'GetMsiProductVersion
'=======================================================================================================

'Get the reference name that is used to reference the product under add / remove programs
'Alternatively this could be taken from the cached .msi by reading the 'SetupExeArpId' value from the 'Property' table
Function GetArpProductname (sUninstallSubkey)
    Dim hDefKey
    Dim sSubKeyName, sName, sValue, sVer
    On Error Resume Next
    
    GetArpProductname = sUninstallSubkey
    hDefKey = HKEY_LOCAL_MACHINE
    sSubKeyName = REG_ARP & sUninstallSubkey & "\"
    sName = "DisplayName"
    
    If RegReadExpStringValue(hDefKey, sSubKeyName, sName, sValue) Then 
        If NOT IsNull(sValue) OR sValue = "" Then GetArpProductname = sValue
    Else
        'Try C2Rv2 location
        For Each sVer in dicActiveC2Rv2Versions.Keys
            sSubKeyName = REG_OFFICE & sVer & REG_C2RVIRT_HKLM & sSubKeyName
            If RegReadValue (hDefKey, sSubKeyName, sName, sValue, "REG_EXPAND_SZ") Then
                If NOT IsNull(sValue) OR sValue = "" Then GetArpProductname = sValue
                Exit For
            End If
        Next
    End If

End Function
'=======================================================================================================

'Get the original .msi name (Package Name) by direct read from registry
Function GetRegOriginalMsiName(sProductCodeCompressed,iContext,sSid)
    Dim hDefKey
    Dim sSubKeyName,sName,sValue, sFallBackName
    On Error Resume Next

    'PackageName is only available for per-machine, current user and managed user - not for other (unmanaged) user!
    
    sFallBackName = ""
    If sSid = sCurUserSid Or iContext = MSIINSTALLCONTEXT_MACHINE Or iContext = MSIINSTALLCONTEXT_C2RV2 Or iContext = MSIINSTALLCONTEXT_C2RV3 Then
        hDefKey = GetRegHive(iContext,sSid,False)
        sSubKeyName = GetRegConfigKey(sProductCodeCompressed,iContext,sSid,False) & "SourceList\"
    Else
        'PackageName not available for other (unmanaged) user
        GetRegProductName = "n/a"
        Exit Function
    End If 'sSid = sCurUserSid
        sName = "PackageName"

    If RegReadExpStringValue(hDefKey,sSubKeyName,sName,sValue) Then GetRegOriginalMsiName = sValue  & sFallBackName Else GetRegOriginalMsiName = "-"

End Function 'GetRegOriginalMsiName
'=======================================================================================================

'Get the registered transform(s) by direct read from registry metadata
Function GetRegTransforms(sProductCodeCompressed,iContext,sSid)
    Dim hDefKey
    Dim sSubKeyName,sName,sValue 
    On Error Resume Next

    'Transforms is only available for per-machine and current user - not for other user!
        
    If sSid = sCurUserSid Or iContext = MSIINSTALLCONTEXT_MACHINE Or iContext = MSIINSTALLCONTEXT_C2RV2 Or iContext = MSIINSTALLCONTEXT_C2RV3 Then
        hDefKey = GetRegHive(iContext,sSid,False)
        sSubKeyName = GetRegConfigKey(sProductCodeCompressed,iContext,sSid,False)
        sName = "Transforms"
    Else
        'Transforms not available for other user
        GetRegTransforms = "n/a"
        Exit Function
    End If 'sSid = sCurUserSid

    If RegReadExpStringValue(hDefKey,sSubKeyName,sName,sValue) Then GetRegTransforms = sValue Else GetRegTransforms = "-"

End Function 'GetRegTransforms
'=======================================================================================================

'Get the product PackageCode by direct read from registry metadata
Function GetRegPackageCode(sProductCodeCompressed,iContext,sSid)
    Dim hDefKey
    Dim sSubKeyName,sName,sValue 
    On Error Resume Next

    'PackageCode is only available for per-machine and current user - not for other user!
        
    If sSid = sCurUserSid Or iContext = MSIINSTALLCONTEXT_MACHINE Or iContext = MSIINSTALLCONTEXT_C2RV2 Or iContext = MSIINSTALLCONTEXT_C2RV3 Then
        hDefKey = GetRegHive(iContext,sSid,False)
        sSubKeyName = GetRegConfigKey(sProductCodeCompressed,iContext,sSid,False)
        sName = "PackageCode"
    Else
        'PackageCode not available for other user
        GetRegPackageCode = "n/a"
        Exit Function
    End If 'sSid = sCurUserSid

    If RegReadExpStringValue(hDefKey, sSubKeyName, sName, sValue) Then GetRegPackageCode = sValue Else GetRegPackageCode = "-"

End Function 'GetRegPackageCode
'=======================================================================================================

'Get the ProductName by direct read from registry metadata
Function GetRegProductName(sProductCodeCompressed,iContext,sSid)
    Dim hDefKey
    Dim sSubKeyName,sName,sValue,sFallBackName
    Dim i
    On Error Resume Next

    'ProductName is only available for per-machine, current user and managed user - not for other (unmanaged) user!
    'If not per-machine, managed or SID not sCurUserSid tweak to 'DisplayName' from GlobalConfig key
        
    sFallBackName = ""
    If sSid = sCurUserSid Or iContext = MSIINSTALLCONTEXT_USERMANAGED Or iContext = MSIINSTALLCONTEXT_MACHINE Or iContext = MSIINSTALLCONTEXT_C2RV2 Or iContext = MSIINSTALLCONTEXT_C2RV3 Then
        If Not iContext = MSIINSTALLCONTEXT_USERMANAGED Then
            hDefKey = GetRegHive(iContext,sSid,False)
            sSubKeyName = GetRegConfigKey(sProductCodeCompressed,iContext,sSid,False)
        Else
            hDefKey = GetRegHive(iContext,sSid,True)
            sSubKeyName = GetRegConfigKey(sProductCodeCompressed,iContext,sSid,True)
        End If
        sName = "ProductName"
    Else
        'Use GlobalConfig key to avoid conflict with per-user installs
        hDefKey = HKEY_LOCAL_MACHINE
        sSubKeyName = GetRegConfigKey(sProductCodeCompressed,iContext,sSid,True) & "InstallProperties\"
        sName = "DisplayName"
        sFallBackName = " (DisplayName)"
    End If 'sSid = sCurUserSid

    
    If RegReadExpStringValue(hDefKey,sSubKeyName,sName,sValue) Then GetRegProductName = sValue  & sFallBackName Else GetRegProductName = ERR_CATEGORYERROR & "No DislpayName registered"

End Function
'=======================================================================================================

Function GetRegProductState(sProductCodeCompressed,iContext,sSid)
    Dim hDefKey
    Dim sSubKeyName,sName,sValue
    Dim iTmpContext
    On Error Resume Next
    
    GetRegProductState = "Unknown"
    hDefKey = HKEY_LOCAL_MACHINE
    If iContext = MSIINSTALLCONTEXT_USERMANAGED Then 
        iTmpContext = MSIINSTALLCONTEXT_USERUNMANAGED
        sName = "ManagedLocalPackage"
    Else
        iTmpContext = iContext
        sName = "LocalPackage"
    End If 'iContext = MSIINSTALLCONTEXT_USERMANAGED
    sSubKeyName = GetRegConfigKey(sProductCodeCompressed, iTmpContext, sSid, True) & "InstallProperties\"
    
    If RegKeyExists (hDefKey,sSubKeyName) Then
        If RegValExists(hDefKey,sSubKeyName,sName) Then 
            GetRegProductState = 5 '"Installed"
        Else
            If InStr(sSubKeyName, REG_C2RVIRT_HKLM) > 0 Then
                GetRegProductState = 8 '"Virtualized"
            Else
                GetRegProductState = 1 '"Advertised"
            End If
        End If 'RegValExists
    Else
        GetRegProductState = -1 '"Broken/Unknown"
    End If 'RegKeyExists
    
End Function 'GetRegProductState
'=======================================================================================================

'Check if the GUID has a valid structure
Function IsValidGuid(sGuid,iGuidType)
    Dim i, n
    Dim c
    Dim bValidGuidLength, bValidGuidChar, bValidGuidCase
    On Error Resume Next

    'Set defaults
    IsValidGuid = False
    fGuidCaseWarningOnly = False
    sError = "" : sErrBpa = ""
    bValidGuidLength = True
    bValidGuidChar = True
   
    Select Case iGuidType
    Case GUID_UNCOMPRESSED 'UnCompressed
        If Len(sGuid) = 38 Then
            IsValidGuid = True
            For i = 1 To 38
                bValidGuidCase = True
                c = Mid(sGuid,i,1)
                Select Case i
                Case 1
                    If Not c = "{" Then 
                        IsValidGuid = False
                        bValidGuidChar = False
                        Exit For
                    End If
                Case 10, 15, 20, 25
                    If not c = "-" Then
                        IsValidGuid = False
                        bValidGuidChar = False
                        Exit For
                    End If
                Case 38
                    If not c = "}" Then
                        IsValidGuid = False
                        bValidGuidChar = False
                        Exit For
                    End If
                Case Else
                    n = Asc(c)
                    If Not IsValidGuidChar(n) Then
                        bValidGuidCase = False
                        fGuidCaseWarningOnly = True
                        IsValidGuid = False
                        n = Asc(UCase(c))
                        If Not IsValidGuidChar(n) Then
                            fGuidCaseWarningOnly = False
                            bValidGuidChar = False
                            Exit For
                        End If 'IsValidGuidChar (inner)
                    End If 'IsValidGuidChar (outer)
                End Select
            Next 'i
        Else
            'Invalid length for this GUID type
            'Suppress error if passed in GUID matches 'compressed' length
            If NOT (Len(sGuid)=32) Then bValidGuidLength = False
        End If 'Len(sGuid)
        
    Case  GUID_COMPRESSED
        If Len(sGuid)=32 Then
            IsValidGuid = True
            For i = 1 to 32
                c = Mid(sGuid,i,1)
                bValidGuidCase = True
                n = Asc(c)
                If Not IsValidGuidChar(n) Then 
                    bValidGuidCase = False
                    fGuidCaseWarningOnly = True
                    IsValidGuid = False
                    n = Asc(UCase(c))
                    If Not IsValidGuidChar(n) Then
                        fGuidCaseWarningOnly = False
                        bValidGuidChar = False
                         Exit For
                   End If 'IsValidGuidChar (inner)
                End If 'IsValidGuidChar (outer)
            Next 'i
        Else
            'Invalid length for this GUID type
            bValidGuidLength = False
        End If 'Len
    Case GUID_SQUISHED '"Squished"
        'Not implemented
    Case Else
        'IsValidGuid = False
    End Select
    
    'Log errors 
    If (NOT bValidGuidLength) OR (NOT bValidGuidChar) OR (fGuidCaseWarningOnly) Then 
         sError = ERR_INVALIDGUID & DOT
         sErrBpa = BPA_GUID
    End If 
    
    If fGuidCaseWarningOnly Then
        sError = sError & ERR_GUIDCASE
    End If 'fGuidCaseWarningOnly
    
    If bValidGuidLength = False Then 
        sError = sError & ERR_INVALIDGUIDLENGTH
    End If 'bValidGuidLength  
    
    If bValidGuidChar = False Then 
        sError = sError & ERR_INVALIDGUIDCHAR
    End If 'bValidGuidChar  
End Function 'IsValidGuid
'=======================================================================================================

'Check if the character is in a valid range for a GUID
Function IsValidGuidChar (iAsc)
    If ((iAsc >= 48 AND iAsc <= 57) OR (iAsc >= 65 AND iAsc <= 70)) Then
        IsValidGuidChar = True
    Else
        IsValidGuidChar = False
    End If
End Function
'=======================================================================================================

Function GetContextString(iContext)
    On Error Resume Next

    Select Case iContext
        Case MSIINSTALLCONTEXT_USERMANAGED      : GetContextString = "MSIINSTALLCONTEXT_USERMANAGED"
        Case MSIINSTALLCONTEXT_USERUNMANAGED    : GetContextString = "MSIINSTALLCONTEXT_USERUNMANAGED"
        Case MSIINSTALLCONTEXT_MACHINE          : GetContextString = "MSIINSTALLCONTEXT_MACHINE"
        Case MSIINSTALLCONTEXT_ALL              : GetContextString = "MSIINSTALLCONTEXT_ALL"
        Case MSIINSTALLCONTEXT_C2RV2            : GetContextString = "MSIINSTALLCONTEXT_C2RV2"
        Case MSIINSTALLCONTEXT_C2RV3            : GetContextString = "MSIINSTALLCONTEXT_C2RV3"
        Case Else                               : GetContextString = iContext
    End Select
End Function
'=======================================================================================================

Function GetHiveString(hDefKey)
    On Error Resume Next

    Select Case hDefKey
        Case HKEY_CLASSES_ROOT : GetHiveString = "HKEY_CLASSES_ROOT"
        Case HKEY_CURRENT_USER : GetHiveString = "HKEY_CURRENT_USER"
        Case HKEY_LOCAL_MACHINE : GetHiveString = "HKEY_LOCAL_MACHINE"
        Case HKEY_USERS : GetHiveString = "HKEY_USERS"
        Case Else : GetHiveString = hDefKey
    End Select
End Function 'GetHiveString
'=======================================================================================================

Function GetRegConfigPatchesKey(iContext, sSid, bGlobal)
    Dim sTmpProductCode, sSubKeyName, sKey, sVer
    On Error Resume Next

    sSubKeyName = ""
    Select Case iContext
    Case MSIINSTALLCONTEXT_USERMANAGED
        sSubKeyName = REG_CONTEXTUSERMANAGED & sSid & "\Installer\Patches\"
    Case MSIINSTALLCONTEXT_USERUNMANAGED
        If bGlobal OR NOT sSid = sCurUserSid Then
            sSubKeyName = REG_GLOBALCONFIG & sSid & "\Patches\"
        Else
            sSubKeyName = REG_CONTEXTUSER
        End If
    Case MSIINSTALLCONTEXT_MACHINE, MSIINSTALLCONTEXT_C2RV2, MSIINSTALLCONTEXT_C2RV3
        If bGlobal Then
            sSubKeyName = REG_GLOBALCONFIG & "S-1-5-18\Patches\"
        Else
            sSubKeyName = REG_CONTEXTMACHINE & "\Patches\"
        End If
    Case Else
    End Select
    
    GetRegConfigPatchesKey = Replace(sSubKeyName,"\\","\")
    
End Function 'GetRegConfigPatchesKey
'=======================================================================================================

Function GetRegConfigKey(sProductCode, iContext, sSid, bGlobal)
    Dim sTmpProductCode, sSubKeyName, sKey
    Dim iVM
    On Error Resume Next

    sTmpProductCode = sProductCode
    sSubKeyName = ""
    If NOT sTmpProductCode = "" Then
        If IsValidGuid(sTmpProductCode, GUID_UNCOMPRESSED) Then sTmpProductCode = GetCompressedGuid(sTmpProductCode)
    End If 'NOT sTmpProductCode = ""

    Select Case iContext
    Case MSIINSTALLCONTEXT_USERMANAGED
        sSubKeyName = REG_CONTEXTUSERMANAGED & sSid & "\Installer\Products\" & sTmpProductCode & "\"
    Case MSIINSTALLCONTEXT_USERUNMANAGED
        If bGlobal OR NOT sSid = sCurUserSid Then
            sSubKeyName = REG_GLOBALCONFIG & sSid & "\Products\" & sTmpProductCode & "\"
        Else
            sSubKeyName = REG_CONTEXTUSER & "Products\" & sTmpProductCode & "\"
        End If
    Case MSIINSTALLCONTEXT_MACHINE
        If bGlobal Then
            sSubKeyName = REG_GLOBALCONFIG & "S-1-5-18\Products\" & sTmpProductCode & "\"
        Else
            sSubKeyName = REG_CONTEXTMACHINE & "Products\" & sTmpProductCode & "\"
        End If
    Case MSIINSTALLCONTEXT_C2RV2
        sKey = REG_OFFICE & "15.0" & REG_C2RVIRT_HKLM
        If bGlobal Then
            sSubKeyName = sKey & REG_GLOBALCONFIG & "S-1-5-18\Products\" & sTmpProductCode & "\"
        Else
            sSubKeyName = sKey & "Software\Classes\" & REG_CONTEXTMACHINE & "Products\" & sTmpProductCode & "\"
        End If
    Case MSIINSTALLCONTEXT_C2RV3
        sKey = REG_OFFICE & REG_C2RVIRT_HKLM
        sKey = Replace(sKey, "\\", "\")
        If bGlobal Then
            sSubKeyName = sKey & REG_GLOBALCONFIG & "S-1-5-18\Products\" & sTmpProductCode & "\"
        Else
            sSubKeyName = sKey & "Software\Classes\" & REG_CONTEXTMACHINE & "Products\" & sTmpProductCode & "\"
        End If
    Case Else
    End Select
    
    GetRegConfigKey = Replace(sSubKeyName,"\\","\")
    
End Function 'GetRegConfigKey
'=======================================================================================================

Function GetRegHive(iContext, sSid, bGlobal)
    On Error Resume Next

    
    Select Case iContext
    Case MSIINSTALLCONTEXT_USERMANAGED 
        GetRegHive = HKEY_LOCAL_MACHINE
    Case MSIINSTALLCONTEXT_USERUNMANAGED
        If bGlobal OR NOT sSid = sCurUserSid Then 
            GetRegHive = HKEY_LOCAL_MACHINE
        Else
            GetRegHive = HKEY_CURRENT_USER
        End If
    Case MSIINSTALLCONTEXT_MACHINE
        If bGlobal Then 
            GetRegHive = HKEY_LOCAL_MACHINE
        Else
            GetRegHive = HKEY_CLASSES_ROOT
        End If
    Case MSIINSTALLCONTEXT_C2RV2, MSIINSTALLCONTEXT_C2RV3
        GetRegHive = HKEY_LOCAL_MACHINE
    Case Else
    End Select
End Function 'GetRegHive
'=======================================================================================================

Function RegKeyExists(hDefKey,sSubKeyName)
    Dim arrKeys
    RegKeyExists = (oReg.EnumKey(hDefKey,sSubKeyName,arrKeys) = 0)
End Function
'=======================================================================================================

Function RegValExists(hDefKey,sSubKeyName,sName)
    Dim arrValueTypes, arrValueNames, i
    On Error Resume Next

    RegValExists = False
    If Not RegKeyExists(hDefKey,sSubKeyName) Then
        Exit Function
    End If
    If oReg.EnumValues(hDefKey,sSubKeyName,arrValueNames,arrValueTypes) = 0 AND CheckArray(arrValueNames) Then
        For i = 0 To UBound(arrValueNames) 
            If LCase(arrValueNames(i)) = Trim(LCase(sName)) Then RegValExists = True
        Next 
    Else
        Exit Function
    End If 'oReg.EnumValues
End Function
'=======================================================================================================

'Check access to a registry key
Function RegCheckAccess(hDefKey, sSubKeyName, lAccPermLevel)
    Dim RetVal
    Dim arrValues
    
    RetVal = RegKeyExists(hDefKey,sSubKeyName)
    RetVal = oReg.CheckAccess(hDefKey,sSubKeyName,lAccPermLevel)
    If Not RetVal = 0 AND f64 Then RetVal = oReg.CheckAccess(hDefKey,Wow64Key(hDefKey, sSubKeyName),lAccPermLevel)
    RegCheckAccess = (RetVal = 0)
End Function 'RegReadValue
'=======================================================================================================

'Read the value of a given registry entry
Function RegReadValue(hDefKey, sSubKeyName, sName, sValue, sType)
    Dim RetVal
    Dim arrValues
    
    Select Case UCase(sType)
        Case "1","REG_SZ"
            RetVal = oReg.GetStringValue(hDefKey,sSubKeyName,sName,sValue)
            If Not RetVal = 0 AND f64 Then RetVal = oReg.GetStringValue(hDefKey,Wow64Key(hDefKey, sSubKeyName),sName,sValue)
        Case "2","REG_EXPAND_SZ"
            RetVal = oReg.GetExpandedStringValue(hDefKey,sSubKeyName,sName,sValue)
            If Not RetVal = 0 AND f64 Then RetVal = oReg.GetExpandedStringValue(hDefKey,Wow64Key(hDefKey, sSubKeyName),sName,sValue)
        Case "7","REG_MULTI_SZ"
            RetVal = oReg.GetMultiStringValue(hDefKey,sSubKeyName,sName,arrValues)
            If Not RetVal = 0 AND f64 Then RetVal = oReg.GetMultiStringValue(hDefKey,Wow64Key(hDefKey, sSubKeyName),sName,arrValues)
            If RetVal = 0 Then sValue = Join(arrValues,chr(34))
        Case "4","REG_DWORD"
            RetVal = oReg.GetDWORDValue(hDefKey,sSubKeyName,sName,sValue)
            If Not RetVal = 0 AND f64 Then 
                RetVal = oReg.GetDWORDValue(hDefKey,Wow64Key(hDefKey, sSubKeyName),sName,sValue)
            End If
        Case "3","REG_BINARY"
            RetVal = oReg.GetBinaryValue(hDefKey,sSubKeyName,sName,sValue)
            If Not RetVal = 0 AND f64 Then RetVal = oReg.GetBinaryValue(hDefKey,Wow64Key(hDefKey, sSubKeyName),sName,sValue)
        Case "11","REG_QWORD"
            RetVal = oReg.GetQWORDValue(hDefKey,sSubKeyName,sName,sValue)
            If Not RetVal = 0 AND f64 Then RetVal = oReg.GetQWORDValue(hDefKey,Wow64Key(hDefKey, sSubKeyName),sName,sValue)
        Case Else
            RetVal = -1
    End Select 'sValue
    RegReadValue = (RetVal = 0)
End Function 'RegReadValue
'=======================================================================================================

Function RegReadStringValue(hDefKey,sSubKeyName,sName,sValue)
    Dim RetVal

    RetVal = oReg.GetStringValue(hDefKey,sSubKeyName,sName,sValue)
    If Not RetVal = 0 AND f64 Then RetVal = oReg.GetStringValue(hDefKey,Wow64Key(hDefKey, sSubKeyName),sName,sValue)
    RegReadStringValue = (RetVal = 0)
End Function 'RegReadStringValue
'=======================================================================================================

Function RegReadExpStringValue(hDefKey,sSubKeyName,sName,sValue)
    Dim RetVal

    RetVal = oReg.GetExpandedStringValue(hDefKey,sSubKeyName,sName,sValue)
    If Not RetVal = 0 AND f64 Then RetVal = oReg.GetExpandedStringValue(hDefKey,Wow64Key(hDefKey, sSubKeyName),sName,sValue)
    RegReadExpStringValue = (RetVal = 0)
End Function 'RegReadExpStringValue
'=======================================================================================================

Function RegReadMultiStringValue(hDefKey,sSubKeyName,sName,arrValues)
    Dim RetVal

    RetVal = oReg.GetMultiStringValue(hDefKey,sSubKeyName,sName,arrValues)
    If Not RetVal = 0 AND f64 Then RetVal = oReg.GetMultiStringValue(hDefKey,Wow64Key(hDefKey, sSubKeyName),sName,arrValues)
    RegReadMultiStringValue = (RetVal = 0 AND IsArray(arrValues))
End Function 'RegReadMultiStringValue
'=======================================================================================================

Function RegReadDWordValue(hDefKey,sSubKeyName,sName,sValue)
    Dim RetVal

    RetVal = oReg.GetDWORDValue(hDefKey,sSubKeyName,sName,sValue)
    If Not RetVal = 0 AND f64 Then RetVal = oReg.GetDWORDValue(hDefKey,Wow64Key(hDefKey, sSubKeyName),sName,sValue)
    RegReadDWordValue = (RetVal = 0)
End Function 'RegReadDWordValue
'=======================================================================================================

Function RegReadBinaryValue(hDefKey,sSubKeyName,sName,sValue)
    Dim RetVal

    RetVal = oReg.GetBinaryValue(hDefKey,sSubKeyName,sName,sValue)
    If Not RetVal = 0 AND f64 Then RetVal = oReg.GetBinaryValue(hDefKey,Wow64Key(hDefKey, sSubKeyName),sName,sValue)
    RegReadBinaryValue = (RetVal = 0)
End Function 'RegReadBinaryValue
'=======================================================================================================

Function RegReadQWordValue(hDefKey,sSubKeyName,sName,sValue)
    Dim RetVal

    RetVal = oReg.GetQWORDValue(hDefKey,sSubKeyName,sName,sValue)
    If Not RetVal = 0 AND f64 Then RetVal = oReg.GetQWORDValue(hDefKey,Wow64Key(hDefKey, sSubKeyName),sName,sValue)
    RegReadQWordValue = (RetVal = 0)
End Function 'RegReadQWordValue
'=======================================================================================================

'Enumerate a registry key to return all values
Function RegEnumValues(hDefKey,sSubKeyName,arrNames, arrTypes)
    Dim RetVal, RetVal64
    Dim arrNames32, arrNames64, arrTypes32, arrTypes64
    
    If f64 Then
        RetVal = oReg.EnumValues(hDefKey,sSubKeyName,arrNames32,arrTypes32)
        RetVal64 = oReg.EnumValues(hDefKey,Wow64Key(hDefKey, sSubKeyName),arrNames64,arrTypes64)
        If (RetVal = 0) AND (Not RetVal64 = 0) AND IsArray(arrNames32) AND IsArray(arrTypes32) Then 
            arrNames = arrNames32
            arrTypes = arrTypes32
        End If
        If (Not RetVal = 0) AND (RetVal64 = 0) AND IsArray(arrNames64) AND IsArray(arrTypes64) Then 
            arrNames = arrNames64
            arrTypes = arrTypes64
        End If
        If (RetVal = 0) AND (RetVal64 = 0) AND IsArray(arrNames32) AND IsArray(arrNames64) AND IsArray(arrTypes32) AND IsArray(arrTypes64) Then 
            arrNames = RemoveDuplicates(Split((Join(arrNames32,"\") & "\" & Join(arrNames64,"\")),"\"))
            arrTypes = RemoveDuplicates(Split((Join(arrTypes32,"\") & "\" & Join(arrTypes64,"\")),"\"))
        End If
    Else
        RetVal = oReg.EnumValues(hDefKey,sSubKeyName,arrNames,arrTypes)
    End If 'f64
    RegEnumValues = ((RetVal = 0) OR (RetVal64 = 0)) AND IsArray(arrNames) AND IsArray(arrTypes)
End Function 'RegEnumValues
'=======================================================================================================

'Enumerate a registry key to return all subkeys
Function RegEnumKey(hDefKey,sSubKeyName,arrKeys)
    Dim RetVal, RetVal64
    Dim arrKeys32, arrKeys64
    
    If f64 Then
        RetVal = oReg.EnumKey(hDefKey,sSubKeyName,arrKeys32)
        RetVal64 = oReg.EnumKey(hDefKey,Wow64Key(hDefKey, sSubKeyName),arrKeys64)
        If (RetVal = 0) AND (Not RetVal64 = 0) AND IsArray(arrKeys32) Then arrKeys = arrKeys32
        If (Not RetVal = 0) AND (RetVal64 = 0) AND IsArray(arrKeys64) Then arrKeys = arrKeys64
        If (RetVal = 0) AND (RetVal64 = 0) Then 
            If IsArray(arrKeys32) AND IsArray (arrKeys64) Then 
                arrKeys = RemoveDuplicates(Split((Join(arrKeys32,"\") & "\" & Join(arrKeys64,"\")),"\"))
            ElseIf IsArray(arrKeys64) Then
                arrKeys = arrKeys64
            Else
                arrKeys = arrKeys32
            End If
        End If
    Else
        RetVal = oReg.EnumKey(hDefKey,sSubKeyName,arrKeys)
    End If 'f64
    RegEnumKey = ((RetVal = 0) OR (RetVal64 = 0)) AND IsArray(arrKeys)
End Function 'RegEnumKey
'=======================================================================================================

'Return the alternate regkey location on 64bit environment
Function Wow64Key(hDefKey, sSubKeyName)
    Dim iPos
    Dim sKey, sVer
    Dim fReplaced

    fReplaced = False
    For Each sVer in dicActiveC2Rv2Versions.Keys
        sKey = REG_OFFICE & sVer & REG_C2RVIRT_HKLM
        If InStr(sSubKeyName, sKey) > 0 Then
            sSubKeyName = Replace(sSubKeyName, sKey, "")
            fReplaced = True
            Exit For
        End If
    Next
    Select Case hDefKey
        Case HKCU
            If Left(sSubKeyName,17) = "Software\Classes\" Then
                Wow64Key = Left(sSubKeyName,17) & "Wow6432Node\" & Right(sSubKeyName,Len(sSubKeyName)-17)
            Else
                iPos = InStr(sSubKeyName,"\")
                Wow64Key = Left(sSubKeyName,iPos) & "Wow6432Node\" & Right(sSubKeyName,Len(sSubKeyName)-iPos)
            End If
        Case HKLM
            If Left(sSubKeyName,17) = "Software\Classes\" Then
                Wow64Key = Left(sSubKeyName,17) & "Wow6432Node\" & Right(sSubKeyName,Len(sSubKeyName)-17)
            Else
                iPos = InStr(sSubKeyName,"\")
                Wow64Key = Left(sSubKeyName,iPos) & "Wow6432Node\" & Right(sSubKeyName,Len(sSubKeyName)-iPos)
            End If
        Case Else
            Wow64Key = "Wow6432Node\" & sSubKeyName
    End Select 'hDefKey
    If fReplaced Then
        sSubKeyName = sKey & sSubKeyName
        Wow64Key = sKey & Wow64Key
    End If
End Function 'Wow64Key
'=======================================================================================================

'64 bit aware wrapper to return the requested folder 
Function GetFolderPath(sPath)
    GetFolderPath = True
    If oFso.FolderExists(sPath) Then Exit Function
    If f64 AND oFso.FolderExists(Wow64Folder(sPath)) Then
        sPath = Wow64Folder(sPath)
        Exit Function
    End If
    GetFolderPath = False
End Function 'GetFolderPath
'=======================================================================================================

'Enumerates subfolder names of a folder and returns True if subfolders exist
Function EnumFolderNames (sFolder, arrSubFolders)
    Dim Folder, Subfolder
    Dim sSubFolders
    
    If oFso.FolderExists(sFolder) Then
        Set Folder = oFso.GetFolder(sFolder)
        For Each Subfolder in Folder.Subfolders
            sSubFolders = sSubFolders & Subfolder.Name & ","
        Next 'Subfolder
    End If
    If f64 AND oFso.FolderExists(Wow64Folder(sFolder)) Then
        Set Folder = oFso.GetFolder(Wow64Folder(sFolder))
        For Each Subfolder in Folder.Subfolders
            sSubFolders = sSubFolders & Subfolder.Name & ","
        Next 'Subfolder
    End If
    If Len(sSubFolders)>0 Then arrSubFolders = RemoveDuplicates(Split(Left(sSubFolders,Len(sSubFolders)-1),","))
    EnumFolderNames = Len(sSubFolders)>0
End Function 'EnumFolderNames
'=======================================================================================================

'Enumerates subfolders of a folder and returns True if subfolders exist
Function EnumFolders (sFolder, arrSubFolders)
    Dim Folder, Subfolder
    Dim sSubFolders
    
    If oFso.FolderExists(sFolder) Then
        Set Folder = oFso.GetFolder(sFolder)
        For Each Subfolder in Folder.Subfolders
            sSubFolders = sSubFolders & Subfolder.Path & ","
        Next 'Subfolder
    End If
    If f64 AND oFso.FolderExists(Wow64Folder(sFolder)) Then
        Set Folder = oFso.GetFolder(Wow64Folder(sFolder))
        For Each Subfolder in Folder.Subfolders
            sSubFolders = sSubFolders & Subfolder.Path & ","
        Next 'Subfolder
    End If
    If Len(sSubFolders)>0 Then arrSubFolders = RemoveDuplicates(Split(Left(sSubFolders,Len(sSubFolders)-1),","))
    EnumFolders = Len(sSubFolders)>0
End Function 'EnumFolders
'=======================================================================================================

Sub GetFolderStructure (Folder)
    Dim SubFolder
    
    For Each SubFolder in Folder.SubFolders
        ReDim Preserve arrMseFolders(UBound(arrMseFolders)+1)
        arrMseFolders(UBound(arrMseFolders)) = SubFolder.Path
        GetFolderStructure SubFolder
    Next 'SubFolder
End Sub 'GetFolderStructure
'=======================================================================================================

'Remove duplicate entries from a one dimensional array
Function RemoveDuplicates(Array)
    Dim Item
    Dim oDic
    
    Set oDic = CreateObject("Scripting.Dictionary")
    For Each Item in Array
        If Not oDic.Exists(Item) Then oDic.Add Item,Item
    Next 'Item
    RemoveDuplicates = oDic.Keys
End Function 'RemoveDuplicates
'=======================================================================================================

'Identify user SID's on the system
Function GetUserSids(sContext)
    Dim i, n, iRetVal
    Dim arrKeys
    On Error Resume Next

    sUserName = oShell.ExpandEnvironmentStrings ("%USERNAME%") 
    sDomain = oShell.ExpandEnvironmentStrings ("%USERDOMAIN%") 

    ReDim arrKeys(-1)
    Select Case sContext
    Case "Current"
        If RegEnumKey(HKCU,"Software\Microsoft\Protected Storage System Provider\", arrKeys) Then
            sCurUserSid = arrKeys(0)
        Else
            sCurUserSid = GetObject ("winmgmts:\\.\root\cimv2:Win32_UserAccount.Domain='" & sDomain & "',Name='" & sUserName & "'").SID
        End If 'RegEnumKey
    
    'Add SID's that are not "S-1-5-18" (Len("S-1-5-18") = 8
    Case "UserUnmanaged"
        If RegEnumKey(HKLM, REG_GLOBALCONFIG, arrKeys) Then
            n = 0
            For i = 0 To UBound(arrKeys)
                If Len(arrKeys(i)) > 8 Then
                    Redim Preserve arrUUSids(n)
                    arrUUSids(n) = arrKeys(i)
                    n = n + 1
                End If 'Len(arrKeys)
            Next 'i
        End If 'RegEnumKey
    Case "UserManaged"
        If RegEnumKey(HKLM, REG_CONTEXTUSERMANAGED, arrKeys) Then
            n = 0
            For i = 0 To UBound(arrKeys)
                If Len(arrKeys(i)) > 8 Then
                    Redim Preserve arrUMSids(n)
                    arrUMSids(n) = arrKeys(i)
                    n = n + 1
                End If 'Len(arrKeys)
            Next 'i
        End If 'RegEnumKey
    Case Else
    End Select
    
End Function
'=======================================================================================================

Function GetArrayPosition (Arr, sProductCode)
    Dim iPos
    On Error Resume Next

    GetArrayPosition = -1
    'Need to allow exception for lower case only violations 'fGuidCaseWarningOnly'
    If CheckArray(Arr) And (IsValidGuid(sProductCode,GUID_UNCOMPRESSED) OR fGuidCaseWarningOnly) Then
        For iPos = 0 To UBound(Arr)
            If Arr(iPos,COL_PRODUCTCODE) = sProductCode Then 
                GetArrayPosition = iPos
                Exit For
            End If
        Next 'iPos
    End If 'CheckArray
    If iPos = -1 Then WriteDebug sActiveSub, "Warning: Invalid ArrayPosition for " & sProductCode & " - Stack: " & sStack
End Function
'=======================================================================================================

Function GetArrayPositionFromPattern (Arr, sProductCodePattern)
    Dim iPos
    On Error Resume Next

    GetArrayPositionFromPattern = -1
    'Need to allow exception for lower case only violations 'fGuidCaseWarningOnly'
    If CheckArray(Arr) Then
        For iPos = 0 To UBound(Arr)
            If InStr(Arr(iPos,COL_PRODUCTCODE), sProductCodePattern) > 0 Then 
                GetArrayPositionFromPattern = iPos
                Exit For
            End If
        Next 'iPos
    End If 'CheckArray
End Function
'=======================================================================================================

Sub InitMasterArray
    On Error Resume Next
    
    'Since ReDim cannot preserve the data on the first dimension determine the needed total first.
    If CheckArray(arrAllProducts) Then iPCount = iPCount + (UBound(arrAllProducts) + 1)
    If CheckArray(arrMProducts) Then iPCount = iPCount + (UBound(arrMProducts) + 1)
    If CheckArray(arrUUProducts) Then iPCount = iPCount + (UBound(arrUUProducts) + 1)
    If CheckArray(arrUMProducts) Then iPCount = iPCount + (UBound(arrUMProducts) + 1)
    If CheckArray(arrMVProducts) Then iPCount = iPCount + (UBound(arrMVProducts) + 1 )
    
    ReDim arrMaster(iPCount-1,UBOUND_MASTER)
    
    If CheckArray(arrAllProducts) Then CopyToMaster arrAllProducts
    If CheckArray(arrMProducts) Then CopyToMaster arrMProducts
    If CheckArray(arrUUProducts) Then CopyToMaster arrUUProducts
    If CheckArray(arrUMProducts) Then CopyToMaster arrUMProducts
    If CheckArray(arrMVProducts) Then CopyToMaster arrMVProducts
End Sub
'=======================================================================================================

Function CheckArray(Arr)
    Dim sTmp
    Dim iRetVal
    On Error Resume Next
    
    CheckArray = True
    If Not IsArray(Arr) Then 
        CheckArray = False
        Exit Function
    End If
    If IsNull(Arr) Then 
        CheckArray = False
        Exit Function
    End If
    If IsEmpty(Arr) Then 
        CheckArray = False
        Exit Function
    End If
    If UBound(Arr) = -1 Then 
        If Not Err = 0 Then 
            Err.Clear
            Redim Arr(-1)
        End If 'Err
        CheckArray = False
        Exit Function
    End If
End Function
 
'=======================================================================================================

Function CheckObject(Obj)
    On Error Resume Next
    
    CheckObject = True
    If Not IsObject(Obj) Then 
        CheckObject = False
        Exit Function
    End If
    If IsNull(Obj) Then 
        CheckObject = False
        Exit Function
    End If
    If IsEmpty(Obj) Then 
        CheckObject = False
        Exit Function
    End If
    If Not Obj.Count > 0 Then 
        If Not Err = 0 Then
            Err.Clear
            Set Obj = Nothing
        End If 'Err
        CheckObject = False
        Exit Function
    End If
End Function
'=======================================================================================================

Sub CheckError(sModule,sErrorHandler)
    Dim sErr
    If Not Err = 0 Then 
        sErr = GetErrorDescription(Err)
        If Not sErrorHandler = "" Then 
            sErrorHandler = sModule & sErrorHandler
            ErrorRelay(sErrorHandler)
        End If
    End If 'Err = 0
    Err.Clear
End Sub
'=======================================================================================================

Sub ErrorRelay(sErrorHandler)
    Select Case (sErrorHandler)
        Case "CheckPreReq_ErrorHandler" : CheckPreReq_ErrorHandler
        Case "FindAllProducts_ErrorHandler3x" : FindAllProducts_ErrorHandler3x
        Case "FindProducts1","FindProducts2","FindProducts4" : FindProducts_ErrorHandler Int(Right(sErrorHandler,1))
        Case Else
    End Select
End Sub
'=======================================================================================================

Function GetErrorDescription (Err)
    If Not Err = 0 Then 
        GetErrorDescription = "Source: " & Err.Source & "; Err# (Hex): " & Hex( Err ) & _
                              "; Err# (Dec): " & Err & "; Description : " & Err.Description & _
                              "; Stack: " & sStack
    End If 'Err = 0
End Function
'=======================================================================================================

Sub ParseCmdLine
    Dim iCnt, iArgCnt
    
    'Handle settings from the .ini section
    If fLogFull Then
        fLogChainedDetails = True
        fFileInventory = True
        fFeatureTree = True
    End If 'fLogFull
    
    If fLogVerbose Then
        fLogChainedDetails = True
        fFeatureTree = True
    End If 'fLogVerbose
    
    iArgCnt = Wscript.Arguments.Count
    If iArgCnt>0 Then
        For iCnt = 0 To iArgCnt-1
            Select Case UCase(Wscript.Arguments(iCnt))
            
            Case "/A","-A","/ALL","-ALL","ALL"
                fListNonOfficeProducts = True
            Case "/BASIC","-BASIC","BASIC"
                fBasicMode = True
            Case "/DCS","-DCS","/DISALLOWCSCRIPT","-DISALLOWCSCRIPT","DISALLOWCSCRIPT"
                fDisallowCScript = True
            Case "/F","-F","/FULL","-FULL","FULL"
                fLogFull = True
                fLogChainedDetails = True
                fFileInventory = True
                fFeatureTree = True
            Case "/FI","-FI","/FILEINVENTORY","-FILEINVENTORY","FILEINVENTORY"
                fFileInventory = True
            Case "/FT","-FT","/FEATURETREE","-FEATURETREE","FEATURETREE"
                fFeatureTree = True
            Case "/L","-L","/LOGFOLDER","-LOGFOLDER","LOGFOLDER"
                If iArgCnt > iCnt Then sPathOutputFolder = Wscript.Arguments(iCnt + 1)
            Case "/LD","-LD","/LOGCHAINEDDETAILS","-LOGCHAINEDDETAILS"
                fLogChainedDetails = True
            Case "/LV","-LV","/L*V","-L*V","/LOGVERBOSE","-LOGVERBOSE","/VERBOSE","-VERBOSE"
                fLogVerbose = True
                fLogChainedDetails = True
                fFeatureTree = True
            Case "/Q","-Q","/QUIET","-QUIET","QUIET"
                fQuiet = True
            Case "/?","-?","?"
                Wscript.Echo vbCrLf & _
                 "ROIScan (Robust Office Inventory Scan) - Version " & SCRIPTBUILD & vbCrLf & _
                 "Copyright (c) 2008,2009,2010 Microsoft Corporation. All Rights Reserved." & vbCrLf & vbCrLf & _
                 "Inventory tool for to create a log of " & vbCrLf & "installed Microsoft Office applications." & vbCrLf & _
                 "Supports Office 2000, XP, 2003, 2007, 2010, 2013, 2016, O365 " & vbCrLf & vbCrLf & _
                 "Usage:" & vbTab & "ROIScan.vbs [Options]" & vbCrLf & vbCrLf & _
                 " /?" & vbTab & vbTab & vbTab & "Display this help"& vbCrLf &_
                 " /All" & vbTab & vbTab & vbTab & "Include non Office products" & vbCrLf &_
                 " /Basic" & vbTab & vbTab & vbTab & "Log less product details" & vbCrLf & _
                 " /Full" & vbTab & vbTab & vbTab & "Log all additional product details" & vbCrLf & _
                 " /LogVerbose" & vbTab & vbTab & "Log additional product details" & vbCrLf & _
                 " /LogChainedDetails" & vbTab & "Log details for chained .msi packages" & vbCrLf & _
                 " /FeatureTree" & vbTab & vbTab & "Include treeview like feature listing" & vbCrLf & _
                 " /FileInventory" & vbTab & vbTab & "Scan all file version" & vbCrLf & _
                 " /Logfolder <Path>" & vbTab & "Custom directory for log file" & vbCrLf &_
                 " /Quiet" & vbTab & vbTab & vbTab & "Don't open log when done"
                Wscript.Quit
            Case Else
            End Select
        Next 'iCnt
    End If 'iArgCnt>0
End Sub
'=======================================================================================================

Sub CleanUp
    On Error Resume Next
    Set oReg = Nothing
    TextStream.Close
    Set TextStream = Nothing
    Set oFso = Nothing
    Set oMsi = Nothing
    Set oShell = Nothing
End Sub
'=======================================================================================================
