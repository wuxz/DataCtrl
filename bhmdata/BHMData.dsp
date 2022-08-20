# Microsoft Developer Studio Project File - Name="BHMData" - Package Owner=<4>
# Microsoft Developer Studio Generated Build File, Format Version 6.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Dynamic-Link Library" 0x0102

CFG=BHMData - Win32 Debug
!MESSAGE This is not a valid makefile. To build this project using NMAKE,
!MESSAGE use the Export Makefile command and run
!MESSAGE 
!MESSAGE NMAKE /f "BHMData.mak".
!MESSAGE 
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "BHMData.mak" CFG="BHMData - Win32 Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "BHMData - Win32 Release" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "BHMData - Win32 Debug" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "BHMData - Win32 Unicode Debug" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "BHMData - Win32 Unicode Release" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE 

# Begin Project
# PROP AllowPerConfigDependencies 0
# PROP Scc_ProjName ""
# PROP Scc_LocalPath ""
CPP=cl.exe
MTL=midl.exe
RSC=rc.exe

!IF  "$(CFG)" == "BHMData - Win32 Release"

# PROP BASE Use_MFC 2
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "Release"
# PROP BASE Intermediate_Dir "Release"
# PROP BASE Target_Ext "ocx"
# PROP BASE Target_Dir ""
# PROP Use_MFC 1
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "Release"
# PROP Intermediate_Dir "Release"
# PROP Target_Ext "ocx"
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MD /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_WINDLL" /D "_AFXDLL" /Yu"stdafx.h" /FD /c
# ADD CPP /nologo /MT /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "_WINDLL" /Yu"stdafx.h" /FD /c
# ADD BASE MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x804 /d "NDEBUG" /d "_AFXDLL"
# ADD RSC /l 0x804 /d "NDEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 /nologo /subsystem:windows /dll /machine:I386
# ADD LINK32 /nologo /subsystem:windows /dll /machine:I386
# Begin Custom Build - Registering ActiveX Control...
OutDir=.\Release
TargetPath=.\Release\BHMData.ocx
InputPath=.\Release\BHMData.ocx
SOURCE="$(InputPath)"

"$(OutDir)\regsvr32.trg" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	regsvr32 /s /c "$(TargetPath)" 
	echo regsvr32 exec. time > "$(OutDir)\regsvr32.trg" 
	
# End Custom Build
# Begin Special Build Tool
SOURCE="$(InputPath)"
PostBuild_Desc=Copying file into share directory
PostBuild_Cmds=copy release\*.ocx \temp
# End Special Build Tool

!ELSEIF  "$(CFG)" == "BHMData - Win32 Debug"

# PROP BASE Use_MFC 2
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir "Debug"
# PROP BASE Intermediate_Dir "Debug"
# PROP BASE Target_Ext "ocx"
# PROP BASE Target_Dir ""
# PROP Use_MFC 2
# PROP Use_Debug_Libraries 1
# PROP Output_Dir "Debug"
# PROP Intermediate_Dir "Debug"
# PROP Target_Ext "ocx"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MDd /W3 /Gm /GX /ZI /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_WINDLL" /D "_AFXDLL" /Yu"stdafx.h" /FD /GZ /c
# ADD CPP /nologo /MDd /W3 /Gm /GX /ZI /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_WINDLL" /D "_AFXDLL" /D "_MBCS" /D "_USRDLL" /FR /Yu"stdafx.h" /FD /GZ /c
# ADD BASE MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x804 /d "_DEBUG" /d "_AFXDLL"
# ADD RSC /l 0x804 /d "_DEBUG" /d "_AFXDLL"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 /nologo /subsystem:windows /dll /debug /machine:I386 /pdbtype:sept
# ADD LINK32 /nologo /subsystem:windows /dll /debug /machine:I386 /pdbtype:sept
# Begin Custom Build - Registering ActiveX Control...
OutDir=.\Debug
TargetPath=.\Debug\BHMData.ocx
InputPath=.\Debug\BHMData.ocx
SOURCE="$(InputPath)"

"$(OutDir)\regsvr32.trg" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	regsvr32 /s /c "$(TargetPath)" 
	echo regsvr32 exec. time > "$(OutDir)\regsvr32.trg" 
	
# End Custom Build

!ELSEIF  "$(CFG)" == "BHMData - Win32 Unicode Debug"

# PROP BASE Use_MFC 2
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir "DebugU"
# PROP BASE Intermediate_Dir "DebugU"
# PROP BASE Target_Ext "ocx"
# PROP BASE Target_Dir ""
# PROP Use_MFC 2
# PROP Use_Debug_Libraries 1
# PROP Output_Dir "DebugU"
# PROP Intermediate_Dir "DebugU"
# PROP Target_Ext "ocx"
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MDd /W3 /Gm /GX /ZI /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_WINDLL" /D "_AFXDLL" /D "_MBCS" /D "_USRDLL" /Yu"stdafx.h" /FD /GZ /c
# ADD CPP /nologo /MDd /W3 /Gm /GX /ZI /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_WINDLL" /D "_AFXDLL" /D "_USRDLL" /D "_UNICODE" /Yu"stdafx.h" /FD /GZ /c
# ADD BASE MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x804 /d "_DEBUG" /d "_AFXDLL"
# ADD RSC /l 0x804 /d "_DEBUG" /d "_AFXDLL"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 /nologo /subsystem:windows /dll /debug /machine:I386 /pdbtype:sept
# ADD LINK32 /nologo /subsystem:windows /dll /debug /machine:I386 /pdbtype:sept
# Begin Custom Build - Registering ActiveX Control...
OutDir=.\DebugU
TargetPath=.\DebugU\BHMData.ocx
InputPath=.\DebugU\BHMData.ocx
SOURCE="$(InputPath)"

"$(OutDir)\regsvr32.trg" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	regsvr32 /s /c "$(TargetPath)" 
	echo regsvr32 exec. time > "$(OutDir)\regsvr32.trg" 
	
# End Custom Build

!ELSEIF  "$(CFG)" == "BHMData - Win32 Unicode Release"

# PROP BASE Use_MFC 2
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "ReleaseU"
# PROP BASE Intermediate_Dir "ReleaseU"
# PROP BASE Target_Ext "ocx"
# PROP BASE Target_Dir ""
# PROP Use_MFC 2
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "ReleaseU"
# PROP Intermediate_Dir "ReleaseU"
# PROP Target_Ext "ocx"
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MD /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_WINDLL" /D "_AFXDLL" /D "_MBCS" /D "_USRDLL" /Yu"stdafx.h" /FD /c
# ADD CPP /nologo /MD /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_WINDLL" /D "_AFXDLL" /D "_USRDLL" /D "_UNICODE" /Yu"stdafx.h" /FD /c
# ADD BASE MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x804 /d "NDEBUG" /d "_AFXDLL"
# ADD RSC /l 0x804 /d "NDEBUG" /d "_AFXDLL"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 /nologo /subsystem:windows /dll /machine:I386
# ADD LINK32 /nologo /subsystem:windows /dll /machine:I386
# Begin Custom Build - Registering ActiveX Control...
OutDir=.\ReleaseU
TargetPath=.\ReleaseU\BHMData.ocx
InputPath=.\ReleaseU\BHMData.ocx
SOURCE="$(InputPath)"

"$(OutDir)\regsvr32.trg" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	regsvr32 /s /c "$(TargetPath)" 
	echo regsvr32 exec. time > "$(OutDir)\regsvr32.trg" 
	
# End Custom Build

!ENDIF 

# Begin Target

# Name "BHMData - Win32 Release"
# Name "BHMData - Win32 Debug"
# Name "BHMData - Win32 Unicode Debug"
# Name "BHMData - Win32 Unicode Release"
# Begin Group "Source Files"

# PROP Default_Filter "cpp;c;cxx;rc;def;r;odl;idl;hpj;bat"
# Begin Source File

SOURCE=.\AddColumnDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\BHMData.cpp
# End Source File
# Begin Source File

SOURCE=.\BHMData.def
# End Source File
# Begin Source File

SOURCE=.\BHMData.hpj

!IF  "$(CFG)" == "BHMData - Win32 Release"

# PROP Ignore_Default_Tool 1
# Begin Custom Build - Building Help ...
OutDir=.\Release
InputPath=.\BHMData.hpj
InputName=BHMData

BuildCmds= \
	start /wait hcw /C /E /M "$(InputName).hpj" \
	if errorlevel 1 goto :Error \
	if not exist "$(InputName).hlp" goto :Error \
	copy "$(InputName).hlp" $(OutDir) \
	goto :done \
	:Error \
	echo $(InputName).hpj(1) : error: Problem encountered creating help file \
	type $(InputName).log \
	:done \
	

"$(InputName).hlp" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
   $(BuildCmds)

"$(OutDir)\$(InputName).hlp" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
   $(BuildCmds)
# End Custom Build

!ELSEIF  "$(CFG)" == "BHMData - Win32 Debug"

# PROP Ignore_Default_Tool 1
# Begin Custom Build - Building Help ...
OutDir=.\Debug
InputPath=.\BHMData.hpj
InputName=BHMData

BuildCmds= \
	start /wait hcw /C /E /M "$(InputName).hpj" \
	if errorlevel 1 goto :Error \
	if not exist "$(InputName).hlp" goto :Error \
	copy "$(InputName).hlp" $(OutDir) \
	goto :done \
	:Error \
	echo $(InputName).hpj(1) : error: Problem encountered creating help file \
	type $(InputName).log \
	:done \
	

"$(InputName).hlp" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
   $(BuildCmds)

"$(OutDir)\$(InputName).hlp" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
   $(BuildCmds)
# End Custom Build

!ELSEIF  "$(CFG)" == "BHMData - Win32 Unicode Debug"

# PROP BASE Ignore_Default_Tool 1
# PROP Ignore_Default_Tool 1
# Begin Custom Build - Building Help ...
OutDir=.\DebugU
InputPath=.\BHMData.hpj
InputName=BHMData

BuildCmds= \
	start /wait hcw /C /E /M "$(InputName).hpj" \
	if errorlevel 1 goto :Error \
	if not exist "$(InputName).hlp" goto :Error \
	copy "$(InputName).hlp" $(OutDir) \
	goto :done \
	:Error \
	echo $(InputName).hpj(1) : error: Problem encountered creating help file \
	type $(InputName).log \
	:done \
	

"$(InputName).hlp" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
   $(BuildCmds)

"$(OutDir)\$(InputName).hlp" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
   $(BuildCmds)
# End Custom Build

!ELSEIF  "$(CFG)" == "BHMData - Win32 Unicode Release"

# PROP BASE Ignore_Default_Tool 1
# PROP Ignore_Default_Tool 1
# Begin Custom Build - Building Help ...
OutDir=.\ReleaseU
InputPath=.\BHMData.hpj
InputName=BHMData

BuildCmds= \
	start /wait hcw /C /E /M "$(InputName).hpj" \
	if errorlevel 1 goto :Error \
	if not exist "$(InputName).hlp" goto :Error \
	copy "$(InputName).hlp" $(OutDir) \
	goto :done \
	:Error \
	echo $(InputName).hpj(1) : error: Problem encountered creating help file \
	type $(InputName).log \
	:done \
	

"$(InputName).hlp" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
   $(BuildCmds)

"$(OutDir)\$(InputName).hlp" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
   $(BuildCmds)
# End Custom Build

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\BHMData.odl
# End Source File
# Begin Source File

SOURCE=.\BHMData.rc
# End Source File
# Begin Source File

SOURCE=.\BHMDBColumnPropPage.cpp
# End Source File
# Begin Source File

SOURCE=.\BHMDBComboCtl.cpp
# End Source File
# Begin Source File

SOURCE=.\BHMDBComboPpgGeneral.cpp
# End Source File
# Begin Source File

SOURCE=.\BHMDBDropDownCtl.cpp
# End Source File
# Begin Source File

SOURCE=.\BHMDBDropDownPpgGeneral.cpp
# End Source File
# Begin Source File

SOURCE=.\BHMDBGridCtl.cpp
# End Source File
# Begin Source File

SOURCE=.\BHMDBGridPpgGeneral.cpp
# End Source File
# Begin Source File

SOURCE=.\BHMDBGroupsPropPage.cpp
# End Source File
# Begin Source File

SOURCE=.\Column.cpp
# End Source File
# Begin Source File

SOURCE=.\Columns.cpp
# End Source File
# Begin Source File

SOURCE=.\ComboEdit.cpp
# End Source File
# Begin Source File

SOURCE=..\testgrid\DateMaskEdit.cpp
# End Source File
# Begin Source File

SOURCE=.\DropDownRealWnd.cpp
# End Source File
# Begin Source File

SOURCE=.\FieldsDlg.cpp
# End Source File
# Begin Source File

SOURCE=..\testgrid\Grid.cpp
# End Source File
# Begin Source File

SOURCE=..\testgrid\GridCheckBoxCtrl.cpp
# End Source File
# Begin Source File

SOURCE=..\testgrid\GridComboCtrl.cpp
# End Source File
# Begin Source File

SOURCE=..\testgrid\GridComboEdit.cpp
# End Source File
# Begin Source File

SOURCE=..\testgrid\GridControl.cpp
# End Source File
# Begin Source File

SOURCE=..\testgrid\GridDateMaskCtrl.cpp
# End Source File
# Begin Source File

SOURCE=..\testgrid\GridEditBtnCtrl.cpp
# End Source File
# Begin Source File

SOURCE=..\testgrid\GridEditCtrl.cpp
# End Source File
# Begin Source File

SOURCE=.\GridInner.cpp
# End Source File
# Begin Source File

SOURCE=.\GridInPpg.cpp
# End Source File
# Begin Source File

SOURCE=.\GridInput.cpp
# End Source File
# Begin Source File

SOURCE=.\GridInputDlg.cpp
# End Source File
# Begin Source File

SOURCE=..\testgrid\GridNumCtrl.cpp
# End Source File
# Begin Source File

SOURCE=.\Group.cpp
# End Source File
# Begin Source File

SOURCE=.\Groups.cpp
# End Source File
# Begin Source File

SOURCE=..\testgrid\MaskEdit.cpp
# End Source File
# Begin Source File

SOURCE=..\testgrid\NumEdit.cpp
# End Source File
# Begin Source File

SOURCE=.\SelBookmarks.cpp
# End Source File
# Begin Source File

SOURCE=.\StdAfx.cpp
# ADD CPP /Yc"stdafx.h"
# End Source File
# End Group
# Begin Group "Header Files"

# PROP Default_Filter "h;hpp;hxx;hm;inl"
# Begin Source File

SOURCE=.\AddColumnDlg.h
# End Source File
# Begin Source File

SOURCE=.\BHMData.h
# End Source File
# Begin Source File

SOURCE=.\BHMDBColumnPropPage.h
# End Source File
# Begin Source File

SOURCE=.\BHMDBComboCtl.h
# End Source File
# Begin Source File

SOURCE=.\BHMDBComboPpgGeneral.h
# End Source File
# Begin Source File

SOURCE=.\BHMDBDropDownCtl.h
# End Source File
# Begin Source File

SOURCE=.\BHMDBDropDownPpgGeneral.h
# End Source File
# Begin Source File

SOURCE=.\BHMDBGridCtl.h
# End Source File
# Begin Source File

SOURCE=.\BHMDBGridPpgGeneral.h
# End Source File
# Begin Source File

SOURCE=.\BHMDBGroupsPropPage.h
# End Source File
# Begin Source File

SOURCE=.\Column.h
# End Source File
# Begin Source File

SOURCE=.\Columns.h
# End Source File
# Begin Source File

SOURCE=.\ComboEdit.h
# End Source File
# Begin Source File

SOURCE=.\DropDownRealWnd.h
# End Source File
# Begin Source File

SOURCE=.\FieldsDlg.h
# End Source File
# Begin Source File

SOURCE=.\GridInner.h
# End Source File
# Begin Source File

SOURCE=.\GridInPpg.h
# End Source File
# Begin Source File

SOURCE=.\GridInput.h
# End Source File
# Begin Source File

SOURCE=.\GridInputDlg.h
# End Source File
# Begin Source File

SOURCE=.\Group.h
# End Source File
# Begin Source File

SOURCE=.\Groups.h
# End Source File
# Begin Source File

SOURCE=.\Resource.h

!IF  "$(CFG)" == "BHMData - Win32 Release"

# PROP Ignore_Default_Tool 1
# Begin Custom Build - Building Help Include File ...
TargetName=BHMData
InputPath=.\Resource.h

"hlp\$(TargetName).hm" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	echo. >"hlp\$(TargetName).hm" 
	echo // Commands (ID_* and IDM_*) >>"hlp\$(TargetName).hm" 
	makehm ID_,HID_,0x10000 IDM_,HIDM_,0x10000 resource.h >>"hlp\$(TargetName).hm" 
	echo // Prompts (IDP_*) >>"hlp\$(TargetName).hm" 
	makehm IDP_,HIDP_,0x30000 resource.h >>"hlp\$(TargetName).hm" 
	echo // Resources (IDR_*) >>"hlp\$(TargetName).hm" 
	makehm IDR_,HIDR_,0x20000 resource.h >>"hlp\$(TargetName).hm" 
	echo // Dialogs (IDD_*) >>"hlp\$(TargetName).hm" 
	makehm IDD_,HIDD_,0x20000 resource.h >>"hlp\$(TargetName).hm" 
	echo // Frame Controls (IDW_*) >>"hlp\$(TargetName).hm" 
	makehm IDW_,HIDW_,0x50000 resource.h >>"hlp\$(TargetName).hm" 
	
# End Custom Build

!ELSEIF  "$(CFG)" == "BHMData - Win32 Debug"

# PROP Ignore_Default_Tool 1
# Begin Custom Build - Building Help Include File ...
TargetName=BHMData
InputPath=.\Resource.h

"hlp\$(TargetName).hm" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	echo. >"hlp\$(TargetName).hm" 
	echo // Commands (ID_* and IDM_*) >>"hlp\$(TargetName).hm" 
	makehm ID_,HID_,0x10000 IDM_,HIDM_,0x10000 resource.h >>"hlp\$(TargetName).hm" 
	echo // Prompts (IDP_*) >>"hlp\$(TargetName).hm" 
	makehm IDP_,HIDP_,0x30000 resource.h >>"hlp\$(TargetName).hm" 
	echo // Resources (IDR_*) >>"hlp\$(TargetName).hm" 
	makehm IDR_,HIDR_,0x20000 resource.h >>"hlp\$(TargetName).hm" 
	echo // Dialogs (IDD_*) >>"hlp\$(TargetName).hm" 
	makehm IDD_,HIDD_,0x20000 resource.h >>"hlp\$(TargetName).hm" 
	echo // Frame Controls (IDW_*) >>"hlp\$(TargetName).hm" 
	makehm IDW_,HIDW_,0x50000 resource.h >>"hlp\$(TargetName).hm" 
	
# End Custom Build

!ELSEIF  "$(CFG)" == "BHMData - Win32 Unicode Debug"

# PROP BASE Ignore_Default_Tool 1
# PROP Ignore_Default_Tool 1
# Begin Custom Build - Building Help Include File ...
TargetName=BHMData
InputPath=.\Resource.h

"hlp\$(TargetName).hm" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	echo. >"hlp\$(TargetName).hm" 
	echo // Commands (ID_* and IDM_*) >>"hlp\$(TargetName).hm" 
	makehm ID_,HID_,0x10000 IDM_,HIDM_,0x10000 resource.h >>"hlp\$(TargetName).hm" 
	echo // Prompts (IDP_*) >>"hlp\$(TargetName).hm" 
	makehm IDP_,HIDP_,0x30000 resource.h >>"hlp\$(TargetName).hm" 
	echo // Resources (IDR_*) >>"hlp\$(TargetName).hm" 
	makehm IDR_,HIDR_,0x20000 resource.h >>"hlp\$(TargetName).hm" 
	echo // Dialogs (IDD_*) >>"hlp\$(TargetName).hm" 
	makehm IDD_,HIDD_,0x20000 resource.h >>"hlp\$(TargetName).hm" 
	echo // Frame Controls (IDW_*) >>"hlp\$(TargetName).hm" 
	makehm IDW_,HIDW_,0x50000 resource.h >>"hlp\$(TargetName).hm" 
	
# End Custom Build

!ELSEIF  "$(CFG)" == "BHMData - Win32 Unicode Release"

# PROP BASE Ignore_Default_Tool 1
# PROP Ignore_Default_Tool 1
# Begin Custom Build - Building Help Include File ...
TargetName=BHMData
InputPath=.\Resource.h

"hlp\$(TargetName).hm" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	echo. >"hlp\$(TargetName).hm" 
	echo // Commands (ID_* and IDM_*) >>"hlp\$(TargetName).hm" 
	makehm ID_,HID_,0x10000 IDM_,HIDM_,0x10000 resource.h >>"hlp\$(TargetName).hm" 
	echo // Prompts (IDP_*) >>"hlp\$(TargetName).hm" 
	makehm IDP_,HIDP_,0x30000 resource.h >>"hlp\$(TargetName).hm" 
	echo // Resources (IDR_*) >>"hlp\$(TargetName).hm" 
	makehm IDR_,HIDR_,0x20000 resource.h >>"hlp\$(TargetName).hm" 
	echo // Dialogs (IDD_*) >>"hlp\$(TargetName).hm" 
	makehm IDD_,HIDD_,0x20000 resource.h >>"hlp\$(TargetName).hm" 
	echo // Frame Controls (IDW_*) >>"hlp\$(TargetName).hm" 
	makehm IDW_,HIDW_,0x50000 resource.h >>"hlp\$(TargetName).hm" 
	
# End Custom Build

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\SelBookmarks.h
# End Source File
# Begin Source File

SOURCE=.\StdAfx.h
# End Source File
# End Group
# Begin Group "Resource Files"

# PROP Default_Filter "ico;cur;bmp;dlg;rc2;rct;bin;rgs;gif;jpg;jpeg;jpe"
# Begin Source File

SOURCE=.\BHMData.ico
# End Source File
# Begin Source File

SOURCE=.\BHMDBComboCtl.bmp
# End Source File
# Begin Source File

SOURCE=.\BHMDBDropDownCtl.bmp
# End Source File
# Begin Source File

SOURCE=.\BHMDBGridCtl.bmp
# End Source File
# Begin Source File

SOURCE=.\res\bitmap1.bmp
# End Source File
# Begin Source File

SOURCE=.\Hlp\Bullet.bmp
# End Source File
# Begin Source File

SOURCE=.\res\Dragmove.cur
# End Source File
# Begin Source File

SOURCE=.\Header.bmp
# End Source File
# End Group
# Begin Source File

SOURCE=.\BHMData.lic

!IF  "$(CFG)" == "BHMData - Win32 Release"

# PROP Ignore_Default_Tool 1
# Begin Custom Build - Copying license file...
OutDir=.\Release
ProjDir=.
TargetName=BHMData
InputPath=.\BHMData.lic

"$(OutDir)\$(TargetName).lic" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	copy "$(ProjDir)\$(TargetName).lic" "$(OutDir)"

# End Custom Build

!ELSEIF  "$(CFG)" == "BHMData - Win32 Debug"

# PROP Ignore_Default_Tool 1
# Begin Custom Build - Copying license file...
OutDir=.\Debug
ProjDir=.
TargetName=BHMData
InputPath=.\BHMData.lic

"$(OutDir)\$(TargetName).lic" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	copy "$(ProjDir)\$(TargetName).lic" "$(OutDir)"

# End Custom Build

!ELSEIF  "$(CFG)" == "BHMData - Win32 Unicode Debug"

# PROP BASE Ignore_Default_Tool 1
# PROP Ignore_Default_Tool 1
# Begin Custom Build - Copying license file...
OutDir=.\DebugU
ProjDir=.
TargetName=BHMData
InputPath=.\BHMData.lic

"$(OutDir)\$(TargetName).lic" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	copy "$(ProjDir)\$(TargetName).lic" "$(OutDir)"

# End Custom Build

!ELSEIF  "$(CFG)" == "BHMData - Win32 Unicode Release"

# PROP BASE Ignore_Default_Tool 1
# PROP Ignore_Default_Tool 1
# Begin Custom Build - Copying license file...
OutDir=.\ReleaseU
ProjDir=.
TargetName=BHMData
InputPath=.\BHMData.lic

"$(OutDir)\$(TargetName).lic" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	copy "$(ProjDir)\$(TargetName).lic" "$(OutDir)"

# End Custom Build

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\Hlp\BHMData.rtf
# End Source File
# Begin Source File

SOURCE=.\ReadMe.txt
# End Source File
# End Target
# End Project
