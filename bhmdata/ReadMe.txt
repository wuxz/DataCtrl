========================================================================
		ActiveX Control DLL : BHMDATA
========================================================================

ControlWizard has created this project for your BHMDATA ActiveX Control
DLL, which contains 3 controls.

This skeleton project not only demonstrates the basics of writing an
ActiveX Control, but is also a starting point for writing the specific
features of your control.

This file contains a summary of what you will find in each of the files
that make up your BHMDATA ActiveX Control DLL.

BHMData.dsp
    This file (the project file) contains information at the project level and
    is used to build a single project or subproject. Other users can share the
    project (.dsp) file, but they should export the makefiles locally.

BHMData.h
	This is the main include file for the ActiveX Control DLL.  It
	includes other project-specific includes such as resource.h.

BHMData.cpp
	This is the main source file that contains code for DLL initialization,
	termination and other bookkeeping.

BHMData.rc
	This is a listing of the Microsoft Windows resources that the project
	uses.  This file can be directly edited with the Visual C++ resource
	editor.

BHMData.def
	This file contains information about the ActiveX Control DLL that
	must be provided to run with Microsoft Windows.

BHMData.clw
	This file contains information used by ClassWizard to edit existing
	classes or add new classes.  ClassWizard also uses this file to store
	information needed to generate and edit message maps and dialog data
	maps and to generate prototype member functions.

BHMData.odl
	This file contains the Object Description Language source code for the
	type library of your controls.

/////////////////////////////////////////////////////////////////////////////
BHMDBGrid control:

BHMDBGridCtl.h
	This file contains the declaration of the CBHMDBGridCtrl C++ class.

BHMDBGridCtl.cpp
	This file contains the implementation of the CBHMDBGridCtrl C++ class.

BHMDBGridPpgGeneral.h
	This file contains the declaration of the CBHMDBGridPropPageGeneral C++ class.

BHMDBGridPpgGeneral.cpp
	This file contains the implementation of the CBHMDBGridPropPageGeneral C++ class.

BHMDBGridCtl.bmp
	This file contains a bitmap that a container will use to represent the
	CBHMDBGridCtrl control when it appears on a tool palette.  This bitmap
	is included by the main resource file BHMData.rc.

/////////////////////////////////////////////////////////////////////////////
BHMDBDropDown control:

BHMDBDropDownCtl.h
	This file contains the declaration of the CBHMDBDropDownCtrl C++ class.

BHMDBDropDownCtl.cpp
	This file contains the implementation of the CBHMDBDropDownCtrl C++ class.

BHMDBDropDownPpgGeneral.h
	This file contains the declaration of the CBHMDBDropDownPropPageGeneral C++ class.

BHMDBDropDownPpgGeneral.cpp
	This file contains the implementation of the CBHMDBDropDownPropPageGeneral C++ class.

BHMDBDropDownCtl.bmp
	This file contains a bitmap that a container will use to represent the
	CBHMDBDropDownCtrl control when it appears on a tool palette.  This bitmap
	is included by the main resource file BHMData.rc.

/////////////////////////////////////////////////////////////////////////////
BHMDBCombo control:

BHMDBComboCtl.h
	This file contains the declaration of the CBHMDBComboCtrl C++ class.

BHMDBComboCtl.cpp
	This file contains the implementation of the CBHMDBComboCtrl C++ class.

BHMDBComboPpgGeneral.h
	This file contains the declaration of the CBHMDBComboPropPageGeneral C++ class.

BHMDBComboPpgGeneral.cpp
	This file contains the implementation of the CBHMDBComboPropPageGeneral C++ class.

BHMDBComboCtl.bmp
	This file contains a bitmap that a container will use to represent the
	CBHMDBComboCtrl control when it appears on a tool palette.  This bitmap
	is included by the main resource file BHMData.rc.

/////////////////////////////////////////////////////////////////////////////
Help Support:

BHMData.hpj
	This file is the Help Project file used by the Help compiler to create
	your ActiveX Control's Help file.

*.bmp
	These are bitmap files required by the standard Help file topics for
	Microsoft Foundation Class Library standard commands.  These files are
	located in the HLP subdirectory.

BHMData.rtf
	This file contains the standard help topics for the common properties,
	events and methods supported by many ActiveX Controls.  You can edit
	this to add or remove additional control specific topics.  This file is
	located in the HLP subdirectory.

/////////////////////////////////////////////////////////////////////////////
Licensing support:
BHMData.lic
	This is the user license file.  This file must be present in the same
	directory as the control's DLL to allow an instance of the control to be
	created in a design-time environment.  Typically, you will distribute
	this file with your control, but your customers will not distribute it.

/////////////////////////////////////////////////////////////////////////////
Other standard files:

stdafx.h, stdafx.cpp
	These files are used to build a precompiled header (PCH) file
	named stdafx.pch and a precompiled types (PCT) file named stdafx.obj.

resource.h
	This is the standard header file, which defines new resource IDs.
	The Visual C++ resource editor reads and updates this file.

/////////////////////////////////////////////////////////////////////////////
Other notes:

ControlWizard uses "TODO:" to indicate parts of the source code you
should add to or customize.

/////////////////////////////////////////////////////////////////////////////
