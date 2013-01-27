RASCW -- The RASC and CASC programs for ranking, scaling and correlation of stratigraphic events

Current Version: RASCWV20 for Windows

* The RASC program is for automated ranking and scaling of biostratigraphic events.
* Resulting bio-zonations are used for CASC correlation between stratigraphic sections.
* RASCW Version 20 is a user-friendly Windows version with new graphical tools. 
* RASCW is particularly useful for long-distance correlation between exploratory wells. 
* Executable files for RASC & CASC versions with manual and documentation were made available on the website of the International Commission on Stratigraphy and, since 2009, on a website maintained by the University of Oslo (http://www.nhm.uio.no/norlex/rasc).

Notes for RASCW V20 Software Development Environment and Source Code

1. Development platform is: (1) Windows XP; (2) Visual Basic 6.0 Enterprise Edition SP6 The source code has also been tested in Windows 7 operating system and can be compiled and running well. If you want to compile the code in higher version than Visual Basic 6.0, the corresponding migration and modification works are needed.

2. The Visual Basic Project file for RASCW V20 Source Code Package is "RascwV20.vbp", which includes 64 forms and 2 modules: Form=frmCascinput.frm

    Form=frmRascW.frm

    Form=MDIFrmCascRasc.frm

    Form=frmDicInput.frm

    Form=frmOpenCum.frm

    Form=frmsaveDBS.frm

    Form=FrmFrondPage.frm

    Form=frmRevise.frm

    Form=frmOpenRank.frm

    Form=frmDataTable2.frm

    Form=FrmFrondPage1.frm

    Form=frmDic.frm

    Form=frmEditWells.frm

    Form=frmMakedatEvent.frm

    Form=frmMakedatWell.frm

    Form=frmSortingW.frm

    Form=frmMakedatHeader.frm

    Form=frmSorting.frm

    Form=frmOpenTable_RC.frm

    Form=frmOpenDBS.frm

    Form=frmOpenIPS.frm

    Form=frmCor.frm

    Form=frmReviseCor.frm

    Form=frmChartShow.frm

    Form=frmSelectEvent.frm

    Form=frmOpenTable.frm

    Form=frmSelectWell.frm

    Form=frmOpenMake.frm

    Form=frmWells2.frm

    Form=frmwellPlot2.frm

    Form=frmVar.frm

    Form=frmDenRank.frm

    Form=frmscatter.frm

    Form=frmDialog.frm

    Form=dlgsetpath.frm

    Form=frmsumTable.frm

    Form=frmChartTable.frm

    Form=frmOpenOutputFiles.frm

    Form=frmCASCParameter.frm

    Form=frmDepthDem.frm

    Form=frmDen.frm

    Form=frmOpenDem.frm

    Form=frmOpenDen.frm

    Form=frmFirstOrderDepthDiff.frm

    Form=frmOpenRan1.frm

    Form=frmOpenDI1.frm

    Form=frmDiffHistogram.frm

    Form=frmTransDepthDiffQQplot.frm

    Form=frmOpenDF2.frm

    Form=frmDepthDiffQQplot.frm

    Form=frmOpenDF1.frm

    Form=ToolBar.frm

    Form=frmRan3.frm

    Form=frmRan4.frm

    Form=frmRan1.frm

    Form=frmRan2.frm

    Form=frmscatterSC2.frm

    Form=frmVar2.frm

    Form=frmscatterDE.frm

    Form=frmscatterDE2.frm

    Form=frmwellPlot.frm

    Form=frmWells.frm

    Form=frmDataTable.frm

    Form=frmCum.frm

    Module=Module3; Module3.bas

    Module=Module1; apifunc.bas

    IconForm="MDIFrmCascRasc"

    Startup="MDIFrmCascRasc"

    HelpFile="RASCW.hlp"
    where Startup form is "MDIFrmCascRasc" and system help file is "RASCW.hlp".

    A compiled excutable file "rascwV20.exe" is included in this package.

3. An ActiveX Control named "Olectra Chart 6.0" developed by APEX Software Company is applied to display charts and diagrams for RASC/CASC results. So the developer should buy the licence from APEX Software Company before his programming.

4. Free download and open access software "QSCreator.jar" and "editpadlite.exe" are included in the package for data exchange and editing.

5. Dataset "14cen.dat" and "27cen.dat" are included in the package for debugging and testing.

6. Current default project file directory is "D:\RASCWV20". If you have install the package in different directory, please change the project directory in Visual Basic project parameters.

Thank you and good luck.
