; ***** BEGIN LICENSE BLOCK *****
; Version: MPL 1.1
;
; The contents of this file are subject to the Mozilla Public License Version
; 1.1 (the "License"); you may not use this file except in compliance with
; the License. You may obtain a copy of the License at
; http://www.mozilla.org/MPL/
;
; Software distributed under the License is distributed on an "AS IS" basis,
; WITHOUT WARRANTY OF ANY KIND, either express or implied. See the License
; for the specific language governing rights and limitations under the
; License.
;
; The Original Code is c2iExplorer (for Visual Basic 6).
;
; The Initial Developer of the Original Code is
; c2i - Richard Clark Ingénierie Informatique (www.c2i.fr).
; Portions created by the Initial Developer are Copyright (C) 2000
; the Initial Developer. All Rights Reserved.
;
; Contributor(s):
; René Rhéaume (rener@mediom.qc.ca)
;
; ***** END LICENSE BLOCK *****

!ifdef VB6
  !define VersionVB 6
!endif
!ifndef VersionVB
  !define VersionVB 5
!endif
!define Revision 157
!define CheminBase "I:\rene\Visual Basic\c2iExplorer"
!define UninstRegKey "Software\Microsoft\Windows\CurrentVersion\Uninstall\c2iexplorerVB${VersionVB}"
!define NomApp "c2iExplorer"
!define Titre "${NomApp} 1.60.${Revision} for VB${VersionVB}"
!define UninstProg "uninst-c2iexplorer.exe"
!define en ;installation en anglais
!include "E:\Progra~1\NSIS\nsisconf.nsi" ;nécessaire si appelé de l'invite de commandes

!macro Removec2iExRegKeys
  DeleteRegKey HKEY_CLASSES_ROOT "c2iexplorer.cElement"
  DeleteRegKey HKEY_CLASSES_ROOT "c2iexplorer.cElements"
  DeleteRegKey HKEY_CLASSES_ROOT "c2iexplorer.cExplorer"
  DeleteRegKey HKEY_CLASSES_ROOT "c2iexplorer.Connectc2iExplorer"
  DeleteRegKey HKEY_CLASSES_ROOT "c2iexplorer.UDAddCode"
  DeleteRegKey HKEY_CLASSES_ROOT "c2iexplorer.UDBiblio"
  DeleteRegKey HKEY_CLASSES_ROOT "c2iexplorer.UDClassViewer"
  DeleteRegKey HKEY_CLASSES_ROOT "c2iexplorer.UDCompteur"
  DeleteRegKey HKEY_CLASSES_ROOT "c2iexplorer.UDExplorer"
  DeleteRegKey HKEY_CLASSES_ROOT "c2iexplorer.UDMsgBox"
!macroend

Name "${Titre}"
Caption "${Titre}"
OutFile "${CheminBase}\SourceForge\Fichiers\c2iExplorer-1.60.${Revision}-vb${VersionVB}.i586.exe"
CRCCheck On

Icon "${CheminBase}\Source\c2iExplorer.ico"
EnabledBitmap "${CheminBase}\Source\checked16.bmp"
DisabledBitmap "${CheminBase}\Source\unchecked16.bmp"

LicenseText "Read this before continuing installation."
LicenseData "${CheminBase}\Source\Modifications René Rhéaume.txt"
ComponentText "Select ${NomApp} components you want to install."
DirText "Specify the folder where Visual Basic ${VersionVB}.0 is installed on your computer."
UninstallText "This will delete ${NomApp} from your computer. Click on 'Uninstall' to uninstall or on 'Cancel' to exit."

!ifdef VB6
  InstallDir "$PROGRAMFILES\Microsoft Visual Studio\VB98"
!else
  InstallDir "$PROGRAMFILES\DevStudio\Vb"
!endif
InstallDirRegKey HKEY_LOCAL_MACHINE "${UninstRegKey}" "UninstallString"
InstType "Standard French"
InstType "Standard English"
InstType "Full"

Section "${NomApp} for VB${VersionVB}"
  SectionIn 1,2,3
  SetCompress Auto
  ; SetOverwrite ifnewer
  ; Placé en commentaire à cause d'un règlement de sourceforge.net
  ; SetOutPath $SYSDIR
  ; UnRegDLL $SYSDIR\mscomctl.ocx
  ; File "C:\Windows\Bureau\mdinote-setup\mscomctl.ocx"
  ; RegDLL $SYSDIR\mscomctl.ocx
  SetOverwrite On
  SetOutPath "$INSTDIR\Wizards"
  File "${CheminBase}\Source\FredJust modifications.txt"
  File "${CheminBase}\Source\Modifications René Rhéaume.txt"
  File "${CheminBase}\ReleaseVB${VersionVB}\c2iexplorer.dll"
  SetOutPath "$INSTDIR\Wizards\Data"
  File "${CheminBase}\Data\*.*"
  SetOutPath "$INSTDIR\Wizards\Html"
  File "${CheminBase}\Html\[]c2i*.*m*"
  File "${CheminBase}\Html\[fr]c2i*.htm"
  File "${CheminBase}\Html\prev.htm"
  SetOutPath "$INSTDIR\Wizards\Html\Img"
  File "${CheminBase}\Html\Img\*.*"
  SetOutPath "$INSTDIR\Wizards\Lang"
  File "${CheminBase}\Lang\*.lng"
  !insertmacro Removec2iExRegKeys
  WriteINIStr "$WINDIR\vbaddin.ini" "Add-Ins32" "c2iexplorer.Connectc2iExplorer" "1"
  RegDLL "$INSTDIR\Wizards\c2iexplorer.dll"
SectionEnd

SectionDivider "English translation"

Section "English user interface"
  SectionIn 2,3
  WriteINIStr "$INSTDIR\Wizards\DATA\c2iExplorer.ini" "General" "Language" "c2iExplorer-en.lng"
SectionEnd

Section "English report templates"
  SectionIn 2,3
  SetOutPath "$INSTDIR\Wizards\Html"
  File "${CheminBase}\Html\[en]c2i*.htm"
SectionEnd

SectionDivider "Programmer stuff"

Section "VB 5/6 source code"
  SectionIn 3
  SetOutPath "$INSTDIR\Wizards\c2iexplorer-source"
  ;Tri selon le type de fichier afin d'optimiser la compression
  File "${CheminBase}\Source\*.dob"
  File "${CheminBase}\Source\*.frm"
  File "${CheminBase}\Source\*.cls"
  File "${CheminBase}\Source\*.bas"
  File "${CheminBase}\Source\*.ctl"
  File "${CheminBase}\Source\*.vbp"
  File "${CheminBase}\Source\*.rtf"
  File "${CheminBase}\Source\*.htm"
  File "${CheminBase}\Source\*.nsi"
  File "${CheminBase}\Source\*.txt"
  File "${CheminBase}\Source\*.scc"
  File "${CheminBase}\Source\*.doc"
  File "${CheminBase}\Source\*.dox"
  File "${CheminBase}\Source\*.frx"
  File "${CheminBase}\Source\*.ico"
  File "${CheminBase}\Source\*.res"
  File "${CheminBase}\Source\*.ctx"
  File "${CheminBase}\Source\*.bmp"
  File "${CheminBase}\Source\*.gz"

  Exec "$WINDIR\explorer.exe /n,$INSTDIR\Wizards\c2iexplorer-source"
SectionEnd

Section "Localisation Dev Kit (LDK)"
  SectionIn 3
  SetOutPath "$INSTDIR\Wizards\LDK"
  File "${CheminBase}\Lang\*.lng"
  File "${CheminBase}\LDK\*.html"
  File "${CheminBase}\LDK\*.css"
  Exec "$WINDIR\explorer.exe /n,$INSTDIR\Wizards\LDK"
SectionEnd

Section -PostInstall
  WriteRegStr HKEY_LOCAL_MACHINE "${UninstRegKey}" "DisplayName" "${Titre}"
  WriteRegStr HKEY_LOCAL_MACHINE "${UninstRegKey}" "UninstallString" "$INSTDIR\${UninstProg}"
  Delete "$INSTDIR\${UninstProg}"
  WriteUninstaller ${UninstProg}
  BringToFront
SectionEnd

Section Uninstall
  CopyFiles "$INSTDIR\Wizards\DATA\c2iExplorer.ini" "$DESKTOP" 1
  DeleteRegKey HKEY_LOCAL_MACHINE "${UninstRegKey}"
  DeleteRegKey HKEY_CURRENT_USER "Software\VB and VBA Program Settings\c2iExplorer"
  DeleteRegKey HKEY_CURRENT_USER "Software\VB and VBA Program Settings\MesCompléments\c2iExplorer"
  WriteINIStr "$WINDIR\vbaddin.ini" "Add-Ins32" "c2iexplorer.Connectc2iExplorer" "0"
  UnRegDLL "$INSTDIR\Wizards\c2iexplorer.dll"
  !insertmacro Removec2iExRegKeys
  Delete "$INSTDIR\Wizards\UDMsgBox.vbd"
  Delete "$INSTDIR\Wizards\UDCompteur.vbd"
  Delete "$INSTDIR\Wizards\UDBiblio.vbd"
  Delete "$INSTDIR\Wizards\UDClassViewer.vbd"
  Delete "$INSTDIR\Wizards\UDAddCode.vbd"
  Delete "$INSTDIR\Wizards\UDExplorer.vbd"
  Delete "$INSTDIR\Wizards\c2iExplorer.c2i"
  Delete "$INSTDIR\Wizards\FredJust modifications.txt"
  Delete "$INSTDIR\Wizards\duree.txt"
  Delete "$INSTDIR\Wizards\Modifications René Rhéaume.txt"
  Delete "$INSTDIR\Wizards\c2iExplorer.dll"
  Delete "$INSTDIR\Wizards\c2iExplorer.lib"
  Delete "$INSTDIR\Wizards\c2iExplorer.exp"
  Delete "$INSTDIR\Wizards\DATA\duree.txt"
  Delete "$INSTDIR\Wizards\DATA\TYPES.TXT"
  Delete "$INSTDIR\Wizards\DATA\c2iExplorer.ini"
  Delete "$INSTDIR\Wizards\DATA\BASE.C2I"
  Delete "$INSTDIR\Wizards\DATA\C2I.MDB"
  Delete "$INSTDIR\Wizards\DATA\c2ifr.mdb"
  Delete "$INSTDIR\Wizards\DATA\codelib.dtd"
  Delete "$INSTDIR\Wizards\HTML\*c2i Description complète.htm"
  Delete "$INSTDIR\Wizards\HTML\*c2i Liste des components avec descriptions.htm"
  Delete "$INSTDIR\Wizards\HTML\*c2i Liste des components avec nbr lignes de code.htm"
  Delete "$INSTDIR\Wizards\HTML\*c2i Liste des composants avec descriptions.htm"
  Delete "$INSTDIR\Wizards\HTML\*c2i Liste des composants avec nbr lignes de code.htm"
  Delete "$INSTDIR\Wizards\HTML\*c2i vierge.htm"
  Delete "$INSTDIR\Wizards\HTML\c2i xml.htm"
  Delete "$INSTDIR\Wizards\HTML\*c2i xml.xml"
  Delete "$INSTDIR\Wizards\HTML\*c2iEssai.htm"
  Delete "$INSTDIR\Wizards\HTML\*c2iMod1.htm"
  Delete "$INSTDIR\Wizards\HTML\*c2iMod2.htm"
  Delete "$INSTDIR\Wizards\HTML\*c2iMod3.htm"
  Delete "$INSTDIR\Wizards\HTML\*c2iMod4.htm"
  Delete "$INSTDIR\Wizards\HTML\prev.htm"
  Delete "$INSTDIR\Wizards\HTML\IMG\ActiveXControl.gif"
  Delete "$INSTDIR\Wizards\HTML\IMG\ActiveXDesigner.gif"
  Delete "$INSTDIR\Wizards\HTML\IMG\ActiveXDll.gif"
  Delete "$INSTDIR\Wizards\HTML\IMG\ActiveXExe.gif"
  Delete "$INSTDIR\Wizards\HTML\IMG\API.GIF"
  Delete "$INSTDIR\Wizards\HTML\IMG\ClassModule.gif"
  Delete "$INSTDIR\Wizards\HTML\IMG\constante.gif"
  Delete "$INSTDIR\Wizards\HTML\IMG\DocObject.gif"
  Delete "$INSTDIR\Wizards\HTML\IMG\evenement.gif"
  Delete "$INSTDIR\Wizards\HTML\IMG\Iconnu.gif"
  Delete "$INSTDIR\Wizards\HTML\IMG\LOGO.GIF"
  Delete "$INSTDIR\Wizards\HTML\IMG\METHODE.GIF"
  Delete "$INSTDIR\Wizards\HTML\IMG\MSForm.gif"
  Delete "$INSTDIR\Wizards\HTML\IMG\PropPage.gif"
  Delete "$INSTDIR\Wizards\HTML\IMG\propriete.gif"
  Delete "$INSTDIR\Wizards\HTML\IMG\ResFile.gif"
  Delete "$INSTDIR\Wizards\HTML\IMG\StdModule.gif"
  Delete "$INSTDIR\Wizards\HTML\IMG\UserControl.gif"
  Delete "$INSTDIR\Wizards\HTML\IMG\VARIABLE.GIF"
  Delete "$INSTDIR\Wizards\HTML\IMG\VBChildForm.gif"
  Delete "$INSTDIR\Wizards\HTML\IMG\VBForm.gif"
  Delete "$INSTDIR\Wizards\HTML\IMG\VBMDIForm.gif"
  Delete "$INSTDIR\Wizards\HTML\IMG\VBproject.gif"
  Delete "$INSTDIR\Wizards\Lang\c2iExplorer-fr.lng"
  Delete "$INSTDIR\Wizards\Lang\c2iExplorer-en.lng"
  Delete "$INSTDIR\Wizards\LDK\c2iExplorer-fr.lng"
  Delete "$INSTDIR\Wizards\LDK\c2iExplorer-en.lng"
  Delete "$INSTDIR\Wizards\LDK\c2iexplorer_ldk.css"
  Delete "$INSTDIR\Wizards\LDK\c2iexplorer_ldk_fr.html"
  Delete "$INSTDIR\Wizards\LDK\c2iexplorer_ldk_en.html"
  Delete "$INSTDIR\Wizards\LDK\gnu_fdl_en.html"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\boilerplate-c2iExplorer.txt"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\FredJust modifications.txt"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\Modifications René Rhéaume.txt"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\Gestfich.bas"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\UDAddCode.dob"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\MClasses.bas"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\MExtraitIcone.bas"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\MPrint.bas"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\checked16.bmp"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\unchecked16.bmp"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\cElement.cls"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\cElements.cls"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\cExplorer.cls"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\Connectc2iExplorer.cls"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\c2iContainerInt.ctl"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\c2iContainerInt.ctx"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\UDBiblio.dob"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\UDClassViewer.dob"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\UDCompteur.dob"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\UDExplorer.dob"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\UDMsgBox.dob"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\UDAddCode.dox"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\UDBiblio.dox"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\UDClassViewer.dox"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\UDCompteur.dox"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\UDExplorer.dox"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\UDMsgBox.dox"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\frmBiblio.frm"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\frmDuree.frm"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\frmDuree.frx"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\frmPrint.frm"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\frmBiblio.frx"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\frmPrint.frx"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\MPL-1.1.html"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\classgenerator.ico"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\c2iExplorer.ico"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\c2iExplorer-VB5.nsi"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\c2iExplorer-VB6.nsi"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\c2iExplorer.nsi"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\c2iExplorer.RES"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\MSSCCPRJ.SCC"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\c2iExplorer-vb5.vbp"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\c2iExplorer-vb6.vbp"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\c2iExplorer-vb?.vbw"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\archive.zip"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\archive.tar.gz"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\cDlgCom.cls"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\images.zip"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\RapportModificationsRR.*"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\InsertCode.txt"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\InsertCode.doc"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\MiseAJourCode.doc"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\SauvegarderBibliotheque.doc"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\frmErreurINI.frm"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\frmErreurINI.frx"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\frmOptions.fr?"
  Delete "$INSTDIR\Wizards\c2iexplorer-source\c2iExplorer-gif.tar.gz"
  RmDir "$INSTDIR\Wizards\Data"
  RmDir "$INSTDIR\Wizards\Html\Img"
  RmDir "$INSTDIR\Wizards\Html"
  RmDir "$INSTDIR\Wizards\Lang"
  RmDir "$INSTDIR\Wizards\LDK"
  RmDir "$INSTDIR\Wizards\c2iexplorer-source"
  Delete "$INSTDIR\${UninstProg}"
  MessageBox "MB_OK|MB_ICONINFORMATION" "${NomApp} is now uninstalled. However, settings file «c2iExplorer.ini» has been copied onto the desktop for archiving and settings restoration purposes."
SectionEnd