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
; Fred Just
; René Rhéaume (rener@mediom.qc.ca)
; 
; ***** END LICENSE BLOCK *****

; <!DOCTYPE Nullsoft-SuperPiMPScript 1.91>
; Script NSIS pour c2iExplorer 1.60.${Revision} VB5/6 Beta

!ifndef VersionVB
!define VersionVB 5
!endif
;!define VB6
!define Revision 126

Name "c2iExplorer 1.60.${Revision} Beta pour VB${VersionVB}"
Caption "c2iExplorer 1.60.${Revision} Beta pour VB${VersionVB}"
CRCCheck On
LicenseText "Lisez et approuvez la licence ci-dessous avant de continuer l'installation."
LicenseData "C:\Mes documents\c2iexplorer\Source\Modifications René Rhéaume.txt"
ComponentText "Sélectionnez les composants de c2iExplorer que vous souhaitez installer."
DirText "Indiquez dans quel dossier est installé Visual Basic ${VersionVB}.0 sur votre ordinateur."
UninstallText "Ceci supprimera c2iExplorer de votre ordinateur. Cliquez sur «Désinstaller» pour désinstaller ou sur «Annuler» pour quitter."
!packhdr "temp.dat" "c:\progra~1\console\upx120w\upx.exe -q --best --compress-icons=1 temp.dat"
;UninstallExeName uninst-c2iexplorer.exe
OutFile "C:\Mes documents\c2iexplorer\c2iexplorer-VB${VersionVB}.exe"
Icon "C:\Mes documents\c2iexplorer\Source\c2iExplorer.ico"
EnabledBitmap "C:\Mes documents\c2iexplorer\Source\checked16.bmp"
DisabledBitmap "C:\Mes documents\c2iexplorer\Source\unchecked16.bmp"
!ifdef VB6
InstallDir "$PROGRAMFILES\Microsoft Visual Studio\VB98"
!else
InstallDir "$PROGRAMFILES\DevStudio\Vb"
!endif
InstType "Standard"
InstType "Complète"
InstType /NOCUSTOM

Section "c2iExplorer pour VB${VersionVB}"
	SectionIn 12
	SetCompress Auto
	; SetOverwrite ifnewer
	; Placé en commentaire à cause d'un règlement de sourceforge.net
	; SetOutPath $SYSDIR
	; UnRegDLL $SYSDIR\mscomctl.ocx
	; File "C:\Windows\Bureau\mdinote-setup\mscomctl.ocx"
	; RegDLL $SYSDIR\mscomctl.ocx
	SetOverwrite On
	SetOutPath "$INSTDIR\Wizards"
	File "C:\Mes documents\c2iexplorer\Source\FredJust modifications.txt"
	File "C:\Mes documents\c2iexplorer\Source\Modifications René Rhéaume.txt"
	File "C:\Mes documents\c2iexplorer\ReleaseVB${VersionVB}\c2iexplorer.dll"
	File "C:\Mes documents\c2iexplorer\ReleaseVB${VersionVB}\c2iexplorer.lib"
	File "C:\Mes documents\c2iexplorer\ReleaseVB${VersionVB}\c2iexplorer.exp"
	File "C:\Mes documents\c2iexplorer\ReleaseVB${VersionVB}\UD*.vbd"
	SetOutPath "$INSTDIR\Wizards\Data"
	File "C:\Mes documents\c2iexplorer\Data\*.*"
	SetOutPath "$INSTDIR\Wizards\Html"
	File "C:\Mes documents\c2iexplorer\Html\*.*"
	SetOutPath "$INSTDIR\Wizards\Html\Img"
	File "C:\Mes documents\c2iexplorer\Html\Img\*.*"
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
	WriteINIStr "$WINDIR\vbaddin.ini" "Add-Ins32" "c2iexplorer.Connectc2iExplorer" "1"
	RegDLL "$INSTDIR\Wizards\c2iexplorer.dll"
SectionEnd

Section "Code source VB 5/6"
	SectionIn 2
	SetCompress Auto
	SetOverwrite On
	SetOutPath "$INSTDIR\Wizards\c2iexplorer-source"
	File "C:\Mes documents\c2iexplorer\Source\*.*"
	Exec "$WINDIR\explorer.exe /n,$INSTDIR\Wizards\c2iexplorer-source"
SectionEnd

Section -PostInstall
	WriteRegStr HKEY_LOCAL_MACHINE "Software\Microsoft\Windows\CurrentVersion\Uninstall\c2iexplorerVB${VersionVB}" "DisplayName" "c2i Explorer pour VB${VersionVB}"
	WriteRegStr HKEY_LOCAL_MACHINE "Software\Microsoft\Windows\CurrentVersion\Uninstall\c2iexplorerVB${VersionVB}" "UninstallString" "$INSTDIR\uninst-c2iexplorer.exe"
	Delete "$INSTDIR\uninst-c2iexplorer.exe"
	WriteUninstaller uninst-c2iexplorer.exe
	BringToFront
SectionEnd

Section Uninstall
	CopyFiles "$INSTDIR\Wizards\DATA\c2iExplorer.ini" "$DESKTOP" 1
	DeleteRegKey HKEY_LOCAL_MACHINE "Software\Microsoft\Windows\CurrentVersion\Uninstall\c2iexplorerVB${VersionVB}"
	DeleteRegKey HKEY_CURRENT_USER "Software\VB and VBA Program Settings\MesCompléments\c2iExplorer"
WriteINIStr "$WINDIR\vbaddin.ini" "Add-Ins32" "c2iexplorer.Connectc2iExplorer" "0"
	UnRegDLL "$INSTDIR\Wizards\c2iexplorer.dll"
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
	Delete "$INSTDIR\Wizards\HTML\c2i Description complète.htm"
	Delete "$INSTDIR\Wizards\HTML\c2i Liste des components avec descriptions.htm"
	Delete "$INSTDIR\Wizards\HTML\c2i Liste des components avec nbr lignes de code.htm"
	Delete "$INSTDIR\Wizards\HTML\c2i Liste des composants avec descriptions.htm"
	Delete "$INSTDIR\Wizards\HTML\c2i Liste des composants avec nbr lignes de code.htm"
	Delete "$INSTDIR\Wizards\HTML\c2i vierge.htm"
	Delete "$INSTDIR\Wizards\HTML\c2i xml.htm"
	Delete "$INSTDIR\Wizards\HTML\c2i xml.xml"
	Delete "$INSTDIR\Wizards\HTML\c2iEssai.htm"
	Delete "$INSTDIR\Wizards\HTML\c2iMod1.htm"
	Delete "$INSTDIR\Wizards\HTML\c2iMod2.htm"
	Delete "$INSTDIR\Wizards\HTML\c2iMod3.htm"
	Delete "$INSTDIR\Wizards\HTML\c2iMod4.htm"
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
	Delete "$INSTDIR\Wizards\c2iexplorer-source\cDlgCom.cls"
	Delete "$INSTDIR\Wizards\c2iexplorer-source\images.zip"
	RmDir "$INSTDIR\Wizards\Data"
	RmDir "$INSTDIR\Wizards\Html\Img"
	RmDir "$INSTDIR\Wizards\Html"
	RmDir "$INSTDIR\Wizards\c2iexplorer-source"
	Delete "$INSTDIR\uninst-c2iexplorer.exe"
	MessageBox "MB_OK|MB_ICONINFORMATION" "c2iExplorer est maintenant désinstallé. Cependant, le fichier de préférences «c2iExplorer.ini» a été copié sur le bureau à des fins d'archivage et de restauration de préférences."
SectionEnd
