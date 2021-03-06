﻿Imports System
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Security

' Allgemeine Informationen über eine Assembly werden über folgende 
' Attribute gesteuert. Ändern Sie diese Attributwerte, um die Informationen zu ändern,
' die einer Assembly zugeordnet sind.

' Werte der Assemblyattribute überprüfen

<Assembly: AssemblyTitle("wpContacts Session Outlook Add-In")>
<Assembly: AssemblyDescription("Outlook Add-In zum Syncronisieren von Kontakten mit der Session-Datenbank")>
<Assembly: AssemblyCompany("WebPoint Internet Solutions Waller Alexander")>
<Assembly: AssemblyProduct("wpContacts Session Outlook Add-In")>
<Assembly: AssemblyCopyright("Copyright ©  2021 Alexander Waller GNU AGPL")>
<Assembly: AssemblyTrademark("")> 

' Durch Festlegen von "ComVisible" auf "false" werden die Typen in dieser Assembly unsichtbar 
' für COM-Komponenten. Wenn Sie auf einen Typ in dieser Assembly von 
' COM aus zugreifen müssen, legen Sie das ComVisible-Attribut für diesen Typ auf "true" fest.
<Assembly: ComVisible(False)>

'Die folgende GUID bestimmt die ID der Typbibliothek, wenn dieses Projekt für COM verfügbar gemacht wird.
<Assembly: Guid("7f9eab0c-fcd9-47d5-90ed-b53fecf30ac3")>

' Versionsinformationen für eine Assembly bestehen aus den folgenden vier Werten:
'
'      Hauptversion
'      Nebenversion 
'      Buildnummer
'      Revision
'
' Sie können alle Werte angeben oder die standardmäßigen Build- und Revisionsnummern 
' übernehmen, indem Sie "*" eingeben:
' <Assembly: AssemblyVersion("1.0.*")> 

'### https://stackoverflow.com/questions/356543/can-i-automatically-increment-the-file-build-version-when-using-visual-studio
<Assembly: AssemblyVersion("1.6.*")>
'<Assembly: AssemblyFileVersion("1.5.0.0")>
<Assembly: log4net.Config.XmlConfigurator(ConfigFile:="config.log4net", Watch:=True)>

Friend Module DesignTimeConstants
    Public Const RibbonTypeSerializer As String = "Microsoft.VisualStudio.Tools.Office.Ribbon.Serialization.RibbonTypeCodeDomSerializer, Microsoft.VisualStudio.Tools.Office.Designer, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
    Public Const RibbonBaseTypeSerializer As String = "System.ComponentModel.Design.Serialization.TypeCodeDomSerializer, System.Design"
    Public Const RibbonDesigner As String = "Microsoft.VisualStudio.Tools.Office.Ribbon.Design.RibbonDesigner, Microsoft.VisualStudio.Tools.Office.Designer, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
End Module
