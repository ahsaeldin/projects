Imports System.Resources

Imports System
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Runtime.CompilerServices

' General Information about an assembly is controlled through the following 
' set of attributes. Change these attribute values to modify the information
' associated with an assembly.

' Review the values of the assembly attributes

<Assembly: AssemblyTitle("Cacadu")> 
<Assembly: AssemblyDescription("Create scheduled tasks to be executed according to specific times or specific triggers.")> 
<Assembly: AssemblyCompany("Conderella Soft www.conderella.com")> 
<Assembly: AssemblyProduct("Cacadu")> 
<Assembly: AssemblyCopyright("Copyright © 2012 , Conderella. All rights reserved.")> 
<Assembly: AssemblyTrademark("")> 

<Assembly: ComVisible(False)> 

'The following GUID is for the ID of the typelib if this project is exposed to COM
<Assembly: Guid("f8aba858-1a9e-4b1e-a735-067a78918811")> 

' Version information for an assembly consists of the following four values:
'
'      Major Version
'      Minor Version 
'      Build Number
'      Revision
'
'First number is the major version.
'Second number is the minor version.
'Third number is the day of new major or minor version.
'Last number is the revision number of a fix major/minor/date version.
'Note:Last number is related to the third only, and for every change in major/minor, the last number return to zero.
'Note:Third number is related to major/minor changes.
<Assembly: AssemblyVersion("1.0.12927.1")> 
<Assembly: AssemblyFileVersion("1.0.12927.1")> 
<Assembly: NeutralResourcesLanguageAttribute("")> 

'No need as I'm using nunit right now. NOOOOOOOOO NEEEEEEEED
'#If DEBUG Then
'<Assembly: System.Runtime.CompilerServices.InternalsVisibleTo("Cacadu Unit Test Project")> 
'#End If

#Region "The only way to merge <remove all other projects from Eazfusctor first>"
'<Assembly: Obfuscation(Feature:="ilmerge custom parameters: /closed", Exclude:=False)> 
'<Assembly: Obfuscation(Feature:="merge with Balora.dll", Exclude:=False)> 
'<Assembly: Obfuscation(Feature:="merge with Cacadore.dll", Exclude:=False)> 
'<Assembly: Obfuscation(Feature:="merge with TecDAL.dll", Exclude:=False)> 
'<Assembly: Obfuscation(Feature:="merge with Helpers.dll", Exclude:=False)> 
#End Region

<Assembly: Obfuscation(Feature:="encrypt symbol names with password AS2#$SDFGYUI789dfgwe$%^#$%GHFSADFGWQEsdfgyu4%^$%@#$Dghsdfg*@#GHAGdrfga5q6345HOR^&U#B$%^WYV$%^BW%$&SRYNR%^YADSFGWQEfjrtgQWEQW$%KJL:5*4%8(9#$6&3@2&0%LDTYUDdfghsd5J4FG2JI8j7H6j3JH1y4K87J8675&*g&gher55$8668354259%$SD54FG4sdgdfAge", Exclude:=False)> 
<Assembly: Obfuscation(Feature:="code control flow obfuscation", Exclude:=False)> 
<Assembly: Obfuscation(Feature:="Apply to *.netprotector.* when internal: renaming", Exclude:=True, ApplyToMembers:=True)> 
<Assembly: Obfuscation(Feature:="Apply to *.IntelliProtectorService.* when internal or enum: renaming", Exclude:=True, ApplyToMembers:=True)> 


