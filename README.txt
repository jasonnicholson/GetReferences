Author: Jason Nicholson
Date Started: 2011-Sept-12
Last Date of README Update: 2011-Sept-18



Goal: 
This project is for creating a console application that obtains the references and their configurations of SolidWorks parts, assemblies, and drawings. 



Compiling Prerequisites: 
You must have the SolidWorks Document Manger Key and the SolidWorks Document Manager DLL to compile this code.  You may obtain a key from the SolidWorks API division.  Visit: http://www.solidworks.com/sw/support/apisupport.htm for more info.  You need a C# compiler with .NET 4.  I used the Microsoft Visual C# 2010 Express.


Binary Use Prerequisites: 
You must have .NET 4.  I am not sure if this will run with mono.  You must have the SolidWorks Document Manager DLL installed.  You may obtain it from SolidWorks.com in the download section.  You will have to sign in to the Customer Portal to get the download file.  From my understand of the license agreement I signed with SolidWorks for the Document Manager, I have the right to distribute the DLL and my code as long as I don't distribute my key.  Until I get the licensing completely figured out, I will not distribute the SwDocumentMgr.dll with my project.  You may obtain Document Manager instructions from the SolidWorks API Help "Getting Started" section.  At the time of writing this, here is the address of the "Getting Started" page of the Document Manager: http://help.solidworks.com/2011/English/api/swdocmgrapi/SolidWorks.Interop.swdocumentmgr_GettingStartedSWDocMgrAPI.html?id=a6e51b17163d4174b8c2c25c0f0afdac#ApplicationBasics



Where to get the Binary: 
its located in GetReferences\GetReferences\bin\Release\GetReferences.exe


Documentation:
An example use of GetReferences is located in: \Documentation\
If you call GetReferences.exe with no arguments, then the syntax usage displays.