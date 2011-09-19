using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SwDocumentMgr;


namespace GetReferences
{
    class Program
    {
        static SwDMApplication3 dmDocManager;
        static SwDMClassFactory dmClassFact;

        //  You must obtain the key directly from SolidWorks API division
        const string SolidWorksDocumentManagerKey = "<Your License Key Here>";

        static void Main(string[] args)
        {
            
            //Takes Care of input checking and input parsing
            string docPath;
            bool quietMode;
            switch (args.Length)
            {
                case 1:
                    quietMode = false;
                    docPath = args[0];
                    break;
                case 2:
                    if (args[0] != "/q")
                    {
                        quietMode = false;
                        inputError(quietMode);
                        return;
                    }
                    quietMode = true;
                    docPath = args[1];
                    break;
                default:
                    quietMode = false;
                    inputError(quietMode);
                    return;
            }


            //Get Document Type
            SwDmDocumentType docType = setDocType(docPath);
            if (docType == SwDmDocumentType.swDmDocumentUnknown)
            {
                inputError(quietMode);
                return;
            }

            //Variable initialization
            SwDMDocument dmDoc;
            SwDmDocumentOpenError OpenError;

            ////Prerequisites
            dmClassFact = new SwDMClassFactory();
            dmDocManager = dmClassFact.GetApplication(SolidWorksDocumentManagerKey) as SwDMApplication3;


            //Open the Document       
            dmDoc = dmDocManager.GetDocument(docPath, docType, true, out OpenError) as SwDMDocument;

            //Check that a SolidWorks file is open
            if (dmDoc != null)
            {
                switch (docType)
                {
                    case SwDmDocumentType.swDmDocumentDrawing:
                        try
                        {
                            GetDrawingReferences(dmDoc);
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("\"" + dmDoc.FullName+ "\"\t\"" + "{0}" + "\tFile is internally damaged, .Net error occurred, or GetReferences.exe has a Bug." +"\"", e);
                            inputError(quietMode);
                        }
                        break;
                    case SwDmDocumentType.swDmDocumentPart:
                    case SwDmDocumentType.swDmDocumentAssembly:
                        try
                        {
                            GetPartOrAssemblyReferences(dmDoc);
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("\"" + dmDoc.FullName + "\"\t\"" + "{0} Exception caught.  File is internally damaged, .Net error occurred, or GetReferences.exe has a Bug." + "\"", e);
                            inputError(quietMode);
                        }
                        break;
                    default:
                        inputError(quietMode);
                        return;
                }
                dmDoc.CloseDoc();
            }
            else
            {
                switch (OpenError)
                {
                    case SwDmDocumentOpenError.swDmDocumentOpenErrorFail:
                        Console.WriteLine(docPath + "\tFile failed to open; reasons could be related to permissions, the file is in use, or the file is corrupted.");
                        inputError(quietMode);
                        break;
                    case SwDmDocumentOpenError.swDmDocumentOpenErrorFileNotFound:
                        Console.WriteLine(docPath + "\tFile not found");
                        inputError(quietMode);
                        break;
                    case SwDmDocumentOpenError.swDmDocumentOpenErrorNonSW:
                        Console.Write(docPath + "\tNon-SolidWorks file was opened");
                        inputError(quietMode);
                        break;
                    default:
                        Console.WriteLine(docPath + "\tAn unknown errror occurred.  Something is wrong with the code of \"GetReferences\"");
                        inputError(quietMode);
                        break;
                }

            }
        }

        static void inputError(bool quietMode)
        {
            if (quietMode)
                return;

            Console.WriteLine(@"
Syntax 
    [option] [ParentFilePath]
Output
    ""ParentFilePath""  ""ParentConfig""    ""ChildFilePath""   ""ChildConfig""

Only one Filename is accepted. If the path has spaces use quotes around it.
Note that the file must have one of the following file extensions: .sldasm,
.slddrw, .sldprt, .asm, .drw, or .prt.

Options
    /q      Quiet mode.  Suppresses the current message.  It does
            not suppress the one line error messages related to problems
            opening SolidWorks Files.  Quiet mode is useful for batch files
            when you are directing the output to a file.  The main error 
            message is suppressed but you are still informed about problems 
            opening files.

Version 2011-Sept-19 13:21
Written and Maintained by Jason Nicholson
http://github.com/jasonnicholson/GetReferences");
        }




        static SwDmDocumentType setDocType(string docPath)
        {
            string fileExtension = System.IO.Path.GetExtension(docPath);

            //Notice no break statement is needed because I used return to get out of the switch statement.
            switch (fileExtension.ToUpper())
            {
                case ".SLDASM":
                case ".ASM":
                    return SwDmDocumentType.swDmDocumentAssembly;
                case ".SLDDRW":
                case ".DRW":
                    return SwDmDocumentType.swDmDocumentDrawing;
                case ".SLDPRT":
                case ".PRT":
                    return SwDmDocumentType.swDmDocumentPart;
                default:
                    return SwDmDocumentType.swDmDocumentUnknown;
            }

        }




        static void GetDrawingReferences(SwDMDocument dmDoc)
        {
            SwDMDocument13 dmDrw = dmDoc as SwDMDocument13;
            SwDMSearchOption dmSearchOptions = dmDocManager.GetSearchOptionObject();
            object vBrokenRefs;
            object vIsVirtual;
            object vTimeStamp;

            string[] reference = dmDrw.GetAllExternalReferences4(dmSearchOptions, out vBrokenRefs, out vIsVirtual, out vTimeStamp);
            //SwDmReferenceStatus[] dmBrokenRefs = vBrokenRefs as SwDmReferenceStatus[];
            
            //check for no references
            if (reference == null)
            {
                Console.WriteLine("\"" + dmDoc.FullName + "\"\t");
                return;
            }

            SwDMDocument10 dmDrw2 = dmDoc as SwDMDocument10;
            object[] dmDrwViews = dmDrw.GetViews() as object[];
            SwDMView view;
            string[] referenceFileNames = new string[dmDrwViews.Length];
            string[] configurations = new string[dmDrwViews.Length];
            string[] referenceAndConfigList = new string[dmDrwViews.Length];
            for (int i = 0; i < dmDrwViews.Length; i++)
            {
                view = dmDrwViews[i] as SwDMView;
                referenceFileNames[i] = view.ReferencedDocument;
                configurations[i] = view.ReferencedConfiguration;
                for (int j = 0; j < reference.Length; j++)
                {
                    if (System.IO.Path.GetFileName(reference[j]).ToUpper() == referenceFileNames[i].ToUpper())
                    {
                        referenceAndConfigList[i] = "\"" + dmDoc.FullName + "\"\t\t\"" + reference[j] + "\"\t\"" + configurations[i] + "\""; //+ Enum.GetName(SwDmReferenceStatus,dmBrokenRefs[j]) ;
                        break;
                    }
                }

            }

            referenceAndConfigList = referenceAndConfigList.Distinct().ToArray();
            foreach (string outputLine in referenceAndConfigList)
            {
                Console.WriteLine(outputLine);
            }
        }




        static void GetPartOrAssemblyReferences(SwDMDocument dmDoc)
        {
            SwDMDocument15 dmPartOrAssembly = dmDoc as SwDMDocument15;
            SwDMExternalReferenceOption dmExternalReferencesOption = dmDocManager.GetExternalReferenceOptionObject();
            SwDMSearchOption dmSearchOptions = dmDocManager.GetSearchOptionObject();
            dmExternalReferencesOption.SearchOption = dmSearchOptions;
            dmExternalReferencesOption.NeedSuppress = true;
            int numberOfExternalReferences;
                        
            string[] configurationNames = dmPartOrAssembly.ConfigurationManager.GetConfigurationNames();
            foreach (string parentConfiguration in configurationNames)
            {
                dmExternalReferencesOption.Configuration = parentConfiguration;
                numberOfExternalReferences = dmPartOrAssembly.GetExternalFeatureReferences(ref dmExternalReferencesOption);
                
                //check for no references
                if (numberOfExternalReferences == 0)
                {
                    Console.WriteLine("\"" + dmPartOrAssembly.FullName + "\"\t\"" + parentConfiguration + "\"");
                    break;
                }
                string[] referenceAndConfigList = new string[numberOfExternalReferences];
                for (int i = 0; i < numberOfExternalReferences; i++)
                {
                    referenceAndConfigList[i] = "\"" + dmPartOrAssembly.FullName + "\"\t\"" + parentConfiguration + "\"\t\"" + dmExternalReferencesOption.ExternalReferences[i] + "\"\t\"" + dmExternalReferencesOption.ReferencedConfigurations[i] + "\"";
                }
                referenceAndConfigList = referenceAndConfigList.Distinct().ToArray();
                foreach (string outputLine in referenceAndConfigList)
                {
                    Console.WriteLine(outputLine);
                }
            }
        }
    }
}