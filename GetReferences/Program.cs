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

        //  You must obtain this directly from SolidWorks API division
        static const string SolidWorksDocumentManagerKey ="<Your License Key>";

        static void Main(string[] args)
        {
            //Check that only only one file argument is given.
            if ((args.Length != 1))
            {
                Console.WriteLine("Only one Filename is accepted. Make sure if the path has spaces you use quotes around it.");
                return;
            }
            string docPath = args[0];

            //string docPath = @"C:\PracticeFiles\00688.DRW";
            //string docPath = @"C:\PracticeFiles\02967.SLDDRW";
            //string docPath = @"C:\PracticeFiles\02967.SLDPRT";
            //string docPath = @"C:\PracticeFiles\03000.SLDASM";
            //string docPath = @"C:\PracticeFiles\2-8255P13.PRT";
            //string docPath = @"C:\PracticeFiles\3-5681.PRT";
            //string docPath = @"C:\PracticeFiles\3x1DellBlindmate.asm";
            //string docPath = @"C:\PracticeFiles\Draw1.slddrw";
                        
            //Variable initialization
            SwDMDocument dmDoc;
            SwDmDocumentOpenError OpenError;

            ////Prerequisites
            dmClassFact = new SwDMClassFactory();
            dmDocManager = dmClassFact.GetApplication(SolidWorksDocumentManagerKey) as SwDMApplication3;
            
            //Get Document Type
            SwDmDocumentType docType = new SwDmDocumentType();
            docType = setDocType(docPath);
    
            //Open the Document       
            dmDoc = dmDocManager.GetDocument(docPath, docType, true, out OpenError) as SwDMDocument;
            
            //Check that a SolidWorks file is open
            if (dmDoc != null)
            {
                //Get the references
                if (docType == SwDmDocumentType.swDmDocumentDrawing)
                {
                    GetDrawingReferences(dmDoc);
                }
                else if (docType == SwDmDocumentType.swDmDocumentPart || docType == SwDmDocumentType.swDmDocumentAssembly)
                {
                    GetPartOrAssemblyReferences(dmDoc);
                }
                dmDoc.CloseDoc();
            }
            else
            {
                Console.WriteLine("\"" + docPath + "\"\tUnable to open document. Error: " + OpenError);
            }
        }




        static SwDmDocumentType setDocType(string docPath)
        {
            string fileExtension = docPath.Substring((docPath.Length - 7), 4);

            if (fileExtension.ToUpper() == ".SLD")
            {
                string fileExtension2 = docPath.Substring((docPath.Length - 7), 7);

                if (fileExtension2.ToUpper() == ".SLDPRT")
                {
                    return SwDmDocumentType.swDmDocumentPart;
                }
                else if (fileExtension2.ToUpper() == ".SLDASM")
                {
                    return SwDmDocumentType.swDmDocumentAssembly;
                }
                else if (fileExtension2.ToUpper() == ".SLDDRW")
                {
                    return SwDmDocumentType.swDmDocumentDrawing;
                }
                else
                {
                    return SwDmDocumentType.swDmDocumentUnknown;
                }
            }
            else
            {
                string fileExtension2 = docPath.Substring((docPath.Length - 4), 4);

                if (fileExtension2.ToUpper() == ".PRT")
                {
                    return SwDmDocumentType.swDmDocumentPart;
                }
                else if (fileExtension2.ToUpper() == ".ASM")
                {
                    return SwDmDocumentType.swDmDocumentAssembly;
                }
                else if (fileExtension2.ToUpper() == ".DRW")
                {
                    return SwDmDocumentType.swDmDocumentDrawing;
                }
                else
                {
                    return SwDmDocumentType.swDmDocumentUnknown;
                }
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
            SwDmReferenceStatus[] dmBrokenRefs = vBrokenRefs as SwDmReferenceStatus[];

            SwDMDocument10 dmDrw2 = dmDoc as SwDMDocument10;
            object[] dmDrwViews = dmDrw.GetViews() as object[];
            SwDMView view;
            string[] referenceFileNames = new string[dmDrwViews.Length];
            string[] configurations = new string[dmDrwViews.Length];
            string[] referenceAndConfigList = new string[dmDrwViews.Length];
            for (int i=0; i < dmDrwViews.Length; i++)
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