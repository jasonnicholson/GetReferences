using System;
using System.Collections.Generic;
using SwDocumentMgr;
using System.IO;


namespace GetReferences
{
    class Program
    {
        static SwDMApplication3 dmDocManager;
        static SwDMClassFactory dmClassFact;

        //  You must obtain the key directly from SolidWorks API division
        const string SolidWorksDocumentManagerKey = "Insert Key Here";

        static void Main(string[] args)
        {
                //Takes Care of input checking and input parsing
                string docPath;
                bool quietMode;
                switch (args.Length)
                {
                    case 1:
                        quietMode = false;
                        if (args[0].Contains("*") || args[0].Contains("?"))
                        {
                            inputError(quietMode);
                            return;
                        }
                        docPath = Path.GetFullPath(args[0]);
                        break;
                    case 2:
                        if (args[0] != "/q")
                        {
                            quietMode = false;
                            inputError(quietMode);
                            return;
                        }
                        quietMode = true;
                        if (args[1].Contains("*") || args[1].Contains("?"))
                        {
                            inputError(quietMode);
                            return;
                        }
                        docPath = Path.GetFullPath(args[1]);
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
                SwDMDocument15 dmDoc;
                SwDmDocumentOpenError OpenError;

                ////Prerequisites
                dmClassFact = new SwDMClassFactory();
                dmDocManager = dmClassFact.GetApplication(SolidWorksDocumentManagerKey) as SwDMApplication3;


                //Open the Document       
                dmDoc = dmDocManager.GetDocument(docPath, docType, true, out OpenError) as SwDMDocument15;

                //Check that a SolidWorks file is open
                if (dmDoc != null)
                {
                    switch (docType)
                    {
                        case SwDmDocumentType.swDmDocumentDrawing:
                            try
                            {
                                GetDrawingReferences(ref dmDoc);
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine("\"" + dmDoc.FullName + "\"\t\"" + e.Message.Replace(Environment.NewLine, " ") + e.StackTrace.Replace(Environment.NewLine, " ") + "File is internally damaged, .Net error occurred, or GetReferences.exe has a Bug." + "\"");
                                inputError(quietMode);
                            }
                            break;
                        case SwDmDocumentType.swDmDocumentPart:
                            try
                            {
                                GetPartReferences(ref dmDoc);
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine("\"" + dmDoc.FullName + "\"\t\"" + e.Message.Replace(Environment.NewLine, " ") + e.StackTrace.Replace(Environment.NewLine, " ") + "File is internally damaged, .Net error occurred, or GetReferences.exe has a Bug." + "\"");
                                inputError(quietMode);
                            }
                            break;
                        case SwDmDocumentType.swDmDocumentAssembly:
                            try
                            {
                                GetAssemblyReferences(ref dmDoc);
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine("\"" + dmDoc.FullName + "\"\t\"" + e.Message.Replace(Environment.NewLine, " ") + e.StackTrace.Replace(Environment.NewLine, " ") + "File is internally damaged, .Net error occurred, or GetReferences.exe has a Bug." + "\"");
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
                            Console.WriteLine("\"" + docPath + "\"\t\"" + "File failed to open; reasons could be related to permissions, the file is in use, or the file is corrupted." + "\"");
                            inputError(quietMode);
                            break;
                        case SwDmDocumentOpenError.swDmDocumentOpenErrorFileNotFound:
                            Console.WriteLine("\"" + docPath + "\"\t\"" + "File not found" + "\"");
                            inputError(quietMode);
                            break;
                        case SwDmDocumentOpenError.swDmDocumentOpenErrorNonSW:
                            Console.Write("\"" + docPath + "\"\t\"" + "Non-SolidWorks file was opened" + "\"");
                            inputError(quietMode);
                            break;
                        default:
                            Console.WriteLine("\"" + docPath + "\"\t\"" + "An unknown errror occurred.  Something is wrong with the code of GetReferences" + "\"");
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

Output for Parts
    ""ParentFilePath""      ""ChildFilePath""

Output for Assemblies
    ""ParentFilePath""      ""ChildFilePath""   

Output for Assemblies
    ""ParentFilePath""      ""ChildFilePath""


Only one Filename is accepted.  No wildcars allowed. If the path has spaces 
use quotes around it.  Note that the file must have one of the following 
file extensions: .sldasm, .slddrw, .sldprt, .asm, .drw, or .prt.  The output
is tab delimited.  This makes it easy to redirect the output to a text file
that can be opened as spreadsheet.

Options
    /q      Quiet mode.  Suppresses the current message.  It does
            not suppress the one line error messages related to problems
            opening SolidWorks Files.  Quiet mode is useful for batch files
            when you are directing the output to a file.  The main error 
            message is suppressed but you are still informed about problems 
            opening files.

Version 2011-Oct-2 15:54
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




        static void GetDrawingReferences(ref SwDMDocument15 dmDrw)
        {
            SwDMSearchOption dmSearchOptions = dmDocManager.GetSearchOptionObject();
            object oBrokenRefs;
            object oIsVirtual;
            object oTimeStamp;

            string[] references = dmDrw.GetAllExternalReferences4(dmSearchOptions, out oBrokenRefs, out oIsVirtual, out oTimeStamp);

            if (references == null || references.Length == 0)
            {
                Console.WriteLine("\"" + dmDrw.FullName + "\"\t");
                return;
            }


            foreach (string reference in references)
            {
                Console.WriteLine("\"" + dmDrw.FullName + "\"\t\t\"" + reference + "\"\t");
            }
        }






        static void GetAssemblyReferences(ref SwDMDocument15 dmAsm)
        {



            SwDMSearchOption dmSearchOptions = dmDocManager.GetSearchOptionObject();
            object oBrokenRefVar;
            object oIsVirtual;
            object oTimeStamp;
            string[] references;

            references = dmAsm.GetAllExternalReferences4(dmSearchOptions, out oBrokenRefVar, out oIsVirtual, out oTimeStamp);

            if (references == null || references.Length == 0)
            {
                Console.WriteLine("\"" + dmAsm.FullName + "\"\t");
                return;
            }

            for (int i = 0; i < references.Length; i++)
            {
                Console.WriteLine("\"" + dmAsm.FullName + "\"\t\t\"" + references[i] + "\"\t");
            }
        }




        static void GetPartReferences(ref SwDMDocument15 dmPart)
        {
            SwDMSearchOption dmSearchOptions = dmDocManager.GetSearchOptionObject();
            object oBrokenRefVar;
            object oIsVirtual;
            object oTimeStamp;
            string[] references;

            references = dmPart.GetAllExternalReferences4(dmSearchOptions, out oBrokenRefVar, out oIsVirtual, out oTimeStamp);

            if (references == null || references.Length == 0)
            {
                Console.WriteLine("\"" + dmPart.FullName + "\"\t");
                return;
            }

            for (int i = 0; i < references.Length; i++)
            {
                Console.WriteLine("\"" + dmPart.FullName + "\"\t\t\"" + references[i] + "\"");
            }
        }

    }
}