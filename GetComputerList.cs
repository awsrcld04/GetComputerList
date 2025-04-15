using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.DirectoryServices;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using Microsoft.Win32;

namespace GetComputerList
{
    class GCLMain
    {
        struct CMDArguments
        {
            public string strQueryFilter;
            public bool bParseCmdArguments;
            public string strCountArgument;
            public bool bLDAPPathArgument;
            public bool bEventLogStartStop;
        }

        static bool funcLicenseCheck()
        {
            string strLicenseString = "";
            bool bValidLicense = false;

            try
            {
                TextReader tr = new StreamReader("sotfwlic.dat");

                try
                {
                    strLicenseString = tr.ReadLine();

                    if (strLicenseString.Length > 0 & strLicenseString.Length < 29)
                    {
                        // [DebugLine] Console.WriteLine("if: " + strLicenseString);
                        Console.WriteLine("Invalid license");

                        tr.Close(); // close license file

                        return bValidLicense;
                    }
                    else
                    {
                        tr.Close(); // close license file
                        // [DebugLine] Console.WriteLine("else: " + strLicenseString);

                        string strMonthTemp = ""; // to convert the month into the proper number
                        string strDate;

                        //Month
                        strMonthTemp = strLicenseString.Substring(7, 1);
                        if (strMonthTemp == "A")
                        {
                            strMonthTemp = "10";
                        }
                        if (strMonthTemp == "B")
                        {
                            strMonthTemp = "11";
                        }
                        if (strMonthTemp == "C")
                        {
                            strMonthTemp = "12";
                        }
                        strDate = strMonthTemp;

                        //Day
                        strDate = strDate + "/" + strLicenseString.Substring(16, 1);
                        strDate = strDate + strLicenseString.Substring(6, 1);

                        // Year
                        strDate = strDate + "/" + strLicenseString.Substring(24, 1);
                        strDate = strDate + strLicenseString.Substring(4, 1);
                        strDate = strDate + strLicenseString.Substring(1, 2);

                        // [DebugLine] Console.WriteLine(strDate);
                        // [DebugLine] Console.WriteLine(DateTime.Today.ToString());
                        DateTime dtLicenseDate = DateTime.Parse(strDate);
                        // [DebugLine]Console.WriteLine(dtLicenseDate.ToString());

                        if (dtLicenseDate >= DateTime.Today)
                        {
                            bValidLicense = true;
                        }
                        else
                        {
                            Console.WriteLine("License expired.");
                        }

                        return bValidLicense;
                    }

                } //end of try block on tr.ReadLine

                catch
                {
                    // [DebugLine] Console.WriteLine("catch on tr.Readline");
                    Console.WriteLine("Invalid license");
                    tr.Close();
                    return bValidLicense;

                } //end of catch block on tr.ReadLine

            } // end of try block on new StreamReader("sotfwlic.dat")

            catch (System.Exception ex)
            {
                // [DebugLine] System.Console.WriteLine("{0} exception caught here.", ex.GetType().ToString());
                
                // [DebugLine] System.Console.WriteLine(ex.Message);
                
                if (ex.Message.StartsWith("Could not find file"))
                {
                    Console.WriteLine("License file not found.");
                }

                return bValidLicense;

            } // end of catch block on new StreamReader("sotfwlic.dat")

        } // LicenseCheck

        static bool funcLicenseActivation()
        {
            try
            {
                if (funcCheckForFile("TurboActivate.dll"))
                {
                    if (funcCheckForFile("TurboActivate.dat"))
                    {
                        TurboActivate.VersionGUID = "4935355894e0da3d4465e86.37472852";

                        if (TurboActivate.IsActivated())
                        {
                            return true;
                        }
                        else
                        {
                            Console.WriteLine("A license for this product has not been activated.");
                            return false;
                        }
                    }
                    else
                    {
                        Console.WriteLine("TurboActivate.dat is required and could not be found.");
                        return false;
                    }
                }
                else
                {
                    Console.WriteLine("TurboActivate.dll is required and could not be found.");
                    return false;
                }
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
                return false;
            }
        }

        static void funcLogToEventLog(string strAppName, string strEventMsg, int intEventType)
        {
            string sLog;

            sLog = "Application";

            if (!EventLog.SourceExists(strAppName))
                EventLog.CreateEventSource(strAppName, sLog);

            //EventLog.WriteEntry(strAppName, strEventMsg);
            EventLog.WriteEntry(strAppName, strEventMsg, EventLogEntryType.Information, intEventType);

        } // LogToEventLog

        static void funcPrintParameterWarning()
        {
            Console.WriteLine("A parameter must be specified to run GetComputerList.");
            Console.WriteLine("Run GetComputerList -? to get the parameter syntax.");
        }

        static void funcPrintParameterSyntax()
        {
            Console.WriteLine("GetComputerList (c) 2011 SystemsAdminPro.com");
            Console.WriteLine();
            Console.WriteLine("Parameter syntax:");
            Console.WriteLine();
            Console.WriteLine("Use the following for the first parameter:");
            Console.WriteLine("-run                required parameter");
            Console.WriteLine();
            Console.WriteLine("Use one of the following for the second parameter:");
            Console.WriteLine("-all                for All computer objects");
            Console.WriteLine("-allservers         for All server computer objects");
            Console.WriteLine("-allworkstations    for All workstation computer objects");
            Console.WriteLine();
            Console.WriteLine("Use one of the following as additional parameters:");
            Console.WriteLine("-count              to include a count with the list");
            Console.WriteLine("-countonly          to only get a count of the number of computers");
            Console.WriteLine("-ldappath           to get the LDAP path from Active Directory");
            Console.WriteLine("-evlog              to log start/stop of GetComputerList to the eventlog");
            Console.WriteLine();
            Console.WriteLine("Examples:");
            Console.WriteLine("GetComputerList -all");
            Console.WriteLine("GetComputerList -allservers -countonly");
            Console.WriteLine("GetComputerList -allservers -count -ldappath");
        } // funcPrintParameterSyntax

        static CMDArguments funcParseCmdArguments(string[] cmdargs)
        {
            CMDArguments objCMDArguments = new CMDArguments();

            try
            {
                objCMDArguments.bParseCmdArguments = false;

                if (cmdargs[0] == "-run" & cmdargs.Length > 1)
                {
                    // [ Comment] Search filter strings for DirectorySearcher object filter
                    string strFilterAll = "(&(objectclass=computer))";
                    string strFilterAllServers = "(&(&(&(sAMAccountType=805306369)(objectCategory=computer)(|(operatingSystem=Windows Server 2008*)(operatingSystem=Windows Server 2003*)(operatingSystem=Windows 2000 Server*)(operatingSystem=Windows NT*)))))";
                    string strFilterAllWorkstations = "(&(&(&(sAMAccountType=805306369)(objectCategory=computer)(|(operatingSystem=Windows XP Pro*)(operatingSystem=Windows 7*)))))";

                    for (int i = 1; i < cmdargs.Length; i++)
                    {
                        if (i == 1)
                        {
                            if (cmdargs[i] == "-all")
                            {
                                objCMDArguments.strQueryFilter = strFilterAll;
                                objCMDArguments.bParseCmdArguments = true;
                            }

                            if (cmdargs[i] == "-allservers")
                            {
                                objCMDArguments.strQueryFilter = strFilterAllServers;
                                objCMDArguments.bParseCmdArguments = true;
                            }

                            if (cmdargs[i] == "-allworkstations")
                            {
                                objCMDArguments.strQueryFilter = strFilterAllWorkstations;
                                objCMDArguments.bParseCmdArguments = true;
                            }
                        }
                        if (i > 0)
                        {

                            if (cmdargs[i] == "-count")
                            {
                                objCMDArguments.strCountArgument = "-count";
                            }

                            if (cmdargs[i] == "-countonly")
                            {
                                objCMDArguments.strCountArgument = "-countOnly";
                            }

                            if (cmdargs[i] == "-ldappath")
                            {
                                objCMDArguments.bLDAPPathArgument = true;
                            }

                            if (cmdargs[i] == "-evlog")
                            {
                                objCMDArguments.bEventLogStartStop = true;
                            }
                        }
                    }
                }
                else
                {
                    objCMDArguments.bParseCmdArguments = false;
                }
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
                objCMDArguments.bParseCmdArguments = false;
            }

            return objCMDArguments;
        }

        static void funcProgramExecution(CMDArguments objCMDArguments2)
        {
            try
            {
                funcProgramRegistryTag("GetComputerList");

                if (objCMDArguments2.bEventLogStartStop)
                {
                    funcLogToEventLog("GetComputerList", "GetComputerList started successfully.", 1301);
                }

                // [Comment] Get local domain context
                string rootDSE;

                System.DirectoryServices.DirectorySearcher objrootDSESearcher = new System.DirectoryServices.DirectorySearcher();
                rootDSE = objrootDSESearcher.SearchRoot.Path;
                // [DebugLine]Console.WriteLine(rootDSE);

                // [Comment] Construct DirectorySearcher object using rootDSE string
                System.DirectoryServices.DirectoryEntry objrootDSEentry = new System.DirectoryServices.DirectoryEntry(rootDSE);
                System.DirectoryServices.DirectorySearcher objComputerObjectSearcher = new System.DirectoryServices.DirectorySearcher(objrootDSEentry);
                // [DebugLine]Console.WriteLine(objComputerObjectSearcher.SearchRoot.Path);

                // [Comment] Add filter to DirectorySearcher object
                objComputerObjectSearcher.Filter = (objCMDArguments2.strQueryFilter);

                // [Comment] Execute query, return results, display name and path values
                System.DirectoryServices.SearchResultCollection objComputerResults = objComputerObjectSearcher.FindAll();
                // [DebugLine]Console.WriteLine(objComputerResults.Count.ToString());
                if (objCMDArguments2.strCountArgument == "-countOnly")
                {
                    Console.WriteLine("Count: " + objComputerResults.Count.ToString());
                }
                else
                {
                    string objComputerDEvalues;
                    string objComputerNameValue;
                    int intStrPosFirst = 3;
                    int intStrPosLast;

                    if (objCMDArguments2.strCountArgument == "-count")
                    {
                        Console.WriteLine("Count: " + objComputerResults.Count.ToString());
                    }
                    foreach (System.DirectoryServices.SearchResult objComputer in objComputerResults)
                    {
                        System.DirectoryServices.DirectoryEntry objComputerDE = new System.DirectoryServices.DirectoryEntry(objComputer.Path);
                        intStrPosLast = objComputerDE.Name.Length;
                        objComputerNameValue = objComputerDE.Name.Substring(intStrPosFirst, intStrPosLast - intStrPosFirst);
                        if (!objCMDArguments2.bLDAPPathArgument)
                        {
                            Console.WriteLine(objComputerNameValue);
                        }
                        else
                        {
                            objComputerDEvalues = objComputerNameValue + "\t" + objComputerDE.Path;
                            Console.WriteLine(objComputerDEvalues);
                        }
                    }
                }
                if (objCMDArguments2.bEventLogStartStop)
                {
                    funcLogToEventLog("GetComputerList", "GetComputerList stopped.", 1302);
                }
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }
        }

        static void funcProgramRegistryTag(string strProgramName)
        {
            try
            {
                string strRegistryProfilesPath = "SOFTWARE";
                RegistryKey objRootKey = Microsoft.Win32.Registry.LocalMachine;
                RegistryKey objSoftwareKey = objRootKey.OpenSubKey(strRegistryProfilesPath, true);
                RegistryKey objSystemsAdminProKey = objSoftwareKey.OpenSubKey("SystemsAdminPro", true);
                if (objSystemsAdminProKey == null)
                {
                    objSystemsAdminProKey = objSoftwareKey.CreateSubKey("SystemsAdminPro");
                }
                if (objSystemsAdminProKey != null)
                {
                    if (objSystemsAdminProKey.GetValue(strProgramName) == null)
                        objSystemsAdminProKey.SetValue(strProgramName, "1", RegistryValueKind.String);
                }
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }
        }

        static void funcGetFuncCatchCode(string strFunctionName, Exception currentex)
        {
            string strCatchCode = "";

            Dictionary<string, string> dCatchTable = new Dictionary<string, string>();
            dCatchTable.Add("funcGetFuncCatchCode", "f0");
            dCatchTable.Add("funcLicenseCheck", "f1");
            dCatchTable.Add("funcPrintParameterWarning", "f2");
            dCatchTable.Add("funcPrintParameterSyntax", "f3");
            dCatchTable.Add("funcParseCmdArguments", "f4");
            dCatchTable.Add("funcProgramExecution", "f5");
            dCatchTable.Add("funcProgramRegistryTag", "f6");
            dCatchTable.Add("funcCreateDSSearcher", "f7");
            dCatchTable.Add("funcCreatePrincipalContext", "f8");
            dCatchTable.Add("funcCheckNameExclusion", "f9");
            dCatchTable.Add("funcMoveDisabledAccounts", "f10");
            dCatchTable.Add("funcFindAccountsToDisable", "f11");
            dCatchTable.Add("funcCheckLastLogin", "f12");
            dCatchTable.Add("funcRemoveUserFromGroup", "f13");
            dCatchTable.Add("funcToEventLog", "f14");
            dCatchTable.Add("funcCheckForFile", "f15");
            dCatchTable.Add("funcCheckForOU", "f16");
            dCatchTable.Add("funcWriteToErrorLog", "f17");

            if (dCatchTable.ContainsKey(strFunctionName))
            {
                strCatchCode = "err" + dCatchTable[strFunctionName] + ": ";
            }

            //[DebugLine] Console.WriteLine(strCatchCode + currentex.GetType().ToString());
            //[DebugLine] Console.WriteLine(strCatchCode + currentex.Message);

            funcWriteToErrorLog(strCatchCode + currentex.GetType().ToString());
            funcWriteToErrorLog(strCatchCode + currentex.Message);

        }

        static void funcWriteToErrorLog(string strErrorMessage)
        {
            try
            {
                FileStream newFileStream = new FileStream("Err-DisabledAccessManager.log", FileMode.Append, FileAccess.Write);
                TextWriter twErrorLog = new StreamWriter(newFileStream);

                DateTime dtNow = DateTime.Now;

                string dtFormat = "MMddyyyy HH:mm:ss";

                twErrorLog.WriteLine("{0} \t {1}", dtNow.ToString(dtFormat), strErrorMessage);

                twErrorLog.Close();
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }

        }

        static bool funcCheckForOU(string strOUPath)
        {
            try
            {
                string strDEPath = "";

                if (!strOUPath.Contains("LDAP://"))
                {
                    strDEPath = "LDAP://" + strOUPath;
                }
                else
                {
                    strDEPath = strOUPath;
                }

                if (DirectoryEntry.Exists(strDEPath))
                {
                    return true;
                }
                else
                {
                    return false;
                }

            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
                return false;
            }
        }

        static bool funcCheckForFile(string strInputFileName)
        {
            try
            {
                if (System.IO.File.Exists(strInputFileName))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
                return false;
            }
        }

        static void Main(string[] args)
        {
            try
            {
                //if (funcLicenseCheck())
                if (funcLicenseActivation())
                {
                    if (args.Length == 0)
                    {
                        funcPrintParameterWarning();
                    }
                    else
                    {
                        if (args[0] == "-?")
                        {
                            funcPrintParameterSyntax();
                        }
                        else
                        {
                            string[] arrArgs = args;
                            CMDArguments objArgumentsProcessed = funcParseCmdArguments(arrArgs);

                            if (objArgumentsProcessed.bParseCmdArguments)
                            {
                                funcProgramExecution(objArgumentsProcessed);
                            }
                            else
                            {
                                funcPrintParameterWarning();
                            } // check objArgumentsProcessed.bParseCmdArguments
                        } // check args[0] = "-?"
                    } // check args.Length == 0
                } // funcLicenseCheck()
            }
            catch (Exception ex)
            {
                Console.WriteLine("errm0: {0}", ex.Message);
            }
        }

    } // class GCLMain

} // namespace GetComputerList
