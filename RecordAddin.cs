using TRIM.SDK;
using log4net;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace EY.US.RecordAddin
{
    public class RecordAddin : ITrimAddIn
    {
        private StringBuilder m_errorMsg;

        string strclientID = string.Empty;
        string strclientName = string.Empty;
        string strengMetID = string.Empty;
        string strengID = string.Empty;
        string strengName = string.Empty;

        public override string ErrorMessage { get { return m_errorMsg.ToString(); } }

        public override void ExecuteLink(int cmdId, TrimMainObject forObject, ref bool itemWasChanged)
        {

        }

        public override void ExecuteLink(int cmdId, TrimMainObjectSearch forTaggedObjects)
        {

        }

        public override TrimMenuLink[] GetMenuLinks()
        {
            return new TrimMenuLink[0];
        }

        public override void Initialise(Database db)
        {
            Logger.Setup();
            log4net.ILog logger = LogManager.GetLogger(Constants.logName);
            try
            {
                //File.AppendAllText(@"C:\users\pramachandran\desktop\RecordAddin.log", "Initializing");
                m_errorMsg = new StringBuilder();
                //SDKLoader.load();
                logger.Info("Initialisation complete");
                //File.AppendAllText(@"C:\users\pramachandran\desktop\RecordAddin.log", "Initialized");
            }
            catch (Exception ex)
            {
                logger.Error("Initialisation Error", ex);
            }
        }

        public override bool IsMenuItemEnabled(int cmdId, TrimMainObject forObject)
        {
            return false;
        }

        public override void PostDelete(TrimMainObject deletedObject)
        {

        }

        public override void PostSave(TrimMainObject savedObject, bool itemWasJustCreated)
        {

        }

        public override bool PreDelete(TrimMainObject modifiedObject)
        {
            return true;
        }

        #region Notinuse
        //public override bool PreSave(TrimMainObject modifiedObject)
        //{

        //    //#region new
        //    //Database dB = modifiedObject.Database; bool updateStatus = false;
        //    //Record newUserRec = modifiedObject as Record;
        //    //RecordType moRecType = newUserRec.RecordType;
        //    //File.AppendAllText(@"C:\users\pramachandran\desktop\RecordAddin.log", "Presave" + Environment.NewLine);
        //    //try
        //    //{
        //    //    Logger.Setup();
        //    //    log4net.ILog logger = LogManager.GetLogger(Constants.logName);
        //    //    try
        //    //    {

        //    //        BuildFieldNames();

        //    //        string engID = string.Empty;

        //    //        if (moRecType.Name.Equals(Constants.ClientRec) || (moRecType.Name.Equals(Constants.InternalFirmRec)))
        //    //        {
        //    //            File.AppendAllText(@"C:\users\pramachandran\desktop\RecordAddin.log", "Processing Record type - " + moRecType.Name + Environment.NewLine);
        //    //            logger.Info("Processing Record type - " + moRecType.Name);
        //    //            string clientID = string.Empty; Location clientLoc = null; string clientName = string.Empty;

        //    //            FieldDefinition fdEngID = new FieldDefinition(dB, strengID);
        //    //            FieldDefinition fdClientId = new FieldDefinition(dB, strclientID);

        //    //            File.AppendAllText(@"C:\users\pramachandran\desktop\RecordAddin.log", "Eng ID loaded" + Environment.NewLine);

        //    //            clientLoc = newUserRec.Client;

        //    //            File.AppendAllText(@"C:\users\pramachandran\desktop\RecordAddin.log", "Client Name loaded - " + clientLoc.Name + Environment.NewLine);

        //    //            clientID = newUserRec.GetFieldValueAsString(fdClientId, StringDisplayType.TreeColumn, false);

        //    //            //File.AppendAllText(@"C:\users\pramachandran\desktop\RecordAddin.log", "Getting Value from ClientName location- " + newUserRec.GetFieldValue(fdClientName).ToString() + Environment.NewLine);

        //    //            if (moRecType.Name.Equals(Constants.InternalFirmRec))
        //    //            {
        //    //                engID = newUserRec.GetFieldValueAsString(fdEngID, StringDisplayType.TreeColumn, false);
        //    //                logger.Info("Engagement ID - " + engID);
        //    //            }

        //    //            logger.Info("Client Name URI - " + clientLoc.Uri.UriAsString);
        //    //            File.AppendAllText(@"C:\users\pramachandran\desktop\RecordAddin.log", "client name URI - " + clientLoc.Uri.UriAsString + Environment.NewLine);

        //    //            if (!string.IsNullOrEmpty(engID))
        //    //            {
        //    //                try
        //    //                {
        //    //                    string engName = string.Empty;
        //    //                    //FieldDefinition fdClientID = new FieldDefinition(dB, strclientID);
        //    //                    FieldDefinition fdEngName = new FieldDefinition(dB, strengName);

        //    //                    Record EngLocRec = null; //Record CliLocRec = null;

        //    //                    TrimMainObjectSearch recSearch = new TrimMainObjectSearch(dB, BaseObjectTypes.Record);
        //    //                    //recSearch.SetSearchString(fdEngID.SearchClauseName + ":" + engID);
        //    //                    recSearch.SetSearchString("number:" + engID);
        //    //                    logger.Info("search query - " + recSearch.SearchString);
        //    //                    File.AppendAllText(@"C:\users\pramachandran\desktop\RecordAddin.log", "search query - " + recSearch.SearchString + Environment.NewLine);
        //    //                    foreach (Record rec in recSearch)
        //    //                    {
        //    //                        EngLocRec = rec;
        //    //                        break;
        //    //                    }

        //    //                    if (EngLocRec != null)
        //    //                    {
        //    //                        //engName = EngLocRec.GetFieldValueAsString(fdEngName, StringDisplayType.TreeColumn, false);
        //    //                        engName = EngLocRec.Title;
        //    //                        logger.Info("eng. name : " + engName);
        //    //                        File.AppendAllText(@"C:\users\pramachandran\desktop\RecordAddin.log", "Eng Name - " + engName + Environment.NewLine);

        //    //                        clientID = EngLocRec.GetFieldValueAsString(fdClientId, StringDisplayType.TreeColumn, false);
        //    //                        logger.Info("client ID : " + clientID);
        //    //                        File.AppendAllText(@"C:\users\pramachandran\desktop\RecordAddin.log", "Client ID - " + clientID + Environment.NewLine);

        //    //                        TrimMainObjectSearch clientLocSearch = new TrimMainObjectSearch(dB, BaseObjectTypes.Location);
        //    //                        clientLocSearch.SetSearchString("clientID:" + clientID);
        //    //                        //This may return all locations having matching value for 'Client ID'

        //    //                        logger.Info(clientLocSearch.SearchString);
        //    //                        TrimURI clientLocURI = null;
        //    //                        foreach (Location loc in clientLocSearch)
        //    //                        {
        //    //                            clientLoc = loc;
        //    //                            clientLocURI = loc.Uri;
        //    //                            break;
        //    //                        }
        //    //                        if (clientLoc != null)
        //    //                        {
        //    //                            //clientName = CliLocRec.GetFieldValueAsString(fdClientName, StringDisplayType.TreeColumn, false);
        //    //                            clientName = clientLoc.GivenNames;
        //    //                            logger.Info("client name : " + clientName);
        //    //                        }

        //    //                        if (System.Windows.Forms.MessageBox.Show(Constants.popupMsg + fdEngName.Name + " = " + engName + "\n" +
        //    //                            fdClientId.Name + " = " + clientID + "\n" +
        //    //                            clientLoc.Name + " = " + clientName + "\n" +
        //    //                            "\nClick 'Yes' to save", "Verify", System.Windows.Forms.MessageBoxButtons.YesNo).Equals(DialogResult.Yes))
        //    //                        {
        //    //                            //newUserRec.SetFieldValue(fdClientName, new UserFieldValue(clientLocURI));
        //    //                            newUserRec.SetFieldValue(fdEngName, new UserFieldValue(engName));
        //    //                            newUserRec.SetFieldValue(fdClientId, new UserFieldValue(clientID));
        //    //                            updateStatus = true;
        //    //                        }
        //    //                        else
        //    //                        {
        //    //                            m_errorMsg.Clear();
        //    //                            m_errorMsg.Append(Constants.SaveCancel);
        //    //                            logger.Info(ErrorMessage);
        //    //                            updateStatus = false;
        //    //                        }
        //    //                    }
        //    //                }
        //    //                catch (Exception ex)
        //    //                {
        //    //                    File.AppendAllText(@"C:\users\pramachandran\desktop\RecordAddin.log", "ERROR - " + ex.Message + Environment.NewLine);
        //    //                    m_errorMsg.Clear();
        //    //                    m_errorMsg.Append(ex.Message);
        //    //                    logger.Error("Presave-EngID", ex);
        //    //                    updateStatus = false;
        //    //                }
        //    //            }
        //    //            else if (clientLoc != null)
        //    //            {
        //    //                try
        //    //                {
        //    //                    FieldDefinition fdClientIDNumber = new FieldDefinition(dB, strClientIDNum);
        //    //                    FieldDefinition fdClientName = new FieldDefinition(dB, strclientName);

        //    //                    File.AppendAllText(@"C:\users\pramachandran\desktop\RecordAddin.log", "Client ID Number - " + clientID + Environment.NewLine);

        //    //                    clientID = clientLoc.GetFieldValue(fdClientIDNumber).AsString();
        //    //                    clientName = clientLoc.Name;

        //    //                    logger.Info("Client ID - " + clientID);

        //    //                    if (System.Windows.Forms.MessageBox.Show(fdClientIDNumber.Name + " = " + clientID + "\n\n" + fdClientName.Name + " = " + clientName + "\n\n" + "Click 'Yes' to save", "Verify", System.Windows.Forms.MessageBoxButtons.YesNo).Equals(DialogResult.Yes))
        //    //                    {
        //    //                        newUserRec.SetFieldValue(fdClientIDNumber, new UserFieldValue(clientID));
        //    //                        newUserRec.SetFieldValue(fdClientName, new UserFieldValue(clientName));
        //    //                        updateStatus = true;
        //    //                    }
        //    //                    else
        //    //                    {
        //    //                        m_errorMsg.Clear();
        //    //                        m_errorMsg.Append(Constants.SaveCancel);
        //    //                        logger.Info(ErrorMessage);
        //    //                        updateStatus = false;
        //    //                    }

        //    //                }
        //    //                catch (Exception ex)
        //    //                {
        //    //                    File.AppendAllText(@"C:\users\pramachandran\desktop\RecordAddin.log", "ERROR - " + ex.Message);
        //    //                    m_errorMsg.Clear();
        //    //                    m_errorMsg.Append(ex.Message);
        //    //                    logger.Error("PreSave-ClientID", ex);
        //    //                    updateStatus = false;
        //    //                }
        //    //            }
        //    //            else
        //    //            {
        //    //                File.AppendAllText(@"C:\users\pramachandran\desktop\RecordAddin.log", "ERROR - NO MATCH");
        //    //                logger.Info(Constants.NoMatch);
        //    //            }
        //    //        }

        //    //        else
        //    //        {
        //    //            updateStatus = true;
        //    //        }
        //    //    }
        //    //    catch (Exception ex)
        //    //    {
        //    //        File.AppendAllText(@"C:\users\pramachandran\desktop\RecordAddin.log", "ERROR - " + ex.Message);
        //    //        logger.Error("PreSave", ex);
        //    //    }
        //    //}

        //    //catch (Exception ex)
        //    //{
        //    //    File.AppendAllText(@"C:\users\pramachandran\desktop\RecordAddin.log", "ERROR - " + ex.Message);
        //    //    EventLog elog = new EventLog();
        //    //    EventLog.CreateEventSource("RecordAddin", "Application");
        //    //    EventLog.WriteEntry("RecordAddin", ex.Message, EventLogEntryType.Error);
        //    //}
        //    //return updateStatus;
        //    #endregion
        //}
        #endregion

        public override bool PreSave(TrimMainObject modifiedObject)
        {
            Database dB = modifiedObject.Database;
            bool updateStatus = false; bool changesApplied = false;
            Record newUserRec = modifiedObject as Record;
            RecordType recType = newUserRec.RecordType;
            //Constants.writelogInLocal("Event processing starts");

            m_errorMsg.Clear();
            //File.WriteAllText(@"C:\Program Files\Micro Focus\Content Manager\RecordAddinLogs\RAD.txt", "Current User - " + dB.CurrentUser.Name);

            //Check if the user is not 'TRIMServices'. Else the record Addin interferes with the Evt processor during a SAVE.

            if (dB.CurrentUser.Name != "TRIMServices") 
            {
                try
                {
                    Logger.Setup();
                    log4net.ILog logger = LogManager.GetLogger(Constants.logName);
                    try
                    {
                        strengMetID = Properties.Settings.Default.IpEngID;

                        strclientID = Properties.Settings.Default.OpClientID;
                        strclientName = Properties.Settings.Default.OpClientName;
                        strengID = Properties.Settings.Default.OpEngID;
                        strengName = Properties.Settings.Default.OpEngName;


                        if ((recType.Name.Equals(Constants.ClientRec)) || (recType.Name.Equals(Constants.InternalFirmRec)))
                        {
                            //Constants.writelogInLocal("Processing Client Records");
                            logger.Info("Processing Record type - " + recType.Name);
                            string clientID = string.Empty; string engID = string.Empty; string clientName = string.Empty; string engName = string.Empty;
                            string engURI = string.Empty; string engRecNumber = string.Empty;

                            Location clientLoc = null; Record RecEngMetadata = null;

                            FieldDefinition fdEngMetID = new FieldDefinition(dB, strengMetID);
                            FieldDefinition fdEngID = new FieldDefinition(dB, strengID);
                            FieldDefinition fdClientId = new FieldDefinition(dB, strclientID);
                            FieldDefinition fdClientName = new FieldDefinition(dB, strclientName);
                            FieldDefinition fdEngName = new FieldDefinition(dB, strengName);

                            string userMessage = string.Empty;

                            #region load info from EngID
                            //check if Eng info is entered in the form & Populate both engagement ID and engagement Name here
                            UserFieldValue ufvEngId = newUserRec.GetFieldValue(fdEngMetID);

                            //MessageBox.Show("Eng ID Value = " + ufvEngId.ToString()); //tempsection
                            //ufvEngId = new UserFieldValue("E-65848517"); //tempsection

                            if (ufvEngId != null)
                            {
                                //Constants.writelogInLocal("Fetching details from Engagement metadata");

                                engID = ufvEngId.ToString();
                                engURI = engID;
                                logger.Info("Fetching details from Engagement metadata");
                                TrimMainObjectSearch engRecSearch = new TrimMainObjectSearch(dB, BaseObjectTypes.Record);
                                engRecSearch.SetSearchString("uri:" + engID);
                                foreach (Record engMetaRecord in engRecSearch)
                                {
                                    RecEngMetadata = engMetaRecord;
                                    engRecNumber = RecEngMetadata.Number;
                                    engName = engMetaRecord.Title;
                                    engID = engMetaRecord.Number;
                                    break;
                                }
                                if (RecEngMetadata.Client != null)
                                {
                                    clientLoc = RecEngMetadata.Client;
                                    clientName = RecEngMetadata.Client.FullFormattedName;
                                    clientID = RecEngMetadata.GetFieldValueAsString(fdClientId, StringDisplayType.TreeColumn, false);
                                }
                                userMessage = Constants.popupMsg + strclientName + " : " + clientName + "\n" + strclientID + " : " + clientID + "\n" + strengID + " : " + engID + "\n" + strengName + " : " + engName + "\n\nClick 'Yes' to save";
                                changesApplied = true;
                            }
                            #endregion

                            #region Load info from Client

                            //check if client info is entered for 'Client Records' and throw a warning if not specified.
                            //else if (recType.Name.Equals(Constants.ClientRec) && newUserRec.Client.Equals(null))
                            //{
                            //    MessageBox.Show("Client cannot be blank. Need to be specified.", "Warning");
                            //}

                            //check if client info is entered in the form & Populate both client ID and client Name
                            else if (newUserRec.Client != null)
                            {
                                //Constants.writelogInLocal("Fetching details from Client metadata");
                                logger.Info("Fetching details from Client Organisation Location");
                                clientLoc = newUserRec.Client; clientName = clientLoc.FullFormattedName;
                                logger.Info("Client Location Name - " + clientLoc.FullFormattedName);
                                clientID = clientLoc.GetFieldValueAsString(fdClientId, StringDisplayType.TreeColumn, false);
                                userMessage = Constants.popupMsg + strclientName + " : " + clientName + "\n" + strclientID + " : " + clientID + "\n\nClick 'Yes' to save";
                                changesApplied = true;
                            }

                            ///Added this condition for a scenario where user clears the Client property from 
                            ///an interal firm record. This was making our adding retain the CLient Name & ID
                            ///earlier. This condition will make them empty without showing a popup.
                            else if (newUserRec.Client == null)
                            {
                                clientID = newUserRec.GetFieldValueAsString(fdClientId, StringDisplayType.TreeColumn, false);
                                clientName = newUserRec.GetFieldValueAsString(fdClientName, StringDisplayType.TreeColumn, false);

                                if (!string.IsNullOrEmpty(clientID) || !string.IsNullOrEmpty(clientName))
                                {
                                    logger.Info("Client Property is NULL");
                                    clientName = string.Empty;
                                    clientID = string.Empty;
                                    newUserRec.SetFieldValue(fdClientName, new UserFieldValue(clientName));
                                    newUserRec.SetFieldValue(fdClientId, new UserFieldValue(clientID));
                                    updateStatus = true;
                                }
                            }

                            #endregion

                            #region Show popup to user and save the details that are not null or empty
                            if (changesApplied)
                            {
                                if (System.Windows.Forms.MessageBox.Show(userMessage, "Verify", System.Windows.Forms.MessageBoxButtons.YesNo).Equals(DialogResult.Yes))
                                {
                                    if (!string.IsNullOrEmpty(clientID))
                                    {
                                        logger.Info(fdClientId.Name + " : " + clientID);
                                        newUserRec.SetFieldValue(fdClientId, new UserFieldValue(clientID));
                                    }

                                    if (clientLoc != null)
                                    {
                                        logger.Info(fdClientName.Name + "(location) : " + clientLoc.FullFormattedName);
                                        newUserRec.Client = clientLoc; //Mandatory field in the US_Digital dataset
                                    }

                                    if (!string.IsNullOrEmpty(clientName))
                                    {
                                        logger.Info(fdClientName.Name + " : " + clientName);
                                        newUserRec.SetFieldValue(fdClientName, new UserFieldValue(clientName));
                                    }

                                    //Check if Engagement Record number is cleared. If so apply empty string or else set value
                                    if (!string.IsNullOrEmpty(engRecNumber))
                                    {
                                        logger.Info(fdEngMetID.Name + "(Record Number) : " + engRecNumber);
                                        //engRecNumber = RecEngMetadata.Uri.UriAsString; //tempsection
                                        newUserRec.SetFieldValue(fdEngMetID, new UserFieldValue(engRecNumber));
                                        logger.Info(fdEngID.Name + " : " + engRecNumber);
                                        newUserRec.SetFieldValue(fdEngID, new UserFieldValue(engRecNumber));
                                    }
                                    else
                                    {
                                        logger.Info(fdEngMetID.Name + "(Record Number) : " + engRecNumber);
                                        //engRecNumber = RecEngMetadata.Uri.UriAsString; //tempsection
                                        //newUserRec.SetFieldValue(fdEngMetID, null);
                                        logger.Info(fdEngID.Name + " : " + engRecNumber);
                                        newUserRec.SetFieldValue(fdEngID, new UserFieldValue(string.Empty));
                                    }

                                    //Check if Engagement name is cleared. If so apply empty string or else set value
                                    if (!string.IsNullOrEmpty(engName))
                                    {
                                        logger.Info(fdEngName.Name + " : " + engName);
                                        newUserRec.SetFieldValue(fdEngName, new UserFieldValue(engName));
                                    }
                                    else
                                    {
                                        logger.Info(fdEngName.Name + " : " + engName);
                                        newUserRec.SetFieldValue(fdEngName, new UserFieldValue(string.Empty));
                                    }

                                    logger.Info("Record ready to be saved");
                                    updateStatus = true;
                                }
                                else
                                {
                                    m_errorMsg.Clear();
                                    m_errorMsg.Append(Constants.SaveCancel);
                                    logger.Info(ErrorMessage);
                                    updateStatus = false;
                                }
                            }
                            #endregion

                            else
                            {
                                updateStatus = true;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        m_errorMsg.Append(ex.Message);
                        logger.Error("PreSave Error", ex);
                    }

                }
                catch (Exception ex)
                {
                    m_errorMsg.Append(ex.Message);
                }
            }
            else
            {
                updateStatus = true;
            }

            return updateStatus;

        }

        public override bool SelectFieldValue(FieldDefinition field, TrimMainObject trimObject, string previousValue)
        {
            return false;
        }

        public override void Setup(TrimMainObject newObject)
        {

        }

        public override bool SupportsField(FieldDefinition field)
        {
            //if (field.Name.ToUpper().Equals("CLIENT ID"))// || field.Name.ToUpper().Equals("CLIENT NAME"))
            //    return true;
            //else
            return false;
        }

        public override bool VerifyFieldValue(FieldDefinition field, TrimMainObject trimObject, string newValue)
        {
            return true;
        }
    }
}
