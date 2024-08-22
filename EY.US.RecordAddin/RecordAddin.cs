using EY.US.RecordAddin.Properties;
using EY.US.RecordAddin;
using log4net;
using System.Text;
using System.Windows.Forms;
using System;
using TRIM.SDK;
using System.Collections;

namespace EY.US.RecordAddin
{
    public class RecordAddin : ITrimAddIn
    {
        private StringBuilder m_errorMsg;

        private string strclientID = string.Empty;

        private string strclientName = string.Empty;

        private string strengMetID = string.Empty;

        private string strengID = string.Empty;

        private string strengName = string.Empty;

        private bool keyChanged = false;

        public override string ErrorMessage => m_errorMsg.ToString();

        public override void ExecuteLink(int cmdId, TrimMainObject forObject, ref bool itemWasChanged)
        {
        }

        public override void ExecuteLink(int cmdId, TrimMainObjectSearch forTaggedObjects)
        {
        }

        public override TrimMenuLink[] GetMenuLinks()
        {
            return (TrimMenuLink[])(object)new TrimMenuLink[0];
        }

        public override void Initialise(Database db)
        {
            Logger.Setup();
            ILog logger = LogManager.GetLogger(Constants.logName);
            try
            {
                m_errorMsg = new StringBuilder();
                logger.Info((object)"Initialisation complete");
            }
            catch (Exception ex)
            {
                logger.Error((object)"Initialisation Error", ex);
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

        public override bool PreSave(TrimMainObject modifiedObject)
        {           
            Database database = ((TrimPropertySet)modifiedObject).Database;
            bool result = false;
            bool flag = false;
            Record val = (Record)(object)((modifiedObject is Record) ? modifiedObject : null);
            RecordType recordType = val.RecordType;
            m_errorMsg.Clear();
            if (((TrimMainObject)database.CurrentUser).Name != "TRIMServices")
            {
                try
                {
                    Logger.Setup();
                    ILog logger = LogManager.GetLogger(Constants.logName);
                    try
                    {
                        strengMetID = Settings.Default.IpEngID;
                        strclientID = Settings.Default.OpClientID;
                        strclientName = Settings.Default.OpClientName;
                        strengID = Settings.Default.OpEngID;
                        strengName = Settings.Default.OpEngName;
                        if (recordType.Name.Equals(Constants.ClientRec))
                        {
                            if (keyChanged || val.Client == null)
                            {
                                logger.Info((object)("Processing Record type - " + recordType.Name));
                                string text = string.Empty;
                                string empty = string.Empty;
                                string text2 = string.Empty;
                                string text3 = string.Empty;
                                string empty2 = string.Empty;
                                string text4 = string.Empty;
                                Location val2 = null;
                                Record val3 = null;
                                FieldDefinition val4 = new FieldDefinition(database, strengMetID);
                                FieldDefinition val5 = new FieldDefinition(database, strengID);
                                FieldDefinition val6 = new FieldDefinition(database, strclientID);
                                FieldDefinition val7 = new FieldDefinition(database, strclientName);
                                FieldDefinition val8 = new FieldDefinition(database, strengName);
                                string text5 = string.Empty;
                                UserFieldValue fieldValue = val.GetFieldValue(val4);
                                if (fieldValue != null)
                                {
                                    empty = ((object)fieldValue).ToString();
                                    empty2 = empty;
                                    logger.Info((object)"Fetching details from Engagement metadata");
                                    TrimMainObjectSearch val9 = new TrimMainObjectSearch(database, (BaseObjectTypes)6);
                                    val9.SetSearchString("uri:" + empty2);
                                    {
                                        IEnumerator enumerator = val9.GetEnumerator();
                                        try
                                        {
                                            if (enumerator.MoveNext())
                                            {
                                                Record val10 = (Record)enumerator.Current;
                                                val3 = val10;
                                                text4 = val3.Number;
                                                text3 = val10.Title;
                                                empty = val10.Number;
                                            }
                                        }
                                        finally
                                        {
                                            IDisposable disposable = enumerator as IDisposable;
                                            if (disposable != null)
                                            {
                                                disposable.Dispose();
                                            }
                                        }
                                    }
                                    if (val3.Client != null)
                                    {
                                        val2 = val3.Client;
                                        text2 = val3.Client.FullFormattedName;
                                        text = val3.GetFieldValueAsString(val6, (StringDisplayType)2, false);
                                    }
                                    text5 = Constants.popupMsg + strclientName + " : " + text2 + "\n" + strclientID + " : " + text + "\n" + strengID + " : " + empty + "\n" + strengName + " : " + text3 + "\n\nClick 'Yes' to save";
                                    flag = true;
                                }
                                else if (val.Client != null)
                                {
                                    val2 = val.Client;
                                    text2 = val2.FullFormattedName;
                                    text = val2.GetFieldValueAsString(val6, (StringDisplayType)2, false);
                                    text5 = Constants.popupMsg + strclientName + " : " + text2 + "\n" + strclientID + " : " + text + "\n\nClick 'Yes' to save";
                                    flag = true;
                                }
                                else
                                {
                                    flag = false;
                                    m_errorMsg.Clear();
                                    m_errorMsg.Append(Constants.MissingInfo);
                                    result = false;
                                }
                                if (flag)
                                {
                                    if (MessageBox.Show(text5, "Verify", MessageBoxButtons.YesNo).Equals(DialogResult.Yes))
                                    {
                                        if (!string.IsNullOrEmpty(text))
                                        {
                                            logger.Info((object)(val6.Name + " : " + text));
                                            val.SetFieldValue(val6, new UserFieldValue(text));
                                        }
                                        if (val2 != null)
                                        {
                                            logger.Info((object)(val7.Name + "(location) : " + val2.FullFormattedName));
                                            val.Client = val2;
                                        }
                                        if (!string.IsNullOrEmpty(text2))
                                        {
                                            logger.Info((object)(val7.Name + " : " + text2));
                                            val.SetFieldValue(val7, new UserFieldValue(text2));
                                        }
                                        if (!string.IsNullOrEmpty(text4))
                                        {
                                            logger.Info((object)(val4.Name + "(Record Number) : " + text4));
                                            val.SetFieldValue(val4, new UserFieldValue(text4));
                                            logger.Info((object)(val5.Name + " : " + text4));
                                            val.SetFieldValue(val5, new UserFieldValue(text4));
                                        }
                                        else
                                        {
                                            logger.Info((object)(val4.Name + "(Record Number) : " + text4));
                                            logger.Info((object)(val5.Name + " : " + text4));
                                            val.SetFieldValue(val5, new UserFieldValue(string.Empty));
                                        }
                                        if (!string.IsNullOrEmpty(text3))
                                        {
                                            logger.Info((object)(val8.Name + " : " + text3));
                                            val.SetFieldValue(val8, new UserFieldValue(text3));
                                        }
                                        else
                                        {
                                            logger.Info((object)(val8.Name + " : " + text3));
                                            val.SetFieldValue(val8, new UserFieldValue(string.Empty));
                                        }
                                        logger.Info((object)"Record ready to be saved");
                                        result = true;
                                    }
                                    else
                                    {
                                        m_errorMsg.Clear();
                                        m_errorMsg.Append(Constants.SaveCancel);
                                        logger.Info((object)((ITrimAddIn)this).ErrorMessage);
                                        result = false;
                                    }
                                    keyChanged = false;
                                }
                            }
                            else
                            {
                                result = true;
                            }
                        }
                        if (recordType.Name.Equals(Constants.InternalFirmRec))
                        {
                            logger.Info((object)("Processing Record type - " + recordType.Name));
                            string text6 = string.Empty;
                            string empty3 = string.Empty;
                            string text7 = string.Empty;
                            string text8 = string.Empty;
                            string empty4 = string.Empty;
                            string text9 = string.Empty;
                            Location val11 = null;
                            Record val12 = null;
                            FieldDefinition val13 = new FieldDefinition(database, strengMetID);
                            FieldDefinition val14 = new FieldDefinition(database, strengID);
                            FieldDefinition val15 = new FieldDefinition(database, strclientID);
                            FieldDefinition val16 = new FieldDefinition(database, strclientName);
                            FieldDefinition val17 = new FieldDefinition(database, strengName);
                            string text10 = string.Empty;
                            UserFieldValue fieldValue2 = val.GetFieldValue(val13);
                            if (keyChanged || val.Client == null)
                            {
                                if (fieldValue2 != null)
                                {
                                    empty3 = ((object)fieldValue2).ToString();
                                    empty4 = empty3;
                                    logger.Info((object)"Fetching details from Engagement metadata");
                                    TrimMainObjectSearch val18 = new TrimMainObjectSearch(database, (BaseObjectTypes)6);
                                    val18.SetSearchString("uri:" + empty3);
                                    {
                                        IEnumerator enumerator2 = val18.GetEnumerator();
                                        try
                                        {
                                            if (enumerator2.MoveNext())
                                            {
                                                Record val19 = (Record)enumerator2.Current;
                                                val12 = val19;
                                                text9 = val12.Number;
                                                text8 = val19.Title;
                                                empty3 = val19.Number;
                                            }
                                        }
                                        finally
                                        {
                                            IDisposable disposable2 = enumerator2 as IDisposable;
                                            if (disposable2 != null)
                                            {
                                                disposable2.Dispose();
                                            }
                                        }
                                    }
                                    if (val12.Client != null)
                                    {
                                        val11 = val12.Client;
                                        text7 = val12.Client.FullFormattedName;
                                        text6 = val12.GetFieldValueAsString(val15, (StringDisplayType)2, false);
                                    }
                                    text10 = Constants.popupMsg + strclientName + " : " + text7 + "\n" + strclientID + " : " + text6 + "\n" + strengID + " : " + empty3 + "\n" + strengName + " : " + text8 + "\n\nClick 'Yes' to save";
                                    flag = true;
                                }
                                else if (val.Client != null)
                                {
                                    logger.Info((object)"Fetching details from Client Organisation Location");
                                    val11 = val.Client;
                                    text7 = val11.FullFormattedName;
                                    logger.Info((object)("Client Location Name - " + val11.FullFormattedName));
                                    text6 = val11.GetFieldValueAsString(val15, (StringDisplayType)2, false);
                                    text10 = Constants.popupMsg + strclientName + " : " + text7 + "\n" + strclientID + " : " + text6 + "\n\nClick 'Yes' to save";
                                    flag = true;
                                }
                                else
                                {
                                    logger.Info((object)"Client Property is NULL");
                                    text7 = string.Empty;
                                    text6 = string.Empty;
                                    text8 = string.Empty;
                                    val.SetFieldValue(val16, new UserFieldValue(text7));
                                    val.SetFieldValue(val15, new UserFieldValue(text6));
                                    val.SetFieldValue(val17, new UserFieldValue(text8));
                                    val.SetFieldValue(val14, new UserFieldValue(empty3));
                                    result = true;
                                }
                            }
                            if (flag)
                            {
                                if (MessageBox.Show(text10, "Verify", MessageBoxButtons.YesNo).Equals(DialogResult.Yes))
                                {
                                    if (!string.IsNullOrEmpty(text6))
                                    {
                                        logger.Info((object)(val15.Name + " : " + text6));
                                        val.SetFieldValue(val15, new UserFieldValue(text6));
                                    }
                                    if (val11 != null)
                                    {
                                        logger.Info((object)(val16.Name + "(location) : " + val11.FullFormattedName));
                                        val.Client = val11;
                                    }
                                    if (!string.IsNullOrEmpty(text7))
                                    {
                                        logger.Info((object)(val16.Name + " : " + text7));
                                        val.SetFieldValue(val16, new UserFieldValue(text7));
                                    }
                                    if (!string.IsNullOrEmpty(text9))
                                    {
                                        logger.Info((object)(val13.Name + "(Record Number) : " + text9));
                                        val.SetFieldValue(val13, new UserFieldValue(text9));
                                        logger.Info((object)(val14.Name + " : " + text9));
                                        val.SetFieldValue(val14, new UserFieldValue(text9));
                                    }
                                    else
                                    {
                                        logger.Info((object)(val13.Name + "(Record Number) : " + text9));
                                        logger.Info((object)(val14.Name + " : " + text9));
                                        val.SetFieldValue(val14, new UserFieldValue(string.Empty));
                                    }
                                    if (!string.IsNullOrEmpty(text8))
                                    {
                                        logger.Info((object)(val17.Name + " : " + text8));
                                        val.SetFieldValue(val17, new UserFieldValue(text8));
                                    }
                                    else
                                    {
                                        logger.Info((object)(val17.Name + " : " + text8));
                                        val.SetFieldValue(val17, new UserFieldValue(string.Empty));
                                    }
                                    logger.Info((object)"Record ready to be saved");
                                    result = true;
                                }
                                else
                                {
                                    m_errorMsg.Clear();
                                    m_errorMsg.Append(Constants.SaveCancel);
                                    logger.Info((object)((ITrimAddIn)this).ErrorMessage);
                                    result = false;
                                }
                                keyChanged = false;
                            }
                            else
                            {
                                result = true;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        m_errorMsg.Append(ex.Message);
                        logger.Error((object)"PreSave Error", ex);
                    }
                }
                catch (Exception ex2)
                {
                    m_errorMsg.Append(ex2.Message);
                }
            }
            else
            {
                result = true;
            }
            return result;
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
            //IL_001d: Unknown result type (might be due to invalid IL or missing references)
            //IL_0023: Expected O, but got Unknown
            strengMetID = Settings.Default.IpEngID;
            FieldDefinition val = new FieldDefinition(((TrimPropertySet)field).Database, strengMetID);
            if (field.Name.Equals(val.Name))
            {
                return true;
            }
            return false;
        }

        public override bool VerifyFieldValue(FieldDefinition field, TrimMainObject trimObject, string newValue)
        {
            //IL_0061: Unknown result type (might be due to invalid IL or missing references)
            //IL_0068: Expected O, but got Unknown
            Record val = (Record)(object)((trimObject is Record) ? trimObject : null);
            string empty = string.Empty;
            strclientID = Settings.Default.OpClientID;
            string name = val.RecordType.Name;
            try
            {
                if (!((TrimObject)val).Uri.UriAsString.Equals("0"))
                {
                    UserFieldValue fieldValue = val.GetFieldValue(field);
                    FieldDefinition val2 = new FieldDefinition(((TrimPropertySet)field).Database, strclientID);
                    UserFieldValue fieldValue2 = val.GetFieldValue(val2);
                    if (fieldValue != null)
                    {
                        empty = ((object)fieldValue).ToString().ToUpper();
                        if (!empty.Equals(newValue.ToUpper()) || val.Client == null)
                        {
                            keyChanged = true;
                        }
                    }
                    else if (fieldValue == null && !newValue.Equals("0"))
                    {
                        keyChanged = true;
                    }
                    if (fieldValue2 != null)
                    {
                        string text = ((object)fieldValue2).ToString();
                        if (val.Client != null && text != val.Client.IdNumber)
                        {
                            keyChanged = true;
                        }
                    }
                }
                else
                {
                    keyChanged = true;
                }
            }
            catch (Exception ex)
            {
                if (name.Equals("Client Records"))
                {
                    MessageBox.Show(ex.Message);
                }
            }
            return true;
        }
    }
}