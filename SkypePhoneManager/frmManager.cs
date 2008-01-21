#region Imports

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using SKYPE4COMLib;
using System.Configuration;

#endregion

namespace SkypePhoneManager
{
    /// <summary>
    /// The main form for the application.
    /// All the main Skype event handlers are here
    /// </summary>
    public partial class frmManager : Form
    {
        #region Constructor

        public frmManager()
        {
            InitializeComponent();
        }

        #endregion

        #region Class Data

        private Skype _objSkype;
        private SKYPE4COMLib.SkypeClass cSkype;
        private string _strMobileUser;

        private bool _blnAttached;
        private bool _blnIsOnline;
        private bool _blnWasAttached;
        private bool _blnPendingSilentModeStartup;

        private string _strShortCutNum1;
        private string _strShortCutNum2;
        private string _strShortCutNum3;

        private int _intIncomingCallId;
        private int _intOutgoingCallId;

        #endregion

        #region Startup

        private void Form1_Load(object sender, EventArgs e)
        {
            // Start up our application
            WriteToLog("Starting SkypePhone Manager v1.3");
            WriteToLog(@"http://leon.mvps.org/SkypePhone/");
            WriteToLog();
            // Get the mobile user account
            WriteToLog("Getting mobile user account details from config");
            this._strMobileUser = ConfigurationManager.AppSettings["SkypeMobileUsername"];
            if (string.IsNullOrEmpty(this._strMobileUser))
            {
                MessageBox.Show("Error! You must edit the config file for this application first and specify the Skype username to forward to!",
                    "Skype Forwarder Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }

            // Get the fast dial numbers
            WriteToLog("Loading in quick call numbers");
            this._strShortCutNum1 = ConfigurationManager.AppSettings["ShortCut1"];
            this._strShortCutNum2 = ConfigurationManager.AppSettings["ShortCut2"];
            this._strShortCutNum3 = ConfigurationManager.AppSettings["ShortCut3"];
            
            // Attach to Skype
            WriteToLog("Attaching to Skype");
            this._objSkype = new Skype();

            // Set up the event handlers
            WriteToLog("Attaching to Skype events");
            cSkype = new SkypeClass();
            cSkype._ISkypeEvents_Event_AttachmentStatus += new _ISkypeEvents_AttachmentStatusEventHandler(OurAttachmentStatus);
            cSkype._ISkypeEvents_Event_ConnectionStatus += new _ISkypeEvents_ConnectionStatusEventHandler(OurConnectionStatus);
            this._objSkype.MessageStatus += new _ISkypeEvents_MessageStatusEventHandler(Skype_MessageStatus);
            this._objSkype.CallStatus += new _ISkypeEvents_CallStatusEventHandler(Skype_CallStatus);

            try
            {
                // Attach to Skype4COM
                cSkype.Attach(7, false);
            }
            catch (Exception)
            {
                // All Skype Logic uses TRY for safety
            }

            try
            {
                if (!_objSkype.Client.IsRunning)
                {
                    _objSkype.Client.Start(false, true);
                }
            }
            catch (Exception)
            {
                // All Skype Logic uses TRY for safety
            }
        }

        public void OurAttachmentStatus(TAttachmentStatus status)
        {
            _blnAttached = false;

            // DEBUG: Write Attachment Status to Window
            //WriteToLog("Attachment Status: " + cSkype.Convert.AttachmentStatusToText(status));
            //WriteToLog(" - " + status.ToString() + Environment.NewLine);

            if (status == TAttachmentStatus.apiAttachAvailable)
            {
                try
                {
                    // This attaches to the Skype4COM class statement
                    cSkype.Attach(7, true);
                }
                catch (Exception)
                {
                    // All Skype Logic uses TRY for safety
                }
            }
            else
                if (status == TAttachmentStatus.apiAttachSuccess)
                {
                    try
                    {
                        System.Windows.Forms.Application.DoEvents();
                        _objSkype.Attach(7, false);
                    }
                    catch (Exception)
                    {
                        // All Skype Logic uses TRY for safety
                    }

                    _blnAttached = true;
                    _blnWasAttached = true;

                    // If we have a queued Silent Mode request, We are attached, process it now
                    if (_blnPendingSilentModeStartup)
                    {
                        _blnPendingSilentModeStartup = false;
                        try
                        {
                            if (!_objSkype.SilentMode) _objSkype.SilentMode = true;
                        }
                        catch (Exception)
                        {
                            // All Skype Logic uses TRY for safety
                        }
                    }
                }
        }

        public void OurConnectionStatus(TConnectionStatus status)
        {
            _blnIsOnline = false;

            // DEBUG: Write Connection Status to Window
            //WriteToLog("Connection Status: " + cSkype.Convert.ConnectionStatusToText(status));
            //WriteToLog(" - " + status.ToString() + Environment.NewLine);

            if (status == TConnectionStatus.conOnline)
            {
                _blnIsOnline = true;
                WriteToLog("Connected to Skype user " + _objSkype.CurrentUserHandle);
                WriteToLog("Now listening for events...");
                WriteToLog();
            }
        }

        #endregion

        #region SkypeIn Functionality

        private void Skype_CallStatus(SKYPE4COMLib.Call aCall, TCallStatus aStatus)
        {
            switch (aStatus)
            {
                case TCallStatus.clsRinging:
                    // A call is ringing, see if it is us calling or
                    // someone calling us
                    if (aCall.Type == TCallType.cltIncomingP2P || aCall.Type == TCallType.cltIncomingPSTN)
                    {
                        // Make sure this isn't a SkypeOut call request
                        if (aCall.PartnerHandle.Equals(_strMobileUser))
                        {
                            WriteToLog("SkypeOut call from " + _strMobileUser);
                            WriteToLog("Allowing call to be forwarded by Skype");
                            return;
                        }
                        // Incoming call, we need to initiate forwarding
                        WriteToLog("Answering call from " + aCall.PartnerHandle);
                        aCall.Answer();
                        this._intIncomingCallId = aCall.Id;
                        aCall.set_InputDevice(TCallIoDeviceType.callIoDeviceTypeSoundcard, "default");
                        aCall.set_InputDevice(TCallIoDeviceType.callIoDeviceTypePort, "1");
                        string strWavPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\PleaseWait2.wav";
                        aCall.set_InputDevice(TCallIoDeviceType.callIoDeviceTypeFile, strWavPath);
                        System.Threading.Thread.Sleep(8000);
                        aCall.set_InputDevice(TCallIoDeviceType.callIoDeviceTypeFile, "");
                        WriteToLog("Placing call on hold");
                        aCall.Hold();
                        
                        WriteToLog("Calling mobile user");
                        System.Threading.Thread.Sleep(500);
                        try
                        {
                            SKYPE4COMLib.Call oCall = _objSkype.PlaceCall(this._strMobileUser, null, null, null);
                            this._intOutgoingCallId = oCall.Id;
                        }
                        catch (Exception ex)
                        {
                            WriteToLog("Error trying to call mobile: " + ex.Message);
                        }
                    }
                    break;
                case TCallStatus.clsInProgress:
                    // We have a new call opened. Make sure it's our outgoing call
                    if (aCall.Type == TCallType.cltOutgoingP2P)
                    {
                        // Yes, the target user has answered.
                        WriteToLog("Mobile user has answered, attempting to join calls");
                        foreach (SKYPE4COMLib.Call objCall in _objSkype.ActiveCalls)
                        {
                            if (objCall.Id == this._intIncomingCallId)
                            {
                                WriteToLog("Joining the calls...");
                                objCall.Join(aCall.Id);
                                WriteToLog("Taking incoming call off hold");
                                objCall.Resume();
                            }
                        }
                    }
                    break;
                case TCallStatus.clsFinished:
                    // Someone has hung up, end the call
                    WriteToLog("Someone has hung up. Attempting to end the conference");
                    foreach (SKYPE4COMLib.Conference objConf in _objSkype.Conferences)
                    {
                        foreach (SKYPE4COMLib.Call objCall in objConf.Calls)
                        {
                            if (objCall.Id == this._intIncomingCallId || objCall.Id == this._intOutgoingCallId)
                            {
                                System.Threading.Thread.Sleep(500);
                                try
                                {
                                    objCall.Finish();
                                }
                                catch (Exception) { }
                                try
                                {
                                    objConf.Finish();
                                }
                                catch (Exception) { }
                            }
                        }
                    }
                    break;
                default:
                    // Something else?
                    if ((aCall.Type == TCallType.cltOutgoingP2P || aCall.Type == TCallType.cltOutgoingPSTN) && (
                            aCall.Status == TCallStatus.clsCancelled ||
                            aCall.Status == TCallStatus.clsFailed ||
                            aCall.Status == TCallStatus.clsMissed ||
                            aCall.Status == TCallStatus.clsRefused ||
                            aCall.Status == TCallStatus.clsVoicemailPlayingGreeting ||
                            aCall.Status == TCallStatus.clsVoicemailRecording
                        )
                       )
                    {
                        WriteToLog("Error calling target user: " + _objSkype.Convert.CallStatusToText(aCall.Status));
                        WriteToLog("Redirecting to voicemail");
                        // End the other call
                        foreach (SKYPE4COMLib.Call objCall in _objSkype.ActiveCalls)
                        {
                            if (objCall.Id == this._intOutgoingCallId)
                            {
                                try
                                {
                                    objCall.Finish();
                                }
                                catch (Exception ex)
                                {
                                    WriteToLog("Error trying to end voicemail call: " + ex.Message);
                                }
                            }
                        }
                        // Now redirect the incoming call
                        foreach (SKYPE4COMLib.Call objCall in _objSkype.ActiveCalls)
                        {
                            if (objCall.Id == this._intIncomingCallId)
                            {
                                System.Threading.Thread.Sleep(500);
                                try
                                {
                                    //objCall.Resume();
                                    objCall.RedirectToVoicemail();
                                    objCall.Finish();
                                    //objCall.Status = TCallStatus.clsFinished;
                                }
                                catch (Exception ex)
                                {
                                    WriteToLog("Error trying to divert to voicemail: " + ex.Message);
                                    objCall.Finish();
                                }
                            }
                        }
                    }
                    break;
            }
        }

        #endregion

        #region SkypeOut Functionality

        private void Skype_MessageStatus(SKYPE4COMLib.ChatMessage pMessage, TChatMessageStatus aStatus)
        {
            if (aStatus == TChatMessageStatus.cmsReceived && pMessage.Type == TChatMessageType.cmeSaid)
            {
                // Make sure the request came from the mobile account
                if (pMessage.Sender.Handle.Equals(this._strMobileUser))
                {
                    // OK, they want to make a change request
                    try
                    {
                        WriteToLog("Forwarding change request received from " + this._strMobileUser);
                        string strNewNum = "";
                        switch (pMessage.Body.ToLower())
                        {
                            case "off":
                                // Switch off forwarding
                                _objSkype.CurrentUserProfile.CallApplyCF = false;
                                _objSkype.CurrentUserProfile.CallForwardRules = "";
                                break;
                            case "1":
                                // Speed switch 1
                                strNewNum = _strShortCutNum1;
                                break;
                            case "2":
                                // Speed switch 1
                                strNewNum = _strShortCutNum2;
                                break;
                            case "3":
                                // Speed switch 1
                                strNewNum = _strShortCutNum3;
                                break;
                            default:
                                strNewNum = pMessage.Body;
                                break;
                        }

                        if (string.IsNullOrEmpty(strNewNum))
                        {
                            pMessage.Chat.SendMessage("Switched off call forwarding");
                            WriteToLog("Switched off call forwarding");
                        }
                        else
                        {
                            _objSkype.CurrentUserProfile.CallApplyCF = true;
                            _objSkype.CurrentUserProfile.CallForwardRules = "0,60," + strNewNum;
                            _objSkype.CurrentUserProfile.CallNoAnswerTimeout = 5;
                            pMessage.Chat.SendMessage("Reset call forwarding to " + strNewNum);
                            WriteToLog("Changed SkypeOut forwarding to " + strNewNum);
                        }
                    }
                    catch (Exception ex)
                    {
                        pMessage.Chat.SendMessage("Error resetting call forwarding to " + pMessage.Body + ": " + ex.Message);
                    }
                }
                else
                {
                    // Someone else? Log and ignore
                    WriteToLog("Chat message received from " + pMessage.Sender.Handle + " was ignored");
                    WriteToLog("Message was: " + pMessage.Body);
                }
            }
        }

        #endregion

        #region Logging

        private void WriteToLog(string pMessage)
        {
            this.txtLog.Text += DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss:") + " " + pMessage + Environment.NewLine;
            this.txtLog.ScrollToCaret();
        }

        private void WriteToLog()
        {
            this.txtLog.Text += Environment.NewLine;
            this.txtLog.ScrollToCaret();
        }

        #endregion
    }
}