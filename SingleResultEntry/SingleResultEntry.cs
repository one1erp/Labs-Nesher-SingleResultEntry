using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Common;
using DAL;
using LSExtensionWindowLib;
using LSSERVICEPROVIDERLib;
using MSXML;
using Telerik.WinControls.UI;
using XmlService;

namespace SingleResultEntry
{
    /// <summary>
    /// התוכנית מחזיקה 2 פאנלים 
    ///  עבור הזנת תוצאה עם מיהול ועבור הזנה ללא מיהול
    /// </summary>
    [ComVisible(true)]
    [ProgId("SingleResultEntry.SingleResultEntry")]
    public partial class SingleResultEntry : UserControl, IExtensionWindow
    {



        #region Ctor

        public SingleResultEntry()
        {
            InitializeComponent();
            BackColor = Color.FromName("Control");
        }

        #endregion

        #region private members

        private Result _currentResult;
        private DOMDocument _getResultXmlRes;
        private INautilusDBConnection _ntlsCon;
        private IExtensionWindowSite2 _ntlsSite;
        private INautilusProcessXML _processXml;
        private bool _withDilution;
        private IDataLayer dal;
        //ספי אמר שהערך הזה ייכנס כאשר המשתמש יסמן גדול מערכי מקסימום
        private const string maxValue = "999999";

        private bool WithDilution
        {
            get { return _withDilution; }
            set
            {
                _withDilution = value;
                timerFocus.Start();
            }
        }

        #endregion

        #region Implementation of IExtensionWindow

        public bool CloseQuery()
        {
            if (dal != null) dal.Close();
            return true;
        }

        public void Internationalise()
        {
        }

        public void SetSite(object site)
        {
            _ntlsSite = (IExtensionWindowSite2)site;
            _ntlsSite.SetWindowInternalName("הזנת תוצאה");
            _ntlsSite.SetWindowRegistryName("הזנת תוצאה");
            _ntlsSite.SetWindowTitle("הזנת תוצאה");
        }

        public void PreDisplay()
        {
            Utils.CreateConstring(_ntlsCon);
            dal = new DataLayer();
            dal.Connect();
            btnClean_Click(null, null);
            //בכדי שנאוטילוס יעשה פוקוס צריך להפעיל טיימר
            timerFocus.Start();

        }

        public WindowButtonsType GetButtons()
        {
            return WindowButtonsType.windowButtonsNone;
        }

        public bool SaveData()
        {
            return false; //???
        }

        public void SetServiceProvider(object serviceProvider)
        {
            sp = serviceProvider as NautilusServiceProvider;
            _processXml = Utils.GetXmlProcessor(sp);
            _ntlsCon = Utils.GetNtlsCon(sp);
        }

        public void SetParameters(string parameters)
        {


            WithDilution = parameters == "Dilution";

            if (parameters == "WaterDilution")
            {
                WithDilution = true;
                SetWaterDilutionText();

            }
            else if (parameters == "Water")
            {
                WithDilution = false;
                SetWaterDilutionText();
            }

            //כאשר יישלח הפרמטר הנ"ל יעלה הפאנל של הזנת תוצאה עם מיהול
            panelDilution.Visible = WithDilution;
            regularPanel.Visible = !WithDilution;
            txtResultID.Select();
            if (WithDilution)
            {
                d_txtResultID.Select();

                //רישום של כל RADIO BUTTON
                IEnumerable<RadRadioButton> radioBtnList = GetRadioBtnList();
                foreach (RadRadioButton radRadioButton in radioBtnList)
                {
                    radRadioButton.ToggleStateChanged += radRadioButton_ToggleStateChanged;
                }
            }
        }

        private void SetWaterDilutionText()
        {
            radRadioButton1.Text = "100ml";
            radRadioButton2.Text = "10ml";
            radRadioButton4.Text = "1ml";
            radRadioButton5.Text = "0.1ml";
            radRadioButton6.Text = "0.01ml";
            radRadioButton7.Text = "0.001ml";
            radRadioButton8.Text = "0.0001ml";
        }


        public void Setup()
        {
        }

        public WindowRefreshType DataChange()
        {
            return WindowRefreshType.windowRefreshNone;
        }

        public WindowRefreshType ViewRefresh()
        {
            return WindowRefreshType.windowRefreshNone;
        }

        public void refresh()
        {
        }

        public void SaveSettings(int hKey)
        {
        }

        public void RestoreSettings(int hKey)
        {
        }

        #endregion

        #region Common

        #region  events

        private void WindowExtension_Load(object sender, EventArgs e)
        {
            InitialMode();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            //   waitnigTimer.Stop();
            System.Media.SystemSounds.Beep.Play();
            InitialMode();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {

            if (WithDilution)
            {
                SetResultValue(d_spinResultValue, d_cbBig);
            }
            else
            {
                SetResultValue(spinResultValue, cbNotCount);
            }
        }

        private void btnClean_Click(object sender, EventArgs e)
        {
            InitialMode();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            _ntlsSite.CloseWindow();
        }

        #region New region

        private void txtResultID_KeyDown(object sender, KeyEventArgs e)
        {

            try
            {

                //Application.UseWaitCursor = true;


                var senderTxtResult = (RadTextBox)sender;
                if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Tab)
                {

                    var value = senderTxtResult.Text.Trim();
                    int resultId;
                    if ((value.StartsWith("R", StringComparison.OrdinalIgnoreCase) || value.StartsWith("ר"))
                        && int.TryParse(value.Replace(value[0], ' ').Trim(), out resultId))
                    {

                        EnableButtons(false);
                        _currentResult = dal.IsGoodResultForEntry(resultId, 'C', 'V', 'U', 'S', 'P');
                        if (_currentResult != null)
                        {
                            string clientName;
                            try
                            {
                                clientName = _currentResult.Test.Aliquot.Sample.Sdg.Client.Name;
                            }
                            catch (Exception e1)
                            {
                                Logger.WriteLogFile(e1);
                                clientName = "";
                            }




                            senderTxtResult.Enabled = false;


                            if (!WithDilution)
                            {
                                lbl_client.Text = clientName;

                                txtResultName.Text = _currentResult.Name;
                                spinResultValue.Enabled = true;
                                spinResultValue.Text = "";
                                lblError.Text = "";
                                cbNotCount.Enabled = true;
                                spinResultValue.Select();
                            }
                            else
                            {

                                lbl_dilutionClient.Text = clientName;
                                d_txtResultName.Text = _currentResult.Name;
                                d_lblDuplicates.Text = GetDuplicates().ToString();
                                d_spinResultValue.Select();
                                d_cbBig.Enabled = true;
                                d_lblError.Text = "";
                                d_groupBox1.Enabled = true;
                                d_spinResultValue.Enabled = true;
                                d_spinResultValue.Text = "";
                                d_btnOk.Enabled = false;
                            }
                        }
                        else
                        {
                            BadValue(senderTxtResult, ".מספר בדיקה לא קיים, או אינו בסטטוס הנכון");

                            //או סטטוס
                        }
                        EnableButtons(true);
                    }
                    else
                    {
                        BadValue(senderTxtResult, "מספר בדיקה אינו מתאים");
                    }
                }
            }


            catch (Exception ex)
            {
                Logger.WriteLogFile(ex);
            }
            finally
            {
                //    Application.UseWaitCursor = false;
            }
        }

        private void BadValue(RadTextBox senderTxtResult, string errorMessage)
        {
            EnableButtons(false);
            string message = errorMessage;
            //   lblError.Text = message;
            //   d_lblError.Text = message;
            System.Media.SystemSounds.Beep.Play();
            System.Media.SystemSounds.Beep.Play();
            System.Media.SystemSounds.Beep.Play();
            System.Media.SystemSounds.Beep.Play();
            MessageBox.Show(errorMessage);
            senderTxtResult.Focus();

            senderTxtResult.Text = string.Empty;



        }

        #endregion

        private void spinResultValue_KeyDown(object sender, KeyEventArgs e)
        {
            if (WithDilution) return;
            if (e.KeyValue == 13 && spinResultValue.Text.Length > 0)
            {
                var senderSpinEditor = (RadSpinEditor)sender;
                SetResultValue(senderSpinEditor, cbNotCount);
            }
        }

        #endregion

        #region XML

        public void Close()
        {
        }

        private bool CreateResultEntryXml(bool isDefaultValue, RadSpinEditor senderSpinEditor)
        {

            var objDom = new DOMDocument();

            IXMLDOMElement objResultRequest = objDom.createElement("result-request");
            objDom.appendChild(objResultRequest);

            IXMLDOMElement objLoad = objDom.createElement("load");
            objLoad.setAttribute("entity", "TEST");

            objLoad.setAttribute("id", _currentResult.TestId);
            objLoad.setAttribute("mode", "entry");

            objResultRequest.appendChild(objLoad);

            if (!isDefaultValue)
            {
                IXMLDOMElement objResultEntryElem = ObjResultEntryElem(objDom, senderSpinEditor.Text);
                objLoad.appendChild(objResultEntryElem);
            }
            else
            {
                //  IXMLDOMElement objResultDefaultElem = ObjResultDefaultElem(objDom);

                IXMLDOMElement objResultDefaultElem = ObjResultEntryElem(objDom, maxValue);
                objLoad.appendChild(objResultDefaultElem);
            }
            var res = new DOMDocument();

            _processXml.ProcessXMLWithResponse(objDom, res);

            // For testing
            //  objDom.save(@"C:\temp\docResultEntry.xml");
            // res.save(@"C:\temp\resResultEntry.xml");


            IXMLDOMNode answer = res.getElementsByTagName("errors")[0];
            if (answer != null)
                Logger.WriteLogFile(answer.ToString() + " " + answer.text, true);

            return answer == null;
        }


        /// <summary>
        /// Add element result-default to xml
        /// </summary>
        /// <param name="objDom">Base DOMDocument</param>
        /// <returns>new element</returns>
        private IXMLDOMElement ObjResultDefaultElem(DOMDocument objDom)
        {
            IXMLDOMElement objResultEntryElem = objDom.createElement("result-default");
            objResultEntryElem.setAttribute("result-id", _currentResult.ResultId);
            return objResultEntryElem;
        }

        /// <summary>
        /// Add element result-entry to xml
        /// </summary>
        /// <param name="objDom">Base DOMDocument</param>
        /// <param name="value">result entry value</param>
        /// <returns>new element</returns>
        private IXMLDOMElement ObjResultEntryElem(DOMDocument objDom, string value)
        {
            IXMLDOMElement objResultEntryElem = objDom.createElement("result-entry");
            objResultEntryElem.setAttribute("result-id", _currentResult.ResultId);
            objResultEntryElem.setAttribute("original-result", value);
            return objResultEntryElem;
        }

        #endregion


        private void SetResultValue(RadSpinEditor senderSpinEditor, RadCheckBox checkBox)
        {
            try
            {

                // Application.UseWaitCursor = true;

                EnableButtons(false);

                if (!WithDilution)
                {
                    CalculateDilutionFactor();
                    bool succeed = CreateResultEntryXml(checkBox.Checked, senderSpinEditor);
                    EnableButtons(true);
                    if (succeed)
                    {
                        //    ReCalculate();

                        //  waitnigTimer.Start();
                        senderSpinEditor.Enabled = false;
                        lblError.Text = "";
                        pictureBox1.Enabled = true;
                        // pictureBox1.Image = new Bitmap("../../approved.jpg");
                        txtResultID.Enabled = false;
                        btnClose.Enabled = false;
                        btnClean.Enabled = false;
                        cbNotCount.Enabled = false;
                        btnOk.Enabled = false;
                        btnReject.Enabled = false;
                        regularPanel.Enabled = false;
                        timer1_Tick(null, null);
                    }
                    else
                    {
                        MessageBox.Show("תוצאה אינה תקינה");
                        senderSpinEditor.Select();
                    }
                }
                else
                {
                    bool succeed = CreateResultEntryXml(checkBox.Checked, senderSpinEditor);
                    EnableButtons(true);
                    if (succeed)
                    {
                        //  ReCalculate();
                        //     waitnigTimer.Start();
                        senderSpinEditor.Enabled = false;
                        d_pictureBox.Enabled = true;
                        //   d_pictureBox.Image = new Bitmap("../../approved.jpg");
                        d_groupBox1.RemoveRedBorder();
                        btnClose.Enabled = false;
                        d_btnClean.Enabled = false;
                        d_btnReject.Enabled = false;
                        d_btnOk.Enabled = false;
                        d_lblError.Text = "";
                        panelDilution.Enabled = false;
                        //תמיד זה היה קורה שהטיימר נגמר מכיוון שספי אמר להוריד את הטיימר קראתי לפונקציה
                        timer1_Tick(null, null);
                    }
                    else
                    {
                        //   d_pictureBox.Image = new Bitmap("../../rejected.jpg");
                        d_lblError.Text = "";
                        MessageBox.Show("תוצאה אינה תקינה");

                    }
                }
            }
            catch (Exception exception)
            {

                Logger.WriteLogFile(exception);
            }
            finally
            {
                //    Cursor.Current = Cursors.Default;
            }
        }


        private void InitialMode()
        {


            if (WithDilution)
                InitialDilutionPanel();
            else
                InitialwithoutDilutionPanel();

            //  Cursor.Current = Cursors.Default;
        }

        private void EnableButtons(bool flag)
        {
            if (!WithDilution)
            {
                btnClean.Enabled = flag;
                btnClose.Enabled = flag;
                btnOk.Enabled = flag;
                btnReject.Enabled = flag;

            }
            else
            {
                d_btnClean.Enabled = flag;
                d_btnClose.Enabled = flag;
            }
        }

        private void ReCalculate()
        {
            var summuriesResults = _currentResult.Test.Results.Where(result => result.ResultType == "S").ToList();
            foreach (Result result in summuriesResults)
            {
                FireEvent(result);
            }
        }

        private void FireEvent(Result result)
        {

            var fireEventXmlHandler = new FireEventXmlHandler(sp);
            fireEventXmlHandler.CreateFireEventXml("RESULT", result.ResultId, "recalculate");
            var success = fireEventXmlHandler.ProcssXml();
            if (!success)
            {
                //TODO:Check inner message
                Logger.WriteLogFile(fireEventXmlHandler.ErrorResponse, true);
                MessageBox.Show(fireEventXmlHandler.ErrorResponse);
            }

        }

        #endregion

        #region Panel without dilution

        #region Events

        private void spinResultValue_TextChanged(object sender, EventArgs e)
        {

            try
            {
                if (spinResultValue.Text.Length > 0 && int.Parse(spinResultValue.Text) > 0 || spinResultValue.Value > 0)
                {
                    btnOk.Enabled = true;
                    btnReject.Enabled = true;

                    cbNotCount.Enabled = false;
                }
                else
                {
                    btnOk.Enabled = false;
                    btnReject.Enabled = false;

                    cbNotCount.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                Logger.WriteLogFile(ex);
            }
        }

        private void cbNotCount_ToggleStateChanged(object sender, StateChangedEventArgs args)
        {
            if (cbNotCount.Checked)
            {
                spinResultValue.Value = 0;
                spinResultValue.Text = "";
                spinResultValue.Enabled = false;
                btnOk.Enabled = true;
                btnReject.Enabled = true;

            }
            else
            {
                btnOk.Enabled = false;
                btnReject.Enabled = false;

                spinResultValue.Enabled = true;
                spinResultValue.Focus();
            }
        }

        #endregion

        private void InitialwithoutDilutionPanel()
        {
            lbl_client.Text = "";
            txtResultName.Text = string.Empty;
            cbNotCount.Checked = false;
            txtResultID.Text = string.Empty;
            txtResultID.Enabled = true;
            spinResultValue.Value = 0;
            spinResultValue.Enabled = false;
            btnOk.Enabled = false;
            btnReject.Enabled = false;

            btnClean.Enabled = true;
            cbNotCount.Enabled = false;
            pictureBox1.Enabled = true;
            pictureBox1.Image = null;
            btnReject.Enabled = false;
            btnClose.Enabled = true;
            lblError.Text = "";
            _getResultXmlRes = null;
            regularPanel.Enabled = true;
            txtResultID.Select();
            txtResultID.Focus();

        }

        private void CalculateDilutionFactor()
        {
            if (_currentResult.ResultTemplate.Extension == null || _currentResult.DilutionFactor == null) return;

            if (_currentResult.ResultTemplate.Extension.Name == "Dilution Calculation Extension" &&
                _currentResult.DilutionFactor > 0)
            {
                spinResultValue.Value = spinResultValue.Value * (decimal)_currentResult.DilutionFactor;
            }
        }

        #endregion

        #region DilutinPanel

        private int oldTag = 1;
        private NautilusServiceProvider sp;


        private void btnReject_Click(object sender, EventArgs e)
        {
            _currentResult.Status = "X";
            dal.SaveChanges();
            InitialMode();
            //SetResultValue(d_spinResultValue, d_cbBig);
            //AuthorisatioRejectnXml();
            //  ReCalculate();
            panelDilution.Enabled = true;
        }

        private void d_cbBig_ToggleStateChanged(object sender, StateChangedEventArgs args)
        {
            if (d_cbBig.Checked)
            {
                d_spinResultValue.Value = 0;
                d_spinResultValue.Text = "";
                d_spinResultValue.Enabled = false;
                d_groupBox1.Enabled = false;
                d_btnReject.Enabled = false;
                InitRadioBtnList();
                d_btnOk.Enabled = true;
            }
            else
            {
                d_groupBox1.Enabled = true;
                d_btnReject.Enabled = false;
                d_spinResultValue.Enabled = true;
                d_spinResultValue.Focus();
                d_groupBox1.RemoveRedBorder();
                d_btnOk.Enabled = false;
            }
        }

        private void radRadioButton_ToggleStateChanged(object sender, StateChangedEventArgs args)
        {
            try
            {


                d_btnOk.Enabled = true;
                var radioButton = (RadRadioButton)sender;
                if (radioButton.IsChecked)
                {
                    //The tag is defined in design
                    int tag = int.Parse(radioButton.Tag.ToString());
                    if (tag > 0)
                    {
                        decimal value = (d_spinResultValue.Value / oldTag) * tag;
                        oldTag = tag;
                        d_spinResultValue.Value = value;
                        d_groupBox1.SetRedBorder();
                        ;
                    }
                    d_btnOk.Enabled = true;
                    d_btnReject.Enabled = true;
                    d_spinResultValue.Enabled = false;
                }
            }
            catch (Exception ex)
            {
                Logger.WriteLogFile(ex);
                One1.Controls.CustomMessageBox.Show("error");
                throw;
            }
        }

        private void d_spinResultValue_TextChanged(object sender, EventArgs e)
        {
            var spinEditor = (RadSpinEditor)sender;
            try
            {
                if (spinEditor.Text.Length > 0 && int.Parse(spinEditor.Text) > 0 || spinEditor.Value > 0)
                {
                    d_groupBox1.Enabled = true;
                    d_cbBig.Enabled = false;
                    d_groupBox1.SetRedBorder();
                }
                else
                {
                    d_cbBig.Enabled = true;
                    d_groupBox1.Enabled = false;
                    d_btnReject.Enabled = false;
                    d_groupBox1.Enabled = false;
                    d_groupBox1.RemoveRedBorder();
                    InitRadioBtnList();
                }
                d_btnOk.Enabled = false;
            }
            catch (Exception ex)
            {
                Logger.WriteLogFile(ex);
            }
        }

        private void AuthorisatioRejectnXml()
        {
            var objDom = new DOMDocument();

            IXMLDOMElement objResultauthorization = objDom.createElement("result-request");
            objDom.appendChild(objResultauthorization);

            IXMLDOMElement objLoad = objDom.createElement("load");
            objLoad.setAttribute("entity", "TEST");

            objLoad.setAttribute("id", _currentResult.TestId);
            objLoad.setAttribute("mode", "authorisation");

            objResultauthorization.appendChild(objLoad);

            IXMLDOMElement objResultReject = objDom.createElement("result-reject");
            objResultReject.setAttribute("result-id", _currentResult.ResultId);
            objLoad.appendChild(objResultReject);
            var resXML = new DOMDocument();

            string s = _processXml.ProcessXMLWithResponse(objDom, resXML);
            var success = s == "";


            //For testing

            //   objDom.save(@"C:\temp\docENTITYAu.xml");
            // resXML.save(@"C:\temp\resENTITYAu.xml");
        }

        private void InitialDilutionPanel()
        {
            lbl_dilutionClient.Text = "";
            d_txtResultID.Text = string.Empty;
            d_txtResultID.Enabled = true;
            d_txtResultID.Focus();
            d_spinResultValue.Value = 0;
            d_spinResultValue.Enabled = false;
            d_btnClean.Enabled = true;
            d_txtResultName.Text = string.Empty;
            d_btnReject.Enabled = false;
            d_cbBig.Checked = false;
            d_cbBig.Enabled = false;
            d_pictureBox.Enabled = true;
            d_pictureBox.Image = null;
            d_btnClose.Enabled = true;
            d_lblError.Text = "";
            _getResultXmlRes = null;
            d_groupBox1.Enabled = false;
            d_lblDuplicates.Text = "";
            InitRadioBtnList();
            d_groupBox1.RemoveRedBorder();
            oldTag = 1;
            d_btnOk.Enabled = false;
            panelDilution.Enabled = true;
            d_txtResultID.Select();
            d_txtResultID.Focus();

        }


        /// <summary>
        /// Init radio buttons mode
        /// </summary>
        private void InitRadioBtnList()
        {

            //Set all radioButtons false
            GetRadioBtnList().Foreach(x => x.IsChecked = false);
        }

        private int GetDuplicates()
        {
            return dal.GetResultDuplicates(_currentResult.TestId, _currentResult.ResultTemplateId);
        }

        private IEnumerable<RadRadioButton> GetRadioBtnList()
        {

            //Do it radioButton list
            List<RadRadioButton> rbl = d_groupBox1.Controls.OfType<RadRadioButton>().ToList();
            return rbl;
        }

        #endregion

        private void SingleResultEntry_Resize(object sender, EventArgs e)
        {
            lblHeader.Location = new Point(Width / 2 - lblHeader.Width / 2, lblHeader.Location.Y);
            panelDilution.Location = new Point(Width / 2 - panelDilution.Width / 2, panelDilution.Location.Y);
            regularPanel.Location = new Point(Width / 2 - regularPanel.Width / 2, regularPanel.Location.Y);

        }

        private void radButton1_Click(object sender, EventArgs e)
        {


            _currentResult.Status = "X";
            dal.SaveChanges();
            InitialMode();
            //SetResultValue(spinResultValue, cbNotCount);
            //AuthorisatioRejectnXml();
            //    ReCalculate();
            panelDilution.Enabled = true;
        }

        private void timerFocus_Tick(object sender, EventArgs e)
        {
            if (WithDilution)
            {
                d_txtResultID.Focus();
            }
            else
            {
                txtResultID.Focus();
            }
            timerFocus.Stop();

        }






    }
}